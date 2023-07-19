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

// TODO: Сейчас Paragraph.Recalculate_FastWholeParagraph работает только на добавлении текста, надо переделать
//       алгоритм определения изменений, чтобы данная функция работала и при других изменениях.

// Import
var c_oAscLineDrawingRule = AscCommon.c_oAscLineDrawingRule;
var align_Left = AscCommon.align_Left;
var hdrftr_Header = AscCommon.hdrftr_Header;
var hdrftr_Footer = AscCommon.hdrftr_Footer;
var c_oAscFormatPainterState = AscCommon.c_oAscFormatPainterState;
var changestype_None = AscCommon.changestype_None;
var changestype_Paragraph_Content = AscCommon.changestype_Paragraph_Content;
var changestype_2_Element_and_Type = AscCommon.changestype_2_Element_and_Type;
var changestype_2_ElementsArray_and_Type = AscCommon.changestype_2_ElementsArray_and_Type;
var g_oTableId = AscCommon.g_oTableId;
var History = AscCommon.History;

var c_oAscHAnchor = Asc.c_oAscHAnchor;
var c_oAscXAlign = Asc.c_oAscXAlign;
var c_oAscYAlign = Asc.c_oAscYAlign;
var c_oAscVAnchor = Asc.c_oAscVAnchor;
var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;
var c_oAscSectionBreakType    = Asc.c_oAscSectionBreakType;

var Page_Width     = 210;
var Page_Height    = 297;

var X_Left_Margin   = 30;  // 3   cm
var X_Right_Margin  = 15;  // 1.5 cm
var Y_Bottom_Margin = 20;  // 2   cm
var Y_Top_Margin    = 20;  // 2   cm

var X_Right_Field  = Page_Width  - X_Right_Margin;
var Y_Bottom_Field = Page_Height - Y_Bottom_Margin;

var docpostype_Content        = 0x00;
var docpostype_HdrFtr         = 0x02;
var docpostype_DrawingObjects = 0x03;
var docpostype_Footnotes      = 0x04;
var docpostype_Endnotes       = 0x05;

var selectionflag_Common        = 0x000;
var selectionflag_Numbering     = 0x001; // Выделена нумерация
var selectionflag_NumberingCur  = 0x002; // Выделена нумерация и данный параграф является текущим

var search_Common              = 0x0000; // Поиск в простом тексте
var search_Header              = 0x0100; // Поиск в верхнем колонтитуле
var search_Footer              = 0x0200; // Поиск в нижнем колонтитуле
var search_Footnote            = 0x0400; // Поиск в сноске
var search_Endnote             = 0x0800; // Поиск в концевой сноске

var search_HdrFtr_All          = 0x0001; // Поиск в колонтитуле, который находится на всех страницах
var search_HdrFtr_All_no_First = 0x0002; // Поиск в колонтитуле, который находится на всех страницах, кроме первой
var search_HdrFtr_First        = 0x0003; // Поиск в колонтитуле, который находится только на первой страниц
var search_HdrFtr_Even         = 0x0004; // Поиск в колонтитуле, который находится только на четных страницах
var search_HdrFtr_Odd          = 0x0005; // Поиск в колонтитуле, который находится только на нечетных страницах, включая первую
var search_HdrFtr_Odd_no_First = 0x0006; // Поиск в колонтитуле, который находится только на нечетных страницах, кроме первой

// Типы которые возвращают классы CParagraph и CTable после пересчета страницы
var recalcresult_NextElement = 0x01; // Пересчитываем следующий элемент
var recalcresult_PrevPage    = 0x02; // Пересчитываем заново предыдущую страницу
var recalcresult_CurPage     = 0x04; // Пересчитываем заново текущую страницу
var recalcresult_NextPage    = 0x08; // Пересчитываем следующую страницу
var recalcresult_NextLine    = 0x10; // Пересчитываем следующую строку
var recalcresult_CurLine     = 0x20; // Пересчитываем текущую строку
var recalcresult_CurPagePara = 0x40; // Специальный случай, когда мы встретили картинку в начале параграфа
var recalcresult_ParaMath    = 0x80; // Пересчитываем заново неинлайновую формулу

var recalcresultflags_Column            = 0x010000; // Пересчитываем только колонку
var recalcresultflags_Page              = 0x020000; // Пересчитываем всю страницу
var recalcresultflags_LastFromNewPage   = 0x040000; // Используется совсместно с recalcresult_NextPage, означает, что начало последнего элемента нужно перенести на новую страницу
var recalcresultflags_LastFromNewColumn = 0x080000; // Используется совсместно с recalcresult_NextPage, означает, что начало последнего элемента нужно перенести на новую колонку
var recalcresultflags_Footnotes         = 0x010000; // Сообщаем, что необходимо пересчитать сноски на данной странице

// Типы которые возвращают классы CDocument и CDocumentContent после пересчета страницы
var recalcresult2_End      = 0x00; // Документ рассчитан до конца
var recalcresult2_NextPage = 0x01; // Рассчет нужно продолжить
var recalcresult2_CurPage  = 0x02; // Нужно заново пересчитать данную страницу

var document_EditingType_Common = 0x00; // Обычный режим редактирования
var document_EditingType_Review = 0x01; // Режим рецензирования

var keydownflags_PreventDefault  = 0x0001;
var keydownflags_PreventKeyPress = 0x0002;

var keydownresult_PreventNothing  = 0x0000;
var keydownresult_PreventDefault  = 0x0001;
var keydownresult_PreventKeyPress = 0x0002;
var keydownresult_PreventAll      = 0xFFFF;

var MEASUREMENT_MAX_MM_VALUE = 1000; // Маскимальное значение в мм, используемое в документе (MS Word) - 55,87 см, или 558,7 мм.

var document_recalcresult_FastFlag = 0x1000;

var document_recalcresult_NoRecal       = 0x0000;
var document_recalcresult_FastRange     = 0x0001 | document_recalcresult_FastFlag;
var document_recalcresult_FastParagraph = 0x0002 | document_recalcresult_FastFlag;
var document_recalcresult_LongRecalc    = 0x00FF;

AscWord.ViewPositionType = {
	Common         : 0x00,
	Cursor         : 0x01,
	SelectionStart : 0x02,
	SelectionEnd   : 0x03
};

function CDocumentColumnProps()
{
    this.W     = 0;
    this.Space = 0;
}
CDocumentColumnProps.prototype.put_W = function(W)
{
    this.W = W;
};
CDocumentColumnProps.prototype.get_W = function()
{
    return this.W;
};
CDocumentColumnProps.prototype.put_Space = function(Space)
{
    this.Space = Space;
};
CDocumentColumnProps.prototype.get_Space = function()
{
    return this.Space;
};

function CDocumentColumnsProps()
{
    this.EqualWidth = true;
    this.Num        = 1;
    this.Sep        = false;
    this.Space      = 30;

    this.Cols       = [];

    this.TotalWidth = 230;
}
CDocumentColumnsProps.prototype.From_SectPr = function(SectPr)
{
    var Columns = SectPr.Columns;

    this.TotalWidth = SectPr.GetContentFrameWidth();
    this.EqualWidth = Columns.EqualWidth;
    this.Num        = Columns.Num;
    this.Sep        = Columns.Sep;
    this.Space      = Columns.Space;

    for (var Index = 0, Count = Columns.Cols.length; Index < Count; ++Index)
    {
        var Col = new CDocumentColumnProps();
        Col.put_W(Columns.Cols[Index].W);
        Col.put_Space(Columns.Cols[Index].Space);
        this.Cols[Index] = Col;
    }
};
CDocumentColumnsProps.prototype.get_EqualWidth = function()
{
    return this.EqualWidth;
};
CDocumentColumnsProps.prototype.put_EqualWidth = function(EqualWidth)
{
    this.EqualWidth = EqualWidth;
};
CDocumentColumnsProps.prototype.get_Num = function()
{
    return this.Num;
};
CDocumentColumnsProps.prototype.put_Num = function(Num)
{
    this.Num = Num;
};
CDocumentColumnsProps.prototype.get_Sep = function()
{
    return this.Sep;
};
CDocumentColumnsProps.prototype.put_Sep = function(Sep)
{
    this.Sep = Sep;
};
CDocumentColumnsProps.prototype.get_Space = function()
{
    return this.Space;
};
CDocumentColumnsProps.prototype.put_Space = function(Space)
{
    this.Space = Space;
};
CDocumentColumnsProps.prototype.get_ColsCount = function()
{
    return this.Cols.length;
};
CDocumentColumnsProps.prototype.get_Col = function(Index)
{
    return this.Cols[Index];
};
CDocumentColumnsProps.prototype.put_Col = function(Index, Col)
{
    this.Cols[Index] = Col;
};
CDocumentColumnsProps.prototype.put_ColByValue = function(Index, W, Space)
{
    var Col = new CDocumentColumnProps();
    Col.put_W(W);
    Col.put_Space(Space);
    this.Cols[Index] = Col;
};
CDocumentColumnsProps.prototype.get_TotalWidth = function()
{
    return this.TotalWidth;
};

function CDocumentSectionProps(oSectPr, oLogicDocument)
{
    if (oSectPr && oLogicDocument)
    {
        this.W      = oSectPr.GetPageWidth();
        this.H      = oSectPr.GetPageHeight();
        this.Orient = oSectPr.GetOrientation();

        this.Left   = oSectPr.GetPageMarginLeft();
        this.Top    = oSectPr.GetPageMarginTop();
        this.Right  = oSectPr.GetPageMarginRight();
        this.Bottom = oSectPr.GetPageMarginBottom();

        this.Header = oSectPr.GetPageMarginHeader();
        this.Footer = oSectPr.GetPageMarginFooter();

		this.Gutter        = oSectPr.GetGutter();
		this.GutterRTL     = oSectPr.IsGutterRTL();
		this.GutterAtTop   = oLogicDocument.IsGutterAtTop();
		this.MirrorMargins = oLogicDocument.IsMirrorMargins();
    }
    else
    {
        this.W      = undefined;
        this.H      = undefined;
        this.Orient = undefined;

        this.Left   = undefined;
        this.Top    = undefined;
        this.Right  = undefined;
        this.Bottom = undefined;

        this.Header = undefined;
        this.Footer = undefined;

        this.Gutter        = undefined;
        this.GutterRTL     = undefined;
        this.GutterAtTop   = undefined;
        this.MirrorMargins = undefined;
    }
}
CDocumentSectionProps.prototype.get_W = function()
{
    return this.W;
};
CDocumentSectionProps.prototype.put_W = function(W)
{
    this.W = W;
};
CDocumentSectionProps.prototype.get_H = function()
{
    return this.H;
};
CDocumentSectionProps.prototype.put_H = function(H)
{
    this.H = H;
};
CDocumentSectionProps.prototype.get_Orientation = function()
{
    return this.Orient;
};
CDocumentSectionProps.prototype.put_Orientation = function(Orient)
{
    this.Orient = Orient;
};
CDocumentSectionProps.prototype.get_LeftMargin = function()
{
    return this.Left;
};
CDocumentSectionProps.prototype.put_LeftMargin = function(Left)
{
    this.Left = Left;
};
CDocumentSectionProps.prototype.get_TopMargin = function()
{
    return this.Top;
};
CDocumentSectionProps.prototype.put_TopMargin = function(Top)
{
    this.Top = Top;
};
CDocumentSectionProps.prototype.get_RightMargin = function()
{
    return this.Right;
};
CDocumentSectionProps.prototype.put_RightMargin = function(Right)
{
    this.Right = Right;
};
CDocumentSectionProps.prototype.get_BottomMargin = function()
{
    return this.Bottom;
};
CDocumentSectionProps.prototype.put_BottomMargin = function(Bottom)
{
    this.Bottom = Bottom;
};
CDocumentSectionProps.prototype.get_HeaderDistance = function()
{
    return this.Header;
};
CDocumentSectionProps.prototype.put_HeaderDistance = function(Header)
{
    this.Header = Header;
};
CDocumentSectionProps.prototype.get_FooterDistance = function()
{
    return this.Footer;
};
CDocumentSectionProps.prototype.put_FooterDistance = function(Footer)
{
    this.Footer = Footer;
};
CDocumentSectionProps.prototype.get_Gutter = function()
{
	return this.Gutter;
};
CDocumentSectionProps.prototype.put_Gutter = function(nGutter)
{
	this.Gutter = nGutter;
};
CDocumentSectionProps.prototype.get_GutterRTL = function()
{
	return this.GutterRTL;
};
CDocumentSectionProps.prototype.put_GutterRTL = function(isRTL)
{
	this.GutterRTL = isRTL;
};
CDocumentSectionProps.prototype.get_GutterAtTop = function()
{
	return this.GutterAtTop;
};
CDocumentSectionProps.prototype.put_GutterAtTop = function(isAtTop)
{
	this.GutterAtTop = isAtTop;
};
CDocumentSectionProps.prototype.get_MirrorMargins = function()
{
	return this.MirrorMargins;
};
CDocumentSectionProps.prototype.put_MirrorMargins = function(isMirrorMargins)
{
	this.MirrorMargins = isMirrorMargins;
};

function CDocumentRecalculateState()
{
    this.Id           = null;

    this.PageIndex    = 0;
    this.SectionIndex = 0;
    this.ColumnIndex  = 0;
    this.Start        = true;
    this.StartIndex   = 0;
    this.StartPage    = 0;
	this.Endnotes     = false;

	this.TimerStartTime = 0;
	this.TimerStartPage = 0;

	this.StartPagesCount   = 2;     // количество страниц, которые мы обсчитываем в первый раз без таймеров
    this.ResetStartElement = false;
    this.MainStartPos      = -1;

    this.UseRecursion = true;
    this.Continue     = false; // параметр сигнализирующий, о том что нужно продолжить пересчет (для нерекурсивного метода)
}

function CDocumentRecalculateHdrFtrPageCountState()
{
	this.Id        = null;
	this.PageIndex = 0;
	this.PageCount = -1;
}

function Document_Recalculate_Page()
{
    var LogicDocument = editor.WordControl.m_oLogicDocument;
	LogicDocument.FullRecalc.TimerStartTime = performance.now();
    LogicDocument.Recalculate_Page();
}

function Document_Recalculate_HdrFtrPageCount()
{
	var LogicDocument = editor.WordControl.m_oLogicDocument;
	LogicDocument.private_RecalculateHdrFtrPageCountUpdate();
}

function CDocumentPageSection()
{
    this.Pos    = 0;
    this.EndPos = -1;

    this.Y      = 0;
    this.YLimit = 0;

    this.YLimit2 = 0;

    this.Index   = -1;

    this.Columns = [];
    this.ColumnsSep = false;

    this.IterationsCount       = 0;
    this.CurrentY              = 0;
    this.RecalculateBottomLine = true;
    this.CanDecrease           = true;
    this.WasIncrease           = false; // Было ли хоть раз увеличение
    this.IterationStep         = 10;
    this.IterationDirection    = 0;
}
/**
 * Инициализируем параметры данной секции
 * @param {number} nPageAbs
 * @param {CSectionPr} oSectPr
 * @param {Number} nSectionIndex
 */
CDocumentPageSection.prototype.Init = function(nPageAbs, oSectPr, nSectionIndex)
{
	var oFrame  = oSectPr.GetContentFrame(nPageAbs);
	var nX      = oFrame.Left;
	var nXLimit = oFrame.Right;

	for (var nCurColumn = 0, nColumnsCount = oSectPr.GetColumnsCount(); nCurColumn < nColumnsCount; ++nCurColumn)
	{
		this.Columns[nCurColumn] = new CDocumentPageColumn();

		this.Columns[nCurColumn].X      = nX;
		this.Columns[nCurColumn].XLimit = nColumnsCount - 1 === nCurColumn ? nXLimit : nX + oSectPr.GetColumnWidth(nCurColumn);

		nX += oSectPr.GetColumnWidth(nCurColumn) + oSectPr.GetColumnSpace(nCurColumn);
	}
	this.ColumnsSep = oSectPr.GetColumnSep();

	this.Y       = oFrame.Top;
	this.YLimit  = oFrame.Bottom;
	this.YLimit2 = oFrame.Bottom;
	this.Index   = nSectionIndex;
};
CDocumentPageSection.prototype.Copy = function()
{
    var NewSection = new CDocumentPageSection();

    NewSection.Pos    = this.Pos;
    NewSection.EndPos = this.EndPos;
    NewSection.Y      = this.Y;
    NewSection.YLimit = this.YLimit;

    for (var ColumnIndex = 0, Count = this.Columns.length; ColumnIndex < Count; ++ColumnIndex)
    {
		NewSection.Columns[ColumnIndex] = this.Columns[ColumnIndex].Copy();
    }

    return NewSection;
};
CDocumentPageSection.prototype.Shift = function(Dx, Dy)
{
    this.Y      += Dy;
    this.YLimit += Dy;

    for (var ColumnIndex = 0, Count = this.Columns.length; ColumnIndex < Count; ++ColumnIndex)
    {
        this.Columns[ColumnIndex].Shift(Dx, Dy);
    }
};
/**
 * Происходи ли процесс расчета нижней границы разрыва секции на текущей странице
 * @returns {boolean}
 */
CDocumentPageSection.prototype.IsCalculatingSectionBottomLine = function()
{
    return (this.IterationsCount > 0 && true === this.RecalculateBottomLine);
};
/**
 * Можно ли расчитывать нижнюю границу разрыва секции на текущей странице
 * @returns {boolean}
 */
CDocumentPageSection.prototype.CanRecalculateBottomLine = function()
{
    return this.RecalculateBottomLine;
};
/**
 * Запрещаем возможность расчета нижней границы разрыва секции на текущей страницы
 */
CDocumentPageSection.prototype.ForbidRecalculateBottomLine = function()
{
	this.RecalculateBottomLine = false;
};
CDocumentPageSection.prototype.GetY = function()
{
    return this.Y;
};
CDocumentPageSection.prototype.GetYLimit = function()
{
    if (0 === this.IterationsCount)
        return this.YLimit;
    else
        return this.CurrentY;
};
/**
 * Производим шаг рассчета нижней границы рарзрыва секции
 * @param {boolean} isIncrease
 * @returns {number}
 */
CDocumentPageSection.prototype.IterateBottomLineCalculation = function(isIncrease)
{
	// Алгоритм следующий:
	// На первом шаге мы прогнозируем положение границы по уже имеющемуся объему текста и
	// ширине колонок.
	// Далее мы сдвигаем границу на значение IterationStep, вверх или вниз в зависимости
	// от результата расчета страницы. При перемене направления сдвига мы всегда уменьшаем шаг
	// в 2 раза. Также мы можем уменьшать шаг в 2 раза, если сдвигаем границу вверх, и хотябы
	// раз до этого двигали ее вниз. Останавливаем итерацию при попытке подвинуть границу наверх,
	// когда шаг итерации становится менее 2мм.

	if (0 === this.IterationsCount)
	{
		// Пытаемся заранее спрогнозировать позицию, где должно быть разделение. Учитывая, что колонки могут быть разной
		// ширины, мы расчитываем суммарную занимаемую текстом область. Делим ее по колонкам, с учетом их суммарной
		// ширины.

		var nSumArea = 0, nSumWidth = 0;
		for (var nColumnIndex = 0, nColumnsCount = this.Columns.length; nColumnIndex < nColumnsCount; ++nColumnIndex)
		{
			var oColumn = this.Columns[nColumnIndex];
			if (true !== oColumn.Empty)
				nSumArea += (oColumn.Bounds.Bottom - this.Y) * (oColumn.XLimit - oColumn.X);

			nSumWidth += oColumn.XLimit - oColumn.X;
		}

		if (nSumWidth > 0.001)
			this.CurrentY = this.Y + nSumArea / nSumWidth;
		else
			this.CurrentY = this.Y;
	}
	else
	{
		if (false === isIncrease)
		{
			if (this.IterationDirection > 0 || this.WasIncrease)
				this.IterationStep /= 2;

			this.CurrentY -= this.IterationStep;
			this.IterationDirection = -1;

			if (this.CurrentY < this.Y)
			{
				// Такое может быть, когда у нас всего одна строка в начале страницы, которую мы размещаем всегда
				this.CurrentY    = this.Y;
				this.CanDecrease = false;
			}
		}
		else
		{
			if (this.IterationDirection < 0)
				this.IterationStep /= 2;

			this.CurrentY += this.IterationStep;
			this.IterationDirection = 1;

			this.WasIncrease = true;
		}
	}

	if (this.IterationStep < 2)
		this.CanDecrease = false;

	this.CurrentY = Math.min(this.CurrentY, this.YLimit2);

	this.IterationsCount++;
	return this.CurrentY;
};
CDocumentPageSection.prototype.Reset_Columns = function()
{
    for (var ColumnIndex = 0, Count = this.Columns.length; ColumnIndex < Count; ++ColumnIndex)
    {
        this.Columns[ColumnIndex].Reset();
    }
};
/**
 * Можем ли мы провести еще одну итерацию с уменьшением нижней границы
 * @returns {boolean}
 */
CDocumentPageSection.prototype.CanDecreaseBottomLine = function()
{
	return this.CanDecrease;
};
CDocumentPageSection.prototype.CanIncreaseBottomLine = function()
{
	// Данная функция не должна возвращать false, если возвращает, значит неправильно работает алгоритм по вычислению
	// нижней границы continuous секции
	// if (!(this.YLimit2 - this.CurrentY > 0.001))
	// 	console.log("Bad continuous section calculate");

	return (this.YLimit2 - this.CurrentY > 0.001);
};

function CDocumentPageColumn()
{
    this.Bounds = new CDocumentBounds(0, 0, 0, 0);
    this.Pos    = 0;
    this.EndPos = -1;
    this.Empty  = true;

    this.X      = 0;
    this.Y      = 0;
    this.XLimit = 0;
    this.YLimit = 0;

    this.SpaceBefore = 0;
    this.SpaceAfter  = 0;
}
CDocumentPageColumn.prototype.Copy = function()
{
    var NewColumn = new CDocumentPageColumn();

    NewColumn.Bounds.CopyFrom(this.Bounds);
    NewColumn.Pos    = this.Pos;
    NewColumn.EndPos = this.EndPos;
    NewColumn.X      = this.X;
    NewColumn.Y      = this.Y;
    NewColumn.XLimit = this.XLimit;
    NewColumn.YLimit = this.YLimit;

    return NewColumn;
};
CDocumentPageColumn.prototype.Shift = function(Dx, Dy)
{
    this.X      += Dx;
    this.XLimit += Dx;
    this.Y      += Dy;
    this.YLimit += Dy;

    this.Bounds.Shift(Dx, Dy);
};
CDocumentPageColumn.prototype.Reset = function()
{
    this.Bounds.Reset();
    this.Pos    = 0;
    this.EndPos = -1;
    this.Empty  = true;

    this.X      = 0;
    this.Y      = 0;
    this.XLimit = 0;
    this.YLimit = 0;
};
CDocumentPageColumn.prototype.IsEmpty = function()
{
	return this.Empty;
};

function CDocumentPage()
{
    this.Width   = 0;
    this.Height  = 0;
    this.Margins =
    {
        Left   : 0,
        Right  : 0,
        Top    : 0,
        Bottom : 0
    };

    this.Bounds = new CDocumentBounds(0,0,0,0);
    this.Pos    = 0;
    this.EndPos = 0;

	this.X      = 0;
	this.Y      = 0;
	this.XLimit = 0;
	this.YLimit = 0;

	this.OriginX      = 0; // Начальные значения X, Y без учета использования функции Shift
	this.OriginY      = 0; // Используется, например, при расчете позиции автофигуры внутри ячейки таблицы,
	this.OriginXLimit = 0; // которая имеет вертикальное выравнивание по центру или по низу
	this.OriginTLimit = 0;

    this.Sections = [];

    this.EndSectionParas = [];

    this.ResetStartElement = false;

    this.Frames     = [];
    this.FlowTables = [];
}

CDocumentPage.prototype.GetStartPos = function()
{
	return this.Pos;
};
CDocumentPage.prototype.GetEndPos = function()
{
	return this.EndPos;
};
CDocumentPage.prototype.GetSection = function(nIndex)
{
	return this.Sections[nIndex] ? this.Sections[nIndex] : null;
};
CDocumentPage.prototype.Update_Limits = function(Limits)
{
	this.X      = Limits.X;
	this.XLimit = Limits.XLimit;
	this.Y      = Limits.Y;
	this.YLimit = Limits.YLimit;

	this.OriginX      = Limits.X;
	this.OriginY      = Limits.Y;
	this.OriginXLimit = Limits.XLimit;
	this.OriginYLimit = Limits.YLimit;
};
CDocumentPage.prototype.Shift = function(Dx, Dy)
{
    this.X      += Dx;
    this.XLimit += Dx;
    this.Y      += Dy;
    this.YLimit += Dy;

    this.Bounds.Shift(Dx, Dy);

    for (var SectionIndex = 0, Count = this.Sections.length; SectionIndex < Count; ++SectionIndex)
    {
        this.Sections[SectionIndex].Shift(Dx, Dy);
    }
};
CDocumentPage.prototype.Check_EndSectionPara = function(Element)
{
    var Count = this.EndSectionParas.length;
    for ( var Index = 0; Index < Count; Index++ )
    {
        if ( Element === this.EndSectionParas[Index] )
            return true;
    }

    return false;
};
CDocumentPage.prototype.Copy = function()
{
    var NewPage = new CDocumentPage();

    NewPage.Width          = this.Width;
    NewPage.Height         = this.Height;
    NewPage.Margins.Left   = this.Margins.Left;
    NewPage.Margins.Right  = this.Margins.Right;
    NewPage.Margins.Top    = this.Margins.Top;
    NewPage.Margins.Bottom = this.Margins.Bottom;

    NewPage.Bounds.CopyFrom(this.Bounds);
    NewPage.Pos    = this.Pos;
    NewPage.EndPos = this.EndPos;
    NewPage.X      = this.X;
    NewPage.Y      = this.Y;
    NewPage.XLimit = this.XLimit;
    NewPage.YLimit = this.YLimit;

    for (var SectionIndex = 0, Count = this.Sections.length; SectionIndex < Count; ++SectionIndex)
    {
        NewPage.Sections[SectionIndex] = this.Sections[SectionIndex].Copy();
    }

    return NewPage;
};
CDocumentPage.prototype.AddFrame = function(oFrame)
{
	if (-1 !== this.private_GetFrameIndex(oFrame.StartIndex))
		return -1;

	this.Frames.push(oFrame);
};
CDocumentPage.prototype.RemoveFrame = function(nStartIndex)
{
	var nPos = this.private_GetFrameIndex(nStartIndex);
	if (-1 === nPos)
		return;

	this.Frames.splice(nPos, 1);
};
CDocumentPage.prototype.private_GetFrameIndex = function(nStartIndex)
{
	for (var nIndex = 0, nCount = this.Frames.length; nIndex < nCount; ++nIndex)
	{
		if (nStartIndex === this.Frames[nIndex].StartIndex)
			return nIndex;
	}

	return -1;
};
CDocumentPage.prototype.AddFlowTable = function(oTable)
{
	if (-1 !== this.private_GetFlowTableIndex(oTable))
		return;

	this.FlowTables.push(oTable);
};
CDocumentPage.prototype.RemoveFlowTable = function(oTable)
{
	var nPos = this.private_GetFlowTableIndex(oTable);
	if (-1 === nPos)
		return;

	this.FlowTables.splice(nPos, 1);
};
CDocumentPage.prototype.private_GetFlowTableIndex = function(oTable)
{
	for (var nIndex = 0, nCount = this.FlowTables.length; nIndex < nCount; ++nIndex)
	{
		if (oTable === this.FlowTables[nIndex])
			return nIndex;
	}

	return -1;
};
CDocumentPage.prototype.IsFlowTable = function(oElement)
{
	return (-1 !== this.private_GetFlowTableIndex(oElement));
};
CDocumentPage.prototype.IsFrame = function(oElement)
{
	var nIndex = oElement.GetIndex();
	for (var nFrameIndex = 0, nFramesCount = this.Frames.length; nFrameIndex < nFramesCount; ++nFrameIndex)
	{
		if (this.Frames[nFrameIndex].StartIndex <= nIndex && nIndex < this.Frames[nFrameIndex].StartIndex + this.Frames[nFrameIndex].FlowCount)
			return true;
	}

	return false;
};
CDocumentPage.prototype.CheckFrameClipStart = function(nIndex, oGraphics, oDrawingDocument)
{
	for (var sId in this.Frames)
	{
		var oFrame = this.Frames[sId];

		if (oFrame.StartIndex === nIndex)
		{
			var nPixelError = oDrawingDocument.GetMMPerDot(1);

			var nL = oFrame.CalculatedFrame.L2 - nPixelError;
			var nT = oFrame.CalculatedFrame.T2 - nPixelError;
			var nH = oFrame.CalculatedFrame.H2 + 2 * nPixelError;
			var nW = oFrame.CalculatedFrame.W2 + 2 * nPixelError;

			oGraphics.SaveGrState();
			oGraphics.AddClipRect(nL, nT, nW, nH);
			return;
		}
	}
};
CDocumentPage.prototype.CheckFrameClipStart = function(nIndex, oGraphics)
{
	for (var sId in this.Frames)
	{
		var oFrame = this.Frames[sId];

		if (oFrame.StartIndex + oFrame.FlowCount - 1 === nIndex)
		{
			oGraphics.RestoreGrState();
			return
		}
	}
};


function CStatistics(LogicDocument)
{
    this.LogicDocument = LogicDocument;
    this.Api           = LogicDocument.Get_Api();

    this.Id       = null; // Id таймера для подсчета всего кроме страниц
    this.PagesId  = null; // Id таймера для подсчета страниц

    this.StartPos = 0;

    this.Pages           = 0;
    this.Words           = 0;
    this.Paragraphs      = 0;
    this.SymbolsWOSpaces = 0;
    this.SymbolsWhSpaces = 0;
}

CStatistics.prototype =
{
//-----------------------------------------------------------------------------------
// Функции для запуска и остановки сбора статистики
//-----------------------------------------------------------------------------------
    Start : function()
    {
        this.StartPos = 0;
        this.CurPage  = 0;

        this.Pages           = 0;
        this.Words           = 0;
        this.Paragraphs      = 0;
        this.SymbolsWOSpaces = 0;
        this.SymbolsWhSpaces = 0;


        var LogicDocument = this.LogicDocument;
        this.PagesId = setTimeout(function(){LogicDocument.Statistics_GetPagesInfo();}, 1);
        this.Id      = setTimeout(function(){LogicDocument.Statistics_GetParagraphsInfo();}, 1);
        this.Send();
    },

    Next_ParagraphsInfo : function(StartPos)
    {
        this.StartPos = StartPos;
        var LogicDocument = this.LogicDocument;
        clearTimeout(this.Id);
        this.Id = setTimeout(function(){LogicDocument.Statistics_GetParagraphsInfo();}, 1);
        this.Send();
    },

    Next_PagesInfo : function()
    {
        var LogicDocument = this.LogicDocument;
        clearTimeout(this.PagesId);
        this.PagesId = setTimeout(function(){LogicDocument.Statistics_GetPagesInfo();}, 100);
        this.Send();
    },

    Stop_PagesInfo : function()
    {
        if (null !== this.PagesId)
        {
            clearTimeout(this.PagesId);
            this.PagesId = null;
        }

        this.Check_Stop();
    },

    Stop_ParagraphsInfo : function()
    {
        if (null != this.Id)
        {
            clearTimeout(this.Id);
            this.Id = null;
        }

        this.Check_Stop();
    },

    Check_Stop : function()
    {
        if (null === this.Id && null === this.PagesId)
        {
            this.Send();
            this.Api.sync_GetDocInfoEndCallback();
        }
    },

    Send : function()
    {
        var Stats =
        {
            PageCount      : this.Pages,
            WordsCount     : this.Words,
            ParagraphCount : this.Paragraphs,
            SymbolsCount   : this.SymbolsWOSpaces,
            SymbolsWSCount : this.SymbolsWhSpaces
        };

        this.Api.sync_DocInfoCallback(Stats);
    },
//-----------------------------------------------------------------------------------
// Функции для пополнения статистики
//-----------------------------------------------------------------------------------
    Add_Paragraph : function (Count)
    {
        if ( "undefined" != typeof( Count ) )
            this.Paragraphs += Count;
        else
            this.Paragraphs++;
    },

    Add_Word : function(Count)
    {
        if ( "undefined" != typeof( Count ) )
            this.Words += Count;
        else
            this.Words++;
    },

    Update_Pages : function(PagesCount)
    {
        this.Pages = PagesCount;
    },

    Add_Symbol : function(bSpace)
    {
        this.SymbolsWhSpaces++;
        if ( true != bSpace )
            this.SymbolsWOSpaces++;
    }
};

function CDocumentRecalcInfo()
{
    this.FlowObject                = null;   // Текущий float-объект, который мы пересчитываем
    this.FlowObjectPageBreakBefore = false;  // Нужно ли перед float-объектом поставить pagebreak
    this.FlowObjectPage            = 0;      // Количество обработанных страниц
    this.FlowObjectElementsCount   = 0;      // Количество элементов подряд идущих в рамке (только для рамок)
    this.RecalcResult              = recalcresult_NextElement;

    this.WidowControlParagraph     = null;   // Параграф, который мы пересчитываем из-за висячих строк
    this.WidowControlLine          = -1;     // Номер строки, перед которой надо поставить разрыв страницы
    this.WidowControlReset         = false;  //

    this.KeepNextParagraph         = null;    // Параграф, который надо пересчитать из-за того, что следующий начался с новой страницы
    this.KeepNextEndParagraph      = null;    // Параграф, на котором определилось, что нужно делать пересчет предыдущих

    this.FrameRecalc               = false;   // Пересчитываем ли рамку
    this.ParaMath                  = null;

    this.FootnoteReference         = null;    // Ссылка на сноску, которую мы пересчитываем
    this.FootnotePage              = 0;
    this.FootnoteColumn            = 0;

    this.AdditionalInfo            = null;

    this.NeedRecalculateFromStart  = false;
    this.Paused                    = false;
}

CDocumentRecalcInfo.prototype =
{
    Reset : function()
    {
        this.FlowObject                = null;
        this.FlowObjectPageBreakBefore = false;
        this.FlowObjectPage            = 0;
        this.FlowObjectElementsCount   = 0;
        this.RecalcResult              = recalcresult_NextElement;

        this.WidowControlParagraph     = null;
        this.WidowControlLine          = -1;
        this.WidowControlReset         = false;

        this.KeepNextParagraph         = null;
        this.KeepNextEndParagraph      = null;

        this.ParaMath                  = null;

        this.FootnoteReference         = null;
        this.FootnotePage              = 0;
        this.FootnoteColumn            = 0;
    },

    Can_RecalcObject : function()
    {
        if (null === this.FlowObject && null === this.WidowControlParagraph && null === this.KeepNextParagraph && null == this.ParaMath && null === this.FootnoteReference)
            return true;

        return false;
    },

    Can_RecalcWidowControl : function()
    {
        // TODO: Пока нельзя отдельно проверять объекты, вызывающие пересчет страниц, потому что возможны зависания.
        //       Надо обдумать новую более грамотную схему, при которой можно будет вызывать независимые пересчеты заново.
        return this.Can_RecalcObject();
    },

    Set_FlowObject : function(Object, RelPage, RecalcResult, ElementsCount, AdditionalInfo)
    {
        this.FlowObject              = Object;
        this.FlowObjectPage          = RelPage;
        this.FlowObjectElementsCount = ElementsCount;
        this.RecalcResult            = RecalcResult;
        this.AdditionalInfo          = AdditionalInfo;
    },

    Set_ParaMath : function(Object)
    {
        this.ParaMath = Object;
    },

    Check_ParaMath: function(ParaMath)
    {
        if ( ParaMath === this.ParaMath )
            return true;

        return false;
    },

    Check_FlowObject : function(FlowObject)
    {
        if ( FlowObject === this.FlowObject )
            return true;

        return false;
    },

    Set_PageBreakBefore : function(Value)
    {
        this.FlowObjectPageBreakBefore = Value;
    },

    Is_PageBreakBefore : function()
    {
        return this.FlowObjectPageBreakBefore;
    },

    Set_KeepNext : function(Paragraph, EndParagraph)
    {
        this.KeepNextParagraph    = Paragraph;
        this.KeepNextEndParagraph = EndParagraph;
    },

    Check_KeepNext : function(Paragraph)
    {
        if ( Paragraph === this.KeepNextParagraph )
            return true;

        return false;
    },

    Check_KeepNextEnd : function(Paragraph)
    {
        if (Paragraph === this.KeepNextEndParagraph)
            return true;

        return false;
    },

    Need_ResetWidowControl : function()
    {
        this.WidowControlReset = true;
    },

    Reset_WidowControl : function()
    {
        if (true === this.WidowControlReset)
        {
            this.WidowControlParagraph = null;
            this.WidowControlLine      = -1;
            this.WidowControlReset     = false;
        }
    },

    Set_WidowControl : function(Paragraph, Line)
    {
        this.WidowControlParagraph = Paragraph;
        this.WidowControlLine      = Line;
    },

    Check_WidowControl : function(Paragraph, Line)
    {
        if (Paragraph === this.WidowControlParagraph && Line === this.WidowControlLine)
            return true;

        return false;
    },

    Set_FrameRecalc  : function(Value)
    {
        this.FrameRecalc = Value;
    },

    Set_FootnoteReference : function(oFootnoteReference, nPageAbs, nColumnAbs)
    {
        this.FootnoteReference = oFootnoteReference;
        this.FootnotePage      = nPageAbs;
		this.FootnoteColumn    = nColumnAbs;
    },

    Check_FootnoteReference : function(oFootnoteReference)
    {
        return (this.FootnoteReference === oFootnoteReference ? true : false);
    },

    Reset_FootnoteReference : function()
    {
        this.FootnoteReference         = null;
		this.FootnotePage              = 0;
		this.FootnoteColumn            = 0;
		this.FlowObjectPageBreakBefore = false;
    },

	Set_NeedRecalculateFromStart : function(isNeedRecalculate)
	{
		this.NeedRecalculateFromStart = isNeedRecalculate;
	},

	Is_NeedRecalculateFromStart : function()
	{
		return this.NeedRecalculateFromStart;
	}

};

function CDocumentFieldsManager()
{
	this.m_aFields = [];

	this.m_oMailMergeFields = {};

	this.m_aComplexFields = [];
	this.m_oCurrentComplexField = null;
	
	this.ComplexFieldCounter = 0;
}

CDocumentFieldsManager.prototype.GetNewComplexFieldId = function()
{
	return "" + (++this.ComplexFieldCounter);
};
CDocumentFieldsManager.prototype.Register_Field = function(oField)
{
    this.m_aFields.push(oField);

    var nFieldType = oField.Get_FieldType();
    if (fieldtype_MERGEFIELD === nFieldType)
    {
        var sName = oField.Get_Argument(0);
        if (undefined !== sName)
        {
            if (undefined === this.m_oMailMergeFields[sName])
                this.m_oMailMergeFields[sName] = [];

            this.m_oMailMergeFields[sName].push(oField);
        }
    }
};
CDocumentFieldsManager.prototype.Update_MailMergeFields = function(Map)
{
    for (var FieldName in this.m_oMailMergeFields)
    {
        for (var Index = 0, Count = this.m_oMailMergeFields[FieldName].length; Index < Count; Index++)
        {
            var oField = this.m_oMailMergeFields[FieldName][Index];
            oField.Map_MailMerge(Map[FieldName]);
        }
    }
};
CDocumentFieldsManager.prototype.Replace_MailMergeFields = function(Map)
{
    for (var FieldName in this.m_oMailMergeFields)
    {
        for (var Index = 0, Count = this.m_oMailMergeFields[FieldName].length; Index < Count; Index++)
        {
            var oField = this.m_oMailMergeFields[FieldName][Index];
            oField.Replace_MailMerge(Map[FieldName]);
        }
    }
};
CDocumentFieldsManager.prototype.Restore_MailMergeTemplate = function()
{
    if (true === AscCommon.CollaborativeEditing.Is_SingleUser())
    {
        // В такой ситуации мы восстанавливаем стандартный текст полей. Чтобы это сделать нам сначала надо пробежаться
        // по всем таким полям, определить параграфы, в которых необходимо восстановить поля. Залочить эти параграфы,
        // и произвести восстановление полей.

        var FieldsToRestore    = [];
        var FieldsRemain       = [];
        var ParagrapsToRestore = [];
        for (var FieldName in this.m_oMailMergeFields)
        {
            for (var Index = 0, Count = this.m_oMailMergeFields[FieldName].length; Index < Count; Index++)
            {
                var oField = this.m_oMailMergeFields[FieldName][Index];
                var oFieldPara = oField.GetParagraph();

                if (oFieldPara && oField.Is_NeedRestoreTemplate())
                {
                    var bNeedAddPara = true;
                    for (var ParaIndex = 0, ParasCount = ParagrapsToRestore.length; ParaIndex < ParasCount; ParaIndex++)
                    {
                        if (oFieldPara === ParagrapsToRestore[ParaIndex])
                        {
                            bNeedAddPara = false;
                            break;
                        }
                    }
                    if (true === bNeedAddPara)
                        ParagrapsToRestore.push(oFieldPara);

                    FieldsToRestore.push(oField);
                }
                else
                {
                    FieldsRemain.push(oField);
                }
            }
        }

        var LogicDocument = (ParagrapsToRestore.length > 0 ? ParagrapsToRestore[0].LogicDocument : null);
        if (LogicDocument && false === LogicDocument.Document_Is_SelectionLocked(changestype_None, { Type : changestype_2_ElementsArray_and_Type, Elements : ParagrapsToRestore, CheckType : changestype_Paragraph_Content }))
        {
            LogicDocument.StartAction(AscDFH.historydescription_Document_RestoreFieldTemplateText);
            for (var nIndex = 0, nCount = FieldsToRestore.length; nIndex < nCount; nIndex++)
            {
                var oField = FieldsToRestore[nIndex];
                oField.Restore_StandardTemplate();
            }

            for (var nIndex = 0, nCount = FieldsRemain.length; nIndex < nCount; nIndex++)
            {
                var oField = FieldsRemain[nIndex];
                oField.Restore_Template();
            }
            LogicDocument.FinalizeAction();
        }
        else
        {
            // Восстанавливать ничего не надо просто возващаем значения тимплейтов
            for (var FieldName in this.m_oMailMergeFields)
            {
                for (var Index = 0, Count = this.m_oMailMergeFields[FieldName].length; Index < Count; Index++)
                {
                    var oField = this.m_oMailMergeFields[FieldName][Index];
                    oField.Restore_Template();
                }
            }
        }
    }
    else
    {
        for (var FieldName in this.m_oMailMergeFields)
        {
            for (var Index = 0, Count = this.m_oMailMergeFields[FieldName].length; Index < Count; Index++)
            {
                var oField = this.m_oMailMergeFields[FieldName][Index];
                oField.Restore_Template();
            }
        }
    }
};
CDocumentFieldsManager.prototype.GetAllFieldsByType = function(nType)
{
	var arrFields = [];
	for (var nIndex = 0, nCount = this.m_aFields.length; nIndex < nCount; ++nIndex)
	{
		var oField = this.m_aFields[nIndex];
		if (nType === oField.Get_FieldType() && oField.IsUseInDocument())
		{
			arrFields.push(oField);
		}
	}
	return arrFields;
};
CDocumentFieldsManager.prototype.RegisterComplexField = function(oComplexField)
{
	this.m_aComplexFields.push(oComplexField);
};
CDocumentFieldsManager.prototype.GetCurrentComplexField = function()
{
	return this.m_oCurrentComplexField;
};
CDocumentFieldsManager.prototype.SetCurrentComplexField = function(oComplexField)
{
	if (this.m_oCurrentComplexField === oComplexField)
		return false;

	if (this.m_oCurrentComplexField)
		this.m_oCurrentComplexField.SetCurrent(false);

	this.m_oCurrentComplexField = oComplexField;

	if (this.m_oCurrentComplexField)
		this.m_oCurrentComplexField.SetCurrent(true);

	return true;
};

var selected_None              = -1;
var selected_DrawingObject     = 0;
var selected_DrawingObjectText = 1;

function CSelectedElementsInfo(oPr)
{
	this.m_bSkipTOC = !!(oPr && oPr.SkipTOC);
	this.m_bCheckAllSelection = !!(oPr && oPr.CheckAllSelection); // Проверять все выделение, или только 1 элемент
    this.m_bAddAllComplexFields = !!(oPr && oPr.AddAllComplexFields);


	this.m_bTable             = false; // Находится курсор или выделение целиком в какой-нибудь таблице
	this.m_bMixedSelection    = false; // Попадает ли в выделение одновременно несколько элементов
	this.m_nDrawing           = selected_None;
	this.m_pParagraph         = null;  // Параграф, в котором находится выделение
	this.m_oMath              = null;  // Формула, в которой находится выделение
	this.m_oHyperlink         = null;  // Гиперссылка, в которой находится выделение
	this.m_oField             = null;  // Поле, в котором находится выделение
	this.m_oCell              = null;  // Выделенная ячейка (специальная ситуация, когда выделена ровно одна ячейка)
	this.m_oBlockLevelSdt     = null;  // Если мы находимся в классе CBlockLevelSdt
	this.m_oInlineLevelSdt    = null;  // Если мы находимся в классе CInlineLevelSdt (важно, что мы находимся внутри класса)
	this.m_bSdtOverDrawing    = false; // Последний контент контрол обернут вокруг Drawing
	this.m_arrSdts            = [];    // Список всех контейнеров, попавших в селект
	this.m_arrComplexFields   = [];
	this.m_oPageNum           = null;
	this.m_oPagesCount        = null;
	this.m_nReviewFlags       = 0x00; // 0 - common, 1 - add, 2 - remove, 3 - change properties
	this.m_oPresentationField = null;
	this.m_arrMoveMarks       = [];
	this.m_arrFootEndNoteRefs = [];
	this.m_bFixedFormShape    = false; // Специальная ситуация, когда выделена автофигура вокруг fixed-form, но при этом мы в m_oInlineLevelSdt мы возвращаем эту форму (хотя в ней не находимся по факту)
}
CSelectedElementsInfo.prototype.Reset = function()
{
    this.m_bSelection      = false;
    this.m_bTable          = false;
    this.m_bMixedSelection = false;
    this.m_nDrawing        = -1;
};
CSelectedElementsInfo.prototype.SetMath = function(Math)
{
    this.m_oMath = Math;
};
CSelectedElementsInfo.prototype.SetField = function(Field)
{
    this.m_oField = Field;
};
CSelectedElementsInfo.prototype.GetMath = function()
{
    return this.m_oMath;
};
CSelectedElementsInfo.prototype.GetField = function()
{
    return this.m_oField;
};
CSelectedElementsInfo.prototype.SetTable = function()
{
    this.m_bTable = true;
};
CSelectedElementsInfo.prototype.SetDrawing = function(nDrawing)
{
    this.m_nDrawing = nDrawing;

    this.m_bSdtOverDrawing = true;
};
CSelectedElementsInfo.prototype.IsDrawingObjSelected = function()
{
    return ( this.m_nDrawing === selected_DrawingObject ? true : false );
};
CSelectedElementsInfo.prototype.SetMixedSelection = function()
{
    this.m_bMixedSelection = true;
};
CSelectedElementsInfo.prototype.IsInTable = function()
{
    return this.m_bTable;
};
CSelectedElementsInfo.prototype.IsMixedSelection = function()
{
    return this.m_bMixedSelection;
};
CSelectedElementsInfo.prototype.SetSingleCell = function(Cell)
{
    this.m_oCell = Cell;
};
CSelectedElementsInfo.prototype.GetSingleCell = function()
{
    return this.m_oCell;
};
CSelectedElementsInfo.prototype.IsSkipTOC = function()
{
	return this.m_bSkipTOC;
};
CSelectedElementsInfo.prototype.SetParagraph = function(Para)
{
	this.m_pParagraph = Para;
};
CSelectedElementsInfo.prototype.GetParagraph = function()
{
	return this.m_pParagraph;
};
CSelectedElementsInfo.prototype.SetBlockLevelSdt = function(oSdt)
{
	this.m_oBlockLevelSdt = oSdt;
	this.m_arrSdts.push(oSdt);
	this.m_bSdtOverDrawing = false;
};
/**
 * @returns {?CBlockLevelSdt}
 */
CSelectedElementsInfo.prototype.GetBlockLevelSdt = function()
{
	return this.m_oBlockLevelSdt;
};
CSelectedElementsInfo.prototype.SetInlineLevelSdt = function(oSdt)
{
	this.m_oInlineLevelSdt = oSdt;
	this.m_arrSdts.push(oSdt);
	this.m_bSdtOverDrawing = false;
};
/**
 * @returns {?CInlineLevelSdt}
 */
CSelectedElementsInfo.prototype.GetInlineLevelSdt = function()
{
	return this.m_oInlineLevelSdt;
};
/**
 * @returns {?CInlineLevelSdt}
 */
CSelectedElementsInfo.prototype.GetForm = function()
{
	if (this.m_oInlineLevelSdt && this.m_oInlineLevelSdt.IsForm())
		return this.m_oInlineLevelSdt;

	return null;
};
CSelectedElementsInfo.prototype.IsSdtOverDrawing = function()
{
	return this.m_bSdtOverDrawing;
};
CSelectedElementsInfo.prototype.GetAllSdts = function()
{
	return this.m_arrSdts;
};
CSelectedElementsInfo.prototype.CanDeleteBlockSdts = function()
{
	for (var nIndex = 0, nCount = this.m_arrSdts.length; nIndex < nCount; ++nIndex)
	{
		if (this.m_arrSdts[nIndex].IsBlockLevel() && !this.m_arrSdts[nIndex].CanBeDeleted() && this.m_arrSdts[nIndex].IsSelectedAll())
			return false;
	}

	return true;
};
CSelectedElementsInfo.prototype.CanEditBlockSdts = function()
{
	for (var nIndex = 0, nCount = this.m_arrSdts.length; nIndex < nCount; ++nIndex)
	{
		let contentControl = this.m_arrSdts[nIndex];
		contentControl.SkipSpecialContentControlLock(true);
		contentControl.SkipFillingFormModeCheck(true);
		
		let canEdit = (!contentControl.IsBlockLevel() || contentControl.CanBeEdited());
		contentControl.SkipSpecialContentControlLock(false);
		contentControl.SkipFillingFormModeCheck(false);
		
		if (!canEdit)
			return false;
	}

	return true;
};
CSelectedElementsInfo.prototype.CanDeleteInlineSdts = function()
{
	for (var nIndex = 0, nCount = this.m_arrSdts.length; nIndex < nCount; ++nIndex)
	{
		if (this.m_arrSdts[nIndex].IsInlineLevel() && !this.m_arrSdts[nIndex].CanBeDeleted() && this.m_arrSdts[nIndex].IsSelectedAll())
			return false;
	}

	return true;
};
CSelectedElementsInfo.prototype.CanEditInlineSdts = function()
{
	for (var nIndex = 0, nCount = this.m_arrSdts.length; nIndex < nCount; ++nIndex)
	{
		let contentControl = this.m_arrSdts[nIndex];
		contentControl.SkipSpecialContentControlLock(true);
		contentControl.SkipFillingFormModeCheck(true);

		let canEdit = (!contentControl.IsInlineLevel() || contentControl.CanBeEdited());
		contentControl.SkipFillingFormModeCheck(false);
		contentControl.SkipSpecialContentControlLock(false);
		
		if (!canEdit)
			return false;
	}

	return true;
};
CSelectedElementsInfo.prototype.SetComplexFields = function(arrComplexFields)
{
    if(this.m_bAddAllComplexFields)
    {
        var nAddIndex, nIndex;
        for(nAddIndex = 0; nAddIndex < arrComplexFields.length; ++nAddIndex)
        {
            var oAddField = arrComplexFields[nAddIndex];
            for(nIndex = 0; nIndex < this.m_arrComplexFields.length; ++nIndex)
            {
                if(this.m_arrComplexFields[nIndex] === oAddField)
                {
                    break;
                }
            }
            if(nIndex === this.m_arrComplexFields.length)
            {
                this.m_arrComplexFields.push(oAddField);
            }
        }
    }
    else
    {
        this.m_arrComplexFields = arrComplexFields;
    }
};
CSelectedElementsInfo.prototype.GetComplexFields = function()
{
	return this.m_arrComplexFields;
};
CSelectedElementsInfo.prototype.GetAllTablesOfFigures = function()
{
    var aTOF = [];
	for (var nIndex = this.m_arrComplexFields.length - 1; nIndex >= 0; --nIndex)
	{
		var oComplexField = this.m_arrComplexFields[nIndex];
		var oInstruction  = oComplexField.GetInstruction();

		if (AscCommonWord.fieldtype_TOC === oInstruction.GetType() && oInstruction.IsTableOfFigures())//TOC field can be table of contents or table or figures depending on flags
        {
            aTOF.push(oComplexField);
        }
	}
	return aTOF;
};
CSelectedElementsInfo.prototype.GetTableOfContents = function()
{
    for (var nIndex = this.m_arrComplexFields.length - 1; nIndex >= 0; --nIndex)
    {
        var oComplexField = this.m_arrComplexFields[nIndex];
        var oInstruction  = oComplexField.GetInstruction();

        if (AscCommonWord.fieldtype_TOC === oInstruction.GetType() && oInstruction.IsTableOfContents())//TOC field can be table of contents or table or figures depending on flags
        {
            return oComplexField;
        }
    }

    if (this.m_oBlockLevelSdt && this.m_oBlockLevelSdt.IsBuiltInTableOfContents())
        return this.m_oBlockLevelSdt;

    return null;
};
CSelectedElementsInfo.prototype.SetHyperlink = function(oHyperlink)
{
	this.m_oHyperlink = oHyperlink;
};
CSelectedElementsInfo.prototype.GetHyperlink = function()
{
	if (this.m_oHyperlink)
		return this.m_oHyperlink;

	for (var nIndex = 0, nCount = this.m_arrComplexFields.length; nIndex < nCount; ++nIndex)
	{
		var oInstruction = this.m_arrComplexFields[nIndex].GetInstruction();
		if (oInstruction && 
			(fieldtype_HYPERLINK === oInstruction.GetType() 
			|| fieldtype_REF === oInstruction.GetType() && oInstruction.GetHyperlink()
			|| fieldtype_NOTEREF === oInstruction.GetType() && oInstruction.GetHyperlink()))
			return oInstruction;
	}

	return null;
};
CSelectedElementsInfo.prototype.SetPageNum = function(oElement)
{
	this.m_oPageNum = oElement;
};
CSelectedElementsInfo.prototype.GetPageNum = function()
{
	return this.m_oPageNum;
};
CSelectedElementsInfo.prototype.SetPagesCount = function(oElement)
{
	this.m_oPagesCount = oElement;
};
CSelectedElementsInfo.prototype.GetPagesCount = function()
{
	return this.m_oPagesCount;
};
CSelectedElementsInfo.prototype.IsCheckAllSelection = function()
{
	return this.m_bCheckAllSelection;
};
CSelectedElementsInfo.prototype.RegisterReviewType = function(nReviewType)
{
	switch (nReviewType)
	{
		case reviewtype_Common:
		{
			this.m_nReviewFlags |= 1;
			break;
		}
		case reviewtype_Add:
		{
			this.m_nReviewFlags |= 2;
			break;
		}
		case reviewtype_Remove:
		{
			this.m_nReviewFlags |= 4;
			break;
		}
	}
};
CSelectedElementsInfo.prototype.HaveAddedInReview = function()
{
	return !!(this.m_nReviewFlags & 2);
};
CSelectedElementsInfo.prototype.HaveRemovedInReview = function()
{
	return !!(this.m_nReviewFlags & 4);
};
CSelectedElementsInfo.prototype.HaveNotReviewedContent = function()
{
	return !!(this.m_nReviewFlags & 1);
};
CSelectedElementsInfo.prototype.SetPresentationField = function(oField)
{
	this.m_oPresentationField = oField;
};
CSelectedElementsInfo.prototype.GetPresentationField = function()
{
	return this.m_oPresentationField;
};
CSelectedElementsInfo.prototype.RegisterTrackMoveMark = function(oMoveMark)
{
	this.m_arrMoveMarks.push(oMoveMark);
};
CSelectedElementsInfo.prototype.GetTrackMoveMarks = function()
{
	return this.m_arrMoveMarks;
};
CSelectedElementsInfo.prototype.RegisterFootEndNoteRef = function(oRef)
{
	this.m_arrFootEndNoteRefs.push(oRef);
};
CSelectedElementsInfo.prototype.GetFootEndNoteRefs  = function()
{
	return this.m_arrFootEndNoteRefs;
};
CSelectedElementsInfo.prototype.SetFixedFormShape = function(isFixedFormShape)
{
	this.FixedFormShape = isFixedFormShape;
};
CSelectedElementsInfo.prototype.IsFixedFormShape = function()
{
	return this.FixedFormShape;
};

/**
 * Основной класс для работы с документом в Word.
 * @param DrawingDocument
 * @param isMainLogicDocument
 * @constructor
 * @extends {CDocumentContentBase}
 */
function CDocument(DrawingDocument, isMainLogicDocument)
{
	CDocumentContentBase.call(this);

    //------------------------------------------------------------------------------------------------------------------
    //  Сохраняем ссылки на глобальные объекты
    //------------------------------------------------------------------------------------------------------------------
    this.History              = History;
    this.IdCounter            = AscCommon.g_oIdCounter;
    this.TableId              = AscCommon.g_oTableId;
    this.CollaborativeEditing = (("undefined" !== typeof(AscCommon.CWordCollaborativeEditing) && AscCommon.CollaborativeEditing instanceof AscCommon.CWordCollaborativeEditing) ? AscCommon.CollaborativeEditing : null);
    this.Api                  = editor;
    //------------------------------------------------------------------------------------------------------------------
    //  Выставляем ссылки на главный класс
    //------------------------------------------------------------------------------------------------------------------
    if (false !== isMainLogicDocument)
    {
        if (this.History)
            this.History.Set_LogicDocument(this);

        if (this.CollaborativeEditing)
            this.CollaborativeEditing.m_oLogicDocument = this;

        if (DrawingDocument)
        	DrawingDocument.m_oLogicDocument = this;
    }
	this.MainDocument = false !== isMainLogicDocument;
    //__________________________________________________________________________________________________________________

    this.Id = this.IdCounter.Get_NewId();
	//Props
	this.App = null;
	this.Core = null;
    this.CustomProperties = null;
    this.CustomXmls = [];

    // Сначала настраиваем размеры страницы и поля
    this.SectPr = new CSectionPr(this);
    this.SectionsInfo = new CDocumentSectionsInfo();

	// Режим рецензирования
	this.TrackRevisions = null; // Локальный флаг рецензирования, который перекрывает флаг Settings.TrackRevisions, если сам не null
	this.TrackRevisionsManager = new AscWord.CTrackRevisionsManager(this);

	this.Settings = new AscWord.CDocumentSettings();

	this.Layouts = {
		Print : new AscWord.CDocumentPrintView(this),
		Read  : new AscWord.CDocumentReadView(this)
	};

	this.Layout = this.Layouts.Print;

	this.Content[0] = new Paragraph(DrawingDocument, this);
    this.Content[0].Set_DocumentNext(null);
    this.Content[0].Set_DocumentPrev(null);

    this.CurPos  =
    {
        X          : 0,
        Y          : 0,
        ContentPos : 0, // в зависимости, от параметра Type: позиция в Document.Content
        RealX      : 0, // позиция курсора, без учета расположения букв
        RealY      : 0, // это актуально для клавиш вверх и вниз
        Type       : docpostype_Content,
		CC         : null,
		MainCC     : null,

		SetCC : function(oCC)
		{
			this.CC     = oCC;
			this.MainCC = oCC && oCC.IsForm() ? oCC.GetMainForm() : oCC;
		},
		ResetCC : function()
		{
			this.CC     = null;
			this.MainCC = null;
		},
		IsInCC : function()
		{
			return !!(this.CC);
		},
		CheckHitInCC : function(nX, nY, nPageAbs)
		{
			if (!this.MainCC)
				return true;

			return this.MainCC.CheckHitInContentControlByXY(nX, nY, nPageAbs);
		},
		CorrectXYToHitInCC : function(nX, nY, nPageAbs)
		{
			if (!this.MainCC)
				return null;

			return this.MainCC.CorrectXYToHitIn(nX, nY, nPageAbs)
		}
    };

	// TODO: Пока временно так сделаем, в будущем надо переделать в общий класс позиции документа
	this.FocusCC = null;

	this.MathTrackHandler = new AscWord.CMathTrackHandler(DrawingDocument, this.Api);

	this.Selection =
    {
        Start    : false,
        Use      : false,
        StartPos : 0,
        EndPos   : 0,
        Flag     : selectionflag_Common,
        Data     : null,
        UpdateOnRecalc : false,
		WordSelected : false,
        DragDrop : { Flag : 0, Data : null }  // 0 - не drag-n-drop, и мы его проверяем, 1 - drag-n-drop, -1 - не проверять drag-n-drop
    };

	this.Action = {
		Start           : false,
		Depth           : 0,
		PointsCount     : 0,
		Description     : AscDFH.historyitem_type_Unknown,
		Recalculate     : false,
		CancelAction    : false,
		UpdateSelection : false,
		UpdateInterface : false,
		UpdateRulers    : false,
		UpdateUndoRedo  : false,
		UpdateStates    : false,
		Redraw          : {
			Start : undefined,
			End   : undefined
		},
		UndoRedo        : false, // В текущий момент идет действие Undo/Redo

		Additional : {}
	};

    if (false !== isMainLogicDocument)
        this.ColumnsMarkup = new CColumnsMarkup();

    // Здесь мы храним инфрмацию, связанную с разбивкой на страницы и самими страницами
    this.Pages = [];

    this.RecalcInfo = new CDocumentRecalcInfo();

    this.RecalcId     = 0; // Номер пересчета
    this.FullRecalc   = new CDocumentRecalculateState(); // Объект полного пересчета
    this.HdrFtrRecalc = new CDocumentRecalculateHdrFtrPageCountState(); // Объект дополнительного пересчета колонтитулов после полного пересчета

    this.TurnOffRecalc          = 0;
    this.TurnOffInterfaceEvents = false;
    this.TurnOffRecalcCurPos    = false;

    this.CheckEmptyElementsOnSelection = true; // При выделении проверять или нет пустой параграф в конце/начале выделения.
	
	this.Numbering           = new AscWord.CNumbering();               // Форматный класс для хранения всех нумераций согласно формату
	this.NumberingApplicator = new AscWord.CNumberingApplicator(this); // Класс для применения нумерации к текущему выделение
	this.NumberingCollection = new AscWord.CNumberingCollection(this); // Класс, хранящий нумерации, используемые в документе
	
    this.Styles    = new CStyles();
    this.Styles.Set_LogicDocument(this);

    this.DrawingDocument = DrawingDocument;

    this.NeedUpdateTarget = false;
	this.ViewPosition     = null;  // Позиция, куда мы должны проскроллиться после пересчета (если задана)

    // Флаг, который контролирует нужно ли обновлять наш курсор у остальных редакторов нашего документа.
    // Также следим за частотой обновления, чтобы оно проходило не чаще одного раза в секунду.
    this.NeedUpdateTargetForCollaboration = true;
    this.LastUpdateTargetTime             = 0;

    this.ReindexStartPos = 0;

    // Класс для работы с колонтитулами
    this.HdrFtr = new CHeaderFooterController(this, this.DrawingDocument);

    // Класс для работы с поиском
    this.SearchInfo =
    {
        Id       : null,
        StartPos : 0,
        CurPage  : 0,
        String   : null
    };

    // Позция каретки
    this.TargetPos =
    {
        X       : 0,
        Y       : 0,
        PageNum : 0
    };


    // Класс для работы со статискикой документа
    this.Statistics = new CStatistics( this );

    this.HighlightColor = null;

	if (typeof AscCommon.CComments !== "undefined")
		this.Comments = new AscCommon.CComments(this);

    this.Lock = new AscCommon.CLock();

    this.m_oContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)

    // Массив укзателей на все инлайновые графические объекты
    this.DrawingObjects = null;

    if (typeof CGraphicObjects !== "undefined")
        this.DrawingObjects = new CGraphicObjects(this, this.DrawingDocument, this.Api);

    this.theme          = AscFormat.GenerateDefaultTheme(this);
    this.clrSchemeMap   = AscFormat.GenerateDefaultColorMap();

    // Класс для работы с поиском и заменой в документе
    this.SearchEngine = null;
    if (typeof AscCommonWord.CDocumentSearch !== "undefined")
        this.SearchEngine = new AscCommonWord.CDocumentSearch(this);

    // Параграфы, в которых есть ошибки в орфографии (объект с ключом - Id параграфа)
    this.Spelling = new AscCommonWord.CDocumentSpellChecker();

    // Дополнительные настройки
	this.ForceHideCCTrack          = false; // Насильно запрещаем отрисовку рамок у ContentControl
    this.UseTextShd                = true;  // Использовать ли заливку текста
    this.ForceCopySectPr           = false; // Копировать ли настройки секции, если родительский класс параграфа не документ
    this.CopyNumberingMap          = null;  // Мап старый индекс -> новый индекс для копирования нумерации
    this.CheckLanguageOnTextAdd    = false; // Проверять ли язык при добавлении текста в ран
	this.RemoveCommentsOnPreDelete = true;  // Удалять ли комментарий при удалении объекта
	this.CheckInlineSdtOnDelete    = null;  // Проверяем заданный InlineSdt на удалении символов внутри него
	this.DragAndDropAction         = false; // Происходит ли сейчас действие drag-n-drop
	this.RecalcTableHeader         = false; // Пересчитываем ли сейчас заголовок таблицы
	this.TrackMoveId               = null;  // Идентификатор переноса внутри рецензирования
	this.TrackMoveRelocation       = false; // Перенос ранее перенесенного текста внутри рецензирования
	this.RemoveEmptySelection      = true;  // При обновлении селекта, если он пустой тогда сбрасываем его
	this.MoveDrawing               = false; // Происходит ли сейчас перенос автофигуры
	this.PrintSelection            = false; // Печатаем выделенный фрагмент
	this.CheckFormPlaceHolder      = true;  // Выполняем ли специальную обработку для плейсхолдеров у форм
	this.MathInputType             = Asc.c_oAscMathInputType.Unicode;
	this.ForceDrawPlaceHolders     = null;  // true/false - насильно заставляем рисовать или не рисовать плейсхолдеры и подсветку,
	this.ForceDrawFormHighlight    = null;  // null - редактор решает рисовать или нет в зависимости от других параметров
	this.ConcatParagraphsOnRemove  = false; // Во время удаления объединять ли первый и последний параграфы
	this.StartCheckTextFormFormat  = false; // Флаг, что в данный момент мы уже проверяем формат текстовых форм, чтобы не вызывать повторно
	this.StartCheckCCPlaceholder   = false; //
	this.CompileStyleOnLoad        = false; // Компилировать ли принудительно стили во время загрузки
	this.SmartParagraphSelection   = true;  // Выделять ли автоматически знак параграфа, когда все содержимое параграфа выделено


	this.DrawTableMode = {
		Start  : false,
		Draw   : false,
		Erase  : false,
		StartX : -1,
		StartY : -1,
		EndX   : -1,
		EndY   : -1,
		Page   : -1,

		TablesOnPage   : [],
		Table          : null,
		TablePageStart : -1,
		TablePageEnd   : -1,

		UpdateTablePages : function()
		{
			if (this.Table)
			{
				var nPageStart = this.Page;
				var nPageEnd   = this.Page;

				if (!(this.Table.Parent instanceof CDocument))
				{
					nPageStart = this.Table.Parent.GetPageIndexByXYAndPageAbs(this.StartX, this.StartY, this.Page);
					nPageEnd   = this.Table.Parent.GetPageIndexByXYAndPageAbs(this.EndX, this.EndY, this.Page);
				}

				this.TablePageStart = this.Table.Parent.private_GetElementPageIndexByXY(this.Table.GetIndex(), this.StartX, this.StartY, nPageStart);
				this.TablePageEnd   = this.Table.Parent.private_GetElementPageIndexByXY(this.Table.GetIndex(), this.EndX, this.EndY, nPageEnd);
			}
			else
			{
				this.TablePageStart = -1;
				this.TablePageEnd   = -1;
			}
		},

		CheckSelectedTable : function()
		{
			var nStartX = this.StartX;
			var nEndX   = this.EndX;

			if (nEndX < nStartX)
			{
				nStartX = this.EndX;
				nEndX   = this.StartX;
			}

			var nStartY = this.StartY;
			var nEndY   = this.EndY;

			if (nEndY < nStartY)
			{
				nStartY = this.EndY;
				nEndY   = this.StartY;
			}

			if (this.Erase)
			{
				this.Table = null;

				var arrTables = this.TablesOnPage;
				for (var nTableIndex = 0, nTablesCount = arrTables.length; nTableIndex < nTablesCount; ++nTableIndex)
				{
					var oBounds = arrTables[nTableIndex].Table.GetPageBounds(arrTables[nTableIndex].Page);

					if (((nStartX < oBounds.Left && oBounds.Left < nEndX) || (nStartX < oBounds.Right && oBounds.Right < nEndX))
						&& ((nStartY < oBounds.Top && oBounds.Top < nEndY) || (nStartY < oBounds.Bottom && oBounds.Bottom < nEndY)))
					{
						this.Table = arrTables[nTableIndex].Table;
						return;
					}

					if ((((oBounds.Left < nStartX && nStartX < oBounds.Right) || (oBounds.Left < nEndX && nEndX < oBounds.Right))
						&& ((oBounds.Top < nStartY && nStartY < oBounds.Bottom) || (oBounds.Top < nEndY && nEndY < oBounds.Bottom)))
						|| (oBounds.Left < nStartX && nEndX < oBounds.Right && nStartY < oBounds.Top && oBounds.Bottom < nEndY)
						|| (nStartX < oBounds.Left && oBounds.Right < nEndX && oBounds.Top < nStartY && nEndY < oBounds.Bottom))
					{
						this.Table = arrTables[nTableIndex].Table;
					}
				}
			}
		}
	};

	// Параметры для случая, когда мы не можем сразу перерисовать треки и нужно перерисовывать их на таймере пересчета
	this.NeedUpdateTracksOnRecalc = false;
	this.NeedUpdateTracksParams   = {
		Selection      : false,
		EmptySelection : false
	};

    // Мап для рассылки
    this.MailMergeMap             = null;
    this.MailMergePreview         = false;
    this.MailMergeFieldsHighlight = false; // Подсвечивать ли все поля связанные с рассылкой
    this.MailMergeFields          = [];

    // Класс, управляющий полями
    this.FieldsManager = new CDocumentFieldsManager();

    // Класс, управляющий закладками
	if (typeof CBookmarksManager !== "undefined")
		this.BookmarksManager = new CBookmarksManager(this);

    this.DocumentOutline = new CDocumentOutline(this);

    this.GlossaryDocument = new CGlossaryDocument(this);

	this.AutoCorrectSettings = new AscCommon.CAutoCorrectSettings();

    // Контролируем изменения интерфейса
    this.ChangedStyles      = []; // Объект с Id стилями, которые были изменены/удалены/добавлены
	this.TurnOffPanelStyles = 0;  // == 0 - можно обновлять панельку со стилями, != 0 - нельзя обновлять

    // Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
    this.TableId.Add(this, this.Id);

	// Объект для составного ввода текста
	this.CompositeInput = null;

	// Нужно ли проверять тип лока у ContentControl при проверке залоченности выделенных объектов
	this.CheckContentControlsLock = true;

	this.ViewModeInReview = {
		mode                : Asc.c_oAscDisplayModeInReview.Edit,
		isFastCollaboration : false,
	};

	this.OFormDocument           = AscCommon.IsSupportOFormFeature() && false !== isMainLogicDocument ? new AscOForm.OForm(this) : null;
	this.FormsManager            = new AscWord.CFormsManager(this);
	this.CurrentForm             = null;
	this.CurrentFormFixed        = false;
	this.HighlightRequiredFields = false;  // Выделяем ли обязательные поля
	this.RequiredFieldsBorder    = new CDocumentBorder();
	this.RequiredFieldsBorder.SetSimpleColor(255, 0, 0);

	this.SectionEndMark = {}; // Храним пересчитанные метки конца секции

	this.LineNumbersInfo = new CDocumentLineNumbersInfo(this); // Класс для кэширования информации для отрисовки нумерации строк

	// Класс для работы со сносками
	this.Footnotes               = new CFootnotesController(this);
	this.LogicDocumentController = new CLogicDocumentController(this);
	this.DrawingsController      = new CDrawingsController(this, this.DrawingObjects);
	this.HeaderFooterController  = new CHdrFtrController(this, this.HdrFtr);
	this.Endnotes                = new CEndnotesController(this);

	this.Controller = this.LogicDocumentController;

    // Кэш для некоторых запросов, которые могут за раз делаться несколько раз
    this.AllParagraphsList = null;
    this.AllFootnotesList  = null; // TODO: Переделать
    this.AllEndnotesList   = null;

	//------------------------------------------------------------------------------------------------------------------
	//  Check StartCollaborationEditing
	//------------------------------------------------------------------------------------------------------------------
	if (false !== isMainLogicDocument && this.CollaborativeEditing && !this.CollaborativeEditing.Is_SingleUser())
	{
		this.StartCollaborationEditing();
	}
	//__________________________________________________________________________________________________________________
}
CDocument.prototype = Object.create(CDocumentContentBase.prototype);
CDocument.prototype.constructor = CDocument;

CDocument.prototype.IsDocumentEditor = function()
{
	return true;
};
CDocument.prototype.IsPresentationEditor = function()
{
	return false;
};
CDocument.prototype.IsSpreadSheetEditor = function()
{
	return false;
};
CDocument.prototype.Init                           = function()
{

};
CDocument.prototype.IsLoadingDocument = function()
{
	return (AscCommon.g_oIdCounter.m_bLoad);
};
CDocument.prototype.On_EndLoad                     = function()
{
	this.Start_SilentMode();

	this.ClearListsCache();

	this.TrackRevisionsManager.InitSelectedChanges();

    // Обновляем информацию о секциях
    this.UpdateAllSectionsInfo();

	this.SectionsInfo.RemoveEmptyHdrFtrs();

    // Проверяем последний параграф на наличие секции
    this.Check_SectionLastParagraph();

    this.Styles.OnEndDocumentLoad(this);

    // Обновляем массив позиций для комментариев
    this.Comments.UpdateAll();

	// ВАЖНО: Вызываем данную функцию после обновления секций
	this.private_UpdateFieldsOnEndLoad();
	this.RemoveSelection();

    // Перемещаем курсор в начало документа
	this.SetDocPosType(docpostype_Content);
	this.MoveCursorToStartOfDocument();

    if (this.Api.DocInfo)
    {
        var TemplateReplacementData = this.Api.DocInfo.get_TemplateReplacement();
        if (null !== TemplateReplacementData)
        {
            this.private_ProcessTemplateReplacement(TemplateReplacementData);
        }
    }
    if (null !== this.CollaborativeEditing)
    {
        this.Set_FastCollaborativeEditing(true);
    }
	
	this.FormsManager.OnEndLoad();

	if (this.OFormDocument)
		this.OFormDocument.onEndLoad();

	this.End_SilentMode();
};
CDocument.prototype.UpdateDefaultsDependingOnCompatibility = function()
{
	this.Styles.UpdateDefaultsDependingOnCompatibility(this.GetCompatibilityMode());
};
CDocument.prototype.private_UpdateFieldsOnEndLoad = function()
{
	//let nTime = performance.now();

	let arrHdrFtrs = this.SectionsInfo.GetAllHdrFtrs();
	let arrFields  = [];
	for (let nIndex = 0, nCount = arrHdrFtrs.length; nIndex < nCount; ++nIndex)
	{
		let oContent = arrHdrFtrs[nIndex].GetContent();
		oContent.ProcessComplexFields();
		oContent.GetAllFields(false, arrFields);
	}

	// Мы пока не умеем обрабытывать атоматически обновляющиеся FldSimple поля в колонтитулах, поэтому
	// преобразуем их в состатовное поле
	for (let nFieldIndex = 0, nFieldsCount = arrFields.length; nFieldIndex < nFieldsCount; ++nFieldIndex)
	{
		let oField = arrFields[nFieldIndex];
		if (oField instanceof ParaField
			&& (fieldtype_PAGECOUNT === oField.GetFieldType() || fieldtype_PAGENUM === oField.GetFieldType()))
		{
			let oComplexField = oField.ReplaceWithComplexField();
			if (oComplexField)
				arrFields[nFieldIndex] = oComplexField;
		}
	}

	let openedAt = this.Api ? this.Api.openedAt : undefined;
	if (undefined === openedAt)
		return;
	
	// Для правильного обновления полей нам нужно, чтобы стили компилировались в обход флага загрузки
	this.CompileStyleOnLoad = true;

	this.ProcessComplexFields();
	this.controller_GetAllFields(false, arrFields);

	for (var nFieldIndex = 0, nFieldsCount = arrFields.length; nFieldIndex < nFieldsCount; ++nFieldIndex)
	{
		let oField = arrFields[nFieldIndex];
		if (oField instanceof CComplexField
			&& oField.GetInstruction()
			&& fieldtype_TIME === oField.GetInstruction().Type)
		{
			oField.UpdateTIME(openedAt);
		}
	}
	
	this.CompileStyleOnLoad = false;

	//console.log("FieldUpdateTime : " + ((performance.now() - nTime) / 1000) + "s");
};
CDocument.prototype.Add_TestDocument               = function()
{
    this.Content   = [];
    var Text       = ["Comparison view helps you track down memory leaks, by displaying which objects have been correctly cleaned up by the garbage collector. Generally used to record and compare two (or more) memory snapshots of before and after an operation. The idea is that inspecting the delta in freed memory and reference count lets you confirm the presence and cause of a memory leak.", "Containment view provides a better view of object structure, helping us analyse objects referenced in the global namespace (i.e. window) to find out what is keeping them around. It lets you analyse closures and dive into your objects at a low level.", "Dominators view helps confirm that no unexpected references to objects are still hanging around (i.e that they are well contained) and that deletion/garbage collection is actually working."];
    var ParasCount = 50;
    var RunsCount  = Text.length;
    for (var ParaIndex = 0; ParaIndex < ParasCount; ParaIndex++)
    {
        var Para = new Paragraph(this.DrawingDocument, this);
        //var Run = new ParaRun(Para);
        for (var RunIndex = 0; RunIndex < RunsCount; RunIndex++)
        {

            var String    = Text[RunIndex];
            var StringLen = String.length;
            for (var TextIndex = 0; TextIndex < StringLen; TextIndex++)
            {
                var Run         = new ParaRun(Para);
                var TextElement = String[TextIndex];

                Run.AddText(TextElement);
                Para.AddToContent(0, Run);
            }


        }
        //Para.Add_ToContent(0, Run);
        this.Internal_Content_Add(this.Content.length, Para);
    }

	this.RecalculateFromStart(true);
};
CDocument.prototype.LoadEmptyDocument              = function()
{
    this.DrawingDocument.TargetStart();
    this.Recalculate();

    this.UpdateInterfaceParaPr();
    this.UpdateInterfaceTextPr();
};
CDocument.prototype.Set_CurrentElement = function(Index, bUpdateStates)
{
	var OldDocPosType = this.CurPos.Type;

	var ContentPos = Math.max(0, Math.min(this.Content.length - 1, Index));

	this.SetDocPosType(docpostype_Content);
	this.CurPos.ContentPos = Math.max(0, Math.min(this.Content.length - 1, Index));

	this.ResetWordSelection();
	if (true === this.Content[ContentPos].IsSelectionUse())
	{
		this.Selection.Flag     = selectionflag_Common;
		this.Selection.Use      = true;
		this.Selection.StartPos = ContentPos;
		this.Selection.EndPos   = ContentPos;
	}
	else
		this.RemoveSelection();

	if (false != bUpdateStates)
	{
		this.Document_UpdateInterfaceState();
		this.Document_UpdateRulersState();
		this.Document_UpdateSelectionState();
	}

	this.UpdateTracks();

	if (docpostype_HdrFtr === OldDocPosType)
	{
		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();
	}
};
CDocument.prototype.IsThisElementCurrent          = function()
{
    return true;
};
CDocument.prototype.Get_PageContentStartPos = function(nPageAbs, nContentIndex)
{
	let oSectPr = this.Layout.GetSection(nPageAbs, nContentIndex);
	return this.Layout.GetPageContentFrame(nPageAbs, oSectPr);
};
CDocument.prototype.Get_PageContentStartPos2 = function(StartPageIndex, StartColumnIndex, ElementPageIndex, ElementIndex)
{
	let oSectPr = this.Layout.GetSection(StartPageIndex, ElementIndex);

	let ColumnsCount = oSectPr.GetColumnsCount();
	let ColumnAbs    = (StartColumnIndex + ElementPageIndex) - ((StartColumnIndex + ElementPageIndex) / ColumnsCount | 0) * ColumnsCount;
	let PageAbs      = StartPageIndex + ((StartColumnIndex + ElementPageIndex) / ColumnsCount | 0);

	return this.Layout.GetColumnContentFrame(PageAbs, ColumnAbs, oSectPr);
};
CDocument.prototype.Get_PageLimits = function(nPageIndex)
{
	var nIndex  = this.Pages[nPageIndex] ? this.Pages[nPageIndex].Pos : 0;
	var oSectPr = this.SectionsInfo.Get_SectPr(nIndex).SectPr;

	return {
		X      : 0,
		Y      : 0,
		XLimit : oSectPr.GetPageWidth(),
		YLimit : oSectPr.GetPageHeight()
	};
};
CDocument.prototype.Get_PageFields = function(nPageIndex, isHdrFtr, oSectPr)
{
	let oPage = this.Pages[nPageIndex];

	if (!oSectPr)
	{
		let nIndex = oPage ? oPage.Pos : 0;
		oSectPr    = this.SectionsInfo.Get_SectPr(nIndex).SectPr;
	}

	var oFrame  = oSectPr.GetContentFrame(nPageIndex);
	if (!oPage)
	{
		return {
			X      : oFrame.Left,
			Y      : oFrame.Top,
			XLimit : oFrame.Right,
			YLimit : oFrame.Bottom
		};
	}

	var nTop    = oFrame.Top;
	var nBottom = oFrame.Bottom;

	if (!isHdrFtr)
	{
		var oHdrFtrLine = this.HdrFtr.GetHdrFtrLines(nPageIndex);

		if (null !== oHdrFtrLine.Top && oHdrFtrLine.Top > nTop)
			nTop = oHdrFtrLine.Top;

		if (null !== oHdrFtrLine.Bottom && oHdrFtrLine.Bottom < nBottom)
			nBottom = oHdrFtrLine.Bottom;
	}

	return {
		X      : oFrame.Left,
		Y      : nTop,
		XLimit : oFrame.Right,
		YLimit : nBottom
	};
};
CDocument.prototype.Get_ColumnFields = function(nElementIndex, nColumnIndex, nPageIndex)
{
	if (undefined === nElementIndex)
	{
		if (!this.Page[nPageIndex])
		{
			return {
				X      : 0,
				XLimit : 0
			};
		}

		nPageIndex = this.Pages[nPageIndex].Pos;
	}

	var oSectPr = this.SectionsInfo.Get_SectPr(nElementIndex).SectPr;
	var oFrame  = oSectPr.GetContentFrame(nPageIndex);

	var X      = oFrame.Left;
	var XLimit = oFrame.Right;

	var ColumnsCount = oSectPr.GetColumnsCount();
	if (nColumnIndex >= ColumnsCount)
		nColumnIndex = ColumnsCount - 1;

	for (var ColIndex = 0; ColIndex < nColumnIndex; ++ColIndex)
	{
		X += oSectPr.GetColumnWidth(ColIndex);
		X += oSectPr.GetColumnSpace(ColIndex);
	}

	if (ColumnsCount - 1 !== nColumnIndex)
		XLimit = X + oSectPr.GetColumnWidth(nColumnIndex);

	return {
		X      : X,
		XLimit : XLimit
	};
};
CDocument.prototype.Get_Theme                      = function()
{
    return this.theme;
};
CDocument.prototype.GetTheme = function()
{
	return this.Get_Theme();
};
CDocument.prototype.Get_ColorMap                   = function()
{
    return this.clrSchemeMap;
};
CDocument.prototype.GetColorMap = function()
{
	return this.Get_ColorMap();
};
/**
 * Начинаем новое действие, связанное с изменением документа
 * @param {number} nDescription
 * @param {object} [oSelectionState=null] - начальное состояние селекта, до начала действия
 */
CDocument.prototype.StartAction = function(nDescription, oSelectionState)
{
	var isNewPoint = this.History.Create_NewPoint(nDescription, oSelectionState);

	if (true === this.Action.Start)
	{
		this.Action.Depth++;

		if (isNewPoint)
			this.Action.PointsCount++;
	}
	else
	{
		this.Action.Start           = true;
		this.Action.Depth           = 0;
		this.Action.PointsCount     = isNewPoint ? 1 : 0;
		this.Action.Recalculate     = false;
		this.Action.Description     = nDescription;
		this.Action.UpdateSelection = false;
		this.Action.UpdateInterface = false;
		this.Action.UpdateRulers    = false;
		this.Action.UpdateUndoRedo  = false;
		this.Action.UpdateTracks    = false;
		this.Action.Redraw.Start    = undefined;
		this.Action.Redraw.End      = undefined;
		this.Action.Additional      = {};
	}
};
/**
 * В процессе ли какое-либо действие
 * @returns {boolean}
 */
CDocument.prototype.IsActionStarted = function()
{
	return this.Action.Start;
};
/**
 * Сообщаем документу, что потребуется пересчет
 * @param {boolean} [isForceRecalculate = false] Запускать ли пересчет прямо сейчас
 */
CDocument.prototype.Recalculate = function(isForceRecalculate)
{
	if (this.Action.Start && true !== isForceRecalculate)
		this.Action.Recalculate = true;
	else
		this.private_Recalculate();
};
/**
 * Сообщаем документу, что потребуется обновить состояние селекта
 * @param {boolean} [isRemoveEmptySelection = true]
 */
CDocument.prototype.UpdateSelection = function(isRemoveEmptySelection)
{
	if (false === isRemoveEmptySelection)
	{
		this.RemoveEmptySelection = false;
		this.private_UpdateSelection();
		this.RemoveEmptySelection = true;
	}
	else if (this.Action.Start)
	{
		this.Action.UpdateSelection = true;
	}
	else
	{
		this.private_UpdateSelection();
	}
};
CDocument.prototype.UpdatePlaceholders = function ()
{
	if (this.Action.Start)
	{
		this.Action.UpdatePlaceholders = true;
    }
	else
	{
		this.private_UpdatePlaceholders();
	}
};
/**
 * Сообщаем документу, что потребуется обновить состояние интерфейса
 * @param {?boolean} bSaveCurRevisionChange
 * @param {boolean} [isExternalTrigger=false] - обновление интерфейса вызвано не действиями данного пользователя (например, совместкой)
 */
CDocument.prototype.UpdateInterface = function(bSaveCurRevisionChange, isExternalTrigger)
{
	if (undefined !== bSaveCurRevisionChange || true === isExternalTrigger)
		this.private_UpdateInterface(bSaveCurRevisionChange, isExternalTrigger);
	else if (this.Action.Start)
		this.Action.UpdateInterface = true;
	else
		this.private_UpdateInterface();
};
/**
 * Сообщаем документу, что потребуется обновить линейки
 */
CDocument.prototype.UpdateRulers = function()
{
	if (this.Action.Start)
		this.Action.UpdateRulers = true;
	else
		this.private_UpdateRulers();
};
/**
 * Сообщаем документу, что потребуется обновить состояние кнопки Undo/Redo
 */
CDocument.prototype.UpdateUndoRedo = function()
{
	if (this.Action.Start)
		this.Action.UpdateUndoRedo = true;
	else
		this.private_UpdateUndoRedo();
};
/**
 * Сообщаем документу, что нужно обновить треки
 */
CDocument.prototype.UpdateTracks = function()
{
	if (this.Action.Start)
		this.Action.UpdateTracks = true;
	else
		this.private_UpdateDocumentTracks();
};
/**
 * Перерисовываем заданные страницы
 * @param {number} nStartPage
 * @param {number} nEndPage
 */
CDocument.prototype.Redraw = function(nStartPage, nEndPage)
{
	if (this.Action.Start)
	{
		if (undefined !== nStartPage && undefined !== nEndPage && -1 !== this.Action.Redraw.Start)
		{
			if (undefined === this.Action.Redraw.Start || nStartPage < this.Action.Redraw.Start)
				this.Action.Redraw.Start = nStartPage;

			if (undefined === this.Action.Redraw.End || nEndPage > this.Action.Redraw.End)
				this.Action.Redraw.End = nEndPage;
		}
		else
		{
			this.Action.Redraw.Start = -1;
			this.Action.Redraw.End   = -1;
		}
	}
	else
	{
		this.private_Redraw(nStartPage, nEndPage);
	}
};
CDocument.prototype.private_Redraw = function(nStartPage, nEndPage)
{
	if (-1 !== nStartPage && -1 !== nEndPage)
	{
		var _nStartPage = Math.max(0, nStartPage);
		var _nEndPage   = Math.min(this.DrawingDocument.m_lCountCalculatePages - 1, nEndPage);

		for (var nCurPage = _nStartPage; nCurPage <= _nEndPage; ++nCurPage)
			this.DrawingDocument.OnRepaintPage(nCurPage);
	}
	else
	{
		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();
	}
};
/**
 * Завершаем действие
 * @param {boolean} [isCheckEmptyAction=true] Нужно ли проверять, что действие ничего не делало
 */
CDocument.prototype.FinalizeAction = function(isCheckEmptyAction)
{
	if (!this.IsActionStarted())
		return;

	if (this.Action.Depth > 0)
	{
		// Если на каждое действие создавалась точка, но последнее действие оказалось пустым
		if (this.Action.Depth + 1 === this.Action.PointsCount && this.History.Is_LastPointEmpty())
		{
			this.History.Remove_LastPoint();
			this.Action.PointsCount--;
		}

		this.Action.Depth--;
		return;
	}

	this.private_CheckAdditionalOnFinalize();

	var isAllPointsEmpty = true;
	if (this.Action.CancelAction)
	{
		let arrChanges = [];
		for (var nIndex = 0, nPointsCount = this.Action.PointsCount; nIndex < nPointsCount; ++nIndex)
		{
			arrChanges = arrChanges.concat(this.History.Undo());
		}

		this.RecalculateByChanges(arrChanges);
	}
	else if (false !== isCheckEmptyAction)
	{
		for (var nIndex = 0, nPointsCount = this.Action.PointsCount; nIndex < nPointsCount; ++nIndex)
		{
			if (this.History.Is_LastPointEmpty())
			{
				this.History.Remove_LastPoint();
			}
			else
			{
				isAllPointsEmpty = false;
				break;
			}
		}
	}
	else
	{
		isAllPointsEmpty = false;
	}

	if (!isAllPointsEmpty)
	{
		if (this.Action.Recalculate)
		{
			var nRecalcResult = this.private_Recalculate();

			if (this.Action.Additional.FormAutoFit)
				this.private_FinalizeFormAutoFit(nRecalcResult & document_recalcresult_FastFlag);

		}
		else if (undefined !== this.Action.Redraw.Start && undefined !== this.Action.Redraw.End)
		{
			this.private_Redraw(this.Action.Redraw.Start, this.Action.Redraw.End);
		}
	}

	if (this.IsFillingFormMode())
		this.Api.sync_OnAllRequiredFormsFilled(this.FormsManager.IsAllRequiredFormsFilled());

	var oCurrentParagraph = this.GetCurrentParagraph(false, false)
	if (oCurrentParagraph && oCurrentParagraph.IsInFixedForm())
		oCurrentParagraph.GetParent().CheckFormViewWindow();

	this.Action.Start              = false;
	this.Action.Depth              = 0;
	this.Action.PointsCount        = 0;
	this.Action.Recalculate        = false;
	this.Action.CancelAction       = false;
	this.Action.Redraw.Start       = undefined;
	this.Action.Redraw.End         = undefined;
	this.Action.Additional         = {};
	this.Api.checkChangesSize();
	
	if (this.Action.UpdateStates)
		return;
	
	this.Action.UpdateStates = true;
	
	if (this.Action.UpdateInterface)
		this.private_UpdateInterface();
	
	if (this.Action.UpdateSelection)
		this.private_UpdateSelection();
	
	if (this.Action.UpdateRulers)
		this.private_UpdateRulers();
	
	if (this.Action.UpdateUndoRedo)
		this.private_UpdateUndoRedo();
	
	if (this.Action.UpdateTracks)
		this.private_UpdateDocumentTracks();
	
	if (this.Action.UpdatePlaceholders)
		this.private_UpdatePlaceholders();
	
	this.Action.UpdateSelection    = false;
	this.Action.UpdateInterface    = false;
	this.Action.UpdateRulers       = false;
	this.Action.UpdateUndoRedo     = false;
	this.Action.UpdateTracks       = false;
	this.Action.UpdatePlaceholders = false;
	
	this.Action.UpdateStates = false;
};
/**
 * Сообщаем, что нужно отменить начатое действие
 */
CDocument.prototype.CancelAction = function()
{
	if (!this.IsActionStarted())
		return;

	this.Action.CancelAction = true;
};
CDocument.prototype.StartUndoRedoAction = function()
{
	this.Action.UndoRedo = true;
	
	this.DrawingDocument.EndTrackTable(null, true);
	this.DrawingObjects.TurnOffCheckChartSelection();
	this.BookmarksManager.SetNeedUpdate(true);
};
CDocument.prototype.FinalizeUndoRedoAction = function()
{
	if (!this.Action.UndoRedo)
		return;
	
	if (this.Action.Additional.ContentControlChange)
		this.private_FinalizeContentControlChange();
	
	this.Action.UndoRedo   = false;
	this.Action.Additional = {};
};
CDocument.prototype.private_CheckAdditionalOnFinalize = function()
{
	this.Action.Additional.Start = true;
	
	this.Comments.CheckMarks();

	if (this.Action.Additional.TrackMove)
		this.private_FinalizeRemoveTrackMove();

	if (this.TrackMoveId)
		this.private_FinalizeCheckTrackMove();

	if (this.Action.Additional.ValidateForm)
		this.private_FinalizeValidateForm();

	if (this.Action.CancelAction)
	{
		this.Action.Additional.Start = false;
		return;
	}

	if (this.Action.Additional.FormChange)
		this.private_FinalizeFormChange();

	if (this.Action.Additional.RadioRequired)
		this.private_FinalizeRadioRequired();

	if (this.Action.Additional.CommentPosition)
		this.private_FinalizeUpdateCommentPosition();

	if (this.Action.Additional.ValidateComplexFields)
		this.private_FinalizeValidateComplexFields();

	if (this.Action.Additional.ContentControlChange)
		this.private_FinalizeContentControlChange();
	
	if (this.OFormDocument)
		this.OFormDocument.onEndAction();
	
	this.Action.Additional.Start = false;
};
/**
 * Пересчитываем нумерацию строк
 */
CDocument.prototype.RecalculateLineNumbers = function()
{
	this.UpdateLineNumbersInfo();
	this.Redraw(-1, -1);
};
CDocument.prototype.private_UpdatePlaceholders = function ()
{
	const arrRet = [];
	const graphicPages = this.DrawingObjects.graphicPages;
	for (let i = 0; i < graphicPages.length; i += 1)
	{
		graphicPages[i] && graphicPages[i].getPlaceholdersControls(arrRet);
	}
    if (this.DrawingDocument.placeholders)
    {
      this.DrawingDocument.placeholders.update(arrRet);
    }
};
CDocument.prototype.private_FinalizeRemoveTrackMove = function()
{
	for (var sMoveId in this.Action.Additional.TrackMove)
	{
		var oMarks = this.Action.Additional.TrackMove[sMoveId];

		if (oMarks)
		{
			if (oMarks.To.Start)
				oMarks.To.Start.RemoveThisMarkFromDocument();

			if (oMarks.To.End)
				oMarks.To.End.RemoveThisMarkFromDocument();

			if (oMarks.From.Start)
				oMarks.From.Start.RemoveThisMarkFromDocument();

			if (oMarks.From.End)
				oMarks.From.End.RemoveThisMarkFromDocument();
		}

		this.Action.Recalculate = true;
	}
};
CDocument.prototype.private_FinalizeCheckTrackMove = function()
{
	var sMoveId = this.TrackMoveId;
	var oMarks  = this.TrackRevisionsManager.GetMoveMarks(sMoveId);

	if (oMarks
		&& (!oMarks.To.Start
		|| !oMarks.To.Start.IsUseInDocument()
		|| !oMarks.To.End
		|| !oMarks.To.End.IsUseInDocument()
		|| !oMarks.From.Start
		|| !oMarks.From.Start.IsUseInDocument()
		|| !oMarks.From.End
		|| !oMarks.From.End.IsUseInDocument()))
	{
		// TODO: Вообще лучше переделать схему, потому что она не всегда правильно работает
		//       Если удаляемый элемент попадает в ранее добавленный текст, тогда и метка удалится, хотя такого быть не
		//       должно

		this.TrackMoveId = null;
		this.RemoveTrackMoveMarks(sMoveId);

		if (oMarks.To.Start)
			oMarks.To.Start.RemoveThisMarkFromDocument();

		if (oMarks.To.End)
			oMarks.To.End.RemoveThisMarkFromDocument();

		if (oMarks.From.Start)
			oMarks.From.Start.RemoveThisMarkFromDocument();

		if (oMarks.From.End)
			oMarks.From.End.RemoveThisMarkFromDocument();

		this.Action.Recalculate = true;
	}
};
CDocument.prototype.private_FinalizeValidateForm = function()
{
	let isCompositeInput = this.IsCompositeInputInProgress();

	let isCancelAction = this.Action.CancelAction;

	let arrForms = [];
	for (var sId in this.Action.Additional.ValidateForm)
	{
		let oForm = this.Action.Additional.ValidateForm[sId];
		if (!isCompositeInput && !this.FormsManager.ValidateChangeOnFly(oForm))
			this.Action.CancelAction = true;

		arrForms.push(oForm);
	}

	if (1 === this.Action.PointsCount
		&& !isCancelAction
		&& (AscDFH.historydescription_Document_BackSpaceButton === this.Action.Description
			|| AscDFH.historydescription_Document_DeleteButton === this.Action.Description))
	{
		this.Action.CancelAction = false;
	}

	if (this.Action.CancelAction)
		return;

	// По логике заполнять одновременно более одной формы нельзя
	if (1 === arrForms.length && !arrForms[0].IsPlaceHolder())
		this.History.SetAdditionalFormFilling(arrForms[0], this.Action.PointsCount);
};
CDocument.prototype.private_FinalizeFormChange = function()
{
	this.Action.Additional.FormChangeStart = true;

	for (var sKey in this.Action.Additional.FormChange)
	{
		let oForm = this.Action.Additional.FormChange[sKey];
		
		if (oForm.IsFixedForm() && oForm.IsMainForm() && oForm.IsAutoFitContent())
			this.CheckFormAutoFit(oForm);
		
		this.FormsManager.OnChange(oForm);
		this.Action.Recalculate = true;
	}

	delete this.Action.Additional.FormChangeStart;
};
CDocument.prototype.private_FinalizeFormAutoFit = function(isFastRecalc)
{
	for (var nIndex = 0, nCount = this.Action.Additional.FormAutoFit.length; nIndex < nCount; ++nIndex)
	{
		this.Action.Additional.FormAutoFit[nIndex].ProcessAutoFitContent(isFastRecalc);
	}

	if (this.Action.Additional.FormAutoFit.length)
		this.Action.Recalculate = true;
};
CDocument.prototype.private_FinalizeRadioRequired = function()
{
	for (var sGroupKey in this.Action.Additional.RadioRequired)
	{
		var isRequired = this.Action.Additional.RadioRequired[sGroupKey];

		var arrRadioGroup = this.FormsManager.GetRadioButtons(sGroupKey);
		for (var nIndex = 0, nCount = arrRadioGroup.length; nIndex < nCount; ++nIndex)
		{
			var oRadioButton = arrRadioGroup[nIndex];
			if (oRadioButton.IsFormRequired() !== isRequired)
				oRadioButton.SetFormRequired(isRequired);
		}

		this.Action.Recalculate = true;
	}
};
CDocument.prototype.private_FinalizeUpdateCommentPosition = function()
{
	var oChangedComments = {};
	for (var sId in this.Action.Additional.CommentPosition)
	{
		var oComment = this.Action.Additional.CommentPosition[sId];
		this.Comments.UpdateCommentPosition(oComment, oChangedComments);
	}

	this.Api.sync_ChangeCommentLogicalPosition(oChangedComments, this.Comments.GetCommentsPositionsCount());
};
CDocument.prototype.private_FinalizeContentControlChange = function()
{
	for (let sId in this.Action.Additional.ContentControlChange)
	{
		this.Api.asc_OnChangeContentControl(this.Action.Additional.ContentControlChange[sId]);
	}
};
CDocument.prototype.private_FinalizeCheckFocusAndBlurCC = function()
{

};
/**
 * Данная функция предназначена для отключения пересчета. Это может быть полезно, т.к. редактор всегда запускает
 * пересчет после каждого действия.
 */
CDocument.prototype.TurnOff_Recalculate = function()
{
    this.TurnOffRecalc++;
};
/**
 * Включаем пересчет, если он был отключен.
 * @param {boolean} bRecalculate Запускать ли пересчет
 */
CDocument.prototype.TurnOn_Recalculate = function(bRecalculate)
{
    this.TurnOffRecalc--;

    if (bRecalculate)
        this.Recalculate();
};
/**
 * Проверяем включен ли пересчет.
 * @returns {boolean}
 */
CDocument.prototype.Is_OnRecalculate = function()
{
    if (0 === this.TurnOffRecalc)
        return true;

    return false;
};
/**
 * Пересчитываем весь документ без включения таймеров
 * @param {boolean} isFromStart
 * @param {number} [nPagesCount=-1]
 */
CDocument.prototype.RecalculateAllAtOnce = function(isFromStart, nPagesCount)
{
	//var nStartTime = new Date().getTime();

	let isUseRecursion = this.FullRecalc.UseRecursion;

	this.FullRecalc.UseRecursion = false;
	this.FullRecalc.Continue     = false;

	if (isFromStart)
	{
		this.Reset_RecalculateCache();
		this.RecalculateWithParams({
			Inline   : {Pos : 0, PageNum : 0},
			Flow     : [],
			HdrFtr   : [],
			Drawings : {All : true, Map : {}}
		}, true);
	}
	else
	{
		this.private_Recalculate(undefined, true);
	}

	while (this.FullRecalc.Continue)
	{
		this.FullRecalc.Continue = false;
		this.Recalculate_Page();

		// Если задано количество страниц для пересчета, то считаем, что страница пересчитана, если текущая PageIndex + 2
		if (undefined !== nPagesCount && null !== nPagesCount && nPagesCount > 0 && this.FullRecalc.PageIndex >= nPagesCount + 1)
			break;
	}

	this.FullRecalc.UseRecursion = isUseRecursion;

	//console.log("RecalcTime: " + ((new Date().getTime() - nStartTime) / 1000));
};
/**
 * Запускаем пересчет документа.
 * @param _RecalcData
 * @param [isForceStrictRecalc=false] {boolean} Запускать ли пересчет первый раз без таймера
 * @param [nNoTimerPageIndex=undefined] {number} Номер страницы, до которой мы считаем без таймера
 */
CDocument.prototype.private_Recalculate = function(_RecalcData, isForceStrictRecalc, nNoTimerPageIndex)
{
	if (this.RecalcInfo.Paused)
		this.ResumeRecalculate();

	if (this.RecalcInfo.Is_NeedRecalculateFromStart())
	{
		this.RecalcInfo.Set_NeedRecalculateFromStart(false);
		this.RecalculateFromStart();
		return document_recalcresult_NoRecal;
	}

	this.DocumentOutline.Update();
	this.MathTrackHandler.Update();

    if (true !== this.Is_OnRecalculate())
        return document_recalcresult_NoRecal;

    this.private_ClearSearchOnRecalculate();

    // Обновляем позицию курсора
    this.NeedUpdateTarget = true;

    // Увеличиваем номер пересчета
    this.RecalcId++;

    // Если задан параметр _RecalcData, тогда мы не можем ориентироваться на историю
    if (undefined === _RecalcData)
    {
    	var arrChanges = this.History.GetNonRecalculatedChanges();

    	if (this.private_RecalculateFastRunRange(arrChanges))
			return document_recalcresult_FastRange;

    	if (this.private_RecalculateFastParagraph(arrChanges))
    		return document_recalcresult_FastParagraph;
    }

	// // Recalculation LOG
    // console.log("Regular Recalculation");

    var ChangeIndex = 0;
    var MainChange  = false;

    // Получаем данные об произошедших изменениях
    var RecalcData = History.Get_RecalcData(_RecalcData);

    History.Reset_RecalcIndex();

    this.DrawingObjects.recalculate(RecalcData.Drawings);

    // 1. Пересчитываем все автофигуры, которые нужно пересчитать. Изменения в них ни на что не влияют.
    for (var GraphIndex = 0; GraphIndex < RecalcData.Flow.length; GraphIndex++)
    {
        RecalcData.Flow[GraphIndex].recalculateDocContent();
    }

    // 2. Просмотрим все колонтитулы, которые подверглись изменениям. Найдем стартовую страницу, с которой надо
    //    запустить пересчет.

    var SectPrIndex = -1;
    for (var HdrFtrIndex = 0; HdrFtrIndex < RecalcData.HdrFtr.length; HdrFtrIndex++)
    {
        var HdrFtr    = RecalcData.HdrFtr[HdrFtrIndex];
        var FindIndex = this.SectionsInfo.Find_ByHdrFtr(HdrFtr);

        if (-1 === FindIndex)
            continue;

        // Колонтитул может быть записан в данной секции, но в ней не использоваться. Нам нужно начинать пересчет
        // с места использования данного колонтитула.

        var SectPr     = this.SectionsInfo.Get_SectPr2(FindIndex).SectPr;
        var HdrFtrInfo = SectPr.GetHdrFtrInfo(HdrFtr);

        if (null !== HdrFtrInfo)
        {
            var bHeader = HdrFtrInfo.Header;
            var bFirst  = HdrFtrInfo.First;
            var bEven   = HdrFtrInfo.Even;

            var CheckSectIndex = -1;

            if (true === bFirst)
            {
                var CurSectIndex = FindIndex;
                var SectCount    = this.SectionsInfo.Elements.length;

                while (CurSectIndex < SectCount)
                {
                    var CurSectPr = this.SectionsInfo.Get_SectPr2(CurSectIndex).SectPr;

                    if (FindIndex === CurSectIndex || null === CurSectPr.GetHdrFtr(bHeader, bFirst, bEven))
                    {
                        if (true === CurSectPr.Get_TitlePage())
                        {
                            CheckSectIndex = CurSectIndex;
                            break;
                        }

                    }
                    else
                    {
                        // Если мы попали сюда, значит данный колонтитул нигде не используется
                        break;
                    }

                    CurSectIndex++;
                }
            }
            else if (true === bEven)
            {
                if (true === EvenAndOddHeaders)
                    CheckSectIndex = FindIndex;
            }
            else
            {
                CheckSectIndex = FindIndex;
            }

            if (-1 !== CheckSectIndex && ( -1 === SectPrIndex || CheckSectIndex < SectPrIndex ))
                SectPrIndex = CheckSectIndex;
        }
    }

    // 3. Пересчитываем нумерацию строк отдельно, если нужно
	if (true === RecalcData.LineNumbers)
	{
		this.RecalculateLineNumbers();
	}

	if (-1 === RecalcData.Inline.Pos && -1 === SectPrIndex)
	{
		// Никаких изменений не было ни с самим документом, ни секциями
		ChangeIndex               = -1;
		RecalcData.Inline.PageNum = 0;
	}
	else if (-1 === RecalcData.Inline.Pos)
	{
		// Были изменения только внутри секций
		MainChange = false;

		// Выставляем начало изменений на начало секции
		ChangeIndex               = ( 0 === SectPrIndex ? 0 : this.SectionsInfo.Get_SectPr2(SectPrIndex - 1).Index + 1 );
		RecalcData.Inline.PageNum = 0;
	}
	else if (-1 === SectPrIndex)
	{
		// Изменения произошли только внутри основного документа
		MainChange = true;

		ChangeIndex = RecalcData.Inline.Pos;
	}
	else
	{
		// Изменения произошли и внутри документа, и внутри секций. Смотрим на более ранюю точку начала изменений
		// для секций и основоной части документа.
		MainChange = true;

		ChangeIndex = RecalcData.Inline.Pos;

		var ChangeIndex2 = ( 0 === SectPrIndex ? 0 : this.SectionsInfo.Get_SectPr2(SectPrIndex - 1).Index + 1 );

		if (ChangeIndex2 <= ChangeIndex)
		{
			ChangeIndex               = ChangeIndex2;
			RecalcData.Inline.PageNum = 0;
		}
	}

	// Найдем начальную страницу, с которой мы начнем пересчет
	var StartPage  = 0;
	var StartIndex = 0;

	var isUseTimeout = false;

	if (ChangeIndex < 0 && true === RecalcData.NotesEnd)
	{
		// Сюда мы попадаем при рассчете сносок, которые выходят за пределы самого документа
		StartIndex = this.Content.length;
		StartPage  = RecalcData.NotesEndPage;
		MainChange = true;
	}
	else
	{
		if (ChangeIndex < 0)
		{
			this.DrawingDocument.ClearCachePages();
			this.DrawingDocument.FirePaint();
			this.UpdatePlaceholders();
			return document_recalcresult_LongRecalc;
		}
		else if (ChangeIndex >= this.Content.length)
		{
			ChangeIndex = this.Content.length - 1;
		}

		// Проверяем предыдущие элементы на наличие параметра KeepNext, но не более, чем на 1 страницу
		var nTempPage  = this.private_GetPageByPos(ChangeIndex);
		var nTempIndex = (-1 !== nTempPage && nTempPage > 1) ? this.Pages[nTempPage - 1].Pos : 0;
		while (ChangeIndex > nTempIndex)
		{
			var PrevElement = this.Content[ChangeIndex - 1];
			if (type_Paragraph === PrevElement.GetType() && PrevElement.IsKeepNext())
			{
				ChangeIndex--;
				RecalcData.Inline.PageNum = PrevElement.Get_AbsolutePage(PrevElement.GetPagesCount());
				// считаем, что изменилась последняя страница
			}
			else
			{
				break;
			}
		}

		var ChangedElement = this.Content[ChangeIndex];
		if (ChangedElement.GetPagesCount() > 0 && -1 !== ChangedElement.GetIndex() && ChangedElement.Get_StartPage_Absolute() < RecalcData.Inline.PageNum - 1)
		{
			StartPage  = ChangedElement.GetStartPageForRecalculate(RecalcData.Inline.PageNum - 1);

			if (!this.FullRecalc.Id || StartPage < this.FullRecalc.PageIndex)
				StartIndex = this.Pages[StartPage].Pos;
			else
				StartIndex = this.FullRecalc.StartIndex;
		}
		else
		{
			var nTempPage = this.private_GetPageByPos(ChangeIndex);
			if (-1 !== nTempPage)
			{
				StartPage  = nTempPage;
				StartIndex = this.Pages[nTempPage].Pos;
			}

			if (ChangeIndex === StartIndex && StartPage < RecalcData.Inline.PageNum)
				StartPage = RecalcData.Inline.PageNum - 1;
		}

		// Если у нас уже начался долгий пересчет, тогда мы его останавливаем, и запускаем новый с текущими параметрами.
		// Здесь возможен случай, когда мы долгий пересчет основной части документа останавливаем из-за пересчета
		// колонтитулов, в этом случае параметр MainContentPos не меняется, и мы будем пересчитывать только колонтитулы
		// либо до страницы, на которой они приводят к изменению основную часть документа, либо до страницы, где
		// остановился предыдущий пересчет.

		if (null !== this.FullRecalc.Id)
		{
			// В такой ситуации не надо запускать заново пересчет
			if (this.FullRecalc.StartIndex < StartIndex || (this.FullRecalc.StartIndex === StartIndex && this.FullRecalc.PageIndex <= StartPage))
			{
				// // Recalculation LOG
				// console.log("No need to recalc");
				this.UpdatePlaceholders();
				return document_recalcresult_LongRecalc;
			}

			clearTimeout(this.FullRecalc.Id);
			this.FullRecalc.Id = null;
			this.DrawingDocument.OnEndRecalculate(false);
		}
		else if (null !== this.HdrFtrRecalc.Id)
		{
			clearTimeout(this.HdrFtrRecalc.Id);
			this.HdrFtrRecalc.Id = null;
			this.DrawingDocument.OnEndRecalculate(false);
		}
	}

	this.HdrFtrRecalc.PageCount = -1;

    // Очищаем данные пересчета
    this.RecalcInfo.Reset();

    this.FullRecalc.PageIndex         = StartPage;
    this.FullRecalc.SectionIndex      = 0;
    this.FullRecalc.ColumnIndex       = 0;
    this.FullRecalc.StartIndex        = StartIndex;
    this.FullRecalc.Start             = true;
    this.FullRecalc.StartPage         = StartPage;
    this.FullRecalc.ResetStartElement = this.private_RecalculateIsNewSection(StartPage, StartIndex);
    this.FullRecalc.Endnotes          = this.Endnotes.IsContinueRecalculateFromPrevPage(StartPage);
    this.FullRecalc.StartPagesCount   = undefined !== nNoTimerPageIndex ? Math.min(100, Math.max(nNoTimerPageIndex - StartPage, 2)) : 2;
	this.FullRecalc.StartTime         = performance.now();
	this.FullRecalc.TimerStartTime    = this.FullRecalc.StartTime;
	this.FullRecalc.TimerStartPage    = StartPage;


	// Если у нас произошли какие-либо изменения с основной частью документа, тогда начинаем его пересчитывать сразу,
    // а если изменения касались только секций, тогда пересчитываем основную часть документа только с того места, где
    // остановился предыдущий пересчет, либо с того места, где изменения секций приводят к пересчету документа.
    if (true === MainChange)
	{
		if (StartIndex > 0)
		{
			// В текущей схеме нам достаточно обновить у предыдущего элемента RecalcId. По хорошему
			// это надо делать у всех элементов до StartIndex
			let lastParagraph = this.Content[StartIndex - 1].GetLastParagraph();
			if (lastParagraph)
				lastParagraph.UpdateEndInfoRecalcId();
		}
		
		this.FullRecalc.MainStartPos = StartIndex;
	}

    this.DrawingDocument.OnStartRecalculate(StartPage);


    // TODO: По большому счету всегда можно запускать пересчет с нулевым таймаутом, но сейчас при переносе картинок и
	//       некоторых других действиях нам важно, чтобы пересчет первый раз сработал сразу. Поэтому запускаем пересчет
	//       на таймере ТОЛЬКО если он уже был запущен на таймере до этого. Если избавиться от первого условия, то
	//       запускать на таймере можно будет всегда.
    if (isUseTimeout && !isForceStrictRecalc)
	{
		this.FullRecalc.Id = setTimeout(Document_Recalculate_Page, 0);
	}
    else
	{
		this.Recalculate_Page();
	}
	this.UpdatePlaceholders();
    return document_recalcresult_LongRecalc;
};
/**
 * Запускаем пересчет документа.
 * @param oRecalcData
 * @param [isForceStrictRecalc=false] {boolean} Запускать ли пересчет первый раз без таймера
 * @param [nNoTimerPageIndex=undefined] {number} Номер страницы, до которой мы считаем без таймера
 */
CDocument.prototype.RecalculateWithParams = function(oRecalcData, isForceStrictRecalc, nNoTimerPageIndex)
{
	this.private_Recalculate(oRecalcData, isForceStrictRecalc, nNoTimerPageIndex);
};
CDocument.prototype.RecalculateByChanges = function(arrChanges, nStartIndex, nEndIndex, isForceStrictRecalc, nNoTimerPageIndex)
{
	this.private_ClearSearchOnRecalculate();

	// Обновляем позицию курсора
	this.NeedUpdateTarget = true;

	// Увеличиваем номер пересчета
	this.RecalcId++;

	if (this.private_RecalculateFastRunRange(arrChanges, nStartIndex, nEndIndex))
		return;

	if (this.private_RecalculateFastParagraph(arrChanges, nStartIndex, nEndIndex))
		return;

	var oRecalcData = this.History.Get_RecalcData(null, arrChanges, nStartIndex, nEndIndex);
	this.RecalculateWithParams(oRecalcData, isForceStrictRecalc, nNoTimerPageIndex);
};
CDocument.prototype.private_RecalculateFastRunRange = function(arrChanges, nStartIndex, nEndIndex)
{
	var _nStartIndex = undefined !== nStartIndex ? nStartIndex : 0;
	var _nEndIndex   = undefined !== nEndIndex ? nEndIndex : arrChanges.length - 1;

	var oRun = null;
	for (var nIndex = _nStartIndex; nIndex <= _nEndIndex; ++nIndex)
	{
		var oChange = arrChanges[nIndex];

		if (oChange.IsDescriptionChange())
			continue;

		if (!oRun)
			oRun = oChange.GetClass();
		else if (oRun !== oChange.GetClass())
			return false;
	}

	if (!oRun || !(oRun instanceof ParaRun) || !oRun.GetParagraph())
		return false;

	var oParaPos = oRun.GetSimpleChangesRange(arrChanges, _nStartIndex, _nEndIndex);
	if (oParaPos)
	{
		var oParagraph = oRun.GetParagraph();
		var nPageIndex = oParagraph.RecalculateFastRunRange(oParaPos);

		if (-1 !== nPageIndex && this.Pages[nPageIndex])
		{
			// Если за данным параграфом следовал пустой параграф с новой секцией, тогда его тоже надо пересчитать.
			var NextElement = oParagraph.Get_DocumentNext();
			if (null !== NextElement && true === this.Pages[nPageIndex].Check_EndSectionPara(NextElement))
				this.private_RecalculateEmptySectionParagraph(NextElement, oParagraph, nPageIndex, oParagraph.Get_AbsoluteColumn(oParagraph.Get_PagesCount() - 1), oParagraph.Get_ColumnsCount());

			// Перерисуем страницу, на которой произошли изменения
			this.DrawingDocument.OnRecalculatePage(nPageIndex, this.Pages[nPageIndex]);
			this.DrawingDocument.OnEndRecalculate(false, true);
			this.History.Reset_RecalcIndex();
			this.private_UpdateCursorXY(true, true);

			if (oParagraph.Parent && oParagraph.Parent.GetTopDocumentContent)
			{
				var oTopDocument = oParagraph.Parent.GetTopDocumentContent();
				if (oTopDocument instanceof CFootEndnote)
					oTopDocument.OnFastRecalculate();
			}

			// // Recalculation LOG
			// console.log("Fast Recalculation RunRange, PageIndex=" + nPageIndex);
			const oGraphicPage = this.DrawingObjects.graphicPages[nPageIndex];
			if (!!(oGraphicPage && oGraphicPage.getAllDrawings().length))
			{
				this.UpdatePlaceholders();
			}
			return true;
		}
	}

	return false;
};
CDocument.prototype.private_RecalculateFastParagraph = function(arrChanges, nStartIndex, nEndIndex)
{
	var _nStartIndex = undefined !== nStartIndex ? nStartIndex : 0;
	var _nEndIndex   = undefined !== nEndIndex ? nEndIndex : arrChanges.length - 1;


	// Смотрим, чтобы изменения происходили только внутри параграфов. Если есть изменение,
	// которое не возвращает параграф, значит возвращаем null.
	// А также проверяем, что каждое из этих изменений влияет только на параграф.
	var arrParagraphs = [];
	for (var nIndex = _nStartIndex; nIndex <= _nEndIndex; ++nIndex)
	{
		var oChange = arrChanges[nIndex];
		var oClass  = oChange.GetClass();
		var oPara   = null;

		if (oClass instanceof Paragraph)
			oPara = oClass;
		else if (oClass instanceof AscCommon.CTableId || oClass instanceof AscCommon.CComments)
			continue;
		else if (oClass.GetParagraph)
			oPara = oClass.GetParagraph();
		else
			return false;

		// Такое может быть, если класс еще не приписан ни к какому параграфу. Либо класс дальше небудет
		// использован, либо его добавят в параграф и в этом изменении мы отметим параграф
		// Поэтому мы не отказываемся от быстрого пересчета в данной ситуации
		if (!oPara)
			continue;

		if (!oClass.IsParagraphSimpleChanges || !oClass.IsParagraphSimpleChanges(oChange))
			return false;

		var isAdd = true;
		for (var nParaIndex = 0, nParasCount = arrParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
		{
			if (oPara === arrParagraphs[nParaIndex])
			{
				isAdd = false;
				break;
			}
		}

		if (isAdd)
			arrParagraphs.push(oPara);
	}

	if (arrParagraphs.length > 0)
	{
		var oFastPages     = {};
		var bCanFastRecalc = true;
		for (var nSimpleIndex = 0, nSimplesCount = arrParagraphs.length; nSimpleIndex < nSimplesCount; ++nSimpleIndex)
		{
			var oSimplePara  = arrParagraphs[nSimpleIndex];
			var arrFastPages = oSimplePara.Recalculate_FastWholeParagraph();
			if (!arrFastPages || arrFastPages.length <= 0)
			{
				bCanFastRecalc = false;
				break;
			}

			for (var nFastPageIndex = 0, nFastPagesCount = arrFastPages.length; nFastPageIndex < nFastPagesCount; ++nFastPageIndex)
			{
				oFastPages[arrFastPages[nFastPageIndex]] = arrFastPages[nFastPageIndex];

				if (!this.Pages[arrFastPages[nFastPageIndex]])
				{
					bCanFastRecalc = false;
					break;
				}
			}

			if (!bCanFastRecalc)
				break;

			// Если изменения произошли на последней странице параграфа, и за данным параграфом следовал
			// пустой параграф с новой секцией, тогда его тоже надо пересчитать.
			var oNextElement  = oSimplePara.Get_DocumentNext();
			var nLastFastPage = arrFastPages[arrFastPages.length - 1];
			if (null !== oNextElement && true === this.Pages[nLastFastPage].Check_EndSectionPara(oNextElement))
				this.private_RecalculateEmptySectionParagraph(oNextElement, oSimplePara, nLastFastPage, oSimplePara.GetAbsoluteColumn(oSimplePara.GetPagesCount() - 1), oSimplePara.GetColumnsCount());
		}


		if (bCanFastRecalc)
		{
			let bUpdatePlaceholders = false;
			for (var nPageIndex in oFastPages)
			{
				// // Recalculation LOG
				// console.log("Fast Recalculation Paragraph, PageIndex=" + nPageIndex);
				this.DrawingDocument.OnRecalculatePage(oFastPages[nPageIndex], this.Pages[nPageIndex]);
				if (!bUpdatePlaceholders)
				{
					const oGraphicPage = this.DrawingObjects.graphicPages[nPageIndex];
					bUpdatePlaceholders = !!(oGraphicPage && oGraphicPage.getAllDrawings().length);
				}
			}

			this.DrawingDocument.OnEndRecalculate(false, true);
			this.History.Reset_RecalcIndex();
			this.private_UpdateCursorXY(true, true);

			for (var nSimpleIndex = 0, nSimplesCount = arrParagraphs.length; nSimpleIndex < nSimplesCount; ++nSimpleIndex)
			{
				var oSimplePara = arrParagraphs[nSimpleIndex];
				if (oSimplePara.Parent && oSimplePara.Parent.GetTopDocumentContent)
				{
					var oTopDocument = oSimplePara.Parent.GetTopDocumentContent();
					if (oTopDocument instanceof CFootEndnote)
						oTopDocument.OnFastRecalculate();
				}
			}
			if (bUpdatePlaceholders)
            {
            	this.UpdatePlaceholders();
			}
			return true;
		}
	}

	return false;
};
/**
 * Пересчитываем следующую страницу.
 */
CDocument.prototype.Recalculate_Page = function()
{
    var SectionIndex = this.FullRecalc.SectionIndex;
    var PageIndex    = this.FullRecalc.PageIndex;
    var ColumnIndex  = this.FullRecalc.ColumnIndex;
    var bStart       = this.FullRecalc.Start;        // флаг, который говорит о том, рассчитываем мы эту страницу первый раз или нет (за один общий пересчет)
    var StartIndex   = this.FullRecalc.StartIndex;

    //var nStartTime = (new Date()).getTime();

    if (0 === SectionIndex && 0 === ColumnIndex && true === bStart)
    {
        var OldPage = ( undefined !== this.Pages[PageIndex] ? this.Pages[PageIndex] : null );

        if (true === bStart)
        {
            var Page              = new CDocumentPage();
            this.Pages[PageIndex] = Page;
            Page.Pos              = StartIndex;

            if (true === this.HdrFtr.Recalculate(PageIndex))
                this.FullRecalc.MainStartPos = StartIndex;

            var SectPr = this.Layout.GetSectionByPos(StartIndex);
			var oFrame = SectPr.GetContentFrame(PageIndex);

            Page.Width          = SectPr.PageSize.W;
            Page.Height         = SectPr.PageSize.H;
            Page.Margins.Left   = oFrame.Left;
            Page.Margins.Top    = oFrame.Top;
            Page.Margins.Right  = oFrame.Right;
            Page.Margins.Bottom = oFrame.Bottom;

            Page.Sections[0] = new CDocumentPageSection();
            Page.Sections[0].Init(PageIndex, SectPr, this.Layout.GetSectionIndex(SectPr));
        }

        var Count = this.Content.length;

        // Проверяем нужно ли пересчитывать основную часть документа на данной странице
        var MainStartPos = this.FullRecalc.MainStartPos;
        if (null !== OldPage && ( -1 === MainStartPos || MainStartPos > StartIndex ))
        {
            if (OldPage.EndPos >= Count - 1 && PageIndex - this.Content[Count - 1].Get_StartPage_Absolute() >= this.Content[Count - 1].GetPagesCount() - 1)
            {
				// // Recalculation LOG
				// console.log("HdrFtr Recalculation " + PageIndex);

                this.Pages[PageIndex] = OldPage;
                this.DrawingDocument.OnRecalculatePage(PageIndex, this.Pages[PageIndex]);

                this.private_CheckCurPage();
                this.DrawingDocument.OnEndRecalculate(true);
                this.DrawingObjects.onEndRecalculateDocument(this.Pages.length);

                if (true === this.Selection.UpdateOnRecalc)
                {
                    this.Selection.UpdateOnRecalc = false;
                    this.DrawingDocument.OnSelectEnd();
                }

                this.FullRecalc.Id           = null;
                this.FullRecalc.MainStartPos = -1;

                return;
            }
            else if (undefined !== this.Pages[PageIndex + 1])
            {
				// // Recalculation LOG
				// console.log("HdrFtr Recalculation " + PageIndex);

                // Переходим к следующей странице
                this.Pages[PageIndex] = OldPage;
                this.DrawingDocument.OnRecalculatePage(PageIndex, this.Pages[PageIndex]);

                this.FullRecalc.PageIndex         = PageIndex + 1;
                this.FullRecalc.Start             = true;
                this.FullRecalc.StartIndex        = this.Pages[PageIndex + 1].Pos;
                this.FullRecalc.ResetStartElement = false;

                var CurSectInfo  = this.Layout.GetSectionInfo(this.Pages[PageIndex + 1].Pos);
                var PrevSectInfo = this.Layout.GetSectionInfo(this.Pages[PageIndex].EndPos);

                if (PrevSectInfo !== CurSectInfo)
                    this.FullRecalc.ResetStartElement = true;

				if (this.FullRecalc.UseRecursion)
				{
					if (window["NATIVE_EDITOR_ENJINE_SYNC_RECALC"] === true)
					{
						if (this.private_IsStartTimeoutOnRecalc(PageIndex))
						{
							if (window["native"] && window["native"]["WC_CheckSuspendRecalculate"] !== undefined)
							{
								//if (window["native"]["WC_CheckSuspendRecalculate"]())
								//    return;

								this.FullRecalc.Id             = setTimeout(Document_Recalculate_Page, 10);
								this.FullRecalc.TimerStartPage = PageIndex + 1;
								return;
							}
						}

						this.Recalculate_Page();
						return;
					}


					if (this.private_IsStartTimeoutOnRecalc(PageIndex))
					{
						// console.log("Page " + _PageIndex);
						// console.log("Pages delta " + (PageIndex + 1 - this.FullRecalc.TimerStartPage));
						this.FullRecalc.Id             = setTimeout(Document_Recalculate_Page, 20);
						this.FullRecalc.TimerStartPage = PageIndex + 1;
					}
					else
					{
						this.Recalculate_Page();
					}
				}
				else
				{
					this.FullRecalc.Continue = true;
				}

                return;
            }
        }
        else
        {
            if (true === bStart)
            {
                this.Pages.length = PageIndex + 1;

                this.DrawingObjects.createGraphicPage(PageIndex);
                this.DrawingObjects.resetDrawingArrays(PageIndex, this);
            }
        }

		// // Recalculation LOG
		// console.log("Regular Recalculation" + PageIndex);

        var StartPos = this.Get_PageContentStartPos(PageIndex, StartIndex);

		this.Footnotes.Reset(PageIndex, this.Layout.GetSectionByPos(StartIndex));

        this.Pages[PageIndex].ResetStartElement = this.FullRecalc.ResetStartElement;
        this.Pages[PageIndex].X                 = StartPos.X;
        this.Pages[PageIndex].XLimit            = StartPos.XLimit;
        this.Pages[PageIndex].Y                 = StartPos.Y;
        this.Pages[PageIndex].YLimit            = StartPos.YLimit;

        this.Pages[PageIndex].Sections[0].Y      = StartPos.Y;
        this.Pages[PageIndex].Sections[0].YLimit = StartPos.YLimit;
        this.Pages[PageIndex].Sections[0].Pos    = StartIndex;
        this.Pages[PageIndex].Sections[0].EndPos = StartIndex;
    }

	this.Endnotes.Reset(PageIndex, SectionIndex, ColumnIndex);

    this.Recalculate_PageColumn();

	//console.log("PageIndex " + PageIndex + " " + ((new Date()).getTime() - nStartTime)/ 1000);
};
/**
 * Пересчитываем следующую колоноку.
 */
CDocument.prototype.Recalculate_PageColumn                   = function()
{
    var PageIndex          = this.FullRecalc.PageIndex;
    var SectionIndex       = this.FullRecalc.SectionIndex;
    var ColumnIndex        = this.FullRecalc.ColumnIndex;
    var StartIndex         = this.FullRecalc.StartIndex;
    var bResetStartElement = this.FullRecalc.ResetStartElement;

	// // Recalculation LOG
	// console.log("Page " + PageIndex + " Section " + SectionIndex + " Column " + ColumnIndex + " Element " + StartIndex);
	// console.log(this.RecalcInfo);

    var StartPos = this.Get_PageContentStartPos2(PageIndex, ColumnIndex, 0, StartIndex);

    var X      = StartPos.X;
    var Y      = StartPos.Y;
    var YLimit = StartPos.YLimit;
    var XLimit = StartPos.XLimit;

    var Page        = this.Pages[PageIndex];
    var PageSection = Page.Sections[SectionIndex];
    var PageColumn  = PageSection.Columns[ColumnIndex];

    PageColumn.X           = X;
    PageColumn.XLimit      = XLimit;
    PageColumn.Y           = Y;
    PageColumn.YLimit      = YLimit;
    PageColumn.Pos         = StartIndex;
    PageColumn.Empty       = false;
    PageColumn.SpaceBefore = StartPos.ColumnSpaceBefore;
    PageColumn.SpaceAfter  = StartPos.ColumnSpaceAfter;
    PageColumn.Endnotes    = this.FullRecalc.Endnotes;

    this.Footnotes.ContinueElementsFromPreviousColumn(PageIndex, ColumnIndex, Y, YLimit);

    var SectPr       = this.Layout.GetSectionByPos(StartIndex);
    var ColumnsCount = SectPr.GetColumnsCount();

    var bReDraw             = true;
    var bContinue           = false;
    var _PageIndex          = PageIndex;
    var _ColumnIndex        = ColumnIndex;
    var _StartIndex         = StartIndex;
    var _SectionIndex       = SectionIndex;
    var _bStart             = false;
    var _bResetStartElement = false;
    var _bEndnotesContinue  = false;

    var Count = this.Content.length;
    var Index;

	var nEndIndex = Count;

	var isEndEndnoteRecalc = false;
	if (this.FullRecalc.Endnotes)
	{
		var nEndnoteSectionIndex = 0;
		var oEndnoteSectPr       = null;

		if (PageIndex <= 0)
		{
			// Такого не должно быть
			nEndnoteSectionIndex = this.Layout.GetSectionIndex(SectPr);
			oEndnoteSectPr       = SectPr;
		}
		else
		{
			if (ColumnIndex > 0)
				nEndnoteSectionIndex = this.Endnotes.GetLastSectionIndexOnPage(PageIndex);
			else
				nEndnoteSectionIndex = this.Endnotes.GetLastSectionIndexOnPage(PageIndex - 1);

			oEndnoteSectPr = this.SectionsInfo.Get(nEndnoteSectionIndex).SectPr;
		}

		var nEndnoteRecalcResult = this.Endnotes.Recalculate(X, Y, XLimit, YLimit - this.Footnotes.GetHeight(PageIndex, ColumnIndex), PageIndex, ColumnIndex, ColumnsCount, oEndnoteSectPr, nEndnoteSectionIndex, StartIndex >= Count);
		if (recalcresult2_End === nEndnoteRecalcResult)
		{
			PageColumn.EndPos = -1;

			// Сноски закончились на данной странице
			Y = this.Endnotes.GetPageBounds(PageIndex, ColumnIndex, this.Layout.GetSectionIndex(SectPr)).Bottom;
			PageColumn.Bounds.Bottom = Y;
			_bEndnotesContinue = false;

			if (StartIndex < Count)
			{
				isEndEndnoteRecalc = true;
			}
		}
		else
		{
			PageColumn.EndPos = StartIndex - 1;

			if (true === PageSection.IsCalculatingSectionBottomLine() && ColumnIndex >= ColumnsCount - 1)
			{
				PageSection.IterateBottomLineCalculation(true);

				bContinue           = true;
				_PageIndex          = PageIndex;
				_SectionIndex       = SectionIndex;
				_ColumnIndex        = 0;
				_StartIndex         = this.Pages[_PageIndex].Sections[_SectionIndex].Columns[0].Pos;
				_bStart             = false;
				_bResetStartElement = 0 === SectionIndex ? Page.ResetStartElement : true;
				_bEndnotesContinue  = this.Pages[_PageIndex].Sections[_SectionIndex].Columns[0].Endnotes === true;

				this.Pages[_PageIndex].Sections[_SectionIndex].Reset_Columns();

				bReDraw = false;
			}
			else
			{
				bContinue           = true;
				_PageIndex          = ColumnIndex >= ColumnsCount - 1 ? PageIndex + 1 : PageIndex;
				_ColumnIndex        = ColumnIndex >= ColumnsCount - 1 ? 0 : ColumnIndex + 1;
				_SectionIndex       = ColumnIndex >= ColumnsCount - 1 ? 0 : SectionIndex;
				_bEndnotesContinue  = true;
				_bStart             = true;
				_bResetStartElement = true;
				_bEndnotesContinue  = true;
			}

			nEndIndex = StartIndex; // Выставляем так, чтобы ничего не пересчитывать из основной части документа
		}
	}

    for (Index = StartIndex; Index < nEndIndex; ++Index)
    {
        // Пересчитываем элемент документа
        var Element = this.Content[Index];

        var RecalcResult = recalcresult_NextElement;
        var bFlow        = false;

        if (!isEndEndnoteRecalc)
		{
			var oFramePr = null;
			if (!Element.IsInline() || (!Element.IsParagraph() && (oFramePr = Element.GetFramePr()) && !oFramePr.IsInline()))
			{
				bFlow = true;

				// Проверяем PageBreak и ColumnBreak в предыдущей строке
				var isPageBreakOnPrevLine   = false;
				var isColumnBreakOnPrevLine = false;

				var PrevElement = Element.Get_DocumentPrev();

				if (null !== PrevElement && type_Paragraph === PrevElement.Get_Type() && true === PrevElement.Is_Empty() && undefined !== PrevElement.Get_SectionPr())
					PrevElement = PrevElement.Get_DocumentPrev();

				if (null !== PrevElement && type_Paragraph === PrevElement.Get_Type())
				{
					var bNeedPageBreak = true;
					if (undefined !== PrevElement.Get_SectionPr())
					{
						var PrevSectPr = PrevElement.Get_SectionPr();
						var CurSectPr  = this.SectionsInfo.Get_SectPr(Index).SectPr;
						if (c_oAscSectionBreakType.Continuous !== CurSectPr.Get_Type() || true !== CurSectPr.Compare_PageSize(PrevSectPr))
							bNeedPageBreak = false;
					}

					var EndLine = PrevElement.Pages[PrevElement.Pages.length - 1].EndLine;
					if (true === bNeedPageBreak && -1 !== EndLine && PrevElement.Lines[EndLine].Info & paralineinfo_BreakRealPage && Index !== Page.Pos)
						isPageBreakOnPrevLine = true;

					if (-1 !== EndLine && !(PrevElement.Lines[EndLine].Info & paralineinfo_BreakRealPage) && PrevElement.Lines[EndLine].Info & paralineinfo_BreakPage && Index !== PageColumn.Pos)
						isColumnBreakOnPrevLine = true;
				}

				if (true === isColumnBreakOnPrevLine)
				{
					RecalcResult = recalcresult_NextPage | recalcresultflags_LastFromNewColumn;
				}
				else if (true === isPageBreakOnPrevLine)
				{
					RecalcResult = recalcresult_NextPage | recalcresultflags_LastFromNewPage;
				}
				else
				{
					var RecalcInfo = {
						Element           : Element,
						X                 : X,
						Y                 : Y,
						XLimit            : XLimit,
						YLimit            : YLimit,
						PageIndex         : PageIndex,
						SectionIndex      : SectionIndex,
						ColumnIndex       : ColumnIndex,
						Index             : Index,
						StartIndex        : StartIndex,
						ColumnsCount      : ColumnsCount,
						ResetStartElement : bResetStartElement,
						RecalcResult      : RecalcResult
					};

					if (Element.IsTable() && !Element.IsInline())
						this.private_RecalculateFlowTable(RecalcInfo);
					else
						this.private_RecalculateFlowParagraph(RecalcInfo);

					Index        = RecalcInfo.Index;
					RecalcResult = RecalcInfo.RecalcResult;
				}
			}
			else
			{
				if ((0 === Index && 0 === PageIndex && 0 === ColumnIndex) || Index != StartIndex || (Index === StartIndex && true === bResetStartElement))
				{
					Element.Set_DocumentIndex(Index);
					Element.Reset(X, Y, XLimit, YLimit, PageIndex, ColumnIndex, ColumnsCount);
				}

				// Делаем как в Word: Обработаем особый случай, когда на данном параграфе заканчивается секция, и он
				// пустой. В такой ситуации этот параграф не добавляет смещения по Y, и сам приписывается в конец
				// предыдущего элемента. Второй подряд идущий такой же параграф обсчитывается по обычному.

				var SectInfoElement = this.SectionsInfo.Get_SectPr(Index);
				var PrevElement     = this.Content[Index - 1]; // может быть undefined, но в следующем условии сразу стоит проверка на Index > 0
				if (Index > 0 && (Index !== StartIndex || true !== bResetStartElement) && Index === SectInfoElement.Index && true === Element.IsEmpty() && (type_Paragraph !== PrevElement.GetType() || undefined === PrevElement.Get_SectionPr()))
				{
					RecalcResult = recalcresult_NextElement;

					this.private_RecalculateEmptySectionParagraph(Element, PrevElement, PageIndex, ColumnIndex, ColumnsCount);

					// Добавим в список особых параграфов
					this.Pages[PageIndex].EndSectionParas.push(Element);

					// Выставляем этот флаг, чтобы у нас не менялось значение по Y
					bFlow = true;
				}
				else
				{
					var ElementPageIndex = this.private_GetElementPageIndex(Index, PageIndex, ColumnIndex, ColumnsCount);
					RecalcResult         = Element.Recalculate_Page(ElementPageIndex);
				}
			}

			if (true != bFlow && (RecalcResult & recalcresult_NextElement || RecalcResult & recalcresult_NextPage))
			{
				var ElementPageIndex = this.private_GetElementPageIndex(Index, PageIndex, ColumnIndex, ColumnsCount);
				Y                    = Element.Get_PageBounds(ElementPageIndex).Bottom;
			}

			PageColumn.Bounds.Bottom = Y;
		}

        if (RecalcResult & recalcresult_CurPage)
        {
            bReDraw       = false;
            bContinue     = true;
            _PageIndex    = PageIndex;
            _SectionIndex = SectionIndex;
            _bStart       = false;

            if (RecalcResult & recalcresultflags_Column)
            {
                _ColumnIndex = ColumnIndex;
                _StartIndex  = StartIndex;
            }
            else
            {
                _ColumnIndex = 0;
                _StartIndex  = this.Pages[_PageIndex].Sections[_SectionIndex].Columns[0].Pos;
            }

            break;
        }
        else if (RecalcResult & recalcresult_NextElement)
		{
			var CurSectInfo = this.Layout.GetSectionInfo(Index);

			if (Index < Count - 1)
			{
				var NextSectInfo = this.Layout.GetSectionInfo(Index + 1);
				if (CurSectInfo !== NextSectInfo)
				{
					if (!isEndEndnoteRecalc)
					{
						PageColumn.EndPos  = Index;
						PageSection.EndPos = Index;
						Page.EndPos        = Index;

						if (this.Endnotes.HaveEndnotes(CurSectInfo.SectPr, false))
						{
							var nSectionIndexAbs = this.Layout.GetSectionIndex(CurSectInfo.SectPr);
							this.Endnotes.FillSection(PageIndex, ColumnIndex, CurSectInfo.SectPr, nSectionIndexAbs, false);
							var nEndnoteRecalcResult = this.Endnotes.Recalculate(X, Y, XLimit, YLimit - this.Footnotes.GetHeight(PageIndex, ColumnIndex), PageIndex, ColumnIndex, ColumnsCount, CurSectInfo.SectPr, nSectionIndexAbs, true);

							// Сноски закончились на данной странице
							Y = this.Endnotes.GetPageBounds(PageIndex, ColumnIndex, nSectionIndexAbs).Bottom;
							PageColumn.Bounds.Bottom = Y;

							if (recalcresult2_End !== nEndnoteRecalcResult)
							{
								if (true === PageSection.IsCalculatingSectionBottomLine() && ColumnIndex >= ColumnsCount - 1)
								{
									PageSection.IterateBottomLineCalculation(true);

									bContinue           = true;
									_PageIndex          = PageIndex;
									_SectionIndex       = SectionIndex;
									_ColumnIndex        = 0;
									_StartIndex         = this.Pages[_PageIndex].Sections[_SectionIndex].Columns[0].Pos;
									_bStart             = false;
									_bResetStartElement = 0 === SectionIndex ? Page.ResetStartElement : true;
									_bEndnotesContinue  = this.Pages[_PageIndex].Sections[_SectionIndex].Columns[0].Endnotes === true;

									this.Pages[_PageIndex].Sections[_SectionIndex].Reset_Columns();

									bReDraw = false;
								}
								else
								{
									bContinue           = true;
									_StartIndex         = Index;
									_PageIndex          = ColumnIndex >= ColumnsCount - 1 ? PageIndex + 1 : PageIndex;
									_ColumnIndex        = ColumnIndex >= ColumnsCount - 1 ? 0 : ColumnIndex + 1;
									_SectionIndex       = ColumnIndex >= ColumnsCount - 1 ? 0 : SectionIndex;
									_bEndnotesContinue  = true;
									_bStart             = true;
									_bResetStartElement = true;
								}

								break;
							}
						}
						else
						{
							this.Endnotes.ClearSection(this.Layout.GetSectionIndex(CurSectInfo.SectPr));
						}
					}

					if (c_oAscSectionBreakType.Continuous === NextSectInfo.SectPr.Get_Type() && true === CurSectInfo.SectPr.Compare_PageSize(NextSectInfo.SectPr) && this.Footnotes.IsEmptyPage(PageIndex))
					{
						// Новая секция начинается на данной странице. Нам надо получить новые поля данной секции, но
						// на данной странице мы будем использовать только новые горизонтальные поля, а поля по вертикали
						// используем от предыдущей секции.

						var SectionY = Y;
						for (var TempColumnIndex = 0; TempColumnIndex < ColumnsCount; ++TempColumnIndex)
						{
							if (PageSection.Columns[TempColumnIndex].Bounds.Bottom > SectionY)
								SectionY = PageSection.Columns[TempColumnIndex].Bounds.Bottom;
						}

						PageSection.YLimit = SectionY;

						if ((!PageSection.IsCalculatingSectionBottomLine() || PageSection.CanDecreaseBottomLine()) && ColumnsCount > 1 && PageSection.CanRecalculateBottomLine())
						{
							PageSection.IterateBottomLineCalculation(false);

							bContinue           = true;
							_PageIndex          = PageIndex;
							_SectionIndex       = SectionIndex;
							_ColumnIndex        = 0;
							_StartIndex         = this.Pages[_PageIndex].Sections[_SectionIndex].Columns[0].Pos;
							_bStart             = false;
							_bResetStartElement = 0 === SectionIndex ? Page.ResetStartElement : true;

							this.Pages[_PageIndex].Sections[_SectionIndex].Reset_Columns();

							break;
						}
						else
						{
							bContinue           = true;
							_PageIndex          = PageIndex;
							_SectionIndex       = SectionIndex + 1;
							_ColumnIndex        = 0;
							_StartIndex         = Index + 1;
							_bStart             = false;
							_bResetStartElement = true;

							var NewPageSection = new CDocumentPageSection();
							NewPageSection.Init(PageIndex, NextSectInfo.SectPr, this.Layout.GetSectionIndex(NextSectInfo.SectPr));
							NewPageSection.Pos           = Index;
							NewPageSection.EndPos        = Index;
							NewPageSection.Y             = SectionY + 0.001;
							NewPageSection.YLimit        = this.Pages[PageIndex].YLimit;
							NewPageSection.YLimit2       = this.Pages[PageIndex].YLimit;
							Page.Sections[_SectionIndex] = NewPageSection;
							// YLimit, YLimit2 проставляем здесь, потому что в функции Init учитываются настройки уже
							// новой секции, а нам нужно расчет вести с учетом отступов самой первой секции
							break;
						}
					}
					else
					{
						bContinue           = true;
						_PageIndex          = PageIndex + 1;
						_SectionIndex       = 0;
						_ColumnIndex        = 0;
						_StartIndex         = Index + 1;
						_bStart             = true;
						_bResetStartElement = true;
						break;
					}
				}
			}
			else if (this.Endnotes.HaveEndnotes(CurSectInfo.SectPr, true))
			{
				var nSectionIndexAbs = this.Layout.GetSectionIndex(CurSectInfo.SectPr);
				this.Endnotes.FillSection(PageIndex, ColumnIndex, CurSectInfo.SectPr, nSectionIndexAbs, true);
				var nEndnoteRecalcResult = this.Endnotes.Recalculate(X, Y, XLimit, YLimit - this.Footnotes.GetHeight(PageIndex, ColumnIndex), PageIndex, ColumnIndex, ColumnsCount, CurSectInfo.SectPr, nSectionIndexAbs, true);
				if (recalcresult2_End === nEndnoteRecalcResult)
				{
					// Сноски закончились на данной странице
					Y = this.Endnotes.GetPageBounds(PageIndex, ColumnIndex, nSectionIndexAbs).Bottom;
				}
				else
				{
					_bEndnotesContinue = true;
				}
			}
			else
			{
				this.Endnotes.ClearSection(this.Layout.GetSectionIndex(CurSectInfo.SectPr));
			}
		}
        else if (RecalcResult & recalcresult_NextPage)
        {
            if (true === PageSection.IsCalculatingSectionBottomLine() && PageSection.CanIncreaseBottomLine() && (RecalcResult & recalcresultflags_LastFromNewPage || ColumnIndex >= ColumnsCount - 1))
            {
                PageSection.IterateBottomLineCalculation(true);

                bContinue           = true;
                _PageIndex          = PageIndex;
                _SectionIndex       = SectionIndex;
                _ColumnIndex        = 0;
                _StartIndex         = this.Pages[_PageIndex].Sections[_SectionIndex].Columns[0].Pos;
                _bStart             = false;
                _bResetStartElement = 0 === SectionIndex ? Page.ResetStartElement : true;

                this.Pages[_PageIndex].Sections[_SectionIndex].Reset_Columns();

                bReDraw = false;
                break;
            }
            else if (RecalcResult & recalcresultflags_LastFromNewColumn)
            {
                PageColumn.EndPos  = Index - 1;
                PageSection.EndPos = Index - 1;
                Page.EndPos        = Index - 1;

                bContinue           = true;
                _SectionIndex       = SectionIndex;
                _ColumnIndex        = ColumnIndex + 1;
                _PageIndex          = PageIndex;
                _StartIndex         = Index;
                _bStart             = true;
                _bResetStartElement = true;

                if (_ColumnIndex >= ColumnsCount)
                {
                    _SectionIndex = 0;
                    _ColumnIndex  = 0;
                    _PageIndex    = PageIndex + 1;
                }
                else
                {
                    bReDraw = false;
                }

                break;
            }
            else if (RecalcResult & recalcresultflags_LastFromNewPage)
            {
                PageColumn.EndPos  = Index - 1;
                PageSection.EndPos = Index - 1;
                Page.EndPos        = Index - 1;

                bContinue = true;

                _SectionIndex       = 0;
                _ColumnIndex        = 0;
                _PageIndex          = PageIndex + 1;
                _StartIndex         = Index;
                _bStart             = true;
                _bResetStartElement = true;

                if (PageColumn.EndPos === PageColumn.Pos)
                {
                    var Element          = this.Content[PageColumn.Pos];
                    var ElementPageIndex = this.private_GetElementPageIndex(Index, PageIndex, ColumnIndex, ColumnsCount);
                    if (true === Element.IsEmptyPage(ElementPageIndex))
                        PageColumn.Empty = true;
                }

                for (var TempColumnIndex = ColumnIndex + 1; TempColumnIndex < ColumnsCount; ++TempColumnIndex)
                {
                    PageSection.Columns[TempColumnIndex].Empty  = true;
                    PageSection.Columns[TempColumnIndex].Pos    = Index;
                    PageSection.Columns[TempColumnIndex].EndPos = Index - 1;
                }

                break;
            }
            else if (RecalcResult & recalcresultflags_Page)
            {
                PageColumn.EndPos  = Index;
                PageSection.EndPos = Index;
                Page.EndPos        = Index;

                bContinue = true;

                _SectionIndex = 0;
                _ColumnIndex  = 0;
                _PageIndex    = PageIndex + 1;
                _StartIndex   = Index;
                _bStart       = true;

                if (PageColumn.EndPos === PageColumn.Pos)
                {
                    var Element          = this.Content[PageColumn.Pos];
                    var ElementPageIndex = this.private_GetElementPageIndex(Index, PageIndex, ColumnIndex, ColumnsCount);
                    if (true === Element.IsEmptyPage(ElementPageIndex))
                        PageColumn.Empty = true;
                }

                for (var TempColumnIndex = ColumnIndex + 1; TempColumnIndex < ColumnsCount; ++TempColumnIndex)
                {
                    var ElementPageIndex = this.private_GetElementPageIndex(Index, PageIndex, TempColumnIndex, ColumnsCount);
                    this.Content[Index].Recalculate_SkipPage(ElementPageIndex);

                    PageSection.Columns[TempColumnIndex].Empty  = true;
                    PageSection.Columns[TempColumnIndex].Pos    = Index;
                    PageSection.Columns[TempColumnIndex].EndPos = Index - 1;
                }

                break;
            }
            else
            {
                PageColumn.EndPos  = Index;
                PageSection.EndPos = Index;
                Page.EndPos        = Index;

                bContinue = true;

                _ColumnIndex = ColumnIndex + 1;
                if (_ColumnIndex >= ColumnsCount)
                {
                    _SectionIndex = 0;
                    _ColumnIndex  = 0;
                    _PageIndex    = PageIndex + 1;
                }
                else
                {
                    bReDraw = false;
                }

                _StartIndex = Index;
                _bStart     = true;

                if (PageColumn.EndPos === PageColumn.Pos)
                {
                    var Element          = this.Content[PageColumn.Pos];
                    var ElementPageIndex = this.private_GetElementPageIndex(Index, PageIndex, ColumnIndex, ColumnsCount);
                    if (true === Element.IsEmptyPage(ElementPageIndex))
                        PageColumn.Empty = true;
                }

                break;
            }
        }
        else if (RecalcResult & recalcresult_PrevPage)
        {
            bReDraw   = false;
            bContinue = true;
            if (RecalcResult & recalcresultflags_Column)
            {
                if (0 === ColumnIndex)
                {
                    _PageIndex    = Math.max(PageIndex - 1, 0);
                    _SectionIndex = this.Pages[_PageIndex].Sections.length - 1;
                    _ColumnIndex  = this.Pages[_PageIndex].Sections[_SectionIndex].Columns.length - 1;
                }
                else
                {
                    _PageIndex    = PageIndex;
                    _ColumnIndex  = ColumnIndex - 1;
                    _SectionIndex = SectionIndex;
                }
            }
            else
            {
                if (_SectionIndex > 0)
                {
                    // Сюда мы никогда не должны попадать
                }

                _PageIndex    = Math.max(PageIndex - 1, 0);
                _SectionIndex = this.Pages[_PageIndex].Sections.length - 1;
                _ColumnIndex  = 0;
            }

            _StartIndex = this.Pages[_PageIndex].Sections[_SectionIndex].Columns[_ColumnIndex].Pos;
            _bStart     = false;
            break;
        }

        if (docpostype_Content === this.GetDocPosType() && Index === this.CurPos.ContentPos)
        {
            if (type_Paragraph === Element.GetType())
                this.CurPage = PageIndex;
            else
                this.CurPage = PageIndex; // TODO: переделать
        }

        if (docpostype_Content === this.GetDocPosType() && ((true !== this.Selection.Use && Index === this.CurPos.ContentPos + 1) || (true === this.Selection.Use && Index === (Math.max(this.Selection.EndPos, this.Selection.StartPos) + 1))))
            this.private_UpdateCursorXY(true, true);
    }

    if (Index >= Count)
    {
        // До перерисовки селекта должны выставить
        Page.EndPos        = Count - 1;
        PageSection.EndPos = Count - 1;
        PageColumn.EndPos  = Count - 1;
    }

    if (Index >= Count || _PageIndex > PageIndex || _ColumnIndex > ColumnIndex)
    {
        this.private_RecalculateShiftFootnotes(PageIndex, ColumnIndex, Y, SectPr);
    }

    if (true === bReDraw)
    {
        this.DrawingDocument.OnRecalculatePage(PageIndex, this.Pages[PageIndex]);
    }

    if (Index >= Count)
    {
		// Пересчет основной части документа закончен. Возможна ситуация, при которой последние сноски с данной
		// страницы переносятся на следующую (т.е. остались непересчитанные сноски). Эти сноски нужно пересчитать
		if (_bEndnotesContinue || this.Footnotes.HaveContinuesFootnotes(PageIndex, ColumnIndex))
		{
			bContinue     = true;
			_PageIndex    = ColumnIndex >= ColumnsCount - 1 ? PageIndex + 1 : PageIndex;
			_ColumnIndex  = ColumnIndex >= ColumnsCount - 1 ? 0 : ColumnIndex + 1;
			_SectionIndex = ColumnIndex >= ColumnsCount - 1 ? 0 : SectionIndex;
			_bStart       = true;
			_StartIndex   = Count;
		}
		else
		{
			// console.log("Recalc time : " + ((performance.now() - this.FullRecalc.StartTime) / 1000));

			this.FullRecalc.Id           = null;
			this.FullRecalc.MainStartPos = -1;

			this.private_CheckUnusedFields();
			this.private_CheckCurPage();
			this.DrawingDocument.OnEndRecalculate(true);
			this.DrawingObjects.onEndRecalculateDocument(this.Pages.length);

			if (true === this.Selection.UpdateOnRecalc)
			{
				this.Selection.UpdateOnRecalc = false;
				this.DrawingDocument.OnSelectEnd();
			}

			// Основной пересчет окончен, если в колонтитулах есть элемент с количеством страниц, тогда нам надо
			// запустить дополнительный пересчет колонтитулов.
			// Если так случилось, что после повторного полного пересчета, вызванного изменением количества страниц и
			// изменением метрик колонтитула, новое количество страниц стало меньше, чем раньше, тогда мы не пересчитываем
			// дальше, чтобы избежать зацикливаний.

			if (-1 === this.HdrFtrRecalc.PageCount || this.HdrFtrRecalc.PageCount < this.Pages.length)
			{
				this.HdrFtrRecalc.PageCount = this.Pages.length;
				var nPageCountStartPage     = this.HdrFtr.HavePageCountElement();
				if (-1 !== nPageCountStartPage)
				{
					this.DrawingDocument.OnStartRecalculate(nPageCountStartPage);
					this.HdrFtrRecalc.PageIndex = nPageCountStartPage;
					this.private_RecalculateHdrFtrPageCountUpdate();
				}
			}
		}
    }

	if (this.NeedUpdateTracksOnRecalc)
	{
		this.private_UpdateTracks(this.NeedUpdateTracksParams.Selection, this.NeedUpdateTracksParams.EmptySelection);
	}

    if (true === bContinue)
    {
        this.FullRecalc.PageIndex         = _PageIndex;
        this.FullRecalc.SectionIndex      = _SectionIndex;
        this.FullRecalc.ColumnIndex       = _ColumnIndex;
        this.FullRecalc.StartIndex        = _StartIndex;
        this.FullRecalc.Start             = _bStart;
        this.FullRecalc.ResetStartElement = _bResetStartElement;
        this.FullRecalc.MainStartPos      = _StartIndex;
        this.FullRecalc.Endnotes          = _bEndnotesContinue;

        if (this.FullRecalc.UseRecursion)
		{
            if (window["NATIVE_EDITOR_ENJINE_SYNC_RECALC"] === true)
            {
				if (this.private_IsStartTimeoutOnRecalc(_PageIndex))
                {
                    if (window["native"] && window["native"]["WC_CheckSuspendRecalculate"] !== undefined)
					{
						//if (window["native"]["WC_CheckSuspendRecalculate"]())
						//    return;

						this.FullRecalc.Id             = setTimeout(Document_Recalculate_Page, 10);
						this.FullRecalc.TimerStartPage = _PageIndex;
						return;
					}
                }

                this.Recalculate_Page();
                return;
            }

			if (this.private_IsStartTimeoutOnRecalc(_PageIndex))
			{
				// console.log("Page " + _PageIndex);
				// console.log("Pages delta " + (_PageIndex - this.FullRecalc.TimerStartPage));
				this.FullRecalc.Id             = setTimeout(Document_Recalculate_Page, 20);
				this.FullRecalc.TimerStartPage = _PageIndex;
			}
			else
			{
				this.Recalculate_Page();
			}
		}
		else
		{
			this.FullRecalc.Continue = true;
		}
	}
	else
	{
		this.UpdatePlaceholders();
	}
};
CDocument.prototype.private_IsStartTimeoutOnRecalc = function(nPageAbs)
{
	let nTime = this.Layout.GetCalculateTimeLimit();
	return ((nPageAbs > this.FullRecalc.StartPage + this.FullRecalc.StartPagesCount
		&& (performance.now() - this.FullRecalc.TimerStartTime > nTime
			|| nPageAbs > this.FullRecalc.TimerStartPage + 50)));

	// if (nRes)
	// {
	// 	console.log("Page        " + nPageAbs);
	// 	console.log("Pages Delta " + (nPageAbs - this.FullRecalc.TimerStartPage));
	// }
	// return nRes;
};
CDocument.prototype.private_RecalculateIsNewSection = function(nPageAbs, nContentIndex)
{
	// Определим, является ли данная страница первой в новой секции
	var bNewSection = ( 0 === nPageAbs ? true : false );
	if (0 !== nPageAbs)
	{
		var PrevStartIndex = this.Pages[nPageAbs - 1].Pos;
		var CurSectInfo    = this.SectionsInfo.Get_SectPr(nContentIndex);
		var PrevSectInfo   = this.SectionsInfo.Get_SectPr(PrevStartIndex);

		if (PrevSectInfo !== CurSectInfo && (c_oAscSectionBreakType.Continuous !== CurSectInfo.SectPr.Get_Type() || true !== CurSectInfo.SectPr.Compare_PageSize(PrevSectInfo.SectPr) ))
			bNewSection = true;
	}

	return bNewSection;
};
CDocument.prototype.private_RecalculateShiftFootnotes = function(nPageAbs, nColumnAbs, dY, oSectPr)
{
	var nPosType = oSectPr.GetFootnotePos();

	// DocEnd, SectEnd ненужные константы по логике, но Word воспринимает их
	// именно как BeneathText, в то время как все остальное (даже константа не по формату)
	// воспринимает как PageBottom.
	if (Asc.c_oAscFootnotePos.BeneathText === nPosType || Asc.c_oAscFootnotePos.DocEnd === nPosType || Asc.c_oAscFootnotePos.SectEnd === nPosType)
	{
		this.Footnotes.Shift(nPageAbs, nColumnAbs, 0, dY);
	}
	else
	{
		var dFootnotesHeight = this.Footnotes.GetHeight(nPageAbs, nColumnAbs);
		var oPageMetrics     = this.Get_PageContentStartPos(nPageAbs);
		this.Footnotes.Shift(nPageAbs, nColumnAbs, 0, oPageMetrics.YLimit - dFootnotesHeight);
	}
};
CDocument.prototype.private_RecalculateFlowTable             = function(RecalcInfo)
{
    var Element            = RecalcInfo.Element;
    var X                  = RecalcInfo.X;
    var Y                  = RecalcInfo.Y;
    var XLimit             = RecalcInfo.XLimit;
    var YLimit             = RecalcInfo.YLimit;
    var PageIndex          = RecalcInfo.PageIndex;
    var ColumnIndex        = RecalcInfo.ColumnIndex;
    var Index              = RecalcInfo.Index;
    var StartIndex         = RecalcInfo.StartIndex;
    var ColumnsCount       = RecalcInfo.ColumnsCount;
    var bResetStartElement = RecalcInfo.ResetStartElement;
    var RecalcResult       = RecalcInfo.RecalcResult;

    // Когда у нас колонки мы не разбиваем плавающую таблицу на страницы.
    var isColumns = ColumnsCount > 1 ? true : false;
    if (true === isColumns)
        YLimit = 10000;

    if (true === this.RecalcInfo.Can_RecalcObject())
    {
        var ElementPageIndex = 0;
        if ((0 === Index && 0 === PageIndex) || Index !== StartIndex || (Index === StartIndex && true === bResetStartElement))
        {
        	// См. баги #42163 #55288
			let oPageSection = this.Pages[PageIndex].Sections[RecalcInfo.SectionIndex];
			if (ColumnIndex > 0 && oPageSection.IsCalculatingSectionBottomLine())
				Y = oPageSection.Y;

			Element.Set_DocumentIndex(Index);
            Element.Reset(X, Y, XLimit, YLimit, PageIndex, ColumnIndex, ColumnsCount);
            ElementPageIndex = 0;
        }
        else
        {
            ElementPageIndex = PageIndex - Element.PageNum;
        }

        var TempRecalcResult = Element.Recalculate_Page(ElementPageIndex);
        this.RecalcInfo.Set_FlowObject(Element, ElementPageIndex, TempRecalcResult, -1, {
            X      : X,
            Y      : Y,
            XLimit : XLimit,
            YLimit : YLimit
        });

        if (((0 === Index && 0 === PageIndex) || Index != StartIndex) && true != Element.IsContentOnFirstPage() && true !== isColumns)
        {
            this.RecalcInfo.Set_PageBreakBefore(true);
            RecalcResult = recalcresult_NextPage | recalcresultflags_LastFromNewPage;
        }
        else
        {
            this.Pages[PageIndex].AddFlowTable(Element);
            this.DrawingObjects.addFloatTable(new CFlowTable(Element, PageIndex));
            RecalcResult = recalcresult_CurPage;
        }
    }
    else if (true === this.RecalcInfo.Check_FlowObject(Element))
    {
        if (Element.PageNum === PageIndex)
        {
            if (true === this.RecalcInfo.Is_PageBreakBefore())
            {
                this.RecalcInfo.Reset();
                RecalcResult = recalcresult_NextPage | recalcresultflags_LastFromNewPage;
            }
            else
            {
                X      = this.RecalcInfo.AdditionalInfo.X;
                Y      = this.RecalcInfo.AdditionalInfo.Y;
                XLimit = this.RecalcInfo.AdditionalInfo.XLimit;
                YLimit = this.RecalcInfo.AdditionalInfo.YLimit;

                // Пересчет нужнен для обновления номеров колонок и страниц
                Element.Reset(X, Y, XLimit, YLimit, PageIndex, ColumnIndex, ColumnsCount, this.Pages[PageIndex].Sections[this.Pages[PageIndex].Sections.length - 1].Y);
                RecalcResult = Element.Recalculate_Page(0);
                this.RecalcInfo.FlowObjectPage++;

                if (true === isColumns)
                    RecalcResult = recalcresult_NextElement;

                if (RecalcResult & recalcresult_NextElement)
                    this.RecalcInfo.Reset();
            }
        }
        else if (Element.PageNum > PageIndex || (this.RecalcInfo.FlowObjectPage <= 0 && Element.PageNum < PageIndex))
        {
        	this.Pages[PageIndex - 1].RemoveFlowTable(Element);
            this.DrawingObjects.removeFloatTableById(PageIndex - 1, Element.Get_Id());
            this.RecalcInfo.Set_PageBreakBefore(true);
            RecalcResult = recalcresult_PrevPage;
        }
        else
        {
            RecalcResult = Element.Recalculate_Page(PageIndex - Element.PageNum);
            this.RecalcInfo.FlowObjectPage++;
			this.Pages[PageIndex].AddFlowTable(Element);
            this.DrawingObjects.addFloatTable(new CFlowTable(Element, PageIndex));

            if (RecalcResult & recalcresult_NextElement)
                this.RecalcInfo.Reset();
        }
    }
    else
    {
        RecalcResult = recalcresult_NextElement;
    }

    RecalcInfo.Index        = Index;
    RecalcInfo.RecalcResult = RecalcResult;
};
CDocument.prototype.private_RecalculateFlowParagraph         = function(RecalcInfo)
{
    var Element            = RecalcInfo.Element;
    var X                  = RecalcInfo.X;
    var Y                  = RecalcInfo.Y;
    var XLimit             = RecalcInfo.XLimit;
    var YLimit             = RecalcInfo.YLimit;
    var PageIndex          = RecalcInfo.PageIndex;
    var ColumnIndex        = RecalcInfo.ColumnIndex;
    var Index              = RecalcInfo.Index;
    var StartIndex         = RecalcInfo.StartIndex;
    var ColumnsCount       = RecalcInfo.ColumnsCount;
    var bResetStartElement = RecalcInfo.ResetStartElement;
    var RecalcResult       = RecalcInfo.RecalcResult;

    var Page = this.Pages[PageIndex];

    if (true === this.RecalcInfo.Can_RecalcObject())
    {
        var FramePr = Element.GetFramePr();

        var FlowCount = this.CountElementsInFrame(Index);

        var Page_W = Page.Width;
        var Page_H = Page.Height;

        var Page_Field_L = Page.Margins.Left;
        var Page_Field_R = Page.Margins.Right;
        var Page_Field_T = Page.Margins.Top;
        var Page_Field_B = Page.Margins.Bottom;

        var Column_Field_L = X;
        var Column_Field_R = XLimit;

        //--------------------------------------------------------------------------------------------------
        // 1. Рассчитаем размер рамки
        //--------------------------------------------------------------------------------------------------
        var FrameH = 0;
        var FrameW = -1;

        var Frame_XLimit = FramePr.Get_W();
        var Frame_YLimit = FramePr.Get_H();

		var FrameHRule = (undefined === FramePr.HRule ? (undefined !== Frame_YLimit ? Asc.linerule_AtLeast : Asc.linerule_Auto) : FramePr.HRule);

        if (undefined === Frame_XLimit || 0 === AscCommon.MMToTwips(Frame_XLimit))
            Frame_XLimit = Page_Field_R - Page_Field_L;

        if (undefined === Frame_YLimit || 0 === AscCommon.MMToTwips(Frame_YLimit) || Asc.linerule_Auto === FrameHRule)
            Frame_YLimit = Page_H;

        for (var TempIndex = Index; TempIndex < Index + FlowCount; ++TempIndex)
        {
            var TempElement = this.Content[TempIndex];
            TempElement.Set_DocumentIndex(TempIndex);

			var ElementPageIndex = 0;
			if ((0 === TempIndex && 0 === PageIndex) || TempIndex != StartIndex || (TempIndex === StartIndex && true === bResetStartElement))
			{
				TempElement.Reset(0, FrameH, Frame_XLimit, Asc.NoYLimit, PageIndex, ColumnIndex, ColumnsCount);
			}
			else
			{
				ElementPageIndex = PageIndex - Element.PageNum;
			}

            RecalcResult = TempElement.Recalculate_Page(ElementPageIndex);

            if (!(RecalcResult & recalcresult_NextElement))
                break;

            FrameH = TempElement.Get_PageBounds(0).Bottom;
        }

		var oTableInfo = Element.GetMaxTableGridWidth();

		var nMaxGapLeft           = oTableInfo.GapLeft;
		var nMaxGridWidth         = oTableInfo.GridWidth;
		var nMaxGridWidthRightGap = oTableInfo.GridWidth + oTableInfo.GapRight;

		for (var nTempIndex = Index + 1; nTempIndex < Index + FlowCount; ++nTempIndex)
		{
			var oTempTableInfo = this.Content[nTempIndex].GetMaxTableGridWidth();

			if (oTempTableInfo.GapLeft > nMaxGapLeft)
				nMaxGapLeft = oTempTableInfo.GapLeft;

			if (oTempTableInfo.GridWidth > nMaxGridWidth)
				nMaxGridWidth = oTempTableInfo.GridWidth;

			if (oTempTableInfo.GridWidth + oTempTableInfo.GapRight > nMaxGridWidthRightGap)
				nMaxGridWidthRightGap = oTempTableInfo.GridWidth + oTempTableInfo.GapRight;
		}

        if (Element.IsParagraph() && -1 === FrameW && 1 === FlowCount && 1 === Element.Lines.length && (undefined === FramePr.Get_W() || 0 === AscCommon.MMToTwips(FramePr.Get_W())))
        {
			FrameW     = Element.GetAutoWidthForDropCap();
			var ParaPr = Element.Get_CompiledPr2(false).ParaPr;
			FrameW += ParaPr.Ind.Left + ParaPr.Ind.FirstLine;

            // Если прилегание в данном случае не к левой стороне, тогда пересчитываем параграф,
            // с учетом того, что ширина буквицы должна быть FrameW
            if (align_Left != ParaPr.Jc)
            {
                Element.Reset(0, 0, FrameW, Frame_YLimit, PageIndex, ColumnIndex, ColumnsCount);
                Element.Recalculate_Page(0);
                FrameH = TempElement.Get_PageBounds(0).Bottom;
            }
        }
        else if (-1 === FrameW)
        {
			if (Element.IsTable() && (!FramePr.Get_W() || 0 === AscCommon.MMToTwips(FramePr.Get_W())))
			{
				FrameW = nMaxGridWidth;
			}
			else
			{
				FrameW = Frame_XLimit;
			}
        }

		var nGapLeft  = nMaxGapLeft;
		var nGapRight = nMaxGridWidthRightGap > FrameW && nMaxGridWidth < FrameW + 0.001 ? nMaxGridWidthRightGap - FrameW : 0;

		if (0 !== AscCommon.MMToTwips(FramePr.H) && ((Asc.linerule_AtLeast === FrameHRule && FrameH < FramePr.H) || Asc.linerule_Exact === FrameHRule))
        {
            FrameH = FramePr.H;
        }
        //--------------------------------------------------------------------------------------------------
        // 2. Рассчитаем положение рамки
        //--------------------------------------------------------------------------------------------------

        // Теперь зная размеры рамки можем рассчитать ее позицию
        var FrameHAnchor = ( FramePr.HAnchor === undefined ? c_oAscHAnchor.Margin : FramePr.HAnchor );
        var FrameVAnchor = ( FramePr.VAnchor === undefined ? c_oAscVAnchor.Text : FramePr.VAnchor );

        // Рассчитаем положение по горизонтали
        var FrameX = 0;
        if (undefined !== FramePr.XAlign || undefined === FramePr.X)
        {
            var XAlign = c_oAscXAlign.Left;
            if (undefined !== FramePr.XAlign)
                XAlign = FramePr.XAlign;

            switch (FrameHAnchor)
            {
                case c_oAscHAnchor.Page   :
                {
                    switch (XAlign)
                    {
                        case c_oAscXAlign.Inside  :
                        case c_oAscXAlign.Outside :
                        case c_oAscXAlign.Left    :
                            FrameX = Page_Field_L - FrameW;
                            break;
                        case c_oAscXAlign.Right   :
                            FrameX = Page_Field_R;
                            break;
                        case c_oAscXAlign.Center  :
                            FrameX = (Page_W - FrameW) / 2;
                            break;
                    }

                    break;
                }
                case c_oAscHAnchor.Text   :
                case c_oAscHAnchor.Margin :
                {
                    switch (XAlign)
                    {
                        case c_oAscXAlign.Inside  :
                        case c_oAscXAlign.Outside :
                        case c_oAscXAlign.Left    :
                            FrameX = Column_Field_L;
                            break;
                        case c_oAscXAlign.Right   :
                            FrameX = Column_Field_R - FrameW;
                            break;
                        case c_oAscXAlign.Center  :
                            FrameX = (Column_Field_R + Column_Field_L - FrameW) / 2;
                            break;
                    }

                    break;
                }
            }

        }
        else
        {
            switch (FrameHAnchor)
            {
                case c_oAscHAnchor.Page   :
                    FrameX = FramePr.X;
                    break;
                case c_oAscHAnchor.Text   :
                case c_oAscHAnchor.Margin :
                    FrameX = Page_Field_L + FramePr.X;
                    break;
            }
        }

        if (FrameX < 0)
            FrameX = 0;

        // Рассчитаем положение по вертикали
        var FrameY = 0;
        if (undefined !== FramePr.YAlign)
        {
            var YAlign = FramePr.YAlign;
            // Случай c_oAscYAlign.Inline не обрабатывается, потому что такие параграфы считаются Inline

            switch (FrameVAnchor)
            {
                case c_oAscVAnchor.Page   :
                {
                    switch (YAlign)
                    {
                        case c_oAscYAlign.Inside  :
                        case c_oAscYAlign.Outside :
                        case c_oAscYAlign.Top     :
                            FrameY = 0;
                            break;
                        case c_oAscYAlign.Bottom  :
                            FrameY = Page_H - FrameH;
                            break;
                        case c_oAscYAlign.Center  :
                            FrameY = (Page_H - FrameH) / 2;
                            break;
                    }

                    break;
                }
                case c_oAscVAnchor.Text   :
                {
                    FrameY = Y;
                    break;
                }
                case c_oAscVAnchor.Margin :
                {
                    switch (YAlign)
                    {
                        case c_oAscYAlign.Inside  :
                        case c_oAscYAlign.Outside :
                        case c_oAscYAlign.Top     :
                            FrameY = Page_Field_T;
                            break;
                        case c_oAscYAlign.Bottom  :
                            FrameY = Page_Field_B - FrameH;
                            break;
                        case c_oAscYAlign.Center  :
                            FrameY = (Page_Field_B + Page_Field_T - FrameH) / 2;
                            break;
                    }

                    break;
                }
            }
        }
        else
        {
            var FramePrY = 0;
            if (undefined !== FramePr.Y)
                FramePrY = FramePr.Y;

            switch (FrameVAnchor)
            {
				case c_oAscVAnchor.Page   :
					FrameY = FramePrY;
					break;
				case c_oAscVAnchor.Text   :
					FrameY = FramePrY + Y;
					break;
				case c_oAscVAnchor.Margin :

					// Если Y не задано, либо ровно 0, тогда MSWord считает это как привязка к тексту (баг 52903)
					if (undefined === FramePr.Y || 0 === AscCommon.MMToTwips(FramePr.Y))
						FrameY = Y + 0.001; // Погрешность, чтобы не сдвигалась предыдущая строка из-за обтекания
					else
						FrameY = FramePrY + Page_Field_T;

					break;
            }
        }

        if (FrameH + FrameY > Page_H)
            FrameY = Page_H - FrameH;

        // TODO: Пересмотреть, почему эти погрешности возникают
        // Избавляемся от погрешности
        FrameY += 0.001;
        FrameH -= 0.002;

        if (FrameY < 0)
            FrameY = 0;

        var FrameBounds;
        if (this.Content[Index].IsParagraph())
        	FrameBounds = this.Content[Index].Get_FrameBounds(FrameX, FrameY, FrameW, FrameH);
        else
        	FrameBounds = this.Content[Index].GetFirstParagraph().Get_FrameBounds(FrameX, FrameY, FrameW, FrameH);

        var FrameX2 = FrameBounds.X - nGapLeft,
			FrameY2 = FrameBounds.Y,
			FrameW2 = FrameBounds.W + nGapLeft + nGapRight,
			FrameH2 = FrameBounds.H;

        if (!(RecalcResult & recalcresult_NextElement))
        {
            // Ничего не делаем здесь, пересчитываем текущую страницу заново, либо предыдущую

            if (RecalcResult & recalcresult_PrevPage)
                this.RecalcInfo.Set_FrameRecalc(false);

            // TODO: Если мы заново пересчитываем текущую страницу, проверить надо ли обнулять параметр RecalcInfo.FrameRecalc
        }
        else if ((FrameY2 + FrameH2 > Page_H || Y > Page_H - 0.001) && Index != StartIndex)
        {
            this.RecalcInfo.Set_FrameRecalc(true);
            RecalcResult = recalcresult_NextPage | recalcresultflags_LastFromNewColumn;
        }
        else
        {
        	var oCalculatedFrame = new CCalculatedFrame(FramePr, FrameX, FrameY, FrameW, FrameH, FrameX2, FrameY2, FrameW2, FrameH2, PageIndex, Index, FlowCount);

            this.RecalcInfo.Set_FrameRecalc(false);
            for (var TempIndex = Index; TempIndex < Index + FlowCount; ++TempIndex)
            {
                var TempElement = this.Content[TempIndex];
                TempElement.Shift(TempElement.GetPagesCount() - 1, FrameX, FrameY);
                TempElement.SetCalculatedFrame(oCalculatedFrame);
            }

            var FrameDx = ( undefined === FramePr.HSpace ? 0 : FramePr.HSpace );
            var FrameDy = ( undefined === FramePr.VSpace ? 0 : FramePr.VSpace );

            this.Pages[PageIndex].AddFrame(oCalculatedFrame);
            this.DrawingObjects.addFloatTable(new CFlowParagraph(Element, FrameX2, FrameY2, FrameW2, FrameH2, FrameDx, FrameDy, Index, FlowCount, FramePr.Wrap));

            Index += FlowCount - 1;

            if (FrameY >= Y && FrameX >= X - 0.001)
            {
                RecalcResult = recalcresult_NextElement;
            }
            else
            {
                this.RecalcInfo.Set_FlowObject(Element, PageIndex, recalcresult_NextElement, FlowCount);
                RecalcResult = recalcresult_CurPage | recalcresultflags_Page;
            }
        }
    }
    else if (true === this.RecalcInfo.Check_FlowObject(Element) && true === this.RecalcInfo.Is_PageBreakBefore())
    {
        this.RecalcInfo.Reset();
        this.RecalcInfo.Set_FrameRecalc(true);
        RecalcResult = recalcresult_NextPage | recalcresultflags_LastFromNewPage;
    }
    else if (true === this.RecalcInfo.Check_FlowObject(Element))
    {
        if (this.RecalcInfo.FlowObjectPage !== PageIndex)
        {
            // Номер страницы не такой же (должен быть +1), значит нам надо заново персесчитать предыдущую страницу
            // с условием, что данная рамка начнется с новой страницы
            this.RecalcInfo.Set_PageBreakBefore(true);
            this.Pages[this.RecalcInfo.FlowObjectPage].RemoveFrame(Index);
            this.DrawingObjects.removeFloatTableById(this.RecalcInfo.FlowObjectPage, Element.Get_Id());
            RecalcResult = recalcresult_PrevPage | recalcresultflags_Page;
        }
        else
        {
            // Все нормально рассчиталось

            // В случае колонок может так случится, что логическое место окажется в другой колонке, поэтому мы делаем
            // Reset для обновления логической позиции, но физическую позицию не меняем.
            var FlowCount = this.RecalcInfo.FlowObjectElementsCount;
            for (var TempIndex = Index; TempIndex < Index + FlowCount; ++TempIndex)
            {
                var TempElement = this.Content[TempIndex];
                TempElement.Reset(TempElement.X, TempElement.Y, TempElement.XLimit, TempElement.YLimit, TempElement.PageNum, ColumnIndex, ColumnsCount);
            }


            Index += this.RecalcInfo.FlowObjectElementsCount - 1;
            this.RecalcInfo.Reset();
            RecalcResult = recalcresult_NextElement;
        }
    }
    else
    {
        // В случае колонок может так случится, что логическое место окажется в другой колонке, поэтому мы делаем
        // Reset для обновления логической позиции, но физическую позицию не меняем.
        var FlowCount = this.CountElementsInFrame(Index);
        for (var TempIndex = Index; TempIndex < Index + FlowCount; ++TempIndex)
        {
            var TempElement = this.Content[TempIndex];
            TempElement.Reset(TempElement.X, TempElement.Y, TempElement.XLimit, TempElement.YLimit, PageIndex, ColumnIndex, ColumnsCount);
        }

        RecalcResult = recalcresult_NextElement;
    }

    RecalcInfo.Index        = Index;
    RecalcInfo.RecalcResult = RecalcResult;
};
CDocument.prototype.private_RecalculateHdrFtrPageCountUpdate = function()
{
	this.HdrFtrRecalc.Id = null;

	var nPageAbs    = this.HdrFtrRecalc.PageIndex;
	var nPagesCount = this.Pages.length;

	while (nPageAbs < nPagesCount)
	{
		var Result = this.HdrFtr.RecalculatePageCountUpdate(nPageAbs, nPagesCount);
		if (null === Result)
		{
			nPageAbs++;
		}
		else if (false === Result)
		{
			this.DrawingDocument.OnRecalculatePage(nPageAbs, this.Pages[nPageAbs]);

			if (window["NATIVE_EDITOR_ENJINE_SYNC_RECALC"] === true)
			{
				if (nPageAbs >= this.HdrFtrRecalc.PageIndex + 5 && window["native"] && window["native"]["WC_CheckSuspendRecalculate"] !== undefined)
				{
					//if (window["native"]["WC_CheckSuspendRecalculate"]())
					//    return;1

					this.HdrFtrRecalc.PageIndex = nPageAbs + 1;
					this.HdrFtrRecalc.Id        = setTimeout(Document_Recalculate_HdrFtrPageCount, 20);
					return;
				}

				nPageAbs++;
			}
			else
			{
				if (nPageAbs < this.HdrFtrRecalc.PageIndex + 5)
				{
					nPageAbs++;
				}
				else
				{
					this.HdrFtrRecalc.PageIndex = nPageAbs + 1;
					this.HdrFtrRecalc.Id        = setTimeout(Document_Recalculate_HdrFtrPageCount, 20);
					return;
				}
			}
		}
		else
		{
			for (var nCurPage = nPageAbs + 1; nCurPage < nPagesCount; ++nCurPage)
			{
				this.HdrFtr.UpdatePagesCount(nCurPage, nPagesCount);
			}

			this.RecalcInfo.Reset();
			this.FullRecalc.PageIndex         = nPageAbs;
			this.FullRecalc.SectionIndex      = 0;
			this.FullRecalc.ColumnIndex       = 0;
			this.FullRecalc.StartIndex        = this.Pages[nPageAbs].Pos;
			this.FullRecalc.Start             = true;
			this.FullRecalc.StartPage         = nPageAbs;
			this.FullRecalc.ResetStartElement = this.private_RecalculateIsNewSection(nPageAbs, this.Pages[nPageAbs].Pos);
			this.FullRecalc.Endnotes          = this.Endnotes.IsContinueRecalculateFromPrevPage(nPageAbs);
			this.FullRecalc.MainStartPos      = this.Pages[nPageAbs].Pos;

			this.DrawingDocument.OnStartRecalculate(nPageAbs);
			this.Recalculate_Page();
			return;
		}
	}

	if (nPageAbs >= nPagesCount)
		this.DrawingDocument.OnEndRecalculate(false, true);
};
CDocument.prototype.private_CheckUnusedFields = function()
{
	if (this.Content.length <= 0)
		return;

	var oLastPara = this.Content[this.Content.length - 1];
	if (type_Paragraph !== oLastPara.GetType())
		return;

	var oInfo = oLastPara.GetEndInfoByPage(oLastPara.GetPagesCount() - 1);
	for (var nIndex = 0, nCount = oInfo.ComplexFields.length; nIndex < nCount; ++nIndex)
	{
		var oComplexField = oInfo.ComplexFields[nIndex].ComplexField;
		var oBeginChar    = oComplexField.GetBeginChar();
		var oEndChar      = oComplexField.GetEndChar();
		var oSeparateChar = oComplexField.GetSeparateChar();

		if (oBeginChar)
			oBeginChar.SetUse(false);

		if (oEndChar)
			oEndChar.SetUse(false);

		if (oSeparateChar)
			oSeparateChar.SetUse(false);
	}
};
CDocument.prototype.private_ClearSearchOnRecalculate = function()
{
	if (!this.SearchEngine.ClearOnRecalc)
		return;

	this.ClearSearch();
};
CDocument.prototype.IsCalculatingContinuousSectionBottomLine = function()
{
	var nPageIndex    = this.FullRecalc.PageIndex;
	var nSectionIndex = this.FullRecalc.SectionIndex;

	if (!this.Pages[nPageIndex] || !this.Pages[nPageIndex].Sections[nSectionIndex])
		return false;

	var oPageSection = this.Pages[nPageIndex].Sections[nSectionIndex];
	return oPageSection.IsCalculatingSectionBottomLine();
};
/**
 * Получаем номер рассчитанной страницы, с которой начинается заданный элемент
 * @param nElementPos
 * @returns {number}
 */
CDocument.prototype.private_GetPageByPos = function(nElementPos)
{
	var nResultPage = -1;
	for (var nCurPage = 0, nPagesCount = this.Pages.length; nCurPage < nPagesCount; ++nCurPage)
	{
		if (nElementPos > this.Pages[nCurPage].Pos)
			nResultPage = nCurPage;
		else
			break;
	}

	return nResultPage;
};
CDocument.prototype.OnColumnBreak_WhileRecalculate           = function()
{
    var PageIndex    = this.FullRecalc.PageIndex;
    var SectionIndex = this.FullRecalc.SectionIndex;

    if (this.Pages[PageIndex] && this.Pages[PageIndex].Sections[SectionIndex])
        this.Pages[PageIndex].Sections[SectionIndex].ForbidRecalculateBottomLine();
};
CDocument.prototype.Reset_RecalculateCache                   = function()
{
    this.SectionsInfo.Reset_HdrFtrRecalculateCache();

    var Count = this.Content.length;
    for (var Index = 0; Index < Count; Index++)
    {
        this.Content[Index].Reset_RecalculateCache();
    }

    this.Footnotes.ResetRecalculateCache();
};
/**
 * Получаем идентификатор текущего пересчета
 * @returns {number}
 */
CDocument.prototype.GetRecalcId = function()
{
	return this.RecalcId;
};
/**
 * Останавливаем процесс пересчета (если пересчет был запущен и он долгий)
 */
CDocument.prototype.StopRecalculate = function()
{
	if (this.FullRecalc.Id)
	{
		clearTimeout(this.FullRecalc.Id);
		this.FullRecalc.Id = null;
		this.DrawingDocument.OnEndRecalculate(false);
		this.RecalcInfo.Set_NeedRecalculateFromStart(true);
	}
};
/**
 * Временно приостанавливаем долгий пересчет
 */
CDocument.prototype.PauseRecalculate = function()
{
	if (this.FullRecalc.Id)
	{
		clearTimeout(this.FullRecalc.Id);
		this.FullRecalc.Id     = null;
		this.RecalcInfo.Paused = true;
	}
};
/**
 * Возобнавляем долгий пересчет
 */
CDocument.prototype.ResumeRecalculate = function()
{
	if (this.RecalcInfo.Paused)
	{
		this.FullRecalc.Id = setTimeout(Document_Recalculate_Page, 10);
		this.RecalcInfo.Paused = false;
	}
};
CDocument.prototype.OnContentReDraw                          = function(StartPage, EndPage)
{
    this.ReDraw(StartPage, EndPage);
};
CDocument.prototype.CheckTargetUpdate = function()
{
	// Проверим можно ли вообще пересчитывать текущее положение.
	if (this.DrawingDocument.UpdateTargetFromPaint === true)
	{
		if (true === this.DrawingDocument.UpdateTargetCheck)
			this.NeedUpdateTarget = this.DrawingDocument.UpdateTargetCheck;
		this.DrawingDocument.UpdateTargetCheck = false;
	}

	if (!this.NeedUpdateTarget)
		return;

	if (this.ViewPosition)
	{
		this.CheckViewPosition();
	}
	else
	{
		if (this.Controller.CanUpdateTarget() && !this.IsMovingTableBorder())
		{
			// Обновляем курсор сначала, чтобы обновить текущую страницу
			this.RecalculateCurPos();
			this.NeedUpdateTarget = false;
		}
	}
};
CDocument.prototype.CheckViewPosition = function()
{
	if (!this.ViewPosition || !this.ViewPosition.AnchorPos)
	{
		this.RecalculateCurPos();
		return;
	}
	
	let anchorPos = this.ViewPosition.AnchorPos;
	let alignTop  = this.ViewPosition.AlignTop;
	let distance  = this.ViewPosition.Distance;
	
	if (!anchorPos[0] || this !== anchorPos[0].Class)
	{
		this.RecalculateCurPos();
		
		this.ViewPosition     = null;
		this.NeedUpdateTarget = false;
		return;
	}
	
	let nInDocumentPosition = anchorPos[0].Position;
	if (this.FullRecalc.Id && this.FullRecalc.StartIndex <= nInDocumentPosition)
		return;
	
	this.ViewPosition     = null;
	this.NeedUpdateTarget = false;
	
	function GetXY(docPos)
	{
		let run = docPos[docPos.length - 1].Class;
		if (!run || !(run instanceof AscWord.CRun))
			return {Page : 0, Y : 0, X : 0, H : 0};
		
		let paragraph = run.GetParagraph();
		
		let state = paragraph.SaveSelectionState();
		paragraph.RemoveSelection();
		
		run.SetThisElementCurrentInParagraph();
		run.State.ContentPos = docPos[docPos.length - 1].Position;
		
		let posInfo = paragraph.RecalculateCurPos(false, false, false, true);
		paragraph.LoadSelectionState(state);
		
		return {
			Page : posInfo.PageNum,
			X    : 0,
			Y    : posInfo.Y,
			H    : posInfo.Height
		}
	}
	
	let anchor = GetXY(anchorPos);
	if (alignTop)
		this.DrawingDocument.m_oWordControl.ScrollToAbsolutePosition(anchor.X, anchor.Y - distance, anchor.Page);
	else
		this.DrawingDocument.m_oWordControl.ScrollToAbsolutePosition(anchor.X, anchor.Y + distance, anchor.Page, true);
	
	this.RecalculateCurPos();
};
CDocument.prototype.RecalculateCurPos = function()
{
	if (true === this.TurnOffRecalcCurPos)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	this.DrawingDocument.UpdateTargetTransform(null);

	this.Controller.RecalculateCurPos();

	// TODO: Здесь добавлено обновление линейки, чтобы обновлялись границы рамки при наборе текста, а также чтобы
	//       обновлялись поля колонтитулов при наборе текста.
	this.Document_UpdateRulersState();
};
CDocument.prototype.private_CheckCurPage = function()
{
	if (true === this.TurnOffRecalcCurPos)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	var nCurPage = this.Controller.GetCurPage();
	if (-1 !== nCurPage)
		this.CurPage = nCurPage;
};
CDocument.prototype.Set_TargetPos = function(X, Y, PageNum)
{
	this.TargetPos.X       = X;
	this.TargetPos.Y       = Y;
	this.TargetPos.PageNum = PageNum;
};
/**
 * Запрос на перерисовку заданного отрезка страниц.
 * @param StartPage
 * @param EndPage
 */
CDocument.prototype.ReDraw                                   = function(StartPage, EndPage)
{
    if ("undefined" === typeof(StartPage))
        StartPage = 0;
    if ("undefined" === typeof(EndPage))
        EndPage = this.DrawingDocument.m_lCountCalculatePages;

    for (var CurPage = StartPage; CurPage <= EndPage; CurPage++)
        this.DrawingDocument.OnRepaintPage(CurPage);
};
CDocument.prototype.DrawPage                                 = function(nPageIndex, pGraphics)
{
    this.Draw(nPageIndex, pGraphics);
};
CDocument.prototype.CanDrawPage = function(nPageAbs)
{
	if (null !== this.FullRecalc.Id && nPageAbs >= this.FullRecalc.PageIndex - 1)
		return false;

	return true;
};
/**
 * Перерисовка заданной страницы документа.
 * @param nPageIndex
 * @param pGraphics
 */
CDocument.prototype.Draw                                     = function(nPageIndex, pGraphics)
{
    // TODO: Пока делаем обновление курсоров при каждой отрисовке. Необходимо поменять
    this.CollaborativeEditing.Update_ForeignCursorsPositions();
    //--------------------------------------------------------------------------------------------------------------

    this.Comments.Reset_Drawing(nPageIndex);

    // Определим секцию
    var Page_StartPos = this.Pages[nPageIndex].Pos;
    var SectPr        = this.SectionsInfo.Get_SectPr(Page_StartPos).SectPr;

	if (docpostype_HdrFtr !== this.CurPos.Type && !this.IsViewMode())
		pGraphics.Start_GlobalAlpha();

	// Рисуем границы вокруг страницы (если границы надо рисовать под текстом)
	if (section_borders_ZOrderBack === SectPr.Get_Borders_ZOrder())
		this.DrawPageBorders(pGraphics, SectPr, nPageIndex);

	this.HdrFtr.Draw(nPageIndex, pGraphics);

	// Рисуем содержимое документа на данной странице
	if (docpostype_HdrFtr === this.CurPos.Type)
		pGraphics.put_GlobalAlpha(true, 0.4);
	else if (!this.IsViewMode())
		pGraphics.End_GlobalAlpha();

    this.DrawingObjects.drawBehindDoc(nPageIndex, pGraphics);

    this.Footnotes.Draw(nPageIndex, pGraphics);

    var Page = this.Pages[nPageIndex];
    for (var SectionIndex = 0, SectionsCount = Page.Sections.length; SectionIndex < SectionsCount; ++SectionIndex)
    {
        var PageSection = Page.Sections[SectionIndex];

        this.Endnotes.Draw(nPageIndex, PageSection.Index, pGraphics);

        for (var ColumnIndex = 0, ColumnsCount = PageSection.Columns.length; ColumnIndex < ColumnsCount; ++ColumnIndex)
        {
            var Column         = PageSection.Columns[ColumnIndex];
            var ColumnStartPos = Column.Pos;
            var ColumnEndPos   = Column.EndPos;

            if (true === PageSection.ColumnsSep && ColumnIndex > 0 && !Column.IsEmpty())
			{

				var SepX = (Column.X + PageSection.Columns[ColumnIndex - 1].XLimit) / 2;
				pGraphics.p_color(0, 0, 0, 255);
				pGraphics.drawVerLine(c_oAscLineDrawingRule.Left, SepX, PageSection.Y, PageSection.YLimit, 0.75 * g_dKoef_pt_to_mm);
			}

            if (ColumnsCount > 1)
            {
                pGraphics.SaveGrState();

                var X    = ColumnIndex === 0 ? 0 : Column.X - Column.SpaceBefore / 2;
                var XEnd = (ColumnIndex >= ColumnsCount - 1 ? Page.Width : Column.XLimit + Column.SpaceAfter / 2);
                pGraphics.AddClipRect(X, 0, XEnd - X, Page.Height);
            }

            for (var ContentPos = ColumnStartPos; ContentPos <= ColumnEndPos; ++ContentPos)
            {
            	var oElement = this.Content[ContentPos];
            	if (Page.IsFlowTable(oElement) || Page.IsFrame(oElement))
            		continue;

				var ElementPageIndex = this.private_GetElementPageIndex(ContentPos, nPageIndex, ColumnIndex, ColumnsCount);
				this.Content[ContentPos].Draw(ElementPageIndex, pGraphics);
            }

            if (ColumnsCount > 1)
            {
                pGraphics.RestoreGrState();
            }

			for (var nFlowTableIndex = 0, nFlowTablesCount = Page.FlowTables.length; nFlowTableIndex < nFlowTablesCount; ++nFlowTableIndex)
			{
				var oTable      = Page.FlowTables[nFlowTableIndex];
				var nTableIndex = oTable.GetIndex();

				if (ColumnStartPos <= nTableIndex && nTableIndex <= ColumnEndPos)
				{
					var nElementPageIndex = this.private_GetElementPageIndex(oTable.GetIndex(), nPageIndex, ColumnIndex, ColumnsCount);
					oTable.Draw(nElementPageIndex, pGraphics);
				}
			}

			var nPixelError = this.GetDrawingDocument().GetMMPerDot(1);
			for (var nFrameIndex = 0, nFramesCount = Page.Frames.length; nFrameIndex < nFramesCount; ++nFrameIndex)
			{
				var oFrame = Page.Frames[nFrameIndex];

				if (ColumnStartPos <= oFrame.StartIndex && oFrame.StartIndex <= ColumnEndPos)
				{
					var nL = oFrame.L2 - nPixelError;
					var nT = oFrame.T2 - nPixelError;
					var nH = oFrame.H2 + 2 * nPixelError;
					var nW = oFrame.W2 + 2 * nPixelError;

					pGraphics.SaveGrState();
					pGraphics.AddClipRect(nL, nT, nW, nH);

					for (var nFlowIndex = oFrame.StartIndex; nFlowIndex < oFrame.StartIndex + oFrame.FlowCount; ++nFlowIndex)
					{
						var nElementPageIndex = this.private_GetElementPageIndex(nFlowIndex, nPageIndex, ColumnIndex, ColumnsCount);
						this.Content[nFlowIndex].Draw(nElementPageIndex, pGraphics);
					}

					pGraphics.RestoreGrState();
				}
			}
        }
    }

    this.DrawingObjects.drawBeforeObjects(nPageIndex, pGraphics);

    // Рисуем границы вокруг страницы (если границы надо рисовать перед текстом)
    if (section_borders_ZOrderFront === SectPr.Get_Borders_ZOrder())
        this.DrawPageBorders(pGraphics, SectPr, nPageIndex);

    if (docpostype_HdrFtr === this.CurPos.Type)
    {
        pGraphics.put_GlobalAlpha(false, 1.0);

        // Рисуем колонтитулы
        var SectIndex = this.SectionsInfo.Get_Index(Page_StartPos);
        var SectCount = this.SectionsInfo.Get_Count();

        var SectIndex = ( 1 === SectCount ? -1 : SectIndex );

        var Header = this.HdrFtr.Pages[nPageIndex].Header;
        var Footer = this.HdrFtr.Pages[nPageIndex].Footer;

        var RepH = ( null === Header || null !== SectPr.GetHdrFtrInfo(Header) ? false : true );
        var RepF = ( null === Footer || null !== SectPr.GetHdrFtrInfo(Footer) ? false : true );

        var HeaderInfo = undefined;
        if (null !== Header && undefined !== Header.RecalcInfo.PageNumInfo[nPageIndex])
        {
            var bFirst = Header.RecalcInfo.PageNumInfo[nPageIndex].bFirst;
            var bEven  = Header.RecalcInfo.PageNumInfo[nPageIndex].bEven;

            var HeaderSectPr = Header.RecalcInfo.SectPr[nPageIndex];

            if (undefined !== HeaderSectPr)
                bFirst = ( true === bFirst && true === HeaderSectPr.Get_TitlePage() ? true : false );

            HeaderInfo = {bFirst : bFirst, bEven : bEven};
        }

        var FooterInfo = undefined;
        if (null !== Footer && undefined !== Footer.RecalcInfo.PageNumInfo[nPageIndex])
        {
            var bFirst = Footer.RecalcInfo.PageNumInfo[nPageIndex].bFirst;
            var bEven  = Footer.RecalcInfo.PageNumInfo[nPageIndex].bEven;

            var FooterSectPr = Footer.RecalcInfo.SectPr[nPageIndex];

            if (undefined !== FooterSectPr)
                bFirst = ( true === bFirst && true === FooterSectPr.Get_TitlePage() ? true : false );

            FooterInfo = {bFirst : bFirst, bEven : bEven};
        }

		var oHdrFtrLine = this.HdrFtr.GetHdrFtrLines(nPageIndex);
		var nHeaderY    = this.Pages[nPageIndex].Y;
		if (null !== oHdrFtrLine.Top && oHdrFtrLine.Top > nHeaderY)
			nHeaderY = oHdrFtrLine.Top;

		var nFooterY = this.Pages[nPageIndex].YLimit;
		if (null !== oHdrFtrLine.Bottom && oHdrFtrLine.Bottom < nFooterY)
			nFooterY = oHdrFtrLine.Bottom;

        pGraphics.DrawHeaderEdit(nHeaderY, this.HdrFtr.Lock.Get_Type(), SectIndex, RepH, HeaderInfo);
        pGraphics.DrawFooterEdit(nFooterY, this.HdrFtr.Lock.Get_Type(), SectIndex, RepF, FooterInfo);
    }
};
CDocument.prototype.DrawPageBorders = function(Graphics, oSectPr, nPageIndex)
{
	var LBorder = oSectPr.Get_Borders_Left();
	var TBorder = oSectPr.Get_Borders_Top();
	var RBorder = oSectPr.Get_Borders_Right();
	var BBorder = oSectPr.Get_Borders_Bottom();

	var W = oSectPr.GetPageWidth();
	var H = oSectPr.GetPageHeight();

	// Порядок отрисовки границ всегда одинаковый вне зависимости от цветы и толщины: сначала вертикальные,
	// потом горизонтальные поверх вертикальных

	if (section_borders_OffsetFromPage === oSectPr.GetBordersOffsetFrom())
	{
		// Рисуем левую границу
		if (border_None !== LBorder.Value)
		{
			var X  = LBorder.Space + LBorder.Size / 2;
			var Y0 = ( border_None !== TBorder.Value ? TBorder.Space + TBorder.Size / 2 : 0 );
			var Y1 = ( border_None !== BBorder.Value ? H - BBorder.Space - BBorder.Size / 2 : H );

			Graphics.p_color(LBorder.Color.r, LBorder.Color.g, LBorder.Color.b, 255);
			Graphics.drawVerLine(c_oAscLineDrawingRule.Center, X, Y0, Y1, LBorder.Size);
		}

		// Рисуем правую границу
		if (border_None !== RBorder.Value)
		{
			var X  = W - RBorder.Space - RBorder.Size / 2;
			var Y0 = ( border_None !== TBorder.Value ? TBorder.Space + TBorder.Size / 2 : 0 );
			var Y1 = ( border_None !== BBorder.Value ? H - BBorder.Space - BBorder.Size / 2 : H );

			Graphics.p_color(RBorder.Color.r, RBorder.Color.g, RBorder.Color.b, 255);
			Graphics.drawVerLine(c_oAscLineDrawingRule.Center, X, Y0, Y1, RBorder.Size);
		}

		// Рисуем верхнюю границу
		if (border_None !== TBorder.Value)
		{
			var Y  = TBorder.Space;
			var X0 = ( border_None !== LBorder.Value ? LBorder.Space + LBorder.Size / 2 : 0 );
			var X1 = ( border_None !== RBorder.Value ? W - RBorder.Space - RBorder.Size / 2 : W );

			Graphics.p_color(TBorder.Color.r, TBorder.Color.g, TBorder.Color.b, 255);
			Graphics.drawHorLineExt(c_oAscLineDrawingRule.Top, Y, X0, X1, TBorder.Size, ( border_None !== LBorder.Value ? -LBorder.Size / 2 : 0 ), ( border_None !== RBorder.Value ? RBorder.Size / 2 : 0 ));
		}

		// Рисуем нижнюю границу
		if (border_None !== BBorder.Value)
		{
			var Y  = H - BBorder.Space;
			var X0 = ( border_None !== LBorder.Value ? LBorder.Space + LBorder.Size / 2 : 0 );
			var X1 = ( border_None !== RBorder.Value ? W - RBorder.Space - RBorder.Size / 2 : W );

			Graphics.p_color(BBorder.Color.r, BBorder.Color.g, BBorder.Color.b, 255);
			Graphics.drawHorLineExt(c_oAscLineDrawingRule.Bottom, Y, X0, X1, BBorder.Size, ( border_None !== LBorder.Value ? -LBorder.Size / 2 : 0 ), ( border_None !== RBorder.Value ? RBorder.Size / 2 : 0 ));
		}
	}
	else
	{
		var oFrame = oSectPr.GetContentFrame(nPageIndex);

		var _X0 = oFrame.Left;
		var _X1 = oFrame.Right;
		var _Y0 = oFrame.Top;
		var _Y1 = oFrame.Bottom;

		// Рисуем левую границу
		if (border_None !== LBorder.Value)
		{
			var X  = _X0 - LBorder.Space;
			var Y0 = ( border_None !== TBorder.Value ? _Y0 - TBorder.Space - TBorder.Size / 2 : _Y0 );
			var Y1 = ( border_None !== BBorder.Value ? _Y1 + BBorder.Space + BBorder.Size / 2 : _Y1 );

			Graphics.p_color(LBorder.Color.r, LBorder.Color.g, LBorder.Color.b, 255);
			Graphics.drawVerLine(c_oAscLineDrawingRule.Right, X, Y0, Y1, LBorder.Size);
		}

		// Рисуем правую границу
		if (border_None !== RBorder.Value)
		{
			var X  = _X1 + RBorder.Space;
			var Y0 = ( border_None !== TBorder.Value ? _Y0 - TBorder.Space - TBorder.Size / 2 : _Y0 );
			var Y1 = ( border_None !== BBorder.Value ? _Y1 + BBorder.Space + BBorder.Size / 2 : _Y1 );

			Graphics.p_color(RBorder.Color.r, RBorder.Color.g, RBorder.Color.b, 255);
			Graphics.drawVerLine(c_oAscLineDrawingRule.Left, X, Y0, Y1, RBorder.Size);
		}

		// Рисуем верхнюю границу
		if (border_None !== TBorder.Value)
		{
			var Y  = _Y0 - TBorder.Space;
			var X0 = ( border_None !== LBorder.Value ? _X0 - LBorder.Space : _X0 );
			var X1 = ( border_None !== RBorder.Value ? _X1 + RBorder.Space : _X1 );

			Graphics.p_color(TBorder.Color.r, TBorder.Color.g, TBorder.Color.b, 255);
			Graphics.drawHorLineExt(c_oAscLineDrawingRule.Bottom, Y, X0, X1, TBorder.Size, ( border_None !== LBorder.Value ? -LBorder.Size : 0 ), ( border_None !== RBorder.Value ? RBorder.Size : 0 ));
		}

		// Рисуем нижнюю границу
		if (border_None !== BBorder.Value)
		{
			var Y  = _Y1 + BBorder.Space;
			var X0 = ( border_None !== LBorder.Value ? _X0 - LBorder.Space : _X0 );
			var X1 = ( border_None !== RBorder.Value ? _X1 + RBorder.Space : _X1 );

			Graphics.p_color(BBorder.Color.r, BBorder.Color.g, BBorder.Color.b, 255);
			Graphics.drawHorLineExt(c_oAscLineDrawingRule.Top, Y, X0, X1, BBorder.Size, ( border_None !== LBorder.Value ? -LBorder.Size : 0 ), ( border_None !== RBorder.Value ? RBorder.Size : 0 ));
		}
	}

	// TODO: Реализовать:
	//       1. ArtBorders
	//       2. Различные типы обычных границ. Причем, если пересакающиеся границы имеют одинаковый тип и размер,
	//          тогда надо специально отрисовывать места соединения данных линий.

};
/**
 *
 * @param bRecalculate
 * @param bForceAdd - добавляем параграф, пропуская всякие проверки типа пустого параграфа с нумерацией.
 */
CDocument.prototype.AddNewParagraph = function(bRecalculate, bForceAdd)
{
	this.Controller.AddNewParagraph(bRecalculate, bForceAdd);

	if (false !== bRecalculate)
		this.Recalculate();

	this.UpdateInterface();
	this.private_UpdateCursorXY(true, true);
};
/**
 * Расширяем документ до точки (X,Y) с помощью новых параграфов.
 * Y0 - низ последнего параграфа, YLimit - предел страницы.
 * @param X
 * @param Y
 */
CDocument.prototype.Extend_ToPos = function(X, Y)
{
	if (!this.CanPerformAction())
		return;

    var LastPara  = this.GetLastParagraph();
    var LastPara2 = LastPara;

    this.StartAction(AscDFH.historydescription_Document_DocumentExtendToPos);
    this.History.Set_Additional_ExtendDocumentToPos();

    while (true)
    {
        var NewParagraph = new Paragraph(this.DrawingDocument, this);
        var NewRun       = new ParaRun(NewParagraph, false);
        NewParagraph.Add_ToContent(0, NewRun);

        var StyleId = LastPara.Style_Get();
        var NextId  = undefined;

        if (undefined != StyleId)
        {
            NextId = this.Styles.Get_Next(StyleId);

            if (null === NextId || undefined === NextId)
                NextId = StyleId;
        }

        // Простое добавление стиля, без дополнительных действий
        if (NextId === this.Styles.Get_Default_Paragraph() || NextId === this.Styles.Get_Default_ParaList())
            NewParagraph.Style_Remove();
        else
            NewParagraph.Style_Add(NextId, true);

        if (undefined != LastPara.TextPr.Value.FontSize || undefined !== LastPara.TextPr.Value.RFonts.Ascii)
        {
            var TextPr        = new CTextPr();
            TextPr.FontSize   = LastPara.TextPr.Value.FontSize;
            TextPr.FontSizeCS = LastPara.TextPr.Value.FontSize;
            TextPr.RFonts     = LastPara.TextPr.Value.RFonts.Copy();
            NewParagraph.SelectAll();
            NewParagraph.Apply_TextPr(TextPr);
        }

        var CurPage = LastPara.Pages.length - 1;
        var X0      = LastPara.Pages[CurPage].X;
        var Y0      = LastPara.Pages[CurPage].Bounds.Bottom;
        var XLimit  = LastPara.Pages[CurPage].XLimit;
        var YLimit  = LastPara.Pages[CurPage].YLimit;
        var PageNum = LastPara.PageNum;

		this.AddToContent(this.Content.length, NewParagraph, false);

        NewParagraph.Reset(X0, Y0, XLimit, YLimit, PageNum);
        var RecalcResult = NewParagraph.Recalculate_Page(0);

        if (recalcresult_NextElement != RecalcResult)
        {
        	this.RemoveFromContent(this.Content.length - 1, 1, false);
            break;
        }

        if (NewParagraph.Pages[0].Bounds.Bottom > Y)
            break;

        LastPara = NewParagraph;
    }

    LastPara = this.Content[this.Content.length - 1];

    if (LastPara != LastPara2 || false === this.Document_Is_SelectionLocked(changestype_None, {
            Type      : changestype_2_Element_and_Type,
            Element   : LastPara,
            CheckType : changestype_Paragraph_Content
        }))
    {
        // Теперь нам нужно вставить таб по X
        LastPara.Extend_ToPos(X);
    }

    LastPara.MoveCursorToEndPos();
    LastPara.Document_SetThisElementCurrent(true);

    this.Recalculate();
    this.FinalizeAction();
};
CDocument.prototype.GroupGraphicObjects = function()
{
	if (this.CanGroup())
	{
		this.DrawingObjects.groupSelectedObjects();
	}
};
CDocument.prototype.UnGroupGraphicObjects = function()
{
	if (this.CanUnGroup())
	{
		this.DrawingObjects.unGroupSelectedObjects();
	}
};
CDocument.prototype.StartChangeWrapPolygon = function()
{
	this.DrawingObjects.startChangeWrapPolygon();
};
CDocument.prototype.CanChangeWrapPolygon = function()
{
	return this.DrawingObjects.canChangeWrapPolygon();
};
CDocument.prototype.CanGroup = function()
{
	return this.DrawingObjects.canGroup();
};
CDocument.prototype.CanUnGroup = function()
{
	return this.DrawingObjects.canUnGroup();
};
CDocument.prototype.AddInlineImage = function(W, H, Img, Chart, bFlow)
{
	if (undefined === bFlow)
		bFlow = false;


    this.TurnOff_InterfaceEvents();
    this.Controller.AddInlineImage(W, H, Img, Chart, bFlow);
    this.TurnOn_InterfaceEvents(true);
};

CDocument.prototype.AddImages = function(aImages){
    this.Controller.AddImages(aImages);
};

CDocument.prototype.AddOleObject  = function(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory)
{
	return this.Controller.AddOleObject(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory);
};
CDocument.prototype.EditOleObject = function(oOleObject, sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory)
{
    oOleObject.editExternal(sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory);
};
CDocument.prototype.AddTextArt = function(nStyle)
{
	this.Controller.AddTextArt(nStyle);
};

CDocument.prototype.AddSignatureLine = function(oSignatureDrawing){
    this.Controller.AddSignatureLine(oSignatureDrawing);
};

CDocument.prototype.EditChart = function(Chart)
{
	this.Controller.EditChart(Chart);
};
CDocument.prototype.GetChartObject = function(type)
{
    var W = null, H = null;
    if(type != null)
    {
        var oTargetContent = this.DrawingObjects.getTargetDocContent();
        if(oTargetContent && !oTargetContent.bPresentation)
        {
            W = oTargetContent.XLimit;
            H = W;
        }
        else
        {
            var oColumnSize = this.GetColumnSize();
            if(oColumnSize)
            {
                W = oColumnSize.W;
                H = oColumnSize.H;
            }
        }
    }
    return this.DrawingObjects.getChartObject(type, W, H);

};
CDocument.prototype.GetImageDataFromSelection = function()
{
    return this.DrawingObjects.getImageDataFromSelection();

};
CDocument.prototype.PutImageToSelection = function(sImageSrc, nWidth, nHeight, replaceMode)
{
    return this.DrawingObjects.putImageToSelection(sImageSrc, nWidth, nHeight, replaceMode);

};
/**
 * Добавляем таблицу в текущую позицию курсора
 * @param {number} nCols
 * @param {number} nRows
 * @param {number} [nMode=0] 0 - делим параграф в текущей точке, -1 добавляем до текущего параграфа, 1 - добавляем после текущего параграфа
 * @returns {?CTable}
 */
CDocument.prototype.AddInlineTable = function(nCols, nRows, nMode)
{
	if (nCols <= 0 || nRows <= 0)
		return null;

	var oTable = this.Controller.AddInlineTable(nCols, nRows, nMode);

	this.Recalculate();
	this.UpdateSelection();
	this.UpdateInterface();
	this.UpdateRulers();

	return oTable;
};
CDocument.prototype.GetTableForPreview = function()
{
	return AscCommon.ExecuteNoHistory(function()
	{
		let nCols  = 5, nRows = 5;
		let _x_mar = 10;
		let _y_mar = 10;
		let _r_mar = 10;
		let _b_mar = 10;
		let _pageW = 297;
		let _pageH = 210;
		
		let W = (_pageW - _x_mar - _r_mar);
		let H = (_pageH - _y_mar - _b_mar);
		
		let arrGrid = [];
		for (let nIndex = 0; nIndex < nCols; ++nIndex)
			arrGrid[nIndex] = W / nCols;
		
		let oDocumentContent = new CDocumentContent();
		oDocumentContent.SetLogicDocument(this);
		
		let oTable = new CTable(this.GetDrawingDocument(), oDocumentContent, true, nRows, nCols, arrGrid, false);
		oTable.Reset(_x_mar, _y_mar, 1000, 1000, 0, 0, 1);
		oTable.Set_Props({
			TableDefaultMargins : {Top : 0, Bottom : 0},
			TableLayout         : c_oAscTableLayout.Fixed
		});
		
		for (let nCurRow = 0, nRowsCount = oTable.GetRowsCount(); nCurRow < nRowsCount; ++nCurRow)
			oTable.GetRow(nCurRow).SetHeight(H / nRows, Asc.linerule_AtLeast);
		
		return oTable;
	}, this, this, []);
};
CDocument.prototype.CheckTableForPreview = function(oTable)
{};
CDocument.prototype.AddDropCap = function(bInText)
{
	// Определим параграф, к которому мы будем добавлять буквицу
	var Pos = -1;

	if (false === this.Selection.Use)
		Pos = this.CurPos.ContentPos;
	else if (true === this.Selection.Use && this.Selection.StartPos <= this.Selection.EndPos)
		Pos = this.Selection.StartPos;
	else if (true === this.Selection.Use && this.Selection.StartPos > this.Selection.EndPos)
		Pos = this.Selection.EndPos;

	if (-1 === Pos || !this.Content[Pos].IsParagraph())
		return;

	var OldParagraph = this.Content[Pos];

	if (this.IsSelectionUse() && !OldParagraph.CheckSelectionForDropCap())
		this.RemoveSelection();

	if (!this.IsSelectionUse() && !OldParagraph.SelectFirstLetter())
		return;

	if (OldParagraph.Lines.length <= 0)
		return;

	if (!this.IsSelectionLocked(changestype_Paragraph_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_AddDropCap);

		var NewParagraph = new Paragraph(this.DrawingDocument, this);

		var TextPr = OldParagraph.Split_DropCap(NewParagraph);
		var Before = OldParagraph.Get_CompiledPr().ParaPr.Spacing.Before;
		var LineH  = OldParagraph.Lines[0].Bottom - OldParagraph.Lines[0].Top - Before;
		var LineTA = OldParagraph.Lines[0].Metrics.TextAscent2;
		var LineTD = OldParagraph.Lines[0].Metrics.TextDescent + OldParagraph.Lines[0].Metrics.LineGap;

		var FramePr = new CFramePr();
		FramePr.Init_Default_DropCap(bInText);
		NewParagraph.Set_FrameParaPr(OldParagraph);
		NewParagraph.Set_FramePr2(FramePr);
		NewParagraph.Update_DropCapByLines(TextPr, NewParagraph.Pr.FramePr.Lines, LineH, LineTA, LineTD, Before);

		this.Internal_Content_Add(Pos, NewParagraph);
		NewParagraph.MoveCursorToEndPos();

		this.RemoveSelection();
		this.CurPos.ContentPos = Pos;
		this.SetDocPosType(docpostype_Content);

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateRulers();
		this.UpdateTracks();
		this.FinalizeAction();
	}
};
CDocument.prototype.RemoveDropCap = function(bDropCap)
{
	var Pos = -1;

	if (false === this.Selection.Use && type_Paragraph === this.Content[this.CurPos.ContentPos].GetType())
		Pos = this.CurPos.ContentPos;
	else if (true === this.Selection.Use && this.Selection.StartPos <= this.Selection.EndPos && type_Paragraph === this.Content[this.Selection.StartPos].GetType())
		Pos = this.Selection.StartPos;
	else if (true === this.Selection.Use && this.Selection.StartPos > this.Selection.EndPos && type_Paragraph === this.Content[this.Selection.EndPos].GetType())
		Pos = this.Selection.EndPos;

	if (-1 === Pos)
		return;

	var Para    = this.Content[Pos];
	var FramePr = Para.Get_FramePr();

	// Возможно буквицой является предыдущий параграф
	if (undefined === FramePr && true === bDropCap)
	{
		var Prev = Para.Get_DocumentPrev();
		if (null != Prev && type_Paragraph === Prev.GetType())
		{
			var PrevFramePr = Prev.Get_FramePr();
			if (undefined != PrevFramePr && undefined != PrevFramePr.DropCap)
			{
				Para    = Prev;
				FramePr = PrevFramePr;
			}
			else
				return;
		}
		else
			return;
	}

	if (undefined === FramePr)
		return;

	var FrameParas = Para.Internal_Get_FrameParagraphs();

	if (false === bDropCap)
	{
		if (false === this.Document_Is_SelectionLocked(changestype_None, {
				Type      : changestype_2_ElementsArray_and_Type,
				Elements  : FrameParas,
				CheckType : changestype_Paragraph_Content
			}))
		{
			this.StartAction(AscDFH.historydescription_Document_RemoveDropCap);
			var Count = FrameParas.length;
			for (var Index = 0; Index < Count; Index++)
			{
				FrameParas[Index].Set_FramePr(undefined, true);
			}

			this.Recalculate();
			this.UpdateInterface();
			this.UpdateRulers();
			this.FinalizeAction();
		}
	}
	else
	{
		// Сначала найдем параграф, к которому относится буквица
		var Next = Para.Get_DocumentNext();
		var Last = Para;
		while (null != Next)
		{
			if (type_Paragraph != Next.GetType() || undefined === Next.Get_FramePr() || true != FramePr.Compare(Next.Get_FramePr()))
				break;

			Last = Next;
			Next = Next.Get_DocumentNext();
		}

		if (null != Next && type_Paragraph === Next.GetType())
		{
			FrameParas.push(Next);
			if (false === this.Document_Is_SelectionLocked(changestype_None, {
					Type      : changestype_2_ElementsArray_and_Type,
					Elements  : FrameParas,
					CheckType : changestype_Paragraph_Content
				}))
			{
				this.StartAction(AscDFH.historydescription_Document_RemoveDropCap);

				// Удалим ненужный элемент
				FrameParas.splice(FrameParas.length - 1, 1);

				// Передвинем курсор в начало следующего параграфа, и рассчитаем текстовые настройки и расстояния между строк
				Next.MoveCursorToStartPos();
				var Spacing = Next.Get_CompiledPr2(false).ParaPr.Spacing.Copy();
				var TextPr  = Next.GetFirstRunPr();

				var Count = FrameParas.length;
				for (var Index = 0; Index < Count; Index++)
				{
					var FramePara = FrameParas[Index];
					FramePara.Set_FramePr(undefined, true);
					FramePara.Set_Spacing(Spacing, true);
					FramePara.SelectAll();
					FramePara.Clear_TextFormatting();
					FramePara.Apply_TextPr(TextPr, undefined);
				}


				Next.CopyPr(Last);
				Last.Concat(Next);

				this.Internal_Content_Remove(Next.Index, 1);

				Last.MoveCursorToStartPos();
				Last.Document_SetThisElementCurrent(true);

				this.Recalculate();
				this.UpdateInterface();
				this.UpdateRulers();
				this.FinalizeAction();
			}
		}
		else
		{
			if (false === this.Document_Is_SelectionLocked(changestype_None, {
					Type      : changestype_2_ElementsArray_and_Type,
					Elements  : FrameParas,
					CheckType : changestype_Paragraph_Content
				}))
			{
				this.StartAction(AscDFH.historydescription_Document_RemoveDropCap);
				var Count = FrameParas.length;
				for (var Index = 0; Index < Count; Index++)
				{
					FrameParas[Index].Set_FramePr(undefined, true);
				}

				this.Recalculate();
				this.UpdateInterface();
				this.UpdateRulers();
				this.FinalizeAction();
			}
		}
	}
};
CDocument.prototype.private_CheckFramePrLastParagraph = function()
{
    var Count = this.Content.length;
    if (Count <= 0)
        return;

    var Element = this.Content[Count - 1];
    if (type_Paragraph === Element.GetType() && undefined !== Element.Get_FramePr())
    {
        Element.Set_FramePr(undefined, true);
    }
};
CDocument.prototype.CheckRange = function(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, PageNum, Inner, bMathWrap)
{
	var HdrFtrRanges = this.HdrFtr.CheckRange(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, PageNum, bMathWrap);
	return this.DrawingObjects.CheckRange(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, PageNum, HdrFtrRanges, null, bMathWrap);
};
CDocument.prototype.AddToParagraph = function(oParaItem, bRecalculate)
{
	if (!oParaItem)
		return;

	if (this.IsNumberingSelection() && para_TextPr !== oParaItem.Type)
		this.RemoveSelectedNumberingOnTextAdd();
	
	this.Controller.AddToParagraph(oParaItem, bRecalculate);
};
CDocument.prototype.RemoveSelectedNumberingOnTextAdd = function()
{
	if (!this.IsNumberingSelection())
		return;
	
	this.RemoveSelection();
	
	let paragraph = this.GetCurrentParagraph();
	if (!paragraph)
		return;
	
	let numPr = paragraph.GetNumPr();
	if (!numPr)
		return;
	
	let num = this.Numbering.GetNum(numPr.NumId);
	if (!num)
		return;
	
	let numLvl = num.GetLvl(numPr.Lvl);
	let paraPr = numLvl.GetParaPr();
	if (!paraPr
		|| !paraPr.Ind
		|| undefined === paraPr.Ind.FirstLine
		|| undefined === paraPr.Ind.Left)
		return;
	
	if (this.Styles.Get_Default_ParaList() === paragraph.GetParagraphStyle())
		paragraph.SetParagraphStyleById(null);

	paragraph.RemoveNumPr();
	paragraph.SetParagraphIndent({FirstLine : 0, Left : paraPr.Ind.FirstLine + paraPr.Ind.Left});
};
/**
 * Очищаем форматирование внутри селекта
 * {boolean} [isClearParaPr=true] Очищать ли настройки параграфа
 * {boolean} [isClearTextPr=true] Очищать ли настройки текста
 */
CDocument.prototype.ClearParagraphFormatting = function(isClearParaPr, isClearTextPr)
{
	if (false !== isClearParaPr)
		isClearParaPr = true;

	if (false !== isClearTextPr)
		isClearTextPr = true;

	this.Controller.ClearParagraphFormatting(isClearParaPr, isClearTextPr);

	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.CheckSubFormBeforeRemove = function(nDirection)
{
	if (!this.IsFillingFormMode() || this.IsTextSelectionUse())
		return;

	let oForm = this.GetSelectedElementsInfo().GetInlineLevelSdt();
	let oMainForm;
	if (!oForm || !oForm.IsForm() || !(oMainForm = oForm.GetMainForm()) || oMainForm === oForm)
		return;

	if (!((nDirection < 0 && oForm.IsCursorAtBegin()) || (nDirection > 0 && oForm.IsCursorAtEnd())))
		return;

	let oNextForm = nDirection > 0 ? oForm.GetNextSubForm() : oForm.GetPrevSubForm();
	if (!oNextForm || oNextForm === oForm)
		return;

	oNextForm.SetThisElementCurrent();

	if (nDirection > 0)
		oNextForm.MoveCursorToStartPos();
	else
		oNextForm.MoveCursorToEndPos();
};
CDocument.prototype.Remove = function(nDirection, isRemoveWholeElement, bRemoveOnlySelection, bOnTextAdd, isWord, isCheckInlineLevelSdt)
{
	if (undefined === nDirection)
		nDirection = 1;

	if (undefined === isRemoveWholeElement)
		isRemoveWholeElement = false;

	if (undefined === bRemoveOnlySelection)
		bRemoveOnlySelection = false;

	if (undefined === bOnTextAdd)
		bOnTextAdd = false;

	if (undefined === isWord)
		isWord = false;

	if (undefined === isCheckInlineLevelSdt)
		isCheckInlineLevelSdt = false;

	if (isCheckInlineLevelSdt)
	{
		// Эта проверка важна для работы с fixed-form, чтобы CC не удалялся при простом backspace/remove/cut
		let selectedInfo = this.GetSelectedElementsInfo();
		let inlineSdt    = selectedInfo.GetInlineLevelSdt();
		if (inlineSdt && inlineSdt.IsForm() && inlineSdt.IsMainForm())
			this.CheckInlineSdtOnDelete = inlineSdt;
	}

	this.Controller.Remove(nDirection, isRemoveWholeElement, bRemoveOnlySelection, bOnTextAdd, isWord);

	this.CheckInlineSdtOnDelete = null;

	this.Recalculate();
	this.UpdateInterface();
	this.UpdateRulers();
	this.UpdateTracks();
};
CDocument.prototype.RemoveBeforePaste = function()
{
	if ((docpostype_DrawingObjects === this.GetDocPosType() && null === this.DrawingObjects.getTargetDocContent())
		|| (docpostype_HdrFtr === this.GetDocPosType() && this.HdrFtr.CurHdrFtr && docpostype_DrawingObjects === this.HdrFtr.CurHdrFtr.GetContent().GetDocPosType() && null === this.DrawingObjects.getTargetDocContent()))
		this.RemoveSelection();
	else
		this.Remove(1, true, true, true);
};
CDocument.prototype.RemoveDrawingObjectById = function(sObjectId)
{
    var oState = this.SaveDocumentState();
    var oDrawing = AscCommon.g_oTableId.Get_ById(sObjectId);
    oDrawing.Set_CurrentElement(false, null);
    if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Remove, null, true, this.IsFormFieldEditing()))
    {
        this.StartAction(AscDFH.historydescription_Document_BackSpaceButton);
        this.Remove(-1, true, false, false, false, true);
        this.FinalizeAction();
    }
    this.LoadDocumentState(oState);
};
CDocument.prototype.RemoveDrawingObjects = function(arrObjectsId)
{
    var oState = this.SaveDocumentState();
    var oDrawing = AscCommon.g_oTableId.Get_ById(arrObjectsId[0]);
    if(!oDrawing)
    {
        return;
    }
    oDrawing.Set_CurrentElement(false, null);
    var aObjectsToDelete = [];
    for(var nObject = 0; nObject < arrObjectsId.length; ++nObject)
    {
        oDrawing = AscCommon.g_oTableId.Get_ById(arrObjectsId[nObject]);
        if(oDrawing)
        {
            aObjectsToDelete.push(oDrawing);
            if(oDrawing.group)
            {
                oDrawing.getMainGroup().select(this.DrawingObjects);
            }
            else
            {
                oDrawing.select(this.DrawingObjects);
            }
        }
    }
    if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Remove, null, true, this.IsFormFieldEditing()))
    {
        this.StartAction(AscDFH.historydescription_Document_BackSpaceButton);
        for(nObject = 0; nObject < aObjectsToDelete.length; ++nObject)
        {
            oDrawing = aObjectsToDelete[nObject];
            oDrawing.Set_CurrentElement(false, null);
            this.Remove(-1, true, false, false, false, true);
        }
        this.FinalizeAction();
    }
    this.LoadDocumentState(oState);
};
CDocument.prototype.GetCursorPosXY = function()
{
	return this.Controller.GetCursorPosXY();
};
/**
 * Получаем точное физическое положение курсора
 * @returns {{X: number, Y: number}}
 */
CDocument.prototype.GetCursorRealPosition = function()
{
	return {
		X : this.CurPos.RealX,
		Y : this.CurPos.RealY
	};
};
CDocument.prototype.CorrectCursorPosition = function(isForward)
{
	if (this.IsFillingFormMode())
	{
		let oInfo = this.GetSelectedElementsInfo();
		let oForm = oInfo.GetInlineLevelSdt();

		// Селект может быть внутри составных форм в режиме заполнения, а вот курсор там не должен находиться
		if (oForm && oForm.IsForm() && oForm.IsComplexForm())
		{
			let oSubForm = oForm.GetSubFormFromCurrentPosition(isForward);
			if (oSubForm)
			{
				oSubForm.SetThisElementCurrent(true);

				if (isForward)
					oSubForm.MoveCursorToStartPos();
				else
					oSubForm.MoveCursorToEndPos();
			}
		}
	}
};
CDocument.prototype.MoveCursorToStartOfDocument = function()
{
	var nDocPosType = this.GetDocPosType();

	if (nDocPosType === docpostype_DrawingObjects)
		this.EndDrawingEditing();
	else if (nDocPosType === docpostype_Footnotes)
		this.EndFootnotesEditing();
	else if (nDocPosType === docpostype_Endnotes)
		this.EndEndnotesEditing();
	else if (nDocPosType === docpostype_HdrFtr)
		this.EndHdrFtrEditing();

	this.RemoveSelection();
	this.SetDocPosType(docpostype_Content);
	this.MoveCursorToStartPos(false);

	this.private_CheckCursorPosInFillingFormMode();
};
CDocument.prototype.MoveCursorToStartPos = function(AddToSelect)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();

	this.Controller.MoveCursorToStartPos(AddToSelect);

	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.MoveCursorToEndPos = function(AddToSelect)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();

	this.Controller.MoveCursorToEndPos(AddToSelect);

	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.MoveCursorLeft = function(AddToSelect, Word)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();

	if (undefined === Word || null === Word)
		Word = false;

	this.Controller.MoveCursorLeft(AddToSelect, Word);

	this.CorrectCursorPosition(false);
	this.Document_UpdateInterfaceState();
	this.Document_UpdateRulersState();
	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.MoveCursorRight = function(AddToSelect, Word, FromPaste)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();

	if (undefined === Word || null === Word)
		Word = false;

	this.Controller.MoveCursorRight(AddToSelect, Word, FromPaste);

	this.CorrectCursorPosition(true);
	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.MoveCursorUp = function(AddToSelect, CtrlKey)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();
	this.Controller.MoveCursorUp(AddToSelect, CtrlKey);
};
CDocument.prototype.MoveCursorDown = function(AddToSelect, CtrlKey)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();
	this.Controller.MoveCursorDown(AddToSelect, CtrlKey);
};
CDocument.prototype.MoveCursorToEndOfLine = function(AddToSelect)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();

	this.Controller.MoveCursorToEndOfLine(AddToSelect);

	this.Document_UpdateInterfaceState();
	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.MoveCursorToStartOfLine = function(AddToSelect)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();

	this.Controller.MoveCursorToStartOfLine(AddToSelect);

	this.Document_UpdateInterfaceState();
	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.MoveCursorToXY = function(X, Y, AddToSelect)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();
	this.Controller.MoveCursorToXY(X, Y, this.CurPage, AddToSelect);
};
CDocument.prototype.MoveCursorToCell = function(bNext)
{
	this.ResetWordSelection();
	this.private_UpdateTargetForCollaboration();
	this.Controller.MoveCursorToCell(bNext);
};
CDocument.prototype.GoToSignature = function(sGuid)
{
    this.DrawingObjects.goToSignature(sGuid);
};
CDocument.prototype.MoveCursorToPageStart = function()
{
	if (docpostype_Content !== this.GetDocPosType())
	{
		this.RemoveSelection();
		this.SetDocPosType(docpostype_Content);
	}

	this.MoveCursorToXY(0, 0, false);
	this.UpdateInterface();
	this.UpdateSelection();
};
CDocument.prototype.MoveCursorToPageEnd = function()
{
	if (docpostype_Content !== this.GetDocPosType())
	{
		this.RemoveSelection();
		this.SetDocPosType(docpostype_Content);
	}

	this.MoveCursorToXY(0, 10000, false);
	this.UpdateInterface();
	this.UpdateSelection();
};
/**
 * Переходим на заданную страницу, если на ней нет ничего из основной части документа, то идем вперед
 * пока не найдем подходящую, если не нашли, то ищем в обратном направлении первую подходящую страницу
 * (т.е. пустые страницы или страницы состоящие целиком из сносок мы пропускаем и на такие не переходим)
 * @param nPageAbs
 * @returns {number}
 */
CDocument.prototype.GoToPage = function(nPageAbs)
{
	if ((this.FullRecalc.Id && this.FullRecalc.PageIndex - 1 <= nPageAbs)
		|| this.Pages.length < 0)
		return;

	var nLastPage = this.FullRecalc.Id ? Math.min(this.FullRecalc.PageIndex - 1, this.Pages.length - 1) : this.Pages.length - 1;
	var nCurPage    = nPageAbs;
	while (nCurPage <= nLastPage)
	{
		var oPage = this.Pages[nCurPage];
		if (-1 !== oPage.Pos && oPage.Pos <= oPage.EndPos)
			break;

		nCurPage++;
	}

	if (nCurPage > nLastPage)
	{
		nCurPage = nPageAbs;
		while (nCurPage > 0)
		{
			var oPage = this.Pages[nCurPage];
			if (-1 !== oPage.Pos && oPage.Pos <= oPage.EndPos)
				break;

			nCurPage--;
		}
	}

	this.RemoveSelection();
	this.SetDocPosType(docpostype_Content);
	
	const step = 3;
	let x = 0;
	let y = 0;
	let xLimit = this.Pages[nCurPage].XLimit;
	let yLimit = this.Pages[nCurPage].YLimit;
	
	while (this.Get_CurPage() !== nCurPage)
	{
		this.Set_CurPage(nCurPage);
		this.MoveCursorToXY(x, y, false);
		this.RecalculateCurPos();
		
		x += step;
		if (x > xLimit)
		{
			x = 0;
			y += step;
			
			if (y > yLimit)
				break;
		}
	}
	this.UpdateSelection();

	return nCurPage;
};
CDocument.prototype.SetParagraphAlign = function(Align)
{
	var SelectedInfo = this.GetSelectedElementsInfo();
	var Math         = SelectedInfo.GetMath();
	if (null !== Math && true !== Math.Is_Inline())
	{
		Math.Set_Align(Align);
	}
	else
	{
		this.Controller.SetParagraphAlign(Align);
	}

	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphSpacing = function(Spacing)
{
	this.Controller.SetParagraphSpacing(Spacing);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphTabs = function(Tabs)
{
	this.Controller.SetParagraphTabs(Tabs);
	this.Recalculate();
	this.UpdateSelection();
	this.UpdateInterface();
	this.Api.Update_ParaTab(AscCommonWord.Default_Tab_Stop, Tabs);
};
CDocument.prototype.SetParagraphIndent = function(Ind)
{
	this.Controller.SetParagraphIndent(Ind);
	this.Recalculate();
	this.UpdateSelection();
	this.UpdateInterface();
};
CDocument.prototype.SetParagraphNumbering = function(numInfo)
{
	let result = this.NumberingApplicator.Apply(numInfo);
	if (result)
	{
		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.UpdateStylePanel();
		
		let paraPr = this.GetCalculatedParaPr();
		let numPr  = paraPr.NumPr;
		return (numPr && numPr.IsValid() ? numPr : null);
	}
	else
	{
		return null;
	}
};
CDocument.prototype.SetParagraphOutlineLvl = function(nLvl)
{
	var arrParagraphs = this.GetSelectedParagraphs();
	for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
	{
		arrParagraphs[nIndex].SetOutlineLvl(nLvl);
		arrParagraphs[nIndex].UpdateDocumentOutline();
	}

	this.Recalculate();
};
CDocument.prototype.SetParagraphSuppressLineNumbers = function(isSuppress)
{
	var arrParagraphs = this.GetSelectedParagraphs();
	for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
	{
		arrParagraphs[nIndex].SetSuppressLineNumbers(isSuppress);
	}
};
CDocument.prototype.SetParagraphShd = function(Shd)
{
	this.Controller.SetParagraphShd(Shd);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
/**
 * Выставляем стиль для выделенных параграфов
 * @param {string} sName - название стиля
 * @param {boolean} [isCheckLinkedStyle=false] - если true и если выделен текст внутри одного параграфа, то мы выставляем линкованный стиль текста, если он есть
 */
CDocument.prototype.SetParagraphStyle = function(sName, isCheckLinkedStyle)
{
	if (isCheckLinkedStyle && this.IsSelectionUse())
	{
		var sStyleId = this.Styles.GetStyleIdByName(sName);
		var oStyle   = this.Styles.Get(sStyleId);

		var arrCurrentParagraphs = this.GetCurrentParagraph(false, true);
		if (1 === arrCurrentParagraphs.length && arrCurrentParagraphs[0].IsSelectedSingleElement() && true !== arrCurrentParagraphs[0].IsSelectedAll() && oStyle && oStyle.GetLink())
		{
			var oLinkedStyle = this.Styles.Get(oStyle.GetLink());
			if (oLinkedStyle && styletype_Character === oLinkedStyle.GetType())
			{
				var oTextPr = new CTextPr();
				oTextPr.Set_FromObject({RStyle : oLinkedStyle.GetId()}, true);
				arrCurrentParagraphs[0].ApplyTextPr(oTextPr);
				this.Recalculate();
				this.Document_UpdateSelectionState();
				this.Document_UpdateInterfaceState();
				return;
			}
		}
	}

	var oParaPr = this.GetCalculatedParaPr();
	if (oParaPr.PStyle && this.Styles.Get(oParaPr.PStyle) && this.Styles.Get(oParaPr.PStyle).GetName() === sName)
		this.Controller.ClearParagraphFormatting(false, true);
	else
		this.Controller.SetParagraphStyle(sName);

	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphContextualSpacing = function(Value)
{
	this.Controller.SetParagraphContextualSpacing(Value);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphPageBreakBefore = function(Value)
{
	this.Controller.SetParagraphPageBreakBefore(Value);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphKeepLines = function(Value)
{
	this.Controller.SetParagraphKeepLines(Value);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphKeepNext = function(Value)
{
	this.Controller.SetParagraphKeepNext(Value);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphWidowControl = function(Value)
{
	this.Controller.SetParagraphWidowControl(Value);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphBorders = function(Borders)
{
	this.Controller.SetParagraphBorders(Borders);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetParagraphFramePr = function(FramePr, bDelete)
{
	this.Controller.SetParagraphFramePr(FramePr, bDelete);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateRulersState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.IncreaseDecreaseFontSize = function(bIncrease)
{
	this.Controller.IncreaseDecreaseFontSize(bIncrease);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.IncreaseDecreaseIndent = function(bIncrease)
{
	this.Controller.IncreaseDecreaseIndent(bIncrease);
};
CDocument.prototype.SetParagraphHighlight = function(IsColor, r, g, b)
{
	if (true === this.IsTextSelectionUse())
	{
		if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
		{
			this.StartAction(AscDFH.historydescription_Document_SetTextHighlight);

			if (false === IsColor)
				this.AddToParagraph(new ParaTextPr({HighLight : highlight_None}));
			else
				this.AddToParagraph(new ParaTextPr({HighLight : new CDocumentColor(r, g, b)}));

			this.Redraw();
			this.UpdateInterface();
			this.FinalizeAction();
		}
	}
	else
	{
		if (false === IsColor)
			this.HighlightColor = highlight_None;
		else
			this.HighlightColor = new CDocumentColor(r, g, b);
	}
};

CDocument.prototype.GetSelectedDrawingObjectsCount = function()
{
    return this.DrawingObjects.getSelectedDrawingObjectsCount();
};
CDocument.prototype.PutShapesAlign = function(type, align)
{
    return this.DrawingObjects.putShapesAlign(type, align);
};
CDocument.prototype.DistributeDrawingsHorizontally = function(align)
{
    return this.DrawingObjects.distributeHor(align);
};
CDocument.prototype.DistributeDrawingsVertically = function(align)
{
    return this.DrawingObjects.distributeVer(align);
};
CDocument.prototype.SetImageProps = function(Props)
{
	this.Controller.SetImageProps(Props);
	this.Recalculate();
	if(this.Api.noCreatePoint)
	{
		this.UpdateInterface();
	}
	this.UpdateSelection();
};
CDocument.prototype.ShapeApply = function(shapeProps)
{
	this.DrawingObjects.shapeApply(shapeProps);
};
CDocument.prototype.SelectDrawings = function(arrDrawings, oTargetContent)
{
	this.private_UpdateTargetForCollaboration();

	var nCount = arrDrawings.length;
	if (nCount > 1 && arrDrawings[0].IsInline())
		return;

	for (var nIndex = 0; nIndex < nCount; ++nIndex)
	{
		if (!arrDrawings[nIndex].IsUseInDocument())
			return;
	}

	this.DrawingObjects.resetSelection();
	var oHdrFtr = oTargetContent.IsHdrFtr(true);
	if (oHdrFtr)
	{
		oHdrFtr.Content.SetDocPosType(docpostype_DrawingObjects);
		oHdrFtr.Set_CurrentElement(false);
	}
	else
	{
		this.SetDocPosType(docpostype_DrawingObjects);
	}

	for (var i = 0; i < nCount; ++i)
	{
		this.DrawingObjects.selectObject(arrDrawings[i].GraphicObj, 0);
	}
};
CDocument.prototype.SetTableProps = function(Props)
{
	this.Controller.SetTableProps(Props);
	this.Recalculate();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateRulersState();
	this.Document_UpdateSelectionState();
};
/**
 * Получаем рассчитанные настройки параграфа (полностью заполненные)
 * @returns {CParaPr}
 */
CDocument.prototype.GetCalculatedParaPr = function()
{
	return this.Controller.GetCalculatedParaPr();
};
/**
 * Получаем рассчитанные настройки текста (полностью заполненные)
 * @returns {CTextPr}
 */
CDocument.prototype.GetCalculatedTextPr = function()
{
	var ret = this.Controller.GetCalculatedTextPr();
    if(ret)
    {
        var oTheme = this.Get_Theme();
        if(oTheme)
        {
            ret.ReplaceThemeFonts(oTheme.themeElements.fontScheme);
        }
        return ret;
    }
};
/**
 * Получаем прямые настройки параграфа, т.е. которые выставлены непосредственно у параграфа, без учета стилей
 * @returns {CParaPr}
 */
CDocument.prototype.GetDirectTextPr = function()
{
	return this.Controller.GetDirectTextPr();
};
/**
 * Получаем прямые настройки текста, т.е. которые выставлены непосредственно у рана, без учета стилей
 * @returns {CTextPr}
 */
CDocument.prototype.GetDirectParaPr = function()
{
	return this.Controller.GetDirectParaPr();
};
CDocument.prototype.Get_PageSizesByDrawingObjects = function()
{
	return this.DrawingObjects.getPageSizesByDrawingObjects();
};
/**
 * Выставляем поля документа
 * @param oMargins {{Left : number, Top : number, Right : number, Bottom : number}}
 * @param isFromRuler {boolean} пришло ли изменение из линейки
 */
CDocument.prototype.SetDocumentMargin = function(oMargins, isFromRuler)
{
	// TODO: Document.Set_DocumentOrientation Сделать в зависимости от выделения

	var nCurPos = this.CurPos.ContentPos;
	var oSectPr = this.SectionsInfo.Get_SectPr(nCurPos).SectPr;

	var L = oMargins.Left;
	var T = oMargins.Top;
	var R = undefined === oMargins.Right ? undefined : oSectPr.GetPageWidth() - oMargins.Right;
	var B = undefined === oMargins.Bottom ? undefined : oSectPr.GetPageHeight() - oMargins.Bottom;

	if (isFromRuler)
	{
		if (this.IsMirrorMargins() && 1 === this.CurPage % 2)
		{
			L = oSectPr.GetPageWidth() - oMargins.Right;
			R = oMargins.Left;
		}

		var nGutter = oSectPr.GetGutter();
		if (nGutter > 0.001)
		{
			if (this.IsGutterAtTop())
				T = Math.max(0, T - nGutter);
			else if (oSectPr.IsGutterRTL())
				R = Math.max(0, R - nGutter);
			else
				L = Math.max(0, L - nGutter);
		}

		if (oSectPr.GetPageMarginTop() < 0 && T > 0)
			T = -T;
	}

	oSectPr.SetPageMargins(L, T, R, B);
	this.DrawingObjects.CheckAutoFit();

	this.Recalculate();
	this.UpdateSelection();
	this.UpdateInterface();
	this.UpdateRulers();
};
CDocument.prototype.Set_DocumentPageSize = function(W, H, bNoRecalc)
{
	// TODO: Document.Set_DocumentOrientation Сделать в зависимости от выделения

	var CurPos = this.CurPos.ContentPos;
	var SectPr = this.SectionsInfo.Get_SectPr(CurPos).SectPr;

	SectPr.SetPageSize(W, H);

	this.DrawingObjects.CheckAutoFit();
	if (true != bNoRecalc)
	{
		this.Recalculate();
		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateRulersState();
	}
};
CDocument.prototype.Get_DocumentPageSize = function()
{
	// TODO: Document.Get_DocumentOrientation Сделать в зависимости от выделения

	var CurPos             = this.CurPos.ContentPos;
	var SectionInfoElement = this.SectionsInfo.Get_SectPr(CurPos);

	if (undefined === SectionInfoElement)
		return true;

	var SectPr = SectionInfoElement.SectPr;

	return {W : SectPr.GetPageWidth(), H : SectPr.GetPageHeight()};
};
CDocument.prototype.Set_DocumentOrientation = function(Orientation, bNoRecalc)
{
	// TODO: Document.Set_DocumentOrientation Сделать в зависимости от выделения

	var CurPos = this.CurPos.ContentPos;
	var SectPr = this.SectionsInfo.Get_SectPr(CurPos).SectPr;

	SectPr.SetOrientation(Orientation, true);

	this.DrawingObjects.CheckAutoFit();
	if (true != bNoRecalc)
	{
		this.Recalculate();
		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateRulersState();
	}
};
CDocument.prototype.Get_DocumentOrientation = function()
{
	// TODO: Document.Get_DocumentOrientation Сделать в зависимости от выделения

	var CurPos             = this.CurPos.ContentPos;
	var SectionInfoElement = this.SectionsInfo.Get_SectPr(CurPos);

	if (undefined === SectionInfoElement)
		return true;

	var SectPr = SectionInfoElement.SectPr;

	return ( SectPr.GetOrientation() === Asc.c_oAscPageOrientation.PagePortrait ? true : false );
};
CDocument.prototype.Set_DocumentDefaultTab = function(DTab)
{
	this.History.Add(new CChangesDocumentDefaultTab(this, AscCommonWord.Default_Tab_Stop, DTab));
	AscCommonWord.Default_Tab_Stop = DTab;
};
CDocument.prototype.Set_DocumentEvenAndOddHeaders = function(Value)
{
	if (Value !== EvenAndOddHeaders)
	{
		this.History.Add(new CChangesDocumentEvenAndOddHeaders(this, EvenAndOddHeaders, Value));
		EvenAndOddHeaders = Value;
	}
};
CDocument.prototype.RemoveHdrFtr = function(nPageAbs, isHeader)
{
	let oHeader = this.HdrFtr.GetHdrFtr(nPageAbs, isHeader);
	if (!oHeader)
		return;

	if (!this.IsSelectionLocked(AscCommon.changestype_HdrFtr))
	{
		this.StartAction(AscDFH.historydescription_Document_RemoveHdrFtr);

		let nSectPos = this.SectionsInfo.Find_ByHdrFtr(oHeader);
		if (0 === nSectPos)
		{
			if (oHeader === this.HdrFtr.GetCurHdrFtr())
				this.EndHdrFtrEditing(true);

			let oSectPr = this.SectionsInfo.Get(nSectPos).SectPr;
			oSectPr.RemoveHeader(oHeader);
		}
		else
		{
			oHeader.UpdateContentToDefaults();
		}

		this.UpdateSelection();
		this.UpdateInterface();
		this.Recalculate();
		this.FinalizeAction();
	}
};
CDocument.prototype.CanAddDropCap = function()
{
	if (docpostype_Content !== this.GetDocPosType())
		return false;
	
	let paragraph = null;
	if (!this.IsSelectionUse() && this.Content[this.CurPos.ContentPos].IsParagraph())
		paragraph = this.Content[this.CurPos.ContentPos];
	else if (this.IsSelectionUse() && this.Selection.StartPos <= this.Selection.EndPos && this.Content[this.Selection.StartPos].IsParagraph())
		paragraph = this.Content[this.Selection.StartPos];
	else if (this.IsSelectionUse() && this.Selection.StartPos > this.Selection.EndPos && this.Content[this.Selection.EndPos].IsParagraph())
		paragraph = this.Content[this.Selection.EndPos];
	
	if (!paragraph || paragraph.GetFramePr())
		return false;
	
	let prev = paragraph.Get_DocumentPrev();
	return (paragraph.CanAddDropCap()
		&& (!prev
			|| !prev.IsParagraph()
			|| !prev.GetFramePr()
			|| !prev.GetFramePr().DropCap));
};
/**
 * Обновляем данные в интерфейсе о свойствах графики (картинки, автофигуры).
 * @param Flag
 * @returns {*}
 */
CDocument.prototype.Interface_Update_DrawingPr = function(Flag)
{
	var DrawingPr = this.DrawingObjects.Get_Props();

	if (true === Flag)
		return DrawingPr;
	else
	{
		if (this.Api)
		{
			for (var i = 0; i < DrawingPr.length; ++i)
				this.Api.sync_ImgPropCallback(DrawingPr[i]);
		}
	}
	if (Flag)
		return null;
};
/**
 * Обновляем данные в интерфейсе о свойствах таблицы.
 * @param Flag
 * @returns {*}
 */
CDocument.prototype.Interface_Update_TablePr = function(Flag)
{
	var TablePr = null;
	if (docpostype_Content == this.GetDocPosType() && ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos) || false == this.Selection.Use))
	{
		if (true == this.Selection.Use)
			TablePr = this.Content[this.Selection.StartPos].GetTableProps();
		else
			TablePr = this.Content[this.CurPos.ContentPos].GetTableProps();
	}

	if (null !== TablePr)
		TablePr.CanBeFlow = true;

	if (true === Flag)
		return TablePr;
	else if (null != TablePr)
		this.Api.sync_TblPropCallback(TablePr);
};
/**
 * Обновляем данные в интерфейсе о свойствах колонотитулов.
 */
CDocument.prototype.Interface_Update_HdrFtrPr = function()
{
	if (docpostype_HdrFtr === this.GetDocPosType())
	{
		this.Api.sync_HeadersAndFootersPropCallback(this.HdrFtr.Get_Props());
	}
};
CDocument.prototype.Internal_GetContentPosByXY = function(X, Y, nCurPage, ColumnsInfo)
{
    if (!ColumnsInfo)
        ColumnsInfo = {Column : 0, ColumnsCount : 1};

    if (undefined === nCurPage || null === nCurPage)
		nCurPage = this.CurPage;

    if (nCurPage >= this.Pages.length)
    	nCurPage = this.Pages.length - 1;
    else if (nCurPage < 0)
    	nCurPage = 0;

    // Сначала проверим попадание в плавающие объекты
	var oFlow    = this.DrawingObjects.getTableByXY(X, Y, nCurPage, this);
	var nFlowPos = this.private_GetContentIndexByFlowObject(oFlow, X, Y);
	if (-1 !== nFlowPos)
	{
		var oElement             = this.Content[nFlowPos];
		ColumnsInfo.Column       = oElement.GetStartColumn();
		ColumnsInfo.ColumnsCount = oElement.GetColumnsCount();
		return nFlowPos;
	}

    // Теперь проверим пустые параграфы с окончанием секций
    var SectCount = this.Pages[nCurPage].EndSectionParas.length;
    for (var Index = 0; Index < SectCount; ++Index)
    {
        var Item   = this.Pages[nCurPage].EndSectionParas[Index];
        var Bounds = Item.Pages[0].Bounds;

        if (Y < Bounds.Bottom && Y > Bounds.Top && X > Bounds.Left && X < Bounds.Right)
        {
            var Element              = this.Content[Item.Index];
            ColumnsInfo.Column       = Element.Get_StartColumn();
            ColumnsInfo.ColumnsCount = Element.Get_ColumnsCount();
            return Item.Index;
        }
    }

    // Сначала мы определим секцию и колонку, в которую попали
    var Page = this.Pages[nCurPage];

    var SectionIndex = 0;
    for (var SectionsCount = Page.Sections.length; SectionIndex < SectionsCount - 1; ++SectionIndex)
    {
        if (Y < Page.Sections[SectionIndex + 1].Y)
            break;
    }

    var PageSection  = this.Pages[nCurPage].Sections[SectionIndex];
    var ColumnsCount = PageSection.Columns.length;
    var ColumnIndex  = 0;
    for (; ColumnIndex < ColumnsCount - 1; ++ColumnIndex)
    {
        if (X < (PageSection.Columns[ColumnIndex].XLimit + PageSection.Columns[ColumnIndex + 1].X) / 2)
            break;
    }

    // TODO: Разобраться с ситуацией, когда пустые колонки стоят не только в конце
    while (ColumnIndex > 0 && (true === PageSection.Columns[ColumnIndex].Empty || PageSection.Columns[ColumnIndex].EndPos < PageSection.Columns[ColumnIndex].Pos))
        ColumnIndex--;

    ColumnsInfo.Column       = ColumnIndex;
    ColumnsInfo.ColumnsCount = ColumnsCount;

    var Column   = PageSection.Columns[ColumnIndex];
    var StartPos = Column.Pos;
    var EndPos   = Column.EndPos;

    // Сохраним позиции всех Inline элементов на данной странице
    var InlineElements = [];
    for (var Index = StartPos; Index <= EndPos; Index++)
    {
        var Item = this.Content[Index];

        var PrevItem       = Item.Get_DocumentPrev();
        var bEmptySectPara = (type_Paragraph === Item.GetType() && undefined !== Item.Get_SectionPr() && true === Item.IsEmpty() && null !== PrevItem && (type_Paragraph !== PrevItem.GetType() || undefined === PrevItem.Get_SectionPr())) ? true : false;

        if (!Page.IsFlowTable(Item) && !Page.IsFrame(Item) && (type_Paragraph !== Item.GetType() || false === bEmptySectPara))
            InlineElements.push(Index);
    }

    var Count = InlineElements.length;
    if (Count <= 0)
        return Math.min(Math.max(0, Page.EndPos), this.Content.length - 1);

    for (var Pos = 0; Pos < Count - 1; ++Pos)
    {
        var Item = this.Content[InlineElements[Pos + 1]];

        var PageBounds = Item.GetPageBounds(0);
        if (Y < PageBounds.Top)
            return InlineElements[Pos];

        if (Item.GetPagesCount() > 1)
        {
            if (true !== Item.IsStartFromNewPage())
                return InlineElements[Pos + 1];

            return InlineElements[Pos];
        }

        if (Pos === Count - 2)
        {
            // Такое возможно, если страница заканчивается Flow-таблицей
            return InlineElements[Count - 1];
        }
    }

    return InlineElements[0];
};
CDocument.prototype.RemoveSelection = function(bNoCheckDrawing)
{
	// Не отправляем во время действия, потому что во время некоторых действий это может происходить по нескольку раз
	// Если появится необходимость, то это надо обработать на окончании действия
	if (this.IsTextSelectionUse() && !this.Action.Start)
		this.private_OnSelectionCancel();
	
	this.ResetWordSelection();
	this.Controller.RemoveSelection(bNoCheckDrawing);
};
CDocument.prototype.IsSelectionEmpty = function(bCheckHidden)
{
	return this.Controller.IsSelectionEmpty(bCheckHidden);
};

CDocument.prototype.IsTextSelectionInSmartArt = function () {
  return this.DrawingObjects.isTextSelectionInSmartArt();
}
CDocument.prototype.DrawSelectionOnPage = function(PageAbs)
{
	this.DrawingDocument.UpdateTargetTransform(null);
	this.DrawingDocument.SetTextSelectionOutline(false);
	this.Controller.DrawSelectionOnPage(PageAbs);
};
CDocument.prototype.GetSelectionBounds = function()
{
	return this.Controller.GetSelectionBounds();
};
CDocument.prototype.Selection_SetStart         = function(X, Y, MouseEvent)
{
	this.ResetWordSelection();

    var bInText      = (null === this.IsInText(X, Y, this.CurPage) ? false : true);
    var bTableBorder = (null === this.IsTableBorder(X, Y, this.CurPage) ? false : true);
    var nInDrawing   = this.DrawingObjects.IsInDrawingObject(X, Y, this.CurPage, this);
	var bFlowTable   = (null === this.DrawingObjects.getTableByXY(X, Y, this.CurPage, this) ? false : true);

	// Сначала посмотрим, попалили мы в текстовый селект (но при этом не в границу таблицы и не более чем одинарным кликом)
	if (this.CanDragAndDrop()
		&& -1 !== this.Selection.DragDrop.Flag
		&& MouseEvent.ClickCount <= 1
		&& false === bTableBorder
		&& (nInDrawing < 0
			|| (nInDrawing === DRAWING_ARRAY_TYPE_BEHIND && true === bInText)
			|| (nInDrawing > -1
				&& (docpostype_DrawingObjects === this.CurPos.Type || (docpostype_HdrFtr === this.CurPos.Type && this.HdrFtr.CurHdrFtr && docpostype_DrawingObjects === this.HdrFtr.CurHdrFtr.Content.CurPos.Type))
				&& true === this.DrawingObjects.isSelectedText()
				&& null !== this.DrawingObjects.getMajorParaDrawing()
				&& this.DrawingObjects.getGraphicInfoUnderCursor(this.CurPage, X, Y).cursorType === "text"))
		&& true === this.CheckPosInSelection(X, Y, this.CurPage, undefined))
	{
		// Здесь мы сразу не начинаем перемещение текста. Его мы начинаем, курсор хотя бы немного изменит свою позицию,
		// это проверяется на MouseMove.
		// TODO: В ворде кроме изменения положения мыши, также запускается таймер для стартования drag-n-drop по времени,
		//       его можно здесь вставить.

		this.Selection.DragDrop.Flag = 1;
		this.Selection.DragDrop.Data = {X : X, Y : Y, PageNum : this.CurPage};
		return;
	}

    var bCheckHdrFtr = true;
	var bFootnotes   = false;
	var bEndnotes    = false;

	var nDocPosType = this.GetDocPosType();
    if (docpostype_HdrFtr === nDocPosType)
    {
        bCheckHdrFtr         = false;
        this.Selection.Start = true;
        this.Selection.Use   = true;
        if (false != this.HdrFtr.Selection_SetStart(X, Y, this.CurPage, MouseEvent, false))
            return;

        this.Selection.Start = false;
        this.Selection.Use   = false;
        this.DrawingDocument.ClearCachePages();
        this.DrawingDocument.FirePaint();
        this.DrawingDocument.EndTrackTable(null, true);
    }
	else
	{
		bFootnotes = this.Footnotes.CheckHitInFootnote(X, Y, this.CurPage);
		bEndnotes  = this.Endnotes.CheckHitInEndnote(X, Y, this.CurPage);
	}

	// На странице нет ничего из основной части документа
	if (-1 === this.Pages[this.CurPage].Pos || this.Pages[this.CurPage].Pos > this.Pages[this.CurPage].EndPos)
	{
		if (!bEndnotes && !this.Endnotes.IsEmptyPage(this.CurPage))
			bEndnotes = true;
		else if (!bFootnotes && !this.Footnotes.IsEmptyPage(this.CurPage))
			bFootnotes = true;
	}

    var PageMetrics = this.Get_PageContentStartPos(this.CurPage, this.Pages[this.CurPage].Pos);

	var oldDocPosType = this.GetDocPosType();
    // Проверяем, не попали ли мы в колонтитул (если мы попадаем в Flow-объект, то попадание в колонтитул не проверяем)
    if (true != bFlowTable && nInDrawing < 0 && true === bCheckHdrFtr && MouseEvent.ClickCount >= 2 && ( Y <= PageMetrics.Y || Y > PageMetrics.YLimit ))
    {
        // Если был селект, тогда убираем его
        if (true === this.Selection.Use)
            this.RemoveSelection();

        this.SetDocPosType(docpostype_HdrFtr);

        // Переходим к работе с колонтитулами
        MouseEvent.ClickCount = 1;
        this.HdrFtr.Selection_SetStart(X, Y, this.CurPage, MouseEvent, true);
        this.Interface_Update_HdrFtrPr();

        this.DrawingDocument.ClearCachePages();
        this.DrawingDocument.FirePaint();
        this.DrawingDocument.EndTrackTable(null, true);
    }
	else if (true !== bFlowTable && nInDrawing < 0 && true === bFootnotes)
	{
		this.RemoveSelection();
        this.Selection.Start = true;
        this.Selection.Use   = true;

		this.SetDocPosType(docpostype_Footnotes);
		this.Footnotes.StartSelection(X, Y, this.CurPage, MouseEvent);
	}
	else if (true !== bFlowTable && nInDrawing < 0 && true === bEndnotes)
	{
		this.RemoveSelection();
		this.Selection.Start = true;
		this.Selection.Use   = true;

		this.SetDocPosType(docpostype_Endnotes);
		this.Endnotes.StartSelection(X, Y, this.CurPage, MouseEvent);
	}
    else if (nInDrawing === DRAWING_ARRAY_TYPE_BEFORE || nInDrawing === DRAWING_ARRAY_TYPE_INLINE || ( false === bTableBorder && false === bInText && nInDrawing >= 0 ))
    {
        if (docpostype_DrawingObjects != this.CurPos.Type)
            this.RemoveSelection();

        // Прячем курсор
        this.DrawingDocument.TargetEnd();
        this.DrawingDocument.SetCurrentPage(this.CurPage);

        this.Selection.Use   = true;
        this.Selection.Start = true;
        this.Selection.Flag  = selectionflag_Common;
        this.SetDocPosType(docpostype_DrawingObjects);
        this.DrawingObjects.OnMouseDown(MouseEvent, X, Y, this.CurPage);
    }
    else
    {
        var bOldSelectionIsCommon = true;

        if (docpostype_DrawingObjects === this.CurPos.Type && (true != this.IsInDrawing(X, Y, this.CurPage) || ( nInDrawing === DRAWING_ARRAY_TYPE_BEHIND && true === bInText )))
        {
            this.DrawingObjects.OnMouseDown(MouseEvent, X, Y, this.CurPage);
            this.DrawingObjects.resetSelection();
            bOldSelectionIsCommon = false;
        }

        var ContentPos = this.Internal_GetContentPosByXY(X, Y);

        if (docpostype_Content != this.CurPos.Type)
        {
            this.SetDocPosType(docpostype_Content);
            this.CurPos.ContentPos = ContentPos;
            bOldSelectionIsCommon  = false;
        }

		var SelectionUse_old = this.Selection.Use;
		var Item             = this.Content[ContentPos];
		var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, this.CurPage);
		var bTableBorder     = null === Item.IsTableBorder(X, Y, ElementPageIndex) ? false : true;

        // Убираем селект, кроме случаев либо текущего параграфа, либо при движении границ внутри таблицы
        if (!(true === SelectionUse_old && true === MouseEvent.ShiftKey && true === bOldSelectionIsCommon))
        {
            if ((selectionflag_Common != this.Selection.Flag) || ( true === this.Selection.Use && MouseEvent.ClickCount <= 1 && true != bTableBorder ))
                this.RemoveSelection();
        }

        this.Selection.Use   = true;
        this.Selection.Start = true;
        this.Selection.Flag  = selectionflag_Common;

		if (true === SelectionUse_old && true === MouseEvent.ShiftKey && true === bOldSelectionIsCommon)
		{
			this.Selection_SetEnd(X, Y, {Type : AscCommon.g_mouse_event_type_up, ClickCount : 1});
			this.Selection.Use    = true;
			this.Selection.Start  = true;
			this.Selection.EndPos = ContentPos;
			this.Selection.Data   = null;
		}
		else
		{
			var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, this.CurPage);
			Item.Selection_SetStart(X, Y, ElementPageIndex, MouseEvent, bTableBorder);

			if (this.IsNumberingSelection())
			{
				// TODO: Можно сделать передвигание нумерации как в Word
				return;
			}

			Item.Selection_SetEnd(X, Y, ElementPageIndex, {
				Type       : AscCommon.g_mouse_event_type_move,
				ClickCount : 1
			}, bTableBorder);

			if (true !== bTableBorder)
			{
				this.Selection.Use      = true;
				this.Selection.StartPos = ContentPos;
				this.Selection.EndPos   = ContentPos;
				this.Selection.Data     = null;

				this.CurPos.ContentPos = ContentPos;

				if (type_Paragraph === Item.GetType() && true === MouseEvent.CtrlKey)
				{
					var oHyperlink   = Item.CheckHyperlink(X, Y, ElementPageIndex);
					var oPageRefLink = Item.CheckPageRefLink(X, Y, ElementPageIndex);
					if (null != oHyperlink)
					{
						this.Selection.Data = {
							Hyperlink : oHyperlink
						};
					}
					else if (null !== oPageRefLink)
					{
						this.Selection.Data = {
							PageRef : oPageRefLink
						};
					}
				}
			}
			else
			{
				this.Selection.Data = {
					TableBorder : true,
					Pos         : ContentPos,
					Selection   : SelectionUse_old
				};
			}
		}
	}

	//при переходе из колонтитула в контент(и обратно) необходимо скрывать иконку с/в
	var newDocPosType = this.GetDocPosType();
    if((docpostype_HdrFtr === newDocPosType && docpostype_Content === oldDocPosType) || (docpostype_Content === newDocPosType && docpostype_HdrFtr === oldDocPosType))
    {
		window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
    }
};
/**
 * Данная функция может использоваться как при движении, так и при окончательном выставлении селекта.
 * Если bEnd = true, тогда это конец селекта.
 * @param X
 * @param Y
 * @param MouseEvent
 */
CDocument.prototype.Selection_SetEnd = function(X, Y, MouseEvent)
{
	this.ResetWordSelection();

	// Работаем с колонтитулом
	if (docpostype_HdrFtr === this.CurPos.Type)
    {
        this.HdrFtr.Selection_SetEnd(X, Y, this.CurPage, MouseEvent);
        if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
        {
            if (true != this.DrawingObjects.isPolylineAddition())
                this.Selection.Start = false;
            else
                this.Selection.Start = true;
        }

        return;
    }
    else if (docpostype_DrawingObjects === this.CurPos.Type)
    {
        if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
        {
            this.DrawingObjects.OnMouseUp(MouseEvent, X, Y, this.CurPage);

            if (true != this.DrawingObjects.isPolylineAddition())
            {
                this.Selection.Start = false;
                this.Selection.Use   = true;
            }
            else
            {
                this.Selection.Start = true;
                this.Selection.Use   = true;
            }
        }
        else
        {
            this.DrawingObjects.OnMouseMove(MouseEvent, X, Y, this.CurPage);
        }
        return;
    }
	else if (docpostype_Footnotes === this.CurPos.Type)
	{
		this.Footnotes.EndSelection(X, Y, this.CurPage, MouseEvent);

		if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
			this.Selection.Start = false;

		return;
	}
	else if (docpostype_Endnotes === this.CurPos.Type)
	{
		this.Endnotes.EndSelection(X, Y, this.CurPage, MouseEvent);

		if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
			this.Selection.Start = false;

		return;
	}

    // Обрабатываем движение границы у таблиц
    if (true === this.IsMovingTableBorder())
    {
		var nPos = this.Selection.Data.Pos;

		// Убираем селект раньше, чтобы при создании точки в истории не сохранялось состояние передвижения границы таблицы
		if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
		{
			this.Selection.Start = false;
			this.Selection.Use   = this.Selection.Data.Selection;
			this.Selection.Data  = null;
		}

		var Item             = this.Content[nPos];
        var ElementPageIndex = this.private_GetElementPageIndexByXY(nPos, X, Y, this.CurPage);
        Item.Selection_SetEnd(X, Y, ElementPageIndex, MouseEvent, true);
        return;
    }

    if (false === this.Selection.Use)
        return;

    var ContentPos = this.Internal_GetContentPosByXY(X, Y);

    this.CurPos.ContentPos = ContentPos;
    var OldEndPos          = this.Selection.EndPos;
    this.Selection.EndPos  = ContentPos;

    // Удалим отметки о старом селекте
    if (OldEndPos < this.Selection.StartPos && OldEndPos < this.Selection.EndPos)
    {
        var TempLimit = Math.min(this.Selection.StartPos, this.Selection.EndPos);
        for (var Index = OldEndPos; Index < TempLimit; Index++)
        {
        	this.Content[Index].RemoveSelection();
        }
    }
    else if (OldEndPos > this.Selection.StartPos && OldEndPos > this.Selection.EndPos)
    {
        var TempLimit = Math.max(this.Selection.StartPos, this.Selection.EndPos);
        for (var Index = TempLimit + 1; Index <= OldEndPos; Index++)
        {
        	this.Content[Index].RemoveSelection();
        }
    }

    // Направление селекта: 1 - прямое, -1 - обратное, 0 - отмечен 1 элемент документа
    var Direction = ( ContentPos > this.Selection.StartPos ? 1 : ( ContentPos < this.Selection.StartPos ? -1 : 0 )  );

    if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
    	this.StopSelection();

    var Start, End;
    if (0 == Direction)
    {
        var Item             = this.Content[this.Selection.StartPos];
        var ElementPageIndex = this.private_GetElementPageIndexByXY(this.Selection.StartPos, X, Y, this.CurPage);
        Item.Selection_SetEnd(X, Y, ElementPageIndex, MouseEvent);

        if (this.IsNumberingSelection())
		{
			// Ничего не делаем
		}
        else if (false === Item.IsSelectionUse())
        {
            this.Selection.Use = false;

            if (this.IsInText(X, Y, this.CurPage))
			{
				if (null != this.Selection.Data && this.Selection.Data.Hyperlink)
				{
					var oHyperlink    = this.Selection.Data.Hyperlink;
					var sBookmarkName = oHyperlink.GetAnchor();
					var sValue        = oHyperlink.GetValue();

					if (oHyperlink.IsTopOfDocument())
					{
						this.MoveCursorToStartOfDocument();
					}
					else if (sValue)
					{
						this.Api.sync_HyperlinkClickCallback(sBookmarkName ? sValue + "#" + sBookmarkName : sValue);

						oHyperlink.SetVisited(true);
						for (var PageIdx = Item.Get_AbsolutePage(0); PageIdx < Item.Get_AbsolutePage(0) + Item.Get_PagesCount(); PageIdx++)
							this.DrawingDocument.OnRecalculatePage(PageIdx, this.Pages[PageIdx]);

						this.DrawingDocument.OnEndRecalculate(false, true);
					}
					else if (sBookmarkName)
					{
						var oBookmark = this.BookmarksManager.GetBookmarkByName(sBookmarkName);
						if (oBookmark)
							oBookmark[0].GoToBookmark();
					}
				}
				else if (null !== this.Selection.Data && this.Selection.Data.PageRef)
				{
					var oInstruction = this.Selection.Data.PageRef.GetInstruction();
					if (oInstruction && fieldtype_PAGEREF === oInstruction.GetType())
					{
						var oBookmark = this.BookmarksManager.GetBookmarkByName(oInstruction.GetBookmarkName());
						if (oBookmark)
							oBookmark[0].GoToBookmark();
					}
				}
			}
        }
        else
        {
            this.Selection.Use = true;
        }

        return;
    }
    else if (Direction > 0)
    {
        Start = this.Selection.StartPos;
        End   = this.Selection.EndPos;
    }
    else
    {
        End   = this.Selection.StartPos;
        Start = this.Selection.EndPos;
    }

    // TODO: Разрулить пустой селект
    // Чтобы не было эффекта, когда ничего не поселекчено, а при удалении соединяются параграфы

	for (var Index = Start; Index <= End; Index++)
	{
		var Item = this.Content[Index];
		Item.SetSelectionUse(true);

		switch (Index)
		{
			case Start:

				Item.SetSelectionToBeginEnd(Direction > 0 ? false : true, false);
				break;

			case End:

				Item.SetSelectionToBeginEnd(Direction > 0 ? true : false, true);
				break;

			default:

				Item.SelectAll(Direction);
				break;
		}
	}

    var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, this.CurPage);
    this.Content[ContentPos].Selection_SetEnd(X, Y, ElementPageIndex, MouseEvent);

    // Проверяем, чтобы у нас в селект не попали элементы, в которых не выделено ничего
    if (true === this.Content[End].IsSelectionEmpty() && true === this.CheckEmptyElementsOnSelection)
    {
        this.Content[End].RemoveSelection();
        End--;
    }

    if (Start != End && true === this.Content[Start].IsSelectionEmpty() && true === this.CheckEmptyElementsOnSelection)
    {
        this.Content[Start].RemoveSelection();
        Start++;
    }

    if (Direction > 0)
    {
        this.Selection.StartPos = Start;
        this.Selection.EndPos   = End;
    }
    else
    {
        this.Selection.StartPos = End;
        this.Selection.EndPos   = Start;
    }
};
/**
 * Получаем направление селекта
 * @returns {number} Возвращается направление селекта. 1 - нормальное направление, -1 - обратное
 */
CDocument.prototype.GetSelectDirection = function()
{
	var oStartPos = this.GetContentPosition(true, true);
	var oEndPos   = this.GetContentPosition(true, false);

	for (var nPos = 0, nLen = Math.min(oStartPos.length, oEndPos.length); nPos < nLen; ++nPos)
	{
		if (!oEndPos[nPos] || !oStartPos[nPos] || oStartPos[nPos].Class !== oEndPos[nPos].Class)
			return 1;

		if (oStartPos[nPos].Position < oEndPos[nPos].Position)
			return 1;
		else if (oStartPos[nPos].Position > oEndPos[nPos].Position)
			return -1;
	}

	return 1;
};
CDocument.prototype.IsMovingTableBorder = function()
{
	return this.Controller.IsMovingTableBorder();
};
/**
 * Проверяем попали ли мы в селект.
 * @param X
 * @param Y
 * @param PageAbs
 * @param NearPos
 * @returns {boolean}
 */
CDocument.prototype.CheckPosInSelection = function(X, Y, PageAbs, NearPos)
{
	return this.Controller.CheckPosInSelection(X, Y, PageAbs, NearPos);
};
/**
 * Выделяем все содержимое, в зависимости от текущего положения курсора.
 */
CDocument.prototype.SelectAll = function()
{
	if (this.IsFillingFormMode())
	{
		var oSelectedElementsInfo = this.GetSelectedElementsInfo();

		var oSdt = oSelectedElementsInfo.GetInlineLevelSdt();
		if (!oSdt)
			oSdt = oSelectedElementsInfo.GetBlockLevelSdt();

		if (oSdt)
		{
			oSdt.SelectContentControl();
			this.UpdateSelection();
			this.UpdateInterface();
		}

		return;
	}

	this.private_UpdateTargetForCollaboration();

	this.ResetWordSelection();
	this.Controller.SelectAll();

	// TODO: Пока делаем Start = true, чтобы при Ctrl+A не происходил переход к концу селекта, надо будет
	//       сделать по нормальному
	this.Selection.Start = true;
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateRulersState();
	this.Selection.Start = false;

	// Отдельно обрабатываем это событие, потому что внутри него идет проверка на this.Selection.Start !== true
	this.Document_UpdateCopyCutState();

	this.private_UpdateCursorXY(true, true);
};
/**
 * Выделяем текущее слово
 * @return {boolean}
 */
CDocument.prototype.SelectCurrentWord = function()
{
	if (this.IsSelectionUse() && !this.IsTextSelectionUse())
		return false;

	if (this.IsSelectionUse())
		this.RemoveSelection();

	var oParagraph = this.GetCurrentParagraph();
	if (!oParagraph)
		return false;

	oParagraph.SelectCurrentWord();

	if (this.IsSelectionEmpty())
	{
		this.RemoveSelection();
		return false;
	}

	return true;
};
/**
 * Выделяем все элементы в заданном диапазоне
 * @param {number} nStartPos
 * @param {number} nEndPos
 */
CDocument.prototype.SelectRange = function(nStartPos, nEndPos)
{
	this.RemoveSelection();

	this.DrawingDocument.SelectEnabled(true);
	this.DrawingDocument.TargetEnd();

	this.SetDocPosType(docpostype_Content);

	this.Selection.Use   = true;
	this.Selection.Start = false;
	this.Selection.Flag  = selectionflag_Common;

	this.Selection.StartPos = Math.max(0, Math.min(nStartPos, this.Content.length - 1));
	this.Selection.EndPos   = Math.max(this.Selection.StartPos, Math.min(nEndPos, this.Content.length - 1));

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].SelectAll();
	}

	// TODO: Пока делаем Start = true, чтобы при Ctrl+A не происходил переход к концу селекта, надо будет
	//       сделать по нормальному
	this.Selection.Start = true;
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateRulersState();
	this.Selection.Start = false;

	// Отдельно обрабатываем это событие, потому что внутри него идет проверка на this.Selection.Start !== true
	this.Document_UpdateCopyCutState();

	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.OnEndTextDrag = function(NearPos, bCopy)
{
    if (true === this.Comments.Is_Use())
    {
        this.SelectComment(null, false);
	    this.Api.sync_HideComment();
    }

    // Сначала нам надо проверить попадаем ли мы обратно в выделенный текст, если да, тогда ничего не делаем,
    // а если нет, тогда удаляем выделенный текст и вставляем его в заданное место.

    if (true === this.CheckPosInSelection(0, 0, 0, NearPos))
    {
        this.RemoveSelection();

        // Нам надо снять селект и поставить курсор в то место, где была ближайшая позиция
        var Paragraph = NearPos.Paragraph;
        Paragraph.Cursor_MoveToNearPos(NearPos);
        Paragraph.Document_SetThisElementCurrent(false);

        this.Document_UpdateSelectionState();
        this.Document_UpdateInterfaceState();
        this.Document_UpdateRulersState();
    }
    else
    {
    	var bCancelTrackMove = false;
    	var sPrevTrackModeId = null;
    	if (this.IsTrackRevisions())
		{
			var oInfo = this.GetSelectedElementsInfo({CheckAllSelection : true});
			if (!oInfo.HaveNotReviewedContent() && !oInfo.HaveAddedInReview() && oInfo.HaveRemovedInReview())
				return;

			if (oInfo.HaveRemovedInReview() || !oInfo.HaveNotReviewedContent())
				bCancelTrackMove = true;

			sPrevTrackModeId = this.private_CheckTrackMoveSelection(oInfo);
			if (sPrevTrackModeId)
			{
				this.SelectTrackMove(sPrevTrackModeId, false, false, false);
				this.TrackMoveRelocation = true;
			}
		}

        // Создаем сразу точку в истории, т.к. при выполнении функции GetSelectedContent нам надо, чтобы данная
        // точка уже набивалась изменениями. Если из-за совместного редактирования действие сделать невозможно будет,
        // тогда последнюю точку удаляем.
        this.StartAction(AscDFH.historydescription_Document_DragText);

        NearPos.Paragraph.Check_NearestPos(NearPos);

		var oDocPos = this.AnchorPositionToDocumentPosition(NearPos);
		this.TrackDocumentPositions([oDocPos]);

		if (!bCopy)
		{
			this.DragAndDropAction = true;
			this.TrackMoveId       = sPrevTrackModeId ? sPrevTrackModeId : (this.IsTrackRevisions() && !bCancelTrackMove ? this.TrackRevisionsManager.GetNewMoveId() : null);
		}
		else
		{
			this.TrackMoveId = null;
		}


		// Получим копию выделенной части документа, которую надо перенести в новое место, одновременно с этим
		// удаляем эту выделенную часть (если надо).

		var DocContent = this.GetSelectedContent(true);

		if (!DocContent.CanInsert(NearPos))
		{
			this.History.Remove_LastPoint();
			NearPos.Paragraph.Clear_NearestPosArray();

			this.DragAndDropAction   = false;
			this.TrackMoveId         = null;
			this.TrackMoveRelocation = false;

			this.FinalizeAction(false);
			return;
		}

		var Para = NearPos.Paragraph;

		// Нам нужно отдельно проверить локи для контент контролов, поэтому мы отключаем стандартную проверку
		// в функции IsSelectionLocked и отдельно проверяем вставляемую часть и место куда мы вставляем на наличие
		// залоченных контент контролов
		this.SetCheckContentControlsLock(false);

		var oSelectInfo = this.GetSelectedElementsInfo({CheckAllSelection : true});
		var arrSdts     = oSelectInfo.GetAllSdts();
		if (arrSdts.length > 0 && !bCopy)
		{
			for (var nIndex = 0, nCount = arrSdts.length; nIndex < nCount; ++nIndex)
			{
				var oSdt = arrSdts[nIndex];
				if ((!oSdt.CanBeDeleted() && oSdt.IsSelectedAll()) || (!oSdt.CanBeEdited() && !oSdt.IsSelectedAll()))
				{
					this.DragAndDropAction   = false;
					this.TrackMoveId         = null;
					this.TrackMoveRelocation = false;

					this.FinalizeAction();
					this.SetCheckContentControlsLock(true);
					return;
				}
			}
		}

		var isLocked = false;

		var oParaNearPos = Para.Get_ParaNearestPos(NearPos);
		var oLastClass   = oParaNearPos.Classes[oParaNearPos.Classes.length - 1];
		if (oLastClass instanceof ParaRun)
		{
			arrSdts = oLastClass.GetParentContentControls();
			if (arrSdts.length > 0 && !arrSdts[arrSdts.length - 1].CanBeEdited())
				isLocked = true;
		}

        // Если мы копируем, тогда не надо проверять выделенные параграфы, а если переносим, тогда проверяем
        var CheckChangesType = (true !== bCopy ? AscCommon.changestype_Document_Content : changestype_None);
        if (!isLocked && !this.IsSelectionLocked(CheckChangesType, {
                Type      : changestype_2_ElementsArray_and_Type,
                Elements  : [Para],
                CheckType : changestype_Paragraph_Content
            }, true))
        {
        	if (this.TrackMoveId && !this.TrackMoveRelocation)
			{
				var arrParagraphs = this.GetSelectedParagraphs();
				if (arrParagraphs.length > 0)
				{
					arrParagraphs[arrParagraphs.length - 1].AddTrackMoveMark(true, false, this.TrackMoveId);
					arrParagraphs[0].AddTrackMoveMark(true, true, this.TrackMoveId);
				}
			}

            // Если надо удаляем выделенную часть (пересчет отключаем на время удаления)
            if (true !== bCopy)
            {
                this.TurnOff_Recalculate();
                this.TurnOff_InterfaceEvents();
                this.Remove(1, false, false, true);
                this.TurnOn_Recalculate(false);
                this.TurnOn_InterfaceEvents(false);

                if (false === Para.IsUseInDocument())
                {
					this.DragAndDropAction   = false;
					this.TrackMoveId         = null;
					this.TrackMoveRelocation = false;

                    this.Document_Undo();
                    this.History.Clear_Redo();
					this.SetCheckContentControlsLock(true);
					this.FinalizeAction(false);
                    return;
                }
            }

            this.RemoveSelection(true);

			this.RefreshDocumentPositions([oDocPos]);
			var oTempNearPos = this.DocumentPositionToAnchorPosition(oDocPos);
			if (oTempNearPos)
				NearPos = oTempNearPos;

			this.CheckFormPlaceHolder = false;
			DocContent.SetNewCommentsGuid(bCopy);
			DocContent.Insert(NearPos, true);

			this.Recalculate();
            this.UpdateSelection();
            this.UpdateInterface();
            this.UpdateRulers();
			this.UpdateTracks();
			this.FinalizeAction();

			this.CheckFormPlaceHolder = true;
		}
        else
		{
			this.History.Remove_LastPoint();
			NearPos.Paragraph.Clear_NearestPosArray();
			this.FinalizeAction(false);
		}

		this.SetCheckContentControlsLock(true);

		this.DragAndDropAction   = false;
		this.TrackMoveId         = null;
		this.TrackMoveRelocation = false;
	}
};
/**
 * В данной функции мы получаем выделенную часть документа в формате класса CSelectedContent.
 * @param bUseHistory - нужна ли запись в историю созданных классов. (при drag-n-drop нужна, при копировании не нужна)
 * @param oPr {{SaveNumberingValues : boolean}} дополнительные настройки
 * @returns {AscCommonWord.CSelectedContent}
 */
CDocument.prototype.GetSelectedContent = function(bUseHistory, oPr)
{
	// При копировании нам не нужно, чтобы новые классы помечались как созданные в рецензировании, а при перетаскивании
	// нужно.
	var isTrack = this.IsTrackRevisions();
	var isLocalTrack = false;
	if (isTrack && !bUseHistory)
	{
		isLocalTrack = this.GetLocalTrackRevisions();
		this.SetLocalTrackRevisions(false);
	}

	var bNeedTurnOffTableId = g_oTableId.m_bTurnOff === false && true !== bUseHistory;
	if (!bUseHistory)
		History.TurnOff();

	if (bNeedTurnOffTableId)
	{
		g_oTableId.m_bTurnOff = true;
	}

	var oSelectedContent = new AscCommonWord.CSelectedContent();

	if (oPr)
	{
		if (undefined !== oPr.SaveNumberingValues)
			oSelectedContent.SetSaveNumberingValues(oPr.SaveNumberingValues);
	}

	oSelectedContent.SetMoveTrack(isTrack, this.TrackMoveId);
	this.Controller.GetSelectedContent(oSelectedContent);
	oSelectedContent.EndCollect(this);

	if (!bUseHistory)
		History.TurnOn();

	if (bNeedTurnOffTableId)
	{
		g_oTableId.m_bTurnOff = false;
	}

	if (false !== isLocalTrack)
		this.SetLocalTrackRevisions(isLocalTrack);

	return oSelectedContent;
};
CDocument.prototype.UpdateCursorType = function(X, Y, PageAbs, MouseEvent)
{
	if (null !== this.FullRecalc.Id && this.FullRecalc.PageIndex <= PageAbs)
		return;

	this.Api.sync_MouseMoveStartCallback();

	this.DrawingDocument.OnDrawContentControl(null, AscCommon.ContentControlTrack.Hover);

	if (undefined !== X && null !== X)
	{
		var nDocPosType = this.GetDocPosType();
		if (docpostype_HdrFtr === nDocPosType)
		{
			this.HeaderFooterController.UpdateCursorType(X, Y, PageAbs, MouseEvent);
		}
		else if (docpostype_DrawingObjects === nDocPosType)
		{
			this.DrawingsController.UpdateCursorType(X, Y, PageAbs, MouseEvent);
		}
		else
		{
			if (true === this.Footnotes.CheckHitInFootnote(X, Y, PageAbs))
				this.Footnotes.UpdateCursorType(X, Y, PageAbs, MouseEvent);
			else if (this.Endnotes.CheckHitInEndnote(X, Y, PageAbs))
				this.Endnotes.UpdateCursorType(X, Y, PageAbs, MouseEvent);
			else
				this.LogicDocumentController.UpdateCursorType(X, Y, PageAbs, MouseEvent);
		}
	}

	this.Api.sync_MouseMoveEndCallback();
};
CDocument.prototype.CloseAllWindowsPopups = function()
{
	// Закрываем всплывашки на MouseMove
	this.Api.sync_MouseMoveStartCallback();
	this.Api.sync_MouseMoveCallback(new AscCommon.CMouseMoveData());
	this.Api.sync_MouseMoveEndCallback();

	// Закрываем балун с рецензированием
	this.TrackRevisionsManager.BeginCollectChanges(false);
	this.TrackRevisionsManager.EndCollectChanges();

	// Закрываем балун с комментариями
	this.SelectComment(null, false);
	this.Api.sync_HideComment();
};
/**
 * Проверяем попадание в границу таблицы.
 * @param X
 * @param Y
 * @param PageIndex
 * @returns {?CTable} null - не попали, а если попали возвращается указатель на таблицу
 */
CDocument.prototype.IsTableBorder = function(X, Y, PageIndex)
{
	if (PageIndex >= this.Pages.length || PageIndex < 0)
		return null;

	if (docpostype_HdrFtr === this.GetDocPosType())
	{
		return this.HdrFtr.IsTableBorder(X, Y, PageIndex);
	}
	else
	{
		if (-1 != this.DrawingObjects.IsInDrawingObject(X, Y, PageIndex, this))
		{
			return null;
		}
		else if (true === this.Footnotes.CheckHitInFootnote(X, Y, PageIndex))
		{
			return this.Footnotes.IsTableBorder(X, Y, PageIndex);
		}
		else if (this.Endnotes.CheckHitInEndnote(X, Y, PageIndex))
		{
			return this.Endnotes.IsTableBorder(X, Y, PageIndex);
		}
		else
		{
			var ColumnsInfo      = {};
			var ElementPos       = this.Internal_GetContentPosByXY(X, Y, PageIndex, ColumnsInfo);
			var Element          = this.Content[ElementPos];
			var ElementPageIndex = this.private_GetElementPageIndex(ElementPos, PageIndex, ColumnsInfo.Column, ColumnsInfo.Column, ColumnsInfo.ColumnsCount);
			return Element.IsTableBorder(X, Y, ElementPageIndex);
		}
	}
};
/**
 * Проверяем, происходит ли сейчас выделение ячеек какой-либо таблицы
 * @returns {boolean}
 */
CDocument.prototype.IsTableCellSelection = function()
{
	return this.Controller.IsTableCellSelection();
};
/**
 * Проверяем, попали ли мы четко в текст (не лежащий в автофигуре)
 * @param X
 * @param Y
 * @param PageIndex
 * @returns {*}
 */
CDocument.prototype.IsInText = function(X, Y, PageIndex)
{
	if (PageIndex >= this.Pages.length || PageIndex < 0)
		return null;

	if (docpostype_HdrFtr === this.GetDocPosType())
	{
		return this.HdrFtr.IsInText(X, Y, PageIndex);
	}
	else
	{
		if (true === this.Footnotes.CheckHitInFootnote(X, Y, PageIndex))
		{
			return this.Footnotes.IsInText(X, Y, PageIndex);
		}
		else if (this.Endnotes.CheckHitInEndnote(X, Y, PageIndex))
		{
			return this.Endnotes.IsInText(X, Y, PageIndex);
		}
		else
		{
			var ContentPos       = this.Internal_GetContentPosByXY(X, Y, PageIndex);
			var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, PageIndex);
			var Item             = this.Content[ContentPos];
			return Item.IsInText(X, Y, ElementPageIndex);
		}
	}
};
CDocument.prototype.Get_ParentTextTransform = function()
{
	return null;
};
/**
 * Проверяем, попали ли мы в автофигуру данного DocumentContent
 * @param X
 * @param Y
 * @param PageIndex
 * @returns {*}
 */
CDocument.prototype.IsInDrawing = function(X, Y, PageIndex)
{
	if (docpostype_HdrFtr === this.GetDocPosType())
	{
		return this.HdrFtr.IsInDrawing(X, Y, PageIndex);
	}
	else
	{
		if (-1 != this.DrawingObjects.IsInDrawingObject(X, Y, this.CurPage, this))
		{
			return true;
		}
		else if (true === this.Footnotes.CheckHitInFootnote(X, Y, PageIndex))
		{
			return this.Footnotes.IsInDrawing(X, Y, PageIndex);
		}
		else if (this.Endnotes.CheckHitInEndnote(X, Y, PageIndex))
		{
			return this.Endnotes.IsInDrawing(X, Y, PageIndex);
		}
		else
		{
			var ContentPos       = this.Internal_GetContentPosByXY(X, Y, PageIndex);
			var Item             = this.Content[ContentPos];
			var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, PageIndex);
			return Item.IsInDrawing(X, Y, ElementPageIndex);
		}
	}
};
/**
 * Проверяем, попали ли мы в форму
 * @param X
 * @param Y
 * @param nPageAbs
 * @returns {boolean}
 */
CDocument.prototype.IsInForm = function(X, Y, nPageAbs)
{
	var oAnchor = this.Get_NearestPos(nPageAbs, X, Y);
	if (!oAnchor || !oAnchor.Paragraph)
		return false;

	var oClass = oAnchor.Paragraph.GetClassByPos(oAnchor.ContentPos);
	if (!oClass || !(oClass instanceof ParaRun))
		return false;

	var oForm = oClass.GetParentForm();
	if (!oForm)
		return false;

	let oMainForm = oForm.GetMainForm();
	return oMainForm.CheckHitInContentControlByXY(X, Y, nPageAbs);
};
/**
 * Проверяем, попали ли мы в контейнер
 * @param X
 * @param Y
 * @param nPageAbs
 * @returns {boolean}
 */
CDocument.prototype.IsInContentControl = function(X, Y, nPageAbs)
{
	var oAnchor = this.Get_NearestPos(nPageAbs, X, Y);
	if (!oAnchor || !oAnchor.Paragraph)
		return false;

	var oClass = oAnchor.Paragraph.GetClassByPos(oAnchor.ContentPos);
	if (!oClass || !(oClass instanceof ParaRun))
		return false;

	var arrContentControls = oClass.GetParentContentControls();
	for (var nIndex = 0, nCount = arrContentControls.length; nIndex < nCount; ++nIndex)
	{
		if (arrContentControls[nIndex].CheckHitInContentControlByXY(X, Y, nPageAbs))
			return true;
	}

	return false;
};
CDocument.prototype.IsUseInDocument = function(Id)
{
	if (!this.MainDocument)
		return false;
	
	if (undefined === Id || null === Id)
		return true;

	for (let nIndex = 0, nCount = this.GetElementsCount(); nIndex < nCount; ++nIndex)
	{
		if (Id === this.GetElement(nIndex).GetId())
			return true;
	}

	return false;
};
CDocument.prototype.OnKeyDown = function(e)
{
	// if (e.KeyCode < 34 || e.KeyCode > 40)
	// 	this.CloseAllWindowsPopups();

    var OldRecalcId = this.RecalcId;

    // Если мы только что расширяли документ двойным щелчком, то сохраняем это действие
    if (true === this.History.Is_ExtendDocumentToPos())
        this.History.ClearAdditional();

    // Сбрасываем текущий элемент в поиске
    if (this.SearchEngine.Count > 0)
        this.SearchEngine.ResetCurrent();

    var bUpdateSelection = true;
    var bRetValue        = keydownresult_PreventNothing;
    const bIsMacOs       = AscCommon.AscBrowser.isMacOs;

    var nShortcutAction = this.Api.getShortcut(e);
    switch (nShortcutAction)
	{
		case Asc.c_oAscDocumentShortcutType.InsertPageBreak:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, false, false))
			{
				this.StartAction(AscDFH.historydescription_Document_EnterButton);
				this.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
				this.FinalizeAction();
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.InsertLineBreak:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, false, this.IsFormFieldEditing()))
			{
				this.StartAction(AscDFH.historydescription_Document_EnterButton);
				this.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line));
				this.FinalizeAction();
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.InsertColumnBreak:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, false, false))
			{
				this.StartAction(AscDFH.historydescription_Document_EnterButton);
				this.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Column));
				this.FinalizeAction();
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.ResetChar:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, false))
			{
				this.StartAction(AscDFH.historydescription_Document_Shortcut_ClearFormatting);
				this.ClearParagraphFormatting(false, true);
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventNothing;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.NonBreakingSpace:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, this.IsFormFieldEditing()))
			{
				this.StartAction(AscDFH.historydescription_Document_Shortcut_AddNonBreakingSpace);
				this.DrawingDocument.TargetStart();
				this.DrawingDocument.TargetShow();
				this.AddToParagraph(new AscWord.CRunText(0x00A0));
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventNothing;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.ApplyHeading1:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			{
				this.StartAction(AscDFH.historydescription_Document_Shortcut_SetStyleHeading1);
				this.SetParagraphStyle("Heading 1");
				this.UpdateInterface();
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.ApplyHeading2:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			{
				this.StartAction(AscDFH.historydescription_Document_Shortcut_SetStyleHeading2);
				this.SetParagraphStyle("Heading 2");
				this.UpdateInterface();
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.ApplyHeading3:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Properties))
			{
				this.StartAction(AscDFH.historydescription_Document_Shortcut_SetStyleHeading3);
				this.SetParagraphStyle("Heading 3");
				this.UpdateInterface();
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.Strikeout:
		{
			var oTextPr = this.GetCalculatedTextPr();
			if (oTextPr)
			{
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
				{
					this.StartAction(AscDFH.historydescription_Document_SetTextStrikeoutHotKey);
					this.AddToParagraph(new ParaTextPr({Strikeout : oTextPr.Strikeout !== true}));
					this.UpdateInterface();
					this.FinalizeAction();
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.ShowAll:
		{
			var isShow = this.Api.get_ShowParaMarks();
			this.Api.put_ShowParaMarks(!isShow);
			this.Api.sync_ShowParaMarks();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.EditSelectAll:
		{
			this.SelectAll();
			bUpdateSelection = false;
			bRetValue        = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.Bold:
		{
			var oTextPr = this.GetCalculatedTextPr();
			if (oTextPr)
			{
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
				{
					this.StartAction(AscDFH.historydescription_Document_SetTextBoldHotKey);
					this.AddToParagraph(new ParaTextPr({Bold : oTextPr.Bold !== true}));
					this.UpdateInterface();
					this.FinalizeAction();
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.CopyFormat:
		{
			this.Document_Format_Copy();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.CopyrightSign:
		{
			this.private_AddSymbolByShortcut(0x00A9);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.InsertEndnoteNow:
		{
			this.AddEndnote();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.CenterPara:
		{
			this.private_ToggleParagraphAlignByHotkey(AscCommon.align_Center);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.EuroSign:
		{
			this.private_AddSymbolByShortcut(0x20AC);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.InsertFootnoteNow:
		{
			this.AddFootnote();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.Italic:
		{
			var oTextPr = this.GetCalculatedTextPr();
			if (oTextPr)
			{
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
				{
					this.StartAction(AscDFH.historydescription_Document_SetTextItalicHotKey);
					this.AddToParagraph(new ParaTextPr({Italic : oTextPr.Italic !== true}));
					this.UpdateInterface();
					this.FinalizeAction();
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.JustifyPara:
		{
			this.private_ToggleParagraphAlignByHotkey(AscCommon.align_Justify);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.InsertHyperlink:
		{
			if (true === this.CanAddHyperlink(false) && this.CanEdit())
				this.Api.sync_DialogAddHyperlink();

			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.ApplyListBullet:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			{
				this.StartAction(AscDFH.historydescription_Document_SetParagraphNumberingHotKey);
				
				let numObject = AscWord.GetNumberingObjectByDeprecatedTypes(0, 1);
				if (numObject)
					this.SetParagraphNumbering(numObject);
				
				this.UpdateInterface();
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.LeftPara:
		{
			this.private_ToggleParagraphAlignByHotkey(align_Left);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.Indent:
		{
			this.IncreaseIndent();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.UnIndent:
		{
			this.DecreaseIndent();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.PrintPreviewAndPrint:
		{
			this.DrawingDocument.m_oWordControl.m_oApi.onPrint();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.InsertPageNumber:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			{
				this.StartAction(AscDFH.historydescription_Document_AddPageNumHotKey);
				this.AddToParagraph(new AscWord.CRunPageNum());
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.RightPara:
		{
			this.private_ToggleParagraphAlignByHotkey(AscCommon.align_Right);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.RegisteredSign:
		{
			this.private_AddSymbolByShortcut(0x00AE);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.Save:
		{
			if (!this.IsViewMode())
				this.Api.asc_Save(false);

			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.TrademarkSign:
		{
			this.private_AddSymbolByShortcut(0x2122);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.Underline:
		{
			var oTextPr = this.GetCalculatedTextPr();
			if (oTextPr)
			{
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
				{
					this.StartAction(AscDFH.historydescription_Document_SetTextUnderlineHotKey);
					this.AddToParagraph(new ParaTextPr({Underline : oTextPr.Underline !== true}));
					this.UpdateInterface();
					this.FinalizeAction();
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.PasteFormat:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
			{
				this.StartAction(AscDFH.historydescription_Document_FormatPasteHotKey);
				this.Document_Format_Paste();
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.EditRedo:
		{
			if (this.CanEdit() || this.IsEditCommentsMode() || this.IsFillingFormMode())
				this.Document_Redo();

			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.EditUndo:
		{
			if ((this.CanEdit() || this.IsEditCommentsMode() || this.IsFillingFormMode()) && !this.IsViewModeInReview())
				this.Document_Undo();

			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.EmDash:
		{
			this.private_AddSymbolByShortcut(0x2014);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.EnDash:
		{
			this.private_AddSymbolByShortcut(0x2013);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.UpdateFields:
		{
			this.UpdateFields(true);
			bUpdateSelection = false;
			bRetValue        = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.InsertEquation:
		{
			var oSelectedInfo = this.GetSelectedElementsInfo();
			var oMath         = oSelectedInfo.GetMath();
			if (null === oMath)
				this.Api.asc_AddMath();

			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.Superscript:
		{
			var oTextPr = this.GetCalculatedTextPr();
			if (oTextPr)
			{
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
				{
					this.StartAction(AscDFH.historydescription_Document_SetTextVertAlignHotKey2);
					this.AddToParagraph(new ParaTextPr({VertAlign : oTextPr.VertAlign === AscCommon.vertalign_SuperScript ? AscCommon.vertalign_Baseline : AscCommon.vertalign_SuperScript}));
					this.UpdateInterface();
					this.FinalizeAction();
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.NonBreakingHyphen:
		{
			if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true))
			{
				this.StartAction(AscDFH.historydescription_Document_MinusButton);

				this.DrawingDocument.TargetStart();
				this.DrawingDocument.TargetShow();

				this.AddToParagraph(AscWord.CreateNonBreakingHyphen());
				this.FinalizeAction();
			}
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.SoftHyphen:
		{
			// TODO: Реализовать
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.HorizontalEllipsis:
		{
			this.private_AddSymbolByShortcut(0x2026);
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.Subscript:
		{
			var oTextPr = this.GetCalculatedTextPr();
			if (oTextPr)
			{
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
				{
					this.StartAction(AscDFH.historydescription_Document_SetTextVertAlignHotKey3);
					this.AddToParagraph(new ParaTextPr({VertAlign : oTextPr.VertAlign === AscCommon.vertalign_SubScript ? AscCommon.vertalign_Baseline : AscCommon.vertalign_SubScript}));
					this.UpdateInterface();
					this.FinalizeAction();
				}
				bRetValue = keydownresult_PreventAll;
			}
			break;
		}
		case Asc.c_oAscDocumentShortcutType.DecreaseFontSize:
		{
			this.Api.FontSizeOut();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		case Asc.c_oAscDocumentShortcutType.IncreaseFontSize:
		{
			this.Api.FontSizeIn();
			bRetValue = keydownresult_PreventAll;
			break;
		}
		default:
		{
			var oCustom = this.Api.getCustomShortcutAction(nShortcutAction);
			if (oCustom)
			{
				if (AscCommon.c_oAscCustomShortcutType.Symbol === oCustom.Type)
				{
					this.Api["asc_insertSymbol"](oCustom.Font, oCustom.CharCode);
				}

			}

			break;
		}
	}

    if (!nShortcutAction)
	{
		if (e.KeyCode === 8) // BackSpace
		{
			this.CheckSubFormBeforeRemove(-1);

			if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Remove, null, true, this.IsFormFieldEditing()))
			{
				this.StartAction(AscDFH.historydescription_Document_BackSpaceButton);
				this.Remove(-1, true, false, false, e.CtrlKey, true);
				this.FinalizeAction();
				this.private_UpdateCursorXY(true, true);
			}
			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 9) // Tab
		{
			if (this.IsFillingFormMode())
			{
				this.DrawingDocument.UpdateTargetFromPaint = true;
				this.ResetWordSelection();
				this.private_UpdateTargetForCollaboration();
				this.MoveToFillingForm(true !== e.ShiftKey);
				this.private_CheckCursorPosInFillingFormMode();
				this.CheckComplexFieldsInSelection();
			}
			else
			{
				var SelectedInfo = this.GetSelectedElementsInfo();

				if (null !== SelectedInfo.GetMath())
				{
					var ParaMath  = SelectedInfo.GetMath();
					var Paragraph = ParaMath.GetParagraph();
					if (Paragraph && false === this.Document_Is_SelectionLocked(changestype_None, {
						Type      : changestype_2_Element_and_Type,
						Element   : Paragraph,
						CheckType : changestype_Paragraph_Content
					}))
					{
						this.StartAction(AscDFH.historydescription_Document_AddTabToMath);
						ParaMath.HandleTab(!e.ShiftKey);
						this.Recalculate();
						this.FinalizeAction();
					}
				}
				else if (true === SelectedInfo.IsInTable() && true != e.CtrlKey)
				{
					this.MoveCursorToCell(true === e.ShiftKey ? false : true);
				}
				else if (true === SelectedInfo.IsDrawingObjSelected() && true != e.CtrlKey)
				{
					this.DrawingObjects.selectNextObject((e.ShiftKey === true ? -1 : 1));
				}
				else
				{
					if (true === SelectedInfo.IsMixedSelection())
					{
						if (true === e.ShiftKey)
							this.DecreaseIndent();
						else
							this.IncreaseIndent();
					}
					else
					{
						var Paragraph = SelectedInfo.GetParagraph();
						var ParaPr    = Paragraph ? Paragraph.Get_CompiledPr2(false).ParaPr : null;
						if (null != Paragraph && (true === Paragraph.IsCursorAtBegin() || true === Paragraph.IsSelectionFromStart()) && (undefined != Paragraph.GetNumPr() || (true != Paragraph.IsEmpty() && ParaPr.Tabs.Tabs.length <= 0)))
						{
							if (false === this.Document_Is_SelectionLocked(changestype_None, {
								Type      : changestype_2_Element_and_Type,
								Element   : Paragraph,
								CheckType : AscCommon.changestype_Paragraph_Properties
							}))
							{
								this.StartAction(AscDFH.historydescription_Document_MoveParagraphByTab);
								Paragraph.Add_Tab(e.ShiftKey);
								this.Recalculate();
								this.UpdateInterface();
								this.UpdateSelection();
								this.FinalizeAction();
							}
						}
						else if (false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content))
						{
							this.StartAction(AscDFH.historydescription_Document_AddTab);
							this.AddToParagraph(new AscWord.CRunTab());
							this.FinalizeAction();
						}
					}
				}
			}

			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 13) // Enter
		{
			let oSelectedInfo = this.GetSelectedElementsInfo();
			let inlineSdt;

			var Hyperlink = this.IsCursorInHyperlink(false);
			if (Hyperlink)
			{
				var sBookmarkName = Hyperlink.GetAnchor();
				var sValue        = Hyperlink.GetValue();

				if (Hyperlink.IsTopOfDocument())
				{
					this.MoveCursorToStartOfDocument();
				}
				else if (sValue)
				{
					this.Api.sync_HyperlinkClickCallback(sBookmarkName ? sValue + "#" + sBookmarkName : sValue);
					Hyperlink.SetVisited(true);

					// TODO: Пока сделаем так, потом надо будет переделать
					this.DrawingDocument.ClearCachePages();
					this.DrawingDocument.FirePaint();
				}
				else if (sBookmarkName)
				{
					var oBookmark = this.BookmarksManager.GetBookmarkByName(sBookmarkName);
					if (oBookmark)
						oBookmark[0].GoToBookmark();
				}
			}
			else if ((inlineSdt = oSelectedInfo.GetInlineLevelSdt()) && inlineSdt.IsForm() && inlineSdt.IsTextForm() && inlineSdt.IsMultiLineForm())
			{
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, this.IsFormFieldEditing()))
				{
					this.StartAction(AscDFH.historydescription_Document_EnterButton);
					this.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Line));
					this.Recalculate();
					this.FinalizeAction();
				}
			}
			else
			{
				var CheckType = AscCommon.changestype_Document_Content_Add;

				var bCanPerform = true;
				if ((oSelectedInfo.GetInlineLevelSdt() && !oSelectedInfo.IsSdtOverDrawing() && (!e.ShiftKey || e.CtrlKey)) || (oSelectedInfo.GetField() && oSelectedInfo.GetField().IsFillingForm()))
					bCanPerform = false;

				if (bCanPerform && (docpostype_DrawingObjects === this.CurPos.Type ||
					(docpostype_HdrFtr === this.CurPos.Type && null != this.HdrFtr.CurHdrFtr && docpostype_DrawingObjects === this.HdrFtr.CurHdrFtr.Content.CurPos.Type)))
				{
					var oTargetDocContent = this.DrawingObjects.getTargetDocContent();
					if (!oTargetDocContent)
					{
						var nRet    = this.DrawingObjects.handleEnter();
						bCanPerform = (nRet === 0);
					}

					if (this.DrawingObjects.selection && null !== this.DrawingObjects.selection.cropSelection)
						CheckType = AscCommon.changestype_Drawing_Props;
				}

				if (bCanPerform && false === this.Document_Is_SelectionLocked(CheckType, null, false, this.IsFormFieldEditing()))
				{
					this.StartAction(AscDFH.historydescription_Document_EnterButton);

					let oMath = oSelectedInfo.GetMath();
					if (oMath)
					{
						if (oMath.Is_InInnerContent())
						{
							oMath.Handle_AddNewLine();
							oMath.ProcessAutoCorrect();
						}
						else
						{
							oMath.ProcessAutoCorrect();
							this.AddNewParagraph();
						}
					}
					else
					{
						this.AddNewParagraph();
					}
					this.Recalculate();
					this.FinalizeAction();
				}
			}

			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 27) // Esc
		{
			// TODO: Когда вернем вызов этой функции на любое нажатие клавиши, тут убрать
			this.CloseAllWindowsPopups();

			// 1. Если начался drag-n-drop сбрасываем его.
			// 2. Если у нас сейчас происходит выделение маркером, тогда его отменяем
			// 3. Если у нас сейчас происходит форматирование по образцу, тогда его отменяем
			// 4. Если у нас выделена автофигура (в колонтитуле или документе), тогда снимаем выделение с нее
			// 5. Если мы просто находимся в колонтитуле (автофигура не выделена) выходим из колонтитула
			if (this.Api.isDrawTablePen || this.Api.isDrawTableErase)
			{
				this.Api.isDrawTablePen && this.Api.sync_TableDrawModeCallback(false);
				this.Api.isDrawTableErase && this.Api.sync_TableEraseModeCallback(false);
				this.UpdateCursorType(this.CurPos.RealX, this.CurPos.RealY, this.CurPage, new AscCommon.CMouseEventHandler());
			}
			else if (true === this.DrawingDocument.IsTrackText())
			{
				// Сбрасываем проверку Drag-n-Drop
				this.Selection.DragDrop.Flag = 0;
				this.Selection.DragDrop.Data = null;

				this.DrawingDocument.CancelTrackText();
			}
			else if (true === this.Api.isMarkerFormat)
			{
				this.Api.sync_MarkerFormatCallback(false);
				this.UpdateCursorType(this.CurPos.RealX, this.CurPos.RealY, this.CurPage, new AscCommon.CMouseEventHandler());
			}
			else if (this.Api.isFormatPainterOn())
			{
				this.Api.sync_PaintFormatCallback(c_oAscFormatPainterState.kOff);
				this.UpdateCursorType(this.CurPos.RealX, this.CurPos.RealY, this.CurPage, new AscCommon.CMouseEventHandler());
			}
			else if (this.Api.isStartAddShape)
			{
				this.Api.sync_StartAddShapeCallback(false);
				this.Api.sync_EndAddShape();
				this.DrawingObjects.endTrackNewShape();
				this.UpdateCursorType(this.CurPos.RealX, this.CurPos.RealY, this.CurPage, new AscCommon.CMouseEventHandler());
			}
			else if (this.Api.isInkDrawerOn())
			{
				this.Api.stopInkDrawer();
			}
			else if (this.IsFormFieldEditing())
			{
				this.EndFormEditing();
			}
			else if (docpostype_DrawingObjects === this.CurPos.Type || (docpostype_HdrFtr === this.CurPos.Type && null != this.HdrFtr.CurHdrFtr && docpostype_DrawingObjects === this.HdrFtr.CurHdrFtr.Content.CurPos.Type))
			{
				this.EndDrawingEditing();
			}
			else if (docpostype_HdrFtr == this.CurPos.Type)
			{
				this.EndHdrFtrEditing(true);
			}
			else if (this.Api.isEyedropperStarted())
			{
				this.Api.cancelEyedropper();
				this.UpdateCursorType(this.CurPos.RealX, this.CurPos.RealY, this.CurPage, new AscCommon.CMouseEventHandler());
			}
			if (window['AscCommon'].g_specialPasteHelper.showSpecialPasteButton)
			{
				window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
			}

			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 32) // Space
		{
			this.private_CheckForbiddenPlaceOnTextAdd();

			var oSelectedInfo = this.GetSelectedElementsInfo();
			var oMath         = oSelectedInfo.GetMath();
			var oInlineSdt    = oSelectedInfo.GetInlineLevelSdt();
			var oBlockSdt     = oSelectedInfo.GetBlockLevelSdt();

			var oCheckBox;

			if (oInlineSdt && oInlineSdt.IsCheckBox())
				oCheckBox = oInlineSdt;
			else if (oBlockSdt && oBlockSdt.IsCheckBox())
				oCheckBox = oBlockSdt;

			let isFormFieldEditing = this.IsFormFieldEditing();
			if (oCheckBox)
			{
				oCheckBox.SkipSpecialContentControlLock(true);
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, isFormFieldEditing))
				{
					this.StartAction(AscDFH.historydescription_Document_SpaceButton);
					oCheckBox.ToggleCheckBox();
					this.Recalculate();
					this.FinalizeAction();
				}
				oCheckBox.SkipSpecialContentControlLock(false);
			}
			else
			{
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, isFormFieldEditing))
				{
					this.StartAction(AscDFH.historydescription_Document_SpaceButton);

					// Если мы находимся в формуле, тогда пытаемся выполнить автозамену
					if (null !== oMath && true === oMath.Make_AutoCorrect())
					{
						// Ничего тут не делаем. Все делается в автозамене
					}
					else
					{
						this.DrawingDocument.TargetStart();
						this.DrawingDocument.TargetShow();

						this.CheckLanguageOnTextAdd = true;
						this.AddToParagraph(new AscWord.CRunSpace());
						this.CheckLanguageOnTextAdd = false;
					}
					this.FinalizeAction();
				}
			}

			bRetValue = keydownresult_PreventNothing;
		}
		else if (e.KeyCode === 33) // PgUp
		{
			if (true === e.AltKey)
			{
				var MouseEvent = new AscCommon.CMouseEventHandler();

				MouseEvent.ClickCount = 1;
				MouseEvent.Type       = AscCommon.g_mouse_event_type_down;

				this.CurPage--;

				if (this.CurPage < 0)
					this.CurPage = 0;

				this.Selection_SetStart(0, 0, MouseEvent);

				MouseEvent.Type = AscCommon.g_mouse_event_type_up;
				this.Selection_SetEnd(0, 0, MouseEvent);
			}
			else
			{
				if (docpostype_HdrFtr === this.CurPos.Type)
				{
					if (true === this.HdrFtr.GoTo_PrevHdrFtr())
					{
						this.Document_UpdateSelectionState();
						this.Document_UpdateInterfaceState();
					}
				}
				else
				{
					if (this.Controller !== this.LogicDocumentController)
					{
						this.RemoveSelection();
						this.SetDocPosType(docpostype_Content);
					}

					this.MoveCursorPageUp(true === e.ShiftKey, true === e.CtrlKey);
				}
			}

			this.private_CheckCursorPosInFillingFormMode();
			this.CheckComplexFieldsInSelection();
			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 34) // PgDn
		{
			if (true === e.AltKey)
			{
				var MouseEvent = new AscCommon.CMouseEventHandler();

				MouseEvent.ClickCount = 1;
				MouseEvent.Type       = AscCommon.g_mouse_event_type_down;

				this.CurPage++;

				// TODO: переделать данную проверку
				if (this.CurPage >= this.DrawingDocument.m_lPagesCount)
					this.CurPage = this.DrawingDocument.m_lPagesCount - 1;

				this.Selection_SetStart(0, 0, MouseEvent);

				MouseEvent.Type = AscCommon.g_mouse_event_type_up;
				this.Selection_SetEnd(0, 0, MouseEvent);
			}
			else
			{
				if (docpostype_HdrFtr === this.CurPos.Type)
				{
					if (true === this.HdrFtr.GoTo_NextHdrFtr())
					{
						this.Document_UpdateSelectionState();
						this.Document_UpdateInterfaceState();
					}
				}
				else
				{
					if (this.Controller !== this.LogicDocumentController)
					{
						this.RemoveSelection();
						this.SetDocPosType(docpostype_Content);
					}

					this.MoveCursorPageDown(true === e.ShiftKey, true === e.CtrlKey);
				}
			}

			this.private_CheckCursorPosInFillingFormMode();
			this.CheckComplexFieldsInSelection();
			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 35) // End
		{
			if (true === e.CtrlKey) // Ctrl + End - переход в конец документа
			{
				this.MoveCursorToEndPos(true === e.ShiftKey);
			}
			else // Переходим в конец строки
			{
				this.MoveCursorToEndOfLine(true === e.ShiftKey);
			}

			this.Document_UpdateInterfaceState();
			this.Document_UpdateRulersState();

			this.private_CheckCursorPosInFillingFormMode();
			this.CheckComplexFieldsInSelection();
			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 36) // Home
		{
			if (true === e.CtrlKey) // Ctrl + Home - переход в начало документа
			{
				this.MoveCursorToStartPos(true === e.ShiftKey);
			}
			else // Переходим в начало строки
			{
				this.MoveCursorToStartOfLine(true === e.ShiftKey);
			}

			this.Document_UpdateInterfaceState();
			this.Document_UpdateRulersState();

			this.private_CheckCursorPosInFillingFormMode();
			this.CheckComplexFieldsInSelection();
			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 37) // Left Arrow
		{
			// Чтобы при зажатой клавише курсор не пропадал
			if (true !== e.ShiftKey)
				this.DrawingDocument.TargetStart();

			this.DrawingDocument.UpdateTargetFromPaint = true;
			this.MoveCursorLeft(true === e.ShiftKey, true === e.CtrlKey);
			this.private_CheckCursorPosInFillingFormMode();
			this.CheckComplexFieldsInSelection();
			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 38) // Top Arrow
		{
			if (this.IsFillingFormMode())
			{
				var oSelectedInfo = this.GetSelectedElementsInfo();
				var oForm         = oSelectedInfo.GetForm();

				if (oForm && !oForm.IsComboBox() && !oForm.IsDropDownList())
					oForm = null;

				if (oForm)
					this.TurnComboBoxFormValue(oForm, false);
			}
			else
			{
				// TODO: Реализовать Ctrl + Up/ Ctrl + Shift + Up
				// Чтобы при зажатой клавише курсор не пропадал
				if (true !== e.ShiftKey)
					this.DrawingDocument.TargetStart();

				this.DrawingDocument.UpdateTargetFromPaint = true;
				this.MoveCursorUp(true === e.ShiftKey, true === e.CtrlKey);
				this.private_CheckCursorPosInFillingFormMode();
				this.CheckComplexFieldsInSelection();
			}

			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 39) // Right Arrow
		{
			// Чтобы при зажатой клавише курсор не пропадал
			if (true !== e.ShiftKey)
				this.DrawingDocument.TargetStart();

			this.DrawingDocument.UpdateTargetFromPaint = true;
			this.MoveCursorRight(true === e.ShiftKey, true === e.CtrlKey);
			this.private_CheckCursorPosInFillingFormMode();
			this.CheckComplexFieldsInSelection();
			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 40) // Bottom Arrow
		{
			if (this.IsFillingFormMode())
			{
				var oSelectedInfo = this.GetSelectedElementsInfo();
				var oForm         = oSelectedInfo.GetForm();

				if (oForm && !oForm.IsComboBox() && !oForm.IsDropDownList())
					oForm = null;

				if (oForm)
					this.TurnComboBoxFormValue(oForm, true);
			}
			else
			{
				// TODO: Реализовать Ctrl + Down/ Ctrl + Shift + Down
				// Чтобы при зажатой клавише курсор не пропадал
				if (true !== e.ShiftKey)
					this.DrawingDocument.TargetStart();

				this.DrawingDocument.UpdateTargetFromPaint = true;
				this.MoveCursorDown(true === e.ShiftKey, true === e.CtrlKey);
				this.private_CheckCursorPosInFillingFormMode();
				this.CheckComplexFieldsInSelection();
			}
			bRetValue = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 46) // Delete
		{
			if (true !== e.ShiftKey)
			{
				this.CheckSubFormBeforeRemove(1);

				if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Delete, null, true, this.IsFormFieldEditing()))
				{
					this.StartAction(AscDFH.historydescription_Document_DeleteButton);
					this.Remove(1, false, false, false, e.CtrlKey, true);
					this.FinalizeAction();
					this.private_UpdateCursorXY(true, true);
				}
				bRetValue = keydownresult_PreventAll;
			}
		}
		else if (e.KeyCode === 88 && true === e.AltKey && true !== e.ShiftKey && (bIsMacOs && true === e.IsCtrl() || !bIsMacOs)) //Alt + X on Windows or Alt + Ctrl + X on Mac
		{
			var arrParagraphs = this.GetSelectedParagraphs();
			if (arrParagraphs.length === 1)
			{
				var oParagraph = arrParagraphs[0];
				var ListForUnicode = [];
				var oSettings = {
					IsForMathPart: 0,
					textForUnicode: "",
					fFlagForUnicode: 0,
					bBreak: false,
					nDirection: 0
				};
				var elStart = oParagraph.Get_ParaContentPos(true, true);
				while (true)
				{
					var cur = oParagraph.Get_ClassesByPos(elStart);
					var oItemRun = cur[cur.length - 1];
					if (oItemRun === undefined || oItemRun.Selection === undefined || oItemRun === oItemParent) break;
					if ((oItemRun.Get_Type() === 49 || oItemRun.Get_Type() === 39) && (oItemRun.Parent === oItemParent || oItemParent == null))
					{
						var oItemParent = oItemRun.Parent;
						oItemRun.CollectTextToUnicode(ListForUnicode, oSettings);
						if (ListForUnicode.length > 6) break;
						if (oItemRun.Selection.EndPos === oItemRun.Content.length && oItemRun.Selection.Use === true && (oSettings.nDirection === 0 || oSettings.nDirection === 1))
						{
							oSettings.nDirection = 1;
							elStart.Data[elStart.Depth - 2]++;
						}
						else if (oItemRun.Selection.EndPos === 0 && oItemRun.Selection.Use === true && (oSettings.nDirection === 0 || oSettings.nDirection === -1))
						{
							oSettings.nDirection = -1;
							elStart.Data[elStart.Depth - 2]--;
						}
						else
							break;
					}
					else
					{
						if (oItemRun.Selection.Use === true && oItemRun.Selection.StartPos !== oItemRun.Selection.EndPos/* && oItemRun.Parent !== oItemParent*/)
							ListForUnicode.splice(0, ListForUnicode.length);
						break;
					}
				}
			
				for (var nPos = 0; nPos < ListForUnicode.length; nPos++)
				{
					if ((ListForUnicode[nPos].value === undefined || ListForUnicode.length > 6 || oSettings.bBreak === true)
						|| (ListForUnicode.length > 1 && !((ListForUnicode[nPos].value <= 0x46 && ListForUnicode[nPos].value >= 0x41) 
						|| (ListForUnicode[nPos].value <= 0x39 && ListForUnicode[nPos].value >= 0x30)
						|| (ListForUnicode[nPos].value <= 0x66 && ListForUnicode[nPos].value >= 0x61))))
					{
						ListForUnicode.splice(0, ListForUnicode.length);
					}
				}
				if (oSettings.nDirection === -1)
					ListForUnicode.reverse();
				var textAfterChange = "";
				if (ListForUnicode.length <= 6 && ListForUnicode.length !== 0)
				{
					if (ListForUnicode.length !== 1 && ListForUnicode.length <= 6)
					{
						textAfterChange = parseInt(oSettings.textForUnicode, 16);
						if (!isNaN(textAfterChange) && textAfterChange > 0x1F && textAfterChange < 0x110000 && !(textAfterChange >= 0xD800 && textAfterChange <= 0xDFFF))
						{
							if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, false))
							{
								this.StartAction(AscDFH.historydescription_Document_AddLetter);					
								if (oParagraph.LogicDocument.IsTrackRevisions() && ListForUnicode[0].oRun.ReviewType === 0)
								{
									var newDelRun = ListForUnicode[0].oRun.Copy();
									var newAddRun = ListForUnicode[0].oRun.Copy();
									newAddRun.ClearContent();
									newDelRun.ClearContent();
									for (var i = 0; i < ListForUnicode.length; i++)
										newDelRun.private_AddItemToRun(i, ListForUnicode[i].oRun.Content[ListForUnicode[i].currentPos]);
									
									if (oSettings.IsForMathPart === -1)
									{
										var oItem = new CMathText();
										oItem.add(textAfterChange);
										newAddRun.private_AddItemToRun(0, oItem);
									}
									else
									{
										newAddRun.private_AddItemToRun(0, new AscWord.CRunText(textAfterChange));
									}
									newDelRun.SetReviewType(reviewtype_Remove, true);
									newAddRun.SetReviewType(reviewtype_Add, true);
									var oParent = ListForUnicode[0].oRun.Get_Parent();
									var nMovePos = 0;
									for (var i = ListForUnicode.length - 1; i >= 0; i--)
										ListForUnicode[i].oRun.RemoveFromContent(ListForUnicode[i].currentPos, 1, true);
									
									if (ListForUnicode[ListForUnicode.length - 1].oRun === ListForUnicode[0].oRun && ListForUnicode[0].oRun.Content !== 0 && ListForUnicode[0].currentPos !== 0)
										nMovePos = 1;
									var oRunNew = ListForUnicode[ListForUnicode.length - 1].oRun.private_SplitRunInCurPos();
									var oPosItem = oParagraph.Get_PosByElement(ListForUnicode[ListForUnicode.length - 1].oRun);
									
									oParent.Add_ToContent(oPosItem.Data[oPosItem.Depth - 1] + nMovePos, newAddRun);
									oParent.Add_ToContent(oPosItem.Data[oPosItem.Depth - 1] + nMovePos, newDelRun);
								}
								else
								{
									var copyReviewInfo = ListForUnicode[0].oRun.GetReviewInfo().Copy();
									for (var i = ListForUnicode.length - 1; i >= 0; i--)
										ListForUnicode[i].oRun.RemoveFromContent(ListForUnicode[i].currentPos, 1, true);
									if (oSettings.IsForMathPart === -1)
									{
										var oItem = new CMathText();
										oItem.add(textAfterChange);
										ListForUnicode[0].oRun.private_AddItemToRun(ListForUnicode[0].currentPos, oItem);
									}
									else if (AscCommon.IsSpace(textAfterChange))
										ListForUnicode[0].oRun.private_AddItemToRun(ListForUnicode[0].currentPos, new AscWord.CRunSpace(textAfterChange));
									else
										ListForUnicode[0].oRun.private_AddItemToRun(ListForUnicode[0].currentPos, new AscWord.CRunText(textAfterChange));
									if (ListForUnicode[0].oRun.ReviewType === 2)
									{
										var oPosItem = oParagraph.Get_PosByElement(ListForUnicode[0].oRun);
										oPosItem.Data[oPosItem.Depth - 1]--;
										var oItemPrev = oParagraph.Get_ClassesByPos(oPosItem);
										oItemPrev = oItemPrev[oItemPrev.length - 1];
										var oValue;
										if (oItemPrev != null && (oItemPrev.Type === 49 || oItemPrev.Type === 39) && oItemPrev.Content != null && oItemPrev.Content.length !== 0)
										{
											oSettings.IsForMathPart === -1 ? oValue = oItemPrev.Content[oItemPrev.Content.length - 1].getCodeChr() : oValue = oItemPrev.Content[oItemPrev.Content.length - 1].Value;
											if (oValue === textAfterChange
												&& oItemPrev.ReviewType === 1 && ListForUnicode[0].oRun.ReviewType === 2
												&& this.CompareReviewInfo(oItemPrev.ReviewInfo, copyReviewInfo))
												oItemPrev.RemoveFromContent(oItemPrev.Content.length - 1, 1, true);
										}
									}
									ListForUnicode[0].oRun.Selection.Use = true;
									ListForUnicode[0].oRun.Selection.StartPos = ListForUnicode[0].currentPos;
									ListForUnicode[0].oRun.Selection.EndPos = ListForUnicode[0].currentPos + 1;
								}

								this.UpdateSelection();
								this.Recalculate();
								this.UpdateInterface();
								this.UpdateRulers();
								this.UpdateTracks();

								this.FinalizeAction();
							}
						}
					}
					else if (ListForUnicode.length === 1)
					{
						if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, false))
						{
							textAfterChange = AscCommon.IntToHex(ListForUnicode[0].value).toUpperCase();

							this.StartAction(AscDFH.historydescription_Document_AddLetter);
							if (oParagraph.LogicDocument.IsTrackRevisions() && ListForUnicode[0].oRun.ReviewType === 0)
							{
								if (oSettings.IsForMathPart === -1)
								{
									var newDelRun = ListForUnicode[0].oRun.Copy();
									var newAddRun = ListForUnicode[0].oRun.Copy();
									newAddRun.ClearContent();
									newDelRun.ClearContent();
									newDelRun.private_AddItemToRun(0, ListForUnicode[0].oRun.Content[ListForUnicode[0].currentPos]);
									for (var i = 0; i < textAfterChange.length; i++)
									{
										var oItem = new CMathText();
										oItem.add(textAfterChange.charCodeAt(i));
										newAddRun.private_AddItemToRun(i, oItem);
									}
									newDelRun.SetReviewType(reviewtype_Remove, true);
									newAddRun.SetReviewType(reviewtype_Add, true);
									var oParent = ListForUnicode[0].oRun.Get_Parent();
									var nMovePos = ListForUnicode[0].currentPos === 0 ? 0 : 1;
									ListForUnicode[0].oRun.RemoveFromContent(ListForUnicode[0].currentPos, 1, true);
									var oRunNew = ListForUnicode[0].oRun.private_SplitRunInCurPos();
									if (nMovePos === 0)
									{
										oParent.Add_ToContent(oParent.CurPos + nMovePos, newDelRun);
										oParent.Add_ToContent(oParent.CurPos + nMovePos, newAddRun);
									}
									else
									{
										oParent.Add_ToContent(oParent.CurPos + nMovePos, newAddRun);
										oParent.Add_ToContent(oParent.CurPos + nMovePos, newDelRun);
									}
								}
								else
								{
									var newDelRun = new ParaRun;
									var newAddRun = new ParaRun;
									newDelRun.private_AddItemToRun(0, ListForUnicode[0].oRun.Content[ListForUnicode[0].currentPos]);
									for (var i = 0; i < textAfterChange.length; i++)
									{
										newAddRun.private_AddItemToRun(i, new AscWord.CRunText(textAfterChange.charCodeAt(i)));
									}
									newDelRun.SetReviewType(reviewtype_Remove, true);
									newAddRun.SetReviewType(reviewtype_Add, true);
									ListForUnicode[0].oRun.RemoveFromContent(ListForUnicode[0].currentPos, 1, true);
									oParagraph.Add(newDelRun);
									oParagraph.Add(newAddRun);
								}
							}
							else
							{
								ListForUnicode[0].oRun.RemoveFromContent(ListForUnicode[0].currentPos, 1, true);
								for (var i = 0; i < textAfterChange.length; i++)
								{
									if (oSettings.IsForMathPart === -1)
									{
										var oItem = new CMathText();
										oItem.add(textAfterChange.charCodeAt(i));
										ListForUnicode[0].oRun.private_AddItemToRun(ListForUnicode[0].currentPos + i, oItem);
									}
									else
										ListForUnicode[0].oRun.private_AddItemToRun(ListForUnicode[0].currentPos + i, new AscWord.CRunText(textAfterChange.charCodeAt(i)));
								}
								
								ListForUnicode[0].oRun.Selection.Use = true;
								ListForUnicode[0].oRun.Selection.StartPos = ListForUnicode[0].currentPos;
								ListForUnicode[0].oRun.Selection.EndPos = ListForUnicode[0].currentPos + textAfterChange.length;
							}

							this.UpdateSelection();
							this.Recalculate();
							this.UpdateInterface();
							this.UpdateRulers();
							this.UpdateTracks();

							this.FinalizeAction();
						}
					}
				}
			}
		}
		else if ((e.KeyCode === 93 && !e.MacCmdKey)
			|| (AscCommon.AscBrowser.isOpera && (57351 === e.KeyCode))
			|| (e.KeyCode === 121 && true === e.ShiftKey)) // Shift + F10 - контекстное меню
		{
			var X_abs, Y_abs, oPosition, ConvertedPos;
			if (this.DrawingObjects.selectedObjects.length > 0)
			{
				oPosition    = this.DrawingObjects.getContextMenuPosition(this.CurPage);
				ConvertedPos = this.DrawingDocument.ConvertCoordsToCursorWR(oPosition.X, oPosition.Y, oPosition.PageIndex);
			}
			else
			{
				ConvertedPos = this.DrawingDocument.ConvertCoordsToCursorWR(this.TargetPos.X, this.TargetPos.Y, this.TargetPos.PageNum);
			}
			X_abs = ConvertedPos.X;
			Y_abs = ConvertedPos.Y;

			this.Api.sync_ContextMenuCallback({Type : Asc.c_oAscContextMenuTypes.Common, X_abs : X_abs, Y_abs : Y_abs});

			bUpdateSelection = false;
			bRetValue        = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 144) // Num Lock
		{
			// Ничего не делаем
			bUpdateSelection = false;
			bRetValue        = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 145) // Scroll Lock
		{
			// Ничего не делаем
			bUpdateSelection = false;
			bRetValue        = keydownresult_PreventAll;
		}
		else if (e.KeyCode === 12288) // Space
		{
			if (false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content, null, true, this.IsFormFieldEditing()))
			{
				this.StartAction(AscDFH.historydescription_Document_SpaceButton);

				this.DrawingDocument.TargetStart();
				this.DrawingDocument.TargetShow();

				this.CheckLanguageOnTextAdd = true;
				this.AddToParagraph(new AscWord.CRunSpace());
				this.CheckLanguageOnTextAdd = false;

				this.FinalizeAction();
			}

			bRetValue = keydownresult_PreventAll;
		}
	}

    // Если был пересчет, значит были изменения, а вместе с ними пересылается и новая позиция курсора
    if (bRetValue & keydownresult_PreventKeyPress && OldRecalcId === this.RecalcId)
        this.private_UpdateTargetForCollaboration();

    if (bRetValue & keydownflags_PreventKeyPress && true === bUpdateSelection)
        this.Document_UpdateSelectionState();

    return bRetValue;
};
CDocument.prototype.CompareReviewInfo = function(ReviewInfo1, ReviewInfo2)
{
	return (ReviewInfo1.UserName === ReviewInfo2.UserName && ReviewInfo1.UserId === ReviewInfo2.UserId
		&& (ReviewInfo1.DateTime + 500 >= ReviewInfo2.DateTime && ReviewInfo1.DateTime - 500 <= ReviewInfo2.DateTime));
};
CDocument.prototype.private_AddSymbolByShortcut = function(nCode)
{
	this.private_CheckForbiddenPlaceOnTextAdd();

	if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, this.IsFormFieldEditing()))
	{
		this.StartAction(AscDFH.historydescription_Document_AddEuroLetter);
		this.DrawingDocument.TargetStart();
		this.DrawingDocument.TargetShow();
		this.AddToParagraph(new AscWord.CRunText(nCode));
		this.FinalizeAction();
	}
};
CDocument.prototype.OnKeyPress = function(e)
{
	var Code;
	if (null != e.Which)
		Code = e.Which;
	else if (e.KeyCode)
		Code = e.KeyCode;
	else
		Code = 0;//special char

	if (Code > 0x20)
		return this.EnterText(Code);

	return false;
};
CDocument.prototype.EnterText = function(value)
{
	if (undefined === value
		|| null === value
		|| (Array.isArray(value) && !value.length))
		return false;

	let codePoints = typeof(value) === "string" ? value.codePointsArray() : value;

	this.private_CheckForbiddenPlaceOnTextAdd();

	if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_AddText, null, true, this.IsFormFieldEditing()))
		return false;

	this.StartAction(AscDFH.historydescription_Document_AddLetter);

	this.DrawingDocument.TargetStart();
	this.DrawingDocument.TargetShow();

	this.CheckLanguageOnTextAdd = true;

	if (Array.isArray(codePoints))
	{
		for (let index = 0, count = codePoints.length; index < count; ++index)
		{
			let codePoint = codePoints[index];
			this.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
		}
	}
	else
	{
		let codePoint = codePoints;
		this.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
	}

	this.CheckLanguageOnTextAdd = false;

	this.UpdateSelection();
	this.FinalizeAction();

	return true;
};
CDocument.prototype.CorrectEnterText = function(oldValue, newValue)
{
	if (undefined === oldValue
		|| null === oldValue
		|| (Array.isArray(oldValue) && !oldValue.length))
		return this.EnterText(newValue);

	let newCodePoints = typeof(newValue) === "string" ? newValue.codePointsArray() : newValue;
	let oldCodePoints = typeof(oldValue) === "string" ? oldValue.codePointsArray() : oldValue;

	if (this.IsSelectionUse())
		return false;

	if (!Array.isArray(oldCodePoints))
		oldCodePoints = [oldCodePoints];

	let paragraph = this.GetCurrentParagraph();
	if (!paragraph)
		return false;

	let contentPos = paragraph.GetContentPosition(false, false);
	let run, inRunPos;
	for (let index = contentPos.length - 1; index >= 0; --index)
	{
		if (contentPos[index].Class instanceof AscWord.CRun)
		{
			run      = contentPos[index].Class;
			inRunPos = contentPos[index].Position;
			break;
		}
	}

	if (!run)
		return false;

	if (!this.History.CheckAsYouTypeEnterText(run, inRunPos, oldCodePoints[oldCodePoints.length - 1]))
		return false;

	if (undefined === newCodePoints || null === newCodePoints)
		newCodePoints = [];
	else if (!Array.isArray(newCodePoints))
		newCodePoints = [newCodePoints];

	let oldText = "";
	for (let index = 0, count = oldCodePoints.length; index < count; ++index)
	{
		oldText += String.fromCodePoint(oldCodePoints[index]);
	}

	let state     = this.SaveDocumentState();
	let maxShifts = oldCodePoints.length;
	let selectedText;
	while (maxShifts >= 0)
	{
		this.MoveCursorLeft(true, false);
		selectedText = this.GetSelectedText(true);

		if (!selectedText || selectedText === oldText)
			break;

		maxShifts--;
	}

	if (selectedText !== oldText || this.IsSelectionLocked(AscCommon.changestype_Paragraph_AddText, null, true, this.IsFormFieldEditing()))
	{
		this.LoadDocumentState(state);
		return false;
	}

	this.StartAction(AscDFH.historydescription_Document_CorrectEnterText);

	this.DrawingDocument.TargetStart();
	this.DrawingDocument.TargetShow();

	this.Remove(1, true, false, true);

	for (let index = 0, count = newCodePoints.length; index < count; ++index)
	{
		let codePoint = newCodePoints[index];
		this.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
	}

	this.UpdateSelection();
	this.FinalizeAction();

	return true;
};
CDocument.prototype.OnMouseDown = function(e, X, Y, PageIndex)
{
	if (PageIndex < 0)
		return;

	this.private_UpdateTargetForCollaboration();

	// Сбрасываем проверку Drag-n-Drop
	this.Selection.DragDrop.Flag = 0;
	this.Selection.DragDrop.Data = null;

	// Сбрасываем текущий элемент в поиске
	if (this.SearchEngine.Count > 0)
		this.SearchEngine.ResetCurrent();

	this.CurPos.ResetCC();

	// Обработка правой кнопки мыши происходит на событии MouseUp
	if (AscCommon.g_mouse_button_right === e.Button)
		return;

	if (this.DrawTableMode.Draw || this.DrawTableMode.Erase)
	{
		this.DrawTableMode.Start  = true;
		this.DrawTableMode.StartX = X;
		this.DrawTableMode.StartY = Y;
		this.DrawTableMode.Page   = PageIndex;


		var arrTables = this.GetAllTablesOnPage(PageIndex);

		this.DrawTableMode.TablesOnPage = arrTables;

		var oElement     = null;
		var nMinDistance = null;
		var isInside     = false;
		for (var nTableIndex = 0, nTablesCount = arrTables.length; nTableIndex < nTablesCount; ++nTableIndex)
		{
			var oBounds = arrTables[nTableIndex].Table.GetPageBounds(arrTables[nTableIndex].Page);

			var nTempMin;
			var isTempInside = false;
			if (oBounds.Left < X && X < oBounds.Right)
			{
				if (oBounds.Top < Y && Y < oBounds.Bottom)
				{
					nTempMin     = Math.min(Math.abs(oBounds.Left - X), Math.abs(X - oBounds.Right), Math.abs(oBounds.Top - Y), Math.abs(Y - oBounds.Bottom));
					isTempInside = true;
				}
				else
				{
					nTempMin = Math.min(Math.abs(oBounds.Top - Y), Math.abs(Y - oBounds.Bottom));
				}
			}
			else
			{
				if (oBounds.Top < Y && Y < oBounds.Bottom)
				{
					nTempMin = Math.min(Math.abs(oBounds.Left - X), Math.abs(X - oBounds.Right));
				}
				else
				{
					nTempMin = Math.max(Math.min(Math.abs(oBounds.Top - Y), Math.abs(Y - oBounds.Bottom)), Math.min(Math.abs(oBounds.Left - X), Math.abs(X - oBounds.Right)));
				}
			}

			if (null == nMinDistance || (nTempMin < nMinDistance && (!isInside || isTempInside)))
			{
				oElement     = arrTables[nTableIndex].Table;
				nMinDistance = nTempMin;

				if (isTempInside)
					isInside = true;
			}
		}

		if (this.DrawTableMode.Draw && !isInside && nMinDistance > 5)
			oElement = null;

		if (oElement)
			this.DrawTableMode.Table = oElement;

		this.DrawTableMode.UpdateTablePages();

		return;
	}

	// Если мы только что расширяли документ двойным щелчком, то отменяем это действие
	if (true === this.History.Is_ExtendDocumentToPos())
	{
		this.Document_Undo();

		// Заглушка на случай "неудачного" пересчета, когда после него страниц меньше, чем ту, по которое мы кликаем
		if (PageIndex >= this.Pages.length)
		{
			this.RemoveSelection();
			return;
		}
	}

	var OldCurPage = this.CurPage;
	this.CurPage   = PageIndex;

	if ((this.Api.isStartAddShape || this.Api.isInkDrawerOn()) && (docpostype_HdrFtr !== this.CurPos.Type || null !== this.HdrFtr.CurHdrFtr))
	{
		if (docpostype_HdrFtr !== this.CurPos.Type)
		{
			this.SetDocPosType(docpostype_DrawingObjects);
			this.Selection.Use   = true;
			this.Selection.Start = true;
		}
		else
		{
			this.Selection.Use   = true;
			this.Selection.Start = true;

			this.HdrFtr.WaitMouseDown = false;
			var CurHdrFtr             = this.HdrFtr.CurHdrFtr;
			var DocContent            = CurHdrFtr.Content;

			DocContent.SetDocPosType(docpostype_DrawingObjects);
			DocContent.Selection.Use   = true;
			DocContent.Selection.Start = true;
		}

		if (this.Api.isStartAddShape && !this.DrawingObjects.isPolylineAddition())
			this.DrawingObjects.startAddShape(this.Api.addShapePreset);

		this.DrawingObjects.OnMouseDown(e, X, Y, this.CurPage);
	}
	else
	{
		if (true === e.ShiftKey &&
			( (docpostype_DrawingObjects !== this.CurPos.Type && !(docpostype_HdrFtr === this.CurPos.Type && this.HdrFtr.CurHdrFtr && this.HdrFtr.CurHdrFtr.Content.CurPos.Type === docpostype_DrawingObjects))
			|| true === this.DrawingObjects.checkTextObject(X, Y, PageIndex) ))
		{
			if (true === this.IsSelectionUse())
				this.Selection.Start = false;
			else
				this.StartSelectionFromCurPos();

			this.Selection_SetEnd(X, Y, e);
			this.Document_UpdateSelectionState();
			return;
		}

		if (!this.IsFillingFormMode())
		{
			var oSelectedContent = this.GetSelectedElementsInfo();
			var oInlineSdt       = oSelectedContent.GetInlineLevelSdt();
			var oBlockSdt        = oSelectedContent.GetBlockLevelSdt();

			if ((oInlineSdt && oInlineSdt.IsForm() && oInlineSdt.IsCheckBox()) || (oBlockSdt && oBlockSdt.IsForm() && oBlockSdt.IsCheckBox()))
				this.CurPos.SetCC((oInlineSdt && oInlineSdt.IsForm() && oInlineSdt.IsCheckBox()) ? oInlineSdt : oBlockSdt);
		}

		this.Selection_SetStart(X, Y, e);
		
		var oSelectedContent = this.GetSelectedElementsInfo();
		var oInlineSdt       = oSelectedContent.GetInlineLevelSdt();
		var oBlockSdt        = oSelectedContent.GetBlockLevelSdt();

		if ((oInlineSdt && oInlineSdt.IsCheckBox()) || (oBlockSdt && oBlockSdt.IsCheckBox()))
		{
			var oCC = (oInlineSdt && oInlineSdt.IsCheckBox()) ? oInlineSdt : oBlockSdt;
			if (oCC.CheckHitInContentControlByXY(X, Y, PageIndex) && (!oCC.IsForm() || this.IsFillingFormMode() || oCC === this.CurPos.CC))
			{
				this.CurPos.SetCC(oCC);
				oCC.SkipSpecialContentControlLock(true);
				if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, this.IsFillingFormMode() && this.CheckOFormUserMaster(oCC)))
				{
					this.RemoveTextSelection();
					this.StartAction();
					oCC.ToggleCheckBox();
					this.Recalculate();
					this.UpdateTracks();
					this.FinalizeAction();
				}
				oCC.SkipSpecialContentControlLock(false);
			}
			this.UpdateSelection();
		}

		if (this.IsFillingFormMode() && (oBlockSdt || oInlineSdt))
		{
			this.CurPos.SetCC(oInlineSdt ? oInlineSdt : oBlockSdt);
		}

		if (e.ClickCount <= 1 && 1 !== this.Selection.DragDrop.Flag)
		{
			this.private_UpdateSelectionOnMouseEvent(X, Y, this.CurPage, e);
		}
	}
};
CDocument.prototype.OnMouseUp = function(e, X, Y, PageIndex)
{
	if (PageIndex < 0)
		return;

	if (this.IsFillingFormMode() && this.CurPos.IsInCC() && !this.CurPos.CheckHitInCC(X, Y, PageIndex))
	{
		var oCorrectedPos = this.CurPos.CorrectXYToHitInCC(X, Y, PageIndex);
		if (!oCorrectedPos)
		{
			if (this.Selection.Start)
				this.StopSelection();

			return;
		}

		X = oCorrectedPos.X;
		Y = oCorrectedPos.Y;
	}


	if (this.DrawTableMode.Draw || this.DrawTableMode.Erase)
	{
		if (!this.DrawTableMode.Start)
			return;

		this.DrawTableMode.Start = false;

		if (PageIndex !== this.DrawTableMode.Page)
			return;

		this.DrawTableMode.EndX = X;
		this.DrawTableMode.EndY = Y;
		this.DrawTableMode.UpdateTablePages();

		this.DrawTable();

		this.DrawTableMode.StartX = -1;
		this.DrawTableMode.StartY = -1;
		this.DrawTableMode.EndX   = -1;
		this.DrawTableMode.EndY   = -1;
		this.DrawTableMode.Page   = -1;
		this.DrawTableMode.Table  = null;

		this.DrawingDocument.OnUpdateOverlay();
		return;
	}

	this.private_UpdateTargetForCollaboration();

	if (1 === this.Selection.DragDrop.Flag)
	{
		this.Selection.DragDrop.Flag = -1;
		var OldCurPage               = this.CurPage;
		this.CurPage                 = this.Selection.DragDrop.Data.PageNum;
		this.Selection_SetStart(this.Selection.DragDrop.Data.X, this.Selection.DragDrop.Data.Y, e);
		this.Selection.DragDrop.Flag = 0;
		this.Selection.DragDrop.Data = null;
	}

	if (AscCommon.g_mouse_button_right === e.Button)
	{
		return this.private_HandleMouseRightClick(X, Y, PageIndex);
	}
	else if (AscCommon.g_mouse_button_left === e.Button)
	{
		if (true === this.Comments.Is_Use())
		{
			// Проверяем не попали ли мы в комментарий
			var arrComments = this.Comments.GetByXY(PageIndex, X, Y);

			var CommentsX     = null;
			var CommentsY     = null;
			var arrCommentsId = [];

			for (var nCommentIndex = 0, nCommentsCount = arrComments.length; nCommentIndex < nCommentsCount; ++nCommentIndex)
			{
				var Comment = arrComments[nCommentIndex];
				if (null != Comment && (this.Comments.IsUseSolved() || !Comment.IsSolved()))
				{
					if (null === CommentsX && Comment.GetRangeStart())
					{
						var oAnchorPoint = this.private_GetCommentWorldAnchorPoint(Comment);
						this.SelectComment(Comment.Get_Id(), false);

						CommentsX = oAnchorPoint.X;
						CommentsY = oAnchorPoint.Y;
					}

					arrCommentsId.push(Comment.Get_Id());
				}
			}

			if (null !== CommentsX && null !== CommentsY && arrCommentsId.length > 0)
			{
				this.Api.sync_ShowComment(arrCommentsId, CommentsX, CommentsY);
			}
			else
			{
				this.SelectComment(null, false);
				this.Api.sync_HideComment();
			}
		}
	}

	var isUpdateTarget = true;
	if (true === this.Selection.Start)
	{
		this.CurPage         = PageIndex;
		this.Selection.Start = false;
		this.Selection_SetEnd(X, Y, e);
		this.CheckComplexFieldsInSelection();
		isUpdateTarget = this.private_UpdateSelectionOnMouseEvent(X, Y, this.CurPage, e);

		if (this.Api.isFormatPainterOn())
		{
			if (false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content))
			{
				this.StartAction(AscDFH.historydescription_Document_FormatPasteHotKey2);
				this.Document_Format_Paste();
				this.FinalizeAction();
			}

			if (this.Api.canTurnOffFormatPainter())
				this.Api.sync_PaintFormatCallback(c_oAscFormatPainterState.kOff);
		}
		else if (true === this.Api.isMarkerFormat && true === this.IsTextSelectionUse())
		{
			if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_TextProperties))
			{
				this.StartAction(AscDFH.historydescription_Document_SetTextHighlight2);
				var ParaItem = null;
				if (this.HighlightColor != highlight_None)
				{
					var TextPr = this.GetCalculatedTextPr();
					if ("undefined" === typeof( TextPr.HighLight ) || null === TextPr.HighLight || highlight_None === TextPr.HighLight ||
						this.HighlightColor.r != TextPr.HighLight.r || this.HighlightColor.g != TextPr.HighLight.g || this.HighlightColor.b != TextPr.HighLight.b)
						ParaItem = new ParaTextPr({HighLight : this.HighlightColor});
					else
						ParaItem = new ParaTextPr({HighLight : highlight_None});
				}
				else
				{
					ParaItem = new ParaTextPr({HighLight : this.HighlightColor});
				}

				this.AddToParagraph(ParaItem);
				this.UpdateSelection();
				this.Recalculate();
				this.FinalizeAction();

				this.MoveCursorToXY(X, Y, false);
				this.Api.sync_MarkerFormatCallback(true);
			}
		}
	}

	this.private_CheckCursorPosInFillingFormMode();
	this.private_UpdateCursorXY(true, true, isUpdateTarget);
	
	this.UpdateSelection();
	this.UpdateInterface();
	this.UpdateRulers();
};
CDocument.prototype.OnMouseMove = function(e, X, Y, PageIndex)
{
	if (PageIndex < 0)
		return;

	if (this.DrawTableMode.Start
		&& (PageIndex === this.DrawTableMode.Page)
		&& (this.DrawTableMode.Draw || this.DrawTableMode.Erase))
	{
		this.DrawTableMode.EndX = X;
		this.DrawTableMode.EndY = Y;
		this.DrawTableMode.CheckSelectedTable();
		this.DrawTableMode.UpdateTablePages();
	}


	if (true === this.Selection.Start)
		this.private_UpdateTargetForCollaboration();

	this.UpdateCursorType(X, Y, PageIndex, e);
	this.CollaborativeEditing.Check_ForeignCursorsLabels(X, Y, PageIndex);

	if (this.DrawTableMode.Draw || this.DrawTableMode.Erase)
	{
	    if (this.DrawTableMode.Start)
            this.DrawingDocument.OnUpdateOverlay();

		return;
	}

	if (1 === this.Selection.DragDrop.Flag)
	{
		// Если курсор не изменил позицию, тогда ничего не делаем, а если изменил, тогда стартуем Drag-n-Drop
		if (Math.abs(this.Selection.DragDrop.Data.X - X) > 0.001 || Math.abs(this.Selection.DragDrop.Data.Y - Y) > 0.001)
		{
			this.Selection.DragDrop.Flag = 0;
			this.Selection.DragDrop.Data = null;

			// Вызываем стандартное событие mouseMove, чтобы сбросить различные подсказки, если они были
			this.Api.sync_MouseMoveStartCallback();
			this.Api.sync_MouseMoveCallback(new AscCommon.CMouseMoveData());
			this.Api.sync_MouseMoveEndCallback();

			this.DrawingDocument.StartTrackText();
		}

		return;
	}

	if (this.IsFillingFormMode() && this.CurPos.IsInCC() && !this.CurPos.CheckHitInCC(X, Y, PageIndex))
	{
		var oCorrectedPos = this.CurPos.CorrectXYToHitInCC(X, Y, PageIndex);
		if (!oCorrectedPos)
			return;

		X = oCorrectedPos.X;
		Y = oCorrectedPos.Y;
	}

	if (true === this.Selection.Use && true === this.Selection.Start)
	{
		this.CurPage = PageIndex;
		this.Selection_SetEnd(X, Y, e);
		this.private_UpdateSelectionOnMouseEvent(X, Y, this.CurPage, e);
	}
};
CDocument.prototype.private_HandleMouseRightClick = function(X, Y, nPageIndex)
{
	if (this.IsStartSelection())
		return;

	if (this.private_HandleMouseRightClickOnTableTrack(X, Y, nPageIndex))
		return;

	if (this.private_HandleMouseRightClickOnHeaderFooter(X, Y, nPageIndex))
		return;

	if (!this.CheckPosInSelection(X, Y, nPageIndex, undefined))
	{
		this.CurPage = nPageIndex;

		var oMouseEvent = {
			// TODO : Если в MouseEvent будет использоваться что-то кроме ClickCount, Type и CtrlKey, добавить здесь
			ClickCount : 1,
			Type       : AscCommon.g_mouse_event_type_down,
			CtrlKey    : false,
			Button     : AscCommon.g_mouse_button_right
		};

		this.Selection_SetStart(X, Y, oMouseEvent);
		oMouseEvent.Type = AscCommon.g_mouse_event_type_up;
		this.Selection_SetEnd(X, Y, oMouseEvent);

		this.UpdateSelection();
		this.UpdateRulers();
		this.UpdateInterface();
	}

	var oAppCoords = this.DrawingDocument.ConvertCoordsToCursorWR(X, Y, nPageIndex);

	this.Api.sync_ContextMenuCallback({
		Type  : Asc.c_oAscContextMenuTypes.Common,
		X_abs : oAppCoords.X,
		Y_abs : oAppCoords.Y
	});

	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.private_HandleMouseRightClickOnTableTrack = function(X, Y, nPageIndex)
{
	if (this.DrawingDocument.IsCursorInTableCur(X, Y, nPageIndex))
	{
		var oTable = this.DrawingDocument.TableOutlineDr.TableOutline.Table;
		oTable.SelectAll();
		oTable.Document_SetThisElementCurrent(false);

		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();

		var oAppCoords = this.DrawingDocument.ConvertCoordsToCursorWR(X, Y, nPageIndex);

		this.Api.sync_ContextMenuCallback({
			Type  : Asc.c_oAscContextMenuTypes.Common,
			X_abs : oAppCoords.X,
			Y_abs : oAppCoords.Y
		});

		return true;
	}

	return false;
};
CDocument.prototype.private_HandleMouseRightClickOnHeaderFooter = function(X, Y, nPageIndex)
{
	if (docpostype_HdrFtr !== this.GetDocPosType()
		&& -1 === this.DrawingObjects.IsInDrawingObject(X, Y, nPageIndex, this)
		&& !this.DrawingObjects.getTableByXY(X, Y, nPageIndex, this))
	{
		var oPageMetrics = this.Get_PageContentStartPos(nPageIndex, this.Pages[nPageIndex].Pos);

		if (Y <= oPageMetrics.Y || Y > oPageMetrics.YLimit)
		{
			var oAppCoords = this.DrawingDocument.ConvertCoordsToCursorWR(X, Y, nPageIndex);

			this.Api.sync_ContextMenuCallback({
				Type    : Asc.c_oAscContextMenuTypes.ChangeHdrFtr,
				X_abs   : oAppCoords.X,
				Y_abs   : oAppCoords.Y,
				Header  : Y <= oPageMetrics.Y,
				PageNum : nPageIndex
			});

			return true;
		}
	}

	return false;
};
CDocument.prototype.private_UpdateSelectionOnMouseEvent = function(nX, nY, nCurPage, oMouseEvent)
{
	var isHitHdrFtr = docpostype_HdrFtr !== this.GetDocPosType() && oMouseEvent && AscCommon.g_mouse_event_type_move !== oMouseEvent.Type ? this.private_IsHitInHdrFtr(nY, nCurPage) : false;
	if (isHitHdrFtr)
		this.Api.asc_LockScrollToTarget(true);

	// TODO: По логике это не нужно (убрать на develop)
	if (oMouseEvent && AscCommon.g_mouse_event_type_down === oMouseEvent.Type)
		this.RecalculateCurPos();

	this.UpdateSelection();

	if (isHitHdrFtr)
		this.Api.asc_LockScrollToTarget(false);

	return !isHitHdrFtr;
};
CDocument.prototype.private_IsHitInHdrFtr = function(nY, nCurPage)
{
	if (undefined === nCurPage)
		nCurPage = this.CurPage;

	if (!this.Pages[nCurPage])
		return false;

	var oPageMetrics = this.Get_PageContentStartPos(nCurPage, this.Pages[nCurPage].Pos);

	return (nY <= oPageMetrics.Y || nY > oPageMetrics.YLimit);
};
CDocument.prototype.private_CheckForbiddenPlaceOnTextAdd = function()
{
	let isFormFieldEditing = this.IsFormFieldEditing();
	let oSelectedInfo      = this.GetSelectedElementsInfo();

	var oInlineSdt = oSelectedInfo.GetInlineLevelSdt();
	var oBlockSdt  = oSelectedInfo.GetBlockLevelSdt();

	var oCheckBox;
	if (oInlineSdt && oInlineSdt.IsCheckBox())
		oCheckBox = oInlineSdt;
	else if (oBlockSdt && oBlockSdt.IsCheckBox())
		oCheckBox = oBlockSdt;

	if (!isFormFieldEditing && oCheckBox)
	{
		this.RemoveSelection();
		oCheckBox.MoveCursorOutsideForm(oCheckBox.IsCursorAtBegin());
	}
};
/**
 * Проверяем будет ли добавление текста на ивенте KeyDown
 * @param e
 * @returns {Number[]} Массив юникодных значений
 */
CDocument.prototype.GetAddedTextOnKeyDown = function(e)
{
	if (e.KeyCode === 32) // Space
	{
		var oSelectedInfo = this.GetSelectedElementsInfo();
		var oMath         = oSelectedInfo.GetMath();

		if (!oMath)
		{
			if (true === e.ShiftKey && true === e.CtrlKey)
				return [0x00A0];
		}
	}
	else if (e.KeyCode == 69 && true === e.CtrlKey) // Ctrl + E + ...
	{
		if (true === e.AltKey) // Ctrl + Alt + E - добавляем знак евро €
			return [0x20AC];
	}
	else if (e.KeyCode == 189) // Клавиша Num-
	{
		if (true === e.CtrlKey && true === e.ShiftKey)
			return [0x2013];
	}

	return [];
};
CDocument.prototype.Get_Numbering = function()
{
	return this.Numbering;
};
/**
 * @returns {AscWord.CNumbering}
 */
CDocument.prototype.GetNumbering = function()
{
	return this.Numbering;
};
CDocument.prototype.GetNumberingManager = function()
{
	return this.GetNumbering();
};
/**
 * @returns {AscWord.NumberingCollection}
 */
CDocument.prototype.GetNumberingCollection = function()
{
	return this.NumberingCollection;
};
/**
 * Получаем стиль по выделенному фрагменту
 */
CDocument.prototype.GetStyleFromFormatting = function()
{
	return this.Controller.GetStyleFromFormatting();
};
/**
 * Добавляем новый стиль (или заменяем старый с таким же названием).
 * И сразу применяем его к выделенному фрагменту.
 */
CDocument.prototype.Add_NewStyle = function(oStyle)
{
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Document_Styles, {Type : AscCommon.changestype_2_AdditionalTypes, Types : [AscCommon.changestype_Paragraph_Properties]}))
	{
		this.StartAction(AscDFH.historydescription_Document_AddNewStyle);
		var NewStyle = this.Styles.Create_StyleFromInterface(oStyle);
		this.SetParagraphStyle(NewStyle.Get_Name());
		this.Recalculate();
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
/**
 * Удаляем заданный стиль по имени.
 */
CDocument.prototype.Remove_Style = function(sStyleName)
{
	var StyleId = this.Styles.GetStyleIdByName(sStyleName);
	// Сначала проверим есть ли стиль с таким именем
	if (null == StyleId)
		return;

	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Document_Styles))
	{
		this.StartAction(AscDFH.historydescription_Document_RemoveStyle);
		this.Styles.Remove_StyleFromInterface(StyleId);
		this.Recalculate();
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
/**
 * Удаляем все недефолтовые стили в документе.
 */
CDocument.prototype.Remove_AllCustomStyles = function()
{
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Document_Styles))
	{
		this.StartAction(AscDFH.historydescription_Document_RemoveAllCustomStyles);
		this.Styles.Remove_AllCustomStylesFromInterface();
		this.Recalculate();
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
/**
 * Проверяем является ли заданный стиль дефолтовым.
 */
CDocument.prototype.IsStyleDefaultByName = function(sName)
{
	return this.Styles.IsStyleDefaultByName(sName);
};
/**
 * Проверяем изменен ли дефолтовый стиль.
 */
CDocument.prototype.Is_DefaultStyleChanged = function(sName)
{
	return this.Styles.Is_DefaultStyleChanged(sName);
};
CDocument.prototype.Get_Styles = function()
{
	return this.Styles;
};
/**
 * Получаем ссылку на менеджер стилей
 * @returns {CStyles}
 */
CDocument.prototype.GetStyles = function()
{
	return this.Styles;
};
/**
 * Получаем ссылку на менеджер стилей
 * @returns {CStyles}
 */
CDocument.prototype.GetStyleManager = function()
{
	return this.Styles;
};
CDocument.prototype.GetAllTableStyles = function()
{
	return this.GetStyles().GetAllTableStyles();
};
CDocument.prototype.CopyStyle = function()
{
	return this.Styles.CopyStyle();
};
CDocument.prototype.IsStylesUseInDocument = function(styleIds)
{
	let styleManager = this.GetStyleManager();
	let allStyles    = styleManager.GetAllStyles();

	let mapStyleId = {};
	for (let index = 0, count = allStyles.length; index < count; ++index)
	{
		mapStyleId[allStyles[index].GetId()] = 0;
	}

	let allParagraphs = this.GetAllParagraphs()
	for (let index = 0, count = allParagraphs.length; index < count; ++index)
	{
		let paraStyleId = allParagraphs[index].GetParagraphStyle();
		if (undefined !== mapStyleId[paraStyleId])
			mapStyleId[paraStyleId]++;
	}

	let allTables = this.GetAllTables();
	for (let index = 0, count = allTables.length; index < count; ++index)
	{
		let tableStyleId = allTables[index].GetTableStyle();
		if (undefined !== mapStyleId[tableStyleId])
			mapStyleId[tableStyleId]++;
	}

	this.CheckAllRunContent(function(run)
	{
		let runStyleId = run.GetRStyle();
		if (undefined !== mapStyleId[runStyleId])
			mapStyleId[runStyleId]++;
	});


	let result = [];
	for (let index = 0, count = styleIds.length; index < count; ++index)
	{
		let currentId = styleIds[index];
		let isUse     = false;

		if (styleManager.IsStyleDefaultById(currentId) || (undefined !== mapStyleId[currentId] && mapStyleId[currentId] > 0))
		{
			isUse = true;
		}
		else
		{
			let relatedStyles = styleManager.GetRelatedStyles(currentId);
			for (let relatedIndex = 0, relatedCount = relatedStyles.length; relatedIndex < relatedCount; ++relatedIndex)
			{
				let relatedId = relatedStyles[relatedIndex].GetId();
				if (undefined !== mapStyleId[relatedId] && mapStyleId[relatedId] > 0)
				{
					isUse = true;
					break;
				}
			}
		}

		result.push(isUse);
	}

	return result;
};
CDocument.prototype.Get_TableStyleForPara = function()
{
	return null;
};
CDocument.prototype.Get_ShapeStyleForPara = function()
{
	return null;
};
CDocument.prototype.Get_TextBackGroundColor = function()
{
	return undefined;
};
CDocument.prototype.Select_DrawingObject = function(Id)
{
	let drawingObject = AscCommon.g_oTableId.GetById(Id);
	if (!drawingObject || !drawingObject.IsUseInDocument())
		return;

	const oDocContent = drawingObject.GetDocumentContent();
	if(!oDocContent)
		return;

	this.RemoveSelection();

	// Прячем курсор
	this.DrawingDocument.TargetEnd();
	this.DrawingDocument.SetCurrentPage(this.CurPage);

	const oHdrFtr = oDocContent.IsHdrFtr(true);
	if (!oHdrFtr)
	{
		this.Selection.Start = false;
		this.Selection.Use   = true;
		this.SetDocPosType(docpostype_DrawingObjects);
	}
	else
	{
		this.SetDocPosType(docpostype_HdrFtr);
		this.Selection.Start = false;
		this.Selection.Use   = true;
		const oDocContent = oHdrFtr.Content;
		oDocContent.SetDocPosType(docpostype_DrawingObjects);
		oDocContent.Selection.Use   = true;
		oDocContent.Selection.Start = true;
	}

	this.DrawingObjects.selectById(Id, this.CurPage);

	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
};
/**
 * Получем ближайшую возможную позицию курсора.
 * @param PageNum
 * @param X
 * @param Y
 * @param bAnchor
 * @param Drawing
 * @returns {*}
 */
CDocument.prototype.Get_NearestPos = function(PageNum, X, Y, bAnchor, Drawing)
{
	if (undefined === bAnchor)
		bAnchor = false;

	if (docpostype_HdrFtr === this.GetDocPosType())
		return this.HdrFtr.Get_NearestPos(PageNum, X, Y, bAnchor, Drawing);

	var bInText    = (null === this.IsInText(X, Y, PageNum) ? false : true);
	var nInDrawing = this.DrawingObjects.IsInDrawingObject(X, Y, PageNum, this);

	if (true != bAnchor)
	{
		// Проверяем попадание в графические объекты
		var NearestPos = this.DrawingObjects.getNearestPos(X, Y, PageNum, Drawing);
		if ((nInDrawing === DRAWING_ARRAY_TYPE_BEFORE || nInDrawing === DRAWING_ARRAY_TYPE_INLINE || (false === bInText && nInDrawing >= 0)) && null != NearestPos)
			return NearestPos;

		NearestPos = true === this.Footnotes.CheckHitInFootnote(X, Y, PageNum) ? this.Footnotes.GetNearestPos(X, Y, PageNum, false, Drawing) : null;
		if (null !== NearestPos)
			return NearestPos;

		NearestPos = this.Endnotes.CheckHitInEndnote(X, Y, PageNum) ? this.Endnotes.GetNearestPos(X, Y, PageNum, false, Drawing) : null;
		if (NearestPos)
			return NearestPos;
	}

	if (!this.Pages.length)
	{
		return {
			X          : 0,
			Y          : 0,
			Height     : 0,
			PageNum    : 0,
			Internal   : {Line : 0, Page : 0, Range : 0},
			Transform  : null,
			Paragraph  : null,
			ContentPos : null,
			SearchPos  : null
		};
	}

	var ContentPos = this.Internal_GetContentPosByXY(X, Y, PageNum);

	// Делаем логику как в ворде
	if (true === bAnchor && ContentPos > 0 && PageNum > 0 && ContentPos === this.Pages[PageNum].Pos && ContentPos === this.Pages[PageNum - 1].EndPos && this.Pages[PageNum].EndPos > this.Pages[PageNum].Pos && type_Paragraph === this.Content[ContentPos].GetType() && true === this.Content[ContentPos].IsContentOnFirstPage())
		ContentPos++;

	var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, PageNum);
	return this.Content[ContentPos].GetNearestPos(ElementPageIndex, X, Y, bAnchor, Drawing);
};
CDocument.prototype.Internal_Content_Add = function(Position, NewObject, isCorrectContent)
{
	// Position = this.Content.length  допускается
	if (Position < 0 || Position > this.Content.length)
		return;

	var PrevObj = this.Content[Position - 1] ? this.Content[Position - 1] : null;
	var NextObj = this.Content[Position] ? this.Content[Position] : null;

	this.private_RecalculateNumbering([NewObject]);
	this.History.Add(new CChangesDocumentAddItem(this, Position, [NewObject]));
	this.Content.splice(Position, 0, NewObject);
	this.private_UpdateSelectionPosOnAdd(Position);
	NewObject.Set_Parent(this);
	NewObject.Set_DocumentNext(NextObj);
	NewObject.Set_DocumentPrev(PrevObj);

	if (null != PrevObj)
		PrevObj.Set_DocumentNext(NewObject);

	if (null != NextObj)
		NextObj.Set_DocumentPrev(NewObject);

	// Обновим информацию о секциях
	this.SectionsInfo.Update_OnAdd(Position, [NewObject]);

	if (false !== isCorrectContent)
	{
		// В последнем параграфе не должно быть разрыва секции
		this.Check_SectionLastParagraph();

		// Проверим, что последний элемент - параграф
		if (!this.Content[this.Content.length - 1].IsParagraph())
			this.Internal_Content_Add(this.Content.length, new Paragraph(this.DrawingDocument, this), false);
	}

	// Запоминаем, что нам нужно произвести переиндексацию элементов
	this.private_ReindexContent(Position);

	if (type_Paragraph === NewObject.GetType())
		this.DocumentOutline.CheckParagraph(NewObject);
};
CDocument.prototype.Internal_Content_Remove = function(Position, Count, isCorrectContent)
{
	var ChangePos = -1;

	if (Position < 0 || Position >= this.Content.length || Count <= 0)
		return -1;

	var PrevObj = this.Content[Position - 1] ? this.Content[Position - 1] : null;
	var NextObj = this.Content[Position + Count] ? this.Content[Position + Count] : null;

	for (var Index = 0; Index < Count; Index++)
	{
		this.Content[Position + Index].PreDelete();
	}

	this.History.Add(new CChangesDocumentRemoveItem(this, Position, this.Content.slice(Position, Position + Count)));
	var Elements = this.Content.splice(Position, Count);
	this.private_RecalculateNumbering(Elements);
	this.private_UpdateSelectionPosOnRemove(Position, Count);

	if (null != PrevObj)
		PrevObj.Set_DocumentNext(NextObj);

	if (null != NextObj)
		NextObj.Set_DocumentPrev(PrevObj);

	if (false !== isCorrectContent)
	{
		// Проверим последний параграф
		this.Check_SectionLastParagraph();

		// Проверим, что последний элемент - параграф
		if (this.Content.length <= 0 || !this.Content[this.Content.length - 1].IsParagraph())
			this.Internal_Content_Add(this.Content.length, new Paragraph(this.DrawingDocument, this));
	}

	// Обновим информацию о секциях
	this.SectionsInfo.Update_OnRemove(Position, Count, true);

	// Проверим не является ли рамкой последний параграф
	this.private_CheckFramePrLastParagraph();

	// Запоминаем, что нам нужно произвести переиндексацию элементов
	this.private_ReindexContent(Position);

	return ChangePos;
};
CDocument.prototype.Clear_ContentChanges = function()
{
	this.m_oContentChanges.Clear();
};
CDocument.prototype.Add_ContentChanges = function(Changes)
{
	this.m_oContentChanges.Add(Changes);
};
CDocument.prototype.Refresh_ContentChanges = function()
{
	this.m_oContentChanges.Refresh();
};
/**
 * @param AlignV 0 - Верх, 1 - низ
 * @param AlignH стандартные значения align_Left, align_Center, align_Right
 */
CDocument.prototype.Document_AddPageNum = function(AlignV, AlignH)
{
	if (AlignV >= 0)
	{
		var PageIndex = this.CurPage;
		if (docpostype_HdrFtr === this.GetDocPosType())
			PageIndex = this.HdrFtr.Get_CurPage();

		if (PageIndex < 0)
			PageIndex = this.CurPage;

		this.Create_HdrFtrWidthPageNum(PageIndex, AlignV, AlignH);
	}
	else
	{
		this.AddToParagraph(new AscWord.CRunPageNum());
	}

	this.Document_UpdateInterfaceState();
};
CDocument.prototype.Document_SetHdrFtrFirstPage = function(Value)
{
	var oCurHdrFtr, nCurPage;

	if (!(oCurHdrFtr = this.HdrFtr.GetCurHdrFtr())
		|| -1 === (nCurPage = oCurHdrFtr.GetPage()))
		return;

	var Index  = this.Pages[nCurPage].Pos;
	var SectPr = this.SectionsInfo.Get_SectPr(Index).SectPr;

	SectPr.Set_TitlePage(Value);

	if (true === Value)
	{
		// Если мы добавляем разные колонтитулы для первой страницы, а этих колонтитулов нет, тогда создаем их
		var FirstSectPr = this.SectionsInfo.Get_SectPr2(0).SectPr;
		var FirstHeader = FirstSectPr.Get_Header_First();
		var FirstFooter = FirstSectPr.Get_Footer_First();

		if (null === FirstHeader)
		{
			var Header = new CHeaderFooter(this.HdrFtr, this, this.DrawingDocument, hdrftr_Header);
			FirstSectPr.Set_Header_First(Header);

			this.HdrFtr.Set_CurHdrFtr(Header);
		}
		else
			this.HdrFtr.Set_CurHdrFtr(FirstHeader);

		if (null === FirstFooter)
		{
			var Footer = new CHeaderFooter(this.HdrFtr, this, this.DrawingDocument, hdrftr_Footer);
			FirstSectPr.Set_Footer_First(Footer);
		}
	}
	else
	{
		var TempSectPr = SectPr;
		var TempIndex  = Index;
		while (null === TempSectPr.Get_Header_Default())
		{
			TempIndex--;
			if (TempIndex < 0)
				break;

			TempSectPr = this.SectionsInfo.Get_SectPr(TempIndex).SectPr;
		}

		var oHeader = TempSectPr.Get_Header_Default();
		if (!oHeader)
		{
			var oHeader = new CHeaderFooter(this.HdrFtr, this, this.DrawingDocument, hdrftr_Header);
			TempSectPr.Set_Header_Default(oHeader);
		}

		this.HdrFtr.Set_CurHdrFtr(TempSectPr.Get_Header_Default());
	}


	this.Recalculate();

	if (null !== this.HdrFtr.CurHdrFtr)
	{
		this.HdrFtr.CurHdrFtr.Content.MoveCursorToStartPos();
		this.HdrFtr.CurHdrFtr.SetPage(nCurPage);
	}

	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.Document_SetHdrFtrEvenAndOddHeaders = function(Value)
{
	this.Set_DocumentEvenAndOddHeaders(Value);

	var FirstSectPr;
	if (true === Value)
	{
		// Если мы добавляем разные колонтитулы для четных и нечетных страниц, а этих колонтитулов нет, тогда
		// создаем их в самой первой секции
		FirstSectPr = this.SectionsInfo.Get_SectPr2(0).SectPr;
		if (null === FirstSectPr.Get_Header_Even())
		{
			var Header = new CHeaderFooter(this.HdrFtr, this, this.DrawingDocument, hdrftr_Header);
			FirstSectPr.Set_Header_Even(Header);
		}

		if (null === FirstSectPr.Get_Footer_Even())
		{
			var Footer = new CHeaderFooter(this.HdrFtr, this, this.DrawingDocument, hdrftr_Footer);
			FirstSectPr.Set_Footer_Even(Footer);
		}
	}
	else
	{
		FirstSectPr = this.SectionsInfo.Get_SectPr2(0).SectPr;
	}

	if (null !== FirstSectPr.Get_Header_First() && true === FirstSectPr.TitlePage)
	{
		this.HdrFtr.Set_CurHdrFtr(FirstSectPr.Get_Header_First());
	}
	else
	{
		var oHeader = FirstSectPr.Get_Header_Default();
		if (!oHeader)
		{
			oHeader = new CHeaderFooter(this.HdrFtr, this, this.DrawingDocument, hdrftr_Header);
			FirstSectPr.Set_Header_Default(oHeader);
		}

		this.HdrFtr.Set_CurHdrFtr(oHeader);
	}


	this.Recalculate();

	if (null !== this.HdrFtr.CurHdrFtr)
		this.HdrFtr.CurHdrFtr.Content.MoveCursorToStartPos();

	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.Document_SetHdrFtrDistance = function(Value)
{
	var CurHdrFtr = this.HdrFtr.CurHdrFtr;

	if (null === CurHdrFtr)
		return;

	var CurPage = CurHdrFtr.RecalcInfo.CurPage;
	if (-1 === CurPage)
		return;

	var Index  = this.Pages[CurPage].Pos;
	var SectPr = this.SectionsInfo.Get_SectPr(Index).SectPr;

	if (hdrftr_Header === CurHdrFtr.Type)
		SectPr.SetPageMarginHeader(Value);
	else
		SectPr.SetPageMarginFooter(Value);

	this.Recalculate();

	this.Document_UpdateRulersState();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
};
CDocument.prototype.Document_SetHdrFtrBounds = function(Y0, Y1)
{
	var CurHdrFtr = this.HdrFtr.CurHdrFtr;

	if (null === CurHdrFtr)
		return;

	var CurPage = CurHdrFtr.RecalcInfo.CurPage;
	if (-1 === CurPage)
		return;

	var Index  = this.Pages[CurPage].Pos;
	var SectPr = this.SectionsInfo.Get_SectPr(Index).SectPr;
	var Bounds = CurHdrFtr.Get_Bounds();

	if (hdrftr_Header === CurHdrFtr.Type)
	{
		if (null !== Y0)
			SectPr.SetPageMarginHeader(Y0);

		if (null !== Y1)
			SectPr.SetPageMargins(undefined, Y1, undefined, undefined);
	}
	else
	{
		if (null !== Y0)
		{
			var H   = Bounds.Bottom - Bounds.Top;
			var _Y1 = Y0 + H;

			SectPr.SetPageMarginFooter(SectPr.GetPageHeight() - _Y1);
		}
	}

	this.Recalculate();

	this.Document_UpdateRulersState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.Document_SetHdrFtrLink = function(bLinkToPrevious)
{
	var CurHdrFtr = this.HdrFtr.CurHdrFtr;
	if (docpostype_HdrFtr !== this.GetDocPosType() || null === CurHdrFtr || -1 === CurHdrFtr.RecalcInfo.CurPage)
		return;

	var PageIndex = CurHdrFtr.RecalcInfo.CurPage;

	var Index  = this.Pages[PageIndex].Pos;
	var SectPr = this.SectionsInfo.Get_SectPr(Index).SectPr;

	// У самой первой секции не может быть повторяющихся колонтитулов, поэтому не делаем ничего
	if (SectPr === this.SectionsInfo.Get_SectPr2(0).SectPr)
		return;

	// Определим тип колонтитула, в котором мы находимся
	var SectionPageInfo = this.Get_SectionPageNumInfo(PageIndex);

	var bFirst  = ( true === SectionPageInfo.bFirst && true === SectPr.Get_TitlePage() ? true : false );
	var bEven   = ( true === SectionPageInfo.bEven && true === EvenAndOddHeaders ? true : false );
	var bHeader = ( hdrftr_Header === CurHdrFtr.Type ? true : false );

	var _CurHdrFtr = SectPr.GetHdrFtr(bHeader, bFirst, bEven);

	if (true === bLinkToPrevious)
	{
		// Если нам надо повторять колонтитул, а он уже изначально повторяющийся, тогда не делаем ничего
		if (null === _CurHdrFtr)
			return;

		// Очистим селект
		_CurHdrFtr.RemoveSelection();

		// Просто удаляем запись о данном колонтитуле в секции
		SectPr.Set_HdrFtr(bHeader, bFirst, bEven, null);

		var HdrFtr = this.Get_SectionHdrFtr(PageIndex, bFirst, bEven);

		// Заглушка. Вообще такого не должно быть, чтобы был колонтитул не в первой секции, и не было в первой,
		// но, на всякий случай, обработаем такую ситуацию.
		if (true === bHeader)
		{
			if (null === HdrFtr.Header)
				CurHdrFtr = this.Create_SectionHdrFtr(hdrftr_Header, PageIndex);
			else
				CurHdrFtr = HdrFtr.Header;
		}
		else
		{
			if (null === HdrFtr.Footer)
				CurHdrFtr = this.Create_SectionHdrFtr(hdrftr_Footer, PageIndex);
			else
				CurHdrFtr = HdrFtr.Footer;
		}


		this.HdrFtr.Set_CurHdrFtr(CurHdrFtr);
		this.HdrFtr.CurHdrFtr.MoveCursorToStartPos(false);
	}
	else
	{
		// Если данный колонтитул уже не повторяющийся, тогда ничего не делаем
		if (null !== _CurHdrFtr)
			return;

		var NewHdrFtr = CurHdrFtr.Copy();
		SectPr.Set_HdrFtr(bHeader, bFirst, bEven, NewHdrFtr);
		this.HdrFtr.Set_CurHdrFtr(NewHdrFtr);
		this.HdrFtr.CurHdrFtr.MoveCursorToStartPos(false);
	}

	this.Recalculate();

	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.SetSectionStartPage = function(nStartPage)
{
	var oCurHdrFtr = this.HdrFtr.CurHdrFtr;
	if (!oCurHdrFtr)
		return;

	var nCurPage = oCurHdrFtr.RecalcInfo.CurPage;
	if (-1 === nCurPage)
		return;

	var nIndex  = this.Pages[nCurPage].Pos;
	var oSectPr = this.SectionsInfo.Get_SectPr(nIndex).SectPr;

	oSectPr.Set_PageNum_Start(nStartPage);

	this.Recalculate();

	this.Document_UpdateRulersState();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
};
CDocument.prototype.Document_Format_Copy = function()
{
	this.Api.checkFormatPainterData();
};
CDocument.prototype.EndHdrFtrEditing = function(bCanStayOnPage)
{
	if (docpostype_HdrFtr === this.GetDocPosType())
	{
		this.SetDocPosType(docpostype_Content);
		var CurHdrFtr = this.HdrFtr.Get_CurHdrFtr();
		if (null === CurHdrFtr || undefined === CurHdrFtr || true !== bCanStayOnPage)
		{
			this.MoveCursorToStartPos(false);
		}
		else
		{
			CurHdrFtr.RemoveSelection();

			if (hdrftr_Header == CurHdrFtr.Type)
				this.MoveCursorToXY(0, 0, false);
			else
				this.MoveCursorToXY(0, 100000, false); // TODO: Переделать здесь по-нормальному
		}

		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();

		this.Document_UpdateRulersState();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
	}
};
CDocument.prototype.EndFootnotesEditing = function()
{
	if (docpostype_Footnotes === this.GetDocPosType())
	{
		this.SetDocPosType(docpostype_Content);

		this.MoveCursorToStartPos(false);

		// TODO: Не всегда можно в данной функции использовать MoveCursorToXY, потому что
		//       данная страница еще может быть не рассчитана.
		//this.MoveCursorToXY(0, 0, false);

		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();

		this.Document_UpdateRulersState();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
	}
};
CDocument.prototype.EndEndnotesEditing = function()
{
	if (docpostype_Endnotes === this.GetDocPosType())
	{
		this.SetDocPosType(docpostype_Content);

		this.MoveCursorToStartPos(false);

		// TODO: Не всегда можно в данной функции использовать MoveCursorToXY, потому что
		//       данная страница еще может быть не рассчитана.
		//this.MoveCursorToXY(0, 0, false);

		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();

		this.Document_UpdateRulersState();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
	}
};
CDocument.prototype.EndDrawingEditing = function()
{
	if (docpostype_DrawingObjects === this.GetDocPosType() || (docpostype_HdrFtr === this.GetDocPosType() && null != this.HdrFtr.CurHdrFtr && docpostype_DrawingObjects === this.HdrFtr.CurHdrFtr.Content.CurPos.Type ))
	{
		this.DrawingObjects.resetSelection2();
		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
		this.private_UpdateCursorXY(true, true);
	}
};
CDocument.prototype.EndFormEditing = function()
{
	var oSelectedInfo = this.GetSelectedElementsInfo();
	var oInlineSdt    = oSelectedInfo.GetInlineLevelSdt();

	if (oInlineSdt && oInlineSdt.IsForm())
	{
		if (oInlineSdt.IsFixedForm())
			this.EndDrawingEditing();
		else
			oInlineSdt.MoveCursorOutsideElement(true);

		this.UpdateSelection();
		this.UpdateInterface();
	}
};
CDocument.prototype.Document_Format_Paste = function()
{
	let oData = this.Api.getFormatPainterData();
	if(!oData)
		return;
	this.Controller.PasteFormatting(oData);
	this.Recalculate();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
};
CDocument.prototype.IsTableCellContent = function(isReturnCell)
{
	if (true === isReturnCell)
		return null;

	return false;
};
CDocument.prototype.Check_AutoFit = function()
{
	return false;
};
CDocument.prototype.Is_TopDocument = function(bReturnTopDocument)
{
	if (true === bReturnTopDocument)
		return this;

	return true;
};
CDocument.prototype.IsInTable = function(bReturnTopTable)
{
	if (true === bReturnTopTable)
		return null;

	return false;
};
CDocument.prototype.Is_DrawingShape = function(bRetShape)
{
	if (bRetShape === true)
	{
		return null;
	}
	return false;
};
CDocument.prototype.IsSelectionUse = function()
{
	return this.Controller.IsSelectionUse();
};
CDocument.prototype.IsNumberingSelection = function()
{
	return this.Controller.IsNumberingSelection();
};
CDocument.prototype.IsTextSelectionUse = function()
{
	return this.Controller.IsTextSelectionUse();
};
CDocument.prototype.GetCurPosXY = function()
{
	var TempXY = this.Controller.GetCurPosXY();
	this.private_CheckCurPage();
	return {X : TempXY.X, Y : TempXY.Y, PageNum : this.CurPage};
};
/**
 * Возвращаем выделенный текст, если в выделении не более 1 параграфа, и там нет картинок, нумерации страниц и т.д.
 * @param bClearText
 * @param oPr
 * @returns {?string}
 */
CDocument.prototype.GetSelectedText = function(bClearText, oPr)
{
	if (undefined === oPr)
		oPr = {};

	if (undefined === bClearText)
		bClearText = false;

	return this.Controller.GetSelectedText(bClearText, oPr);
};
/**
 * Получаем текущий параграф (или граничные параграфы селекта)
 * @param bIgnoreSelection Если true, тогда используется текущая позиция, даже если есть селект
 * @param bReturnSelectedArray (Используется, только если bIgnoreSelection==false) Если true, тогда возвращаем массив из
 * из параграфов, которые попали в выделение.
 * @param {object} oPr
 *  oPr.ReplacePlaceHolder {boolean} - Если текущий параграф в плейсхолдере, тогда заменяем его нормальным контентом
 *  oPr.ReturnSelectedTable {boolean} - Если идет выделение по ячейкам, тогда возвращаем таблицу
 *  oPr.FirstInSelection {boolean} - Если true, и bReturnSelectedArray = false, то мы возвращаем первый параграф в выделении (а не тот, с которого начали)
 *  oPr.LastInSelection {boolean} - Если true, и bReturnSelectedArray = false, то мы возвращаем последний параграф в выделении
 * @returns {Paragraph | [Paragraph]}
 */
CDocument.prototype.GetCurrentParagraph = function(bIgnoreSelection, bReturnSelectedArray, oPr)
{
	if (true !== bIgnoreSelection && true === bReturnSelectedArray)
	{
		var arrSelectedParagraphs = [];
		this.Controller.GetCurrentParagraph(bIgnoreSelection, arrSelectedParagraphs, oPr);
		return arrSelectedParagraphs;
	}
	else
	{
		return this.Controller.GetCurrentParagraph(bIgnoreSelection, null, oPr);
	}
};
/**
 * Возвращаем массив параграфов, попавших в селект
 * @returns {Paragraph[]}
 */
CDocument.prototype.GetSelectedParagraphs = function()
{
	return this.GetCurrentParagraph(false, true);
};
/**
 * Получаем текущую таблицу
 * @returns {?CTable}
 */
CDocument.prototype.GetCurrentTable = function()
{
	var arrTables = this.GetCurrentTablesStack();
	return arrTables.length > 0 ? arrTables[arrTables.length - 1] : null;
};
/**
 * Получаем массив таблицц, в которых мы находимся
 * @returns {CTable[]}
 */
CDocument.prototype.GetCurrentTablesStack = function()
{
	var arrTables = [];
	this.Controller.GetCurrentTablesStack(arrTables)
	return arrTables;
};
/**
 * Получаем массив автофигур, попавших в текстовый селект
 * @returns {ParaDrawing[]}
 */
CDocument.prototype.GetSelectedDrawingObjectsInText = function()
{
	let arrDrawings = [];
	let arrParagraphs = this.GetSelectedParagraphs();
	for (let nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
	{
		arrParagraphs[nIndex].CheckRunContent(function(oRun)
		{
			oRun.GetSelectedDrawingObjectsInText(arrDrawings);
		});
	}

	return arrDrawings;
};
/**
 * Получаем информацию о текущем выделении
 * @returns {CSelectedElementsInfo}
 */
CDocument.prototype.GetSelectedElementsInfo = function(oPr)
{
	var oInfo = new CSelectedElementsInfo(oPr);
	this.Controller.GetSelectedElementsInfo(oInfo);
	return oInfo;
};
/**
 * Получаем текущее слово (или часть текущего слова)
 * @param nDirection -1 - часть до курсора, 1 - часть после курсора, 0 (или не задано) слово целиком
 * @returns {string}
 */
CDocument.prototype.GetCurrentWord = function(nDirection)
{
	var oParagraph = this.GetCurrentParagraph();
	if (!oParagraph)
		return "";

	return oParagraph.GetCurrentWord(nDirection);
};
/**
 * Заменяем текущее слово (или часть текущего слова) на заданную строку
 * @param nDirection {number} -1 - часть до курсора, 1 - часть после курсора, 0 (или не задано) слово целиком
 * @param sReplace {string}
 * @returns {boolean}
 */
CDocument.prototype.ReplaceCurrentWord = function(nDirection, sReplace)
{
	var oParagraph = this.GetCurrentParagraph();
	if (!oParagraph)
		return false;

	var bResult = false;
	if (!this.IsSelectionLocked(AscCommon.changestype_None, {
		Type      : AscCommon.changestype_2_Element_and_Type,
		Element   : oParagraph,
		CheckType : AscCommon.changestype_Paragraph_Content
	}))
	{
		this.StartAction(AscDFH.historydescription_Document_ReplaceCurrentWord);
		bResult = oParagraph.ReplaceCurrentWord(nDirection, sReplace);
		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	}

	return bResult;
};
/**
 * Получаем текущее слово (или часть текущего слова)
 * @param nDirection -1 - часть до курсора, 1 - часть после курсора, 0 (или не задано) слово целиком
 * @returns {string}
 */
CDocument.prototype.GetCurrentSentence = function(nDirection)
{
	let paragraph = this.GetCurrentParagraph();
	if (!paragraph)
		return "";
	
	return paragraph.GetCurrentSentence(nDirection);
};
/**
 * Заменяем текущее предложения (или часть текущего предложения) на заданную строку
 * @param nDirection {number} -1 - часть до курсора, 1 - часть после курсора, 0 (или не задано) слово целиком
 * @param sReplace {string}
 * @returns {boolean}
 */
CDocument.prototype.ReplaceCurrentSentence = function(nDirection, sReplace)
{
	let paragraph = this.GetCurrentParagraph();
	if (!paragraph)
		return false;
	
	let result = false;
	if (!this.IsSelectionLocked(AscCommon.changestype_None, {
		Type      : AscCommon.changestype_2_Element_and_Type,
		Element   : paragraph,
		CheckType : AscCommon.changestype_Paragraph_Content
	}))
	{
		this.StartAction(AscDFH.historydescription_Document_ReplaceCurrentWord);
		result = paragraph.ReplaceCurrentSentence(nDirection, sReplace);
		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	}
	
	return result;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с таблицами
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.AddTableRow = function(bBefore, nCount)
{
	this.Controller.AddTableRow(bBefore, nCount);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.AddTableColumn = function(bBefore, nCount)
{
	this.Controller.AddTableColumn(bBefore, nCount);
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.RemoveTableRow = function()
{
	this.Controller.RemoveTableRow();
	this.Recalculate();
	this.UpdateSelection();
	this.UpdateInterface();
};
CDocument.prototype.RemoveTableColumn = function()
{
	this.Controller.RemoveTableColumn();
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.MergeTableCells = function()
{
	var isLocalTrackRevisions = this.GetLocalTrackRevisions();
	this.SetLocalTrackRevisions(false);
	this.Controller.MergeTableCells();
	this.SetLocalTrackRevisions(isLocalTrackRevisions);
	this.Recalculate();
	this.UpdateSelection();
	this.UpdateInterface();
};
CDocument.prototype.SplitTableCells = function(Cols, Rows)
{
	var isLocalTrackRevisions = this.GetLocalTrackRevisions();
	this.SetLocalTrackRevisions(false);
	this.Controller.SplitTableCells(Cols, Rows);
	this.SetLocalTrackRevisions(isLocalTrackRevisions);
	this.Recalculate();
	this.UpdateSelection();
	this.UpdateInterface();
};
CDocument.prototype.RemoveTableCells = function()
{
	this.Controller.RemoveTableCells();
	this.Recalculate();
	this.UpdateSelection();
	this.UpdateInterface();
};
CDocument.prototype.RemoveTable = function()
{
	this.Controller.RemoveTable();
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateRulersState();
};
CDocument.prototype.SelectTable = function(Type)
{
	this.Controller.SelectTable(Type);
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
	//this.Document_UpdateRulersState();
};
CDocument.prototype.CanMergeTableCells = function()
{
	return this.Controller.CanMergeTableCells();
};
CDocument.prototype.CanSplitTableCells = function()
{
	return this.Controller.CanSplitTableCells();
};
CDocument.prototype.CheckTableCoincidence = function(Table)
{
	return false;
};
CDocument.prototype.DistributeTableCells = function(isHorizontally)
{
	if (!this.Controller.DistributeTableCells(isHorizontally))
		return false;

	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();

	return true;
};
//----------------------------------------------------------------------------------------------------------------------
// Дополнительные функции
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.Document_CreateFontMap = function()
{
	var FontMap = {};
	this.SectionsInfo.Document_CreateFontMap(FontMap);

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		this.Content[Index].CreateFontMap(FontMap);
	}
	AscFormat.checkThemeFonts(FontMap, this.theme.themeElements.fontScheme);
	return FontMap;
};
CDocument.prototype.Document_CreateFontCharMap = function(FontCharMap)
{
	this.SectionsInfo.Document_CreateFontCharMap(FontCharMap);
	this.DrawingObjects.documentCreateFontCharMap(FontCharMap);

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		this.Content[Index].CreateFontCharMap(FontCharMap);
	}
};
CDocument.prototype.Document_Get_AllFontNames = function()
{
	var AllFonts = {};

	this.SectionsInfo.Document_Get_AllFontNames(AllFonts);
	this.Numbering.GetAllFontNames(AllFonts);
	this.Styles.Document_Get_AllFontNames(AllFonts);
	this.theme.Document_Get_AllFontNames(AllFonts);

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetAllFontNames(AllFonts);
	}
	AscFormat.checkThemeFonts(AllFonts, this.theme.themeElements.fontScheme);
	return AllFonts;
};
CDocument.prototype.Document_UpdateInterfaceState = function(bSaveCurRevisionChange)
{
	this.UpdateInterface(bSaveCurRevisionChange);
};
CDocument.prototype.private_UpdateInterface = function(isSaveCurrentReviewChange, isExternalTrigger)
{
	let oApi = this.GetApi();
	if (!oApi || !oApi.isDocumentLoadComplete || true === AscCommon.g_oIdCounter.m_bLoad || true === AscCommon.g_oIdCounter.m_bRead)
		return;

	if (true === this.TurnOffInterfaceEvents)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
	{
		let oThis = this;
		oApi.checkLongActionCallback(function(){oThis.UpdateInterface();});
		return;
	}

	// Удаляем весь список
	oApi.sync_BeginCatchSelectedElements();

	// Уберем из интерфейса записи о том где мы находимся (параграф, таблица, картинка или колонтитул)
	oApi.ClearPropObjCallback();

	this.Controller.UpdateInterfaceState();

	// Сообщаем, что список составлен
	oApi.sync_EndCatchSelectedElements(isExternalTrigger);

	this.UpdateSelectedReviewChanges(isSaveCurrentReviewChange);

	this.Document_UpdateUndoRedoState();
	this.Document_UpdateCanAddHyperlinkState();
	this.Document_UpdateSectionPr();
	this.UpdateStylePanel();
	this.UpdateNumberingPanel();
};
CDocument.prototype.private_UpdateRulers = function()
{
	if (true === this.TurnOffInterfaceEvents)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	this.DrawingDocument.Set_RulerState_Start();
	this.Controller.UpdateRulersState();
	this.DrawingDocument.Set_RulerState_End();
};
CDocument.prototype.Document_UpdateRulersState = function()
{
	this.UpdateRulers();
};
CDocument.prototype.Document_UpdateRulersStateBySection = function(nPos)
{
	// В данной функции мы уже точно знаем, что нам секцию нужно выбирать исходя из текущего параграфа
	var nCurPos = undefined === nPos ? ( this.Selection.Use === true ? this.Selection.EndPos : this.CurPos.ContentPos ) : nPos;

	var oSectPr = this.SectionsInfo.Get_SectPr(nCurPos).SectPr;
	if (oSectPr.GetColumnsCount() > 1)
	{
		this.ColumnsMarkup.UpdateFromSectPr(oSectPr, this.CurPage);

		var oElement = this.Content[nCurPos];
		if (oElement.IsParagraph())
			this.ColumnsMarkup.SetCurCol(oElement.Get_CurrentColumn());

		this.DrawingDocument.Set_RulerState_Columns(this.ColumnsMarkup);
	}
	else
	{
		var oFrame = oSectPr.GetContentFrame(this.CurPage);
		this.DrawingDocument.Set_RulerState_Paragraph({L : oFrame.Left, T : oFrame.Top, R : oFrame.Right, B : oFrame.Bottom}, true);
	}
};
CDocument.prototype.Document_UpdateSelectionState = function()
{
	this.UpdateSelection();
};
CDocument.prototype.private_UpdateSelection = function()
{
	if (true === this.TurnOffInterfaceEvents)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	this.DrawingDocument.UpdateTargetTransform(null);

	// TODO: Надо подумать как это вынести в верхнюю функцию внутренние реализации типа controller_UpdateSelectionState
	//       потому что они все очень похожие.

	this.Controller.UpdateSelectionState();

	if (this.IsFillingFormMode() && !this.IsInFormField())
		this.DrawingDocument.TargetEnd();

	this.UpdateDocumentOutlinePosition();

	// Обновим состояние кнопок Copy/Cut
	this.Document_UpdateCopyCutState();
};
CDocument.prototype.UpdateDocumentOutlinePosition = function()
{
	if (this.DocumentOutline.IsUse())
	{
		if (this.Controller !== this.LogicDocumentController)
		{
			this.DocumentOutline.UpdateCurrentPosition(null);
		}
		else
		{
			var oCurrentParagraph = this.GetCurrentParagraph(false, false);
			this.DocumentOutline.UpdateCurrentPosition(oCurrentParagraph.GetDocumentPositionFromObject());
		}
	}
};
CDocument.prototype.private_UpdateDocumentTracks = function()
{
	if (true === this.TurnOffInterfaceEvents)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	this.private_UpdateTracks(this.IsSelectionUse(), this.IsSelectionEmpty());
};
CDocument.prototype.private_UpdateTracks = function(bSelection, bEmptySelection)
{
	var Pos = (true === this.Selection.Use && selectionflag_Numbering !== this.Selection.Flag ? this.Selection.EndPos : this.CurPos.ContentPos);
	if (docpostype_Content === this.GetDocPosType() && !(Pos >= 0 && (null === this.FullRecalc.Id || this.FullRecalc.StartIndex > Pos)))
	{
		this.NeedUpdateTracksOnRecalc = true;

		this.NeedUpdateTracksParams.Selection      = bSelection;
		this.NeedUpdateTracksParams.EmptySelection = bEmptySelection;

		return;
	}

	var isNeedRedraw = false;

	this.NeedUpdateTracksOnRecalc = false;

	var oSelectedInfo = this.GetSelectedElementsInfo();

	var oMath = oSelectedInfo.GetMath();
	this.MathTrackHandler.SetTrackObject(this.IsShowEquationTrack() ? oMath : null, this.CurPage, false === bSelection || true === bEmptySelection)

	var oBlockLevelSdt  = oSelectedInfo.GetBlockLevelSdt();
	var oInlineLevelSdt = oSelectedInfo.GetInlineLevelSdt();

	var oCurrentForm = null;
	if (oInlineLevelSdt)
	{
		if (oInlineLevelSdt.IsForm())
			oCurrentForm = oInlineLevelSdt;
		
		if (!oInlineLevelSdt.IsForm() || !oInlineLevelSdt.IsFixedForm() || this.IsFillingOFormMode())
			oInlineLevelSdt.DrawContentControlsTrack(AscCommon.ContentControlTrack.In);
	}
	else if (oBlockLevelSdt)
	{
		oBlockLevelSdt.DrawContentControlsTrack(AscCommon.ContentControlTrack.In);
	}
	else
	{
		var oForm = null, oMajorParaDrawing;
		if (docpostype_DrawingObjects === this.GetDocPosType()
			&& (oMajorParaDrawing = this.DrawingObjects.getMajorParaDrawing())
			&& this.DrawingObjects.getTargetDocContent())
			oForm = oMajorParaDrawing.GetInnerForm();

		if (oForm)
			oForm.DrawContentControlsTrack(AscCommon.ContentControlTrack.In);
		else
			this.DrawingDocument.OnDrawContentControl(null, AscCommon.ContentControlTrack.In);
	}

	this.UpdateContentControlFocusState(oInlineLevelSdt ? oInlineLevelSdt : (oBlockLevelSdt ? oBlockLevelSdt : null));

	if (this.private_SetCurrentSpecialForm(oCurrentForm))
	{
		isNeedRedraw = true;
	}

	var oField = oSelectedInfo.GetField();
	if (null !== oField && (fieldtype_MERGEFIELD !== oField.Get_FieldType() || true !== this.MailMergeFieldsHighlight))
	{
		var aBounds = oField.Get_Bounds();
		this.DrawingDocument.Update_FieldTrack(true, aBounds);
	}
	else
	{
		this.DrawingDocument.Update_FieldTrack(false);

		var arrComplexFields = oSelectedInfo.GetComplexFields();
		if ((arrComplexFields.length > 0 && this.FieldsManager.SetCurrentComplexField(arrComplexFields[arrComplexFields.length - 1]))
			|| (arrComplexFields.length <= 0 && this.FieldsManager.SetCurrentComplexField(null)))
		{
			isNeedRedraw = true;
		}
	}

	if (isNeedRedraw)
	{
		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();
	}
};
CDocument.prototype.UpdateContentControlFocusState = function(oCC)
{
	if (this.FocusCC === oCC)
		return;

	if (this.FocusCC)
	{
		this.CheckTextFormFormatOnBlur(this.FocusCC);
		this.CheckContentControlPlaceholderOnBlur(this.FocusCC);
	}

	if (this.FocusCC && this.Api)
		this.Api.asc_OnBlurContentControl(this.FocusCC);

	if (oCC && this.Api)
		this.Api.asc_OnFocusContentControl(oCC);

	this.FocusCC = oCC;
};
CDocument.prototype.CheckTextFormFormatOnBlur = function(oForm)
{
	if (!oForm
		|| !oForm.IsUseInDocument()
		|| !oForm.IsForm()
		|| !oForm.IsTextForm()
		|| oForm.IsComplexForm()
		|| oForm.IsPlaceHolder()
		|| this.StartCheckTextFormFormat)
		return;

	this.StartCheckTextFormFormat = true;

	let sInnerText  = oForm.GetInnerText();
	let oTextFormPr = oForm.GetTextFormPr();
	let sText       = sInnerText;

	let format = oTextFormPr.GetFormat();
	if (!format.Check(sInnerText))
		sText = format.Correct(sInnerText);

	if (!format.Check(sText))
	{
		this.Api.sendEvent("asc_onError", Asc.c_oAscError.ID.TextFormWrongFormat, Asc.c_oAscError.Level.NoCritical, oTextFormPr);

		if (this.CollaborativeEditing.Is_SingleUser() || !this.CollaborativeEditing.Is_Fast())
		{
			let arrChanges = this.History.UndoFormFilling(oForm);
			if (arrChanges.length)
				this.RecalculateByChanges(arrChanges);
		}
	}
	else if (sText !== sInnerText)
	{
		this.private_UpdateFormInnerText(oForm, sText);
	}
	else if (oForm.IsEmpty() && !oForm.IsPlaceHolder())
	{
		this.private_UpdateFormToPlaceholder(oForm);
	}

	this.History.ClearFormFillingInfo();

	this.StartCheckTextFormFormat = false;
};
CDocument.prototype.CheckContentControlPlaceholderOnBlur = function(contentControl)
{
	if (!contentControl
		|| !contentControl.IsUseInDocument()
		|| contentControl.IsForm()
		|| contentControl.IsPlaceHolder()
		|| !contentControl.IsEmpty()
		|| this.StartCheckCCPlaceholder)
		return;
	
	this.StartCheckCCPlaceholder = true;
	
	let state = this.SaveDocumentState();
	contentControl.SelectContentControl();
	contentControl.SkipSpecialContentControlLock(true);
	
	if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, this.IsFillingFormMode()))
	{
		contentControl.SkipSpecialContentControlLock(false);
		this.LoadDocumentState(state);
		this.StartCheckCCPlaceholder = false;
		return;
	}
	contentControl.SkipSpecialContentControlLock(false);
	contentControl.MoveCursorToContentControl(true);
	this.StartAction(AscDFH.historydescription_Document_FillContentControlPlaceholderOnBlur, this.GetSelectionState());
	
	contentControl.ReplaceContentWithPlaceHolder();
	this.LoadDocumentState(state);
	this.Recalculate();
	this.UpdateInterface();
	this.FinalizeAction();
	
	this.StartCheckCCPlaceholder = false;
};
CDocument.prototype.private_UpdateFormInnerText = function(form, text)
{
	let paragraph = form.GetParagraph();
	if (!paragraph || form.IsPlaceHolder())
		return;

	if (!this.IsSelectionLocked(AscCommon.changestype_None, {
		Type      : AscCommon.changestype_2_Element_and_Type,
		Element   : paragraph,
		CheckType : AscCommon.changestype_Paragraph_Content
	}, true, this.IsFillingFormMode()))
	{
		this.StartAction(AscDFH.historydescription_Document_CorrectFormTextByFormat);

		let run = form.MakeSingleRunElement(true);
		run.AddText(text);
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
};
CDocument.prototype.private_UpdateFormToPlaceholder = function(form)
{
	let paragraph = form.GetParagraph();
	if (!paragraph || form.IsPlaceHolder())
		return;

	if (!this.IsSelectionLocked(AscCommon.changestype_None, {
		Type      : AscCommon.changestype_2_Element_and_Type,
		Element   : paragraph,
		CheckType : AscCommon.changestype_Paragraph_Content
	}, true, true))
	{
		this.StartAction(AscDFH.historydescription_Document_CorrectFormTextByFormat);
		form.ReplaceContentWithPlaceHolder(false, true);
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
};
CDocument.prototype.UpdateSelectedReviewChanges = function(isSaveCurrentReviewChange)
{
	this.TrackRevisionsManager.BeginCollectChanges(isSaveCurrentReviewChange);
	this.Controller.CollectSelectedReviewChanges(this.TrackRevisionsManager);
	this.TrackRevisionsManager.EndCollectChanges();
};
CDocument.prototype.Document_UpdateUndoRedoState = function()
{
	this.UpdateUndoRedo();
};
CDocument.prototype.private_UpdateUndoRedo = function()
{
	if (true === this.TurnOffInterfaceEvents)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	// TODO: Возможно стоит перенсти эту проверку в класс CHistory и присылать
	//       данные события при изменении значения History.Index

	// Проверяем состояние Undo/Redo

	if (this.IsViewModeInReview())
	{
		this.Api.sync_CanUndoCallback(false);
		this.Api.sync_CanRedoCallback(false);
	}
	else
	{
		var bCanUndo = this.History.Can_Undo();
		if (true !== bCanUndo && this.Api && this.CollaborativeEditing && true === this.CollaborativeEditing.Is_Fast() && true !== this.CollaborativeEditing.Is_SingleUser())
			bCanUndo = this.CollaborativeEditing.CanUndo();

		this.Api.sync_CanUndoCallback(bCanUndo);
		this.Api.sync_CanRedoCallback(this.History.Can_Redo());
		this.Api.CheckChangedDocument();
	}
};
CDocument.prototype.Document_UpdateCopyCutState = function()
{
	if (true === this.TurnOffInterfaceEvents || !this.Api)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	// Во время работы селекта не обновляем состояние
	if (true === this.Selection.Start)
		return;

	this.Api.sync_CanCopyCutCallback(this.Can_CopyCut());
};
CDocument.prototype.Document_UpdateCanAddHyperlinkState = function()
{
	if (true === this.TurnOffInterfaceEvents)
		return;

	if (true === AscCommon.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	// Проверяем можно ли добавить гиперссылку
	this.Api.sync_CanAddHyperlinkCallback(this.CanAddHyperlink(false));
};
CDocument.prototype.Document_UpdateSectionPr = function()
{
	if (true === this.TurnOffInterfaceEvents)
		return;

	if (true === this.CollaborativeEditing.Get_GlobalLockSelection())
		return;

	// Обновляем ориентацию страницы
	this.Api.sync_PageOrientCallback(this.Get_DocumentOrientation());

	// Обновляем размер страницы
	var PageSize = this.Get_DocumentPageSize();
	this.Api.sync_DocSizeCallback(PageSize.W, PageSize.H);

	// Обновляем настройки колонок
	var CurPos = this.CurPos.ContentPos;
	var SectPr = this.SectionsInfo.Get_SectPr(CurPos).SectPr;

	if (SectPr)
	{
		var ColumnsPr = new CDocumentColumnsProps();
		ColumnsPr.From_SectPr(SectPr);
		this.Api.sync_ColumnsPropsCallback(ColumnsPr);
		this.Api.sync_LineNumbersPropsCollback(SectPr.GetLineNumbers());
		this.Api.sync_SectionPropsCallback(new CDocumentSectionProps(SectPr, this));
	}
};
CDocument.prototype.Get_ColumnsProps = function()
{
	// Обновляем настройки колонок
	var CurPos = this.CurPos.ContentPos;
	var SectPr = this.SectionsInfo.Get_SectPr(CurPos).SectPr;

	var ColumnsPr = new CDocumentColumnsProps();
	if (SectPr)
	{
		ColumnsPr.From_SectPr(SectPr);
	}

	return ColumnsPr;
};


CDocument.prototype.GetWatermarkProps = function()
{
    var SectionPageInfo = this.Get_SectionPageNumInfo(this.CurPage);
    var bFirst = SectionPageInfo.bFirst;
    var bEven  = SectionPageInfo.bEven;
    var HdrFtr = this.Get_SectionHdrFtr(this.CurPage, false, false);
    var Header = HdrFtr.Header;
    var oProps;
    if (null === Header)
    {
        oProps = new Asc.CAscWatermarkProperties();
        oProps.put_Type(Asc.c_oAscWatermarkType.None);
        oProps.put_Api(this.Api);
        return oProps;
    }
    var oWatermark = Header.FindWatermark();
    if(oWatermark)
    {
        oProps = oWatermark.GetWatermarkProps();
        oProps.put_Api(this.Api);
        return oProps;
    }
    oProps = new Asc.CAscWatermarkProperties();
    oProps.put_Type(Asc.c_oAscWatermarkType.None);
    oProps.put_Api(this.Api);
    return oProps;
};

CDocument.prototype.SetWatermarkProps = function(oProps)
{
	if (!this.CanPerformAction())
		return;

	this.StartAction(AscDFH.historydescription_Document_AddWatermark);
	const SectionPageInfo = this.Get_SectionPageNumInfo(this.CurPage);
	const bFirst = SectionPageInfo.bFirst;
	const bEven  = SectionPageInfo.bEven;
	const HdrFtr = this.Get_SectionHdrFtr(this.CurPage, bFirst, bEven);
	let Header = HdrFtr.Header;
	if(null === Header)
	{
		if(Asc.c_oAscWatermarkType.None === oProps.get_Type())
		{
			this.FinalizeAction(true);
			return;
		}
		Header = this.Create_SectionHdrFtr(hdrftr_Header, this.CurPage);
	}
	let oWatermark = Header.FindWatermark();
	if(oWatermark)
	{
		if(oWatermark.GraphicObj.selected)
		{
			this.RemoveSelection(true);
		}
		oWatermark.Remove_FromDocument(false);
	}
	oWatermark = this.DrawingObjects.createWatermark(oProps);
	if(oWatermark)
	{
		const oDocState = this.Get_SelectionState2();
		const oContent = Header.Content;
		let oWatermarkCC = null;
		const aAllContentControls = oContent.GetAllContentControls();
		const nCount = aAllContentControls.length;
		for(let nContentControl = 0; nContentControl < nCount; ++nContentControl)
		{
			let oContentControl = aAllContentControls[nContentControl];
			let oDocPart = oContentControl.Pr.DocPartObj;
			if(oDocPart.Gallery === "Watermarks" && oDocPart.Unique)
			{
				oWatermarkCC = oContentControl;
				break;
			}
		}
		if(!oWatermarkCC)
		{
			oWatermarkCC = oContent.AddContentControl(c_oAscSdtLevelType.Inline);
			oWatermarkCC.SetDocPartObj(undefined, "Watermarks", true);
		}
		if(oWatermarkCC.IsBlockLevel())
		{
			oWatermarkCC.AddToParagraph(oWatermark);
		}
		else
		{
			oWatermarkCC.Add(oWatermark);
		}
		this.Set_SelectionState2(oDocState);
	}
	this.Recalculate();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
	this.Document_UpdateRulersState();
	this.FinalizeAction(true);
};

/**
 * Отключаем отсылку сообщений в интерфейс.
 */
CDocument.prototype.TurnOff_InterfaceEvents = function()
{
	this.TurnOffInterfaceEvents = true;
};
/**
 * Включаем отсылку сообщений в интерфейс.
 *
 * @param {bool} bUpdate Обновлять ли интерфейс
 */
CDocument.prototype.TurnOn_InterfaceEvents = function(bUpdate)
{
	this.TurnOffInterfaceEvents = false;

	if (true === bUpdate)
	{
		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
		this.Document_UpdateRulersState();
	}
};
CDocument.prototype.TurnOff_RecalculateCurPos = function()
{
	this.TurnOffRecalcCurPos = true;
};
CDocument.prototype.TurnOn_RecalculateCurPos = function(bUpdate)
{
	this.TurnOffRecalcCurPos = false;

	if (true === bUpdate)
		this.Document_UpdateSelectionState();
};
CDocument.prototype.Can_CopyCut = function()
{
	var bCanCopyCut = false;

	var LogicDocument  = null;
	var DrawingObjects = null;

	var nDocPosType = this.GetDocPosType();
	if (docpostype_HdrFtr === nDocPosType)
	{
		var CurHdrFtr = this.HdrFtr.CurHdrFtr;
		if (null !== CurHdrFtr)
		{
			if (docpostype_DrawingObjects === CurHdrFtr.Content.GetDocPosType())
				DrawingObjects = this.DrawingObjects;
			else
				LogicDocument = CurHdrFtr.Content;
		}
	}
	else if (docpostype_DrawingObjects === nDocPosType)
	{
		DrawingObjects = this.DrawingObjects;
	}
	else if (docpostype_Footnotes === nDocPosType)
	{
		if (0 === this.Footnotes.Selection.Direction)
		{
			var oFootnote = this.Footnotes.GetCurFootnote();
			if (oFootnote)
			{
				if (docpostype_DrawingObjects === oFootnote.GetDocPosType())
				{
					DrawingObjects = this.DrawingObjects;
				}
				else
				{
					LogicDocument = oFootnote;
				}
			}
		}
		else
		{
			return true;
		}
	}
	else if (docpostype_Endnotes === nDocPosType)
	{
		if (0 === this.Endnotes.Selection.Direction)
		{
			var oEndnote = this.Endnotes.GetCurEndnote();
			if (oEndnote)
			{
				if (docpostype_DrawingObjects === oEndnote.GetDocPosType())
				{
					DrawingObjects = this.DrawingObjects;
				}
				else
				{
					LogicDocument = oEndnote;
				}
			}
		}
		else
		{
			return true;
		}
	}
	else //if (docpostype_Content === nDocPosType)
	{
		LogicDocument = this;
	}

	if (null !== DrawingObjects)
	{
		if (true === DrawingObjects.isSelectedText())
			LogicDocument = DrawingObjects.getTargetDocContent();
		else
			bCanCopyCut = true;
	}

	if (null !== LogicDocument)
	{
		if (true === LogicDocument.IsSelectionUse() && true !== LogicDocument.IsSelectionEmpty(true))
		{
			if (selectionflag_Numbering === LogicDocument.Selection.Flag)
				bCanCopyCut = false;
			else if (LogicDocument.Selection.StartPos !== LogicDocument.Selection.EndPos)
				bCanCopyCut = true;
			else
				bCanCopyCut = LogicDocument.Content[LogicDocument.Selection.StartPos].Can_CopyCut();
		}
	}

	return bCanCopyCut;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с номерами страниц
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.Get_StartPage_Absolute = function()
{
	return 0;
};
CDocument.prototype.Get_StartPage_Relative = function()
{
	return 0;
};
CDocument.prototype.Get_AbsolutePage = function(CurPage)
{
	return CurPage;
};
CDocument.prototype.Set_CurPage = function(PageNum)
{
	this.CurPage = Math.min(this.Pages.length - 1, Math.max(0, PageNum));
};
CDocument.prototype.Get_CurPage = function()
{
	// Работаем с колонтитулом
	if (docpostype_HdrFtr === this.GetDocPosType())
		return this.HdrFtr.Get_CurPage();

	return this.CurPage;
};
//----------------------------------------------------------------------------------------------------------------------
// Undo/Redo функции
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.Create_NewHistoryPoint = function(nDescription, oSelectionState)
{
	this.History.Create_NewPoint(nDescription, oSelectionState);
};
CDocument.prototype.Document_Undo = function(Options)
{
	// Проверка !this.IsViewModeInReview() нужна, чтобы во вьювере было переключение режима просмотра рецензирования
	if (true === AscCommon.CollaborativeEditing.Get_GlobalLock() && true !== this.IsFillingFormMode() && !this.IsViewModeInReview())
		return;

	// Нужно сбрасывать, т.к. после Undo/Redo данные списки у нас будут в глобальной таблице, но не такие, какие нужны
	this.NumberingApplicator.ResetLast();

	if (true !== this.History.Can_Undo() && this.Api && this.CollaborativeEditing && true === this.CollaborativeEditing.Is_Fast() && true !== this.CollaborativeEditing.Is_SingleUser())
	{
		if (this.CollaborativeEditing.CanUndo() && true === this.Api.canSave)
		{
			this.CollaborativeEditing.Set_GlobalLock(true);
			this.Api.forceSaveUndoRequest = true;
		}
	}
	else
	{
		if (this.History.Can_Undo())
		{
			this.StartUndoRedoAction();
			let changes = this.History.Undo(Options);
			this.UpdateAfterUndoRedo(changes);
			this.FinalizeUndoRedoAction();
		}
	}

	if (this.IsFillingFormMode())
		this.Api.sync_OnAllRequiredFormsFilled(this.FormsManager.IsAllRequiredFormsFilled());
};
CDocument.prototype.Document_Redo = function()
{
	if (true === AscCommon.CollaborativeEditing.Get_GlobalLock() && true !== this.IsFillingFormMode())
		return;

	// Нужно сбрасывать, т.к. после Undo/Redo данные списки у нас будут в глобальной таблице, но не такие, какие нужны
	this.NumberingApplicator.ResetLast();

	if (this.History.Can_Redo())
	{
		this.StartUndoRedoAction();
		let changes = this.History.Redo();
		this.UpdateAfterUndoRedo(changes);
		this.FinalizeUndoRedoAction();
	}

	if (this.IsFillingFormMode())
		this.Api.sync_OnAllRequiredFormsFilled(this.FormsManager.IsAllRequiredFormsFilled());
};
CDocument.prototype.UpdateAfterUndoRedo = function(changes)
{
	if (this.OFormDocument)
		this.OFormDocument.onUndoRedo();
	
	this.DocumentOutline.UpdateAll(); // TODO: надо бы подумать как переделать на более легкий пересчет
	this.Comments.UpdateAll();        // TODO: Надо переделать как на Start/Finalize
	this.DrawingObjects.TurnOnCheckChartSelection();
	this.RecalculateByChanges(changes);

	this.UpdateSelection();
	this.UpdateInterface();
	this.UpdateRulers();
};
CDocument.prototype.GetSelectionState = function()
{
	var DocState    = {};
	DocState.CurPos = {
		X          : this.CurPos.X,
		Y          : this.CurPos.Y,
		ContentPos : this.CurPos.ContentPos,
		RealX      : this.CurPos.RealX,
		RealY      : this.CurPos.RealY,
		Type       : this.CurPos.Type
	};

	DocState.Selection = {

		Start    : false,
		Use      : this.Selection.Use,
		StartPos : this.Selection.StartPos,
		EndPos   : this.Selection.EndPos,
		Flag     : this.Selection.Flag,
		Data     : this.Selection.Data
	};

	DocState.CurPage    = this.CurPage;
	DocState.CurComment = this.Comments.Get_CurrentId();

	var State = null;
	if ((true === this.Api.isStartAddShape || this.Api.isInkDrawerOn()) && docpostype_DrawingObjects === this.GetDocPosType())
	{
		DocState.CurPos.Type     = docpostype_Content;
		DocState.Selection.Start = false;
		DocState.Selection.Use   = false;

		this.Content[DocState.CurPos.ContentPos].RemoveSelection();
		State = this.Content[this.CurPos.ContentPos].GetSelectionState();
	}
	else
	{
		State = this.Controller.GetSelectionState();
	}

	if (null != this.Selection.Data && true === this.Selection.Data.TableBorder)
	{
		DocState.Selection.Data = null;
	}

	State.push(DocState);

	return State;
};
CDocument.prototype.SetSelectionState = function(State)
{
	this.RemoveSelection();

	if (docpostype_DrawingObjects === this.GetDocPosType())
		this.DrawingObjects.resetSelection();

	if (State.length <= 0)
		return;

	var DocState = State[State.length - 1];

	this.CurPos.X          = DocState.CurPos.X;
	this.CurPos.Y          = DocState.CurPos.Y;
	this.CurPos.ContentPos = DocState.CurPos.ContentPos;
	this.CurPos.RealX      = DocState.CurPos.RealX;
	this.CurPos.RealY      = DocState.CurPos.RealY;
	this.SetDocPosType(DocState.CurPos.Type);

	this.Selection.Start    = false;
	this.Selection.Use      = DocState.Selection.Use;
	this.Selection.StartPos = DocState.Selection.StartPos;
	this.Selection.EndPos   = DocState.Selection.EndPos;
	this.Selection.Flag     = DocState.Selection.Flag;
	this.Selection.Data     = DocState.Selection.Data;

	this.Selection.DragDrop.Flag = 0;
	this.Selection.DragDrop.Data = null;

	this.CurPage = DocState.CurPage;
	this.Comments.Set_Current(DocState.CurComment);

	this.Controller.SetSelectionState(State, State.length - 2);
};
CDocument.prototype.Get_ParentObject_or_DocumentPos = function(Index)
{
	return {Type : AscDFH.historyitem_recalctype_Inline, Data : Index};
};
CDocument.prototype.Refresh_RecalcData = function(oData)
{
	var nChangePos = -1;
	switch (oData.Type)
	{
		case AscDFH.historyitem_Document_AddItem:
		case AscDFH.historyitem_Document_RemoveItem:
		{
			if (oData instanceof CChangesDocumentAddItem || oData instanceof CChangesDocumentRemoveItem)
				nChangePos = oData.GetMinPos();
			else
				nChangePos = oData.Pos;

			break;
		}

		case AscDFH.historyitem_Document_DefaultTab:
		case AscDFH.historyitem_Document_EvenAndOddHeaders:
		case AscDFH.historyitem_Document_MathSettings:
		case AscDFH.historyitem_Document_Settings_GutterAtTop:
		case AscDFH.historyitem_Document_Settings_MirrorMargins:
		{
			nChangePos = 0;
			break;
		}
	}

	if (-1 !== nChangePos)
	{
		this.History.RecalcData_Add({
			Type : AscDFH.historyitem_recalctype_Inline,
			Data : {Pos : nChangePos, PageNum : 0}
		});
	}
};
CDocument.prototype.Refresh_RecalcData2 = function(nIndex, nPageRel)
{
	if (-1 === nIndex)
		return;

	this.History.RecalcData_Add({
		Type : AscDFH.historyitem_recalctype_Inline,
		Data : {
			Pos     : nIndex,
			PageNum : nPageRel
		}
	});
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с гиперссылками
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.AddHyperlink = function(HyperProps)
{
	// Проверку, возможно ли добавить гиперссылку, должны были вызвать до этой функции
	if (HyperProps.Text && "" !== HyperProps.Text && true === this.IsSelectionUse() && this.GetSelectedText() !== HyperProps.Text)
	{
		// Корректировка в данном случае пройдет при добавлении гиперссылки.
		var SelectionInfo = this.GetSelectedElementsInfo();
		var Para          = SelectionInfo.GetParagraph();
		if (null !== Para)
			HyperProps.TextPr = Para.Get_TextPr(Para.Get_ParaContentPos(true, true));

		this.Remove(1, false, false, true);
		this.RemoveTextSelection();
	}

	let hyperlink = this.Controller.AddHyperlink(HyperProps);
	if (hyperlink)
	{
		this.RemoveSelection();
		hyperlink.MoveCursorOutsideElement(false);
	}
	
	this.Recalculate();
	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
};
CDocument.prototype.ModifyHyperlink = function(oHyperProps)
{
	var sText    = oHyperProps.get_Text(),
		sValue   = oHyperProps.get_Value(),
		sToolTip = oHyperProps.get_ToolTip(),
		sAnchor  = oHyperProps.get_Bookmark();

	var oClass = oHyperProps.get_InternalHyperlink();
	if (oClass instanceof ParaHyperlink)
	{
		var oHyperlink = oClass;

		if (undefined !== sAnchor && null !== sAnchor && "" !== sAnchor)
		{
			oHyperlink.SetAnchor(sAnchor);
			oHyperlink.SetValue("");
		}
		else if (undefined !== sValue && null !== sValue)
		{
			oHyperlink.SetValue(sValue);
			oHyperlink.SetAnchor("");
		}

		if (undefined !== sToolTip && null !== sToolTip)
			oHyperlink.SetToolTip(sToolTip);

		if (null !== sText)
		{
			var oHyperRun = new ParaRun(oHyperlink.GetParagraph());
			oHyperRun.Set_Pr(oHyperlink.GetTextPr().Copy());
			oHyperRun.Set_Color(undefined);
			oHyperRun.SetUnderline(undefined);
			oHyperRun.Set_RStyle(this.GetStyles().GetDefaultHyperlink());
			oHyperRun.AddText(sText);

			oHyperlink.RemoveSelection();
			oHyperlink.RemoveAll();
			oHyperlink.AddToContent(0, oHyperRun, false);

			this.RemoveSelection();
			oHyperlink.MoveCursorOutsideElement(false);
		}
	}
	else if (oClass instanceof CFieldInstructionHYPERLINK)
	{
		var oInstruction = oClass;
		var oComplexField = oInstruction.GetComplexField();
		if (!oComplexField || oComplexField)
		{
			if (undefined !== sAnchor && null !== sAnchor && "" !== sAnchor)
			{
				oInstruction.SetBookmarkName(sAnchor);
				oInstruction.SetLink("");
			}
			else if (undefined !== sValue && null !== sValue)
			{
				oInstruction.SetLink(sValue);
				oInstruction.SetBookmarkName("");
			}

			if (undefined !== sToolTip && null !== sToolTip)
				oInstruction.SetToolTip(sToolTip);

			oComplexField.SelectFieldCode();
			var sInstruction = oInstruction.ToString();
			for (var oIterator = sInstruction.getUnicodeIterator(); oIterator.check (); oIterator.next())
			{
				this.AddToParagraph(new ParaInstrText(oIterator.value()));
			}

			if (null !== sText)
			{
				oComplexField.SelectFieldValue();
				var oTextPr = this.GetDirectTextPr();
				this.AddText(sText);
				oComplexField.SelectFieldValue();
				this.AddToParagraph(new ParaTextPr(oTextPr));
			}
			oComplexField.MoveCursorOutsideElement(false);
		}
	}
	else
	{
		return;
	}

	this.Recalculate();
    this.Document_UpdateSelectionState();
    this.Document_UpdateInterfaceState();
};
CDocument.prototype.RemoveHyperlink = function(oHyperProps)
{
	var oClass = oHyperProps.get_InternalHyperlink();
	if (oClass instanceof ParaHyperlink)
	{
		var oHyperlink = oClass;

		var oParent      = oHyperlink.GetParent();
		var nPosInParent = oHyperlink.GetPosInParent(oParent);

		if (!oParent)
			return;

		oParent.RemoveFromContent(nPosInParent, 1);

		var oTextPr       = new CTextPr();
		oTextPr.RStyle    = null;
		oTextPr.Underline = null;
		oTextPr.Color     = null;
		oTextPr.Unifill   = null;

		var oElement = null;
		for (var nPos = 0, nCount = oHyperlink.GetElementsCount(); nPos < nCount; ++nPos)
		{
			oElement = oHyperlink.GetElement(nPos);
			oParent.AddToContent(nPosInParent + nPos, oElement);
			oElement.ApplyTextPr(oTextPr, undefined, true);
		}

		this.RemoveSelection();
		if (oElement)
		{
			oElement.SetThisElementCurrent();
			oElement.MoveCursorToEndPos();
		}
	}
	else if (oClass instanceof CFieldInstructionHYPERLINK)
	{
		var oInstruction = oClass;
		var oComplexField = oInstruction.GetComplexField();
		if (!oComplexField || oComplexField)
		{
			var oTextPr       = new CTextPr();
			oTextPr.RStyle    = null;
			oTextPr.Underline = null;
			oTextPr.Color     = null;
			oTextPr.Unifill   = null;

			oComplexField.SelectFieldValue();
			this.AddToParagraph(new ParaTextPr(oTextPr));

			oComplexField.RemoveFieldWrap();
		}
	}
	else
	{
		return;
	}

	//this.Controller.RemoveHyperlink();
	this.Recalculate();
	this.Document_UpdateSelectionState();
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.CanAddHyperlink = function(bCheckInHyperlink)
{
	return this.Controller.CanAddHyperlink(bCheckInHyperlink);
};
/**
 * Проверяем, находимся ли мы в гиперссылке сейчас.
 * @param bCheckEnd
 * @returns {*}
 */
CDocument.prototype.IsCursorInHyperlink = function(bCheckEnd)
{
	if (undefined === bCheckEnd)
		bCheckEnd = true;

	return this.Controller.IsCursorInHyperlink(bCheckEnd);
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с совместным редактирования
//----------------------------------------------------------------------------------------------------------------------
/**
 * Проверяем можно ли выполнять действия
 *
 * Данная функция вызывается в основной функции проверки лока совместного редактирования, но есть действия, которые
 * не требуют проверки лока совместки, в таких случая перед выполнением действия надо вызывать эту проверку
 * @param [isIgnoreCanEditFlag=false]
 * @returns {boolean}
 */
CDocument.prototype.CanPerformAction = function(isIgnoreCanEditFlag)
{
	return !((!this.CanEdit() && true !== isIgnoreCanEditFlag) || (true === this.CollaborativeEditing.Get_GlobalLock()));
};
CDocument.prototype.Document_Is_SelectionLocked = function(CheckType, AdditionalData, DontLockInFastMode, isIgnoreCanEditFlag, fCallback)
{
	if (!this.CanPerformAction(isIgnoreCanEditFlag))
	{
		if (fCallback)
			fCallback(true);

		return true;
	}

	this.CollaborativeEditing.OnStart_CheckLock();

	this.private_DocumentIsSelectionLocked(CheckType);

	if (AdditionalData)
	{
		if (undefined !== AdditionalData.length)
		{
			for (var nIndex = 0, nCount = AdditionalData.length; nIndex < nCount; ++nIndex)
			{
				this.private_IsSelectionLockedAdditional(AdditionalData[nIndex]);
			}
		}
		else
		{
			this.private_IsSelectionLockedAdditional(AdditionalData);
		}
	}

	var isLocked = this.CollaborativeEditing.OnEnd_CheckLock(DontLockInFastMode, fCallback);

	if (true === isLocked && !fCallback)
	{
		this.UpdateSelection();
		this.UpdateInterface();
	}

	return isLocked;
};
CDocument.prototype.private_IsSelectionLockedAdditional = function(oAdditionalData)
{
	if (oAdditionalData)
	{
		if (AscCommon.changestype_2_InlineObjectMove === oAdditionalData.Type)
		{
			var PageNum = oAdditionalData.PageNum;
			var X       = oAdditionalData.X;
			var Y       = oAdditionalData.Y;

			var NearestPara = this.Get_NearestPos(PageNum, X, Y).Paragraph;
			NearestPara.Document_Is_SelectionLocked(AscCommon.changestype_Document_Content);
		}
		else if (AscCommon.changestype_2_HdrFtr === oAdditionalData.Type)
		{
			this.HdrFtr.Document_Is_SelectionLocked(AscCommon.changestype_HdrFtr);
		}
		else if (AscCommon.changestype_2_Comment === oAdditionalData.Type)
		{
			this.Comments.Document_Is_SelectionLocked(oAdditionalData.Id);
		}
		else if (AscCommon.changestype_2_Element_and_Type === oAdditionalData.Type)
		{
			oAdditionalData.Element.Document_Is_SelectionLocked(oAdditionalData.CheckType, false);
		}
		else if (AscCommon.changestype_2_ElementsArray_and_Type === oAdditionalData.Type)
		{
			var Count = oAdditionalData.Elements.length;
			for (var Index = 0; Index < Count; Index++)
			{
				oAdditionalData.Elements[Index].Document_Is_SelectionLocked(oAdditionalData.CheckType, false);
			}
		}
		else if (AscCommon.changestype_2_Element_and_Type_Array === oAdditionalData.Type)
		{
			var nCount = Math.min(oAdditionalData.Elements.length, oAdditionalData.CheckTypes.length);
			for (var nIndex = 0; nIndex < nCount; ++nIndex)
			{
				oAdditionalData.Elements[nIndex].Document_Is_SelectionLocked(oAdditionalData.CheckTypes[nIndex], false);
			}
		}
		else if (AscCommon.changestype_2_AdditionalTypes === oAdditionalData.Type)
		{
			var Count = oAdditionalData.Types.length;
			for (var Index = 0; Index < Count; ++Index)
			{
				this.private_DocumentIsSelectionLocked(oAdditionalData.Types[Index]);
			}
		}
	}
};
CDocument.prototype.IsSelectionLocked = function(nCheckType, oAdditionalData, isDontLockInFastMode, isIgnoreCanEditFlag, fCallback)
{
	return this.Document_Is_SelectionLocked(nCheckType, oAdditionalData, isDontLockInFastMode, isIgnoreCanEditFlag, fCallback);
};
/**
 * Начинаем составную проверку на залоченность объектов
 * @param [isIgnoreCanEditFlag=false] игнорируем ли запрет на редактирование
 * @returns {boolean} началась ли проверка залоченности
 */
CDocument.prototype.StartSelectionLockCheck = function(isIgnoreCanEditFlag)
{
	if (!this.CanEdit() && true !== isIgnoreCanEditFlag)
		return false;

	if (true === this.CollaborativeEditing.Get_GlobalLock())
		return false;

	this.CollaborativeEditing.OnStart_CheckLock();

	return true;
};
/**
 * Производим шаг проверки залоченности с заданными типом
 * @param nCheckType {number} тип проверки
 * @param isClearOnLock {boolean} очищать ли общий массив локов от последней проверки, в случае залоченности
 * @returns {boolean} залочен ли редактор на выполнение данного шага
 */
CDocument.prototype.ProcessSelectionLockCheck = function(nCheckType, isClearOnLock)
{
	if (isClearOnLock)
		this.CollaborativeEditing.OnStartCheckLockInstance();

	this.private_DocumentIsSelectionLocked(nCheckType);

	if (isClearOnLock)
		return this.CollaborativeEditing.OnEndCheckLockInstance();

	return false;
};
/**
 * Заканчиваем процесс составной проверки залоченности объектов
 * @param [isDontLockInFastMode=false] {boolean} нужно ли лочить в быстром режиме совместного редактирования
 * @returns {boolean} залочен ли редактор на выполнение данного составного действия
 */
CDocument.prototype.EndSelectionLockCheck = function(isDontLockInFastMode)
{
	var isLocked = this.CollaborativeEditing.OnEnd_CheckLock(isDontLockInFastMode);
	if (isLocked)
	{
		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();
	}

	return isLocked;
};
CDocument.prototype.private_DocumentIsSelectionLocked = function(CheckType)
{
	if (AscCommon.changestype_None !== CheckType)
	{
		if (AscCommon.changestype_Document_SectPr === CheckType)
		{
			this.Lock.Check(this.Get_Id());
		}
		else if (AscCommon.changestype_Document_Styles === CheckType)
		{
			this.Styles.Lock.Check(this.Styles.Get_Id());
		}
		else if (AscCommon.changestype_ColorScheme === CheckType)
		{
			this.DrawingObjects.Lock.Check(this.DrawingObjects.Get_Id());
		}
		else if (AscCommon.changestype_CorePr === CheckType)
		{
			if (this.Core)
			{
				this.Core.Lock.Check(this.Core.Get_Id());
			}
		}
		else if (AscCommon.changestype_DocumentProtection === CheckType)
		{
			if (this.Settings.DocumentProtection)
			{
				this.Settings.DocumentProtection.Lock.Check(this.Settings.DocumentProtection.Get_Id());
			}
		}
		else if (AscCommon.changestype_Document_Settings === CheckType)
		{
			this.Lock.Check(this.GetId());
		}
		else
		{
			this.Controller.IsSelectionLocked(CheckType);
		}
	}
};
CDocument.prototype.controller_IsSelectionLocked = function(CheckType)
{
	switch (this.Selection.Flag)
	{
		case selectionflag_Common :
		{
			if (true === this.Selection.Use)
			{
				var StartPos = ( this.Selection.StartPos > this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos );
				var EndPos   = ( this.Selection.StartPos > this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos );

				if (StartPos != EndPos && AscCommon.changestype_Delete === CheckType)
					CheckType = AscCommon.changestype_Remove;

				for (var Index = StartPos; Index <= EndPos; Index++)
					this.Content[Index].Document_Is_SelectionLocked(CheckType);
			}
			else
			{
				var CurElement = this.Content[this.CurPos.ContentPos];

				if (AscCommon.changestype_Document_Content_Add === CheckType && CurElement.IsParagraph() && CurElement.IsCursorAtEnd() && CurElement.Lock.Is_Locked())
					AscCommon.CollaborativeEditing.Add_CheckLock(false);
				else
					this.Content[this.CurPos.ContentPos].Document_Is_SelectionLocked(CheckType);
			}

			break;
		}
		case selectionflag_Numbering:
		{
			var oNumPr = this.Selection.Data.CurPara.GetNumPr();
			if (oNumPr)
			{
				var oNum = this.GetNumbering().GetNum(oNumPr.NumId);
				oNum.IsSelectionLocked(CheckType);
			}

			this.Content[this.CurPos.ContentPos].Document_Is_SelectionLocked(CheckType);

			break;
		}
	}
};
CDocument.prototype.Get_SelectionState2 = function()
{
	this.RemoveSelection();

	// Сохраняем Id ближайшего элемента в текущем классе
	var State       = new CDocumentSelectionState();
	var nDocPosType = this.GetDocPosType();
	if (docpostype_HdrFtr === nDocPosType)
	{
		State.Type = docpostype_HdrFtr;

		if (null != this.HdrFtr.CurHdrFtr)
			State.Id = this.HdrFtr.CurHdrFtr.Get_Id();
		else
			State.Id = null;
	}
	else if (docpostype_DrawingObjects === nDocPosType)
	{
		// TODO: запрашиваем параграф текущего выделенного элемента
		var X       = 0;
		var Y       = 0;
		var PageNum = this.CurPage;

		var ContentPos = this.Internal_GetContentPosByXY(X, Y, PageNum);

		State.Type = docpostype_Content;
		State.Id   = this.Content[ContentPos].GetId();
	}
	else if (docpostype_Footnotes === nDocPosType)
	{
		var oFootnote = this.Footnotes.GetCurFootnote();

		State.Type = docpostype_Footnotes;
		State.Id   = oFootnote ? oFootnote.GetId() : null;
	}
	else if (docpostype_Endnotes === nDocPosType)
	{
		var oEndnote = this.Footnotes.GetCurEndnote();

		State.Type = docpostype_Endnotes;
		State.Id   = oEndnote ? oEndnote.GetId() : null;
	}
	else // if (docpostype_Content === nDocPosType)
	{
		State.Id   = this.Get_Id();
		State.Type = docpostype_Content;

		var Element = this.Content[this.CurPos.ContentPos];
		State.Data  = Element.Get_SelectionState2();
	}

	return State;
};
CDocument.prototype.Set_SelectionState2 = function(State)
{
	this.RemoveSelection();

	var Id = State.Id;
	if (docpostype_HdrFtr === State.Type)
	{
		this.SetDocPosType(docpostype_HdrFtr);

		if (null === Id || true != this.HdrFtr.Set_CurHdrFtr_ById(Id))
		{
			this.SetDocPosType(docpostype_Content);
			this.CurPos.ContentPos = 0;

			this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
		}
	}
	else if (docpostype_Footnotes === State.Type)
	{
		this.SetDocPosType(docpostype_Footnotes);
		var oFootnote = g_oTableId.Get_ById(State.Id);
		if (oFootnote && true === this.Footnotes.IsUseInDocument(State.Id))
		{
			this.Footnotes.private_SetCurrentFootnoteNoSelection(oFootnote);
			oFootnote.MoveCursorToStartPos(false);
		}
		else
		{
			this.EndFootnotesEditing();
		}
	}
	else if (docpostype_Endnotes === State.Type)
	{
		this.SetDocPosType(docpostype_Endnotes);
		var oEndnote = g_oTableId.Get_ById(State.Id);
		if (oEndnote && true === this.Endnotes.IsUseInDocument(State.Id))
		{
			this.Endnotes.private_SetCurrentEndnoteNoSelection(oEndnote);
			oEndnote.MoveCursorToStartPos(false);
		}
		else
		{
			this.EndEndnotesEditing();
		}
	}
	else // if ( docpostype_Content === State.Type )
	{
		var CurId = State.Data.Id;

		var bFlag = false;

		var Pos = 0;

		// Найдем элемент с Id = CurId
		var Count = this.Content.length;
		for (Pos = 0; Pos < Count; Pos++)
		{
			if (this.Content[Pos].GetId() == CurId)
			{
				bFlag = true;
				break;
			}
		}

		if (true !== bFlag)
		{
			var TempElement = g_oTableId.Get_ById(CurId);
			Pos             = ( null != TempElement ? TempElement.Index : 0 );
			Pos             = Math.max(0, Math.min(Pos, this.Content.length - 1));
		}

		this.Selection.Start    = false;
		this.Selection.Use      = false;
		this.Selection.StartPos = Pos;
		this.Selection.EndPos   = Pos;
		this.Selection.Flag     = selectionflag_Common;

		this.SetDocPosType(docpostype_Content);
		this.CurPos.ContentPos = Pos;

		if (true !== bFlag)
			this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
		else
			this.Content[this.CurPos.ContentPos].SetSelectionState2(State.Data);
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с комментариями
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.AddComment = function(CommentData, isForceGlobal)
{
	if (true === isForceGlobal || true != this.CanAddComment())
	{
		CommentData.Set_QuoteText(null);
		var Comment = new AscCommon.CComment(this.Comments, CommentData);
		this.Comments.Add(Comment);

		// Обновляем информацию для Undo/Redo
		this.UpdateInterface();
	}
	else
	{
		// TODO: Как будет реализовано добавление комментариев к объекту, добавить проверку на выделение объекта

		var QuotedText = this.GetSelectedText(false);
		if (null === QuotedText || "" === QuotedText)
		{
			var oParagraph = this.GetCurrentParagraph();
			if (oParagraph && oParagraph.SelectCurrentWord())
			{
				QuotedText = this.GetSelectedText(false);
				if (null === QuotedText)
					QuotedText = "";
			}
			else
			{
				QuotedText = "";
			}
		}
		CommentData.Set_QuoteText(QuotedText);

		var Comment = new AscCommon.CComment(this.Comments, CommentData);
		this.Comments.Add(Comment);
		this.Controller.AddComment(Comment);

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
	}

	return Comment;
};
CDocument.prototype.EditComment = function(Id, CommentData)
{
	if (!this.Comments.IsUseSolved() && this.Comments.Get_CurrentId() === Id)
	{
		var oComment = this.Comments.Get_ById(Id);
		if (oComment && !oComment.IsSolved() && CommentData.IsSolved())
		{
			this.Comments.Set_Current(null);
			this.Api.sync_HideComment();
			this.DrawingDocument.ClearCachePages();
			this.DrawingDocument.FirePaint();
		}
	}

	this.Comments.Set_CommentData(Id, CommentData);
	this.Document_UpdateInterfaceState();
};
CDocument.prototype.RemoveComment = function(Id, bSendEvent, bRecalculate)
{
	if (null === Id)
		return;

	if (true === this.Comments.Remove_ById(Id))
	{
		if (true === bRecalculate)
		{
			// TODO: Продумать как избавиться от пересчета при удалении комментария
			this.Recalculate();
			this.Document_UpdateInterfaceState();
		}

		if (true === bSendEvent)
			this.Api.sync_RemoveComment(Id);
	}
};
CDocument.prototype.CanAddComment = function()
{
	if (!this.CanEdit() && !this.IsEditCommentsMode())
		return false;

	if (true !== this.Comments.Is_Use())
		return false;

	return this.Controller.CanAddComment();
};
CDocument.prototype.SelectComment = function(sId, isScrollToComment)
{
	var sOldId = this.Comments.Get_CurrentId();
	this.Comments.Set_Current(sId);

	var oComment = this.Comments.Get_ById(sId);
	if (isScrollToComment && oComment)
	{
		this.RemoveSelection();
		oComment.MoveCursorToStart();
		this.UpdateSelection();
		this.UpdateInterface();
	}

	if (sOldId !== sId)
	{
		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();
	}
};
CDocument.prototype.ShowComment = function(arrId)
{
	var CommentsX     = null;
	var CommentsY     = null;
	var arrCommentsId = [];

	for (var nIndex = 0, nCount = arrId.length; nIndex < nCount; ++nIndex)
	{
		var Comment = this.Comments.Get_ById(arrId[nIndex]);
		if (Comment)
		{
			if (null === CommentsX || null === CommentsY)
			{
				var oCoords = this.private_GetCommentWorldAnchorPoint(Comment);
				CommentsX = oCoords.X;
				CommentsY = oCoords.Y;
			}

			arrCommentsId.push(Comment.Get_Id());
		}

	}

	if (null !== CommentsX && null !== CommentsY && arrCommentsId.length > 0)
	{
		this.Api.sync_ShowComment(arrCommentsId, CommentsX, CommentsY);
	}
	else
	{
		this.Api.sync_HideComment();
	}
};
CDocument.prototype.ShowComments = function(isShowSolved)
{
	if (false !== isShowSolved)
		isShowSolved = true;

	this.Comments.Set_Use(true);
	this.Comments.SetUseSolved(isShowSolved);
	this.DrawingDocument.ClearCachePages();
	this.DrawingDocument.FirePaint();
};
CDocument.prototype.HideComments = function()
{
	this.Comments.Set_Use(false);
	this.Comments.SetUseSolved(false);
	this.Comments.Set_Current(null);
	this.DrawingDocument.ClearCachePages();
	this.DrawingDocument.FirePaint();
};
CDocument.prototype.private_GetCommentWorldAnchorPoint = function(oComment)
{
	var nPage      = oComment.m_oStartInfo.PageNum;
	var nY         = oComment.m_oStartInfo.Y;
	var nX         = this.Get_PageLimits(nPage).XLimit;
	var oStartMark = this.TableId.Get_ById(oComment.GetRangeStart());

	if (!oStartMark)
		return {X : 0, Y : 0};

	var oCommentParagraph = oStartMark.GetParagraph();
	if (oCommentParagraph)
	{
		var oTransform = oCommentParagraph.Get_ParentTextTransform();
		if (oTransform)
			nY = oTransform.TransformPointY(oComment.m_oStartInfo.X, oComment.m_oStartInfo.Y);
	}

	if (window["NATIVE_EDITOR_ENJINE"])
		nX = oComment.m_oStartInfo.X;

	var oCoords = this.DrawingDocument.ConvertCoordsToCursorWR(nX, nY, nPage);
	return {X : oCoords.X, Y : oCoords.Y};
};
CDocument.prototype.GetPrevElementEndInfo = function(CurElement)
{
    var PrevElement = CurElement.Get_DocumentPrev();

    if (null !== PrevElement && undefined !== PrevElement)
        return PrevElement.GetEndInfo();
    else
        return null;
};
CDocument.prototype.GetSelectionAnchorPos = function()
{
	var Result = this.Controller.GetSelectionAnchorPos();

	var PageLimit = this.Get_PageLimits(Result.Page);
	Result.X0     = PageLimit.X;
	Result.X1     = PageLimit.XLimit;

	var Coords0 = this.DrawingDocument.ConvertCoordsToCursorWR(Result.X0, Result.Y, Result.Page);
	var Coords1 = this.DrawingDocument.ConvertCoordsToCursorWR(Result.X1, Result.Y, Result.Page);

	return {X0 : Coords0.X, X1 : Coords1.X, Y : Coords0.Y};
};
/**
 * Получаем все комментарии по заданным параметрам
 * @param isMine {boolean} удаляем только все свои комментарии
 * @param isCurrent {boolean} удаляем только все комментарии, находящиеся в текущей позиции
 * @returns {Array.string}
 */
CDocument.prototype.GetAllComments = function(isMine, isCurrent)
{
	var arrCommentsId = [];

	if (isCurrent)
	{
		if (true === this.Comments.Is_Use())
		{
			var arrParagraphs = this.GetSelectedParagraphs();
			var oComments     = {};

			for (var nParaIndex = 0, nParasCount = arrParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
			{
				arrParagraphs[nParaIndex].GetCurrentComments(oComments);
			}

			for (var sCommentId in oComments)
			{
				var oComment = this.Comments.Get_ById(sCommentId);
				if (oComment && (!isMine || oComment.IsCurrentUser()))
					arrCommentsId.push(oComment.GetId());
			}
		}
	}
	else
	{
		var oComments = this.Comments.GetAllComments();
		for (var sId in oComments)
		{
			var oComment = oComments[sId];
			if (oComment && (!isMine || oComment.IsCurrentUser()))
				arrCommentsId.push(oComment.GetId());
		}
	}

	return arrCommentsId;
};
/**
 * Сохраняем комментарий, у которого изменилась позиция
 * @param {CComment} oComment
 */
CDocument.prototype.UpdateCommentPosition = function(oComment)
{
	if (AscCommon.g_oIdCounter.m_bLoad)
		return;

	if (!this.Action.Start)
	{
		var oChangedComments = this.Comments.UpdateCommentPosition(oComment);
		
		if (this.Api)
			this.Api.sync_ChangeCommentLogicalPosition(oChangedComments, this.Comments.GetCommentsPositionsCount());

		return;
	}

	if (!this.Action.Additional.CommentPosition)
		this.Action.Additional.CommentPosition = {};

	this.Action.Additional.CommentPosition[oComment.GetId()] = oComment;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с textbox
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.TextBox_Put = function(sText, rFonts)
{
	if (false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_AddTextFromTextBox);

		// Отключаем пересчет, включим перед последним добавлением. Поскольку,
		// у нас все добавляется в 1 параграф, так можно делать.
		this.Start_SilentMode();

		if (undefined === rFonts)
		{
			this.AddText(sText);
		}
		else
		{
			var Para = this.GetCurrentParagraph();
			if (null === Para)
				return;

			var RunPr = Para.Get_TextPr();
			if (null === RunPr || undefined === RunPr)
				RunPr = new CTextPr();

			RunPr.RFonts = rFonts;

			var Run = new ParaRun(Para);
			Run.Set_Pr(RunPr);
			Run.AddText(sText);
			Para.Add(Run);
		}

		this.End_SilentMode(true);

		this.FinalizeAction();
	}
};
//----------------------------------------------------------------------------------------------------------------------
// события вьюера
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.Viewer_OnChangePosition = function()
{
	var oComment = this.Comments.Get_Current();
	if (oComment)
	{
		var oAnchorPoint = this.private_GetCommentWorldAnchorPoint(oComment);
		this.Api.sync_UpdateCommentPosition(oComment.GetId(), oAnchorPoint.X, oAnchorPoint.Y);
	}
    window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Update_Position();
	this.TrackRevisionsManager.UpdateSelectedChangesPosition(this.Api);
	this.MathTrackHandler.OnChangePosition();
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с секциями
//----------------------------------------------------------------------------------------------------------------------
/**
 * Обновляем информацию о всех секциях в данном документе
 */
CDocument.prototype.UpdateAllSectionsInfo = function()
{
	this.SectionsInfo.Clear();

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var Element = this.Content[Index];
		if (type_Paragraph === Element.GetType() && undefined !== Element.Get_SectionPr())
			this.SectionsInfo.Add(Element.Get_SectionPr(), Index);
	}

	this.SectionsInfo.Add(this.SectPr, Count);

	// Когда полностью обновляются секции надо пересчитывать с самого начала
	this.RecalcInfo.Set_NeedRecalculateFromStart(true);
};
/**
 * Обновляем информацию о заданной секции
 * @param oSectPr {CSectionPr} - Если не задано, значит добавляется новая секция
 * @param oNewSectPr {CSectionPr} - Если не задано, тогда секция удаляется
 * @param isCheckHdrFtr {boolean} - Проверять ли колонтитулы при удалении секции
 */
CDocument.prototype.UpdateSectionInfo = function(oSectPr, oNewSectPr, isCheckHdrFtr)
{
	if (!this.SectionsInfo.UpdateSection(oSectPr, oNewSectPr, isCheckHdrFtr))
		this.UpdateAllSectionsInfo();
};
CDocument.prototype.Check_SectionLastParagraph = function()
{
	var Count = this.Content.length;
	if (Count <= 0)
		return;

	var Element = this.Content[Count - 1];
	if (type_Paragraph === Element.GetType() && undefined !== Element.Get_SectionPr())
		this.Internal_Content_Add(Count, new Paragraph(this.DrawingDocument, this), false);
};
CDocument.prototype.Add_SectionBreak = function(SectionBreakType)
{
	if (docpostype_Content !== this.GetDocPosType())
		return false;

	if (true === this.Selection.Use)
	{
		// Если у нас есть селект, тогда ставим курсор в начало селекта
		this.MoveCursorLeft(false, false);
	}

	var nContentPos = this.CurPos.ContentPos;
	var oElement    = this.Content[nContentPos];
	var oCurSectPr  = this.SectionsInfo.Get_SectPr(nContentPos).SectPr;
	if (oElement.IsParagraph())
	{
		// Если мы стоим в параграфе, тогда делим данный параграф на 2 в текущей точке(даже если мы стоим в начале
		// или в конце параграфа) и к первому параграфу приписываем конец секции

		var oNewParagraph = oElement.Split();
		oNewParagraph.MoveCursorToStartPos(false);

		this.AddToContent(nContentPos + 1, oNewParagraph);
		this.CurPos.ContentPos = nContentPos + 1;

		// Заметим, что после функции Split, у параграфа Element не может быть окончания секции, т.к. если она
		// была в нем изначально, тогда после функции Split, окончание секции перенеслось в новый параграф.
	}
	else if (oElement.IsTable())
	{
		// Если мы стоим в таблице, тогда делим данную таблицу на 2 по текущему ряду(текущий ряд попадает во
		// вторую таблицу). Вставляем между таблицами параграф, и к этому параграфу приписываем окончание
		// секции. Если мы стоим в первой строке таблицы, таблицу делить не надо, достаточно добавить новый
		// параграф перед ней.

		var oNewParagraph = new Paragraph(this.DrawingDocument, this);
		var oNewTable     = oElement.Split();

		if (null === oNewTable)
		{
			this.AddToContent(nContentPos, oNewParagraph);
			this.CurPos.ContentPos = nContentPos + 1;
		}
		else
		{
			this.AddToContent(nContentPos + 1, oNewParagraph);
			this.AddToContent(nContentPos + 2, oNewTable);
			this.CurPos.ContentPos = nContentPos + 2;
		}

		this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);

		oElement = oNewParagraph;
	}
	else
	{
		return false;
	}

	var oSectPr = new CSectionPr(this);

	// В данном месте мы ставим разрыв секции. Чтобы до текущего места ничего не изменилось, мы у новой
	// для новой секции копируем все настройки из старой, а в старую секцию выставляем приходящий тип
	// разрыва секций. Заметим, что поскольку мы делаем все так, чтобы до текущей страницы ничего не
	// изменилось, надо сохранить эту информацию для пересчета, для этого мы помечаем все следующие изменения
	// как не влияющие на пересчет.

	this.History.MinorChanges = true;

	oSectPr.Copy(oCurSectPr);
	oCurSectPr.Set_Type(SectionBreakType);
	oCurSectPr.Set_PageNum_Start(-1);
	oCurSectPr.Clear_AllHdrFtr();

	this.History.MinorChanges = false;

	oElement.Set_SectionPr(oSectPr);
	oElement.Refresh_RecalcData2(0, 0);

	this.Recalculate();
	this.UpdateInterface();
	this.UpdateSelection();

	return true;
};
CDocument.prototype.Get_SectionFirstPage = function(SectIndex, Page_abs)
{
	if (SectIndex <= 0)
		return 0;

	var StartIndex = this.SectionsInfo.Get_SectPr2(SectIndex - 1).Index;

	// Ищем номер страницы, на которой закончилась предыдущая секция
	var CurPage = Page_abs;
	for (; CurPage > 0; CurPage--)
	{
		if (this.Pages[CurPage].EndPos >= StartIndex && this.Pages[CurPage].Pos <= StartIndex)
			break;
	}

	return CurPage + 1;
};
CDocument.prototype.Get_SectionPageNumInfo = function(Page_abs)
{
	var PageNumInfo = this.Get_SectionPageNumInfo2(Page_abs);

	var FP = PageNumInfo.FirstPage;
	var CP = PageNumInfo.CurPage;

	// Первая страница учитывается, только если параграф, идущий сразу за разрывом секции, начинается с новой страницы
	var bCheckFP  = true;
	var SectIndex = PageNumInfo.SectIndex;
	if (SectIndex > 0)
	{
		var CurSectInfo  = this.SectionsInfo.Get_SectPr2(SectIndex);
		var PrevSectInfo = this.SectionsInfo.Get_SectPr2(SectIndex - 1);

		if (CurSectInfo !== PrevSectInfo && c_oAscSectionBreakType.Continuous === CurSectInfo.SectPr.Get_Type() && true === CurSectInfo.SectPr.Compare_PageSize(PrevSectInfo.SectPr))
		{
			var ElementIndex = PrevSectInfo.Index;
			if (ElementIndex < this.Content.length - 1 && true !== this.Content[ElementIndex + 1].IsStartFromNewPage())
				bCheckFP = false;
		}
	}


	var bFirst = ( FP === CP && true === bCheckFP ? true : false );
	var bEven  = ( 0 === CP % 2 ? true : false ); // Четность/нечетность проверяется по текущему номеру страницы в секции, с учетом нумерации в секциях

	return new CSectionPageNumInfo(FP, CP, bFirst, bEven, Page_abs);
};
CDocument.prototype.Get_SectionPageNumInfo2 = function(Page_abs)
{
	var StartIndex = 0;

	// Такое может случится при первом рассчете документа, и когда мы находимся в автофигуре
	if (undefined !== this.Pages[Page_abs])
		StartIndex = this.Pages[Page_abs].Pos;

	var SectIndex      = this.SectionsInfo.Get_Index(StartIndex);
	var StartSectIndex = SectIndex;

	if (0 === SectIndex)
	{
		var PageNumStart = this.SectionsInfo.Get_SectPr2(0).SectPr.Get_PageNum_Start();
		var BT           = this.SectionsInfo.Get_SectPr2(0).SectPr.Get_Type();

		// Нумерация начинается с 1, если начальное значение не задано. Заметим, что в Word нумерация может начинаться и
		// со значения 0, а все отрицательные значения воспринимаются как продолжение нумерации с предыдущей секции.
		if (PageNumStart < 0)
			PageNumStart = 1;

		if ((c_oAscSectionBreakType.OddPage === BT && 0 === PageNumStart % 2) || (c_oAscSectionBreakType.EvenPage === BT && 1 === PageNumStart % 2))
			PageNumStart++;

		return {FirstPage : PageNumStart, CurPage : Page_abs + PageNumStart, SectIndex : StartSectIndex};
	}

	var SectionFirstPage = this.Get_SectionFirstPage(SectIndex, Page_abs);

	var FirstPage    = SectionFirstPage;
	var PageNumStart = this.SectionsInfo.Get_SectPr2(SectIndex).SectPr.Get_PageNum_Start();
	var BreakType    = this.SectionsInfo.Get_SectPr2(SectIndex).SectPr.Get_Type();

	var StartInfo = [];
	StartInfo.push({FirstPage : FirstPage, BreakType : BreakType});

	while ((PageNumStart < 0 || c_oAscSectionBreakType.Continuous === BreakType) && SectIndex > 0)
	{
		SectIndex--;

		FirstPage    = this.Get_SectionFirstPage(SectIndex, Page_abs);
		PageNumStart = this.SectionsInfo.Get_SectPr2(SectIndex).SectPr.Get_PageNum_Start();
		BreakType    = this.SectionsInfo.Get_SectPr2(SectIndex).SectPr.Get_Type();

		StartInfo.push({FirstPage : FirstPage, BreakType : BreakType});
	}

	// Нумерация начинается с 1, если начальное значение не задано. Заметим, что в Word нумерация может начинаться и
	// со значения 0, а все отрицательные значения воспринимаются как продолжение нумерации с предыдущей секции.
	if (PageNumStart < 0)
		PageNumStart = 1;

	var InfoCount = StartInfo.length;
	var InfoIndex = InfoCount - 1;

	var FP     = StartInfo[InfoIndex].FirstPage;
	var BT     = StartInfo[InfoIndex].BreakType;
	var PrevFP = StartInfo[InfoIndex].FirstPage;

	while (InfoIndex >= 0)
	{
		FP = StartInfo[InfoIndex].FirstPage;
		BT = StartInfo[InfoIndex].BreakType;

		PageNumStart += FP - PrevFP;
		PrevFP = FP;

		if ((c_oAscSectionBreakType.OddPage === BT && 0 === PageNumStart % 2) || (c_oAscSectionBreakType.EvenPage === BT && 1 === PageNumStart % 2))
			PageNumStart++;

		InfoIndex--;
	}

	if (FP > Page_abs)
		Page_abs = FP;

	var _FP = PageNumStart;
	var _CP = PageNumStart + Page_abs - FP;	// TODO: Здесь есть баг с рассчетом (чтобы его поправить добавил следующий комментарий,
											// но такой рассчет оказался неверным)
											// + 1 потому что FP начинает считать от 1, а Page_abs от 0

	return {FirstPage : _FP, CurPage : _CP, SectIndex : StartSectIndex};
};
CDocument.prototype.Get_SectionHdrFtr = function(nPageAbs, isFirst, isEven)
{
	return this.Layout.GetSectionHdrFtr(nPageAbs, isFirst, isEven);
};
CDocument.prototype.Create_SectionHdrFtr = function(Type, PageIndex)
{
	// Данная функция используется, когда у нас нет колонтитула вообще. Это значит, что его нет ни в 1 секции. Следовательно,
	// создаем колонтитул в первой секции, а в остальных он будет повторяться. По текущей секции нам надо будет
	// определить какой конкретно колонтитул мы собираемся создать.

	var SectionPageInfo = this.Get_SectionPageNumInfo(PageIndex);

	var _bFirst = SectionPageInfo.bFirst;
	var _bEven  = SectionPageInfo.bEven;

	var StartIndex = this.Pages[PageIndex].Pos;
	var SectIndex  = this.SectionsInfo.Get_Index(StartIndex);
	var CurSectPr  = this.SectionsInfo.Get_SectPr2(SectIndex).SectPr;

	var bEven  = ( true === _bEven && true === EvenAndOddHeaders ? true : false );
	var bFirst = ( true === _bFirst && true === CurSectPr.TitlePage ? true : false );

	var SectPr = this.SectionsInfo.Get_SectPr2(0).SectPr;
	var HdrFtr = new CHeaderFooter(this.HdrFtr, this, this.DrawingDocument, Type);

	if (hdrftr_Header === Type)
	{
		if (true === bFirst)
			SectPr.Set_Header_First(HdrFtr);
		else if (true === bEven)
			SectPr.Set_Header_Even(HdrFtr);
		else
			SectPr.Set_Header_Default(HdrFtr);
	}
	else
	{
		if (true === bFirst)
			SectPr.Set_Footer_First(HdrFtr);
		else if (true === bEven)
			SectPr.Set_Footer_Even(HdrFtr);
		else
			SectPr.Set_Footer_Default(HdrFtr);
	}

	return HdrFtr;
};
CDocument.prototype.On_SectionChange = function(_SectPr)
{
	var Index = this.SectionsInfo.Find(_SectPr);
	if (-1 === Index)
		return;

	var SectPr  = null;
	var HeaderF = null, HeaderD = null, HeaderE = null, FooterF = null, FooterD = null, FooterE = null;

	while (Index >= 0)
	{
		SectPr = this.SectionsInfo.Get_SectPr2(Index).SectPr;

		if (null === HeaderF)
			HeaderF = SectPr.Get_Header_First();

		if (null === HeaderD)
			HeaderD = SectPr.Get_Header_Default();

		if (null === HeaderE)
			HeaderE = SectPr.Get_Header_Even();

		if (null === FooterF)
			FooterF = SectPr.Get_Footer_First();

		if (null === FooterD)
			FooterD = SectPr.Get_Footer_Default();

		if (null === FooterE)
			FooterE = SectPr.Get_Footer_Even();

		Index--;
	}

	if (null !== HeaderF)
		HeaderF.Refresh_RecalcData_BySection(_SectPr);

	if (null !== HeaderD)
		HeaderD.Refresh_RecalcData_BySection(_SectPr);

	if (null !== HeaderE)
		HeaderE.Refresh_RecalcData_BySection(_SectPr);

	if (null !== FooterF)
		FooterF.Refresh_RecalcData_BySection(_SectPr);

	if (null !== FooterD)
		FooterD.Refresh_RecalcData_BySection(_SectPr);

	if (null !== FooterE)
		FooterE.Refresh_RecalcData_BySection(_SectPr);
};
CDocument.prototype.Create_HdrFtrWidthPageNum = function(PageIndex, AlignV, AlignH)
{
	// Определим четность страницы и является ли она первой в данной секции. Заметим, что четность страницы
	// отсчитывается от начала текущей секции и не зависит от настроек нумерации страниц для данной секции.
	var SectionPageInfo = this.Get_SectionPageNumInfo(PageIndex);

	var bFirst = SectionPageInfo.bFirst;
	var bEven  = SectionPageInfo.bEven;

	// Запросим нужный нам колонтитул
	var HdrFtr = this.Get_SectionHdrFtr(PageIndex, bFirst, bEven);

	switch (AlignV)
	{
		case hdrftr_Header :
		{
			var Header = HdrFtr.Header;

			if (null === Header)
				Header = this.Create_SectionHdrFtr(hdrftr_Header, PageIndex);

			Header.AddPageNum(AlignH);

			break;
		}

		case hdrftr_Footer :
		{
			var Footer = HdrFtr.Footer;

			if (null === Footer)
				Footer = this.Create_SectionHdrFtr(hdrftr_Footer, PageIndex);

			Footer.AddPageNum(AlignH);

			break;
		}
	}

	this.Recalculate();
};
CDocument.prototype.GetCurrentSectionPr = function()
{
	var oSectPr = this.Controller.GetCurrentSectionPr();
	if (null === oSectPr)
		return this.controller_GetCurrentSectionPr();

	return oSectPr;
};
CDocument.prototype.GetFirstElementInSection = function(SectionIndex)
{
	if (SectionIndex <= 0)
		return this.Content[0] ? this.Content[0] : null;

	var nElementPos = this.SectionsInfo.Get_SectPr2(SectionIndex - 1).Index + 1;
	return this.Content[nElementPos] ? this.Content[nElementPos] : null;
};
CDocument.prototype.GetSectionIndexByElementIndex = function(ElementIndex)
{
	return this.SectionsInfo.Get_Index(ElementIndex);
};
CDocument.prototype.GetSectionsCount = function()
{
	return this.SectionsInfo.GetSectionsCount();
};
CDocument.prototype.GetSectionsInfo = function()
{
	return this.SectionsInfo;
};
CDocument.prototype.GetLastSection = function()
{
	return this.SectPr;
};
/**
 * Определяем использовать ли заливку текста в особых случаях, когда вызывается заливка параграфа
 * @param isUse {boolean}
 */
CDocument.prototype.SetUseTextShd = function(isUse)
{
	this.UseTextShd = isUse;
};
CDocument.prototype.RecalculateFromStart = function(bUpdateStates)
{
	var RecalculateData = {
		Inline   : {Pos : 0, PageNum : 0},
		Flow     : [],
		HdrFtr   : [],
		Drawings : {All : true, Map : {}},
		Tables   : []
	};

	this.Reset_RecalculateCache();
	this.RecalculateWithParams(RecalculateData, true);

	if (true === bUpdateStates)
	{
		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
	}
};
CDocument.prototype.Register_Field = function(oField)
{
	this.FieldsManager.Register_Field(oField);
};
CDocument.prototype.Is_MailMergePreviewResult = function()
{
	return this.MailMergePreview;
};
CDocument.prototype.Is_HighlightMailMergeFields = function()
{
	return this.MailMergeFieldsHighlight;
};
CDocument.prototype.CompareDrawingsLogicPositions = function(Drawing1, Drawing2)
{
	var ParentPara1 = Drawing1.GetParagraph();
	var ParentPara2 = Drawing2.GetParagraph();

	if (!ParentPara1 || !ParentPara2 || !ParentPara1.Parent || !ParentPara2.Parent)
		return 0;

	var TopDocument1 = ParentPara1.Parent.Is_TopDocument(true);
	var TopDocument2 = ParentPara2.Parent.Is_TopDocument(true);

	if (TopDocument1 !== TopDocument2 || !TopDocument1)
		return 0;

	var TopElement1 = ParentPara1.GetTopElement();
	var TopElement2 = ParentPara2.GetTopElement();

	if (!TopElement1 || !TopElement2)
		return 0;

	var TopIndex1 = TopElement1.Get_Index();
	var TopIndex2 = TopElement2.Get_Index();

	if (TopIndex1 < TopIndex2)
		return 1;
	else if (TopIndex1 > TopIndex2)
		return -1;

	if (undefined !== TopDocument1.Content[TopIndex1])
	{
		var CompareEngine = new CDocumentCompareDrawingsLogicPositions(Drawing1, Drawing2);
		TopDocument1.Content[TopIndex1].CompareDrawingsLogicPositions(CompareEngine);
		return CompareEngine.Result;
	}

	return 0;
};
CDocument.prototype.GetTopElement = function()
{
	return null;
};
CDocument.prototype.private_StopSelection = function()
{
	this.Selection.Start = false;
};
CDocument.prototype.private_UpdateCurPage = function()
{
	if (true === this.TurnOffRecalcCurPos)
		return;

	this.private_CheckCurPage();
};
CDocument.prototype.private_UpdateCursorXY = function(bUpdateX, bUpdateY, isUpdateTarget)
{
	if (undefined === isUpdateTarget)
		isUpdateTarget = true;

	this.private_UpdateCurPage();

	var NewCursorPos = null;

	if (true !== this.IsSelectionUse() || true === this.IsSelectionEmpty())
	{
		this.DrawingDocument.UpdateTargetTransform(null);
		NewCursorPos = this.Controller.RecalculateCurPos(bUpdateX, bUpdateY, isUpdateTarget);
		if (NewCursorPos && NewCursorPos.Transform)
		{
			var x = NewCursorPos.Transform.TransformPointX(NewCursorPos.X, NewCursorPos.Y);
			var y = NewCursorPos.Transform.TransformPointY(NewCursorPos.X, NewCursorPos.Y);

			NewCursorPos.X = x;
			NewCursorPos.Y = y;
		}
	}
	else
	{
		// Обновляем всегда по концу селекта
		var SelectionBounds = this.GetSelectionBounds();
		if (null !== SelectionBounds)
		{
			NewCursorPos = {};
			if (-1 === SelectionBounds.Direction)
			{
				NewCursorPos.X = SelectionBounds.Start.X;
				NewCursorPos.Y = SelectionBounds.Start.Y;

				if (this.CurPage < SelectionBounds.Start.Page)
					NewCursorPos.Y = this.Pages[this.CurPage].Height;
			}
			else
			{
				NewCursorPos.X = SelectionBounds.End.X + SelectionBounds.End.W;
				NewCursorPos.Y = SelectionBounds.End.Y + SelectionBounds.End.H;

				if (this.CurPage > SelectionBounds.End.Page)
					NewCursorPos.Y = 0;
			}
		}
	}

	if (null === NewCursorPos || undefined === NewCursorPos)
		return;

	if (bUpdateX)
		this.CurPos.RealX = NewCursorPos.X;

	if (bUpdateY)
		this.CurPos.RealY = NewCursorPos.Y;

	if (true === this.Selection.Use && true !== this.Selection.Start)
		this.private_OnSelectionEnd();

	this.private_CheckCursorInPlaceHolder();
};
CDocument.prototype.private_MoveCursorDown = function(StartX, StartY, AddToSelect)
{
	var StartPage = this.CurPage;
	var CurY      = StartY;

	// Если данная страница еще не успела пересчитаться, тогда не даем смещаться
	if (StartPage >= this.Pages.length)
		return true;

	var PageH = this.Pages[this.CurPage].Height;

	this.TurnOff_InterfaceEvents();
	this.CheckEmptyElementsOnSelection = false;

	var Result = false;
	while (true)
	{
		CurY += 0.1;

		if (CurY > PageH || this.CurPage > StartPage)
		{
			this.CurPage = StartPage;
			if (this.Pages.length - 1 <= this.CurPage)
			{
				Result = false;
				break;
			}
			else
			{
				this.CurPage++;
				StartY = 0;

				var NewPage = this.CurPage;
				var bBreak = false;
				while (true)
				{
					this.MoveCursorToXY(StartX, StartY, AddToSelect);
					this.private_UpdateCursorXY(false, true, false);

					if (this.CurPage < NewPage)
					{
						StartY += 0.1;
						if (StartY > this.Pages[NewPage].Height)
						{
							NewPage++;
							if (this.Pages.length - 1 < NewPage)
							{
								Result = false;
								break;
							}

							StartY = 0;
						}
						else
						{
							StartY += 0.1;
						}
						this.CurPage = NewPage;
					}
					else
					{
						Result = true;
						bBreak = true;
						break;
					}
				}

				if (false === Result || true === bBreak)
					break;

				CurY = StartY;
			}
		}



		this.MoveCursorToXY(StartX, CurY, AddToSelect);
		this.private_UpdateCursorXY(false, true, false);

		if (this.CurPos.RealY > StartY + 0.001)
		{
			Result = true;
			break;
		}
	}

	this.CheckEmptyElementsOnSelection = true;
	this.TurnOn_InterfaceEvents(true);

	this.private_UpdateCursorXY(false, true, true);

	return Result;
};
CDocument.prototype.private_MoveCursorUp = function(StartX, StartY, AddToSelect)
{
	var StartPage = this.CurPage;
	var CurY      = StartY;

	// Если данная страница еще не успела пересчитаться, тогда не даем смещаться
	if (StartPage >= this.Pages.length)
		return true;

	var PageH = this.Pages[this.CurPage].Height;

	this.TurnOff_InterfaceEvents();
	this.CheckEmptyElementsOnSelection = false;

	var Result = false;
	while (true)
	{
		CurY -= 0.1;

		if (CurY < 0 || this.CurPage < StartPage)
		{
			this.CurPage = StartPage;
			if (0 === this.CurPage)
			{
				Result = false;
				break;
			}
			else
			{
				this.CurPage--;
				StartY = this.Pages[this.CurPage].Height;

				var NewPage = this.CurPage;
				var bBreak  = false;
				while (true)
				{
					this.MoveCursorToXY(StartX, StartY, AddToSelect);
					this.private_UpdateCursorXY(false, true, false);

					if (this.CurPage > NewPage)
					{
						this.CurPage = NewPage;
						StartY -= 0.1;

						if (StartY < 0)
						{
							// Защита от полностью пустой страницы
							if (0 === this.CurPage)
							{
								Result = false;
								bBreak = true;
								break;
							}
							else
							{
								this.CurPage--;
								StartY = this.Pages[this.CurPage].Height;

								NewPage = this.CurPage;
							}
						}
					}
					else
					{
						Result = true;
						bBreak = true;
						break;
					}
				}

				if (false === Result || true === bBreak)
					break;

				CurY = StartY;
			}
		}

		this.MoveCursorToXY(StartX, CurY, AddToSelect);
		this.private_UpdateCursorXY(false, true, false);

		if (this.CurPos.RealY < StartY - 0.001)
		{
			Result = true;
			break;
		}
	}

	this.CheckEmptyElementsOnSelection = true;
	this.TurnOn_InterfaceEvents(true);
	this.private_UpdateCursorXY(false, true, true);
	return Result;
};
CDocument.prototype.MoveCursorPageDown = function(AddToSelect, NextPage)
{
	if (this.Pages.length <= 0)
		return;

	if (true === this.IsSelectionUse() && true !== AddToSelect)
		this.MoveCursorRight(false, false);

	var bStopSelection = false;
	if (true !== this.IsSelectionUse() && true === AddToSelect && true !== NextPage)
	{
		bStopSelection = true;
		this.StartSelectionFromCurPos();
	}

	this.private_UpdateCursorXY(false, true);
	var Result = this.private_MoveCursorPageDown(this.CurPos.RealX, this.CurPos.RealY, AddToSelect, NextPage);

	if (true === AddToSelect && true !== NextPage && true !== Result)
		this.MoveCursorToEndPos(true);

	if (bStopSelection)
		this.private_StopSelection();
};
CDocument.prototype.private_MoveCursorPageDown = function(StartX, StartY, AddToSelect, NextPage)
{
	if (this.Pages.length <= 0)
		return true;

	var Dy        = 0;
	var StartPage = this.CurPage;
	if (NextPage)
	{
		if (this.CurPage >= this.Pages.length - 1)
		{
			this.MoveCursorToXY(0, 0, false);
			this.private_UpdateCursorXY(false, true);
			return;
		}

		this.CurPage++;
		StartX = 0;
		StartY = 0;
	}
	else
	{
		if (this.CurPage >= this.Pages.length)
		{
			this.CurPage    = this.Pages.length - 1;
			var LastElement = this.Content[this.Pages[this.CurPage].EndPos];
			Dy              = LastElement.GetPageBounds(LastElement.GetPagesCount() - 1).Bottom;
			StartPage       = this.CurPage;
		}
		else
		{
			Dy = this.DrawingDocument.GetVisibleMMHeight();
			if (StartY + Dy > this.Get_PageLimits(this.CurPage).YLimit)
			{
				this.CurPage++;
				var PageH = this.Get_PageLimits(this.CurPage).YLimit;
				Dy -= PageH - StartY;
				StartY    = 0;
				while (Dy > PageH)
				{
					Dy -= PageH;
					this.CurPage++;
				}

				if (this.CurPage >= this.Pages.length)
				{
					this.CurPage    = this.Pages.length - 1;
					var LastElement = this.Content[this.Pages[this.CurPage].EndPos];
					Dy              = LastElement.GetPageBounds(LastElement.GetPagesCount() - 1).Bottom;
				}

				StartPage = this.CurPage;
			}
		}
	}

	this.MoveCursorToXY(StartX, StartY + Dy, AddToSelect);
	this.private_UpdateCursorXY(false, true);

	if (this.CurPage > StartPage || (this.CurPos.RealY > StartY + 0.001 && this.CurPage === StartPage))
		return true;

	return this.private_MoveCursorDown(this.CurPos.RealX, this.CurPos.RealY, AddToSelect);
};
CDocument.prototype.MoveCursorPageUp = function(AddToSelect, PrevPage)
{
	if (this.Pages.length <= 0)
		return;

	if (true === this.IsSelectionUse() && true !== AddToSelect)
		this.MoveCursorRight(false, false);

	var bStopSelection = false;
	if (true !== this.IsSelectionUse() && true === AddToSelect && true !== PrevPage)
	{
		bStopSelection = true;
		this.StartSelectionFromCurPos();
	}

	this.private_UpdateCursorXY(false, true);
	var Result = this.private_MoveCursorPageUp(this.CurPos.RealX, this.CurPos.RealY, AddToSelect, PrevPage);

	if (true === AddToSelect && true !== PrevPage && true !== Result)
		this.MoveCursorToEndPos(true);

	if (bStopSelection)
		this.private_StopSelection();
};
CDocument.prototype.private_MoveCursorPageUp = function(StartX, StartY, AddToSelect, PrevPage)
{
	if (this.Pages.length <= 0)
		return true;

	if (this.CurPage >= this.Pages.length)
	{
		this.CurPage    = this.Pages.length - 1;
		var LastElement = this.Content[this.Pages[this.CurPage].EndPos];
		StartY          = LastElement.GetPageBounds(LastElement.GetPagesCount() - 1).Bottom;
	}

	var Dy        = 0;
	var StartPage = this.CurPage;
	if (PrevPage)
	{
		if (this.CurPage <= 0)
		{
			this.MoveCursorToXY(0, 0, false);
			this.private_UpdateCursorXY(false, true);
			return;
		}

		this.CurPage--;
		StartX = 0;
		StartY = 0;
	}
	else
	{
		Dy = this.DrawingDocument.GetVisibleMMHeight();
		if (StartY - Dy < 0)
		{
			this.CurPage--;
			var PageH = this.Get_PageLimits(this.CurPage).YLimit;

			Dy -= StartY;
			StartY = PageH;
			while (Dy > PageH)
			{
				Dy -= PageH;
				this.CurPage--;
			}

			if (this.CurPage < 0)
			{
				this.CurPage = 0;
				var oElement = this.Content[0];
				Dy           = PageH - oElement.GetPageBounds(oElement.GetPagesCount() - 1).Top;
			}

			StartPage = this.CurPage;
		}
	}

	this.MoveCursorToXY(StartX, StartY - Dy, AddToSelect);
	this.private_UpdateCursorXY(false, true);

	if (this.CurPage < StartPage || (this.CurPos.RealY < StartY - 0.001 && this.CurPage === StartPage))
		return true;

	return this.private_MoveCursorUp(this.CurPos.RealX, this.CurPos.RealY, AddToSelect);
};
CDocument.prototype.private_ProcessTemplateReplacement = function(TemplateReplacementData)
{
	let oProps = new AscCommon.CSearchSettings();
	oProps.SetMatchCase(true);

	for (var Id in TemplateReplacementData)
	{
		oProps.SetText(Id);
		this.Search(oProps, false);
		this.SearchEngine.ReplaceAll(TemplateReplacementData[Id], false);
	}
};
CDocument.prototype.private_CheckCursorInPlaceHolder = function()
{
	let placeholderOwner = this.GetPlaceHolderObject();
	if (placeholderOwner)
	{
		if (placeholderOwner instanceof CInlineLevelSdt || placeholderOwner instanceof CBlockLevelSdt)
		{
			placeholderOwner.SelectContentControl();
		}
		else if (placeholderOwner instanceof ParaMath)
		{
			placeholderOwner.SelectAllInCurrentMathContent();
		}
	}
};

CDocument.prototype.ResetWordSelection = function()
{
	this.Selection.WordSelected = false;
};
CDocument.prototype.SetWordSelection = function(isWord)
{
	this.Selection.WordSelected = isWord;
};
CDocument.prototype.IsWordSelection = function()
{
	return this.Selection.WordSelected;
};
CDocument.prototype.IsStartSelection = function()
{
	return this.Selection.Start;
};
CDocument.prototype.Get_EditingType = function()
{
	return this.EditingType;
};
CDocument.prototype.Set_EditingType = function(EditingType)
{
	this.EditingType = EditingType;
};
CDocument.prototype.GetTrackRevisionsManager = function()
{
	return this.TrackRevisionsManager;
};
CDocument.prototype.Start_SilentMode = function()
{
	this.TurnOff_Recalculate();
	this.TurnOff_InterfaceEvents();
	this.TurnOff_RecalculateCurPos();
};
CDocument.prototype.End_SilentMode = function(bUpdate)
{
	this.TurnOn_Recalculate(bUpdate);
	this.TurnOn_RecalculateCurPos(bUpdate);
	this.TurnOn_InterfaceEvents(bUpdate);
};
/**
 * Стартуем режим, в котором новые классы не регистрируются в истории, а также отключается рецензирование
 * Возвращается текущее состояние редактора
 * @returns {object}
 */
CDocument.prototype.StartNoHistoryMode = function()
{
	var oState = {};

	oState.LocalTrackRevisions = this.GetLocalTrackRevisions();

	this.History.TurnOff();
	this.TableId.TurnOff();
	this.SetLocalTrackRevisions(false);

	return oState;
};
/**
 * Останавливаем режим "без истории"
 * @param oState
 */
CDocument.prototype.EndNoHistoryMode = function(oState)
{
	this.History.TurnOn();
	this.TableId.TurnOn();
	this.SetLocalTrackRevisions(oState.LocalTrackRevisions);
};
/**
 * Начинаем селект с текущей точки. Если селект уже есть, тогда ничего не делаем.
 */
CDocument.prototype.StartSelectionFromCurPos = function()
{
	if (true === this.IsSelectionUse())
		return true;

	this.Selection.Use   = true;
	this.Selection.Start = false;

	this.Controller.StartSelectionFromCurPos();
};
CDocument.prototype.Is_TrackingDrawingObjects = function()
{
	return this.DrawingObjects.Check_TrackObjects();
};
CDocument.prototype.Add_ChangedStyle = function(arrStylesId)
{
	for (var nIndex = 0, nCount = arrStylesId.length; nIndex < nCount; nIndex++)
	{
		this.ChangedStyles[arrStylesId[nIndex]] = true;
	}
};
CDocument.prototype.UpdateStylePanel = function()
{
	if (0 !== this.TurnOffPanelStyles)
		return;

	var bNeedUpdate = false;
	for (var StyleId in this.ChangedStyles)
	{
		bNeedUpdate = true;
		break;
	}

	this.ChangedStyles = {};

	if (true === bNeedUpdate)
	{
		this.Api.GenerateStyles();
	}
};
CDocument.prototype.LockPanelStyles = function()
{
	this.TurnOffPanelStyles++;
};
CDocument.prototype.UnlockPanelStyles = function(isUpdate)
{
	this.TurnOffPanelStyles = Math.max(0, this.TurnOffPanelStyles - 1);

	if (true === isUpdate)
		this.UpdateStylePanel();
};
CDocument.prototype.UpdateNumberingPanel = function()
{
	this.NumberingCollection.UpdatePanel();
};
CDocument.prototype.GetAllParagraphs = function(Props, ParaArray)
{
	if (Props && true === Props.OnlyMainDocument && true === Props.All && null !== this.AllParagraphsList)
		return this.AllParagraphsList;

	if (!ParaArray)
		ParaArray = [];

	if (!Props)
		Props = {All : true};

	if (true === Props.OnlyMainDocument)
	{
		var Count = this.Content.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var Element = this.Content[Index];
			Element.GetAllParagraphs(Props, ParaArray);
		}
	}
	else
	{
		this.SectionsInfo.GetAllParagraphs(Props, ParaArray);

		var Count = this.Content.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var Element = this.Content[Index];
			Element.GetAllParagraphs(Props, ParaArray);
		}

		this.Footnotes.GetAllParagraphs(Props, ParaArray);
		this.Endnotes.GetAllParagraphs(Props, ParaArray);
	}

	if (Props && true === Props.OnlyMainDocument && true === Props.All)
		this.AllParagraphsList = ParaArray;

	return ParaArray;
};
CDocument.prototype.GetAllTables = function(oProps, arrTables)
{
	if (!arrTables)
		arrTables = [];

	if (!oProps)
		oProps = {};

	if (oProps && true === oProps.OnlyMainDocument)
	{
		for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			this.Content[nIndex].GetAllTables(oProps, arrTables);
		}
	}
	else
	{
		this.SectionsInfo.GetAllTables(oProps, arrTables);

		for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			this.Content[nIndex].GetAllTables(oProps, arrTables);
		}

		this.Footnotes.GetAllTables(oProps, arrTables);
		this.Endnotes.GetAllTables(oProps, arrTables);
	}

	return arrTables;
};
CDocument.prototype.GetAllNumberedParagraphs = function()
{
	var oProps = {};
	oProps.OnlyMainDocument = true;
	oProps.Shapes = false;
	oProps.DoNotAddRemoved = true;
	oProps.Numbering = true;
	oProps.NumPr = [];
	var oNumbering = this.GetNumbering();
	var oNumMap = oNumbering.Num, nLvl;
	for(var sNumId in oNumMap)
	{
		var oNum = oNumbering.Num[sNumId];
		for(nLvl = 0; nLvl < 9; ++nLvl)
		{
			var oLvl = oNum.GetLvl(nLvl);
			if(oLvl && oLvl.IsNumbered())
			{
				oProps.NumPr.push(new CNumPr(oNum.GetId(), nLvl));
			}
		}
	}
	return this.GetAllParagraphs(oProps);
};
CDocument.prototype.GetFootNotesFirstParagraphs = function()
{
	return this.Footnotes.GetFirstParagraphs();
};
CDocument.prototype.GetEndNotesFirstParagraphs = function()
{
	return this.Endnotes.GetFirstParagraphs();
};
CDocument.prototype.GetAllCaptionParagraphs = function(sCaption)
{
	var aFields = [];
	if(!(typeof sCaption === "string" && sCaption.length > 0))
	{
		return [];
	}
	this.GetAllSeqFieldsByType(sCaption, aFields);
	var aParagraphs = [];
	for(var nField = 0; nField < aFields.length; ++nField)
	{
		var oField = aFields[nField];
		if(oField)
		{
            if(oField.BeginChar)
            {
                var oRun       = oField.BeginChar.GetRun();
                var oParagraph = oRun.GetParagraph();
                if(oParagraph)
                {
                    aParagraphs.push(oParagraph);
                }
            }
            else if(oField.Paragraph)
            {
                aParagraphs.push(oField.Paragraph);
            }
		}
	}
	return aParagraphs;
};
CDocument.prototype.GetAllUsedParagraphStyles = function ()
{
    var aParagraphs = this.GetAllParagraphs({All: true});
    var nIndex, nCount = aParagraphs.length;
    var oStylesMap = {}, sStyleId;
    var oStyles = this.GetStyles();
    var aStyles = [];
    var oStyle;
    for(nIndex = 0; nIndex < nCount; ++nIndex)
    {
        sStyleId = aParagraphs[nIndex].GetParagraphStyle();
        if(sStyleId)
        {
            oStylesMap[sStyleId] = true;
        }
        else
        {
            sStyleId = oStyles.GetDefaultParagraph();
            oStylesMap[sStyleId] = true;
        }
    }
    for(sStyleId in oStylesMap)
    {
        oStyle = oStyles.Get(sStyleId);
        if(oStyle)
        {
            aStyles.push(oStyles.Get(sStyleId));
        }
    }
    return aStyles;
};
CDocument.prototype.TurnOffHistory = function()
{
	this.History.TurnOff();
	this.TableId.TurnOff();
};
CDocument.prototype.TurnOnHistory = function()
{
	this.TableId.TurnOn();
	this.History.TurnOn();
};
CDocument.prototype.Get_SectPr = function(nContentPos)
{
	return this.Layout.GetSectionByPos(nContentPos);
};
CDocument.prototype.Add_ToContent = function(Pos, Item, isCorrectContent)
{
	this.Internal_Content_Add(Pos, Item, isCorrectContent);
};
CDocument.prototype.Remove_FromContent = function(Pos, Count, isCorrectContent)
{
	this.Internal_Content_Remove(Pos, Count, isCorrectContent);
};
CDocument.prototype.Set_FastCollaborativeEditing = function(isOn)
{
	this.CollaborativeEditing.Set_Fast(isOn);

	if (this.Api && c_oAscCollaborativeMarksShowType.LastChanges === this.Api.GetCollaborativeMarksShowType())
		this.Api.SetCollaborativeMarksShowType(c_oAscCollaborativeMarksShowType.All);
};
CDocument.prototype.Continue_FastCollaborativeEditing = function()
{
	if (true === this.CollaborativeEditing.Get_GlobalLock())
	{
		if (this.Api.forceSaveUndoRequest)
			this.Api.asc_Save(true);

		return;
	}

	if (this.Api.isLongAction())
		return;

	if (true !== this.CollaborativeEditing.Is_Fast() || true === this.CollaborativeEditing.Is_SingleUser())
		return;

	if (true === this.IsMovingTableBorder() || true === this.Api.isStartAddShape || this.DrawingObjects.isTrackingDrawings() || this.Api.isOpenedChartFrame)
		return;

	var HaveChanges = this.History.Have_Changes(true);
	if (true !== HaveChanges && (true === this.CollaborativeEditing.Have_OtherChanges() || 0 !== this.CollaborativeEditing.getOwnLocksLength()))
	{
		// Принимаем чужие изменения. Своих нет, но функцию отсылки надо вызвать, чтобы снять локи.
		this.CollaborativeEditing.Apply_Changes();
		this.CollaborativeEditing.Send_Changes();
	}
	else if (true === HaveChanges || true === this.CollaborativeEditing.Have_OtherChanges())
	{
		this.Api.asc_Save(true);
	}

	var CurTime = new Date().getTime();
	if (true === this.NeedUpdateTargetForCollaboration && (CurTime - this.LastUpdateTargetTime > 1000))
	{
		this.NeedUpdateTargetForCollaboration = false;
		if (true !== HaveChanges)
		{
			var CursorInfo = this.History.Get_DocumentPositionBinary();
			if (null !== CursorInfo)
			{
				this.Api.CoAuthoringApi.sendCursor(CursorInfo);
				this.LastUpdateTargetTime = CurTime;
			}
		}
		else
		{
			this.LastUpdateTargetTime = CurTime;
		}
	}
};
CDocument.prototype.Save_DocumentStateBeforeLoadChanges = function(isRemoveSelection, isStoreViewPosition)
{
	var State = {};

	State.CurPos =
	{
		X     : this.CurPos.X,
		Y     : this.CurPos.Y,
		RealX : this.CurPos.RealX,
		RealY : this.CurPos.RealY,
		Type  : this.CurPos.Type
	};

	State.Selection =
	{
		Start          : this.Selection.Start,
		Use            : this.Selection.Use,
		Flag           : this.Selection.Flag,
		UpdateOnRecalc : this.Selection.UpdateOnRecalc,
		DragDrop       : {
			Flag : this.Selection.DragDrop.Flag,
			Data : null === this.Selection.DragDrop.Data ? null : {
				X       : this.Selection.DragDrop.Data.X,
				Y       : this.Selection.DragDrop.Data.Y,
				PageNum : this.Selection.DragDrop.Data.PageNum
			}
		}
	};

	State.SingleCell = this.GetSelectedElementsInfo().GetSingleCell();
	State.Pos        = [];
	State.StartPos   = [];
	State.EndPos     = [];

	this.Controller.SaveDocumentStateBeforeLoadChanges(State);
	
	if (isStoreViewPosition)
		this.private_StoreViewPositions(State);

	// TODO: Разобраться зачем здесь делается RemoveSelection, по логике надо вынести за пределы данной функции
	if (false !== isRemoveSelection)
		this.RemoveSelection();

	this.CollaborativeEditing.WatchDocumentPositionsByState(State);

	return State;
};
CDocument.prototype.GetDocumentPositionByXY = function(pageIndex, x, y)
{
	let anchorPos = this.Get_NearestPos(pageIndex, x, y);

	if (anchorPos
		&& anchorPos.Paragraph
		&& anchorPos.Paragraph.IsUseInDocument()
		&& anchorPos.ContentPos)
	{
		let run = anchorPos.Paragraph.GetClassByPos(anchorPos.ContentPos);
		if (run && run instanceof AscWord.CRun)
		{
			let posInRun = anchorPos.ContentPos.Get(anchorPos.ContentPos.GetDepth());

			let docPos = run.GetDocumentPositionFromObject();
			docPos.push({Class : run, Position : posInRun});
			return docPos;
		}
	}

	return null;
};
CDocument.prototype.private_GetXYByDocumentPosition = function(docPos)
{
	let run = docPos[docPos.length - 1].Class;
	if (!run || !(run instanceof AscWord.CRun))
		return null;

	let paragraph = run.GetParagraph();

	let state = paragraph.SaveSelectionState();
	paragraph.RemoveSelection();

	run.SetThisElementCurrentInParagraph();
	run.State.ContentPos = docPos[docPos.length - 1].Position;

	let posInfo = paragraph.RecalculateCurPos(false, false, false, true);
	paragraph.LoadSelectionState(state);

	return {
		Page : posInfo.PageNum,
		X    : 0,
		Y    : posInfo.Y,
		H    : posInfo.Height
	}
};
CDocument.prototype.private_StoreViewPositions = function(state)
{
	if (this.ViewPosition)
	{
		// Сюда попадаем, когда мы еще не успели обработать предыдущую расчитанную позицию, но при этом
		// прошло несколько пересчетов, которые могли поменять значения позиции, поэтому мы должны использовать
		// начально расчитанные значения
		state.AnchorAlignTop = this.ViewPosition.AlignTop;
		state.AnchorDistance = this.ViewPosition.Distance;
		state.AnchorType     = this.ViewPosition.Type;
		state.AnchorPos      = this.ViewPosition.AnchorPos;
		return;
	}
	
	let viewPort = this.DrawingDocument.GetVisibleRegion();
	if (!viewPort)
		return;
	
	let selectionBounds = this.GetSelectionBounds();
	
	let cursorY, cursorH, cursorPage = -1;
	let anchorType = AscWord.ViewPositionType.Common;
	if (selectionBounds)
	{
		if (this.IsSelectionUse())
		{
			if (selectionBounds.Direction > 0)
			{
				cursorY    = selectionBounds.End.Y;
				cursorH    = selectionBounds.End.H;
				cursorPage = selectionBounds.End.Page;
				anchorType = AscWord.ViewPositionType.SelectionEnd;
			}
			else
			{
				cursorY    = selectionBounds.Start.Y;
				cursorH    = selectionBounds.Start.H;
				cursorPage = selectionBounds.Start.Page;
				anchorType = AscWord.ViewPositionType.SelectionEnd;
			}
		}
		else
		{
			cursorY    = selectionBounds.Start.Y;
			cursorH    = selectionBounds.End.Y - selectionBounds.Start.Y;
			cursorPage = selectionBounds.Start.Page;
			anchorType = AscWord.ViewPositionType.Cursor;
		}
	}
	
	// TODO: Решить проблему, когда видно больше 2 страниц и курсор находится на средней странице
	if (-1 !== cursorPage
		&& ((viewPort[0].Page === cursorPage && cursorY + cursorH > viewPort[0].Y)
			|| (viewPort[1].Page === cursorPage && cursorY < viewPort[1].Y)))
	{
		let distance = 0;
		let alignTop = true;
		
		if (viewPort[0].Page === cursorPage)
		{
			distance = cursorY - viewPort[0].Y;
			alignTop = true;
		}
		else
		{
			distance = viewPort[1].Y - cursorY;
			alignTop = false;
		}
		
		state.AnchorAlignTop = alignTop;
		state.AnchorDistance = distance
		state.AnchorType     = anchorType;
		state.AnchorPos      = null;
	}
	else
	{
		let anchorPos = this.GetDocumentPositionByXY(viewPort[0].Page, 0, viewPort[0].Y);
		if (!anchorPos)
			return;
		
		let xyInfo = this.private_GetXYByDocumentPosition(anchorPos);
		
		let _anchorPos = anchorPos;
		if (viewPort[0].Page === viewPort[1].Page)
		{
			let pageIndex = viewPort[0].Page;
			
			let y0 = viewPort[0].Y;
			let y1 = viewPort[1].Y;
			let y  = y0;
			
			let _xyInfo = this.private_GetXYByDocumentPosition(_anchorPos);
			while (_xyInfo.Y < y0 && y < y1)
			{
				y += 10;
				_anchorPos = this.GetDocumentPositionByXY(pageIndex, 0, y);
				if (!_anchorPos)
					continue;
				
				_xyInfo = this.private_GetXYByDocumentPosition(_anchorPos);
			}
		}
		else
		{
			let pageIndex = viewPort[0].Page;
			
			let y0 = viewPort[0].Y;
			let y1 = this.Pages[pageIndex] ? this.Pages[pageIndex].Height : 297;
			let y  = y0;
			
			
			let _xyInfo = this.private_GetXYByDocumentPosition(_anchorPos);
			while (_xyInfo.Y < y0 && y < y1)
			{
				y += 10;
				_anchorPos = this.GetDocumentPositionByXY(pageIndex, 0, y);
				if (!_anchorPos)
					continue;
				
				_xyInfo = this.private_GetXYByDocumentPosition(_anchorPos);
			}
			
			if (y >= y1)
				_anchorPos = this.GetDocumentPositionByXY(pageIndex + 1, 0, 0);
		}
		
		if (_anchorPos)
		{
			anchorPos = _anchorPos;
			xyInfo = this.private_GetXYByDocumentPosition(anchorPos);
		}
		
		// TODO: Надо проверить, что совпали страница viewPort[0].Page и xyInfo.Page
		
		state.AnchorType     = AscWord.ViewPositionType.Common;
		state.AnchorPos      = anchorPos;
		state.AnchorAlignTop = true;
		state.AnchorDistance = xyInfo.Y - viewPort[0].Y;
	}
};
CDocument.prototype.Load_DocumentStateAfterLoadChanges = function(State)
{
	this.CollaborativeEditing.UpdateDocumentPositionsByState(State);

	this.RemoveSelection();

	this.CurPos.X     = State.CurPos.X;
	this.CurPos.Y     = State.CurPos.Y;
	this.CurPos.RealX = State.CurPos.RealX;
	this.CurPos.RealY = State.CurPos.RealY;
	this.SetDocPosType(State.CurPos.Type);

	this.Selection.Start          = State.Selection.Start;
	this.Selection.Use            = State.Selection.Use;
	this.Selection.Flag           = State.Selection.Flag;
	this.Selection.UpdateOnRecalc = State.Selection.UpdateOnRecalc;
	this.Selection.DragDrop.Flag  = State.Selection.DragDrop.Flag;
	this.Selection.DragDrop.Data  = State.Selection.DragDrop.Data === null ? null : {
		X       : State.Selection.DragDrop.Data.X,
		Y       : State.Selection.DragDrop.Data.Y,
		PageNum : State.Selection.DragDrop.Data.PageNum
	};

	this.Controller.RestoreDocumentStateAfterLoadChanges(State);

	if (true === this.Selection.Use && null !== State.SingleCell && undefined !== State.SingleCell)
	{
		var Cell  = State.SingleCell;
		var Table = Cell.Get_Table();
		if (Table && true === Table.IsUseInDocument())
		{
			Table.Set_CurCell(Cell);
			Table.RemoveSelection();
			Table.SelectTable(c_oAscTableSelectionType.Cell);
		}
	}

	if (undefined !== State.AnchorType)
	{
		this.ViewPosition = {
			AnchorPos : State.AnchorPos,
			AlignTop  : State.AnchorAlignTop,
			Distance  : State.AnchorDistance,
			Type      : State.AnchorType
		};
		
		if (AscWord.ViewPositionType.Cursor === this.ViewPosition.Type)
			this.ViewPosition.AnchorPos = State.Pos;
		else if (AscWord.ViewPositionType.SelectionStart === this.ViewPosition.Type)
			this.ViewPosition.AnchorPos = State.StartPos;
		else if (AscWord.ViewPositionType.SelectionEnd === this.ViewPosition.Type)
			this.ViewPosition.AnchorPos = State.EndPos;
	}
	else
	{
		this.ViewPosition = null;
	}

	this.UpdateSelection();

};
CDocument.prototype.SaveDocumentState = function(isRemoveSelection)
{
	return this.Save_DocumentStateBeforeLoadChanges(isRemoveSelection);
};
CDocument.prototype.LoadDocumentState = function(oState)
{
	return this.Load_DocumentStateAfterLoadChanges(oState);
};
CDocument.prototype.SetContentSelection = function(StartDocPos, EndDocPos, Depth, StartFlag, EndFlag)
{
	if ((0 === StartFlag && (!StartDocPos[Depth] || this !== StartDocPos[Depth].Class)) || (0 === EndFlag && (!EndDocPos[Depth] || this !== EndDocPos[Depth].Class)))
		return;

	var StartPos = 0, EndPos = 0;
	switch (StartFlag)
	{
		case 0 :
			StartPos = StartDocPos[Depth].Position;
			break;
		case 1 :
			StartPos = 0;
			break;
		case -1:
			StartPos = this.Content.length - 1;
			break;
	}

	switch (EndFlag)
	{
		case 0 :
			EndPos = EndDocPos[Depth].Position;
			break;
		case 1 :
			EndPos = 0;
			break;
		case -1:
			EndPos = this.Content.length - 1;
			break;
	}

	var _StartDocPos = StartDocPos, _StartFlag = StartFlag;
	if (null !== StartDocPos && true === StartDocPos[Depth].Deleted)
	{
		if (StartPos < this.Content.length)
		{
			_StartDocPos = null;
			_StartFlag   = 1;
		}
		else if (StartPos > 0)
		{
			StartPos--;
			_StartDocPos = null;
			_StartFlag   = -1;
		}
		else
		{
			// Такого не должно быть
			return;
		}
	}

	var _EndDocPos = EndDocPos, _EndFlag = EndFlag;
	if (null !== EndDocPos && true === EndDocPos[Depth].Deleted)
	{
		if (EndPos < this.Content.length)
		{
			_EndDocPos = null;
			_EndFlag   = 1;
		}
		else if (EndPos > 0)
		{
			EndPos--;
			_EndDocPos = null;
			_EndFlag   = -1;
		}
		else
		{
			// Такого не должно быть
			return;
		}
	}

	this.Selection.StartPos = StartPos;
	this.Selection.EndPos   = EndPos;

	if (StartPos !== EndPos)
	{
		this.Content[StartPos].SetContentSelection(_StartDocPos, null, Depth + 1, _StartFlag, StartPos > EndPos ? 1 : -1);
		this.Content[EndPos].SetContentSelection(null, _EndDocPos, Depth + 1, StartPos > EndPos ? -1 : 1, _EndFlag);

		var _StartPos = StartPos;
		var _EndPos   = EndPos;
		var Direction = 1;

		if (_StartPos > _EndPos)
		{
			_StartPos = EndPos;
			_EndPos   = StartPos;
			Direction = -1;
		}

		for (var CurPos = _StartPos + 1; CurPos < _EndPos; CurPos++)
		{
			this.Content[CurPos].SelectAll(Direction);
		}
	}
	else
	{
		this.Content[StartPos].SetContentSelection(_StartDocPos, _EndDocPos, Depth + 1, _StartFlag, _EndFlag);
	}
};
CDocument.prototype.GetContentPosition = function(bSelection, bStart, PosArray)
{
	if (undefined === PosArray)
		PosArray = [];

	var oSelection = this.private_GetSelectionPos();

	var Pos = (true === bSelection ? (true === bStart ? oSelection.Start : oSelection.End) : this.CurPos.ContentPos);
	PosArray.push({Class : this, Position : Pos});

	if (undefined !== this.Content[Pos] && this.Content[Pos].GetContentPosition)
		this.Content[Pos].GetContentPosition(bSelection, bStart, PosArray);

	return PosArray;
};
CDocument.prototype.SetContentPosition = function(DocPos, Depth, Flag)
{
	if (0 === Flag && (!DocPos[Depth] || this !== DocPos[Depth].Class))
		return;

	var Pos = 0;
	switch (Flag)
	{
		case 0 :
			Pos = DocPos[Depth].Position;
			break;
		case 1 :
			Pos = 0;
			break;
		case -1:
			Pos = this.Content.length - 1;
			break;
	}

	var _DocPos = DocPos, _Flag = Flag;
	if (null !== DocPos && true === DocPos[Depth].Deleted)
	{
		if (Pos < this.Content.length)
		{
			_DocPos = null;
			_Flag   = 1;
		}
		else if (Pos > 0)
		{
			Pos--;
			_DocPos = null;
			_Flag   = -1;
		}
		else
		{
			// Такого не должно быть
			return;
		}
	}

	this.CurPos.ContentPos = Pos;

	if (this.Content[Pos])
		this.Content[Pos].SetContentPosition(_DocPos, Depth + 1, _Flag);
};
CDocument.prototype.GetDocumentPositionFromObject = function(arrPos)
{
	if (!arrPos)
		arrPos = [];

	return arrPos;
};
CDocument.prototype.Get_CursorLogicPosition = function()
{
	var nDocPosType = this.GetDocPosType();
	if (docpostype_HdrFtr === nDocPosType)
	{
		var HdrFtr = this.HdrFtr.Get_CurHdrFtr();
		if (HdrFtr)
		{
			return this.private_GetLogicDocumentPosition(HdrFtr.Get_DocumentContent());
		}
	}
	else if (docpostype_Footnotes === nDocPosType)
	{
		var oFootnote = this.Footnotes.GetCurFootnote();
		if (oFootnote)
			return this.private_GetLogicDocumentPosition(oFootnote);
	}
	else if (docpostype_Endnotes === nDocPosType)
	{
		var oEndnote = this.Endnotes.GetCurEndnote();
		if (oEndnote)
			return this.private_GetLogicDocumentPosition(oEndnote);
	}
	else
	{
		return this.private_GetLogicDocumentPosition(this);
	}

	return null;
};
CDocument.prototype.private_GetLogicDocumentPosition = function(LogicDocument)
{
	if (!LogicDocument)
		return null;

	if (docpostype_DrawingObjects === LogicDocument.GetDocPosType())
	{
		var ParaDrawing    = this.DrawingObjects.getMajorParaDrawing();
		var DrawingContent = this.DrawingObjects.getTargetDocContent();
		if (!ParaDrawing)
			return null;

		if (DrawingContent)
		{
			return DrawingContent.GetContentPosition(DrawingContent.IsSelectionUse(), false);
		}
		else
		{
			var Run = ParaDrawing.Get_Run();
			if (null === Run)
				return null;

			var DrawingInRunPos = Run.Get_DrawingObjectSimplePos(ParaDrawing.Get_Id());
			if (-1 === DrawingInRunPos)
				return null;

			var DocPos = [{Class : Run, Position : DrawingInRunPos}];
			Run.GetDocumentPositionFromObject(DocPos);
			return DocPos;
		}
	}
	else
	{
		return LogicDocument.GetContentPosition(LogicDocument.IsSelectionUse(), false);
	}
};
CDocument.prototype.Get_DocumentPositionInfoForCollaborative = function()
{
	var DocPos = this.Get_CursorLogicPosition();
	if (!DocPos || DocPos.length <= 0)
		return null;

	var Last = DocPos[DocPos.length - 1];
	if (!(Last.Class instanceof ParaRun))
		return null;

	return Last;
};
CDocument.prototype.Update_ForeignCursor = function(CursorInfo, UserId, Show, UserShortId)
{
	if (!this.Api.User)
		return;

	if (UserId === this.Api.CoAuthoringApi.getUserConnectionId())
		return;

	var sUserName = this.Api.CoAuthoringApi.getParticipantName(UserId);

	// "" - это означает, что курсор нужно удалить
	if (!CursorInfo || "" === CursorInfo || !AscCommon.UserInfoParser.isUserVisible(sUserName))
	{
		this.Remove_ForeignCursor(UserId);
		return;
	}

	var Changes = new AscCommon.CCollaborativeChanges();
	var Reader  = Changes.GetStream(CursorInfo);

	var RunId    = Reader.GetString2();
	var InRunPos = Reader.GetLong();

	var Run = this.TableId.Get_ById(RunId);
	if (!Run)
	{
		this.Remove_ForeignCursor(UserId);
		return;
	}

	var CursorPos = [{Class : Run, Position : InRunPos}];
	Run.GetDocumentPositionFromObject(CursorPos);
	this.CollaborativeEditing.Add_ForeignCursor(UserId, CursorPos, UserShortId);

	if (true === Show)
		this.CollaborativeEditing.Update_ForeignCursorPosition(UserId, Run, InRunPos, true);
};
CDocument.prototype.Remove_ForeignCursor = function(UserId)
{
	this.CollaborativeEditing.Remove_ForeignCursor(UserId);
};
CDocument.prototype.private_UpdateTargetForCollaboration = function()
{
	this.NeedUpdateTargetForCollaboration = true;
};
CDocument.prototype.GetHdrFtr = function()
{
	return this.HdrFtr;
};
CDocument.prototype.Get_DrawingDocument = function()
{
	return this.DrawingDocument;
};
CDocument.prototype.GetDrawingDocument = function()
{
	return this.DrawingDocument;
};
CDocument.prototype.GetDrawingObjects = function()
{
	return this.DrawingObjects;
};
CDocument.prototype.Get_Api = function()
{
	return this.Api;
};
CDocument.prototype.GetApi = function()
{
	return this.Api;
};
CDocument.prototype.Get_IdCounter = function()
{
	return this.IdCounter;
};
CDocument.prototype.GetIdCounter = function()
{
	return this.IdCounter;
};
CDocument.prototype.Get_TableId = function()
{
	return this.TableId;
};
CDocument.prototype.GetTableId = function()
{
	return this.TableId;
};
CDocument.prototype.Get_History = function()
{
	return this.History;
};
CDocument.prototype.GetHistory = function()
{
	return this.History;
};
CDocument.prototype.Get_CollaborativeEditing = function()
{
	return this.CollaborativeEditing;
};
CDocument.prototype.GetDocumentOutline = function()
{
	return this.DocumentOutline;
};
CDocument.prototype.private_CorrectDocumentPosition = function()
{
	if (this.CurPos.ContentPos < 0 || this.CurPos.ContentPos >= this.Content.length)
	{
		this.RemoveSelection();
		if (this.CurPos.ContentPos < 0)
		{
			this.CurPos.ContentPos = 0;
			this.Content[0].MoveCursorToStartPos(false);
		}
		else
		{
			this.CurPos.ContentPos = this.Content.length - 1;
			this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false);
		}
	}
};
CDocument.prototype.private_ToggleParagraphAlignByHotkey = function(Align)
{
	var SelectedInfo = this.GetSelectedElementsInfo();
	var Math         = SelectedInfo.GetMath();
	if (null !== Math && true !== Math.Is_Inline())
	{
		var MathAlign = Math.Get_Align();
		if (Align !== MathAlign)
		{
			if (false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content))
			{
				this.StartAction(AscDFH.historydescription_Document_SetParagraphAlignHotKey);
				Math.Set_Align(Align);
				this.Recalculate();
				this.UpdateInterface();
				this.FinalizeAction();
			}
		}
	}
	else
	{
		var ParaPr = this.GetCalculatedParaPr();
		if (null != ParaPr)
		{
			if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Properties))
			{
				this.StartAction(AscDFH.historydescription_Document_SetParagraphAlignHotKey);
				this.SetParagraphAlign(ParaPr.Jc === Align ? (Align === align_Left ? AscCommon.align_Justify : align_Left) : Align);
				this.UpdateInterface();
				this.FinalizeAction();
			}
		}
	}
};
CDocument.prototype.Is_ShowParagraphMarks = function()
{
	return this.Api.ShowParaMarks;
};
CDocument.prototype.Set_ShowParagraphMarks = function(isShow, isRedraw)
{
	this.Api.ShowParaMarks = isShow;

	if (true === isRedraw)
	{
		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();
	}
};
CDocument.prototype.Get_StyleNameById = function(StyleId)
{
	if (!this.Styles)
		return "";

	var Style = this.Styles.Get(StyleId);
	if (!Style)
		return "";

	return Style.Get_Name();
};
CDocument.prototype.private_GetElementPageIndex = function(ElementPos, PageIndex, ColumnIndex, ColumnsCount)
{
	var Element = this.Content[ElementPos];
	if (!Element)
		return 0;

	var StartPage   = Element.Get_StartPage_Relative();
	var StartColumn = Element.Get_StartColumn();

	return ColumnIndex - StartColumn + (PageIndex - StartPage) * ColumnsCount;
};
CDocument.prototype.private_GetElementPageIndexByXY = function(ElementPos, X, Y, PageIndex)
{
	var Element = this.Content[ElementPos];
	if (!Element)
		return 0;

	var Page = this.Pages[PageIndex];
	if (!Page)
		return 0;

	var PageSection = null;
	for (var SectionIndex = 0, SectionsCount = Page.Sections.length; SectionIndex < SectionsCount; ++SectionIndex)
	{
		if (Page.Sections[SectionIndex].Pos <= ElementPos && ElementPos <= Page.Sections[SectionIndex].EndPos)
		{
			PageSection = Page.Sections[SectionIndex];
			break;
		}
	}

	if (!PageSection)
		return 0;

	var ElementStartPage   = Element.Get_StartPage_Relative();
	var ElementStartColumn = Element.Get_StartColumn();
	var ElementPagesCount  = Element.Get_PagesCount();

	var ColumnsCount = PageSection.Columns.length;
	var StartColumn  = 0;
	var EndColumn    = ColumnsCount - 1;

	if (PageIndex === ElementStartPage)
	{
		StartColumn = Element.Get_StartColumn();
		EndColumn   = Math.min(ElementStartColumn + ElementPagesCount - 1, ColumnsCount - 1);
	}
	else
	{
		StartColumn = 0;
		EndColumn   = Math.min(ElementPagesCount - ElementStartColumn + (PageIndex - ElementStartPage) * ColumnsCount, ColumnsCount - 1);
	}

	// TODO: Разобраться с ситуацией, когда пустые колонки стоят не только в конце
	while (true === PageSection.Columns[EndColumn].Empty && EndColumn > StartColumn)
		EndColumn--;

	var ResultColumn = EndColumn;
	for (var ColumnIndex = StartColumn; ColumnIndex < EndColumn; ++ColumnIndex)
	{
		if (X < (PageSection.Columns[ColumnIndex].XLimit + PageSection.Columns[ColumnIndex + 1].X) / 2)
		{
			ResultColumn = ColumnIndex;
			break;
		}
	}

	return this.private_GetElementPageIndex(ElementPos, PageIndex, ResultColumn, ColumnsCount);
};
CDocument.prototype.Get_DocumentPagePositionByContentPosition = function(ContentPosition)
{
	if (!ContentPosition)
		return null;

	var Para    = null;
	var ParaPos = 0;
	var Count   = ContentPosition.length;
	for (; ParaPos < Count; ++ParaPos)
	{
		var Element = ContentPosition[ParaPos].Class;
		if (Element instanceof Paragraph)
		{
			Para = Element;
			break;
		}
	}

	if (!Para)
		return null;

	var ParaContentPos = new AscWord.CParagraphContentPos();
	for (var Pos = ParaPos; Pos < Count; ++Pos)
	{
		ParaContentPos.Update(ContentPosition[Pos].Position, Pos - ParaPos);
	}

	var ParaPos = Para.Get_ParaPosByContentPos(ParaContentPos);
	if (!ParaPos)
		return;

	var Result    = new CDocumentPagePosition();
	Result.Page   = Para.Get_AbsolutePage(ParaPos.Page);
	Result.Column = Para.Get_AbsoluteColumn(ParaPos.Page);

	return Result;
};
CDocument.prototype.private_GetPageSectionByContentPosition = function(PageIndex, ContentPosition)
{
	var Page = this.Pages[PageIndex];
	if (!Page || !Page.Sections || Page.Sections.length <= 1)
		return 0;

	var SectionIndex = 0;
	for (var SectionsCount = Page.Sections.length; SectionIndex < SectionsCount; ++SectionIndex)
	{
		var Section = Page.Sections[SectionIndex];
		if (Section.Pos <= ContentPosition && ContentPosition <= Section.EndPos)
			break;
	}

	if (SectionIndex >= Page.Sections.length)
		return 0;

	return SectionIndex;
};
CDocument.prototype.Update_ColumnsMarkupFromRuler = function(oNewMarkup)
{
	var oSectPr = oNewMarkup.SectPr;
	if (!oSectPr)
		return;

	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr))
	{
		this.StartAction(AscDFH.historydescription_Document_SetColumnsFromRuler);

		oSectPr.Set_Columns_EqualWidth(oNewMarkup.EqualWidth);

		var nLeft  = oNewMarkup.X;
		var nRight = oSectPr.GetPageWidth() - oNewMarkup.R;

		if (this.IsMirrorMargins() && 1 === oNewMarkup.PageIndex % 2)
		{
			nLeft  = oSectPr.GetPageWidth() - oNewMarkup.R;
			nRight = oNewMarkup.X;
		}

		var nGutter = oSectPr.GetGutter();
		if (nGutter > 0.001 && !this.IsGutterAtTop())
		{
			if (oSectPr.IsGutterRTL())
				nRight = Math.max(0, nRight - nGutter);
			else
				nLeft = Math.max(0, nLeft - nGutter);
		}

		oSectPr.SetPageMargins(nLeft, undefined, nRight, undefined);

		if (false === oNewMarkup.EqualWidth)
		{
			for (var Index = 0, Count = oNewMarkup.Cols.length; Index < Count; ++Index)
			{
				oSectPr.Set_Columns_Col(Index, oNewMarkup.Cols[Index].W, oNewMarkup.Cols[Index].Space);
			}
		}
		else
		{
			oSectPr.Set_Columns_Space(oNewMarkup.Space);
			oSectPr.Set_Columns_Num(oNewMarkup.Num);
		}

		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.UpdateRulers();
		this.FinalizeAction();
	}
};
CDocument.prototype.Set_ColumnsProps = function(ColumnsProps)
{
	if (this.IsSelectionUse())
	{
		if (docpostype_DrawingObjects === this.GetDocPosType())
			return;

		// К селекту мы применяем колонки не так как ворд.
		// Элементы попавшие в селект полностью входят в  новую секцию, даже если они выделены частично.

		var nStartPos  = this.Selection.StartPos;
		var nEndPos    = this.Selection.EndPos;
		var nDirection = 1;

		if (nEndPos < nStartPos)
		{
			nStartPos  = this.Selection.EndPos;
			nEndPos    = this.Selection.StartPos;
			nDirection = -1;
		}

		var oStartSectPr = this.SectionsInfo.Get_SectPr(nStartPos).SectPr;
		var oEndSectPr   = this.SectionsInfo.Get_SectPr(nEndPos).SectPr;
		if (!oStartSectPr || !oEndSectPr || (oStartSectPr === oEndSectPr && oStartSectPr.IsEqualColumnProps(ColumnsProps)))
			return;

		if (this.IsSelectionLocked(AscCommon.changestype_Document_SectPr))
			return;

		this.StartAction(AscDFH.historydescription_Document_SetColumnsProps);

		var oEndParagraph = null;
		if (type_Paragraph !== this.Content[nEndPos].GetType())
		{
			oEndParagraph = new Paragraph(this.DrawingDocument, this);
			this.Add_ToContent(nEndPos + 1, oEndParagraph);
		}
		else
		{
			oEndParagraph = this.Content[nEndPos];
		}

		if (nStartPos > 0
			&& (type_Paragraph !== this.Content[nStartPos - 1].GetType()
			|| !this.Content[nStartPos - 1].Get_SectionPr()))
		{
			var oSectPr = new CSectionPr(this);
			oSectPr.Copy(oStartSectPr, false);

			var oStartParagraph = new Paragraph(this.DrawingDocument, this);
			this.Add_ToContent(nStartPos, oStartParagraph);
			oStartParagraph.Set_SectionPr(oSectPr, true);

			nStartPos++;
			nEndPos++;
		}

		if (nEndPos !== this.Content.length - 1)
		{
			oEndSectPr.Set_Type(c_oAscSectionBreakType.Continuous);
			var oSectPr = new CSectionPr(this);
			oSectPr.Copy(oEndSectPr, false);
			oEndParagraph.Set_SectionPr(oSectPr, true);
			oSectPr.SetColumnProps(ColumnsProps);
		}
		else
		{
			oEndSectPr.Set_Type(c_oAscSectionBreakType.Continuous);
			oEndSectPr.SetColumnProps(ColumnsProps);
		}

		for (var nIndex = nStartPos; nIndex < nEndPos; ++nIndex)
		{
			var oElement = this.Content[nIndex];
			if (type_Paragraph === oElement.GetType())
			{
				var oCurSectPr = oElement.Get_SectionPr();
				if (oCurSectPr)
					oCurSectPr.SetColumnProps(ColumnsProps);
			}
		}

		if (nDirection >= 0)
		{
			this.Selection.StartPos = nStartPos;
			this.Selection.EndPos   = nEndPos;
		}
		else
		{
			this.Selection.StartPos = nEndPos;
			this.Selection.EndPos   = nStartPos;
		}

		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.UpdateRulers();
		this.FinalizeAction();
	}
	else
	{
		var CurPos = this.CurPos.ContentPos;
		var SectPr = this.SectionsInfo.Get_SectPr(CurPos).SectPr;

		if (!SectPr)
			return;

		if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr))
		{
			this.StartAction(AscDFH.historydescription_Document_SetColumnsProps);

			SectPr.SetColumnProps(ColumnsProps);

			this.Recalculate();
			this.UpdateSelection();
			this.UpdateInterface();
			this.UpdateRulers();
			this.FinalizeAction();
		}
	}
};
CDocument.prototype.GetTopDocumentContent = function(isOneLevel)
{
	return this;
};
CDocument.prototype.Set_SectionProps = function(Props)
{
	var CurPos = this.CurPos.ContentPos;
	var SectPr = this.SectionsInfo.Get_SectPr(CurPos).SectPr;

	if (SectPr && false === this.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr))
	{
		this.StartAction(AscDFH.historydescription_Document_SetSectionProps);

		if (undefined !== Props.get_W() || undefined !== Props.get_H())
		{
			var PageW = undefined !== Props.get_W() ? Props.get_W() : SectPr.GetPageWidth();
			var PageH = undefined !== Props.get_H() ? Props.get_H() : SectPr.GetPageHeight();
			SectPr.SetPageSize(PageW, PageH);
		}

		if (undefined !== Props.get_LeftMargin() || undefined !== Props.get_TopMargin() || undefined !== Props.get_RightMargin() || undefined !== Props.get_BottomMargin())
		{
			// Внутри функции идет разруливания, если какое то значение undefined
			SectPr.SetPageMargins(Props.get_LeftMargin(), Props.get_TopMargin(), Props.get_RightMargin(), Props.get_BottomMargin());
		}

		if (undefined !== Props.get_HeaderDistance())
		{
			SectPr.SetPageMarginHeader(Props.get_HeaderDistance());
		}

		if (undefined !== Props.get_FooterDistance())
		{
			SectPr.SetPageMarginFooter(Props.get_FooterDistance());
		}

		if (undefined !== Props.get_Gutter())
			SectPr.SetGutter(Props.get_Gutter());

		if (undefined !== Props.get_GutterRTL())
			SectPr.SetGutterRTL(Props.get_GutterRTL());

		if (undefined !== Props.get_GutterAtTop())
			this.SetGutterAtTop(Props.get_GutterAtTop());

		if (undefined !== Props.get_MirrorMargins())
			this.SetMirrorMargins(Props.get_MirrorMargins());

		if (undefined !== Props.get_Orientation())
			SectPr.SetOrientation(Props.get_Orientation(), false);

		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.UpdateRulers();
		this.FinalizeAction();
	}
};
CDocument.prototype.Get_SectionProps = function()
{
	var CurPos = this.CurPos.ContentPos;
	var SectPr = this.SectionsInfo.Get_SectPr(CurPos).SectPr;

	return new Asc.CDocumentSectionProps(SectPr, this);
};
/**
 * Получаем ширину текущей колонки
 * @returns {number}
 */
CDocument.prototype.GetCurrentColumnWidth = function()
{
	var nCurPos = 0;
	if (this.Controller === this.LogicDocumentController)
		nCurPos = this.Selection.Use ? this.Selection.EndPos : this.CurPos.ContentPos;
	else
		nCurPos = this.CurPos.ContentPos;

	var oSectPr       = this.SectionsInfo.Get_SectPr(nCurPos).SectPr;
	var nColumnsCount = oSectPr.Get_ColumnsCount();

	if (nColumnsCount > 1)
	{
		var oParagraph = this.GetCurrentParagraph();
		if (!oParagraph)
			return 0;

		var nCurrentColumn = oParagraph.Get_CurrentColumn();
		return oSectPr.Get_ColumnWidth(nCurrentColumn);
	}

	return oSectPr.GetContentFrameWidth();
};
CDocument.prototype.Get_FirstParagraph = function()
{
	if (type_Paragraph == this.Content[0].GetType())
		return this.Content[0];
	else if (type_Table == this.Content[0].GetType())
		return this.Content[0].Get_FirstParagraph();

	return null;
};
/**
 * Обработчик нажатия кнопки IncreaseIndent в меню.
 */
CDocument.prototype.IncreaseIndent = function()
{
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Properties))
	{
		this.StartAction(AscDFH.historydescription_Document_IncParagraphIndent);
		this.IncreaseDecreaseIndent(true);
		this.UpdateSelection();
		this.UpdateInterface();
		this.Recalculate();
		this.FinalizeAction();
	}
};
/**
 * Обработчик нажатия кнопки DecreaseIndent в меню.
 */
CDocument.prototype.DecreaseIndent = function()
{
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Properties))
	{
		this.StartAction(AscDFH.historydescription_Document_DecParagraphIndent);
		this.IncreaseDecreaseIndent(false);
		this.UpdateSelection();
		this.UpdateInterface();
		this.Recalculate();
		this.FinalizeAction();
	}
};
/**
 * Получаем размеры текущей колонки (используется при вставке изображения).
 */
CDocument.prototype.GetColumnSize = function()
{
	return this.Controller.GetColumnSize();
};
CDocument.prototype.private_OnSelectionEnd = function()
{
	this.Api.sendEvent("asc_onSelectionEnd");
};
CDocument.prototype.private_OnSelectionCancel = function()
{
	this.Api.sendEvent("asc_onSelectionCancel");
};
CDocument.prototype.AddPageCount = function()
{
	if (false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_AddPageCount);
		this.AddToParagraph(new AscWord.CRunPagesCount(this.Pages.length));
		this.FinalizeAction();
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Settings
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.GetDocumentSettings = function()
{
	return this.Settings;
};
CDocument.prototype.GetCompatibilityMode = function()
{
	return this.Settings.CompatibilityMode;
};
CDocument.prototype.GetSdtGlobalColor = function()
{
	return this.Settings.SdtSettings.Color;
};
CDocument.prototype.SetSdtGlobalColor = function(r, g, b)
{
	var oNewColor = new CDocumentColor(r, g, b);
	if (!oNewColor.Compare(this.Settings.SdtSettings.Color))
	{
		var oNewSettings = this.Settings.SdtSettings.Copy();
		oNewSettings.Color = oNewColor;

		this.History.Add(new CChangesDocumentSdtGlobalSettings(this, this.Settings.SdtSettings, oNewSettings));
		this.Settings.SdtSettings = oNewSettings;

		this.OnChangeSdtGlobalSettings();
	}
};
CDocument.prototype.GetSdtGlobalShowHighlight = function()
{
	return this.Settings.SdtSettings.ShowHighlight;
};
CDocument.prototype.SetSdtGlobalShowHighlight = function(isShow)
{
	if (this.Settings.SdtSettings.ShowHighlight !== isShow)
	{
		var oNewSettings = this.Settings.SdtSettings.Copy();
		oNewSettings.ShowHighlight = isShow;

		this.History.Add(new CChangesDocumentSdtGlobalSettings(this, this.Settings.SdtSettings, oNewSettings));
		this.Settings.SdtSettings = oNewSettings;

		this.OnChangeSdtGlobalSettings();
	}
};
CDocument.prototype.OnChangeSdtGlobalSettings = function()
{
	this.GetApi().sync_OnChangeSdtGlobalSettings();
};
CDocument.prototype.IsSplitPageBreakAndParaMark = function()
{
	return this.Settings.SplitPageBreakAndParaMark;
};
CDocument.prototype.IsDoNotExpandShiftReturn = function()
{
	return this.Settings.DoNotExpandShiftReturn;
};
CDocument.prototype.IsBalanceSingleByteDoubleByteWidth = function()
{
	return (this.Settings.BalanceSingleByteDoubleByteWidth && (this.Styles.IsValidDefaultEastAsiaFont() || this.Settings.UseFELayout));
};
CDocument.prototype.IsUnderlineTrailSpace = function()
{
	return this.Settings.UlTrailSpace;
};
/**
 * Проверяем все ли параметры SdtSettings выставлены по умолчанию
 * @returns {boolean}
 */
CDocument.prototype.IsSdtGlobalSettingsDefault = function()
{
	return this.Settings.SdtSettings.IsDefault();
};
CDocument.prototype.GetSpecialFormsHighlight = function()
{
	if (!this.Settings.SpecialFormsSettings.Highlight)
		return new AscCommonWord.CDocumentColor(201, 200, 255);

	return this.Settings.SpecialFormsSettings.Highlight;
};
CDocument.prototype.SetSpecialFormsHighlight = function(r, g, b)
{
	if ((undefined === r || null === r) && (undefined === this.Settings.SpecialFormsSettings.Highlight || !this.Settings.SpecialFormsSettings.Highlight.IsAuto()))
	{
		var oNewSettings = this.Settings.SpecialFormsSettings.Copy();
		oNewSettings.Highlight = new AscCommonWord.CDocumentColor(0, 0, 0, true);

		this.History.Add(new CChangesDocumentSpecialFormsGlobalSettings(this, this.Settings.SpecialFormsSettings, oNewSettings));
		this.Settings.SpecialFormsSettings = oNewSettings;

		this.OnChangeSpecialFormsGlobalSettings();
	}
	else if (undefined !== r && null !== r)
	{
		var oNewColor = new CDocumentColor(r, g, b);
		if (!oNewColor.IsEqual(this.Settings.SpecialFormsSettings.Highlight))
		{
			var oNewSettings = this.Settings.SpecialFormsSettings.Copy();
			oNewSettings.Highlight = oNewColor;

			this.History.Add(new CChangesDocumentSpecialFormsGlobalSettings(this, this.Settings.SpecialFormsSettings, oNewSettings));
			this.Settings.SpecialFormsSettings = oNewSettings;

			this.OnChangeSpecialFormsGlobalSettings();
		}
	}
};
CDocument.prototype.OnChangeSpecialFormsGlobalSettings = function()
{
	this.GetApi().sync_OnChangeSpecialFormsGlobalSettings();
};
CDocument.prototype.IsSpecialFormsSettingsDefault = function()
{
	return this.Settings.SpecialFormsSettings.IsDefault();
};
CDocument.prototype.SetProtection = function(props)
{
	if (this.Settings.DocumentProtection) {
		this.Settings.DocumentProtection.setProps(props);
	}
};

CDocument.prototype.GetDocumentLayout = function()
{
	return this.Layout;
};
CDocument.prototype.SetDocumentReadMode = function(nW, nH, nScale)
{
	this.Layouts.Read.Set(nW, nH, nScale);
	this.Layout = this.Layouts.Read;

	this.CheckAllRunContent(function(oRun)
	{
		oRun.Recalc_CompiledPr(true);
	});

	this.RecalculateFromStart(true);
};
CDocument.prototype.SetDocumentPrintMode = function()
{
	this.Layout = this.Layouts.Print;

	this.CheckAllRunContent(function(oRun)
	{
		oRun.Recalc_CompiledPr(true);
	});

	this.RecalculateFromStart(true);
};
CDocument.prototype.private_SetCurrentSpecialForm = function(oForm)
{
	if (this.CurrentForm === oForm && ((!oForm && !this.CurrentFormFixed) || (oForm && this.CurrentFormFixed === oForm.IsFixedForm())))
		return false;

	if (this.CurrentForm)
		this.CurrentForm.SetCurrent(false);

	this.CurrentForm = oForm;
	this.CurrentFormFixed = oForm ? oForm.IsFixedForm() : false;

	if (this.CurrentForm)
		this.CurrentForm.SetCurrent(true);
	
	return true;
};
/**
 * Добавляем специальный контейнер в виде чекбокса
 * @param oPr {?AscWord.CSdtCheckBoxPr}
 * @returns {CInlineLevelSdt | CBlockLevelSdt}
 */
CDocument.prototype.AddContentControlCheckBox = function(oPr)
{
	this.RemoveTextSelection();

	if (!oPr)
		oPr = new AscWord.CSdtCheckBoxPr();

	var oTextPr = this.GetDirectTextPr();
	var oCC = this.AddContentControl(c_oAscSdtLevelType.Inline);
	if (!oCC)
		return;

	oCC.ApplyCheckBoxPr(oPr, oTextPr);
	oCC.MoveCursorToStartPos();

	this.UpdateSelection();
	this.UpdateTracks();

	return oCC;
};
/**
 * Добавляем специальный контейнер для картинок
 * @returns {CInlineLevelSdt | CBlockLevelSdt}
 */
CDocument.prototype.AddContentControlPicture = function()
{
	this.RemoveTextSelection();

	var oCC = this.AddContentControl(c_oAscSdtLevelType.Inline);
	if (!oCC)
		return null;

	oCC.SetPlaceholderText(AscCommon.translateManager.getValue("Click to load image"));
	oCC.ApplyPicturePr(true);
	return oCC;
};
/**
 * Добавляем контйенер с полем для спискам
 * @param oPr {?AscWord.CSdtComboBoxPr}
 */
CDocument.prototype.AddContentControlComboBox = function(oPr)
{
	this.RemoveTextSelection();

	if (!oPr)
	{
		oPr = new AscWord.CSdtComboBoxPr();
		oPr.AddItem(AscCommon.translateManager.getValue("Choose an item"), "");
	}

	var oCC = this.AddContentControl(c_oAscSdtLevelType.Inline);
	if (!oCC)
		return null;

	oCC.ApplyComboBoxPr(oPr);
	oCC.SelectContentControl();
	return oCC;
};
/**
 * Добавляем контейнер с выпалающим списком
 * @param oPr {?AscWord.CSdtComboBoxPr}
 */
CDocument.prototype.AddContentControlDropDownList = function(oPr)
{
	this.RemoveTextSelection();

	if (!oPr)
	{
		oPr = new AscWord.CSdtComboBoxPr();
		oPr.AddItem(AscCommon.translateManager.getValue("Choose an item"), "");
	}

	var oCC = this.AddContentControl(c_oAscSdtLevelType.Inline);
	if (!oCC)
		return null;

	oCC.ApplyDropDownListPr(oPr);
	oCC.SelectContentControl();
	return oCC;
};
/**
 * Добавляем специальный контейнер с выбором даты
 * @param oPr {?AscWord.CSdtDatePickerPr}
 */
CDocument.prototype.AddContentControlDatePicker = function(oPr)
{
	this.RemoveTextSelection();

	if (!oPr)
		oPr = new AscWord.CSdtDatePickerPr();

	var oCC = this.AddContentControl(c_oAscSdtLevelType.Inline);
	if (!oCC)
		return null;
	
	oCC.SetPlaceholder(c_oAscDefaultPlaceholderName.DateTime);
	oCC.ApplyDatePickerPr(oPr);
	oCC.SelectContentControl();
	return oCC;
};
/**
 * Добавляем специальную текстовую форму
 * @param oPr {?AscWord.CSdtTextFormPr}
 * @returns {CInlineLevelSdt | CBlockLevelSdt | null}
 */
CDocument.prototype.AddContentControlTextForm = function(oPr)
{
	if (!oPr)
		oPr = new AscWord.CSdtTextFormPr();

	var sText   = this.GetSelectedText(false, {SkipPlaceholder : true});
	var oTextPr = this.GetDirectTextPr();

	if (this.IsTextSelectionUse())
		this.RemoveBeforePaste();
	else if (this.IsSelectionUse())
		this.RemoveSelection();

	var oCC = this.AddContentControl(c_oAscSdtLevelType.Inline);

	if (oPr.Comb)
	{
		if (!oPr.MaxCharacters || !oPr.MaxCharacters <= 0)
			oPr.MaxCharacters = 12;

		var oDocPart
		if (oPr.CombPlaceholderSymbol)
			oCC.SetPlaceholderText(String.fromCharCode(oPr.CombPlaceholderSymbol));
		else
			oCC.SetPlaceholder(c_oAscDefaultPlaceholderName.TextForm);

		if (oDocPart && oPr.CombPlaceholderFont)
		{
			oDocPart.SelectAll();
			oDocPart.AddToParagraph(new ParaTextPr({
				RFonts : {
					Ascii    : {Name : oPr.CombPlaceholderFont, Index : -1},
					EastAsia : {Name : oPr.CombPlaceholderFont, Index : -1},
					HAnsi    : {Name : oPr.CombPlaceholderFont, Index : -1},
					CS       : {Name : oPr.CombPlaceholderFont, Index : -1}
				}
			}));
			oDocPart.RemoveSelection();
		}
	}

	if (!oCC)
		return null;

	oCC.ApplyTextFormPr(oPr);
	oCC.MoveCursorToStartPos();

	if (sText && oCC instanceof CInlineLevelSdt)
	{
		oCC.ReplacePlaceHolderWithContent();
		var oRun = oCC.MakeSingleRunElement(false);
		oRun.AddText(sText);
		oRun.ApplyTextPr(oTextPr);
		oCC.SelectContentControl();
	}

	this.UpdateSelection();
	this.UpdateTracks();

	return oCC;
};
/**
 * Добавляем специальную текстовую форму
 * @param {?AscWord.CSdtComplexFormPr} oPr
 * @param {?AscWord.CSdtFormPr} formPr
 * @returns {?CInlineLevelSdt}
 */
CDocument.prototype.AddComplexForm = function(oPr, formPr)
{
	if (!oPr)
		oPr = new AscWord.CSdtComplexFormPr();

	let sText   = this.GetSelectedText(false, {SkipPlaceholder : true});
	let oTextPr = this.GetDirectTextPr();

	if (this.IsTextSelectionUse())
		this.RemoveBeforePaste();
	else if (this.IsSelectionUse())
		this.RemoveSelection();

	let oCC = this.AddContentControl(c_oAscSdtLevelType.Inline);
	if (!oCC)
		return null;

	oCC.SetComplexFormPr(oPr);
	oCC.MoveCursorToStartPos();

	let _formPr = formPr ? formPr : new AscWord.CSdtFormPr();
	oCC.SetFormPr(_formPr);

	if (sText)
	{
		oCC.ReplacePlaceHolderWithContent();
		let oRun = oCC.MakeSingleRunElement(false);
		oRun.AddText(sText);
		oRun.ApplyTextPr(oTextPr);
		oCC.SelectContentControl();
	}
	
	if (_formPr.GetFixed())
		oCC.ConvertFormToFixed();

	this.UpdateSelection();
	this.UpdateTracks();

	return oCC;
};
/**
 * Выставляем плейсхолдер для заданного контент контрола
 * @param {string} sText
 * @param {CBlockLevelSdt | CInlineLevelSdt} oCC
 */
CDocument.prototype.SetContentControlTextPlaceholder = function(sText, oCC)
{
	if (!oCC)
		return;

	if (!sText)
		sText = String.fromCharCode(nbsp_charcode, nbsp_charcode, nbsp_charcode, nbsp_charcode);

	var oGlossary = this.GetGlossaryDocument();

	if (!this.IsSelectionLocked(AscCommon.changestype_None, {
		Type      : AscCommon.changestype_2_Element_and_Type,
		Element   : oGlossary,
		CheckType : AscCommon.changestype_Document_Content
	}))
	{
		this.StartAction(AscDFH.historydescription_Document_SetContentControlTextPlaceholder);

		oCC.SetPlaceholderText(sText);
		if (oCC.IsPlaceHolder())
		{
			this.Recalculate();
			this.UpdateInterface();
			this.UpdateSelection();
		}

		this.FinalizeAction();
	}
};
/**
 * Выставляем текст у заданного заданного контент контрола
 * @param {string} sText
 * @param {CBlockLevelSdt | CInlineLevelSdt} oCC
 */
CDocument.prototype.SetContentControlText = function(sText, oCC)
{
	if (!oCC || oCC.IsCheckBox() || oCC.IsDropDownList() || oCC.IsPicture())
		return;

	if (!sText)
		return;

	oCC.SelectContentControl();

	if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, false, this.IsFormFieldEditing()))
	{
		this.StartAction(AscDFH.historydescription_Document_SetContentControlText);

		oCC.ReplacePlaceHolderWithContent();

		var oRun;
		if (oCC.IsBlockLevel())
			oRun = this.Content.MakeSingleParagraphContent().MakeSingleRunParagraph(true);
		else if (oCC.IsInlineLevel())
			oRun = oCC.MakeSingleRunElement(true);

		if (oRun)
			oRun.AddText(sText);

		this.RemoveSelection();
		oCC.MoveCursorToContentControl(false);

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();

		this.FinalizeAction();
	}
};
/**
 * Выставляем глобальный параметр, находится ли переплет наверху документа
 * @param {boolean} isGutterAtTop
 */
CDocument.prototype.SetGutterAtTop = function(isGutterAtTop)
{
	if (isGutterAtTop !== this.Settings.GutterAtTop)
	{
		this.History.Add(new CChangesDocumentSettingsGutterAtTop(this, this.Settings.GutterAtTop, isGutterAtTop));
		this.Settings.GutterAtTop = isGutterAtTop;
	}
};
/**
 * Проверяем находится ли переплет наверху документа
 * @returns {boolean}
 */
CDocument.prototype.IsGutterAtTop = function()
{
	return this.Settings.GutterAtTop;
};
/**
 * Выставляем зеркальные отступы
 * @param {boolean} isMirror
 */
CDocument.prototype.SetMirrorMargins = function(isMirror)
{
	if (isMirror !== this.Settings.MirrorMargins)
	{
		this.History.Add(new CChangesDocumentSettingsMirrorMargins(this, this.Settings.MirrorIndent, isMirror));
		this.Settings.MirrorMargins = isMirror;
	}
};
/**
 * Проверяем зеркальные ли отступы
 * @returns {boolean}
 */
CDocument.prototype.IsMirrorMargins = function()
{
	return this.Settings.MirrorMargins;
};
//----------------------------------------------------------------------------------------------------------------------
// Math
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.Set_MathProps = function(MathProps)
{
	var SelectedInfo = this.GetSelectedElementsInfo();
	if (null !== SelectedInfo.GetMath() && false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_SetMathProps);

		var ParaMath = SelectedInfo.GetMath();
		ParaMath.Set_MenuProps(MathProps);

		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Statistics (Функции для работы со статистикой)
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.Statistics_Start = function()
{
	this.Statistics.Start();
};
CDocument.prototype.Statistics_GetParagraphsInfo = function()
{
	var Count   = this.Content.length;
	var CurPage = this.Statistics.CurPage;

	var Index    = 0;
	var CurIndex = 0;
	for (Index = this.Statistics.StartPos; Index < Count; ++Index, ++CurIndex)
	{
		var Element = this.Content[Index];
		Element.CollectDocumentStatistics(this.Statistics);

		if (CurIndex > 20)
		{
			this.Statistics.Next_ParagraphsInfo(Index + 1);
			break;
		}
	}

	if (Index >= Count)
		this.Statistics.Stop_ParagraphsInfo();
};
CDocument.prototype.Statistics_GetPagesInfo = function()
{
	this.Statistics.Update_Pages(this.Pages.length);

	if (null !== this.FullRecalc.Id)
	{
		this.Statistics.Next_PagesInfo();
	}
	else
	{
		for (var CurPage = 0, PagesCount = this.Pages.length; CurPage < PagesCount; ++CurPage)
		{
			this.DrawingObjects.documentStatistics(CurPage, this.Statistics);
		}

		this.Statistics.Stop_PagesInfo();
	}
};
CDocument.prototype.Statistics_Stop = function()
{
	this.Statistics.Stop();
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с MailMerge
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.Start_MailMerge = function(MailMergeMap, arrFields)
{
	this.EndPreview_MailMergeResult();

	this.MailMergeMap    = MailMergeMap;
	this.MailMergeFields = arrFields;
	this.Api.sync_HighlightMailMergeFields(this.MailMergeFieldsHighlight);
	this.Api.sync_StartMailMerge();
};
CDocument.prototype.Get_MailMergeReceptionsCount = function()
{
	if (null === this.MailMergeMap || !this.MailMergeMap)
		return 0;

	return this.MailMergeMap.length;
};
CDocument.prototype.Get_MailMergeFieldsNameList = function()
{
	if (this.Get_MailMergeReceptionsCount() <= 0)
		return this.MailMergeFields;

	// Предполагаем, что в первом элементе перечислены все поля
	var Element = this.MailMergeMap[0];
	var aList = [];
	for (var sId in Element)
	{
		aList.push(sId);
	}

	return aList;
};
CDocument.prototype.Add_MailMergeField = function(Name)
{
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_AddMailMergeField);

		var oField = new ParaField(fieldtype_MERGEFIELD, [Name], []);
		var oRun = new ParaRun();
		oRun.AddText("«" + Name + "»");
		oField.Add_ToContent(0, oRun);

		this.Register_Field(oField);
		this.AddToParagraph(oField);
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
CDocument.prototype.Set_HightlighMailMergeFields = function(Value)
{
	if (Value !== this.MailMergeFieldsHighlight)
	{
		this.MailMergeFieldsHighlight = Value;
		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();
		this.DrawingDocument.Update_FieldTrack(false);
		this.Api.sync_HighlightMailMergeFields(this.MailMergeFieldsHighlight);
	}
};
CDocument.prototype.Preview_MailMergeResult = function(Index)
{
	if (null === this.MailMergeMap)
		return;

	if (true !== this.MailMergePreview)
	{
		this.MailMergePreview = true;
		this.RemoveSelection();
		AscCommon.CollaborativeEditing.Set_GlobalLock(true);
	}

	this.FieldsManager.Update_MailMergeFields(this.MailMergeMap[Index]);
	this.RecalculateFromStart(true);

	this.Api.sync_PreviewMailMergeResult(Index);
};
CDocument.prototype.EndPreview_MailMergeResult = function()
{
	if (null === this.MailMergeMap || true !== this.MailMergePreview)
		return;

	this.MailMergePreview = false;
	AscCommon.CollaborativeEditing.Set_GlobalLock(false);

	this.FieldsManager.Restore_MailMergeTemplate();
	this.RecalculateFromStart(true);

	this.Api.sync_EndPreviewMailMergeResult();
};
CDocument.prototype.Get_MailMergeReceptionsList = function()
{
	var aList = [];

	var aHeaders = [];
	var nCount = this.MailMergeMap.length;
	if (nCount <= 0)
		return [this.MailMergeFields];

	for (var sId in this.MailMergeMap[0])
		aHeaders.push(sId);

	var nHeadersCount = aHeaders.length;

	aList.push(aHeaders);
	for (var nIndex = 0; nIndex < nCount; nIndex++)
	{
		var aReception = [];
		var oReception = this.MailMergeMap[nIndex];
		for (var nHeaderIndex = 0; nHeaderIndex < nHeadersCount; nHeaderIndex++)
		{
			var sValue = oReception[aHeaders[nHeaderIndex]];
			aReception.push(sValue ? sValue : "");
		}

		aList.push(aReception);
	}

	return aList;
};
CDocument.prototype.Get_MailMergeFieldValue = function(nIndex, sName)
{
	if (null === this.MailMergeMap)
		return null;

	return this.MailMergeMap[nIndex][sName];
};
CDocument.prototype.Get_MailMergedDocument = function(_nStartIndex, _nEndIndex)
{
	var nStartIndex = (undefined !== _nStartIndex ? Math.max(0, _nStartIndex) : 0);
	var nEndIndex   = (undefined !== _nEndIndex   ? Math.min(_nEndIndex, this.MailMergeMap.length - 1) : this.MailMergeMap.length - 1);

	if (null === this.MailMergeMap || nStartIndex > nEndIndex || nStartIndex >= this.MailMergeMap.length)
		return null;

	AscCommon.History.TurnOff();
	AscCommon.g_oTableId.TurnOff();

	var LogicDocument = new CDocument(undefined, false);
	AscCommon.History.Document = this;

	// Копируем стили, они все одинаковые для всех документов
	LogicDocument.Styles = this.Styles.Copy();

	// Нумерацию придется повторить для каждого отдельного файла
	LogicDocument.Numbering.Clear();

	LogicDocument.DrawingDocument = this.DrawingDocument;

	LogicDocument.theme = this.theme.createDuplicate();
	LogicDocument.clrSchemeMap   = this.clrSchemeMap.createDuplicate();

	var FieldsManager = this.FieldsManager;

	var ContentCount = this.Content.length;
	var OverallIndex = 0;
	this.ForceCopySectPr = true;

	for (var Index = nStartIndex; Index <= nEndIndex; Index++)
	{
		// Подменяем ссылку на менеджер полей, чтобы скопированные поля регистрировались в новом классе
		this.FieldsManager = LogicDocument.FieldsManager;
		var NewNumbering = this.Numbering.CopyAllNums(LogicDocument.Numbering);

		LogicDocument.Numbering.AppendAbstractNums(NewNumbering.AbstractNum);
		LogicDocument.Numbering.AppendNums(NewNumbering.Num);

		this.CopyNumberingMap = NewNumbering.NumMap;

		for (var ContentIndex = 0; ContentIndex < ContentCount; ContentIndex++)
		{
			LogicDocument.Content[OverallIndex++] = this.Content[ContentIndex].Copy(LogicDocument, this.DrawingDocument);

			if (type_Paragraph === this.Content[ContentIndex].Get_Type())
			{
				var ParaSectPr = this.Content[ContentIndex].Get_SectionPr();
				if (ParaSectPr)
				{
					var NewParaSectPr = new CSectionPr();
					NewParaSectPr.Copy(ParaSectPr, true);
					LogicDocument.Content[OverallIndex - 1].Set_SectionPr(NewParaSectPr, false);
				}
			}
		}

		// Добавляем дополнительный параграф с окончанием секции
		var SectionPara = new Paragraph(this.DrawingDocument, this);
		var SectPr = new CSectionPr(LogicDocument);
		SectPr.Copy(this.SectPr, true);
		SectPr.Set_Type(c_oAscSectionBreakType.NextPage);
		SectionPara.Set_SectionPr(SectPr, false);
		LogicDocument.Content[OverallIndex++] = SectionPara;

		LogicDocument.FieldsManager.Replace_MailMergeFields(this.MailMergeMap[Index]);
	}

	this.CopyNumberingMap = null;
	this.ForceCopySectPr  = false;

	// Добавляем дополнительный параграф в самом конце для последней секции документа
	var SectPara = new Paragraph(this.DrawingDocument, this);
	LogicDocument.Content[OverallIndex++] = SectPara;
	LogicDocument.SectPr.Copy(this.SectPr);
	LogicDocument.SectPr.Set_Type(c_oAscSectionBreakType.Continuous);

	for (var Index = 0, Count = LogicDocument.Content.length; Index < Count; Index++)
	{
		if (0 === Index)
			LogicDocument.Content[Index].Prev = null;
		else
			LogicDocument.Content[Index].Prev = LogicDocument.Content[Index - 1];

		if (Count - 1 === Index)
			LogicDocument.Content[Index].Next = null;
		else
			LogicDocument.Content[Index].Next = LogicDocument.Content[Index + 1];

		LogicDocument.Content[Index].Parent = LogicDocument;
	}

	this.FieldsManager = FieldsManager;
	AscCommon.g_oTableId.TurnOn();
	AscCommon.History.TurnOn();

	return LogicDocument;
};
CDocument.prototype.ContinueTrackRevisions = function()
{
	this.TrackRevisionsManager.ContinueTrackRevisions();
};
CDocument.prototype.SetTrackRevisions = function(bTrack)
{
	return this.SetLocalTrackRevisions(bTrack);
};
/**
 * Проверяем включено ли рецезирование в файле
 * @return {boolean}
 */
CDocument.prototype.IsTrackRevisions = function()
{
	if (null === this.TrackRevisions)
		return (true === this.Settings.TrackRevisions);

	return this.TrackRevisions;
};
/**
 * Устанавливаем локальное значение TrackChanges
 * @param {boolean|null} isTrack
 * @param {boolean} [isUpdateInterface=false] посылаем сообщение в интерфейс об изменении данного флага
 */
CDocument.prototype.SetLocalTrackRevisions = function(isTrack, isUpdateInterface)
{
	this.TrackRevisions = isTrack;

	if (true === isUpdateInterface)
		this.private_OnTrackRevisionsChange();
};
/**
 * Получаем локальный флаг рецензирования
 * @return {null|boolean}
 */
CDocument.prototype.GetLocalTrackRevisions = function()
{
	return this.TrackRevisions;
};
/**
 * Устанавливаем глобальный флаг рецензирования в файле
 * @param {boolean} isTrack
 * @param {boolean} [isUpdateInterface=false] посылаем сообщение в интерфейс об изменении данного флага
 */
CDocument.prototype.SetGlobalTrackRevisions = function(isTrack, isUpdateInterface)
{
	if (isTrack !== this.Settings.TrackRevisions)
	{
		this.History.Add(new CChangesDocumentSettingsTrackRevisions(this, this.Settings.TrackRevisions, isTrack, this.GetUserId()));
		this.Settings.TrackRevisions = isTrack;
	}

	if (true === isUpdateInterface)
		this.private_OnTrackRevisionsChange();
};
/**
 * Получаем глобальный флаг рецензирования в файле
 * @returns {boolean}
 */
CDocument.prototype.GetGlobalTrackRevisions = function()
{
	return this.Settings.TrackRevisions;
};
CDocument.prototype.private_OnTrackRevisionsChange = function(sUserId)
{
	this.Api.sync_OnTrackRevisionsChange(this.TrackRevisions, this.Settings.TrackRevisions, sUserId);
};
CDocument.prototype.GetNextRevisionChange = function()
{
	this.TrackRevisionsManager.ContinueTrackRevisions();
	var oChange = this.TrackRevisionsManager.GetNextChange();
	if (oChange)
	{
		this.RemoveSelection();
		this.private_SelectRevisionChange(oChange);
		this.UpdateSelection(false);
		this.UpdateInterface(true);
	}
};
CDocument.prototype.GetPrevRevisionChange = function()
{
	this.TrackRevisionsManager.ContinueTrackRevisions();
	var oChange = this.TrackRevisionsManager.GetPrevChange();
	if (oChange)
	{
		this.RemoveSelection();
		this.private_SelectRevisionChange(oChange);
		this.UpdateSelection(false);
		this.UpdateInterface(true);
	}
};
CDocument.prototype.GetRevisionsChangeElement = function(nDirection, oCurrentElement)
{
	return this.private_GetRevisionsChangeElement(nDirection, oCurrentElement).GetFoundedElement();
};
CDocument.prototype.private_GetRevisionsChangeElement = function(nDirection, oCurrentElement)
{
	var oSearchEngine = new CRevisionsChangeParagraphSearchEngine(nDirection, oCurrentElement, this.TrackRevisionsManager);
	if (null === oCurrentElement)
	{
		oCurrentElement = this.GetCurrentParagraph(false, false, {ReturnSelectedTable : true});
		if (null === oCurrentElement)
			return oSearchEngine;

		oSearchEngine.SetCurrentElement(oCurrentElement);
		oSearchEngine.SetFoundedElement(oCurrentElement);
		if (true === oSearchEngine.IsFound())
			return oSearchEngine;
	}

	var oFootnote = oCurrentElement.Parent ? oCurrentElement.Parent.GetTopDocumentContent() : null;
	if (!(oFootnote instanceof CFootEndnote))
		oFootnote = null;

	var HdrFtr = oCurrentElement.GetHdrFtr();
	if (null !== HdrFtr)
	{
		this.private_GetRevisionsChangeElementInHdrFtr(oSearchEngine, HdrFtr);

		if (nDirection > 0)
		{
			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInDocument(oSearchEngine, 0);

			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInFootnotes(oSearchEngine, null);
		}
		else
		{
			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInFootnotes(oSearchEngine, null);

			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInDocument(oSearchEngine, this.Content.length - 1);
		}


		if (true !== oSearchEngine.IsFound())
			this.private_GetRevisionsChangeElementInHdrFtr(oSearchEngine, null);
	}
	else if (oFootnote)
	{
		this.private_GetRevisionsChangeElementInFootnotes(oSearchEngine, oFootnote);

		if (nDirection > 0)
		{
			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInHdrFtr(oSearchEngine, null);

			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInDocument(oSearchEngine, 0);
		}
		else
		{
			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInDocument(oSearchEngine, this.Content.length - 1);

			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInHdrFtr(oSearchEngine, null);
		}

		if (true !== oSearchEngine.IsFound())
			this.private_GetRevisionsChangeElementInFootnotes(oSearchEngine, null);
	}
	else
	{
		var Pos = (true === this.Selection.Use && docpostype_DrawingObjects !== this.GetDocPosType() ? (this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos) : this.CurPos.ContentPos);
		this.private_GetRevisionsChangeElementInDocument(oSearchEngine, Pos);

		if (nDirection > 0)
		{
			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInFootnotes(oSearchEngine, null);

			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInHdrFtr(oSearchEngine, null);
		}
		else
		{
			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInHdrFtr(oSearchEngine, null);

			if (true !== oSearchEngine.IsFound())
				this.private_GetRevisionsChangeElementInFootnotes(oSearchEngine, null);
		}

		if (true !== oSearchEngine.IsFound())
			this.private_GetRevisionsChangeElementInDocument(oSearchEngine, nDirection > 0 ? 0 : this.Content.length - 1);
	}

	return oSearchEngine;
};
CDocument.prototype.private_GetRevisionsChangeElementInDocument = function(SearchEngine, Pos)
{
	var Direction = SearchEngine.GetDirection();
	this.Content[Pos].GetRevisionsChangeElement(SearchEngine);
	while (true !== SearchEngine.IsFound())
	{
		Pos = (Direction > 0 ? Pos + 1 : Pos - 1);
		if (Pos >= this.Content.length || Pos < 0)
			break;

		this.Content[Pos].GetRevisionsChangeElement(SearchEngine);
	}
};
CDocument.prototype.private_GetRevisionsChangeElementInHdrFtr = function(SearchEngine, HdrFtr)
{
	var AllHdrFtrs = this.SectionsInfo.GetAllHdrFtrs();
	var Count = AllHdrFtrs.length;

	if (Count <= 0)
		return;

	var Pos = -1;
	if (null !== HdrFtr)
	{
		for (var Index = 0; Index < Count; ++Index)
		{
			if (HdrFtr === AllHdrFtrs[Index])
			{
				Pos = Index;
				break;
			}
		}
	}

	var Direction = SearchEngine.GetDirection();
	if (Pos < 0 || Pos >= Count)
	{
		if (Direction > 0)
			Pos = 0;
		else
			Pos = Count - 1;
	}

	AllHdrFtrs[Pos].GetRevisionsChangeElement(SearchEngine);
	while (true !== SearchEngine.IsFound())
	{
		Pos = (Direction > 0 ? Pos + 1 : Pos - 1);
		if (Pos >= Count || Pos < 0)
			break;

		AllHdrFtrs[Pos].GetRevisionsChangeElement(SearchEngine);
	}
};
CDocument.prototype.private_GetRevisionsChangeElementInFootnotesList = function(SearchEngine, oFootnote, arrFootnotes)
{
	var nCount = arrFootnotes.length;
	if (nCount <= 0)
		return;

	var nPos = -1;
	if (oFootnote)
	{
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			if (arrFootnotes[nPos] === oFootnote)
			{
				nPos = nIndex;
				break;
			}
		}
	}

	var nDirection = SearchEngine.GetDirection();
	if (nPos < 0 || nPos >= nCount)
	{
		if (nDirection > 0)
			nPos = 0;
		else
			nPos = nCount - 1;
	}

	arrFootnotes[nPos].GetRevisionsChangeElement(SearchEngine);
	while (true !== SearchEngine.IsFound())
	{
		nPos = (nDirection > 0 ? nPos + 1 : nPos - 1);
		if (nPos >= nCount || nPos < 0)
			break;

		arrFootnotes[nPos].GetRevisionsChangeElement(SearchEngine);
	}
};
CDocument.prototype.private_GetRevisionsChangeElementInFootnotes = function(oSearchEngine, oFootnote)
{
	if (oSearchEngine.GetDirection() > 0)
	{
		if (true !== oSearchEngine.IsFound())
			this.private_GetRevisionsChangeElementInFootnotesList(oSearchEngine, oFootnote, this.GetFootnotesList(null, null));

		if (true !== oSearchEngine.IsFound())
			this.private_GetRevisionsChangeElementInFootnotesList(oSearchEngine, oFootnote, this.GetEndnotesList(null, null));
	}
	else
	{
		if (true !== oSearchEngine.IsFound())
			this.private_GetRevisionsChangeElementInFootnotesList(oSearchEngine, oFootnote, this.GetEndnotesList(null, null));

		if (true !== oSearchEngine.IsFound())
			this.private_GetRevisionsChangeElementInFootnotesList(oSearchEngine, oFootnote, this.GetFootnotesList(null, null));
	}
};
CDocument.prototype.private_SelectRevisionChange = function(oChange, isSkipCompleteCheck)
{
	if (oChange && oChange.get_Paragraph())
	{
		this.RemoveSelection();

		if (oChange.IsComplexChange())
		{
			if (oChange.IsMove())
			{
				this.SelectTrackMove(oChange.GetMoveId(), oChange.IsMoveFrom(), false, false);
			}
		}
		else
		{
			var oElement = oChange.get_Paragraph();

			if (true !== isSkipCompleteCheck && this.TrackRevisionsManager.CompleteTrackChangesForElements([oElement]))
				return;

			if (oElement instanceof Paragraph)
			{
				// Текущую позицию нужно выставить до селекта
				oElement.Set_ParaContentPos(oChange.get_StartPos(), false, -1, -1);
				oElement.Selection.Use = true;
				oElement.Set_SelectionContentPos(oChange.get_StartPos(), oChange.get_EndPos(), false);
				oElement.Document_SetThisElementCurrent(false);
			}
			else if (oElement instanceof CTable)
			{
				oElement.SelectRows(oChange.get_StartPos(), oChange.get_EndPos());
				oElement.Document_SetThisElementCurrent(false);
			}
		}

		this.CheckComplexFieldsInSelection();
	}
};
CDocument.prototype.AcceptRevisionChange = function(oChange)
{
	if (oChange)
	{
		var arrRelatedParas = this.TrackRevisionsManager.GetChangeRelatedParagraphs(oChange, true);

		if (this.TrackRevisionsManager.CompleteTrackChangesForElements(arrRelatedParas))
		{
			this.Document_UpdateInterfaceState();
			this.Document_UpdateSelectionState();
			return;
		}

		if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : arrRelatedParas,
			CheckType : AscCommon.changestype_Paragraph_Properties
		}))
		{
			this.StartAction(AscDFH.historydescription_Document_AcceptRevisionChange);

			var isTrackRevision = this.GetLocalTrackRevisions();
			if (false !== isTrackRevision)
				this.SetLocalTrackRevisions(false);

			if (oChange.IsComplexChange())
			{
				if (oChange.IsMove())
					this.private_ProcessMoveReview(oChange, true);
			}
			else
			{
				this.private_SelectRevisionChange(oChange);
				this.AcceptRevisionChanges(oChange.GetType(), false);
			}

			if (false !== isTrackRevision)
				this.SetLocalTrackRevisions(isTrackRevision);

			this.FinalizeAction();
		}
	}
};
CDocument.prototype.private_ProcessMoveReview = function(oChange, isAccept)
{
	var oTrackRevisionsManager = this.GetTrackRevisionsManager();

	var sMoveId     = oChange.GetMoveId();
	var isMovedDown = oChange.IsMovedDown();

	var oThis = this;

	var oTrackMove = oTrackRevisionsManager.StartProcessReviewMove(sMoveId, oChange.GetUserId());
	function privateProcessChanges(isFrom)
	{
		oTrackMove.SetFrom(isFrom);
		oThis.SelectTrackMove(sMoveId, isFrom, false, false);

		if (isAccept)
			oThis.AcceptRevisionChanges(c_oAscRevisionsChangeType.MoveMark, false);
		else
			oThis.RejectRevisionChanges(c_oAscRevisionsChangeType.MoveMark, false);
	}

	if (isMovedDown)
	{
		privateProcessChanges(false);
		privateProcessChanges(true);
	}
	else
	{
		privateProcessChanges(true);
		privateProcessChanges(false);
	}

	oTrackRevisionsManager.EndProcessReviewMove();
};
CDocument.prototype.RejectRevisionChange = function(oChange)
{
	if (oChange)
	{
		var arrRelatedParas = this.TrackRevisionsManager.GetChangeRelatedParagraphs(oChange, false);

		if (this.TrackRevisionsManager.CompleteTrackChangesForElements(arrRelatedParas))
		{
			this.Document_UpdateInterfaceState();
			this.Document_UpdateSelectionState();
			return;
		}

		if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : arrRelatedParas,
			CheckType : AscCommon.changestype_Paragraph_Properties
		}))
		{
			this.StartAction(AscDFH.historydescription_Document_RejectRevisionChange);

			var isTrackRevision = this.GetLocalTrackRevisions();
			if (false !== isTrackRevision)
				this.SetLocalTrackRevisions(false);

			if (oChange.IsComplexChange())
			{
				if (oChange.IsMove())
					this.private_ProcessMoveReview(oChange, false);
			}
			else
			{
				this.private_SelectRevisionChange(oChange);
				this.RejectRevisionChanges(oChange.GetType(), false);
			}

			if (false !== isTrackRevision)
				this.SetLocalTrackRevisions(isTrackRevision);

			this.FinalizeAction();
		}
	}
};
CDocument.prototype.AcceptRevisionChangesBySelection = function()
{
	var CurrentChange = this.TrackRevisionsManager.GetCurrentChange();
	if (null !== CurrentChange)
	{
		this.AcceptRevisionChange(CurrentChange);
	}
	else if (this.IsSelectionUse())
	{
		var sMoveId = this.CheckTrackMoveInSelection();
		if (sMoveId)
		{
			var oChange = this.TrackRevisionsManager.GetMoveMarkChange(sMoveId, true, false);
			if (oChange)
			{
				oChange = this.TrackRevisionsManager.CollectMoveChange(oChange);
				return this.AcceptRevisionChange(oChange);
			}
		}

		var SelectedParagraphs = this.GetAllParagraphs({Selected : true});
		var RelatedParas       = this.TrackRevisionsManager.Get_AllChangesRelatedParagraphsBySelectedParagraphs(SelectedParagraphs, true);
		if (SelectedParagraphs.length > 0
			&& !this.IsSelectionLocked(AscCommon.changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : RelatedParas,
			CheckType : AscCommon.changestype_Paragraph_Content
		}))
		{
			this.StartAction(AscDFH.historydescription_Document_AcceptRevisionChangesBySelection);

			var isTrackRevision = this.GetLocalTrackRevisions();
			if (false !== isTrackRevision)
				this.SetLocalTrackRevisions(false);

			this.AcceptRevisionChanges(undefined, false);

			if (false !== isTrackRevision)
				this.SetLocalTrackRevisions(isTrackRevision);

			this.FinalizeAction();
		}
	}
	else
	{
		this.private_AcceptRejectChangeByCurrentPosition(true);
	}

	this.TrackRevisionsManager.ClearCurrentChange();
};
CDocument.prototype.RejectRevisionChangesBySelection = function()
{
	var CurrentChange = this.TrackRevisionsManager.GetCurrentChange();
	if (null !== CurrentChange)
	{
		this.RejectRevisionChange(CurrentChange);
	}
	else if (this.IsSelectionUse())
	{
		var sMoveId = this.CheckTrackMoveInSelection();
		if (sMoveId)
		{
			var oChange = this.TrackRevisionsManager.GetMoveMarkChange(sMoveId, true, false);
			if (oChange)
			{
				oChange = this.TrackRevisionsManager.CollectMoveChange(oChange);
				return this.RejectRevisionChange(oChange);
			}
		}

		var SelectedParagraphs = this.GetAllParagraphs({Selected : true});
		var RelatedParas       = this.TrackRevisionsManager.Get_AllChangesRelatedParagraphsBySelectedParagraphs(SelectedParagraphs, false);
		if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : RelatedParas,
			CheckType : AscCommon.changestype_Paragraph_Content
		}))
		{
			this.StartAction(AscDFH.historydescription_Document_RejectRevisionChangesBySelection);

			var isTrackRevision = this.GetLocalTrackRevisions();
			if (false !== isTrackRevision)
				this.SetLocalTrackRevisions(false);

			this.RejectRevisionChanges(undefined, false);

			if (false !== isTrackRevision)
				this.SetLocalTrackRevisions(isTrackRevision);

			this.FinalizeAction();
		}
	}
	else
	{
		this.private_AcceptRejectChangeByCurrentPosition(false);
	}

	this.TrackRevisionsManager.ClearCurrentChange();
};
CDocument.prototype.private_AcceptRejectChangeByCurrentPosition = function(isAccept)
{
	this.TrackRevisionsManager.ContinueTrackRevisions();

	let currentChanges = this.TrackRevisionsManager.GetSelectedChanges();

	let change = null;
	for (let index = 0, count = currentChanges.length; index < count; ++index)
	{
		if (!change || change.GetWeight() < currentChanges[index].GetWeight())
			change = currentChanges[index];
	}

	let state = this.SaveDocumentState(false);

	if (change)
	{
		if (isAccept)
			this.AcceptRevisionChange(change);
		else
			this.RejectRevisionChange(change);
	}

	this.LoadDocumentState(state);
};
CDocument.prototype.AcceptAllRevisionChanges = function(isSkipCheckLock, isCheckEmptyAction)
{
	var _isCheckEmptyAction = (false !== isCheckEmptyAction);

	var RelatedParas = this.TrackRevisionsManager.Get_AllChangesRelatedParagraphs(true);
	if (true === isSkipCheckLock || false === this.IsSelectionLocked(AscCommon.changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : RelatedParas,
			CheckType : AscCommon.changestype_Paragraph_Properties
		}))
	{
		this.StartAction(AscDFH.historydescription_Document_AcceptAllRevisionChanges);

		var isTrackRevision = this.GetLocalTrackRevisions();
		if (false !== isTrackRevision)
			this.SetLocalTrackRevisions(false);

		var LogicDocuments = this.TrackRevisionsManager.Get_AllChangesLogicDocuments();
		for (var LogicDocId in LogicDocuments)
		{
			var LogicDoc = AscCommon.g_oTableId.Get_ById(LogicDocId);
			if (LogicDoc)
			{
				LogicDoc.AcceptRevisionChanges(undefined, true);
			}
		}

		if (false !== isTrackRevision)
			this.SetLocalTrackRevisions(isTrackRevision);

		if (true !== isSkipCheckLock && true === this.History.Is_LastPointEmpty())
		{
			this.FinalizeAction(_isCheckEmptyAction);
			return;
		}

		this.RemoveSelection();
		this.private_CorrectDocumentPosition();
		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction(_isCheckEmptyAction);
	}
};
CDocument.prototype.RejectAllRevisionChanges = function(isSkipCheckLock, isCheckEmptyAction)
{
	var _isCheckEmptyAction = (false !== isCheckEmptyAction);

	var RelatedParas = this.TrackRevisionsManager.Get_AllChangesRelatedParagraphs(false);
	if (true === isSkipCheckLock || false === this.IsSelectionLocked(AscCommon.changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : RelatedParas,
			CheckType : AscCommon.changestype_Paragraph_Properties
		}))
	{
		this.StartAction(AscDFH.historydescription_Document_RejectAllRevisionChanges);

		var isTrackRevision = this.GetLocalTrackRevisions();
		if (false !== isTrackRevision)
			this.SetLocalTrackRevisions(false);

		this.private_RejectAllRevisionChanges();

		if (false !== isTrackRevision)
			this.SetLocalTrackRevisions(isTrackRevision);

		if (true !== isSkipCheckLock && true === this.History.Is_LastPointEmpty())
		{
			this.FinalizeAction(_isCheckEmptyAction);
			return;
		}

		this.RemoveSelection();
		this.private_CorrectDocumentPosition();
		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction(_isCheckEmptyAction);
	}
};
CDocument.prototype.private_RejectAllRevisionChanges = function()
{
	var LogicDocuments = this.TrackRevisionsManager.Get_AllChangesLogicDocuments();
	for (var LogicDocId in LogicDocuments)
	{
		var LogicDoc = this.TableId.Get_ById(LogicDocId);
		if (LogicDoc)
		{
			LogicDoc.RejectRevisionChanges(undefined, true);
		}
	}
};
CDocument.prototype.AcceptRevisionChanges = function(nType, bAll)
{
	// Принимаем все изменения, которые попали в селект.
	// Принимаем изменения в следующей последовательности:
	// 1. Изменение настроек параграфа
	// 2. Изменение настроек текста
	// 3. Добавление/удаление текста
	// 4. Добавление/удаление параграфа
	if (docpostype_Content === this.CurPos.Type || true === bAll)
	{
		this.private_AcceptRevisionChanges(nType, bAll);
	}
	else if (docpostype_HdrFtr === this.CurPos.Type)
	{
		this.HdrFtr.AcceptRevisionChanges(nType, bAll);
	}
	else if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.DrawingObjects.AcceptRevisionChanges(nType, bAll);
	}
	else if (docpostype_Footnotes === this.CurPos.Type)
	{
		this.Footnotes.AcceptRevisionChanges(nType, bAll);
	}
	else if (docpostype_Endnotes === this.CurPos.Type)
	{
		this.Endnotes.AcceptRevisionChanges(nType, bAll);
	}


	if (true !== bAll)
	{
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
	}
};
CDocument.prototype.RejectRevisionChanges = function(nType, bAll)
{
	// Отменяем все изменения, которые попали в селект.
	// Отменяем изменения в следующей последовательности:
	// 1. Изменение настроек параграфа
	// 2. Изменение настроек текста
	// 3. Добавление/удаление текста
	// 4. Добавление/удаление параграфа

	if (docpostype_Content === this.CurPos.Type || true === bAll)
	{
		this.private_RejectRevisionChanges(nType, bAll);
	}
	else if (docpostype_HdrFtr === this.CurPos.Type)
	{
		this.HdrFtr.RejectRevisionChanges(nType, bAll);
	}
	else if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.DrawingObjects.RejectRevisionChanges(nType, bAll);
	}
	else if (docpostype_Footnotes === this.CurPos.Type)
	{
		this.Footnotes.RejectRevisionChanges(nType, bAll);
	}
	else if (docpostype_Endnotes === this.CurPos.Type)
	{
		this.Endnotes.RejectRevisionChanges(nType, bAll);
	}

	if (true !== bAll)
	{
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
	}
};
CDocument.prototype.HaveRevisionChanges = function(isCheckOwnChanges)
{
	this.TrackRevisionsManager.ContinueTrackRevisions();

	if (true === isCheckOwnChanges)
		return this.TrackRevisionsManager.Have_Changes();
	else
		return this.TrackRevisionsManager.HaveOtherUsersChanges();
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с составным вводом
//----------------------------------------------------------------------------------------------------------------------
/**
 * Сообщаем о начале составного ввода текста.
 * @returns {boolean} Начался или нет составной ввод.
 */
CDocument.prototype.Begin_CompositeInput = function()
{
	var bResult = false;
	if (false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content, null, true, this.IsFormFieldEditing()))
	{
		this.StartAction(AscDFH.historydescription_Document_CompositeInput);
		this.DrawingObjects.CreateDocContent();
		this.DrawingDocument.TargetStart();
		this.DrawingDocument.TargetShow();

		if (true === this.IsSelectionUse())
		{
			if (docpostype_DrawingObjects === this.GetDocPosType() && null === this.DrawingObjects.getTargetDocContent())
				this.RemoveSelection();
			else
				this.Remove(1, true, false, true);
		}

		var oPara = this.GetCurrentParagraph();
		if (oPara)
		{
			var oRun = oPara.Get_ElementByPos(oPara.Get_ParaContentPos(false, false));

			if (oRun instanceof ParaRun)
			{
				var oRunParent = oRun.GetParent();
				if (oRunParent instanceof CInlineLevelSdt && oRunParent.IsPlaceHolder())
				{
					oRunParent.ReplacePlaceHolderWithContent(false);
					oRun = oRunParent.GetElement(0);
				}
			}

			if (oRun instanceof ParaRun)
			{
				let oNewRun = null;
				let isCheck = false;
				if (!oRun.GetParentForm())
				{
					isCheck = true;
					oNewRun = oRun.CheckRunBeforeAdd();
					if (!oNewRun)
						oNewRun = oRun.private_SplitRunInCurPos();
				}

				let prevRun = null;
				if (oNewRun && oNewRun !== oRun)
				{
					prevRun = oRun;
					oRun = oNewRun;
					oRun.Make_ThisElementCurrent();
				}

				this.CompositeInput = {
					Run     : oRun,
					Pos     : oRun.State.ContentPos,
					Length  : 0,
					CanUndo : true,
					Check   : isCheck,
					PrevRun : prevRun
				};

				oRun.Set_CompositeInput(this.CompositeInput);

				bResult = true;
			}
		}

		this.FinalizeAction(false);
	}

	return bResult;
};
CDocument.prototype.CheckCurrentTextObjectExtends = function()
{
	var oController = this.DrawingObjects;
	if (oController)
	{
		oController.checkCurrentTextObjectExtends();
	}
};
CDocument.prototype.Replace_CompositeText = function(arrCharCodes)
{
	if (null === this.CompositeInput)
		return;

	arrCharCodes = typeof(arrCharCodes) === "string" ? arrCharCodes.codePointsArray() : arrCharCodes;

	if (!this.CompositeInput.Length && !arrCharCodes.length)
		return;

	this.StartAction(AscDFH.historydescription_Document_CompositeInputReplace);

	if (this.CompositeInput.Check && arrCharCodes.length)
	{
		let prevRun = this.CompositeInput.PrevRun;
		let isCS = AscCommon.IsComplexScript(arrCharCodes[0]);

		this.CompositeInput.Run.ApplyComplexScript(isCS);
		if (prevRun
			&& isCS !== prevRun.IsCS()
			&& prevRun.IsOnlyCommonTextScript())
		{
			prevRun.ApplyComplexScript(isCS);
		}

		this.CompositeInput.Check = false;
	}

	this.Start_SilentMode();
	this.private_RemoveCompositeText(this.CompositeInput.Length, true);
	for (var nIndex = 0, nCount = arrCharCodes.length; nIndex < nCount; ++nIndex)
	{
		this.private_AddCompositeText(arrCharCodes[nIndex], true);
	}

	this.End_SilentMode(false);

	let oRun  = this.CompositeInput.Run;
	let oForm = oRun.GetParentForm();
	if (oForm)
	{
		oForm.TrimTextForm();

		if (oRun.IsEmpty())
		{
			oForm.ReplaceContentWithPlaceHolder();
			AscCommon.g_inputContext.externalEndCompositeInput();
		}
	}

	this.CheckCurrentTextObjectExtends();
	this.Recalculate();
	this.UpdateSelection();
	this.UpdateUndoRedo();
	this.FinalizeAction(false);

	this.private_UpdateCursorXY(true, true);

	if (!this.CompositeInput)
		return;

	if (!this.History.CheckUnionLastPoints())
		this.CompositeInput.CanUndo = false;
};
CDocument.prototype.Set_CursorPosInCompositeText = function(nPos)
{
	if (null === this.CompositeInput)
		return;

	var oRun = this.CompositeInput.Run;

	var nInRunPos = Math.max(Math.min(this.CompositeInput.Pos + nPos, this.CompositeInput.Pos + this.CompositeInput.Length, oRun.Content.length), this.CompositeInput.Pos);
	oRun.State.ContentPos = nInRunPos;
	this.Document_UpdateSelectionState();
};
CDocument.prototype.Get_CursorPosInCompositeText = function()
{
	if (null === this.CompositeInput)
		return 0;

	var oRun = this.CompositeInput.Run;
	var nInRunPos = oRun.State.ContentPos;
	var nPos = Math.min(this.CompositeInput.Length, Math.max(0, nInRunPos - this.CompositeInput.Pos));
	return nPos;
};
CDocument.prototype.End_CompositeInput = function()
{
	if (null === this.CompositeInput)
		return;

	var nLen = this.CompositeInput.Length;

	var oRun = this.CompositeInput.Run;
	oRun.Set_CompositeInput(null);

	let oParentForm;
	if ((0 === nLen && this.CompositeInput.CanUndo)
		|| ((oParentForm = oRun.GetParentForm()) && !this.FormsManager.ValidateChangeOnFly(oParentForm)))
	{
		let arrChanges = this.History.UndoCompositeInput();
		if (arrChanges)
		{
			this.History.ClearRedo();
			this.UpdateAfterUndoRedo(arrChanges);
		}
	}

	this.CompositeInput = null;

	var oController = this.DrawingObjects;
    if(oController)
    {
        var oTargetTextObject = AscFormat.getTargetTextObject(oController);
        if(oTargetTextObject && oTargetTextObject.txWarpStructNoTransform)
        {
            oTargetTextObject.recalcInfo.recalculateTxBoxContent = true;
            oTargetTextObject.recalculateText();
        }
    }

	// Обновление интерфейса здесь обязательно, т.к. на нем должно сработать Api.CheckChangedDocument
	this.UpdateInterface();

    this.private_UpdateCursorXY(true, true);

	this.DrawingDocument.ClearCachePages();
	this.DrawingDocument.FirePaint();
};
CDocument.prototype.Get_MaxCursorPosInCompositeText = function()
{
	if (null === this.CompositeInput)
		return 0;

	return this.CompositeInput.Length;
};
CDocument.prototype.private_AddCompositeText = function(nCharCode, bSkipCheckExtents)
{
	var oRun = this.CompositeInput.Run;
	var nPos = this.CompositeInput.Pos + this.CompositeInput.Length;
	var oChar;

	if (para_Math_Run === oRun.Type)
	{
		oChar = new CMathText();
		oChar.add(nCharCode);
	}
	else
	{
		if (AscCommon.IsSpace(nCharCode))
			oChar = new AscWord.CRunSpace(nCharCode);
		else
			oChar = new AscWord.CRunText(nCharCode);
	}

	oRun.AddToContent(nPos, oChar, true);
	this.CompositeInput.Length++;

	var oForm = oRun.GetParentForm();
	if (oForm && oForm.IsAutoFitContent())
		this.CheckFormAutoFit(oForm);
    if (!bSkipCheckExtents)
    {
        this.CheckCurrentTextObjectExtends();
    }
	this.Recalculate();
	this.UpdateSelection();
};
CDocument.prototype.private_RemoveCompositeText = function(nCount, bSkipCheckExtents)
{
	var oRun = this.CompositeInput.Run;
	var nPos = this.CompositeInput.Pos + this.CompositeInput.Length;

	var nDelCount = Math.max(0, Math.min(nCount, this.CompositeInput.Length, oRun.Content.length, nPos));
	oRun.Remove_FromContent(nPos - nDelCount, nDelCount, true);
	this.CompositeInput.Length -= nDelCount;

	var oForm = oRun.GetParentForm();
	if (oForm && oForm.IsAutoFitContent())
		this.CheckFormAutoFit(oForm);

    if (!bSkipCheckExtents)
    {
        this.CheckCurrentTextObjectExtends();
    }
	this.Recalculate();
	this.UpdateSelection();
};
CDocument.prototype.Check_CompositeInputRun = function()
{
	if (null === this.CompositeInput)
		return;

	var oRun = this.CompositeInput.Run;
	if (true !== oRun.IsUseInDocument())
		AscCommon.g_inputContext.externalEndCompositeInput();
};
CDocument.prototype.Is_CursorInsideCompositeText = function()
{
	if (null === this.CompositeInput)
		return false;

	var oCurrentParagraph = this.GetCurrentParagraph();
	if (!oCurrentParagraph)
		return false;

	var oParaPos   = oCurrentParagraph.Get_ParaContentPos(false, false, false);
	var arrClasses = oCurrentParagraph.Get_ClassesByPos(oParaPos);

	if (arrClasses.length <= 0 || arrClasses[arrClasses.length - 1] !== this.CompositeInput.Run)
		return false;

	var nInRunPos = oParaPos.Get(oParaPos.GetDepth());
	if (nInRunPos >= this.CompositeInput.Pos && nInRunPos <= this.CompositeInput.Pos + this.CompositeInput.Length)
		return true;

	return false;
};
CDocument.prototype.IsCompositeInputInProgress = function()
{
	return (!!this.CompositeInput);
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы со сносками
//----------------------------------------------------------------------------------------------------------------------
/**
 * Переходим к редактированию сносок на заданной странице. Если сносок на заданной странице нет, тогда ничего не делаем.
 * @param {number} nPageIndex
 * @returns {boolean}
 */
CDocument.prototype.GotoFootnotesOnPage = function(nPageIndex)
{
	if (this.Footnotes.IsEmptyPage(nPageIndex))
		return false;

	this.SetDocPosType(docpostype_Footnotes);
	this.Footnotes.GotoPage(nPageIndex);
	this.Document_UpdateSelectionState();

	return true;
};
CDocument.prototype.AddFootnote = function(sText)
{
	var nDocPosType = this.GetDocPosType();
	if (docpostype_Content !== nDocPosType && docpostype_Footnotes !== nDocPosType)
		return null;

	var oInfo = this.GetSelectedElementsInfo();
	if (oInfo.GetMath())
		return null;

	if (false === this.Document_Is_SelectionLocked(changestype_Paragraph_Content))
	{
		let oFootnote = null;
		this.StartAction(AscDFH.historydescription_Document_AddFootnote);

		if (docpostype_Content === nDocPosType)
		{
			oFootnote = this.Footnotes.CreateFootnote();
			oFootnote.AddDefaultFootnoteContent(sText);
			if (true === this.IsSelectionUse())
			{
				this.MoveCursorRight(false, false, false);
				this.RemoveSelection();
			}

			if (sText)
				this.AddToParagraph(new AscWord.CRunFootnoteReference(oFootnote, sText));
			else
				this.AddToParagraph(new AscWord.CRunFootnoteReference(oFootnote));

			this.SetDocPosType(docpostype_Footnotes);
			this.Footnotes.Set_CurrentElement(true, 0, oFootnote);
		}
		else if (docpostype_Footnotes === nDocPosType)
		{
			this.Footnotes.AddFootnoteRef();
			this.Recalculate();
		}
		this.FinalizeAction();

		return oFootnote;
	}
	
	return null;
};
CDocument.prototype.RemoveAllFootnotes = function(bRemoveFootnotes, bRemoveEndnotes)
{
	var nDocPosType = this.GetDocPosType();

	var oEngine = new CDocumentFootnotesRangeEngine(true);
	oEngine.Init(null, null, bRemoveFootnotes, bRemoveEndnotes);

	var arrParagraphs = this.GetAllParagraphs({OnlyMainDocument : true, All : true});
	for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
	{
		arrParagraphs[nIndex].GetFootnotesList(oEngine);
	}

	var arrFootnotes  = oEngine.GetRange();
	var arrParagraphs = oEngine.GetParagraphs();
	var arrRuns       = oEngine.GetRuns();
	var arrRefs       = oEngine.GetRefs();

	if (arrRuns.length !== arrRefs.length || arrFootnotes.length !== arrRuns.length)
		return;

	if (false === this.Document_Is_SelectionLocked(changestype_None, { Type : changestype_2_ElementsArray_and_Type, Elements : arrParagraphs, CheckType : changestype_Paragraph_Content }, true))
	{
		this.StartAction(AscDFH.historydescription_Document_RemoveAllFootnotes);

		var sDefaultStyleId = this.GetStyles().GetDefaultFootnoteReference();
		for (var nIndex = 0, nCount = arrFootnotes.length; nIndex < nCount; ++nIndex)
		{
			var oRef = arrRefs[nIndex];
			var oRun = arrRuns[nIndex];

			oRun.RemoveElement(oRef);

			if (oRun.GetRStyle() === sDefaultStyleId)
				oRun.SetRStyle(undefined);
		}

		this.Recalculate();

		if (docpostype_Footnotes === nDocPosType)
		{
			this.CurPos.Type = docpostype_Content;
			if (arrRuns.length > 0)
				arrRuns[0].Make_ThisElementCurrent();
		}
		this.FinalizeAction();
	}
};
CDocument.prototype.GotoFootnote = function(isNext)
{
	var nDocPosType = this.GetDocPosType();

	if (docpostype_Footnotes === nDocPosType)
	{
		if (isNext)
			this.Footnotes.GotoNextFootnote();
		else
			this.Footnotes.GotoPrevFootnote();

		this.Document_UpdateSelectionState();
		this.Document_UpdateInterfaceState();

		return;
	}

	if (docpostype_HdrFtr === nDocPosType)
	{
		this.EndHdrFtrEditing(true);
	}
	else if (docpostype_Endnotes === nDocPosType)
	{
		this.EndEndnotesEditing();
	}
	else if (docpostype_DrawingObjects === nDocPosType)
	{
		this.DrawingObjects.resetSelection2();
		this.private_UpdateCursorXY(true, true);

		this.CurPos.Type = docpostype_Content;

		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
	}

	return this.GotoFootnoteRef(isNext, true, true, false);
};
/**
 * @return {CFootnotesController}
 */
CDocument.prototype.GetFootnotesController = function()
{
	return this.Footnotes;
};
/**
 * @return {CEndnotesController}
 */
CDocument.prototype.GetEndnotesController = function()
{
	return this.Endnotes;
};
CDocument.prototype.SetFootnotePr = function(oFootnotePr, bApplyToAll)
{
	var nNumStart   = oFootnotePr.get_NumStart();
	var nNumRestart = oFootnotePr.get_NumRestart();
	var nNumFormat  = oFootnotePr.get_NumFormat();
	var nPos        = oFootnotePr.get_Pos();
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Document_SectPr))
	{
		this.StartAction(AscDFH.historydescription_Document_SetFootnotePr);

		if (bApplyToAll)
		{
			for (var nIndex = 0, nCount = this.SectionsInfo.GetSectionsCount(); nIndex < nCount; ++nIndex)
			{
				var oSectPr = this.SectionsInfo.Get_SectPr2(nIndex).SectPr;
				if (undefined !== nNumStart)
					oSectPr.SetFootnoteNumStart(nNumStart);

				if (undefined !== nNumRestart)
					oSectPr.SetFootnoteNumRestart(nNumRestart);

				if (undefined !== nNumFormat)
					oSectPr.SetFootnoteNumFormat(nNumFormat);

				if (undefined !== nPos)
					oSectPr.SetFootnotePos(nPos);
			}
		}
		else
		{
			var oSectPr = this.GetCurrentSectionPr();
			if (undefined !== nNumStart)
				oSectPr.SetFootnoteNumStart(nNumStart);

			if (undefined !== nNumRestart)
				oSectPr.SetFootnoteNumRestart(nNumRestart);

			if (undefined !== nNumFormat)
				oSectPr.SetFootnoteNumFormat(nNumFormat);

			if (undefined !== nPos)
				oSectPr.SetFootnotePos(nPos);
		}

		this.Recalculate();
		this.FinalizeAction();
	}
};
CDocument.prototype.GetFootnotePr = function()
{
	var oSectPr     = this.GetCurrentSectionPr();
	var oFootnotePr = new Asc.CAscFootnotePr();
	oFootnotePr.put_Pos(oSectPr.GetFootnotePos());
	oFootnotePr.put_NumStart(oSectPr.GetFootnoteNumStart());
	oFootnotePr.put_NumRestart(oSectPr.GetFootnoteNumRestart());
	oFootnotePr.put_NumFormat(oSectPr.GetFootnoteNumFormat());
	return oFootnotePr;
};
CDocument.prototype.IsCursorInFootnote = function()
{
	return (docpostype_Footnotes === this.GetDocPosType());
};
CDocument.prototype.GetFootnotesList = function(oFirstFootnote, oLastFootnote)
{
	// TODO: Возможно стоит пересмотреть схему составления списка на следующую
	//       У сносок добавить Prev/Next и обновлять при добавлении/удалении сносок

	if (null === oFirstFootnote && null === oLastFootnote && null !== this.AllFootnotesList)
		return this.AllFootnotesList;

	var arrFootnotes = CDocumentContentBase.prototype.GetFootnotesList.call(this, oFirstFootnote, oLastFootnote, false);

	if (null === oFirstFootnote && null === oLastFootnote)
		this.AllFootnotesList = arrFootnotes;

	return arrFootnotes;
};
CDocument.prototype.GetEndnotesList = function(oFirstEndnote, oLastEndnote)
{
	// TODO: Возможно стоит пересмотреть схему составления списка на следующую
	//       У сносок добавить Prev/Next и обновлять при добавлении/удалении сносок

	if (null === oFirstEndnote && null === oLastEndnote && null !== this.AllEndnotesList)
		return this.AllEndnotesList;

	var arrEndnotes = CDocumentContentBase.prototype.GetFootnotesList.call(this, oFirstEndnote, oLastEndnote, true);

	if (null === oFirstEndnote && null === oLastEndnote)
		this.AllEndnotesList = arrEndnotes;

	return arrEndnotes;
};
CDocument.prototype.AddEndnote = function(sText)
{
	var nDocPosType = this.GetDocPosType();
	if (docpostype_Content !== nDocPosType && docpostype_Endnotes !== nDocPosType)
		return null;

	var oInfo = this.GetSelectedElementsInfo();
	if (oInfo.GetMath())
		return null;

	if (!this.IsSelectionLocked(changestype_Paragraph_Content))
	{
		let oEndnote = null;
		this.StartAction(AscDFH.historydescription_Document_AddEndnote);

		if (docpostype_Content === nDocPosType)
		{
			oEndnote = this.Endnotes.CreateEndnote();
			oEndnote.AddDefaultEndnoteContent(sText);

			if (true === this.IsSelectionUse())
			{
				this.MoveCursorRight(false, false, false);
				this.RemoveSelection();
			}

			if (sText)
				this.AddToParagraph(new AscWord.CRunEndnoteReference(oEndnote, sText));
			else
				this.AddToParagraph(new AscWord.CRunEndnoteReference(oEndnote));

			this.SetDocPosType(docpostype_Endnotes);
			this.Endnotes.Set_CurrentElement(true, 0, oEndnote);
		}
		else if (docpostype_Endnotes === nDocPosType)
		{
			this.Endnotes.AddEndnoteRef();
		}

		this.Recalculate();
		this.FinalizeAction();

		return oEndnote;
	}

	return null;
};
CDocument.prototype.GotoEndnote = function(isNext)
{
	var nDocPosType = this.GetDocPosType();

	if (docpostype_Endnotes === nDocPosType)
	{
		if (isNext)
			this.Endnotes.GotoNextEndnote();
		else
			this.Endnotes.GotoPrevEndnote();

		this.UpdateSelection();
		this.UpdateInterface();

		return;
	}

	if (docpostype_HdrFtr === nDocPosType)
	{
		this.EndHdrFtrEditing(true);
	}
	else if (docpostype_Footnotes === nDocPosType)
	{
		this.EndFootnotesEditing();
	}
	else if (docpostype_DrawingObjects === nDocPosType)
	{
		this.DrawingObjects.resetSelection2();
		this.private_UpdateCursorXY(true, true);

		this.CurPos.Type = docpostype_Content;

		this.Document_UpdateInterfaceState();
		this.Document_UpdateSelectionState();
	}

	return this.GotoFootnoteRef(isNext, true, false, true);
};
CDocument.prototype.IsCursorInEndnote = function()
{
	return (docpostype_Endnotes === this.GetDocPosType());
};
CDocument.prototype.GetEndnotePr = function()
{
	var oSectPr    = this.GetCurrentSectionPr();
	var oEndnotePr = new Asc.CAscFootnotePr();
	oEndnotePr.put_Pos(this.Endnotes.GetEndnotePrPos());
	oEndnotePr.put_NumStart(oSectPr.GetEndnoteNumStart());
	oEndnotePr.put_NumRestart(oSectPr.GetEndnoteNumRestart());
	oEndnotePr.put_NumFormat(oSectPr.GetEndnoteNumFormat());
	return oEndnotePr;
};
CDocument.prototype.SetEndnotePr = function(oEndnotePr, bApplyToAll)
{
	var nNumStart   = oEndnotePr.get_NumStart();
	var nNumRestart = oEndnotePr.get_NumRestart();
	var nNumFormat  = oEndnotePr.get_NumFormat();
	var nPos        = oEndnotePr.get_Pos();

	if (!this.IsSelectionLocked(AscCommon.changestype_Document_SectPr))
	{
		this.StartAction(AscDFH.historydescription_Document_SetEndnotePr);

		if (undefined !== nPos)
			this.Endnotes.SetEndnotePrPos(nPos);

		if (bApplyToAll)
		{
			for (var nIndex = 0, nCount = this.SectionsInfo.GetSectionsCount(); nIndex < nCount; ++nIndex)
			{
				var oSectPr = this.SectionsInfo.Get_SectPr2(nIndex).SectPr;
				if (undefined !== nNumStart)
					oSectPr.SetEndnoteNumStart(nNumStart);

				if (undefined !== nNumRestart)
					oSectPr.SetEndnoteNumRestart(nNumRestart);

				if (undefined !== nNumFormat)
					oSectPr.SetEndnoteNumFormat(nNumFormat);
			}
		}
		else
		{
			var oSectPr = this.GetCurrentSectionPr();
			if (undefined !== nNumStart)
				oSectPr.SetEndnoteNumStart(nNumStart);

			if (undefined !== nNumRestart)
				oSectPr.SetEndnoteNumRestart(nNumRestart);

			if (undefined !== nNumFormat)
				oSectPr.SetEndnoteNumFormat(nNumFormat);
		}

		this.Recalculate();
		this.FinalizeAction();
	}
};
CDocument.prototype.ConvertFootnoteType = function(isCurrent, isFootnotes, isEndnotes)
{
	var nDocPosType  = this.GetDocPosType();
	var oCurFootnote = null;

	if (docpostype_Footnotes === nDocPosType || docpostype_Endnotes === nDocPosType)
		oCurFootnote = nDocPosType === docpostype_Footnotes ? this.Footnotes.GetCurFootnote() : this.Endnotes.GetCurEndnote();

	var arrFootnotes  = [];
	var arrParagraphs = [];
	var arrRuns       = [];
	var arrRefs       = [];

	if (isCurrent)
	{
		var nDocPosType = this.GetDocPosType();

		if (docpostype_Footnotes === nDocPosType || docpostype_Endnotes === nDocPosType)
		{
			var oFootnote = oCurFootnote;
			var oRef = oFootnote.GetRef();
			if (!oRef || !oRef.GetRun() || !oRef.GetRun().GetParagraph())
				return;

			arrFootnotes  = [oFootnote];
			arrParagraphs = [oRef.GetRun().GetParagraph()];
			arrRuns       = [oRef.GetRun()];
			arrRefs       = [oRef];
		}
		else
		{
			if (docpostype_Content !== nDocPosType || this.IsSelectionUse())
				return;

			var oParagraph = this.GetCurrentParagraph();
			if (!oParagraph)
				return;

			var oNextElement = oParagraph.GetNextRunElement();
			var oPrevElement = oParagraph.GetPrevRunElement();

			var oRef = null;
			if (oNextElement && (para_FootnoteReference === oNextElement.Type || para_EndnoteReference === oNextElement.Type))
				oRef = oNextElement;
			else if (oPrevElement && (para_FootnoteReference === oPrevElement.Type || para_EndnoteReference === oPrevElement.Type))
				oRef = oPrevElement;

			if (!oRef)
				return;

			var oRunInfo = oParagraph.GetRunByElement(oRef);
			if (!oRunInfo)
				return;

			arrFootnotes  = [oRef.GetFootnote()];
			arrParagraphs = [oParagraph];
			arrRuns       = [oRunInfo.Run];
			arrRefs       = [oRef];
		}
	}
	else
	{
		var oEngine = new CDocumentFootnotesRangeEngine(true);
		oEngine.Init(null, null, isFootnotes, isEndnotes);

		var arrParagraphs = this.GetAllParagraphs({OnlyMainDocument : true, All : true});
		for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
		{
			arrParagraphs[nIndex].GetFootnotesList(oEngine);
		}

		arrFootnotes  = oEngine.GetRange();
		arrParagraphs = oEngine.GetParagraphs();
		arrRuns       = oEngine.GetRuns();
		arrRefs       = oEngine.GetRefs();
	}


	if (arrRuns.length !== arrRefs.length || arrFootnotes.length !== arrRuns.length)
		return;

	if (false === this.Document_Is_SelectionLocked(changestype_None, { Type : changestype_2_ElementsArray_and_Type, Elements : arrParagraphs, CheckType : changestype_Paragraph_Content }, true))
	{
		this.StartAction(AscDFH.historydescription_Document_ConvertFootnoteType);

		for (var nIndex = 0, nCount = arrFootnotes.length; nIndex < nCount; ++nIndex)
		{
			var oRef = arrRefs[nIndex];
			var oRun = arrRuns[nIndex];

			var oFootnote = oRef.GetFootnote();
			if (para_FootnoteReference === oRef.Type)
			{
				this.Footnotes.RemoveFootnote(oFootnote);
				this.Endnotes.AddEndnote(oFootnote);
				oFootnote.ConvertFootnoteType(false);
				oRun.ConvertFootnoteType(false, this.GetStyles(), oFootnote, oRef);
			}
			else if (para_EndnoteReference === oRef.Type)
			{
				this.Endnotes.RemoveEndnote(oFootnote);
				this.Footnotes.AddFootnote(oFootnote);
				oFootnote.ConvertFootnoteType(true);
				oRun.ConvertFootnoteType(true, this.GetStyles(), oFootnote, oRef);
			}
		}

		if (oCurFootnote)
			oCurFootnote.SetThisElementCurrent(false);

		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
CDocument.prototype.TurnOffCheckChartSelection = function()
{
	if (this.DrawingObjects)
	{
		this.DrawingObjects.TurnOffCheckChartSelection();
	}
};
CDocument.prototype.TurnOnCheckChartSelection = function()
{
	if (this.DrawingObjects)
	{
		this.DrawingObjects.TurnOnCheckChartSelection();
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Функции, которые вызываются из CLogicDocumentController
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.controller_CanUpdateTarget = function()
{
	var nPos = this.private_GetSelectionPos(true).End;

	if (null != this.FullRecalc.Id && this.FullRecalc.StartIndex < this.CurPos.ContentPos)
	{
		return false;
	}
	else if (null !== this.FullRecalc.Id && this.FullRecalc.StartIndex === nPos)
	{
		var oElement     = this.Content[nPos];
		var nElementPage = this.private_GetElementPageIndex(nPos, this.FullRecalc.PageIndex, this.FullRecalc.ColumnIndex, oElement.Get_ColumnsCount());
		return oElement.CanUpdateTarget(nElementPage);
	}

	return true;
};
CDocument.prototype.controller_RecalculateCurPos = function(bUpdateX, bUpdateY, isUpdateTarget)
{
	if (this.controller_CanUpdateTarget())
	{
		this.private_CheckCurPage();
		var nPos = this.private_GetSelectionPos(true).End;
		return this.Content[nPos].RecalculateCurPos(bUpdateX, bUpdateY, isUpdateTarget);
	}

	return {X : 0, Y : 0, Height : 0, PageNum : 0, Internal : {Line : 0, Page : 0, Range : 0}, Transform : null};
};
CDocument.prototype.controller_GetCurPage = function()
{
	var nPos = this.private_GetSelectionPos(true).End;
	if (nPos >= 0 && (null === this.FullRecalc.Id || this.FullRecalc.StartIndex > nPos))
		return this.Content[nPos].Get_CurrentPage_Absolute();

	return -1;
};
CDocument.prototype.controller_AddNewParagraph = function(bRecalculate, bForceAdd)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	// Сначала удаляем заселекченую часть
	if (true === this.Selection.Use)
	{
		this.Remove(1, true, false, true);
	}

	// Добавляем новый параграф
	var Item = this.Content[this.CurPos.ContentPos];

	// Если мы внутри параграфа, тогда:
	// 1. Если мы в середине параграфа, разделяем данный параграф на 2.
	//    При этом полностью копируем все настройки из исходного параграфа.
	// 2. Если мы в конце данного параграфа, тогда добавляем новый пустой параграф.
	//    Стиль у него проставляем такой какой указан у текущего в Style.Next.
	//    Если при этом у нового параграфа стиль будет такой же как и у старого,
	//    в том числе если стиля нет у обоих, тогда копируем еще все прямые настройки.
	//    (Т.е. если стили разные, а у исходный параграф был параграфом со списком, тогда
	//    новый параграф будет без списка). Кроме этого у предыдущего параграфа пытаемся
	//    выполнить автозамены
	if (Item.IsParagraph())
	{
		var isCheckAutoCorrect = false;
		let numPr = Item.GetNumPr();

		// Если текущий параграф пустой и с нумерацией, тогда удаляем нумерацию и отступы левый и первой строки
		if (true !== bForceAdd && numPr && true === Item.IsEmpty({SkipNewLine : true}) && true === Item.IsCursorAtBegin())
		{
			if (numPr.Lvl <= 0)
			{
				Item.RemoveNumPr();
				Item.Set_Ind({FirstLine : undefined, Left : undefined, Right : Item.Pr.Ind.Right}, true);
			}
			else
			{
				Item.SetNumPr(numPr.NumId, numPr.Lvl - 1);
			}
		}
		else
		{
			var ItemReviewType = Item.GetReviewType();
			// Создаем новый параграф
			var NewParagraph   = new Paragraph(this.DrawingDocument, this);

			if (Item.IsCursorAtBegin())
			{
				// Продолжаем (в плане настроек) новый параграф
				Item.Continue(NewParagraph);

				NewParagraph.Correct_Content();
				NewParagraph.MoveCursorToStartPos();

				var nContentPos = this.CurPos.ContentPos;
				this.AddToContent(nContentPos, NewParagraph);
				this.CurPos.ContentPos = nContentPos + 1;
			}
			else
			{
				if (true === Item.IsCursorAtEnd())
				{
					if (!Item.Lock.Is_Locked())
						isCheckAutoCorrect = true;

					var StyleId = Item.Style_Get();
					var NextId  = undefined;

					if (undefined != StyleId)
					{
						NextId = this.Styles.Get_Next(StyleId);

						var oNextStyle = this.Styles.Get(NextId);
						if (!NextId || !oNextStyle || !oNextStyle.IsParagraphStyle())
							NextId = StyleId;
					}

					if (StyleId === NextId)
					{
						// Продолжаем (в плане настроек) новый параграф
						Item.Continue(NewParagraph);
					}
					else
					{
						// Простое добавление стиля, без дополнительных действий
						if (NextId === this.Styles.Get_Default_Paragraph())
							NewParagraph.Style_Remove();
						else
							NewParagraph.Style_Add(NextId, true);
					}

					var SectPr = Item.Get_SectionPr();
					if (undefined !== SectPr)
					{
						Item.Set_SectionPr(undefined);
						NewParagraph.Set_SectionPr(SectPr);
					}

					var LastRun = Item.Content[Item.Content.length - 1];
					if (LastRun && LastRun.Pr.Lang && LastRun.Pr.Lang.Val)
					{
						NewParagraph.SelectAll();
						NewParagraph.Add(new ParaTextPr({Lang : LastRun.Pr.Lang.Copy()}));
						NewParagraph.RemoveSelection();
					}
				}
				else
				{
					Item.Split(NewParagraph);
				}

				NewParagraph.Correct_Content();
				NewParagraph.MoveCursorToStartPos();

				var nContentPos = this.CurPos.ContentPos + 1;
				this.AddToContent(nContentPos, NewParagraph);
				this.CurPos.ContentPos = nContentPos;
			}

			if (true === this.IsTrackRevisions())
			{
				Item.RemovePrChange();
				NewParagraph.SetReviewType(ItemReviewType);
				Item.SetReviewType(reviewtype_Add);
			}
			else if (reviewtype_Common !== ItemReviewType)
			{
				NewParagraph.SetReviewType(ItemReviewType);
				Item.SetReviewType(reviewtype_Common);
			}
            NewParagraph.CheckSignatureLinesOnAdd();
		}

		if (isCheckAutoCorrect)
		{
			var nContentPos = this.CurPos.ContentPos;

			var oParaEndRun = Item.GetParaEndRun();
			if (oParaEndRun)
				oParaEndRun.ProcessAutoCorrectOnParaEnd();

			this.CurPos.ContentPos = nContentPos;
		}
	}
	else if (type_Table === Item.GetType() || type_BlockLevelSdt === Item.GetType())
	{
		// Если мы находимся в начале первого параграфа первой ячейки, и
		// данная таблица - первый элемент, тогда добавляем параграф до таблицы.

		if (0 === this.CurPos.ContentPos && Item.IsCursorAtBegin(true))
		{
			// Создаем новый параграф
			var NewParagraph = new Paragraph(this.DrawingDocument, this);
			this.Internal_Content_Add(0, NewParagraph);
			this.CurPos.ContentPos = 0;

			if (true === this.IsTrackRevisions())
			{
				NewParagraph.RemovePrChange();
				NewParagraph.SetReviewType(reviewtype_Add);
			}
		}
		else if (this.Content.length - 1 === this.CurPos.ContentPos && Item.IsCursorAtEnd())
		{
			var oNewParagraph = new Paragraph(this.DrawingDocument, this);
			this.Internal_Content_Add(this.Content.length, oNewParagraph);
			this.CurPos.ContentPos = this.Content.length - 1;

			if (this.IsTrackRevisions())
			{
				oNewParagraph.RemovePrChange();
				oNewParagraph.SetReviewType(reviewtype_Add);
			}
		}
		else
		{
			Item.AddNewParagraph();
		}
	}
};
CDocument.prototype.controller_AddInlineImage = function(W, H, Img, Chart, bFlow)
{
	if (true == this.Selection.Use)
		this.Remove(1, true);

	var Item = this.Content[this.CurPos.ContentPos];
	if (type_Paragraph == Item.GetType())
	{
		var Drawing;
		if (!AscCommon.isRealObject(Chart))
		{
			Drawing   = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
			var Image = this.DrawingObjects.createImage(Img, 0, 0, W, H);
			Image.setParent(Drawing);
			Drawing.Set_GraphicObject(Image);
		}
		else
		{
			Drawing   = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
			var Image = this.DrawingObjects.getChartSpace2(Chart, null);
			Image.setParent(Drawing);
			Drawing.Set_GraphicObject(Image);
			Drawing.setExtent(Image.spPr.xfrm.extX, Image.spPr.xfrm.extY);
		}
		if (true === bFlow)
		{
			Drawing.Set_DrawingType(drawing_Anchor);
			Drawing.Set_WrappingType(WRAPPING_TYPE_SQUARE);
			Drawing.Set_BehindDoc(false);
			Drawing.Set_Distance(3.2, 0, 3.2, 0);
			Drawing.Set_PositionH(Asc.c_oAscRelativeFromH.Column, false, 0, false);
			Drawing.Set_PositionV(Asc.c_oAscRelativeFromV.Paragraph, false, 0, false);
		}
		this.AddToParagraph(Drawing);
		this.Select_DrawingObject(Drawing.Get_Id());
	}
	else
	{
		Item.AddInlineImage(W, H, Img, Chart, bFlow);
	}
};
CDocument.prototype.AddPlaceholderImages = function (aImages, oPlaceholder)
{
	if (oPlaceholder && undefined !== oPlaceholder.id && aImages.length === 1 && aImages[0].Image)
	{
		const oController = this.Api.getGraphicController();
		const oPlaceholderTarget = AscCommon.g_oTableId.Get_ById(oPlaceholder.id);
		if (oPlaceholderTarget)
		{
			if (oPlaceholderTarget.isObjectInSmartArt && oPlaceholderTarget.isObjectInSmartArt())
			{
				
				if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Drawing_Props, undefined, false, false))
				{
					this.StartAction(AscDFH.historydescription_Document_AddPlaceholderImages);
					oPlaceholderTarget.applyImagePlaceholderCallback && oPlaceholderTarget.applyImagePlaceholderCallback(aImages, oPlaceholder);
					const nDrawingPage = oPlaceholderTarget.Get_AbsolutePage();
					if (AscFormat.isRealNumber(nDrawingPage))
					{
						oPlaceholderTarget.Set_CurrentElement(false, nDrawingPage, true);
					}
					this.Document_UpdateSelectionState();
					this.Document_UpdateUndoRedoState();
					this.Recalculate();
					this.FinalizeAction();
				}
			}
		}
	}
};
CDocument.prototype.controller_AddImages = function(aImages)
{
    if (true === this.Selection.Use)
        this.Remove(1, true);

    var Item = this.Content[this.CurPos.ContentPos];
    if (type_Paragraph === Item.GetType())
    {
        var Drawing, W, H;
        var ColumnSize = this.GetColumnSize();
        for(var i = 0; i < aImages.length; ++i){

            W = Math.max(1, ColumnSize.W);
            H = Math.max(1, ColumnSize.H);

			var _image = aImages[i];
			if(_image.Image)
			{
				var __w = Math.max((_image.Image.width * AscCommon.g_dKoef_pix_to_mm), 1);
				var __h = Math.max((_image.Image.height * AscCommon.g_dKoef_pix_to_mm), 1);
				W      = Math.max(5, Math.min(W, __w));
				H      = Math.max(5, Math.min((W * __h / __w)));
				Drawing   = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
				var Image = this.DrawingObjects.createImage(_image.src, 0, 0, W, H);
				Image.setParent(Drawing);
				Drawing.Set_GraphicObject(Image);
				this.AddToParagraph(Drawing);
			}
        }
        if(aImages.length === 1)
        {
            if(Drawing)
            {
                this.Select_DrawingObject(Drawing.Get_Id());
            }
        }
    }
    else
    {
        Item.AddImages(aImages);
    }
};
CDocument.prototype.controller_AddOleObject = function(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory)
{
	if (true == this.Selection.Use)
		this.Remove(1, true);

	let Item = this.Content[this.CurPos.ContentPos];
	let Drawing;
	if (type_Paragraph == Item.GetType())
	{
		Drawing   = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
		let Image = this.DrawingObjects.createOleObject(Data, sApplicationId, Img, 0, 0, W, H, nWidthPix, nHeightPix, arrImagesForAddToHistory);
		Image.setParent(Drawing);
		Drawing.Set_GraphicObject(Image);
		this.AddToParagraph(Drawing);
        if(bSelect !== false)
        {
            this.Select_DrawingObject(Drawing.Get_Id());
        }
	}
	else
	{
		Drawing = Item.AddOleObject(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory);
	}
	return Drawing;
};
CDocument.prototype.controller_AddTextArt = function(nStyle)
{
	var Item = this.Content[this.CurPos.ContentPos];
	if (type_Paragraph == Item.GetType())
	{
		var Drawing = new ParaDrawing(1828800 / 36000, 1828800 / 36000, null, this.DrawingDocument, this, null);
		var TextArt = this.DrawingObjects.createTextArt(nStyle, true);
		TextArt.setParent(Drawing);
		Drawing.Set_GraphicObject(TextArt);
		Drawing.Set_DrawingType(drawing_Anchor);
		Drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
		Drawing.Set_BehindDoc(false);
		Drawing.Set_Distance(3.2, 0, 3.2, 0);
		Drawing.Set_PositionH(Asc.c_oAscRelativeFromH.Column, false, 0, false);
		Drawing.Set_PositionV(Asc.c_oAscRelativeFromV.Paragraph, false, 0, false);

		if (true == this.Selection.Use)
			this.Remove(1, true);

		this.AddToParagraph(Drawing);
		if (TextArt.bSelectedText)
		{
			this.Select_DrawingObject(Drawing.Get_Id());
		}
		else
		{
			var oContent = Drawing.GraphicObj.getDocContent();
			oContent.Content[0].Document_SetThisElementCurrent(false);
			this.SelectAll();
		}
	}
	else
	{
		Item.AddTextArt(nStyle);
	}
};
CDocument.prototype.controller_AddSignatureLine = function(oSignatureDrawing)
{
	var Item = this.Content[this.CurPos.ContentPos];
	if (type_Paragraph == Item.GetType())
	{
		var Drawing = oSignatureDrawing;

		if (true == this.Selection.Use)
			this.Remove(1, true);

		this.AddToParagraph(Drawing);
        this.Select_DrawingObject(Drawing.Get_Id());
	}
	else
	{
		Item.AddSignatureLine(oSignatureDrawing);
	}
};
CDocument.prototype.controller_AddInlineTable = function(nCols, nRows, nMode)
{
	if (this.CurPos.ContentPos < 0)
		return null;

	// Сначала удаляем заселекченую часть
	if (true === this.Selection.Use)
	{
		this.Remove(1, true);
	}

	// Добавляем таблицу
	var Item = this.Content[this.CurPos.ContentPos];

	// Если мы внутри параграфа, тогда разрываем его и на месте разрыва добавляем таблицу.
	// А если мы внутри таблицы, тогда добавляем таблицу внутрь текущей таблицы.
	if (Item.IsParagraph())
	{
		// Ширину таблицы делаем по минимальной ширине колонки.
		var oPage  = this.Pages[this.CurPage];
		var SectPr = this.SectionsInfo.Get_SectPr(this.CurPos.ContentPos).SectPr;

		var PageFields = this.Get_PageFields(this.CurPage);

		var nAdd = this.GetCompatibilityMode() <= AscCommon.document_compatibility_mode_Word14 ?  2 * 1.9 : 0;

		// Создаем новую таблицу
		var W    = (PageFields.XLimit - PageFields.X + nAdd);
		var Grid = [];

		if (SectPr.Get_ColumnsCount() > 1)
		{
			for (var CurCol = 0, ColsCount = SectPr.Get_ColumnsCount(); CurCol < ColsCount; ++CurCol)
			{
				var ColumnWidth = SectPr.Get_ColumnWidth(CurCol);
				if (W > ColumnWidth)
					W = ColumnWidth;
			}

			W += nAdd;
		}

		W = Math.max(W, nCols * 2 * 1.9);

		for (var Index = 0; Index < nCols; Index++)
			Grid[Index] = W / nCols;

		var NewTable = new CTable(this.DrawingDocument, this, true, nRows, nCols, Grid);
		NewTable.SetParagraphPrOnAdd(Item);

		var nContentPos = this.CurPos.ContentPos;
		if (true === Item.IsCursorAtBegin() && undefined === Item.Get_SectionPr())
		{
			NewTable.MoveCursorToStartPos(false);
			this.AddToContent(nContentPos, NewTable);
			this.CurPos.ContentPos = nContentPos;
		}
		else
		{
			if (nMode < 0)
			{
				NewTable.MoveCursorToStartPos(false);

				if (Item.GetCurrentParaPos().Page > 0 && oPage && nContentPos === oPage.Pos)
				{
					this.AddToContent(nContentPos + 1, NewTable);
					this.CurPos.ContentPos = nContentPos + 1;
				}
				else
				{
					this.AddToContent(nContentPos, NewTable);
					this.CurPos.ContentPos = nContentPos;
				}
			}
			else if (nMode > 0)
			{
				NewTable.MoveCursorToStartPos(false);
				this.AddToContent(nContentPos + 1, NewTable);
				this.CurPos.ContentPos = nContentPos + 1;
			}
			else
			{
				var NewParagraph = new Paragraph(this.DrawingDocument, this);
				Item.Split(NewParagraph);

				this.AddToContent(nContentPos + 1, NewParagraph);

				NewTable.MoveCursorToStartPos(false);
				this.AddToContent(nContentPos + 1, NewTable);
				this.CurPos.ContentPos = nContentPos + 1;
			}
		}

		return NewTable;
	}
	else
	{
		return Item.AddInlineTable(nCols, nRows, nMode);
	}

	this.Recalculate();

	return null;
};
CDocument.prototype.controller_ClearParagraphFormatting = function(isClearParaPr, isClearTextPr)
{
	if (true === this.Selection.Use)
	{
		if (selectionflag_Common === this.Selection.Flag)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (StartPos > EndPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				this.Content[Index].ClearParagraphFormatting(isClearParaPr, isClearTextPr);
			}
		}
	}
	else
	{
		this.Content[this.CurPos.ContentPos].ClearParagraphFormatting(isClearParaPr, isClearTextPr);
	}
};
CDocument.prototype.controller_AddToParagraph = function(ParaItem, bRecalculate)
{
	if (true === this.Selection.Use)
	{
		var nSpaceCharCode = -1;
		if (this.IsWordSelection())
		{
			var sText = this.GetSelectedText();
			if (sText.length > 1 && AscCommon.IsSpace(sText.charCodeAt(sText.length - 1)))
				nSpaceCharCode = sText.charCodeAt(sText.length - 1);
		}

		var Type = ParaItem.Get_Type();
		switch (Type)
		{
			case para_Math:
			case para_NewLine:
			case para_Text:
			case para_Space:
			case para_Tab:
			case para_PageNum:
			case para_Field:
			case para_FootnoteReference:
			case para_FootnoteRef:
			case para_Separator:
			case para_ContinuationSeparator:
			case para_InstrText:
			case para_EndnoteReference:
			case para_EndnoteRef:
			{
				if (ParaItem instanceof AscCommonWord.MathMenu)
				{
					var oInfo = this.GetSelectedElementsInfo();
					if (oInfo.GetMath())
					{
						var oMath = oInfo.GetMath();
						if (!oMath.IsParentEquationPlaceholder())
							ParaItem.SetText(oMath.Copy(true));
					}
					else if (!oInfo.IsMixedSelection())
					{
						ParaItem.SetText(this.GetSelectedText({MathAdd : true}));
					}
				}

				// Если у нас что-то заселекчено и мы вводим текст или пробел
				// и т.д., тогда сначала удаляем весь селект.
				this.Remove(1, true, false, true);

				if (-1 !== nSpaceCharCode)
				{
					this.AddToParagraph(new AscWord.CRunSpace(nSpaceCharCode));
					this.MoveCursorLeft(false, false);
				}

				break;
			}
			case para_TextPr:
			{
				switch (this.Selection.Flag)
				{
					case selectionflag_Common:
					{
						// Текстовые настройки применяем ко всем параграфам, попавшим
						// в селект.
						var StartPos = this.Selection.StartPos;
						var EndPos   = this.Selection.EndPos;
						if (EndPos < StartPos)
						{
							var Temp = StartPos;
							StartPos = EndPos;
							EndPos   = Temp;
						}

						for (var Index = StartPos; Index <= EndPos; Index++)
						{
							this.Content[Index].Add(ParaItem.Copy());
						}

						if (false != bRecalculate)
						{
							// Если в TextPr только HighLight, тогда не надо ничего пересчитывать, только перерисовываем
							if (true === ParaItem.Value.Check_NeedRecalc())
							{
								this.Recalculate();
							}
							else
							{
								// Просто перерисовываем нужные страницы
								var StartPage = this.Content[StartPos].Get_StartPage_Absolute();
								var EndPage   = this.Content[EndPos].Get_StartPage_Absolute() + this.Content[EndPos].GetPagesCount() - 1;
								this.ReDraw(StartPage, EndPage);
							}
						}

						break;
					}
					case selectionflag_Numbering:
					{
						// Текстовые настройки применяем к конкретной нумерации
						if (!this.Selection.Data || !this.Selection.Data.CurPara)
							break;

						if (undefined != ParaItem.Value.FontFamily)
						{
							var FName  = ParaItem.Value.FontFamily.Name;
							var FIndex = ParaItem.Value.FontFamily.Index;

							ParaItem.Value.RFonts          = new CRFonts();
							ParaItem.Value.RFonts.Ascii    = {Name : FName, Index : FIndex};
							ParaItem.Value.RFonts.EastAsia = {Name : FName, Index : FIndex};
							ParaItem.Value.RFonts.HAnsi    = {Name : FName, Index : FIndex};
							ParaItem.Value.RFonts.CS       = {Name : FName, Index : FIndex};
						}

						var oNumPr = this.Selection.Data.CurPara.GetNumPr();
						var oNum   = this.GetNumbering().GetNum(oNumPr.NumId);
						oNum.ApplyTextPr(oNumPr.Lvl, ParaItem.Value);
						break;
					}
				}

				this.Document_UpdateSelectionState();
				this.Document_UpdateUndoRedoState();

				return;
			}
		}
	}

	var nContentPos = this.CurPos.ContentPos;

	var Item     = this.Content[nContentPos];
	var ItemType = Item.GetType();

	if (para_NewLine === ParaItem.Type && true === ParaItem.IsPageOrColumnBreak())
	{
		if (type_Paragraph === ItemType)
		{
			if (true === Item.IsCursorAtBegin())
			{
				if (ParaItem.IsColumnBreak())
				{
					this.Content[this.CurPos.ContentPos].Add(ParaItem);
				}
				else
				{
					this.AddNewParagraph(undefined, true);

					if (this.Content[nContentPos] && this.Content[nContentPos].IsParagraph())
					{
						this.Content[nContentPos].AddToParagraph(ParaItem);
						this.Content[nContentPos].Clear_Formatting();
					}

					this.CurPos.ContentPos = nContentPos + 1;
				}
			}
			else
			{
				if (ParaItem.IsColumnBreak())
				{
					var oCurElement = this.Content[this.CurPos.ContentPos];
					if (oCurElement && type_Paragraph === oCurElement.Get_Type() && oCurElement.IsColumnBreakOnLeft())
					{
						oCurElement.AddToParagraph(ParaItem);
					}
					else
					{
						this.AddNewParagraph(undefined, true);

						nContentPos = this.CurPos.ContentPos;
						if (this.Content[nContentPos] && this.Content[nContentPos].IsParagraph())
						{
							this.Content[nContentPos].MoveCursorToStartPos(false);
							this.Content[nContentPos].AddToParagraph(ParaItem);
						}
					}
				}
				else
				{
					this.AddNewParagraph(undefined, true);
					this.CurPos.ContentPos = nContentPos + 1;
					this.Content[nContentPos + 1].MoveCursorToStartPos();
					this.AddNewParagraph(undefined, true);

					if (this.Content[nContentPos + 1] && this.Content[nContentPos + 1].IsParagraph())
					{
						this.Content[nContentPos + 1].AddToParagraph(ParaItem);
						this.Content[nContentPos + 1].Clear_Formatting();
					}

					this.CurPos.ContentPos = nContentPos + 2;
					this.Content[nContentPos + 1].MoveCursorToStartPos();
				}
			}

			if (false != bRecalculate)
			{
				this.Recalculate();

				Item.CurPos.RealX = Item.CurPos.X;
				Item.CurPos.RealY = Item.CurPos.Y;
			}
		}
		else if (type_BlockLevelSdt === Item.GetType())
		{
			Item.AddToParagraph(ParaItem);
		}
		else if (Item.IsTable())
		{
			if (!ParaItem.IsPageBreak() && Item.IsInnerTable())
			{
				Item.AddToParagraph(ParaItem);
			}
			else
			{
				var oNewTable = Item.Split();
				var oNewPara  = new Paragraph(this.DrawingDocument, this);

				if (ParaItem.IsPageBreak())
					oNewPara.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));

				var nCurPos = this.CurPos.ContentPos;
				if (oNewTable)
				{
					this.AddToContent(nCurPos + 1, oNewTable);
					this.AddToContent(nCurPos + 1, oNewPara);
					this.CurPos.ContentPos = nCurPos + 1;
				}
				else
				{
					this.AddToContent(nCurPos, oNewPara);
					this.CurPos.ContentPos = nCurPos;
				}

				this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
			}

			this.Recalculate();
			this.UpdateInterface();
			this.UpdateSelection();
			return;
		}
	}
	else
	{
		Item.AddToParagraph(ParaItem);

		if (false != bRecalculate && type_Paragraph == Item.GetType())
		{
			if (para_TextPr === ParaItem.Type && false === ParaItem.Value.Check_NeedRecalc())
			{
				// Просто перерисовываем нужные страницы
				var StartPage = Item.Get_StartPage_Absolute();
				var EndPage   = StartPage + Item.Pages.length - 1;
				this.ReDraw(StartPage, EndPage);
			}
			else
			{
				this.Recalculate();
			}

			if (false === this.TurnOffRecalcCurPos)
			{
				Item.RecalculateCurPos();
				Item.CurPos.RealX = Item.CurPos.X;
				Item.CurPos.RealY = Item.CurPos.Y;
			}
		}

		this.UpdateSelection();
		this.UpdateInterface();
	}

	// Специальная заглушка для функции TextBox_Put
	if (true === this.Is_OnRecalculate())
		this.Document_UpdateUndoRedoState();
};
CDocument.prototype.controller_Remove = function(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord)
{
	this.private_Remove(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord);
};
CDocument.prototype.controller_GetCursorPosXY = function()
{
	if (true === this.Selection.Use)
	{
		if (selectionflag_Common === this.Selection.Flag)
			return this.Content[this.Selection.EndPos].GetCursorPosXY();

		return {X : 0, Y : 0};
	}
	else
	{
		return this.Content[this.CurPos.ContentPos].GetCursorPosXY();
	}
};
CDocument.prototype.controller_MoveCursorToStartPos = function(AddToSelect)
{
	if (true === AddToSelect)
	{
		var StartPos = ( true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
		var EndPos   = 0;

		this.Selection.Start    = false;
		this.Selection.Use      = true;
		this.Selection.StartPos = StartPos;
		this.Selection.EndPos   = EndPos;
		this.Selection.Flag     = selectionflag_Common;

		this.CurPos.ContentPos = 0;
		this.SetDocPosType(docpostype_Content);

		for (var Index = StartPos - 1; Index >= EndPos; Index--)
		{
			this.Content[Index].SelectAll(-1);
		}

		this.Content[StartPos].MoveCursorToStartPos(true);
	}
	else
	{
		this.RemoveSelection();

		this.Selection.Start    = false;
		this.Selection.Use      = false;
		this.Selection.StartPos = 0;
		this.Selection.EndPos   = 0;
		this.Selection.Flag     = selectionflag_Common;

		this.CurPos.ContentPos = 0;
		this.SetDocPosType(docpostype_Content);
		this.Content[0].MoveCursorToStartPos(false);
	}
};
CDocument.prototype.controller_MoveCursorToEndPos = function(AddToSelect)
{
	if (true === AddToSelect)
	{
		var StartPos = ( true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
		var EndPos   = this.Content.length - 1;

		this.Selection.Start    = false;
		this.Selection.Use      = true;
		this.Selection.StartPos = StartPos;
		this.Selection.EndPos   = EndPos;
		this.Selection.Flag     = selectionflag_Common;

		this.CurPos.ContentPos = this.Content.length - 1;
		this.SetDocPosType(docpostype_Content);

		for (var Index = StartPos + 1; Index <= EndPos; Index++)
		{
			this.Content[Index].SelectAll(1);
		}

		this.Content[StartPos].MoveCursorToEndPos(true, false);
	}
	else
	{
		this.RemoveSelection();

		this.Selection.Start    = false;
		this.Selection.Use      = false;
		this.Selection.StartPos = 0;
		this.Selection.EndPos   = 0;
		this.Selection.Flag     = selectionflag_Common;

		this.CurPos.ContentPos = this.Content.length - 1;
		this.SetDocPosType(docpostype_Content);
		this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false);
	}
};
CDocument.prototype.controller_MoveCursorLeft = function(AddToSelect, Word)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	this.RemoveNumberingSelection();
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			if (!this.Content[this.Selection.EndPos].MoveCursorLeft(true, Word)
				&& 0 !== this.Selection.EndPos)
			{
				if (this.Selection.EndPos > this.Selection.StartPos)
					this.Content[this.Selection.EndPos].RemoveSelection();

				this.Selection.EndPos--;
				this.CurPos.ContentPos = this.Selection.EndPos;

				var Item = this.Content[this.Selection.EndPos];
				if (this.Selection.StartPos <= this.Selection.EndPos)
					Item.MoveCursorLeft(true, Word);
				else
					Item.MoveCursorLeftWithSelectionFromEnd(Word);
			}

			// Проверяем не обнулился ли селект в последнем элементе. Такое могло быть, если была
			// заселекчена одна буква в последнем параграфе, а мы убрали селект последним действием.
			if (this.Selection.EndPos != this.Selection.StartPos && false === this.Content[this.Selection.EndPos].IsSelectionUse())
			{
				// Такая ситуация возможна только при прямом селекте (сверху вниз), поэтому вычитаем
				this.Selection.EndPos--;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}

			// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
			if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
			{
				this.Selection.Use     = false;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}
		}
		else
		{
			// Нам нужно переместить курсор в левый край селекта, и отменить весь селект
			var Start = this.Selection.StartPos;
			if (Start > this.Selection.EndPos)
				Start = this.Selection.EndPos;

			this.CurPos.ContentPos = Start;
			if (false === this.Content[this.CurPos.ContentPos].MoveCursorLeft(false, Word))
			{
				if (this.CurPos.ContentPos > 0)
				{
					this.CurPos.ContentPos--;
					this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false, false);
				}
			}

			this.RemoveSelection();
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = this.CurPos.ContentPos;
			this.Selection.EndPos   = this.CurPos.ContentPos;

			if (false === this.Content[this.CurPos.ContentPos].MoveCursorLeft(true, Word))
			{
				// Нужно перейти в конец предыдущего элемент
				if (0 != this.CurPos.ContentPos)
				{
					this.CurPos.ContentPos--;
					this.Selection.EndPos = this.CurPos.ContentPos;

					var Item = this.Content[this.CurPos.ContentPos];
					Item.MoveCursorLeftWithSelectionFromEnd(Word);
				}
			}

			// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
			if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
			{
				this.Selection.Use     = false;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}
		}
		else
		{
			if (false === this.Content[this.CurPos.ContentPos].MoveCursorLeft(false, Word))
			{
				// Нужно перейти в конец предыдущего элемент
				if (0 != this.CurPos.ContentPos)
				{
					this.CurPos.ContentPos--;
					this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false, false);
				}
			}
		}
	}
};
CDocument.prototype.controller_MoveCursorRight = function(AddToSelect, Word)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	this.RemoveNumberingSelection();
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			if (!this.Content[this.Selection.EndPos].MoveCursorRight(true, Word)
				&& this.Content.length - 1 !== this.Selection.EndPos)
			{
				if (this.Selection.EndPos < this.Selection.StartPos)
					this.Content[this.Selection.EndPos].RemoveSelection();

				this.Selection.EndPos++;
				this.CurPos.ContentPos = this.Selection.EndPos;

				var Item = this.Content[this.Selection.EndPos];
				if (this.Selection.StartPos >= this.Selection.EndPos)
					Item.MoveCursorRight(true, Word);
				else
					Item.MoveCursorRightWithSelectionFromStart(Word);
			}

			// Проверяем не обнулился ли селект в последнем параграфе. Такое могло быть, если была
			// заселекчена одна буква в последнем параграфе, а мы убрали селект последним действием.
			if (this.Selection.EndPos != this.Selection.StartPos && false === this.Content[this.Selection.EndPos].IsSelectionUse())
			{
				// Такая ситуация возможна только при обратном селекте (снизу вверх), поэтому вычитаем
				this.Selection.EndPos++;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}

			// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
			if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
			{
				this.Selection.Use     = false;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}
		}
		else
		{
			// Нам нужно переместить курсор в правый край селекта, и отменить весь селект
			var End = this.Selection.EndPos;
			if (End < this.Selection.StartPos)
				End = this.Selection.StartPos;

			this.CurPos.ContentPos = End;

			if (true === this.Content[this.CurPos.ContentPos].IsSelectionToEnd() && this.CurPos.ContentPos < this.Content.length - 1)
			{
				this.CurPos.ContentPos = End + 1;
				this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
			}
			else
			{
				this.Content[this.CurPos.ContentPos].MoveCursorRight(false, Word);
			}

			this.RemoveSelection();
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = this.CurPos.ContentPos;
			this.Selection.EndPos   = this.CurPos.ContentPos;

			if (false === this.Content[this.CurPos.ContentPos].MoveCursorRight(true, Word))
			{
				// Нужно перейти в конец предыдущего элемента
				if (this.Content.length - 1 != this.CurPos.ContentPos)
				{
					this.CurPos.ContentPos++;
					this.Selection.EndPos = this.CurPos.ContentPos;

					var Item = this.Content[this.CurPos.ContentPos];
					Item.MoveCursorRightWithSelectionFromStart(Word);
				}
			}

			// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
			if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
			{
				this.Selection.Use     = false;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}
		}
		else
		{
			if (false === this.Content[this.CurPos.ContentPos].MoveCursorRight(false, Word))
			{
				// Нужно перейти в начало следующего элемента
				if (this.Content.length - 1 != this.CurPos.ContentPos)
				{
					this.CurPos.ContentPos++;
					this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
				}
			}
		}
	}
};
CDocument.prototype.controller_MoveCursorUp = function(AddToSelect)
{
	if (true === this.IsSelectionUse() && true !== AddToSelect)
	{
		var oCurParagraph = this.GetCurrentParagraph(true);
		this.MoveCursorLeft(false, false);

		if (this.GetCurrentParagraph(true) !== oCurParagraph)
			return;
	}

	var bStopSelection = false;
	if (true !== this.IsSelectionUse() && true === AddToSelect)
	{
		bStopSelection = true;
		this.StartSelectionFromCurPos();
	}

	this.private_UpdateCursorXY(false, true);
	var Result = this.private_MoveCursorUp(this.CurPos.RealX, this.CurPos.RealY, AddToSelect);

	// TODO: Вообще Word селектит до начала данной колонки в таком случае, а не до начала документа
	if (true === AddToSelect && true !== Result)
		this.MoveCursorToStartPos(true);

	if (bStopSelection)
		this.private_StopSelection();
};
CDocument.prototype.controller_MoveCursorDown = function(AddToSelect)
{
	if (true === this.IsSelectionUse() && true !== AddToSelect)
	{
		var oCurParagraph = this.GetCurrentParagraph(true);
		this.MoveCursorRight(false, false);

		if (this.GetCurrentParagraph(true) !== oCurParagraph)
			return;
	}

	var bStopSelection = false;
	if (true !== this.IsSelectionUse() && true === AddToSelect)
	{
		bStopSelection = true;
		this.StartSelectionFromCurPos();
	}

	this.private_UpdateCursorXY(false, true);
	var Result = this.private_MoveCursorDown(this.CurPos.RealX, this.CurPos.RealY, AddToSelect);

	if (true === AddToSelect && true !== Result)
		this.MoveCursorToEndPos(true);

	if (bStopSelection)
		this.private_StopSelection();
};
CDocument.prototype.controller_MoveCursorToEndOfLine = function(AddToSelect)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	this.RemoveNumberingSelection();
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			var Item = this.Content[this.Selection.EndPos];
			Item.MoveCursorToEndOfLine(AddToSelect);

			// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
			if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
			{
				this.Selection.Use     = false;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}
		}
		else
		{
			var Pos                = ( this.Selection.EndPos >= this.Selection.StartPos ? this.Selection.EndPos : this.Selection.StartPos );
			this.CurPos.ContentPos = Pos;

			var Item = this.Content[Pos];
			Item.MoveCursorToEndOfLine(AddToSelect);

			this.RemoveSelection();
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = this.CurPos.ContentPos;
			this.Selection.EndPos   = this.CurPos.ContentPos;

			var Item = this.Content[this.CurPos.ContentPos];
			Item.MoveCursorToEndOfLine(AddToSelect);

			// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
			if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
			{
				this.Selection.Use     = false;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.MoveCursorToEndOfLine(AddToSelect);
		}
	}

	this.Document_UpdateInterfaceState();
	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.controller_MoveCursorToStartOfLine = function(AddToSelect)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	this.RemoveNumberingSelection();
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			var Item = this.Content[this.Selection.EndPos];
			Item.MoveCursorToStartOfLine(AddToSelect);

			// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
			if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
			{
				this.Selection.Use     = false;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}
		}
		else
		{
			var Pos                = ( this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos );
			this.CurPos.ContentPos = Pos;

			var Item = this.Content[Pos];
			Item.MoveCursorToStartOfLine(AddToSelect);

			this.RemoveSelection();
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = this.CurPos.ContentPos;
			this.Selection.EndPos   = this.CurPos.ContentPos;

			var Item = this.Content[this.CurPos.ContentPos];
			Item.MoveCursorToStartOfLine(AddToSelect);

			// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
			if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
			{
				this.Selection.Use     = false;
				this.CurPos.ContentPos = this.Selection.EndPos;
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.MoveCursorToStartOfLine(AddToSelect);
		}
	}

	this.Document_UpdateInterfaceState();
	this.private_UpdateCursorXY(true, true);
};
CDocument.prototype.controller_MoveCursorToXY = function(X, Y, PageAbs, AddToSelect)
{
	this.CurPage = PageAbs;

	this.RemoveNumberingSelection();
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			this.Selection_SetEnd(X, Y, true);
		}
		else
		{
			this.RemoveSelection();

			var ContentPos         = this.Internal_GetContentPosByXY(X, Y);
			this.CurPos.ContentPos = ContentPos;
			var ElementPageIndex   = this.private_GetElementPageIndexByXY(ContentPos, X, Y, PageAbs);
			this.Content[ContentPos].MoveCursorToXY(X, Y, false, false, ElementPageIndex);

			this.Document_UpdateInterfaceState();
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			this.StartSelectionFromCurPos();
			var oMouseEvent  = new AscCommon.CMouseEventHandler();
			oMouseEvent.Type = AscCommon.g_mouse_event_type_up;
			this.Selection_SetEnd(X, Y, oMouseEvent);
		}
		else
		{
			var ContentPos         = this.Internal_GetContentPosByXY(X, Y);
			this.CurPos.ContentPos = ContentPos;
			var ElementPageIndex   = this.private_GetElementPageIndexByXY(ContentPos, X, Y, PageAbs);
			this.Content[ContentPos].MoveCursorToXY(X, Y, false, false, ElementPageIndex);

			this.Document_UpdateInterfaceState();
		}
	}
};
CDocument.prototype.controller_MoveCursorToCell = function(bNext)
{
	if (true === this.Selection.Use)
	{
		if (this.Selection.StartPos === this.Selection.EndPos)
			this.Content[this.Selection.StartPos].MoveCursorToCell(bNext);
	}
	else
	{
		this.Content[this.CurPos.ContentPos].MoveCursorToCell(bNext);
	}
};
CDocument.prototype.controller_SetParagraphAlign = function(Align)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphAlign(Align);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphAlign(Align);
	}
};
CDocument.prototype.controller_SetParagraphSpacing = function(Spacing)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphSpacing(Spacing);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphSpacing(Spacing);
	}
};
CDocument.prototype.controller_SetParagraphTabs = function(Tabs)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphTabs(Tabs);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphTabs(Tabs);
	}
};
CDocument.prototype.controller_SetParagraphIndent = function(Ind)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphIndent(Ind);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphIndent(Ind);
	}
};
CDocument.prototype.controller_SetParagraphShd = function(Shd)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;

		if (true === this.UseTextShd && StartPos === EndPos && type_Paragraph === this.Content[StartPos].GetType() && false === this.Content[StartPos].Selection_CheckParaEnd() && selectionflag_Common === this.Selection.Flag)
		{
			this.AddToParagraph(new ParaTextPr({Shd : Shd}));
		}
		else
		{
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphShd(Shd);
			}
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphShd(Shd);
	}
};
CDocument.prototype.controller_SetParagraphStyle = function(Name)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		this.RemoveNumberingSelection();

		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphStyle(Name);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphStyle(Name);
	}
};
CDocument.prototype.controller_SetParagraphContextualSpacing = function(Value)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphContextualSpacing(Value);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphContextualSpacing(Value);
	}
};
CDocument.prototype.controller_SetParagraphPageBreakBefore = function(Value)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphPageBreakBefore(Value);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphPageBreakBefore(Value);
	}
};
CDocument.prototype.controller_SetParagraphKeepLines = function(Value)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphKeepLines(Value);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphKeepLines(Value);
	}
};
CDocument.prototype.controller_SetParagraphKeepNext = function(Value)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphKeepNext(Value);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphKeepNext(Value);
	}
};
CDocument.prototype.controller_SetParagraphWidowControl = function(Value)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphWidowControl(Value);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphWidowControl(Value);
	}
};
CDocument.prototype.controller_SetParagraphBorders = function(Borders)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphBorders(Borders);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		if (type_Paragraph == Item.GetType())
		{
			// Мы должны выставить границу для всех параграфов, входящих в текущую группу параграфов
			// с одинаковыми границами

			var StartPos = Item.Index;
			var EndPos   = Item.Index;
			var CurBrd   = Item.Get_CompiledPr().ParaPr.Brd;

			while (true != CurBrd.First)
			{
				StartPos--;
				if (StartPos < 0)
				{
					StartPos = 0;
					break;
				}

				var TempItem = this.Content[StartPos];
				if (type_Paragraph !== TempItem.GetType())
				{
					StartPos++;
					break;
				}

				CurBrd = TempItem.Get_CompiledPr().ParaPr.Brd;
			}

			CurBrd = Item.Get_CompiledPr().ParaPr.Brd;
			while (true != CurBrd.Last)
			{
				EndPos++;
				if (EndPos >= this.Content.length)
				{
					EndPos = this.Content.length - 1;
					break;
				}

				var TempItem = this.Content[EndPos];
				if (type_Paragraph !== TempItem.GetType())
				{
					EndPos--;
					break;
				}

				CurBrd = TempItem.Get_CompiledPr().ParaPr.Brd;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
				this.Content[Index].SetParagraphBorders(Borders);
		}
		else
		{
			Item.SetParagraphBorders(Borders);
		}
	}
};
CDocument.prototype.controller_SetParagraphFramePr = function(FramePr, bDelete)
{
	if (true === this.Selection.Use)
	{
		if (this.Selection.StartPos === this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		{
			this.Content[this.Selection.StartPos].SetParagraphFramePr(FramePr, bDelete);
			return;
		}

		// Проверим, если у нас все выделенные элементы - параграфы, с одинаковыми настройками
		// FramePr, тогда мы можем применить новую настройку FramePr

		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;

		if (StartPos > EndPos)
		{
			StartPos = this.Selection.EndPos;
			EndPos   = this.Selection.StartPos;
		}

		var Element = this.Content[StartPos];

		if (type_Paragraph !== Element.GetType() || undefined === Element.Get_FramePr())
			return;

		var FramePr = Element.Get_FramePr();
		for (var Pos = StartPos + 1; Pos < EndPos; Pos++)
		{
			var TempElement = this.Content[Pos];

			if (type_Paragraph !== TempElement.GetType() || undefined === TempElement.Get_FramePr() || true != FramePr.Compare(TempElement.Get_FramePr()))
				return;
		}

		// Раз дошли до сюда, значит можно у всех выделенных параграфов менять настройку рамки
		var FrameParas = this.Content[StartPos].Internal_Get_FrameParagraphs();
		var FrameCount = FrameParas.length;
		for (var Pos = 0; Pos < FrameCount; Pos++)
		{
			FrameParas[Pos].Set_FramePr(FramePr, bDelete);
		}
	}
	else
	{
		var Element = this.Content[this.CurPos.ContentPos];

		if (type_Paragraph !== Element.GetType())
		{
			Element.SetParagraphFramePr(FramePr, bDelete);
			return;
		}

		// Возможно, предыдущий элемент является буквицей
		if (undefined === Element.Get_FramePr())
		{
			var PrevElement = Element.Get_DocumentPrev();

			if (type_Paragraph !== PrevElement.GetType() || undefined === PrevElement.Get_FramePr() || undefined === PrevElement.Get_FramePr().DropCap)
				return;

			Element = PrevElement;
		}

		var FrameParas = Element.Internal_Get_FrameParagraphs();
		var FrameCount = FrameParas.length;
		for (var Pos = 0; Pos < FrameCount; Pos++)
		{
			FrameParas[Pos].Set_FramePr(FramePr, bDelete);
		}
	}
};
CDocument.prototype.controller_IncreaseDecreaseFontSize = function(bIncrease)
{
	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		switch (this.Selection.Flag)
		{
			case selectionflag_Common:
			{
				var StartPos = this.Selection.StartPos;
				var EndPos   = this.Selection.EndPos;
				if (EndPos < StartPos)
				{
					var Temp = StartPos;
					StartPos = EndPos;
					EndPos   = Temp;
				}

				for (var Index = StartPos; Index <= EndPos; Index++)
				{
					var Item = this.Content[Index];
					Item.IncreaseDecreaseFontSize(bIncrease);
				}
				break;
			}
			case  selectionflag_Numbering:
			{
				var NewFontSize = this.GetCalculatedTextPr().GetIncDecFontSize(bIncrease);
				var TextPr      = new CTextPr();
				TextPr.FontSize = NewFontSize;
				this.AddToParagraph(new ParaTextPr(TextPr), true);
				break;
			}
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.IncreaseDecreaseFontSize(bIncrease);
	}
};
CDocument.prototype.controller_IncreaseDecreaseIndent = function(bIncrease)
{
	if (true === this.Selection.Use && selectionflag_Common === this.Selection.Flag)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			this.Content[Index].IncreaseDecreaseIndent(bIncrease);
		}
	}
	else
	{
		this.Content[this.CurPos.ContentPos].IncreaseDecreaseIndent(bIncrease);
	}
};
CDocument.prototype.controller_SetImageProps = function(Props)
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Table == this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Table == this.Content[this.CurPos.ContentPos].GetType()))
	{
		if (true == this.Selection.Use)
			this.Content[this.Selection.StartPos].SetImageProps(Props);
		else
			this.Content[this.CurPos.ContentPos].SetImageProps(Props);
	}
};
CDocument.prototype.controller_SetTableProps = function(Props)
{
	var Pos = -1;
	if (true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos)
		Pos = this.Selection.StartPos;
	else if (false === this.Selection.Use)
		Pos = this.CurPos.ContentPos;

	if (-1 !== Pos)
		this.Content[Pos].SetTableProps(Props);
};
CDocument.prototype.controller_GetCalculatedParaPr = function()
{
	var Result_ParaPr = new CParaPr();
	if (true === this.Selection.Use && selectionflag_Common === this.Selection.Flag)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		var StartPr = this.Content[StartPos].GetCalculatedParaPr();
		var Pr      = StartPr.Copy();
		Pr.Locked   = StartPr.Locked;

		for (var Index = StartPos + 1; Index <= EndPos; Index++)
		{
			var Item   = this.Content[Index];
			var TempPr = Item.GetCalculatedParaPr();
			Pr         = Pr.Compare(TempPr);
		}

		if (undefined === Pr.Ind.Left)
			Pr.Ind.Left = StartPr.Ind.Left;

		if (undefined === Pr.Ind.Right)
			Pr.Ind.Right = StartPr.Ind.Right;

		if (undefined === Pr.Ind.FirstLine)
			Pr.Ind.FirstLine = StartPr.Ind.FirstLine;

		Result_ParaPr             = Pr;
		Result_ParaPr.CanAddTable = true !== Pr.Locked;

		// Если мы находимся в рамке, тогда дополняем ее свойства настройками границы и настройкой текста (если это буквица)
		if (undefined != Result_ParaPr.FramePr && type_Paragraph === this.Content[StartPos].GetType())
		{
			this.Content[StartPos].Supplement_FramePr(Result_ParaPr.FramePr);
		}
		else if (StartPos === EndPos && StartPos > 0 && type_Paragraph === this.Content[StartPos - 1].GetType())
		{
			var PrevFrame = this.Content[StartPos - 1].Get_FramePr();
			if (undefined != PrevFrame && undefined != PrevFrame.DropCap)
			{
				Result_ParaPr.FramePr = PrevFrame.Copy();
				this.Content[StartPos - 1].Supplement_FramePr(Result_ParaPr.FramePr);
			}
		}

	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		if (type_Paragraph == Item.GetType())
		{
			Result_ParaPr             = Item.GetCalculatedParaPr().Copy();
			Result_ParaPr.CanAddTable = true === Result_ParaPr.Locked ? Item.IsCursorAtEnd() : true;

			// Если мы находимся в рамке, тогда дополняем ее свойства настройками границы и настройкой текста (если это буквица)
			if (undefined != Result_ParaPr.FramePr)
			{
				Item.Supplement_FramePr(Result_ParaPr.FramePr);
			}
			else if (this.CurPos.ContentPos > 0 && type_Paragraph === this.Content[this.CurPos.ContentPos - 1].GetType())
			{
				var PrevFrame = this.Content[this.CurPos.ContentPos - 1].Get_FramePr();
				if (undefined != PrevFrame && undefined != PrevFrame.DropCap)
				{
					Result_ParaPr.FramePr = PrevFrame.Copy();
					this.Content[this.CurPos.ContentPos - 1].Supplement_FramePr(Result_ParaPr.FramePr);
				}
			}
		}
		else
		{
			Result_ParaPr = Item.GetCalculatedParaPr();
		}
	}

	if (Result_ParaPr.Shd && Result_ParaPr.Shd.Unifill)
	{
		Result_ParaPr.Shd.Unifill.check(this.theme, this.Get_ColorMap());
	}

	return Result_ParaPr;
};
CDocument.prototype.controller_GetCalculatedTextPr = function()
{
	var Result_TextPr = null;
	if (true === this.Selection.Use)
	{
		var VisTextPr;
		switch (this.Selection.Flag)
		{
			case selectionflag_Common:
			{
				var StartPos = this.Selection.StartPos;
				var EndPos   = this.Selection.EndPos;
				if (EndPos < StartPos)
				{
					var Temp = StartPos;
					StartPos = EndPos;
					EndPos   = Temp;
				}

				VisTextPr = this.Content[StartPos].GetCalculatedTextPr();

				for (var Index = StartPos + 1; Index <= EndPos; Index++)
				{
					var CurPr = this.Content[Index].GetCalculatedTextPr();
					VisTextPr = VisTextPr.Compare(CurPr);
				}

				break;
			}
			case selectionflag_Numbering:
			{
				if (!this.Selection.Data || !this.Selection.Data.CurPara)
					break;

				VisTextPr = this.Selection.Data.CurPara.GetNumberingTextPr();
				break;
			}
		}

		Result_TextPr = VisTextPr;
	}
	else
	{
		Result_TextPr = this.Content[this.CurPos.ContentPos].GetCalculatedTextPr();
	}

	return Result_TextPr;
};
CDocument.prototype.controller_GetDirectParaPr = function()
{
	var Result_ParaPr = null;

	if (true === this.Selection.Use)
	{
		switch (this.Selection.Flag)
		{
			case selectionflag_Common:
			{
				var StartPos = this.Selection.StartPos;
				if (this.Selection.EndPos < StartPos)
					StartPos = this.Selection.EndPos;

				var Item      = this.Content[StartPos];
				Result_ParaPr = Item.GetDirectParaPr();

				break;
			}
			case selectionflag_Numbering:
			{
				if (!this.Selection.Data || !this.Selection.Data.CurPara)
					break;

				var oNumPr    = this.Selection.Data.CurPara.GetNumPr();
				Result_ParaPr = this.GetNumbering().GetNum(oNumPr.NumId).GetLvl(oNumPr.Lvl).GetParaPr();

				break;
			}
		}
	}
	else
	{
		var Item      = this.Content[this.CurPos.ContentPos];
		Result_ParaPr = Item.GetDirectParaPr();
	}

	return Result_ParaPr;
};
CDocument.prototype.controller_GetDirectTextPr = function()
{
	var Result_TextPr = null;

	if (true === this.Selection.Use)
	{
		var VisTextPr;
		switch (this.Selection.Flag)
		{
			case selectionflag_Common:
			{
				var StartPos = this.Selection.StartPos;
				if (this.Selection.EndPos < StartPos)
					StartPos = this.Selection.EndPos;

				var Item  = this.Content[StartPos];
				VisTextPr = Item.GetDirectTextPr();

				break;
			}
			case selectionflag_Numbering:
			{
				if (!this.Selection.Data || !this.Selection.Data.CurPara)
					break;

				var oNumPr = this.Selection.Data.CurPara.GetNumPr();
				VisTextPr  = this.GetNumbering().GetNum(oNumPr.NumId).GetLvl(oNumPr.Lvl).GetTextPr();

				break;
			}
		}

		Result_TextPr = VisTextPr;
	}
	else
	{
		var Item      = this.Content[this.CurPos.ContentPos];
		Result_TextPr = Item.GetDirectTextPr();
	}

	return Result_TextPr;
};
CDocument.prototype.controller_RemoveSelection = function(bNoCheckDrawing)
{
	if (true === this.Selection.Use)
	{
		switch (this.Selection.Flag)
		{
			case selectionflag_Common:
			{
				var Start = this.Selection.StartPos;
				var End   = this.Selection.EndPos;

				if (Start > End)
				{
					var Temp = Start;
					Start    = End;
					End      = Temp;
				}

				Start = Math.max(0, Start);
				End   = Math.min(this.Content.length - 1, End);

				for (var Index = Start; Index <= End; Index++)
				{
					this.Content[Index].RemoveSelection();
				}

				this.Selection.Use   = false;
				this.Selection.Start = false;

				this.Selection.StartPos = 0;
				this.Selection.EndPos   = 0;

				// Убираем селект и возвращаем курсор
				this.DrawingDocument.SelectEnabled(false);
				this.DrawingDocument.TargetStart();
				this.DrawingDocument.TargetShow();

				break;
			}
			case selectionflag_Numbering:
			{
				if (!this.Selection.Data)
					break;

				for (var nIndex = 0, nCount = this.Selection.Data.Paragraphs.length; nIndex < nCount; ++nIndex)
				{
					this.Selection.Data.Paragraphs[nIndex].RemoveSelection();
				}

				if (this.Selection.Data.CurPara)
				{
					this.Selection.Data.CurPara.RemoveSelection();
					this.Selection.Data.CurPara.MoveCursorToStartPos();
				}

				var Start = this.Selection.StartPos;
				var End   = this.Selection.EndPos;

				if (Start > End)
				{
					var Temp = Start;
					Start    = End;
					End      = Temp;
				}

				Start = Math.max(0, Start);
				End   = Math.min(this.Content.length - 1, End);

				for (var Index = Start; Index <= End; Index++)
				{
					this.Content[Index].RemoveSelection();
				}

				this.Selection.Use   = false;
				this.Selection.Start = false;
				this.Selection.Flag  = selectionflag_Common;

				// Убираем селект и возвращаем курсор
				this.DrawingDocument.SelectEnabled(false);
				this.DrawingDocument.TargetStart();
				this.DrawingDocument.TargetShow();

				break;
			}
		}
	}
};
CDocument.prototype.controller_IsSelectionEmpty = function(bCheckHidden)
{
	if (true === this.Selection.Use)
	{
		if (selectionflag_Numbering == this.Selection.Flag)
			return false;
		else if (true === this.IsMovingTableBorder())
			return false;
		else
		{
			if (this.Selection.StartPos === this.Selection.EndPos)
				return this.Content[this.Selection.StartPos].IsSelectionEmpty(bCheckHidden);
			else
				return false;
		}
	}

	return true;
};
CDocument.prototype.controller_DrawSelectionOnPage = function(PageAbs)
{
	if (true !== this.Selection.Use)
		return;

	var Page = this.Pages[PageAbs];
	for (var SectionIndex = 0, SectionsCount = Page.Sections.length; SectionIndex < SectionsCount; ++SectionIndex)
	{
		var PageSection = Page.Sections[SectionIndex];
		for (var ColumnIndex = 0, ColumnsCount = PageSection.Columns.length; ColumnIndex < ColumnsCount; ++ColumnIndex)
		{
			this.DrawingDocument.UpdateTargetTransform(null);

			var Pos_start = PageSection.Pos;
			var Pos_end   = PageSection.EndPos;

			switch (this.Selection.Flag)
			{
				case selectionflag_Common:
				{
					var Start = this.Selection.StartPos;
					var End   = this.Selection.EndPos;

					if (Start > End)
					{
						Start = this.Selection.EndPos;
						End   = this.Selection.StartPos;
					}

					var Start = Math.max(Start, Pos_start);
					var End   = Math.min(End, Pos_end);

					for (var Index = Start; Index <= End; ++Index)
					{
						var ElementPage = this.private_GetElementPageIndex(Index, PageAbs, ColumnIndex, ColumnsCount);
						this.Content[Index].DrawSelectionOnPage(ElementPage);
					}

					if (PageAbs >= 2 && End < this.Pages[PageAbs - 2].EndPos)
					{
						this.Selection.UpdateOnRecalc = false;
						this.DrawingDocument.OnSelectEnd();
					}

					break;
				}
				case selectionflag_Numbering:
				{
					if (!this.Selection.Data)
						break;

					for (var nIndex = 0, nCount = this.Selection.Data.Paragraphs.length; nIndex < nCount; ++nIndex)
					{
						var oPara = this.Selection.Data.Paragraphs[nIndex];
						if (oPara.GetNumberingPage(true) === PageAbs)
							oPara.DrawSelectionOnPage(oPara.GetNumberingPage(false));
					}

					if (PageAbs >= 2 && this.Selection.Data[this.Selection.Data.length - 1] < this.Pages[PageAbs - 2].EndPos)
					{
						this.Selection.UpdateOnRecalc = false;
						this.DrawingDocument.OnSelectEnd();
					}

					break;
				}
			}
		}
	}
};
CDocument.prototype.controller_GetSelectionBounds = function()
{
	if (true === this.Selection.Use && selectionflag_Common === this.Selection.Flag)
	{
		var Start = this.Selection.StartPos;
		var End   = this.Selection.EndPos;

		if (Start > End)
		{
			Start = this.Selection.EndPos;
			End   = this.Selection.StartPos;
		}

		if (Start === End)
			return this.Content[Start].GetSelectionBounds();
		else
		{
			var Result       = {};
			Result.Start     = this.Content[Start].GetSelectionBounds().Start;
			Result.End       = this.Content[End].GetSelectionBounds().End;
			Result.Direction = (this.Selection.StartPos > this.Selection.EndPos ? -1 : 1);
			return Result;
		}
	}
	else if (this.Content[this.CurPos.ContentPos])
	{
		return this.Content[this.CurPos.ContentPos].GetSelectionBounds();
	}

	return null;
};
CDocument.prototype.controller_IsMovingTableBorder = function()
{
	if (null != this.Selection.Data && true === this.Selection.Data.TableBorder)
		return true;

	return false;
};
CDocument.prototype.controller_CheckPosInSelection = function(X, Y, PageAbs, NearPos)
{
	if (this.Footnotes.CheckHitInFootnote(X, Y, PageAbs) || this.Endnotes.CheckHitInEndnote(X, Y, PageAbs))
		return false;

	if (true === this.Selection.Use)
	{
		switch (this.Selection.Flag)
		{
			case selectionflag_Common:
			{
				var Start = this.Selection.StartPos;
				var End   = this.Selection.EndPos;

				if (Start > End)
				{
					Start = this.Selection.EndPos;
					End   = this.Selection.StartPos;
				}

				if (undefined !== NearPos)
				{
					for (var Index = Start; Index <= End; Index++)
					{
						if (true === this.Content[Index].CheckPosInSelection(0, 0, 0, NearPos))
							return true;
					}

					return false;
				}
				else
				{
					var ContentPos = this.Internal_GetContentPosByXY(X, Y, PageAbs, NearPos);
					if (ContentPos > Start && ContentPos < End)
					{
						return true;
					}
					else if (ContentPos < Start || ContentPos > End)
					{
						return false;
					}
					else
					{
						var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, PageAbs);
						return this.Content[ContentPos].CheckPosInSelection(X, Y, ElementPageIndex, undefined);
					}
				}
			}
			case selectionflag_Numbering:
			{
				return false;
			}
		}

		return false;
	}

	return false;
};
CDocument.prototype.controller_SelectAll = function()
{
	if (true === this.Selection.Use)
		this.RemoveSelection();

	this.DrawingDocument.SelectEnabled(true);
	this.DrawingDocument.TargetEnd();

	this.SetDocPosType(docpostype_Content);

	this.Selection.Use   = true;
	this.Selection.Start = false;
	this.Selection.Flag  = selectionflag_Common;

	this.Selection.StartPos = 0;
	this.Selection.EndPos   = this.Content.length - 1;

	for (var Index = 0; Index < this.Content.length; Index++)
	{
		this.Content[Index].SelectAll();
	}
};
CDocument.prototype.controller_GetSelectedContent = function(oSelectedContent)
{
	if (true !== this.Selection.Use || this.Selection.Flag !== selectionflag_Common)
		return;

	var StartPos = this.Selection.StartPos;
	var EndPos   = this.Selection.EndPos;
	if (StartPos > EndPos)
	{
		StartPos = this.Selection.EndPos;
		EndPos   = this.Selection.StartPos;
	}

	for (var Index = StartPos; Index <= EndPos; Index++)
	{
		this.Content[Index].GetSelectedContent(oSelectedContent);
	}

	oSelectedContent.SetLastSection(this.SectionsInfo.Get_SectPr(EndPos).SectPr);
};
CDocument.prototype.controller_UpdateCursorType = function(X, Y, PageAbs, MouseEvent)
{
	var bInText      = (null === this.IsInText(X, Y, PageAbs) ? false : true);
	var bTableBorder = (null === this.IsTableBorder(X, Y, PageAbs) ? false : true);

	// Ничего не делаем
	if (true === this.DrawingObjects.updateCursorType(PageAbs, X, Y, MouseEvent, ( true === bInText || true === bTableBorder ? true : false )))
		return;

	var ContentPos       = this.Internal_GetContentPosByXY(X, Y, PageAbs);
	var Item             = this.Content[ContentPos];
	var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, PageAbs);
	Item.UpdateCursorType(X, Y, ElementPageIndex);
};
CDocument.prototype.controller_PasteFormatting = function(oData)
{
	if (true === this.Selection.Use)
	{
		if (selectionflag_Common === this.Selection.Flag)
		{
			var Start = this.Selection.StartPos;
			var End   = this.Selection.EndPos;
			if (Start > End)
			{
				Start = this.Selection.EndPos;
				End   = this.Selection.StartPos;
			}

			for (var Pos = Start; Pos <= End; Pos++)
			{
				this.Content[Pos].PasteFormatting(oData);
			}
		}
	}
	else
	{
		this.Content[this.CurPos.ContentPos].PasteFormatting(oData);
	}
};
CDocument.prototype.controller_IsSelectionUse = function()
{
	if (true === this.Selection.Use)
		return true;

	return false;
};
CDocument.prototype.controller_IsNumberingSelection = function()
{
	return CDocumentContentBase.prototype.IsNumberingSelection.apply(this, arguments);
};
CDocument.prototype.controller_IsTextSelectionUse = function()
{
	return this.Selection.Use;
};
CDocument.prototype.controller_GetCurPosXY = function()
{
	if (true === this.Selection.Use)
	{
		if (selectionflag_Numbering === this.Selection.Flag)
			return {X : 0, Y : 0};
		else
			return this.Content[this.Selection.EndPos].GetCurPosXY();
	}
	else
	{
		return this.Content[this.CurPos.ContentPos].GetCurPosXY();
	}
};
CDocument.prototype.controller_GetSelectedText = function(bClearText, oPr)
{
	if ((true === this.Selection.Use && selectionflag_Common === this.Selection.Flag) || false === this.Selection.Use)
	{
		if (true === bClearText && this.Selection.StartPos === this.Selection.EndPos)
		{
			var Pos = ( true == this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
			return this.Content[Pos].GetSelectedText(true, oPr);
		}
		else if (false === bClearText)
		{
			var StartPos = ( true == this.Selection.Use ? Math.min(this.Selection.StartPos, this.Selection.EndPos) : this.CurPos.ContentPos );
			var EndPos   = ( true == this.Selection.Use ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : this.CurPos.ContentPos );

			var ResultText = "";

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				ResultText += this.Content[Index].GetSelectedText(false, oPr);
			}

			return ResultText;
		}
	}

	return null;
};
CDocument.prototype.controller_GetCurrentParagraph = function(bIgnoreSelection, arrSelectedParagraphs, oPr)
{
	if (null !== arrSelectedParagraphs)
	{
		if (true === this.Selection.Use)
		{
			var nStartPos = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
			var nEndPos   = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;
			for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
			{
				this.Content[nPos].GetCurrentParagraph(false, arrSelectedParagraphs, oPr);
			}
		}
		else
		{
			this.Content[this.CurPos.ContentPos].GetCurrentParagraph(false, arrSelectedParagraphs, oPr);
		}
	}
	else
	{
		let pos = this.CurPos.ContentPos;
		if (this.Selection.Use && true !== bIgnoreSelection)
		{
			pos = this.Selection.StartPos;
			if (oPr)
			{
				if (oPr.FirstInSelection)
					pos = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
				else if (oPr.LastInSelection)
					pos = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;
			}
		}
		
		if (pos < 0 || pos >= this.Content.length)
			return null;
		
		return this.Content[pos].GetCurrentParagraph(bIgnoreSelection, null, oPr);
	}
};
CDocument.prototype.controller_GetCurrentTablesStack = function(arrTables)
{
	if (true === this.Selection.Use)
	{
		if (this.Selection.StartPos !== this.Selection.EndPos)
			return arrTables;
		else
			return this.Content[this.Selection.StartPos].GetCurrentTablesStack(arrTables);
	}
	else
	{
		return this.Content[this.CurPos.ContentPos].GetCurrentTablesStack(arrTables);
	}
};
CDocument.prototype.controller_GetSelectedElementsInfo = function(oInfo)
{
	if (true === this.Selection.Use)
	{
		if (selectionflag_Numbering === this.Selection.Flag)
		{
			if (this.Selection.Data && this.Selection.Data.CurPara)
			{
				this.Selection.Data.CurPara.GetSelectedElementsInfo(oInfo);
			}
		}
		else
		{
			if (this.Selection.StartPos !== this.Selection.EndPos)
				oInfo.SetMixedSelection();

			if (oInfo.IsCheckAllSelection() || this.Selection.StartPos === this.Selection.EndPos)
			{
				var nStart = this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
				var nEnd   = this.Selection.StartPos < this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;

				for (var nPos = nStart; nPos <= nEnd; ++nPos)
				{
					this.Content[nPos].GetSelectedElementsInfo(oInfo);
				}
			}
		}
	}
	else
	{
		this.Content[this.CurPos.ContentPos].GetSelectedElementsInfo(oInfo);
	}
};
CDocument.prototype.controller_AddTableRow = function(bBefore, nCount)
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		this.Content[Pos].AddTableRow(bBefore, nCount);

		if (false === this.Selection.Use && true === this.Content[Pos].IsSelectionUse())
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = Pos;
			this.Selection.EndPos   = Pos;
		}
	}
};
CDocument.prototype.controller_AddTableColumn = function(bBefore, nCount)
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		this.Content[Pos].AddTableColumn(bBefore, nCount);

		if (false === this.Selection.Use && true === this.Content[Pos].IsSelectionUse())
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = Pos;
			this.Selection.EndPos   = Pos;
		}
	}
};
CDocument.prototype.controller_RemoveTableRow = function()
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		if (false === this.Content[Pos].RemoveTableRow())
			this.controller_RemoveTable();
	}
};
CDocument.prototype.controller_RemoveTableColumn = function()
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		if (false === this.Content[Pos].RemoveTableColumn())
			this.controller_RemoveTable();
	}
};
CDocument.prototype.controller_MergeTableCells = function()
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		this.Content[Pos].MergeTableCells();
	}
};
CDocument.prototype.controller_DistributeTableCells = function(isHorizontally)
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		return this.Content[Pos].DistributeTableCells(isHorizontally);
	}

	return false;
};
CDocument.prototype.controller_SplitTableCells = function(nCols, nRows)
{
	var nPos = true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos;
	this.Content[nPos].SplitTableCells(nCols, nRows);
};
CDocument.prototype.controller_RemoveTableCells = function()
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && !this.Content[this.Selection.StartPos].IsParagraph())
		|| (false == this.Selection.Use && !this.Content[this.CurPos.ContentPos].IsParagraph()))
	{
		var nPos = true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos;
		if (false === this.Content[nPos].RemoveTableCells())
			this.controller_RemoveTable();
	}
};
CDocument.prototype.controller_RemoveTable = function()
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		var Table = this.Content[Pos];

		if (type_Table === Table.GetType())
		{
			if (true === Table.IsInnerTable())
			{
				Table.RemoveInnerTable();
			}
			else
			{
				var isNeedRemoveTable = true;
				if (this.IsTrackRevisions())
				{
					this.Content[Pos].SelectAll();
					isNeedRemoveTable = !this.Content[Pos].RemoveTableRow();
				}

				if (isNeedRemoveTable)
				{
					this.RemoveSelection();
					Table.PreDelete();
					this.Internal_Content_Remove(Pos, 1);

					if (Pos >= this.Content.length - 1)
						Pos--;

					if (Pos < 0)
						Pos = 0;

					this.SetDocPosType(docpostype_Content);
					this.CurPos.ContentPos = Pos;
					this.Content[Pos].MoveCursorToStartPos();
				}
			}
		}
		else
		{
			Table.RemoveTable();
		}
	}
};
CDocument.prototype.controller_SelectTable = function(Type)
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		this.Content[Pos].SelectTable(Type);
		if (false === this.Selection.Use && true === this.Content[Pos].IsSelectionUse())
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = Pos;
			this.Selection.EndPos   = Pos;
		}
	}
};
CDocument.prototype.controller_CanMergeTableCells = function()
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		return this.Content[Pos].CanMergeTableCells();
	}

	return false;
};
CDocument.prototype.controller_CanSplitTableCells = function()
{
	var nPos = true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos;
	return this.Content[nPos].CanSplitTableCells();
};
CDocument.prototype.controller_UpdateInterfaceState = function()
{
	if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
		|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
	{
		this.Interface_Update_TablePr();

		if (true == this.Selection.Use)
			this.Content[this.Selection.StartPos].Document_UpdateInterfaceState();
		else
			this.Content[this.CurPos.ContentPos].Document_UpdateInterfaceState();
	}
	else
	{
		this.UpdateInterfaceParaPr();
		this.UpdateInterfaceTextPr();

		// Если у нас в выделении находится 1 параграф, или курсор находится в параграфе
		if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph == this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph == this.Content[this.CurPos.ContentPos].GetType() ) ))
		{
			if (true == this.Selection.Use)
				this.Content[this.Selection.StartPos].Document_UpdateInterfaceState();
			else
				this.Content[this.CurPos.ContentPos].Document_UpdateInterfaceState();
		}
	}
};
CDocument.prototype.controller_UpdateRulersState = function()
{
	this.Document_UpdateRulersStateBySection();

	if (true === this.Selection.Use)
	{
		var oSelection = this.private_GetSelectionPos(false);

		if (oSelection.Start === oSelection.End && type_Paragraph !== this.Content[oSelection.Start].GetType())
		{
			var PagePos = this.Get_DocumentPagePositionByContentPosition(this.GetContentPosition(true, true));

			var Page   = PagePos ? PagePos.Page : this.CurPage;
			var Column = PagePos ? PagePos.Column : 0;

			var ElementPos       = oSelection.Start;
			var Element          = this.Content[ElementPos];
			var ElementPageIndex = this.private_GetElementPageIndex(ElementPos, Page, Column, Element.Get_ColumnsCount());
			Element.Document_UpdateRulersState(ElementPageIndex);
		}
		else
		{
			var StartPos = oSelection.Start;
			var EndPos   = oSelection.End;

			var FramePr = undefined;

			for (var Pos = StartPos; Pos <= EndPos; Pos++)
			{
				var Element = this.Content[Pos];
				if (type_Paragraph != Element.GetType())
				{
					FramePr = undefined;
					break;
				}
				else
				{
					var TempFramePr = Element.Get_FramePr();
					if (undefined === FramePr)
					{
						if (undefined === TempFramePr)
							break;

						FramePr = TempFramePr;
					}
					else if (undefined === TempFramePr || false === FramePr.Compare(TempFramePr))
					{
						FramePr = undefined;
						break;
					}
				}
			}

			if (undefined !== FramePr)
				this.Content[StartPos].Document_UpdateRulersState();
		}
	}
	else
	{
		this.private_CheckCurPage();

		if (this.CurPos.ContentPos >= 0 && (null === this.FullRecalc.Id || this.FullRecalc.StartIndex > this.CurPos.ContentPos))
		{
			var PagePos = this.Get_DocumentPagePositionByContentPosition(this.GetContentPosition(false));

			var Page   = PagePos ? PagePos.Page : this.CurPage;
			var Column = PagePos ? PagePos.Column : 0;

			var ElementPos       = this.CurPos.ContentPos;
			var Element          = this.Content[ElementPos];
			var ElementPageIndex = this.private_GetElementPageIndex(ElementPos, Page, Column, Element.Get_ColumnsCount());
			Element.Document_UpdateRulersState(ElementPageIndex);
		}
	}
};
CDocument.prototype.HandleOformSelectionInEditMode = function()
{
	if (this.IsFillingFormMode())
		return false;
	
	let selectionInfo = this.GetSelectedElementsInfo();
	let form          = selectionInfo.GetInlineLevelSdt();
	
	if (!form || !form.IsForm() || form.IsComplexForm())
		return false;
	
	form.SelectContentControl();
	this.DrawingDocument.SelectEnabled(false);
	this.DrawingDocument.TargetEnd();
	return true;
};
CDocument.prototype.controller_UpdateSelectionState = function()
{
	if (this.HandleOformSelectionInEditMode())
		return;
	
	if (true === this.Selection.Use)
	{
		// Выделение нумерации
		if (selectionflag_Numbering == this.Selection.Flag)
		{
			this.DrawingDocument.TargetEnd();
			this.DrawingDocument.SelectEnabled(true);
			this.DrawingDocument.SelectShow();
		}
		// Обрабатываем движение границы у таблиц
		else if (true === this.IsMovingTableBorder())
		{
			// Убираем курсор, если он был
			this.DrawingDocument.TargetEnd();
			this.DrawingDocument.SetCurrentPage(this.CurPage);
		}
		else
		{
			if (false === this.IsSelectionEmpty() || !this.RemoveEmptySelection)
			{
				if (true !== this.Selection.Start)
				{
					this.private_CheckCurPage();
					this.RecalculateCurPos();
				}
				this.private_UpdateTracks(true, false);

				this.DrawingDocument.TargetEnd();
				this.DrawingDocument.SelectEnabled(true);
				this.DrawingDocument.SelectShow();
			}
			else
			{
				if (true !== this.Selection.Start)
				{
					this.RemoveSelection();
				}

				this.private_CheckCurPage();
				this.RecalculateCurPos();
				this.private_UpdateTracks(true, true);

				this.DrawingDocument.SelectEnabled(false);
				this.DrawingDocument.TargetStart();
				this.DrawingDocument.TargetShow();

				if (this.IsFillingFormMode())
				{
					var oContentControl = this.GetContentControl();
					if (oContentControl && oContentControl.IsCheckBox())
						this.DrawingDocument.TargetEnd();
				}
			}
		}
	}
	else
	{
		this.RemoveSelection();
		this.private_CheckCurPage();
		this.RecalculateCurPos();
		this.private_UpdateTracks(false, false);

		this.DrawingDocument.SelectEnabled(false);
		this.DrawingDocument.TargetShow();

		if (this.IsFillingFormMode())
		{
			var oContentControl = this.GetContentControl();
			if (oContentControl && oContentControl.IsCheckBox())
				this.DrawingDocument.TargetEnd();
		}
	}
};
CDocument.prototype.controller_GetSelectionState = function()
{
	var State;
	if (true === this.Selection.Use)
	{
		if (this.controller_IsNumberingSelection() || this.controller_IsMovingTableBorder())
		{
			State = [];
		}
		else
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (StartPos > EndPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			State = [];

			var TempState = [];
			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				TempState.push(this.Content[Index].GetSelectionState());
			}

			State.push(TempState);
		}
	}
	else
	{
		State = this.Content[this.CurPos.ContentPos].GetSelectionState();
	}

	return State;
};
CDocument.prototype.controller_SetSelectionState = function(State, StateIndex)
{
	if (true === this.Selection.Use)
	{
		// Выделение нумерации
		if (selectionflag_Numbering == this.Selection.Flag)
		{
			if (type_Paragraph === this.Content[this.Selection.StartPos].Get_Type())
			{
				var NumPr = this.Content[this.Selection.StartPos].GetNumPr();
				if (undefined !== NumPr)
					this.SelectNumbering(NumPr, this.Content[this.Selection.StartPos]);
				else
					this.RemoveSelection();
			}
			else
				this.RemoveSelection();
		}
		else
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (StartPos > EndPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			var CurState = State[StateIndex];
			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				this.Content[Index].SetSelectionState(CurState[Index - StartPos], CurState[Index - StartPos].length - 1);
			}
		}
	}
	else
	{
		this.Content[this.CurPos.ContentPos].SetSelectionState(State, StateIndex);
	}
};
CDocument.prototype.controller_AddHyperlink = function(Props)
{
	if (false === this.Selection.Use || this.Selection.StartPos === this.Selection.EndPos)
	{
		var Pos = ( true == this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
		return this.Content[Pos].AddHyperlink(Props);
	}
	
	return null;
};
CDocument.prototype.controller_ModifyHyperlink = function(Props)
{
	if (false == this.Selection.Use || this.Selection.StartPos == this.Selection.EndPos)
	{
		var Pos = ( true == this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
		this.Content[Pos].ModifyHyperlink(Props);
	}
};
CDocument.prototype.controller_RemoveHyperlink = function()
{
	if (false == this.Selection.Use || this.Selection.StartPos == this.Selection.EndPos)
	{
		var Pos = ( true == this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
		this.Content[Pos].RemoveHyperlink();
	}
};
CDocument.prototype.controller_CanAddHyperlink = function(bCheckInHyperlink)
{
	if (true === this.Selection.Use)
	{
		if (selectionflag_Common === this.Selection.Flag)
		{
			if (this.Selection.StartPos != this.Selection.EndPos)
				return false;

			return this.Content[this.Selection.StartPos].CanAddHyperlink(bCheckInHyperlink);
		}
	}
	else
	{
		return this.Content[this.CurPos.ContentPos].CanAddHyperlink(bCheckInHyperlink);
	}

	return false;
};
CDocument.prototype.controller_IsCursorInHyperlink = function(bCheckEnd)
{
	if (true === this.Selection.Use)
	{
		if (selectionflag_Common === this.Selection.Flag)
		{
			if (this.Selection.StartPos != this.Selection.EndPos)
				return null;

			return this.Content[this.Selection.StartPos].IsCursorInHyperlink(bCheckEnd);
		}
	}
	else
	{
		return this.Content[this.CurPos.ContentPos].IsCursorInHyperlink(bCheckEnd);
	}

	return null;
};
CDocument.prototype.controller_AddComment = function(Comment)
{
	if (selectionflag_Numbering === this.Selection.Flag)
		return;

	if (true === this.Selection.Use)
	{
		var StartPos, EndPos;
		if (this.Selection.StartPos < this.Selection.EndPos)
		{
			StartPos = this.Selection.StartPos;
			EndPos   = this.Selection.EndPos;
		}
		else
		{
			StartPos = this.Selection.EndPos;
			EndPos   = this.Selection.StartPos;
		}

		if (StartPos === EndPos)
		{
			this.Content[StartPos].AddComment(Comment, true, true);
		}
		else
		{
			this.Content[StartPos].AddComment(Comment, true, false);
			this.Content[EndPos].AddComment(Comment, false, true);
		}
	}
	else
	{
		this.Content[this.CurPos.ContentPos].AddComment(Comment, true, true);
	}
};
CDocument.prototype.controller_CanAddComment = function()
{
	if (selectionflag_Common === this.Selection.Flag)
	{
		if (true === this.Selection.Use && this.Selection.StartPos != this.Selection.EndPos)
		{
			return true;
		}
		else
		{
			var Pos     = ( this.Selection.Use === true ? this.Selection.StartPos : this.CurPos.ContentPos );
			var Element = this.Content[Pos];
			return Element.CanAddComment();
		}
	}

	return false;
};
CDocument.prototype.controller_GetSelectionAnchorPos = function()
{
	var Pos = ( true === this.Selection.Use ? ( this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos ) : this.CurPos.ContentPos );
	return this.Content[Pos].GetSelectionAnchorPos();
};
CDocument.prototype.controller_StartSelectionFromCurPos = function()
{
	this.Selection.StartPos = this.CurPos.ContentPos;
	this.Selection.EndPos   = this.CurPos.ContentPos;
	this.Content[this.CurPos.ContentPos].StartSelectionFromCurPos();
};
CDocument.prototype.controller_SaveDocumentStateBeforeLoadChanges = function(State)
{
	State.Pos      = this.GetContentPosition(false, false);
	State.StartPos = this.GetContentPosition(true, true);
	State.EndPos   = this.GetContentPosition(true, false);
};
CDocument.prototype.controller_RestoreDocumentStateAfterLoadChanges = function(State)
{
	if (true === this.Selection.Use)
	{
		this.SetContentPosition(State.StartPos, 0, 0);
		this.SetContentSelection(State.StartPos, State.EndPos, 0, 0, 0);
	}
	else
	{
		this.SetContentPosition(State.Pos, 0, 0);
		this.NeedUpdateTarget = true;
	}
};
CDocument.prototype.controller_GetColumnSize = function()
{
	var nContentPos = true === this.Selection.Use ? ( this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos ) : this.CurPos.ContentPos;

	var oPagePos   = this.Get_DocumentPagePositionByContentPosition(this.GetContentPosition(this.Selection.Use, false));
	var nColumnAbs = oPagePos ? oPagePos.Column : 0;
	var nPageAbs   = oPagePos ? oPagePos.Page : 0;

	var oSectPr = this.Get_SectPr(nContentPos);
	var oFrame  = oSectPr.GetContentFrame(nPageAbs);

	var Y      = oFrame.Top;
	var YLimit = oFrame.Bottom;
	var X      = oFrame.Left;
	var XLimit = oFrame.Right;

	var ColumnsCount = oSectPr.GetColumnsCount();
	for (var ColumnIndex = 0; ColumnIndex < nColumnAbs; ++ColumnIndex)
	{
		X += oSectPr.GetColumnWidth(ColumnIndex);
		X += oSectPr.GetColumnSpace(ColumnIndex);
	}

	if (ColumnsCount - 1 !== nColumnAbs)
		XLimit = X + oSectPr.GetColumnWidth(nColumnAbs);

	return {
		W : XLimit - X,
		H : YLimit - Y
	};
};
CDocument.prototype.controller_GetCurrentSectionPr = function()
{
	var nContentPos = this.CurPos.ContentPos;
	return this.SectionsInfo.Get_SectPr(nContentPos).SectPr;
};
CDocument.prototype.controller_AddContentControl = function(nContentControlType)
{
	return this.private_AddContentControl(nContentControlType);
};
CDocument.prototype.controller_GetStyleFromFormatting = function()
{
	if (true == this.Selection.Use)
	{
		if (this.Selection.StartPos > this.Selection.EndPos)
			return this.Content[this.Selection.EndPos].GetStyleFromFormatting();
		else
			return this.Content[this.Selection.StartPos].GetStyleFromFormatting();
	}
	else
	{
		return this.Content[this.CurPos.ContentPos].GetStyleFromFormatting();
	}
};
CDocument.prototype.controller_GetSimilarNumbering = function(oContinueEngine)
{
	this.GetSimilarNumbering(oContinueEngine);
};
CDocument.prototype.controller_GetPlaceHolderObject = function()
{
	var nCurPos = this.CurPos.ContentPos;
	if (this.Selection.Use)
	{
		if (this.Selection.StartPos === this.Selection.EndPos)
			nCurPos = this.Selection.StartPos;
		else
			return null;
	}

	return this.Content[nCurPos].GetPlaceHolderObject();
};
CDocument.prototype.controller_GetAllFields = function(isUseSelection, arrFields)
{
	if (!arrFields)
		arrFields = [];

	var nStartPos = isUseSelection ?
		(this.Selection.Use ?
			(this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos)
			: this.CurPos.ContentPos)
		: 0;

	var nEndPos = isUseSelection ?
		(this.Selection.Use ?
			(this.Selection.StartPos < this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos)
			: this.CurPos.ContentPos)
		: this.Content.length - 1;

	for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
	{
		this.Content[nIndex].GetAllFields(isUseSelection, arrFields);
	}

	return arrFields;
};
CDocument.prototype.controller_IsTableCellSelection = function()
{
	return (this.Selection.Use && this.Selection.StartPos === this.Selection.EndPos && this.Content[this.Selection.StartPos].IsTable() && this.Content[this.Selection.StartPos].IsTableCellSelection());
};
CDocument.prototype.controller_CollectSelectedReviewChanges = function(oTrackManager)
{
	var isUseSelection = this.Selection.Use;

	var nStartPos = isUseSelection ? this.Selection.StartPos : this.CurPos.ContentPos;
	var nEndPos   = isUseSelection ? this.Selection.EndPos   : this.CurPos.ContentPos;
	if (nStartPos > nEndPos)
	{
		var nTemp = nStartPos;
		nStartPos = nEndPos;
		nEndPos   = nTemp;
	}

	for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		this.Content[nPos].CollectSelectedReviewChanges(oTrackManager);
	}
};
//----------------------------------------------------------------------------------------------------------------------
//
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.RemoveTextSelection = function()
{
	this.Controller.RemoveTextSelection();
};
CDocument.prototype.AddFormTextField = function(sName, sDefaultText)
{
	if (false === this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_AddMailMergeField);

		var oField = new ParaField(fieldtype_FORMTEXT);
		var oRun = new ParaRun();
		oField.SetFormFieldName(sName);
		oField.SetFormFieldDefaultText(sDefaultText);
		oRun.AddText(sDefaultText);
		oField.Add_ToContent(0, oRun);

		this.Register_Field(oField);
		this.AddToParagraph(oField);
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
CDocument.prototype.GetAllFormTextFields = function()
{
	return this.FieldsManager.GetAllFieldsByType(fieldtype_FORMTEXT);
};
CDocument.prototype.IsFillingFormMode = function()
{
	return this.Api && this.Api.isRestrictionForms();
};
CDocument.prototype.IsFillingOFormMode = function()
{
	if (!this.IsFillingFormMode())
		return false;

	var oApi = this.GetApi();
	return !!(oApi.DocInfo &&  oApi.DocInfo.Format && (this.Api.DocInfo.Format === "oform" || this.Api.DocInfo.Format === "docxf"));
};
CDocument.prototype.CheckOFormUserMaster = function(form)
{
	let currentUserMaster = this.GetCurrentOFormUserMaster();
	if (!currentUserMaster || !form || !form.IsForm())
		return true;
	
	if (!form.IsMainForm())
		form = form.GetMainForm();
	
	let fieldMaster = form.GetFieldMaster();
	if (!fieldMaster)
		return false;
	
	return fieldMaster.getFirstUser() === currentUserMaster;
};
CDocument.prototype.IsInFormField = function(isAllowComplexForm, isCheckCurrentUser)
{
	var oSelectedInfo = this.GetSelectedElementsInfo();
	var oField        = oSelectedInfo.GetField();
	var oInlineSdt    = oSelectedInfo.GetInlineLevelSdt();
	var oBlockSdt     = oSelectedInfo.GetBlockLevelSdt();
	
	if (isCheckCurrentUser
		&& oInlineSdt
		&& (!oInlineSdt.IsComplexForm() || isAllowComplexForm))
		return this.CheckOFormUserMaster(oInlineSdt);

	// оInlineSdt отдает нам нижний уровень, если на нем у нас ComplexField, значит мы находимся внутри текста,
	// в такой ситуации мы отдаем, что ме не находимся в форме, чтобы запретить редактирование в этой части формы
	// Когда мы будем находится внутри простой формы, находящейся в сложной, то oInlineSdt вернет именно проостую форму
	return !!(oBlockSdt || (oInlineSdt && (!oInlineSdt.IsComplexForm() || isAllowComplexForm)) || (oField && fieldtype_FORMTEXT === oField.Get_FieldType()));
};
CDocument.prototype.IsFormFieldEditing = function()
{
	return (this.IsFillingFormMode() && this.IsInFormField(false, true));
};
CDocument.prototype.MoveToFillingForm = function(isNext)
{
	var oInfo = this.GetSelectedElementsInfo();
	var oForm;
	if ((oForm = oInfo.GetInlineLevelSdt()) || (oForm = oInfo.GetBlockLevelSdt()))
		oForm.SelectContentControl();
	else if ((oForm = oInfo.GetField()) instanceof ParaField)
		oForm.SelectThisElement();

	this.DrawingObjects.resetDrawStateBeforeAction();
	var oRes = null;
	if (docpostype_DrawingObjects === this.GetDocPosType())
	{
		var oParaDrawing = this.DrawingObjects.getMajorParaDrawing();

		oRes = oParaDrawing.FindNextFillingForm(isNext, true);
		if (!oRes)
		{
			this.DrawingObjects.resetSelection();
			oParaDrawing.GoTo_Text(true !== isNext, false);

			// В случаях, когда у нас автофигура внутри другой автофигуры
			var arrPassedParagraphs = [];
			var oParagraph = this.GetCurrentParagraph();
			var oShape     = oParagraph.GetParent() ? oParagraph.GetParent().Is_DrawingShape(true) : null;

			while (oShape)
			{
				oParaDrawing = oShape.parent;
				var isBreak = false;
				for (var nIndex = 0, nCount = arrPassedParagraphs.length; nIndex < nCount; ++nIndex)
				{
					if (arrPassedParagraphs[nIndex] === oParagraph)
					{
						isBreak = true;
						break;
					}
				}

				if (isBreak)
					break;

				arrPassedParagraphs.push(oParagraph);
				oRes = oParaDrawing.FindNextFillingForm(isNext, true);

				if (oRes === oForm)
					oRes = null;

				if (oRes)
				{
					break;
				}
				else
				{
					oParaDrawing.GoTo_Text(true !== isNext, false);
					oParagraph = this.GetCurrentParagraph();
					oShape     = oParagraph.GetParent() ? oParagraph.GetParent().Is_DrawingShape(true) : null;
				}

			}
		}
	}

	if (!oRes)
	{
		var nDocPosType = this.GetDocPosType();
		if (docpostype_Content === nDocPosType)
		{
			oRes = this.FindNextFillingForm(isNext, true, true);

			if (!oRes)
				oRes = this.Footnotes.FindNextFillingForm(isNext, false);

			if (!oRes)
				oRes = this.Endnotes.FindNextFillingForm(isNext, false);

			if (!oRes)
				oRes = this.SectionsInfo.FindNextFillingForm(isNext, null);

			if (!oRes)
				oRes = this.FindNextFillingForm(isNext, false, false);
		}
		else if (docpostype_HdrFtr === nDocPosType)
		{
			oRes = this.SectionsInfo.FindNextFillingForm(isNext, this.HdrFtr.GetCurHdrFtr());

			if (!oRes)
				oRes = this.FindNextFillingForm(isNext, false, false);

			if (!oRes)
				oRes = this.Footnotes.FindNextFillingForm(isNext, false);

			if (!oRes)
				oRes = this.Endnotes.FindNextFillingForm(isNext, false);

			if (!oRes)
				oRes = this.SectionsInfo.FindNextFillingForm(isNext, null);
		}
		else if (docpostype_Footnotes === nDocPosType)
		{
			oRes = this.Footnotes.FindNextFillingForm(isNext, true);

			if (!oRes)
				oRes = this.Endnotes.FindNextFillingForm(isNext, false);

			if (!oRes)
				oRes = this.SectionsInfo.FindNextFillingForm(isNext, null);

			if (!oRes)
				oRes = this.FindNextFillingForm(isNext, false, false);

			if (!oRes)
				oRes = this.Footnotes.FindNextFillingForm(isNext, false);
		}
		else if (docpostype_Endnotes === nDocPosType)
		{
			oRes = this.Endnotes.FindNextFillingForm(isNext, true);

			if (!oRes)
				oRes = this.SectionsInfo.FindNextFillingForm(isNext, null);

			if (!oRes)
				oRes = this.FindNextFillingForm(isNext, false, false);

			if (!oRes)
				oRes = this.Footnotes.FindNextFillingForm(isNext, false);

			if (!oRes)
				oRes = this.Endnotes.FindNextFillingForm(isNext, false);
		}
	}

	if (oRes)
	{
		this.RemoveSelection();

		if (oRes instanceof CBlockLevelSdt || oRes instanceof CInlineLevelSdt)
		{
			oRes.SelectContentControl();
		}
		else if (oRes instanceof ParaField)
		{
			oRes.SelectThisElement();
		}
	}
};
CDocument.prototype.TurnComboBoxFormValue = function(oForm, isNext)
{
	if (!(oForm instanceof CInlineLevelSdt) || (!oForm.IsComboBox() && !oForm.IsDropDownList()))
		return;

	var sValue = oForm.GetSelectedText(true, true);
	var oComboBoxPr = oForm.IsComboBox() ? oForm.GetComboBoxPr() : oForm.GetDropDownListPr();

	var nIndex = oComboBoxPr.FindByText(sValue);
	var nCount = oComboBoxPr.GetItemsCount();
	if (nCount <= 0)
		return;

	var nNewIndex = 0;
	if (-1 !== nIndex)
	{
		if (isNext)
		{
			nNewIndex = nIndex + 1;
			if (nNewIndex >= nCount)
				nNewIndex = 0;
		}
		else
		{
			nNewIndex = nIndex - 1;
			if (nNewIndex < 0)
				nNewIndex = nCount - 1;
		}
	}

	var sNewValue = oComboBoxPr.GetItemValue(nNewIndex);
	oForm.SelectContentControl();
	oForm.SkipSpecialContentControlLock(true);
	if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content, null, true, this.IsFormFieldEditing()))
	{
		this.StartAction();
		oForm.SelectListItem(sNewValue);
		oForm.SelectContentControl();
		this.Recalculate();
		this.FinalizeAction();
	}
	oForm.SkipSpecialContentControlLock(false);
};
CDocument.prototype.OnContentControlTrackEnd = function(Id, NearestPos, isCopy)
{
	return this.OnEndTextDrag(NearestPos, isCopy);
};
CDocument.prototype.AddContentControl = function(nContentControlType)
{
	if (this.IsDrawingSelected() && !this.DrawingObjects.getTargetDocContent())
	{
		this.DrawingObjects.resetDrawStateBeforeAction();
		var oDrawing = this.DrawingObjects.getMajorParaDrawing();
		if (oDrawing)
			oDrawing.SelectAsText();
	}

	return this.Controller.AddContentControl(nContentControlType);
};
CDocument.prototype.GetAllContentControls = function()
{
	var arrContentControls = [];

	this.SectionsInfo.GetAllContentControls(arrContentControls);

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetAllContentControls(arrContentControls);
	}
	return arrContentControls;
};
CDocument.prototype.RemoveContentControl = function(Id)
{
	var oContentControl = this.TableId.Get_ById(Id);
	if (!oContentControl || !oContentControl.GetContentControlType)
		return;
	
	// TODO: По хорошему надо сделать метод у CStdBase и перенести реализации в соответствующие классы
	this.RemoveSelection();
	if (oContentControl.IsBlockLevel())
	{
		if (!oContentControl.Parent)
			return;
		
		var oDocContent = oContentControl.Parent;
		oDocContent.Update_ContentIndexing();
		var nIndex = oContentControl.GetIndex();

		var nCurPos = oDocContent.CurPos.ContentPos;
		oDocContent.Remove_FromContent(nIndex, 1);

		if (nIndex === nCurPos)
		{
			if (nIndex >= oDocContent.GetElementsCount())
			{
				oDocContent.MoveCursorToEndPos();
			}
			else
			{
				oDocContent.CurPos.ContentPos = Math.max(0, Math.min(oDocContent.GetElementsCount() - 1, nIndex));;
				oDocContent.Content[oDocContent.CurPos.ContentPos].MoveCursorToStartPos();
			}
		}
		else if (nIndex < nCurPos)
		{
			oDocContent.CurPos.ContentPos = Math.max(0, Math.min(oDocContent.GetElementsCount() - 1, nCurPos - 1));
		}
	}
	else if (oContentControl.IsInlineLevel())
	{
		oContentControl.RemoveThisFromParent(true);
	}
};
CDocument.prototype.RemoveContentControlWrapper = function(Id)
{
	if (!this.CanPerformAction())
		return;
	
	let contentControl = this.TableId.Get_ById(Id);
	if (!contentControl)
		return;
	
	this.StartAction();
	
	if (contentControl.IsForm())
		contentControl.RemoveFormPr();
	
	if (contentControl.IsPlaceHolder())
		this.RemoveContentControl(Id);
	else
		contentControl.RemoveContentControlWrapper();
	
	this.Recalculate();
	this.UpdateInterface();
	this.UpdateSelection();
	this.FinalizeAction();
};
CDocument.prototype.GetContentControl = function(Id)
{
	if (undefined === Id || null === Id)
	{
		var oInfo          = this.GetSelectedElementsInfo({SkipTOC : true});
		var oInlineControl = oInfo.GetInlineLevelSdt();
		var oBlockControl  = oInfo.GetBlockLevelSdt();

		if (oInlineControl)
			return oInlineControl;
		else if (oBlockControl)
			return oBlockControl;

		return null;
	}

	return this.TableId.Get_ById(Id);
};
CDocument.prototype.ClearContentControl = function(Id)
{
	var oContentControl = this.TableId.Get_ById(Id);
	if (!oContentControl)
		return null;

	this.RemoveSelection();

	if (oContentControl.GetContentControlType
		&& (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType()
		|| c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType()))
	{
		oContentControl.ClearContentControl();
		oContentControl.SetThisElementCurrent();
		oContentControl.MoveCursorToStartPos();
	}

	return oContentControl;
};
/**
 * Выделяем содержимое внутри заданного контейнера
 * @param {string} sId
 */
CDocument.prototype.SelectContentControl = function(sId)
{
	var oContentControl = this.TableId.Get_ById(sId);
	if (!oContentControl)
		return;

	if (oContentControl.GetContentControlType
		&& (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType()
		|| c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType()))
	{
		this.CheckFormPlaceHolder = false;
		this.RemoveSelection();
		oContentControl.SelectContentControl();
		this.UpdateSelection();
		this.UpdateRulers();
		this.UpdateInterface();
		this.UpdateTracks();

		this.private_UpdateCursorXY(true, true);
		this.CheckFormPlaceHolder = true;
	}
};
/**
 * Передвигаем курсор в заданный блочный элемент
 * @param {string} sId
 * @param {boolean} [isBegin=true]
 */
CDocument.prototype.MoveCursorToContentControl = function(sId, isBegin)
{
	var oContentControl = this.TableId.Get_ById(sId);
	if (!oContentControl)
		return;

	if (oContentControl.GetContentControlType
		&& (c_oAscSdtLevelType.Block === oContentControl.GetContentControlType()
		|| c_oAscSdtLevelType.Inline === oContentControl.GetContentControlType()))
	{
		this.RemoveSelection();

		if (false !== isBegin)
			isBegin = true;

		oContentControl.MoveCursorToContentControl(isBegin);
		this.UpdateSelection();
		this.UpdateRulers();
		this.UpdateInterface();
		this.UpdateTracks();

		this.private_UpdateCursorXY(true, true);
	}
};
CDocument.prototype.GetAllSignatures = function()
{
    var aSignatures = [];
    this.CheckRunContent(function(oRun)
    {
        var aContent = oRun.Content;
        if(Array.isArray(aContent))
        {
            var nItem, oItem, oSp;
            for(nItem = 0; nItem < aContent.length; ++nItem)
            {
                oItem = aContent[nItem];
                if(oItem && oItem.GraphicObj)
                {
                    oSp = oItem.GraphicObj;
                    if(oSp.signatureLine)
                    {
                        aSignatures.push(oSp.signatureLine);
                    }
                }
            }
        }
    });
    return aSignatures;
};
CDocument.prototype.CallSignatureDblClickEvent = function(sGuid)
{
    var ret = [], allSpr = [];
    allSpr = allSpr.concat(allSpr.concat(this.DrawingObjects.getAllSignatures2(ret, this.DrawingObjects.getDrawingArray())));
    for(var i = 0; i < allSpr.length; ++i)
    {
        if(allSpr[i].getSignatureLineGuid() === sGuid)
        {
            this.Api.sendEvent("asc_onSignatureDblClick", sGuid, allSpr[i].extX, allSpr[i].extY);
        }
    }
};
CDocument.prototype.SetCheckContentControlsLock = function(isLocked)
{
	this.CheckContentControlsLock = isLocked;
};
CDocument.prototype.IsCheckContentControlsLock = function()
{
	return this.CheckContentControlsLock;
};
CDocument.prototype.IsEditCommentsMode = function()
{
	return this.Api.isRestrictionComments();
};
CDocument.prototype.IsViewMode = function()
{
	return this.Api.getViewMode();
};
CDocument.prototype.IsEditSignaturesMode = function()
{
	return this.Api.isRestrictionSignatures();
};
CDocument.prototype.IsViewModeInEditor = function()
{
	return this.Api.isRestrictionView();
};
CDocument.prototype.CanEdit = function()
{
	return this.Api.canEdit();
};
CDocument.prototype.private_CheckCursorPosInFillingFormMode = function()
{
	// В oform не сдвигаемся автоматически
	if (this.IsFillingFormMode() && !this.IsInFormField(true) && !this.IsFillingOFormMode())
	{
		this.MoveToFillingForm(true);
		this.UpdateSelection();
		this.UpdateInterface();
	}
};
CDocument.prototype.OnEndLoadScript = function()
{
	this.UpdateAllSectionsInfo();
	this.Check_SectionLastParagraph();
	this.Styles.Check_StyleNumberingOnLoad(this.Numbering);

	var arrParagraphs = this.GetAllParagraphs({All : true});
	for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
	{
		arrParagraphs[nIndex].Recalc_CompiledPr();
		arrParagraphs[nIndex].Recalc_RunsCompiledPr();
	}
};
CDocument.prototype.BeginViewModeInReview = function(nMode)
{
	if (this.IsViewModeInReview())
		this.EndViewModeInReview();

	this.ViewModeInReview.mode = undefined !== nMode ? nMode : Asc.c_oAscDisplayModeInReview.Final;

	this.ViewModeInReview.isFastCollaboration = this.CollaborativeEditing.Is_Fast();
	this.CollaborativeEditing.Set_Fast(false);

	this.History.SaveRedoPoints();
	if (this.ViewModeInReview.mode === Asc.c_oAscDisplayModeInReview.Final)
		this.AcceptAllRevisionChanges(true, false);
	else
		this.RejectAllRevisionChanges(true, false);

	this.CollaborativeEditing.Set_GlobalLock(true);
};
CDocument.prototype.EndViewModeInReview = function(nMode)
{
	if (!this.IsViewModeInReview())
	{
		this.ViewModeInReview.mode = undefined !== nMode ? nMode : Asc.c_oAscDisplayModeInReview.Edit;
		return;
	}

	this.CollaborativeEditing.Set_GlobalLock(false);

	this.Document_Undo();
	this.History.Clear_Redo();
	this.History.PopRedoPoints();

	if (this.ViewModeInReview.isFastCollaboration)
		this.CollaborativeEditing.Set_Fast(true);

	this.ViewModeInReview.mode = undefined !== nMode ? nMode : Asc.c_oAscDisplayModeInReview.Edit;
};
CDocument.prototype.IsViewModeInReview = function()
{
	return (Asc.c_oAscDisplayModeInReview.Final === this.ViewModeInReview.mode
		|| Asc.c_oAscDisplayModeInReview.Original === this.ViewModeInReview.mode);
};
CDocument.prototype.IsSimpleMarkupInReview = function()
{
	return (this.ViewModeInReview.mode === Asc.c_oAscDisplayModeInReview.Simple);
};
CDocument.prototype.SetDisplayModeInReview = function(nMode, isSendEvent)
{
	// У нас есть два режима редактирования Edit, Simple и два режима просмотра Final, Original
	if (this.ViewModeInReview.mode === nMode)
		return;

	if (isSendEvent)
		this.Api.sendEvent("asc_onChangeDisplayModeInReview", nMode);

	if (Asc.c_oAscDisplayModeInReview.Edit === nMode || Asc.c_oAscDisplayModeInReview.Simple === nMode)
	{
		this.EndViewModeInReview(nMode);
	}
	else
	{
		this.BeginViewModeInReview(nMode);
	}

	this.UpdateInterface();
};
CDocument.prototype.GetDisplayModeInReview = function()
{
	return this.ViewModeInReview.mode;
};
CDocument.prototype.StartCollaborationEditing = function()
{
	if (this.DrawingDocument)
	{
		this.DrawingDocument.Start_CollaborationEditing();

		if (this.IsViewModeInReview())
			this.SetDisplayModeInReview(Asc.c_oAscDisplayModeInReview.Edit, true);
	}
};
CDocument.prototype.AddField = function(nType, oPr)
{
	if (fieldtype_PAGENUM === nType)
	{
		return this.AddFieldWithInstruction("PAGE");
	}
	else if (fieldtype_TOC === nType)
	{
		return this.AddFieldWithInstruction("TOC");
	}
	else if (fieldtype_PAGEREF === nType)
	{
		return this.AddFieldWithInstruction("PAGEREF Test \\p");
	}

	return false;
};
CDocument.prototype.AddFieldWithInstruction = function(sInstruction)
{
	var oParagraph = this.GetCurrentParagraph(false, false, {ReplacePlaceHolder : true});
	if (!oParagraph)
		return null;

    var oBeginChar    = new ParaFieldChar(fldchartype_Begin, this),
        oSeparateChar = new ParaFieldChar(fldchartype_Separate, this),
        oEndChar      = new ParaFieldChar(fldchartype_End, this);

    var oRun = new ParaRun();
    oRun.AddToContent(-1, oBeginChar);
    oRun.AddInstrText(sInstruction);
    oRun.AddToContent(-1, oSeparateChar);
    oRun.AddToContent(-1, oEndChar);
    oParagraph.Add(oRun);

    oBeginChar.SetRun(oRun);
    oSeparateChar.SetRun(oRun);
    oEndChar.SetRun(oRun);

    var oComplexField = oBeginChar.GetComplexField();
    oComplexField.SetBeginChar(oBeginChar);
    oComplexField.SetInstructionLine(sInstruction);
    oComplexField.SetSeparateChar(oSeparateChar);
    oComplexField.SetEndChar(oEndChar);
    oComplexField.Update(false);
    return oComplexField;
};
CDocument.prototype.AddDateTime = function(oPr)
{
	if (!oPr)
		return;

	var nLang = oPr.get_Lang();
	if (!AscFormat.isRealNumber(nLang))
		nLang = 1033;

	if (oPr.get_Update())
	{
		if (!this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			this.StartAction(AscDFH.historydescription_Document_AddDateTimeField);

			var oComplexField = this.AddFieldWithInstruction("TIME \\@ \"" + oPr.get_Format() + "\"");

			var oBeginChar = oComplexField.GetBeginChar();
			var oRun       = oBeginChar.GetRun();
			oRun.Set_Lang_Val(nLang);

			this.Recalculate();
			this.UpdateInterface();
			this.UpdateSelection();
			this.FinalizeAction();
		}
	}
	else
	{
		let oSettings = new AscCommon.CAddTextSettings();
		let oTextPr   = new CTextPr();
		oTextPr.Set_FromObject({Lang : {Val : nLang}});
		oSettings.SetTextPr(oTextPr);
		oSettings.MoveCursorOutside(true);

		this.AddTextWithPr(oPr.get_String(), oSettings);
	}
};
CDocument.prototype.ValidateComplexField = function(oComplexField)
{
	if (this.IsLoadingDocument()
		|| !this.IsActionStarted())
		return;

	if (!this.Action.Additional.ValidateComplexFields)
		this.Action.Additional.ValidateComplexFields = [];

	for (var nIndex = 0, nCount = this.Action.Additional.ValidateComplexFields.length; nIndex < nCount; ++nIndex)
	{
		if (oComplexField === this.Action.Additional.ValidateComplexFields[nIndex])
			return;
	}

	this.Action.Additional.ValidateComplexFields.push(oComplexField);
};
CDocument.prototype.private_FinalizeValidateComplexFields = function()
{
	for (var nIndex = 0, nCount = this.Action.Additional.ValidateComplexFields.length; nIndex < nCount; ++nIndex)
	{
		var oComplexField = this.Action.Additional.ValidateComplexFields[nIndex];
		if (!oComplexField.IsValid())
		{
			oComplexField.RemoveFieldChars();
			this.Action.Recalculate = true;
		}
	}
};
CDocument.prototype.AddRefToParagraph = function(oParagraph, nType, bHyperlink, bAboveBelow, sSeparator)
{
	if(false === this.IsSelectionLocked(AscCommon.changestype_Document_Content, {
		Type      : changestype_2_ElementsArray_and_Type,
		Elements  : [oParagraph],
		CheckType : changestype_Paragraph_Content
	}))
	{
		this.StartAction(AscDFH.historydescription_Document_AddCrossRef);
		var sBookmarkName = oParagraph.AddBookmarkForRef();
		if(sBookmarkName)
		{
			this.private_AddRefToBookmark(sBookmarkName, nType, bHyperlink, bAboveBelow, sSeparator);
		}
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
};
CDocument.prototype.AddRefToBookmark = function(sBookmarkName, nType, bHyperlink, bAboveBelow, sSeparator)
{
	if(false === this.IsSelectionLocked(AscCommon.changestype_Document_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_AddCrossRef);
		this.private_AddRefToBookmark(sBookmarkName, nType, bHyperlink, bAboveBelow, sSeparator);
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
};
CDocument.prototype.private_AddRefToBookmark = function(sBookmarkName, nType, bHyperlink, bAboveBelow, sSeparator)
{
	if(!(typeof sBookmarkName === "string" && sBookmarkName.length > 0))
	{
		return;
	}
	var sInstr = "";
	var sSuffix = "";
	if(bHyperlink)
	{
		sSuffix += " \\h";
	}
	if(bAboveBelow && nType !== Asc.c_oAscDocumentRefenceToType.AboveBelow)
	{
		sSuffix += " \\p";
	}
	if(typeof sSeparator === "string" && sSeparator.length > 0)
	{
		sSuffix += " \\d " + sSeparator;
	}
	switch (nType)
	{
		case Asc.c_oAscDocumentRefenceToType.PageNum:
		{
			sInstr = " PAGEREF " + sBookmarkName;
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.Text:
		case Asc.c_oAscDocumentRefenceToType.OnlyCaptionText:
		case Asc.c_oAscDocumentRefenceToType.OnlyLabelAndNumber:
		{
			sInstr = " REF " + sBookmarkName + " ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.ParaNum:
		{
			sInstr = " REF " + sBookmarkName + " \\r ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.ParaNumNoContext:
		{
			sInstr = " REF " + sBookmarkName + " \\n ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.ParaNumFullContex:
		{
			sInstr = " REF " + sBookmarkName + " \\w ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.AboveBelow:
		{
			sInstr = " REF " + sBookmarkName + " \\p ";
            sInstr += sSuffix;
			break;
		}
	}
	var oComplexField = this.AddFieldWithInstruction(sInstr);
    if(nType === Asc.c_oAscDocumentRefenceToType.PageNum)
    {
        this.FullRecalc.Continue = false;
        this.FullRecalc.UseRecursion = false;
        this.private_Recalculate(undefined, true);
        while (this.FullRecalc.Continue)
        {
            this.FullRecalc.Continue = false;
            this.Recalculate_Page();
        }
        this.FullRecalc.UseRecursion = true;
        oComplexField.Update(false);
    }
    return oComplexField;
};
CDocument.prototype.AddNoteRefToParagraph = function(oParagraph, nType, bHyperlink, bAboveBelow)
{
	if(!oParagraph.Parent)
	{
		return;
	}
	var oFootnote = oParagraph.Parent.IsFootnote(true);
	if(!oFootnote)
	{
		return;
	}
	var oRef = oFootnote.Ref;
	if(!oRef)
	{
		return;
	}
	var oRun = oRef.Run;
	if(!oRun)
	{
		return;
	}
	var oRefParagraph = oRun.Paragraph;
	if(!oRefParagraph)
	{
		return;
	}
    if(Asc.c_oAscDocumentRefenceToType.AboveBelow === nType)
    {
        var bIsHdrFtr = (docpostype_HdrFtr === this.GetDocPosType());
        var bIsRefHdrFtr = oRefParagraph.Parent && oRefParagraph.Parent.IsHdrFtr(false);
        if(bIsHdrFtr !== bIsRefHdrFtr)
        {
            return;
        }
    }
	if(false === this.IsSelectionLocked(AscCommon.changestype_Document_Content, {
		Type      : changestype_2_ElementsArray_and_Type,
		Elements  : [oRefParagraph],
		CheckType : changestype_Paragraph_Content
	}))
	{
		this.StartAction(AscDFH.historydescription_Document_AddCrossRef);
		var sBookmarkName;
        if(Asc.c_oAscDocumentRefenceToType.PageNum === nType)
        {
            sBookmarkName = oParagraph.AddBookmarkForRef();
        }
        else
        {
            sBookmarkName = oParagraph.AddBookmarkForNoteRef();
        }
		if(sBookmarkName)
		{
			this.private_AddNoteRefToBookmark(sBookmarkName, nType, bHyperlink, bAboveBelow);
			this.Recalculate();
		}
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
};
CDocument.prototype.private_AddNoteRefToBookmark = function(sBookmarkName, nType, bHyperlink, bAboveBelow)
{
	if(!(typeof sBookmarkName === "string" && sBookmarkName.length > 0))
	{
		return;
	}
	var sInstr = "";
	var sSuffix = "";
	if(bHyperlink)
	{
		sSuffix += " \\h";
	}
	if(bAboveBelow && nType !== Asc.c_oAscDocumentRefenceToType.AboveBelow)
	{
		sSuffix += " \\p";
	}
	switch (nType)
	{
		case Asc.c_oAscDocumentRefenceToType.NoteNumber:
		{
			sInstr = " NOTEREF " + sBookmarkName;
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.PageNum:
		{
			sInstr = " PAGEREF " + sBookmarkName;
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.NoteNumberFormatted:
		{
			sInstr = " NOTEREF " + sBookmarkName + " \\f ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.AboveBelow:
		{
			sInstr = " NOTEREF " + sBookmarkName + " \\p ";
            sInstr += sSuffix;
			break;
		}
	}
	var oComplexField = this.AddFieldWithInstruction(sInstr);
    if(nType !== Asc.c_oAscDocumentRefenceToType.AboveBelow)
    {
        this.FullRecalc.Continue = false;
        this.FullRecalc.UseRecursion = false;
        this.private_Recalculate(undefined, true);
        while (this.FullRecalc.Continue)
        {
            this.FullRecalc.Continue = false;
            this.Recalculate_Page();
        }
        this.FullRecalc.UseRecursion = true;
        oComplexField.Update(false);
    }
    return oComplexField;
};
CDocument.prototype.AddRefToCaption = function(sCaption, oParagraph, nType, bHyperlink, bAboveBelow)
{
	if(nType  === Asc.c_oAscDocumentRefenceToType.PageNum || 
		nType === Asc.c_oAscDocumentRefenceToType.AboveBelow ||
		nType ===  Asc.c_oAscDocumentRefenceToType.Text)
	{
		this.AddRefToParagraph(oParagraph, nType, bHyperlink, bAboveBelow);
		return;
	}
	this.StartAction(AscDFH.historydescription_Document_AddCrossRef);
	var sBookmarkName = null;
	switch(nType)
	{
		case Asc.c_oAscDocumentRefenceToType.OnlyLabelAndNumber:
		{
			sBookmarkName = oParagraph.AddBookmarkForCaption(sCaption, false);
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.OnlyCaptionText:
		{
			sBookmarkName = oParagraph.AddBookmarkForCaption(sCaption, true);
			break;
		}
	}
	if(sBookmarkName)
	{
		this.private_AddRefToBookmark(sBookmarkName, nType, bHyperlink, bAboveBelow, null);
	}
	this.UpdateInterface();
	this.UpdateSelection();
	this.FinalizeAction();
};
CDocument.prototype.private_CreateComplexFieldRun = function(sInstruction, oParagraph)
{
    var oBeginChar    = new ParaFieldChar(fldchartype_Begin, this),
        oSeparateChar = new ParaFieldChar(fldchartype_Separate, this),
        oEndChar      = new ParaFieldChar(fldchartype_End, this);

    var oRun = new ParaRun();
    oRun.AddToContent(-1, oBeginChar);
    oRun.AddInstrText(sInstruction);
    oRun.AddToContent(-1, oSeparateChar);
    oRun.AddToContent(-1, oEndChar);
    oParagraph.Add(oRun);

    oBeginChar.SetRun(oRun);
    oSeparateChar.SetRun(oRun);
    oEndChar.SetRun(oRun);

    var oComplexField = oBeginChar.GetComplexField();
    oComplexField.SetBeginChar(oBeginChar);
    oComplexField.SetInstructionLine(sInstruction);
    oComplexField.SetSeparateChar(oSeparateChar);
    oComplexField.SetEndChar(oEndChar);
    oComplexField.Update(false);
    return oComplexField;
};
CDocument.prototype.UpdateComplexField = function(oField)
{
	if (!oField)
		return;

	oField.Update(true);
};
CDocument.prototype.GetCurrentComplexFields = function()
{
	var oParagraph = this.GetCurrentParagraph();
	if (!oParagraph)
		return [];

	return oParagraph.GetCurrentComplexFields();
};
CDocument.prototype.IsFastCollaborationBeforeViewModeInReview = function()
{
	return this.ViewModeInReview.isFastCollaboration;
};
CDocument.prototype.CheckComplexFieldsInSelection = function()
{
	if (true !== this.Selection.Use || this.Controller !== this.LogicDocumentController)
		return;

	// Смотрим сколько полей открытых в начальной и конечной позициях.
	var oStartPos = this.GetContentPosition(true, true);
	var oEndPos   = this.GetContentPosition(true, false);

	var arrStartComplexFields = this.GetComplexFieldsByContentPos(oStartPos);
	var arrEndComplexFields   = this.GetComplexFieldsByContentPos(oEndPos);

	var arrStartFields = [];
	for (var nIndex = 0, nCount = arrStartComplexFields.length; nIndex < nCount; ++nIndex)
	{
		var bAdd = true;
		for (var nIndex2 = 0, nCount2 = arrEndComplexFields.length; nIndex2 < nCount2; ++nIndex2)
		{
			if (arrStartComplexFields[nIndex] === arrEndComplexFields[nIndex2])
			{
				bAdd = false;
				break;
			}
		}

		if (bAdd)
		{
			arrStartFields.push(arrStartComplexFields[nIndex]);
		}
	}

	var arrEndFields = [];
	for (var nIndex = 0, nCount = arrEndComplexFields.length; nIndex < nCount; ++nIndex)
	{
		var bAdd = true;
		for (var nIndex2 = 0, nCount2 = arrStartComplexFields.length; nIndex2 < nCount2; ++nIndex2)
		{
			if (arrEndComplexFields[nIndex] === arrStartComplexFields[nIndex2])
			{
				bAdd = false;
				break;
			}
		}

		if (bAdd)
		{
			arrEndFields.push(arrEndComplexFields[nIndex]);
		}
	}

	var nDirection = this.GetSelectDirection();
	if (nDirection > 0)
	{
		if (arrStartFields.length > 0 && arrStartFields[0].IsValid())
			oStartPos = arrStartFields[0].GetStartDocumentPosition();

		if (arrEndFields.length > 0 && arrEndFields[0].IsValid())
			oEndPos = arrEndFields[0].GetEndDocumentPosition();
	}
	else
	{
		if (arrStartFields.length > 0 && arrStartFields[0].IsValid())
			oStartPos = arrStartFields[0].GetEndDocumentPosition();

		if (arrEndFields.length > 0 && arrEndFields[0].IsValid())
			oEndPos = arrEndFields[0].GetStartDocumentPosition();
	}

	this.SetContentSelection(oStartPos, oEndPos, 0, 0, 0);
};
/**
 * Получаем текст, который лежит внутри заданного сложного поля. Если внутри текста есть объекты, то возвращается null.
 * @param {CComplexField} oComplexField
 * @returns {?string}
 */
CDocument.prototype.GetComplexFieldTextValue = function(oComplexField)
{
	if (!oComplexField)
		return null;

	var oState = this.SaveDocumentState();
	oComplexField.SelectFieldValue();
	var sResult = this.GetSelectedText();
	this.LoadDocumentState(oState);
	this.Document_UpdateSelectionState();

	return sResult;
};
/**
 * Получаем прямые текстовые настройки заданного сложного поля.
 * @param oComplexField
 * @returns {CTextPr}
 */
CDocument.prototype.GetComplexFieldTextPr = function(oComplexField)
{
	if (!oComplexField)
		return new CTextPr();

	var oState = this.SaveDocumentState();
	oComplexField.SelectFieldValue();
	var oTextPr = this.GetDirectTextPr();
	this.LoadDocumentState(oState);
	this.Document_UpdateSelectionState();

	return oTextPr;
};
CDocument.prototype.GetFieldsManager = function()
{
	return this.FieldsManager;
};
CDocument.prototype.GetComplexFieldsByContentPos = function(oDocPos)
{
	if (!oDocPos)
		return [];

	var oCurrentDocPos = this.GetContentPosition(false);
	this.SetContentPosition(oDocPos, 0, 0);

	var oCurrentParagraph = this.controller_GetCurrentParagraph(true, null);
	if (!oCurrentParagraph)
		return [];

	var arrComplexFields = oCurrentParagraph.GetCurrentComplexFields();
	this.SetContentPosition(oCurrentDocPos, 0, 0);

	return arrComplexFields;
};
/**
 * Получаем массив все полей в документе (простых и сложных)
 * @param {boolean} isUseSelection
 * @returns {Array}
 */
CDocument.prototype.GetAllFields = function(isUseSelection)
{
	var arrFields = [];

	if (isUseSelection)
	{
		this.Controller.GetAllFields(isUseSelection, arrFields);
	}
	else
	{
		// По сноскам и автофигурам мы пробегаемся при поиске в основной части документа

		this.LogicDocumentController.GetAllFields(false, arrFields);

		var arrHdrFtrs = this.SectionsInfo.GetAllHdrFtrs();
		for (var nIndex = 0, nCount = arrHdrFtrs.length; nIndex < nCount; ++nIndex)
		{
			arrHdrFtrs[nIndex].GetContent().GetAllFields(true, arrFields);
		}
	}

	return arrFields;
};
/**
 * Получаем список всех полей заданного аддона
 * @param {boolean} [bySelection=false]
 * @returns {Array.CComplexField}
 */
CDocument.prototype.GetAllAddinFields = function(bySelection)
{
	let allFields   = this.GetAllFields(bySelection);
	let addinFields = [];
	allFields.forEach(function(field)
	{
		if ((field instanceof AscWord.CComplexField) && field.IsAddin())
			addinFields.push(field);
	});
	
	return addinFields;
};
/**
 * Create Addin field by data
 * @param {AscWord.CAddinFieldData} data
 */
CDocument.prototype.AddAddinField = function(data)
{
	if (!data)
		return;
	
	let instruction = data.GetValue();
	if (!instruction)
		return;
	
	let innerText = data.GetContent();
	if (!innerText)
		innerText = "    ";
	
	if (this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
		return;
	
	this.StartAction(AscDFH.historydescription_Document_AddAddinField);
	
	if (this.IsTextSelectionUse())
		this.RemoveBeforePaste();
	else if (this.IsSelectionUse())
		this.RemoveSelection();
	
	let field = this.AddFieldWithInstruction(" ADDIN " + instruction + " ");
	if (field)
	{
		field.SelectFieldValue();
		this.AddText(innerText);
	}
	
	this.Recalculate();
	this.UpdateInterface();
	this.UpdateSelection();
	this.FinalizeAction();
};
/**
 * Update addin fields
 * @param arrData
 */
CDocument.prototype.UpdateAddinFieldsByData = function(arrData)
{
	if (!arrData || !Array.isArray(arrData))
		return;
	
	let allFields   = this.GetAllFields();
	let paragraphs  = [];
	let addinFields = {};
	arrData.forEach(function(data)
	{
		let fieldId = data.GetFieldId();
		for (let index = 0, count = allFields.length; index < count; ++index)
		{
			if (allFields[index] instanceof AscWord.CComplexField && allFields[index].GetFieldId() === fieldId)
			{
				addinFields[fieldId] = allFields[index];
				break;
			}
		}
	});
	
	for (let fieldId in addinFields)
	{
		let field = addinFields[fieldId];
		field.SelectField();
		let selectedParagraphs = this.GetSelectedParagraphs();
		for (let index = 0, count = selectedParagraphs.length; index < count; ++index)
		{
			if (-1 === paragraphs.indexOf(selectedParagraphs[index]))
				paragraphs.push(selectedParagraphs[index]);
		}
	}
	
	if (this.IsSelectionLocked(changestype_None, {
		Type      : changestype_2_ElementsArray_and_Type,
		Elements  : paragraphs,
		CheckType : AscCommon.changestype_Paragraph_Content
	}))
		return;
	
	this.StartAction(AscDFH.historydescription_Document_UpdateAddinFields);
	
	let logicDocument = this;
	arrData.forEach(function(data)
	{
		let field = addinFields[data.GetFieldId()];
		if (!field)
			return;
		
		let value = data.GetValue();
		if (value)
			field.ChangeInstruction(" ADDIN " + value + " ");
		
		let content = data.GetContent();
		if (content)
		{
			field.SelectFieldValue();
			logicDocument.Remove();
			logicDocument.AddText(content);
		}
	});
	
	this.Recalculate();
	this.UpdateInterface();
	this.UpdateSelection();
	this.FinalizeAction();
};
/**
 * Remove field wrapper
 * @param {string} [fieldId=undefined] if not specified then remove wrapper from current field
 */
CDocument.prototype.RemoveComplexFieldWrapper = function(fieldId)
{
	let field;
	if (fieldId)
	{
		let allFields = this.GetAllFields();
		for (let index = 0, count = allFields.length; index < count; ++index)
		{
			if (allFields[index] instanceof AscWord.CComplexField && allFields[index].GetFieldId() === fieldId)
			{
				field = allFields[index];
				break;
			}
		}
	}
	else
	{
		field = this.GetCurrentComplexField();
	}
	
	if (!field || !(field instanceof AscWord.CComplexField))
		return;
	
	field.SelectField();
	let paragraphs = this.GetSelectedParagraphs();
	
	if (this.IsSelectionLocked(changestype_None, {
		Type      : changestype_2_ElementsArray_and_Type,
		Elements  : paragraphs,
		CheckType : AscCommon.changestype_Paragraph_Content
	}))
		return;
	
	this.StartAction(AscDFH.historydescription_Document_RemoveComplexFieldWrapper);
	
	field.RemoveFieldWrap();
	
	this.Recalculate();
	this.UpdateInterface();
	this.UpdateSelection();
	this.FinalizeAction();
};
/**
 * Update fields in the document by selection or all
 * @param isBySelection {boolean}
 */
CDocument.prototype.UpdateFields = function(isBySelection)
{
	var arrFields = this.GetAllFields(isBySelection);

	if (arrFields.length <= 0)
	{
		var oInfo = this.GetSelectedElementsInfo();
		arrFields = oInfo.GetComplexFields();
	}

	var oDocState = this.SaveDocumentState();

	var arrParagraphs = [];
	for (var nIndex = 0, nCount = arrFields.length; nIndex < nCount; ++nIndex)
	{
		var oField = arrFields[nIndex];
		if (oField instanceof CComplexField)
		{
			oField.SelectField();
			arrParagraphs = arrParagraphs.concat(this.GetCurrentParagraph(false, true));
		}
		else if (oField instanceof ParaField)
		{
			if (oField.GetParagraph())
				arrParagraphs.push(oField.GetParagraph());
		}
	}

	if (!this.Document_Is_SelectionLocked(changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : arrParagraphs,
			CheckType : changestype_Paragraph_Content
		}))
	{
		this.StartAction(AscDFH.historydescription_Document_UpdateFields);

		// TODO: Функция работает плохо. Обновляются вообще все поля, даже вложенные в другие
		//       Вложенные сами по себе обновятся при обновлении внешних
		for (var nIndex = 0, nCount = arrFields.length; nIndex < nCount; ++nIndex)
		{
			arrFields[nIndex].Update(false, false);
		}

		this.LoadDocumentState(oDocState);

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}

};
/**
 * Получаем ссылку на класс, управляющий закладками
 * @returns {CBookmarksManager}
 */
CDocument.prototype.GetBookmarksManager = function()
{
	return this.BookmarksManager;
};
CDocument.prototype.UpdateBookmarks = function()
{
	if (!this.BookmarksManager.IsNeedUpdate())
		return;

	this.BookmarksManager.BeginCollectingProcess();

	this.SectionsInfo.UpdateBookmarks(this.BookmarksManager);
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].UpdateBookmarks(this.BookmarksManager);
	}
	this.Footnotes.UpdateBookmarks(this.BookmarksManager);
	this.Endnotes.UpdateBookmarks(this.BookmarksManager);

	this.BookmarksManager.EndCollectingProcess();
};
CDocument.prototype.AddBookmark = function(sName)
{
	var arrBookmarkChars = this.BookmarksManager.GetBookmarkByName(sName);

	var oStartPara = null,
		oEndPara   = null;
	var arrParagraphs = [];
	if (true === this.IsSelectionUse())
	{
		var arrSelectedParagraphs = this.GetCurrentParagraph(false, true);
		if (arrSelectedParagraphs.length > 1)
		{
			arrParagraphs.push(arrSelectedParagraphs[0]);
			arrParagraphs.push(arrSelectedParagraphs[arrSelectedParagraphs.length - 1]);

			oStartPara = arrSelectedParagraphs[0];
			oEndPara   = arrSelectedParagraphs[arrSelectedParagraphs.length - 1];
		}
		else if (arrSelectedParagraphs.length === 1)
		{
			arrParagraphs.push(arrSelectedParagraphs[0]);

			oStartPara = arrSelectedParagraphs[0];
			oEndPara   = arrSelectedParagraphs[0];
		}
	}
	else
	{
		var oCurrentParagraph = this.GetCurrentParagraph(true);
		if (oCurrentParagraph)
		{
			arrParagraphs.push(oCurrentParagraph);
			oStartPara = oCurrentParagraph;
		}
	}

	if (arrParagraphs.length <= 0)
		return;

	if (arrBookmarkChars)
	{
		var oStartPara = arrBookmarkChars[0].GetParagraph();
		var oEndPara   = arrBookmarkChars[1].GetParagraph();

		if (oStartPara)
			arrParagraphs.push(oStartPara);

		if (oEndPara)
			arrParagraphs.push(oEndPara);
	}

	if (false === this.Document_Is_SelectionLocked(changestype_None, {Type : AscCommon.changestype_2_ElementsArray_and_Type, Elements : arrParagraphs, CheckType : changestype_Paragraph_Content}, true))
	{
		this.StartAction(AscDFH.historydescription_Document_AddBookmark);

		if (this.BookmarksManager.HaveBookmark(sName))
			this.private_RemoveBookmark(sName);

		var sBookmarkId = this.BookmarksManager.GetNewBookmarkId();

		if (true === this.IsSelectionUse())
		{
			var nDirection = this.GetSelectDirection();
			if (nDirection > 0)
			{
				oEndPara.AddBookmarkChar(new CParagraphBookmark(false, sBookmarkId, sName), true, false);
				oStartPara.AddBookmarkChar(new CParagraphBookmark(true, sBookmarkId, sName), true, true);
			}
			else
			{
				oEndPara.AddBookmarkChar(new CParagraphBookmark(false, sBookmarkId, sName), true, true);
				oStartPara.AddBookmarkChar(new CParagraphBookmark(true, sBookmarkId, sName), true, false);
			}
		}
		else
		{
			oStartPara.AddBookmarkChar(new CParagraphBookmark(false, sBookmarkId, sName), false);
			oStartPara.AddBookmarkChar(new CParagraphBookmark(true, sBookmarkId, sName), false);
		}

		// TODO: Здесь добавляются просто метки закладок, нужно сделать упрощенный пересчет
		this.Recalculate();
		this.FinalizeAction();
	}

};
CDocument.prototype.RemoveBookmark = function(sName)
{
	var arrBookmarkChars = this.BookmarksManager.GetBookmarkByName(sName);

	var arrParagraphs = [];
	if (arrBookmarkChars)
	{
		var oStartPara = arrBookmarkChars[0].GetParagraph();
		var oEndPara   = arrBookmarkChars[1].GetParagraph();

		if (oStartPara)
			arrParagraphs.push(oStartPara);

		if (oEndPara)
			arrParagraphs.push(oEndPara);
	}

	if (false === this.Document_Is_SelectionLocked(changestype_None, {Type : AscCommon.changestype_2_ElementsArray_and_Type, Elements : arrParagraphs, CheckType : changestype_Paragraph_Content}, true))
	{
		this.StartAction(AscDFH.historydescription_Document_RemoveBookmark);

		this.private_RemoveBookmark(sName);

		// TODO: Здесь добавляются просто метки закладок, нужно сделать упрощенный пересчет
		this.Recalculate();
		this.FinalizeAction();
	}
};
CDocument.prototype.GoToBookmark = function(sName, isSelect)
{
	if (isSelect)
		this.BookmarksManager.SelectBookmark(sName);
	else
		this.BookmarksManager.GoToBookmark(sName);
};
CDocument.prototype.private_RemoveBookmark = function(sName)
{
	var arrBookmarkChars = this.BookmarksManager.GetBookmarkByName(sName);
	if (!arrBookmarkChars)
		return;

	arrBookmarkChars[1].RemoveBookmark();
	arrBookmarkChars[0].RemoveBookmark();
};
CDocument.prototype.AddTableOfContents = function(sHeading, oPr, oSdt)
{
	var oStyles     = this.GetStyles();
	var nStylesType = oPr ? oPr.get_StylesType() : Asc.c_oAscTOCStylesType.Current;

	var isNeedChangeStyles = (Asc.c_oAscTOCStylesType.Current !== nStylesType && nStylesType !== oStyles.GetTOCStylesType());

	var isLocked = true;
	if (isNeedChangeStyles)
	{
		isLocked = this.IsSelectionLocked(AscCommon.changestype_Document_Content, {
			Type : AscCommon.changestype_2_AdditionalTypes, Types : [AscCommon.changestype_Document_Styles]
		});
	}
	else
	{
		isLocked = this.IsSelectionLocked(AscCommon.changestype_Document_Content)
	}

	if (!isLocked)
	{
		this.StartAction(AscDFH.historydescription_Document_AddTableOfContents);

		if (this.DrawingObjects.selectedObjects.length > 0)
		{
			var oContent = this.DrawingObjects.getTargetDocContent();
			if (!oContent || oContent.bPresentation)
			{
				this.DrawingObjects.resetInternalSelection();
			}
		}

		if (oSdt instanceof CBlockLevelSdt)
		{
			oSdt.ClearContentControl();
		}
		else
		{
			this.Remove(1, true, true, true);
			oSdt = this.AddContentControl(c_oAscSdtLevelType.Block);

			if (sHeading)
			{
				var oParagraph = oSdt.GetCurrentParagraph();
				oParagraph.Style_Add(this.Get_Styles().GetDefaultTOCHeading());
				for (var oIterator = sHeading.getUnicodeIterator(); oIterator.check(); oIterator.next())
				{
					var nCharCode = oIterator.value();
					if (AscCommon.IsSpace(nCharCode))
						oParagraph.Add(new AscWord.CRunSpace(nCharCode));
					else
						oParagraph.Add(new AscWord.CRunText(nCharCode));
				}
				oSdt.AddNewParagraph(false, true);
			}
		}

		oSdt.SetThisElementCurrent();

		var sInstruction = "TOC \\o \"1-3\" \\h \\z \\u";
		if (oPr)
		{
			var oInstruction = new CFieldInstructionTOC();
			oInstruction.SetPr(oPr);
			sInstruction = oInstruction.ToString();

			if (isNeedChangeStyles)
				oStyles.SetTOCStylesType(nStylesType);
		}

		this.AddFieldWithInstruction(sInstruction);
		oSdt.SetDocPartObj(undefined, "Table of Contents", true);

		var oNextParagraph = oSdt.GetNextParagraph();
		if (oNextParagraph)
		{
			oNextParagraph.MoveCursorToStartPos(false);
			oNextParagraph.Document_SetThisElementCurrent();
		}
		else
		{
			oSdt.MoveCursorToEndPos(false);
		}

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.UpdateRulers();
		this.FinalizeAction();
	}
};
CDocument.prototype.AddTableOfFigures = function(oPr)
{
    var oStyles     = this.GetStyles();
    var nStylesType = oPr ? oPr.get_StylesType() : Asc.c_oAscTOFStylesType.Current;

    var isNeedChangeStyles = (Asc.c_oAscTOFStylesType.Current !== nStylesType && nStylesType !== oStyles.GetTOFStyleType());

    var isLocked = true;
    if (isNeedChangeStyles)
    {
        isLocked = this.IsSelectionLocked(AscCommon.changestype_Document_Content, {
            Type : AscCommon.changestype_2_AdditionalTypes, Types : [AscCommon.changestype_Document_Styles]
        });
    }
    else
    {
        isLocked = this.IsSelectionLocked(AscCommon.changestype_Document_Content)
    }
    if (!isLocked)
    {
        this.StartAction(AscDFH.historydescription_Document_AddTableOfContents);
        if (this.DrawingObjects.selectedObjects.length > 0)
        {
            var oContent = this.DrawingObjects.getTargetDocContent();
            if (!oContent || oContent.bPresentation)
            {
                this.DrawingObjects.resetInternalSelection();
            }
        }
        var sInstruction = "TOC \\h \\z \\u";
        if (oPr)
        {
            var oInstruction = new CFieldInstructionTOC();
            oInstruction.SetPr(oPr);
            sInstruction = oInstruction.ToString();
        }
        this.Remove(1, true, true, true);
        var oCurParagraph = this.GetCurrentParagraph();
        if(oCurParagraph)
        {
            if(!oCurParagraph.IsEmpty())
            {
                this.AddNewParagraph(false, false);
            }
        }
        else
        {
            this.AddNewParagraph(false, false);
        }
        var oComplexField = this.AddFieldWithInstruction(sInstruction);
        if(oComplexField)
        {
            if (oPr)
            {
                if (isNeedChangeStyles)
                    oStyles.SetTOFStyleType(nStylesType);
                oComplexField.Update();
                oComplexField.MoveCursorOutsideElement(false);
                var oNextParagraph;
                var oParagraph = this.GetCurrentParagraph();
                if(oParagraph)
                {
                    oNextParagraph = oParagraph.GetNextParagraph();
                    if (oNextParagraph)
                    {
                        oNextParagraph.MoveCursorToStartPos(false);
                        oNextParagraph.Document_SetThisElementCurrent();
                    }
                    else
                    {
                        oParagraph.MoveCursorToEndPos(false);
                    }
                }
            }
            this.Recalculate();
            this.UpdateInterface();
            this.UpdateSelection();
            this.UpdateRulers();
            this.FinalizeAction();
        }
        else
        {
            this.FinalizeAction();
            this.Document_Undo();
        }
    }
};
CDocument.prototype.GetPage = function(nPageAbs)
{
	return this.Pages[nPageAbs];
};
CDocument.prototype.GetPagesCount = function()
{
	return this.Pages.length;
};
/**
 * Данная функция получает первую таблицу TOC по схеме Word
 * @param {boolean} [isCurrent=false] Ищем только в текущем месте или с начала документа
 * @returns {CBlockLevelSdt | CComplexField | null}
 */
CDocument.prototype.GetTableOfContents = function(isCurrent)
{
	if (true === isCurrent)
	{
		var oSelectedInfo = this.GetSelectedElementsInfo();
		return oSelectedInfo.GetTableOfContents();
	}
	else
	{
		// 1. Ищем среди CBlockLevelSdt с параметром Unique = true
		// 2. Ищем среди CBlockLevelSdt
		// 3. Ищем потом просто в сложных полях

		for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			var oResult = this.Content[nIndex].GetTableOfContents(true, false);
			if (oResult)
				return oResult;
		}

		for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			var oResult = this.Content[nIndex].GetTableOfContents(false, true);
			if (oResult)
				return oResult;
		}
	}

	return null;
};
/**
 * Данная функция получает все таблицы TOC в документе по схеме Word
 * @returns {Array}
 */
CDocument.prototype.GetAllTablesOfContentsInDoc = function()
{
	// 1. Ищем среди CBlockLevelSdt с параметром Unique = true
	// 2. Ищем среди CBlockLevelSdt
	// 3. Ищем потом просто в сложных полях

	var aTOC = [];
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		var oResult = this.Content[nIndex].GetTableOfContents(true, false);
		if (oResult)
			aTOC.indexOf(oResult) === -1 && aTOC.push(oResult);
	}

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		var oResult = this.Content[nIndex].GetTableOfContents(false, true);
		if (oResult)
			aTOC.indexOf(oResult) === -1 && aTOC.push(oResult);
	}

	return aTOC;
};
/**
 * Данная функция получает все таблицы TOF в документе по схеме Word
 * @returns {Array}
 */
CDocument.prototype.GetAllTablesOfFiguresInDoc = function()
{
	var aTOF = [];
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		this.Content[nIndex].GetTablesOfFigures(aTOF);

	return aTOF;
};
CDocument.prototype.GetAllTablesOfFigures = function(isCurrent)
{
    if (true === isCurrent)
    {
        var oSelectedInfo = this.GetSelectedElementsInfo({CheckAllSelection: true, AddAllComplexFields: true});
        return oSelectedInfo.GetAllTablesOfFigures();
    }
    else
    {
        var aResult = [];
        for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
        {
            this.Content[nIndex].GetTablesOfFigures(aResult);
            if (aResult.length > 0)
                return aResult;
        }
        return aResult;
    }
}
/**
 * Получаем текущее сложное поле
 * @returns {CComplexField | AscWord.CRunPageNum | AscWord.CRunPagesCount | null}
 */
CDocument.prototype.GetCurrentComplexField = function()
{
	if (this.IsSelectionUse())
	{
		let arrFields = this.GetAllFields(true);
		if (arrFields.length > 0)
			return arrFields[arrFields.length - 1];
	}
	else
	{
		var oSelectedInfo    = this.GetSelectedElementsInfo();
		var arrComplexFields = oSelectedInfo.GetComplexFields();

		if (arrComplexFields.length > 0)
			return arrComplexFields[arrComplexFields.length - 1];

		var oPageNum = oSelectedInfo.GetPageNum();
		if (oPageNum)
			return oPageNum;

		var oPagesCount = oSelectedInfo.GetPagesCount();
		if (oPagesCount)
			return oPagesCount;
	}

	return null;
};
/**
 * Очищаем закэшированные списки, которые чаще всего используются
 */
CDocument.prototype.ClearListsCache = function()
{
	this.AllParagraphsList = null;
	this.AllFootnotesList  = null;
	this.AllEndnotesList   = null;
};
/**
 * Получаем массив положений, к которым можно привязать гиперссылку
 * @returns {CHyperlinkAnchor[]}
 */
CDocument.prototype.GetHyperlinkAnchors = function()
{
	var arrAnchors = [];

	var arrOutline = [];
	this.GetOutlineParagraphs(arrOutline, {SkipEmptyParagraphs : true, OutlineStart : 1, OutlineEnd : 9});
	var nIndex = 0, nCount = arrOutline.length;
	for (nIndex = 0; nIndex < nCount; ++nIndex)
	{
		arrAnchors.push(new CHyperlinkAnchor(c_oAscHyperlinkAnchor.Heading, arrOutline[nIndex]));
	}

	this.BookmarksManager.Update();
	nCount = this.BookmarksManager.GetCount();
	for (nIndex = 0; nIndex < nCount; ++nIndex)
	{
		var sName =  this.BookmarksManager.GetName(nIndex);
		if (!this.BookmarksManager.IsHiddenBookmark(sName) && !this.BookmarksManager.IsInternalUseBookmark(sName))
			arrAnchors.push(new CHyperlinkAnchor(c_oAscHyperlinkAnchor.Bookmark, sName));
	}

	return arrAnchors;
};
/**
 * Получаем текущую выделенную нумерацию
 * @param [isCheckSelection=false] {boolean} - проверять ли по селекту
 * @returns {?CNumPr} Если выделено несколько параграфов с одним NumId, но с разными уровнями, то Lvl = null
 */
CDocument.prototype.GetSelectedNum = function(isCheckSelection)
{
	if (true === isCheckSelection)
	{
		var arrParagraphs = this.GetCurrentParagraph(false, true);
		if (arrParagraphs.length <= 0)
			return null;

		var oStartNumPr = arrParagraphs[0].GetNumPr();
		if (!oStartNumPr)
			return null;

		var oNumPr = oStartNumPr.Copy();
		for (var nIndex = 1, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
		{
			var oCurNumPr = arrParagraphs[nIndex].GetNumPr();
			if (!oCurNumPr || oCurNumPr.NumId !== oStartNumPr.NumId)
				return null;

			if (oCurNumPr.Lvl !== oStartNumPr.Lvl)
				oNumPr.Lvl = null;
		}

		return oNumPr;
	}
	else
	{
		var oCurrentPara = this.GetCurrentParagraph(true);
		if (!oCurrentPara)
			return null;

		var oDocContent = oCurrentPara.GetParent();
		if (!oDocContent)
			return null;

		var oTopDocContent = oDocContent.GetTopDocumentContent();
		if (!oTopDocContent || !oTopDocContent.IsNumberingSelection())
			return null;

		oCurrentPara = oTopDocContent.Selection.Data.CurPara;
		if (!oCurrentPara)
			return null;

		return oCurrentPara.GetNumPr();
	}
};
/**
 * Продолжаем нумерацию в текущем параграфе
 * @returns {boolean}
 */
CDocument.prototype.ContinueNumbering = function()
{
	var isNumberingSelection = this.IsNumberingSelection();

	var oParagraph = this.GetCurrentParagraph(true);
	if (!oParagraph || !oParagraph.GetNumPr() || (this.IsSelectionUse() && !isNumberingSelection))
		return false;

	var oNumPr = oParagraph.GetNumPr();

	var oEngine = new CDocumentNumberingContinueEngine(oParagraph, oNumPr, this.GetNumbering());
	this.Controller.GetSimilarNumbering(oEngine);
	var oSimilarNumPr = oEngine.GetNumPr();
	if (!oSimilarNumPr || oSimilarNumPr.NumId === oNumPr.NumId)
		return false;

	var bFind                 = false;
	var arrParagraphs         = this.GetAllParagraphsByNumbering(new CNumPr(oNumPr.NumId, null));
	var arrParagraphsToChange = [];

	var isRelated = this.GetNumbering().GetNum(oNumPr.NumId).IsHaveRelatedLvlText();

	for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
	{
		var oPara = arrParagraphs[nIndex];

		if (!bFind && oPara === oParagraph)
			bFind = true;

		if (bFind)
		{
			var oCurNumPr = oPara.GetNumPr();
			if (!oCurNumPr)
				continue;

			if (!isRelated)
			{
				if (oCurNumPr.Lvl < oNumPr.Lvl)
					break;
				else if (oCurNumPr.Lvl > oNumPr.Lvl)
					continue;
			}

			arrParagraphsToChange.push(oPara);
		}
	}

	if (!this.Document_Is_SelectionLocked(changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : arrParagraphsToChange,
			CheckType : AscCommon.changestype_Paragraph_Properties
		}))
	{
		this.StartAction(AscDFH.historydescription_Document_ContinueNumbering);

		for (var nIndex = 0, nCount = arrParagraphsToChange.length; nIndex < nCount; ++nIndex)
		{
			var oPara     = arrParagraphsToChange[nIndex];
			var oCurNumPr = oPara.GetNumPr();

			oPara.SetNumPr(oSimilarNumPr.NumId, oCurNumPr.Lvl);
		}

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}

	return true;
};
/**
 * Начинаем нумерацию занового с текущего параграфа с заданного значения
 * @param [nRestartValue=1] {number}
 * @returns {boolean}
 */
CDocument.prototype.RestartNumbering = function(nRestartValue)
{
	if (undefined === nRestartValue || null === nRestartValue)
		nRestartValue = 1;

	var isNumberingSelection = this.IsNumberingSelection();

	var oParagraph = this.GetCurrentParagraph(true);
	if (!oParagraph || !oParagraph.GetNumPr() || (this.IsSelectionUse() && !isNumberingSelection))
		return false;

	var oNumPr = oParagraph.GetNumPr();

	var bFind                 = false;
	var nPrevLvl              = null;
	var isFirstParaOnLvl      = false;
	var arrParagraphs         = this.GetAllParagraphsByNumbering(new CNumPr(oNumPr.NumId, null));
	var arrParagraphsToChange = [];
	var isRelated             = this.Numbering.GetNum(oNumPr.NumId).IsHaveRelatedLvlText();

	for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
	{
		var oPara = arrParagraphs[nIndex];

		if (!bFind && oPara === oParagraph)
		{
			bFind = true;

			if (null === nPrevLvl || nPrevLvl < oNumPr.Lvl)
				isFirstParaOnLvl = true;
		}

		var oCurNumPr = oPara.GetNumPr();

		if (bFind)
		{
			if (!oCurNumPr)
				continue;

			if (0 !== oNumPr.Lvl && !isRelated)
			{
				if (oCurNumPr.Lvl < oNumPr.Lvl)
					break;
				else if (oCurNumPr.Lvl > oNumPr.Lvl)
					continue;
			}

			arrParagraphsToChange.push(oPara);
		}

		if (oCurNumPr)
			nPrevLvl = oCurNumPr.Lvl;
	}

	var oNum = this.Numbering.GetNum(oNumPr.NumId);
	if (isFirstParaOnLvl)
	{
		if (Asc.c_oAscNumberingFormat.Bullet === oNum.GetLvl(oNumPr.Lvl).GetFormat())
			return;

		if (!this.Document_Is_SelectionLocked(changestype_None, {
				Type      : changestype_2_ElementsArray_and_Type,
				Elements  : [oNum],
				CheckType : AscCommon.changestype_Paragraph_Properties
			}))
		{
			this.StartAction(AscDFH.historydescription_Document_RestartNumbering);

			var oLvl   = oNum.GetLvl(oNumPr.Lvl).Copy();
			oLvl.Start = nRestartValue;
			oNum.SetLvl(oLvl, oNumPr.Lvl);

			this.Recalculate();
			this.UpdateInterface();
			this.UpdateSelection();
			this.FinalizeAction();
		}
	}
	else
	{
		if (!this.Document_Is_SelectionLocked(changestype_None, {
				Type      : changestype_2_ElementsArray_and_Type,
				Elements  : arrParagraphsToChange,
				CheckType : AscCommon.changestype_Paragraph_Properties
			}))
		{
			this.StartAction(AscDFH.historydescription_Document_RestartNumbering);

			var oNewNum = oNum.Copy();
			if (Asc.c_oAscNumberingFormat.Bullet !== oNum.GetLvl(oNumPr.Lvl).GetFormat())
			{
				var oLvl   = oNewNum.GetLvl(oNumPr.Lvl).Copy();
				oLvl.Start = nRestartValue;
				oNewNum.SetLvl(oLvl, oNumPr.Lvl);
			}

			var sNewId = oNewNum.GetId();

			for (var nIndex = 0, nCount = arrParagraphsToChange.length; nIndex < nCount; ++nIndex)
			{
				var oPara     = arrParagraphsToChange[nIndex];
				var oCurNumPr = oPara.GetNumPr();

				oPara.SetNumPr(sNewId, oCurNumPr.Lvl);
			}

			if (isNumberingSelection)
				this.SelectNumbering(oParagraph.GetNumPr(), oParagraph);

			this.Recalculate();
			this.UpdateInterface();
			this.UpdateSelection();
			this.FinalizeAction();
		}
	}

	return true;
};
CDocument.prototype.GetAutoCorrectSettings = function()
{
	return this.AutoCorrectSettings;
};
/**
 * Устанавливаем настройку автосоздания маркированных списков
 * @param isAuto {boolean}
 */
CDocument.prototype.SetAutomaticBulletedLists = function(isAuto)
{
	this.AutoCorrectSettings.SetAutomaticBulletedLists(isAuto);
};
/**
 * Запрашиваем настройку автосоздания маркированных списков
 * @returns {boolean}
 */
CDocument.prototype.IsAutomaticBulletedLists = function()
{
	return this.AutoCorrectSettings.IsAutomaticBulletedLists();
};
/**
 * Устанавливаем настройку автосоздания нумерованных списков
 * @param isAuto {boolean}
 */
CDocument.prototype.SetAutomaticNumberedLists = function(isAuto)
{
	this.AutoCorrectSettings.SetAutomaticNumberedLists(isAuto);
};
/**
 * Запрашиваем настройку автосоздания нумерованных списков
 * @returns {boolean}
 */
CDocument.prototype.IsAutomaticNumberedLists = function()
{
	return this.AutoCorrectSettings.IsAutomaticNumberedLists();
};
/**
 * Устанавливаем параметр автозамены: заменять ли прямые кавычки "умными"
 * @param isSmartQuotes {boolean}
 */
CDocument.prototype.SetAutoCorrectSmartQuotes = function(isSmartQuotes)
{
	this.AutoCorrectSettings.SetSmartQuotes(isSmartQuotes);
};
/**
 * Запрашиваем настройку автозамены: заменять ли прямые кавычки "умными"
 * @returns {boolean}
 */
CDocument.prototype.IsAutoCorrectSmartQuotes = function()
{
	return this.AutoCorrectSettings.IsSmartQuotes();
};
/**
 * Устанавливаем параметр автозамены двух дефисов на тире
 * @param isReplace {boolean}
 */
CDocument.prototype.SetAutoCorrectHyphensWithDash = function(isReplace)
{
	this.AutoCorrectSettings.SetHyphensWithDash(isReplace);
};
/**
 * Запрашиваем настройку автозамены двух дефисов на тире
 * @returns {boolean}
 */
CDocument.prototype.IsAutoCorrectHyphensWithDash = function()
{
	return this.AutoCorrectSettings.IsHyphensWithDash();
};
/**
 * Запрашиваем настройку автозамены для французской пунктуации
 * @returns {boolean}
 */
CDocument.prototype.IsAutoCorrectFrenchPunctuation = function()
{
	return this.AutoCorrectSettings.IsFrenchPunctuation();
};
/**
 * Запрашиваем настройку автозамены двойного пробела на точку
 * @returns {boolean}
 */
CDocument.prototype.IsAutoCorrectDoubleSpaceWithPeriod = function()
{
	return this.AutoCorrectSettings.IsDoubleSpaceWithPeriod();
};
/**
 * Выставляем настройку атозамены двойного пробела на точку
 * @param {boolean} isCorrect
 */
CDocument.prototype.SetAutoCorrectDoubleSpaceWithPeriod = function(isCorrect)
{
	this.AutoCorrectSettings.SetDoubleSpaceWithPeriod(isCorrect);
};
/**
 * Выставляем настройку атозамены для первого символа предложения
 * @param {boolean} isCorrect
 */
CDocument.prototype.SetAutoCorrectFirstLetterOfSentences = function(isCorrect)
{
	this.AutoCorrectSettings.SetFirstLetterOfSentences(isCorrect);
};
/**
 * Запрашиваем настройку атозамены для первого символа предложения
 * @return {boolean}
 */
CDocument.prototype.IsAutoCorrectFirstLetterOfSentences = function()
{
	return this.AutoCorrectSettings.IsFirstLetterOfSentences();
};
/**
 * Выставляем настройку атозамены для первого символа в ячейке таблицы
 * @param {boolean} isCorrect
 */
CDocument.prototype.SetAutoCorrectFirstLetterOfCells = function(isCorrect)
{
	this.AutoCorrectSettings.SetFirstLetterOfCells(isCorrect);
}
/**
 * Запрашиваем настройку атозамены для первого символа ячейки таблицы
 * @return {boolean}
 */
CDocument.prototype.IsAutoCorrectFirstLetterOfCells = function()
{
	return this.AutoCorrectSettings.IsFirstLetterOfCells();
};
/**
 * Выставляем настройку атозамены для гиперссылок
 * @param {boolean} isCorrect
 */
CDocument.prototype.SetAutoCorrectHyperlinks = function(isCorrect)
{
	this.AutoCorrectSettings.SetHyperlinks(isCorrect);
};
/**
 * Запрашиваем настройку атозамены для гиперссылок
 * @return {boolean}
 */
CDocument.prototype.IsAutoCorrectHyperlinks = function()
{
	return this.AutoCorrectSettings.IsHyperlinks();
};
/**
 * Получаем идентификатор текущего пользователя
 * @param [isConnectionId=false] {boolean} true - Id соединения пользователя или false - Id пользователя
 * @returns {string}
 */
CDocument.prototype.GetUserId = function(isConnectionId)
{
	var oApi = this.GetApi();

	if (isConnectionId)
		return oApi.CoAuthoringApi ? oApi.CoAuthoringApi.getUserConnectionId() : "";

	return oApi.DocInfo ? oApi.DocInfo.get_UserId() : "";
};
/**
 * Получаем метки селекта, в зависимости от типа селекта
 * @param [isSaveDirection=true] {boolean} Сохраняем направление селекта, либо разворачиваем от меньшего к большему
 * @returns {{Start: number, End: number}}
 */
CDocument.prototype.private_GetSelectionPos = function(isSaveDirection)
{
	var nStartPos = 0;
	var nEndPos   = 0;

	if (!this.controller_IsSelectionUse())
	{
		nStartPos = this.CurPos.ContentPos;
		nEndPos   = this.CurPos.ContentPos;
	}
	else if (this.controller_IsMovingTableBorder())
	{
		nStartPos = this.Selection.Data.Pos;
		nEndPos   = this.Selection.Data.Pos;
	}
	else if (this.controller_IsNumberingSelection())
	{
		this.UpdateContentIndexing();

		var oTopElement = this.Selection.Data.CurPara.GetTopElement();
		if (oTopElement)
		{
			nStartPos = oTopElement.GetIndex();
			nEndPos   = nStartPos;
		}
	}
	else
	{
		nStartPos = this.Selection.StartPos;
		nEndPos   = this.Selection.EndPos;
	}

	if (false === isSaveDirection && nStartPos > nEndPos)
	{
		var nTemp = nStartPos;
		nStartPos = nEndPos;
		nEndPos   = nTemp;
	}

	return {Start : nStartPos, End : nEndPos};
};
/**
 * Получаем объект, внутри которого в данный момент используется PlaceHolder (т.е. мы там стоим курсором или он выделен целиком)
 * @returns {?object}
 */
CDocument.prototype.GetPlaceHolderObject = function()
{
	return this.Controller.GetPlaceHolderObject();
};
/**
 * Добавляем пустую страницу в текущее положение курсор
 */
CDocument.prototype.AddBlankPage = function()
{
	if (this.LogicDocumentController === this.Controller)
	{
		if (!this.Document_Is_SelectionLocked(AscCommon.changestype_Document_Content))
		{
			this.StartAction(AscDFH.historydescription_Document_AddBlankPage);

			if (this.IsSelectionUse())
			{
				this.MoveCursorLeft(false, false);
				this.RemoveSelection();
			}

			var oElement = this.Content[this.CurPos.ContentPos];
			if (oElement.IsParagraph())
			{
				if (this.CurPos.ContentPos === this.Content.length - 1 && oElement.IsCursorAtEnd())
				{
					var oBreakParagraph = oElement.Split();
					var oEmptyParagraph = oElement.Split();

					oBreakParagraph.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
					this.AddToContent(this.CurPos.ContentPos + 1, oBreakParagraph);
					this.AddToContent(this.CurPos.ContentPos + 2, oEmptyParagraph);

					this.CurPos.ContentPos = this.CurPos.ContentPos + 2;
				}
				else
				{
					var oNext   = oElement.Split();
					var oBreak1 = oElement.Split();
					var oBreak2 = oElement.Split();
					var oEmpty  = oElement.Split();

					oBreak1.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
					oBreak2.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));

					this.AddToContent(this.CurPos.ContentPos + 1, oNext);
					this.AddToContent(this.CurPos.ContentPos + 1, oBreak2);
					this.AddToContent(this.CurPos.ContentPos + 1, oEmpty);
					this.AddToContent(this.CurPos.ContentPos + 1, oBreak1);

					this.CurPos.ContentPos = this.CurPos.ContentPos + 2;
				}

				this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
			}
			else if (oElement.IsTable())
			{
				var oNewTable = oElement.Split();
				var oBreak1   = new Paragraph(this.DrawingDocument, this);
				var oEmpty    = new Paragraph(this.DrawingDocument, this);
				var oBreak2   = new Paragraph(this.DrawingDocument, this);

				oBreak1.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));
				oBreak2.AddToParagraph(new AscWord.CRunBreak(AscWord.break_Page));

				if (!oNewTable)
				{
					this.AddToContent(this.CurPos.ContentPos, oBreak2);
					this.AddToContent(this.CurPos.ContentPos, oEmpty);
					this.AddToContent(this.CurPos.ContentPos, oBreak1);
					this.CurPos.ContentPos = this.CurPos.ContentPos + 1;
				}
				else
				{
					this.AddToContent(this.CurPos.ContentPos + 1, oNewTable);
					this.AddToContent(this.CurPos.ContentPos + 1, oBreak2);
					this.AddToContent(this.CurPos.ContentPos + 1, oEmpty);
					this.AddToContent(this.CurPos.ContentPos + 1, oBreak1);
					this.CurPos.ContentPos = this.CurPos.ContentPos + 2;
				}

				this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
			}


			this.Recalculate();
			this.UpdateInterface();
			this.UpdateSelection();
			this.FinalizeAction();
		}
	}
};
/**
 * Получаем формулу в текущей ячейке таблицы
 * {boolean} [isReturnField=false] - возвращаем само поле
 * @returns {string | CComplexField}
 */
CDocument.prototype.GetTableCellFormula = function(isReturnField)
{
	var sDefault = isReturnField ? null : "=";

	var oParagraph = this.GetCurrentParagraph();
	if (!oParagraph)
		return sDefault;

	var oParaParent = oParagraph.GetParent();
	if (!oParaParent)
		return sDefault;

	var oCell = oParaParent.IsTableCellContent(true);
	if (!oCell)
		return sDefault;

	var oCellContent = oCell.GetContent();
	var arrAllFields = oCellContent.GetAllFields(false);

	for (var nIndex = 0, nCount = arrAllFields.length; nIndex < nCount; ++nIndex)
	{
		var oField = arrAllFields[nIndex];
		if (oField instanceof CComplexField && oField.GetInstruction() && fieldtype_FORMULA === oField.GetInstruction().Type)
		{
			if (isReturnField)
				return oField;

			return oField.InstructionLine;
		}
	}

	if (isReturnField)
		return null;

	var oRow   = oCell.GetRow();
	var oTable = oCell.GetTable();

	if (!oRow || !oTable)
		return sDefault;

	var nCurCell = oCell.GetIndex();
	var nCurRow  = oRow.GetIndex();

	var isLeft = false;
	for (var nCellIndex = 0; nCellIndex < nCurCell; ++nCellIndex)
	{
		var oTempCell        = oRow.GetCell(nCellIndex);
		var oTempCellContent = oTempCell.GetContent();

		oTempCellContent.SetApplyToAll(true);
		var sCellText = oTempCellContent.GetSelectedText();
		oTempCellContent.SetApplyToAll(false);

		if (!isNaN(parseInt(sCellText)))
		{
			isLeft = true;
			break;
		}
	}

	var arrColumnCells = oCell.GetColumn();
	var isAbove = false;
	for (var nCellIndex = 0, nCellsCount = arrColumnCells.length; nCellIndex < nCellsCount; ++nCellIndex)
	{
		var oTempCell = arrColumnCells[nCellIndex];
		if (oTempCell === oCell)
			break;

		var oTempCellContent = oTempCell.GetContent();

		oTempCellContent.SetApplyToAll(true);
		var sCellText = oTempCellContent.GetSelectedText();
		oTempCellContent.SetApplyToAll(false);

		if (!isNaN(parseInt(sCellText)))
		{
			isAbove = true;
			break;
		}
	}

	if (isAbove)
		return "=SUM(ABOVE)";

	if (isLeft)
		return "=SUM(LEFT)";

	return sDefault;
};
/**
 * Добавляем формулу к текущей ячейке таблицы
 * @param {string} sFormula
 */
CDocument.prototype.AddTableCellFormula = function(sFormula)
{
	if (!sFormula || "=" !== sFormula.charAt(0))
		return;

	var oField = this.GetTableCellFormula(true);
	if (!oField)
	{
		if (!this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			this.StartAction(AscDFH.historydescription_Document_AddTableFormula);

			if (this.IsSelectionUse())
			{
				if (!this.IsTableCellSelection())
					this.Remove(1, false, false, true);

				this.RemoveSelection();
			}

			this.AddFieldWithInstruction(sFormula);
			this.Recalculate();
			this.UpdateInterface();
			this.UpdateSelection();
			this.FinalizeAction();
		}
	}
	else
	{
		if (!this.Document_Is_SelectionLocked(AscCommon.changestype_Paragraph_Content))
		{
			this.StartAction(AscDFH.historydescription_Document_ChangeTableFormula);
			oField.ChangeInstruction(sFormula);
			oField.Update();
			oField.MoveCursorOutsideElement(false);
			this.Recalculate();
			this.UpdateInterface();
			this.UpdateSelection();
			this.FinalizeAction();
		}
	}

};
/**
 * Разбираем sInstrLine на формулу и формат числа
 * @param {string} sInstrLine
 * @returns {Array}
 */
CDocument.prototype.ParseTableFormulaInstrLine = function(sInstrLine)
{
    var oParser = new CFieldInstructionParser();
    var oResult = oParser.GetInstructionClass(sInstrLine);
    if(oResult && oResult instanceof CFieldInstructionFORMULA)
    {
        return [oResult.Formula ? "=" + oResult.Formula : "=", oResult.Format && oResult.Format.sFormat ? oResult.Format.sFormat : ""];
    }
    return ["", ""];
};


CDocument.prototype.AddCaption = function(oPr)
{
	if (!this.CanPerformAction())
		return;

    this.StartAction(AscDFH.historydescription_Document_AddCaption);
    var NewParagraph;
    if(this.CurPos.Type === docpostype_DrawingObjects)
    {
        if(this.DrawingObjects.selectedObjects.length === 1)
        {
            var oDrawing = this.DrawingObjects.selectedObjects[0].parent;
            if(oDrawing.Is_Inline())
            {
                NewParagraph = new Paragraph(this.DrawingDocument, oDrawing.DocumentContent);
                NewParagraph.SetParagraphStyle("Caption");
                oDrawing.DocumentContent.Internal_Content_Add(oPr.get_Before() ? oDrawing.Get_ParentParagraph().Index : (oDrawing.Get_ParentParagraph().Index + 1), NewParagraph, true);
            }
            else
            {

				var oNearestPos = this.Get_NearestPos(oDrawing.PageNum, oDrawing.X, oDrawing.Y, true, null);
				if(oNearestPos)
				{
                    var oNewDrawing = new ParaDrawing(1828800 / 36000, 1828800 / 36000, null, this.DrawingDocument, this, null);
                    var oShape = this.DrawingObjects.createTextArt(0, true, null, "");
                    oShape.setParent(oNewDrawing);
                    oNewDrawing.Set_GraphicObject(oShape);
                    oNewDrawing.Set_DrawingType(drawing_Anchor);
                    oNewDrawing.Set_WrappingType(oDrawing.wrappingType);
                    oNewDrawing.Set_BehindDoc(oDrawing.behindDoc);
                    oNewDrawing.Set_Distance(3.2, 0, 3.2, 0);
					oNearestPos.Paragraph.Check_NearestPos(oNearestPos);
					oNearestPos.Page = oDrawing.PageNum;
                    oShape.spPr.xfrm.setOffX(0.0);
                    oShape.spPr.xfrm.setOffY(0.0);
                    oShape.spPr.xfrm.setExtX(Math.max(25.4, oDrawing.Extent.W));
                    oShape.spPr.xfrm.setExtY(12.7);
                    var oBodyPr;
                    oBodyPr = oShape.getBodyPr().createDuplicate();
                    oBodyPr.wrap = AscFormat.nTWTSquare;
                    oBodyPr.lIns = 0.0;
                    oBodyPr.tIns = 0.0;
                    oBodyPr.rIns = 0.0;
                    oBodyPr.bIns = 0.0;
                    var dInset = 1.6;
                    var Y = oDrawing.Y + oDrawing.Extent.H + dInset;
                    if(oPr.get_Before())
					{
                        oBodyPr.textFit = new AscFormat.CTextFit();
                        oBodyPr.textFit.type = AscFormat.text_fit_No;
                        Y = oDrawing.Y - oShape.spPr.xfrm.extY - dInset;
					}
                    oNewDrawing.Set_XYForAdd(oDrawing.X, Y, oNearestPos, oDrawing.PageNum);
                    oShape.setBodyPr(oBodyPr);
                    oNewDrawing.Set_Parent(oNearestPos.Paragraph);
					oNewDrawing.AddToDocument(oNearestPos);
					oNewDrawing.CheckWH();
					var oContent = oShape.getDocContent();
					NewParagraph = oContent.Content[0];
                    NewParagraph.RemoveFromContent(0, NewParagraph.Content.length - 1);
                    NewParagraph.Clear_Formatting();
                    NewParagraph.SetParagraphStyle("Caption");
				}
            }
        }
    }
    else
    {
        if(this.Selection.Use)
        {
            var oTable = this.Content[this.Selection.StartPos];
            NewParagraph = new Paragraph(this.DrawingDocument, this);
            NewParagraph.SetParagraphStyle("Caption");
            this.Internal_Content_Add(oPr.get_Before() ? oTable.Index : (oTable.Index + 1), NewParagraph, true);
        }
        else
        {
            NewParagraph = new Paragraph(this.DrawingDocument, this);
            NewParagraph.SetParagraphStyle("Caption");
            this.Internal_Content_Add(oPr.get_Before() ? this.CurPos.ContentPos : (this.CurPos.ContentPos + 1), NewParagraph, true);
        }
    }
    if(NewParagraph)
    {
        var NewRun;
        var nCurPos = 0;
        var oComplexField;
        var oBeginChar, oSeparateChar, oEndChar;
        var aFieldsToUpdate = [];
        if(!oPr.get_ExcludeLabel())
        {
            var sLabel = oPr.get_Label();
            if(typeof sLabel === "string" && sLabel.length > 0)
            {
                NewRun = new ParaRun(NewParagraph, false);
                NewRun.AddText(sLabel + " ");
                NewParagraph.Internal_Content_Add(nCurPos++, NewRun, false);
            }
        }
        if(oPr.get_IncludeChapterNumber())
        {
            var nHeadingLvl = oPr.get_HeadingLvl();
            if(AscFormat.isRealNumber(nHeadingLvl))
            {
                oBeginChar    = new ParaFieldChar(fldchartype_Begin, this);
                oSeparateChar = new ParaFieldChar(fldchartype_Separate, this);
                oEndChar      = new ParaFieldChar(fldchartype_End, this);
                NewRun = new ParaRun();
                NewRun.AddToContent(-1, oBeginChar);

                var sStyleId = this.Styles.GetDefaultHeading(nHeadingLvl - 1);
                var oStyle = this.Styles.Get(sStyleId);
                NewRun.AddInstrText(" STYLEREF \"" + oStyle.GetName() + "\" \\s");
                NewRun.AddToContent(-1, oSeparateChar);
                NewRun.AddToContent(-1, oEndChar);
                oBeginChar.SetRun(NewRun);
                oSeparateChar.SetRun(NewRun);
                oEndChar.SetRun(NewRun);
                NewParagraph.Internal_Content_Add(nCurPos++, NewRun, false);
                oComplexField = oBeginChar.GetComplexField();
                oComplexField.SetBeginChar(oBeginChar);
                oComplexField.SetInstructionLine(" STYLEREF \"" + oStyle.GetName() + "\" \\s");
                oComplexField.SetSeparateChar(oSeparateChar);
                oComplexField.SetEndChar(oEndChar);
                aFieldsToUpdate.push(oComplexField);
            }
            var sSeparator = oPr.get_Separator();
            if(!sSeparator || sSeparator.length === 0)
            {
                sSeparator = " ";
            }
            NewRun = new ParaRun(NewParagraph, false);
            NewRun.AddText(sSeparator);
            NewParagraph.Internal_Content_Add(nCurPos++, NewRun, false);
        }

        var sInstruction = " SEQ " + oPr.getLabelForInstruction() + " \\* " + oPr.get_FormatGeneral() + " ";
        if(AscFormat.isRealNumber(nHeadingLvl))
        {
            sInstruction += ("\\s " + nHeadingLvl);
        }
        oBeginChar    = new ParaFieldChar(fldchartype_Begin, this);
        oSeparateChar = new ParaFieldChar(fldchartype_Separate, this);
        oEndChar      = new ParaFieldChar(fldchartype_End, this);
        NewRun = new ParaRun();
        NewRun.AddToContent(-1, oBeginChar);
        NewRun.AddInstrText(sInstruction);
        NewRun.AddToContent(-1, oSeparateChar);
        NewRun.AddToContent(-1, oEndChar);
        oBeginChar.SetRun(NewRun);
        oSeparateChar.SetRun(NewRun);
        oEndChar.SetRun(NewRun);
        NewParagraph.Internal_Content_Add(nCurPos++, NewRun, false);
        oComplexField = oBeginChar.GetComplexField();
        oComplexField.SetBeginChar(oBeginChar);
        oComplexField.SetInstructionLine(sInstruction);
        oComplexField.SetSeparateChar(oSeparateChar);
        oComplexField.SetEndChar(oEndChar);
        var sAdditional = oPr.get_Additional();
        if(typeof sAdditional === "string" && sAdditional.length > 0)
        {
            NewRun = new ParaRun(NewParagraph, false);
            NewRun.AddText(sAdditional + " ");
            NewParagraph.Internal_Content_Add(nCurPos++, NewRun, false);
        }


        //Update added field
        aFieldsToUpdate.push(oComplexField);
        for(nField = aFieldsToUpdate.length - 1; nField > -1; --nField)
        {
            aFieldsToUpdate[nField].Update(false, false);
        }

        //Update all fields from current sequence
        var aFields = [];
        this.GetAllSeqFieldsByType(oPr.get_Label(), aFields);
        var arrParagraphs = [];
        for (var nIndex = 0, nCount = aFields.length; nIndex < nCount; ++nIndex)
        {
            var oField = aFields[nIndex];
            if (oField instanceof CComplexField)
            {
                oField.SelectField();
                arrParagraphs = arrParagraphs.concat(this.GetCurrentParagraph(false, true));
            }
            else if (oField instanceof ParaField)
            {
                if (oField.GetParagraph())
                    arrParagraphs.push(oField.GetParagraph());
            }
        }
        var nField;
        if(arrParagraphs.length > 0)
        {
            if (!this.Document_Is_SelectionLocked(changestype_None, {
                Type      : changestype_2_ElementsArray_and_Type,
                Elements  : arrParagraphs,
                CheckType : changestype_Paragraph_Content
            }))
            {
                for(nField = 0; nField < aFields.length; ++nField)
                {
                    if(aFields[nField].Update)
                    {
                        aFields[nField].Update(false, false);
                    }
                }
            }
        }
        NewParagraph.MoveCursorToEndPos();
        NewParagraph.Document_SetThisElementCurrent(true);
    }
    this.Recalculate();
    this.FinalizeAction();
};
/**
 * Выделяем перемещенный или удаленный после перемещения текст
 * @param sMoveId {string} идентификатор переноса
 * @param isFrom {boolean} true - выделяем удаленный текст, false - выделяем перемещенный текст
 * @param [isSetCurrentChange=false] {boolean} Выставлять или нет данное изменение текущим
 * @param [isUpdateInterface=true] {boolean} Обновлять ли интерфейс
 */
CDocument.prototype.SelectTrackMove = function(sMoveId, isFrom, isSetCurrentChange, isUpdateInterface)
{
	var oManager = this.GetTrackRevisionsManager();

	function private_GetDocumentPosition(oMark)
	{
		if (!oMark.IsUseInDocument())
			return null;

		if (oMark instanceof CParaRevisionMove)
		{
			return oMark.GetDocumentPositionFromObject();
		}
		else if (oMark instanceof AscWord.CRunRevisionMove && oMark.GetRun())
		{
			var oRun      = oMark.GetRun();
			var arrPos    = oRun.GetDocumentPositionFromObject();
			var nInRunPos = oRun.GetElementPosition(oMark);

			if (oMark.IsStart())
				arrPos.push({Class : oRun, Position : nInRunPos});
			else
				arrPos.push({Class : oRun, Position : nInRunPos + 1});

			return arrPos;
		}

		return null;
	}

	// TODO: Можно отказаться вообще от MoveMarks в пользу сбора изменений через функцию GetAllMoveChanges

	var oMarks = oManager.GetMoveMarks(sMoveId);
	if (oMarks)
	{
		var oStart = isFrom ? oMarks.From.Start : oMarks.To.Start;
		var oEnd   = isFrom ? oMarks.From.End : oMarks.To.End;

		if (oStart && oEnd)
		{
			var oStartDocPos = private_GetDocumentPosition(oStart);
			var oEndDocPos   = private_GetDocumentPosition(oEnd);

			if (oStartDocPos && oEndDocPos && oStartDocPos[0].Class === oEndDocPos[0].Class)
			{
				var oDocContent = oStartDocPos[0].Class;
				oDocContent.SetSelectionByContentPositions(oStartDocPos, oEndDocPos);

				if (isSetCurrentChange)
				{
					var oChange = oStart.GetReviewChange();
					if (oChange)
						oManager.SetCurrentChange(oManager.CollectMoveChange(oChange));
				}
			}
			else
			{
				this.RemoveSelection();
			}

			this.UpdateSelection();

			if (false !== isUpdateInterface)
				this.UpdateInterface(true);
		}
	}
};
/**
 * Удаляем метки переноса
 * @param sMoveId
 */
CDocument.prototype.RemoveTrackMoveMarks = function(sMoveId)
{
	if (sMoveId === this.TrackMoveId)
		return;

	var oManager = this.GetTrackRevisionsManager();
	var oMarks   = oManager.GetMoveMarks(sMoveId);

	if (!oMarks)
		return;

	var oTrackMoveProcess = oManager.GetProcessTrackMove();
	if (oTrackMoveProcess && oTrackMoveProcess.GetMoveId() === sMoveId)
		return;

	var oDocState = this.GetSelectionState();

	this.SelectTrackMove(sMoveId, true, false, false);
	this.AcceptRevisionChanges(c_oAscRevisionsChangeType.MoveMarkRemove, false);

	this.SelectTrackMove(sMoveId, false, false, false);
	this.AcceptRevisionChanges(c_oAscRevisionsChangeType.MoveMarkRemove, false);

	this.SetSelectionState(oDocState);

	if (!this.Action.Start)
		return;

	if (!this.Action.Additional.TrackMove)
		this.Action.Additional.TrackMove = {};


	if (this.Action.Additional.TrackMove[sMoveId])
		return;

	this.Action.Additional.TrackMove[sMoveId] = oMarks;
};
/**
 * Выставляем параметр для насильного запрета отрисовки рамок для ContentControl
 * @param {boolean} isHide
 */
CDocument.prototype.SetForceHideContentControlTrack = function(isHide)
{
	this.ForceHideCCTrack = isHide;
	this.UpdateTracks();
};
/**
 * Выставлен ли насильный запрет отрисовки рамок ContentControl
 * @returns {boolean}
 */
CDocument.prototype.IsForceHideContentControlTrack = function()
{
	return this.ForceHideCCTrack;
};
/**
 * Проверяем, выделен ли перенесенный текст
 * @param {?CSelectedElementsInfo} oSelectedInfo
 * @returns {string | null} id переноса
 */
CDocument.prototype.private_CheckTrackMoveSelection = function(oSelectedInfo)
{
	if (!oSelectedInfo)
		oSelectedInfo = this.GetSelectedElementsInfo({CheckAllSelection : true});

	if (oSelectedInfo.HaveNotReviewedContent() || oSelectedInfo.HaveRemovedInReview() || !oSelectedInfo.HaveAddedInReview())
		return null;

	var arrParagraphs = this.GetSelectedParagraphs();

	if (arrParagraphs.length <= 0)
		return null;

	var sMarkS = arrParagraphs[0].CheckTrackMoveMarkInSelection(true);
	var sMarkE = arrParagraphs[arrParagraphs.length - 1].CheckTrackMoveMarkInSelection(false);

	if (sMarkS === sMarkE)
		return sMarkS;

	return null;
};
/**
 * Проверяем попадает ли в выделение какой-нибудь перенос
 * @returns {string | null} id переноса
 */
CDocument.prototype.CheckTrackMoveInSelection = function()
{
	var oSelectedInfo = this.GetSelectedElementsInfo({CheckAllSelection : true});

	var arrMoveMarks = oSelectedInfo.GetTrackMoveMarks();
	if (arrMoveMarks.length > 0)
		return arrMoveMarks[0].GetMarkId();

	if (!oSelectedInfo.HaveRemovedInReview() && !oSelectedInfo.HaveAddedInReview())
		return null;

	var arrParagraphs = this.GetSelectedParagraphs();

	if (arrParagraphs.length <= 0)
		return null;

	var sMarkS = arrParagraphs[0].CheckTrackMoveMarkInSelection(true);
	var sMarkE = arrParagraphs[arrParagraphs.length - 1].CheckTrackMoveMarkInSelection(false);

	if (sMarkS && sMarkS === sMarkE)
		return sMarkS;

	sMarkS = arrParagraphs[0].CheckTrackMoveMarkInSelection(true, false);
	sMarkE = arrParagraphs[arrParagraphs.length - 1].CheckTrackMoveMarkInSelection(false, false);

	if (sMarkS && sMarkS === sMarkE)
		return sMarkS;

	return null;
};
/**
 * Сообщаем, что в конце действия нужно будет проверить размер содержимого формы
 * @param {CSdtBase} oForm
 */
CDocument.prototype.CheckFormAutoFit = function(oForm)
{
	if (!this.Action.Additional.FormAutoFit)
		this.Action.Additional.FormAutoFit = [];

	this.Action.Additional.FormAutoFit.push(oForm);
};
/**
 * Выставляем настройку выделять знак параграфа, когда выделено все его содержимое
 * @param {boolean} isUse
 */
CDocument.prototype.SetSmartParagraphSelection = function(isUse)
{
	this.SmartParagraphSelection = isUse;
};
/**
 * Выделять ли знак параграфа, когда выделено все его содержимое
 * @returns {boolean}
 */
CDocument.prototype.IsSmartParagraphSelection = function()
{
	return this.SmartParagraphSelection;
};
/**
 * Функция для рисования таблицы с помощью мыши
 */
CDocument.prototype.DrawTable = function()
{
	if (!this.DrawTableMode.Draw && !this.DrawTableMode.Erase)
		return;

	if (!this.DrawTableMode.Table)
	{
		if (!this.DrawTableMode.Draw || Math.abs(this.DrawTableMode.StartX - this.DrawTableMode.EndX) < 1 || Math.abs(this.DrawTableMode.StartY - this.DrawTableMode.EndY) < 1)
		{
			if (this.DrawTableMode.StartX === this.DrawTableMode.EndX && this.DrawTableMode.StartY === this.DrawTableMode.EndY)
			{
				this.DrawTableMode.Draw && this.Api.sync_TableDrawModeCallback(false);
				this.DrawTableMode.Erase && this.Api.sync_TableEraseModeCallback(false);
				this.UpdateCursorType(this.CurPos.RealX, this.CurPos.RealY, this.CurPage, new AscCommon.CMouseEventHandler());
			}

			return;
		}

		var oSelectionState = this.GetSelectionState();

		this.RemoveSelection();

		this.CurPage = this.DrawTableMode.Page;
		this.MoveCursorToXY(this.DrawTableMode.StartX, this.DrawTableMode.StartY, false);

		if (!this.IsSelectionLocked(AscCommon.changestype_Document_Content_Add))
		{
			this.StartAction(AscDFH.historydescription_Document_DrawNewTable, oSelectionState);

			var oTable = this.AddInlineTable(1, 1, -1);
			if (oTable && oTable.GetRowsCount() > 0)
			{
				oTable.Set_Inline(false);
				oTable.Set_Distance(3.2, 0, 3.2, 0);
				oTable.Set_PositionH(c_oAscHAnchor.Page, false, Math.min(this.DrawTableMode.StartX, this.DrawTableMode.EndX));
				oTable.Set_PositionV(c_oAscVAnchor.Page, false, Math.min(this.DrawTableMode.StartY, this.DrawTableMode.EndY));
				oTable.Set_TableW(tblwidth_Mm, Math.abs(this.DrawTableMode.EndX - this.DrawTableMode.StartX));
				oTable.GetRow(0).SetHeight(Math.abs(this.DrawTableMode.EndY - this.DrawTableMode.StartY), Asc.linerule_AtLeast);
			}

			this.FinalizeAction();
			this.Api.SetTableDrawMode(true);
		}

		return;
	}

	if (!this.IsSelectionLocked(changestype_None, {
			Type      : changestype_2_ElementsArray_and_Type,
			Elements  : [this.DrawTableMode.Table],
			CheckType : AscCommon.changestype_Table_Properties
		}))
	{
		var isDraw  = this.DrawTableMode.Draw;
		var isErase = this.DrawTableMode.Erase;

		var oTable = this.DrawTableMode.Table;

		this.StartAction(AscDFH.historydescription_Document_DrawTable);
		var bKeepDrawMode = oTable.DrawTableCells(this.DrawTableMode.StartX, this.DrawTableMode.StartY, this.DrawTableMode.EndX, this.DrawTableMode.EndY, this.DrawTableMode.TablePageStart, this.DrawTableMode.TablePageEnd, isDraw);

		if (bKeepDrawMode === false)
		{
			isDraw  = false;
			isErase = false;
		}

		if (oTable.GetRowsCount() <= 0 && oTable.GetParent())
		{
			var oParentDocContent = oTable.GetParent();

			this.RemoveSelection();
			oTable.PreDelete();

			var nPos = oTable.GetIndex();
			oParentDocContent.RemoveFromContent(nPos, 1);

			oParentDocContent.SetDocPosType(docpostype_Content);

			if (nPos >= oParentDocContent.Content.length)
			{
				oParentDocContent.CurPos.ContentPos = nPos - 1;
				oParentDocContent.Content[nPos - 1].MoveCursorToEndPos();
			}
			else
			{
				if (nPos < 0)
					nPos = 0;

				oParentDocContent.CurPos.ContentPos = nPos;
				oParentDocContent.Content[nPos].MoveCursorToEndPos();
			}
		}

		this.Recalculate();
		this.FinalizeAction();

		if (isDraw)
			this.Api.SetTableDrawMode(true);
		else if (isErase)
			this.Api.SetTableEraseMode(true);
		this.UpdateCursorType(this.CurPos.RealX, this.CurPos.RealY, this.CurPage, new AscCommon.CMouseEventHandler());
	}
};
/**
 * Добавляем текст в текущую позицию с заданными текстовыми настройками
 * @param {string} sText
 * @param {?AscCommon.CAddTextSettings} oSettings
 */
CDocument.prototype.AddTextWithPr = function(sText, oSettings)
{
	if (!oSettings)
		oSettings = new AscCommon.CAddTextSettings();

	if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_AddText))
	{
		this.StartAction(AscDFH.historydescription_Document_AddTextWithProperties);

		this.RemoveBeforePaste();

		var oCurrentTextPr = this.GetDirectTextPr();

		var oParagraph = this.GetCurrentParagraph();
		if (oParagraph && oParagraph.GetParent())
		{
			if (oSettings.IsWrapWithSpaces())
			{
				let arrCodePoints = sText.codePointsArray();
				if (arrCodePoints.length)
				{
					let isSpaceBefore = !AscCommon.IsSpace(arrCodePoints[0]);
					let isSpaceAfter  = !AscCommon.IsSpace(arrCodePoints[arrCodePoints.length - 1]);

					if (isSpaceAfter)
					{
						let oRunItem = oParagraph.GetNextRunElement();
						if (!oRunItem || oRunItem.IsSpace() || oRunItem.IsTab() || oRunItem.IsBreak() || oRunItem.IsParaEnd())
							isSpaceAfter = false;
					}
					if (isSpaceBefore)
					{
						let oRunItem = oParagraph.GetPrevRunElement();
						if (!oRunItem || oRunItem.IsSpace() || oRunItem.IsTab() || oRunItem.IsBreak())
							isSpaceBefore = false;
					}

					if (isSpaceAfter)
						sText += " ";
					if (isSpaceBefore)
						sText = " " + sText;
				}
			}

			var oTempPara = new Paragraph(this.GetDrawingDocument(), oParagraph.GetParent());
			var oRun      = new ParaRun(oTempPara, false);
			oRun.AddText(sText);
			oTempPara.AddToContent(0, oRun);

			oRun.SetPr(oCurrentTextPr.Copy());

			let oTextPr = oSettings.GetTextPr();
			if (oTextPr)
				oRun.ApplyPr(oTextPr);

			var oAnchorPos = oParagraph.GetCurrentAnchorPosition();

			var oSelectedContent = new AscCommonWord.CSelectedContent();
			oSelectedContent.Add(new AscCommonWord.CSelectedElement(oTempPara, false));
			oSelectedContent.EndCollect(this);
			oSelectedContent.ForceInlineInsert();
			oSelectedContent.PlaceCursorInLastInsertedRun(!oSettings.IsMoveCursorOutside());
			oSelectedContent.Insert(oAnchorPos);
            this.CheckCurrentTextObjectExtends();
		}

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
};
/**
 * Добавляем спец символ
 * @param oPr
 */
CDocument.prototype.AddSpecialSymbol = function(oPr)
{
	var oItem;
	if (oPr)
	{
		if (true === oPr["NonBreakingHyphen"])
		{
			oItem = AscWord.CreateNonBreakingHyphen();
		}
	}

	if (!oItem)
		return;

	if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_AddText))
	{
		this.StartAction(AscDFH.historydescription_Document_AddTextWithProperties);

		this.AddToParagraph(oItem);

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
};
/**
 * Список позиций, которые мы собираемся отслеживать
 * @param arrPositions
 */
CDocument.prototype.TrackDocumentPositions = function(arrPositions)
{
	this.CollaborativeEditing.Clear_DocumentPositions();

	for (var nIndex = 0, nCount = arrPositions.length; nIndex < nCount; ++nIndex)
	{
		this.CollaborativeEditing.Add_DocumentPosition(arrPositions[nIndex]);
	}
};
/**
 * Обновляем отслеживаемые позиции
 * @param arrPositions
 */
CDocument.prototype.RefreshDocumentPositions = function(arrPositions)
{
	for (var nIndex = 0, nCount = arrPositions.length; nIndex < nCount; ++nIndex)
	{
		this.CollaborativeEditing.Update_DocumentPosition(arrPositions[nIndex]);
	}
};
/**
 * Конвертируем позицию в документе в специальную позицию NearestPos
 * @param oDocPos
 * @returns {NearestPos}
 */
CDocument.prototype.DocumentPositionToAnchorPosition = function(oDocPos)
{
	var oRun, nInRunPos = -1;
	for (var nIndex = oDocPos.length - 1; nIndex >= 0; --nIndex)
	{
		if (oDocPos[nIndex].Class instanceof ParaRun)
		{
			oRun      = oDocPos[nIndex].Class;
			nInRunPos = oDocPos[nIndex].Position;
			break;
		}
	}

	if (!oRun)
		return;

	var oParaContentPos = oRun.GetParagraphContentPosFromObject(nInRunPos);
	var oParagraph      = oRun.GetParagraph();

	if (!oParagraph || !oParaContentPos)
		return;

	return oParaContentPos.ToAnchorPos(oParagraph);
};
/**
 * Конвертируем NearestPos в позицию документа
 * @param oAnchorPos {NearestPos}
 * @returns {*}
 */
CDocument.prototype.AnchorPositionToDocumentPosition = function(oAnchorPos)
{
	var oParagraph = oAnchorPos.Paragraph;
	var oRun       = oParagraph.GetElementByPos(oAnchorPos.ContentPos);

	if (!oRun || !(oRun instanceof ParaRun))
		return null;

	var nInRunPos = oAnchorPos.ContentPos.Get(oAnchorPos.ContentPos.GetDepth());

	var oDocPos = oRun.GetDocumentPositionFromObject();
	oDocPos.push({Class : oRun, Position : nInRunPos});

	return oDocPos;
};
/**
 * Выделена ли сейчас автофигура (проверяем с учетом колонтитулов)
 * @returns {boolean}
 */
CDocument.prototype.IsDrawingSelected = function()
{
	return (docpostype_DrawingObjects === this.GetDocPosType() || (docpostype_HdrFtr === this.CurPos.Type && null != this.HdrFtr.CurHdrFtr && docpostype_DrawingObjects === this.HdrFtr.CurHdrFtr.Content.CurPos.Type));
};
/**
 * Получаем список таблицы на заданной абсолютной странице
 * @param {number} nPageAbs
 * @returns {Array}
 */
CDocument.prototype.GetAllTablesOnPage = function(nPageAbs)
{
	var arrTables = [];

	if (!this.Pages[nPageAbs])
		return [];

	if (docpostype_HdrFtr === this.GetDocPosType())
	{
		this.HdrFtr.GetAllTablesOnPage(nPageAbs, arrTables);
	}
	else
	{
		var nStartPos = this.Pages[nPageAbs].Pos;
		var nEndPos   = this.Pages[nPageAbs].EndPos;

		for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
		{
			this.Content[nPos].GetAllTablesOnPage(nPageAbs, arrTables);
		}

		this.Footnotes.GetAllTablesOnPage(nPageAbs, arrTables);
		this.Endnotes.GetAllTablesOnPage(nPageAbs, arrTables);
	}

	return arrTables;
};
/**
 * Конвертируем формулу из старого формата в новый
 * @param {ParaDrawing} oEquation
 * @param {boolean} isAll
 */
CDocument.prototype.ConvertEquationToMath = function(oEquation, isAll)
{
	if (isAll)
	{
		var arrParagraphs = [];
		var arrDrawings = this.GetAllDrawingObjects();

		for (var nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
		{
			if (arrDrawings[nIndex].IsMathEquation())
			{
				var oParagraph = arrDrawings[nIndex].GetParagraph();
				var isAdd = true;
				for (var nParaIndex = 0, nParasCount = arrParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
				{
					if (arrParagraphs[nParaIndex] === oParagraph)
					{
						isAdd = false;
						break;
					}
				}

				if (isAdd)
					arrParagraphs.push(oParagraph);
			}
		}

		if (arrParagraphs.length < 0)
			return;

		if (!this.IsSelectionLocked(AscCommon.changestype_None, {
				Type      : AscCommon.changestype_2_ElementsArray_and_Type,
				Elements  : arrParagraphs,
				CheckType : AscCommon.changestype_Paragraph_Content
			}))
		{
			this.StartAction(AscDFH.historydescription_Document_ConvertOldEquation);
			this.RemoveSelection();

			for (var nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
			{
				if (arrDrawings[nIndex].IsMathEquation())
					arrDrawings[nIndex].ConvertToMath(arrDrawings[nIndex] === oEquation);
			}

			this.Recalculate();
			this.UpdateSelection();
			this.UpdateInterface();
			this.FinalizeAction();
		}
	}
	else if (oEquation)
	{
		if (!oEquation.IsMathEquation())
			return;

		var oParagraph = oEquation.GetParagraph();

		if (!this.IsSelectionLocked(AscCommon.changestype_None, {
				Type      : AscCommon.changestype_2_Element_and_Type,
				Element   : oParagraph,
				CheckType : AscCommon.changestype_Paragraph_Content
			}))
		{
			this.StartAction(AscDFH.historydescription_Document_ConvertOldEquation);
			this.RemoveSelection();

			oEquation.ConvertToMath(true);

			this.Recalculate();
			this.UpdateSelection();
			this.UpdateInterface();
			this.FinalizeAction();
		}
	}
};
CDocument.prototype.ConvertMathDisplayMode = function(isInline)
{
	let oMath = this.GetCurrentMath();
	if (!oMath)
		return;

	let oParagraph = oMath.GetParagraph();
	if (!oParagraph)
		return;

	if (!this.IsSelectionLocked(AscCommon.changestype_None, {
		Type      : changestype_2_ElementsArray_and_Type,
		Elements  : [oParagraph],
		CheckType : AscCommon.changestype_Paragraph_Content
	}, true, false))
	{
		this.StartAction(AscDFH.historydescription_Document_ConvertMathDisplayMode);

		if (isInline)
			oMath.ConvertToInlineMode();
		else
			oMath.ConvertToDisplayMode();

		this.Recalculate();
		this.UpdateSelection();
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
CDocument.prototype.GetCurrentMath = function()
{
	return this.GetSelectedElementsInfo().GetMath();
};
/**
 * @returns {CGlossaryDocument}
 */
CDocument.prototype.GetGlossaryDocument = function()
{
	return this.GlossaryDocument;
};
/**
 * Добавляем формулу
 */
CDocument.prototype.AddParaMath = function(nType)
{
	if ((undefined === nType || c_oAscMathType.Default_Text === nType) && (!this.IsSelectionUse() || !this.IsTextSelectionUse()))
	{
		this.Remove(1, true, false, true);
		var oCC = this.AddContentControl(c_oAscSdtLevelType.Inline);
		if (!oCC)
			return;

		oCC.ApplyContentControlEquationPr();
		oCC.SelectContentControl();
	}
	else
	{
		let textPr = this.GetDirectTextPr();
		this.AddToParagraph(new AscCommonWord.MathMenu(nType, textPr));
	}

	this.UpdateSelection();
	this.UpdateInterface();
};
CDocument.prototype.OnChangeContentControl = function(oControl)
{
	if (!this.Action.Start && !this.Action.UndoRedo)
		return;

	if (!this.Action.Additional.ContentControlChange)
		this.Action.Additional.ContentControlChange = {};

	let sId = oControl.GetId();
	if (this.Action.Additional.ContentControlChange[sId])
		return;

	this.Action.Additional.ContentControlChange[sId] = oControl;
};
/**
 * @returns {?AscOForm.CDocument}
 */
CDocument.prototype.GetOFormDocument = function()
{
	return this.OFormDocument;
};
/**
 * @returns {?AscOForm.CUserMaster}
 */
CDocument.prototype.GetCurrentOFormUserMaster = function()
{
	return this.OFormDocument ? this.OFormDocument.getCurrentUserMaster() : null;
};
/**
 * @returns {AscWord.CFormsManager}
 */
CDocument.prototype.GetFormsManager = function()
{
	return this.FormsManager;
};
/**
 * Сохраняем информацию о том, что форма с заданным ключом была изменена
 * @param {CInlineLevelSdt | CBlockLevelSdt} oForm
 */
CDocument.prototype.OnChangeForm = function(oForm)
{
	if (!oForm
		|| !oForm.IsForm()
		|| !this.Action.Start
		|| (this.Action.Additional && true === this.Action.Additional.FormChangeStart))
		return;

	let sKey = oForm.IsRadioButton() ? oForm.GetRadioButtonGroupKey() : oForm.GetFormKey();

	if (!this.Action.Additional.ValidateForm)
		this.Action.Additional.ValidateForm = {};

	this.Action.Additional.ValidateForm[oForm.GetId()] = oForm;

	let oMainForm = oForm.GetMainForm();
	if (oForm !== oMainForm)
	{
		sKey  = oMainForm.GetFormKey();
		oForm = oMainForm;
	}

	if (!sKey)
		return;

	if (!this.Action.Additional.FormChange)
		this.Action.Additional.FormChange = {};

	if (this.Action.Additional.FormChange[sKey])
		return;

	this.Action.Additional.FormChange[sKey] = oForm;
};
/**
 * Удаляем все дополнительные обработки по изменению значения формы в конце действия
 */
CDocument.prototype.ClearActionOnChangeForm = function()
{
	if (this.Action.Additional.FormChange)
		delete this.Action.Additional.FormChange;
};
/**
 * Сохраняем изменение, что радиогруппа должна иметь заданный статус Required
 * @param {string} sGroupKey
 * @param {boolean} isRequired
 */
CDocument.prototype.OnChangeRadioRequired = function(sGroupKey, isRequired)
{
	if (!this.Action.Start)
		return;
	
	if (!this.Action.Additional.RadioRequired)
		this.Action.Additional.RadioRequired = {};

	this.Action.Additional.RadioRequired[sGroupKey] = isRequired;
};
/**
 * Очищаем все специальные формы до плейсхолдеров
 * @param {boolean} [isClearAllContentControls=false] Очищаем ли вообще все контролы
 */
CDocument.prototype.ClearAllSpecialForms = function(isClearAllContentControls)
{
	var arrContentControls;
	if (isClearAllContentControls)
		arrContentControls = this.GetAllContentControls();
	else
		arrContentControls = this.FormsManager.GetAllForms();

	var arrParagraphs = [];
	for (var nIndex = 0, nCount = arrContentControls.length; nIndex < nCount; ++nIndex)
	{
		var oControl = arrContentControls[nIndex];
		if (oControl.IsInlineLevel() && oControl.GetParagraph())
			arrParagraphs.push(oControl.GetParagraph());
		else if (oControl.IsBlockLevel())
			oControl.GetAllParagraphs({}, arrParagraphs);

		oControl.SkipSpecialContentControlLock(true);
	}

	var oCurrentControl = isClearAllContentControls ? this.GetSelectedElementsInfo().GetInlineLevelSdt() : this.GetSelectedElementsInfo().GetForm();

	if (!this.IsSelectionLocked(AscCommon.changestype_None, {
		Type : changestype_2_ElementsArray_and_Type,
		Elements : arrParagraphs,
		CheckType : AscCommon.changestype_Paragraph_Content
	}, true, this.IsFillingFormMode()))
	{
		this.StartAction(AscDFH.historydescription_Document_ClearAllSpecialForms);

		if (oCurrentControl)
			this.RemoveSelection();

		for (var nIndex = 0, nCount = arrContentControls.length; nIndex < nCount; ++nIndex)
		{
			var oControl = arrContentControls[nIndex];

			if (oControl.IsBuiltInTableOfContents() || oControl.IsBuiltInWatermark())
				continue;

			if (!oCurrentControl && oControl.IsUseInDocument())
				oCurrentControl = oControl;

			oControl.ClearContentControlExt();
			oControl.RemoveSelection();
		}


		if (oCurrentControl)
			oCurrentControl.SelectContentControl();

		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
	
	for (let iCC = 0, nCC = arrContentControls.length; iCC < nCC; ++iCC)
	{
		arrContentControls[iCC].SkipSpecialContentControlLock(false);
	}
};
/**
 * Конвертируем
 * @param sId
 * @param isToFixed
 * @returns {boolean}
 */
CDocument.prototype.ConvertFormFixedType = function(sId, isToFixed)
{
	var oForm = this.GetContentControl(sId);
	if (!oForm || !oForm.IsForm())
		return false;

	oForm = oForm.GetMainForm();
	if (!oForm)
		return false;

	var isLocked   = false;
	var oParagraph = oForm.GetParagraph();
	if (oParagraph)
	{
		if (isToFixed && oParagraph.GetParentShape())
			return false;
		
		isLocked = this.IsSelectionLocked(AscCommon.changestype_None, {
			Type      : AscCommon.changestype_2_ElementsArray_and_Type,
			Elements  : [oParagraph],
			CheckType : AscCommon.changestype_Paragraph_Properties
		}, false, false);
	}

	if (!isLocked)
	{
		this.StartAction(AscDFH.historydescription_Document_ConvertFormFixedType);

		if (isToFixed)
		{
			let drawing = oForm.ConvertFormToFixed();
			
			this.RemoveSelection();
			if (drawing)
				drawing.SelectAsDrawing();
		}
		else
		{
			oForm.ConvertFormToInline();
			this.RemoveSelection();
			oForm.SelectContentControl();
		}
		
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.UpdateTracks();
		this.FinalizeAction();

		return true;
	}

	return false;
};
/**
 * Подсвечиваем ли обязательные поля
 * @returns {boolean}
 */
CDocument.prototype.IsHighlightRequiredFields = function()
{
	return this.HighlightRequiredFields;
};
/**
 * Выставляем нужно ли подсвечивать обязательные поля
 * @param {boolean} isHighlight
 */
CDocument.prototype.SetHighlightRequiredFields = function(isHighlight)
{
	this.HighlightRequiredFields = isHighlight;
};
/**
 * Получаем границу подсветки для обязательных полей
 * @returns {CDocumentBorder}
 */
CDocument.prototype.GetRequiredFieldsBorder = function()
{
	return this.RequiredFieldsBorder;
};
/**
 *  Функция, которая используется для отрисовки символа конца секции
 * @param nType {number} тип разрыва секции
 * @returns {object}
 */
CDocument.prototype.GetSectionEndMarkPr = function(nType)
{
	if (this.SectionEndMark[nType])
	{
		return this.SectionEndMark[nType];
	}
	else
	{
		var strSectionBreak = "";
		switch (nType)
		{
			case c_oAscSectionBreakType.Column     :
				strSectionBreak = " End of Section ";
				break;
			case c_oAscSectionBreakType.Continuous :
				strSectionBreak = " Section Break (Continuous) ";
				break;
			case c_oAscSectionBreakType.EvenPage   :
				strSectionBreak = " Section Break (Even Page) ";
				break;
			case c_oAscSectionBreakType.NextPage   :
				strSectionBreak = " Section Break (Next Page) ";
				break;
			case c_oAscSectionBreakType.OddPage    :
				strSectionBreak = " Section Break (Odd Page) ";
				break;
		}

		g_oTextMeasurer.SetFont({
			FontFamily : {Name : "Courier New", Index : -1},
			FontSize   : 8,
			Italic     : false,
			Bold       : false
		});

		this.SectionEndMark[nType] = {
			String      : strSectionBreak,
			Widths      : [],
			ColonSymbol : 58, // ":"
			ColonWidth  : 0,
			StringWidth : 0
		};

		var oPr = this.SectionEndMark[nType];
		for (var nPos = 0, nLen = strSectionBreak.length; nPos < nLen; ++nPos)
		{
			var Val = g_oTextMeasurer.Measure(strSectionBreak[nPos]).Width;

			oPr.StringWidth += Val;
			oPr.Widths[nPos] = Val;
		}

		var strSymbol = ":";
		oPr.ColonWidth = g_oTextMeasurer.Measure(strSymbol).Width * 2 / 3;

		return oPr;
	}
};
/**
 * Получаем закэшированную информацию для отрисовки нумерации строк
 * @returns {CDocumentLineNumbersInfo}
 */
CDocument.prototype.GetLineNumbersInfo = function()
{
	return this.LineNumbersInfo;
};
/**
 * Устанавливаем настройки для нумерации строк
 * @param {Asc.c_oAscSectionApplyType} nApplyType
 * @param {?CSectionLnNumType} oProps  (если null или undefined, то из заданных разделов удаляем нумерацию строк)
 */
CDocument.prototype.SetLineNumbersProps = function(nApplyType, oProps)
{
	if (!this.IsSelectionLocked(AscCommon.changestype_Document_SectPr))
	{
		this.StartAction(AscDFH.historydescription_Document_SetLineNumbersProps);

		var arrSectPr = this.GetSectionsByApplyType(nApplyType);
		if (arrSectPr.length > 0)
		{
			if (undefined === oProps || null === oProps)
			{
				for (var nIndex = 0, nCount = this.SectionsInfo.GetCount(); nIndex < nCount; ++nIndex)
				{
					var oSectPr = this.SectionsInfo.Get(nIndex).SectPr;
					oSectPr.RemoveLineNumbers();
				}
			}
			else
			{
				var nCountBy  = oProps.GetCountBy();
				var nDistance = oProps.GetDistance();
				var nStart    = oProps.GetStart();
				var nRestart  = oProps.GetRestart();

				for (var nIndex = 0, nCount = arrSectPr.length; nIndex < nCount; ++nIndex)
				{
					var oSectPr = arrSectPr[nIndex];

					var _nCountBy  = undefined === nCountBy ? oSectPr.GetLineNumbersCountBy() : nCountBy;
					var _nDistance = undefined === nDistance ? oSectPr.GetLineNumbersDistance() : nDistance;
					var _nStart    = undefined === nStart ? oSectPr.GetLineNumbersStart() : nStart;
					var _nRestart  = undefined === nRestart ? oSectPr.GetLineNumbersRestart() : nRestart;

					oSectPr.SetLineNumbers(_nCountBy, _nDistance, _nStart, _nRestart);
				}
			}
		}

		this.Recalculate();
		this.UpdateInterface();
		this.FinalizeAction();
	}
};
/**
 * Получаем настройки нумерации строк текущего раздела
 * @returns {?CSectionLnNumType|null}
 */
CDocument.prototype.GetLineNumbersProps = function()
{
	var oSectPr = this.Layout.GetSectionByPos(this.CurPos.ContentPos);
	return oSectPr.HaveLineNumbers() ? oSectPr.GetLineNumbers().Copy() : null;
};
/**
 * Получаем список секции на основе по заданному типу
 * @param {Asc.c_oAscSectionApplyType} nType
 * @returns {Array.CSectionPr}
 */
CDocument.prototype.GetSectionsByApplyType = function(nType)
{
	if (Asc.c_oAscSectionApplyType.Current === nType)
	{
		if (docpostype_Content !== this.GetDocPosType())
			return [];

		if (this.IsSelectionUse())
		{
			var arrParagraphs = this.GetSelectedParagraphs();
			if (arrParagraphs.length > 0)
			{
				var oSectPrS = arrParagraphs[0].GetDocumentSectPr();
				var oSectPrE = arrParagraphs[arrParagraphs.length - 1].GetDocumentSectPr();

				var nStartIndex = this.SectionsInfo.Find(oSectPrS);
				var nEndIndex   = this.SectionsInfo.Find(oSectPrE);
				if (-1 === nStartIndex || -1 === nEndIndex)
					return [oSectPrE];

				if (nStartIndex > nEndIndex)
				{
					var nTemp = nStartIndex;
					nStartIndex = nEndIndex;
					nEndIndex   = nTemp;
				}

				var arrSectPr = [];
				for (var nIndex = nStartIndex; nIndex <= nEndIndex; ++nIndex)
				{
					arrSectPr.push(this.SectionsInfo.Get(nIndex).SectPr);
				}
				return arrSectPr;
			}
		}
		else
		{

			var oParagraph = this.GetCurrentParagraph();
			if (!oParagraph)
				return [];

			var oSectPr = oParagraph.GetDocumentSectPr();
			if (oSectPr)
				return [oSectPr];
		}
	}
	else if (Asc.c_oAscSectionApplyType.ToEnd === nType)
	{
		if (docpostype_Content !== this.GetDocPosType())
			return [];

		var oParagraph = this.GetCurrentParagraph();
		if (!oParagraph)
			return [];

		var oSectPr = oParagraph.GetDocumentSectPr();
		if (!oSectPr)
			return [];

		var nIndex = this.SectionsInfo.Find(oSectPr);
		if (-1 === nIndex)
			return [oSectPr];

		var arrSectPr = [];
		for (nCount = this.SectionsInfo.GetCount(); nIndex < nCount; ++nIndex)
		{
			arrSectPr.push(this.SectionsInfo.Get(nIndex).SectPr);
		}
		return arrSectPr;
	}
	else if (Asc.c_oAscSectionApplyType.All === nType)
	{
		var arrSectPr = [];
		for (var nIndex = 0, nCount = this.SectionsInfo.GetCount(); nIndex < nCount; ++nIndex)
		{
			arrSectPr.push(this.SectionsInfo.Get(nIndex).SectPr);
		}
		return arrSectPr;
	}

	return [];
};
/**
 * Получаем разделитель для списков
 * @returns {?string}
 */
CDocument.prototype.GetListSeparator = function()
{
	return this.Settings.ListSeparator ? this.Settings.ListSeparator : ",";
};
/**
 * Получаем разделитель десятичного числа
 * @returns {?string}
 */
CDocument.prototype.GetDecimalSymbol = function()
{
	return this.Settings.DecimalSymbol ? this.Settings.DecimalSymbol : ".";
};
/**
 * Меняем регистр выделенного текста
 * @param {Asc.c_oAscChangeTextCaseType} nCaseType
 */
CDocument.prototype.ChangeTextCase = function(nCaseType)
{
	let oState = null;
	if (!this.IsSelectionUse())
	{
		oState = this.SaveDocumentState(false);
		this.SelectCurrentWord();
	}

	if (!this.IsSelectionUse() || !this.IsTextSelectionUse())
	{
		if (oState)
			this.LoadDocumentState(oState);

		return;
	}

	if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_ChangeTextCase);

		let oChangeEngine = new AscCommonWord.CChangeTextCaseEngine(nCaseType);
		oChangeEngine.ProcessParagraphs(this.GetSelectedParagraphs());

		if (oState)
			this.LoadDocumentState(oState);

		this.UpdateSelection();
		this.Recalculate();
		this.FinalizeAction();
	}
	else if (oState)
	{
		this.LoadDocumentState(oState);
		this.UpdateSelection();
	}
};
/**
 * @returns {boolean}
 */
CDocument.prototype.IsShowShapeAdjustments = function()
{
	return (!!this.CanEdit());
};
/**
 * Рисовать ли трек у таблицы и давать ли возможность таскать границы
 * @returns {boolean}
 */
CDocument.prototype.IsShowTableAdjustments = function()
{
	return (!!this.CanEdit());
};
/**
 * Рисовать ли трек у таблицы и давать ли возможность таскать границы
 * @returns {boolean}
 */
CDocument.prototype.IsShowEquationTrack = function()
{
	return (!!this.CanEdit());
};
/**
 * Можем ли перетаскивать текст
 * @returns {boolean}
 */
CDocument.prototype.CanDragAndDrop = function()
{
	return (!!this.CanEdit()) &&
    !this.IsTextSelectionInSmartArt();
};
/**
 * Конвертируем выделенный текст в таблицу
 * @param oProps
 */
CDocument.prototype.ConvertTextToTable = function(oProps)
{
	if (!this.IsTextSelectionUse() || this.IsTableCellSelection())
		return;

	if (!this.IsSelectionLocked(AscCommon.changestype_Document_Contentt))
	{
		this.StartAction(AscDFH.historydescription_Document_ConvertTextToTable);

		var oSelectedContent = this.GetSelectedContent(true);
		if (!this.private_ConvertTextToTable(oProps, oSelectedContent))
		{
			this.History.Remove_LastPoint();
			this.FinalizeAction();
			return;
		}

		this.RemoveBeforePaste();
		var oParagraph  = this.GetCurrentParagraph();
		var oDocContent = oParagraph ? oParagraph.GetParent() : null;
		if (oSelectedContent
			&& oSelectedContent.Elements.length
			&& oDocContent
			&& oParagraph.IsUseInDocument()
			&& oDocContent.GetElement(oParagraph.GetIndex()) === oParagraph)
		{
			var nInsertPos = 0;
			if (oParagraph.IsCursorAtBegin())
			{
				nInsertPos = oParagraph.GetIndex();
			}
			else if (oParagraph.IsCursorAtEnd())
			{
				nInsertPos = oParagraph.GetIndex() + 1;
				oParagraph = null;
			}
			else
			{
				var oNewParagraph = oParagraph.Split();

				nInsertPos = oParagraph.GetIndex() + 1;
				oParagraph = null;

				if (!oNewParagraph.IsEmpty())
				{
					oDocContent.AddToContent(nInsertPos, oNewParagraph);
				}
			}

			var nCount = oSelectedContent.Elements.length;
			for (var nIndex = 0; nIndex < nCount; ++nIndex)
			{
				var oElement = oSelectedContent.Elements[nIndex].Element;
				oDocContent.AddToContent(nInsertPos + nIndex, oElement);
				oElement.SelectAll();
			}

			if (oParagraph && oParagraph.IsEmpty())
				oDocContent.RemoveFromContent(oParagraph.GetIndex(), 1);

			oDocContent.Selection.Use      = true;
			oDocContent.Selection.StartPos = nInsertPos;
			oDocContent.Selection.EndPos   = nInsertPos + nCount - 1;
			oDocContent.SetThisElementCurrent();
		}

		this.UpdateSelection();
		this.Recalculate();
		this.FinalizeAction();
	}
};
CDocument.prototype.private_ConvertTextToTable = function(oProps, oSelectedContent)
{
	var nCols = oProps.get_ColsCount();
	var nRows = oProps.get_RowsCount();

	var oEngine = new CTextToTableEngine();
	oEngine.SetConvertMode(oProps.get_SeparatorType(), oProps.get_Separator(), oProps.get_ColsCount());
	for (var nIndex = 0, nCount = oSelectedContent.Elements.length; nIndex < nCount; ++nIndex)
	{
		var oElement = oSelectedContent.Elements[nIndex].Element;
		if (oElement.IsBlockLevelSdt() && !oEngine.GetContentControl())
			oEngine.SetContentControl(oElement);

		oElement.CalculateTextToTable(oEngine);
	}
	oEngine.FinalizeConvert();

	var arrRows = oEngine.GetRows();
	if (arrRows.length <= 0)
		return false;

	nRows = Math.max(nRows, arrRows.length);

	var oFirstItem = oSelectedContent.Elements[0].Element.GetParent();
	var haveTable = oSelectedContent.HaveTable();
	oSelectedContent.Reset();
	var oItem = (oFirstItem === this) ? oFirstItem : this.GetCurrentParagraph().GetParent();
	var Grid = [];
	var W;
	if (oItem === this)
	{
		var SectPr = this.SectionsInfo.Get_SectPr(this.CurPos.ContentPos).SectPr;
		var PageFields = this.Get_PageFields(this.CurPage);
		var nAdd = this.GetCompatibilityMode() <= AscCommon.document_compatibility_mode_Word14 ?  2 * 1.9 : 0;
		W = (PageFields.XLimit - PageFields.X + nAdd);
		if (SectPr.Get_ColumnsCount() > 1)
			{
				for (var CurCol = 0, ColsCount = SectPr.Get_ColumnsCount(); CurCol < ColsCount; ++CurCol)
				{
					var ColumnWidth = SectPr.Get_ColumnWidth(CurCol);
					if (W > ColumnWidth)
						W = ColumnWidth;
				}
	
				W += nAdd;
			}
		W = Math.max(W, nCols * 2 * 1.9);
	}
	else
	{
		var nAdd = this.GetCompatibilityMode() <= AscCommon.document_compatibility_mode_Word14 && oItem.IsTableCellContent() ?  2 * 1.9 : 0;
		W = Math.max(oItem.XLimit - oItem.X + nAdd, nCols * 2 * 1.9);
	}

	for (var Index = 0; Index < nCols; Index++)
		Grid[Index] = W / nCols;
	
	var oTable = new CTable(this.DrawingDocument, this, true, nRows, nCols, Grid);
	for (var nCurRow = 0, nRowsCount = arrRows.length; nCurRow < nRowsCount; ++nCurRow)
	{
		var oRow = oTable.GetRow(nCurRow);
		if (!oRow)
			break;

		for (var nCurCell = 0, nCellsCount = arrRows[nCurRow].length; nCurCell < nCellsCount; ++nCurCell)
		{
			var oCell = oRow.GetCell(nCurCell);
			if (!oCell)
				break;

			var oCellContent = oCell.GetContent();
			oCellContent.Internal_Content_RemoveAll();
			for (var nItemPos = 0, nItemsCount = arrRows[nCurRow][nCurCell].length; nItemPos < nItemsCount; ++nItemPos)
				oCellContent.Internal_Content_Add(nItemPos, arrRows[nCurRow][nCurCell][nItemPos], false);
		}
	}

	oTable.SelectAll();
	if (oProps.get_AutoFitType() !== Asc.c_oAscTextToTableAutoFitType.Window)
	{
		var width = oProps.get_Fit();
		var oTableProps = new Asc.CTableProp();
		oTableProps.put_CellsWidth((width > 0) ? width : null);
		oTableProps.put_CellSelect(true);
		if (width === -10 && !haveTable)
			oTableProps.put_TableLayout(1);
		
		oTable.SetTableProps(oTableProps);
	}

	var oCC = oEngine.GetContentControl();
	if (oCC)
	{
		var oDocContent = oCC.GetContent();
		oDocContent.Internal_Content_RemoveAll();
		oDocContent.AddToContent(0, oTable, false);
		oSelectedContent.Add(new AscCommonWord.CSelectedElement(oCC, true));
	}
	else
	{
		oSelectedContent.Add(new AscCommonWord.CSelectedElement(oTable, true));
	}

	return true;
};
/**
* Подготовка к преобразованию текста в таблицу
 * @param oProps
 */
CDocument.prototype.GetConvertTextToTableProps = function(oProps)
{
	if (!oProps)
	{
		oProps = new Asc.CAscTextToTableProperties(this.GetSelectedContent(false));

		var oEngine = new CTextToTableEngine();
		oEngine.SetCheckSeparatorMode();
		for (var nIndex = 0, nCount = oProps.Selected.Elements.length; nIndex < nCount; ++nIndex)
		{
			var oElement = oProps.Selected.Elements[nIndex].Element;
			oElement.CalculateTextToTable(oEngine);
		}

		if (oEngine.HaveTab())
		{
			oProps.put_SeparatorType(Asc.c_oAscTextToTableSeparator.Tab);
		}
		else if (oEngine.HaveSemicolon())
		{
			oProps.put_SeparatorType(Asc.c_oAscTextToTableSeparator.Symbol);
			oProps.put_Separator(0x003B);
		}
	}
	oProps.CalculateTableSize();
	return oProps;
};
/**
 * Конвертируем текущую таблицу в текст
 * @param oProps
 */
CDocument.prototype.ConvertTableToText = function(oProps)
{
	var oTable = this.GetCurrentTable();
	if (!oTable)
		return;

	if (!this.IsSelectionLocked(AscCommon.changestype_None, {
		Type      : changestype_2_ElementsArray_and_Type,
		Elements  : [oTable],
		CheckType : changestype_Paragraph_Content
	}))
	{
		this.StartAction(AscDFH.historydescription_Document_ConvertTableToText);
		var FramePr = null;
		var NewContent = this.private_ConvertTableToText(oTable, oProps);
		var oNewContent = new AscCommonWord.CSelectedContent();

		if (!oTable.IsInline())
		{
			FramePr = new CFramePr();
			FramePr.HAnchor = oTable.PositionH.RelativeFrom;
			FramePr.VAnchor = oTable.PositionV.RelativeFrom;
			FramePr.X = oTable.PositionH.Value;
			FramePr.Y = oTable.PositionV.Value;
		}

		if (NewContent.before)
			oNewContent.Add(new AscCommonWord.CSelectedElement(NewContent.before, true));

		for (var i = 0; i < NewContent.content.length; i++)
		{
			oNewContent.Add(new AscCommonWord.CSelectedElement(NewContent.content[i], true));
			if (FramePr)
			{
				if (NewContent.content[i].IsParagraph())
				{
					NewContent.content[i].Set_FramePr2(FramePr);
				}
				else if (NewContent.content[i].IsTable() || NewContent.content[i].IsBlockLevelSdt())
				{
					var ArrParagraphs = NewContent.content[i].GetAllParagraphs();
					ArrParagraphs.forEach(function(item) {
						item.Set_FramePr2(FramePr);
					});
				}
			}
		}

		if (NewContent.after)
			oNewContent.Add(new AscCommonWord.CSelectedElement(NewContent.after, true));
		
		var oParent = oTable.GetParent();
		if (oNewContent && oParent)
		{
			var nIndex     = oTable.GetIndex();
			var oParagraph = new Paragraph(this.GetDrawingDocument());

			oParent.RemoveFromContent(nIndex, 1, true);
			oParent.AddToContent(nIndex, oParagraph, true);

			oParagraph.Document_SetThisElementCurrent(false);

			oNewContent.EndCollect(this);
			var oAnchorPos = oParagraph.GetCurrentAnchorPosition();
			if (oAnchorPos && oNewContent.CanInsert(oAnchorPos))
			{
				oNewContent.Insert(oAnchorPos, true);

				if (oNewContent.Elements[oNewContent.Elements.length - 1].Element.IsTable() && !oTable.IsInline())
					oNewContent.Elements[oNewContent.Elements.length - 1].Element.Internal_UpdateFlowPosition(oTable.PositionH.Value, oTable.PositionV.Value);

				let nParaPos = oParagraph.GetIndex();
				if (-1 !== nParaPos)
					oParent.RemoveFromContent(nParaPos, 1, true);
			}
		}

		this.UpdateSelection();
		this.Recalculate();
		this.FinalizeAction();
	}
};
CDocument.prototype.private_ConvertTableToText = function(oTable, oProps)
{
	if (oTable)
	{
		// если хоть одна строка выделена не до конца, то преобразуем всю таблицу (можно проверять только последнюю в выделении строку, а не все)
		// если выделено несколько строк, но до конца или от самого начала, то только эти строки
		var oSelectedRows = oTable.GetSelectedRowsRange();
		oSelectedRows.IsSelectionToEnd = oTable.IsSelectionToEnd();
		var oSelectionArr = oTable.GetSelectionArray();
		var isConvertAll = oTable.IsSelectedAll() || !oTable.Selection.Use || (!oSelectionArr[0].Row && !oSelectionArr[0].Cell && oSelectedRows.IsSelectionToEnd);
		if (!isConvertAll)
		{
			var oFirstCell;
			var	oLastCell;
			if (oSelectedRows.Start === oSelectedRows.End)
			{
				if (oTable.Selection.StartPos.Pos.Cell === oTable.Selection.EndPos.Pos.Cell && oTable.GetRow(oSelectedRows.Start).GetCellsCount() > 1)
				{
					isConvertAll = true;
				}
				else
				{
					oLastCell = Math.max(oTable.Selection.StartPos.Pos.Cell, oTable.Selection.EndPos.Pos.Cell);
					oFirstCell = Math.min(oTable.Selection.StartPos.Pos.Cell, oTable.Selection.EndPos.Pos.Cell);
				}
	
			}
			else
			{
				// надо посомотреть первую или последнюю строку, выделены ли они целиком или нет
				for (var i = 0; i < oSelectionArr.length; i++)
				{
					if (oSelectionArr[i].Row == oSelectedRows.End)
					{
						oFirstCell = oSelectionArr[i].Cell;
						break;
					}
				}
				oLastCell = oSelectionArr[oSelectionArr.length - 1].Cell;
			}
			isConvertAll = ( !oFirstCell && oLastCell === (oTable.GetRow(oSelectedRows.End).GetCellsCount() - 1) ) ? false : true;
		}

		var TableC = oTable.Copy();
		var NewContent = {before: null, content: [], after: null};
		if (isConvertAll)
		{
			oSelectedRows.Start = 0;
			oSelectedRows.End = TableC.GetRowsCount() - 1;
		}
		for (var i = oSelectedRows.Start; i <= oSelectedRows.End; i++)
		{
			var oRow = TableC.GetRow(i);
			var Tabs = oProps.type == 2 ? new CParaTabs() : null;
			var oNewParagraph = new Paragraph(this.DrawingDocument, this);
			NewContent.content.push(oNewParagraph);
			var bAdd = true;
			for (var k = 0; k < oRow.GetCellsCount(); k++)
			{
				var oCell = oRow.GetCell(k);
				var oCDocumentContent = oCell.GetContent();
				var isNewPar = false;
				for (var j = 0; j < oCDocumentContent.GetElementsCount(); j++)
				{
					var oElement = oCDocumentContent.GetElement(j);
					switch (oElement.GetType()) {
						case type_Paragraph:
							if (isNewPar)
							{
								oNewParagraph = new Paragraph(this.DrawingDocument, this);
								NewContent.content.push(oNewParagraph);
							}
							else
							{
								isNewPar = true;
							}
							oNewParagraph.Concat(oElement, true);
							if (oProps.type !== 1)
								oNewParagraph.SetParagraphAlign(1);
							break;
						case type_Table:
							var oNestedContent = (oProps.nested) ? this.private_ConvertTableToText(oElement, oProps) : {content:[oElement]};
							if (j == 0 && NewContent.content[NewContent.content.length-1].IsEmpty() && bAdd)
								NewContent.content.pop();

							NewContent.content = NewContent.content.concat(oNestedContent.content);
							isNewPar = true;
							break;
						default:
							NewContent.content.push(oElement);
							isNewPar = true;
							if (j == 0 && NewContent.content[NewContent.content.length-1].IsEmpty() && bAdd)
								NewContent.content.pop();
							break;
					}
					bAdd = false;
					
				}
				if (k !== oRow.GetCellsCount() - 1)
				{
					var oRun = new ParaRun(oNewParagraph, false);
					var oText;
					switch (oProps.type) {
						case 2:
							oText = new AscWord.CRunTab();
							var pos = (oCell.Metrics.X_cell_end >> 0) + 2;
							Tabs.Add(new CParaTab(tab_Left, pos, Asc.c_oAscTabLeader.None));
							break;
						case 1:
							oNewParagraph = new Paragraph(this.DrawingDocument, this);
							NewContent.content.push(oNewParagraph);
							break;
						default:
							oText = new AscWord.CRunText(oProps.separator);
							break;
					}
					if (oText)
					{
						oRun.Add(oText);
						oNewParagraph.Internal_Content_Add((oNewParagraph.Content.length - 1 > 0 ? oNewParagraph.Content.length - 1 : oNewParagraph.Content.length), oRun);
					}
				}
				else if (oProps.type == 2)
				{
					oNewParagraph.SetParagraphTabs(Tabs);
				}
			}
		}

		if (!isConvertAll)
		{
			for (var i = oSelectedRows.End; i >= oSelectedRows.Start; i--)
			{
				TableC.RemoveTableRow(i);
			}
			if (!oSelectedRows.IsSelectionToEnd && oSelectedRows.Start) {
				var oNewTable = TableC.Split();
				NewContent.after = oNewTable;
				NewContent.before = TableC;
			}
			if (oSelectedRows.IsSelectionToEnd)
			{
				NewContent.before = TableC;
			}
			else if (!oSelectedRows.Start)
			{
				NewContent.after = TableC;
			}
		}

		for (var i = 0; i < NewContent.content.length; i++)
		{
			if (NewContent.content[i].IsParagraph())
				NewContent.content[i].CorrectContent();
		}
	
	}
	return NewContent;
};
/**
 * Получаем класс, управляющий комментариями
 * @returns {?AscCommonWord.CComments}
 */
CDocument.prototype.GetCommentsManager = function()
{
	return this.Comments;
};
/**
 * Убираем все, что связано с формами
 * @param {boolean} [isUseHistory=false] использовать ли историю и создать точку для этого действия
 */
CDocument.prototype.DocxfToDocx = function(isUseHistory)
{
	if (isUseHistory)
		this.StartAction(AscDFH.historydescription_Document_Docxf_To_Docx);

	let arrForms = this.FormsManager.GetAllForms();
	for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
	{
		let oForm = arrForms[nIndex];

		var oShape, oParaDrawing;

		if (oForm.IsComplexForm())
			oForm.SetComplexFormPr(undefined);

		if (oForm.IsFixedForm()
			&& (oShape = oForm.GetFixedFormWrapperShape())
			&& (oParaDrawing = oShape.parent)
			&& oParaDrawing instanceof ParaDrawing)
		{
			var oParentParagraph = oParaDrawing.GetParagraph();
			if (oParentParagraph && oParentParagraph.Parent && oParentParagraph.Parent.Is_DrawingShape())
				oForm.ConvertFormToInline();
			else
				oParaDrawing.SetForm(false);
		}

		var oCheckBoxPr;
		if ((oCheckBoxPr = oForm.GetCheckBoxPr()) && oCheckBoxPr.GetGroupKey())
		{
			var oNewPr = oCheckBoxPr.Copy();
			oNewPr.SetGroupKey(undefined);
			oForm.SetCheckBoxPr(oNewPr);
		}

		if (oForm.GetPictureFormPr())
			oForm.SetPictureFormPr(undefined);

		if (oForm.GetTextFormPr())
			oForm.SetTextFormPr(undefined);

		oForm.SetFormPr(undefined);
	}

	if (isUseHistory)
	{
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateSelection();
		this.FinalizeAction();
	}
};
//----------------------------------------------------------------------------------------------------------------------
// SpellCheck
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.SetDefaultLanguage = function(nNewLang)
{
	let oStyles  = this.GetStyles();
	let nOldLang = oStyles.Default.TextPr.Lang.Val;

	if (nOldLang !== nNewLang)
	{
		this.History.Add(new CChangesDocumentDefaultLanguage(this, nOldLang, nNewLang));
		oStyles.Default.TextPr.Lang.Val = nNewLang;

		this.RestartSpellCheck();
		this.UpdateInterface();
	}
};
CDocument.prototype.GetDefaultLanguage = function()
{
	return this.GetStyles().Default.TextPr.Lang.Val;
};
CDocument.prototype.RestartSpellCheck = function()
{
	this.Spelling.Reset();

	// TODO: добавить обработку в автофигурах
	// TODO: Надо добавить обработку во всех контроллерах

	this.SectionsInfo.RestartSpellCheck();

	for (let nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].RestartSpellCheck();
	}
};
CDocument.prototype.StopSpellCheck = function()
{
	this.Spelling.Reset();
};
CDocument.prototype.ContinueSpellCheck = function()
{
	this.Spelling.ContinueSpellCheck();
};
CDocument.prototype.TurnOffSpellCheck = function()
{
	this.Spelling.TurnOff();
};
CDocument.prototype.TurnOnSpellCheck = function()
{
	this.Spelling.TurnOn();
};
CDocument.prototype.GetSpellCheckManager = function()
{
	return this.Spelling;
};
//----------------------------------------------------------------------------------------------------------------------
// Search
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.Search = function(oProps, bDraw)
{
	//let nStartTime = performance.now();

	if (this.SearchEngine.Compare(oProps))
		return this.SearchEngine;

	this.SearchEngine.Clear();
	this.SearchEngine.Set(oProps);

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].Search(this.SearchEngine, search_Common);
	}

	this.SectionsInfo.Search(this.SearchEngine);

	var arrFootnotes = this.GetFootnotesList(null, null);
	this.SearchEngine.SetFootnotes(arrFootnotes);
	for (var nIndex = 0, nCount = arrFootnotes.length; nIndex < nCount; ++nIndex)
	{
		arrFootnotes[nIndex].Search(this.SearchEngine, search_Footnote);
	}

	var arrEndnotes = this.GetEndnotesList(null, null);
	this.SearchEngine.SetEndnotes(arrEndnotes);
	for (var nIndex = 0, nCount = arrEndnotes.length; nIndex < nCount; ++nIndex)
	{
		arrEndnotes[nIndex].Search(this.SearchEngine, search_Endnote);
	}

	if (false !== bDraw)
		this.Redraw(-1, -1);

	//console.log("Search logic: " + ((performance.now() - nStartTime) / 1000) + " s");

	return this.SearchEngine;
};
CDocument.prototype.ClearSearch = function()
{
	let isPrevSearch = this.SearchEngine.Count > 0;

	this.SearchEngine.Clear();

	if (isPrevSearch)
	{
		this.Api.sync_SearchEndCallback();
		this.DrawingDocument.ClearCachePages();
		this.DrawingDocument.FirePaint();
	}
};
CDocument.prototype.SelectSearchElement = function(Id)
{
	this.RemoveSelection();
	this.SearchEngine.Select(Id, true);
	this.RecalculateCurPos();

	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
	this.Document_UpdateRulersState();
};
CDocument.prototype.ReplaceSearchElement = function(NewStr, bAll, Id, bInterfaceEvent)
{
	var bResult = false;

	var oState = this.SaveDocumentState();

	var arrReplaceId = [];
	if (bAll)
	{
		if (this.StartSelectionLockCheck())
		{
			for (var sElementId in this.SearchEngine.Elements)
			{
				this.SelectSearchElement(sElementId);
				if (!this.ProcessSelectionLockCheck(AscCommon.changestype_Paragraph_Content, true))
					arrReplaceId.push(sElementId);
			}

			if (this.EndSelectionLockCheck())
				arrReplaceId.length = 0;
		}
	}
	else
	{
		this.SelectSearchElement(Id);

		if (this.StartSelectionLockCheck())
		{
			this.ProcessSelectionLockCheck(AscCommon.changestype_Paragraph_Content);
			if (!this.EndSelectionLockCheck(true))
				arrReplaceId.push(Id);
		}
	}

	var nOverallCount  = this.SearchEngine.Count;
	var nReplacedCount = arrReplaceId.length;

	if (nReplacedCount > 0)
	{
		this.StartAction(bAll ? AscDFH.historydescription_Document_ReplaceAll : AscDFH.historydescription_Document_ReplaceSingle);

		for (var nIndex = 0; nIndex < nReplacedCount; ++nIndex)
		{
			this.SearchEngine.Replace(NewStr, arrReplaceId[nIndex], false);
		}

		// Если остались элементы поиска, тогда не очищаем текущий поиск
		if (this.SearchEngine.Count)
			this.SearchEngine.ClearOnRecalc = false;

		this.Recalculate(true);
		this.SearchEngine.ClearOnRecalc = true;
		this.FinalizeAction();
	}

	if (bAll && false !== bInterfaceEvent)
		this.Api.sync_ReplaceAllCallback(nReplacedCount, nOverallCount);

	if (nReplacedCount)
		bResult = true;

	this.LoadDocumentState(oState);

	this.Document_UpdateInterfaceState();
	this.Document_UpdateSelectionState();
	this.Document_UpdateRulersState();

	return bResult;
};
CDocument.prototype.GetSearchElementId = function(bNext)
{
	var Id = null;

	this.SearchEngine.SetDirection(bNext);

	this.DrawingObjects.resetDrawStateBeforeAction();
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		var ParaDrawing = this.DrawingObjects.getMajorParaDrawing();

		Id = ParaDrawing.GetSearchElementId(bNext, true);
		if (null != Id)
			return Id;

		this.DrawingObjects.resetSelection();
		ParaDrawing.GoTo_Text(true !== bNext, false);
	}

	if (docpostype_Content === this.CurPos.Type)
	{
		Id = this.private_GetSearchIdInMainDocument(true);

		if (null === Id)
			Id = this.private_GetSearchIdInFootnotes(false);

		if (null === Id)
			Id = this.private_GetSearchIdInEndnotes(false);

		if (null === Id)
			Id = this.private_GetSearchIdInHdrFtr(false);

		if (null === Id)
			Id = this.private_GetSearchIdInMainDocument(false);
	}
	else if (docpostype_HdrFtr === this.CurPos.Type)
	{
		Id = this.private_GetSearchIdInHdrFtr(true);

		if (null === Id)
			Id = this.private_GetSearchIdInMainDocument(false);

		if (null === Id)
			Id = this.private_GetSearchIdInFootnotes(false);

		if (null === Id)
			Id = this.private_GetSearchIdInEndnotes(false);

		if (null === Id)
			Id = this.private_GetSearchIdInHdrFtr(false);
	}
	else if (docpostype_Footnotes === this.CurPos.Type)
	{
		Id = this.private_GetSearchIdInFootnotes(true);

		if (null === Id)
			Id = this.private_GetSearchIdInEndnotes(false);

		if (null === Id)
			Id = this.private_GetSearchIdInHdrFtr(false);

		if (null === Id)
			Id = this.private_GetSearchIdInMainDocument(false);

		if (null === Id)
			Id = this.private_GetSearchIdInFootnotes(false);
	}
	else if (docpostype_Endnotes === this.CurPos.Type)
	{
		Id = this.private_GetSearchIdInEndnotes(true);

		if (null === Id)
			Id = this.private_GetSearchIdInHdrFtr(false);

		if (null === Id)
			Id = this.private_GetSearchIdInMainDocument(false);

		if (null === Id)
			Id = this.private_GetSearchIdInFootnotes(false);

		if (null === Id)
			Id = this.private_GetSearchIdInEndnotes(false);
	}

	return Id;
};
CDocument.prototype.HighlightSearchResults = function(bSelection)
{
	var OldValue = this.SearchEngine.Selection;
	if ( OldValue === bSelection )
		return;

	this.SearchEngine.Selection = bSelection;
	this.DrawingDocument.ClearCachePages();
	this.DrawingDocument.FirePaint();
};
CDocument.prototype.IsHighlightSearchResults = function()
{
	return this.SearchEngine.Selection;
};
CDocument.prototype.private_GetSearchIdInMainDocument = function(isCurrent)
{
	var Id    = null;
	var bNext = this.SearchEngine.GetDirection();
	var Pos   = this.CurPos.ContentPos;
	if (true === this.Selection.Use && selectionflag_Common === this.Selection.Flag)
		Pos = bNext ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : Math.min(this.Selection.StartPos, this.Selection.EndPos);

	if (true !== isCurrent)
		Pos = bNext ? 0 : this.Content.length - 1;

	if (true === bNext)
	{
		Id = this.Content[Pos].GetSearchElementId(true, isCurrent);

		if (null != Id)
			return Id;

		Pos++;

		var Count = this.Content.length;
		while (Pos < Count)
		{
			Id = this.Content[Pos].GetSearchElementId(true, false);

			if (null != Id)
				return Id;

			Pos++;
		}
	}
	else
	{
		Id = this.Content[Pos].GetSearchElementId(false, isCurrent);

		if (null != Id)
			return Id;

		Pos--;

		while (Pos >= 0)
		{
			Id = this.Content[Pos].GetSearchElementId(false, false);

			if (null != Id)
				return Id;

			Pos--;
		}
	}

	return Id;
};
CDocument.prototype.private_GetSearchIdInHdrFtr = function(isCurrent)
{
	return this.SectionsInfo.GetSearchElementId(this.SearchEngine.GetDirection(), isCurrent ? this.HdrFtr.CurHdrFtr : null);
};
CDocument.prototype.private_GetSearchIdInFootnotes = function(isCurrent)
{
	var bNext        = this.SearchEngine.GetDirection();
	var oCurFootnote = this.Footnotes.CurFootnote;

	var arrFootnotes = this.SearchEngine.GetFootnotes();
	var nCurPos      = -1;
	var nCount       = arrFootnotes.length;

	if (nCount <= 0)
		return null;

	if (isCurrent)
	{
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			if (arrFootnotes[nIndex] === oCurFootnote)
			{
				nCurPos = nIndex;
				break;
			}
		}
	}

	if (-1 === nCurPos)
	{
		nCurPos      = bNext ? 0 : nCount - 1;
		oCurFootnote = arrFootnotes[nCurPos];
		isCurrent    = false;
	}

	var Id = oCurFootnote.GetSearchElementId(bNext, isCurrent);
	if (null !== Id)
		return Id;

	if (true === bNext)
	{
		for (var nIndex = nCurPos + 1; nIndex < nCount; ++nIndex)
		{
			Id = arrFootnotes[nIndex].GetSearchElementId(bNext, false);
			if (null != Id)
				return Id;
		}
	}
	else
	{
		for (var nIndex = nCurPos - 1; nIndex >= 0; --nIndex)
		{
			Id = arrFootnotes[nIndex].GetSearchElementId(bNext, false);
			if (null != Id)
				return Id;
		}
	}

	return null;
};
CDocument.prototype.private_GetSearchIdInEndnotes = function(isCurrent)
{
	var bNext       = this.SearchEngine.GetDirection();
	var oCurEndnote = this.Endnotes.CurEndnote;

	var arrEndnotes = this.SearchEngine.GetEndnotes();
	var nCurPos     = -1;
	var nCount      = arrEndnotes.length;

	if (nCount <= 0)
		return null;

	if (isCurrent)
	{
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			if (arrEndnotes[nIndex] === oCurEndnote)
			{
				nCurPos = nIndex;
				break;
			}
		}
	}

	if (-1 === nCurPos)
	{
		nCurPos     = bNext ? 0 : nCount - 1;
		oCurEndnote = arrEndnotes[nCurPos];
		isCurrent   = false;
	}

	var Id = oCurEndnote.GetSearchElementId(bNext, isCurrent);
	if (null !== Id)
		return Id;

	if (true === bNext)
	{
		for (var nIndex = nCurPos + 1; nIndex < nCount; ++nIndex)
		{
			Id = arrEndnotes[nIndex].GetSearchElementId(bNext, false);
			if (null != Id)
				return Id;
		}
	}
	else
	{
		for (var nIndex = nCurPos - 1; nIndex >= 0; --nIndex)
		{
			Id = arrEndnotes[nIndex].GetSearchElementId(bNext, false);
			if (null != Id)
				return Id;
		}
	}

	return null;
};
//----------------------------------------------------------------------------------------------------------------------
CDocument.prototype.SetupBeforeNativePrint = function(layoutOptions, oGraphics)
{
	if (!layoutOptions || !oGraphics)
		return;

	if (!this.StoredOptions)
		this.StoredOptions = {};

	this.StoredOptions.DrawFormHighlight = this.ForceDrawFormHighlight;
	this.StoredOptions.DrawPlaceHolders  = this.ForceDrawPlaceHolders;
	this.StoredOptions.Graphics          = oGraphics;
	this.StoredOptions.GraphicsPrintMode = oGraphics.isPrintMode;

	if (true === layoutOptions["drawPlaceHolders"] || false === layoutOptions["drawPlaceHolders"])
		this.ForceDrawPlaceHolders = layoutOptions["drawPlaceHolders"];

	if (true === layoutOptions["drawFormHighlight"] || false === layoutOptions["drawFormHighlight"])
		this.ForceDrawFormHighlight = layoutOptions["drawFormHighlight"];

	if (true === layoutOptions["isPrint"] || false === layoutOptions["isPrint"])
		oGraphics.isPrintMode = layoutOptions["isPrint"];
};
CDocument.prototype.RestoreAfterNativePrint = function()
{
	if (!this.StoredOptions)
		return;

	if (undefined !== this.StoredOptions.DrawFormHighlight)
	{
		this.ForceDrawFormHighlight = this.StoredOptions.DrawFormHighlight;
		delete this.StoredOptions.DrawFormHighlight;
	}

	if (undefined !== this.StoredOptions.DrawPlaceHolders)
	{
		this.ForceDrawPlaceHolders = this.StoredOptions.DrawPlaceHolders;
		delete this.StoredOptions.DrawPlaceHolders;
	}

	if (this.StoredOptions.Graphics)
	{
		if (undefined !== this.StoredOptions.GraphicsPrintMode)
		{
			this.StoredOptions.Graphics.isPrintMode = this.StoredOptions.GraphicsPrintMode;
			delete this.StoredOptions.GraphicsPrintMode;
		}

		delete this.StoredOptions.Graphics;
	}
};
CDocument.prototype.GetImageFormSelection = function()
{
    return this.DrawingObjects.getImageDataFromSelection();
};
/**
 * Проходим по всем ранам документа (в отличие от функции CheckRunContent, проходящей только по основной части
 * документа)
 */
CDocument.prototype.CheckAllRunContent = function(fCheck)
{
	function private_check(run)
	{
		for (let pos = 0, count = run.GetElementsCount(); pos < count; ++pos)
		{
			let item = run.GetElement(pos);
			if (item.IsDrawing())
			{
				if (item.CheckRunContent(private_check))
					return true;
			}
		}

		return fCheck(run);
	}

	this.CheckRunContent(private_check);
	this.SectionsInfo.CheckRunContent(private_check);
	this.Footnotes.CheckRunContent(private_check);
	this.Endnotes.CheckRunContent(private_check);
};

/**
 * @param {Asc.c_oAscMathInputType} nType
 */
CDocument.prototype.SetMathInputType = function(nType)
{
	this.MathInputType = nType;
};
/**
 * @returns {Asc.c_oAscMathInputType}
 */
CDocument.prototype.GetMathInputType = function()
{
	return this.MathInputType;
};

/**
 * Функция конвертации вида формулы из линейного в профессиональный и наоборот
 * @param {boolean} isToLinear
 */
CDocument.prototype.ConvertMathView = function(isToLinear)
{
	var oInfo = this.GetSelectedElementsInfo();
	var oMath = oInfo.GetMath();
	if (!oMath)
		return;
	
	if (!this.IsSelectionLocked(AscCommon.changestype_Paragraph_Content))
	{
		this.StartAction(AscDFH.historydescription_Document_ConvertMathView);
		let nInputType = this.Api.getMathInputType();
		
		if (!this.IsTextSelectionUse())
		{
			this.RemoveTextSelection();
			oMath.ConvertView(isToLinear, nInputType);
		}
		else
		{
			oMath.ConvertViewBySelection(isToLinear, nInputType);
		}
		
		this.Recalculate();
		this.UpdateInterface();
		this.UpdateTracks();
		this.FinalizeAction();
	}
};
/**
 * Функция конвертации вида всех формул из линейного в профессиональный и наоборот
 * @param {boolean} isToLinear
 */
CDocument.prototype.ConvertAllMathView = function(isToLinear)
{
	let allMaths          = [];
	let allParagraphs     = this.GetAllParagraphs();
	let paragraphsToCheck = [];
	
	for (let paraIndex = 0, parasCount = allParagraphs.length; paraIndex < parasCount; ++paraIndex)
	{
		let paragraph  = allParagraphs[paraIndex];
		let mathsCount = allMaths.length;
		paragraph.GetAllParaMaths(allMaths);
		
		if (mathsCount !== allMaths.length)
			paragraphsToCheck.push(paragraph);
	}
	
	if (!allMaths.length || !paragraphsToCheck.length)
		return;
	
	if (this.IsSelectionLocked(AscCommon.changestype_None, {
		Type      : AscCommon.changestype_2_ElementsArray_and_Type,
		Elements  : paragraphsToCheck,
		CheckType : AscCommon.changestype_Paragraph_Content
	}))
		return;
	
	this.StartAction(AscDFH.historydescription_Document_ConvertMathView);
	let inputType = this.Api.getMathInputType();
	
	for (let mathIndex = 0, mathsCount = allMaths.length; mathIndex < mathsCount; ++mathIndex)
	{
		allMaths[mathIndex].ConvertView(isToLinear, inputType);
	}
	
	this.Recalculate();
	this.UpdateInterface();
	this.UpdateTracks();
	this.FinalizeAction();
};
CDocument.prototype.IsCheckFormPlaceholder = function()
{
	if (!this.IsFillingFormMode())
		return false;
	
	return this.CheckFormPlaceHolder;
};

function CDocumentSelectionState()
{
    this.Id        = null;
    this.Type      = docpostype_Content;
    this.Data      = {}; // Объект с текущей позицией
}

function CDocumentSectionsInfo()
{
    this.Elements = [];
}

CDocumentSectionsInfo.prototype =
{
    Add : function( SectPr, Index )
    {
        this.Elements.push( new CDocumentSectionsInfoElement( SectPr, Index ) );
    },

	GetSectionsCount : function()
    {
        return this.Elements.length;
    },

    Clear : function()
    {
        this.Elements.length = 0;
    },

    Find_ByHdrFtr : function(HdrFtr)
    {
		if (!HdrFtr)
			return -1;
		
        var Count = this.Elements.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            var SectPr = this.Elements[Index].SectPr;

            if ( HdrFtr === SectPr.Get_Header_First() || HdrFtr === SectPr.Get_Header_Default() || HdrFtr === SectPr.Get_Header_Even() ||
                 HdrFtr === SectPr.Get_Footer_First() || HdrFtr === SectPr.Get_Footer_Default() || HdrFtr === SectPr.Get_Footer_Even() )
                    return Index;
        }

        return -1;
    },

    Reset_HdrFtrRecalculateCache : function()
    {
        var Count = this.Elements.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            var SectPr = this.Elements[Index].SectPr;

            if ( null != SectPr.HeaderFirst )
                SectPr.HeaderFirst.Reset_RecalculateCache();

            if ( null != SectPr.HeaderDefault )
                SectPr.HeaderDefault.Reset_RecalculateCache();

            if ( null != SectPr.HeaderEven )
                SectPr.HeaderEven.Reset_RecalculateCache();

            if ( null != SectPr.FooterFirst )
                SectPr.FooterFirst.Reset_RecalculateCache();

            if ( null != SectPr.FooterDefault )
                SectPr.FooterDefault.Reset_RecalculateCache();

            if ( null != SectPr.FooterEven )
                SectPr.FooterEven.Reset_RecalculateCache();
        }
    },

    GetAllParagraphs : function(Props, ParaArray)
    {
        var Count = this.Elements.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            var SectPr = this.Elements[Index].SectPr;

            if ( null != SectPr.HeaderFirst )
                SectPr.HeaderFirst.GetAllParagraphs(Props, ParaArray);

            if ( null != SectPr.HeaderDefault )
                SectPr.HeaderDefault.GetAllParagraphs(Props, ParaArray);

            if ( null != SectPr.HeaderEven )
                SectPr.HeaderEven.GetAllParagraphs(Props, ParaArray);

            if ( null != SectPr.FooterFirst )
                SectPr.FooterFirst.GetAllParagraphs(Props, ParaArray);

            if ( null != SectPr.FooterDefault )
                SectPr.FooterDefault.GetAllParagraphs(Props, ParaArray);

            if ( null != SectPr.FooterEven )
                SectPr.FooterEven.GetAllParagraphs(Props, ParaArray);
        }
    },

	GetAllTables : function(oProps, arrTables)
	{
		var Count = this.Elements.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var SectPr = this.Elements[Index].SectPr;

			if (null != SectPr.HeaderFirst)
				SectPr.HeaderFirst.GetAllTables(oProps, arrTables);

			if (null != SectPr.HeaderDefault)
				SectPr.HeaderDefault.GetAllTables(oProps, arrTables);

			if (null != SectPr.HeaderEven)
				SectPr.HeaderEven.GetAllTables(oProps, arrTables);

			if (null != SectPr.FooterFirst)
				SectPr.FooterFirst.GetAllTables(oProps, arrTables);

			if (null != SectPr.FooterDefault)
				SectPr.FooterDefault.GetAllTables(oProps, arrTables);

			if (null != SectPr.FooterEven)
				SectPr.FooterEven.GetAllTables(oProps, arrTables);
		}
	},

	GetAllDrawingObjects : function(arrDrawings)
    {
        for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
        {
            var SectPr = this.Elements[nIndex].SectPr;

            if (null != SectPr.HeaderFirst)
                SectPr.HeaderFirst.GetAllDrawingObjects(arrDrawings);

            if (null != SectPr.HeaderDefault)
                SectPr.HeaderDefault.GetAllDrawingObjects(arrDrawings);

            if (null != SectPr.HeaderEven)
                SectPr.HeaderEven.GetAllDrawingObjects(arrDrawings);

            if (null != SectPr.FooterFirst)
                SectPr.FooterFirst.GetAllDrawingObjects(arrDrawings);

            if (null != SectPr.FooterDefault)
                SectPr.FooterDefault.GetAllDrawingObjects(arrDrawings);

            if (null != SectPr.FooterEven)
                SectPr.FooterEven.GetAllDrawingObjects(arrDrawings);
        }
	},
	
	UpdateBookmarks : function(oBookmarkManager)
	{
        for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
        {
            var SectPr = this.Elements[nIndex].SectPr;

            if (null != SectPr.HeaderFirst)
                SectPr.HeaderFirst.UpdateBookmarks(oBookmarkManager);

            if (null != SectPr.HeaderDefault)
                SectPr.HeaderDefault.UpdateBookmarks(oBookmarkManager);

            if (null != SectPr.HeaderEven)
                SectPr.HeaderEven.UpdateBookmarks(oBookmarkManager);

            if (null != SectPr.FooterFirst)
                SectPr.FooterFirst.UpdateBookmarks(oBookmarkManager);

            if (null != SectPr.FooterDefault)
                SectPr.FooterDefault.UpdateBookmarks(oBookmarkManager);

            if (null != SectPr.FooterEven)
                SectPr.FooterEven.UpdateBookmarks(oBookmarkManager);
        }
	},

    Document_CreateFontMap : function(FontMap)
    {
        var Count = this.Elements.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            var SectPr = this.Elements[Index].SectPr;

            if ( null != SectPr.HeaderFirst )
                SectPr.HeaderFirst.Document_CreateFontMap(FontMap);

            if ( null != SectPr.HeaderDefault )
                SectPr.HeaderDefault.Document_CreateFontMap(FontMap);

            if ( null != SectPr.HeaderEven )
                SectPr.HeaderEven.Document_CreateFontMap(FontMap);

            if ( null != SectPr.FooterFirst )
                SectPr.FooterFirst.Document_CreateFontMap(FontMap);

            if ( null != SectPr.FooterDefault )
                SectPr.FooterDefault.Document_CreateFontMap(FontMap);

            if ( null != SectPr.FooterEven )
                SectPr.FooterEven.Document_CreateFontMap(FontMap);
        }
    },

    Document_CreateFontCharMap : function(FontCharMap)
    {
        var Count = this.Elements.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            var SectPr = this.Elements[Index].SectPr;

            if ( null != SectPr.HeaderFirst )
                SectPr.HeaderFirst.Document_CreateFontCharMap(FontCharMap);

            if ( null != SectPr.HeaderDefault )
                SectPr.HeaderDefault.Document_CreateFontCharMap(FontCharMap);

            if ( null != SectPr.HeaderEven )
                SectPr.HeaderEven.Document_CreateFontCharMap(FontCharMap);

            if ( null != SectPr.FooterFirst )
                SectPr.FooterFirst.Document_CreateFontCharMap(FontCharMap);

            if ( null != SectPr.FooterDefault )
                SectPr.FooterDefault.Document_CreateFontCharMap(FontCharMap);

            if ( null != SectPr.FooterEven )
                SectPr.FooterEven.Document_CreateFontCharMap(FontCharMap);
        }
    },

    Document_Get_AllFontNames : function ( AllFonts )
    {
        var Count = this.Elements.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            var SectPr = this.Elements[Index].SectPr;

            if ( null != SectPr.HeaderFirst )
                SectPr.HeaderFirst.Document_Get_AllFontNames(AllFonts);

            if ( null != SectPr.HeaderDefault )
                SectPr.HeaderDefault.Document_Get_AllFontNames(AllFonts);

            if ( null != SectPr.HeaderEven )
                SectPr.HeaderEven.Document_Get_AllFontNames(AllFonts);

            if ( null != SectPr.FooterFirst )
                SectPr.FooterFirst.Document_Get_AllFontNames(AllFonts);

            if ( null != SectPr.FooterDefault )
                SectPr.FooterDefault.Document_Get_AllFontNames(AllFonts);

            if ( null != SectPr.FooterEven )
                SectPr.FooterEven.Document_Get_AllFontNames(AllFonts);
        }
    },

    Get_Index : function(Index)
    {
        var Count = this.Elements.length;

        for ( var Pos = 0; Pos < Count; Pos++ )
        {
            if ( Index <= this.Elements[Pos].Index )
                return Pos;
        }

        // Последний элемент здесь это всегда конечная секция документа
        return (Count - 1);
    },

    Get_Count : function()
    {
        return this.Elements.length;
    },

    Get_SectPr : function(Index)
    {
    	return this.GetByContentPos(Index);
    },

    Get_SectPr2 : function(Index)
    {
        return this.Elements[Index];
    },

    Find : function(SectPr)
    {
        var Count = this.Elements.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            var Element = this.Elements[Index];
            if ( Element.SectPr === SectPr )
                return Index;
        }

        return -1;
    },

    Update_OnAdd : function(Pos, Items)
    {
        var Count = Items.length;
        var Len = this.Elements.length;

        // Сначала обновим старые метки
        for (var Index = 0; Index < Len; Index++)
        {
            if ( this.Elements[Index].Index >= Pos )
                this.Elements[Index].Index += Count;
        }

        // Если среди новых элементов были параграфы с настройками секции, тогда добавим их здесь
        for (var Index = 0; Index < Count; Index++ )
        {
            var Item = Items[Index];
            var SectPr = ( type_Paragraph === Item.GetType() ? Item.Get_SectionPr() : undefined );

            if ( undefined !== SectPr )
            {
                var TempPos = 0;
                for ( ; TempPos < Len; TempPos++ )
                {
                    if ( Pos + Index <= this.Elements[TempPos].Index )
                        break;
                }

                this.Elements.splice( TempPos, 0, new CDocumentSectionsInfoElement( SectPr, Pos + Index ) );
                Len++;
            }
        }
    },

    Update_OnRemove : function(Pos, Count, bCheckHdrFtr)
    {
        var Len = this.Elements.length;

        for (var Index = 0; Index < Len; Index++)
        {
            var CurPos = this.Elements[Index].Index;

            if (CurPos >= Pos && CurPos < Pos + Count)
            {
                // Копируем поведение Word: Если у следующей секции не задан вообще ни один колонтитул,
                // тогда копируем ссылки на колонтитулы из удаляемой секции. Если задан хоть один колонтитул,
                // тогда этого не делаем.
                if (true === bCheckHdrFtr && Index < Len - 1)
                {
                    var CurrSectPr = this.Elements[Index].SectPr;
                    var NextSectPr = this.Elements[Index + 1].SectPr;
                    if (true === NextSectPr.IsAllHdrFtrNull() && true !== CurrSectPr.IsAllHdrFtrNull())
                    {
                        NextSectPr.Set_Header_First(CurrSectPr.Get_Header_First());
                        NextSectPr.Set_Header_Even(CurrSectPr.Get_Header_Even());
                        NextSectPr.Set_Header_Default(CurrSectPr.Get_Header_Default());
                        NextSectPr.Set_Footer_First(CurrSectPr.Get_Footer_First());
                        NextSectPr.Set_Footer_Even(CurrSectPr.Get_Footer_Even());
                        NextSectPr.Set_Footer_Default(CurrSectPr.Get_Footer_Default());
                    }
                }

                this.Elements.splice(Index, 1);
                Len--;
                Index--;


            }
            else if (CurPos >= Pos + Count)
                this.Elements[Index].Index -= Count;
        }
    }
};
CDocumentSectionsInfo.prototype.GetCount = function()
{
	return this.Elements.length;
};
/**
 * Получаем секцию по заданному номеру
 * @param {number} nIndex
 * @returns {CDocumentSectionsInfoElement}
 */
CDocumentSectionsInfo.prototype.Get = function(nIndex)
{
	return this.Elements[nIndex];
};
/**
 * Получаем секцию по заданной позиции контента
 * @param {number} nContentPos
 * @returns {CDocumentSectionsInfoElement}
 */
CDocumentSectionsInfo.prototype.GetByContentPos = function(nContentPos)
{
	var nCount = this.Elements.length;
	for (var nPos = 0; nPos < nCount; ++nPos)
	{
		if (nContentPos <= this.Elements[nPos].Index)
			return this.Elements[nPos];
	}

	// Последний элемент здесь это всегда конечная секция документа
	return this.Elements[nCount - 1];
};
/**
 * Получаем массив всех колонтитулов, используемых в данном документе
 * @returns {Array.CHeaderFooter}
 */
CDocumentSectionsInfo.prototype.GetAllHdrFtrs = function()
{
	var HdrFtrs = [];

	var Count = this.Elements.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var SectPr = this.Elements[Index].SectPr;
		SectPr.GetAllHdrFtrs(HdrFtrs);
	}

	return HdrFtrs;
};
CDocumentSectionsInfo.prototype.GetAllContentControls = function(arrContentControls)
{
	for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
	{
		var SectPr = this.Elements[nIndex].SectPr;

		if (null != SectPr.HeaderFirst)
			SectPr.HeaderFirst.GetAllContentControls(arrContentControls);

		if (null != SectPr.HeaderDefault)
			SectPr.HeaderDefault.GetAllContentControls(arrContentControls);

		if (null != SectPr.HeaderEven)
			SectPr.HeaderEven.GetAllContentControls(arrContentControls);

		if (null != SectPr.FooterFirst)
			SectPr.FooterFirst.GetAllContentControls(arrContentControls);

		if (null != SectPr.FooterDefault)
			SectPr.FooterDefault.GetAllContentControls(arrContentControls);

		if (null != SectPr.FooterEven)
			SectPr.FooterEven.GetAllContentControls(arrContentControls);
	}
};
/**
 * Обновляем заданную секцию
 * @param oSectPr {CSectionPr} - Секция, которую нужно обновить
 * @param oNewSectPr {?CSectionPr} - Либо новое значение секции, либо undefined для удалении секции
 * @param isCheckHdrFtr {boolean} - Нужно ли проверять колонтитулы при удалении секции
 * @returns {boolean} Если не смогли обновить, возвращаем false
 */
CDocumentSectionsInfo.prototype.UpdateSection = function(oSectPr, oNewSectPr, isCheckHdrFtr)
{
	if (oSectPr === oNewSectPr || !oSectPr)
		return false;

	for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
	{
		if (oSectPr === this.Elements[nIndex].SectPr)
		{
			if (!oNewSectPr)
			{
				// Копируем поведение Word: Если у следующей секции не задан вообще ни один колонтитул,
				// тогда копируем ссылки на колонтитулы из удаляемой секции. Если задан хоть один колонтитул,
				// тогда этого не делаем.
				if (true === isCheckHdrFtr && nIndex < nCount - 1)
				{
					var oCurrSectPr = this.Elements[nIndex].SectPr;
					var oNextSectPr = this.Elements[nIndex + 1].SectPr;

					if (true === oNextSectPr.IsAllHdrFtrNull() && true !== oCurrSectPr.IsAllHdrFtrNull())
					{
						oNextSectPr.Set_Header_First(oCurrSectPr.Get_Header_First());
						oNextSectPr.Set_Header_Even(oCurrSectPr.Get_Header_Even());
						oNextSectPr.Set_Header_Default(oCurrSectPr.Get_Header_Default());
						oNextSectPr.Set_Footer_First(oCurrSectPr.Get_Footer_First());
						oNextSectPr.Set_Footer_Even(oCurrSectPr.Get_Footer_Even());
						oNextSectPr.Set_Footer_Default(oCurrSectPr.Get_Footer_Default());
					}
				}

				this.Elements.splice(nIndex, 1);
			}
			else
			{
				this.Elements[nIndex].SectPr = oNewSectPr;
			}

			return true;
		}
	}

	return false;
};
CDocumentSectionsInfo.prototype.private_GetHdrFtrsArray = function(oCurHdrFtr)
{
	var isEvenOdd = EvenAndOddHeaders;

	var nCurPos    = -1;
	var arrHdrFtrs = [];
	for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
	{
		var oSectPr = this.Elements[nIndex].SectPr;
		var isFirst = oSectPr.Get_TitlePage();

		var oHeaderFirst   = oSectPr.Get_Header_First();
		var oHeaderEven    = oSectPr.Get_Header_Even();
		var oHeaderDefault = oSectPr.Get_Header_Default();
		var oFooterFirst   = oSectPr.Get_Footer_First();
		var oFooterEven    = oSectPr.Get_Footer_Even();
		var oFooterDefault = oSectPr.Get_Footer_Default();

		if (oHeaderFirst && isFirst)
			arrHdrFtrs.push(oHeaderFirst);

		if (oHeaderEven && isEvenOdd)
			arrHdrFtrs.push(oHeaderEven);

		if (oHeaderDefault)
			arrHdrFtrs.push(oHeaderDefault);

		if (oFooterFirst && isFirst)
			arrHdrFtrs.push(oFooterFirst);

		if (oFooterEven && isEvenOdd)
			arrHdrFtrs.push(oFooterEven);

		if (oFooterDefault)
			arrHdrFtrs.push(oFooterDefault);
	}

	if (oCurHdrFtr)
	{
		for (var nIndex = 0, nCount = arrHdrFtrs.length; nIndex < nCount; ++nIndex)
		{
			if (oCurHdrFtr === arrHdrFtrs[nIndex])
			{
				nCurPos = nIndex;
				break;
			}
		}
	}

	return {
		HdrFtrs : arrHdrFtrs,
		CurPos  : nCurPos
	};
};
CDocumentSectionsInfo.prototype.FindNextFillingForm = function(isNext, oCurHdrFtr)
{
	var oInfo = this.private_GetHdrFtrsArray(oCurHdrFtr);

	var arrHdrFtrs = oInfo.HdrFtrs;
	var nCurPos    = oInfo.CurPos;

	var nCount = arrHdrFtrs.length;

	var isCurrent = true;
	if (-1 === nCurPos)
	{
		isCurrent = false;
		nCurPos   = isNext ? 0 : arrHdrFtrs.length - 1;
		if (arrHdrFtrs[nCurPos])
			oCurHdrFtr = arrHdrFtrs[nCurPos];
	}

	if (nCurPos >= 0 && nCurPos <= nCount - 1)
	{
		var oRes = oCurHdrFtr.GetContent().FindNextFillingForm(isNext, isCurrent, isCurrent);
		if (oRes)
			return oRes;

		if (isNext)
		{
			for (var nIndex = nCurPos + 1; nIndex < nCount; ++nIndex)
			{
				oRes = arrHdrFtrs[nIndex].GetContent().FindNextFillingForm(isNext, false);

				if (oRes)
					return oRes;
			}
		}
		else
		{
			for (var nIndex = nCurPos - 1; nIndex >= 0; --nIndex)
			{
				oRes = arrHdrFtrs[nIndex].GetContent().FindNextFillingForm(isNext, false);

				if (oRes)
					return oRes;
			}
		}
	}

	return null;
};
CDocumentSectionsInfo.prototype.RestartSpellCheck = function()
{
	var bEvenOdd = EvenAndOddHeaders;
	for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
	{
		var SectPr = this.Elements[nIndex].SectPr;
		var bFirst = SectPr.Get_TitlePage();

		if (null != SectPr.HeaderFirst && true === bFirst)
			SectPr.HeaderFirst.RestartSpellCheck();

		if (null != SectPr.HeaderEven && true === bEvenOdd)
			SectPr.HeaderEven.RestartSpellCheck();

		if (null != SectPr.HeaderDefault)
			SectPr.HeaderDefault.RestartSpellCheck();

		if (null != SectPr.FooterFirst && true === bFirst)
			SectPr.FooterFirst.RestartSpellCheck();

		if (null != SectPr.FooterEven && true === bEvenOdd)
			SectPr.FooterEven.RestartSpellCheck();

		if (null != SectPr.FooterDefault)
			SectPr.FooterDefault.RestartSpellCheck();
	}
};
CDocumentSectionsInfo.prototype.RemoveEmptyHdrFtrs = function()
{
	for (let nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
	{
		let oSectPr = this.Elements[nIndex].SectPr;
		oSectPr.RemoveEmptyHdrFtrs();
	}
};
CDocumentSectionsInfo.prototype.CheckRunContent = function(fCheck)
{
	let headers = this.GetAllHdrFtrs();
	for (let index = 0, count = headers.length; index < count; ++index)
	{
		headers[index].GetContent().CheckRunContent(fCheck);
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Search
//----------------------------------------------------------------------------------------------------------------------
CDocumentSectionsInfo.prototype.Search = function(oSearchEngine)
{
	var bEvenOdd = EvenAndOddHeaders;
	for (var nIndex = 0, nCount = this.Elements.length; nIndex < nCount; ++nIndex)
	{
		var oSectPr = this.Elements[nIndex].SectPr;
		var bFirst  = oSectPr.Get_TitlePage();

		if (oSectPr.HeaderFirst && true === bFirst)
			oSectPr.HeaderFirst.Search(oSearchEngine, search_Header);

		if (oSectPr.HeaderEven && true === bEvenOdd)
			oSectPr.HeaderEven.Search(oSearchEngine, search_Header);

		if (oSectPr.HeaderDefault)
			oSectPr.HeaderDefault.Search(oSearchEngine, search_Header);

		if (oSectPr.FooterFirst && true === bFirst)
			oSectPr.FooterFirst.Search(oSearchEngine, search_Footer);

		if (oSectPr.FooterEven && true === bEvenOdd)
			oSectPr.FooterEven.Search(oSearchEngine, search_Footer);

		if (oSectPr.FooterDefault)
			oSectPr.FooterDefault.Search(oSearchEngine, search_Footer);
	}
};
CDocumentSectionsInfo.prototype.GetSearchElementId = function(bNext, CurHdrFtr)
{
	var HdrFtrs = [];
	var CurPos  = -1;

	var bEvenOdd = EvenAndOddHeaders;
	var Count    = this.Elements.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var SectPr = this.Elements[Index].SectPr;
		var bFirst = SectPr.Get_TitlePage();

		if (null != SectPr.HeaderFirst && true === bFirst)
		{
			HdrFtrs.push(SectPr.HeaderFirst);

			if (CurHdrFtr === SectPr.HeaderFirst)
				CurPos = HdrFtrs.length - 1;
		}

		if (null != SectPr.HeaderEven && true === bEvenOdd)
		{
			HdrFtrs.push(SectPr.HeaderEven);

			if (CurHdrFtr === SectPr.HeaderEven)
				CurPos = HdrFtrs.length - 1;
		}

		if (null != SectPr.HeaderDefault)
		{
			HdrFtrs.push(SectPr.HeaderDefault);

			if (CurHdrFtr === SectPr.HeaderDefault)
				CurPos = HdrFtrs.length - 1;
		}

		if (null != SectPr.FooterFirst && true === bFirst)
		{
			HdrFtrs.push(SectPr.FooterFirst);

			if (CurHdrFtr === SectPr.FooterFirst)
				CurPos = HdrFtrs.length - 1;
		}

		if (null != SectPr.FooterEven && true === bEvenOdd)
		{
			HdrFtrs.push(SectPr.FooterEven);

			if (CurHdrFtr === SectPr.FooterEven)
				CurPos = HdrFtrs.length - 1;
		}

		if (null != SectPr.FooterDefault)
		{
			HdrFtrs.push(SectPr.FooterDefault);

			if (CurHdrFtr === SectPr.FooterDefault)
				CurPos = HdrFtrs.length - 1;
		}
	}

	var Count = HdrFtrs.length;

	var isCurrent = true;
	if (-1 === CurPos)
	{
		isCurrent = false;
		CurPos    = bNext ? 0 : HdrFtrs.length - 1;
		if (HdrFtrs[CurPos])
			CurHdrFtr = HdrFtrs[CurPos];
	}

	if (CurPos >= 0 && CurPos <= HdrFtrs.length - 1)
	{
		var Id = CurHdrFtr.GetSearchElementId(bNext, isCurrent);
		if (null != Id)
			return Id;

		if (true === bNext)
		{
			for (var Index = CurPos + 1; Index < Count; Index++)
			{
				Id = HdrFtrs[Index].GetSearchElementId(bNext, false);

				if (null != Id)
					return Id;
			}
		}
		else
		{
			for (var Index = CurPos - 1; Index >= 0; Index--)
			{
				Id = HdrFtrs[Index].GetSearchElementId(bNext, false);

				if (null != Id)
					return Id;
			}
		}
	}

	return null;
};
//----------------------------------------------------------------------------------------------------------------------

function CDocumentSectionsInfoElement(SectPr, Index)
{
    this.SectPr = SectPr;
    this.Index  = Index;
}

function CDocumentCompareDrawingsLogicPositions(Drawing1, Drawing2)
{
    this.Drawing1 = Drawing1;
    this.Drawing2 = Drawing2;
    this.Result   = 0;
}


/**
 * Класс для обработки (принятия/отклонения) изменения связанного с переносом
 * @param sMoveId
 * @param sUserId
 * @constructor
 */
function CTrackRevisionsMoveProcessEngine(sMoveId, sUserId)
{
	this.MoveId        = sMoveId; // Идентификатор обрабаываемого переноса
	this.UserId        = sUserId; // Идентификатор пользователя, сделавшего перенос
	this.From          = false;   // Фаза обработки (обрабатываем удаленную или вставленную часть)
	this.MovesToDelete = {};      // Если в процессе обработки встретились отметки других переносов, тогда мы превратим такие переносы в обычный добавлненый/удаленный текст
}
/**
 * Меняем фазу проверки
 * @param {boolean} isFrom
 */
CTrackRevisionsMoveProcessEngine.prototype.SetFrom = function(isFrom)
{
	this.From = isFrom;
};
/**
 * Получаем индентификатор переноса
 * @returns {string}
 */
CTrackRevisionsMoveProcessEngine.prototype.GetMoveId = function()
{
	return this.MoveId;
};
/**
 * Проверяем, происходит ли сейчас обработка удаленного текста во время переноса
 * @returns {boolean}
 */
CTrackRevisionsMoveProcessEngine.prototype.IsFrom = function()
{
	return this.From;
};
/**
 * Регистрируем наличие меток другого переноса во время обработки заданного переноса
 * @param sMoveId {string}
 */
CTrackRevisionsMoveProcessEngine.prototype.RegisterOtherMove = function(sMoveId)
{
	this.MovesToDelete[sMoveId] = sMoveId;
};
/**
 * Получаем идентификатор пользователя, сделавшего перенос, который сейчас обрабатывается
 * @returns {string}
 */
CTrackRevisionsMoveProcessEngine.prototype.GetUserId = function()
{
	return this.UserId;
};


function CRevisionsChangeParagraphSearchEngine(nDirection, oCurrentElement, oTrackManager)
{
    this.TrackManager   = oTrackManager;
    this.Direction      = nDirection;
    this.CurrentElement = oCurrentElement;
    this.CurrentFound   = false;
    this.Element        = null;
}
CRevisionsChangeParagraphSearchEngine.prototype.SetCurrentFound = function()
{
    this.CurrentFound = true;
};
CRevisionsChangeParagraphSearchEngine.prototype.SetCurrentElement = function(oElement)
{
    this.CurrentElement = oElement;
};
/**
 * Нашли ли мы текущий элемент
 * @returns {boolean}
 */
CRevisionsChangeParagraphSearchEngine.prototype.IsCurrentFound = function()
{
    return this.CurrentFound;
};
CRevisionsChangeParagraphSearchEngine.prototype.GetCurrentElement = function()
{
    return this.CurrentElement;
};
CRevisionsChangeParagraphSearchEngine.prototype.SetFoundedElement = function(oElement)
{
    if (this.TrackManager.GetElementChanges(oElement.GetId()).length > 0)
        this.Element = oElement;
};
/**
 * Нашли ли мы элемент с рецензированием
 * @returns {boolean}
 */
CRevisionsChangeParagraphSearchEngine.prototype.IsFound = function()
{
    return (null !== this.Element);
};
CRevisionsChangeParagraphSearchEngine.prototype.GetFoundedElement = function()
{
    return this.Element;
};
CRevisionsChangeParagraphSearchEngine.prototype.GetDirection = function()
{
    return this.Direction;
};

function CDocumentPagePosition()
{
    this.Page   = 0;
    this.Column = 0;
}

function CDocumentNumberingInfoCounter()
{
	this.Nums    = {}; // Список Num, которые использовались. Нужно для обработки startOverride
	this.NumInfo = new Array(9);
	this.PrevLvl = -1;

	for (var nIndex = 0; nIndex < 9; ++nIndex)
		this.NumInfo[nIndex] = undefined;
}
CDocumentNumberingInfoCounter.prototype.CheckNum = function(oNum)
{
	if (this.Nums[oNum.GetId()])
		return false;

	this.Nums[oNum.GetId()] = oNum;

	return true;
};

/**
 * Класс для рассчета значение номера для нумерации заданного параграфа
 * @param oPara {Paragraph}
 * @param oNumPr {CNumPr}
 * @param oNumbering {AscWord.CNumbering}
 * @constructor
 */
function CDocumentNumberingInfoEngine(oPara, oNumPr, oNumbering)
{
	this.Paragraph   = oPara;
	this.NumId       = oNumPr.NumId;
	this.Lvl         = oNumPr.Lvl !== undefined ? oNumPr.Lvl : 0;
	this.Numbering   = oNumbering;
	this.NumInfo     = new Array(this.Lvl + 1);
	this.Restart     = [-1, -1, -1, -1, -1, -1, -1, -1, -1]; // Этот параметр контролирует уровень, начиная с которого делаем рестарт для текущего уровня
	this.PrevLvl     = -1;
	this.Found       = false;
	this.AbstractNum = null;
	this.Start       = [];

	this.FinalCounter  = new CDocumentNumberingInfoCounter();
	this.SourceCounter = new CDocumentNumberingInfoCounter();

	for (var nIndex = 0; nIndex < 9; ++nIndex)
		this.NumInfo[nIndex] = undefined;

	if (!this.Numbering)
		return;

	var oNum = this.Numbering.GetNum(this.NumId);
	if (!oNum)
		return;

	var oAbstractNum = oNum.GetAbstractNum();
	if (oAbstractNum)
	{
		for (var nLvl = 0; nLvl < 9; ++nLvl)
		{
			this.Restart[nLvl] = oAbstractNum.GetLvl(nLvl).GetRestart();
			this.Start[nLvl]   = oAbstractNum.GetLvl(nLvl).GetStart() - 1;
		}
	}

	this.AbstractNum = oAbstractNum;
}

/**
 * Проверяем закончилось ли вычисление номера
 * @returns {boolean}
 */
CDocumentNumberingInfoEngine.prototype.IsStop = function()
{
    return this.Found;
};
/**
 * Проверяем параграф
 * @param oPara {Paragraph}
 */
CDocumentNumberingInfoEngine.prototype.CheckParagraph = function(oPara)
{
	if (!this.Numbering)
		return;

	if (this.Paragraph === oPara)
		this.Found = true;

	var oParaNumPr     = oPara.GetNumPr();
	var oParaNumPrPrev = oPara.GetPrChangeNumPr();
	var isPrChange     = oPara.HavePrChange();

	if (undefined !== oPara.Get_SectionPr() && true === oPara.IsEmpty())
		return;

	var isEqualNumPr = false;
	if (!isPrChange
		|| (isPrChange
			&& oParaNumPr
			&& oParaNumPrPrev
			&& oParaNumPr.NumId === oParaNumPrPrev.NumId
			&& oParaNumPr.Lvl === oParaNumPrPrev.Lvl
		)
	)
	{
		isEqualNumPr = true;
	}

	if (isEqualNumPr)
	{
		if (!oParaNumPr)
			return;

		var oNum         = this.Numbering.GetNum(oParaNumPr.NumId);
		var oAbstractNum = oNum.GetAbstractNum();

		if (oAbstractNum === this.AbstractNum)
		{
			var oReviewType = oPara.GetReviewType();
			var oReviewInfo = oPara.GetReviewInfo();

			if (reviewtype_Common === oReviewType)
			{
				this.private_UpdateCounter(this.FinalCounter, oNum, oParaNumPr.Lvl);
				this.private_UpdateCounter(this.SourceCounter, oNum, oParaNumPr.Lvl);
			}
			else if (reviewtype_Add === oReviewType)
			{
				this.private_UpdateCounter(this.FinalCounter, oNum, oParaNumPr.Lvl);
			}
			else if (reviewtype_Remove === oReviewType)
			{
				if (!oReviewInfo.GetPrevAdded())
					this.private_UpdateCounter(this.SourceCounter, oNum, oParaNumPr.Lvl);
			}
		}
	}
	else
	{
		if (oParaNumPr)
		{
			var oNum         = this.Numbering.GetNum(oParaNumPr.NumId);
			var oAbstractNum = oNum.GetAbstractNum();

			if (oAbstractNum === this.AbstractNum)
			{
				var oReviewType = oPara.GetReviewType();
				var oReviewInfo = oPara.GetReviewInfo();

				if (reviewtype_Common === oReviewType)
				{
					this.private_UpdateCounter(this.FinalCounter, oNum, oParaNumPr.Lvl);
				}
				else if (reviewtype_Add === oReviewType)
				{
					this.private_UpdateCounter(this.FinalCounter, oNum, oParaNumPr.Lvl);
				}
			}
		}

		if (oParaNumPrPrev)
		{
			var oNum = this.Numbering.GetNum(oParaNumPrPrev.NumId);
			if (oNum)
			{
				var oAbstractNum = oNum.GetAbstractNum();
				if (oAbstractNum === this.AbstractNum)
				{
					var oReviewType = oPara.GetReviewType();
					var oReviewInfo = oPara.GetReviewInfo();

					if (reviewtype_Common === oReviewType)
					{
						this.private_UpdateCounter(this.SourceCounter, oNum, oParaNumPrPrev.Lvl);
					}
					else if (reviewtype_Add === oReviewType)
					{
					}
					else if (reviewtype_Remove === oReviewType)
					{
						if (!oReviewInfo.GetPrevAdded())
							this.private_UpdateCounter(this.SourceCounter, oNum, oParaNumPrPrev.Lvl);
					}
				}
			}
		}
	}
};
CDocumentNumberingInfoEngine.prototype.GetNumInfo = function(isFinal)
{
	if (false === isFinal)
		return this.SourceCounter.NumInfo;

	return this.FinalCounter.NumInfo;
};
CDocumentNumberingInfoEngine.prototype.private_UpdateCounter = function(oCounter, oNum, nParaLvl)
{
	if (-1 === oCounter.PrevLvl)
	{
		for (var nLvl = 0; nLvl < 9; ++nLvl)
		{
			oCounter.NumInfo[nLvl] = this.Start[nLvl];
		}

		for (var nLvl = oCounter.PrevLvl + 1; nLvl < nParaLvl; ++nLvl)
		{
			oCounter.NumInfo[nLvl]++;
		}
	}
	else if (nParaLvl < oCounter.PrevLvl)
	{
		for (var nLvl = 0; nLvl < 9; ++nLvl)
		{
			if (nLvl > nParaLvl && 0 !== this.Restart[nLvl] && (-1 === this.Restart[nLvl] || nParaLvl <= this.Restart[nLvl] - 1))
				oCounter.NumInfo[nLvl] = this.Start[nLvl];
		}
	}
	else if (nParaLvl > oCounter.PrevLvl)
	{
		for (var nLvl = oCounter.PrevLvl + 1; nLvl < nParaLvl; ++nLvl)
		{
			oCounter.NumInfo[nLvl]++;
		}
	}

	oCounter.NumInfo[nParaLvl]++;

	if (oCounter.CheckNum(oNum))
	{
		var nForceStart = oNum.GetStartOverride(nParaLvl);
		if (-1 !== nForceStart)
			oCounter.NumInfo[nParaLvl] = nForceStart;
	}

	for (var nIndex = nParaLvl - 1; nIndex >= 0; --nIndex)
	{
		if (undefined === oCounter.NumInfo[nIndex] || 0 === oCounter.NumInfo[nIndex])
			oCounter.NumInfo[nIndex] = 1;
	}

	oCounter.PrevLvl = nParaLvl;
};

function CDocumentFootnotesRangeEngine(bExtendedInfo)
{
	this.m_oFirstFootnote = null; // Если не задана ищем с начала
	this.m_oLastFootnote  = null; // Если не задана ищем до конца
	this.m_bFootnotes     = true;
	this.m_bEndnotes      = true;

	this.m_arrFootnotes  = [];
	this.m_bForceStop    = false;

	this.m_bExtendedInfo = (true === bExtendedInfo ? true : false);
	this.m_arrParagraphs = [];
	this.m_oCurParagraph = null;
	this.m_arrRuns       = [];
	this.m_arrRefs       = [];
}
CDocumentFootnotesRangeEngine.prototype.Init = function(oFirstFootnote, oLastFootnote, isFootnotes, isEndnotes)
{
	this.m_oFirstFootnote = oFirstFootnote ? oFirstFootnote : null;
	this.m_oLastFootnote  = oLastFootnote ? oLastFootnote : null;
	this.m_bFootnotes     = isFootnotes;
	this.m_bEndnotes      = isEndnotes;
};
CDocumentFootnotesRangeEngine.prototype.Add = function(oFootnote, oFootnoteRef, oRun)
{
	if (!oFootnote || true === this.m_bForceStop)
		return;

	if (this.m_arrFootnotes.length <= 0 && null !== this.m_oFirstFootnote)
	{
		if (this.m_oFirstFootnote === oFootnote)
			this.private_AddFootnote(oFootnote, oFootnoteRef, oRun);
		else if (this.m_oLastFootnote === oFootnote)
			this.m_bForceStop = true;
	}
	else if (this.m_arrFootnotes.length >= 1 && null !== this.m_oLastFootnote)
	{
		if (this.m_oLastFootnote !== this.m_arrFootnotes[this.m_arrFootnotes.length - 1])
			this.private_AddFootnote(oFootnote, oFootnoteRef, oRun);
	}
	else
	{
		this.private_AddFootnote(oFootnote, oFootnoteRef, oRun);
	}
};
CDocumentFootnotesRangeEngine.prototype.IsRangeFull = function()
{
	if (true === this.m_bForceStop)
		return true;

	if (null !== this.m_oLastFootnote && this.m_arrFootnotes.length >= 1 && this.m_oLastFootnote === this.m_arrFootnotes[this.m_arrFootnotes.length - 1])
		return true;

	return false;
};
CDocumentFootnotesRangeEngine.prototype.GetRange = function()
{
	return this.m_arrFootnotes;
};
CDocumentFootnotesRangeEngine.prototype.GetParagraphs = function()
{
	return this.m_arrParagraphs;
};
CDocumentFootnotesRangeEngine.prototype.SetCurrentParagraph = function(oParagraph)
{
	this.m_oCurParagraph = oParagraph;
};
CDocumentFootnotesRangeEngine.prototype.private_AddFootnote = function(oFootnote, oFootnoteRef, oRun)
{
	this.m_arrFootnotes.push(oFootnote);

	if (true === this.m_bExtendedInfo)
	{
		this.m_arrRuns.push(oRun);
		this.m_arrRefs.push(oFootnoteRef);

		// Добавляем в общий список используемых параграфов
		var nCount = this.m_arrParagraphs.length;
		if (nCount <= 0 || this.m_arrParagraphs[nCount - 1] !== this.m_oCurParagraph)
			this.m_arrParagraphs[nCount] = this.m_oCurParagraph;
	}
};
CDocumentFootnotesRangeEngine.prototype.GetRuns = function()
{
	return this.m_arrRuns;
};
CDocumentFootnotesRangeEngine.prototype.GetRefs = function()
{
	return this.m_arrRefs;
};
CDocumentFootnotesRangeEngine.prototype.IsCheckFootnotes = function()
{
	return this.m_bFootnotes;
};
CDocumentFootnotesRangeEngine.prototype.IsCheckEndnotes = function()
{
	return this.m_bEndnotes;
};

/**
 * Класс для поиска подходящей нумерации в документе
 * @param oParagraph {Paragraph}
 * @param oNumPr {CNumPr}
 * @param oNumbering {AscWord.CNumbering}
 */
function CDocumentNumberingContinueEngine(oParagraph, oNumPr, oNumbering)
{
	this.Paragraph = oParagraph;
	this.NumPr     = oNumPr;
	this.Numbering = oNumbering;

	this.Found        = false;
	this.SimilarNumPr = null;
	this.LastNumPr    = null; // Для случая нулевого уровня, когда не получилось подобрать подходящий
}
/**
 * Проверем, дошли ли мы до заданного параграфа
 * @returns {boolean}
 */
CDocumentNumberingContinueEngine.prototype.IsFound = function()
{
	return this.Found;
};
/**
 * Проверяем параграф на совпадение нумераций
 * @param oParagraph {Paragraph}
 */
CDocumentNumberingContinueEngine.prototype.CheckParagraph = function(oParagraph)
{
	if (this.IsFound())
		return;

	if (oParagraph === this.Paragraph)
	{
		this.Found = true;
	}
	else
	{
		var oNumPr = oParagraph.GetNumPr();
		if (oNumPr && oNumPr.Lvl === this.NumPr.Lvl)
		{
			if (oNumPr.Lvl > 0)
			{
				this.SimilarNumPr = new CNumPr(oNumPr.NumId, this.NumPr.Lvl);
			}
			else
			{
				var oCurLvl = this.Numbering.GetNum(oNumPr.NumId).GetLvl(0);
				var oLvl    = this.Numbering.GetNum(this.NumPr.NumId).GetLvl(0);

				if (oCurLvl.IsSimilar(oLvl) || (oCurLvl.GetFormat() === oLvl.GetFormat() && Asc.c_oAscNumberingFormat.Bullet === oCurLvl.GetFormat()))
					this.SimilarNumPr = new CNumPr(oNumPr.NumId, 0);
			}

			this.LastNumPr = new CNumPr(oNumPr.NumId, this.NumPr.Lvl);
		}
	}
};
/**
 * Получаем подходящую нумерацию
 * @returns {?CNumPr}
 */
CDocumentNumberingContinueEngine.prototype.GetNumPr = function()
{
	return this.SimilarNumPr ? this.SimilarNumPr : this.LastNumPr;
};

function CDocumentLineNumbersInfo(oLogicDocument)
{
	this.LogicDocument = oLogicDocument;
	this.Widths        = [0, 0, 0, 0, 0, 0, 0, 0, 0];
	this.TextPr        = null;
	this.RecalcId      = null;
}
CDocumentLineNumbersInfo.prototype.GetTextPr = function()
{
	this.private_Update();
	return this.TextPr;
};
CDocumentLineNumbersInfo.prototype.GetWidths = function()
{
	this.private_Update();
	return this.Widths;
};
CDocumentLineNumbersInfo.prototype.private_Update = function()
{
	if (this.RecalcId && this.RecalcId === this.LogicDocument.GetRecalcId())
		return;

	this.RecalcId = this.LogicDocument.GetRecalcId();

	var oPr = this.LogicDocument.GetStyles().Get_Pr(undefined, styletype_Paragraph);
	this.TextPr = oPr.TextPr.Copy();

	var nFontKoef = 1;
	if (this.TextPr.VertAlign !== AscCommon.vertalign_Baseline)
		nFontKoef = AscCommon.vaKSize;

	g_oTextMeasurer.SetTextPr(this.TextPr, this.LogicDocument.GetTheme());
	g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, nFontKoef);
	for (var nDigit = 0; nDigit < 10; ++nDigit)
	{
		this.Widths[nDigit] = g_oTextMeasurer.MeasureCode(0x0030 + nDigit).Width;
	}
};

//-------------------------------------------------------------export---------------------------------------------------
window['Asc']            = window['Asc'] || {};
window['AscCommon']      = window['AscCommon'] || {};
window['AscCommonWord']  = window['AscCommonWord'] || {};
window['AscCommonSlide'] = window['AscCommonSlide'] || {};
window['AscCommonExcel'] = window['AscCommonExcel'] || {};
window['AscWord'] = window['AscWord'] || {};

window['AscWord'].CDocument = CDocument;
window['AscCommonWord'].CDocument = CDocument;
window['AscCommonWord'].docpostype_Content        = docpostype_Content;
window['AscCommonWord'].docpostype_HdrFtr         = docpostype_HdrFtr;
window['AscCommonWord'].docpostype_DrawingObjects = docpostype_DrawingObjects;
window['AscCommonWord'].docpostype_Footnotes      = docpostype_Footnotes;
window['AscCommonWord'].docpostype_Endnotes       = docpostype_Endnotes;

window['AscCommon'].Page_Width = Page_Width;
window['AscCommon'].Page_Height = Page_Height;
window['AscCommon'].X_Left_Margin = X_Left_Margin;
window['AscCommon'].X_Right_Margin = X_Right_Margin;
window['AscCommon'].Y_Bottom_Margin = Y_Bottom_Margin;
window['AscCommon'].Y_Top_Margin = Y_Top_Margin;
window['AscCommon'].selectionflag_Common = selectionflag_Common;

CDocumentColumnProps.prototype['put_W']     = CDocumentColumnProps.prototype.put_W;
CDocumentColumnProps.prototype['get_W']     = CDocumentColumnProps.prototype.get_W;
CDocumentColumnProps.prototype['put_Space'] = CDocumentColumnProps.prototype.put_Space;
CDocumentColumnProps.prototype['get_Space'] = CDocumentColumnProps.prototype.get_Space;

window['Asc']['CDocumentColumnsProps'] = CDocumentColumnsProps;
CDocumentColumnsProps.prototype['get_EqualWidth'] = CDocumentColumnsProps.prototype.get_EqualWidth;
CDocumentColumnsProps.prototype['put_EqualWidth'] = CDocumentColumnsProps.prototype.put_EqualWidth;
CDocumentColumnsProps.prototype['get_Num']        = CDocumentColumnsProps.prototype.get_Num       ;
CDocumentColumnsProps.prototype['put_Num']        = CDocumentColumnsProps.prototype.put_Num       ;
CDocumentColumnsProps.prototype['get_Sep']        = CDocumentColumnsProps.prototype.get_Sep       ;
CDocumentColumnsProps.prototype['put_Sep']        = CDocumentColumnsProps.prototype.put_Sep       ;
CDocumentColumnsProps.prototype['get_Space']      = CDocumentColumnsProps.prototype.get_Space     ;
CDocumentColumnsProps.prototype['put_Space']      = CDocumentColumnsProps.prototype.put_Space     ;
CDocumentColumnsProps.prototype['get_ColsCount']  = CDocumentColumnsProps.prototype.get_ColsCount ;
CDocumentColumnsProps.prototype['get_Col']        = CDocumentColumnsProps.prototype.get_Col       ;
CDocumentColumnsProps.prototype['put_Col']        = CDocumentColumnsProps.prototype.put_Col       ;
CDocumentColumnsProps.prototype['put_ColByValue'] = CDocumentColumnsProps.prototype.put_ColByValue;
CDocumentColumnsProps.prototype['get_TotalWidth'] = CDocumentColumnsProps.prototype.get_TotalWidth;

window['Asc']['CDocumentSectionProps'] = window['Asc'].CDocumentSectionProps = CDocumentSectionProps;
CDocumentSectionProps.prototype["get_W"]              = CDocumentSectionProps.prototype.get_W;
CDocumentSectionProps.prototype["put_W"]              = CDocumentSectionProps.prototype.put_W;
CDocumentSectionProps.prototype["get_H"]              = CDocumentSectionProps.prototype.get_H;
CDocumentSectionProps.prototype["put_H"]              = CDocumentSectionProps.prototype.put_H;
CDocumentSectionProps.prototype["get_Orientation"]    = CDocumentSectionProps.prototype.get_Orientation;
CDocumentSectionProps.prototype["put_Orientation"]    = CDocumentSectionProps.prototype.put_Orientation;
CDocumentSectionProps.prototype["get_LeftMargin"]     = CDocumentSectionProps.prototype.get_LeftMargin;
CDocumentSectionProps.prototype["put_LeftMargin"]     = CDocumentSectionProps.prototype.put_LeftMargin;
CDocumentSectionProps.prototype["get_TopMargin"]      = CDocumentSectionProps.prototype.get_TopMargin;
CDocumentSectionProps.prototype["put_TopMargin"]      = CDocumentSectionProps.prototype.put_TopMargin;
CDocumentSectionProps.prototype["get_RightMargin"]    = CDocumentSectionProps.prototype.get_RightMargin;
CDocumentSectionProps.prototype["put_RightMargin"]    = CDocumentSectionProps.prototype.put_RightMargin;
CDocumentSectionProps.prototype["get_BottomMargin"]   = CDocumentSectionProps.prototype.get_BottomMargin;
CDocumentSectionProps.prototype["put_BottomMargin"]   = CDocumentSectionProps.prototype.put_BottomMargin;
CDocumentSectionProps.prototype["get_HeaderDistance"] = CDocumentSectionProps.prototype.get_HeaderDistance;
CDocumentSectionProps.prototype["put_HeaderDistance"] = CDocumentSectionProps.prototype.put_HeaderDistance;
CDocumentSectionProps.prototype["get_FooterDistance"] = CDocumentSectionProps.prototype.get_FooterDistance;
CDocumentSectionProps.prototype["put_FooterDistance"] = CDocumentSectionProps.prototype.put_FooterDistance;
CDocumentSectionProps.prototype["get_Gutter"]         = CDocumentSectionProps.prototype.get_Gutter;
CDocumentSectionProps.prototype["put_Gutter"]         = CDocumentSectionProps.prototype.put_Gutter;
CDocumentSectionProps.prototype["get_GutterRTL"]      = CDocumentSectionProps.prototype.get_GutterRTL;
CDocumentSectionProps.prototype["put_GutterRTL"]      = CDocumentSectionProps.prototype.put_GutterRTL;
CDocumentSectionProps.prototype["get_GutterAtTop"]    = CDocumentSectionProps.prototype.get_GutterAtTop;
CDocumentSectionProps.prototype["put_GutterAtTop"]    = CDocumentSectionProps.prototype.put_GutterAtTop;
CDocumentSectionProps.prototype["get_MirrorMargins"]  = CDocumentSectionProps.prototype.get_MirrorMargins;
CDocumentSectionProps.prototype["put_MirrorMargins"]  = CDocumentSectionProps.prototype.put_MirrorMargins;
