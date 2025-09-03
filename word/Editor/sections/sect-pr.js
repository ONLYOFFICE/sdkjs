/*
 * (c) Copyright Ascensio System SIA 2010-2025
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

var section_borders_DisplayAllPages     = 0x00;
var section_borders_DisplayFirstPage    = 0x01;
var section_borders_DisplayNotFirstPage = 0x02;

var section_borders_OffsetFromPage = 0x00;
var section_borders_OffsetFromText = 0x01;

var section_borders_ZOrderBack  = 0x00;
var section_borders_ZOrderFront = 0x01;

var section_footnote_RestartContinuous = 0x00;
var section_footnote_RestartEachSect   = 0x01;
var section_footnote_RestartEachPage   = 0x02;

(function()
{
	/**
	 * @param {AscWord.Document} logicDocument
	 * @constructor
	 */
	function SectPr(logicDocument)
	{
		this.Id = AscCommon.g_oIdCounter.Get_NewId();
		
		this.Type        = Asc.c_oAscSectionBreakType.NextPage;
		this.PageSize    = new AscWord.SectionPageSize();
		this.PageMargins = new AscWord.SectionPageMargins();
		
		this.LogicDocument = logicDocument;
		
		this.Borders     = new AscWord.SectionPageBorders();
		this.PageNumType = new AscWord.SectionPageNumType();
		
		this.FooterFirst   = null;
		this.FooterEven    = null;
		this.FooterDefault = null;
		
		this.HeaderFirst   = null;
		this.HeaderEven    = null;
		this.HeaderDefault = null;
		
		this.TitlePage = false;
		this.GutterRTL = false;
		
		this.Columns    = new AscWord.SectionColumns(this);
		this.FootnotePr = new AscWord.FootnotePr();
		this.EndnotePr  = new AscWord.FootnotePr();
		
		this.LnNumType = undefined;
		
		// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
		AscCommon.g_oTableId.Add(this, this.Id);
	}
	
	SectPr.prototype.Get_Id = function()
	{
		return this.Id;
	};
	SectPr.prototype.Copy = function(Other, CopyHdrFtr, oCopyPr)
	{
		if (!Other)
			return;
		
		// Тип
		this.Set_Type(Other.Type);
		
		// Настройки страницы
		this.SetPageSize(Other.PageSize.W, Other.PageSize.H);
		this.SetOrientation(Other.PageSize.Orient, false);
		
		// Настройки отступов
		this.SetPageMargins(Other.PageMargins.Left, Other.PageMargins.Top, Other.PageMargins.Right, Other.PageMargins.Bottom);
		this.SetGutter(Other.PageMargins.Gutter);
		
		// Настройки границ
		this.Set_Borders_Left(Other.Borders.Left);
		this.Set_Borders_Top(Other.Borders.Top);
		this.Set_Borders_Right(Other.Borders.Right);
		this.Set_Borders_Bottom(Other.Borders.Bottom);
		this.Set_Borders_Display(Other.Borders.Display);
		this.SetBordersOffsetFrom(Other.Borders.OffsetFrom);
		this.Set_Borders_ZOrder(Other.Borders.ZOrder);
		this.Set_TitlePage(Other.TitlePage);
		this.SetGutterRTL(Other.GutterRTL);
		
		// Колонтитулы
		if (true === CopyHdrFtr)
		{
			if (Other.HeaderFirst)
				this.Set_Header_First(Other.HeaderFirst.Copy(this.LogicDocument, oCopyPr));
			else
				this.Set_Header_First(null);
			
			if (Other.HeaderEven)
				this.Set_Header_Even(Other.HeaderEven.Copy(this.LogicDocument, oCopyPr));
			else
				this.Set_Header_Even(null);
			
			if (Other.HeaderDefault)
				this.Set_Header_Default(Other.HeaderDefault.Copy(this.LogicDocument, oCopyPr));
			else
				this.Set_Header_Default(null);
			
			if (Other.FooterFirst)
				this.Set_Footer_First(Other.FooterFirst.Copy(this.LogicDocument, oCopyPr));
			else
				this.Set_Footer_First(null);
			
			if (Other.FooterEven)
				this.Set_Footer_Even(Other.FooterEven.Copy(this.LogicDocument, oCopyPr));
			else
				this.Set_Footer_Even(null);
			
			if (Other.FooterDefault)
				this.Set_Footer_Default(Other.FooterDefault.Copy(this.LogicDocument, oCopyPr));
			else
				this.Set_Footer_Default(null);
		}
		else
		{
			this.Set_Header_First(Other.HeaderFirst);
			this.Set_Header_Even(Other.HeaderEven);
			this.Set_Header_Default(Other.HeaderDefault);
			this.Set_Footer_First(Other.FooterFirst);
			this.Set_Footer_Even(Other.FooterEven);
			this.Set_Footer_Default(Other.FooterDefault);
		}
		
		this.SetPageNumStart(Other.PageNumType.Start);
		this.SetPageNumFormat(Other.PageNumType.Format);
		this.SetPageNumChapStyle(Other.PageNumType.ChapStyle);
		this.SetPageNumChapSep(Other.PageNumType.ChapSep);
		
		this.Set_Columns_EqualWidth(Other.Columns.EqualWidth);
		this.Set_Columns_Num(Other.Columns.Num);
		this.Set_Columns_Sep(Other.Columns.Sep);
		this.Set_Columns_Space(Other.Columns.Space);
		
		let columns = [];
		for (let columnIndex = 0, columnCount = Other.Columns.Cols.length; columnIndex < columnCount; ++columnIndex)
		{
			let column = Other.Columns.Cols[columnIndex];
			columns.push(new AscWord.SectionColumn(column.W, column.Space));
		}
		this.Set_Columns_Cols(columns);
		
		this.SetFootnotePos(Other.FootnotePr.Pos);
		this.SetFootnoteNumStart(Other.FootnotePr.NumStart);
		this.SetFootnoteNumRestart(Other.FootnotePr.NumRestart);
		this.SetFootnoteNumFormat(Other.FootnotePr.NumFormat);
		
		if (Other.HaveLineNumbers())
			this.SetLineNumbers(Other.GetLineNumbersCountBy(), Other.GetLineNumbersDistance(), Other.GetLineNumbersStart(), Other.GetLineNumbersRestart());
	};
	SectPr.prototype.Clear_AllHdrFtr = function()
	{
		this.Set_Header_First(null);
		this.Set_Header_Even(null);
		this.Set_Header_Default(null);
		this.Set_Footer_First(null);
		this.Set_Footer_Even(null);
		this.Set_Footer_Default(null);
	};
	SectPr.prototype.GetAllHdrFtrs = function(HdrFtrs)
	{
		if (!HdrFtrs)
			HdrFtrs = [];
		
		if (null !== this.HeaderFirst) HdrFtrs.push(this.HeaderFirst);
		if (null !== this.HeaderEven) HdrFtrs.push(this.HeaderEven);
		if (null !== this.HeaderDefault) HdrFtrs.push(this.HeaderDefault);
		
		if (null !== this.FooterFirst) HdrFtrs.push(this.FooterFirst);
		if (null !== this.FooterEven) HdrFtrs.push(this.FooterEven);
		if (null !== this.FooterDefault) HdrFtrs.push(this.FooterDefault);
		
		return HdrFtrs;
	};
	SectPr.prototype.RemoveEmptyHdrFtrs = function()
	{
		function IsEmpty(oHeader)
		{
			return (oHeader && oHeader.GetContent().GetElementsCount() <= 0);
		}
		
		if (IsEmpty(this.HeaderFirst))
			this.Set_Header_First(null);
		
		if (IsEmpty(this.HeaderEven))
			this.Set_Header_Even(null);
		
		if (IsEmpty(this.HeaderDefault))
			this.Set_Header_Default(null);
		
		if (IsEmpty(this.FooterFirst))
			this.Set_Footer_First(null);
		
		if (IsEmpty(this.FooterEven))
			this.Set_Footer_Even(null);
		
		if (IsEmpty(this.FooterDefault))
			this.Set_Footer_Default(null);
	};
	SectPr.prototype.Compare_PageSize = function(OtherSectPr)
	{
		var ThisPS  = this.PageSize;
		var OtherPS = OtherSectPr.PageSize;
		
		if (Math.abs(ThisPS.W - OtherPS.W) > 0.001 || Math.abs(ThisPS.H - OtherPS.H) > 0.001 || ThisPS.Orient !== OtherPS.Orient)
			return false;
		
		return true;
	};
	SectPr.prototype.Set_Type = function(Type)
	{
		if (this.Type !== Type)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionType(this, this.Type, Type));
			this.Type = Type;
		}
	};
	SectPr.prototype.Get_Type = function()
	{
		return this.Type;
	};
	SectPr.prototype.Set_Borders_Left = function(Border)
	{
		if (true !== this.Borders.Left.Compare(Border))
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionBordersLeft(this, this.Borders.Left, Border));
			this.Borders.Left = Border;
		}
	};
	SectPr.prototype.Get_Borders_Left = function()
	{
		return this.Borders.Left;
	};
	SectPr.prototype.Set_Borders_Top = function(Border)
	{
		if (true !== this.Borders.Top.Compare(Border))
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionBordersTop(this, this.Borders.Top, Border));
			this.Borders.Top = Border;
		}
	};
	SectPr.prototype.Get_Borders_Top = function()
	{
		return this.Borders.Top;
	};
	SectPr.prototype.Set_Borders_Right = function(Border)
	{
		if (true !== this.Borders.Right.Compare(Border))
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionBordersRight(this, this.Borders.Right, Border));
			this.Borders.Right = Border;
		}
	};
	SectPr.prototype.Get_Borders_Right = function()
	{
		return this.Borders.Right;
	};
	SectPr.prototype.Set_Borders_Bottom = function(Border)
	{
		if (true !== this.Borders.Bottom.Compare(Border))
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionBordersBottom(this, this.Borders.Bottom, Border));
			this.Borders.Bottom = Border;
		}
	};
	SectPr.prototype.Get_Borders_Bottom = function()
	{
		return this.Borders.Bottom;
	};
	SectPr.prototype.Set_Borders_Display = function(Display)
	{
		if (Display !== this.Borders.Display)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionBordersDisplay(this, this.Borders.Display, Display));
			this.Borders.Display = Display;
		}
	};
	SectPr.prototype.Get_Borders_Display = function()
	{
		return this.Borders.Display;
	};
	SectPr.prototype.Set_Borders_ZOrder = function(ZOrder)
	{
		if (ZOrder !== this.Borders.ZOrder)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionBordersZOrder(this, this.Borders.ZOrder, ZOrder));
			this.Borders.ZOrder = ZOrder;
		}
	};
	SectPr.prototype.Get_Borders_ZOrder = function()
	{
		return this.Borders.ZOrder;
	};
	SectPr.prototype.Set_Footer_First = function(Footer)
	{
		if (Footer !== this.FooterFirst)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionFooterFirst(this, this.FooterFirst, Footer));
			this.FooterFirst = Footer;
		}
	};
	SectPr.prototype.Get_Footer_First = function()
	{
		return this.FooterFirst;
	};
	SectPr.prototype.Set_Footer_Even = function(Footer)
	{
		if (Footer !== this.FooterEven)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionFooterEven(this, this.FooterEven, Footer));
			this.FooterEven = Footer;
		}
	};
	SectPr.prototype.Get_Footer_Even = function()
	{
		return this.FooterEven;
	};
	SectPr.prototype.Set_Footer_Default = function(Footer)
	{
		if (Footer !== this.FooterDefault)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionFooterDefault(this, this.FooterDefault, Footer));
			this.FooterDefault = Footer;
		}
	};
	SectPr.prototype.Get_Footer_Default = function()
	{
		return this.FooterDefault;
	};
	SectPr.prototype.Set_Header_First = function(Header)
	{
		if (Header !== this.HeaderFirst)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionHeaderFirst(this, this.HeaderFirst, Header));
			this.HeaderFirst = Header;
		}
	};
	SectPr.prototype.Get_Header_First = function()
	{
		return this.HeaderFirst;
	};
	SectPr.prototype.Set_Header_Even = function(Header)
	{
		if (Header !== this.HeaderEven)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionHeaderEven(this, this.HeaderEven, Header));
			this.HeaderEven = Header;
		}
	};
	SectPr.prototype.Get_Header_Even = function()
	{
		return this.HeaderEven;
	};
	SectPr.prototype.Set_Header_Default = function(Header)
	{
		if (Header !== this.HeaderDefault)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionHeaderDefault(this, this.HeaderDefault, Header));
			this.HeaderDefault = Header;
		}
	};
	SectPr.prototype.Get_Header_Default = function()
	{
		return this.HeaderDefault;
	};
	SectPr.prototype.Set_TitlePage = function(Value)
	{
		if (Value !== this.TitlePage)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionTitlePage(this, this.TitlePage, Value));
			this.TitlePage = Value;
		}
	};
	SectPr.prototype.Get_TitlePage = function()
	{
		return this.TitlePage;
	};
	SectPr.prototype.IsTitlePage = function()
	{
		return this.TitlePage;
	};
	SectPr.prototype.IsEvenAndOdd = function()
	{
		return EvenAndOddHeaders;
	};
	SectPr.prototype.GetHdrFtr = function(bHeader, bFirst, bEven)
	{
		if (true === bHeader)
		{
			if (true === bFirst)
				return this.HeaderFirst;
			else if (true === bEven)
				return this.HeaderEven;
			else
				return this.HeaderDefault;
		}
		else
		{
			if (true === bFirst)
				return this.FooterFirst;
			else if (true === bEven)
				return this.FooterEven;
			else
				return this.FooterDefault;
		}
	};
	SectPr.prototype.Set_HdrFtr = function(bHeader, bFirst, bEven, HdrFtr)
	{
		if (true === bHeader)
		{
			if (true === bFirst)
				return this.Set_Header_First(HdrFtr);
			else if (true === bEven)
				return this.Set_Header_Even(HdrFtr);
			else
				return this.Set_Header_Default(HdrFtr);
		}
		else
		{
			if (true === bFirst)
				return this.Set_Footer_First(HdrFtr);
			else if (true === bEven)
				return this.Set_Footer_Even(HdrFtr);
			else
				return this.Set_Footer_Default(HdrFtr);
		}
	};
	SectPr.prototype.GetHdrFtrInfo = function(HdrFtr)
	{
		if (HdrFtr === this.HeaderFirst)
			return {Header : true, First : true, Even : false};
		else if (HdrFtr === this.HeaderEven)
			return {Header : true, First : false, Even : true};
		else if (HdrFtr === this.HeaderDefault)
			return {Header : true, First : false, Even : false};
		else if (HdrFtr === this.FooterFirst)
			return {Header : false, First : true, Even : false};
		else if (HdrFtr === this.FooterEven)
			return {Header : false, First : false, Even : true};
		else if (HdrFtr === this.FooterDefault)
			return {Header : false, First : false, Even : false};
		
		return null;
	};
	SectPr.prototype.IsDefaultPageNum = function()
	{
		return (-1 === this.PageNumType.Start
			&& Asc.c_oAscNumberingFormat.Decimal === this.PageNumType.Format
			&& undefined === this.PageNumType.ChapStyle
			&& undefined === this.PageNumType.ChapSep);
	};
	SectPr.prototype.SetPageNumStart = function(Start)
	{
		if (Start !== this.PageNumType.Start)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionPageNumTypeStart(this, this.PageNumType.Start, Start));
			this.PageNumType.Start = Start;
		}
	};
	SectPr.prototype.GetPageNumStart = function()
	{
		return this.PageNumType.Start;
	};
	SectPr.prototype.SetPageNumFormat = function(format)
	{
		if (format === this.PageNumType.Format)
			return;
		
		AscCommon.History.Add(new AscDFH.CChangesSectionPageNumTypeFormat(this, this.PageNumType.Format, format));
		this.PageNumType.Format = format;
	};
	SectPr.prototype.GetPageNumFormat = function()
	{
		return this.PageNumType.Format;
	};
	SectPr.prototype.SetPageNumChapStyle = function(chapStyle)
	{
		if (chapStyle === this.PageNumType.ChapStyle)
			return;
		
		AscCommon.History.Add(new AscDFH.CChangesSectionPageNumTypeChapStyle(this, this.PageNumType.ChapStyle, chapStyle));
		this.PageNumType.ChapStyle = chapStyle;
	};
	SectPr.prototype.GetPageNumChapStyle = function()
	{
		return this.PageNumType.ChapStyle;
	};
	SectPr.prototype.SetPageNumChapSep = function(chapSep)
	{
		if (chapSep === this.PageNumType.ChapSep)
			return;
		
		AscCommon.History.Add(new AscDFH.CChangesSectionPageNumTypeChapSep(this, this.PageNumType.ChapSep, chapSep));
		this.PageNumType.ChapSep = chapSep;
	};
	SectPr.prototype.GetPageNumChapSep = function()
	{
		return this.PageNumType.ChapSep;
	};
	SectPr.prototype.Get_ColumnWidth = function(ColIndex)
	{
		return this.Columns.Get_ColumnWidth(ColIndex);
	};
	SectPr.prototype.Get_ColumnSpace = function(ColIndex)
	{
		return this.Columns.Get_ColumnSpace(ColIndex);
	};
	SectPr.prototype.Get_ColumnsSep = function()
	{
		return this.Columns.Sep;
	};
	SectPr.prototype.Set_Columns_EqualWidth = function(Equal)
	{
		if (Equal !== this.Columns.Equal)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionColumnsEqualWidth(this, this.Columns.EqualWidth, Equal));
			this.Columns.EqualWidth = Equal;
		}
	};
	SectPr.prototype.Set_Columns_Space = function(Space)
	{
		if (Space !== this.Columns.Space)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionColumnsSpace(this, this.Columns.Space, Space));
			this.Columns.Space = Space;
		}
	};
	SectPr.prototype.Set_Columns_Num = function(_Num)
	{
		var Num = Math.max(_Num, 1);
		
		if (Num !== this.Columns.Num)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionColumnsNum(this, this.Columns.Num, Num));
			this.Columns.Num = Num;
		}
	};
	SectPr.prototype.SetColumnsNum = function(num)
	{
		return this.Set_Columns_Num(num);
	};
	SectPr.prototype.Set_Columns_Sep = function(Sep)
	{
		if (Sep !== this.Columns.Sep)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionColumnsSep(this, this.Columns.Sep, Sep));
			this.Columns.Sep = Sep;
		}
	};
	SectPr.prototype.Set_Columns_Cols = function(Cols)
	{
		AscCommon.History.Add(new AscDFH.CChangesSectionColumnsSetCols(this, this.Columns.Cols, Cols));
		this.Columns.Cols = Cols;
	};
	SectPr.prototype.Set_Columns_Col = function(Index, W, Space)
	{
		var OldCol = this.Columns.Cols[Index];
		
		if (undefined === OldCol || OldCol.Space !== Space || OldCol.W !== W)
		{
			var NewCol   = new AscWord.SectionColumn();
			NewCol.W     = W;
			NewCol.Space = Space;
			
			AscCommon.History.Add(new AscDFH.CChangesSectionColumnsCol(this, OldCol, NewCol, Index));
			this.Columns.Cols[Index] = NewCol;
		}
	};
	SectPr.prototype.Get_LayoutInfo = function()
	{
		// Получаем информацию о колонках в данной секции
		
		var Margins = this.PageMargins;
		var H       = this.PageSize.H;
		var _W      = this.PageSize.W;
		var W       = _W - Margins.Left - Margins.Right;
		
		// Если так случилось, что правое и левое поля в сумме больше ширины, тогда оставляем для документа 1 см ширины.
		if (W < 0)
			W = 10;
		
		var Columns = this.Columns;
		
		var Layout      = new SectionLayoutInfo(Margins.Left, Margins.Top, _W - Margins.Right, H - Margins.Bottom);
		var ColumnsInfo = Layout.Columns;
		
		if (true === Columns.EqualWidth)
		{
			var Num   = Math.max(Columns.Num, 1);
			var Space = Columns.Space;
			
			var ColW = (W - Space * (Num - 1)) / Num;
			
			// Если так случилось, что под колонки места не осталось, тогда делаем колонки шириной 0.3 мм, оставшееся
			// свободное место распределяем под Space, но если и оставшегося места не осталось, тогда Space делаем 0, а
			// колонки пусть выходят за пределы W.
			if (ColW < 0)
			{
				ColW = 0.3;
				
				var __W = W - ColW * Num;
				
				if (_W > 0 && Num > 1)
					Space = _W / (Num - 1);
				else
					Space = 0;
			}
			
			var X = Margins.Left;
			for (var Pos = 0; Pos < Num; Pos++)
			{
				var X0 = X;
				var X1 = X + ColW;
				
				ColumnsInfo.push(new SectionLayoutColumnInfo(X0, X1));
				
				X += ColW + Space;
			}
		}
		else
		{
			var Num = Columns.Cols.length;
			
			// Когда задаются колонки неравномерно, то Word плюет на поля, заданные в документе и ориентируется только
			// по размеру колонок дальше. (если ни 1 колонка не задана, тогда Word добавляет 1 колонку шириной 17.09 см)
			
			if (Num <= 0)
			{
				ColumnsInfo.push(new SectionLayoutColumnInfo(Margins.Left, Margins.Left + 170.9));
			}
			else
			{
				var X = Margins.Left;
				for (var Pos = 0; Pos < Num; Pos++)
				{
					var Col = this.Columns.Cols[Pos];
					var X0  = X;
					var X1  = X + Col.W;
					
					ColumnsInfo.push(new SectionLayoutColumnInfo(X0, X1));
					
					X += Col.W + Col.Space;
				}
			}
		}
		
		return Layout;
	};
	//------------------------------------------------------------------------------------------------------------------
	// Undo/Redo функции
	//------------------------------------------------------------------------------------------------------------------
	SectPr.prototype.Refresh_RecalcData = function(Data)
	{
		let sectionIndex = this.LogicDocument.SectionsInfo.Find(this);
		if (-1 === sectionIndex)
			return;
		
		if (AscDFH.historyitem_Section_LnNumType === Data.Type)
		{
			AscCommon.History.AddLineNumbersToRecalculateData();
			return;
		}
		
		let logicDocument = this.LogicDocument;
		function refreshRecalc(sectionIndex)
		{
			if (0 === sectionIndex)
			{
				logicDocument.Refresh_RecalcData2(0, 0);
			}
			else
			{
				let paragraph = logicDocument.SectionsInfo.Elements[sectionIndex - 1].Paragraph;
				if (paragraph)
				{
					let nextParagraph = paragraph.GetNextParagraph();
					if (nextParagraph)
						paragraph = nextParagraph;
					
					paragraph.Refresh_RecalcData2(0);
				}
			}
		}
		
		// Здесь есть 1 исключение: когда мы добавляем колонтитул для первой страницы, может так получиться, что
		// у данной секции флаг TitlePage = False, а значит пересчет надо запускать с места, где данный колонтитул
		// первый раз начнет использоваться, а не с текущей секции.
		
		if ((AscDFH.historyitem_Section_Header_First === Data.Type || AscDFH.historyitem_Section_Footer_First === Data.Type) && false === this.TitlePage)
		{
			var bHeader       = AscDFH.historyitem_Section_Header_First === Data.Type ? true : false
			var SectionsCount = logicDocument.SectionsInfo.GetSectionsCount();
			while (sectionIndex < SectionsCount - 1)
			{
				++sectionIndex;
				
				let sectPr = logicDocument.SectionsInfo.GetSectPrByIndex(sectionIndex);
				
				// Если в следующей секции свой колонтитул, тогда наш добавленный колонтитул вообще ни на что не влияет
				if ((true === bHeader && null !== sectPr.Get_Header_First()) || (true !== bHeader && null !== sectPr.Get_Footer_First()))
					break;
				
				// Если в следующей секции есть титульная страница, значит мы нашли нужную секцию
				if (true === sectPr.Get_TitlePage())
					refreshRecalc(sectionIndex);
			}
		}
		else
		{
			refreshRecalc(sectionIndex);
		}
		
		// Дополнительно кроме этого мы должны обновить пересчет в колонтитулах, причем только начиная с данной секции
		this.LogicDocument.On_SectionChange(this);
	};
	//------------------------------------------------------------------------------------------------------------------
	// Функции совместного редактирования
	//------------------------------------------------------------------------------------------------------------------
	SectPr.prototype.Write_ToBinary2 = function(Writer)
	{
		Writer.WriteLong(AscDFH.historyitem_type_Section);
		
		// String2  : Id
		// String2  : Id LogicDocument
		// Variable : PageSize
		// Variable : PageMargins
		// Byte     : Type
		// Variable : Borders
		// Колонтитулы не пишем в бинарник, при создании класса они всегда null, а TitlePage = false
		// Variable : PageNumType
		// Variable : AscWord.SectionColumns
		// Variable : AscWord.FootnotePr
		// Bool     : GutterRTL
		
		Writer.WriteString2("" + this.Id);
		Writer.WriteString2("" + this.LogicDocument.Get_Id());
		this.PageSize.Write_ToBinary(Writer);
		this.PageMargins.Write_ToBinary(Writer);
		Writer.WriteByte(this.Type);
		this.Borders.Write_ToBinary(Writer);
		this.PageNumType.Write_ToBinary(Writer);
		this.Columns.Write_ToBinary(Writer);
		this.FootnotePr.WriteToBinary(Writer);
		Writer.WriteBool(this.GutterRTL);
	};
	SectPr.prototype.Read_FromBinary2 = function(Reader)
	{
		// String2  : Id
		// String2  : Id LogicDocument
		// Variable : PageSize
		// Variable : PageMargins
		// Byte     : Type
		// Variable : Borders
		// Колонтитулы не пишем в бинарник, при создании класса они всегда null, а TitlePage = false
		// Variable : PageNumType
		// Variable : AscWord.SectionColumns
		// Variable : AscWord.FootnotePr
		// Bool     : GutterRTL
		
		this.Id            = Reader.GetString2();
		this.LogicDocument = AscCommon.g_oTableId.Get_ById(Reader.GetString2());
		this.PageSize.Read_FromBinary(Reader);
		this.PageMargins.Read_FromBinary(Reader);
		this.Type = Reader.GetByte();
		this.Borders.Read_FromBinary(Reader);
		this.PageNumType.Read_FromBinary(Reader);
		this.Columns.Read_FromBinary(Reader);
		this.FootnotePr.ReadFromBinary(Reader);
		this.GutterRTL = Reader.GetBool();
	};
	//------------------------------------------------------------------------------------------------------------------
	/**
	 * @returns {string}
	 */
	SectPr.prototype.GetId = function()
	{
		return this.Id;
	};
	SectPr.prototype.SetType = function(type)
	{
		return this.Set_Type(type);
	};
	SectPr.prototype.GetType = function()
	{
		return this.Get_Type();
	};
	/**
	 * Проверяем, есть ли хоть один колонтитул в данной секции
	 * @returns {boolean}
	 */
	SectPr.prototype.IsAllHdrFtrNull = function()
	{
		if (null !== this.FooterFirst
			|| null !== this.HeaderFirst
			|| null !== this.FooterDefault
			|| null !== this.HeaderDefault
			|| null !== this.FooterEven
			|| null !== this.HeaderEven)
			return false;
		
		return true;
	};
	SectPr.prototype.GetFootnotePr                     = function()
	{
		return this.FootnotePr;
	};
	SectPr.prototype.SetFootnotePos                    = function(nPos)
	{
		if (this.FootnotePr.Pos !== nPos)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionFootnotePos(this, this.FootnotePr.Pos, nPos));
			this.FootnotePr.Pos = nPos;
		}
	};
	SectPr.prototype.GetFootnotePos                    = function()
	{
		if (undefined === this.FootnotePr.Pos)
			return Asc.c_oAscFootnotePos.PageBottom;
		
		return this.FootnotePr.Pos;
	};
	SectPr.prototype.SetFootnoteNumStart               = function(nStart)
	{
		if (this.FootnotePr.NumStart !== nStart)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionFootnoteNumStart(this, this.FootnotePr.NumStart, nStart));
			this.FootnotePr.NumStart = nStart;
		}
	};
	SectPr.prototype.GetFootnoteNumStart               = function()
	{
		if (undefined === this.FootnotePr.NumStart)
			return 1;
		
		return this.FootnotePr.NumStart;
	};
	SectPr.prototype.SetFootnoteNumRestart             = function(nRestartType)
	{
		if (this.FootnotePr.NumRestart !== nRestartType)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionFootnoteNumRestart(this, this.FootnotePr.NumRestart, nRestartType));
			this.FootnotePr.NumRestart = nRestartType;
		}
	};
	SectPr.prototype.GetFootnoteNumRestart             = function()
	{
		if (undefined === this.FootnotePr.NumRestart)
			return this.private_GetDocumentWideFootnotePr().NumRestart;
		
		return this.FootnotePr.NumRestart;
	};
	SectPr.prototype.SetFootnoteNumFormat              = function(nFormatType)
	{
		if (this.FootnotePr.NumFormat !== nFormatType)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionFootnoteNumFormat(this, this.FootnotePr.NumFormat, nFormatType));
			this.FootnotePr.NumFormat = nFormatType;
		}
	};
	SectPr.prototype.GetFootnoteNumFormat              = function()
	{
		if (undefined === this.FootnotePr.NumFormat)
			return this.private_GetDocumentWideFootnotePr().NumFormat;
		
		return this.FootnotePr.NumFormat;
	};
	SectPr.prototype.private_GetDocumentWideFootnotePr = function()
	{
		return this.LogicDocument.Footnotes.FootnotePr;
	};
	/**
	 * Возвращаем настройки концевых сносок
	 * @return {AscWord.FootnotePr}
	 */
	SectPr.prototype.GetEndnotePr = function()
	{
		return this.EndnotePr;
	};
	SectPr.prototype.SetEndnotePos                    = function(nPos)
	{
		// Pos, заданная в секции не должна использоваться
		if (nPos !== this.EndnotePr.Pos)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionEndnotePos(this, this.EndnotePr.Pos, nPos));
			this.EndnotePr.Pos = nPos;
		}
	};
	SectPr.prototype.GetEndnotePos                    = function()
	{
		// Pos, заданная в секции не должна использоваться
		return this.EndnotePr.Pos;
	};
	SectPr.prototype.SetEndnoteNumStart               = function(nStart)
	{
		if (this.EndnotePr.NumStart !== nStart)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionEndnoteNumStart(this, this.EndnotePr.NumStart, nStart));
			this.EndnotePr.NumStart = nStart;
		}
	};
	SectPr.prototype.GetEndnoteNumStart               = function()
	{
		if (undefined === this.EndnotePr.NumStart)
			return 1;
		
		return this.EndnotePr.NumStart;
	};
	SectPr.prototype.SetEndnoteNumRestart             = function(nRestartType)
	{
		if (this.EndnotePr.NumRestart !== nRestartType)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionEndnoteNumRestart(this, this.EndnotePr.NumRestart, nRestartType));
			this.EndnotePr.NumRestart = nRestartType;
		}
	};
	SectPr.prototype.GetEndnoteNumRestart             = function()
	{
		if (undefined === this.EndnotePr.NumRestart)
			return section_footnote_RestartContinuous;
		
		return this.EndnotePr.NumRestart;
	};
	SectPr.prototype.SetEndnoteNumFormat              = function(nFormatType)
	{
		if (this.EndnotePr.NumFormat !== nFormatType)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionEndnoteNumFormat(this, this.EndnotePr.NumFormat, nFormatType));
			this.EndnotePr.NumFormat = nFormatType;
		}
	};
	SectPr.prototype.GetEndnoteNumFormat              = function()
	{
		if (undefined === this.EndnotePr.NumFormat)
			return Asc.c_oAscNumberingFormat.LowerRoman;
		
		return this.EndnotePr.NumFormat;
	};
	SectPr.prototype.private_GetDocumentWideEndnotePr = function()
	{
		return this.LogicDocument.Endnotes.EndnotePr;
	};
	SectPr.prototype.SetColumnProps                   = function(oColumnsProps)
	{
		var EqualWidth = oColumnsProps.get_EqualWidth();
		this.Set_Columns_EqualWidth(oColumnsProps.get_EqualWidth());
		if (false === EqualWidth)
		{
			var X      = 0;
			var XLimit = this.GetContentFrameWidth();
			
			var Cols          = [];
			var SectionColumn = null;
			var Count         = oColumnsProps.get_ColsCount();
			for (var Index = 0; Index < Count; ++Index)
			{
				var Col             = oColumnsProps.get_Col(Index);
				SectionColumn       = new AscWord.SectionColumn();
				SectionColumn.W     = Col.get_W();
				SectionColumn.Space = Col.get_Space();
				
				if (X + SectionColumn.W > XLimit)
				{
					SectionColumn.W = XLimit - X;
					Cols.push(SectionColumn);
					X += SectionColumn.W;
					break;
				}
				
				X += SectionColumn.W;
				if (Index != Count - 1)
					X += SectionColumn.Space;
				
				Cols.push(SectionColumn);
			}
			
			if (SectionColumn && X < XLimit - 0.001)
			{
				SectionColumn.W += XLimit - X;
			}
			
			this.Set_Columns_Cols(Cols);
			this.Set_Columns_Num(Count);
		}
		else
		{
			this.Set_Columns_Num(oColumnsProps.get_Num());
			this.Set_Columns_Space(oColumnsProps.get_Space());
		}
		
		this.Set_Columns_Sep(oColumnsProps.get_Sep());
	};
	SectPr.prototype.IsEqualColumnProps               = function(oColumnsProps)
	{
		if (oColumnsProps.get_Sep() !== this.Get_ColumnsSep() || oColumnsProps.get_EqualWidth() !== this.IsEqualColumnWidth())
			return false;
		
		if (this.IsEqualColumnWidth())
		{
			if (this.GetColumnCount() !== oColumnsProps.get_Num()
				|| Math.abs(this.GetColumnSpace() - oColumnsProps.get_Space()) > 0.01763)
				return false;
		}
		else
		{
			var nColumnsCount = oColumnsProps.get_ColsCount();
			if (nColumnsCount !== this.GetColumnCount())
				return false;
			
			for (var nIndex = 0; nIndex < nColumnsCount; ++nIndex)
			{
				var oCol = oColumnsProps.get_Col(nIndex);
				if (Math.abs(this.GetColumnWidth(nIndex) - oCol.get_W()) > 0.01763
					|| this.GetColumnSpace(nIndex) !== oCol.get_Space())
					return false;
			}
		}
		
		return true;
	};
	SectPr.prototype.SetGutter                        = function(nGutter)
	{
		if (Math.abs(nGutter - this.PageMargins.Gutter) > 0.001)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionPageMarginsGutter(this, this.PageMargins.Gutter, nGutter));
			this.PageMargins.Gutter = nGutter;
		}
	};
	SectPr.prototype.GetGutter                        = function()
	{
		return this.PageMargins.Gutter;
	};
	SectPr.prototype.SetGutterRTL                     = function(isRTL)
	{
		if (isRTL !== this.GutterRTL)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionGutterRTL(this, this.GutterRTL, isRTL));
			this.GutterRTL = isRTL;
		}
	};
	SectPr.prototype.IsGutterRTL                      = function()
	{
		return this.GutterRTL;
	};
	SectPr.prototype.SetPageMargins                   = function(_L, _T, _R, _B)
	{
		// Значения могут прийти как undefined, в этом случае мы поля со значением undefined не меняем
		var L = (undefined !== _L ? _L : this.PageMargins.Left);
		var T = (undefined !== _T ? _T : this.PageMargins.Top);
		var R = (undefined !== _R ? _R : this.PageMargins.Right);
		var B = (undefined !== _B ? _B : this.PageMargins.Bottom);
		
		if (Math.abs(L - this.PageMargins.Left) > 0.001 || Math.abs(T - this.PageMargins.Top) > 0.001 || Math.abs(R - this.PageMargins.Right) > 0.001 || Math.abs(B - this.PageMargins.Bottom) > 0.001)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionPageMargins(this, {
				L : this.PageMargins.Left,
				T : this.PageMargins.Top,
				R : this.PageMargins.Right,
				B : this.PageMargins.Bottom
			}, {L : L, T : T, R : R, B : B}));
			
			this.PageMargins.Left   = L;
			this.PageMargins.Top    = T;
			this.PageMargins.Right  = R;
			this.PageMargins.Bottom = B;
		}
	};
	SectPr.prototype.GetPageMarginLeft                = function()
	{
		return this.PageMargins.Left;
	};
	SectPr.prototype.GetPageMarginRight               = function()
	{
		return this.PageMargins.Right;
	};
	SectPr.prototype.GetPageMarginTop                 = function()
	{
		return this.PageMargins.Top;
	};
	SectPr.prototype.GetPageMarginBottom              = function()
	{
		return this.PageMargins.Bottom;
	};
	SectPr.prototype.SetPageSize                      = function(W, H)
	{
		if (Math.abs(W - this.PageSize.W) > 0.001 || Math.abs(H - this.PageSize.H) > 0.001)
		{
			H = Math.max(2.6, H);
			W = Math.max(12.7, W);
			
			AscCommon.History.Add(new AscDFH.CChangesSectionPageSize(this, {W : this.PageSize.W, H : this.PageSize.H}, {W : W, H : H}));
			
			this.PageSize.W = W;
			this.PageSize.H = H;
		}
	};
	SectPr.prototype.GetPageWidth                     = function()
	{
		return this.PageSize.W;
	};
	SectPr.prototype.GetPageHeight                    = function()
	{
		return this.PageSize.H;
	};
	SectPr.prototype.SetOrientation                   = function(Orient, ApplySize)
	{
		var _Orient = this.GetOrientation();
		if (_Orient !== Orient)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionPageOrient(this, this.PageSize.Orient, Orient));
			this.PageSize.Orient = Orient;
			
			if (true === ApplySize)
			{
				// При смене ориентации меняем местами высоту и ширину страницы и изменяем отступы
				
				var W = this.PageSize.W;
				var H = this.PageSize.H;
				
				var L = this.PageMargins.Left;
				var R = this.PageMargins.Right;
				var T = this.PageMargins.Top;
				var B = this.PageMargins.Bottom;
				
				this.SetPageSize(H, W);
				
				if (Asc.c_oAscPageOrientation.PagePortrait === Orient)
					this.SetPageMargins(T, R, B, L);
				else
					this.SetPageMargins(B, L, T, R);
			}
		}
	};
	SectPr.prototype.GetOrientation                   = function()
	{
		if (this.PageSize.W > this.PageSize.H)
			return Asc.c_oAscPageOrientation.PageLandscape;
		
		return Asc.c_oAscPageOrientation.PagePortrait;
	};
	SectPr.prototype.GetColumnCount                   = function()
	{
		return this.Columns.GetCount();
	};
	SectPr.prototype.GetColumnWidth                   = function(nColIndex)
	{
		return this.Columns.Get_ColumnWidth(nColIndex);
	};
	SectPr.prototype.GetMinColumnWidth                = function()
	{
		return this.Columns.GetMinColumnWidth();
	};
	SectPr.prototype.GetColumnSpace                   = function(nColIndex)
	{
		return this.Columns.Get_ColumnSpace(nColIndex);
	};
	SectPr.prototype.GetColumnSep                     = function()
	{
		return this.Columns.Sep;
	};
	SectPr.prototype.IsEqualColumnWidth               = function()
	{
		return this.Columns.EqualWidth;
	};
	SectPr.prototype.SetBordersOffsetFrom             = function(nOffsetFrom)
	{
		if (nOffsetFrom !== this.Borders.OffsetFrom)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionBordersOffsetFrom(this, this.Borders.OffsetFrom, nOffsetFrom));
			this.Borders.OffsetFrom = nOffsetFrom;
		}
	};
	SectPr.prototype.GetBordersOffsetFrom             = function()
	{
		return this.Borders.OffsetFrom;
	};
	SectPr.prototype.SetPageMarginHeader              = function(nHeader)
	{
		if (nHeader !== this.PageMargins.Header)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionPageMarginsHeader(this, this.PageMargins.Header, nHeader));
			this.PageMargins.Header = nHeader;
		}
	};
	SectPr.prototype.GetPageMarginHeader              = function()
	{
		return this.PageMargins.Header;
	};
	SectPr.prototype.SetPageMarginFooter              = function(nFooter)
	{
		if (nFooter !== this.PageMargins.Footer)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionPageMarginsFooter(this, this.PageMargins.Footer, nFooter));
			this.PageMargins.Footer = nFooter;
		}
	};
	SectPr.prototype.GetPageMarginFooter              = function()
	{
		return this.PageMargins.Footer;
	};
	/**
	 * Получаем границы для расположения содержимого документа на заданной страницы
	 * @param nPageAbs {number}
	 * @returns {{Left: number, Top: number, Right: number, Bottom: number}}
	 */
	SectPr.prototype.GetContentFrame = function(nPageAbs)
	{
		var nT = this.GetPageMarginTop();
		var nB = this.GetPageHeight() - this.GetPageMarginBottom();
		var nL = this.GetPageMarginLeft();
		var nR = this.GetPageWidth() - this.GetPageMarginRight();
		
		if (nT < 0)
			nT = -nT;
		
		if (this.LogicDocument && this.LogicDocument.IsMirrorMargins() && 1 === nPageAbs % 2)
		{
			nL = this.GetPageMarginRight();
			nR = this.GetPageWidth() - this.GetPageMarginLeft();
		}
		
		var nGutter = this.GetGutter();
		if (nGutter > 0.001)
		{
			if (this.LogicDocument && this.LogicDocument.IsGutterAtTop())
			{
				nT += nGutter;
			}
			else
			{
				if (this.LogicDocument && this.LogicDocument.IsMirrorMargins() && 1 === nPageAbs % 2)
				{
					if (this.IsGutterRTL())
						nL += nGutter;
					else
						nR -= nGutter;
				}
				else
				{
					if (this.IsGutterRTL())
						nR -= nGutter;
					else
						nL += nGutter;
				}
			}
		}
		
		return {
			Left   : nL,
			Top    : nT,
			Right  : nR,
			Bottom : nB
		};
	};
	/**
	 * Получаем ширину области для расположения содержимого документа
	 * @returns {number}
	 */
	SectPr.prototype.GetContentFrameWidth = function()
	{
		var nFrameWidth = this.GetPageWidth() - this.GetPageMarginLeft() - this.GetPageMarginRight();
		
		var nGutter = this.GetGutter();
		if (nGutter > 0.001 && !(this.LogicDocument && this.LogicDocument.IsGutterAtTop()))
			nFrameWidth -= nGutter;
		
		return nFrameWidth;
	};
	/**
	 * Получаем высоту области для расположения содержимого документа
	 * @returns {number}
	 */
	SectPr.prototype.GetContentFrameHeight = function()
	{
		var nFrameHeight = this.GetPageHeight() - this.GetPageMarginTop() - this.GetPageMarginBottom();
		
		var nGutter = this.GetGutter();
		if (nGutter > 0.001 && this.LogicDocument && this.LogicDocument.IsGutterAtTop())
			nFrameHeight -= nGutter;
		
		return nFrameHeight;
	};
	/**
	 * Есть ли нумерация строк
	 * @returns {boolean}
	 */
	SectPr.prototype.HaveLineNumbers = function()
	{
		return (undefined !== this.LnNumType && undefined !== this.LnNumType.CountBy && this.LnNumType.GetStart() >= 0);
	};
	/**
	 * Добавляем или меняем нумерацию строк
	 * @param nCountBy
	 * @param nDistance
	 * @param nStart
	 * @param nRestartType
	 */
	SectPr.prototype.SetLineNumbers = function(nCountBy, nDistance, nStart, nRestartType)
	{
		if (!this.HaveLineNumbers()
			|| nCountBy !== this.GetLineNumbersCountBy()
			|| nDistance !== this.GetLineNumbersDistance()
			|| nStart !== this.GetLineNumbersStart()
			|| nRestartType !== this.GetLineNumbersRestart()
		)
		{
			var oLnNumType = new AscWord.SectionLnNumType(nCountBy, nDistance, nStart, nRestartType);
			AscCommon.History.Add(new AscDFH.CChangesSectionLnNumType(this, this.LnNumType, oLnNumType));
			this.LnNumType = oLnNumType;
		}
	};
	/**
	 * Получаем класс с настройками нумерации строк
	 * @returns {?AscWord.SectionLnNumType}
	 */
	SectPr.prototype.GetLineNumbers = function()
	{
		if (this.HaveLineNumbers())
			return this.LnNumType;
		
		return undefined;
	};
	/**
	 * Убираем нумерацию строк
	 */
	SectPr.prototype.RemoveLineNumbers = function()
	{
		if (this.LnNumType)
		{
			AscCommon.History.Add(new AscDFH.CChangesSectionLnNumType(this, this.LnNumType, undefined));
			this.LnNumType = undefined;
		}
	};
	SectPr.prototype.GetLineNumbersCountBy  = function()
	{
		return (this.LnNumType && undefined !== this.LnNumType.CountBy ? this.LnNumType.CountBy : 0);
	};
	SectPr.prototype.GetLineNumbersStart    = function()
	{
		return (this.LnNumType && undefined !== this.LnNumType.GetStart() ? this.LnNumType.GetStart() : 0);
	};
	SectPr.prototype.GetLineNumbersRestart  = function()
	{
		return (this.LnNumType && undefined !== this.LnNumType.Restart ? this.LnNumType.Restart : Asc.c_oAscLineNumberRestartType.NewPage);
	};
	SectPr.prototype.GetLineNumbersDistance = function()
	{
		return (this.LnNumType ? this.LnNumType.Distance : undefined);
	};
	SectPr.prototype.RemoveHeader           = function(oHeader)
	{
		if (!oHeader)
			return;
		
		if (this.HeaderDefault === oHeader)
			this.Set_Header_Default(null);
		else if (this.HeaderEven === oHeader)
			this.Set_Header_Even(null);
		else if (this.HeaderFirst === oHeader)
			this.Set_Header_First(null);
		else if (this.FooterDefault === oHeader)
			this.Set_Footer_Default(null);
		else if (this.FooterEven === oHeader)
			this.Set_Footer_Even(null);
		else if (this.FooterFirst === oHeader)
			this.Set_Footer_First(null);
	};
	//------------------------------------------------------------------------------------------------------------------
	/**
	 * @constructor
	 */
	function SectionLayoutColumnInfo(X, XLimit)
	{
		this.X      = X;
		this.XLimit = XLimit;
		
		this.Pos    = 0;
		this.EndPos = 0;
	}
	/**
	 * @constructor
	 */
	function SectionLayoutInfo(X, Y, XLimit, YLimit)
	{
		this.X       = X;
		this.Y       = Y;
		this.XLimit  = XLimit;
		this.YLimit  = YLimit;
		this.Columns = [];
	}
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.SectPr = SectPr;
})();
