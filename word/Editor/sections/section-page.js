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

(function()
{
	/**
	 * Page Size
	 * @constructor
	 */
	function SectionPageSize()
	{
		this.W      = 210;
		this.H      = 297;
		this.Orient = Asc.c_oAscPageOrientation.PagePortrait;
	}
	
	SectionPageSize.prototype.Write_ToBinary = function(Writer)
	{
		Writer.WriteDouble(this.W);
		Writer.WriteDouble(this.H);
		Writer.WriteByte(this.Orient);
	};
	SectionPageSize.prototype.Read_FromBinary = function(Reader)
	{
		this.W      = Reader.GetDouble();
		this.H      = Reader.GetDouble();
		this.Orient = Reader.GetByte();
	};
	
	/**
	 * Page Margins
	 * @constructor
	 */
	function SectionPageMargins()
	{
		this.Left   = 30; // 3 cm
		this.Top    = 20; // 2 cm
		this.Right  = 15; // 1.5 cm
		this.Bottom = 20; // 2 cm
		this.Gutter = 0;  // 0 cm
		
		this.Header = 12.5; // 1.25 cm
		this.Footer = 12.5; // 1.25 cm
	}
	
	SectionPageMargins.prototype.Write_ToBinary  = function(Writer)
	{
		Writer.WriteDouble(this.Left);
		Writer.WriteDouble(this.Top);
		Writer.WriteDouble(this.Right);
		Writer.WriteDouble(this.Bottom);
		Writer.WriteDouble(this.Header);
		Writer.WriteDouble(this.Footer);
		Writer.WriteDouble(this.Gutter);
	};
	SectionPageMargins.prototype.Read_FromBinary = function(Reader)
	{
		this.Left   = Reader.GetDouble();
		this.Top    = Reader.GetDouble();
		this.Right  = Reader.GetDouble();
		this.Bottom = Reader.GetDouble();
		this.Header = Reader.GetDouble();
		this.Footer = Reader.GetDouble();
		this.Gutter = Reader.GetDouble();
	};
	
	/**
	 * Page Borders
	 * @constructor
	 */
	function SectionPageBorders()
	{
		this.Top    = new CDocumentBorder();
		this.Bottom = new CDocumentBorder();
		this.Left   = new CDocumentBorder();
		this.Right  = new CDocumentBorder();
		
		this.Display    = section_borders_DisplayAllPages;
		this.OffsetFrom = section_borders_OffsetFromText;
		this.ZOrder     = section_borders_ZOrderFront;
	}
	
	SectionPageBorders.prototype.Write_ToBinary = function(Writer)
	{
		this.Left.Write_ToBinary(Writer);
		this.Top.Write_ToBinary(Writer);
		this.Right.Write_ToBinary(Writer);
		this.Bottom.Write_ToBinary(Writer);
		Writer.WriteByte(this.Display);
		Writer.WriteByte(this.OffsetFrom);
		Writer.WriteByte(this.ZOrder);
	};
	SectionPageBorders.prototype.Read_FromBinary = function(Reader)
	{
		this.Left.Read_FromBinary(Reader);
		this.Top.Read_FromBinary(Reader);
		this.Right.Read_FromBinary(Reader);
		this.Bottom.Read_FromBinary(Reader);
		
		this.Display    = Reader.GetByte();
		this.OffsetFrom = Reader.GetByte();
		this.ZOrder     = Reader.GetByte();
	};
	SectionPageBorders.prototype.IsEmptyBorders = function()
	{
		if (this.Top.IsNone() && this.Bottom.IsNone() && this.Left.IsNone() && this.Right.IsNone())
			return true;
		
		return false;
	};
	
	/**
	 *
	 * @constructor
	 */
	function SectionPageNumType()
	{
		this.Start     = -1;
		this.Format    = Asc.c_oAscNumberingFormat.Decimal;
		this.ChapStyle = undefined;
		this.ChapSep   = undefined;
	}
	
	SectionPageNumType.prototype.Write_ToBinary = function(writer)
	{
		writer.WriteLong(this.Start);
		writer.WriteLong(this.Format);
		
		let startPos = writer.GetCurPosition();
		writer.Skip(4);
		let flags = 0;
		
		if (undefined !== this.ChapStyle)
		{
			writer.WriteLong(this.ChapStyle);
			flags |= 1;
		}
		if (undefined !== this.ChapSep)
		{
			writer.WriteByte(this.ChapSep);
			flags |= 2;
		}
		
		let endPos = writer.GetCurPosition();
		writer.Seek(startPos);
		writer.WriteLong(flags);
		writer.Seek(endPos);
	};
	SectionPageNumType.prototype.Read_FromBinary = function(reader)
	{
		this.Start  = reader.GetLong();
		this.Format = reader.GetLong();
		
		let flags = reader.GetLong();
		if (flags & 1)
			this.ChapStyle = reader.GetLong();
		
		if (flags & 2)
			this.ChapSep = reader.GetByte();
	};
	
	/**
	 * Class for calculating actual page number
	 * @param firstPage
	 * @param currentPage
	 * @param isFirst
	 * @param isEven
	 * @param pageNum
	 * @param invalid
	 * @constructor
	 */
	function SectionPageNumInfo(firstPage, currentPage, isFirst, isEven, pageNum, invalid)
	{
		this.FirstPage = firstPage;
		this.CurPage   = currentPage;
		this.bFirst    = isFirst;
		this.bEven     = isEven;
		this.PageNum   = pageNum;
		this.Invalid   = invalid !== undefined ? invalid :false;
	}
	
	SectionPageNumInfo.prototype.Compare = function(info)
	{
		return (info
			&& !this.Invalid
			&& !info.Invalid
			&& this.CurPage === info.CurPage
			&& this.bFirst === info.bFirst
			&& this.bEven === info.bEven
			&& this.PageNum === info.PageNum);
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.SectionPageSize    = SectionPageSize;
	AscWord.SectionPageMargins = SectionPageMargins;
	AscWord.SectionPageBorders = SectionPageBorders;
	AscWord.SectionPageNumType = SectionPageNumType;
	AscWord.SectionPageNumInfo = SectionPageNumInfo;
})();
