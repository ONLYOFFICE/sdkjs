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
	 * @constructor
	 */
	function SectionColumn(w, space)
	{
		this.W     = undefined !== w ? w : 0;
		this.Space = undefined !== space ? space :0;
	}
	
	SectionColumn.prototype.Write_ToBinary  = function(Writer)
	{
		// Double : W
		// Double : Space
		
		Writer.WriteDouble(this.W);
		Writer.WriteDouble(this.Space);
	};
	SectionColumn.prototype.Read_FromBinary = function(Reader)
	{
		// Double : W
		// Double : Space
		
		this.W     = Reader.GetDouble();
		this.Space = Reader.GetDouble();
	};
	
	
	/**
	 * @param {AscWord.SectPr} sectPr
	 * @constructor
	 */
	function SectionColumns(sectPr)
	{
		this.SectPr     = sectPr;
		this.EqualWidth = true;
		this.Num        = 1;
		this.Sep        = false;
		this.Space      = 30;
		
		this.Cols = [];
	}
	
	SectionColumns.prototype.Write_ToBinary              = function(Writer)
	{
		// Bool   : Equal
		// Long   : Num
		// Bool   : Sep
		// Double : Space
		
		// Long   : Количество колонок
		// Variable : Сами колонки
		
		Writer.WriteBool(this.EqualWidth);
		Writer.WriteLong(this.Num);
		Writer.WriteBool(this.Sep);
		Writer.WriteDouble(this.Space);
		
		var Count = this.Cols.length;
		Writer.WriteLong(Count);
		
		for (var Pos = 0; Pos < Count; Pos++)
		{
			this.Cols[Pos].Write_ToBinary(Writer);
		}
	};
	SectionColumns.prototype.Read_FromBinary             = function(Reader)
	{
		// Bool   : Equal
		// Long   : Num
		// Bool   : Sep
		// Double : Space
		
		// Long   : Количество колонок
		// Variable : Сами колонки
		
		this.EqualWidth = Reader.GetBool();
		this.Num        = Reader.GetLong();
		this.Sep        = Reader.GetBool();
		this.Space      = Reader.GetDouble();
		
		var Count = Reader.GetLong();
		this.Cols = [];
		
		for (var Pos = 0; Pos < Count; Pos++)
		{
			this.Cols[Pos] = new SectionColumn();
			this.Cols[Pos].Read_FromBinary(Reader);
		}
	};
	SectionColumns.prototype.GetCount                   = function()
	{
		if (true === this.EqualWidth)
		{
			return this.Num;
		}
		else
		{
			if (this.Cols.length <= 0)
				return 1;
			
			return this.Cols.length;
		}
	};
	SectionColumns.prototype.Get_ColumnWidth             = function(ColIndex)
	{
		if (true === this.EqualWidth)
		{
			return this.private_GetEqualColumnWidth();
		}
		else
		{
			if (this.Cols.length <= 0)
				return 0;
			
			ColIndex = Math.max(0, Math.min(this.Cols.length - 1, ColIndex));
			if (ColIndex < 0)
				return 0;
			
			return this.Cols[ColIndex].W;
		}
	};
	SectionColumns.prototype.Get_ColumnSpace             = function(ColIndex)
	{
		if (true === this.EqualWidth)
		{
			return this.Space;
		}
		else
		{
			if (this.Cols.length <= 0)
				return this.Space;
			
			ColIndex = Math.max(0, Math.min(this.Cols.length - 1, ColIndex));
			if (ColIndex < 0)
				return this.Space;
			
			return this.Cols[ColIndex].Space;
		}
	};
	SectionColumns.prototype.GetMinColumnWidth           = function()
	{
		if (true === this.EqualWidth)
		{
			return this.private_GetEqualColumnWidth();
		}
		else
		{
			if (this.Cols.length <= 0)
				return 0;
			
			let nWidth = this.Cols[0].W;
			for (let nColumn = 1, nColumnsCount = this.Cols.length; nColumn < nColumnsCount; ++nColumn)
			{
				if (this.Cols[nColumn].W < nWidth)
					nWidth = this.Cols[nColumn].W;
			}
			
			return nWidth;
		}
	};
	SectionColumns.prototype.private_GetEqualColumnWidth = function()
	{
		let nFrameW = this.SectPr.GetContentFrameWidth();
		return this.Num > 0 ? (nFrameW - this.Space * (this.Num - 1)) / this.Num : nFrameW;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.SectionColumn  = SectionColumn;
	AscWord.SectionColumns = SectionColumns;
})();
