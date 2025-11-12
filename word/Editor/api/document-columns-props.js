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

(function(window)
{
	/**
	 * @constructor
	 */
	function CDocumentColumnsProps()
	{
		this.EqualWidth = true;
		this.Num        = 1;
		this.Sep        = false;
		this.Space      = 30;
		
		this.Cols       = [];
		
		this.TotalWidth = 230;
	}
	CDocumentColumnsProps.fromSectPr = function(sectPr)
	{
		let props = new CDocumentColumnsProps();
		if (!sectPr)
			return props;
		
		var Columns = sectPr.Columns;
		
		props.TotalWidth = sectPr.GetContentFrameWidth();
		props.EqualWidth = Columns.EqualWidth;
		props.Num        = Columns.Num;
		props.Sep        = Columns.Sep;
		props.Space      = Columns.Space;
		
		for (var Index = 0, Count = Columns.Cols.length; Index < Count; ++Index)
		{
			var Col = new Asc.CDocumentColumnProps();
			Col.put_W(Columns.Cols[Index].W);
			Col.put_Space(Columns.Cols[Index].Space);
			props.Cols[Index] = Col;
		}
		
		return props;
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
		var Col = new Asc.CDocumentColumnProps();
		Col.put_W(W);
		Col.put_Space(Space);
		this.Cols[Index] = Col;
	};
	CDocumentColumnsProps.prototype.get_TotalWidth = function()
	{
		return this.TotalWidth;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['Asc']['CDocumentColumnsProps'] = window['Asc'].CDocumentColumnsProps = CDocumentColumnsProps;
	
	CDocumentColumnsProps.prototype['get_EqualWidth'] = CDocumentColumnsProps.prototype.get_EqualWidth;
	CDocumentColumnsProps.prototype['put_EqualWidth'] = CDocumentColumnsProps.prototype.put_EqualWidth;
	CDocumentColumnsProps.prototype['get_Num']        = CDocumentColumnsProps.prototype.get_Num;
	CDocumentColumnsProps.prototype['put_Num']        = CDocumentColumnsProps.prototype.put_Num;
	CDocumentColumnsProps.prototype['get_Sep']        = CDocumentColumnsProps.prototype.get_Sep;
	CDocumentColumnsProps.prototype['put_Sep']        = CDocumentColumnsProps.prototype.put_Sep;
	CDocumentColumnsProps.prototype['get_Space']      = CDocumentColumnsProps.prototype.get_Space;
	CDocumentColumnsProps.prototype['put_Space']      = CDocumentColumnsProps.prototype.put_Space;
	CDocumentColumnsProps.prototype['get_ColsCount']  = CDocumentColumnsProps.prototype.get_ColsCount;
	CDocumentColumnsProps.prototype['get_Col']        = CDocumentColumnsProps.prototype.get_Col;
	CDocumentColumnsProps.prototype['put_Col']        = CDocumentColumnsProps.prototype.put_Col;
	CDocumentColumnsProps.prototype['put_ColByValue'] = CDocumentColumnsProps.prototype.put_ColByValue;
	CDocumentColumnsProps.prototype['get_TotalWidth'] = CDocumentColumnsProps.prototype.get_TotalWidth;
	
})(window);
