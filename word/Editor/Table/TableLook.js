/*
 * (c) Copyright Ascensio System SIA 2010-2022
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

(function(window)
{
	/**
	 * @constructor
	 */
	function CTableLook(bFC, bFR, bLC, bLR, bBH, bBV)
	{
		this.FirstCol = (true === bFC);
		this.FirstRow = (true === bFR);
		this.LastCol  = (true === bLC);
		this.LastRow  = (true === bLR);
		this.BandHor  = (true === bBH);
		this.BandVer  = (true === bBV);
	}
	CTableLook.prototype.Set = function(bFC, bFR, bLC, bLR, bBH, bBV)
	{
		this.FirstCol = (true === bFC);
		this.FirstRow = (true === bFR);
		this.LastCol  = (true === bLC);
		this.LastRow  = (true === bLR);
		this.BandHor  = (true === bBH);
		this.BandVer  = (true === bBV);
	};
	CTableLook.prototype.Copy = function()
	{
		return new CTableLook(this.FirstCol, this.FirstRow, this.LastCol, this.LastRow, this.BandHor, this.BandVer);
	};
	CTableLook.prototype.IsEqual = function(oTableLook)
	{
		return (oTableLook
			&& oTableLook.FirstCol === this.FirstCol
			&& oTableLook.FirstRow === this.FirstRow
			&& oTableLook.LastCol === this.LastCol
			&& oTableLook.LastRow === this.LastRow
			&& oTableLook.BandHor === this.BandHor
			&& oTableLook.BandVer === this.BandVer);
	};
	CTableLook.prototype.IsFirstCol = function()
	{
		return this.FirstCol;
	};
	CTableLook.prototype.SetFirstCol = function(isFirstCol)
	{
		this.FirstCol = isFirstCol;
	};
	CTableLook.prototype.IsFirstRow = function()
	{
		return this.FirstRow;
	};
	CTableLook.prototype.SetFirstRow = function(isFirstRow)
	{
		this.FirstRow = isFirstRow;
	};
	CTableLook.prototype.IsLastCol = function()
	{
		return this.LastCol;
	};
	CTableLook.prototype.SetLastCol = function(isLastCol)
	{
		this.LastCol = isLastCol;
	};
	CTableLook.prototype.IsLastRow = function()
	{
		return this.LastRow;
	};
	CTableLook.prototype.SetLastRow = function(isLastRow)
	{
		this.LastRow = isLastRow;
	};
	CTableLook.prototype.IsBandHor = function()
	{
		return this.BandHor;
	};
	CTableLook.prototype.SetBandHor = function(isBandHor)
	{
		this.BandHor = isBandHor;
	};
	CTableLook.prototype.IsBandVer = function()
	{
		return this.BandVer;
	};
	CTableLook.prototype.SetBandVer = function(isBandVer)
	{
		this.BandVer = isBandVer;
	};
	CTableLook.prototype.Write_ToBinary = function(Writer)
	{
		Writer.WriteBool(this.FirstCol);
		Writer.WriteBool(this.FirstRow);
		Writer.WriteBool(this.LastCol);
		Writer.WriteBool(this.LastRow);
		Writer.WriteBool(this.BandHor);
		Writer.WriteBool(this.BandVer);
	};
	CTableLook.prototype.Read_FromBinary = function(Reader)
	{
		this.FirstCol = Reader.GetBool();
		this.FirstRow = Reader.GetBool();
		this.LastCol  = Reader.GetBool();
		this.LastRow  = Reader.GetBool();
		this.BandHor  = Reader.GetBool();
		this.BandVer  = Reader.GetBool();
	};
	CTableLook.prototype.SetDefault = function()
	{
		this.Set(true, true, false, false, true, false);
	};

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon']            = window['AscCommon'] || {};
	window['AscCommon'].CTableLook = CTableLook;

	//--------------------------------------------------------export for interface--------------------------------------
	window['Asc']                   = window['Asc'] || {};
	window['Asc']['CTablePropLook'] = window['Asc'].CTablePropLook = CTableLook;

	CTableLook.prototype['get_FirstCol'] = CTableLook.prototype.get_FirstCol = CTableLook.prototype.IsFirstCol;
	CTableLook.prototype['put_FirstCol'] = CTableLook.prototype.put_FirstCol = CTableLook.prototype.SetFirstCol;
	CTableLook.prototype['get_FirstRow'] = CTableLook.prototype.get_FirstRow = CTableLook.prototype.IsFirstRow;
	CTableLook.prototype['put_FirstRow'] = CTableLook.prototype.put_FirstRow = CTableLook.prototype.SetFirstRow;
	CTableLook.prototype['get_LastCol']  = CTableLook.prototype.get_LastCol = CTableLook.prototype.IsLastCol;
	CTableLook.prototype['put_LastCol']  = CTableLook.prototype.put_LastCol = CTableLook.prototype.SetLastCol;
	CTableLook.prototype['get_LastRow']  = CTableLook.prototype.get_LastRow = CTableLook.prototype.IsLastRow;
	CTableLook.prototype['put_LastRow']  = CTableLook.prototype.put_LastRow = CTableLook.prototype.SetLastRow;
	CTableLook.prototype['get_BandHor']  = CTableLook.prototype.get_BandHor = CTableLook.prototype.IsBandHor;
	CTableLook.prototype['put_BandHor']  = CTableLook.prototype.put_BandHor = CTableLook.prototype.SetBandHor;
	CTableLook.prototype['get_BandVer']  = CTableLook.prototype.get_BandVer = CTableLook.prototype.IsBandVer;
	CTableLook.prototype['put_BandVer']  = CTableLook.prototype.put_BandVer = CTableLook.prototype.SetBandVer;

})(window);
