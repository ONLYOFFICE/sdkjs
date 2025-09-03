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
	 * Line numbering settings
	 *
	 * @param nCountBy
	 * @param nDistance
	 * @param nStart
	 * @param nRestartType
	 * @constructor
	 */
	function SectionLnNumType(nCountBy, nDistance, nStart, nRestartType)
	{
		// Если задан сам класс, но в нем не задан CountBy, считаем, что нумерация строк не задана
		this.CountBy  = undefined !== nCountBy ? nCountBy : undefined;
		this.Distance = undefined !== nDistance && null !== nDistance ? nDistance : undefined; // В твипсах
		this.Start    = undefined !== nStart && 0 !== nStart ? nStart : undefined;
		this.Restart  = undefined !== nRestartType && Asc.c_oAscLineNumberRestartType.NewPage !== nRestartType ? nRestartType : undefined;
	}
	
	SectionLnNumType.prototype.Copy            = function()
	{
		return new SectionLnNumType(this.CountBy, this.Distance, this.Start, this.Restart);
	};
	SectionLnNumType.prototype.WriteToBinary   = function(oWriter)
	{
		var nStartPos = oWriter.GetCurPosition();
		oWriter.Skip(4);
		var nFlags = 0;
		
		if (undefined !== this.CountBy)
		{
			oWriter.WriteLong(this.CountBy);
			nFlags |= 1;
		}
		
		if (undefined !== this.Distance)
		{
			oWriter.WriteLong(this.Distance);
			nFlags |= 2;
		}
		
		if (undefined !== this.Start)
		{
			oWriter.WriteLong(this.Start);
			nFlags |= 4;
		}
		
		if (undefined !== this.Restart)
		{
			oWriter.WriteLong(this.Restart);
			nFlags |= 8;
		}
		
		var nEndPos = oWriter.GetCurPosition();
		oWriter.Seek(nStartPos);
		oWriter.WriteLong(nFlags);
		oWriter.Seek(nEndPos);
	};
	SectionLnNumType.prototype.ReadFromBinary  = function(oReader)
	{
		var nFlags = oReader.GetLong();
		
		if (nFlags & 1)
			this.CountBy = oReader.GetLong();
		else
			this.CountBy = undefined;
		
		if (nFlags & 2)
			this.Distance = oReader.GetLong();
		else
			this.Distance = undefined;
		
		if (nFlags & 4)
			this.Start = oReader.GetLong();
		else
			this.Start = undefined;
		
		if (nFlags & 8)
			this.Restart = oReader.GetLong();
		else
			this.Restart = undefined;
	};
	SectionLnNumType.prototype.Write_ToBinary  = function(oWriter)
	{
		this.WriteToBinary(oWriter);
	};
	SectionLnNumType.prototype.Read_FromBinary = function(oReader)
	{
		this.ReadFromBinary(oReader);
	};
	SectionLnNumType.prototype.SetCountBy      = function(nCountBy)
	{
		this.CountBy = nCountBy;
	};
	SectionLnNumType.prototype.GetCountBy      = function()
	{
		return this.CountBy;
	};
	SectionLnNumType.prototype.SetDistance     = function(nDistance)
	{
		this.Distance = nDistance;
	};
	SectionLnNumType.prototype.GetDistance     = function()
	{
		return this.Distance;
	};
	SectionLnNumType.prototype.SetStart        = function(nStart)
	{
		this.Start = nStart;
	};
	SectionLnNumType.prototype.GetStart        = function()
	{
		return undefined === this.Start ? 0 : this.Start;
	};
	SectionLnNumType.prototype.SetRestart      = function(nRestart)
	{
		this.Restart = nRestart;
	};
	SectionLnNumType.prototype.GetRestart      = function()
	{
		return (undefined === this.Restart ? Asc.c_oAscLineNumberRestartType.NewPage : this.Restart);
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.SectionLnNumType = SectionLnNumType;
	
	window['Asc']['CSectionLnNumType'] = window['Asc'].CSectionLnNumType = window['Asc'].SectionLnNumType = SectionLnNumType;
	SectionLnNumType.prototype["get_CountBy"]  = SectionLnNumType.prototype.GetCountBy;
	SectionLnNumType.prototype["put_CountBy"]  = SectionLnNumType.prototype.SetCountBy;
	SectionLnNumType.prototype["get_Distance"] = SectionLnNumType.prototype.GetDistance;
	SectionLnNumType.prototype["put_Distance"] = SectionLnNumType.prototype.SetDistance;
	SectionLnNumType.prototype["get_Start"]    = function()
	{
		return undefined === this.Start ? 1 : this.Start + 1;
	};
	SectionLnNumType.prototype["put_Start"]    = function(nStart)
	{
		this.Start = nStart - 1;
	};
	SectionLnNumType.prototype["get_Restart"]  = SectionLnNumType.prototype.GetRestart;
	SectionLnNumType.prototype["put_Restart"]  = SectionLnNumType.prototype.SetRestart;
})();
