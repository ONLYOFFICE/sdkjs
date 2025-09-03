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
	function FootnotePr()
	{
		this.NumRestart = undefined;
		this.NumFormat  = undefined;
		this.NumStart   = undefined;
		this.Pos        = undefined;
	}
	
	FootnotePr.prototype.InitDefault          = function()
	{
		this.NumFormat  = Asc.c_oAscNumberingFormat.Decimal;
		this.NumRestart = section_footnote_RestartContinuous;
		this.NumStart   = 1;
		this.Pos        = Asc.c_oAscFootnotePos && Asc.c_oAscFootnotePos.PageBottom;
	};
	FootnotePr.prototype.InitDefaultEndnotePr = function()
	{
		this.NumFormat  = Asc.c_oAscNumberingFormat.LowerRoman;
		this.NumRestart = section_footnote_RestartContinuous;
		this.NumStart   = 1;
		this.Pos        = Asc.c_oAscEndnotePos && Asc.c_oAscEndnotePos.DocEnd;
	};
	FootnotePr.prototype.WriteToBinary        = function(Writer)
	{
		var StartPos = Writer.GetCurPosition();
		Writer.Skip(4);
		var Flags = 0;
		
		if (undefined !== this.NumFormat)
		{
			Writer.WriteLong(this.NumFormat);
			Flags |= 1;
		}
		
		if (undefined !== this.NumRestart)
		{
			Writer.WriteLong(this.NumRestart);
			Flags |= 2;
		}
		
		if (undefined !== this.NumStart)
		{
			Writer.WriteLong(this.NumStart);
			Flags |= 4;
		}
		
		if (undefined !== this.Pos)
		{
			Writer.WriteLong(this.Pos);
			Flags |= 8;
		}
		
		var EndPos = Writer.GetCurPosition();
		Writer.Seek(StartPos);
		Writer.WriteLong(Flags);
		Writer.Seek(EndPos);
	};
	FootnotePr.prototype.ReadFromBinary       = function(Reader)
	{
		var Flags = Reader.GetLong();
		
		if (Flags & 1)
			this.NumFormat = Reader.GetLong();
		else
			this.NumFormat = undefined;
		
		if (Flags & 2)
			this.NumRestart = Reader.GetLong();
		else
			this.NumRestart = undefined;
		
		if (Flags & 4)
			this.NumStart = Reader.GetLong();
		else
			this.NumStart = undefined;
		
		if (Flags & 8)
			this.Pos = Reader.GetLong();
		else
			this.Pos = undefined;
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.FootnotePr = FootnotePr;
})();
