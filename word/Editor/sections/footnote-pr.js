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
