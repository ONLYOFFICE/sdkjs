/*
 * (c) Copyright Ascensio System SIA 2010-2026
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
	 * Class representing paragraph spacing
	 * @constructor
	 */
	function ParaSpacing()
	{
		this.Line              = undefined; // Расстояние между строками внутри абзаца
		this.LineRule          = undefined; // Тип расстрояния между строками
		this.Before            = undefined; // Дополнительное расстояние до абзаца
		this.BeforePct         = undefined; // Расстояние до абзаца в процентах от высоты строки
		this.BeforeAutoSpacing = undefined; // Использовать ли автоматический расчет расстояния до параграфа
		this.After             = undefined; // Дополнительное расстояние после абзаца
		this.AfterPct          = undefined; // Расстояние после абзаца в процентах от высоты строки
		this.AfterAutoSpacing  = undefined; // Использовать ли автоматический расчет расстояния после параграфа
		this.BeforeLines       = undefined; // Расстояние до абзаца в строках
		this.AfterLines        = undefined; // Расстояние после абзаца в строках
	}

	ParaSpacing.prototype.Copy = function()
	{
		var Spacing = new ParaSpacing();
		Spacing.Line              = this.Line;
		Spacing.LineRule          = this.LineRule;
		Spacing.Before            = this.Before;
		Spacing.BeforeAutoSpacing = this.BeforeAutoSpacing;
		Spacing.After             = this.After;
		Spacing.AfterAutoSpacing  = this.AfterAutoSpacing;
		Spacing.BeforePct         = this.BeforePct;
		Spacing.AfterPct          = this.AfterPct;
		Spacing.BeforeLines       = this.BeforeLines;
		Spacing.AfterLines        = this.AfterLines;
		return Spacing;
	};
	ParaSpacing.prototype.Merge = function(Spacing)
	{
		if ( undefined != Spacing.Line )
			this.Line = Spacing.Line;

		if ( undefined != Spacing.LineRule )
			this.LineRule = Spacing.LineRule;

		if ( undefined != Spacing.Before )
		{
			this.Before = Spacing.Before;
			this.BeforePct = undefined;
		}

		if ( undefined != Spacing.BeforeAutoSpacing )
			this.BeforeAutoSpacing = Spacing.BeforeAutoSpacing;

		if ( undefined != Spacing.After )
		{
			this.After = Spacing.After;
			this.AfterPct = undefined;
		}

		if ( undefined != Spacing.AfterAutoSpacing )
			this.AfterAutoSpacing = Spacing.AfterAutoSpacing;

		if ( undefined != Spacing.BeforePct )
		{
			this.BeforePct = Spacing.BeforePct;
			this.Before = 0;
		}

		if ( undefined != Spacing.AfterPct )
		{
			this.AfterPct = Spacing.AfterPct;
			this.After = 0;
		}

		if (undefined !== Spacing.BeforeLines)
			this.BeforeLines = Spacing.BeforeLines;

		if (undefined !== Spacing.AfterLines)
			this.AfterLines = Spacing.AfterLines;
	};
	ParaSpacing.prototype.Is_Equal = function(Spacing)
	{
		return this.IsEqual(Spacing);
	};
	ParaSpacing.prototype.Set_FromObject = function(Spacing)
	{
		this.Line              = Spacing.Line;
		this.LineRule          = Spacing.LineRule;
		this.Before            = Spacing.Before;
		this.BeforeAutoSpacing = Spacing.BeforeAutoSpacing;
		this.After             = Spacing.After;
		this.AfterAutoSpacing  = Spacing.AfterAutoSpacing;
		this.BeforePct         = Spacing.BeforePct;
		this.AfterPct          = Spacing.AfterPct;
		this.BeforeLines       = Spacing.BeforeLines;
		this.AfterLines        = Spacing.AfterLines;
	};
	ParaSpacing.prototype.Write_ToBinary = function(Writer)
	{
		var StartPos = Writer.GetCurPosition();
		Writer.Skip(4);
		var Flags = 0;

		if ( undefined != this.Line )
		{
			Writer.WriteDouble( this.Line );
			Flags |= 1;
		}

		if ( undefined != this.LineRule )
		{
			Writer.WriteByte( this.LineRule );
			Flags |= 2;
		}

		if ( undefined != this.Before )
		{
			Writer.WriteDouble( this.Before );
			Flags |= 4;
		}

		if ( undefined != this.After )
		{
			Writer.WriteDouble( this.After );
			Flags |= 8;
		}

		if ( undefined != this.AfterAutoSpacing )
		{
			Writer.WriteBool( this.AfterAutoSpacing );
			Flags |= 16;
		}

		if ( undefined != this.BeforeAutoSpacing )
		{
			Writer.WriteBool( this.BeforeAutoSpacing );
			Flags |= 32;
		}

		if ( undefined != this.BeforePct )
		{
			Writer.WriteLong( this.BeforePct );
			Flags |= 64;
		}

		if ( undefined != this.AfterPct )
		{
			Writer.WriteLong( this.AfterPct );
			Flags |= 128;
		}

		if (undefined !== this.BeforeLines)
		{
			Writer.WriteLong(this.BeforeLines);
			Flags |= 256;
		}

		if (undefined !== this.AfterLines)
		{
			Writer.WriteLong(this.AfterLines);
			Flags |= 512;
		}

		var EndPos = Writer.GetCurPosition();
		Writer.Seek( StartPos );
		Writer.WriteLong( Flags );
		Writer.Seek( EndPos );
	};
	ParaSpacing.prototype.Read_FromBinary = function(Reader)
	{
		var Flags = Reader.GetLong();

		if ( Flags & 1 )
			this.Line = Reader.GetDouble();

		if ( Flags & 2 )
			this.LineRule = Reader.GetByte();

		if ( Flags & 4 )
			this.Before = Reader.GetDouble();

		if ( Flags & 8 )
			this.After = Reader.GetDouble();

		if ( Flags & 16 )
			this.AfterAutoSpacing = Reader.GetBool();

		if ( Flags & 32 )
			this.BeforeAutoSpacing = Reader.GetBool();

		if ( Flags & 64 )
			this.BeforePct = Reader.GetLong();

		if ( Flags & 128 )
			this.AfterPct = Reader.GetLong();

		if (Flags & 256)
			this.BeforeLines = Reader.GetLong();

		if (Flags & 512)
			this.AfterLines = Reader.GetLong();
	};
	ParaSpacing.prototype.Get_Diff = function(Spacing)
	{
		var DiffSpacing = new ParaSpacing();

		if (this.Line !== Spacing.Line)
			DiffSpacing.Line = this.Line;

		if (this.LineRule !== Spacing.LineRule)
			DiffSpacing.LineRule = this.LineRule;

		if (this.Before !== Spacing.Before)
			DiffSpacing.Before = this.Before;

		if (this.BeforeAutoSpacing !== Spacing.BeforeAutoSpacing)
			DiffSpacing.BeforeAutoSpacing = this.BeforeAutoSpacing;

		if (this.After !== Spacing.After)
			DiffSpacing.After = this.After;

		if (this.AfterAutoSpacing !== Spacing.AfterAutoSpacing)
			DiffSpacing.AfterAutoSpacing = this.AfterAutoSpacing;

		if (this.BeforePct !== Spacing.BeforePct)
			DiffSpacing.BeforePct = this.BeforePct;

		if (this.AfterPct !== Spacing.AfterPct)
			DiffSpacing.AfterPct = this.AfterPct;

		if (this.BeforeLines !== Spacing.BeforeLines)
			DiffSpacing.BeforeLines = this.BeforeLines;

		if (this.AfterLines !== Spacing.AfterLines)
			DiffSpacing.AfterLines = this.AfterLines;

		return DiffSpacing;
	};
	ParaSpacing.prototype.Is_Empty = function()
	{
		return (undefined === this.Line
			&& undefined === this.LineRule
			&& undefined === this.Before
			&& undefined === this.BeforeAutoSpacing
			&& undefined === this.After
			&& undefined === this.AfterAutoSpacing
			&& undefined === this.BeforePct
			&& undefined === this.AfterPct
			&& undefined === this.BeforeLines
			&& undefined === this.AfterLines);
	};
	ParaSpacing.prototype.IsEmpty = function()
	{
		return this.Is_Empty();
	};
	ParaSpacing.prototype.IsEqual = function(oSpacing)
	{
		return (this.Line === oSpacing.Line
			&& this.LineRule === oSpacing.LineRule
			&& IsEqualNullableFloatNumbers(this.Before, oSpacing.Before)
			&& IsEqualNullableFloatNumbers(this.After, oSpacing.After)
			&& IsEqualNullableFloatNumbers(this.AfterPct, oSpacing.AfterPct)
			&& IsEqualNullableFloatNumbers(this.BeforePct, oSpacing.BeforePct)
			&& this.BeforeAutoSpacing === oSpacing.BeforeAutoSpacing
			&& this.AfterAutoSpacing === oSpacing.AfterAutoSpacing
			&& this.BeforeLines === oSpacing.BeforeLines
			&& this.AfterLines === oSpacing.AfterLines);
	};
	ParaSpacing.prototype.SetLineTwips = function(val)
	{
		if (null === val || undefined === val)
			return;
		
		if (val < 0)
		{
			val           = Math.abs(val);
			this.LineRule = Asc.linerule_Exact;
		}
		
		if (Asc.linerule_Auto === this.LineRule)
			this.Line = val / 240;
		else
			this.Line = g_dKoef_twips_to_mm * val;
	};
	ParaSpacing.prototype.CalculateBefore = function()
	{
		if (true === this.BeforeAutoSpacing)
			return 14 * g_dKoef_pt_to_mm;
		else if (undefined !== this.BeforeLines && null !== this.BeforeLines && 0 !== this.BeforeLines)
			return this.BeforeLines * 240 / 100 * g_dKoef_twips_to_mm;
		
		return this.Before;
	};
	ParaSpacing.prototype.CalculateAfter = function()
	{
		if (true === this.AfterAutoSpacing)
			return 14 * g_dKoef_pt_to_mm;
		else if (undefined !== this.AfterLines && null !== this.AfterLines && 0 !== this.AfterLines)
			return this.AfterLines * 240 / 100 * g_dKoef_twips_to_mm;
		
		return this.After;
	};
	//--------------------------------------------------------export----------------------------------------------------

	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].ParaSpacing = ParaSpacing;

})(window);
