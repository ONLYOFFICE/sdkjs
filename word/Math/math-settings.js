/*
 * (c) Copyright Ascensio System SIA 2010-2024
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
	function MathSettings()
	{
		this.Pr         = new MathPr();
		this.CompiledPr = new MathPr();
		this.DefaultPr  = new MathPr();
		
		this.DefaultPr.SetDefaultPr();
		
		this.bNeedCompile = true;
	}
	MathSettings.prototype.SetPr = function(Pr)
	{
		this.bNeedCompile = true;
		this.Pr.Merge(Pr);
		this.SetCompiledPr();
	};
	MathSettings.prototype.GetPr = function()
	{
		return this.Pr;
	};
	MathSettings.prototype.SetCompiledPr = function()
	{
		if (this.bNeedCompile)
		{
			this.CompiledPr.Merge(this.DefaultPr);
			this.CompiledPr.Merge(this.Pr);
			
			this.bNeedCompile = false;
		}
	};
	MathSettings.prototype.GetPrDispDef = function()
	{
		var Pr;
		if (this.CompiledPr.dispDef == false)
			Pr = this.DefaultPr;
		else
			Pr = this.CompiledPr;
		
		return Pr;
	};
	MathSettings.prototype.Get_WrapIndent = function(WrapState)
	{
		this.SetCompiledPr();
		
		var wrapIndent = 0;
		if (this.wrapRight == false && (WrapState == ALIGN_MARGIN_WRAP || WrapState == ALIGN_WRAP))
			wrapIndent = this.GetPrDispDef().wrapIndent;
		
		return wrapIndent;
	};
	MathSettings.prototype.Get_LeftMargin = function(WrapState)
	{
		this.SetCompiledPr();
		
		var lMargin = 0;
		if (WrapState == ALIGN_MARGIN_WRAP || WrapState == ALIGN_MARGIN)
			lMargin = this.GetPrDispDef().lMargin;
		
		return lMargin;
	};
	MathSettings.prototype.Get_RightMargin = function(WrapState)
	{
		this.SetCompiledPr();
		var rMargin = 0;
		if (WrapState == ALIGN_MARGIN_WRAP || WrapState == ALIGN_MARGIN)
			rMargin = this.GetPrDispDef().rMargin;
		
		return rMargin;
	};
	MathSettings.prototype.Get_IntLim = function()
	{
		this.SetCompiledPr();
		return this.CompiledPr.intLim;
	};
	MathSettings.prototype.Get_NaryLim = function()
	{
		this.SetCompiledPr();
		return this.CompiledPr.naryLim;
	};
	MathSettings.prototype.Get_DefJc = function()
	{
		this.SetCompiledPr();
		return this.GetPrDispDef().defJc;
	};
	MathSettings.prototype.Get_DispDef = function()
	{
		this.SetCompiledPr();
		return this.CompiledPr.dispDef;
	};
	MathSettings.prototype.Get_BrkBin = function()
	{
		this.SetCompiledPr();
		return this.CompiledPr.brkBin;
	};
	MathSettings.prototype.Get_WrapRight = function()
	{
		this.SetCompiledPr();
		return this.CompiledPr.wrapRight;
	};
	MathSettings.prototype.Get_SmallFrac = function()
	{
		this.SetCompiledPr();
		return this.CompiledPr.smallFrac;
	};
	MathSettings.prototype.Get_MenuProps = function()
	{
		return new CMathMenuSettings(this.CompiledPr);
	};
	MathSettings.prototype.Set_MenuProps = function(Props)
	{
		if (Props.BrkBin !== undefined)
		{
			this.Pr.brkBin = Props.BrkBin == c_oAscMathInterfaceSettingsBrkBin.BreakAfter ? BREAK_AFTER : BREAK_BEFORE;
		}
		
		if (Props.Justification !== undefined)
		{
			switch (Props.Justification)
			{
				case c_oAscMathInterfaceSettingsAlign.Justify:
				{
					this.Pr.defJc = AscCommon.align_Justify;
					break;
				}
				case c_oAscMathInterfaceSettingsAlign.Center:
				{
					this.Pr.defJc = AscCommon.align_Center;
					break;
				}
				case c_oAscMathInterfaceSettingsAlign.Left:
				{
					this.Pr.defJc = AscCommon.align_Left;
					break;
				}
				case c_oAscMathInterfaceSettingsAlign.Right:
				{
					this.Pr.defJc = AscCommon.align_Right;
					break;
				}
			}
		}
		
		if (Props.UseSettings !== undefined)
		{
			this.Pr.dispDef = Props.UseSettings;
		}
		
		if (Props.IntLim !== undefined)
		{
			if (Props.IntLim == Asc.c_oAscMathInterfaceNaryLimitLocation.SubSup)
			{
				this.Pr.intLim = NARY_SubSup;
			}
			else if (Props.IntLim == Asc.c_oAscMathInterfaceNaryLimitLocation.UndOvr)
			{
				this.Pr.intLim = NARY_UndOvr;
			}
		}
		
		if (Props.NaryLim !== undefined)
		{
			if (Props.NaryLim == Asc.c_oAscMathInterfaceNaryLimitLocation.SubSup)
			{
				this.Pr.naryLim = NARY_SubSup;
			}
			else if (Props.NaryLim == Asc.c_oAscMathInterfaceNaryLimitLocation.UndOvr)
			{
				this.Pr.naryLim = NARY_UndOvr;
			}
		}
		
		if (Props.LeftMargin !== undefined && Props.LeftMargin == Props.LeftMargin + 0)
		{
			this.Pr.lMargin = Props.LeftMargin;
		}
		
		if (Props.RightMargin !== undefined && Props.RightMargin == Props.RightMargin + 0)
		{
			this.Pr.rMargin = Props.RightMargin;
		}
		
		if (Props.WrapIndent !== undefined && Props.WrapIndent == Props.WrapIndent + 0)
		{
			this.Pr.wrapIndent = Props.WrapIndent;
		}
		
		if (Props.WrapRight !== undefined && Props.WrapRight !== null)
		{
			this.Pr.wrapRight = Props.WrapRight;
		}
		
		this.bNeedCompile = true;
		
	};
	MathSettings.prototype.Write_ToBinary = function(Writer)
	{
		this.Pr.Write_ToBinary(Writer);
	};
	MathSettings.prototype.Read_FromBinary = function(Reader)
	{
		this.Pr.Read_FromBinary(Reader);
		this.bNeedCompile = true;
	};
	
	/**
	 * @constructor
	 */
	function MathPr()
	{
		this.brkBin     = null;
		
		this.defJc      = null;
		this.dispDef    = null;  // свойство: применять/ не применять paragraph settings (в тч defJc)
		
		this.intLim     = null;
		this.naryLim    = null;
		
		this.lMargin    = null;
		this.rMargin    = null;
		this.wrapIndent = null;
		this.wrapRight  = null;
		
		this.smallFrac  = null;
		
		//   не реализовано    //
		
		// for minus operator
		// when brkBin is set to repeat
		this.brkBinSub  = null;
		
		//***** WORD IGNORES followings parameters *****//
		
		// mathFont: в качестве font поддерживается только Cambria Math
		// остальные шрифты  возможно будут поддержаны MS Word в будущем
		
		this.mathFont   = null;
		
		// Default font for math zones
		// Gives a drop-down list of math fonts that can be used as the default math font to be used in the document.
		// Currently only Cambria Math has thorough math support, but others such as the STIX fonts are coming soon.
		
		// http://blogs.msdn.com/b/murrays/archive/2008/10/27/default-document-math-properties.aspx
		
		
		// http://msdn.microsoft.com/en-us/library/ff529906(v=office.12).aspx
		// Word ignores the interSp attribute and fails to write it back out.
		this.interSp    = null;
		// http://msdn.microsoft.com/en-us/library/ff529301(v=office.12).aspx
		// Word does not implement this feature and does not write the intraSp element.
		this.intraSp    = null;
		
		// http://msdn.microsoft.com/en-us/library/ff533406(v=office.12).aspx
		this.postSp     = null;
		this.preSp      = null;
		
		// RichEdit Hot Keys
		// http://blogs.msdn.com/b/murrays/archive/2013/10/30/richedit-hot-keys.aspx
		
		//*********************//
	}
	MathPr.prototype.SetDefaultPr = function()
	{
		this.brkBin     = BREAK_BEFORE;
		this.defJc      = AscCommon.align_Justify;
		this.dispDef    = true;
		this.intLim     = NARY_SubSup;
		this.mathFont   = {Name  : "Cambria Math", Index : -1 };
		this.lMargin    = 0;
		this.naryLim    = NARY_UndOvr;
		this.rMargin    = 0;
		this.smallFrac  = false;
		this.wrapIndent = 25; // mm
		this.wrapRight  = false;
	};
	MathPr.prototype.Merge = function(Pr)
	{
		if (Pr.wrapIndent !== null && Pr.wrapIndent !== undefined)
			this.wrapIndent = Pr.wrapIndent;
		
		if (Pr.lMargin !== null && Pr.lMargin !== undefined)
			this.lMargin = Pr.lMargin;
		
		if (Pr.rMargin !== null && Pr.rMargin !== undefined)
			this.rMargin = Pr.rMargin;
		
		if (Pr.intLim !== null && Pr.intLim !== undefined)
			this.intLim = Pr.intLim;
		
		if (Pr.naryLim !== null && Pr.naryLim !== undefined)
			this.naryLim = Pr.naryLim;
		
		if (Pr.defJc !== null && Pr.defJc !== undefined)
			this.defJc = Pr.defJc;
		
		if (Pr.brkBin !== null && Pr.brkBin !== undefined)
			this.brkBin = Pr.brkBin;
		
		if (Pr.brkBinSub !== null && Pr.brkBinSub !== undefined)
			this.brkBinSub = Pr.brkBinSub;
		
		if (Pr.dispDef !== null && Pr.dispDef !== undefined)
			this.dispDef = Pr.dispDef;
		
		if (Pr.mathFont !== null && Pr.mathFont !== undefined)
			this.mathFont = Pr.mathFont;
		
		if (Pr.wrapRight !== null && Pr.wrapRight !== undefined)
			this.wrapRight = Pr.wrapRight;
		
		if (Pr.smallFrac !== null && Pr.smallFrac !== undefined)
			this.smallFrac = Pr.smallFrac;
	};
	MathPr.prototype.Copy = function()
	{
		var NewPr = new MathPr();
		
		NewPr.brkBin     = this.brkBin;
		NewPr.defJc      = this.defJc;
		NewPr.dispDef    = this.dispDef;
		NewPr.intLim     = this.intLim;
		NewPr.lMargin    = this.lMargin;
		NewPr.naryLim    = this.naryLim;
		NewPr.rMargin    = this.rMargin;
		NewPr.wrapIndent = this.wrapIndent;
		NewPr.brkBinSub  = this.brkBinSub;
		NewPr.interSp    = this.interSp;
		NewPr.intraSp    = this.intraSp;
		NewPr.mathFont   = this.mathFont;
		NewPr.postSp     = this.postSp;
		NewPr.preSp      = this.preSp;
		NewPr.smallFrac  = this.smallFrac;
		NewPr.wrapRight  = this.wrapRight;
		
		return NewPr;
	};
	MathPr.prototype.Write_ToBinary = function(Writer)
	{
		var StartPos = Writer.GetCurPosition();
		Writer.Skip(4);
		var Flags = 0;
		
		if (undefined !== this.brkBin)
		{
			Writer.WriteLong(this.brkBin);
			Flags |= 1;
		}
		
		if (undefined !== this.brkBinSub)
		{
			Writer.WriteLong(this.brkBinSub);
			Flags |= 2;
		}
		
		if (undefined !== this.defJc)
		{
			Writer.WriteLong(this.defJc);
			Flags |= 4;
		}
		
		if (undefined !== this.dispDef)
		{
			Writer.WriteBool(this.dispDef);
			Flags |= 8;
		}
		
		if (undefined !== this.interSp)
		{
			Writer.WriteLong(this.interSp);
			Flags |= 16;
		}
		
		if (undefined !== this.intLim)
		{
			Writer.WriteLong(this.intLim);
			Flags |= 32;
		}
		
		if (undefined !== this.intraSp)
		{
			Writer.WriteLong(this.intraSp);
			Flags |= 64;
		}
		
		if (undefined !== this.lMargin)
		{
			Writer.WriteLong(this.lMargin);
			Flags |= 128;
		}
		
		if (undefined !== this.mathFont)
		{
			Writer.WriteString2(this.mathFont.Name);
			Flags |= 256;
		}
		
		if (undefined !== this.naryLim)
		{
			Writer.WriteLong(this.naryLim);
			Flags |= 512;
		}
		
		if (undefined !== this.postSp)
		{
			Writer.WriteLong(this.postSp);
			Flags |= 1024;
		}
		
		if (undefined !== this.preSp)
		{
			Writer.WriteLong(this.preSp);
			Flags |= 2048;
		}
		
		if (undefined !== this.rMargin)
		{
			Writer.WriteLong(this.rMargin);
			Flags |= 4096;
		}
		
		if (undefined !== this.smallFrac)
		{
			Writer.WriteBool(this.smallFrac);
			Flags |= 8192;
		}
		
		if (undefined !== this.wrapIndent)
		{
			Writer.WriteLong(this.wrapIndent);
			Flags |= 16384;
		}
		
		if (undefined !== this.wrapRight)
		{
			Writer.WriteBool(this.wrapRight);
			Flags |= 32768;
		}
		
		var EndPos = Writer.GetCurPosition();
		Writer.Seek(StartPos);
		Writer.WriteLong(Flags);
		Writer.Seek(EndPos);
	};
	MathPr.prototype.Read_FromBinary = function(Reader)
	{
		var Flags = Reader.GetLong();
		
		if (Flags & 1)
			this.brkBin = Reader.GetLong();
		
		if (Flags & 2)
			this.brkBinSub = Reader.GetLong();
		
		if (Flags & 4)
			this.defJc = Reader.GetLong();
		
		if (Flags & 8)
			this.dispDef = Reader.GetBool();
		
		if (Flags & 16)
			this.interSp = Reader.GetLong();
		
		if (Flags & 32)
			this.intLim = Reader.GetLong();
		
		if (Flags & 64)
			this.intraSp = Reader.GetLong();
		
		if (Flags & 128)
			this.lMargin = Reader.GetLong();
		
		if (Flags & 256)
		{
			this.mathFont = {
				Name  : Reader.GetString2(),
				Index : -1
			};
		}
		
		if (Flags & 512)
			this.naryLim = Reader.GetLong();
		
		if (Flags & 1024)
			this.postSp = Reader.GetLong();
		
		if (Flags & 2048)
			this.preSp = Reader.GetLong();
		
		if (Flags & 4096)
			this.rMargin = Reader.GetLong();
		
		if (Flags & 8192)
			this.smallFrac = Reader.GetBool();
		
		if (Flags & 16384)
			this.wrapIndent = Reader.GetLong();
		
		if (Flags & 32768)
			this.wrapRight = Reader.GetBool();
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].MathSettings = MathSettings;
	window['AscWord'].MathPr       = MathPr;
	
})(window);

