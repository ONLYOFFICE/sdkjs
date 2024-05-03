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

// Import
var g_fontApplication = AscFonts.g_fontApplication;

var g_oTableId      = AscCommon.g_oTableId;
var g_oTextMeasurer = AscCommon.g_oTextMeasurer;
var isRealObject    = AscCommon.isRealObject;
var History         = AscCommon.History;

var HitInLine  = AscFormat.HitInLine;
var MOVE_DELTA = AscFormat.MOVE_DELTA;

var c_oAscRelativeFromH = Asc.c_oAscRelativeFromH;
var c_oAscRelativeFromV = Asc.c_oAscRelativeFromV;
var c_oAscSectionBreakType = Asc.c_oAscSectionBreakType;


var nbsp_charcode = 0x00A0;

var nbsp_string = String.fromCharCode(0x00A0);
var sp_string   = String.fromCharCode(0x0032);

// Suitable Run content for the paragraph simple changes
var g_oSRCFPSC             = [];
g_oSRCFPSC[para_Text]      = 1;
g_oSRCFPSC[para_Space]     = 1;
g_oSRCFPSC[para_End]       = 1;
g_oSRCFPSC[para_Tab]       = 1;
g_oSRCFPSC[para_Sym]       = 1;
g_oSRCFPSC[para_PageCount] = 1;
g_oSRCFPSC[para_FieldChar] = 1;
g_oSRCFPSC[para_InstrText] = 1;
g_oSRCFPSC[para_Bookmark]  = 1;

/**
 * Класс представляющий символ(текст) нумерации параграфа
 * @constructor
 * @extends {AscWord.CRunElementBase}
 */
function ParaNumbering()
{
	AscWord.CRunElementBase.call(this);

	this.Item = null; // Элемент в ране, к которому привязана нумерация
	this.Run  = null; // Ран, к которому привязана нумерация

	this.Line  = 0;
	this.Range = 0;
	this.Page  = 0;

	this.Internal = {
		FinalNumInfo    : undefined,
		FinalCalcValue  : -1,
		FinalNumId      : null,
		FinalNumLvl     : -1,

		SourceNumInfo   : undefined,
		SourceCalcValue : -1,
		SourceNumId     : null,
		SourceNumLvl    : -1,
		SourceWidth     : 0,

		Reset : function()
		{
			this.FinalNumInfo    = undefined;
			this.FinalCalcValue  = -1;
			this.FinalNumId      = null;
			this.FinalNumLvl     = -1;

			this.SourceNumInfo   = undefined;
			this.SourceCalcValue = -1;
			this.SourceNumId     = null;
			this.SourceNumLvl    = -1;
			this.SourceWidth     = 0;
		}
	};
}
ParaNumbering.prototype = Object.create(AscWord.CRunElementBase.prototype);
ParaNumbering.prototype.constructor = ParaNumbering;

ParaNumbering.prototype.Type = para_Numbering;
ParaNumbering.prototype.Draw = function(X, Y, oContext, oNumbering, oTextPr, oTheme, oPrevNumTextPr)
{
	var _X = X;
	if (this.Internal.SourceNumInfo)
	{
		oNumbering.Draw(this.Internal.SourceNumId,this.Internal.SourceNumLvl, _X, Y, oContext, this.Internal.SourceNumInfo, oPrevNumTextPr ? oPrevNumTextPr : oTextPr, oTheme);
		_X += this.Internal.SourceWidth;
	}

	if (this.Internal.FinalNumInfo)
	{
		oNumbering.Draw(this.Internal.FinalNumId,this.Internal.FinalNumLvl, _X, Y, oContext, this.Internal.FinalNumInfo, oTextPr, oTheme);
	}
};
ParaNumbering.prototype.Measure = function (oContext, oNumbering, oTextPr, oTheme, oFinalNumInfo, oFinalNumPr, oSourceNumInfo, oSourceNumPr)
{
	this.Width        = 0;
	this.Height       = 0;
	this.WidthVisible = 0;
	this.WidthNum     = 0;
	this.WidthSuff    = 0;

	this.Internal.Reset();

	if (!oNumbering)
	{
		return {
			Width        : this.Width,
			Height       : this.Height,
			WidthVisible : this.WidthVisible
		}
	}

	var nWidth = 0, nAscent = 0;
	if (oFinalNumInfo && oFinalNumPr && undefined !== oFinalNumInfo[oFinalNumPr.Lvl])
	{
		var oTemp = oNumbering.Measure(oFinalNumPr.NumId, oFinalNumPr.Lvl, oContext, oFinalNumInfo, oTextPr, oTheme);

		this.Internal.FinalNumInfo   = oFinalNumInfo;
		this.Internal.FinalCalcValue = oFinalNumInfo[oFinalNumPr.Lvl];
		this.Internal.FinalNumId     = oFinalNumPr.NumId;
		this.Internal.FinalNumLvl    = oFinalNumPr.Lvl;

		nWidth    = oTemp.Width;
		nAscent   = oTemp.Ascent;
	}

	if (oSourceNumInfo && oSourceNumPr && undefined !== oSourceNumInfo[oSourceNumPr.Lvl])
	{
		var oTemp = oNumbering.Measure(oSourceNumPr.NumId, oSourceNumPr.Lvl, oContext, oSourceNumInfo, oTextPr, oTheme);

		this.Internal.SourceNumInfo   = oSourceNumInfo;
		this.Internal.SourceCalcValue = oSourceNumInfo[oSourceNumPr.Lvl];
		this.Internal.SourceNumId     = oSourceNumPr.NumId;
		this.Internal.SourceNumLvl    = oSourceNumPr.Lvl;
		this.Internal.SourceWidth     = oTemp.Width;
		nWidth += this.Internal.SourceWidth;

		if (nAscent < oTemp.Ascent)
			nAscent = oTemp.Ascent;
	}

	this.Width        = nWidth;
	this.WidthVisible = nWidth;
	this.WidthNum     = nWidth;
	this.WidthSuff    = 0;
	this.Height       = nAscent; // Это не вся высота, а только высота над BaseLine
};
ParaNumbering.prototype.checkRange = function(range, line)
{
	return (this.Item
		&& this.Run
		&& range === this.Range
		&& line === this.Line
	);
};
ParaNumbering.prototype.CanAddNumbering = function()
{
	return false;
};
ParaNumbering.prototype.Copy = function()
{
	return new ParaNumbering();
};
ParaNumbering.prototype.Write_ToBinary = function(Writer)
{
	// Long   : Type
	Writer.WriteLong(this.Type);
};
ParaNumbering.prototype.Read_FromBinary = function(Reader)
{
};
ParaNumbering.prototype.GetCalculatedValue = function()
{
	return this.Internal.FinalCalcValue;
};
ParaNumbering.prototype.GetCalculatedNumInfo = function()
{
	return this.Internal.FinalNumInfo;
};
ParaNumbering.prototype.GetCalculatedNumberingLvl = function()
{
	return this.Internal.FinalNumLvl;
};
ParaNumbering.prototype.GetCalculatedNumId = function()
{
	return this.Internal.FinalNumId;
};
/**
 * Нужно ли отрисовывать исходную нумерацию
 * @returns {boolean}
 */
ParaNumbering.prototype.HaveSourceNumbering = function()
{
	return !!this.Internal.SourceNumInfo;
};
/**
 * Нужно ли отрисовывать финальную нумерацию
 * @returns {boolean}
 */
ParaNumbering.prototype.HaveFinalNumbering = function()
{
	return !!this.Internal.FinalNumInfo;
};
/**
 * Получаем ширину исходной нумерации
 * @returns {number}
 */
ParaNumbering.prototype.GetSourceWidth = function()
{
	return this.Internal.SourceWidth;
};
ParaNumbering.prototype.GetFontSlot = function(oTextPr)
{
	return AscWord.fontslot_Unknown;
};


/**
 * Класс представляющий символ нумерации у параграфа в презентациях
 * @constructor
 * @extends {AscWord.CRunElementBase}
 */
function ParaPresentationNumbering()
{
	AscWord.CRunElementBase.call(this);

    // Эти данные заполняются во время пересчета, перед вызовом Measure
    this.Bullet    = null;
    this.BulletNum = null;
}
ParaPresentationNumbering.prototype = Object.create(AscWord.CRunElementBase.prototype);
ParaPresentationNumbering.prototype.constructor = ParaPresentationNumbering;

ParaPresentationNumbering.prototype.Type = para_PresentationNumbering;
ParaPresentationNumbering.prototype.Draw = function(X, Y, Context, PDSE)
{
	this.Bullet.Draw(X, Y, Context, PDSE);
};
ParaPresentationNumbering.prototype.Measure = function (Context, FirstTextPr, Theme)
{
	this.Width        = 0;
	this.Height       = 0;
	this.WidthVisible = 0;

	var Temp = this.Bullet.Measure(Context, FirstTextPr, this.BulletNum, Theme);

	this.Width        = Temp.Width;
	this.WidthVisible = Temp.Width;
};
ParaPresentationNumbering.prototype.CanAddNumbering = function()
{
	return false;
};
ParaPresentationNumbering.prototype.Copy = function()
{
	return new ParaPresentationNumbering();
};
ParaPresentationNumbering.prototype.Write_ToBinary = function(Writer)
{
	// Long   : Type
	Writer.WriteLong(this.Type);
};
ParaPresentationNumbering.prototype.Read_FromBinary = function(Reader)
{
};
ParaPresentationNumbering.prototype.checkRange = function(Range, Line)
{
	if (null !== this.Item && null !== this.Run && Range === this.Range && Line === this.Line)
		return true;

	return false;
};
