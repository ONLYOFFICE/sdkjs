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

/**
 * Класс представляющий собой настройки текста (сейчас используется как настройка текста для конца параграфа)
 * @param oProps
 * @constructor
 * @extends {AscWord.CRunElementBase}
 */
function ParaTextPr(oProps)
{
	AscWord.CRunElementBase.call(this);

	this.Id = AscCommon.g_oIdCounter.Get_NewId();

	this.Type      = para_TextPr;
	this.Value     = new CTextPr();
	this.Parent    = null;
	this.CalcValue = this.Value;

	this.Width        = 0;
	this.Height       = 0;
	this.WidthVisible = 0;

	if (oProps)
		this.Value.Set_FromObject(oProps);

	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
	AscCommon.g_oTableId.Add(this, this.Id);
}
ParaTextPr.prototype = Object.create(AscWord.CRunElementBase.prototype);
ParaTextPr.prototype.constructor = ParaTextPr;

ParaTextPr.prototype.Type = para_TextPr;
ParaTextPr.prototype.Get_Type = function()
{
	return this.Type;
};
ParaTextPr.prototype.Copy = function()
{
	var ParaTextPr_new = new ParaTextPr();
	ParaTextPr_new.Set_Value(this.Value);
	return ParaTextPr_new;
};
ParaTextPr.prototype.CanAddNumbering = function()
{
	return false;
};
ParaTextPr.prototype.Get_Id = function()
{
	return this.Id;
};
ParaTextPr.prototype.GetParagraph = function()
{
	if (this.Parent instanceof Paragraph)
		return this.Parent;

	return null;
};
ParaTextPr.prototype.IsParagraphSimpleChanges = function()
{
	return true;
};
ParaTextPr.prototype.GetCompiledPr = function()
{
	let textPr;
	if (!this.Parent || !this.Parent.Get_CompiledPr2)
	{
		textPr = new CTextPr();
		textPr.InitDefault();
	}
	else
	{
		textPr = this.Parent.Get_CompiledPr2(false).TextPr.Copy();
		let layoutCoeff = this.Parent.getLayoutFontSizeCoefficient();
		textPr.FontSize   *= layoutCoeff;
		textPr.FontSizeCS *= layoutCoeff;
	}
	
	textPr.Merge(this.Value);
	return textPr;
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для изменения свойств
//----------------------------------------------------------------------------------------------------------------------
ParaTextPr.prototype.Apply_TextPr = function(TextPr)
{
	if (undefined !== TextPr.Bold)
	{
		let _bold = null === TextPr.Bold ? undefined : TextPr.Bold;
		this.SetBold(_bold);
		this.SetBoldCS(_bold);
	}

	if (undefined !== TextPr.Italic)
	{
		let _italic = null === TextPr.Italic ? undefined : TextPr.Italic;
		this.SetItalic(_italic);
		this.SetItalicCS(_italic);
	}

	if (undefined !== TextPr.Strikeout)
		this.SetStrikeout(TextPr.Strikeout);

	if (undefined !== TextPr.Underline)
		this.SetUnderline(TextPr.Underline);

	if (undefined !== TextPr.FontSize)
	{
		let fontSize = null === TextPr.FontSize ? undefined : TextPr.FontSize;
		this.SetFontSize(fontSize);
		this.SetFontSizeCS(fontSize);
	}

	if (undefined !== TextPr.Color)
	{
		this.Set_Color(TextPr.Color);

		if (undefined !== this.Value.Unifill)
		{
			this.Set_Unifill(undefined);
		}

		if (undefined !== this.Value.TextFill)
		{
			this.Set_TextFill(undefined);
		}
	}

	if (undefined !== TextPr.VertAlign)
		this.Set_VertAlign(TextPr.VertAlign);

	if (undefined !== TextPr.HighLight)
		this.Set_HighLight(TextPr.HighLight);

	if (undefined !== TextPr.HighlightColor)
		this.SetHighlightColor(TextPr.HighlightColor);

	if (undefined !== TextPr.RStyle)
		this.Set_RStyle(TextPr.RStyle);

	if (undefined !== TextPr.Spacing)
		this.Set_Spacing(TextPr.Spacing);

	if (undefined !== TextPr.DStrikeout)
		this.Set_DStrikeout(TextPr.DStrikeout);

	if (undefined !== TextPr.Caps)
		this.Set_Caps(TextPr.Caps);

	if (undefined !== TextPr.SmallCaps)
		this.Set_SmallCaps(TextPr.SmallCaps);

	if (undefined !== TextPr.Position)
		this.Set_Position(TextPr.Position);

	if (TextPr.RFonts)
	{
		if (TextPr.FontFamily)
			this.ApplyFontFamily(TextPr.FontFamily.Name);
		else
			this.Set_RFonts2(TextPr.RFonts);
	}

	if (undefined != TextPr.Lang)
		this.Set_Lang(TextPr.Lang);

	if (undefined != TextPr.Unifill)
	{
		this.Set_Unifill(TextPr.Unifill.createDuplicate());
		if (undefined != this.Value.Color)
		{
			this.Set_Color(undefined);
		}
		if (undefined != this.Value.TextFill)
		{
			this.Set_TextFill(undefined);
		}
	}
	if (undefined != TextPr.TextOutline)
	{
		this.Set_TextOutline(TextPr.TextOutline);
	}
	if (undefined != TextPr.TextFill)
	{
		this.Set_TextFill(TextPr.TextFill);
		if (undefined != this.Value.Color)
		{
			this.Set_Color(undefined);
		}
		if (undefined != this.Value.Unifill)
		{
			this.Set_Unifill(undefined);
		}
	}

	if (undefined !== TextPr.Ligatures)
		this.SetLigatures(TextPr.Ligatures);
};
ParaTextPr.prototype.Clear_Style = function()
{
	this.SetBold(undefined);
	this.SetBoldCS(undefined);
	this.SetItalic(undefined);
	this.SetItalicCS(undefined);
	this.SetStrikeout(undefined);
	this.SetUnderline(undefined);
	this.SetFontSize(undefined);
	this.SetFontSizeCS(undefined);

	if (undefined != this.Value.Color)
		this.Set_Color(undefined);

	if (undefined != this.Value.Unifill)
		this.Set_Unifill(undefined);

	if (undefined != this.Value.VertAlign)
		this.Set_VertAlign(undefined);

	if (undefined != this.Value.HighLight)
		this.Set_HighLight(undefined);

	if (undefined != this.Value.HighlightColor)
		this.SetHighlightColor(undefined);

	if (undefined != this.Value.RStyle)
		this.Set_RStyle(undefined);

	if (undefined != this.Value.Spacing)
		this.Set_Spacing(undefined);

	if (undefined != this.Value.DStrikeout)
		this.Set_DStrikeout(undefined);

	if (undefined != this.Value.Caps)
		this.Set_Caps(undefined);

	if (undefined != this.Value.SmallCaps)
		this.Set_SmallCaps(undefined);

	if (undefined != this.Value.Position)
		this.Set_Position(undefined);

	this.SetRFontsAscii(undefined);
	this.SetRFontsHAnsi(undefined);
	this.SetRFontsCS(undefined);
	this.SetRFontsEastAsia(undefined);
	this.SetRFontsHint(undefined);

	if (undefined != this.Value.TextFill)
		this.Set_TextFill(undefined);

	if (undefined != this.Value.TextOutline)
		this.Set_TextOutline(undefined);
};
ParaTextPr.prototype.SetBold = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.Bold === Value)
		return;

	History.Add(new CChangesParaTextPrBold(this, this.Value.Bold, Value));
	this.Value.Bold = Value;
};
ParaTextPr.prototype.SetItalic = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.Italic === Value)
		return;

	History.Add(new CChangesParaTextPrItalic(this, this.Value.Italic, Value));
	this.Value.Italic = Value;
};
ParaTextPr.prototype.SetStrikeout = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.Strikeout === Value)
		return;

	History.Add(new CChangesParaTextPrStrikeout(this, this.Value.Strikeout, Value));
	this.Value.Strikeout = Value;
};
ParaTextPr.prototype.SetUnderline = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.Underline === Value)
		return;

	History.Add(new CChangesParaTextPrUnderline(this, this.Value.Underline, Value));
	this.Value.Underline = Value;
};
ParaTextPr.prototype.SetFontSize = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.FontSize === Value)
		return;

	History.Add(new CChangesParaTextPrFontSize(this, this.Value.FontSize, Value));
	this.Value.FontSize = Value;
};
ParaTextPr.prototype.Set_Color = function(Value)
{
	if (null === Value)
		Value = undefined;

	History.Add(new CChangesParaTextPrColor(this, this.Value.Color, Value));
	this.Value.Color = Value;
};
ParaTextPr.prototype.Set_VertAlign = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.VertAlign === Value)
		return;

	History.Add(new CChangesParaTextPrVertAlign(this, this.Value.VertAlign, Value));
	this.Value.VertAlign = Value;
};
ParaTextPr.prototype.Set_HighLight = function(Value)
{
	if (null === Value)
		Value = undefined;

	History.Add(new CChangesParaTextPrHighLight(this, this.Value.HighLight, Value));
	this.Value.HighLight = Value;
};
ParaTextPr.prototype.SetHighlightColor = function(Value)
{
	if (null === Value)
		Value = undefined;

	History.Add(new CChangesParaTextPrHighlightColor(this, this.Value.HighlightColor, Value));
	this.Value.HighlightColor = Value;
};
ParaTextPr.prototype.Set_RStyle = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.RStyle === Value)
		return;

	History.Add(new CChangesParaTextPrRStyle(this, this.Value.RStyle, Value));
	this.Value.RStyle = Value;
};
ParaTextPr.prototype.SetRStyle = function(styleId)
{
	this.Set_RStyle(styleId);
};
ParaTextPr.prototype.Set_Spacing = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.Spacing === Value)
		return;

	History.Add(new CChangesParaTextPrSpacing(this, this.Value.Spacing, Value));
	this.Value.Spacing = Value;
};
ParaTextPr.prototype.Set_DStrikeout = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.DStrikeout === Value)
		return;

	History.Add(new CChangesParaTextPrDStrikeout(this, this.Value.DStrikeout, Value));
	this.Value.DStrikeout = Value;
};
ParaTextPr.prototype.Set_Caps = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.Caps === Value)
		return;

	History.Add(new CChangesParaTextPrCaps(this, this.Value.Caps, Value));
	this.Value.Caps = Value;
};
ParaTextPr.prototype.Set_SmallCaps = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.SmallCaps === Value)
		return;

	History.Add(new CChangesParaTextPrSmallCaps(this, this.Value.SmallCaps, Value));
	this.Value.SmallCaps = Value;
};
ParaTextPr.prototype.Set_Position = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.Position === Value)
		return;

	History.Add(new CChangesParaTextPrPosition(this, this.Value.Position, Value));
	this.Value.Position = Value;
};
ParaTextPr.prototype.Set_Value = function(Value)
{
	if (!Value || !(Value instanceof CTextPr) || true === this.Value.Is_Equal(Value))
		return;

	History.Add(new CChangesParaTextPrValue(this, this.Value, Value));
	this.Value = Value;
};
ParaTextPr.prototype.Set_RFonts = function(Value)
{
	var _Value = Value ? Value : new CRFonts();
	History.Add(new CChangesParaTextPrRFonts(this, this.Value.RFonts, _Value));
	this.Value.RFonts = _Value;
};
ParaTextPr.prototype.Set_RFonts2 = function(oRFonts)
{
	if (oRFonts)
	{
		if (oRFonts.AsciiTheme)
		{
			this.SetRFontsAscii(undefined);
			this.SetRFontsAsciiTheme(oRFonts.AsciiTheme);
		}
		else if (oRFonts.Ascii)
		{
			this.SetRFontsAscii(oRFonts.Ascii);
			this.SetRFontsAsciiTheme(undefined);
		}
		else
		{
			if (null === oRFonts.Ascii)
				this.SetRFontsAscii(undefined);

			if (null === oRFonts.AsciiTheme)
				this.SetRFontsAsciiTheme(undefined);
		}

		if (oRFonts.HAnsiTheme)
		{
			this.SetRFontsHAnsi(undefined);
			this.SetRFontsHAnsiTheme(oRFonts.HAnsiTheme);
		}
		else if (oRFonts.HAnsi)
		{
			this.SetRFontsHAnsi(oRFonts.HAnsi);
			this.SetRFontsHAnsiTheme(undefined);
		}
		else
		{
			if (null === oRFonts.HAnsi)
				this.SetRFontsHAnsi(undefined);

			if (null === oRFonts.HAnsiTheme)
				this.SetRFontsHAnsiTheme(undefined);
		}

		if (oRFonts.CSTheme)
		{
			this.SetRFontsCS(undefined);
			this.SetRFontsCSTheme(oRFonts.CSTheme);
		}
		else if (oRFonts.CS)
		{
			this.SetRFontsCS(oRFonts.CS);
			this.SetRFontsCSTheme(undefined);
		}
		else
		{
			if (null === oRFonts.CS)
				this.SetRFontsCS(undefined);

			if (null === oRFonts.CSTheme)
				this.SetRFontsCSTheme(undefined);
		}

		if (oRFonts.EastAsiaTheme)
		{
			this.SetRFontsEastAsia(undefined);
			this.SetRFontsEastAsiaTheme(oRFonts.EastAsiaTheme);
		}
		else if (oRFonts.EastAsia)
		{
			this.SetRFontsEastAsia(oRFonts.EastAsia);
			this.SetRFontsEastAsiaTheme(undefined);
		}
		else
		{
			if (null === oRFonts.EastAsia)
				this.SetRFontsEastAsia(undefined);

			if (null === oRFonts.EastAsiaTheme)
				this.SetRFontsEastAsiaTheme(undefined);
		}

		if (undefined !== oRFonts.Hint)
			this.SetRFontsHint(null === oRFonts.Hint ? undefined : oRFonts.Hint);
	}
	else
	{
		this.SetRFontsAscii(undefined);
		this.SetRFontsAsciiTheme(undefined);
		this.SetRFontsHAnsi(undefined);
		this.SetRFontsHAnsiTheme(undefined);
		this.SetRFontsCS(undefined);
		this.SetRFontsCSTheme(undefined);
		this.SetRFontsEastAsia(undefined);
		this.SetRFontsEastAsiaTheme(undefined);
		this.SetRFontsHint(undefined);
	}
};
ParaTextPr.prototype.SetRFontsAscii = function(Value)
{
	if (null === Value)
		Value = undefined;

	History.Add(new CChangesParaTextPrRFontsAscii(this, this.Value.RFonts.Ascii, Value));
	this.Value.RFonts.Ascii = Value;
};
ParaTextPr.prototype.SetRFontsHAnsi = function(Value)
{
	if (null === Value)
		Value = undefined;

	History.Add(new CChangesParaTextPrRFontsHAnsi(this, this.Value.RFonts.HAnsi, Value));
	this.Value.RFonts.HAnsi = Value;
};
ParaTextPr.prototype.SetRFontsCS = function(Value)
{
	if (null === Value)
		Value = undefined;

	History.Add(new CChangesParaTextPrRFontsCS(this, this.Value.RFonts.CS, Value));
	this.Value.RFonts.CS = Value;
};
ParaTextPr.prototype.SetRFontsEastAsia = function(Value)
{
	if (null === Value)
		Value = undefined;

	History.Add(new CChangesParaTextPrRFontsEastAsia(this, this.Value.RFonts.EastAsia, Value));
	this.Value.RFonts.EastAsia = Value;
};
ParaTextPr.prototype.SetRFontsHint = function(Value)
{
	if (null === Value)
		Value = undefined;

	History.Add(new CChangesParaTextPrRFontsHint(this, this.Value.RFonts.Hint, Value));
	this.Value.RFonts.Hint = Value;
};
ParaTextPr.prototype.SetRFontsAsciiTheme = function(sValue)
{
	var _sValue = (!sValue ? undefined : sValue);
	if (_sValue !== this.Value.RFonts.AsciiTheme)
	{
		AscCommon.History.Add(new CChangesParaTextPrRFontsAsciiTheme(this, this.Value.RFonts.AsciiTheme, _sValue));
		this.Value.RFonts.AsciiTheme = _sValue;
	}
};
ParaTextPr.prototype.SetRFontsHAnsiTheme = function(sValue)
{
	var _sValue = (!sValue ? undefined : sValue);
	if (_sValue !== this.Value.RFonts.HAnsiTheme)
	{
		AscCommon.History.Add(new CChangesParaTextPrRFontsHAnsiTheme(this, this.Value.RFonts.HAnsiTheme, _sValue));
		this.Value.RFonts.HAnsiTheme = _sValue;
	}
};
ParaTextPr.prototype.SetRFontsCSTheme = function(sValue)
{
	var _sValue = (!sValue ? undefined : sValue);
	if (_sValue !== this.Value.RFonts.CSTheme)
	{
		AscCommon.History.Add(new CChangesParaTextPrRFontsCSTheme(this, this.Value.RFonts.CSTheme, _sValue));
		this.Value.RFonts.CSTheme = _sValue;
	}
};
ParaTextPr.prototype.SetRFontsEastAsiaTheme = function(sValue)
{
	var _sValue = (!sValue ? undefined : sValue);
	if (_sValue !== this.Value.RFonts.EastAsiaTheme)
	{
		AscCommon.History.Add(new CChangesParaTextPrRFontsEastAsiaTheme(this, this.Value.RFonts.EastAsiaTheme, _sValue));
		this.Value.RFonts.EastAsiaTheme = _sValue;
	}
};
ParaTextPr.prototype.Set_Lang = function(Value)
{
	var _Value = new CLang();
	if (Value)
		_Value.Set_FromObject(Value);

	History.Add(new CChangesParaTextPrLang(this, this.Value.Lang, Value));
	this.Value.Lang = _Value;
};
ParaTextPr.prototype.Set_Lang_Bidi = function(Value)
{
	History.Add(new CChangesParaTextPrLangBidi(this, this.Value.Lang.Bidi, Value));
	this.Value.Lang.Bidi = Value;
};
ParaTextPr.prototype.Set_Lang_EastAsia = function(Value)
{
	History.Add(new CChangesParaTextPrLangEastAsia(this, this.Value.Lang.EastAsia, Value));
	this.Value.Lang.EastAsia = Value;
};
ParaTextPr.prototype.Set_Lang_Val = function(Value)
{
	History.Add(new CChangesParaTextPrLangVal(this, this.Value.Lang.Val, Value));
	this.Value.Lang.Val = Value;
};
ParaTextPr.prototype.Set_Unifill = function(Value)
{
	History.Add(new CChangesParaTextPrUnifill(this, this.Value.Unifill, Value));
	this.Value.Unifill = Value;
};
ParaTextPr.prototype.SetFontSizeCS = function(Value)
{
	if (null === Value)
		Value = undefined;

	if (this.Value.FontSizeCS === Value)
		return;

	History.Add(new CChangesParaTextPrFontSizeCS(this, this.Value.FontSizeCS, Value));
	this.Value.FontSizeCS = Value;
};
ParaTextPr.prototype.Set_TextOutline = function(Value)
{
	History.Add(new CChangesParaTextPrTextOutline(this, this.Value.TextOutline, Value));
	this.Value.TextOutline = Value;
};
ParaTextPr.prototype.Set_TextFill = function(Value)
{
	History.Add(new CChangesParaTextPrTextFill(this, this.Value.TextFill, Value));
	this.Value.TextFill = Value;
};
ParaTextPr.prototype.SetBoldCS = function(isBold)
{
	if (this.Value.BoldCS === isBold)
		return;

	let oChange = new CChangesParaTextPrBoldCS(this, this.Value.BoldCS, isBold);
	AscCommon.History.Add(oChange);
	oChange.Redo();
};
ParaTextPr.prototype.SetItalicCS = function(isItalic)
{
	if (this.Value.ItalicCS === isItalic)
		return;

	let oChange = new CChangesParaTextPrBoldCS(this, this.Value.ItalicCS, isItalic);
	AscCommon.History.Add(oChange);
	oChange.Redo();
};
ParaTextPr.prototype.SetLigatures = function(nType)
{
	if (this.Value.Ligatures === nType)
		return;

	let oChange = new CChangesParaTextPrLigatures(this, this.Value.Ligatures, nType);
	AscCommon.History.Add(oChange);
	oChange.Redo();
};
/**
 * Жестко выставляем заданные настройки
 * @param {CTextPr} textPr
 */
ParaTextPr.prototype.SetPr = function(textPr)
{
	if (!textPr)
		textPr = new CTextPr();

	this.Set_Value(textPr);
};
ParaTextPr.prototype.IncreaseDecreaseFontSize = function(isIncrease)
{
	let oParagraph = this.GetParagraph();
	if (!oParagraph)
		return;

	let oTextPr = oParagraph.GetParaEndCompiledPr();
	this.SetFontSizeCS(oTextPr.GetIncDecFontSizeCS(isIncrease));
	this.SetFontSize(oTextPr.GetIncDecFontSize(isIncrease));
};
ParaTextPr.prototype.ApplyFontFamily = function(sFontName)
{
	this.SetRFontsAscii({Name : sFontName, Index : -1});
	this.SetRFontsHAnsi({Name : sFontName, Index : -1});
	this.SetRFontsCS({Name : sFontName, Index : -1});

	this.SetRFontsAsciiTheme(undefined);
	this.SetRFontsHAnsiTheme(undefined);
	this.SetRFontsCSTheme(undefined);

	this.SetRFontsEastAsia(undefined);
	this.SetRFontsEastAsiaTheme(undefined);
};
//----------------------------------------------------------------------------------------------------------------------
// Undo/Redo функции
//----------------------------------------------------------------------------------------------------------------------
ParaTextPr.prototype.Get_ParentObject_or_DocumentPos = function()
{
	if (null != this.Parent)
		return this.Parent.Get_ParentObject_or_DocumentPos();
};
ParaTextPr.prototype.Refresh_RecalcData = function(Data)
{
	if (undefined !== this.Parent && null !== this.Parent)
		this.Parent.Refresh_RecalcData2();
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с совместным редактирования
//----------------------------------------------------------------------------------------------------------------------
ParaTextPr.prototype.Write_ToBinary = function(Writer)
{
	// Long   : Type
	// String : Id

	Writer.WriteLong(this.Type);
	Writer.WriteString2(this.Id);
};
ParaTextPr.prototype.Write_ToBinary2 = function(Writer)
{
	Writer.WriteLong(AscDFH.historyitem_type_TextPr);

	// Long   : Type
	// String : Id
	// Long   : Value

	Writer.WriteLong(this.Type);
	Writer.WriteString2(this.Id);
	this.Value.Write_ToBinary(Writer);
};
ParaTextPr.prototype.Read_FromBinary2 = function(Reader)
{
	this.Type = Reader.GetLong();
	this.Id   = Reader.GetString2();

	this.Value.Clear();
	this.Value.Read_FromBinary(Reader);
};

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].ParaTextPr = ParaTextPr;
window['AscWord'].ParaTextPr = ParaTextPr;
