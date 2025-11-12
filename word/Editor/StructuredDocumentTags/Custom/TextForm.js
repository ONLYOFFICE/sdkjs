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
	 * Класс с настройками для текстовой формы
	 * @constructor
	 */
	function CSdtTextFormPr(nMax, isComb, nWidth, nSymbol, sFont, oCombBorder)
	{
		this.MaxCharacters         = undefined !== nMax ? nMax : -1;
		this.Comb                  = undefined !== isComb ? isComb : false;
		this.Width                 = nWidth;
		this.WidthRule             = Asc.CombFormWidthRule.Auto;
		this.CombPlaceholderSymbol = nSymbol;
		this.CombPlaceholderFont   = sFont;
		this.CombBorder            = undefined !== oCombBorder ? oCombBorder.Copy() : undefined;
		this.MultiLine             = false;
		this.AutoFit               = false;
		this.Format                = new AscWord.CTextFormFormat();
	}
	CSdtTextFormPr.prototype.Copy = function()
	{
		var oText = new CSdtTextFormPr();

		oText.MaxCharacters         = this.MaxCharacters;
		oText.Comb                  = this.Comb;
		oText.Width                 = this.Width;
		oText.WidthRule             = this.WidthRule;
		oText.CombPlaceholderSymbol = this.CombPlaceholderSymbol;
		oText.CombPlaceholderFont   = this.CombPlaceholderFont;
		oText.CombBorder            = this.CombBorder ? this.CombBorder.Copy() : undefined;
		oText.MultiLine             = this.MultiLine;
		oText.AutoFit               = this.AutoFit;
		oText.Format                = this.Format.Copy();

		return oText;
	};
	CSdtTextFormPr.prototype.IsEqual = function(oOther)
	{
		return (oOther
			&& this.MaxCharacters === oOther.MaxCharacters
			&& this.Comb === oOther.Comb
			&& this.Width === oOther.Width
			&& this.WidthRule === oOther.WidthRule
			&& this.CombPlaceholderSymbol === oOther.CombPlaceholderSymbol
			&& this.CombPlaceholderFont === oOther.CombPlaceholderFont
			&& ((!this.CombBorder && !oOther) || (this.CombBorder && this.CombBorder.IsEqual(oOther)))
			&& this.MultiLine === oOther.MultiLine
			&& this.AutoFit === oOther.AutoFit
			&& this.Format.IsEqual(oOther.Format)
		);
	};
	CSdtTextFormPr.prototype.WriteToBinary = function(oWriter)
	{
		oWriter.WriteLong(this.MaxCharacters);
		oWriter.WriteBool(this.Comb);
		oWriter.WriteLong(this.Width);

		if (undefined !== this.CombPlaceholderSymbol)
		{
			oWriter.WriteBool(true);
			oWriter.WriteLong(this.CombPlaceholderSymbol);
		}
		else
		{
			oWriter.WriteBool(false);
		}

		if (undefined !== this.CombPlaceholderFont)
		{
			oWriter.WriteBool(true);
			oWriter.WriteString2(this.CombPlaceholderFont);
		}
		else
		{
			oWriter.WriteBool(false);
		}

		if (undefined !== this.CombBorder)
		{
			oWriter.WriteBool(true);
			this.CombBorder.WriteToBinary(oWriter);
		}
		else
		{
			oWriter.WriteBool(false);
		}

		oWriter.WriteBool(this.MultiLine);
		oWriter.WriteBool(this.AutoFit);
		oWriter.WriteLong(this.WidthRule);

		this.Format.WriteToBinary(oWriter);
	};
	CSdtTextFormPr.prototype.ReadFromBinary = function(oReader)
	{
		this.MaxCharacters = oReader.GetLong();
		this.Comb          = oReader.GetBool();
		this.Width         = oReader.GetLong();

		if (oReader.GetBool())
			this.CombPlaceholderSymbol = oReader.GetLong();

		if (oReader.GetBool())
			this.CombPlaceholderFont = oReader.GetString2();

		if (oReader.GetBool())
		{
			this.CombBorder = new CDocumentBorder();
			this.CombBorder.ReadFromBinary(oReader);
		}

		this.MultiLine = oReader.GetBool();
		this.AutoFit   = oReader.GetBool();
		this.WidthRule = oReader.GetLong();

		this.Format.ReadFromBinary(oReader);
	};
	CSdtTextFormPr.prototype.Write_ToBinary = function(oWriter)
	{
		this.WriteToBinary(oWriter);
	};
	CSdtTextFormPr.prototype.Read_FromBinary = function(oReader)
	{
		this.ReadFromBinary(oReader);
	};
	CSdtTextFormPr.prototype.GetMaxCharacters = function()
	{
		if (this.Format.IsMask())
			return this.Format.GetMaskLength();

		return this.MaxCharacters;
	};
	CSdtTextFormPr.prototype.SetMaxCharacters = function(nMax)
	{
		if (undefined === nMax || null === nMax || nMax <= 0)
			this.MaxCharacters = -1;
		else
			this.MaxCharacters = nMax | 0;
	};
	CSdtTextFormPr.prototype.GetComb = function()
	{
		return this.Comb;
	};
	CSdtTextFormPr.prototype.SetComb = function(isComb)
	{
		this.Comb = isComb;
	};
	CSdtTextFormPr.prototype.GetWidth = function()
	{
		return this.Width;
	};
	CSdtTextFormPr.prototype.SetWidth = function(nWidth)
	{
		this.Width = nWidth;
	};
	CSdtTextFormPr.prototype.GetPlaceHolderSymbol = function()
	{
		return this.CombPlaceholderSymbol;
	};
	CSdtTextFormPr.prototype.SetPlaceHolderSymbol = function(nCharCode)
	{
		this.CombPlaceholderSymbol = nCharCode;
	};
	CSdtTextFormPr.prototype.GetPlaceHolderFont = function()
	{
		return this.CombPlaceholderFont;
	};
	CSdtTextFormPr.prototype.SetPlaceHolderFont = function(sFont)
	{
		this.CombPlaceholderFont = sFont;
	};
	CSdtTextFormPr.prototype.GetCombBorder = function()
	{
		return this.CombBorder;
	};
	CSdtTextFormPr.prototype.GetAscCombBorder = function()
	{
		if (!this.CombBorder)
			return undefined;

		return (new Asc.asc_CTextBorder(this.CombBorder));
	};
	CSdtTextFormPr.prototype.SetAscCombBorder = function(oAscBorder)
	{
		if (!oAscBorder)
		{
			this.CombBorder = undefined;
		}
		else
		{
			this.CombBorder = new CDocumentBorder();
			this.CombBorder.Set_FromObject(oAscBorder);
		}
	};
	CSdtTextFormPr.prototype.IsComb = function()
	{
		let maxCharacters = this.GetMaxCharacters();
		return !!(this.Comb && undefined !== maxCharacters && maxCharacters <= 1000);
	};
	CSdtTextFormPr.prototype.GetMultiLine = function()
	{
		return (this.MultiLine && !this.IsComb());
	};
	CSdtTextFormPr.prototype.SetMultiLine = function(isMultiLine)
	{
		this.MultiLine = isMultiLine;
	};
	CSdtTextFormPr.prototype.GetAutoFit = function()
	{
		if (this.Comb)
			return false;

		return this.AutoFit;
	};
	CSdtTextFormPr.prototype.SetAutoFit = function(isAutoFit)
	{
		this.AutoFit = isAutoFit;
	};
	CSdtTextFormPr.prototype.GetWidthRule = function()
	{
		return this.WidthRule;
	};
	CSdtTextFormPr.prototype.SetWidthRule = function(nRule)
	{
		this.WidthRule = nRule;
	};
	CSdtTextFormPr.prototype.GetFormatType = function()
	{
		return this.Format.GetType();
	};
	CSdtTextFormPr.prototype.SetDigitFormat = function()
	{
		this.Format.SetDigit();
	};
	CSdtTextFormPr.prototype.SetLetterFormat = function()
	{
		this.Format.SetLetter();
	};
	CSdtTextFormPr.prototype.SetMaskFormat = function(sMask)
	{
		this.Format.SetMask(sMask);
	};
	CSdtTextFormPr.prototype.GetMaskFormat = function()
	{
		return this.Format.GetMask();
	};
	CSdtTextFormPr.prototype.SetRegExpFormat = function(sRegExp)
	{
		this.Format.SetRegExp(sRegExp);
	};
	CSdtTextFormPr.prototype.GetRegExpFormat = function()
	{
		return this.Format.GetRegExp();
	};
	CSdtTextFormPr.prototype.SetNoneFormat = function()
	{
		this.Format.SetNone();
	};
	CSdtTextFormPr.prototype.SetFormatSymbols = function(arrSymbols)
	{
		this.Format.SetSymbols(arrSymbols);
	};
	CSdtTextFormPr.prototype.GetFormatSymbols = function()
	{
		return this.Format.GetSymbols(true);
	};
	CSdtTextFormPr.prototype.CheckFormat = function(sText, isFullCheck)
	{
		return this.Format.Check(sText, isFullCheck);
	};
	CSdtTextFormPr.prototype.CheckFormatOnFly = function(sText)
	{
		return this.Format.CheckOnFly(sText);
	};
	CSdtTextFormPr.prototype.GetFormat = function()
	{
		return this.Format;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommon'] = window['AscCommon'] || {};
	window['AscCommon'].CSdtTextFormPr    = CSdtTextFormPr;
	window['AscCommon']['CSdtTextFormPr'] = CSdtTextFormPr;

	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CSdtTextFormPr = CSdtTextFormPr;

	CSdtTextFormPr.prototype['get_MaxCharacters']     = CSdtTextFormPr.prototype.GetMaxCharacters;
	CSdtTextFormPr.prototype['put_MaxCharacters']     = CSdtTextFormPr.prototype.SetMaxCharacters;
	CSdtTextFormPr.prototype['get_Comb']              = CSdtTextFormPr.prototype.GetComb;
	CSdtTextFormPr.prototype['put_Comb']              = CSdtTextFormPr.prototype.SetComb;
	CSdtTextFormPr.prototype['get_Width']             = CSdtTextFormPr.prototype.GetWidth;
	CSdtTextFormPr.prototype['put_Width']             = CSdtTextFormPr.prototype.SetWidth;
	CSdtTextFormPr.prototype['get_PlaceHolderSymbol'] = CSdtTextFormPr.prototype.GetPlaceHolderSymbol;
	CSdtTextFormPr.prototype['put_PlaceHolderSymbol'] = CSdtTextFormPr.prototype.SetPlaceHolderSymbol;
	CSdtTextFormPr.prototype['get_PlaceHolderFont']   = CSdtTextFormPr.prototype.GetPlaceHolderFont;
	CSdtTextFormPr.prototype['put_PlaceHolderFont']   = CSdtTextFormPr.prototype.SetPlaceHolderFont;
	CSdtTextFormPr.prototype['get_CombBorder']        = CSdtTextFormPr.prototype.GetAscCombBorder;
	CSdtTextFormPr.prototype['put_CombBorder']        = CSdtTextFormPr.prototype.SetAscCombBorder;
	CSdtTextFormPr.prototype['get_MultiLine']         = CSdtTextFormPr.prototype.GetMultiLine;
	CSdtTextFormPr.prototype['put_MultiLine']         = CSdtTextFormPr.prototype.SetMultiLine;
	CSdtTextFormPr.prototype['get_AutoFit']           = CSdtTextFormPr.prototype.GetAutoFit;
	CSdtTextFormPr.prototype['put_AutoFit']           = CSdtTextFormPr.prototype.SetAutoFit;
	CSdtTextFormPr.prototype['get_WidthRule']         = CSdtTextFormPr.prototype.GetWidthRule;
	CSdtTextFormPr.prototype['put_WidthRule']         = CSdtTextFormPr.prototype.SetWidthRule;
	CSdtTextFormPr.prototype['get_FormatType']        = CSdtTextFormPr.prototype.GetFormatType;
	CSdtTextFormPr.prototype['put_NoneFormat']        = CSdtTextFormPr.prototype.SetNoneFormat;
	CSdtTextFormPr.prototype['put_DigitFormat']       = CSdtTextFormPr.prototype.SetDigitFormat;
	CSdtTextFormPr.prototype['put_LetterFormat']      = CSdtTextFormPr.prototype.SetLetterFormat;
	CSdtTextFormPr.prototype['put_MaskFormat']        = CSdtTextFormPr.prototype.SetMaskFormat;
	CSdtTextFormPr.prototype['get_MaskFormat']        = CSdtTextFormPr.prototype.GetMaskFormat;
	CSdtTextFormPr.prototype['put_RegExpFormat']      = CSdtTextFormPr.prototype.SetRegExpFormat;
	CSdtTextFormPr.prototype['get_RegExpFormat']      = CSdtTextFormPr.prototype.GetRegExpFormat;
	CSdtTextFormPr.prototype['get_FormatSymbols']     = CSdtTextFormPr.prototype.GetFormatSymbols;
	CSdtTextFormPr.prototype['put_FormatSymbols']     = CSdtTextFormPr.prototype.SetFormatSymbols;

})(window);
