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
	 * Класс представляющий ссылку на сноску
	 * @param {CFootEndnote} Footnote - Ссылка на сноску
	 * @param {string} CustomMark
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunFootnoteReference(Footnote, CustomMark)
	{
		this.Footnote   = Footnote instanceof AscCommonWord.CFootEndnote ? Footnote : null;
		this.CustomMark = CustomMark ? CustomMark : undefined;

		this.Width        = 0;
		this.WidthVisible = 0;
		this.Number       = 1;
		this.NumFormat    = Asc.c_oAscNumberingFormat.Decimal;
		this.TextAscent   = 0;

		this.Run          = null;
		this.Widths       = [];

		if (this.Footnote)
			this.Footnote.SetRef(this);
	}
	CRunFootnoteReference.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunFootnoteReference.prototype.constructor = CRunFootnoteReference;

	CRunFootnoteReference.prototype.Type = para_FootnoteReference;
	CRunFootnoteReference.prototype.Get_Type = function()
	{
		return para_FootnoteReference;
	};
	CRunFootnoteReference.prototype.Draw = function(X, Y, Context, PDSE)
	{
		if (true === this.IsCustomMarkFollows())
			return;

		var TextPr = this.Run.Get_CompiledPr(false);

		var FontKoef = 1;
		if (TextPr.VertAlign !== AscCommon.vertalign_Baseline)
			FontKoef = AscCommon.vaKSize;

		Context.SetFontSlot(AscWord.fontslot_ASCII, FontKoef);

		var _X = X;
		var T  = this.private_GetString();
		if (this.Widths.length !== T.length)
			return;

		for (var nPos = 0; nPos < T.length; ++nPos)
		{
			var Char = T.charAt(nPos);
			Context.FillText(_X, Y, Char);
			_X += this.Widths[nPos];
		}

		if (editor && editor.ShowParaMarks && Context.DrawFootnoteRect && this.Run)
		{
			Context.p_color(0, 0, 0, 255);
			Context.DrawFootnoteRect(X, PDSE.BaseLine - this.TextAscent, this.GetWidth(), this.TextAscent);
		}
	};
	CRunFootnoteReference.prototype.Measure = function(Context, TextPr, MathInfo, Run)
	{
		this.Run = Run;
		this.private_Measure();
	};
	CRunFootnoteReference.prototype.Copy = function(oPr)
	{
		if (this.Footnote)
		{
			var oFootnote;
			if (oPr && oPr.Comparison)
			{
				oFootnote = oPr.Comparison.createFootNote();
			}
			else
			{
				oFootnote = this.Footnote.Parent.CreateFootnote();
			}
			oFootnote.Copy2(this.Footnote, oPr);
		}

		var oRef = new CRunFootnoteReference(oFootnote);

		oRef.Number    = this.Number;
		oRef.NumFormat = this.NumFormat;

		return oRef;
	};
	CRunFootnoteReference.prototype.IsEqual = function(oElement)
	{
		return (oElement.Type === this.Type && this.Footnote === oElement.Footnote && oElement.CustomMark === this.CustomMark);
	};
	CRunFootnoteReference.prototype.Write_ToBinary = function(Writer)
	{
		// Long   : Type
		// String : FootnoteId
		// Bool : is undefined mark ?
		// false -> String2 : CustomMark
		Writer.WriteLong(this.Type);
		Writer.WriteString2(this.Footnote ? this.Footnote.GetId() : "");

		if (undefined === this.CustomMark)
		{
			Writer.WriteBool(true);
		}
		else
		{
			Writer.WriteBool(false);
			Writer.WriteString2(this.CustomMark);
		}
	};
	CRunFootnoteReference.prototype.Read_FromBinary = function(Reader)
	{
		// String : FootnoteId
		// Bool : is undefined mark ?
		// false -> String2 : CustomMark
		this.Footnote = g_oTableId.Get_ById(Reader.GetString2());
		
		if (this.Footnote)
			this.Footnote.SetRef(this);
		
		if (false === Reader.GetBool())
			this.CustomMark = Reader.GetString2();
	};
	CRunFootnoteReference.prototype.GetFootnote = function()
	{
		return this.Footnote;
	};
	CRunFootnoteReference.prototype.UpdateNumber = function(PRS, isKeepNumber)
	{
		if (this.Footnote && true !== PRS.IsFastRecalculate() && PRS.TopDocument instanceof CDocument)
		{
			var nPageAbs    = PRS.GetPageAbs();
			var nColumnAbs  = PRS.GetColumnAbs();
			var nAdditional = PRS.GetFootnoteReferencesCount(this);
			var oSectPr     = PRS.GetSectPr();
			var nNumFormat  = oSectPr.GetFootnoteNumFormat();

			var oLogicDocument       = this.Footnote.Get_LogicDocument();
			var oFootnotesController = oLogicDocument.GetFootnotesController();

			if (!isKeepNumber)
			{
				this.NumFormat = nNumFormat;
				this.Number    = oFootnotesController.GetFootnoteNumberOnPage(nPageAbs, nColumnAbs, oSectPr) + nAdditional;

				// Если данная сноска не участвует в нумерации, просто уменьшаем ей номер на 1, для упрощения работы
				if (this.IsCustomMarkFollows())
					this.Number--;
			}

			this.private_Measure();
			this.Footnote.SetNumber(this.Number, oSectPr, this.IsCustomMarkFollows());
		}
		else
		{
			this.Number    = 1;
			this.NumFormat = Asc.c_oAscNumberingFormat.Decimal;
			this.private_Measure();
		}
	};
	CRunFootnoteReference.prototype.private_Measure = function()
	{
		if (!this.Run)
			return;

		this.Width        = 0;
		this.WidthVisible = 0;
		this.TextAscent   = 0;

		if (this.IsCustomMarkFollows())
			return;

		let oMeasurer = g_oTextMeasurer;
		let oTextPr   = this.Run.Get_CompiledPr(false);

		var FontKoef = 1;
		if (oTextPr.VertAlign !== AscCommon.vertalign_Baseline)
			FontKoef = AscCommon.vaKSize;

		this.TextAscent = oTextPr.GetTextMetrics(AscWord.fontslot_ASCII, this.Run.GetParagraph().GetTheme()).Ascent;
		let oFontInfo   = oTextPr.GetFontInfo(AscWord.fontslot_ASCII);
		oMeasurer.SetFontInternal(oFontInfo.Name, oFontInfo.Size * FontKoef, oFontInfo.Style);

		var X = 0;
		var T = this.private_GetString();
		this.Widths = [];
		for (var nPos = 0; nPos < T.length; ++nPos)
		{
			var Char  = T.charAt(nPos);
			var CharW = oMeasurer.Measure(Char).Width
			this.Widths.push(CharW);
			X += CharW;
		}

		var ResultWidth   = (Math.max((X + oTextPr.Spacing), 0) * AscWord.TEXTWIDTH_DIVIDER) | 0;
		this.Width        = ResultWidth;
		this.WidthVisible = ResultWidth;
	};
	CRunFootnoteReference.prototype.private_GetString = function()
	{
		if (Asc.c_oAscNumberingFormat.Decimal === this.NumFormat)
			return Numbering_Number_To_String(this.Number);
		if (Asc.c_oAscNumberingFormat.LowerRoman === this.NumFormat)
			return Numbering_Number_To_Roman(this.Number, true);
		else if (Asc.c_oAscNumberingFormat.UpperRoman === this.NumFormat)
			return Numbering_Number_To_Roman(this.Number, false);
		else if (Asc.c_oAscNumberingFormat.LowerLetter === this.NumFormat)
			return Numbering_Number_To_Alpha(this.Number, true);
		else if (Asc.c_oAscNumberingFormat.UpperLetter === this.NumFormat)
			return Numbering_Number_To_Alpha(this.Number, false);
		else// if (Asc.c_oAscNumberingFormat.Decimal === this.NumFormat)
			return Numbering_Number_To_String(this.Number);
	};
	CRunFootnoteReference.prototype.IsCustomMarkFollows = function()
	{
		return (undefined !== this.CustomMark ? true : false);
	};
	CRunFootnoteReference.prototype.GetCustomText = function()
	{
		return this.CustomMark;
	};
	CRunFootnoteReference.prototype.CreateDocumentFontMap = function(FontMap)
	{
		if (this.Footnote)
			this.Footnote.Document_CreateFontMap(FontMap);
	};
	CRunFootnoteReference.prototype.GetAllContentControls = function(arrContentControls)
	{
		if (this.Footnote)
			this.Footnote.GetAllContentControls(arrContentControls);
	};
	CRunFootnoteReference.prototype.GetAllFontNames = function(arrAllFonts)
	{
		if (this.Footnote)
			this.Footnote.Document_Get_AllFontNames(arrAllFonts);
	};
	CRunFootnoteReference.prototype.SetParent = function(oRun)
	{
		this.Run = oRun;
	};
	CRunFootnoteReference.prototype.GetRun = function()
	{
		return this.Run;
	};
	CRunFootnoteReference.prototype.ToSearchElement = function(oProps)
	{
		return new AscCommonWord.CSearchTextSpecialFootnoteMark();
	};
	CRunFootnoteReference.prototype.PreDelete = function()
	{
		let run = this.GetRun();
		let paragraph = run ? run.GetParagraph() : null;
		
		if (!paragraph || !(paragraph.GetTopDocumentContent() instanceof AscWord.CDocument))
			return;
		
		var oFootnote = this.Footnote;
		if (oFootnote)
		{
			oFootnote.PreDelete();
			oFootnote.ClearContent(true);
		}
	};
	CRunFootnoteReference.prototype.IsReference = function()
	{
		return true;
	};
	CRunFootnoteReference.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_ASCII;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunFootnoteReference = CRunFootnoteReference;

})(window);
