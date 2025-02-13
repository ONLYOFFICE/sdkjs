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
	const FLAGS_MARK_NORMAL     = 0x00; // Два бита выделяем под тип (не побитово)
	const FLAGS_MARK_ENDCELL    = 0x01;
	const FLAGS_MARK_ENDSECTION = 0x02;

	const FLAGS_MARK_MASK       = 0x03;

	/**
	 * Класс представляющий символ конца параграфа
	 * @constructor
	 * @extends {AscWord.CRunElementBase}
	 */
	function CRunParagraphMark()
	{
		AscWord.CRunElementBase.call(this);

		this.Grapheme     = AscFonts.NO_GRAPHEME;
		this.SectionEnd   = null;
		this.WidthVisible = 0x00000000 | 0;
		this.Flags        = 0x00000000 | 0;
	}
	CRunParagraphMark.prototype = Object.create(AscWord.CRunElementBase.prototype);
	CRunParagraphMark.prototype.constructor = CRunParagraphMark;

	CRunParagraphMark.prototype.Type = para_End;
	CRunParagraphMark.prototype.IsParaEnd = function()
	{
		return true;
	};
	CRunParagraphMark.prototype.Draw = function(X, Y, Context)
	{
		if (this.SectionEnd)
			return this.private_DrawSectionEnd(X, Y, Context);

		if (!this.Grapheme)
			return;

		let nFontSize = (((this.Flags >> 16) & 0xFFFF) / 64);
		AscFonts.DrawGrapheme(this.Grapheme, Context, X, Y, nFontSize);
	};
	CRunParagraphMark.prototype.Measure = function(oMeasurer, oTextPr, isEndCell)
	{
		this.private_UpdateMarkType(isEndCell ? FLAGS_MARK_ENDCELL : FLAGS_MARK_NORMAL);

		let nUnicode  = isEndCell ? 0x00A4 : 0x00B6;
		let nFontSlot = oTextPr.RTL || oTextPr.CS ? AscWord.fontslot_CS : AscWord.fontslot_ASCII;
		let oFontInfo = oTextPr.GetFontInfo(nFontSlot);
		this.Grapheme = oMeasurer.GetGraphemeByUnicode(nUnicode, oFontInfo.Name, oFontInfo.Style);

		let nFontSize = oFontInfo.Size;
		if (oTextPr.VertAlign !== AscCommon.vertalign_Baseline)
			nFontSize = AscWord.AlignFontSize(nFontSize, AscCommon.vaKSize);

		this.Flags = (this.Flags & 0xFFFF) | (((nFontSize * 64) & 0xFFFF) << 16);

		this.Width        = 0;
		this.WidthVisible = ((AscFonts.GetGraphemeWidth(this.Grapheme) * nFontSize) * AscWord.TEXTWIDTH_DIVIDER) | 0;
	};
	CRunParagraphMark.prototype.GetWidth = function()
	{
		return 0;
	};
	CRunParagraphMark.prototype.UpdateSectionEnd = function(nSectionType, nWidth, oLogicDocument)
	{
		this.private_UpdateMarkType(FLAGS_MARK_ENDSECTION);

		var oPr = oLogicDocument.GetSectionEndMarkPr(nSectionType);

		var nStrWidth = oPr.StringWidth;
		var nSymWidth = oPr.ColonWidth;

		this.SectionEnd = {
			String       : null,
			ColonsCount  : 0,
			ColonWidth   : nSymWidth,
			ColonSymbol  : oPr.ColonSymbol,
			Widths       : []
		};

		if (nWidth - 6 * nSymWidth >= nStrWidth)
		{
			this.SectionEnd.ColonsCount = parseInt((nWidth - nStrWidth) / (2 * nSymWidth));
			this.SectionEnd.String      = oPr.String;

			var nAdd = 0;
			var nResultWidth = 2 * nSymWidth * this.SectionEnd.ColonsCount + nStrWidth;
			if (nResultWidth < nWidth)
			{
				nAdd = (nWidth - nResultWidth) / (2 * this.SectionEnd.ColonsCount + this.SectionEnd.Widths.length);
				this.SectionEnd.ColonWidth += nAdd;
			}

			for (var nPos = 0, nLen = oPr.Widths.length; nPos < nLen; ++nPos)
			{
				this.SectionEnd.Widths[nPos] = oPr.Widths[nPos] + nAdd;
			}
		}
		else
		{
			this.SectionEnd.ColonsCount = parseInt(nWidth / nSymWidth);

			var nResultWidth = nSymWidth * this.SectionEnd.ColonsCount;
			if (nResultWidth < nWidth && this.SectionEnd.ColonsCount > 0)
				this.SectionEnd.ColonWidth += (nWidth - nResultWidth) /this.SectionEnd.ColonsCount ;
		}

		this.WidthVisible = (nWidth * AscWord.TEXTWIDTH_DIVIDER) | 0;
	};
	CRunParagraphMark.prototype.ClearSectionEnd = function()
	{
		this.SectionEnd = null;
	};
	CRunParagraphMark.prototype.CheckMark = function(oParagraph, nRangeW)
	{
		let oSectPr        = oParagraph.Get_SectionPr();
		let oLogicDocument = oParagraph.GetLogicDocument();

		if (!oLogicDocument
			|| oLogicDocument !== oParagraph.GetParent()
			|| !oLogicDocument.IsDocumentEditor())
			oSectPr = null;

		if (oSectPr)
		{
			let oNextSectPr = oLogicDocument.SectionsInfo.Get_SectPr(oParagraph.GetIndex() + 1).SectPr;
			this.UpdateSectionEnd(oNextSectPr.Type, nRangeW, oLogicDocument);
		}
		else
		{
			this.ClearSectionEnd();

			let isEndCell = oParagraph.IsLastParagraphInCell();

			let nType = isEndCell ? FLAGS_MARK_ENDCELL : FLAGS_MARK_NORMAL;
			if (nType !== this.GetMarkType())
			{
				let oMeasurer = AscCommon.g_oTextMeasurer;
				let oTextPr   = oParagraph.GetParaEndCompiledPr();
				oMeasurer.SetTextPr(oTextPr, oParagraph.GetTheme());
				this.Measure(oMeasurer, oTextPr, isEndCell);
			}
		}
	};
	CRunParagraphMark.prototype.CanAddNumbering = function()
	{
		return true;
	};
	CRunParagraphMark.prototype.Copy = function()
	{
		return new CRunParagraphMark();
	};
	CRunParagraphMark.prototype.Write_ToBinary = function(Writer)
	{
		// Long   : Type
		Writer.WriteLong(para_End);
	};
	CRunParagraphMark.prototype.Read_FromBinary = function(Reader)
	{
	};
	CRunParagraphMark.prototype.GetAutoCorrectFlags = function()
	{
		return (AscWord.AUTOCORRECT_FLAGS_FIRST_LETTER_SENTENCE
			| AscWord.AUTOCORRECT_FLAGS_HYPERLINK
			| AscWord.AUTOCORRECT_FLAGS_HYPHEN_WITH_DASH);
	};
	CRunParagraphMark.prototype.ToSearchElement = function(oProps)
	{
		return new AscCommonWord.CSearchTextSpecialParaEnd();
	};
	CRunParagraphMark.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_Unknown;
	};
	CRunParagraphMark.prototype.GetMarkType = function()
	{
		return (this.Flags & FLAGS_MARK_MASK);
	};
	CRunParagraphMark.prototype.getBidiType = function()
	{
		return AscBidi.TYPE.PM;
	};
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	CRunParagraphMark.prototype.private_DrawSectionEnd = function(X, Y, Context)
	{
		Context.b_color1(0, 0, 0, 255);
		Context.p_color(0, 0, 0, 255);
		Context.SetFont({
			FontFamily : {Name : "Courier New", Index : -1},
			FontSize   : 8,
			Italic     : false,
			Bold       : false
		});

		for (var nPos = 0, nCount = this.SectionEnd.ColonsCount; nPos < nCount; ++nPos)
		{
			Context.FillTextCode(X, Y, this.SectionEnd.ColonSymbol);
			X += this.SectionEnd.ColonWidth;
		}

		if (this.SectionEnd.String)
		{
			for (var nPos = 0, nCount = this.SectionEnd.String.length; nPos < nCount; ++nPos)
			{
				Context.FillText(X, Y, this.SectionEnd.String[nPos]);
				X += this.SectionEnd.Widths[nPos];
			}

			for (var nPos = 0, nCount = this.SectionEnd.ColonsCount; nPos < nCount; ++nPos)
			{
				Context.FillTextCode(X, Y, this.SectionEnd.ColonSymbol);
				X += this.SectionEnd.ColonWidth;
			}
		}
	};
	CRunParagraphMark.prototype.private_UpdateMarkType = function(nType)
	{
		this.Flags = (this.Flags & 0xFFFFFFFC) | nType;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunParagraphMark = CRunParagraphMark;

})(window);
