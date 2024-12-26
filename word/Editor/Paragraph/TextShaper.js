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
	const CODEPOINT_TYPE = {
		BASE              : 0,
		LIGATURE          : 1,
		LIGATURE_CONTINUE : 2,
		COMBINING_MARK    : 3
	};

	/**
	 *
	 * @constructor
	 */
	function CParagraphTextShaper()
	{
		AscFonts.CTextShaper.call(this);

		this.Parent    = null;
		this.TextPr    = null;
		this.Temporary = false;
		this.Ligatures = Asc.LigaturesType.None;
		this.Spacing   = 0;
		this.AscFont   = false; // Специальный случай, когда используемый шрифт ASCW3, а не тот, что задан в настройках
	}
	CParagraphTextShaper.prototype = Object.create(AscFonts.CTextShaper.prototype);
	CParagraphTextShaper.prototype.constructor = CParagraphTextShaper;

	CParagraphTextShaper.prototype.Init = function(isTemporary)
	{
		this.Parent    = null;
		this.TextPr    = null;
		this.Temporary = isTemporary;
		this.Ligatures = Asc.LigaturesType.None;
		this.Spacing   = 0;
		this.AscFont   = false;
		
		this.ClearBuffer();
	};
	CParagraphTextShaper.prototype.GetCodePoint = function(oItem)
	{
		let nCodePoint = oItem.GetCodePoint();

		if (this.TextPr && (this.TextPr.Caps || this.TextPr.SmallCaps))
			nCodePoint = (String.fromCharCode(nCodePoint).toUpperCase()).charCodeAt(0);

		return nCodePoint;
	};
	CParagraphTextShaper.prototype.GetFontInfo = function(nFontSlot)
	{
		if (!this.TextPr)
			return AscFonts.DEFAULT_TEXTFONTINFO;

		let oFontInfo = this.TextPr.GetFontInfo(nFontSlot);

		if (this.AscFont)
			oFontInfo.Name = "ASCW3";

		return oFontInfo;
	};
	CParagraphTextShaper.prototype.GetFontSlot = function(nUnicode)
	{
		let oTextPr = this.TextPr;
		if (!oTextPr)
			return AscWord.fontslot_None;

		return AscWord.GetFontSlotByTextPr(nUnicode, oTextPr);
	};
	CParagraphTextShaper.prototype.GetLigaturesType = function()
	{
		return this.Ligatures;
	};
	CParagraphTextShaper.prototype.Shape = function(oParagraph)
	{
		this.Init(false);
		let oThis = this;
		oParagraph.CheckRunContent(function(oRun, nStartPos, nEndPos)
		{
			oThis.HandleRun(oRun, nStartPos, nEndPos);
		});
		this.FlushWord();
	};
	CParagraphTextShaper.prototype.ShapeRange = function(oParagraph, oStart, oEnd, isTemporary)
	{
		this.Init(isTemporary);
		let oThis = this;
		oParagraph.CheckRunContent(function(oRun, nStartPos, nEndPos)
		{
			oThis.HandleRun(oRun, nStartPos, nEndPos);
		}, oStart, oEnd);
		this.FlushWord();
	};
	CParagraphTextShaper.prototype.ShapeRun = function(run)
	{
		this.Init(false);
		this.HandleRun(run, 0, run.GetElementsCount());
		this.FlushWord();
	};
	CParagraphTextShaper.prototype.HandleRun = function(oRun, nStartPos, nEndPos)
	{
		this.private_CheckRun(oRun);

		for (let nPos = nStartPos; nPos < nEndPos; ++nPos)
		{
			let oItem = oRun.GetElement(nPos);
			if (!oItem.IsText())
			{
				this.FlushWord();
			}
			else if (oItem.IsNBSP())
			{
				this.FlushWord();
				this.private_HandleNBSP(oItem);
			}
			else
			{
				this.AppendToString(oItem);

				if (oItem.IsSpaceAfter())
					this.FlushWord();
			}
		}
	};
	CParagraphTextShaper.prototype.FlushGrapheme = function(nGrapheme, nWidth, nCodePointsCount, isLigature)
	{
		if (nCodePointsCount <= 0)
			return;
		
		let curIndex = 0;
		
		if (this.IsRtlDirection())
		{
			if (this.BufferIndex - nCodePointsCount < 0)
				return;
			
			this.BufferIndex -= nCodePointsCount;
			curIndex = this.BufferIndex;
		}
		else
		{
			if (this.BufferIndex + nCodePointsCount - 1 >= this.Buffer.length)
				return;
			
			curIndex = this.BufferIndex;
			this.BufferIndex += nCodePointsCount;
		}
		
		let _nWidth = (nWidth + (this.Spacing / this.FontSize)) / nCodePointsCount;
		if (1 === nCodePointsCount)
		{
			this.private_HandleItem(this.Buffer[curIndex], nGrapheme, _nWidth, this.FontSize, this.FontSlot, CODEPOINT_TYPE.BASE);
		}
		else
		{
			if (this.IsRtlDirection())
			{
				this.private_HandleItem(this.Buffer[curIndex], AscFonts.NO_GRAPHEME, _nWidth, this.FontSize, this.FontSlot, isLigature ? CODEPOINT_TYPE.LIGATURE : CODEPOINT_TYPE.BASE);
				this.private_HandleItem(this.Buffer[curIndex + nCodePointsCount - 1], nGrapheme, _nWidth, this.FontSize, this.FontSlot, isLigature ? CODEPOINT_TYPE.LIGATURE_CONTINUE : CODEPOINT_TYPE.COMBINING_MARK);
			}
			else
			{
				this.private_HandleItem(this.Buffer[curIndex], nGrapheme, _nWidth, this.FontSize, this.FontSlot, isLigature ? CODEPOINT_TYPE.LIGATURE : CODEPOINT_TYPE.BASE);
				this.private_HandleItem(this.Buffer[curIndex + nCodePointsCount - 1], AscFonts.NO_GRAPHEME, _nWidth, this.FontSize, this.FontSlot, isLigature ? CODEPOINT_TYPE.LIGATURE_CONTINUE : CODEPOINT_TYPE.COMBINING_MARK);
			}
			
			for (let nIndex = 1; nIndex < nCodePointsCount - 1; ++nIndex)
			{
				this.private_HandleItem(this.Buffer[++curIndex], AscFonts.NO_GRAPHEME, _nWidth, this.FontSize, AscWord.fontslot_ASCII, isLigature ? CODEPOINT_TYPE.LIGATURE_CONTINUE : CODEPOINT_TYPE.COMBINING_MARK);
			}
		}
	};
	CParagraphTextShaper.prototype.private_CheckRun = function(oRun)
	{
		let oRunParent = oRun.GetParent();
		let oTextPr    = oRun.Get_CompiledPr(false);

		if (this.Parent !== oRunParent || !this.IsEqualTextPr(oTextPr))
			this.FlushWord();

		let oForm      = oRun.GetParentForm();
		let isCombForm = oForm && oForm.IsTextForm() && oForm.GetTextFormPr().IsComb();

		this.Parent    = oRunParent;
		this.TextPr    = oTextPr;
		this.Spacing   = isCombForm ? 0 : oTextPr.Spacing;
		this.Ligatures = isCombForm || Math.abs(this.Spacing) > 0.001 ? Asc.LigaturesType.None : oTextPr.Ligatures;
		this.AscFont   = oRun.IsUseAscFont(oTextPr);
	};
	CParagraphTextShaper.prototype.private_HandleNBSP = function(oItem)
	{
		let oFontInfo = this.TextPr.GetFontInfo(AscWord.fontslot_ASCII);
		let nGrapheme = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(0x00B0, oFontInfo.Name, oFontInfo.Style);
		let nSpace    = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(0x0020, oFontInfo.Name, oFontInfo.Style);
		this.private_HandleItem(oItem, nGrapheme, AscFonts.GetGraphemeWidth(nSpace), oFontInfo.Size, AscWord.fontslot_ASCII, false, false, false);
	};
	CParagraphTextShaper.prototype.private_HandleItem = function(oItem, nGrapheme, nWidth, nFontSize, nFontSlot, nCodePointType)
	{
		if (this.Temporary)
		{
			oItem.ResetTemporaryGrapheme();
			if (nGrapheme !== oItem.GetGrapheme()
				|| nCodePointType !== oItem.GetCodePointType()
				|| Math.abs(nWidth - oItem.GetMeasuredWidth()) > 0.001)
			{
				oItem.SetTemporaryGrapheme(nGrapheme);
				oItem.SetTemporaryCodePointType(nCodePointType);
				oItem.SetTemporaryWidth(nWidth);
			}
		}
		else
		{
			oItem.SetGrapheme(nGrapheme);
			oItem.SetMetrics(nFontSize, nFontSlot, this.TextPr);
			oItem.SetCodePointType(nCodePointType);
			oItem.SetWidth(nWidth);
		}
	};
	CParagraphTextShaper.prototype.IsEqualTextPr = function(oTextPr)
	{
		// Здесь мы не используем стандартную функцию CTextPr.IsEqual, потому что она сравнивает на полное
		// совпадение объектов, а нас интересует только совпадение настроек, которые не мешают сборке текста

		if (!oTextPr || !this.TextPr)
			return false;

		let t = this.TextPr;

		return (t.Bold === oTextPr.Bold
			&& t.BoldCS === oTextPr.BoldCS
			&& t.Italic === oTextPr.Italic
			&& t.ItalicCS === oTextPr.ItalicCS
			&& IsEqualNullableFloatNumbers(t.FontSize, oTextPr.FontSize)
			&& IsEqualNullableFloatNumbers(t.FontSizeCS, oTextPr.FontSizeCS)
			&& t.VertAlign === oTextPr.VertAlign
			&& IsEqualNullableFloatNumbers(t.Spacing, oTextPr.Spacing)
			&& t.SmallCaps === oTextPr.SmallCaps
			&& t.Position === oTextPr.Position
			&& t.CS === oTextPr.CS
			&& t.RTL === oTextPr.RTL
			&& t.Vanish === oTextPr.Vanish
			&& t.Ligatures === oTextPr.Ligatures
			&& t.RFonts.IsEqual(oTextPr.RFonts));
	};
	CParagraphTextShaper.prototype.GetTextScript = function(nUnicode)
	{
		// TODO: Remove it after implementing bigi algorithm
		// Check bugs 66317, 66435
		if (0x060C <= nUnicode && nUnicode <= 0x074A)
			return AscFonts.HB_SCRIPT.HB_SCRIPT_ARABIC;
		
		let script = AscFonts.hb_get_script_by_unicode(nUnicode);
		if (AscFonts.HB_SCRIPT.HB_SCRIPT_COMMON === script && this.TextPr && this.TextPr.CS)
			return AscFonts.HB_SCRIPT.HB_SCRIPT_INHERITED;
		
		return script;
	};
	CParagraphTextShaper.prototype.ShapeRunTextItem = function(item, textPr)
	{
		let fontSlot = item.GetFontSlot(textPr);
		let fontInfo = textPr.GetFontInfo(fontSlot);
		let grapheme = AscCommon.g_oTextMeasurer.GetGraphemeByUnicode(item.GetCodePoint(), fontInfo.Name, fontInfo.Style);
		item.SetGrapheme(grapheme);
		item.SetMetrics(fontInfo.Size, fontSlot, textPr);
		item.SetCodePointType(CODEPOINT_TYPE.BASE);
		item.SetWidth(AscFonts.GetGraphemeWidth(grapheme));
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CODEPOINT_TYPE      = CODEPOINT_TYPE;
	window['AscWord'].ParagraphTextShaper = new CParagraphTextShaper();
	window['AscWord'].stringShaper        = new AscFonts.StringShaper();

})(window);
