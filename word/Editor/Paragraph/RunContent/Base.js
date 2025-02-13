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
	// TODO: Избавиться от использования TEXTWIDTH_DIVIDER константы вне данного модуля
	const TEXTWIDTH_DIVIDER = 16384;

	/**
	 * Базовый класс для элементов, лежащих внутри рана
	 * @constructor
	 */
	function CRunElementBase()
	{
		this.Width = 0x00000000 | 0;
	}

	CRunElementBase.prototype.Type = para_RunBase;
	CRunElementBase.prototype.Get_Type = function()
	{
		return this.Type;
	};
	CRunElementBase.prototype.GetType = function()
	{
		return this.Type;
	};
	CRunElementBase.prototype.Draw = function(nX, nY, oGraphics, PDSE)
	{
	};
	CRunElementBase.prototype.Measure = function(oGraphics, oTextPr)
	{
		this.Width = 0x00000000 | 0;
	};
	CRunElementBase.prototype.GetWidth = function()
	{
		return (this.Width > 0 ? this.Width / TEXTWIDTH_DIVIDER : 0);
	};
	CRunElementBase.prototype.Set_Width = function(nWidth)
	{
		this.Width = (nWidth * TEXTWIDTH_DIVIDER) | 0;
	};
	CRunElementBase.prototype.SetWidth = function(nWidth)
	{
		this.Width = (nWidth * TEXTWIDTH_DIVIDER) | 0;
	};
	/**
	 * Функция GetWidth возвращает ширину объекта, данная функция возвращает ширину объекта, которую он занимает
	 * внутри строки. Фактически, разница только для плавающих автофигур
	 * @returns {number}
	 */
	CRunElementBase.prototype.GetInlineWidth = function()
	{
		return this.GetWidth();
	};
	CRunElementBase.prototype.GetWidthVisible = function()
	{
		if (undefined !== this.WidthVisible)
			return (this.WidthVisible > 0 ? this.WidthVisible / TEXTWIDTH_DIVIDER : 0);

		return (this.Width > 0 ? this.Width / TEXTWIDTH_DIVIDER : 0);
	};
	CRunElementBase.prototype.SetWidthVisible = function(nWidthVisible)
	{
		this.WidthVisible = (nWidthVisible * TEXTWIDTH_DIVIDER) | 0;
	};
	CRunElementBase.prototype.CanAddNumbering = function()
	{
		return true;
	};
	CRunElementBase.prototype.SetParent = function(oParent)
	{
	};
	CRunElementBase.prototype.GetRun = function()
	{
		return null;
	};
	CRunElementBase.prototype.GetInRunPos = function()
	{
		let run = this.GetRun();
		if (!run)
			return -1;
		
		return run.GetElementPosition(this);
	};
	CRunElementBase.prototype.GetInParagraphPos = function()
	{
		let run = this.GetRun();
		if (!run)
			return null;
		
		let paragraph = run.GetParagraph();
		if (!paragraph)
			return null;
		
		let inRunPos = this.GetInRunPos();
		if (-1 === inRunPos)
			return null;
		
		let paraPos = paragraph.GetPosByElement(run);
		if (!paraPos)
			return null;
		
		paraPos.Add(inRunPos);
		return paraPos;
	};
	CRunElementBase.prototype.SetParagraph = function(oParagraph)
	{
	};
	CRunElementBase.prototype.GetParagraph = function()
	{
		let run = this.GetRun();
		return run ? run.GetParagraph() : null;
	};
	CRunElementBase.prototype.IsInPermRange = function()
	{
		let paragraph = this.GetParagraph();
		if (!paragraph)
			return false;
		
		let paraPos = this.GetInParagraphPos();
		if (!paraPos)
			return null;
		
		return paragraph.GetPermRangesByPos(paraPos).length > 0;
	};
	CRunElementBase.prototype.Copy = function()
	{
		return new this.constructor();
	};
	CRunElementBase.prototype.Write_ToBinary = function(Writer)
	{
		// Long : Type
		Writer.WriteLong(this.Type);
	};
	CRunElementBase.prototype.Read_FromBinary = function(Reader)
	{
	};
	/**
	 * Может ли строка начинаться с данного элемента
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.CanBeAtBeginOfLine = function()
	{
		return true;
	};
	/**
	 * Может ли строка заканчиваться данным элементом
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.CanBeAtEndOfLine = function()
	{
		return true;
	};
	/**
	 * Какие мы можем выполнять автозамены на вводе данного элемента
	 * @returns {number}
	 */
	CRunElementBase.prototype.GetAutoCorrectFlags = function()
	{
		return AscWord.AUTOCORRECT_FLAGS_NONE;
	};
	/**
	 * Является ли данный элемент символом пунктуации
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsPunctuation = function()
	{
		return false;
	};
	/**
	 * Проверяем является ли элемент символом точки
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsDot = function()
	{
		return false;
	};
	/**
	 * Проверяем является ли элемент символом знака восклицания
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsExclamationMark = function()
	{
		return false;
	};
	/**
	 * Проверяем является ли элемент символом знака вопроса
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsQuestionMark = function()
	{
		return false;
	};
	/**
	 * Проверяем является ли элемент символом конца предложения
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsSentenceEndMark = function()
	{
		return (this.IsDot() ||this.IsQuestionMark() || this.IsExclamationMark());
	};
	/**
	 * Является ли данный элемент символом дефиса
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsHyphen = function()
	{
		return false;
	};
	/**
	 * @param {CRunElementBase} oElement
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsEqual = function(oElement)
	{
		return (this.Type === oElement.Type)
	};
	/**
	 * Нужно ли ставить разрыв слова после данного элемента
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsSpaceAfter = function()
	{
		return false;
	};
	/**
	 * Нужно ли ставить дефис для автоматического переноса
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.isHyphenAfter = function()
	{
		return false;
	};
	/**
	 * Является ли данный элемент буквой (не цифрой, не знаком пунктуации и т.д.)
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsLetter = function()
	{
		return false;
	};
	/**
	 * Нужно ли сохранять данные этого элемента при сохранении состояния пересчета
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsNeedSaveRecalculateObject = function()
	{
		return false;
	};
	/**
	 * Является ли данный элемент цифрой
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsDigit = function()
	{
		return false;
	};
	/**
	 * Является ли данный элемент пробельным символом
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsSpace = function()
	{
		return false;
	};
	/**
	 * Преобразуем в элемент для поиска
	 * @returns {?CSearchTextItemBase}
	 */
	CRunElementBase.prototype.ToSearchElement = function(oProps)
	{
		return null;
	};
	/**
	 * Преобразуем в элемент матиматического рана
	 * @returns {?CMathBaseText}
	 */
	CRunElementBase.prototype.ToMathElement = function()
	{
		return null;
	};
	/**
	 * Является ли данный элемент автофигурой
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsDrawing = function()
	{
		return false;
	};
	/**
	 * Является ли данный элемент текстовым элементом (но не пробелом и не табом)
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsText = function()
	{
		return false;
	};
	/**
	 * Является ли данный элемент текстовым элементом внутри математического выражения
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsMathText = function()
	{
		return false;
	};
	/**
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsTab = function()
	{
		return false;
	};
	/**
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsParaEnd = function()
	{
		return false;
	};
	/**
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsReference = function()
	{
		return false;
	};
	/**
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsFieldChar = function()
	{
		return false;
	};
	/**
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsInstrText = function()
	{
		return false;
	};
	/**
	 * @returns {AscWord.fontslot_Unknown}
	 */
	CRunElementBase.prototype.GetFontSlot = function(oTextPr)
	{
		return AscWord.fontslot_Unknown;
	};
	/**
	 * Является ли элемент текстом из ComplexScript
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsCS = function()
	{
		return this.GetFontSlot() === AscWord.fontslot_CS;
	};
	/**
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsBreak = function()
	{
		return false;
	};
	/**
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsLigature = function()
	{
		return false;
	};
	/**
	 * @returns {boolean}
	 */
	CRunElementBase.prototype.IsCombiningMark = function()
	{
		return false;
	};
	/**
	 * return {AscBidi.TYPE}
	 */
	CRunElementBase.prototype.getBidiType = function()
	{
		return AscBidi.TYPE.ON;
	};
	/**
	 * @return {number}
	 */
	CRunElementBase.prototype.GetCombWidth = function()
	{
		return this.GetWidth();
	};
	CRunElementBase.prototype.SetGaps = function(nLeftGap, nRightGap, nCellWidth)
	{
	};
	CRunElementBase.prototype.ResetGapBackground = function()
	{
		this.RGapCount     = undefined;
		this.RGapCharCode  = undefined;
		this.RGapCharWidth = undefined;
		this.RGapShift     = undefined;
		this.RGapFontSlot  = undefined;
		this.RGapFont      = undefined;
	};
	CRunElementBase.prototype.SetGapBackground = function(nCount, nCharCode, nCombWidth, oContext, sFont, oTextPr, oTheme, nCombBorderW)
	{
		this.RGapCount    = nCount;
		this.RGapCharCode = nCharCode;
		this.RGapFontSlot = AscWord.GetFontSlotByTextPr(nCharCode, oTextPr);

		if (sFont)
		{
			this.RGapFont = sFont;

			var oCurTextPr = oTextPr.Copy();
			oCurTextPr.SetFontFamily(sFont);

			oContext.SetTextPr(oCurTextPr, oTheme);
			oContext.SetFontSlot(this.RGapFontSlot, oTextPr.getFontCoef());
		}

		this.RGapCharWidth = !nCharCode ? nCombBorderW : Math.max(oContext.MeasureCode(nCharCode).Width + oTextPr.Spacing + nCombBorderW, nCombBorderW);
		this.RGapShift     = Math.max(nCombWidth, this.RGapCharWidth);

		if (sFont)
			oContext.SetTextPr(oTextPr, oTheme);
	};
	CRunElementBase.prototype.DrawGapsBackground = function(X, Y, oGraphics, PDSE, oTextPr)
	{
		if (!this.RGapCharCode)
			return;

		if (this.RGapFont)
		{
			let oCurTextPr = oTextPr.Copy();
			oCurTextPr.SetFontFamily(this.RGapFont);

			oGraphics.SetTextPr(oCurTextPr, PDSE.Theme);
			oGraphics.SetFontSlot(this.RGapFontSlot, oTextPr.getFontCoef());
		}

		if (this.RGap && this.RGapCount)
		{
			X += this.GetWidth();

			let nShift = (this.RGapShift - this.RGapCharWidth) / 2;
			for (let nIndex = 0; nIndex < this.RGapCount; ++nIndex)
			{
				X -= nShift + this.RGapCharWidth;
				oGraphics.FillTextCode(X, Y, this.RGapCharCode);
				X -= nShift;
			}
		}

		if (this.RGapFont)
			oGraphics.SetTextPr(oTextPr, PDSE.Theme);
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].CRunElementBase   = CRunElementBase;
	window['AscWord'].TEXTWIDTH_DIVIDER = TEXTWIDTH_DIVIDER;

})(window);

