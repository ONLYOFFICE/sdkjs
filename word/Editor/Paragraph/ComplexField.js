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

var fldchartype_Begin    = 0;
var fldchartype_Separate = 1;
var fldchartype_End      = 2;

function ParaFieldChar(Type, LogicDocument)
{
	AscWord.CRunElementBase.call(this);
	
	this.LogicDocument = LogicDocument;
	this.Use           = true;
	this.CharType      = undefined === Type ? fldchartype_Begin : Type;
	this.ComplexField  = (this.CharType === fldchartype_Begin) ? new CComplexField(LogicDocument) : null;
	this.fldData       = null;
	this.ffData        = null;
	this.Run           = null;
	this.X             = 0;
	this.Y             = 0;
	this.PageAbs       = 0;
	
	this.showFieldCode = false;

	this.numText  = null;
	this.textPr   = null;
	this.checkBox = null;
	this.hidden   = false;
	
	this.graphemes = [];
	this.widths    = [];
	this.ascent    = 0;
	this.descent   = 0;
}
ParaFieldChar.prototype = Object.create(AscWord.CRunElementBase.prototype);
ParaFieldChar.prototype.constructor = ParaFieldChar;
ParaFieldChar.prototype.Type = para_FieldChar;
ParaFieldChar.prototype.IsFieldChar = function()
{
	return true;
};
ParaFieldChar.prototype.Init = function(Type, LogicDocument)
{
	this.CharType = Type;

	this.LogicDocument = LogicDocument;
	this.ComplexField  = (this.CharType === fldchartype_Begin) ? new CComplexField(this.LogicDocument) : null;
};
ParaFieldChar.prototype.Copy = function()
{
	let oChar = new ParaFieldChar(this.CharType, this.LogicDocument);

	let oComplexField = this.GetComplexField();
	if (oComplexField && oComplexField.IsUpdate())
	{
		oChar.SetComplexField(oComplexField);
		oComplexField.ReplaceChar(oChar);
	}
	//todo fldData
	oChar.ffData = this.ffData ? this.ffData.Copy() : null;

	return oChar;
};
ParaFieldChar.prototype.Measure = function(Context, textPr, sectPr)
{
	if (!this.IsEnd())
		return;
	
	this.textPr = textPr;
	this.sectPr = sectPr;
	this.format = Asc.c_oAscNumberingFormat.UpperRoman;
	
	this.private_UpdateWidth();
};
ParaFieldChar.prototype.Draw = function(x, y, context)
{
	if (!this.IsVisual())
		return;
	
	if (this.numText)
	{
		let fontSize = this.textPr.FontSize * this.textPr.getFontCoef();
		for (let index = 0; index < this.graphemes.length; ++index)
		{
			AscFonts.DrawGrapheme(this.graphemes[index], context, x, y, fontSize);
			x += this.widths[index] * fontSize;
		}
	}
	else if (this.checkBox)
	{
		let shift = 0.75 * g_dKoef_pt_to_mm;
		let penW  = 0.75 * g_dKoef_pt_to_mm;
		
		let width   = this.GetWidth();
		let ascent  = this.ascent / AscWord.TEXTWIDTH_DIVIDER;
		let descent = this.descent / AscWord.TEXTWIDTH_DIVIDER;
		
		let y0 = y - ascent + shift;
		let y1 = y + descent - shift;
		let x0 = x + shift;
		let x1 = x + width - shift;
		
		context.drawHorLineExt(c_oAscLineDrawingRule.Top, y0, x0, x1, penW, 0, 0);
		context.drawHorLineExt(c_oAscLineDrawingRule.Bottom, y1, x0, x1, penW, 0, 0);
		context.drawVerLine(c_oAscLineDrawingRule.Left, x0, y0, y1, penW);
		context.drawVerLine(c_oAscLineDrawingRule.Right, x1, y0, y1, penW);
		
		let ffData  = this.ComplexField.GetBeginChar().GetFFData();
		if (ffData && ffData.isCheckBoxChecked())
		{
			context.p_width(0.5 * g_dKoef_pt_to_mm * 1000);
			let penW_2 = penW / 2;
			context._m(x0 + penW_2, y0 + penW_2);
			context._l(x1 - penW_2, y1 - penW_2);
			context.ds();
			
			context._m(x1 - penW_2, y0 + penW_2);
			context._l(x0 + penW_2, y1 - penW_2);
			context.ds();
		}
	}
};
ParaFieldChar.prototype.IsBegin = function()
{
	return (this.CharType === fldchartype_Begin ? true : false);
};
ParaFieldChar.prototype.IsEnd = function()
{
	return (this.CharType === fldchartype_End ? true : false);
};
ParaFieldChar.prototype.IsSeparate = function()
{
	return (this.CharType === fldchartype_Separate ? true : false);
};
ParaFieldChar.prototype.IsUse = function()
{
	return this.Use;
};
ParaFieldChar.prototype.SetUse = function(isUse)
{
	this.Use = isUse;
};
ParaFieldChar.prototype.GetComplexField = function()
{
	return this.ComplexField;
};
ParaFieldChar.prototype.SetComplexField = function(oComplexField)
{
	this.ComplexField = oComplexField;
};
ParaFieldChar.prototype.Write_ToBinary = function(writer)
{
	// Long : Type
	// Long : CharType
	writer.WriteLong(this.Type);
	writer.WriteLong(this.CharType);
	
	if (this.ffData)
	{
		writer.WriteBool(true);
		this.ffData.toBinary(writer);
	}
	else
	{
		writer.WriteBool(false);
	}
	
	//todo fldData
};
ParaFieldChar.prototype.Read_FromBinary = function(reader)
{
	// Long : CharType
	
	let charType = reader.GetLong();
	
	this.Init(charType, editor.WordControl.m_oLogicDocument);
	
	if (reader.GetBool())
		this.ffData = AscWord.FFData.fromBinary(reader);
	
	//todo fldData
};
ParaFieldChar.prototype.SetParent = function(oParent)
{
	this.Run = oParent;
};
ParaFieldChar.prototype.SetRun = function(oRun)
{
	this.Run = oRun;
};
ParaFieldChar.prototype.GetRun = function()
{
	return this.Run;
};
ParaFieldChar.prototype.GetParagraph = function()
{
	return this.Run ? this.Run.GetParagraph() : null;
};
ParaFieldChar.prototype.SetXY = function(X, Y)
{
	this.X = X;
	this.Y = Y;
};
ParaFieldChar.prototype.GetXY = function()
{
	return {X : this.X, Y : this.Y};
};
ParaFieldChar.prototype.SetPage = function(nPage)
{
	this.PageAbs = nPage;
};
ParaFieldChar.prototype.GetPage = function()
{
	return this.PageAbs;
};
ParaFieldChar.prototype.GetTopDocumentContent = function()
{
	if (!this.Run)
		return null;

	var oParagraph = this.Run.GetParagraph();
	if (!oParagraph)
		return null;

	return oParagraph.Parent.GetTopDocumentContent();
};
/**
 * Специальная функция для работы с полями PAGE NUMPAGES в колонтитулах
 * @param value {number}
 * @param numFormat {Asc.c_oAscNumberingFormat}
 */
ParaFieldChar.prototype.SetNumValue = function(value, numFormat)
{
	if (null === value)
	{
		this.numText = null;
		return;
	}
	
	this.numText = AscCommon.IntToNumberFormat(value, numFormat, {lang: this.textPr && this.textPr.Lang, isFromField: true, isSkipFractPart: true});
	this.private_UpdateWidth();
};
ParaFieldChar.prototype.SetHiddenValue = function(isHidden)
{
	this.hidden = isHidden;
};
ParaFieldChar.prototype.IsHiddenValue = function()
{
	return this.hidden;
};
/**
 * Специальная функция для работы с полями FORUMULA в колонтитулах
 * @param value {number|string}
 */
ParaFieldChar.prototype.SetFormulaValue = function(value)
{
	if (null === value)
	{
		this.numText = null;
		return;
	}
	
	this.numText = "" + value;
	this.private_UpdateWidth();
};
ParaFieldChar.prototype.GetFFData = function()
{
	return this.ffData;
};
ParaFieldChar.prototype.SetFormCheckBox = function()
{
	this.checkBox = true;
	this.private_UpdateWidth();
};
ParaFieldChar.prototype.GetNumFormat = function()
{
	let numFormat = Asc.c_oAscNumberingFormat.Decimal;
	let cf = this.ComplexField;
	let instruction = cf.IsValid() ? cf.GetInstruction() : null;
	if (instruction && instruction.haveNumericFormat())
		numFormat = instruction.getNumericFormat();
	
	return numFormat;
};
ParaFieldChar.prototype.UpdatePageCount = function(pageCount)
{
	let cf = this.ComplexField;
	if (!cf)
		return;
	
	let instruction = cf.GetInstruction();
	if (!instruction)
		return;
	
	let fieldType = instruction.GetType();
	if (fieldType === AscWord.fieldtype_FORMULA)
	{
		let value =  parseInt(cf.CalculateValue());
		if (isNaN(value))
			value = 0;
		
		this.SetFormulaValue(value);
	}
	else if (fieldType === AscWord.fieldtype_NUMPAGES)
	{
		this.SetNumValue(pageCount, this.GetNumFormat());
	}
};
ParaFieldChar.prototype.private_UpdateWidth = function()
{
	let totalWidth = 0;
	
	if (this.numText)
	{
		AscWord.stringShaper.Shape(this.numText.codePointsArray(), this.textPr);
		
		this.graphemes = AscWord.stringShaper.GetGraphemes();
		this.widths    = AscWord.stringShaper.GetWidths();
		
		for (let index = 0; index < this.widths.length; ++index)
		{
			totalWidth += this.widths[index];
		}
		let fontSize = this.textPr.FontSize * this.textPr.getFontCoef();
		totalWidth   = (totalWidth * fontSize * AscWord.TEXTWIDTH_DIVIDER) | 0;
	}
	else if (this.checkBox)
	{
		let fontSize = this.textPr.FontSize * this.textPr.getFontCoef();
		let ffData = this.ComplexField.GetBeginChar().GetFFData();
		if (ffData && !ffData.isCheckBoxAutoSize())
			fontSize = ffData.getCheckBoxSize();
		
		let textPr = this.textPr.Copy();
		textPr.FontSize = fontSize;
		
		AscCommon.g_oTextMeasurer.SetTextPr(textPr);
		AscCommon.g_oTextMeasurer.SetFontSlot(AscWord.fontslot_ASCII, 1);
		
		let textHeight = AscCommon.g_oTextMeasurer.GetHeight();
		totalWidth = (textHeight * AscWord.TEXTWIDTH_DIVIDER) | 0;
		
		// Для совместимости при работе с RecalcObject
		this.graphemes = [];
		this.widths    = [];
		this.ascent    = (AscCommon.g_oTextMeasurer.GetAscender() * AscWord.TEXTWIDTH_DIVIDER) | 0;
		this.descent   = (-AscCommon.g_oTextMeasurer.GetDescender() * AscWord.TEXTWIDTH_DIVIDER) | 0;
	}
	
	this.Width        = totalWidth;
	this.WidthVisible = totalWidth;
};
ParaFieldChar.prototype.IsNumValue = function()
{
	return (this.IsEnd() && null !== this.numText);
};
ParaFieldChar.prototype.IsVisual = function()
{
	return (this.IsEnd() && (null !== this.numText || null !== this.checkBox));
};
ParaFieldChar.prototype.IsFormField = function()
{
	return !!(this.checkBox);
};
ParaFieldChar.prototype.IsNeedSaveRecalculateObject = function()
{
	return this.IsVisual();
};
ParaFieldChar.prototype.SaveRecalculateObject = function(isCopy)
{
	return new FieldCharRecalculateObject(this.Type, this.numText, this.checkBox, this.graphemes, this.widths, this.Width, this.ascent, this.descent, isCopy);
};
ParaFieldChar.prototype.LoadRecalculateObject = function(recalcObj)
{
	this.numText      = recalcObj.numText;
	this.checkBox     = recalcObj.checkBox;
	this.graphemes    = recalcObj.graphemes;
	this.widths       = recalcObj.widths;
	this.Width        = recalcObj.width;
	this.WidthVisible = this.Width;
	this.ascent       = recalcObj.ascent;
	this.descent      = recalcObj.descent;
};
ParaFieldChar.prototype.PrepareRecalculateObject = function()
{
	this.graphemes = [];
	this.widths    = [];
	this.checkBox  = null;
	this.numText   = null;
};
ParaFieldChar.prototype.IsValid = function()
{
	var oRun = this.GetRun();
	return (oRun && oRun.IsUseInDocument() && -1 !== oRun.GetElementPosition(this));
};
ParaFieldChar.prototype.RemoveThisFromDocument = function()
{
	var oRun = this.GetRun();
	var nInRunPos = oRun.GetElementPosition(this);
	if (-1 !== nInRunPos)
		oRun.RemoveFromContent(nInRunPos, 1);
};
ParaFieldChar.prototype.PreDelete = function()
{
	if (this.LogicDocument && this.ComplexField)
		this.LogicDocument.ValidateComplexField(this.ComplexField);
};
ParaFieldChar.prototype.FindNextFillingForm = function(isNext, isCurrent, isStart)
{
	if (!this.ComplexField
		|| !this.ComplexField.IsFormField()
		|| !this.ComplexField.IsFormFieldEnabled())
		return null;
	
	if (isNext)
		return (this.IsBegin() && (!isCurrent || isNext) ? this.ComplexField : null);
	else
		return (this.IsEnd() && (!isCurrent || isNext) ? this.ComplexField : null);
};
ParaFieldChar.prototype.IsShowFieldCode = function()
{
	return this.showFieldCode;
};
ParaFieldChar.prototype.SetShowFieldCode = function(isShow)
{
	this.showFieldCode = isShow;
};
ParaFieldChar.prototype.MoveCursorToChar = function(isBefore)
{
	let run = this.GetRun();
	if (!run)
		return;
	let inRunPos = run.GetElementPosition(this);
	if (-1 === inRunPos)
		return;
	
	if (this.LogicDocument)
		this.LogicDocument.RemoveSelection();
	
	if (false === isBefore)
		inRunPos += 1
	
	run.Make_ThisElementCurrent(false);
	run.SetCursorPosition(inRunPos);
};

/**
 * @constructor
 */
function FieldCharRecalculateObject(type, numText, checkBox, graphemes, widths, totalWidth, ascent, descent, isCopy)
{
	this.type      = type;
	this.numText   = numText;
	this.checkBox  = checkBox;
	this.graphemes = graphemes && isCopy ? graphemes.slice() : graphemes;
	this.widths    = widths && isCopy ? widths.slice() : widths;
	this.width     = totalWidth;
	this.ascent    = ascent;
	this.descent   = descent;
}

/**
 * Класс представляющий символ инструкции сложного поля
 * @param {Number} nCharCode
 * @constructor
 */
function ParaInstrText(nCharCode)
{
	AscWord.CRunElementBase.call(this);

	this.Value        = (undefined !== nCharCode ? nCharCode : 0x00);
	this.Width        = 0x00000000 | 0;
	this.WidthVisible = 0x00000000 | 0;
	this.Run          = null;
	this.Replacement  = null; // Используется, когда InstrText идет в неположенном месте и должно восприниматься как обычный текст
}
ParaInstrText.prototype = Object.create(AscWord.CRunElementBase.prototype);
ParaInstrText.prototype.constructor = ParaInstrText;
ParaInstrText.prototype.Type = para_InstrText;
ParaInstrText.prototype.IsInstrText = function()
{
	return true;
};
ParaInstrText.prototype.Copy = function()
{
	return new ParaInstrText(this.Value);
};
ParaInstrText.prototype.Measure = function(Context, TextPr)
{
};
ParaInstrText.prototype.Draw = function(X, Y, Context)
{
};
ParaInstrText.prototype.Write_ToBinary = function(Writer)
{
	// Long : Type
	// Long : Value
	Writer.WriteLong(this.Type);
	Writer.WriteLong(this.Value);
};
ParaInstrText.prototype.Read_FromBinary = function(Reader)
{
	// Long : Value
	this.Value = Reader.GetLong();
};
ParaInstrText.prototype.SetParent = function(oParent)
{
	this.Run = oParent;
};
ParaInstrText.prototype.SetRun = function(oRun)
{
	this.Run = oRun;
};
ParaInstrText.prototype.GetRun = function()
{
	return this.Run;
};
ParaInstrText.prototype.GetValue = function()
{
	return String.fromCharCode(this.Value);
};
ParaInstrText.prototype.GetCodePoint = function()
{
	return this.Value;
};
ParaInstrText.prototype.GetCharCode = function()
{
	return this.Value;
};
ParaInstrText.prototype.SetCharCode = function(CharCode)
{
	this.Value = CharCode;
};
ParaInstrText.prototype.SetReplacementItem = function(oItem)
{
	this.Replacement = oItem;
};
ParaInstrText.prototype.GetReplacementItem = function()
{
	return this.Replacement;
};

function CComplexField(logicDocument)
{
	this.LogicDocument   = logicDocument;
	this.Current         = false;
	this.BeginChar       = null;
	this.EndChar         = null;
	this.SeparateChar    = null;
	this.InstructionLine = "";
	this.Instruction     = null;
	this.FieldId         = logicDocument && logicDocument.IsDocumentEditor() ? logicDocument.GetFieldsManager().GetNewComplexFieldId() : null;

	this.InstructionLineSrc = "";
	this.InstructionCF      = [];

	this.StartUpdate = false;
}
CComplexField.prototype.GetFieldId = function()
{
	return this.FieldId;
};
CComplexField.prototype.SetCurrent = function(isCurrent)
{
	this.Current = isCurrent;
};
CComplexField.prototype.IsCurrent = function()
{
	return this.Current;
};
CComplexField.prototype.IsUpdate = function()
{
	return (this.StartUpdate > 0);
};
CComplexField.prototype.StartCharsUpdate = function()
{
	++this.StartUpdate;
};
CComplexField.prototype.FinishCharsUpdate = function()
{
	if (this.StartUpdate > 0)
		--this.StartUpdate;
};
CComplexField.prototype.SetInstruction = function(oParaInstr)
{
	this.InstructionLine += oParaInstr.GetValue();
	this.InstructionLineSrc += oParaInstr.GetValue();
};
CComplexField.prototype.SetInstructionCF = function(oCF)
{
	this.InstructionLine += " \\& ";
	this.InstructionLineSrc +=  " \\& ";
	this.InstructionCF.push(oCF);
};
CComplexField.prototype.SetInstructionLine = function(sLine)
{
	this.InstructionLine = sLine;
};
CComplexField.prototype.GetInstructionLine = function()
{
	return this.InstructionLine;
};
CComplexField.prototype.GetBeginChar = function()
{
	return this.BeginChar;
};
CComplexField.prototype.GetEndChar = function()
{
	return this.EndChar;
};
CComplexField.prototype.GetSeparateChar = function()
{
	return this.SeparateChar;
};
CComplexField.prototype.SetBeginChar = function(oChar)
{
	oChar.SetComplexField(this);

	this.BeginChar       = oChar;
	this.SeparateChar    = null;
	this.EndChar         = null;
	this.InstructionLine = "";

	this.InstructionLineSrc = "";
	this.InstructionCF      = [];
};
CComplexField.prototype.SetEndChar = function(oChar)
{
	oChar.SetComplexField(this);

	this.EndChar = oChar;
};
CComplexField.prototype.SetSeparateChar = function(oChar)
{
	oChar.SetComplexField(this);

	this.SeparateChar = oChar;
	this.EndChar      = null;
};
CComplexField.prototype.ReplaceChar = function(oChar)
{
	oChar.SetComplexField(this);

	if (oChar.IsBegin())
		this.BeginChar = oChar;
	else if (oChar.IsSeparate())
		this.SeparateChar = oChar;
	else if (oChar.IsEnd())
		this.EndChar = oChar;
};
CComplexField.prototype.Update = function(isCreateHistoryPoint, isNeedRecalculate)
{
	this.private_CheckNestedComplexFields();
	this.private_UpdateInstruction();

	if (!this.Instruction || !this.IsValid())
		return false;
	
	// TODO: Нужно добавить разделитель, если его нет. Пока не обновляем такие поля
	if (!this.HaveValuePart())
		return false;

	this.SelectFieldValue();

	if (true === isCreateHistoryPoint)
	{
		if (true === this.LogicDocument.Document_Is_SelectionLocked(changestype_Paragraph_Content))
			return false;

		this.LogicDocument.StartAction();
	}

	this.StartCharsUpdate();
	switch (this.Instruction.GetType())
	{
		case AscWord.fieldtype_PAGE:
		case AscWord.fieldtype_PAGENUM:
			this.private_UpdatePAGE();
			break;
		case AscWord.fieldtype_TOC:
			this.private_UpdateTOC();
			break;
		case AscWord.fieldtype_PAGEREF:
			this.private_UpdatePAGEREF();
			break;
		case AscWord.fieldtype_NUMPAGES:
		case AscWord.fieldtype_PAGECOUNT:
			this.private_UpdateNUMPAGES();
			break;
		case AscWord.fieldtype_FORMULA:
			this.private_UpdateFORMULA();
			break;
		case AscWord.fieldtype_SEQ:
			this.private_UpdateSEQ();
			break;
		case AscWord.fieldtype_STYLEREF:
			this.private_UpdateSTYLEREF();
			break;
		case AscWord.fieldtype_TIME:
		case AscWord.fieldtype_DATE:
			this.private_UpdateTIME();
			break;
		case AscWord.fieldtype_REF:
			this.private_UpdateREF();
			break;
		case AscWord.fieldtype_NOTEREF:
			this.private_UpdateNOTEREF();
			break;
		case AscWord.fieldtype_FORMCHECKBOX:
		case AscWord.fieldtype_FORMTEXT:
		case AscWord.fieldtype_ADDIN:
			break;
	}
	this.FinishCharsUpdate();

	if (false !== isNeedRecalculate)
		this.LogicDocument.Recalculate();

	if (isCreateHistoryPoint)
	{
		this.LogicDocument.FinalizeAction();
	}
	
	return true;
};
CComplexField.prototype.CalculateValue = function()
{
	this.private_CheckNestedComplexFields();
	this.private_UpdateInstruction();

	if (!this.Instruction || !this.BeginChar || !this.EndChar || !this.SeparateChar)
		return;

	var sResult = "";
	switch (this.Instruction.GetType())
	{
		case AscWord.fieldtype_PAGE:
		case AscWord.fieldtype_PAGENUM:
			sResult = this.private_CalculatePAGE();
			break;
		case AscWord.fieldtype_TOC:
			sResult = "";
			break;
		case AscWord.fieldtype_PAGEREF:
			sResult = this.private_CalculatePAGEREF();
			break;
		case AscWord.fieldtype_NUMPAGES:
		case AscWord.fieldtype_PAGECOUNT:
			sResult = this.private_CalculateNUMPAGES();
			break;
		case AscWord.fieldtype_FORMULA:
			sResult = this.private_CalculateFORMULA();
			break;
		case AscWord.fieldtype_SEQ:
			sResult = this.private_CalculateSEQ();
			break;
		case AscWord.fieldtype_STYLEREF:
			sResult = this.private_CalculateSTYLEREF();
			break;
		case AscWord.fieldtype_TIME:
		case AscWord.fieldtype_DATE:
			sResult = this.private_CalculateTIME();
			break;
		case AscWord.fieldtype_REF:
			sResult = this.private_CalculateREF();
			break;
		case AscWord.fieldtype_NOTEREF:
			sResult = this.private_CalculateNOTEREF();
			break;
		case AscWord.fieldtype_FORMCHECKBOX:
		case AscWord.fieldtype_FORMTEXT:
		case AscWord.fieldtype_ADDIN:
			sResult = "";
			break;
	}

	return sResult;
};
CComplexField.prototype.UpdateTIME = function(ms)
{
	this.private_CheckNestedComplexFields();
	this.private_UpdateInstruction();

	if (!this.Instruction
		|| !this.BeginChar
		|| !this.EndChar
		|| !this.SeparateChar
		|| (AscWord.fieldtype_TIME !== this.Instruction.GetType() && AscWord.fieldtype_DATE !== this.Instruction.GetType()))
		return;

	this.SelectFieldValue();
	this.private_UpdateTIME(ms);
};
CComplexField.prototype.IsHaveNestedNUMPAGES = function()
{
	for (let index = 0, count = this.InstructionCF.length; index < count; ++index)
	{
		let instruction = this.InstructionCF[index].GetInstruction();
		if (instruction && AscWord.fieldtype_NUMPAGES === instruction.GetType())
			return true;
		
		if (this.InstructionCF[index].IsHaveNestedNUMPAGES())
			return true;
	}
	
	return false;
};

CComplexField.prototype.private_UpdateSEQ = function()
{
	this.LogicDocument.AddText(this.private_CalculateSEQ());
};
CComplexField.prototype.private_CalculateSEQ = function()
{
	return this.Instruction.GetText();
};
CComplexField.prototype.private_UpdateSTYLEREF = function()
{
	this.LogicDocument.AddText(this.private_CalculateSTYLEREF());
};
CComplexField.prototype.private_CalculateSTYLEREF = function()
{
	return this.Instruction.GetText();
};
CComplexField.prototype.private_InsertMessage = function(sMessage, oTextPr)
{
	var oSelectedContent = new AscCommonWord.CSelectedContent();
	var oPara = new AscWord.Paragraph();
	var oRun  = new ParaRun(oPara, false);
	if(oTextPr)
	{
		oRun.Apply_Pr(oTextPr);
	}
	oRun.AddText(sMessage);
	oPara.AddToContent(0, oRun);
	oSelectedContent.Add(new AscCommonWord.CSelectedElement(oPara, false));
	this.private_InsertContent(oSelectedContent);
};
CComplexField.prototype.private_InsertContent = function(oSelectedContent)
{
	this.SelectFieldValue();
	this.LogicDocument.ConcatParagraphsOnRemove = true;
	this.LogicDocument.Remove(1, false, false, true);
	this.LogicDocument.ConcatParagraphsOnRemove = false;

	var oRun       = this.BeginChar.GetRun();
	var oParagraph = oRun.GetParagraph();
	if (oParagraph)
	{
		var oAnchorPos = oParagraph.GetCurrentAnchorPosition();
		if (oAnchorPos && oSelectedContent.CanInsert(oAnchorPos))
		{
			oParagraph.Check_NearestPos(oAnchorPos);
			oSelectedContent.ForceInlineInsert();
			oSelectedContent.Insert(oAnchorPos);
			this.MoveCursorOutsideElement(false);
		}
	}
};
CComplexField.prototype.private_UpdateFORMULA = function()
{
	this.Instruction.Calculate(this.LogicDocument);
	if(this.Instruction.ErrStr !== null)
	{
		var oTextPr = new CTextPr();
		oTextPr.Set_FromObject({Bold: true});
		this.private_InsertMessage(this.Instruction.ErrStr, oTextPr);
	}
	else
	{
		if(this.Instruction.ResultStr !== null)
		{
			this.private_InsertMessage(this.Instruction.ResultStr, null);
		}
	}
};
CComplexField.prototype.private_CalculateFORMULA = function()
{
	this.Instruction.Calculate(this.LogicDocument);
	if (null !== this.Instruction.ErrStr)
		return this.Instruction.ErrStr;

	return this.Instruction.ResultStr;
};
CComplexField.prototype.private_UpdatePAGE = function()
{
	this.LogicDocument.AddText("" + this.private_CalculatePAGE());
};
CComplexField.prototype.private_CalculatePAGE = function()
{
	var oRun       = this.BeginChar.GetRun();
	var oParagraph = oRun.GetParagraph();
	var nInRunPos  = oRun.GetElementPosition(this.BeginChar);
	var nLine      = oRun.GetLineByPosition(nInRunPos);
	var nPage      = oParagraph.getPageByLine(nLine);

	var oLogicDocument = oParagraph.LogicDocument;
	return oLogicDocument.Get_SectionPageNumInfo2(oParagraph.Get_AbsolutePage(nPage)).CurPage;
};
CComplexField.prototype.private_UpdateTOC = function()
{
	this.LogicDocument.GetBookmarksManager().RemoveTOCBookmarks();

	var nTabPos = 9345 / 20 / 72 * 25.4; // Стандартное значение для A4 и обычных полей 3см и 2см
	var oSectPr = this.LogicDocument.GetCurrentSectionPr();
	if (oSectPr)
	{
		if (oSectPr.Get_ColumnsCount() > 1)
		{
			// TODO: Сейчас забирается ширина текущей колонки. По правильному надо читать поля от текущего места
			nTabPos = Math.max(0, Math.min(oSectPr.GetColumnWidth(0), oSectPr.GetPageWidth(), oSectPr.GetContentFrameWidth()));
		}
		else
		{
			nTabPos = Math.max(0, Math.min(oSectPr.GetPageWidth(), oSectPr.GetContentFrameWidth()));
		}
	}

	var oStyles          = this.LogicDocument.Get_Styles();
	var arrOutline;
	var sCaption         = this.Instruction.GetCaption(); //flag c
	var sCaptionOnlyText = this.Instruction.GetCaptionOnlyText();//flag a
	
	var oBookmarksManager = this.LogicDocument.GetBookmarksManager();
	if (undefined === sCaption && undefined !== sCaptionOnlyText)
		sCaption = sCaptionOnlyText;
	
	var oOutlinePr = {
		OutlineStart : this.Instruction.GetHeadingRangeStart(),
		OutlineEnd   : this.Instruction.GetHeadingRangeEnd(),
		Styles       : this.Instruction.GetStylesArray()
	};
	var bTOF = false;
	var bSkipCaptionLbl = false;
	if (undefined !== sCaption)
	{
		bTOF = true;
		var aStyles = this.Instruction.GetStylesArray();
		if(aStyles.length > 0)
		{
			arrOutline = this.LogicDocument.GetOutlineParagraphs(null, oOutlinePr);
		}
		else
		{
			arrOutline = [];
			if (sCaptionOnlyText !== undefined)
				bSkipCaptionLbl = true;
			
			var aParagraphs = this.LogicDocument.GetAllCaptionParagraphs(sCaption);
			var oCurPara;
			for (var nParagraph = 0; nParagraph < aParagraphs.length; ++nParagraph)
			{
				oCurPara = aParagraphs[nParagraph];
				if (!bSkipCaptionLbl || oCurPara.CanAddRefAfterSEQ(sCaption))
					arrOutline.push({Paragraph : oCurPara, Lvl : 0});
			}
		}
	}
	else
	{
		arrOutline = this.LogicDocument.GetOutlineParagraphs(null, oOutlinePr);
	}
	var oSelectedContent = new AscCommonWord.CSelectedContent();

	var isPreserveTabs   = this.Instruction.IsPreserveTabs();
	var sSeparator       = this.Instruction.GetSeparator();

	var nForceTabLeader = this.Instruction.GetForceTabLeader();

	var oTab = new CParaTab(tab_Right, nTabPos, Asc.c_oAscTabLeader.Dot);
	var oPara, oTabs, oRun;
	if (undefined !== nForceTabLeader)
	{
		oTab = new CParaTab(tab_Right, nTabPos, nForceTabLeader);
	}
	else if ((!sSeparator || "" === sSeparator) && arrOutline.length > 0)
	{
		var arrSelectedParagraphs = this.LogicDocument.GetCurrentParagraph(false, true);
		if (arrSelectedParagraphs.length > 0)
		{
			oPara = arrSelectedParagraphs[0];
			oTabs = oPara.GetParagraphTabs();

			if (oTabs.Tabs.length > 0)
				oTab = oTabs.Tabs[oTabs.Tabs.length - 1];
		}
	}

	if (arrOutline.length > 0)
	{
		for (var nIndex = 0, nCount = arrOutline.length; nIndex < nCount; ++nIndex)
		{
			var oSrcParagraph = arrOutline[nIndex].Paragraph;
			var sBookmarkName;
			var oParaForCopy = oSrcParagraph;
			if(bSkipCaptionLbl)
			{
				sBookmarkName = oSrcParagraph.AddBookmarkForCaption(sCaption, true, true);
				if(!sBookmarkName)
				{
					sBookmarkName = oSrcParagraph.AddBookmarkForTOC();
				}
				oBookmarksManager.SelectBookmark(sBookmarkName);
				var oParaSelectedContent = this.LogicDocument.GetSelectedContent(false);
				var oElement = oParaSelectedContent.Elements[0];
				if(oElement && oElement.Element.GetType() === type_Paragraph)
				{
					oParaForCopy = oElement.Element;
				}
			}
			else
			{
				sBookmarkName = oSrcParagraph.AddBookmarkForTOC();
			}
			oPara = oParaForCopy.Copy(null, null, {
				SkipPageBreak         : true,
				SkipLineBreak         : this.Instruction.IsRemoveBreaks(),
				SkipColumnBreak       : true,
				SkipAnchors           : true,
				SkipFootnoteReference : true,
				SkipComplexFields     : true,
				SkipComments          : true,
				SkipBookmarks         : true,
				CopyReviewPr          : false
			});
			oPara.RemovePrChange();
			if(bTOF)
			{
				oPara.Style_Add(oStyles.GetDefaultTOF(), false);
			}
			else
			{
				oPara.Style_Add(oStyles.GetDefaultTOC(arrOutline[nIndex].Lvl), false);
			}
			oPara.SetOutlineLvl(undefined);

			var oClearTextPr = new CTextPr();
			oClearTextPr.Set_FromObject({
				FontSize  : null,
				Unifill   : null,
				Underline : null,
				Color     : null
			});

			// Дополнительно очищаем текстовые настройки, которые были заданы непосредственно в самом стиле
			var oSrcPStylePr = oStyles.Get_Pr(oSrcParagraph.Style_Get(), styletype_Paragraph, oSrcParagraph.Parent.Get_TableStyleForPara(), oSrcParagraph.Parent.Get_ShapeStyleForPara()).TextPr;
			var oDefaultPr   = oStyles.Get_Pr(oStyles.GetDefaultParagraph(), styletype_Paragraph, null, null).TextPr;

			if (oSrcPStylePr.Bold !== oDefaultPr.Bold)
			{
				oClearTextPr.Bold   = null;
				oClearTextPr.BoldCS = null;
			}

			if (oSrcPStylePr.Italic !== oDefaultPr.Italic)
			{
				oClearTextPr.Italic   = null;
				oClearTextPr.ItalicCS = null;
			}

			oPara.SelectAll();
			oPara.ApplyTextPr(oClearTextPr);
			oPara.RemoveSelection();



			var oContainer    = oPara,
				nContainerPos = 0;

			if (this.Instruction.IsHyperlinks())
			{
				var sHyperlinkStyleId = oStyles.GetDefaultHyperlink();

				var oHyperlink = new ParaHyperlink();
				for (var nParaPos = 0, nParaCount = oPara.Content.length - 1; nParaPos < nParaCount; ++nParaPos)
				{
					// TODO: Проверить, нужно ли проставлять этот стиль во внутренние раны
					if (oPara.Content[0] instanceof ParaRun)
						oPara.Content[0].Set_RStyle(sHyperlinkStyleId);

					oHyperlink.Add_ToContent(nParaPos, oPara.Content[0]);
					oPara.Remove_FromContent(0, 1);
				}
				oHyperlink.SetAnchor(sBookmarkName);
				oPara.Add_ToContent(0, oHyperlink);

				oContainer    = oHyperlink;
				nContainerPos = oHyperlink.Content.length;
			}
			else
			{
				// TODO: ParaEnd
				oContainer    = oPara;
				nContainerPos = oPara.Content.length - 1;
			}

			let numTabPos = null;
			if (oSrcParagraph.HaveNumbering() && oSrcParagraph.GetParent())
			{
				var oNumPr     = oSrcParagraph.GetNumPr();
				var oNumbering = this.LogicDocument.GetNumbering();
				var oNumInfo   = oSrcParagraph.GetParent().CalculateNumberingValues(oSrcParagraph, oNumPr);
				var sText      = oNumbering.GetText(oNumPr.NumId, oNumPr.Lvl, oNumInfo);
				var oNumTextPr = oSrcParagraph.GetNumberingCompiledPr();

				var oNumberingRun = new ParaRun(oPara, false);
				oNumberingRun.AddText(sText);

				if (oNumTextPr)
					oNumberingRun.Set_RFonts(oNumTextPr.RFonts);

				oContainer.Add_ToContent(0, oNumberingRun);
				nContainerPos++;

				var oNumTabRun = new ParaRun(oPara, false);

				var oNumLvl  = oNumbering.GetNum(oNumPr.NumId).GetLvl(oNumPr.Lvl);
				var nNumSuff = oNumLvl.GetSuff();

				if (Asc.c_oAscNumberingSuff.Space === nNumSuff)
				{
					oNumTabRun.Add_ToContent(0, new AscWord.CRunSpace());
					oContainer.Add_ToContent(1, oNumTabRun);
					nContainerPos++;
				}
				else if (Asc.c_oAscNumberingSuff.Tab === nNumSuff)
				{
					// Выставление родительского класса нужно для правильного расчета ширины рана с учетом того,
					// что используемые шрифты могут быть заданы в темах
					oPara.SetParent(oSrcParagraph.GetParent());
					
					numTabPos = 0;
					AscWord.ParagraphTextShaper.ShapeRun(oNumberingRun);
					for (let runItemIndex = 0; runItemIndex < oNumberingRun.GetElementsCount(); ++runItemIndex)
						numTabPos += oNumberingRun.GetElement(runItemIndex).GetWidth();
					
					oNumTabRun.Add_ToContent(0, new AscWord.CRunTab());
					oContainer.Add_ToContent(1, oNumTabRun);
					nContainerPos++;
				}
			}

			// Word добавляет табы независимо о наличия Separator и PAGEREF
			oTabs = new CParaTabs();
			oTabs.Add(oTab);

			if ((!isPreserveTabs && oPara.RemoveTabsForTOC()) || null !== numTabPos)
			{
				// В данной ситуации ворд делает следующим образом: он пробегает по параграфу и смотрит, если там есть
				// табы (в контенте, а не в свойствах), тогда он первый таб оставляет, а остальные заменяет на пробелы,
				// при этом в список табов добавляется новый левый таб без заполнителя, отступающий на 1,16 см от левого
				// поля параграфа, т.е. позиция таба зависит от стиля.

				var nFirstTabPos = 11.6;
				var sTOCStyleId  = this.LogicDocument.GetStyles().GetDefaultTOC(arrOutline[nIndex].Lvl);
				if (sTOCStyleId)
				{
					let paraPr = this.LogicDocument.GetStyles().Get_Pr(sTOCStyleId, styletype_Paragraph, null, null).ParaPr;
					if (null !== numTabPos)
						nFirstTabPos = Math.trunc((paraPr.Ind.Left + paraPr.Ind.FirstLine + numTabPos + 4.9) / 5 + 1) * 5;
					else
						nFirstTabPos = 1.6 + (paraPr.Ind.Left + paraPr.Ind.FirstLine);
				}
				
				oTabs.Add(new CParaTab(tab_Left, nFirstTabPos, Asc.c_oAscTabLeader.None));
			}

			oPara.Set_Tabs(oTabs);

			if (!(this.Instruction.IsSkipPageRefLvl(arrOutline[nIndex].Lvl)))
			{
				var oSeparatorRun = new ParaRun(oPara, false);
				if (!sSeparator || "" === sSeparator)
					oSeparatorRun.AddToContent(0, new AscWord.CRunTab());
				else
					oSeparatorRun.AddText(sSeparator.charAt(0));

				oContainer.Add_ToContent(nContainerPos, oSeparatorRun);

				var oPageRefRun = new ParaRun(oPara, false);

				oPageRefRun.AddToContent(-1, new ParaFieldChar(fldchartype_Begin, this.LogicDocument));
				oPageRefRun.AddInstrText("PAGEREF " + sBookmarkName + " \\h");
				oPageRefRun.AddToContent(-1, new ParaFieldChar(fldchartype_Separate, this.LogicDocument));
				oPageRefRun.AddText("" + this.LogicDocument.Get_SectionPageNumInfo2(oSrcParagraph.GetFirstNonEmptyPageAbsolute()).CurPage);
				oPageRefRun.AddToContent(-1, new ParaFieldChar(fldchartype_End, this.LogicDocument));
				oContainer.Add_ToContent(nContainerPos + 1, oPageRefRun);
			}

			oSelectedContent.Add(new AscCommonWord.CSelectedElement(oPara, true));
		}
	}
	else
	{
		var sReplacementText;
		if (bTOF)
		{
			sReplacementText = AscCommon.translateManager.getValue("No table of figures entries found.");
		}
		else
		{
			sReplacementText = AscCommon.translateManager.getValue("No table of contents entries found.");

			let oApi = this.LogicDocument.GetApi();
			oApi.sendEvent("asc_onError", c_oAscError.ID.ComplexFieldEmptyTOC, c_oAscError.Level.NoCritical);
		}

		oPara = new AscWord.Paragraph();
		oRun  = new ParaRun(oPara, false);
		oRun.SetBold(true);
		oRun.AddText(sReplacementText);
		oPara.AddToContent(0, oRun);
		oSelectedContent.Add(new AscCommonWord.CSelectedElement(oPara, true));
	}

	this.SelectFieldValue();
	this.LogicDocument.ConcatParagraphsOnRemove = true;
	this.LogicDocument.Remove(1, false, false, true);
	this.LogicDocument.ConcatParagraphsOnRemove = false;
	oRun = this.BeginChar.GetRun();

	let oParagraph = oRun.GetParagraph();
	let oAnchorPos = {
		Paragraph  : oParagraph,
		ContentPos : oParagraph.Get_ParaContentPos(false, false)
	};
	oParagraph.Check_NearestPos(oAnchorPos);
	oSelectedContent.Insert(oAnchorPos);
};
CComplexField.prototype.private_UpdatePAGEREF = function()
{
	var oBookmarksManager = this.LogicDocument.GetBookmarksManager();
	var oBookmark = oBookmarksManager.GetBookmarkByName(this.Instruction.GetBookmarkName());
	if(!oBookmark)
	{
		var sValue = AscCommon.translateManager.getValue("Error! Bookmark not defined.");
		this.private_InsertError(sValue);
		return;
	}
	this.LogicDocument.AddText(this.private_CalculatePAGEREF());
};
CComplexField.prototype.private_CalculatePAGEREF = function()
{
	var oBookmarksManager = this.LogicDocument.GetBookmarksManager();
	var oBookmark = oBookmarksManager.GetBookmarkByName(this.Instruction.GetBookmarkName());

	var sValue = AscCommon.translateManager.getValue("Error! Bookmark not defined.");
	if (oBookmark)
	{
		var oStartBookmark = oBookmark[0];
		var nBookmarkPage  = oStartBookmark.GetPage() + 1;
		if (this.Instruction.IsPositionRelative())
		{
			if (oStartBookmark.GetPage() === this.SeparateChar.GetPage())
			{
				var oBookmarkXY = oStartBookmark.GetXY();
				var oFieldXY    = this.SeparateChar.GetXY();

				if (Math.abs(oBookmarkXY.Y - oFieldXY.Y) < 0.001)
					sValue = oBookmarkXY.X < oFieldXY.X ? AscCommon.translateManager.getValue("above") : AscCommon.translateManager.getValue("below");
				else if (oBookmarkXY.Y < oFieldXY.Y)
					sValue = AscCommon.translateManager.getValue("above");
				else
					sValue = AscCommon.translateManager.getValue("below");
			}
			else
			{
				sValue = AscCommon.translateManager.getValue("on page ") + nBookmarkPage;
			}
		}
		else
		{
			sValue = (this.LogicDocument.Get_SectionPageNumInfo2(oStartBookmark.GetPage()).CurPage) + "";
		}
	}

	return sValue;
};
CComplexField.prototype.private_UpdateNUMPAGES = function()
{
	this.LogicDocument.AddText("" + this.private_CalculateNUMPAGES());
};
CComplexField.prototype.private_CalculateNUMPAGES = function()
{
	return this.LogicDocument.GetPagesCount();
};
CComplexField.prototype.private_UpdateTIME = function(ms)
{
	var sDate = this.private_CalculateTIME(ms);

	if (sDate)
		this.LogicDocument.AddText(sDate);
};
CComplexField.prototype.private_CalculateTIME = function(ms)
{
	var nLangId = 1033;
	var oSepChar = this.GetSeparateChar();
	if (oSepChar && oSepChar.GetRun())
	{
		var oCompiledTextPr = oSepChar.GetRun().Get_CompiledPr(false);
		nLangId = oCompiledTextPr.Lang.Val;
	}

	var sFormat = this.Instruction.GetFormat();
	var oFormat = AscCommon.oNumFormatCache.get(sFormat, AscCommon.NumFormatType.WordFieldDate);
	var sDate   = "";
	if (oFormat)
	{
		var oCultureInfo = AscCommon.g_aCultureInfos[nLangId];

		if (undefined !== ms)
		{
			var oDateTime = new Asc.cDate(ms);
			sDate         = oFormat.formatToWord(oDateTime.getExcelDate(true) + (oDateTime.getUTCHours() * 60 * 60 + oDateTime.getMinutes() * 60 + oDateTime.getSeconds()) / AscCommonExcel.c_sPerDay, 15, oCultureInfo);

		}
		else
		{
			let oDateTime = new Asc.cDate();
			sDate         = oFormat.formatToWord(oDateTime.getExcelDate(true) + (oDateTime.getHours() * 60 * 60 + oDateTime.getMinutes() * 60 + oDateTime.getSeconds()) / AscCommonExcel.c_sPerDay, 15, oCultureInfo);
		}
	}

	return sDate;
};
CComplexField.prototype.private_InsertError = function(sText)
{
	var oTextPr = new CTextPr();
	oTextPr.Set_FromObject({Bold: true});
	this.private_InsertMessage(sText, oTextPr);
};
CComplexField.prototype.private_GetREFPosValue = function()
{
	var oBookmarksManager = this.LogicDocument.GetBookmarksManager();
	var sBookmarkName = this.Instruction.GetBookmarkName();
	var oBookmark = oBookmarksManager.GetBookmarkByName(sBookmarkName);
	if(!oBookmark)
	{
		return "";
	}
	var oStartBookmark = oBookmark[0];
	var oSrcParagraph = oStartBookmark.Paragraph;
	var oRun       = this.BeginChar.GetRun();
	var oParagraph = oRun.GetParagraph();
	if(!oSrcParagraph || !oParagraph)
	{
		return "";
	}
	var oParent = oParagraph.GetParent();
	var oSrcParent = oSrcParagraph.GetParent();
	if(!oParent || !oSrcParent)
	{
		return "";
	}
	var oTopDoc = oParent.Is_TopDocument(true);
	if(oTopDoc !== oSrcParent.Is_TopDocument(true))
	{
		return "";
	}
	var sPosition = AscCommon.translateManager.getValue("above");
	oRun.Make_ThisElementCurrent(false);
	var aFieldPos = oTopDoc.GetContentPosition();
	this.LogicDocument.TurnOff_InterfaceEvents();
	oBookmarksManager.SelectBookmark(sBookmarkName);
	this.LogicDocument.TurnOn_InterfaceEvents(false);
	var aBookmarkPos = oTopDoc.GetContentPosition(true, false);
	var nIdx, nEnd, oBookmarkPos, oFieldPos;
	for(nIdx = 0, nEnd = Math.min(aFieldPos.length, aBookmarkPos.length); nIdx < nEnd; ++nIdx)
	{
		oBookmarkPos = aBookmarkPos[nIdx];
		oFieldPos = aFieldPos[nIdx];
		if(oBookmarkPos && oFieldPos
		&& oBookmarkPos.Position !== oFieldPos.Position)
		{
			if(oBookmarkPos.Position < oFieldPos.Position)
			{
				sPosition = AscCommon.translateManager.getValue("above");
			}
			else
			{
				sPosition = AscCommon.translateManager.getValue("below");
			}
			break;
		}
	}
	return sPosition;
};
CComplexField.prototype.private_UpdateREF = function()
{
	this.private_InsertContent(this.private_GetREFContent());

};
CComplexField.prototype.private_CalculateREF = function()
{
	let logicDocument = this.LogicDocument;
	
	let state = logicDocument ? logicDocument.SaveDocumentState() : null;
	
	let oSelectedContent = this.private_GetREFContent();
	oSelectedContent.GetText(null);
	
	if (state)
		logicDocument.LoadDocumentState(state);
	
	return oSelectedContent.GetText(null);
};
CComplexField.prototype.private_GetMessageContent = function(sMessage, oTextPr)
{
	var oSelectedContent = new AscCommonWord.CSelectedContent();
	var oPara = new AscWord.Paragraph();
	var oRun  = new ParaRun(oPara, false);
	
	if (this.Instruction && this.Instruction.isMergeFormat() && this.SeparateChar)
		oRun.ApplyPr(this.SeparateChar.GetRun().GetDirectTextPr());
	
	if (oTextPr)
		oRun.ApplyPr(oTextPr);

	oRun.AddText(sMessage);
	oPara.AddToContent(0, oRun);
	oSelectedContent.Add(new AscCommonWord.CSelectedElement(oPara, false));
	return oSelectedContent;
};
CComplexField.prototype.private_GetErrorContent = function(sMessage)
{
	var oTextPr = new CTextPr();
	oTextPr.Set_FromObject({Bold: true});
	return this.private_GetMessageContent(sMessage, oTextPr);
};
CComplexField.prototype.private_GetBookmarkContent = function(sBookmarkName)
{
	var oBookmarksManager = this.LogicDocument.GetBookmarksManager();
	this.LogicDocument.TurnOff_InterfaceEvents();
	oBookmarksManager.SelectBookmark(sBookmarkName);
	this.LogicDocument.TurnOn_InterfaceEvents(false);
	var oSelectedContent = this.LogicDocument.GetSelectedContent(false);
	var aElements = oSelectedContent.Elements;
	var oElement;
	
	let isMergeFormat = this.Instruction.isMergeFormat();
	let textPr        = this.GetFieldValueTextPr();
	let paraTextPr    = new AscWord.ParaTextPr(textPr);
	
	for (var nIndex = 0; nIndex < aElements.length; ++nIndex)
	{
		oElement = aElements[nIndex];
		oElement.Element = oElement.Element.Copy(null, null, {
			SkipPageBreak         : true,
			SkipColumnBreak       : true,
			SkipAnchors           : true,
			SkipFootnoteReference : true,
			SkipComplexFields     : true,
			SkipComments          : true,
			SkipBookmarks         : true,
			SkipFldSimple         : true
		});
		
		if (isMergeFormat)
		{
			oElement.Element.SetApplyToAll(true);
			oElement.Element.AddToParagraph(paraTextPr);
			oElement.Element.SetApplyToAll(false);
		}
	}
	return oSelectedContent;
};
CComplexField.prototype.private_GetREFContent = function()
{
	var sValue = AscCommon.translateManager.getValue("Error! Reference source not found.");
	if(!this.Instruction || this.Instruction.Type !== AscWord.fieldtype_REF)
	{
		return this.private_GetErrorContent(sValue);
	}
	var oBookmarksManager = this.LogicDocument.GetBookmarksManager();
	var sBookmarkName = this.Instruction.GetBookmarkName();
	var oBookmark = oBookmarksManager.GetBookmarkByName(sBookmarkName);
	if(!oBookmark)
	{
		return this.private_GetErrorContent(sValue);
	}
	var oStartBookmark = oBookmark[0];
	var oSrcParagraph = oStartBookmark.Paragraph;
	var oRun       = this.BeginChar.GetRun();
	var oParagraph = oRun.GetParagraph();
	if(!oSrcParagraph || !oParagraph)
	{
		return this.private_GetErrorContent(sValue);
	}
	var oParent = oParagraph.GetParent();
	var oSrcParent = oSrcParagraph.GetParent();
	if(!oParent || !oSrcParent)
	{
		return this.private_GetErrorContent(sValue);
	}
	var sPosition = "";
	if(this.Instruction.IsPosition())
	{
		sPosition = this.private_GetREFPosValue();
	}
	if(this.Instruction.HaveNumberFlag())
	{
		if(!oSrcParagraph.IsNumberedNumbering())
		{
			return this.private_GetMessageContent("0", null);
		}
		var oNumPr     = oSrcParagraph.GetNumPr();
		var oNumbering = this.LogicDocument.GetNumbering();
		var oNumInfo   = oSrcParagraph.GetParent().CalculateNumberingValues(oSrcParagraph, oNumPr);
		var nLvl, oParaNumInfo;
		if(this.Instruction.IsNumber())
		{
			var oParaNumPr = oParagraph.GetNumPr();
			if(oParaNumPr && oParaNumPr.NumId === oNumPr.NumId)
			{
				oParaNumInfo = oParagraph.GetParent().CalculateNumberingValues(oParagraph, oParaNumPr);
				for(nLvl = 0; nLvl <= oNumPr.Lvl && nLvl <= oParaNumPr.Lvl; ++nLvl)
				{
					if(oParaNumInfo[nLvl] !== oNumInfo[nLvl])
					{
						break;
					}
				}
				sValue = "";
				for( ;nLvl <= oNumPr.Lvl; ++nLvl)
				{
					sValue += oNumbering.GetText(oNumPr.NumId, nLvl, oNumInfo, nLvl === oNumPr.Lvl);
				}
			}
			else
			{
				sValue = oNumbering.GetText(oNumPr.NumId, oNumPr.Lvl, oNumInfo, true);
			}
		}
		else if(this.Instruction.IsNumberFullContext())
		{
			sValue = "";
			var sDelimiter = this.Instruction.GetDelimiter();
			for(nLvl = 0; nLvl <= oNumPr.Lvl; ++nLvl)
			{
				sValue += oNumbering.GetText(oNumPr.NumId, nLvl, oNumInfo, nLvl === oNumPr.Lvl);
				if(nLvl !== oNumPr.Lvl && typeof sDelimiter === "string" && sDelimiter.length > 0)
				{
					sValue += sDelimiter;
				}
			}
		}
		else if(this.Instruction.IsNumberNoContext())
		{
			sValue = oNumbering.GetText(oNumPr.NumId, oNumPr.Lvl, oNumInfo, true);
		}
		if(sPosition.length > 0)
		{
			sValue += " ";
			sValue += sPosition;
		}
		return this.private_GetMessageContent(sValue, null);
	}
	else if(this.Instruction.IsPosition() && sPosition.length > 0)
	{
		return this.private_GetMessageContent(sPosition, null);
	}
	else // bookmark content
	{
		return this.private_GetBookmarkContent(sBookmarkName);
	}
	//TODO: Apply formatting from general switches
};
CComplexField.prototype.private_GetNOTEREFContent = function()
{
	var sValue = AscCommon.translateManager.getValue("Error! Bookmark not defined.");
	if(!this.Instruction || this.Instruction.Type !== AscWord.fieldtype_NOTEREF)
	{
		return this.private_GetErrorContent(sValue);
	}
	var oBookmarksManager = this.LogicDocument.GetBookmarksManager();
	var sBookmarkName = this.Instruction.GetBookmarkName();
	var oBookmark = oBookmarksManager.GetBookmarkByName(sBookmarkName);
	if(!oBookmark)
	{
		return this.private_GetErrorContent(sValue);
	}
	//check notes in bookmarked content
	this.LogicDocument.TurnOff_InterfaceEvents();
	oBookmarksManager.SelectBookmark(sBookmarkName);
	this.LogicDocument.TurnOn_InterfaceEvents(false);
	var oSelectionInfo = this.LogicDocument.GetSelectedElementsInfo({CheckAllSelection : true});
	var aFootEndNotes = oSelectionInfo.GetFootEndNoteRefs();
	if(aFootEndNotes.length === 0)
	{
		return this.private_GetErrorContent(sValue);
	}
	var oFootEndNote = aFootEndNotes[0];
	var oStartBookmark = oBookmark[0];
	var oSrcParagraph = oStartBookmark.Paragraph;
	var oRun       = this.BeginChar.GetRun();
	var oParagraph = oRun.GetParagraph();
	if(!oSrcParagraph || !oParagraph)
	{
		return this.private_GetErrorContent(sValue);
	}
	var oSelectedContent = new AscCommonWord.CSelectedContent();
	var oTextPr;
	var oPara = new AscWord.Paragraph();
	var oParent = oParagraph.GetParent();
	var oSrcParent = oSrcParagraph.GetParent();
	if(!oParent || !oSrcParent)
	{
		return this.private_GetErrorContent(sValue);
	}
	if(this.Instruction.IsPosition() && oParent.IsHdrFtr() === oSrcParent.IsHdrFtr())
	{
		sValue = this.private_GetREFPosValue();
		if(typeof sValue === "string" && sValue.length > 0)
		{
			oRun  = new ParaRun(oPara, false);
			oRun.AddText(sValue);
			oPara.AddToContent(0, oRun);
		}
	}
	else
	{
		oRun  = new ParaRun(oPara, false);
		if(this.Instruction.IsFormatting())
		{
			oTextPr = new CTextPr();
			oTextPr.Set_FromObject({VertAlign: AscCommon.vertalign_SuperScript});
			oRun.Apply_Pr(oTextPr);
		}
		oRun.AddText(oFootEndNote.private_GetString());
		oPara.AddToContent(0, oRun);
	}
	oSelectedContent.Add(new AscCommonWord.CSelectedElement(oPara, false));
	return oSelectedContent;
};
CComplexField.prototype.private_UpdateNOTEREF = function()
{
	this.private_InsertContent(this.private_GetNOTEREFContent());
};
CComplexField.prototype.private_CalculateNOTEREF = function()
{
	var oSelectedContent = this.private_GetNOTEREFContent();
	return oSelectedContent.GetText(null);
};
CComplexField.prototype.SelectFieldValue = function()
{
	var oDocument = this.GetTopDocumentContent();
	if (!oDocument)
		return;
	
	if (!this.SeparateChar)
	{
		let endRun = this.EndChar.GetRun();
		endRun.Make_ThisElementCurrent(false);
		endRun.SetCursorPosition(endRun.GetElementPosition(this.EndChar));
		let endPos = oDocument.GetContentPosition(false);
		oDocument.SetSelectionByContentPositions(endPos, endPos);
	}
	else
	{
		oDocument.RemoveSelection();
		
		var oRun = this.SeparateChar.GetRun();
		oRun.Make_ThisElementCurrent(false);
		oRun.SetCursorPosition(oRun.GetElementPosition(this.SeparateChar) + 1);
		var oStartPos = oDocument.GetContentPosition(false);
		
		oRun = this.EndChar.GetRun();
		oRun.Make_ThisElementCurrent(false);
		oRun.SetCursorPosition(oRun.GetElementPosition(this.EndChar));
		var oEndPos = oDocument.GetContentPosition(false);
		
		oDocument.SetSelectionByContentPositions(oStartPos, oEndPos);
	}
};
CComplexField.prototype.SelectFieldCode = function()
{
	var oDocument = this.GetTopDocumentContent();
	if (!oDocument)
		return;

	oDocument.RemoveSelection();

	var oRun = this.BeginChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.BeginChar) + 1);
	var oStartPos = oDocument.GetContentPosition(false);
	
	let endChar = this.SeparateChar ? this.SeparateChar : this.EndChar;
	oRun = endChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(endChar));
	var oEndPos = oDocument.GetContentPosition(false);

	oDocument.SetSelectionByContentPositions(oStartPos, oEndPos);
};
CComplexField.prototype.SelectField = function()
{
	var oDocument = this.GetTopDocumentContent();
	if (!oDocument)
		return;

	oDocument.RemoveSelection();

	var oRun = this.BeginChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.BeginChar));
	var oStartPos = oDocument.GetContentPosition(false);

	oRun = this.EndChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.EndChar) + 1);
	var oEndPos = oDocument.GetContentPosition(false);

	oDocument.RemoveSelection();
	oDocument.SetSelectionByContentPositions(oStartPos, oEndPos);
};
CComplexField.prototype.GetFieldValueText = function()
{
	let logicDocument = this.LogicDocument;
	var oDocument = this.GetTopDocumentContent();
	if (!oDocument || !this.SeparateChar)
		return "";
	
	let state = logicDocument ? logicDocument.SaveDocumentState() : null;
	oDocument.RemoveSelection();

	var oRun = this.SeparateChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.SeparateChar) + 1);
	var oStartPos = oDocument.GetContentPosition(false);

	oRun = this.EndChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.EndChar));
	var oEndPos = oDocument.GetContentPosition(false);

	oDocument.SetSelectionByContentPositions(oStartPos, oEndPos);
	let result = oDocument.GetSelectedText();
	
	if (state)
		logicDocument.LoadDocumentState(state);
	
	return result;
};
CComplexField.prototype.GetFieldValueTextPr = function(isCompiled)
{
	let fieldChar = this.SeparateChar ? this.SeparateChar : this.EndChar;
	if (isCompiled)
	{
		let run       = fieldChar.GetRun();
		let runParent = run.GetParent();
		let runPos    = run.private_GetPosInParent(runParent);
		
		let inRunPos  = run.GetElementPosition(fieldChar);
		if (inRunPos >= run.GetElementsCount() - 1
			&& runParent
			&& runParent.GetElement
			&& runParent.GetElement(runPos + 1) instanceof AscWord.CRun)
		{
			return runParent.GetElement(runPos + 1).getCompiledPr();
		}
		
		return run.getCompiledPr();
	}
	else
	{
		// TODO: Temporary. We select the first visible element in InstrText area and return its direct TextPr
		let logicDocument = this.LogicDocument;
		if (!logicDocument)
			return new AscWord.CTextPr();
		
		let state = logicDocument.SaveDocumentState();
		
		let run = fieldChar.GetRun();
		run.Make_ThisElementCurrent(false);
		run.SetCursorPosition(run.GetElementPosition(fieldChar) + 1);
		
		logicDocument.MoveCursorRight(true, false);
		let textPr = logicDocument.GetDirectTextPr();
		logicDocument.LoadDocumentState(state);
		return textPr;
	}
};
CComplexField.prototype.GetTopDocumentContent = function()
{
	if (!this.BeginChar || !this.EndChar)
		return null;
	
	let topDocument = this.BeginChar.GetTopDocumentContent();
	if (topDocument !== this.EndChar.GetTopDocumentContent())
		return null;
	
	if (this.SeparateChar && topDocument !== this.SeparateChar.GetTopDocumentContent())
		return null;
	
	return topDocument;
};
CComplexField.prototype.IsUse = function()
{
	if (!this.BeginChar)
		return false;

	return this.BeginChar.IsUse();
};
CComplexField.prototype.GetStartDocumentPosition = function()
{
	if (!this.BeginChar)
		return null;

	var oDocument = this.LogicDocument;
	var oState    = oDocument.SaveDocumentState();

	var oRun = this.BeginChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.BeginChar));
	var oDocPos = oDocument.GetContentPosition(false);

	oDocument.LoadDocumentState(oState);

	return oDocPos;
};
CComplexField.prototype.GetEndDocumentPosition = function()
{
	if (!this.EndChar)
		return null;

	var oDocument = this.LogicDocument;
	var oState    = oDocument.SaveDocumentState();

	var oRun = this.EndChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.EndChar) + 1);
	var oDocPos = oDocument.GetContentPosition(false);

	oDocument.LoadDocumentState(oState);

	return oDocPos;
};
CComplexField.prototype.IsValid = function()
{
	return (this.IsUse()
		&& this.BeginChar && this.BeginChar.IsValid()
		&& this.EndChar && this.EndChar.IsValid());
};
CComplexField.prototype.HaveValuePart = function()
{
	return (this.SeparateChar && this.SeparateChar.IsValid());
};
CComplexField.prototype.GetInstruction = function()
{
	this.private_UpdateInstruction();
	return this.Instruction;
};
CComplexField.prototype.private_UpdateInstruction = function()
{
	if ((!this.Instruction || !this.Instruction.CheckInstructionLine(this.InstructionLine)))
	{
		var oParser = new CFieldInstructionParser();
		this.Instruction = oParser.GetInstructionClass(this.InstructionLine);
		this.Instruction.SetComplexField(this);
		this.Instruction.SetInstructionLine(this.InstructionLine);
	}
};
CComplexField.prototype.private_CheckNestedComplexFields = function()
{
	var nCount = this.InstructionCF.length;
	if (nCount > 0)
	{
		this.Instruction     = null;
		this.InstructionLine = this.InstructionLineSrc;

		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			var sValue = this.InstructionCF[nIndex].CalculateValue();
			this.InstructionLine = this.InstructionLine.replace("\\&", sValue);
		}
	}
};
CComplexField.prototype.IsHidden = function()
{
	var oInstruction = this.GetInstruction();
	if (!oInstruction)
		return false;
	
	if (this.SeparateChar && this.SeparateChar.IsHiddenValue())
		return true;
	
	if (!this.BeginChar || !this.SeparateChar)
		return false;
	
	return (AscWord.fieldtype_ASK === oInstruction.GetType());
};
CComplexField.prototype.RemoveFieldWrap = function()
{
	if (!this.IsValid())
		return;
	
	if (!this.HaveValuePart())
		return this.RemoveField();

	this.EndChar.RemoveThisFromDocument();

	var oDocument = this.GetTopDocumentContent();
	if (!oDocument)
		return;

	var oRun = this.BeginChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.BeginChar));
	var oStartPos = oDocument.GetContentPosition(false);

	oRun = this.SeparateChar.GetRun();
	oRun.Make_ThisElementCurrent(false);
	oRun.SetCursorPosition(oRun.GetElementPosition(this.SeparateChar) + 1);
	var oEndPos = oDocument.GetContentPosition(false);

	oDocument.SetSelectionByContentPositions(oStartPos, oEndPos);
	oDocument.Remove();
};
CComplexField.prototype.MoveCursorOutsideElement = function(isBefore)
{
	var oDocument = this.GetTopDocumentContent();
	if (!oDocument)
		return;

	oDocument.RemoveSelection();

	if (isBefore)
	{
		var oRun = this.BeginChar.GetRun();
		oRun.Make_ThisElementCurrent(false);
		oRun.SetCursorPosition(oRun.GetElementPosition(this.BeginChar));

		if (oRun.IsCursorAtBegin())
			oRun.MoveCursorOutsideElement(true);
	}
	else
	{
		var oRun = this.EndChar.GetRun();
		oRun.Make_ThisElementCurrent(false);
		oRun.SetCursorPosition(oRun.GetElementPosition(this.EndChar) + 1);

		if (oRun.IsCursorAtEnd())
			oRun.MoveCursorOutsideElement(false);
	}
};
CComplexField.prototype.RemoveField = function()
{
	if (!this.IsValid())
		return;

	var oDocument = this.GetTopDocumentContent();
	if (!oDocument)
		return;

	this.SelectField();
	oDocument.Remove();
};
CComplexField.prototype.RemoveFieldChars = function()
{
	if (this.BeginChar)
		this.BeginChar.RemoveThisFromDocument();

	if (this.EndChar)
		this.EndChar.RemoveThisFromDocument();

	if (this.SeparateChar)
		this.SeparateChar.RemoveThisFromDocument();
};
/**
 * Выставляем свойства для данного поля
 * @param oPr (зависит от типа данного поля)
 */
CComplexField.prototype.SetPr = function(oPr)
{
	if (!this.IsValid())
		return;

	var oInstruction = this.GetInstruction();
	if (!oInstruction)
		return;

	var oDocument = this.GetTopDocumentContent();
	if (!oDocument)
		return;

	oInstruction.SetPr(oPr);
	var sNewInstruction = oInstruction.ToString();

	this.SelectFieldCode();
	oDocument.Remove();

	var oRun      = this.BeginChar.GetRun();
	var nInRunPos = oRun.GetElementPosition(this.BeginChar) + 1;
	oRun.AddInstrText(sNewInstruction, nInRunPos);
};
/**
 * Изменяем строку инструкции у поля
 * @param {string} sNewInstruction
 */
CComplexField.prototype.ChangeInstruction = function(sNewInstruction)
{
	if (!this.IsValid())
		return;

	var oDocument = this.GetTopDocumentContent();
	if (!oDocument)
		return;

	this.SelectFieldCode();
	oDocument.Remove();

	var oRun      = this.BeginChar.GetRun();
	var nInRunPos = oRun.GetElementPosition(this.BeginChar) + 1;
	oRun.AddInstrText(sNewInstruction, nInRunPos);

	this.Instruction        = null;
	this.InstructionLine    = sNewInstruction;
	this.InstructionLineSrc = sNewInstruction;
	this.InstructionCF      = [];
	this.private_UpdateInstruction();
};
CComplexField.prototype.CheckType = function(type)
{
	if (!this.IsValid())
		return false;
	
	let instruction = this.GetInstruction();
	if (!instruction)
		return false;
	
	// TODO: По-хорошему надо сделать мап типов, которые могут идти без разделителя, и какие не могут
	return instruction.GetType() === type;
};
CComplexField.prototype.IsAddin = function()
{
	return this.CheckType(AscWord.fieldtype_ADDIN);
};
CComplexField.prototype.IsFormField = function()
{
	if (!this.IsValid())
		return false;
	
	let instruction = this.GetInstruction();
	if (!instruction)
		return false;
	
	let type = instruction.GetType();
	return (AscWord.fieldtype_FORMCHECKBOX === type
		|| AscWord.fieldtype_FORMTEXT === type
		|| AscWord.fieldtype_FORMDROPDOWN === type);
};
CComplexField.prototype.IsFormFieldEnabled = function()
{
	if (!this.IsFormField())
		return false;
	
	let ffData = this.BeginChar.GetFFData();
	return (!ffData || ffData.isEnabled());
};
CComplexField.prototype.IsFormCheckBox = function()
{
	return this.CheckType(AscWord.fieldtype_FORMCHECKBOX);
};
CComplexField.prototype.ToggleFormCheckBox = function()
{
	if (!this.IsFormCheckBox())
		return;
	
	let beginChar = this.GetBeginChar();
	let run     = beginChar.GetRun();
	if (!run)
		return;
	
	let inRunPos= run.GetElementPosition(beginChar);
	if (-1 === inRunPos)
		return;
	
	let newChar = beginChar.Copy();
	let ffData  = newChar.GetFFData();
	if (!ffData)
	{
		ffData = new AscWord.FFData();
		newChar.ffData = ffData;
	}
	
	if (!ffData.checkBox)
		ffData.initCheckBox();
	
	ffData.checkBox.checked = !ffData.isCheckBoxChecked();
	
	run.RemoveFromContent(inRunPos, 1);
	run.AddToContent(inRunPos, newChar);
	
	this.ReplaceChar(newChar);
};
/**
 * Получаем список связанных параграфов с данным полем (параграфы содержащие метки поля, это не обязательно будут
 * все параграфы между метками начала и конца)
 * @ returns {Array[AscWord.Paragraph]}
 */
CComplexField.prototype.GetRelatedParagraphs = function()
{
	if (!this.IsValid())
		return [];
	
	let begPara = this.BeginChar.GetParagraph();
	if (!begPara)
		return [];
	
	let result = [begPara];
	if (this.SeparateChar)
	{
		let sepPara = this.SeparateChar.GetParagraph();
		if (sepPara !== begPara)
			result.push(sepPara);
	}
	
	let endPara = this.EndChar.GetParagraph();
	if (-1 === result.indexOf(endPara))
		result.push(endPara);
	
	return result;
};
CComplexField.prototype.IsShowFieldCode = function()
{
	if (!this.IsValid())
		return false;
	
	return this.BeginChar.IsShowFieldCode();
};
CComplexField.prototype.ToggleFieldCodes = function()
{
	let isShowFieldCode = !this.BeginChar.IsShowFieldCode();
	
	this.BeginChar.SetShowFieldCode(isShowFieldCode);
	
	let logicDocument = this.LogicDocument;
	if (!logicDocument)
		return;
	
	let history    = logicDocument.GetHistory();
	let recalcData = history.getRecalcDataByElements([this.BeginChar.GetParagraph()]);
	logicDocument.RecalculateWithParams(recalcData);
	
	if (isShowFieldCode)
	{
		this.BeginChar.MoveCursorToChar(false);
	}
	else
	{
		if (this.SeparateChar)
			this.SeparateChar.MoveCursorToChar(false);
		else
			this.EndChar.MoveCursorToChar(true);
	}
};

function getRefInstruction(sBookmarkName, nType, bHyperlink, bAboveBelow, sSeparator)
{
	if (!(typeof sBookmarkName === "string" && sBookmarkName.length > 0))
	{
		return;
	}
	var sInstr  = "";
	var sSuffix = "";
	if (bHyperlink)
	{
		sSuffix += " \\h";
	}
	if (bAboveBelow && nType !== Asc.c_oAscDocumentRefenceToType.AboveBelow)
	{
		sSuffix += " \\p";
	}
	if (typeof sSeparator === "string" && sSeparator.length > 0)
	{
		sSuffix += " \\d " + sSeparator;
	}
	
	switch (nType)
	{
		case Asc.c_oAscDocumentRefenceToType.PageNum:
		{
			sInstr = " PAGEREF " + sBookmarkName;
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.Text:
		case Asc.c_oAscDocumentRefenceToType.OnlyCaptionText:
		case Asc.c_oAscDocumentRefenceToType.OnlyLabelAndNumber:
		{
			sInstr = " REF " + sBookmarkName + " ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.ParaNum:
		{
			sInstr = " REF " + sBookmarkName + " \\r ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.ParaNumNoContext:
		{
			sInstr = " REF " + sBookmarkName + " \\n ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.ParaNumFullContex:
		{
			sInstr = " REF " + sBookmarkName + " \\w ";
			sInstr += sSuffix;
			break;
		}
		case Asc.c_oAscDocumentRefenceToType.AboveBelow:
		{
			sInstr = " REF " + sBookmarkName + " \\p ";
			sInstr += sSuffix;
			break;
		}
	}
	
	return sInstr;
}

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CComplexField = CComplexField;

window['AscWord'] = window['AscWord'] || {};
window['AscWord'].getRefInstruction = getRefInstruction;
window['AscWord'].CComplexField = CComplexField;
window['AscWord'].ComplexField = CComplexField;
