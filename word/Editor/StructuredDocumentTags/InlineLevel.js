/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
 * User: Ilja.Kirillov
 * Date: 11.05.2017
 * Time: 16:48
 */

/**
 * @constructor
 * @extends {CParagraphContentWithParagraphLikeContent}
 */

var c_oAscSdtLockType = Asc.c_oAscSdtLockType;
function CInlineLevelSdt()
{
	this.Id = AscCommon.g_oIdCounter.Get_NewId();

	CParagraphContentWithParagraphLikeContent.call(this);

	this.Pr   = new CSdtPr();
	this.Type = para_InlineLevelSdt;

	this.BoundsPaths          = null;
	this.BoundsPathsStartPage = -1;

	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
	g_oTableId.Add(this, this.Id);

	this.SkipSpecialLock = false;
	this.Current         = false;
}

CInlineLevelSdt.prototype = Object.create(CParagraphContentWithParagraphLikeContent.prototype);
CInlineLevelSdt.prototype.constructor = CInlineLevelSdt;
AscCommon.ExtendPrototype(CInlineLevelSdt, CSdtBase);

CInlineLevelSdt.prototype.IsInlineLevel = function()
{
	return true;
};
CInlineLevelSdt.prototype.IsBlockLevel = function()
{
	return false;
};
CInlineLevelSdt.prototype.Get_Id = function()
{
	return this.Id;
};
CInlineLevelSdt.prototype.GetId = function()
{
	return this.Get_Id();
};
CInlineLevelSdt.prototype.Add = function(Item)
{
	if (this.IsPlaceHolder() && para_TextPr === Item.Type)
	{
		var oTempTextPr = this.Pr.TextPr.Copy();
		oTempTextPr.Merge(Item.Value);
		this.SetDefaultTextPr(oTempTextPr);
		return;
	}

	this.private_ReplacePlaceHolderWithContent();

	var oTextFormRun;
	if (this.IsTextForm())
	{
		if (para_Tab === Item.Type)
			return CParagraphContentWithParagraphLikeContent.prototype.Add.call(this, new ParaSpace());
		else if (Item.Type !== para_Text && Item.Type !== para_Space)
			return;

		oTextFormRun = this.MakeSingleRunElement(false);

		if (this.Pr.TextForm.MaxCharacters > 0 && oTextFormRun.GetElementsCount() >= this.Pr.TextForm.MaxCharacters)
			return;
	}

	CParagraphContentWithParagraphLikeContent.prototype.Add.apply(this, arguments);
};
CInlineLevelSdt.prototype.Copy = function(isUseSelection, oPr)
{
	var oContentControl = new CInlineLevelSdt();

	var nStartPos = 0;
	var nEndPos   = this.Content.length - 1;

	if (isUseSelection && this.State.Selection.Use)
	{
		nStartPos = this.State.Selection.StartPos;
		nEndPos   = this.State.Selection.EndPos;

		if (nStartPos > nEndPos)
		{
			nStartPos = this.State.Selection.EndPos;
			nEndPos   = this.State.Selection.StartPos;
		}
	}

	if (nStartPos <= nEndPos)
		oContentControl.ClearContent();

	for (var nCurPos = nStartPos; nCurPos <= nEndPos; ++nCurPos)
	{
		var oItem = this.Content[nCurPos];

		if (nStartPos === nEndPos || nEndPos === nCurPos)
			oContentControl.AddToContent(nCurPos - nStartPos, oItem.Copy(isUseSelection, oPr));
		else
			oContentControl.AddToContent(nCurPos - nStartPos, oItem.Copy(false, oPr));
	}

	// ВАЖНО: настройки копируем после копирования содержимого, потому что есть специальные случаи, когда
	//        содержимое дальше меняется в зависимости от настроек (например, для радио кнопок)
	this.private_CopyPrTo(oContentControl, oPr);

	if (oContentControl.IsEmpty())
		oContentControl.ReplaceContentWithPlaceHolder();

	return oContentControl;
};
CInlineLevelSdt.prototype.private_CopyPrTo = function(oContentControl, oPr)
{
	oContentControl.SetDefaultTextPr(this.GetDefaultTextPr());
	oContentControl.SetLabel(this.GetLabel());
	oContentControl.SetTag(this.GetTag());
	oContentControl.SetAlias(this.GetAlias());
	oContentControl.SetContentControlLock(this.GetContentControlLock());
	oContentControl.SetAppearance(this.GetAppearance());
	oContentControl.SetColor(this.GetColor());

	if (undefined !== this.Pr.DocPartObj)
		oContentControl.SetDocPartObj(this.Pr.DocPartObj.Category, this.Pr.DocPartObj.Gallery, this.Pr.DocPartObj.Unique);

	if (undefined !== this.Pr.CheckBox)
	{
		var oCheckBoxPr     = this.Pr.CheckBox;
		var isChangeChecked = false;

		if (this.Pr.CheckBox.GroupKey
			&& true === this.Pr.CheckBox.Checked
			&& this.GetLogicDocument()
			&& !this.GetLogicDocument().DragAndDropAction)
		{
			oCheckBoxPr = this.Pr.CheckBox.Copy();
			oCheckBoxPr.Checked = false;
			isChangeChecked = true;
		}

		oContentControl.SetCheckBoxPr(oCheckBoxPr);

		if (isChangeChecked)
			oContentControl.private_UpdateCheckBoxContent();
	}

	if (undefined !== this.Pr.Picture)
		oContentControl.SetPicturePr(this.Pr.Picture);

	if (undefined !== this.Pr.ComboBox)
		oContentControl.SetComboBoxPr(this.Pr.ComboBox);

	if (undefined !== this.Pr.DropDown)
		oContentControl.SetDropDownListPr(this.Pr.DropDown);

	if (undefined !== this.Pr.Date)
		oContentControl.SetDatePickerPr(this.Pr.Date);

	oContentControl.SetShowingPlcHdr(this.Pr.ShowingPlcHdr);
	oContentControl.SetPlaceholder(this.private_CopyPlaceholder(oPr));
	oContentControl.SetContentControlEquation(this.Pr.Equation);
	oContentControl.SetContentControlTemporary(this.Pr.Temporary);

	if (undefined !== this.Pr.FormPr)
		oContentControl.SetFormPr(this.Pr.FormPr);

	if (undefined !== this.Pr.TextForm)
		oContentControl.SetTextFormPr(this.Pr.TextForm);

	if (undefined !== this.Pr.PictureFormPr)
		oContentControl.SetPictureFormPr(this.Pr.PictureFormPr);
};
CInlineLevelSdt.prototype.GetSelectedContent = function(oSelectedContent)
{
	var oNewElement = new CInlineLevelSdt();
	this.private_CopyPrTo(oNewElement);

	if (this.IsPlaceHolder())
	{
		return oNewElement;
	}
	else
	{
		oNewElement.ReplacePlaceHolderWithContent();

		var nStartPos = this.State.Selection.StartPos;
		var nEndPos   = this.State.Selection.EndPos;

		if (nStartPos > nEndPos)
		{
			nStartPos = this.State.Selection.EndPos;
			nEndPos   = this.State.Selection.StartPos;
		}

		var nItemPos = 0;
		for (var nPos = nStartPos, nItemPos = 0; nPos <= nEndPos; ++nPos)
		{
			var oNewItem = this.Content[nPos].GetSelectedContent(oSelectedContent);
			if (oNewItem)
			{
				oNewElement.AddToContent(nItemPos, oNewItem);
				nItemPos++;
			}
		}

		if (0 === nItemPos)
			return null;

		return oNewElement;
	}
};
CInlineLevelSdt.prototype.GetSelectedText = function(bAll, bClearText, oPr)
{
	if (oPr && oPr.MathAdd && this.IsContentControlEquation() && this.IsPlaceHolder())
		return "";

	return CParagraphContentWithParagraphLikeContent.prototype.GetSelectedText.apply(this, arguments);
};
CInlineLevelSdt.prototype.GetSelectedElementsInfo = function(Info)
{
	Info.SetInlineLevelSdt(this);
	CParagraphContentWithParagraphLikeContent.prototype.GetSelectedElementsInfo.apply(this, arguments);
};
CInlineLevelSdt.prototype.Add_ToContent = function(Pos, Item, UpdatePosition)
{
	History.Add(new CChangesParaFieldAddItem(this, Pos, [Item]));
	CParagraphContentWithParagraphLikeContent.prototype.Add_ToContent.apply(this, arguments);

	var sKey = this.GetFormKey();
	if (sKey)
		this.GetLogicDocument().OnChangeForm(sKey, this);
};
CInlineLevelSdt.prototype.Remove_FromContent = function(Pos, Count, UpdatePosition)
{
	// Получим массив удаляемых элементов
	var DeletedItems = this.Content.slice(Pos, Pos + Count);
	History.Add(new CChangesParaFieldRemoveItem(this, Pos, DeletedItems));
	CParagraphContentWithParagraphLikeContent.prototype.Remove_FromContent.apply(this, arguments);

	var sKey = this.GetFormKey();
	if (sKey)
		this.GetLogicDocument().OnChangeForm(sKey, this);

};
CInlineLevelSdt.prototype.Split = function (ContentPos, Depth)
{
	// Не даем разделять
	return null;
};
CInlineLevelSdt.prototype.CanSplit = function()
{
	return false;
};
CInlineLevelSdt.prototype.Recalculate_Range_Spaces = function(PRSA, _CurLine, _CurRange, _CurPage)
{
	var isFastRangeRecalc = PRSA.IsFastRangeRecalc();

	var CurLine  = _CurLine - this.StartLine;
	var CurRange = (0 === CurLine ? _CurRange - this.StartRange : _CurRange);

	if (0 === CurLine && 0 === CurRange && !isFastRangeRecalc)
		this.Bounds = {};

	var oParagraph = PRSA.Paragraph;
	var Y0         = oParagraph.Lines[_CurLine].Top + oParagraph.Pages[_CurPage].Y;
	var Y1         = oParagraph.Lines[_CurLine].Y + oParagraph.Pages[_CurPage].Y + oParagraph.Lines[_CurLine].Metrics.Descent;
	var X0         = PRSA.X;

	if (0 === CurLine)
		Y0 = oParagraph.Lines[_CurLine].Y + oParagraph.Pages[_CurPage].Y - oParagraph.Lines[_CurLine].Metrics.Ascent;

	if (this.IsForm() && !this.IsPicture() && (this.Content[0] instanceof ParaRun))
	{
		var oRun    = this.Content[0];
		var oTextPr = oRun.Get_CompiledPr(false);

		g_oTextMeasurer.SetTextPr(oTextPr, oParagraph.GetTheme());
		g_oTextMeasurer.SetFontSlot(fontslot_ASCII);

		var nTextHeight  = g_oTextMeasurer.GetHeight();
		var nTextDescent = Math.abs(g_oTextMeasurer.GetDescender());
		var nTextAscent  = nTextHeight - nTextDescent;
		var nYOffset     = oTextPr.Position;

		if (0 === CurLine)
			Y0 = oParagraph.Lines[_CurLine].Y + oParagraph.Pages[_CurPage].Y - nTextAscent - nYOffset;

		Y1 = oParagraph.Lines[_CurLine].Y + oParagraph.Pages[_CurPage].Y + nTextDescent - nYOffset;
	}

	if (!isFastRangeRecalc)
	{
		for (var Key in this.Bounds)
		{
			var __CurLine = (Key | 0) >> 16;
			if (__CurLine === CurLine - 1 && this.Bounds[Key].PageInternal === _CurPage)
			{
				this.Bounds[Key].H = Y0 - this.Bounds[Key].Y;
			}
		}
	}

	CParagraphContentWithParagraphLikeContent.prototype.Recalculate_Range_Spaces.apply(this, arguments);

	var X1 = PRSA.X;

	if (this.IsForm() && this.IsPicture() && Math.abs(X1 - X0) > 0.001)
	{
		var arrDrawings = this.GetAllDrawingObjects();
		if (arrDrawings.length > 0 && arrDrawings[0].IsPicture())
		{
			Y0 = arrDrawings[0].GraphicObj.y;
			Y1 = arrDrawings[0].GraphicObj.y + arrDrawings[0].GraphicObj.extY;
			X0 = arrDrawings[0].GraphicObj.x;
			X1 = arrDrawings[0].GraphicObj.x + arrDrawings[0].GraphicObj.extX;
		}
	}

	if (isFastRangeRecalc && this.Bounds[((CurLine << 16) & 0xFFFF0000) | (CurRange & 0x0000FFFF)])
	{
		var oBounds = this.Bounds[((CurLine << 16) & 0xFFFF0000) | (CurRange & 0x0000FFFF)];
		oBounds.X = X0;
		oBounds.W = X1 - X0;
	}
	else
	{
		this.Bounds[((CurLine << 16) & 0xFFFF0000) | (CurRange & 0x0000FFFF)] = {
			X            : X0,
			W            : X1 - X0,
			Y            : Y0,
			H            : Y1 - Y0,
			TextLineH    : Y1 - Y0,
			Page         : PRSA.Paragraph.Get_AbsolutePage(_CurPage),
			PageInternal : _CurPage
		};
	}

	this.BoundsPaths = null;
};
CInlineLevelSdt.prototype.Draw_HighLights = function(PDSH)
{
	PDSH.AddInlineSdt(this);
	var oGraphics = PDSH.Graphics;

	if (this.IsSkipDraw(oGraphics))
		return this.SkipDraw(PDSH);

	// Для экспорта в PDF записываем поля. Поля, находящиеся в автофигурах, пока не пишем
	var oTransform = this.Get_ParentTextTransform();
	if (!oGraphics.isPrintMode && this.private_IsAddFormFieldToGraphics(oGraphics, oTransform))
	{
		this.SkipDraw(PDSH);

		// Ищем первый ненулевой промежуток, если он на данной странице, тогда сохраняем его в форму
		var oBounds = null;
		for (var Key in this.Bounds)
		{
			if (this.Bounds[Key].W > 0.001 && this.Bounds[Key].H > 0.001)
			{
				if (this.Bounds[Key].PageInternal === PDSH.Page)
					oBounds = this.Bounds[Key];

				var CurLine = PDSH.Line - this.StartLine;
				var CurRange = (0 === CurLine ? PDSH.Range - this.StartRange : PDSH.Range);

				if ((Key | 0) !== ((CurLine << 16) & 0xFFFF0000) | (CurRange & 0x0000FFFF))
					return;

				break;
			}
		}

		if (this.IsFixedForm())
		{
			var oShape       = this.GetParagraph().Parent.Is_DrawingShape(true);
			var oShapeBounds = oShape.getFormRelRect();

			if (oShapeBounds.Page === PDSH.Paragraph.GetAbsolutePage(PDSH.Page))
				oBounds = oShapeBounds;
		}

		var oRun = this.Content[0];
		if (oBounds && oRun)
		{
			var X = oBounds.X;
			var Y = oBounds.Y;

			if (oTransform)
			{
				X = oTransform.TransformPointX(oBounds.X, oBounds.Y);
				Y = oTransform.TransformPointY(oBounds.X, oBounds.Y);
			}

			var oTextPr = oRun.Get_CompiledPr(false);

			g_oTextMeasurer.SetTextPr(oTextPr, PDSH.Paragraph.GetTheme());
			g_oTextMeasurer.SetFontSlot(fontslot_ASCII);

			var nTextHeight  = g_oTextMeasurer.GetHeight();
			var nTextDescent = Math.abs(g_oTextMeasurer.GetDescender());
			var nTextAscent  = nTextHeight - nTextDescent;

			var oColor = oTextPr.GetSimpleTextColor(PDSH.Paragraph.GetTheme(), PDSH.Paragraph.GetColorMap());
			oGraphics.SetTextPr(oTextPr, PDSH.Paragraph.GetTheme());
			oGraphics.b_color1(oColor.r, oColor.g, oColor.b, this.IsPlaceHolder() ? 127 : 255);
			oGraphics.SetFontSlot(fontslot_ASCII); // Именно на этой функции записываются настройки шрифта в метафайл

			// TODO: Заглушка для AdobeReader
			//       Середина по вертикали у поля совпадает со средней точкой bbox, поэтому меняем сдвиги по вертикали
			//       с учетом этого момента

			var nW         = oBounds.W;
			var nH         = oBounds.H;
			var nBaseLine  = nTextAscent;

			if ((this.IsTextForm() || this.IsDropDownList() || this.IsComboBox())
				&& (!this.IsFixedForm() || !this.IsMultiLineForm()))
			{
				var oLimits = g_oTextMeasurer.GetLimitsY();

				var nMidPoint = ((nBaseLine - oLimits.min) + (nBaseLine - oLimits.max)) / 2;

				var nDiff = nH / 2 - nMidPoint;
				if (Math.abs(nDiff) > 0.001)
				{
					Y  -= nDiff;
					nBaseLine += nDiff;
				}
			}

			oGraphics.AddFormField(X, Y, nW, nH, nBaseLine, this);
		}
	}
	else
	{
		CParagraphContentWithParagraphLikeContent.prototype.Draw_HighLights.apply(this, arguments);
	}
};
CInlineLevelSdt.prototype.Draw_Elements = function(PDSE)
{
	if ((!PDSE.Graphics.isPrintMode && this.private_IsAddFormFieldToGraphics(PDSE.Graphics))
		|| (this.IsSkipDraw(PDSE.Graphics)))
		this.SkipDraw(PDSE);
	else
		CParagraphContentWithParagraphLikeContent.prototype.Draw_Elements.apply(this, arguments);
};
CInlineLevelSdt.prototype.Draw_Lines = function(PDSL)
{
	// Не отключаем отрисовку линий для PlaceHolder, т.к. рамка рисуется через данную функцию
	// отключение остальных отрисовок идет внутри ParaRun.Draw_Lines

	if ((!PDSL.Graphics.isPrintMode && this.private_IsAddFormFieldToGraphics(PDSL.Graphics)))
		this.SkipDraw(PDSL);
	else
		CParagraphContentWithParagraphLikeContent.prototype.Draw_Lines.apply(this, arguments);
};
CInlineLevelSdt.prototype.IsSkipDraw = function(oGraphics)
{
	if (!this.IsPlaceHolder() || !this.IsForm())
		return false;

	let oLogicDocument = this.GetLogicDocument();
	if (oLogicDocument)
	{
		if (true === oLogicDocument.ForceDrawPlaceHolders)
			return false;
		else if (false === oLogicDocument.ForceDrawPlaceHolders)
			return true;
	}

	return !!oGraphics.isPrintMode;
};
CInlineLevelSdt.prototype.GetRangeBounds = function(_CurLine, _CurRange)
{
	var CurLine  = _CurLine - this.StartLine;
	var CurRange = (0 === CurLine ? _CurRange - this.StartRange : _CurRange);

	return this.Bounds[((CurLine << 16) & 0xFFFF0000) | (CurRange & 0x0000FFFF)];
};
CInlineLevelSdt.prototype.private_IsAddFormFieldToGraphics = function(oGraphics, oTransform)
{
	var _oTransform = oTransform;
	if (undefined === oTransform)
		_oTransform = this.Get_ParentTextTransform();

	return (this.IsForm() && oGraphics && oGraphics.AddFormField && (!_oTransform || _oTransform.IsIdentity2()));
};
CInlineLevelSdt.prototype.Get_LeftPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
	if (false === UseContentPos && this.Content.length > 0)
	{
		// При переходе в новый контент встаем в его конец
		var CurPos = this.Content.length - 1;
		this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(CurPos, Depth);
		SearchPos.Found = true;
		return true;
	}

	if (this.IsForm() && this.IsPlaceHolder())
		return false;

	var bResult = CParagraphContentWithParagraphLikeContent.prototype.Get_LeftPos.call(this, SearchPos, ContentPos, Depth, UseContentPos);

	if (true !== bResult && this.Paragraph && this.Paragraph.LogicDocument && true === this.Paragraph.LogicDocument.IsFillingFormMode())
	{
		this.Get_StartPos(SearchPos.Pos, Depth);
		SearchPos.Found = true;
		return true;
	}

	return bResult;
};
CInlineLevelSdt.prototype.Get_RightPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
	if (false === UseContentPos && this.Content.length > 0)
	{
		// При переходе в новый контент встаем в его начало
		this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
		SearchPos.Pos.Update(0, Depth);
		SearchPos.Found = true;
		return true;
	}

	if (this.IsForm() && this.IsPlaceHolder())
		return false;

	var bResult = CParagraphContentWithParagraphLikeContent.prototype.Get_RightPos.call(this, SearchPos, ContentPos, Depth, UseContentPos, StepEnd);

	if (true !== bResult && this.Paragraph && this.Paragraph.LogicDocument && true === this.Paragraph.LogicDocument.IsFillingFormMode())
	{
		this.Get_EndPos(false, SearchPos.Pos, Depth);
		SearchPos.Found = true;
		return true;
	}

	return bResult;
};
CInlineLevelSdt.prototype.Remove = function(nDirection, bOnAddText)
{
	if (this.IsPlaceHolder())
	{
		if (!this.CanBeDeleted() && !bOnAddText)
			return true;

		if (bOnAddText || !this.Paragraph.LogicDocument.IsFillingFormMode())
			this.private_ReplacePlaceHolderWithContent();

		return true;
	}

	var bResult = CParagraphContentWithParagraphLikeContent.prototype.Remove.call(this, nDirection, bOnAddText);

	if (this.Is_Empty()
		&& this.Paragraph
		&& this.Paragraph.LogicDocument
		&& this.CanBeEdited()
		&& ((!bOnAddText
		&& true === this.Paragraph.LogicDocument.IsFillingFormMode())
		|| (this === this.Paragraph.LogicDocument.CheckInlineSdtOnDelete)))
	{
		this.private_ReplaceContentWithPlaceHolder();
		return true;
	}

	return bResult;
};
CInlineLevelSdt.prototype.Shift_Range = function(Dx, Dy, _CurLine, _CurRange, _CurPage)
{
	CParagraphContentWithParagraphLikeContent.prototype.Shift_Range.call(this, Dx, Dy, _CurLine, _CurRange, _CurPage);

	var CurLine = _CurLine - this.StartLine;
	var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

	var oRangeBounds = this.Bounds[((CurLine << 16) & 0xFFFF0000) | (CurRange & 0x0000FFFF)];
	if (oRangeBounds)
	{
		oRangeBounds.X += Dx;
		oRangeBounds.Y += Dy;
	}

	if (this.BoundsPaths)
		this.BoundsPaths = null;
};
CInlineLevelSdt.prototype.Get_WordStartPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
	CParagraphContentWithParagraphLikeContent.prototype.Get_WordStartPos.call(this, SearchPos, ContentPos, Depth, UseContentPos);

	if (true !== SearchPos.Found && this.Paragraph && this.Paragraph.LogicDocument && true === this.Paragraph.LogicDocument.IsFillingFormMode())
	{
		this.Get_StartPos(SearchPos.Pos, Depth);
		SearchPos.UpdatePos = true;
		SearchPos.Found     = true;

	}
};
CInlineLevelSdt.prototype.Get_WordEndPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
	CParagraphContentWithParagraphLikeContent.prototype.Get_WordEndPos.call(this, SearchPos, ContentPos, Depth, UseContentPos, StepEnd);

	if (true !== SearchPos.Found && this.Paragraph && this.Paragraph.LogicDocument && true === this.Paragraph.LogicDocument.IsFillingFormMode())
	{
		this.Get_EndPos(false, SearchPos.Pos, Depth);
		SearchPos.UpdatePos = true;
		SearchPos.Found     = true;

	}
};
CInlineLevelSdt.prototype.GetBoundingPolygon = function()
{
	var oHdrFtr     = this.Paragraph.Parent.IsHdrFtr(true);
	var nHdrFtrPage = oHdrFtr ? oHdrFtr.GetContent().GetAbsolutePage(0) : null;

	var StartPage = this.Paragraph.Get_StartPage_Absolute();
	if (null === this.BoundsPaths || StartPage !== this.BoundsPathsStartPage)
	{
		var arrBounds = [], arrRects = [], CurPage = -1;
		for (var Key in this.Bounds)
		{
			if (null !== nHdrFtrPage && nHdrFtrPage !== this.Paragraph.GetAbsolutePage(this.Bounds[Key].PageInternal))
				continue;

			if (CurPage !== this.Bounds[Key].PageInternal)
			{
				arrRects = [];
				arrBounds.push(arrRects);
				CurPage  = this.Bounds[Key].PageInternal;
			}
			this.Bounds[Key].Page = this.Paragraph.GetAbsolutePage(this.Bounds[Key].PageInternal);
			arrRects.push(this.Bounds[Key]);
		}

		this.BoundsPaths = [];
		for (var nIndex = 0, nCount = arrBounds.length; nIndex < nCount; ++nIndex)
		{
			var oPolygon = new AscCommon.CPolygon();
			oPolygon.fill([arrBounds[nIndex]]);
			this.BoundsPaths = this.BoundsPaths.concat(oPolygon.GetPaths(0));
		}

		this.BoundsPathsStartPage = StartPage;
	}

	return this.BoundsPaths;
};
CInlineLevelSdt.prototype.GetBoundingPolygonFirstLineH = function()
{
	for (var nCurLine = 0, nLinesCount = this.protected_GetLinesCount(); nCurLine < nLinesCount; ++nCurLine)
	{
		for (var nCurRange = 0, nRangesCount = this.protected_GetRangesCount(nCurLine); nCurRange < nRangesCount; ++nCurRange)
		{
			var oBounds = this.Bounds[((nCurLine << 16) & 0xFFFF0000) | (nCurRange & 0x0000FFFF)];
			if (!oBounds)
				break;
			else if (oBounds.W > 0)
				return oBounds.TextLineH;
		}
	}

	return 0;
};
/**
 * Функция возвращает рект, содержащий все отрезки данного контрола
 * @return {{X:number, Y:number, W:number, H:number, Page:number, Transform:object}}
 */
CInlineLevelSdt.prototype.GetBoundingRect = function()
{
	var nR = null;
	var nT = null;
	var nB = null;
	var nL = null;

	var nCurPage = -1;

	for (var Key in this.Bounds)
	{
		if (this.Bounds[Key].H > 0.001 && this.Bounds[Key].W > 0.001 && (-1 === nCurPage || nCurPage > this.Bounds[Key].PageInternal))
			nCurPage = this.Bounds[Key].PageInternal;
	}

	if (-1 === nCurPage)
		return {X: 0, Y : 0, W : 0, H : 0, Page : 0, Transform : null};

	var nPageAbs = this.Paragraph.GetAbsolutePage(nCurPage);
	for (var Key in this.Bounds)
	{
		if (nCurPage !== this.Bounds[Key].PageInternal)
			continue;

		if (null === nL || nL > this.Bounds[Key].X)
			nL = this.Bounds[Key].X;

		if (null === nR || nR < this.Bounds[Key].X + this.Bounds[Key].W)
			nR = this.Bounds[Key].X + this.Bounds[Key].W;

		if (null === nT || nT > this.Bounds[Key].Y)
			nT = this.Bounds[Key].Y;

		if (null === nB || nB < this.Bounds[Key].Y + this.Bounds[Key].H)
			nB = this.Bounds[Key].Y + this.Bounds[Key].H;
	}

	if (null === nL || null === nT || null === nR || null === nB)
		return {X: 0, Y : 0, W : 0, H : 0, Page : 0, Transform : null};

	return {
		X : nL,
		Y : nT,
		W : nR - nL,
		H : nB - nT,
		Page : nPageAbs,
		Transform : this.Get_ParentTextTransform()
	};
};
CInlineLevelSdt.prototype.IsFixedForm = function()
{
	if (!this.IsForm() || !this.Paragraph)
		return false;

	var oShape = this.Paragraph.Parent ? this.Paragraph.Parent.Is_DrawingShape(true) : null;
	return !!(oShape && oShape.isForm());
};
CInlineLevelSdt.prototype.GetFixedFormBounds = function(isUsePaddings)
{
	if (!this.Paragraph)
		return {X : 0, Y : 0, W : 0, H : 0, Page : 0};

	var oShape = this.Paragraph.Parent ? this.Paragraph.Parent.Is_DrawingShape(true) : null;
	if (oShape && oShape.isForm())
	{
		return oShape.getFormRelRect(isUsePaddings);
	}

	return {X : 0, Y : 0, W : 0, H : 0, Page : 0};
};
CInlineLevelSdt.prototype.DrawContentControlsTrack = function(isHover, X, Y, nCurPage, isCheckHit)
{
	if (!this.Paragraph && this.Paragraph.LogicDocument)
		return;

	var oLogicDocument = this.Paragraph.LogicDocument;

	if (this.IsContentControlEquation())
		return;

	var oDrawingDocument = oLogicDocument.GetDrawingDocument();
	
	if (undefined !== X && undefined !== Y && undefined !== nCurPage)
	{
		var isHit = false;

		for (var sKey in this.Bounds)
		{
			var oBound = this.Bounds[sKey];
			if (oBound.PageInternal === nCurPage && oBound.X <= X && X <= oBound.X + oBound.W && oBound.Y <= Y && Y <= oBound.Y + oBound.H)
			{
				isHit = true;
				break;
			}
		}

		if (false !== isCheckHit && !isHit)
			return;

		var sHelpText = "";
		if (isHover && this.IsForm() && (sHelpText = this.GetFormPr().HelpText))
		{
			var oMMData   = new AscCommon.CMouseMoveData();
			var oCoords   = oDrawingDocument.ConvertCoordsToCursorWR(X, Y, this.Paragraph.GetAbsolutePage(nCurPage), this.Paragraph.Get_ParentTextTransform());
			oMMData.X_abs = oCoords.X - 5;
			oMMData.Y_abs = oCoords.Y;
			oMMData.Type  = Asc.c_oAscMouseMoveDataTypes.Form;
			oMMData.Text  = sHelpText;
			oLogicDocument.GetApi().sync_MouseMoveCallback(oMMData);
		}
	}

	var oShape = this.Paragraph.Parent ? this.Paragraph.Parent.Is_DrawingShape(true) : null;
	if (this.IsForm() && oShape && oShape.isForm())
	{
		var oPolygon = new AscCommon.CPolygon();
		oPolygon.fill([[oShape.getFormRelRect()]]);
		oDrawingDocument.OnDrawContentControl(this, isHover ? AscCommon.ContentControlTrack.Hover : AscCommon.ContentControlTrack.In, oPolygon.GetPaths(0));
		return;
	}

	if (Asc.c_oAscSdtAppearance.Hidden === this.GetAppearance() || this.Paragraph.LogicDocument.IsForceHideContentControlTrack())
	{
		oDrawingDocument.OnDrawContentControl(null, isHover ? AscCommon.ContentControlTrack.Hover : AscCommon.ContentControlTrack.In);
		return;
	}

	oDrawingDocument.OnDrawContentControl(this, isHover ? AscCommon.ContentControlTrack.Hover : AscCommon.ContentControlTrack.In, this.GetBoundingPolygon());
};
CInlineLevelSdt.prototype.IsDrawContentControlsTrackBounds = function()
{
	var oShd;
	return (!this.IsForm()
		|| !this.IsCurrent()
		|| !(oShd = this.GetFormPr().GetShd())
		|| oShd.IsNil());
};
CInlineLevelSdt.prototype.SelectContentControl = function()
{
	var arrDrawings, oLogicDocument;
	if (this.IsPicture()
		&& (arrDrawings = this.GetAllDrawingObjects())
		&& arrDrawings.length > 0
		&& (oLogicDocument = this.GetLogicDocument()))
	{
		oLogicDocument.Select_DrawingObject(arrDrawings[0].GetId());
		return;
	}

	if (this.IsForm() && this.IsPlaceHolder() && this.GetLogicDocument() && this.GetLogicDocument().CheckFormPlaceHolder)
	{
		this.RemoveSelection();
		this.MoveCursorToStartPos();
		this.SetThisElementCurrent();
	}
	else
	{
		this.SelectThisElement(1);
	}
};
CInlineLevelSdt.prototype.MoveCursorToContentControl = function(isBegin)
{
	this.RemoveSelection();
	this.SetThisElementCurrent();

	if (isBegin)
		this.MoveCursorToStartPos();
	else
		this.MoveCursorToEndPos();
};
CInlineLevelSdt.prototype.MoveCursorToStartPos = function()
{
	if (this.IsPlaceHolder() && !this.IsForm())
	{
		this.SelectContentControl();
	}
	else
	{
		CParagraphContentWithParagraphLikeContent.prototype.MoveCursorToStartPos.apply(this);
	}
};
CInlineLevelSdt.prototype.MoveCursorToEndPos = function(isSelectFromEnd)
{
	if (this.IsPlaceHolder())
	{
		if (this.IsForm())
			CParagraphContentWithParagraphLikeContent.prototype.MoveCursorToStartPos.apply(this, arguments);
		else
			this.SelectContentControl();
	}
	else
	{
		CParagraphContentWithParagraphLikeContent.prototype.MoveCursorToEndPos.apply(this, arguments);
	}
};
CInlineLevelSdt.prototype.RemoveContentControlWrapper = function()
{
	var oParent = this.Get_Parent();
	if (!oParent)
		return {Parent : null, Pos : -1, Count : 0};

	var nElementPos = this.Get_PosInParent(oParent);
	if (-1 === nElementPos)
		return {Parent : null, Pos : -1, Count : 0};

	var nParentCurPos            = oParent instanceof Paragraph ? oParent.CurPos.ContentPos : oParent.State.ContentPos;
	var nParentSelectionStartPos = oParent.Selection.StartPos;
	var nParentSelectionEndPos   = oParent.Selection.EndPos;

	var nCount = this.Content.length;
	oParent.Remove_FromContent(nElementPos, 1);
	for (var nIndex = 0; nIndex < nCount; ++nIndex)
	{
		oParent.Add_ToContent(nElementPos + nIndex, this.Content[nIndex]);
	}

	if (nParentCurPos === nElementPos)
	{
		if (oParent instanceof Paragraph)
			oParent.CurPos.ContentPos = nParentCurPos + this.State.ContentPos;
		else
			oParent.State.ContentPos = nParentCurPos + this.State.ContentPos;

	}
	else if (nParentCurPos > nElementPos)
	{
		if (oParent instanceof Paragraph)
			oParent.CurPos.ContentPos = nParentCurPos + nCount - 1;
		else
			oParent.State.ContentPos = nParentCurPos + nCount - 1;
	}

	if (nParentSelectionStartPos === nElementPos)
		oParent.Selection.StartPos = nParentSelectionStartPos + this.Selection.StartPos;
	else if (nParentSelectionStartPos > nElementPos)
		oParent.Selection.StartPos = nParentSelectionStartPos + nCount - 1;

	if (nParentSelectionEndPos === nElementPos)
		oParent.Selection.EndPos = nParentSelectionEndPos + this.Selection.EndPos;
	else if (nParentSelectionEndPos > nElementPos)
		oParent.Selection.EndPos = nParentSelectionEndPos + nCount - 1;

	this.Remove_FromContent(0, this.Content.length);

	return {Parent : oParent, Pos : nElementPos, Count : nCount};
};
CInlineLevelSdt.prototype.FindNextFillingForm = function(isNext, isCurrent, isStart)
{
	if (isCurrent && (true === this.IsSelectedAll() || this.IsPlaceHolder()))
	{
		if (isNext)
			return CParagraphContentWithParagraphLikeContent.prototype.FindNextFillingForm.apply(this, arguments);

		return null;
	}

	if (!isCurrent && isNext)
		return this;

	var oRes = CParagraphContentWithParagraphLikeContent.prototype.FindNextFillingForm.apply(this, arguments);
	if (!oRes && !isNext)
		return this;

	return null;
};
CInlineLevelSdt.prototype.GetAllContentControls = function(arrContentControls)
{
	arrContentControls.push(this);
	CParagraphContentWithParagraphLikeContent.prototype.GetAllContentControls.apply(this, arguments);
};
CInlineLevelSdt.prototype.Document_UpdateInterfaceState = function()
{
	if (this.Paragraph && this.Paragraph.LogicDocument)
		this.Paragraph.LogicDocument.Api.sync_ContentControlCallback(this.GetContentControlPr());

	CParagraphContentWithParagraphLikeContent.prototype.Document_UpdateInterfaceState.apply(this, arguments);
};
CInlineLevelSdt.prototype.SetParagraph = function(oParagraph)
{
	if (this.GetTextFormPr() && this.GetLogicDocument())
		this.GetLogicDocument().RegisterForm(this);

	CParagraphContentWithParagraphLikeContent.prototype.SetParagraph.apply(this, arguments);
};
CInlineLevelSdt.prototype.Apply_TextPr = function(TextPr, IncFontSize, ApplyToAll)
{
	if (this.IsPlaceHolder() || ApplyToAll || this.IsSelectedAll() || this.IsDropDownList() || this.IsComboBox() || this.IsCheckBox() || this.IsDatePicker() || this.IsTextForm())
	{
		if (undefined !== IncFontSize)
		{
			var oCompiledTextPr = this.Get_CompiledTextPr(false);
			if (oCompiledTextPr)
			{
				var oNewTextPr = new CTextPr();
				oNewTextPr.FontSize   = oCompiledTextPr.GetIncDecFontSize(IncFontSize);
				oNewTextPr.FontSizeCS = oCompiledTextPr.GetIncDecFontSizeCS(IncFontSize);

				var oTempTextPr = this.Pr.TextPr.Copy();
				oTempTextPr.Merge(oNewTextPr);
				this.SetDefaultTextPr(oTempTextPr);
			}
		}
		else
		{
			var oTempTextPr = this.Pr.TextPr.Copy();
			oTempTextPr.Merge(TextPr);
			this.SetDefaultTextPr(oTempTextPr);
		}
	}

	if (this.IsDropDownList() || this.IsComboBox() || this.IsCheckBox() || this.IsDatePicker() || this.IsTextForm())
		CParagraphContentWithParagraphLikeContent.prototype.Apply_TextPr.call(this, TextPr, IncFontSize, true);
	else
		CParagraphContentWithParagraphLikeContent.prototype.Apply_TextPr.call(this, TextPr, IncFontSize, ApplyToAll);
};
CInlineLevelSdt.prototype.CanAddDropCap = function()
{
	if (!this.CanBeEdited() || this.IsPlaceHolder())
		return false;

	return CParagraphContentWithParagraphLikeContent.prototype.CanAddDropCap.apply(this, arguments);
};
CInlineLevelSdt.prototype.CheckSelectionForDropCap = function(isUsePos, oEndPos, nDepth)
{
	return false;
};
/**
 * Активен PlaceHolder сейчас или нет
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsPlaceHolder = function()
{
	return this.Pr.ShowingPlcHdr;
};
CInlineLevelSdt.prototype.private_ReplacePlaceHolderWithContent = function(bMathRun)
{
	if (!this.IsPlaceHolder())
		return;

	this.SetShowingPlcHdr(false);

	var isUseSelection = this.IsSelectionUse();

	this.RemoveSelection();
	this.MoveCursorToStartPos();

	this.RemoveFromContent(0, this.GetElementsCount());

	if (this.IsContentControlEquation())
	{
		var oParaMath = new ParaMath();
		oParaMath.Root.Load_FromMenu(c_oAscMathType.Default_Text, this.GetParagraph());
		oParaMath.Root.Correct_Content(true);
		this.AddToContent(0, oParaMath);
	}
	else
	{
		var oRun = new ParaRun(undefined, bMathRun);
		oRun.SetPr(this.Pr.TextPr.Copy());
		this.AddToContent(0, oRun);
	}

	this.RemoveSelection();
	this.MoveCursorToStartPos();

	if (isUseSelection)
		this.SelectAll();

	if (this.IsContentControlTemporary())
		this.RemoveContentControlWrapper();
};
CInlineLevelSdt.prototype.private_ReplaceContentWithPlaceHolder = function(isSelect)
{
	if (this.IsPlaceHolder())
		return;

	this.SetShowingPlcHdr(true);

	var isUseSelection = this.IsSelectionUse();

	this.private_FillPlaceholderContent();
	this.TrimCombForm();

	if (false !== isSelect)
		this.SelectContentControl();

	if (isUseSelection)
		this.SelectAll();
};
CInlineLevelSdt.prototype.private_FillPlaceholderContent = function()
{
	var isSelection = this.IsSelectionUse();

	this.RemoveFromContent(0, this.GetElementsCount());

	var oParagraph     = this.GetParagraph();
	var oLogicDocument = oParagraph ? oParagraph.GetLogicDocument() : editor.WordControl.m_oLogicDocument;
	var oDocPart       = oLogicDocument.GetGlossaryDocument().GetDocPartByName(this.GetPlaceholder());
	if (oDocPart)
	{
		var oFirstParagraph = oDocPart.GetFirstParagraph();

		if (this.IsContentControlEquation())
		{
			var oParaMath = new ParaMath();
			oParaMath.Root.Load_FromMenu(c_oAscMathType.Default_Text, this.GetParagraph(), null, oFirstParagraph.GetText());
			oParaMath.Root.Correct_Content(true);
			this.AddToContent(0, oParaMath);
		}
		else
		{
			// TODO: Последний Run с ParaEnd не добавляем
			for (var nPos = 0, nCount = oFirstParagraph.Content.length; nPos < nCount - 1; ++nPos)
			{
				this.AddToContent(0, oFirstParagraph.Content[nPos].Copy());
			}

			if (this.IsTextForm() && this.GetTextFormPr().MaxCharacters > 0 && this.Content[0] instanceof ParaRun)
			{
				// В такой ситуации у нас должен быть всего 1 ран в параграфе
				var oRun = this.Content[0];
				if (oRun.Content.length > this.GetTextFormPr().MaxCharacters)
				{
					oRun.RemoveFromContent(this.GetTextFormPr().MaxCharacters, oRun.Content.length - this.GetTextFormPr().MaxCharacters, true);
				}
			}
		}
	}
	else
	{
		if (this.IsContentControlEquation())
		{
			var oParaMath = new ParaMath();
			oParaMath.Root.Load_FromMenu(c_oAscMathType.Default_Text, this.GetParagraph(), null, String.fromCharCode(nbsp_charcode, nbsp_charcode, nbsp_charcode, nbsp_charcode));
			oParaMath.Root.Correct_Content(true);
			this.AddToContent(0, oParaMath);
		}
		else
		{
			var oRun = new ParaRun(oParagraph, false);
			oRun.AddText(String.fromCharCode(nbsp_charcode, nbsp_charcode, nbsp_charcode, nbsp_charcode));
			this.AddToContent(0, oRun);
		}
	}

	// В формах для плейсхолдера мы копируем настройки текста в сам плейсхолдер тоже
	if (this.IsForm())
	{
		for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			if (this.Content[nIndex] instanceof ParaRun)
				this.Content[nIndex].SetPr(this.Pr.TextPr.Copy());
		}
	}

	if (isSelection)
		this.SelectAll(1);
};
CInlineLevelSdt.prototype.Set_SelectionContentPos = function(StartContentPos, EndContentPos, Depth, StartFlag, EndFlag)
{
	if (this.IsPlaceHolder())
	{
		if (this.IsForm() && StartContentPos && EndContentPos && this.GetLogicDocument() && this.GetLogicDocument().CheckFormPlaceHolder)
		{
			this.Get_StartPos(StartContentPos, Depth);
			this.Get_StartPos(EndContentPos, Depth);
			CParagraphContentWithParagraphLikeContent.prototype.Set_SelectionContentPos.call(this, StartContentPos, EndContentPos, Depth, StartFlag, EndFlag);
		}
		else
		{

			if (this.Paragraph && this.Paragraph.GetSelectDirection() > 0)
				this.SelectAll(1);
			else
				this.SelectAll(-1);
		}
	}
	else
	{
		CParagraphContentWithParagraphLikeContent.prototype.Set_SelectionContentPos.apply(this, arguments);
	}
};
CInlineLevelSdt.prototype.Set_ParaContentPos = function(ContentPos, Depth)
{
	if (this.IsPlaceHolder() && this.IsForm() && ContentPos && this.GetLogicDocument() && this.GetLogicDocument().CheckFormPlaceHolder)
	{
		this.Get_StartPos(ContentPos, Depth);
		CParagraphContentWithParagraphLikeContent.prototype.Set_ParaContentPos.call(this, ContentPos, Depth);
	}
	else
	{
		CParagraphContentWithParagraphLikeContent.prototype.Set_ParaContentPos.apply(this, arguments);
	}
};
CInlineLevelSdt.prototype.ReplacePlaceHolderWithContent = function(bMathRun)
{
	this.private_ReplacePlaceHolderWithContent(bMathRun);
};
CInlineLevelSdt.prototype.ReplaceContentWithPlaceHolder = function(isSelect)
{
	this.private_ReplaceContentWithPlaceHolder(isSelect);
};
CInlineLevelSdt.prototype.CorrectContent = function()
{
	if (this.IsForm())
	{
		this.MakeSingleRunElement(false);
	}
	else
	{
		CParagraphContentWithParagraphLikeContent.prototype.CorrectContent.apply(this, arguments);
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Выставление настроек
//----------------------------------------------------------------------------------------------------------------------
CInlineLevelSdt.prototype.GetContentControlType = function()
{
	return c_oAscSdtLevelType.Inline;
};
CInlineLevelSdt.prototype.SetPr = function(oPr)
{
	this.SetAlias(oPr.Alias);
	this.SetTag(oPr.Tag);
	this.SetLabel(oPr.Label);
	this.SetContentControlLock(oPr.Lock);
	this.SetContentControlId(oPr.Id);

	if (undefined !== oPr.DocPartObj)
		this.SetDocPartObj(oPr.DocPartObj.Category, oPr.DocPartObj.Gallery, oPr.DocPartObj.Unique);

	if (undefined !== oPr.Appearance)
		this.SetAppearance(oPr.Appearance);

	if (undefined !== oPr.Color)
		this.SetColor(oPr.Color);
};
/**
 * Выставляем настройки текста по умолчанию для данного контрола
 * @param {CTextPr} oTextPr
 */
CInlineLevelSdt.prototype.SetDefaultTextPr = function(oTextPr)
{
	if (oTextPr && !this.Pr.TextPr.IsEqual(oTextPr))
	{
		History.Add(new CChangesSdtPrTextPr(this, this.Pr.TextPr, oTextPr));
		this.Pr.TextPr = oTextPr;
	}
};
/**
 * Получаем настройки для текста по умолчанию
 * @returns {CTextPr}
 */
CInlineLevelSdt.prototype.GetDefaultTextPr = function()
{
	return this.Pr.TextPr;
};
CInlineLevelSdt.prototype.SetAlias = function(sAlias)
{
	if (sAlias !== this.Pr.Alias)
	{
		History.Add(new CChangesSdtPrAlias(this, this.Pr.Alias, sAlias));
		this.Pr.Alias = sAlias;
	}
};
CInlineLevelSdt.prototype.GetAlias = function()
{
	return (undefined !== this.Pr.Alias ? this.Pr.Alias : "");
};
CInlineLevelSdt.prototype.SetAppearance = function(nType)
{
	if (this.Pr.Appearance !== nType)
	{
		History.Add(new CChangesSdtPrAppearance(this, this.Pr.Appearance, nType));
		this.Pr.Appearance = nType;
	}
};
CInlineLevelSdt.prototype.GetAppearance = function()
{
	return this.Pr.Appearance;
};
CInlineLevelSdt.prototype.SetColor = function(oColor)
{
	if (null === oColor || undefined === oColor)
	{
		if (undefined !== this.Pr.Color)
		{
			History.Add(new CChangesSdtPrColor(this, this.Pr.Color, undefined));
			this.Pr.Color = undefined;
		}
	}
	else
	{
		History.Add(new CChangesSdtPrColor(this, this.Pr.Color, oColor));
		this.Pr.Color = oColor;
	}
};
CInlineLevelSdt.prototype.GetColor = function()
{
	return this.Pr.Color;
};
CInlineLevelSdt.prototype.SetContentControlId = function(Id)
{
	if (this.Pr.Id !== Id)
	{
		History.Add(new CChangesSdtPrId(this, this.Pr.Id, Id));
		this.Pr.Id = Id;
	}
};
CInlineLevelSdt.prototype.GetContentControlId = function()
{
	return this.Pr.Id;
};
CInlineLevelSdt.prototype.SetTag = function(sTag)
{
	if (this.Pr.Tag !== sTag)
	{
		History.Add(new CChangesSdtPrTag(this, this.Pr.Tag, sTag));
		this.Pr.Tag = sTag;
	}
};
CInlineLevelSdt.prototype.GetTag = function()
{
	return (undefined !== this.Pr.Tag ? this.Pr.Tag : "");
};
CInlineLevelSdt.prototype.SetLabel = function(sLabel)
{
	if (this.Pr.Label !== sLabel)
	{
		History.Add(new CChangesSdtPrLabel(this, this.Pr.Label, sLabel));
		this.Pr.Label = sLabel;
	}
};
CInlineLevelSdt.prototype.GetLabel = function()
{
	return (undefined !== this.Pr.Label ? this.Pr.Label : "");
};
CInlineLevelSdt.prototype.SetDocPartObj = function(sCategory, sGallery, isUnique)
{
	History.Add(new CChangesSdtPrDocPartObj(this, this.Pr.DocPartObj, {Category : sCategory, Gallery : sGallery, Unique : isUnique}));
	this.Pr.DocPartObj.Category = sCategory;
	this.Pr.DocPartObj.Gallery  = sGallery;
	this.Pr.DocPartObj.Unique   = isUnique;
};
CInlineLevelSdt.prototype.SetContentControlLock = function(nLockType)
{
	if (this.Pr.Lock !== nLockType)
	{
		History.Add(new CChangesSdtPrLock(this, this.Pr.Lock, nLockType));
		this.Pr.Lock = nLockType;
	}
};
CInlineLevelSdt.prototype.GetContentControlLock = function()
{
	return (undefined !== this.Pr.Lock ? this.Pr.Lock : c_oAscSdtLockType.Unlocked);
};
CInlineLevelSdt.prototype.SetContentControlPr = function(oPr)
{
	if (!oPr)
		return;

	if (oPr && !(oPr instanceof CContentControlPr))
	{
		var oTemp = new CContentControlPr(c_oAscSdtLockType.Inline);
		oTemp.FillFromObject(oPr);
		oPr = oTemp;
	}

	oPr.SetToContentControl(this);

	if (this.IsAutoFitContent())
		this.GetLogicDocument().CheckFormAutoFit(this);
};
CInlineLevelSdt.prototype.GetContentControlPr = function()
{
	var oPr = new CContentControlPr(c_oAscSdtLevelType.Inline);
	oPr.FillFromContentControl(this);
	return oPr;
};
/**
 * Можно ли удалить данный контейнер
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.CanBeDeleted = function()
{
	if (this.IsFixedForm())
		return false;

	return (undefined === this.Pr.Lock || c_oAscSdtLockType.Unlocked === this.Pr.Lock || c_oAscSdtLockType.ContentLocked === this.Pr.Lock);
};
/**
 * Можно ли редактировать данный контейнер
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.CanBeEdited = function()
{
	if (!this.SkipSpecialLock && (this.IsCheckBox() || this.IsPicture() || this.IsDropDownList()))
		return false;

	return (undefined === this.Pr.Lock || c_oAscSdtLockType.Unlocked === this.Pr.Lock || c_oAscSdtLockType.SdtLocked === this.Pr.Lock);
};
/**
 * Проверяем, залочена ли данная форма
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsFormLocked = function()
{
	return !(undefined === this.Pr.Lock || c_oAscSdtLockType.Unlocked === this.Pr.Lock || c_oAscSdtLockType.ContentLocked === this.Pr.Lock);
};
//----------------------------------------------------------------------------------------------------------------------
// Функции совместного редактирования
//----------------------------------------------------------------------------------------------------------------------
CInlineLevelSdt.prototype.Write_ToBinary2 = function(Writer)
{
	Writer.WriteLong(AscDFH.historyitem_type_InlineLevelSdt);

	// String : Id
	// Long   : Количество элементов
	// Array of Strings : массив с Id элементов

	Writer.WriteString2(this.Id);

	var Count = this.Content.length;
	Writer.WriteLong(Count);
	for (var Index = 0; Index < Count; Index++)
		Writer.WriteString2(this.Content[Index].Get_Id());
};
CInlineLevelSdt.prototype.Read_FromBinary2 = function(Reader)
{
	// String : Id
	// Long   : Количество элементов
	// Array of Strings : массив с Id элементов

	this.Id = Reader.GetString2();

	var Count = Reader.GetLong();
	this.Content = [];
	for (var Index = 0; Index < Count; Index++)
	{
		var Element = AscCommon.g_oTableId.Get_ById(Reader.GetString2());
		if (null !== Element)
			this.Content.push(Element);
	}
};
CInlineLevelSdt.prototype.Write_ToBinary = function(Writer)
{
	// Long   : Type
	// String : Id

	Writer.WriteLong(this.Type);
	Writer.WriteString2(this.Id);
};
//----------------------------------------------------------------------------------------------------------------------
CInlineLevelSdt.prototype.IsStopCursorOnEntryExit = function()
{
	return true;
};
CInlineLevelSdt.prototype.GetSelectedContentControls = function(arrContentControls)
{
	arrContentControls.push(this);
	CParagraphContentWithParagraphLikeContent.prototype.GetSelectedContentControls.call(this, arrContentControls);
};
CInlineLevelSdt.prototype.ClearContentControl = function()
{
	var oRun = new ParaRun(this.GetParagraph(), false);
	oRun.SetPr(this.Pr.TextPr.Copy());
	this.AddToContent(0, oRun);
	this.RemoveFromContent(1, this.Content.length - 1);
	this.RemoveSelection();
	this.MoveCursorToStartPos();
};
CInlineLevelSdt.prototype.CanAddComment = function()
{
	if (!this.CanBeDeleted() || (!this.CanBeEdited() && (!this.IsSelectedAll() || this.IsSelectedOnlyThis())))
		return false;

	return CParagraphContentWithParagraphLikeContent.prototype.CanAddComment.apply(this, arguments);
};
/**
 * Проверяем выделен ли только данный элемент
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsSelectedOnlyThis = function()
{
	if (this.Paragraph && this.Paragraph.LogicDocument)
	{
		var oInfo = this.Paragraph.LogicDocument.GetSelectedElementsInfo();
		return (oInfo.GetInlineLevelSdt() === this);
	}

	return false;
};
/**
 * Проверяем, является ли данный контейнер чекбоксом
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsCheckBox = function()
{
	return !!(this.Pr.CheckBox);
};
/**
 * Применяем заданные настройки для чекобокса
 * @param {CSdtCheckBoxPr} oCheckBoxPr
 * @param {CTextPr} oTextPr
 */
CInlineLevelSdt.prototype.ApplyCheckBoxPr = function(oCheckBoxPr, oTextPr)
{
	if (undefined === this.Pr.CheckBox || !this.Pr.CheckBox.IsEqual(oCheckBoxPr))
	{
		if (this.IsPlaceHolder())
			this.private_ReplacePlaceHolderWithContent(false);

		this.SetCheckBoxPr(oCheckBoxPr);
	}

	if (this.IsCheckBox())
	{
		if (oTextPr)
		{
			if (this.Content[0] && para_Run === this.Content[0].Type)
				this.Content[0].SetPr(oTextPr);
		}

		this.SetPlaceholder(undefined);
		this.SetShowingPlcHdr(false);
		this.private_UpdateCheckBoxContent();
	}
};
/**
 * Выставляем настройки чекбокса
 * @param {CSdtCheckBoxPr} oCheckBoxPr
 */
CInlineLevelSdt.prototype.SetCheckBoxPr = function(oCheckBoxPr)
{
	if (undefined === this.Pr.CheckBox || !this.Pr.CheckBox.IsEqual(oCheckBoxPr))
	{
		var _oCheckBox = oCheckBoxPr ? oCheckBoxPr.Copy() : undefined;
		History.Add(new CChangesSdtPrCheckBox(this, this.Pr.CheckBox, _oCheckBox));
		this.Pr.CheckBox = _oCheckBox;
	}
};
/**
 * Получаем настройки для чекбокса
 * @returns {?CSdtCheckBoxPr}
 */
CInlineLevelSdt.prototype.GetCheckBoxPr = function()
{
	return this.Pr.CheckBox;
};
/**
 * Выставляем состояние чекбокса
 * @param {boolean|undefined} [isChecked=undefined]
 */
CInlineLevelSdt.prototype.ToggleCheckBox = function(isChecked)
{
	if (!this.IsCheckBox())
		return;

	if (undefined !== isChecked && this.Pr.CheckBox.Checked === isChecked)
		return;

	var oLogicDocument = this.GetLogicDocument();
	if (oLogicDocument && (this.IsRadioButton() || this.GetFormKey()))
		oLogicDocument.OnChangeForm(this.IsRadioButton() ? this.Pr.CheckBox.GroupKey : this.GetFormKey(), this);

	if (undefined === isChecked && this.IsRadioButton() && true === this.Pr.CheckBox.Checked)
		return;

	this.SetCheckBoxChecked(!this.Pr.CheckBox.Checked);
};
CInlineLevelSdt.prototype.SetCheckBoxChecked = function(isChecked)
{
	History.Add(new CChangesSdtPrCheckBoxChecked(this, this.Pr.CheckBox.Checked, isChecked));
	this.Pr.CheckBox.Checked = isChecked;

	this.private_UpdateCheckBoxContent();
};
/**
 * Выключаем проверку невозможности редактирования данного объекта, из-за того что специальный контейнер
 * @param isSkip {boolean}
 */
CInlineLevelSdt.prototype.SkipSpecialContentControlLock = function(isSkip)
{
	this.SkipSpecialLock = isSkip;
};
/**
 * @retuns {boolean}
 */
CInlineLevelSdt.prototype.IsSkipSpecialContentControlLock = function()
{
	return this.SkipSpecialLock;
};
CInlineLevelSdt.prototype.private_UpdateCheckBoxContent = function()
{
	var isChecked = this.Pr.CheckBox.Checked;

	var oRun;
	if (this.GetParagraph() && this.GetParagraph().GetLogicDocument() && this.GetParagraph().GetLogicDocument().IsTrackRevisions())
	{
		var oFirstRun = this.GetFirstRun();
		var oTextPr   = oFirstRun ? oFirstRun.GetDirectTextPr() : new CTextPr();

		this.SelectAll();
		this.Remove();
		this.RemoveSelection();

		oRun = new ParaRun(this.GetParagraph(), false);
		oRun.SetPr(oTextPr);
		this.AddToContent(0, oRun);

		if (2 === this.Content.length
			&& para_Run === this.Content[0].Type
			&& para_Run === this.Content[1].Type
			&& reviewtype_Add === this.Content[0].GetReviewType()
			&& reviewtype_Remove === this.Content[1].GetReviewType()
			&& this.Content[0].GetReviewInfo().IsCurrentUser()
			&& this.Content[1].GetReviewInfo().IsCurrentUser()
			&& ((isChecked
			&& String.fromCharCode(this.Pr.CheckBox.CheckedSymbol) === this.Content[1].GetText())
			|| (!isChecked
			&& String.fromCharCode(this.Pr.CheckBox.UncheckedSymbol) === this.Content[1].GetText())))
		{
			this.RemoveFromContent(1, 1);
			oRun.SetReviewType(reviewtype_Common);
		}
	}
	else
	{
		oRun = this.MakeSingleRunElement();
		if (!oRun)
			return;
	}

	oRun.AddText(String.fromCharCode(isChecked ? this.Pr.CheckBox.CheckedSymbol : this.Pr.CheckBox.UncheckedSymbol));

	if (isChecked && this.Pr.CheckBox.CheckedFont)
	{
		oRun.Set_RFonts_Ascii({Index : -1, Name : this.Pr.CheckBox.CheckedFont});
		oRun.Set_RFonts_HAnsi({Index : -1, Name : this.Pr.CheckBox.CheckedFont});
		oRun.Set_RFonts_CS({Index : -1, Name : this.Pr.CheckBox.CheckedFont});
		oRun.Set_RFonts_EastAsia({Index : -1, Name : this.Pr.CheckBox.CheckedFont});
	}
	else if (!isChecked && this.Pr.CheckBox.UncheckedFont)
	{
		oRun.Set_RFonts_Ascii({Index : -1, Name : this.Pr.CheckBox.UncheckedFont});
		oRun.Set_RFonts_HAnsi({Index : -1, Name : this.Pr.CheckBox.UncheckedFont});
		oRun.Set_RFonts_CS({Index : -1, Name : this.Pr.CheckBox.UncheckedFont});
		oRun.Set_RFonts_EastAsia({Index : -1, Name : this.Pr.CheckBox.UncheckedFont});
	}
};
/**
 * Проверяем, является ли данный класс специальным контейнером для картинки
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsPicture = function()
{
	return (!!this.Pr.Picture);
};
/**
 * Выставляем настройку того, что это контент контрол с картинкой
 * @param isPicture {boolean}
 */
CInlineLevelSdt.prototype.SetPicturePr = function(isPicture)
{
	if (this.Pr.Picture !== isPicture)
	{
		History.Add(new CChangesSdtPrPicture(this, this.Pr.Picture, isPicture));
		this.Pr.Picture = isPicture;
	}
};
CInlineLevelSdt.prototype.private_UpdatePictureContent = function(_nW, _nH)
{
	if (!this.IsPicture())
		return;

	if (this.IsPlaceHolder())
	{
		this.ReplacePlaceHolderWithContent();
		this.SetShowingPlcHdr(true);
	}

	var arrDrawings = this.GetAllDrawingObjects();

	var oRun = this.MakeSingleRunElement();
	if (!oRun)
		return;

	var oDrawing;
	for (var nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
	{
		if (arrDrawings[nIndex].IsPicture())
		{
			oDrawing = arrDrawings[nIndex];
			break;
		}
	}

	if (!oDrawing)
	{
		var oDrawingObjects = this.Paragraph && this.Paragraph.LogicDocument ? this.Paragraph.LogicDocument.DrawingObjects : null;
		if (!oDrawingObjects)
			return;

		var nW = _nW ? _nW : 50;
		var nH = _nH ? _nH : 50;

		oDrawing   = new ParaDrawing(nW, nH, null, oDrawingObjects, this.Paragraph.LogicDocument, null);
		var oImage = oDrawingObjects.createImage(AscCommon.g_sWordPlaceholderImage, 0, 0, nW, nH);
		oImage.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oImage);
	}

	oRun.AddToContent(0, oDrawing);
};
/**
 * Применяме к данному контейнеру настройку того, что это специальный контейнер для картинок
 * @param isPicture {boolean}
 * @param [nW=-1] {number} если не задано (или значение не положительное), тогда используется стандартное значение
 * @param [nH=-1] {number} если не задано (или значение не положительное), тогда используется стандартное значение
 */
CInlineLevelSdt.prototype.ApplyPicturePr = function(isPicture, nW, nH)
{
	this.SetPicturePr(isPicture);
	this.private_UpdatePictureContent(nW, nH);
};
/**
 * Выделяем изображение, если это специальный контейнер для изображения
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.SelectPicture = function()
{
	if (!this.IsPicture() || !this.GetParagraph() || !this.GetParagraph().GetParent())
		return false;

	var arrDrawings = this.GetAllDrawingObjects();
	if (arrDrawings.length <= 0)
		return false;

	this.GetParagraph().GetParent().Select_DrawingObject(arrDrawings[0].GetId());
	return true;
};
/**
 * Проверяем, является ли данный класс специальной формой для картинок
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsPictureForm = function()
{
	return (this.IsForm() && this.IsPicture() && undefined !== this.Pr.PictureFormPr);
};
/**
 * Выставляем настройк для формы с картинкой
 * @param oPr {CSdtPictureFormPr}
 */
CInlineLevelSdt.prototype.SetPictureFormPr = function(oPr)
{
	if (undefined === this.Pr.PictureFormPr || !this.Pr.PictureFormPr.IsEqual(oPr))
	{
		var _oPr = oPr ? oPr.Copy() : undefined;
		History.Add(new CChangesSdtPrPictureFormPr(this, this.Pr.PictureFormPr, _oPr));
		this.Pr.PictureFormPr = _oPr;
	}
};
/**
 * Получаем настройки для картиночной формы
 * @returns {?CSdtPictureFormPr}
 */
CInlineLevelSdt.prototype.GetPictureFormPr = function()
{
	return this.Pr.PictureFormPr;
};
/**
 * Проверяем является ли данный контейнер специальным для поля со списком
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsComboBox = function()
{
	return (undefined !== this.Pr.ComboBox);
};
/**
 * @param oPr {CSdtComboBoxPr}
 */
CInlineLevelSdt.prototype.SetComboBoxPr = function(oPr)
{
	if (undefined === this.Pr.ComboBox || !this.Pr.ComboBox.IsEqual(oPr))
	{
		var _oPr = oPr ? oPr.Copy() : undefined;
		History.Add(new CChangesSdtPrComboBox(this, this.Pr.ComboBox, _oPr));
		this.Pr.ComboBox = _oPr;
	}
};
/**
 * @returns {?CSdtComboBoxPr}
 */
CInlineLevelSdt.prototype.GetComboBoxPr = function()
{
	return this.Pr.ComboBox;
};
/**
 * Проверяем является ли данный контейнер специальным для выпадающего списка
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsDropDownList = function()
{
	return (undefined !== this.Pr.DropDown);
};
/**
 * @param oPr {CSdtComboBoxPr}
 */
CInlineLevelSdt.prototype.SetDropDownListPr = function(oPr)
{
	if (undefined === this.Pr.DropDown || !this.Pr.DropDown.IsEqual(oPr))
	{
		var _oPr = oPr ? oPr.Copy() : undefined;
		History.Add(new CChangesSdtPrDropDownList(this, this.Pr.DropDown, _oPr));
		this.Pr.DropDown = _oPr;
	}
};
/**
 * @returns {?CSdtComboBoxPr}
 */
CInlineLevelSdt.prototype.GetDropDownListPr = function()
{
	return this.Pr.DropDown;
};
/**
 * Применяем к данному контейнеру настройки того, что это специальный контйенер для поля со списком
 * @param oPr {CSdtComboBoxPr}
 */
CInlineLevelSdt.prototype.ApplyComboBoxPr = function(oPr)
{
	this.SetPlaceholder(c_oAscDefaultPlaceholderName.List);
	if (this.IsPlaceHolder())
		this.private_FillPlaceholderContent();

	this.SetComboBoxPr(oPr);
	this.SelectListItem();
};
/**
 * Применяем к данному контейнеру настройки того, что это специальный контейнер для выпадающего списка
 * @param oPr {CSdtComboBoxPr}
 */
CInlineLevelSdt.prototype.ApplyDropDownListPr = function(oPr)
{
	this.SetPlaceholder(c_oAscDefaultPlaceholderName.List);
	if (this.IsPlaceHolder())
		this.private_FillPlaceholderContent();

	this.SetDropDownListPr(oPr);
	this.SelectListItem();
};
/**
 * Заполняем контейнер текстом в зависимости от выбранного элемента в списке
 * @param sValue {string}
 */
CInlineLevelSdt.prototype.SelectListItem = function(sValue)
{
	var oList = null;
	if (this.IsComboBox())
		oList = this.Pr.ComboBox;
	else if (this.IsDropDownList())
		oList = this.Pr.DropDown;

	if (!oList)
		return;

	var sText = oList.GetTextByValue(sValue);

	if (this.GetParagraph() && this.GetParagraph().GetLogicDocument() && this.GetParagraph().GetLogicDocument().IsTrackRevisions())
	{
		if (!sText && this.IsPlaceHolder())
		{
			this.private_FillPlaceholderContent();
			return;
		}

		var oFirstRun = this.GetFirstRun();
		var oTextPr   = oFirstRun ? oFirstRun.GetDirectTextPr() : new CTextPr();

		if (!this.IsPlaceHolder())
		{
			this.SelectAll();
			this.Remove();
			this.RemoveSelection();
		}
		else
		{
			this.ReplacePlaceHolderWithContent();
		}

		if (!sText && this.IsEmpty())
		{
			this.ReplaceContentWithPlaceHolder();
		}

		if (sText)
		{
			var oRun;
			if (this.IsEmpty())
			{
				oRun = this.MakeSingleRunElement();
				if (!oRun)
					return;

				oRun.SetReviewType(reviewtype_Add);
			}
			else
			{
				oRun = new ParaRun(this.GetParagraph(), false);
				this.AddToContent(this.GetContentLength(), oRun);
			}

			oRun.SetPr(oTextPr);
			oRun.AddText(sText);
		}
	}
	else
	{
		if (null === sText)
		{
			this.ReplaceContentWithPlaceHolder();
		}
		else
		{
			this.ReplacePlaceHolderWithContent();
			var oRun = this.private_UpdateListContent();
			if (oRun)
				oRun.AddText(sText);
		}
	}
};
CInlineLevelSdt.prototype.private_UpdateListContent = function()
{
	if (this.IsPlaceHolder())
		return null;

	return this.MakeSingleRunElement();
};
/**
 * Проверяем является ли данный контейнер специальным для даты
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsDatePicker = function()
{
	return (undefined !== this.Pr.Date);
};
/**
 * @param oPr {CSdtDatePickerPr}
 */
CInlineLevelSdt.prototype.SetDatePickerPr = function(oPr)
{
	if (undefined === this.Pr.Date || !this.Pr.Date.IsEqual(oPr))
	{
		var _oPr = oPr ? oPr.Copy() : undefined;
		History.Add(new CChangesSdtPrDatePicker(this, this.Pr.Date, _oPr));
		this.Pr.Date = _oPr;
	}
};
/**
 * @returns {?CSdtDatePickerPr}
 */
CInlineLevelSdt.prototype.GetDatePickerPr = function()
{
	return this.Pr.Date;
};
/**
 * Применяем к данному контейнеру настройки того, что это специальный контйенер для даты
 * @param oPr {CSdtDatePickerPr}
 */
CInlineLevelSdt.prototype.ApplyDatePickerPr = function(oPr)
{
	this.SetDatePickerPr(oPr);

	if (!this.IsDatePicker())
		return;

	this.SetPlaceholder(c_oAscDefaultPlaceholderName.DateTime);
	if (this.IsPlaceHolder())
		this.private_FillPlaceholderContent();

	this.private_UpdateDatePickerContent();
};
CInlineLevelSdt.prototype.private_UpdateDatePickerContent = function()
{
	if (!this.Pr.Date)
		return;

	if (this.IsPlaceHolder())
		this.ReplacePlaceHolderWithContent();

	var oRun;
	var sText = this.Pr.Date.ToString();
	if (this.GetParagraph() && this.GetParagraph().GetLogicDocument() && this.GetParagraph().GetLogicDocument().IsTrackRevisions())
	{
		if (!sText && this.IsPlaceHolder())
			return;

		var oFirstRun = this.GetFirstRun();
		var oTextPr   = oFirstRun ? oFirstRun.GetDirectTextPr() : new CTextPr();

		if (!this.IsPlaceHolder())
		{
			this.SelectAll();
			this.Remove();
			this.RemoveSelection();
		}
		else
		{
			this.ReplacePlaceHolderWithContent();
		}

		if (!sText && this.IsEmpty())
		{
			this.ReplaceContentWithPlaceHolder();
		}

		if (sText)
		{
			var oRun;
			if (this.IsEmpty())
			{
				oRun = this.MakeSingleRunElement();
				if (!oRun)
					return;

				oRun.SetReviewType(reviewtype_Add);
			}
			else
			{
				oRun = new ParaRun(this.GetParagraph(), false);
				this.AddToContent(this.GetContentLength(), oRun);
			}

			oRun.SetPr(oTextPr);
		}
	}
	else
	{
		oRun = this.MakeSingleRunElement();
	}

	if (oRun)
		oRun.AddText(sText);
};
/**
 * Является ли данный контейнер специальной текстовой формой
 * @returns {boolean}
 */
CInlineLevelSdt.prototype.IsTextForm = function()
{
	return (undefined !== this.Pr.TextForm);
};
CInlineLevelSdt.prototype.SetTextFormPr = function(oPr)
{
	if (undefined === this.Pr.TextForm || !this.Pr.TextForm.IsEqual(oPr))
	{
		var _oPr = oPr ? oPr.Copy() : undefined;
		History.Add(new CChangesSdtPrTextForm(this, this.Pr.TextForm, _oPr));
		this.Pr.TextForm = _oPr;

		this.Recalc_RunsCompiledPr();
	}
};
CInlineLevelSdt.prototype.GetTextFormPr = function()
{
	return this.Pr.TextForm;
};
/**
 * Применяем к данному контейнеру настройки того, что это специальный контйенер для даты
 * @param oPr {CSdtDatePickerPr}
 */
CInlineLevelSdt.prototype.ApplyTextFormPr = function(oPr)
{
	this.SetTextFormPr(oPr);

	if (!this.IsTextForm())
		return;

	if (this.IsPlaceHolder())
		this.private_FillPlaceholderContent();
	else
		this.private_UpdateTextFormContent();
};
CInlineLevelSdt.prototype.private_UpdateTextFormContent = function()
{
	if (!this.Pr.TextForm)
		return;

	if (this.IsPlaceHolder())
		this.ReplacePlaceHolderWithContent();

	this.MakeSingleRunElement();
};
CInlineLevelSdt.prototype.Document_Is_SelectionLocked = function(CheckType)
{
	if (this.GetFormKey() && this.GetLogicDocument())
		this.GetLogicDocument().CheckSelectionLockedByFormKey(CheckType, this.GetFormKey(), this.GetParagraph());

	if (AscCommon.changestype_Paragraph_TextProperties === CheckType
		|| ((AscCommon.changestype_Drawing_Props === CheckType || AscCommon.changestype_Image_Properties === CheckType)
		&& this.IsPicture()))
	{
		this.SkipSpecialContentControlLock(true);
		if (!this.CanBeEdited())
			AscCommon.CollaborativeEditing.Add_CheckLock(true);
		this.SkipSpecialContentControlLock(false);

		return;
	}

	var oLogicDocument = this.Paragraph ? this.Paragraph.LogicDocument : null;
	if (oLogicDocument)
	{
		if (!oLogicDocument.IsCheckContentControlsLock())
			return;

		if (oLogicDocument.IsFillingFormMode()
			&& !this.IsSelectionUse()
			&& ((this.IsPlaceHolder() && (AscCommon.changestype_Remove === CheckType || AscCommon.changestype_Delete === CheckType))
				|| (this.Cursor_Is_Start() && AscCommon.changestype_Remove === CheckType)
				|| (this.Cursor_Is_End() && AscCommon.changestype_Delete === CheckType))
		)
		{
			return AscCommon.CollaborativeEditing.Add_CheckLock(true);
		}
	}

	var nContentControlLock = this.GetContentControlLock();

	if ((AscCommon.changestype_Paragraph_Content === CheckType
		|| AscCommon.changestype_Paragraph_AddText === CheckType
		|| AscCommon.changestype_ContentControl_Add === CheckType
		|| AscCommon.changestype_Remove === CheckType
		|| AscCommon.changestype_Delete === CheckType
		|| AscCommon.changestype_Document_Content === CheckType
		|| AscCommon.changestype_Document_Content_Add === CheckType)
		&& this.IsSelectionUse()
		&& this.IsSelectedAll())
	{
		var bSelectedOnlyThis = false;

		if ((AscCommon.changestype_Remove === CheckType
			|| AscCommon.changestype_Delete === CheckType)
			&& oLogicDocument
			&& oLogicDocument.IsCheckContentControlsLock()
			&& ((this.IsPlaceHolder() && oLogicDocument.IsFillingFormMode())
				|| (!this.CanBeEdited() && (oLogicDocument.IsFillingFormMode() || this.IsFixedForm()))))
		{
			return AscCommon.CollaborativeEditing.Add_CheckLock(true);
		}

		// Если действие происходит на удалении (del/backspace), тогда мы должны проверить, что это не PlaceHolder
		// и что элемент может быть отредактирован и что выделен только данный элемент. (для тех случаев, когда
		// элемент можно привести к PlaceHolder через удаление)
		// Если это происходит на добавлении текста, тогда проверяем, что выделен только данный элемент
		if ((!this.IsPlaceHolder() && this.CanBeEdited()) || (AscCommon.changestype_Remove !== CheckType && AscCommon.changestype_Delete !== CheckType))
		{
			var oInfo = this.Paragraph.LogicDocument.GetSelectedElementsInfo();
			bSelectedOnlyThis = (oInfo.GetInlineLevelSdt() === this);

			if (bSelectedOnlyThis && oInfo.IsFixedFormShape())
				bSelectedOnlyThis = false;
		}

		if (c_oAscSdtLockType.SdtContentLocked === nContentControlLock
			|| (c_oAscSdtLockType.SdtLocked === nContentControlLock && true !== bSelectedOnlyThis)
			|| (!this.CanBeEdited() && true === bSelectedOnlyThis))
		{
			return AscCommon.CollaborativeEditing.Add_CheckLock(true);
		}
	}
	else if ((AscCommon.changestype_Paragraph_Content === CheckType
		|| AscCommon.changestype_Paragraph_AddText === CheckType
		|| AscCommon.changestype_ContentControl_Add === CheckType
		|| AscCommon.changestype_Remove === CheckType
		|| AscCommon.changestype_Delete === CheckType
		|| AscCommon.changestype_Document_Content === CheckType
		|| AscCommon.changestype_Document_Content_Add === CheckType
		|| AscCommon.changestype_Image_Properties === CheckType
		|| AscCommon.changestype_Drawing_Props === CheckType)
		&& !this.CanBeEdited())
	{
		return AscCommon.CollaborativeEditing.Add_CheckLock(true);
	}
};
/**
 * Получаем типа данного контейнера
 * @returns {Asc.c_oAscContentControlSpecificType}
 */
CInlineLevelSdt.prototype.GetSpecificType = function()
{
	if (this.IsCheckBox())
		return Asc.c_oAscContentControlSpecificType.CheckBox;

	if (this.IsPicture())
		return Asc.c_oAscContentControlSpecificType.Picture;

	if (this.IsComboBox())
		return Asc.c_oAscContentControlSpecificType.ComboBox;

	if (this.IsDropDownList())
		return Asc.c_oAscContentControlSpecificType.DropDownList;

	if (this.IsDatePicker())
		return Asc.c_oAscContentControlSpecificType.DateTime;

	return Asc.c_oAscContentControlSpecificType.None;
};
CInlineLevelSdt.prototype.Get_ParentTextTransform = function()
{
	return this.Paragraph.Get_ParentTextTransform();
};
CInlineLevelSdt.prototype.AcceptRevisionChanges = function(Type, bAll)
{
	if (this.IsCheckBox() || this.IsDropDownList() || this.IsComboBox() || this.IsPicture() || this.IsDatePicker() || this.IsTextForm())
	{
		Type = undefined;
		bAll = true;
	}

	CParagraphContentWithParagraphLikeContent.prototype.AcceptRevisionChanges.call(this, Type, bAll);
};
CInlineLevelSdt.prototype.RejectRevisionChanges = function(Type, bAll)
{
	if (this.IsCheckBox() || this.IsDropDownList() || this.IsComboBox() || this.IsPicture() || this.IsDatePicker() || this.IsTextForm())
	{
		Type = undefined;
		bAll = true;
	}

	CParagraphContentWithParagraphLikeContent.prototype.RejectRevisionChanges.call(this, Type, bAll);
};
CInlineLevelSdt.prototype.SetPlaceholder = function(sDocPartName)
{
	if (this.Pr.Placeholder !== sDocPartName)
	{
		History.Add(new CChangesSdtPrPlaceholder(this, this.Pr.Placeholder, sDocPartName));
		this.Pr.Placeholder = sDocPartName;
	}
};
CInlineLevelSdt.prototype.GetPlaceholder = function()
{
	return this.Pr.Placeholder;
};
CInlineLevelSdt.prototype.SetShowingPlcHdr = function(isShow)
{
	if (this.Pr.ShowingPlcHdr !== isShow)
	{
		History.Add(new CChangesSdtPrShowingPlcHdr(this, this.Pr.ShowingPlcHdr, isShow));
		this.Pr.ShowingPlcHdr = isShow;
	}
};
CInlineLevelSdt.prototype.IsShowingPlcHdr = function()
{
	return this.Pr.ShowingPlcHdr;
};
CInlineLevelSdt.prototype.GetLogicDocument = function()
{
	var oParagraph = this.GetParagraph();
	return oParagraph ? oParagraph.GetLogicDocument() : editor.WordControl.m_oLogicDocument;
};
CInlineLevelSdt.prototype.private_OnAddFormPr = function()
{
	this.Recalc_RunsCompiledPr();
};
CInlineLevelSdt.prototype.UpdatePlaceHolderTextPrForForm = function()
{
	if (this.IsPlaceHolder() && this.IsForm() && this.Pr.TextPr)
	{
		for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			if (this.Content[nIndex] instanceof ParaRun)
				this.Content[nIndex].SetPr(this.Pr.TextPr.Copy());
		}
	}
};
CInlineLevelSdt.prototype.CheckHitInContentControlByXY = function(X, Y, nPageAbs, isUseTransform)
{
	var oParagraph = this.GetParagraph();
	if (!oParagraph)
		return false;

	var oTransform = this.Get_ParentTextTransform();

	var _X = X;
	var _Y = Y;
	if (oTransform && false !== isUseTransform)
	{
		oTransform = oTransform.Invert();
		_X = oTransform.TransformPointX(X, Y);
		_Y = oTransform.TransformPointY(X, Y);
	}

	if (this.IsFixedForm())
	{
		var oShape  = oParagraph.Parent.Is_DrawingShape(true);
		var oBounds = oShape.getFormRelRect();

		if (oBounds.Page === nPageAbs && oBounds.X <= _X && _X <= oBounds.X + oBounds.W && oBounds.Y <= _Y && _Y <= oBounds.Y + oBounds.H)
			return true;
	}
	else
	{
		for (var sKey in this.Bounds)
		{
			var oBounds = this.Bounds[sKey];
			if (oParagraph.GetAbsolutePage(oBounds.PageInternal) === nPageAbs && oBounds.X <= _X && _X <= oBounds.X + oBounds.W && oBounds.Y <= _Y && _Y <= oBounds.Y + oBounds.H)
				return true;
		}
	}

	return false;
};
CInlineLevelSdt.prototype.CorrectXYToHitIn = function(X, Y, nPageAbs, isUseTransform)
{
	var oParagraph = this.GetParagraph();
	if (!oParagraph)
		return false;

	var oTransform = this.Get_ParentTextTransform();

	var _X = X;
	var _Y = Y;
	if (oTransform && false !== isUseTransform)
	{
		oTransform = oTransform.Invert();
		_X = oTransform.TransformPointX(X, Y);
		_Y = oTransform.TransformPointY(X, Y);
	}

	function private_Diff(oBound, X, Y)
	{
		var _X = 0;
		var _Y = 0;

		if (oBound.Y <= Y && Y <= oBound.Y + oBound.H)
			_Y = 0;
		else if (Y < oBound.Y)
			_Y = oBound.Y - Y;
		else
			_Y = Y - oBound.Y - oBound.H;

		if (oBound.X <= X && X <= oBound.X + oBound.W)
			_X = 0;
		else if (X < oBound.X)
			_X = oBound.X - X;
		else
			_X = X - oBound.X - oBound.W;

		return {X : _X, Y : _Y};
	}

	var oDiff      = null;
	var oNearBound = null;
	for (var sKey in this.Bounds)
	{
		var oBound = this.Bounds[sKey];
		if (oParagraph.GetAbsolutePage(oBound.PageInternal) === nPageAbs)
		{
			if (!oNearBound)
			{
				oDiff      = private_Diff(oBound, _X, _Y);
				oNearBound = oBound;
			}
			else
			{
				var _oDiff = private_Diff(oBound, _X, _Y);
				if ((0 === oDiff.Y && 0 === _oDiff.Y && _oDiff.X < oDiff.Y)
					|| (0 !== oDiff.Y
						&& (0 === _oDiff.Y
							|| (_oDiff.Y < oDiff.Y)
							|| (Math.abs(_oDiff.Y - oDiff.Y) < 0.001 && _oDiff.X < oDiff.X))))
				{
					oDiff      = _oDiff;
					oNearBound = oBound;
				}
			}

			if (oDiff && 0 === oDiff.Y && 0 === oDiff.X)
				return {X : X, Y : Y};
		}
	}

	if (oNearBound)
	{
		var __X, __Y;

		if (oNearBound.Y <= _Y && _Y <= oNearBound.Y + oNearBound.H)
			__Y = _Y;
		else if (_Y < oNearBound.Y)
			__Y = oNearBound.Y + 0.001;
		else
			__Y = oNearBound.Y + oNearBound.H - 0.001;

		if (oNearBound.X <= _X && _X <= oNearBound.X + oNearBound.W)
			__X = _X;
		else if (_X < oNearBound.X)
			__X = oNearBound.X + 0.001;
		else
			__X = oNearBound.X + oNearBound.W - 0.001;

		if (oTransform)
		{
			oTransform = oTransform.Invert();
			_X = oTransform.TransformPointX(__X, __Y);
			_Y = oTransform.TransformPointY(__X, __Y);
		}
		else
		{
			_X = __X;
			_Y = __Y;
		}

		return {X : _X, Y : _Y};
	}

	return null;
};
CInlineLevelSdt.prototype.IntersectWithRect = function(X, Y, W, H, nPageAbs)
{
	var arrRects = [];

	if (this.IsFixedForm())
	{
		var oShape  = this.Paragraph.Parent.Is_DrawingShape(true);
		var oBounds = oShape.getFormRelRect();

		if (nPageAbs === oBounds.Page)
		{
			var nLeft   = Math.max(X, oBounds.X);
			var nRight  = Math.min(X + W, oBounds.X + oBounds.W);
			var nTop    = Math.max(Y, oBounds.Y);
			var nBottom = Math.min(Y + H, oBounds.Y + oBounds.H);

			if (nLeft < nRight && nTop < nBottom)
			{
				arrRects.push({
					X : nLeft,
					Y : nTop,
					W : nRight - nLeft,
					H : nBottom - nTop
				});
			}
		}
	}
	else
	{
		var oParagraph = this.GetParagraph();
		if (!oParagraph)
			return [];

		for (var sKey in this.Bounds)
		{
			var oBound = this.Bounds[sKey];
			if (oParagraph.GetAbsolutePage(oBound.PageInternal) === nPageAbs)
			{
				var nLeft   = Math.max(X, oBound.X);
				var nRight  = Math.min(X + W, oBound.X + oBound.W);
				var nTop    = Math.max(Y, oBound.Y);
				var nBottom = Math.min(Y + H, oBound.Y + oBound.H);

				if (nLeft < nRight && nTop < nBottom)
				{
					arrRects.push({
						X : nLeft,
						Y : nTop,
						W : nRight - nLeft,
						H : nBottom - nTop
					});
				}
			}
		}
	}

	return arrRects;
};
CInlineLevelSdt.prototype.IsSelectedAll = function(Props)
{
	if (!this.Selection.Use)
		return false;

	if (this.IsPlaceHolder())
		return true;

	return CParagraphContentWithParagraphLikeContent.prototype.IsSelectedAll.apply(this, arguments);
};
CInlineLevelSdt.prototype.IsSelectedFromStart = function()
{
	if (!this.Selection.Use)
		return false;

	if (this.IsPlaceHolder())
		return true;

	return CParagraphContentWithParagraphLikeContent.prototype.IsSelectedFromStart.apply(this, arguments);
};
CInlineLevelSdt.prototype.IsSelectedToEnd = function()
{
	if (!this.Selection.Use)
		return false;

	if (this.IsPlaceHolder())
		return true;

	return CParagraphContentWithParagraphLikeContent.prototype.IsSelectedToEnd.apply(this, arguments);
};
CInlineLevelSdt.prototype.IsFormFilled = function()
{
	if (!this.IsForm())
		return true;

	if (this.IsPlaceHolder())
		return false;

	if (this.IsTextForm())
	{
		var sText = this.GetSelectedText(true);
		var oTextFormPr = this.GetTextFormPr();

		if (oTextFormPr.IsComb())
		{
			var nMaxLen = oTextFormPr.GetMaxCharacters();
			return (nMaxLen === sText.length);
		}
		else
		{
			return (sText !== "");
		}
	}
	else if (this.IsComboBox())
	{
		var sText = this.GetSelectedText(true);
		return (sText !== "");
	}
	else if (this.IsDropDownList())
	{
		var sText = this.GetSelectedText(true);

		var oListPr = this.GetDropDownListPr();
		for (var nIndex = 0, nCount = oListPr.GetItemsCount(); nIndex < nCount; ++nIndex)
		{
			if (sText === oListPr.GetItemDisplayText(nIndex))
				return true;
		}

		return false;
	}
	else if (this.IsPicture())
	{
		// Мы уже проверили, что тут не Placeholder, значит форма заполнена
		return true;
	}
	else if (this.IsCheckBox() && !this.IsRadioButton())
	{
		return this.GetCheckBoxPr().GetChecked();
	}
	else if (this.IsCheckBox() && this.IsRadioButton())
	{
		var oLogicDocument = this.GetLogicDocument();
		if (!oLogicDocument)
			return false;

		var arrRadioGroup = oLogicDocument.GetSpecialRadioButtons(this.GetCheckBoxPr().GetGroupKey());
		for (var nIndex = 0, nCount = arrRadioGroup.length; nIndex < nCount; ++nIndex)
		{
			var oRadioForm = arrRadioGroup[nIndex];
			if (oRadioForm.GetCheckBoxPr().GetChecked())
				return true;
		}

		return false;
	}

	return false;
}
CInlineLevelSdt.prototype.ConvertFormToFixed = function(nW, nH)
{
	var oParagraph        = this.GetParagraph();
	var oParent           = this.GetParent();
	var oParentDocContent = oParagraph ? oParagraph.GetParent() : null;
	var oLogicDocument    = oParagraph ? oParagraph.GetLogicDocument() : null;
	var nPosInParent      = this.GetPosInParent(oParent);

	if (undefined === nW)
	{
		nW = 50;
		nH = 50;

		if (oParent)
		{
			for (var Key in this.Bounds)
			{
				if (this.Bounds[Key].W > 0.001 && this.Bounds[Key].H > 0.001)
				{
					nW = this.Bounds[Key].W + 0.5;
					nH = this.Bounds[Key].H + 0.1;
					break;
				}
			}
		}
	}

	// Для билдера, чтобы мы могли конвертить форму, даже если она нигде не лежит
	if (!oParent)
		return this.private_ConvertFormToFixed(nW, nH);

	if (!oParagraph
		|| !oParent
		|| !oParentDocContent
		|| !oLogicDocument
		|| -1 === nPosInParent
		|| oParagraph.IsInFixedForm())
		return null;

	let oParaDrawing = this.private_ConvertFormToFixed(nW, nH);

	var oRun = new ParaRun(oParagraph, false);
	oRun.AddToContent(0, oParaDrawing);

	// Этот код выравнивает позицию рана по вертикали, чтобы после конвертации типа формы текст внутри автофигуры
	// визуально оставался на месте, но сама настройка позиции по вертикали вызывает много непонятных ситуаций у
	// пользователей (баг 55524)
	//
	// if (this.Content.length > 0 && this.Content[0] instanceof ParaRun)
	// {
	// 	var oInnerRun = this.Content[0];
	// 	var oTextPr   = oInnerRun.Get_CompiledPr(false);
	//
	// 	g_oTextMeasurer.SetTextPr(oTextPr, oParagraph.GetTheme());
	// 	g_oTextMeasurer.SetFontSlot(fontslot_ASCII);
	//
	// 	var nTextDescent = Math.abs(g_oTextMeasurer.GetDescender());
	// 	oRun.Set_Position(oTextPr.Position - nTextDescent);
	// 	oInnerRun.Recalc_CompiledPr(true);
	// }

	oParent.RemoveFromContent(nPosInParent, 1, true);
	oParent.AddToContent(nPosInParent, oRun, true);

	if (this.IsAutoFitContent())
		oLogicDocument.CheckFormAutoFit(this);

	return oParaDrawing;
};
CInlineLevelSdt.prototype.private_ConvertFormToFixed = function(nW, nH)
{
	var oShape = new AscFormat.CShape();
	oShape.setWordShape(true);
	oShape.setBDeleted(false);

	var oSpPr = new AscFormat.CSpPr();
	var oXfrm = new AscFormat.CXfrm();
	oXfrm.setOffX(0);
	oXfrm.setOffY(0);
	oXfrm.setExtX(nW);
	oXfrm.setExtY(nH);
	oSpPr.setXfrm(oXfrm);
	oXfrm.setParent(oSpPr);

	oSpPr.setFill(AscFormat.CreateNoFillUniFill());
	oSpPr.setLn(AscFormat.CreateNoFillLine());
	oSpPr.setGeometry(AscFormat.CreateGeometry("rect"));
	oSpPr.setParent(oShape);
	oShape.setSpPr(oSpPr);

	oShape.createTextBoxContent();
	var oContent = oShape.getDocContent();

	var oInnerParagraph = oContent.GetElement(0);
	if (!oInnerParagraph)
		return null;

	oInnerParagraph.MoveCursorToStartPos();
	oInnerParagraph.Add(this);
	oInnerParagraph.SetParagraphAlign(this.IsCheckBox() ? AscCommon.align_Center : AscCommon.align_Left);
	oInnerParagraph.SetParagraphSpacing({Before : 0, After : 0, Line : 1, LineRule : Asc.linerule_Auto});

	var oBodyPr = oShape.getBodyPr().createDuplicate();

	oBodyPr.spcFirstLastPara = false;
	oBodyPr.vertOverflow     = AscFormat.nOTOwerflow;
	oBodyPr.horzOverflow     = AscFormat.nOTOwerflow;
	oBodyPr.vert             = AscFormat.nVertTThorz;

	oBodyPr.rot         = 0;
	oBodyPr.lIns        = 0.0;
	oBodyPr.tIns        = 0.0;
	oBodyPr.rIns        = 0.0;
	oBodyPr.bIns        = 0.0;
	oBodyPr.numCol      = 1;
	oBodyPr.spcCol      = 0;
	oBodyPr.rtlCol      = 0;
	oBodyPr.fromWordArt = false;
	oBodyPr.anchor      = 1;
	oBodyPr.anchorCtr   = false;
	oBodyPr.forceAA     = false;
	oBodyPr.compatLnSpc = true;
	oBodyPr.prstTxWarp  = null;

	oShape.setBodyPr(oBodyPr);

	var oParaDrawing = new ParaDrawing(oShape.spPr.xfrm.extX, oShape.spPr.xfrm.extY, oShape, editor.WordControl.m_oDrawingDocument, null, null);
	oShape.setParent(oParaDrawing);
	oParaDrawing.Set_DrawingType(drawing_Inline);
	oParaDrawing.SetForm(true);

	var oTextPr = this.GetTextFormPr();
	if (oTextPr && oTextPr.GetMultiLine())
	{
		var oNewTextPr = oTextPr.Copy();
		oNewTextPr.SetMultiLine(false);
		this.SetTextFormPr(oNewTextPr);
	}

	return oParaDrawing;
};
CInlineLevelSdt.prototype.ConvertFormToInline = function()
{
	var oParagraph   = this.GetParagraph();
	var oParent      = this.GetParent();
	var nPosInParent = this.GetPosInParent(oParent);
	if (!oParagraph || !oParent || !oParagraph.IsInFixedForm() || -1 === nPosInParent || this.IsPicture())
		return null;

	var oShape = oParagraph.Parent.Is_DrawingShape(true);
	if (!oShape || !oShape.parent)
		return null;

	var oTextPr = this.GetTextFormPr();
	if (oTextPr && oTextPr.GetMultiLine())
	{
		var oNewTextPr = oTextPr.Copy();
		oNewTextPr.SetMultiLine(false);
		this.SetTextFormPr(oNewTextPr);
	}

	var oParaDrawing = oShape.parent;
	var oRun = oParaDrawing.GetRun();
	if (!oRun || !oRun.GetParent())
	{
		// Это специальная ветка для билдера, чтобы дать возможность конвертить форму, даже если она никуда не добавлена
		oParent.RemoveFromContent(nPosInParent, 1, true);
		this.SetParent(null);
		this.SetParagraph(null);
		return this;
	}

	var oRunParent = oRun.GetParent();
	var nInRunParentPos = oRun.GetPosInParent(oRunParent);
	if (!oRunParent || -1 === nInRunParentPos)
		return null;

	if (1 === oRun.GetElementsCount())
	{
		oParent.RemoveFromContent(nPosInParent, 1, true);
		oRunParent.RemoveFromContent(nInRunParentPos, 1, true);
		oRunParent.AddToContent(nInRunParentPos, this, true);
	}
	else
	{
		var nInRunPos = -1;
		for (var nPos = 0, nRunLen = oRun.GetElementsCount(); nPos < nRunLen; ++nPos)
		{
			if (oRun.GetElement() === oParaDrawing)
			{
				nInRunPos = nPos;
				break;
			}
		}

		if (-1 === nInRunPos)
			return null;

		oParent.RemoveFromContent(nPosInParent, 1, true);
		oRun.RemoveFromContent(nInRunPos, 1);
		oRunParent.AddToContent(nInRunParentPos, this, true);
	}

	var oTextPr = this.GetTextFormPr();
	if (oTextPr && oTextPr.GetMultiLine())
	{
		var oNewTextPr = oTextPr.Copy();
		oNewTextPr.SetMultiLine(false);
		this.SetTextFormPr(oNewTextPr);
	}

	return this;
};
CInlineLevelSdt.prototype.IsMultiLineForm = function()
{
	var oTextFormPr = this.GetTextFormPr();
	if (oTextFormPr)
		return oTextFormPr.GetMultiLine();

	return true;
};
CInlineLevelSdt.prototype.OnChangeFixedFormTrack = function(nW, nH)
{
	if (!this.IsForm())
		return;

	var oParagraph = this.GetParagraph();
	if (!oParagraph || !oParagraph.Parent)
		return;

	var oShape = oParagraph.Parent.Is_DrawingShape(true);
	if (!oShape)
		return;

	if (this.IsCheckBox())
	{
		var oRun = this.Content[0];

		var oTextPr = oRun.Get_CompiledPr(false);

		g_oTextMeasurer.SetTextPr(oTextPr, oParagraph.GetTheme());
		g_oTextMeasurer.SetFontSlot(fontslot_ASCII);

		var nTextHeight = g_oTextMeasurer.GetHeight();

		var nKoef = 1.2 * Math.min(nW, nH * 0.8) / nTextHeight;

		var nFontSize    = oTextPr.FontSize;
		var nNewFontSize = nFontSize * nKoef;
		oRun.Set_FontSize(nNewFontSize);
	}
	else if (this.IsPicture())
	{
		this.private_UpdatePictureFormLayout(nW, nH);
	}
	else if (this.IsTextForm())
	{
		var oTextFormPr = this.GetTextFormPr();
		var nMax = oTextFormPr.GetMaxCharacters();
		if (oTextFormPr.IsComb() && nMax > 0)
		{
			var oNewTextFormPr = oTextFormPr.Copy();
			oNewTextFormPr.SetWidth(AscCommon.MMToTwips(nW / nMax));
			this.SetTextFormPr(oNewTextFormPr);
		}
	}
};
CInlineLevelSdt.prototype.IsAutoFitContent = function()
{
	if (!this.IsForm())
		return false;

	var oTextPr = this.GetTextFormPr();
	if (oTextPr)
		return oTextPr.GetAutoFit();

	var oComboPr = this.GetComboBoxPr();
	if (oComboPr)
		return oComboPr.GetAutoFit();

	return false;
};
CInlineLevelSdt.prototype.ProcessAutoFitContent = function(isFastRecalc)
{
	var oParagraph = this.GetParagraph();
	var oRun       = this.GetElement(0);
	var oTextPr    = this.Get_CompiledTextPr();
	var oShape     = oParagraph.Parent ? oParagraph.Parent.Is_DrawingShape(true) : null;
	if (!oShape || !oShape.isForm())
		return;

	var oShapeBounds = oShape.getFormRelRect(true);

	g_oTextMeasurer.SetTextPr(oTextPr, oParagraph.GetTheme());
	g_oTextMeasurer.SetFontSlot(fontslot_ASCII);

	var nTextHeight = g_oTextMeasurer.GetHeight();
	var nMaxWidth   = oParagraph.RecalculateMinMaxContentWidth(false).Max;
	var nFontSize   = oTextPr.FontSize;

	if (nMaxWidth < 0.001 || nTextHeight < 0.001 || oShapeBounds.W < 0.001 || oShapeBounds.H < 0.001)
		return;

	var nNewFontSize = nFontSize;

	History.TurnOff();
	if (this.IsMultiLineForm())
	{
		const nFontStep = 0.1;

		oParagraph.Recalculate_Page(0);
		var oContentBounds = oParagraph.GetContentBounds(0);
		if (oContentBounds.Bottom - oContentBounds.Top > oShapeBounds.H)
		{
			nNewFontSize = AscCommon.CorrectFontSize(nFontSize, true);
			while (nNewFontSize > 1)
			{
				oRun.Set_FontSize(nNewFontSize);
				oParagraph.Recalculate_Page(0);

				oContentBounds = oParagraph.GetContentBounds(0);
				if (oContentBounds.Bottom - oContentBounds.Top < oShapeBounds.H)
					break;

				nNewFontSize -= nFontStep;
			}
		}
		else
		{
			var nMaxFontSize = this.Pr.TextPr.FontSize;
			if (!nMaxFontSize)
				nMaxFontSize = 12;

			//nNewFontSize = AscCommon.CorrectFontSize(nFontSize, true);
			while (nNewFontSize <= nMaxFontSize)
			{
				oRun.Set_FontSize(nNewFontSize);
				oParagraph.Recalculate_Page(0);

				var oContentBounds = oParagraph.GetContentBounds(0);
				if (oContentBounds.Bottom - oContentBounds.Top > oShapeBounds.H)
				{
					nNewFontSize -= nFontStep;
					oRun.Set_FontSize(nNewFontSize);
					break;
				}

				nNewFontSize += nFontStep;
			}

			nNewFontSize = Math.min(nNewFontSize, nMaxFontSize);
		}

		oParagraph.Recalculate_Page(0);
		oShape.recalcContent();
		oShape.recalculateText();
	}
	else
	{
		nNewFontSize = Math.min(nFontSize * oShapeBounds.H / nTextHeight * 0.9, 100, nFontSize * oShapeBounds.W / nMaxWidth);

		if (!isFastRecalc)
		{
			oRun.Set_FontSize(nNewFontSize);
			oParagraph.Recalculate_Page(0);
			oShape.recalcContent();
			oShape.recalculateText();
		}
		else if (AscCommon.align_Left !== oParagraph.GetParagraphAlign())
		{
			oRun.Set_FontSize(nNewFontSize);
			oParagraph.private_RecalculateFastRange(0, 0);
		}
	}
	// Восстанавливаем старое значение, чтобы в историю все правильно записалось
	oRun.Set_FontSize(nFontSize);

	nNewFontSize = ((nNewFontSize * 100) | 0) / 100;
	History.TurnOn();

	if (Math.abs(nNewFontSize - nFontSize) > 0.001)
		oRun.Set_FontSize(nNewFontSize);
};
CInlineLevelSdt.prototype.UpdatePictureFormLayout = function()
{
	var oBounds = this.GetFixedFormBounds();
	this.private_UpdatePictureFormLayout(oBounds.W, oBounds.H);
};
CInlineLevelSdt.prototype.private_UpdatePictureFormLayout = function(nW, nH)
{
	var arrDrawings = this.GetAllDrawingObjects();
	if (1 !== arrDrawings.length || nW < 0.001 || nH < 0.001)
		return;

	var oDrawing = arrDrawings[0];
	var oShape   = oDrawing.GraphicObj;

	if (this.IsPlaceHolder() || !oDrawing.IsPicture())
	{
		oShape.spPr.xfrm.setExtX(nW);
		oShape.spPr.xfrm.setExtY(nH);
	}
	else
	{
		var oLogicDocument = this.GetLogicDocument();
		if (!oLogicDocument || !oLogicDocument.GetDrawingObjects() || !oLogicDocument.GetApi())
			return;

		var oDrawingProps = oLogicDocument.GetDrawingObjects().getDrawingPropsFromArray([oDrawing.GraphicObj]);
		if (!oDrawingProps || !oDrawingProps.imageProps)
			return;

		var oProps = new Asc.asc_CImgProperty();
		oProps.ImageUrl = oDrawingProps.imageProps.ImageUrl;

		var oOriginSize = oProps.asc_getOriginSize(oLogicDocument.GetApi());
		if (!oOriginSize.asc_getIsCorrect())
			return;

		var nOriginW = oOriginSize.asc_getImageWidth();
		var nOriginH = oOriginSize.asc_getImageHeight();

		var oPictureFormPr = this.GetPictureFormPr();
		if (!oPictureFormPr || nOriginW < 0.001 || nOriginH < 0.001)
			return;

		var nScaleFlag = oPictureFormPr.GetScaleFlag();
		
		var nDstW, nDstH, isCrop = false;

		// TODO: RespectBorders

		if (Asc.c_oAscPictureFormScaleFlag.Never === nScaleFlag
			|| (Asc.c_oAscPictureFormScaleFlag.Smaller === nScaleFlag && (nOriginH > nH || nOriginW > nW))
			|| (Asc.c_oAscPictureFormScaleFlag.Bigger === nScaleFlag && nH > nOriginH && nW > nOriginW))
		{
			nDstW  = nOriginW;
			nDstH  = nOriginH;
			isCrop = true;
		}
		else if (oPictureFormPr.IsConstantProportions())
		{
			var nCoef = Math.min(nW / nOriginW, nH / nOriginH);
			nDstW     = nOriginW * nCoef;
			nDstH     = nOriginH * nCoef;
			isCrop    = true;
		}
		
		if (isCrop)
		{
			var nSpaceX = nW - nDstW;
			var nSpaceY = nH - nDstH;

			var nPadL = oPictureFormPr.GetShiftX() * nSpaceX;
			var nPadT = oPictureFormPr.GetShiftY() * nSpaceY;

			var oSrcRect = new AscFormat.CSrcRect();
			oSrcRect.setLTRB(
				100 * -nPadL / nDstW,
				100 * -nPadT / nDstH,
				100 * (1 + (nSpaceX - nPadL) / nDstW),
				100 * (1 + (nSpaceY - nPadT) / nDstH)
			);
			oShape.setSrcRect(oSrcRect);
		}
		else
		{
			var oSrcRect = new AscFormat.CSrcRect();
			oSrcRect.setLTRB(0, 0, 100, 100);
			oShape.setSrcRect(oSrcRect)
		}


		oShape.spPr.xfrm.setExtX(nW);
		oShape.spPr.xfrm.setExtY(nH);
	}

	oDrawing.SetSizeRelH({RelativeFrom : AscCommon.c_oAscSizeRelFromH.sizerelfromhPage, Percent : 0});
	oDrawing.SetSizeRelV({RelativeFrom : AscCommon.c_oAscSizeRelFromV.sizerelfromvPage, Percent : 0});
	oDrawing.CheckWH();
};
CInlineLevelSdt.prototype.GetFixedFormWrapperShape = function()
{
	var oParagraph = this.GetParagraph();
	if (!oParagraph || !oParagraph.Parent)
		return null;

	var oShape = oParagraph.Parent.Is_DrawingShape(true);
	if (!oShape || !oShape.isForm())
		return null;

	return oShape;
};
CInlineLevelSdt.prototype.UpdateFixedFormSizeByCombWidth = function()
{
	var oParagraph = this.GetParagraph();
	if (!oParagraph || !oParagraph.Parent)
		return;

	var oShape = oParagraph.Parent.Is_DrawingShape(true);
	if (!oShape || !oShape.parent || !oShape.spPr || !oShape.spPr.xfrm)
		return;

	var oTextFormPr = this.GetTextFormPr();
	if (!oTextFormPr || !oTextFormPr.IsComb() || oTextFormPr.GetMaxCharacters() <= 0 || oTextFormPr.GetWidth() <= 0)
		return;

	var oDrawing = oShape.parent;

	oShape.spPr.xfrm.setExtX(AscCommon.TwipsToMM(oTextFormPr.GetWidth()) * oTextFormPr.GetMaxCharacters());
	oDrawing.CheckWH();
};
CInlineLevelSdt.prototype.UpdateFixedFormCombWidthByFormSize = function(oTextFormPr)
{
	if (!oTextFormPr || !oTextFormPr.IsComb() || oTextFormPr.GetMaxCharacters() <= 0)
		return;

	var oParagraph = this.GetParagraph();
	if (!oParagraph || !oParagraph.Parent)
		return;

	var oShape = oParagraph.Parent.Is_DrawingShape(true);
	if (!oShape || !oShape.parent || !oShape.spPr || !oShape.spPr.xfrm)
		return;


	var nW = oShape.spPr.xfrm.extX;
	if (nW < 0.001)
		return;

	oTextFormPr.SetWidth(AscCommon.MMToTwips(nW / oTextFormPr.GetMaxCharacters()));
};
CInlineLevelSdt.prototype.IsFormExceedsBounds = function()
{
	var oParagraph  = this.GetParagraph();
	var oFormBounds = this.GetFixedFormBounds();
	if (oFormBounds.W <= 0.001 || oFormBounds.H <= 0.001 || !oParagraph || !oParagraph.IsRecalculated())
		return false;

	var oParaBounds = oParagraph.GetContentBounds(0);
	return (oParaBounds.Right - oParaBounds.Left > oFormBounds.W || oParaBounds.Bottom - oParaBounds.Top > oFormBounds.H);
};
CInlineLevelSdt.prototype.CheckSpelling = function(oCollector, nDepth)
{
	let isForm = this.IsForm();
	if (isForm)
		oCollector.FlushWord();

	CParagraphContentWithParagraphLikeContent.prototype.CheckSpelling.apply(this, arguments);

	if (isForm)
		oCollector.FlushWord();
};
CInlineLevelSdt.prototype.MoveCursorOutsideForm = function(isBefore)
{
	let oShape;
	if (this.IsForm()
		&& this.Paragraph
		&& (oShape = this.Paragraph.Parent ? this.Paragraph.Parent.Is_DrawingShape(true) : null)
		&& oShape.isForm())
	{
		let oParaDrawing = oShape.GetParaDrawing();
		if (oParaDrawing)
			oParaDrawing.GoTo_Text(isBefore);
	}
	else
	{
		this.MoveCursorOutsideElement(isBefore);
	}
};
CInlineLevelSdt.prototype.TrimCombForm = function()
{
	let oTextFormPr = this.GetTextFormPr();
	if (!oTextFormPr || !oTextFormPr.IsComb())
		return;

	if (this.IsPlaceHolder())
		this.private_FillPlaceholderContent();

	let nMax = oTextFormPr.GetMaxCharacters();

	let oRun = this.MakeSingleRunElement(false);
	if (oRun.GetElementsCount() > nMax)
		oRun.RemoveFromContent(nMax, oRun.GetElementsCount() - nMax);
};

//--------------------------------------------------------export--------------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CInlineLevelSdt = CInlineLevelSdt;
