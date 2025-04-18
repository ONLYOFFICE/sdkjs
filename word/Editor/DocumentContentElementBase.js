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

var type_Unknown = 0x00;

/**
 * Базовый класс для элементов содержимого документа (Paragraph, CTable, CBlockLevelSdt)
 * @param oParent - ссылка на базовый класс
 * @constructor
 */
function CDocumentContentElementBase(oParent)
{
	this.Id = AscCommon.g_oIdCounter.Get_NewId();

	this.Parent = oParent;
	this.Prev   = null;
	this.Next   = null;
	this.Index  = -1; // перед тем как пользоваться этим параметром нужно у родительского класса вызывать this.Parent.Update_ContentIndexing();

	this.Recalculated = false; // Данный параметр управляет тем, были ли у нас произведены изменения с элементом с момента последнего пересчета

	this.X            = 0;
	this.Y            = 0;
	this.XLimit       = 0;
	this.YLimit       = 0;
	this.PageNum      = 0;
	this.ColumnNum    = 0;
	this.ColumnsCount = 0;
	this.UseXLimit    = true;
	this.UseYLimit    = true;
}

CDocumentContentElementBase.prototype.Get_Type = function()
{
	return this.GetType();
};
CDocumentContentElementBase.prototype.GetType = function()
{
	return type_Unknown;
};
/**
 * Является ли данный элемент параграфом
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsParagraph = function()
{
	return (this.GetType() === type_Paragraph);
};
/**
 * Является ли данный элемент таблицей
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsTable = function()
{
	return (this.GetType() === type_Table);
};
/**
 * Является ли данный элемент блочным контейнером
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsBlockLevelSdt = function()
{
	return (this.GetType() === type_BlockLevelSdt);
};
CDocumentContentElementBase.prototype.Is_Inline = function()
{
	return true;
};
CDocumentContentElementBase.prototype.Set_DocumentIndex = function(nIndex)
{
	this.Index = nIndex;
};
CDocumentContentElementBase.prototype.Set_DocumentNext = function(oElement)
{
	this.Next = oElement;
};
CDocumentContentElementBase.prototype.Set_DocumentPrev = function(oElement)
{
	this.Prev = oElement;
};
CDocumentContentElementBase.prototype.Get_DocumentNext = function()
{
	return this.Next;
};
CDocumentContentElementBase.prototype.Get_DocumentPrev = function()
{
	return this.Prev;
};
CDocumentContentElementBase.prototype.GetNextDocumentElement = function()
{
	var oNext = this.Get_DocumentNext();
	if (!oNext && this.Parent && this.Parent.GetNextDocumentElement)
		return this.Parent.GetNextDocumentElement();

	return oNext;
};
CDocumentContentElementBase.prototype.GetPrevDocumentElement = function()
{
	var oPrev = this.Get_DocumentPrev();
	if (!oPrev && this.Parent && this.Parent.GetPrevDocumentElement)
		return this.Parent.GetPrevDocumentElement();

	return oPrev;
};
CDocumentContentElementBase.prototype.GetNextParagraphInDocument = function()
{
	let next = this.GetNextDocumentElement();
	return next ? next.GetFirstParagraph() : null;
};
CDocumentContentElementBase.prototype.GetPrevParagraphInDocument = function()
{
	let prev = this.GetPrevDocumentElement();
	return prev ? prev.GetLastParagraph() : null;
};
CDocumentContentElementBase.prototype.GetParent = function()
{
	return this.Parent;
};
CDocumentContentElementBase.prototype.SetParent = function(oParent)
{
	this.Parent = oParent;
};
CDocumentContentElementBase.prototype.Set_Parent = function(oParent)
{
	this.Parent = oParent;
};
CDocumentContentElementBase.prototype.Get_Parent = function()
{
	return this.Parent;
};
CDocumentContentElementBase.prototype.GetId = function()
{
	return this.Id;
};
CDocumentContentElementBase.prototype.Get_Id = function()
{
	return this.GetId();
};
CDocumentContentElementBase.prototype.Reset = function(X, Y, XLimit, YLimit, PageAbs, ColumnAbs, ColumnsCount)
{
	this.X            = X;
	this.Y            = Y;
	this.XLimit       = XLimit;
	this.YLimit       = YLimit;
	this.PageNum      = PageAbs;
	this.ColumnNum    = ColumnAbs ? ColumnAbs : 0;
	this.ColumnsCount = ColumnsCount ? ColumnsCount : 1;
};
CDocumentContentElementBase.prototype.SetUseXLimit = function(isUse)
{
	this.UseXLimit = isUse;
};
CDocumentContentElementBase.prototype.IsUseXLimit = function()
{
	return this.UseXLimit;
};
CDocumentContentElementBase.prototype.SetUseYLimit = function(isUse)
{
	this.UseYLimit = isUse;
};
CDocumentContentElementBase.prototype.IsUseYLimit = function()
{
	return this.UseYLimit;
};
CDocumentContentElementBase.prototype.Recalculate_Page = function(CurPage)
{
	return recalcresult_NextElement;
};
CDocumentContentElementBase.prototype.Get_PageBounds = function(CurPage)
{
	return new CDocumentBounds(this.X, this.Y, this.XLimit, this.YLimit);
};
CDocumentContentElementBase.prototype.GetContentBounds = function(CurPage)
{
	return new CDocumentBounds(this.X, this.Y, this.XLimit, this.YLimit);
};
CDocumentContentElementBase.prototype.IsEmptyPage = function(nCurPage)
{
	return false;
};
CDocumentContentElementBase.prototype.Reset_RecalculateCache = function()
{
};
/**
 * @param iPage {number} relative page index
 * @returns {recalcresult}
 */
CDocumentContentElementBase.prototype.RecalculateKeepNext = function(iPage)
{
	if (!(this.Parent instanceof CDocument))
		return recalcresult_NextElement;
	
	// Такая настройка срабатывает в единственном случае:
	// У предыдущего параграфа выставлена данная настройка, а текущий параграф сразу начинается с новой страницы
	// ( при этом у него не выставлен флаг "начать с новой страницы", иначе будет зацикливание здесь ).
	if (1 === iPage && this.IsEmptyPage(0))
	{
		// Если у предыдущего параграфа стоит настройка "не отрывать от следующего".
		// И сам параграф не разбит на несколько страниц и не начинается с новой страницы,
		// тогда мы должны пересчитать предыдущую страницу, с учетом того, что предыдущий параграф
		// надо начать с новой страницы.
		let curr = this.Get_DocumentPrev();
		while (curr && curr.IsParagraph() && !curr.Get_SectionPr())
		{
			let currKeepNext = curr.Get_CompiledPr2(false).ParaPr.KeepNext;
			if (!currKeepNext || curr.getPageCount() > 1 || !curr.IsInline() || curr.Check_PageBreak())
				break;
			
			let prev = curr.Get_DocumentPrev();
			if (!prev || (prev.IsParagraph() && prev.Get_SectionPr()))
				break;
			
			if (!prev.IsParagraph() || !prev.Get_CompiledPr2(false).ParaPr.KeepNext)
			{
				if (this.Parent.RecalcInfo.Can_RecalcObject())
				{
					this.Parent.RecalcInfo.Set_KeepNext(curr, this);
					return recalcresult_PrevPage | recalcresultflags_Column;
				}
				
				break;
			}
			
			curr = prev;
		}
	}
	
	if (this.Parent.RecalcInfo.Check_KeepNextEnd(this))
	{
		// Дошли до сюда, значит уже пересчитали данную ситуацию.
		// Делаем Reset здесь, потому что Reset надо делать в том же месте, гды мы запросили пересчет заново.
		this.Parent.RecalcInfo.Reset();
	}
	
	return recalcresult_NextElement;
};
CDocumentContentElementBase.prototype.Write_ToBinary2 = function(Writer)
{
	Writer.WriteLong(AscDFH.historyitem_type_Unknown);
};
CDocumentContentElementBase.prototype.Read_FromBinary2 = function(Reader)
{
};
CDocumentContentElementBase.prototype.Get_PagesCount = function()
{
	return 0;
};
CDocumentContentElementBase.prototype.Document_CreateFontMap = function(FontMap)
{
};
CDocumentContentElementBase.prototype.Document_CreateFontCharMap = function(FontCharMap)
{
};
CDocumentContentElementBase.prototype.Document_Get_AllFontNames = function(AllFonts)
{
};
CDocumentContentElementBase.prototype.IsInText = function(X, Y, CurPage)
{
	return null;
};
CDocumentContentElementBase.prototype.IsInDrawing = function(X, Y, CurPage)
{
	return false;
};
CDocumentContentElementBase.prototype.IsTableBorder = function(X, Y, CurPage)
{
	return null;
};
CDocumentContentElementBase.prototype.UpdateCursorType = function(X, Y, CurPage)
{
};
CDocumentContentElementBase.prototype.Selection_SetStart = function(X, Y, CurPage, MouseEvent, isTableBorder)
{
};
CDocumentContentElementBase.prototype.Selection_SetEnd = function(X, Y, CurPage, MouseEvent, isTableBorder)
{
};
CDocumentContentElementBase.prototype.IsSelectionEmpty = function(isCheckHidden)
{
	return true;
};
CDocumentContentElementBase.prototype.GetSelectedElementsInfo = function(oInfo)
{
};
CDocumentContentElementBase.prototype.Document_UpdateRulersState = function(CurPage)
{
	this.Content.Document_UpdateRulersState(CurPage);
};
CDocumentContentElementBase.prototype.IsSelectionUse = function()
{
	return false;
};
CDocumentContentElementBase.prototype.IsSelectionToEnd = function()
{
	return false;
};
CDocumentContentElementBase.prototype.RemoveSelection = function()
{
};
CDocumentContentElementBase.prototype.SetSelectionUse = function(isUse)
{
};
CDocumentContentElementBase.prototype.SetSelectionToBeginEnd = function(isSelectionStart, isElementStart)
{
};
CDocumentContentElementBase.prototype.SelectAll = function(nDirection)
{
};
CDocumentContentElementBase.prototype.GetCalculatedTextPr = function()
{
	var oTextPr = new CTextPr();
	oTextPr.InitDefault();
	return oTextPr;
};
CDocumentContentElementBase.prototype.GetCalculatedParaPr = function()
{
	var oParaPr = new CParaPr();
	oParaPr.InitDefault();
	return oParaPr;
};
CDocumentContentElementBase.prototype.GetDirectParaPr = function()
{
	return new CParaPr();
};
CDocumentContentElementBase.prototype.GetDirectTextPr = function()
{
	return new CTextPr();
};
CDocumentContentElementBase.prototype.DrawSelectionOnPage = function(CurPage, clipInfo)
{
};
CDocumentContentElementBase.prototype.StopSelection = function()
{
};
CDocumentContentElementBase.prototype.GetSelectionBounds = function()
{
	return {
		Start     : null,
		End       : null,
		Direction : 0
	};
};
CDocumentContentElementBase.prototype.RecalculateCurPos = function(bUpdateX, bUpdateY, isUpdateTarget)
{
	return null;
};
CDocumentContentElementBase.prototype.Can_CopyCut = function()
{
	return false;
};
CDocumentContentElementBase.prototype.CanCopyCut = function()
{
	return this.Can_CopyCut();
};
CDocumentContentElementBase.prototype.CheckPosInSelection = function(X, Y, CurPage, NearPos)
{
	return false;
};
CDocumentContentElementBase.prototype.Get_NearestPos = function(CurPage, X, Y, bAnchor, Drawing)
{
	return {
		X          : 0,
		Y          : 0,
		Height     : 0,
		PageNum    : 0,
		Internal   : {Line : 0, Page : 0, Range : 0},
		Transform  : null,
		Paragraph  : null,
		ContentPos : null,
		SearchPos  : null
	};
};
CDocumentContentElementBase.prototype.GetNearestPos = function(CurPage, X, Y, bAnchor, Drawing)
{
	return this.Get_NearestPos(CurPage, X, Y, bAnchor, Drawing);
};
CDocumentContentElementBase.prototype.CanUpdateTarget = function(CurPage)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorLeftWithSelectionFromEnd = function(Word)
{
};
CDocumentContentElementBase.prototype.MoveCursorLeft = function(AddToSelect, Word)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorRight = function(AddToSelect, Word)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorRightWithSelectionFromStart = function(Word)
{

};
CDocumentContentElementBase.prototype.MoveCursorToStartPos = function(AddToSelect)
{
};
CDocumentContentElementBase.prototype.MoveCursorToEndPos = function(AddToSelect, StartSelectFromEnd)
{
};
CDocumentContentElementBase.prototype.MoveCursorUp = function(AddToSelect)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorUpToLastRow = function(X, Y, AddToSelect)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorDown = function(AddToSelect)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorDownToFirstRow = function(X, Y, AddToSelect)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorToEndOfLine = function(AddToSelect)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorToStartOfLine = function(AddToSelect)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorToXY = function(X, Y, bLine, bDontChangeRealPos, CurPage)
{
	return false;
};
CDocumentContentElementBase.prototype.MoveCursorToCell = function(bNext)
{
	return false;
};
CDocumentContentElementBase.prototype.IsCursorAtBegin = function()
{
	return true;
};
CDocumentContentElementBase.prototype.IsCursorAtEnd = function()
{
	return true;
};
CDocumentContentElementBase.prototype.GetSelectionState = function()
{
	return [];
};
CDocumentContentElementBase.prototype.SetSelectionState = function(State, StateIndex)
{

};
CDocumentContentElementBase.prototype.AddNewParagraph = function()
{

};
CDocumentContentElementBase.prototype.Get_SelectionState2 = function()
{
	return null;
};
CDocumentContentElementBase.prototype.Set_SelectionState2 = function(State)
{
};
CDocumentContentElementBase.prototype.IsStartFromNewPage = function()
{
	return false;
};
CDocumentContentElementBase.prototype.GetAllParagraphs = function(Props, ParaArray)
{
	return [];
};
CDocumentContentElementBase.prototype.GetAllTables = function(oProps, arrTables)
{
	return [];
};
CDocumentContentElementBase.prototype.SetContentSelection = function(StartDocPos, EndDocPos, Depth, StartFlag, EndFlag)
{
};
CDocumentContentElementBase.prototype.GetContentPosition = function(bSelection, bStart, PosArray)
{
};
CDocumentContentElementBase.prototype.SetContentPosition = function(DocPos, Depth, Flag)
{
};
CDocumentContentElementBase.prototype.GetNumberingInfo = function(oNumberingEngine)
{
};
CDocumentContentElementBase.prototype.AddInlineImage = function(W, H, Img, GraphicObject, bFlow)
{
};
CDocumentContentElementBase.prototype.AddImages = function(aImages)
{
};
CDocumentContentElementBase.prototype.AddOleObject = function(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory)
{
	return null;
};
CDocumentContentElementBase.prototype.AddSignatureLine = function(oSignatureDrawing)
{
};
CDocumentContentElementBase.prototype.AddTextArt = function(nStyle)
{
};
CDocumentContentElementBase.prototype.AddInlineTable = function(nCols, nRows, nMode)
{
	return null;
};
CDocumentContentElementBase.prototype.Remove = function(nCount, bOnlyText, bRemoveOnlySelection, bOnAddText, isWord)
{
};
CDocumentContentElementBase.prototype.SetReviewType = function(ReviewType)
{

};
CDocumentContentElementBase.prototype.GetReviewType = function()
{
	return reviewtype_Common;
};
CDocumentContentElementBase.prototype.Is_Empty = function()
{
	return true;
};
CDocumentContentElementBase.prototype.Add = function(oParaItem)
{
	// TODO: Данную функцию нужно заменить на AddToParagraph
};
CDocumentContentElementBase.prototype.PreDelete = function()
{
};
CDocumentContentElementBase.prototype.ClearParagraphFormatting = function(isClearParaPr, isClearTextPr)
{
};
CDocumentContentElementBase.prototype.GetCursorPosXY = function()
{
	return {X : 0, Y : 0};
};
CDocumentContentElementBase.prototype.StartSelectionFromCurPos = function()
{
};
CDocumentContentElementBase.prototype.SetParagraphPr = function(oParaPr)
{
};
CDocumentContentElementBase.prototype.SetParagraphAlign = function(Align)
{
};
CDocumentContentElementBase.prototype.SetParagraphDefaultTabSize = function(TabSize)
{
};
CDocumentContentElementBase.prototype.SetParagraphSpacing = function(Spacing)
{
};
CDocumentContentElementBase.prototype.SetParagraphTabs = function(Tabs)
{
};
CDocumentContentElementBase.prototype.SetParagraphIndent = function(Ind)
{
};
CDocumentContentElementBase.prototype.SetParagraphShd = function(Shd)
{
};
CDocumentContentElementBase.prototype.SetParagraphStyle = function(Name)
{
};
CDocumentContentElementBase.prototype.SetParagraphContextualSpacing = function(Value)
{
};
CDocumentContentElementBase.prototype.SetParagraphPageBreakBefore = function(Value)
{
};
CDocumentContentElementBase.prototype.SetParagraphKeepLines = function(Value)
{
};
CDocumentContentElementBase.prototype.SetParagraphKeepNext = function(Value)
{
};
CDocumentContentElementBase.prototype.SetParagraphWidowControl = function(Value)
{
};
CDocumentContentElementBase.prototype.SetParagraphBorders = function(Borders)
{
};
CDocumentContentElementBase.prototype.SetParagraphFramePr = function(FramePr, bDelete)
{
};
CDocumentContentElementBase.prototype.IncreaseDecreaseFontSize = function(bIncrease)
{
};
CDocumentContentElementBase.prototype.IncreaseDecreaseIndent = function(bIncrease)
{
};
CDocumentContentElementBase.prototype.SetImageProps = function(oProps)
{
};
CDocumentContentElementBase.prototype.SetTableProps = function(oProps)
{
};
CDocumentContentElementBase.prototype.GetSelectedContent = function(oSelectedContent)
{
};
CDocumentContentElementBase.prototype.PasteFormatting = function(oData)
{
};
CDocumentContentElementBase.prototype.GetCurPosXY = function()
{
	return {X : 0, Y : 0};
};
CDocumentContentElementBase.prototype.GetSelectedText = function(bClearText, oPr)
{
	return null;
};
CDocumentContentElementBase.prototype.GetCurrentParagraph = function(bIgnoreSelection, arrSelectedParagraphs)
{
	return null;
};
CDocumentContentElementBase.prototype.GetCurrentTablesStack = function(arrTables)
{
	return arrTables ? arrTables : [];
};
CDocumentContentElementBase.prototype.AddTableRow = function(bBefore, nCount)
{
	return false;
};
CDocumentContentElementBase.prototype.AddTableColumn = function(bBefore, nCount)
{
	return false;
};
CDocumentContentElementBase.prototype.RemoveTableRow = function(nRowIndex)
{
	return false;
};
CDocumentContentElementBase.prototype.RemoveTableColumn = function()
{
	return false;
};
CDocumentContentElementBase.prototype.MergeTableCells = function()
{
	return false;
};
CDocumentContentElementBase.prototype.SplitTableCells = function(nColsCount, nRowsCount)
{
	return false;
};
CDocumentContentElementBase.prototype.RemoveTableCells = function()
{
	return false;
};
CDocumentContentElementBase.prototype.RemoveTable = function()
{
	return false;
};
CDocumentContentElementBase.prototype.SelectTable = function(Type)
{
};
CDocumentContentElementBase.prototype.DistributeTableCells = function(isHorizontally)
{
	return false;
};
CDocumentContentElementBase.prototype.CanMergeTableCells = function()
{
	return false;
};
CDocumentContentElementBase.prototype.CanSplitTableCells = function()
{
	return false;
};
CDocumentContentElementBase.prototype.Document_UpdateInterfaceState = function()
{
};
CDocumentContentElementBase.prototype.Document_UpdateRulersState = function()
{
};
CDocumentContentElementBase.prototype.GetTableProps = function()
{
	return null;
};
CDocumentContentElementBase.prototype.AddHyperlink = function(Props)
{
	return null;
};
CDocumentContentElementBase.prototype.ModifyHyperlink = function(Props)
{
};
CDocumentContentElementBase.prototype.RemoveHyperlink = function()
{
};
CDocumentContentElementBase.prototype.CanAddHyperlink = function(bCheckInHyperlink)
{
	return false;
};
CDocumentContentElementBase.prototype.IsCursorInHyperlink = function(bCheckEnd)
{
	return null;
};
CDocumentContentElementBase.prototype.AddComment = function(Comment, bStart, bEnd)
{
};
CDocumentContentElementBase.prototype.CanAddComment = function()
{
	return false;
};
CDocumentContentElementBase.prototype.GetSelectionAnchorPos = function()
{
	return null;
};
CDocumentContentElementBase.prototype.AddContentControl = function(nContentControlType)
{
	return null;
};
CDocumentContentElementBase.prototype.RecalculateMinMaxContentWidth = function(isRotated)
{
	return {Min : 0, Max : 0};
};
CDocumentContentElementBase.prototype.Shift = function(CurPage, dX, dY)
{
};
CDocumentContentElementBase.prototype.UpdateEndInfo = function()
{
};
CDocumentContentElementBase.prototype.PrepareRecalculateObject = function()
{
};
CDocumentContentElementBase.prototype.SaveRecalculateObject = function()
{
	return null;
};
CDocumentContentElementBase.prototype.LoadRecalculateObject = function(RecalcObj)
{
};
CDocumentContentElementBase.prototype.SetApplyToAll = function(isApplyAll)
{
	this.ApplyToAll = isApplyAll;
};
CDocumentContentElementBase.prototype.IsApplyToAll = function()
{
	return this.ApplyToAll;
};
CDocumentContentElementBase.prototype.RecalculateAllTables = function()
{
};
CDocumentContentElementBase.prototype.GetAllFloatElements = function(FloatObjects)
{
	if (!FloatObjects)
		return [];

	return FloatObjects;
};
CDocumentContentElementBase.prototype.Get_FirstParagraph = function()
{
	return null;
};
CDocumentContentElementBase.prototype.StartFromNewPage = function()
{
};
CDocumentContentElementBase.prototype.CollectDocumentStatistics = function(Stats)
{
};
CDocumentContentElementBase.prototype.CompareDrawingsLogicPositions = function(CompareObject)
{
	return 0;
};
CDocumentContentElementBase.prototype.GetStyleFromFormatting = function()
{
	return null;
};
CDocumentContentElementBase.prototype.GetAllContentControls = function(arrContentControls)
{
};
/**
 * Проверяем выделен ли элемент целиком
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsSelectedAll = function()
{
	return false;
};
CDocumentContentElementBase.prototype.GetLastRangeVisibleBounds = function()
{
	return {
		X        : 0,
		Y        : 0,
		W        : 0,
		H        : 0,
		BaseLine : 0,
		XLimit   : 0
	};
};
CDocumentContentElementBase.prototype.FindNextFillingForm = function(isNext, isCurrent, isStart)
{
	return null;
};
CDocumentContentElementBase.prototype.GetRevisionsChangeElement = function(SearchEngine)
{
	return null;
};
CDocumentContentElementBase.prototype.AcceptRevisionChanges = function(Type, bAll)
{
};
CDocumentContentElementBase.prototype.RejectRevisionChanges = function(Type, bAll)
{
};
CDocumentContentElementBase.prototype.GetDocumentPositionFromObject = function(arrPos)
{
	if (!arrPos)
		arrPos = [];

	var oParent = this.GetParent();
	if (oParent)
	{
		if (arrPos.length > 0)
		{
			arrPos.splice(0, 0, {Class : oParent, Position : this.GetIndex()});
			arrPos = oParent.GetDocumentPositionFromObject(arrPos);
		}
		else
		{
			arrPos = oParent.GetDocumentPositionFromObject(arrPos);
			arrPos.push({Class : oParent, Position : this.GetIndex()});
		}
	}

	return arrPos;
};
/**
 * Получаем массив всех конент контролов, внутри которых лежит данный класс
 * @returns {Array}
 */
CDocumentContentElementBase.prototype.GetParentContentControls = function()
{
	let docPos = this.GetDocumentPositionFromObject();
	
	let contentControls = [];
	for (let pos = 0, len = docPos.length; pos < len; ++pos)
	{
		if (docPos[pos].Class instanceof AscWord.CInlineLevelSdt)
			contentControls.push(docPos[pos].Class);
		else if (docPos[pos].Class instanceof AscWord.CDocumentContent && docPos[pos].Class.Parent instanceof AscWord.CBlockLevelSdt)
			contentControls.push(docPos[pos].Class.Parent);
	}
	
	return contentControls;
};
/**
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsInPlaceholder = function()
{
	let contentControls = this.GetParentContentControls();
	for (let index = 0, count = contentControls.length; index < count; ++index)
	{
		if (contentControls[index].IsPlaceHolder())
			return true;
	}
	
	return false;
};
CDocumentContentElementBase.prototype.Get_Index = function()
{
	return this.GetIndex();
};
CDocumentContentElementBase.prototype.GetOutlineParagraphs = function(arrOutline, oPr)
{
};
//----------------------------------------------------------------------------------------------------------------------
// Функции для работы с номерами страниц
//----------------------------------------------------------------------------------------------------------------------
CDocumentContentElementBase.prototype.Get_StartPage_Absolute = function()
{
	return this.Get_AbsolutePage(0);
};
CDocumentContentElementBase.prototype.Get_StartPage_Relative = function()
{
	return this.PageNum;
};
CDocumentContentElementBase.prototype.Get_StartColumn = function()
{
	return this.ColumnNum;
};
CDocumentContentElementBase.prototype.Get_ColumnsCount = function()
{
	return this.ColumnsCount;
};
CDocumentContentElementBase.prototype.GetStartColumn = function()
{
	return this.ColumnNum;
};
CDocumentContentElementBase.prototype.GetColumnsCount = function()
{
	return this.ColumnsCount;
};
CDocumentContentElementBase.prototype.private_GetRelativePageIndex = function(CurPage)
{
	if (!this.ColumnsCount || 0 === this.ColumnsCount)
		return this.PageNum + CurPage;

	return this.PageNum + ((this.ColumnNum + CurPage) / this.ColumnsCount | 0);
};
CDocumentContentElementBase.prototype.private_GetAbsolutePageIndex = function(CurPage)
{
	return this.Parent.Get_AbsolutePage(this.private_GetRelativePageIndex(CurPage));
};
CDocumentContentElementBase.prototype.Get_AbsolutePage = function(CurPage)
{
	return this.private_GetAbsolutePageIndex(CurPage);
};
CDocumentContentElementBase.prototype.Get_AbsoluteColumn = function(CurPage)
{
	if (this.Parent instanceof CDocument)
		return this.private_GetColumnIndex(CurPage);

	return this.Parent.Get_AbsoluteColumn(this.private_GetRelativePageIndex(CurPage));
};
CDocumentContentElementBase.prototype.private_GetColumnIndex = function(CurPage)
{
	return (this.ColumnNum + CurPage) - (((this.ColumnNum + CurPage) / this.ColumnsCount | 0) * this.ColumnsCount);
};
CDocumentContentElementBase.prototype.Get_CurrentPage_Absolute = function()
{
	return this.private_GetAbsolutePageIndex(0);
};
CDocumentContentElementBase.prototype.Get_CurrentPage_Relative = function()
{
	return this.private_GetRelativePageIndex(0);
};
CDocumentContentElementBase.prototype.GetCurrentPageAbsolute = function()
{
	return this.Get_CurrentPage_Absolute();
};
CDocumentContentElementBase.prototype.GetAbsolutePage = function(CurPage)
{
	return this.private_GetAbsolutePageIndex(CurPage);
};
CDocumentContentElementBase.prototype.GetAbsoluteColumn = function(CurPage)
{
	return this.Get_AbsoluteColumn(CurPage);
};
/**
 * Получаем начальный номер страницы данного элемента относительно родительского класса
 * @returns {number}
 */
CDocumentContentElementBase.prototype.GetStartPageRelative = function()
{
	return this.PageNum;
};
/**
 * Получаем номер страницы, относительно родительского класса
 * @param {number} nCurPage
 * @returns {number}
 */
CDocumentContentElementBase.prototype.GetRelativePage = function(nCurPage)
{
	return this.private_GetRelativePageIndex(nCurPage);
};
/**
 * Получаем обсолютный начальный номер страницы данного элемента
 * @returns {number}
 */
CDocumentContentElementBase.prototype.GetStartPageAbsolute = function()
{
	return this.private_GetAbsolutePageIndex(0);
};
//----------------------------------------------------------------------------------------------------------------------
CDocumentContentElementBase.prototype.GetPagesCount = function()
{
	return this.Get_PagesCount();
};
CDocumentContentElementBase.prototype.getPageCount = function()
{
	return this.Get_PagesCount();
};
CDocumentContentElementBase.prototype.GetIndex = function()
{
	if (!this.Parent)
		return -1;

	this.Parent.Update_ContentIndexing();

	if (this !== this.Parent.GetElement(this.Index))
		this.Index = -1;

	return this.Index;
};
CDocumentContentElementBase.prototype.getPageBounds = function(iPage)
{
	return this.Get_PageBounds(iPage);
};
CDocumentContentElementBase.prototype.GetPageBounds = function(CurPage)
{
	return this.Get_PageBounds(CurPage);
};
CDocumentContentElementBase.prototype.GetNearestPos = function(CurPage, X, Y, bAnchor, Drawing)
{
	return this.Get_NearestPos(CurPage, X, Y, bAnchor, Drawing);
};
CDocumentContentElementBase.prototype.CreateFontMap = function(oFontMap)
{
	return this.Document_CreateFontMap(oFontMap);
};
CDocumentContentElementBase.prototype.CreateFontCharMap = function(oFontCharMap)
{
	return this.Document_CreateFontCharMap(oFontCharMap);
};
CDocumentContentElementBase.prototype.GetAllFontNames = function(FontNames)
{
	return this.Document_Get_AllFontNames(FontNames);
};
CDocumentContentElementBase.prototype.GetSelectionState2 = function()
{
	return this.Get_SelectionState2();
};
CDocumentContentElementBase.prototype.SetSelectionState2 = function(State)
{
	return this.Set_SelectionState2(State);
};
CDocumentContentElementBase.prototype.GetReviewInfo = function()
{
	return new AscWord.ReviewInfo();
};
CDocumentContentElementBase.prototype.SetReviewTypeWithInfo = function(nType, oInfo)
{
};
CDocumentContentElementBase.prototype.IsEmpty = function(oProps)
{
	return this.Is_Empty(oProps);
};
CDocumentContentElementBase.prototype.AddToParagraph = function(oItem)
{
	return this.Add(oItem);
};
CDocumentContentElementBase.prototype.GetAllDrawingObjects = function(AllDrawingObjects)
{
};
CDocumentContentElementBase.prototype.GetAllComments = function(AllComments)
{
};
CDocumentContentElementBase.prototype.GetAllMaths = function(AllMaths)
{
};

/**
 * Find all SEQ complex fields with specified type
 * @param {String} sType - field type
 * @param {Array} aFields - array which accumulates complex fields
 */
CDocumentContentElementBase.prototype.GetAllSeqFieldsByType = function(sType, aFields)
{
};

CDocumentContentElementBase.prototype.UpdateBookmarks = function(oManager)
{
};
/**
 * Получаем текущий TableOfContents, это может быть просто поле или поле вместе с оберткой Sdt
 * @param isUnique ищем с параметром Unique = true
 * @param isCheckFields Проверять ли TableOfContents, заданные через сложные поля
 * @returns {CComplexField | CBlockLevelSdt | null}
 */
CDocumentContentElementBase.prototype.GetTableOfContents = function(isUnique, isCheckFields)
{
	return null;
};
/**
 * Get all tables of figures inside
 * @param arrComplexFields
 */
CDocumentContentElementBase.prototype.GetTablesOfFigures = function(arrComplexFields)
{
};
/**
 * Проверяем у родительского класса выделен ли только один элемент
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsSelectedSingleElement = function()
{
	if (this.Parent)
		return this.Parent.IsSelectedSingleElement();

	return false;
};
/**
 * Получаем последний параграф в данном контенте, если последний элемент не параграф, то запрашиваем у него
 * @returns {?Paragraph}
 */
CDocumentContentElementBase.prototype.GetLastParagraph = function()
{
	return null;
};
/**
 * Получаем первый параграф в данном контенте
 * @returns {?Paragraph}
 */
CDocumentContentElementBase.prototype.GetFirstParagraph = function()
{
	return this.Get_FirstParagraph();
};
/**
 * Получаем параграф, следующий за данным элементом
 * @returns {?Paragraph}
 */
CDocumentContentElementBase.prototype.GetNextParagraph = function()
{
	var oNextElement = this.Get_DocumentNext();

	if (oNextElement)
	{
		if (type_Paragraph === oNextElement.GetType())
			return oNextElement;
		else
			return oNextElement.GetFirstParagraph();
	}

	if (this.Parent && this.Parent.GetNextParagraph)
		return this.Parent.GetNextParagraph();

	return null;
};
/**
 * Получаем параграф, идущий перед данным элементом
 * @returns {?Paragraph}
 */
CDocumentContentElementBase.prototype.GetPrevParagraph = function()
{
	var oPrevElement = this.Get_DocumentPrev();

	if (oPrevElement)
	{
		if (type_Paragraph === oPrevElement.GetType())
			return oPrevElement;
		else
			return oPrevElement.GetLastParagraph();
	}

	if (this.Parent && this.Parent.GetPrevParagraph)
		return this.Parent.GetPrevParagraph();

	return null;
};
/**
 * Собираем список заголовков
 * @param {Paragraph[]} arrOutline
 * @param {object} oPr
 */
CDocumentContentElementBase.prototype.GetOutlineParagraphs = function(arrOutline, oPr)
{
};
/**
 * Вплоть до заданного параграфа ищем последнюю похожую нумерацию
 * @param oContinueEngine {CDocumentNumberingContinueEngine}
 * @returns {AscWord.NumPr | null}
 */
CDocumentContentElementBase.prototype.GetSimilarNumbering = function(oContinueEngine)
{
	return null;
};
/**
 * Переходим к следующей ссылке на сноску
 * @param isNext {boolean} - направление поиска
 * @param isCurrent {boolean} - ищем начиная с текущей позиции или с края элемента
 * @param isStepFootnote {boolean} - ищем сноски на странице
 * @param isStepEndnote {boolean} - ищем концевые сноски
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.GotoFootnoteRef = function(isNext, isCurrent, isStepFootnote, isStepEndnote)
{
	return false;
};
/**
 * Сообщаем, пересчитан ли данный элемент с момента внесения в него изменений
 * @param {boolean} isRecalculated
 */
CDocumentContentElementBase.prototype.SetIsRecalculated = function(isRecalculated)
{
	this.Recalculated = isRecalculated;
};
/**
 * Узнаем рассчитан ли данный элемент
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsRecalculated = function()
{
	return this.Recalculated;
};
/**
 * Проверяем выделен ли сейчас какой-либо плейсхолдер, если да, то возвращаем управляющий объект
 * @returns {?Object}
 */
CDocumentContentElementBase.prototype.GetPlaceHolderObject = function()
{
	return null;
};
/**
 * Получаем массив все полей в документе (простых и сложных)
 * @param isUseSelection {boolean} ищем по селекут или вообще все
 * @param arrFields - массив, который мы заполняем, если не задан, то создается новый и возвращается
 * @returns {Array}
 */
CDocumentContentElementBase.prototype.GetAllFields = function(isUseSelection, arrFields)
{
	return arrFields ? arrFields : [];
};
/**
 * Получаем верхний элемент в документе, в котором лежит данный элемент
 * @returns {?CDocumentContentElementBase}
 */
CDocumentContentElementBase.prototype.GetTopElement = function()
{
	if (!this.Parent)
		return null;

	if (this.Parent === this.Parent.Is_TopDocument(true))
		return this;

	return this.Parent.GetTopElement();
};
/**
 * Получаем верхний DocContent, в котором лежит данный элемент
 * @returns {?CDocumentContentBase}
 */
CDocumentContentElementBase.prototype.GetTopDocumentContent = function()
{
	if (!this.Parent)
		return null;
	
	return this.Parent.GetTopDocumentContent();
};
/**
 * Получаем объект лока данного элемента
 * @returns {AscCommon.CLock}
 */
CDocumentContentElementBase.prototype.GetLock = function()
{
	return this.Lock;
};
/**
 * Если мы находимся в колонтитуле возвращаем его
 * @returns {?CHdrFtr}
 */
CDocumentContentElementBase.prototype.GetHdrFtr = function()
{
	if (this.Parent)
		return this.Parent.IsHdrFtr(true);

	return null;
};
/**
 * Используется ли данный элемент в содержимом документа
 * @param {string} sId - идентификатор внутреннего класса
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsUseInDocument = function(sId)
{
	return false;
};
/**
 * Пробегаемся по все ранам с заданной функцией
 * @param fCheck - функция проверки содержимого рана
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.CheckRunContent = function(fCheck)
{
	return false;
};
CDocumentContentElementBase.prototype.CheckSelectedRunContent = function(fCheck)
{
	return false;
};
/**
 * По заданной странице получаем страницу, с которой нужно начинать расчет
 * @param {number} nPageAbs
 * @return {number}
 */
CDocumentContentElementBase.prototype.GetStartPageForRecalculate = function(nPageAbs)
{
	return nPageAbs;
};
/**
 * Проверяем выделено ли сейчас какое-либо презентационное поле, если да, то возвращаем управляющий объект
 * @returns {?Object}
 */
CDocumentContentElementBase.prototype.GetPresentationField = function()
{
	return null;
};
/**
 * Получаем список таблицы на заданной абсолютной странице
 * @param {number} nPageAbs
 * @param {Array} arrTables
 * @returns {Array}
 */
CDocumentContentElementBase.prototype.GetAllTablesOnPage = function(nPageAbs, arrTables){return arrTables ? arrTables : [];};
/**
 * Обрабатываем сложные поля
 */
CDocumentContentElementBase.prototype.ProcessComplexFields = function() {};
/**
 * Вычисляем EndInfo для всех параграфаов
 */
CDocumentContentElementBase.prototype.RecalculateEndInfo = function() {};
/**
 * Получаем ссылку на глобальный класс документа
 * @returns {CDocument}
 */
CDocumentContentElementBase.prototype.GetLogicDocument = function()
{
	if (!this.LogicDocument && this.Parent && this.Parent.GetLogicDocument)
		this.LogicDocument = this.Parent.GetLogicDocument();
	
	return this.LogicDocument;
};
/**
 * @returns {?CDrawingDocument}
 */
CDocumentContentElementBase.prototype.getDrawingDocument = function()
{
	return Asc.editor.getDrawingDocument();
};
/**
 * @returns {?AscWord.CDocumentSpellChecker}
 */
CDocumentContentElementBase.prototype.getSpelling = function()
{
	let oLogicDocument = this.GetLogicDocument();
	if(oLogicDocument)
	{
		return oLogicDocument.Spelling;
	}
	return null;
};
/**
 * @returns {boolean}
 */
CDocumentContentElementBase.prototype.IsSpellingUse = function()
{
	let oSpelling = this.getSpelling();
	if(!oSpelling)
		return false;
	return oSpelling.Use;
};
/**
 * Получаем настройки рамки для данного элемента
 * @returns {?CFramePr}
 */
CDocumentContentElementBase.prototype.GetFramePr = function(){return null;};
/**
 * Получаем маскимальную ширину таблицы
 * @returns {{GapLeft : {number}, GapRight : {number}, GridWidth : {number}}}
 */
CDocumentContentElementBase.prototype.GetMaxTableGridWidth = function(){return {GapLeft : 0, GapRight : 0, GridWidth : -1};};
/**
 * Обновляем нумерацию строк
 */
CDocumentContentElementBase.prototype.UpdateLineNumbersInfo = function(){};
/**
 * Подсчитываем на сколько элементов разбивается данный элемент с заданным сепаратором
 * @param oEngine {CTextToTableEngine}
 */
CDocumentContentElementBase.prototype.CalculateTextToTable = function(oEngine){};
/**
 * @param arrChanges
 */
CDocumentContentElementBase.prototype.GetSelectedReviewChanges = function(arrChanges, oTrackChanges) {return arrChanges ? arrChanges : [];};
/**
 * Прокидываем наверх событие об изменении содержимого данного элемента
 */
CDocumentContentElementBase.prototype.OnContentChange = function()
{
	if (this.Parent && this.Parent.OnContentChange)
		this.Parent.OnContentChange();
};
/**
 * Get the scale coefficient for the current element depending on the current section and the document layout
 * @returns {number}
 */
CDocumentContentElementBase.prototype.getLayoutScaleCoefficient = function()
{
	let logicDocument = this.GetLogicDocument();
	if (!logicDocument || !logicDocument.IsDocumentEditor() || !this.Get_SectPr)
		return 1;
	
	let layout = logicDocument.Layout;
	logicDocument.Layout = logicDocument.Layouts.Print;
	
	let sectPr = this.Get_SectPr();
	logicDocument.Layout = layout;
	
	if (!sectPr)
		return 1;
	
	return logicDocument.GetDocumentLayout().GetScaleBySection(sectPr);
};
CDocumentContentElementBase.prototype.updateTrackRevisions = function()
{
	AscWord.checkElementInRevision && AscWord.checkElementInRevision(this);
};
CDocumentContentElementBase.prototype.isPreventedPreDelete = function()
{
	let logicDocument = this.GetLogicDocument();
	return !logicDocument || !logicDocument.IsDocumentEditor() || logicDocument.isPreventedPreDelete();
};
CDocumentContentElementBase.prototype.isWholeElementInPermRange = function()
{
	// TODO: В таблицах GetNextDocumentElement/GetPrevDocumentElement не работает, надо проверить не баг ли это
	//       по логике оба варианта должны выдавать одинаковый результат
	
	// let prevPara = this.GetPrevParagraphInDocument();
	// let nextPara = this.GetNextParagraphInDocument();
	//
	// let startRanges = prevPara ? prevPara.GetEndInfo().GetPermRanges() : [];
	// let endRanges   = nextPara ? nextPara.GetEndInfoByPage(-1).GetPermRanges() : [];
	
	let startPara = this.GetFirstParagraph();
	let endPara   = this.GetLastParagraph();
	
	if (!startPara
		|| !endPara
		|| !startPara.IsRecalculated()
		|| !endPara.IsRecalculated())
		return false;
	
	let startInfo = startPara.GetEndInfoByPage(-1);
	let endInfo   = endPara.GetEndInfo();
	
	let startRanges = startInfo ? startInfo.GetPermRanges() : [];
	let endRanges   = endInfo ? endInfo.GetPermRanges() : [];
	
	return AscWord.PermRangesManager.isInPermRange(startRanges, endRanges);
};

//--------------------------------------------------------export--------------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CDocumentContentElementBase = CDocumentContentElementBase;
window['AscCommonWord'].type_Unknown = type_Unknown;
