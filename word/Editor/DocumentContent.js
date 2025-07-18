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
var g_oTableId = AscCommon.g_oTableId;

var c_oAscHAnchor = Asc.c_oAscHAnchor;
var c_oAscXAlign = Asc.c_oAscXAlign;
var c_oAscYAlign = Asc.c_oAscYAlign;
var c_oAscVAnchor = Asc.c_oAscVAnchor;


/**
 * Класс CDocumentContent. Данный класс используется для работы с контентом ячеек таблицы, колонтитулов, сносок,
 * надписей.
 * @param Parent
 * @param DrawingDocument
 * @param X
 * @param Y
 * @param XLimit
 * @param YLimit
 * @param Split
 * @param TurnOffInnerWrap
 * @param bPresentation
 * @constructor
 * @extends {CDocumentContentBase}
 */
function CDocumentContent(Parent, DrawingDocument, X, Y, XLimit, YLimit, Split, TurnOffInnerWrap, bPresentation)
{
	CDocumentContentBase.call(this);

    this.Id = AscCommon.g_oIdCounter.Get_NewId();

	this.X      = X;
	this.Y      = Y;
	this.XLimit = XLimit;
	this.YLimit = YLimit;

    this.UseXLimit = true;
    this.UseYLimit = true;

	this.StartPage    = 0;
	this.StartColumn  = 0;
	this.ColumnsCount = 1;

    this.Parent = Parent;

    this.DrawingDocument = null;
    this.LogicDocument   = null;
    this.DrawingObjects  = null;

    if ( undefined !== DrawingDocument && null !== DrawingDocument )
    {
        this.DrawingDocument = DrawingDocument;

        if ( undefined !== editor && true === editor.isDocumentEditor && !(bPresentation === true) && DrawingDocument.m_oLogicDocument )
        {
            this.LogicDocument   = DrawingDocument.m_oLogicDocument;
            this.DrawingObjects  = DrawingDocument.m_oLogicDocument.DrawingObjects; // Массив укзателей на все инлайновые графические объекты
        }
    }

    this.TurnOffInnerWrap = undefined === TurnOffInnerWrap ? false : TurnOffInnerWrap;

    this.Pages = [];

    this.RecalcInfo = new CDocumentRecalcInfo();

    this.Split = Split; // Разделяем ли на страницы
    this.bPresentation = bPresentation; // Разделяем ли на страницы

    this.Content[0] = new AscWord.Paragraph(this, bPresentation);
    this.Content[0].Correct_Content();
    this.Content[0].Set_DocumentNext( null );
    this.Content[0].Set_DocumentPrev( null );

    this.CurPos  =
    {
        X          : 0,
        Y          : 0,
        ContentPos : 0, // в зависимости, от параметра Type: позиция в Document.Content
        RealX      : 0, // позиция курсора, без учета расположения букв
        RealY      : 0, // это актуально для клавиш вверх и вниз
        Type       : docpostype_Content,
        TableMove  : 0  // специльный параметр для переноса таблиц
    };

    this.Selection =
    {
        Start    : false,
        Use      : false,
        StartPos : 0,
        EndPos   : 0,
        Flag     : selectionflag_Common,
        Data     : null
    };

    this.ClipInfo = [];

    this.ShiftViewX = 0;
    this.ShiftViewY = 0;

    this.ApplyToAll = false; // Специальный параметр, используемый в ячейках таблицы.
                             // True, если ячейка попадает в выделение по ячейкам.

    this.TurnOffRecalc = false;

    this.m_oContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)
    this.StartState = null;
	this.Recalculated = false; // Flag only for the current level

    this.ReindexStartPos = 0;
	
	this.transform = null;

    // Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
    AscCommon.g_oTableId.Add( this, this.Id );

    if(this.bPresentation)
    {
        this.Save_StartState();
    }
}
CDocumentContent.prototype = Object.create(CDocumentContentBase.prototype);
CDocumentContent.prototype.constructor = CDocumentContent;

CDocumentContent.prototype.Save_StartState = function()
{
    this.StartState = new CDocumentContentStartState(this);
};
CDocumentContent.prototype.Copy = function(Parent, DrawingDocument, oPr)
{
	var DC = new CDocumentContent(Parent, DrawingDocument ? DrawingDocument : this.DrawingDocument, 0, 0, 0, 0, this.Split, this.TurnOffInnerWrap, this.bPresentation);

	// Копируем содержимое
	DC.Internal_Content_RemoveAll();

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		DC.Internal_Content_Add(Index, this.Content[Index].Copy(DC, DrawingDocument, oPr), false);
	}

	return DC;
};
CDocumentContent.prototype.Copy2 = function(OtherDC, oPr)
{
	if(this === OtherDC)
	{
		return;
	}
	// Копируем содержимое
	this.Internal_Content_RemoveAll();

	var Count = OtherDC.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		this.Internal_Content_Add(Index, OtherDC.Content[Index].Copy(this, this.DrawingDocument, oPr), false);
	}
};
CDocumentContent.prototype.Copy3 = function(Parent)//для заголовков диаграмм
{
	var DC = new CDocumentContent(Parent, this.DrawingDocument, 0, 0, 0, 0, this.Split, this.TurnOffInnerWrap, true);

	// Копируем содержимое
	DC.Internal_Content_RemoveAll();

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		DC.Internal_Content_Add(Index, this.Content[Index].Copy2(DC), false);
	}
	return DC;
};
CDocumentContent.prototype.IsRecalculated = function()
{
	return (this.Recalculated && this.Pages.length > 0);
};
CDocumentContent.prototype.isDocumentContentInSmartArtShape = function ()
{
	return this.Parent && this.Parent.isObjectInSmartArt && this.Parent.isObjectInSmartArt();
};

CDocumentContent.prototype.createDuplicateForSmartArt = function (oPr, arrayOfTxBody) {
	if (!arrayOfTxBody) {
		return;
	}
	var that = this;
	var arrayOfDC = Array(arrayOfTxBody.length).fill(0).map(function (value, idx) {
		var DC = new CDocumentContent(arrayOfTxBody[idx], that.DrawingDocument, 0, 0, 0, 0, that.Split, that.TurnOffInnerWrap, true);
		DC.Internal_Content_RemoveAll();
		return DC;
	});

	var Count = this.Content.length;
	if (Count % arrayOfDC.length === 0) {
		var amountOfParOfPer = Count / arrayOfDC.length;
	} else {
		amountOfParOfPer = (Count - Count % (arrayOfDC.length)) / (arrayOfDC.length);
	}
	for (var i = 0; i < arrayOfDC.length; i += 1) {
		var DC = arrayOfDC[i];
		var firstParOnThisDC = i * amountOfParOfPer;
		var lastParOnThisDC = firstParOnThisDC + amountOfParOfPer;
		if (i === arrayOfDC.length - 1) {
			lastParOnThisDC = Count;
		}
		for (var j = firstParOnThisDC; j < lastParOnThisDC; j += 1) {
			DC.Internal_Content_Add(DC.Content.length, this.Content[j].createDuplicateForSmartArt(DC, oPr), false);
		}
	}
	arrayOfTxBody.forEach(function (txBody, idx) {
		txBody.setContent(arrayOfDC[idx]);
	})
}
/**
 * В текущем содержимом получаем все комментарии и создаем их копии. Это иногда нужно делать во время функции
 * копирования, чтобы создавать новый уникальный комментарий с новым Id. Обрабатываются ТОЛЬКО комментарии, у которых
 * начало и конец лежат в данном контейнере.
 */
CDocumentContent.prototype.CreateDuplicateComments = function()
{
	var arrComments = this.GetAllComments();

	var oComments = {};
	for (var nIndex = 0, nCount = arrComments.length; nIndex < nCount; ++nIndex)
	{
		var oMark      = arrComments[nIndex].Comment;
		var sCommentId = oMark.GetCommentId();

		if (!oComments[sCommentId])
			oComments[sCommentId] = {Start : null, End : null};

		if (oMark.IsCommentStart())
			oComments[sCommentId].Start = oMark;
		else
			oComments[sCommentId].End = oMark;
	}

	var oLogicDocument = this.LogicDocument ? this.LogicDocument : editor.WordControl.m_oLogicDocument;
	if (!oLogicDocument)
		return;

	var oDocumentComments = oLogicDocument.Comments;
	for (var sId in oComments)
	{
		var oItem = oComments[sId];

		if (oItem.Start && oItem.End)
		{
			var oOldComment = oDocumentComments.GetById(sId);
			if (oOldComment)
			{
				var oNewComment = oOldComment.Copy();
				oDocumentComments.Add(oNewComment);

				oLogicDocument.GetApi().sync_AddComment(oNewComment.GetId(), oNewComment.GetData());

				oItem.Start.SetCommentId(oNewComment.GetId());
				oItem.End.SetCommentId(oNewComment.GetId());
			}
		}
	}
};
/**
 * Устанавливаем родительский класс
 * @param oParent
 */
CDocumentContent.prototype.SetParent = function(oParent)
{
	this.Parent = oParent;
};
//-----------------------------------------------------------------------------------
// Функции, к которым идет обращение из контента
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.Get_PageContentStartPos = function(PageNum)
{
	if (this.Parent)
		return this.Parent.Get_PageContentStartPos(PageNum);

	return {X : 0, Y : 0, XLimit : 210, YLimit : 297};
};
CDocumentContent.prototype.Get_PageContentStartPos2 = function(StartPageIndex, StartColumnIndex, ElementPageIndex, ElementIndex)
{
    return this.Get_PageContentStartPos(StartPageIndex + ElementPageIndex);
};
CDocumentContent.prototype.Get_PageLimits = function(nCurPage)
{
	if (this.Parent && this.Parent.IsCell())
	{
		var Margins = this.Parent.GetMargins();

		var Y      = this.Pages[nCurPage].OriginY - Margins.Top.W;
		var YLimit = this.Pages[nCurPage].OriginYLimit + Margins.Bottom.W;
		var X      = this.Pages[nCurPage].OriginX - Margins.Left.W;
		var XLimit = this.Pages[nCurPage].OriginXLimit + Margins.Right.W;

		return {
			X      : X,
			XLimit : XLimit,
			Y      : Y,
			YLimit : YLimit
		};
	}
	else
	{
		if (!this.Parent || !this.LogicDocument)
			return {X : 0, Y : 0, XLimit : 210, YLimit : 297};

		var nPageAbs = this.GetAbsolutePage(nCurPage);
		var Index    = ( undefined !== this.LogicDocument.Pages[nPageAbs] ? this.LogicDocument.Pages[nPageAbs].Pos : 0 );
		var SectPr   = this.LogicDocument.SectionsInfo.Get_SectPr(Index).SectPr;

		var W = SectPr.GetPageWidth();
		var H = SectPr.GetPageHeight();

		return {X : 0, Y : 0, XLimit : W, YLimit : H};
	}
};
CDocumentContent.prototype.Get_PageFields = function(PageIndex, isHdrFtr, oSectPr)
{
	if (this.Parent && (this.Parent.IsCell() || (undefined !== AscFormat.CShape && this.Parent instanceof AscFormat.CShape)))
	{
		if (PageIndex < this.Pages.length && PageIndex >= 0)
		{
			var Y      = this.Pages[PageIndex].Y;
			var YLimit = this.Pages[PageIndex].YLimit;
			var X      = this.Pages[PageIndex].X;
			var XLimit = this.Pages[PageIndex].XLimit;

			return {X : X, XLimit : XLimit, Y : Y, YLimit : YLimit}
		}
		else
		{
			if (!this.LogicDocument)
				return {X : 0, Y : 0, XLimit : 210, YLimit : 297};

			var Page_abs = this.Get_AbsolutePage(PageIndex);
			var Index    = ( undefined !== this.LogicDocument.Pages[Page_abs] ? this.LogicDocument.Pages[Page_abs].Pos : 0 );
			var SectPr   = this.LogicDocument.SectionsInfo.Get_SectPr(Index).SectPr;

			var W = SectPr.GetPageWidth();
			var H = SectPr.GetPageHeight();

			return {X : 0, Y : 0, XLimit : W, YLimit : H};
		}
	}
	else
	{
		if (!this.Parent || !this.LogicDocument)
			return {X : 0, Y : 0, XLimit : 210, YLimit : 297};

		return this.LogicDocument.Get_PageFields(this.GetAbsolutePage(PageIndex), isHdrFtr);
	}
};
CDocumentContent.prototype.Get_EmptyHeight = function()
{
	var Count = this.Content.length;
	if (Count <= 0)
		return 0;

	var Element = this.Content[Count - 1];

	if (type_Paragraph === Element.GetType())
		return Element.Get_EmptyHeight();
	else
		return 0;
};
/**
 * Inner = true  - запрос пришел из содержимого,
 *         false - запрос пришел от родительского класса
 *         Запрос от родительского класса нужен, например, для колонтитулов, потому
 *         что у них врапится текст не колонтитула, а документа.
 */
CDocumentContent.prototype.CheckRange = function(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, CurPage, isInner, bMathWrap)
{
	if (!this.Parent)
		return [];

	if (undefined === isInner)
		isInner = true;

	if (this.IsBlockLevelSdtContent() && isInner)
		return this.Parent.CheckRange(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, CurPage, true, bMathWrap);

	if (this.LogicDocument && this.LogicDocument.IsDocumentEditor())
	{
		let oDocContent = this;
		if (this.Parent instanceof CBlockLevelSdt)
			oDocContent = this.Parent.Parent;

		if (!isInner || this.IsUseInnerWrap())
			return this.LogicDocument.DrawingObjects.CheckRange(X0, Y0, X1, Y1, _Y0, _Y1, X_lf, X_rf, this.Get_AbsolutePage(CurPage), [], oDocContent, bMathWrap);
	}

	return [];
};
CDocumentContent.prototype.IsUseInnerWrap = function()
{
	if (!this.IsHdrFtr())
		return true;

	let oLogicDocument = this.GetLogicDocument();
	return oLogicDocument ? oLogicDocument.GetCompatibilityMode() >= AscCommon.document_compatibility_mode_Word15 : false;
};
CDocumentContent.prototype.Is_PointInDrawingObjects = function(X, Y, Page_Abs)
{
	return this.LogicDocument && this.LogicDocument.DrawingObjects.pointInObjInDocContent(this, X, Y, Page_Abs);
};
CDocumentContent.prototype.Is_PointInFlowTable = function(X, Y, PageAbs)
{
	return this.LogicDocument && null !== this.LogicDocument.DrawingObjects.getTableByXY(X, Y, PageAbs, this);
};
CDocumentContent.prototype.Get_Numbering = function()
{
	return this.GetNumbering();
};
CDocumentContent.prototype.GetNumbering = function()
{
	if (this.LogicDocument && this.LogicDocument.GetNumbering)
		return this.LogicDocument.GetNumbering();

	return AscCommonWord.DEFAULT_NUMBERING;
};
CDocumentContent.prototype.Get_Styles = function(nLvl)
{
	return this.GetStyles(nLvl);
};
CDocumentContent.prototype.GetStyles = function(nLvl)
{
	if (this.bPresentation && this.Parent)
		return this.Parent.Get_Styles(nLvl);

	if (this.LogicDocument && this.LogicDocument.GetStyles)
		return this.LogicDocument.GetStyles();

	return AscWord.DEFAULT_STYLES;
};
CDocumentContent.prototype.Get_TableStyleForPara = function()
{
	return this.Parent ? this.Parent.Get_TableStyleForPara() : null;
};
CDocumentContent.prototype.Get_ShapeStyleForPara = function()
{
	return this.Parent ? this.Parent.Get_ShapeStyleForPara() : null;
};
CDocumentContent.prototype.Get_TextBackGroundColor = function()
{
	return this.Parent ? this.Parent.Get_TextBackGroundColor() : undefined;
};
CDocumentContent.prototype.Recalc_AllParagraphs_CompiledPr = function()
{
	var Count = this.Content.length;
	for (var Pos = 0; Pos < Count; Pos++)
	{
		var Item = this.Content[Pos];
		if (type_Paragraph === Item.GetType())
		{
			Item.Recalc_CompiledPr();
			Item.Recalc_RunsCompiledPr();
		}
		else if (type_Table === Item.GetType())
			Item.Recalc_CompiledPr2();
	}
};
CDocumentContent.prototype.Set_CurrentElement = function(Index, bUpdateStates)
{
	var ContentPos = Math.max(0, Math.min(this.Content.length - 1, Index));
	this.SetDocPosType(docpostype_Content);

	var CurPos = Math.max(0, Math.min(this.Content.length - 1, Index));

	this.Selection.Use      = false;
	this.Selection.Start    = false;
	this.Selection.Flag     = selectionflag_Common;
	this.Selection.StartPos = CurPos;
	this.Selection.EndPos   = CurPos;
	this.CurPos.ContentPos  = CurPos;

	if (true === this.Content[ContentPos].IsSelectionUse())
	{
		this.Selection.Use      = true;
		this.Selection.StartPos = ContentPos;
		this.Selection.EndPos   = ContentPos;
	}

	this.SetThisElementCurrent(bUpdateStates)
};
CDocumentContent.prototype.IsThisElementCurrent = function()
{
	if (this.Parent)
		return this.Parent.IsThisElementCurrent(this);

	return false;
};
CDocumentContent.prototype.SetThisElementCurrent = function(isUpdateStates)
{
	if (this.Parent)
		this.Parent.Set_CurrentElement(isUpdateStates, this.GetAbsolutePage(0), this);
};
// Получем ближающую возможную позицию курсора
CDocumentContent.prototype.Get_NearestPos = function(CurPage, X, Y, bAnchor, Drawing)
{
	// TODO: Возможно лучше вернуть null, и разобраться с ситуациями, когда Get_NearestPos возвращает null
	if (CurPage < 0)
	{
		Y       = 0;
		CurPage = 0;
	}
	else if (CurPage >= this.Pages.length)
	{
		CurPage = this.Pages.length - 1;
		Y       = 10000;
	}

	var PageAbs = this.Get_AbsolutePage(CurPage);

	if (this.Parent && this.Parent instanceof CHeaderFooter)
	{
		var bInText    = (null === this.IsInText(X, Y, CurPage) ? false : true);
		var nInDrawing = this.LogicDocument.DrawingObjects.IsInDrawingObject(X, Y, PageAbs, this);

		if (true != bAnchor)
		{
			// Проверяем попадание в графические объекты
			var NearestPos = this.LogicDocument.DrawingObjects.getNearestPos(X, Y, PageAbs, Drawing);
			if (( nInDrawing === DRAWING_ARRAY_TYPE_BEFORE || nInDrawing === DRAWING_ARRAY_TYPE_INLINE || ( false === bInText && nInDrawing >= 0 ) ) && null != NearestPos)
				return NearestPos;
		}
	}
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;

	var ContentPos = this.Internal_GetContentPosByXY(X, Y, CurPage);

	// Делаем логику как в ворде
	if (true != bAnchor && (0 < ContentPos || CurPage > 0) && ContentPos === this.Pages[CurPage].Pos && this.Pages[CurPage].EndPos > this.Pages[CurPage].Pos && type_Paragraph === this.Content[ContentPos].GetType() && true === this.Content[ContentPos].IsContentOnFirstPage())
		ContentPos++;

	// Заглушка для плохих Fixed-форм
	var oShape = this.Is_DrawingShape(true);
	if (oShape && oShape.isForm())
		ContentPos = 0;

	var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, CurPage);
	return this.Content[ContentPos].Get_NearestPos(ElementPageIndex, X, Y, bAnchor, Drawing);
};
// Проверяем, описывает ли данный класс содержимое ячейки
CDocumentContent.prototype.IsTableCellContent = function(isReturnCell)
{
	if (!this.Parent)
		return isReturnCell ? null : false;

	return this.Parent.IsCell(isReturnCell);
};
CDocumentContent.prototype.IsLastTableCellInRow = function(isSelection)
{
	if (!this.Parent.IsCell())
		return false;

	return this.Parent.IsLastTableCellInRow(isSelection);
};
/**
 * Проверяем находимся ли мы заголовке хоть какой-либо таблицы
 * @returns {boolean}
 */
CDocumentContent.prototype.IsTableHeader = function()
{
	var oCell = this.IsTableCellContent(true);
	if (oCell)
		return oCell.IsInHeader(true);

	return false;
};
CDocumentContent.prototype.IsTableFirstRowOnNewPage = function()
{
	if (!this.Parent || !this.Parent.IsCell())
		return false;

	return this.Parent.IsTableFirstRowOnNewPage();
};
CDocumentContent.prototype.Check_AutoFit = function()
{
	if (!this.Parent)
		return false;

	return this.Parent.Check_AutoFit();
};
// Проверяем, лежит ли данный класс в таблице
CDocumentContent.prototype.IsInTable = function(isReturnTopTable)
{
	if (!this.Parent)
		return isReturnTopTable ? null : false;

	return this.Parent.IsInTable(isReturnTopTable);
};
// Проверяем, является ли данный класс верхним, по отношению к другим классам DocumentContent, Document
CDocumentContent.prototype.Is_TopDocument = function(isReturnTopDocument)
{
	if (!this.Parent)
		return isReturnTopDocument ? this : true;

	return this.Parent.Is_TopDocument(isReturnTopDocument);
};
CDocumentContent.prototype.IsUseInDocument = function(Id)
{
	if (undefined !== Id && null !== Id)
	{
		let isFound = false;
		for (let nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			if (this.Content[nIndex].Get_Id() === Id)
			{
				isFound = true;
				break;
			}
		}

		if (!isFound)
			return false;
	}

	if (!this.Parent || !this.Parent.IsUseInDocument)
		return false;

	return this.Parent.IsUseInDocument(this.GetId());
};
CDocumentContent.prototype.IsHdrFtr = function(bReturnHdrFtr)
{
	if (this.Parent && this.Parent.IsHdrFtr)
		return this.Parent.IsHdrFtr(bReturnHdrFtr);
	else
		return (bReturnHdrFtr ? null : false);
};
CDocumentContent.prototype.IsFootnote = function(bReturnFootnote)
{
	if (this instanceof CFootEndnote)
		return bReturnFootnote ? this : true;

	if (this.Parent)
		return this.Parent.IsFootnote(bReturnFootnote);

	return (bReturnFootnote ? null : false);
};
CDocumentContent.prototype.Is_DrawingShape = function(bRetShape)
{
	if (this.Parent && this.Parent.Is_DrawingShape)
		return this.Parent.Is_DrawingShape(bRetShape);
	else
		return (bRetShape ? null : false);
};
CDocumentContent.prototype.IsMovingTableBorder = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.DrawingObjects.selectionIsTableBorder();
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (null != this.Selection.Data && true === this.Selection.Data.TableBorder)
			return true;
	}

	return false;
};
CDocumentContent.prototype.CheckTableCoincidence = function(Table)
{
	if (!this.Parent)
		return false;

	return this.Parent.CheckTableCoincidence(Table);
};
//-----------------------------------------------------------------------------------
// Основные функции, к которым идет обращение от родительского класса
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.Reset = function(X, Y, XLimit, YLimit)
{
	this.X      = X;
	this.Y      = Y;
	this.XLimit = XLimit;
	this.YLimit = YLimit;

	// Заглушка для работы курсора в новой таблице
	if (0 === this.CurPos.X && 0 === this.CurPos.Y)
	{
		this.CurPos.X = X;
		this.CurPos.Y = Y;

		this.CurPos.RealX = X;
		this.CurPos.RealY = Y;
	}

	this.ClipInfo = [];
};
CDocumentContent.prototype.SetUseXLimit = function(isUse)
{
	this.UseXLimit = isUse;
};
CDocumentContent.prototype.SetUseYLimit = function(isUse)
{
	this.UseYLimit = isUse;
};
CDocumentContent.prototype.Recalculate                    = function()
{
    if (typeof(editor) !== "undefined" && editor.isDocumentEditor)
    {
        editor.WordControl.m_oLogicDocument.bRecalcDocContent    = true;
        editor.WordControl.m_oLogicDocument.recalcDocumentConten = this;
        editor.WordControl.m_oLogicDocument.Recalculate();
    }
};
CDocumentContent.prototype.Reset_RecalculateCache = function()
{
	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		this.Content[Index].Reset_RecalculateCache();
	}
};
// Пересчитываем отдельную страницу DocumentContent
CDocumentContent.prototype.Recalculate_Page               = function(PageIndex, bStart)
{
	this.Recalculated = true;
	
	this.ShiftViewX = 0;
	this.ShiftViewY = 0;

	var oDocContentRI = this.GetDocumentContentForRecalcInfo();
	var oRecalcInfo   = oDocContentRI.RecalcInfo;

	if (0 === PageIndex && true === bStart && oDocContentRI === this)
	{
		oRecalcInfo.FlowObject                = null;
		oRecalcInfo.FlowObjectPageBreakBefore = false;
	}

    var StartIndex = 0;
    if (PageIndex > 0)
        StartIndex = this.Pages[PageIndex - 1].EndPos;

    if (true === bStart)
    {
        this.Pages.length         = PageIndex;
        this.Pages[PageIndex]     = new CDocumentPage();
        this.Pages[PageIndex].Pos = StartIndex;

        if (this.LogicDocument && this.LogicDocument.IsDocumentEditor() && oDocContentRI === this)
            this.LogicDocument.DrawingObjects.resetDrawingArrays(this.Get_AbsolutePage(PageIndex), oDocContentRI);
    }

    var Count = this.Content.length;

    var StartPos;
    if (0 === PageIndex)
    {
        StartPos = {
            X      : this.X,
            Y      : this.Y,
            XLimit : this.XLimit,
            YLimit : this.YLimit
        }
    }
    else
    {
		StartPos = this.Get_PageContentStartPos(PageIndex);
    }

    this.Pages[PageIndex].Update_Limits(StartPos);

    var X      = StartPos.X;
    var StartY = StartPos.Y;
    var Y      = StartY;
    var YLimit = StartPos.YLimit;
    var XLimit = StartPos.XLimit;

    var Result = recalcresult2_End;

    for (var Index = StartIndex; Index < Count; Index++)
    {
        // Пересчитываем элемент документа
        var Element = this.Content[Index];

        var RecalcResult = recalcresult_NextElement;
        var bFlow        = false;
		let oFramePr;
        if (Element.IsTable() && !Element.IsInline())
        {
            bFlow = true;

            if (true === oRecalcInfo.Can_RecalcObject())
            {
				var ElementPageIndex = 0;
				if ((0 === Index && 0 === PageIndex) || Index !== StartIndex)
				{
					Element.Set_DocumentIndex(Index);
					Element.Reset(X, Y, XLimit, YLimit, PageIndex, 0, 1);
					ElementPageIndex = 0;
				}
				else
				{
					ElementPageIndex = PageIndex - Element.PageNum;
				}

                var TempRecalcResult = Element.Recalculate_Page(ElementPageIndex);

				oRecalcInfo.Set_FlowObject(Element, ElementPageIndex, TempRecalcResult, -1, {
                    X      : X,
                    Y      : Y,
                    XLimit : XLimit,
                    YLimit : YLimit
                });

                if (this.DrawingObjects)
                    this.DrawingObjects.addFloatTable(new CFlowTable(Element, PageIndex));

                RecalcResult = recalcresult_CurPage;
            }
            else if (true === oRecalcInfo.Check_FlowObject(Element))
            {
                // Если у нас текущая страница совпадает с той, которая указана в таблице, тогда пересчитываем дальше
                if (Element.PageNum > PageIndex || (oRecalcInfo.FlowObjectPage <= 0 && Element.PageNum < PageIndex) || Element.PageNum === PageIndex)
                {
                    if (true === oRecalcInfo.FlowObjectPageBreakBefore)
                    {
                        // Добавляем начало таблицы в конец страницы так, чтобы не убралось ничего
                        Element.Set_DocumentIndex(Index);
                        Element.Reset(X, YLimit, XLimit, YLimit, PageIndex, 0, 1);
                        Element.Recalculate_Page(0);

						oRecalcInfo.FlowObjectPage++;
                        RecalcResult = recalcresult_NextPage;
                    }
                    else
                    {
                        var _X      = oRecalcInfo.AdditionalInfo.X;
                        var _Y      = oRecalcInfo.AdditionalInfo.Y;
                        var _XLimit = oRecalcInfo.AdditionalInfo.XLimit;
                        var _YLimit = oRecalcInfo.AdditionalInfo.YLimit;

                        // Пересчет нужнен для обновления номеров колонок и страниц
                        Element.Reset(_X, _Y, _XLimit, _YLimit, PageIndex, 0, 1);
                        RecalcResult = Element.Recalculate_Page(0);
						oRecalcInfo.FlowObjectPage++;

                        if (RecalcResult & recalcresult_NextElement)
							oRecalcInfo.Reset();
                    }
                }
                else
                {
                    RecalcResult = Element.Recalculate_Page(PageIndex - Element.PageNum);
					oRecalcInfo.FlowObjectPage++;

                    if (this.DrawingObjects)
                        this.DrawingObjects.addFloatTable(new CFlowTable(Element, PageIndex));

                    if (RecalcResult & recalcresult_NextElement)
						oRecalcInfo.Reset();
                }
            }
            else
            {
                RecalcResult = recalcresult_NextElement;
            }
        }
		else if (this.CanCalculateFrames() && (!Element.IsInline() || (!Element.IsParagraph() && (oFramePr = Element.GetFramePr()) && !oFramePr.IsInline())))
        {
            bFlow = true;

            if (true === oRecalcInfo.Can_RecalcObject())
            {
                var FramePr = Element.GetFramePr();

                var FlowCount = this.CountElementsInFrame(Index);

                var LD_PageLimits = this.LogicDocument.Get_PageLimits(PageIndex + this.Get_StartPage_Absolute());
                var LD_PageFields = this.LogicDocument.Get_PageFields(PageIndex + this.Get_StartPage_Absolute());

                var Page_W = LD_PageLimits.XLimit;
                var Page_H = LD_PageLimits.YLimit;

                var Page_Field_L = LD_PageFields.X;
                var Page_Field_R = LD_PageFields.XLimit;
                var Page_Field_T = LD_PageFields.Y;
                var Page_Field_B = LD_PageFields.YLimit;

                //--------------------------------------------------------------------------------------------------
                // 1. Рассчитаем размер рамки
                //--------------------------------------------------------------------------------------------------
                var FrameH = 0;
                var FrameW = -1;

                var Frame_XLimit = FramePr.Get_W();
                var Frame_YLimit = FramePr.Get_H();

				var FrameHRule = (undefined === FramePr.HRule ? (undefined !== Frame_YLimit ? Asc.linerule_AtLeast : Asc.linerule_Auto) : FramePr.HRule);

				if (undefined === Frame_XLimit || 0 === AscCommon.MMToTwips(Frame_XLimit))
                    Frame_XLimit = Page_Field_R - Page_Field_L;

                if (undefined === Frame_YLimit || 0 === AscCommon.MMToTwips(Frame_YLimit) || Asc.linerule_Auto === FrameHRule)
                    Frame_YLimit = Page_H;

                for (var TempIndex = Index; TempIndex < Index + FlowCount; TempIndex++)
                {
                    var TempElement = this.Content[TempIndex];
                    TempElement.Set_DocumentIndex(TempIndex);

                    var nElementPageIndex = 0;
                    if (Index != TempIndex || ( true != oRecalcInfo.FrameRecalc && ( ( 0 === Index && 0 === PageIndex ) || Index != StartIndex ) ))
					{
						TempElement.Reset(0, FrameH, Frame_XLimit, Asc.NoYLimit, PageIndex);
					}
					else
					{
						nElementPageIndex = PageIndex - Element.PageNum;
					}

                    TempElement.Recalculate_Page(nElementPageIndex);

                    FrameH = TempElement.Get_PageBounds(PageIndex - TempElement.Get_StartPage_Relative()).Bottom;
                }

				var oTableInfo = Element.GetMaxTableGridWidth();

				var nMaxGapLeft           = oTableInfo.GapLeft;
				var nMaxGridWidth         = oTableInfo.GridWidth;
				var nMaxGridWidthRightGap = oTableInfo.GridWidth + oTableInfo.GapRight;

				for (var nTempIndex = Index + 1; nTempIndex < Index + FlowCount; ++nTempIndex)
				{
					var oTempTableInfo = this.Content[nTempIndex].GetMaxTableGridWidth();

					if (oTempTableInfo.GapLeft > nMaxGapLeft)
						nMaxGapLeft = oTempTableInfo.GapLeft;

					if (oTempTableInfo.GridWidth > nMaxGridWidth)
						nMaxGridWidth = oTempTableInfo.GridWidth;

					if (oTempTableInfo.GridWidth + oTempTableInfo.GapRight > nMaxGridWidthRightGap)
						nMaxGridWidthRightGap = oTempTableInfo.GridWidth + oTempTableInfo.GapRight;
				}

                // Обработаем "авто" ширину рамки. Ширина "авто" может быть в случае, когда значение W в FramePr
                // отсутствует, когда, у нас ровно 1 параграф, с 1 строкой.
                if (Element.IsParagraph() && -1 === FrameW && 1 === FlowCount && 1 === Element.Lines.length && (undefined === FramePr.Get_W() || 0 === AscCommon.MMToTwips(FramePr.Get_W())))
                {
                    FrameW     = Element.GetAutoWidthForDropCap();
                    var ParaPr = Element.Get_CompiledPr2(false).ParaPr;
                    FrameW += ParaPr.Ind.Left + ParaPr.Ind.FirstLine;

                    // Если прилегание в данном случае не к левой стороне, тогда пересчитываем параграф,
                    // с учетом того, что ширина буквицы должна быть FrameW
                    if (AscCommon.align_Left != ParaPr.Jc)
                    {
                        TempElement.Reset(0, 0, FrameW, Frame_YLimit, PageIndex);
                        TempElement.Recalculate_Page(PageIndex);
                        FrameH = TempElement.Get_PageBounds(PageIndex - TempElement.Get_StartPage_Relative()).Bottom;
                    }
                }
				else if (-1 === FrameW)
				{
					if (Element.IsTable() && (!FramePr.Get_W() || 0 === AscCommon.MMToTwips(FramePr.Get_W())))
					{
						FrameW = nMaxGridWidth;
					}
					else
					{
						FrameW = Frame_XLimit;
					}
				}

				var nGapLeft  = nMaxGapLeft;
				var nGapRight = nMaxGridWidthRightGap > FrameW ? nMaxGridWidthRightGap - FrameW : 0;

				if (0 !== AscCommon.MMToTwips(FramePr.H) && ((Asc.linerule_AtLeast === FrameHRule && FrameH < FramePr.H) || Asc.linerule_Exact === FrameHRule))
				{
					FrameH = FramePr.H;
				}

                //--------------------------------------------------------------------------------------------------
                // 2. Рассчитаем положение рамки
                //--------------------------------------------------------------------------------------------------

                // Теперь зная размеры рамки можем рассчитать ее позицию
                var FrameHAnchor = ( FramePr.HAnchor === undefined ? c_oAscHAnchor.Margin : FramePr.HAnchor );
                var FrameVAnchor = ( FramePr.VAnchor === undefined ? c_oAscVAnchor.Text : FramePr.VAnchor );

                // Рассчитаем положение по горизонтали
                var FrameX = 0;
                if (undefined !== FramePr.XAlign || undefined === FramePr.X)
                {
                    var XAlign = c_oAscXAlign.Left;
                    if (undefined !== FramePr.XAlign)
                        XAlign = FramePr.XAlign;

                    switch (FrameHAnchor)
                    {
                        case c_oAscHAnchor.Page   :
                        {
                            switch (XAlign)
                            {
                                case c_oAscXAlign.Inside  :
                                case c_oAscXAlign.Outside :
                                case c_oAscXAlign.Left    :
                                    FrameX = Page_Field_L - FrameW;
                                    break;
                                case c_oAscXAlign.Right   :
                                    FrameX = Page_Field_R;
                                    break;
                                case c_oAscXAlign.Center  :
                                    FrameX = (Page_W - FrameW) / 2;
                                    break;
                            }

                            break;
                        }
                        case c_oAscHAnchor.Text   :
                        case c_oAscHAnchor.Margin :
                        {
                            switch (XAlign)
                            {
                                case c_oAscXAlign.Inside  :
                                case c_oAscXAlign.Outside :
                                case c_oAscXAlign.Left    :
                                    FrameX = Page_Field_L;
                                    break;
                                case c_oAscXAlign.Right   :
                                    FrameX = Page_Field_R - FrameW;
                                    break;
                                case c_oAscXAlign.Center  :
                                    FrameX = (Page_Field_R + Page_Field_L - FrameW) / 2;
                                    break;
                            }

                            break;
                        }
                    }

                }
                else
                {
                    switch (FrameHAnchor)
                    {
                        case c_oAscHAnchor.Page   :
                            FrameX = FramePr.X;
                            break;
                        case c_oAscHAnchor.Text   :
                        case c_oAscHAnchor.Margin :
                            FrameX = Page_Field_L + FramePr.X;
                            break;
                    }
                }

                if (FrameW + FrameX > Page_W)
                    FrameX = Page_W - FrameW;

                if (FrameX < 0)
                    FrameX = 0;

                // Рассчитаем положение по вертикали
                var FrameY = 0;
                if (undefined !== FramePr.YAlign)
                {
                    var YAlign = FramePr.YAlign;
                    // Случай c_oAscYAlign.Inline не обрабатывается, потому что такие параграфы считаются Inline

                    switch (FrameVAnchor)
                    {
                        case c_oAscVAnchor.Page   :
                        {
                            switch (YAlign)
                            {
                                case c_oAscYAlign.Inside  :
                                case c_oAscYAlign.Outside :
                                case c_oAscYAlign.Top     :
                                    FrameY = 0;
                                    break;
                                case c_oAscYAlign.Bottom  :
                                    FrameY = Page_H - FrameH;
                                    break;
                                case c_oAscYAlign.Center  :
                                    FrameY = (Page_H - FrameH) / 2;
                                    break;
                            }

                            break;
                        }
                        case c_oAscVAnchor.Text   :
                        {
                            FrameY = Y;
                            break;
                        }
                        case c_oAscVAnchor.Margin :
                        {
                            switch (YAlign)
                            {
                                case c_oAscYAlign.Inside  :
                                case c_oAscYAlign.Outside :
                                case c_oAscYAlign.Top     :
                                    FrameY = Page_Field_T;
                                    break;
                                case c_oAscYAlign.Bottom  :
                                    FrameY = Page_Field_B - FrameH;
                                    break;
                                case c_oAscYAlign.Center  :
                                    FrameY = (Page_Field_B + Page_Field_T - FrameH) / 2;
                                    break;
                            }

                            break;
                        }
                    }
                }
                else
                {
                    var FramePrY = 0;
                    if (undefined !== FramePr.Y)
                        FramePrY = FramePr.Y;

                    switch (FrameVAnchor)
                    {
                        case c_oAscVAnchor.Page   :
                            FrameY = FramePrY;
                            break;
                        case c_oAscVAnchor.Text   :
                            FrameY = FramePrY + Y;
                            break;
                        case c_oAscVAnchor.Margin :

							// Если Y не задано, либо ровно 0, тогда MSWord считает это как привязка к тексту (баг 52903)
							if (undefined === FramePr.Y || 0 === AscCommon.MMToTwips(FramePr.Y))
								FrameY = Y + 0.001; // Погрешность, чтобы не сдвигалась предыдущая строка из-за обтекания
							else
								FrameY = FramePrY + Page_Field_T;

                            break;
                    }
                }

                if (FrameH + FrameY > Page_H)
                    FrameY = Page_H - FrameH;
	
				// Начинаем рамку с позиции начала следующего твипса, чтобы не было пересечений с объектами, заканчивающимися
				// на текущем твипсе
				// TODO: Проверить, нужно ли уменьшать размер рамки на 2 твипса (ранее уменьшали на 0.002, когда исправляли проблему с пересечением объектов в обтекании)
				let twFrameY = AscCommon.MMToTwips(AscCommon.CorrectMMToTwips(FrameY));
	
				FrameY = AscCommon.TwipsToMM(twFrameY + 1);
				FrameH -= AscCommon.TwipsToMM(2);

                if (FrameY < 0)
                    FrameY = 0;

				var FrameBounds;
				if (this.Content[Index].IsParagraph())
					FrameBounds = this.Content[Index].Get_FrameBounds(FrameX, FrameY, FrameW, FrameH);
				else
					FrameBounds = this.Content[Index].GetFirstParagraph().Get_FrameBounds(FrameX, FrameY, FrameW, FrameH);

				var FrameX2 = FrameBounds.X - nGapLeft,
					FrameY2 = FrameBounds.Y,
					FrameW2 = FrameBounds.W + nGapLeft + nGapRight,
					FrameH2 = FrameBounds.H;

                if ((FrameY2 + FrameH2 > Page_H || Y > Page_H - 0.001 ) && Index != StartIndex)
                {
					oRecalcInfo.Set_FrameRecalc(true);
                    this.Content[Index].StartFromNewPage();
                    RecalcResult = recalcresult_NextPage;
                }
                else
                {
					oRecalcInfo.Set_FrameRecalc(false);
                    for (var TempIndex = Index; TempIndex < Index + FlowCount; TempIndex++)
                    {
                        var TempElement = this.Content[TempIndex];
                        TempElement.Shift(TempElement.GetPagesCount() - 1, FrameX, FrameY);
                        TempElement.SetCalculatedFrame(new CCalculatedFrame(FramePr, FrameX, FrameY, FrameW, FrameH, FrameX2, FrameY2, FrameW2, FrameH2, PageIndex, Index, FlowCount));
                    }

                    var FrameDx = ( undefined === FramePr.HSpace ? 0 : FramePr.HSpace );
                    var FrameDy = ( undefined === FramePr.VSpace ? 0 : FramePr.VSpace );

                    this.DrawingObjects.addFloatTable(new CFlowParagraph(Element, FrameX2, FrameY2, FrameW2, FrameH2, FrameDx, FrameDy, Index, FlowCount, FramePr.Wrap));

                    Index += FlowCount - 1;

                    if (FrameY >= Y)
                        RecalcResult = recalcresult_NextElement;
                    else
                    {
						oRecalcInfo.Set_FlowObject(Element, FlowCount, recalcresult_NextElement, FlowCount);
                        RecalcResult = recalcresult_CurPage;
                    }
                }
            }
            else if (true === oRecalcInfo.Check_FlowObject(Element))
            {
                Index += oRecalcInfo.FlowObjectPage - 1;
				oRecalcInfo.Reset();
                RecalcResult = recalcresult_NextElement;
            }
            else
            {
                // Пропускаем
                RecalcResult = recalcresult_NextElement;
            }
        }
        else
        {
            if (( 0 === Index && 0 === PageIndex ) || Index != StartIndex)
            {
                Element.Set_DocumentIndex(Index);
                Element.Reset(X, Y, XLimit, YLimit, PageIndex, 0, 1);
                Element.SetUseXLimit(this.UseXLimit);
                Element.SetUseYLimit(this.UseYLimit);
            }

            if (this.IsEmptyParagraphAfterTableInTableCell(Index))
            {
                RecalcResult = recalcresult_NextElement;

                this.private_RecalculateEmptySectionParagraph(Element, this.Content[Index - 1], PageIndex, 0, 1);

                // Добавим в список особых параграфов
                this.Pages[PageIndex].EndSectionParas.push(Element);

                // Выставляем этот флаг, чтобы у нас не менялось значение по Y
                bFlow = true;
            }
            else
            {

                var ElementPageIndex = this.private_GetElementPageIndex(Index, PageIndex, 0, 1);
                RecalcResult         = Element.Recalculate_Page(ElementPageIndex);
            }
        }

        if (true != bFlow)
        {
            var ElementPageIndex = this.private_GetElementPageIndex(Index, PageIndex, 0, 1);
            Y                    = Element.Get_PageBounds(ElementPageIndex).Bottom;
        }

        if (RecalcResult & recalcresult_CurPage)
        {
        	if (true === this.IsBlockLevelSdtContent())
        		return recalcresult2_CurPage;

            // Такое не должно приходить в автофигурах, только в таблицах основного документа. Проверка на это находится в параграфе.
            if (RecalcResult & recalcresultflags_Footnotes)
				return recalcresult2_CurPage | recalcresultflags_Column | recalcresultflags_Footnotes;

            return this.Recalculate_Page(PageIndex, false);
        }
        else if (RecalcResult & recalcresult_NextElement)
        {
            // Ничего не делаем
        }
        else if (RecalcResult & recalcresult_NextPage)
        {
            this.Pages[PageIndex].EndPos = Index;
            Result                       = recalcresult2_NextPage;
            break;
        }
    }

    this.Pages[PageIndex].Bounds.Left   = X;
    this.Pages[PageIndex].Bounds.Top    = StartY;
    this.Pages[PageIndex].Bounds.Right  = XLimit;
    this.Pages[PageIndex].Bounds.Bottom = Y;

    if (Index >= Count)
        this.Pages[PageIndex].EndPos = Count - 1;
    return Result;
};
CDocumentContent.prototype.CanCalculateFrames = function()
{
	return (this.Parent && this.Parent instanceof CHeaderFooter);
};
/**
 * Получаем верхний док контент, который используется для пересчета плавающих объектов и различных переносов.
 * Как правило это нужно, если данный класс - это CBlockLevelSdt
 * @returns {?CDocumentContent}
 */
CDocumentContent.prototype.GetDocumentContentForRecalcInfo = function()
{
	var oDocContent = this;
	while (oDocContent.IsBlockLevelSdtContent())
	{
		if (oDocContent.Parent && oDocContent.Parent.GetTopDocumentContent)
			oDocContent = oDocContent.Parent.GetTopDocumentContent(true);
		else
			break;

		if (!oDocContent)
			return this;
	}

	return oDocContent;
};
CDocumentContent.prototype.RecalculateContent = function(fWidth, fHeight, nStartPage)
{
    this.Set_StartPage(nStartPage);
    this.Reset(0, 0, fWidth, 20000);
    var nRecalcResult = recalcresult2_NextPage;
    var nCurPage = 0;
    while ( recalcresult2_End !== nRecalcResult  )
        nRecalcResult = this.Recalculate_Page( nCurPage++, true );
};
CDocumentContent.prototype.RecalculateMinMaxContentWidth = function(isRotated)
{
	var Min   = 0;
	var Max   = 0;
	var Count = this.Content.length;

	if (true === isRotated)
	{
		for (var Pos = 0; Pos < Count; ++Pos)
		{
			var Element   = this.Content[Pos];
			var CurMinMax = Element.RecalculateMinMaxContentWidth(isRotated);

			Min += CurMinMax.Min;
			Max += CurMinMax.Max;
		}
	}
	else
	{
		for (var Pos = 0; Pos < Count; Pos++)
		{
			var Element   = this.Content[Pos];
			var CurMinMax = Element.RecalculateMinMaxContentWidth(isRotated);

			if (Min < CurMinMax.Min)
				Min = CurMinMax.Min;

			if (Max < CurMinMax.Max)
				Max = CurMinMax.Max;
		}
	}

	return {Min : Min, Max : Max};
};
/**
 * Специальная функция, которая проверяет, что по заданному индексу распологается пустой параграф,
 * который идет после таблицы. При этом весь контент сам является контентом ячейки другой таблицы.
 * @param nIndex {number}
 * @returns {boolean}
 */
CDocumentContent.prototype.IsEmptyParagraphAfterTableInTableCell = function(nIndex)
{
	var nCount = this.Content.length;
	if (nIndex !== nCount - 1 || nIndex <= 0)
		return false;

	var oElement     = this.Content[nIndex];
	var oPrevElement = this.Content[nIndex - 1];

	if (!oElement.IsParagraph() || !oPrevElement.IsTable() || !oPrevElement.IsInline() || !oElement.IsEmpty() || !this.IsTableCellContent())
		return false;

	// В 14-ой версии совместимости и ниже, если у такого параграфа есть нумерация, тогда он не считается пустым
	// В версиях совместимости больше 14, даже если у него есть нумерация мы его считаем пустым

	var oLogicDocument = this.GetLogicDocument();
	if (!oLogicDocument)
		return false;

	var nCompatibilityMode = oLogicDocument.GetCompatibilityMode();
	return !(nCompatibilityMode <= AscCommon.document_compatibility_mode_Word14 && oElement.HaveNumbering());
};
CDocumentContent.prototype.SaveRecalculateObject = function()
{
	var RecalcObj = new CDocumentRecalculateObject();
	RecalcObj.Save(this);
	return RecalcObj;
};
CDocumentContent.prototype.LoadRecalculateObject = function(RecalcObj)
{
	RecalcObj.Load(this);
};
CDocumentContent.prototype.PrepareRecalculateObject = function()
{
	this.ClipInfo = [];
	this.Pages    = [];
	var Count     = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		this.Content[Index].PrepareRecalculateObject();
	}
};
CDocumentContent.prototype.ReDraw = function(StartPage, EndPage)
{
	if (!this.Parent)
		return;

	if ("undefined" === typeof(StartPage))
		StartPage = this.Get_StartPage_Absolute();
	if ("undefined" === typeof(EndPage))
		EndPage = StartPage + this.Pages.length - 1;

	this.Parent.OnContentReDraw(StartPage, EndPage);
};
CDocumentContent.prototype.OnContentReDraw = function(StartPage, EndPage)
{
	this.Parent.OnContentReDraw(StartPage, EndPage);
};
CDocumentContent.prototype.Draw                           = function(nPageIndex, pGraphics)
{
	if (!this.IsRecalculated())
		return;
	
    var CurPage = nPageIndex - this.StartPage;
    if (CurPage < 0 || CurPage >= this.Pages.length)
        return;

    pGraphics.Start_Command(AscFormat.DRAW_COMMAND_CONTENT);

	var nPixelError = this.DrawingDocument && this.DrawingDocument.GetMMPerDot(1);
	
	let clipInfo = this.ClipInfo[CurPage];
	if (clipInfo)
	{
		let pageBounds = this.Pages[CurPage].Bounds;
		pGraphics.SaveGrState();
		
		let clipX0 = undefined !== clipInfo.X0 ? clipInfo.X0 : pageBounds.Left - nPixelError;
		let clipX1 = undefined !== clipInfo.X1 ? clipInfo.X1 : pageBounds.Right;
		let clipY0 = undefined !== clipInfo.Y0 ? clipInfo.Y0 : pageBounds.Top - nPixelError;
		let clipY1 = undefined !== clipInfo.Y1 ? clipInfo.Y1 : pageBounds.Bottom;
		pGraphics.AddClipRect(clipX0, clipY0, Math.abs(clipX1 - clipX0), Math.abs(clipY1 - clipY0));
	}
	
	if (this.transform)
	{
		pGraphics.SaveGrState();
		pGraphics.transform3(this.Get_ParentTextTransform());
	}

    var oPage = this.Pages[CurPage];
    for (var nIndex = oPage.Pos; nIndex <= oPage.EndPos; ++nIndex)
    {
    	var oElement = this.Content[nIndex];

    	if (oPage.IsFrame(oElement) || oPage.IsFlowTable(oElement))
    		continue;

        var nElementPageIndex = this.private_GetElementPageIndex(nIndex, CurPage, 0, 1);
		oElement.Draw(nElementPageIndex, pGraphics);
    }

	for (var nFlowTableIndex = 0, nFlowTablesCount = oPage.FlowTables.length; nFlowTableIndex < nFlowTablesCount; ++nFlowTableIndex)
	{
		var oTable = oPage.FlowTables[nFlowTableIndex];

		var nElementPageIndex = this.private_GetElementPageIndex(oTable.GetIndex(), nPageIndex, 0, 1);
		oTable.Draw(nElementPageIndex, pGraphics);
	}

	for (var nFrameIndex = 0, nFramesCount = oPage.Frames.length; nFrameIndex < nFramesCount; ++nFrameIndex)
	{
		var oFrame = oPage.Frames[nFrameIndex];

		var nL = oFrame.L2 - nPixelError;
		var nT = oFrame.T2 - nPixelError;
		var nH = oFrame.H2 + 2 * nPixelError;
		var nW = oFrame.W2 + 2 * nPixelError;

		pGraphics.SaveGrState();
		pGraphics.AddClipRect(nL, nT, nW, nH);

		for (var nFlowIndex = oFrame.StartIndex; nFlowIndex < oFrame.StartIndex + oFrame.FlowCount; ++nFlowIndex)
		{
			var nElementPageIndex = this.private_GetElementPageIndex(nFlowIndex, nPageIndex, 0, 1);
			this.Content[nFlowIndex].Draw(nElementPageIndex, pGraphics);
		}

		pGraphics.RestoreGrState();
	}
	
	if (this.transform)
		pGraphics.RestoreGrState();
	
	if (clipInfo)
		pGraphics.RestoreGrState();

    pGraphics.End_Command();
};
CDocumentContent.prototype.GetAllComments = function(AllComments)
{
	if (undefined === AllComments)
		AllComments = [];

	var Count = this.Content.length;
	for (var Pos = 0; Pos < Count; Pos++)
	{
		var Item = this.Content[Pos];
		Item.GetAllComments(AllComments);
	}

	return AllComments;
};
CDocumentContent.prototype.GetAllMaths = function(AllMaths)
{
	if (undefined === AllMaths)
		AllMaths = [];

	var Count = this.Content.length;
	for (var Pos = 0; Pos < Count; Pos++)
	{
		var Item = this.Content[Pos];
		Item.GetAllMaths(AllMaths);
	}

	return AllMaths;
};
CDocumentContent.prototype.HaveMaths = function()
{
	return this.GetAllMaths().length > 0;
};
CDocumentContent.prototype.GetAllFloatElements = function(FloatObjs)
{
	if (undefined === FloatObjs)
		FloatObjs = [];

	var Count = this.Content.length;
	for (var Pos = 0; Pos < Count; Pos++)
	{
		var Item = this.Content[Pos];

		if (true !== Item.Is_Inline())
			FloatObjs.push(Item);

		Item.GetAllFloatElements(FloatObjs);
	}

	return FloatObjs;
};
CDocumentContent.prototype.Shift = function(CurPage, Dx, Dy, keepClip)
{
	if (!this.IsRecalculated())
		return;
	
	this.Pages[CurPage].Shift(Dx, Dy);

	if (this.ClipInfo[CurPage] && true !== keepClip)
		this.ClipInfo[CurPage].shift(Dx, Dy);
	
	var StartPos = this.Pages[CurPage].Pos;
	var EndPos   = this.Pages[CurPage].EndPos;
	for (var Index = StartPos; Index <= EndPos; Index++)
	{
		var Element          = this.Content[Index];
		var ElementPageIndex = this.private_GetElementPageIndex(Index, CurPage, 0, 1);
		Element.Shift(ElementPageIndex, Dx, Dy);
	}
};
CDocumentContent.prototype.ShiftView = function(nDx, nDy)
{
	this.Shift(0, nDx, nDy, true);
	this.ShiftViewX += nDx;
	this.ShiftViewY += nDy;
};
CDocumentContent.prototype.ResetShiftView = function()
{
	this.Shift(0, -this.ShiftViewX, -this.ShiftViewY, true);
	this.ShiftViewX = 0;
	this.ShiftViewY = 0;
};
CDocumentContent.prototype.ShiftViewToFirstLine = function()
{
	let para = this.GetElement(0);
	if (this.GetElementsCount() <= 0 || !para || !para.IsParagraph())
		return;
	
	let paraState = para.SaveSelectionState();
	para.MoveCursorToStartPos();
	this.ResetShiftView();
	this.CheckFormViewWindow(true);
	para.LoadSelectionState(paraState);
};
CDocumentContent.prototype.CheckShiftView = function()
{
	var oForm = this.GetInnerForm();
	if (!this.LogicDocument
		|| !oForm
		|| oForm.IsCheckBox()
		|| oForm.IsPicture()
		|| (oForm.IsTextForm() && oForm.GetTextFormPr().IsComb())
		|| oForm.IsAutoFitContent()
		|| this.Content.length !== 1
		|| !this.Content[0].IsParagraph())
		return;
	
	this.ShiftViewToFirstLine();
};
CDocumentContent.prototype.CheckFormViewWindow = function()
{
	var isChanged = false;
	var oForm = this.GetInnerForm();
	if (!this.LogicDocument
		|| !oForm
		|| oForm.IsCheckBox()
		|| oForm.IsPicture()
		|| (oForm.IsTextForm() && oForm.GetTextFormPr().IsComb())
		|| oForm.IsAutoFitContent()
		|| this.Content.length !== 1
		|| !this.Content[0].IsParagraph())
		return false;

	var oParagraph  = this.GetElement(0);
	var oPageBounds = this.GetContentBounds(0);
	var oFormBounds = oForm.GetFixedFormBounds(true);

	var nDx = 0, nDy = 0, nPad = 0;

	if (oPageBounds.Right - oPageBounds.Left > oFormBounds.W)
	{
		if (oPageBounds.Left > oFormBounds.X)
			nDx = -oPageBounds.Left + oFormBounds.X;
		else if (oPageBounds.Right < oFormBounds.X + oFormBounds.W)
			nDx = oFormBounds.X + oFormBounds.W - oPageBounds.Right;
	}
	else
	{
		nDx = -this.ShiftViewX;
	}
	
	if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H)
	{
		if (oPageBounds.Top > oFormBounds.Y)
			nDy = -oPageBounds.Top + oFormBounds.Y;
		else if (oPageBounds.Bottom < oFormBounds.Y + oFormBounds.H)
			nDy = oFormBounds.Y + oFormBounds.H - oPageBounds.Bottom;
	}
	else
	{
		nDy = -this.ShiftViewY;
	}
	
	if (Math.abs(nDx) > 0.001 || Math.abs(nDy) > 0.001)
	{
		this.ShiftView(nDx, nDy);
		isChanged = true;
	}
	
	var oCursorPos  = oParagraph.GetCalculatedCurPosXY();
	var oLineBounds = oParagraph.GetLineBounds(oCursorPos.Internal.Line);

	nDx = 0;
	nDy = 0;

	var nCursorT = Math.min(oCursorPos.Y, oLineBounds.Top);
	var nCursorB = Math.max(oCursorPos.Y + oCursorPos.Height, oLineBounds.Bottom);
	var nCursorH = Math.max(0, nCursorB - nCursorT);

	if (oPageBounds.Right - oPageBounds.Left > oFormBounds.W)
	{
		if (oCursorPos.X < oFormBounds.X + nPad)
			nDx = oFormBounds.X + nPad - oCursorPos.X;
		else if (oCursorPos.X > oFormBounds.W - nPad)
			nDx = oFormBounds.W - nPad - oCursorPos.X;
	}

	if (oPageBounds.Bottom - oPageBounds.Top > oFormBounds.H)
	{
		if (nCursorH > oFormBounds.H - nPad)
			nDy = oFormBounds.Y + nPad - nCursorT - (nCursorH - oFormBounds.H - nPad);
		else if (nCursorT < oFormBounds.Y + nPad)
			nDy = oFormBounds.Y + nPad - nCursorT;
		else if (nCursorT + nCursorH > oFormBounds.H - nPad)
			nDy = oFormBounds.H - nPad - nCursorT - nCursorH;
	}
	
	// For multiline form we don't allow horizontal shift, for not multiline - don't allow vertical shift
	if (oForm.IsMultiLineForm())
		nDx = 0;
	else
		nDy = 0;

	if (Math.abs(nDx) > 0.001 || Math.abs(nDy) > 0.001)
	{
		this.ShiftView(nDx, nDy);
		isChanged = true;
	}

	return isChanged;
};
CDocumentContent.prototype.UpdateEndInfo = function()
{
	for (var Index = 0, Count = this.Content.length; Index < Count; Index++)
	{
		this.Content[Index].UpdateEndInfo();
	}
};
CDocumentContent.prototype.RecalculateCurPos = function(bUpdateX, bUpdateY, isUpdateTarget)
{
	var oCurPosInfo = null;

	if (docpostype_Content === this.CurPos.Type)
	{
		let drawingDocument = Asc.editor.getDrawingDocument();
		if (this.transform && drawingDocument)
			drawingDocument.MultiplyTargetTransform(this.transform.CreateDublicate());
		
		if (this.CurPos.ContentPos >= 0 && undefined !== this.Content[this.CurPos.ContentPos])
		{
			this.private_CheckCurPage();

			if (this.CurPage > 0 && true === this.Parent.IsHdrFtr(false))
			{
				this.CurPage = 0;
				this.DrawingDocument.TargetEnd();
			}
			else
			{
				oCurPosInfo = this.Content[this.CurPos.ContentPos].RecalculateCurPos(bUpdateX, bUpdateY, isUpdateTarget);
			}
		}
	}
	else // if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		oCurPosInfo = this.LogicDocument.DrawingObjects.recalculateCurPos(bUpdateX, bUpdateY, isUpdateTarget);
	}

	if (oCurPosInfo)
	{
		if (bUpdateX)
			this.CurPos.RealX = oCurPosInfo.X;

		if (bUpdateY)
			this.CurPos.RealY = oCurPosInfo.Y;
	}

	return oCurPosInfo;
};
CDocumentContent.prototype.Get_PageBounds = function(CurPage, Height, bForceCheckDrawings)
{
	if (!this.IsRecalculated())
		return new AscWord.CDocumentBounds(0, 0, 0, 0);

	if (CurPage < 0)
		CurPage = 0;

	if (CurPage >= this.Pages.length)
		CurPage = this.Pages.length - 1;

	var Bounds  = this.Pages[CurPage].Bounds;
	var PageAbs = this.Get_AbsolutePage(CurPage);

	// В колонтитуле не учитывается.
	if ((true != this.IsHdrFtr(false) && true !== this.IsBlockLevelSdtContent()) || true === bForceCheckDrawings)
	{
		// Учитываем все Drawing-объекты с обтеканием. Объекты без обтекания (над и под текстом) учитываем только в
		// случае, когда начальная точка (левый верхний угол) попадает в this.Y + Height

		var AllDrawingObjects = this.GetAllDrawingObjects();
		var Count             = AllDrawingObjects.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var Obj = AllDrawingObjects[Index];
			if (PageAbs === Obj.GetPageNum())
			{
				var ObjBounds = Obj.Get_Bounds();
				if (true === Obj.Use_TextWrap())
				{
					if (ObjBounds.Bottom > Bounds.Bottom)
						Bounds.Bottom = ObjBounds.Bottom;
				}
				else if (undefined !== Height && ObjBounds.Top < this.Y + Height)
				{
					if (ObjBounds.Bottom >= this.Y + Height)
						Bounds.Bottom = this.Y + Height;
					else if (ObjBounds.Bottom > Bounds.Bottom)
						Bounds.Bottom = ObjBounds.Bottom;
				}
			}
		}

		// Кроме этого пробежимся по всем Flow-таблицам и учтем их границы
		var Count = this.Content.length;
		for (var Index = 0; Index < Count; Index++)
		{
			var Element          = this.Content[Index];
			var ElementPageIndex = this.private_GetElementPageIndex(Index, CurPage, 0, 1);
			if (type_Table === Element.GetType() && true != Element.Is_Inline() && 0 <= ElementPageIndex && ElementPageIndex < Element.Get_PagesCount())
			{
				var TableBounds = Element.Get_PageBounds(ElementPageIndex);
				if (TableBounds.Bottom > Bounds.Bottom)
					Bounds.Bottom = TableBounds.Bottom;
			}
		}
	}

	return Bounds;
};
CDocumentContent.prototype.GetPageBounds = function(nCurPage, nHeight, isForceCheckDrawings)
{
	return this.Get_PageBounds(nCurPage, nHeight, isForceCheckDrawings);
};
CDocumentContent.prototype.GetContentBounds = function(CurPage)
{
	if (!this.IsRecalculated())
		return new AscWord.CDocumentBounds(0, 0, 0, 0);
	
	var oPage = this.Pages[CurPage];
	if (!oPage || oPage.Pos > oPage.EndPos)
		return this.Get_PageBounds(CurPage);

	var oBounds = null;
	for (var nIndex = oPage.Pos; nIndex <= oPage.EndPos; ++nIndex)
	{
		var oElement          = this.Content[nIndex];
		var nElementPageIndex = this.private_GetElementPageIndex(nIndex, CurPage, 0, 1);

		var oElementBounds = oElement.GetContentBounds(nElementPageIndex);

		if (null === oBounds)
		{
			oBounds = oElementBounds.Copy();
		}
		else
		{
			if (oElementBounds.Bottom > oBounds.Bottom)
				oBounds.Bottom = oElementBounds.Bottom;

			if (oElementBounds.Top < oBounds.Top)
				oBounds.Top = oElementBounds.Top;

			if (oElementBounds.Right > oBounds.Right)
				oBounds.Right = oElementBounds.Right;

			if (oElementBounds.Left < oBounds.Left)
				oBounds.Left = oElementBounds.Left;
		}
	}

	return oBounds;
};
CDocumentContent.prototype.Get_PagesCount = function()
{
	return this.Pages.length;
};
CDocumentContent.prototype.GetSummaryHeight = function()
{
	var Height = 0;
	for (var Page = 0; Page < this.Get_PagesCount(); Page++)
	{
		var Bounds = this.Get_PageBounds(Page);
		Height += Bounds.Bottom - Bounds.Top;
	}

	return Height;
};
CDocumentContent.prototype.Get_FirstParagraph = function()
{
	if (this.Content.length > 0)
		return this.Content[0].Get_FirstParagraph();

	return null;
};
CDocumentContent.prototype.GetAllParagraphs = function(Props, ParaArray)
{
	var arrParagraphs = (ParaArray ? ParaArray : []);

	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var Element = this.Content[Index];
		Element.GetAllParagraphs(Props, arrParagraphs);
	}

	return arrParagraphs;
};
CDocumentContent.prototype.GetAllTables = function(oProps, arrTables)
{
	if (!arrTables)
		arrTables = [];

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetAllTables(oProps, arrTables);
	}

	return arrTables;
};

/**
 * Специальный пресет с номером страницы для колонтитула
 * @param nAlignType
 * @param sStyleId
 */
CDocumentContent.prototype.SetHdrFtrPageNum = function(nAlignType, sStyleId)
{
	var arrDrawings = this.GetAllDrawingObjects();
	var oWatermark  = null;

	for (var nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
	{
		if (arrDrawings[nIndex].IsWatermark())
		{
			oWatermark = arrDrawings[nIndex];
			break;
		}
	}

	this.ClearContent(false);

	var oPara1 = new AscWord.Paragraph(this, this.bPresentation === true);
	var oPara2 = new AscWord.Paragraph(this, this.bPresentation === true);

	this.AddToContent(0, oPara1);
	this.AddToContent(1, oPara2);

	oPara1.SetParagraphStyleById(sStyleId);
	oPara2.SetParagraphStyleById(sStyleId);

	if (oWatermark)
	{
		var oSdt = new CInlineLevelSdt();
		oSdt.ReplacePlaceHolderWithContent();
		oSdt.SetDocPartObj(undefined, "Watermarks", true);
		oPara1.AddToContent(0, oSdt);

		var oRun = oSdt.GetFirstRun();
		if (oRun)
			oRun.AddToContent(0, oWatermark);
	}

	oPara1.SetParagraphAlign(nAlignType);
	var oRun = new ParaRun(oPara1, false);
	oRun.AddToContent(0, new AscWord.CRunPageNum());
	oPara1.AddToContent(0, oRun);
};
CDocumentContent.prototype.Clear_Content                 = function()
{
    this.RemoveSelection();

    this.CurPos =
    {
        X          : 0,
        Y          : 0,
        ContentPos : 0,
        RealX      : 0,
        RealY      : 0,
        Type       : docpostype_Content
    };

    this.Selection.Use = false;

    // Удаляем все элементы
    this.Internal_Content_RemoveAll();

    // Добавляем новый параграф
    var Para = new AscWord.Paragraph(this, this.bPresentation === true);
    this.Internal_Content_Add(0, Para);
};
CDocumentContent.prototype.Add_Content = function(OtherContent)
{
	return this.AddContent(OtherContent.Content);
};
/**
 * Полностью очищаем содержимое
 * @param {boolean} [isAddEmptyPara=true] добавлять ли пустой параграф
 */
CDocumentContent.prototype.ClearContent = function(isAddEmptyPara)
{
	this.Selection.Start    = false;
	this.Selection.Use      = false;
	this.Selection.StartPos = 0;
	this.Selection.EndPos   = 0;
	this.Selection.Flag     = selectionflag_Common;
	this.Selection.Data     = null;

	this.CurPos.X          = 0;
	this.CurPos.Y          = 0;
	this.CurPos.ContentPos = 0;
	this.CurPos.RealX      = 0;
	this.CurPos.RealY      = 0;
	this.CurPos.Type       = docpostype_Content;

	this.Internal_Content_RemoveAll();

	if (false !== isAddEmptyPara)
		this.Internal_Content_Add(0, new AscWord.Paragraph(this, this.bPresentation === true));
};
/**
 * Присоединяем к содержимому массив новых элементов (параграфов, таблиц)
 * @param {Paragraph, CTable} arrElements[]
 */
CDocumentContent.prototype.AddContent = function(arrElements)
{
	if (!arrElements || arrElements.length <= 0)
		return;

	if (this.Content.length <= 0 || true === this.IsEmpty())
	{
		if (this.Content.length > 0)
			this.Internal_Content_RemoveAll();

		for (var nIndex = 0, nCount = arrElements.length; nIndex < nCount; ++nIndex)
			this.Internal_Content_Add(nIndex, arrElements[nIndex]);
	}
	else
	{
		this.Content[this.Content.length - 1].Set_DocumentNext(arrElements[0]);
		arrElements[0].Set_DocumentPrev(this.Content[this.Content.length - 1]);

		for (var nIndex = 0, nCount = arrElements.length; nIndex < nCount; ++nIndex)
			this.Internal_Content_Add(this.Content.length, arrElements[nIndex]);
	}
};
CDocumentContent.prototype.Is_Empty = function()
{
	if (this.Content.length > 1 || type_Paragraph !== this.Content[0].GetType())
		return false;

	return this.Content[0].IsEmpty({SkipPlcHldr : false});
};


CDocumentContent.prototype.IsEmpty = function()
{
	return this.Is_Empty();
};
CDocumentContent.prototype.Is_CurrentElementTable = function()
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.isCurrentElementTable();
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Table == this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Table == this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		return true;
	}
	return false;
};
CDocumentContent.prototype.Is_CurrentElementParagraph = function()
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.isCurrentElementParagraph();
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Table == this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Table == this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		return false;
	}

	return true;
};
CDocumentContent.prototype.GetCurrentParagraph = function(bIgnoreSelection, arrSelectedParagraphs, oPr)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.getCurrentParagraph(bIgnoreSelection, arrSelectedParagraphs, oPr);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (arrSelectedParagraphs)
		{
			if (true === this.ApplyToAll)
			{
				for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
				{
					this.Content[nPos].SetApplyToAll(true);
					this.Content[nPos].GetCurrentParagraph(false, arrSelectedParagraphs, oPr);
					this.Content[nPos].SetApplyToAll(false);
				}
			}
			else if (true === this.Selection.Use)
			{
				var nStartPos = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
				var nEndPos   = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;
				for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
				{
					this.Content[nPos].GetCurrentParagraph(false, arrSelectedParagraphs, oPr);
				}
			}
			else
			{
				this.Content[this.CurPos.ContentPos].GetCurrentParagraph(false, arrSelectedParagraphs, oPr);
			}
		}
		else
		{
			let pos = this.CurPos.ContentPos;
			if (this.Selection.Use && true !== bIgnoreSelection)
			{
				pos = this.Selection.StartPos;
				if (oPr)
				{
					if (oPr.FirstInSelection)
						pos = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
					else if (oPr.LastInSelection)
						pos = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;
				}
			}
			
			if (pos < 0 || pos >= this.Content.length)
				return null;
			
			return this.Content[pos].GetCurrentParagraph(bIgnoreSelection, null, oPr);
		}
	}

	return null;
};
CDocumentContent.prototype.GetCurrentTablesStack = function(arrTables)
{
	if (!arrTables)
		arrTables = [];

	if (true === this.Selection.Use)
	{
		if (this.Selection.StartPos === this.Selection.EndPos)
			return this.Content[this.CurPos.ContentPos].GetCurrentTablesStack(arrTables);
		else
			return arrTables;
	}
	else
	{
		return this.Content[this.CurPos.ContentPos].GetCurrentTablesStack(arrTables);
	}
};
CDocumentContent.prototype.IsContentOnFirstPage = function()
{
	if (this.Content.length <= 0)
		return false;

	var Element = this.Content[0];
	return Element.IsContentOnFirstPage();
};
CDocumentContent.prototype.StartFromNewPage = function()
{
	this.Pages.length = 1;
	this.Pages[0]     = new CDocumentPage();

	var Element = this.Content[0];
	Element.StartFromNewPage();
};
CDocumentContent.prototype.Get_ParentTextTransform = function()
{
	let parentTransform = this.Parent ? this.Parent.Get_ParentTextTransform() : null;
	let transform = this.transform ? this.transform.CreateDublicate() : null;
	if (transform && parentTransform)
	{
		global_MatrixTransformer.MultiplyAppend(transform, parentTransform);
		return transform;
	}
	return transform || parentTransform;
};
CDocumentContent.prototype.GetFullTransform = function()
{
	return this.Get_ParentTextTransform();
};
CDocumentContent.prototype.GetCurrentTransform = function()
{
	return this.transform;
};
CDocumentContent.prototype.TransformPoint = function(x, y)
{
	if (!this.transform)
		return {x: x, y : y};
	
	let invert = global_MatrixTransformer.Invert(this.transform);
	return {
		x : invert.TransformPointX(x, y),
		y : invert.TransformPointY(x, y)
	};
};
CDocumentContent.prototype.IsTableBorder = function(X, Y, CurPage)
{
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;
	
	CurPage = Math.max(0, Math.min(this.Pages.length - 1, CurPage));

	var ElementPos       = this.Internal_GetContentPosByXY(X, Y, CurPage);
	var Element          = this.Content[ElementPos];
	var ElementPageIndex = this.private_GetElementPageIndex(ElementPos, CurPage, 0, 1);
	return Element.IsTableBorder(X, Y, ElementPageIndex);
};
CDocumentContent.prototype.IsInText = function(X, Y, CurPage)
{
	if (CurPage < 0 || CurPage >= this.Pages.length)
		CurPage = 0;
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;

	var ContentPos       = this.Internal_GetContentPosByXY(X, Y, CurPage);
	var Item             = this.Content[ContentPos];
	var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, CurPage);
	return Item.IsInText(X, Y, ElementPageIndex);
};
CDocumentContent.prototype.IsInDrawing = function(X, Y, CurPage)
{
	if (-1 != this.DrawingObjects.IsInDrawingObject(X, Y, this.Get_AbsolutePage(CurPage), this))
	{
		return true;
	}
	else
	{
		if (CurPage < 0 || CurPage >= this.Pages.length)
			CurPage = 0;
		
		let point = this.TransformPoint(X, Y);
		X = point.x;
		Y = point.y;
		
		var ContentPos = this.Internal_GetContentPosByXY(X, Y, CurPage);
		var Item       = this.Content[ContentPos];
		if (type_Table == Item.GetType())
		{
			var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, CurPage);
			return Item.IsInDrawing(X, Y, ElementPageIndex);
		}

		return false;
	}
};
CDocumentContent.prototype.Get_CurrentPage_Absolute = function()
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.getCurrentPageAbsolute();
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.IsNumberingSelection())
			return this.Selection.Data.CurPara.Get_CurrentPage_Absolute();

		var Pos = ( true === this.Selection.Use ? this.Selection.EndPos : this.CurPos.ContentPos );
		if (Pos >= 0 && Pos < this.Content.length)
			return this.Content[Pos].Get_CurrentPage_Absolute();
	}

	return 0;
};
CDocumentContent.prototype.Get_CurrentPage_Relative = function()
{
	return this.CurPage;
};
CDocumentContent.prototype.CollectDocumentStatistics = function(Stats)
{
	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var Element = this.Content[Index];
		Element.CollectDocumentStatistics(Stats);
	}
};
CDocumentContent.prototype.Document_CreateFontMap = function(FontMap)
{
	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var Element = this.Content[Index];
		Element.Document_CreateFontMap(FontMap);
	}
};
CDocumentContent.prototype.Document_CreateFontCharMap = function(FontCharMap)
{
	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var Element = this.Content[Index];
		Element.Document_CreateFontCharMap(FontCharMap);
	}
};
CDocumentContent.prototype.Document_Get_AllFontNames = function(fontNames)
{
	if (!fontNames)
		fontNames = [];
		
	var Count = this.Content.length;
	for (var Index = 0; Index < Count; Index++)
	{
		var Element = this.Content[Index];
		Element.Document_Get_AllFontNames(fontNames);
	}
	
	return fontNames;
};
CDocumentContent.prototype.Document_UpdateInterfaceState = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		var oDrawingObjects = this.LogicDocument.DrawingObjects;
		var oTargetTextObject = AscFormat.getTargetTextObject(oDrawingObjects);
		if (oTargetTextObject)
		{
			this.LogicDocument.Interface_Update_DrawingPr();
			oDrawingObjects.documentUpdateInterfaceState();
		}
		else
		{
			oDrawingObjects.resetInterfaceTextPr();
			oDrawingObjects.updateTextPr();
			this.LogicDocument.Interface_Update_DrawingPr();
			oDrawingObjects.updateParentParagraphParaPr();
		}
	}
	else //if (docpostype_Content === this.CurPos.Type)
	{
		if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
			|| (false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType()))
		{
			this.Interface_Update_TablePr();
			if (true == this.Selection.Use)
				this.Content[this.Selection.StartPos].Document_UpdateInterfaceState();
			else
				this.Content[this.CurPos.ContentPos].Document_UpdateInterfaceState();
		}
		else
		{
			this.UpdateInterfaceParaPr();
			this.UpdateInterfaceTextPr();

			// Если у нас в выделении находится 1 параграф, или курсор находится в параграфе
			if ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph == this.Content[this.Selection.StartPos].GetType())
				|| (false == this.Selection.Use && type_Paragraph == this.Content[this.CurPos.ContentPos].GetType()))
			{
				if (true == this.Selection.Use)
					this.Content[this.Selection.StartPos].Document_UpdateInterfaceState();
				else
					this.Content[this.CurPos.ContentPos].Document_UpdateInterfaceState();
			}
		}
	}
};
CDocumentContent.prototype.Document_UpdateRulersState = function(CurPage)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.LogicDocument.DrawingObjects.documentUpdateRulersState(CurPage);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			if (this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType())
			{
				var ElementPos       = this.Selection.StartPos;
				var Element          = this.Content[ElementPos];
				var ElementPageIndex = this.private_GetElementPageIndex(ElementPos, this.CurPage, Element.Get_StartColumn(), Element.Get_ColumnsCount());
				Element.Document_UpdateRulersState(ElementPageIndex);
			}
			else
			{
				var StartPos = ( this.Selection.EndPos <= this.Selection.StartPos ? this.Selection.EndPos : this.Selection.StartPos );
				var EndPos   = ( this.Selection.EndPos <= this.Selection.StartPos ? this.Selection.StartPos : this.Selection.EndPos );

				var FramePr = undefined;

				for (var Pos = StartPos; Pos <= EndPos; Pos++)
				{
					var Element = this.Content[Pos];
					if (type_Paragraph != Element.GetType())
					{
						FramePr = undefined;
						break;
					}
					else
					{
						var TempFramePr = Element.Get_FramePr();
						if (undefined === FramePr)
						{
							if (undefined === TempFramePr)
								break;

							FramePr = TempFramePr;
						}
						else if (undefined === TempFramePr || false === FramePr.Compare(TempFramePr))
						{
							FramePr = undefined;
							break;
						}
					}
				}

				if (undefined !== FramePr)
					this.Content[StartPos].Document_UpdateRulersState();
			}
		}
		else
		{
			var ElementPos       = this.CurPos.ContentPos;
			var Element          = this.Content[ElementPos];
			var ElementPageIndex = this.private_GetElementPageIndex(ElementPos, this.CurPage, Element.Get_StartColumn(), Element.Get_ColumnsCount());
			Element.Document_UpdateRulersState(ElementPageIndex);
		}
	}
};
CDocumentContent.prototype.Can_CopyCut = function()
{
	var bCanCopyCut = false;

	var LogicDocument  = null;
	var DrawingObjects = null;

	// Работаем с колонтитулом
	if (docpostype_DrawingObjects === this.CurPos.Type)
		DrawingObjects = this.DrawingObjects;
	else
		LogicDocument = this;

	if (null !== DrawingObjects)
	{
		if (true === DrawingObjects.isSelectedText())
			LogicDocument = DrawingObjects.getTargetDocContent();
		else
			bCanCopyCut = true;
	}

	if (null !== LogicDocument)
	{
		if (true === LogicDocument.IsSelectionUse() && !LogicDocument.IsSelectionEmpty())
		{
			if (selectionflag_Numbering === LogicDocument.Selection.Flag)
				bCanCopyCut = false;
			else if (LogicDocument.Selection.StartPos !== LogicDocument.Selection.EndPos)
				bCanCopyCut = true;
			else
				bCanCopyCut = LogicDocument.Content[LogicDocument.Selection.StartPos].Can_CopyCut();
		}
	}

	return bCanCopyCut;
};
CDocumentContent.prototype.MoveCursorToStartPos = function(AddToSelect)
{
	if (true === AddToSelect)
	{
		if (docpostype_DrawingObjects === this.CurPos.Type)
		{
			// TODO: Пока ничего не делаем, в дальнейшем надо будет делать в зависимости от селекта внутри
			//       автофигуры: если селект текста внутри, то делать для текста внутри, а если выделена
			//       сама автофигура, тогда мы перемещаем курсор влево от нее в контенте параграфа и выделяем все до конца
		}
		else if (docpostype_Content === this.CurPos.Type)
		{
			var StartPos = ( true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
			var EndPos   = 0;

			this.Selection.Start    = false;
			this.Selection.Use      = true;
			this.Selection.StartPos = StartPos;
			this.Selection.EndPos   = EndPos;
			this.Selection.Flag     = selectionflag_Common;

			this.CurPos.ContentPos = 0;
			this.SetDocPosType(docpostype_Content);

			for (var Index = StartPos - 1; Index >= EndPos; Index--)
			{
				this.Content[Index].SelectAll(-1);
			}

			this.Content[StartPos].MoveCursorToStartPos(true);
		}
	}
	else
	{
		this.RemoveSelection();

		this.Selection.Start    = false;
		this.Selection.Use      = false;
		this.Selection.StartPos = 0;
		this.Selection.EndPos   = 0;
		this.Selection.Flag     = selectionflag_Common;

		this.CurPos.ContentPos = 0;
		this.SetDocPosType(docpostype_Content);
		this.Content[0].MoveCursorToStartPos(false);
	}
};
CDocumentContent.prototype.MoveCursorToEndPos = function(AddToSelect, StartSelectFromEnd)
{
	if (true === AddToSelect)
	{
		if (docpostype_DrawingObjects === this.CurPos.Type)
		{
			// TODO: Пока ничего не делаем, в дальнейшем надо будет делать в зависимости от селекта внутри
			//       автофигуры: если селект текста внутри, то делать для текста внутри, а если выделена
			//       сама автофигура, тогда мы перемещаем курсор влево от нее в контенте параграфа и выделяем все до конца
		}
		else if (docpostype_Content === this.CurPos.Type)
		{
			var StartPos = ( true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos )
			var EndPos   = this.Content.length - 1;

			this.Selection.Start    = false;
			this.Selection.Use      = true;
			this.Selection.StartPos = StartPos;
			this.Selection.EndPos   = EndPos;
			this.Selection.Flag     = selectionflag_Common;

			this.CurPos.ContentPos = this.Content.length - 1;
			this.SetDocPosType(docpostype_Content);

			for (var Index = StartPos + 1; Index <= EndPos; Index++)
			{
				this.Content[Index].SelectAll(1);
			}

			this.Content[StartPos].MoveCursorToEndPos(true);
		}
	}
	else
	{
		if (true === StartSelectFromEnd)
		{
			this.Selection.Start    = false;
			this.Selection.Use      = true;
			this.Selection.StartPos = this.Content.length - 1;
			this.Selection.EndPos   = this.Content.length - 1;
			this.Selection.Flag     = selectionflag_Common;
			this.CurPos.ContentPos  = this.Content.length - 1;
			this.SetDocPosType(docpostype_Content);
			this.Content[this.Content.length - 1].MoveCursorToEndPos(false, true);
		}
		else
		{
			this.RemoveSelection();

			this.Selection.Start    = false;
			this.Selection.Use      = false;
			this.Selection.StartPos = 0;
			this.Selection.EndPos   = 0;
			this.Selection.Flag     = selectionflag_Common;

			this.CurPos.ContentPos = this.Content.length - 1;
			this.SetDocPosType(docpostype_Content);
			this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false);
		}
	}
};
CDocumentContent.prototype.MoveCursorUpToLastRow = function(X, Y, AddToSelect)
{
	this.SetCurPosXY(X, Y);
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;
	
	if (true === AddToSelect)
	{
		if (true !== this.Selection.Use)
		{
			this.CurPos.ContentPos = this.Content.length - 1;
			this.SetDocPosType(docpostype_Content);
			this.Selection.Use      = true;
			this.Selection.Start    = false;
			this.Selection.StartPos = this.CurPos.ContentPos;
			this.Selection.EndPos   = this.CurPos.ContentPos;
			this.Selection.Flag     = selectionflag_Common;

			this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false, true);
			this.Content[this.CurPos.ContentPos].MoveCursorUpToLastRow(X, Y, true);
		}
		else
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Content.length - 1;

			this.CurPos.ContentPos = EndPos;

			// Очистим старый селект кроме начального элемента
			var _S = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
			var _E = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;
			for (var nPos = _S; nPos <= _E; ++nPos)
			{
				if (nPos !== StartPos)
					this.Content[nPos].RemoveSelection();
			}

			if (StartPos === EndPos)
			{
				this.Selection.StartPos = StartPos;
				this.Selection.EndPos   = StartPos;
				this.Content[StartPos].MoveCursorUpToLastRow(X, Y, true);
			}
			else
			{
				this.Content[StartPos].MoveCursorToEndPos(true);
				for (var nPos = StartPos + 1; nPos <= EndPos; ++nPos)
				{
					this.Content[nPos].SelectAll(1);
				}

				this.Content[EndPos].MoveCursorUpToLastRow(X, Y, true);
			}
		}
	}
	else
	{
		this.CurPos.ContentPos = this.Content.length - 1;
		this.Content[this.CurPos.ContentPos].MoveCursorUpToLastRow(X, Y, false);
	}
};
CDocumentContent.prototype.MoveCursorDownToFirstRow = function(X, Y, AddToSelect)
{
	this.SetCurPosXY(X, Y);
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;
	
	if (true === AddToSelect)
	{
		if (true !== this.Selection.Use)
		{
			this.CurPos.ContentPos = 0;
			this.SetDocPosType(docpostype_Content);
			this.Selection.Use      = true;
			this.Selection.Start    = false;
			this.Selection.StartPos = 0;
			this.Selection.EndPos   = 0;
			this.Selection.Flag     = selectionflag_Common;

			this.Content[0].MoveCursorToStartPos(false);
			this.Content[0].MoveCursorDownToFirstRow(X, Y, true);
		}
		else
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = 0;

			this.CurPos.ContentPos = EndPos;

			// Очистим старый селект кроме начального элемента
			var _S = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
			var _E = this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;
			for (var nPos = _S; nPos <= _E; ++nPos)
			{
				if (nPos !== StartPos)
					this.Content[nPos].RemoveSelection();
			}

			if (StartPos === EndPos)
			{
				this.Selection.StartPos = StartPos;
				this.Selection.EndPos   = StartPos;
				this.Content[StartPos].MoveCursorDownToFirstRow(X, Y, true);
			}
			else
			{
				this.Content[StartPos].MoveCursorToStartPos(true);
				for (var nPos = EndPos; nPos < StartPos; ++nPos)
				{
					this.Content[nPos].SelectAll(-1);
				}

				this.Content[EndPos].MoveCursorDownToFirstRow(X, Y, true);
			}
		}
	}
	else
	{
		this.CurPos.ContentPos = 0;
		this.Content[this.CurPos.ContentPos].MoveCursorDownToFirstRow(X, Y, false);
	}
};
CDocumentContent.prototype.MoveCursorToCell = function(bNext)
{
	if (true === this.ApplyToAll)
	{
		if (1 === this.Content.length)
			this.Content[0].MoveCursorToCell(bNext);
	}
	else
	{
		if (docpostype_DrawingObjects == this.CurPos.Type)
		{
			this.LogicDocument.DrawingObjects.cursorMoveToCell(bNext);
		}
		else //if ( docpostype_Content == this.CurPos.Type )
		{
			if (true === this.Selection.Use)
			{
				if (this.Selection.StartPos === this.Selection.EndPos)
					this.Content[this.Selection.StartPos].MoveCursorToCell(bNext);
			}
			else
			{
				this.Content[this.CurPos.ContentPos].MoveCursorToCell(bNext);
			}
		}
	}
};
CDocumentContent.prototype.Set_ClipInfo = function(CurPage, X0, X1, Y0, Y1)
{
	this.ClipInfo[CurPage] = new AscWord.ClipRect(X0, X1, Y0, Y1);
};
CDocumentContent.prototype.IntersectClip = function(clipInfo, pageIndex)
{
	let pageClip = this.ClipInfo[pageIndex];
	if (!clipInfo && pageClip)
		clipInfo = pageClip.clone();
	else if (clipInfo)
		clipInfo.intersect(pageClip);
	
	return clipInfo;
};
CDocumentContent.prototype.GetClip = function(pageIndex)
{
	return this.ClipInfo[pageIndex] ? this.ClipInfo[pageIndex] : null;
};
CDocumentContent.prototype.IsApplyToAll = function()
{
	return this.ApplyToAll;
};
CDocumentContent.prototype.SetApplyToAll = function(isApplyAll)
{
	this.ApplyToAll = isApplyAll;
};
CDocumentContent.prototype.UpdateCursorType = function(X, Y, CurPage)
{
	if (CurPage < 0 || CurPage >= this.Pages.length)
		return this.DrawingDocument.SetCursorType("text", new AscCommon.CMouseMoveData());
	
	var bInText      = (null === this.IsInText(X, Y, CurPage) ? false : true);
	var bTableBorder = (null === this.IsTableBorder(X, Y, CurPage) ? false : true);

	// Ничего не делаем
	if (this.Parent instanceof CHeaderFooter && true === this.LogicDocument.DrawingObjects.updateCursorType(this.Get_AbsolutePage(CurPage), X, Y, {}, ( true === bInText || true === bTableBorder ? true : false )))
		return;
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;
	
	var ContentPos       = this.Internal_GetContentPosByXY(X, Y, CurPage);
	var Item             = this.Content[ContentPos];
	var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, CurPage);
	Item.UpdateCursorType(X, Y, ElementPageIndex);
};
//-----------------------------------------------------------------------------------
// Функции для работы с контентом
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.AddNewParagraph = function(bForceAdd)
{
    if (docpostype_DrawingObjects === this.CurPos.Type)
    {
        return this.DrawingObjects.addNewParagraph();
    }
    else //if ( docpostype_Content == this.CurPos.Type )
    {
        if (this.CurPos.ContentPos < 0)
            return false;

        // Сначала удаляем заселекченую часть
        if (true === this.Selection.Use)
        {
            this.Remove(1, true, false, true);
        }

        // Добавляем новый параграф
        var Item = this.Content[this.CurPos.ContentPos];

        // Если мы внутри параграфа, тогда:
        // 1. Если мы в середине параграфа, разделяем данный параграф на 2.
        //    При этом полностью копируем все настройки из исходного параграфа.
        // 2. Если мы в конце данного параграфа, тогда добавляем новый пустой параграф.
        //    Стиль у него проставляем такой какой указан у текущего в Style.Next.
        //    Если при этом у нового параграфа стиль будет такой же как и у старого,
        //    в том числе если стиля нет у обоих, тогда копируем еще все прямые настройки.
        //    (Т.е. если стили разные, а у исходный параграф был параграфом со списком, тогда
        //    новый параграф будет без списка).
        if (Item.IsParagraph())
        {
        	var isCheckAutoCorrect = false;
			let numPr = Item.GetNumPr();

            // Если текущий параграф пустой и с нумерацией, тогда удаляем нумерацию и отступы левый и первой строки
            if (true !== bForceAdd && numPr && true === Item.IsEmpty({SkipNewLine : true}) && true === Item.IsCursorAtBegin())
            {
				if (numPr.Lvl <= 0)
				{
					Item.RemoveNumPr();
					Item.Set_Ind({FirstLine : undefined, Left : undefined, Right : Item.Pr.Ind.Right}, true);
				}
				else
				{
					Item.SetNumPr(numPr.NumId, numPr.Lvl - 1);
				}
            }
            else
            {
                var ItemReviewType = Item.GetReviewType();

                var NewParagraph   = new AscWord.Paragraph(this, this.bPresentation === true);
				let firstPara, secondPara;
				if (Item.IsCursorAtBegin())
				{
					Item.SplitContent(NewParagraph, false);
					Item.Continue(NewParagraph);
					
					NewParagraph.Correct_Content();
					NewParagraph.MoveCursorToEndPos();
					Item.MoveCursorToStartPos();
					
					var nContentPos = this.CurPos.ContentPos;
					this.AddToContent(nContentPos, NewParagraph);
					this.CurPos.ContentPos = nContentPos + 1;
					
					firstPara  = NewParagraph;
					secondPara = Item;
				}
				else
				{
					if (true === Item.IsCursorAtEnd())
					{
						if (!Item.Lock.Is_Locked())
							isCheckAutoCorrect = true;

						var StyleId = Item.Style_Get();
						var NextId  = undefined;

						if (undefined != StyleId)
						{
							var Styles = this.Parent.Get_Styles();
							NextId     = Styles.Get_Next(StyleId);

							var oNextStyle = Styles.Get(NextId);
							if (!NextId || !oNextStyle || !oNextStyle.IsParagraphStyle())
								NextId = StyleId;
						}
						
						Item.SplitContent(NewParagraph, true);
						if (StyleId === NextId)
						{
							// Продолжаем (в плане настроек) новый параграф
							Item.Continue(NewParagraph);
						}
						else
						{
							// Простое добавление стиля, без дополнительных действий
							if (NextId === this.Get_Styles().Get_Default_Paragraph())
								NewParagraph.Style_Remove();
							else
								NewParagraph.Style_Add(NextId, true);
							
							let bidi = Item.GetDirectParaPr(false).Bidi;
							NewParagraph.SetParagraphBidi(bidi);
						}

						var LastRun = Item.Content[Item.Content.length - 1];
						if (LastRun && LastRun.Pr.Lang && LastRun.Pr.Lang.Val)
						{
							NewParagraph.SelectAll();
							NewParagraph.Add(new ParaTextPr({Lang : LastRun.Pr.Lang.Copy()}));
							NewParagraph.RemoveSelection();
						}
					}
					else
					{
						Item.Split(NewParagraph);
					}

					NewParagraph.Correct_Content();
					NewParagraph.MoveCursorToStartPos();

					var nContentPos = this.CurPos.ContentPos + 1;
					this.AddToContent(nContentPos, NewParagraph);
					this.CurPos.ContentPos = nContentPos;
					
					firstPara  = Item;
					secondPara = NewParagraph;
				}
				
				if (this.IsTrackRevisions())
				{
					firstPara.RemovePrChange();
					firstPara.SetReviewType(reviewtype_Add);
					secondPara.SetReviewType(ItemReviewType);
				}
				else if (reviewtype_Common !== ItemReviewType)
				{
					firstPara.SetReviewType(reviewtype_Common);
					secondPara.SetReviewType(ItemReviewType);
				}
				NewParagraph.CheckSignatureLinesOnAdd();
            }

			if (isCheckAutoCorrect)
			{
				var nContentPos = this.CurPos.ContentPos;

				var oParaEndRun = Item.GetParaEndRun();
				if (oParaEndRun)
					oParaEndRun.ProcessAutoCorrectOnParaEnd();

				this.CurPos.ContentPos = nContentPos;
			}
		}
		else if (type_Table === Item.GetType() || type_BlockLevelSdt === Item.GetType())
		{
			// Если мы находимся в начале первого параграфа первой ячейки, и
			// данная таблица - первый элемент, тогда добавляем параграф до таблицы.
			
			let newPos = -1;
			if (Item.IsCursorAtBegin(true) && (0 === this.CurPos.ContentPos || !this.Content[this.CurPos.ContentPos - 1].IsParagraph()))
				newPos = this.CurPos.ContentPos;
			else if (Item.IsCursorAtEnd() && (this.Content.length - 1 === this.CurPos.ContentPos || !this.Content[this.CurPos.ContentPos + 1].IsParagraph()))
				newPos = this.CurPos.ContentPos + 1;

			if (-1 !== newPos)
			{
				let newParagraph = new AscWord.Paragraph(this);
				this.Internal_Content_Add(newPos, newParagraph);
				this.CurPos.ContentPos = newPos;
				
				if (this.IsTrackRevisions())
				{
					newParagraph.RemovePrChange();
					newParagraph.SetReviewType(reviewtype_Add);
				}
			}
			else
			{
				Item.AddNewParagraph();
			}
		}
	}
};
// Расширяем документ до точки (X,Y) с помощью новых параграфов
// Y0 - низ последнего параграфа, YLimit - предел страницы
CDocumentContent.prototype.Extend_ToPos                       = function(X, Y)
{
	if (!this.LogicDocument || !this.LogicDocument.CanPerformAction || !this.LogicDocument.CanPerformAction())
		return;

	// В сносках нельзя делать ExtendToPos, смотри замечание в CFootnotesController.prototype.EndSelection
	if (this.IsFootnote())
		return;

	if (this.IsBlockLevelSdtContent())
	{
		var oParent = this.Parent.GetParent();
		if (oParent)
			return oParent.Extend_ToPos(X, Y);
	}

    var LastPara  = this.GetLastParagraph();
    var LastPara2 = LastPara;

    this.LogicDocument.StartAction(AscDFH.historydescription_Document_DocumentContentExtendToPos);
    this.LogicDocument.GetHistory().Set_Additional_ExtendDocumentToPos();

    while (true)
    {
        var NewParagraph = new AscWord.Paragraph(this, this.bPresentation === true);
		var NewRun       = new ParaRun(NewParagraph, false);
		NewParagraph.Add_ToContent(0, NewRun);

        var StyleId = LastPara.Style_Get();
        var NextId  = undefined;

		let oStyles = this.GetStyles();
        if (undefined != StyleId)
        {
            NextId = oStyles.Get_Next(StyleId);

            if (null === NextId || undefined === NextId)
                NextId = StyleId;
        }

        // Простое добавление стиля, без дополнительных действий
        if (NextId === oStyles.Get_Default_Paragraph())
            NewParagraph.Style_Remove();
        else
            NewParagraph.Style_Add(NextId, true);

        if (undefined != LastPara.TextPr.Value.FontSize || undefined !== LastPara.TextPr.Value.RFonts.Ascii)
        {
            var TextPr        = new CTextPr();
            TextPr.FontSize   = LastPara.TextPr.Value.FontSize;
            TextPr.FontSizeCS = LastPara.TextPr.Value.FontSize;
            TextPr.RFonts     = LastPara.TextPr.Value.RFonts.Copy();
            NewParagraph.SelectAll();
            NewParagraph.Apply_TextPr(TextPr);
        }

        var CurPage = LastPara.Pages.length - 1;
        var X0      = LastPara.Pages[CurPage].X;
        var Y0      = LastPara.Pages[CurPage].Bounds.Bottom;
        var XLimit  = LastPara.Pages[CurPage].XLimit;
        var YLimit  = LastPara.Pages[CurPage].YLimit;
        var PageNum = LastPara.PageNum;

		this.AddToContent(this.Content.length, NewParagraph, false);

        NewParagraph.Reset(X0, Y0, XLimit, YLimit, PageNum);
        var RecalcResult = NewParagraph.Recalculate_Page(0);

        if (!(RecalcResult & recalcresult_NextElement))
        {
			this.RemoveFromContent(this.Content.length - 1, 1, false);
            break;
        }

        if (NewParagraph.Pages[0].Bounds.Bottom > Y)
            break;

        LastPara = NewParagraph;
    }

    LastPara = this.Content[this.Content.length - 1];

    if (LastPara != LastPara2 || false === this.LogicDocument.Document_Is_SelectionLocked(AscCommon.changestype_None, {
            Type      : AscCommon.changestype_2_Element_and_Type,
            Element   : LastPara,
            CheckType : AscCommon.changestype_Paragraph_Content
        }))
    {
        // Теперь нам нужно вставить таб по X
        LastPara.extendLastLineToPos(X);
    }

    LastPara.MoveCursorToEndPos();
    LastPara.Document_SetThisElementCurrent(true);

    this.LogicDocument.Recalculate();
    this.LogicDocument.FinalizeAction();
};
CDocumentContent.prototype.AddInlineImage = function(W, H, Img, GraphicObject, bFlow)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.DrawingObjects.addInlineImage(W, H, Img, GraphicObject, bFlow);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true == this.Selection.Use)
			this.Remove(1, true);

		var Item = this.Content[this.CurPos.ContentPos];
		if (type_Paragraph == Item.GetType())
		{
			var Drawing;
			if (!AscCommon.isRealObject(GraphicObject))
			{
				Drawing   = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
				var Image = this.DrawingObjects.createImage(Img, 0, 0, W, H);
				Image.setParent(Drawing);
				Drawing.Set_GraphicObject(Image);
			}
			else if (GraphicObject.isSmartArtObject && GraphicObject.isSmartArtObject())
			{
				Drawing   = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
				GraphicObject.setParent(Drawing);
				Drawing.Set_GraphicObject(GraphicObject);
				Drawing.setExtent(GraphicObject.spPr.xfrm.extX, GraphicObject.spPr.xfrm.extY);
			}
			else
			{
				Drawing   = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
				var Image = this.DrawingObjects.getChartSpace2(GraphicObject, null);
				Image.setParent(Drawing);
				Drawing.Set_GraphicObject(Image);
				Drawing.setExtent(Image.spPr.xfrm.extX, Image.spPr.xfrm.extY);
			}
			if (true === bFlow)
			{
				Drawing.Set_DrawingType(drawing_Anchor);
				Drawing.Set_WrappingType(WRAPPING_TYPE_SQUARE);
				Drawing.Set_BehindDoc(false);
				Drawing.Set_Distance(3.2, 0, 3.2, 0);
				Drawing.Set_PositionH(Asc.c_oAscRelativeFromH.Column, false, 0, false);
				Drawing.Set_PositionV(Asc.c_oAscRelativeFromV.Paragraph, false, 0, false);
			}
			this.AddToParagraph(Drawing);
			this.Select_DrawingObject(Drawing.Get_Id());
		}
		else
		{
			Item.AddInlineImage(W, H, Img, GraphicObject, bFlow);
		}
	}
};

CDocumentContent.prototype.AddImages = function(aImages){
    if (docpostype_DrawingObjects === this.CurPos.Type)
    {
        return this.DrawingObjects.addImages(aImages);
    }
    else //if ( docpostype_Content === this.CurPos.Type )
    {
        if (true === this.Selection.Use)
            this.Remove(1, true);

        var Item = this.Content[this.CurPos.ContentPos];
        if (type_Paragraph === Item.GetType())
        {
            var Drawing, W, H;
            var ColumnSize = this.LogicDocument.GetColumnSize();
        	for(var i = 0; i < aImages.length; ++i){

                W = Math.max(1, ColumnSize.W);
                H = Math.max(1, ColumnSize.H);

				var _image = aImages[i];
				if(_image.Image)
				{
					var __w = Math.max((_image.Image.width * AscCommon.g_dKoef_pix_to_mm), 1);
					var __h = Math.max((_image.Image.height * AscCommon.g_dKoef_pix_to_mm), 1);
					W      = Math.max(5, Math.min(W, __w));
					H      = Math.max(5, Math.min((W * __h / __w)));
					Drawing   = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
					var Image = this.DrawingObjects.createImage(_image.src, 0, 0, W, H);
					Image.setParent(Drawing);
					Drawing.Set_GraphicObject(Image);
					this.AddToParagraph(Drawing);
				}
			}

			if(aImages.length === 1)
			{
				if(Drawing)
				{
					this.Select_DrawingObject(Drawing.Get_Id());
				}
			}
        }
        else
        {
            Item.AddImages(aImages);
        }
    }
};

CDocumentContent.prototype.AddOleObject = function(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.DrawingObjects.addOleObject(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true == this.Selection.Use)
			this.Remove(1, true);

		let Item = this.Content[this.CurPos.ContentPos];
		let Drawing;
		if (type_Paragraph == Item.GetType())
		{
			Drawing = new ParaDrawing(W, H, null, this.DrawingDocument, this, null);
			let Image   = this.DrawingObjects.createOleObject(Data, sApplicationId, Img, 0, 0, W, H, nWidthPix, nHeightPix, arrImagesForAddToHistory);
			Image.setParent(Drawing);
			Drawing.Set_GraphicObject(Image);

			this.AddToParagraph(Drawing);
			if(bSelect !== false)
			{
				this.Select_DrawingObject(Drawing.Get_Id());
			}
		}
		else
		{
			Drawing = Item.AddOleObject(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect, arrImagesForAddToHistory);
		}
		return Drawing;
	}
};
CDocumentContent.prototype.AddTextArt = function(nStyle)
{
	if (docpostype_DrawingObjects !== this.CurPos.Type)
	{
		var Item = this.Content[this.CurPos.ContentPos];
		if (type_Paragraph == Item.GetType())
		{
			var Drawing = new ParaDrawing(1828800 / 36000, 1828800 / 36000, null, this.DrawingDocument, this, null);
			var TextArt = this.DrawingObjects.createTextArt(nStyle, true);
			TextArt.setParent(Drawing);
			Drawing.Set_GraphicObject(TextArt);
			Drawing.Set_DrawingType(drawing_Anchor);
			Drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
			Drawing.Set_BehindDoc(false);
			Drawing.Set_Distance(3.2, 0, 3.2, 0);
			Drawing.Set_PositionH(Asc.c_oAscRelativeFromH.Column, false, 0, false);
			Drawing.Set_PositionV(Asc.c_oAscRelativeFromV.Paragraph, false, 0, false);
			if (true == this.Selection.Use)
				this.Remove(1, true);
			this.AddToParagraph(Drawing);
			if (TextArt.bSelectedText)
			{
				this.Select_DrawingObject(Drawing.Get_Id());
			}
			else
			{
				var oContent = Drawing.GraphicObj.getDocContent();
				oContent.Content[0].Document_SetThisElementCurrent(false);
				this.LogicDocument.SelectAll();
			}
		}
		else
		{
			Item.AddTextArt(nStyle);
		}
	}
};
CDocumentContent.prototype.AddSignatureLine = function(oSignatureDrawing)
{
	if (docpostype_DrawingObjects !== this.CurPos.Type)
	{
        var Item = this.Content[this.CurPos.ContentPos];
        if (type_Paragraph == Item.GetType())
        {
            var Drawing = oSignatureDrawing;

            if (true == this.Selection.Use)
                this.Remove(1, true);

            this.AddToParagraph(Drawing);
            this.Select_DrawingObject(Drawing.Get_Id());
        }
        else
        {
            Item.AddSignatureLine(oSignatureDrawing);
        }
	}
};
CDocumentContent.prototype.EditChart = function(Chart)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.editChart(Chart);
	}
};
CDocumentContent.prototype.AddInlineTable = function(nCols, nRows, nMode)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.addInlineTable(nCols, nRows, nMode);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return null;

		if (true === this.Selection.Use)
			this.Remove(1, true);

		// Добавляем таблицу
		var Item = this.Content[this.CurPos.ContentPos];

		// Если мы внутри параграфа, тогда разрываем его и на месте разрыва добавляем таблицу.
		// А если мы внутри таблицы, тогда добавляем таблицу внутрь текущей таблицы.
		if (Item.IsParagraph())
		{
			var oPage = this.Pages[this.CurPage];

			var nAdd = 0;
			if (this.LogicDocument)
				nAdd = this.LogicDocument.GetCompatibilityMode() <= AscCommon.document_compatibility_mode_Word14 && this.IsTableCellContent() ?  2 * 1.9 : 0;

			var W = Math.max(this.XLimit - this.X + nAdd, nCols * 2 * 1.9);

			var Grid = [];

			for (var Index = 0; Index < nCols; Index++)
				Grid[Index] = W / nCols;

			var NewTable = new CTable(this.DrawingDocument, this, true, nRows, nCols, Grid);
			NewTable.SetParagraphPrOnAdd(Item);

			var nContentPos = this.CurPos.ContentPos;
			if (true === Item.IsCursorAtBegin())
			{
				NewTable.MoveCursorToStartPos(false);
				this.AddToContent(nContentPos, NewTable);
				this.CurPos.ContentPos = nContentPos;
			}
			else
			{
				if (nMode < 0)
				{
					NewTable.MoveCursorToStartPos(false);

					if (Item.GetCurrentParaPos().Page > 0 && oPage && nContentPos === oPage.Pos)
					{
						this.AddToContent(nContentPos + 1, NewTable);
						this.CurPos.ContentPos = nContentPos + 1;
					}
					else
					{
						this.AddToContent(nContentPos, NewTable);
						this.CurPos.ContentPos = nContentPos;
					}
				}
				else if (nMode > 0)
				{
					NewTable.MoveCursorToStartPos(false);
					this.AddToContent(nContentPos + 1, NewTable);
					this.CurPos.ContentPos = nContentPos + 1;
				}
				else
				{
					var NewParagraph = new AscWord.Paragraph(this, this.bPresentation === true);
					Item.Split(NewParagraph);

					this.AddToContent(nContentPos + 1, NewParagraph);

					NewTable.MoveCursorToStartPos();
					this.AddToContent(nContentPos + 1, NewTable);
					this.CurPos.ContentPos = nContentPos + 1;
				}
			}

			return NewTable;
		}
		else
		{
			return Item.AddInlineTable(nCols, nRows, nMode);
		}
	}

	return null;
};
CDocumentContent.prototype.AddToParagraph = function(ParaItem, bRecalculate)
{
	if (true === this.ApplyToAll)
	{
		if (para_TextPr === ParaItem.Type)
		{
			for (var Index = 0; Index < this.Content.length; Index++)
			{
				var Item = this.Content[Index];
				Item.SetApplyToAll(true);
				Item.AddToParagraph(ParaItem);
				Item.SetApplyToAll(false);
			}
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.paragraphAdd(ParaItem, bRecalculate);
	}
	else // if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			var nSpaceCharCode = -1;
			if (this.LogicDocument && this.LogicDocument.IsWordSelection())
			{
				var sText = this.LogicDocument.GetSelectedText();
				if (sText.length > 1 && AscCommon.IsSpace(sText.charCodeAt(sText.length - 1)))
					nSpaceCharCode = sText.charCodeAt(sText.length - 1);
			}

			var Type = ParaItem.Get_Type();
			switch (Type)
			{
				case para_Math:
				case para_NewLine:
				case para_Text:
				case para_Space:
				case para_Tab:
				case para_PageNum:
				case para_Field:
				case para_FootnoteReference:
				case para_FootnoteRef:
				case para_Separator:
				case para_ContinuationSeparator:
				case para_InstrText:
				case para_EndnoteReference:
				case para_EndnoteRef:
				{
					if (ParaItem instanceof AscCommonWord.MathMenu)
					{
						var oInfo = this.GetSelectedElementsInfo();
						if (oInfo.GetMath())
						{
							var oMath = oInfo.GetMath();
							if (!oMath.IsParentEquationPlaceholder())
								ParaItem.SetText(oMath.Copy(true));
						}
						else if (!oInfo.IsMixedSelection())
						{
							ParaItem.SetText(this.GetSelectedText());
						}
					}


					// Если у нас что-то заселекчено и мы вводим текст или пробел
					// и т.д., тогда сначала удаляем весь селект.
					this.Remove(1, true, false, true);

					if (-1 !== nSpaceCharCode)
					{
						this.AddToParagraph(new AscWord.CRunSpace(nSpaceCharCode));
						this.MoveCursorLeft(false, false);
					}

					break;
				}
				case para_TextPr:
				{
					switch (this.Selection.Flag)
					{
						case selectionflag_Common:
						{
							// Текстовые настройки применяем ко всем параграфам, попавшим
							// в селект.
							var StartPos = this.Selection.StartPos;
							var EndPos   = this.Selection.EndPos;
							if (EndPos < StartPos)
							{
								var Temp = StartPos;
								StartPos = EndPos;
								EndPos   = Temp;
							}

							for (var Index = StartPos; Index <= EndPos; Index++)
							{
								this.Content[Index].AddToParagraph(ParaItem.Copy());
							}

							if (false != bRecalculate)
							{
								// Если в TextPr только HighLight, тогда не надо ничего пересчитывать, только перерисовываем
								if (true === ParaItem.Value.Check_NeedRecalc())
								{
									this.Recalculate();
								}
								else
								{
									// Просто перерисовываем нужные страницы
									var StartPage = this.Content[StartPos].Get_StartPage_Absolute();
									var EndPage   = this.Content[EndPos].Get_StartPage_Absolute() + this.Content[EndPos].GetPagesCount() - 1;
									this.ReDraw(StartPage, EndPage);
								}
							}

							break;
						}
						case selectionflag_Numbering:
						{
							// Текстовые настройки применяем к конкретной нумерации
							if (!this.Selection.Data || !this.Selection.Data.CurPara)
								break;

							if (undefined != ParaItem.Value.FontFamily)
							{
								var FName  = ParaItem.Value.FontFamily.Name;
								var FIndex = ParaItem.Value.FontFamily.Index;

								ParaItem.Value.RFonts          = new CRFonts();
								ParaItem.Value.RFonts.Ascii    = {Name : FName, Index : FIndex};
								ParaItem.Value.RFonts.EastAsia = {Name : FName, Index : FIndex};
								ParaItem.Value.RFonts.HAnsi    = {Name : FName, Index : FIndex};
								ParaItem.Value.RFonts.CS       = {Name : FName, Index : FIndex};
							}

							var oNumPr = this.Selection.Data.CurPara.GetNumPr();
							var oNum   = this.GetNumbering().GetNum(oNumPr.NumId);
							oNum.ApplyTextPr(oNumPr.Lvl, ParaItem.Value);
							break;
						}
					}

					return;
				}
			}
		}

		var nContentPos = this.CurPos.ContentPos;

		var Item     = this.Content[nContentPos];
		var ItemType = Item.GetType();

		if (para_NewLine === ParaItem.Type && true === ParaItem.IsPageOrColumnBreak())
		{
			if (type_Paragraph === ItemType)
			{
				if (true === Item.IsCursorAtBegin())
				{
					this.AddNewParagraph(true);

					if (this.Content[nContentPos] && this.Content[nContentPos].IsParagraph())
					{
						this.Content[nContentPos].AddToParagraph(ParaItem);
						this.Content[nContentPos].Clear_Formatting();
					}

					this.CurPos.ContentPos = nContentPos + 1;
				}
				else
				{
					this.AddNewParagraph(true);
					this.CurPos.ContentPos = nContentPos + 1;
					this.Content[nContentPos + 1].MoveCursorToStartPos();
					this.AddNewParagraph(true);

					if (this.Content[nContentPos + 1] && this.Content[nContentPos + 1].IsParagraph())
					{
						this.Content[nContentPos + 1].AddToParagraph(ParaItem);
						this.Content[nContentPos + 1].Clear_Formatting();
					}

					this.CurPos.ContentPos = nContentPos + 2;
					this.Content[nContentPos + 1].MoveCursorToStartPos();
				}

				if (false != bRecalculate)
				{
					this.Recalculate();

					Item.CurPos.RealX = Item.CurPos.X;
					Item.CurPos.RealY = Item.CurPos.Y;
				}
			}
			else if (type_BlockLevelSdt === Item.GetType())
			{
				Item.AddToParagraph(ParaItem);
			}
			else if (Item.IsTable())
			{
				// Разрыв страницы вне основного раздела вообще не добавляем
				if (ParaItem.IsPageBreak())
					return;

				if (Item.IsInnerTable())
				{
					Item.AddToParagraph(ParaItem);
				}
				else
				{
					var oNewTable = Item.Split();
					var oNewPara  = new AscWord.Paragraph(this);

					var nCurPos = this.CurPos.ContentPos;
					if (oNewTable)
					{
						this.AddToContent(nCurPos + 1, oNewTable);
						this.AddToContent(nCurPos + 1, oNewPara);
						this.CurPos.ContentPos = nCurPos + 1;
					}
					else
					{
						this.AddToContent(nCurPos, oNewPara);
						this.CurPos.ContentPos = nCurPos;
					}

					this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
				}

				this.Recalculate();
				return;
			}
			else
			{
				return;
			}
		}
		else
		{
			Item.AddToParagraph(ParaItem);

			if (false != bRecalculate)
			{
				if (para_TextPr === ParaItem.Type && false === ParaItem.Value.Check_NeedRecalc())
				{
					// Просто перерисовываем нужные страницы
					var StartPage = Item.Get_StartPage_Absolute();
					var EndPage   = StartPage + Item.GetPagesCount() - 1;
					this.ReDraw(StartPage, EndPage);
				}
				else
				{
					this.Recalculate();
				}

				if (type_Paragraph === ItemType)
				{
					Item.RecalculateCurPos();
					Item.CurPos.RealX = Item.CurPos.X;
					Item.CurPos.RealY = Item.CurPos.Y;
				}
			}
		}
	}
};
CDocumentContent.prototype.ClearParagraphFormatting = function(isClearParaPr, isClearTextPr)
{
	if (false !== isClearParaPr)
		isClearParaPr = true;

	if (false !== isClearTextPr)
		isClearTextPr = true;

	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.ClearParagraphFormatting(isClearParaPr, isClearTextPr);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.paragraphClearFormatting(isClearParaPr, isClearTextPr);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			if (selectionflag_Common === this.Selection.Flag)
			{
				var StartPos = this.Selection.StartPos;
				var EndPos   = this.Selection.EndPos;
				if (StartPos > EndPos)
				{
					var Temp = StartPos;
					StartPos = EndPos;
					EndPos   = Temp;
				}

				for (var Index = StartPos; Index <= EndPos; Index++)
				{
					var Item = this.Content[Index];
					Item.ClearParagraphFormatting(isClearParaPr, isClearTextPr);
				}
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.ClearParagraphFormatting(isClearParaPr, isClearTextPr);
		}
	}
};
CDocumentContent.prototype.Remove = function(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord)
{
	if (true === this.ApplyToAll)
	{
		this.SelectAll();
		this.private_Remove(1, false, true, false, false);

		this.CurPos = {
			X          : 0,
			Y          : 0,
			ContentPos : 0, // в зависимости, от параметра Type: озиция в Document.Content
			RealX      : 0, // позиция курсора, без учета расположения букв
			RealY      : 0, // это актуально для клавиш вверх и вниз
			Type       : docpostype_Content
		};

		this.Selection = {
			Start    : false,
			Use      : false,
			StartPos : 0,
			EndPos   : 0,
			Flag     : selectionflag_Common,
			Data     : null
		};

		return false;
	}

	if (undefined === bRemoveOnlySelection)
		bRemoveOnlySelection = false;

	if (undefined === bOnTextAdd)
		bOnTextAdd = false;

	if (docpostype_DrawingObjects === this.CurPos.Type)
		return this.LogicDocument.DrawingObjects.remove(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord);
	else //if ( docpostype_Content === this.CurPos.Type )
		return this.private_Remove(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord);
};
CDocumentContent.prototype.GetCursorPosXY = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.cursorGetPos();
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			if (selectionflag_Common === this.Selection.Flag)
			{
				return this.Content[this.Selection.EndPos].GetCursorPosXY();
			}

			return {X : 0, Y : 0};
		}
		else
		{
			return this.Content[this.CurPos.ContentPos].GetCursorPosXY();
		}
	}
};
CDocumentContent.prototype.MoveCursorLeft = function(AddToSelect, Word)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.cursorMoveLeft(AddToSelect, Word);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		var ReturnValue = true;

		this.RemoveNumberingSelection();
		if (true === this.Selection.Use)
		{
			if (true === AddToSelect)
			{
				if (false === this.Content[this.Selection.EndPos].MoveCursorLeft(true, Word))
				{
					if (0 !== this.Selection.EndPos)
					{
						if (this.Selection.EndPos > this.Selection.StartPos)
							this.Content[this.Selection.EndPos].RemoveSelection();

						this.Selection.EndPos--;
						this.CurPos.ContentPos = this.Selection.EndPos;

						var Item = this.Content[this.Selection.EndPos];
						if (this.Selection.StartPos <= this.Selection.EndPos)
							Item.MoveCursorLeft(true, Word);
						else
							Item.MoveCursorLeftWithSelectionFromEnd(Word);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				// Проверяем не обнулился ли селект в последнем параграфе. Такое могло быть, если была
				// заселекчена одна буква в последнем параграфе, а мы убрали селект последним действием.
				if (this.Selection.EndPos != this.Selection.StartPos && false === this.Content[this.Selection.EndPos].IsSelectionUse())
				{
					// Такая ситуация возможна только при прямом селекте (сверху вниз), поэтому вычитаем
					this.Selection.EndPos--;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
				{
					this.Selection.Use     = false;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}
			}
			else
			{
				// Нам нужно переместить курсор в левый край селекта, и отменить весь селект
				var Start = this.Selection.StartPos;
				if (Start > this.Selection.EndPos)
					Start = this.Selection.EndPos;

				this.CurPos.ContentPos = Start;
				if (false === this.Content[this.CurPos.ContentPos].MoveCursorLeft(false, Word))
				{
					if (this.CurPos.ContentPos > 0)
					{
						this.CurPos.ContentPos--;
						this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false, false);
					}
					else
					{
						ReturnValue = false;
					}
				}

				this.RemoveSelection();
			}
		}
		else
		{
			if (true === AddToSelect)
			{
				this.Selection.Use      = true;
				this.Selection.StartPos = this.CurPos.ContentPos;
				this.Selection.EndPos   = this.CurPos.ContentPos;

				if (false === this.Content[this.CurPos.ContentPos].MoveCursorLeft(true, Word))
				{
					// Нужно перейти в конец предыдущего элемент
					if (0 != this.CurPos.ContentPos)
					{
						this.CurPos.ContentPos--;
						this.Selection.EndPos = this.CurPos.ContentPos;

						var Item = this.Content[this.CurPos.ContentPos];
						Item.MoveCursorLeftWithSelectionFromEnd(Word);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
				{
					this.Selection.Use     = false;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}
			}
			else
			{
				if (false === this.Content[this.CurPos.ContentPos].MoveCursorLeft(false, Word))
				{
					// Нужно перейти в конец предыдущего элемент
					if (0 != this.CurPos.ContentPos)
					{
						this.CurPos.ContentPos--;
						this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false, false);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}
			}
		}

		return ReturnValue;
	}
};
CDocumentContent.prototype.MoveCursorLeftWithSelectionFromEnd = function(Word)
{
	this.RemoveSelection();

	if (this.Content.length <= 0)
		return;

	this.Selection.Use      = true;
	this.Selection.Start    = false;
	this.Selection.Data     = null;
	this.Selection.Flag     = selectionflag_Common;
	this.Selection.StartPos = this.Content.length - 1;
	this.Selection.EndPos   = this.Content.length - 1;

	this.Content[this.Content.length - 1].MoveCursorLeftWithSelectionFromEnd(Word);
};
CDocumentContent.prototype.MoveCursorRight = function(AddToSelect, Word, FromPaste)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.cursorMoveRight(AddToSelect, Word);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		var ReturnValue = true;

		this.RemoveNumberingSelection();
		if (true === this.Selection.Use)
		{
			if (true === AddToSelect)
			{
				if (false === this.Content[this.Selection.EndPos].MoveCursorRight(true, Word))
				{
					if (this.Content.length - 1 !== this.Selection.EndPos)
					{
						if (this.Selection.EndPos < this.Selection.StartPos)
							this.Content[this.Selection.EndPos].RemoveSelection();

						this.Selection.EndPos++;
						this.CurPos.ContentPos = this.Selection.EndPos;

						var Item = this.Content[this.Selection.EndPos];
						if (this.Selection.StartPos >= this.Selection.EndPos)
							Item.MoveCursorRight(true, Word);
						else
							Item.MoveCursorRightWithSelectionFromStart(Word);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				// Проверяем не обнулился ли селект в последнем параграфе. Такое могло быть, если была
				// заселекчена одна буква в последнем параграфе, а мы убрали селект последним действием.
				if (this.Selection.EndPos != this.Selection.StartPos && false === this.Content[this.Selection.EndPos].IsSelectionUse())
				{
					// Такая ситуация возможна только при обратном селекте (снизу вверх), поэтому вычитаем
					this.Selection.EndPos++;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
				{
					this.Selection.Use     = false;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}
			}
			else
			{
				// Нам нужно переместить курсор в правый край селекта, и отменить весь селект
				var End = this.Selection.EndPos;
				if (End < this.Selection.StartPos)
					End = this.Selection.StartPos;


				this.CurPos.ContentPos = End;

				if (true === this.Content[this.CurPos.ContentPos].IsSelectionToEnd() && this.CurPos.ContentPos < this.Content.length - 1)
				{
					this.CurPos.ContentPos = End + 1;
					this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
				}
				else
				{
					if (!this.Content[this.CurPos.ContentPos].MoveCursorRight(false, Word))
					{
						if (this.CurPos.ContentPos < this.Content.length - 1)
						{
							this.CurPos.ContentPos++;
							this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
						}
						else
						{
							ReturnValue = false;
						}
					}
				}

				this.RemoveSelection();
			}
		}
		else
		{
			if (true === AddToSelect)
			{
				this.Selection.Use      = true;
				this.Selection.StartPos = this.CurPos.ContentPos;
				this.Selection.EndPos   = this.CurPos.ContentPos;

				if (false === this.Content[this.CurPos.ContentPos].MoveCursorRight(true, Word))
				{
					// Нужно перейти в конец предыдущего элемента
					if (this.Content.length - 1 != this.CurPos.ContentPos)
					{
						this.CurPos.ContentPos++;
						this.Selection.EndPos = this.CurPos.ContentPos;

						var Item = this.Content[this.CurPos.ContentPos];
						Item.MoveCursorRightWithSelectionFromStart(Word);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
				{
					this.Selection.Use     = false;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}
			}
			else
			{
				if (false === this.Content[this.CurPos.ContentPos].MoveCursorRight(false, Word))
				{
					// Нужно перейти в начало следующего элемента
					if (this.Content.length - 1 != this.CurPos.ContentPos)
					{
						this.CurPos.ContentPos++;
						this.Content[this.CurPos.ContentPos].MoveCursorToStartPos(false);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}
			}
		}

		return ReturnValue;
	}
};
CDocumentContent.prototype.MoveCursorRightWithSelectionFromStart = function(Word)
{
	this.RemoveSelection();

	if (this.Content.length <= 0)
		return;

	this.Selection.Use      = true;
	this.Selection.Start    = false;
	this.Selection.Data     = null;
	this.Selection.Flag     = selectionflag_Common;
	this.Selection.StartPos = 0;
	this.Selection.EndPos   = 0;

	this.Content[0].MoveCursorRightWithSelectionFromStart(Word);
};
CDocumentContent.prototype.MoveCursorUp = function(AddToSelect)
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
		return this.LogicDocument.DrawingObjects.cursorMoveUp(AddToSelect);
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		var ReturnValue = true;

		this.RemoveNumberingSelection();
		if (true === this.Selection.Use)
		{
			if (true === AddToSelect)
			{
				var SelectDirection = this.Selection.StartPos === this.Selection.EndPos ? 0 : this.Selection.StartPos < this.Selection.EndPos ? 1 : -1;

				var Item = this.Content[this.Selection.EndPos];
				if (false === Item.MoveCursorUp(true))
				{
					if (0 != this.Selection.EndPos)
					{
						var TempXY        = Item.GetCurPosXY();
						this.CurPos.RealX = TempXY.X;
						this.CurPos.RealY = TempXY.Y;

						if (1 === SelectDirection)
							Item.RemoveSelection();

						this.Selection.EndPos--;
						Item = this.Content[this.Selection.EndPos];
						Item.MoveCursorUpToLastRow(this.CurPos.RealX, this.CurPos.RealY, true);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
					this.Selection.Use = false;

				this.CurPos.ContentPos = this.Selection.EndPos;
			}
			else
			{
				// Мы должны переместиться на строку выше, чем начало селекта
				var Start = this.Selection.StartPos;
				if (Start > this.Selection.EndPos)
					Start = this.Selection.EndPos;

				this.CurPos.ContentPos = Start;

				var Item = this.Content[this.CurPos.ContentPos];
				if (false === this.Content[this.CurPos.ContentPos].MoveCursorUp(false))
				{
					if (0 != this.CurPos.ContentPos)
					{
						var TempXY        = Item.GetCurPosXY();
						this.CurPos.RealX = TempXY.X;
						this.CurPos.RealY = TempXY.Y;

						this.CurPos.ContentPos--;
						Item = this.Content[this.CurPos.ContentPos];
						Item.MoveCursorUpToLastRow(this.CurPos.RealX, this.CurPos.RealY, false);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				this.RemoveSelection();
			}
		}
		else
		{
			if (true === AddToSelect)
			{
				this.Selection.Use      = true;
				this.Selection.StartPos = this.CurPos.ContentPos;
				this.Selection.EndPos   = this.CurPos.ContentPos;

				var Item = this.Content[this.CurPos.ContentPos];
				if (false === Item.MoveCursorUp(true))
				{
					if (0 != this.CurPos.ContentPos)
					{
						var TempXY        = Item.GetCurPosXY();
						this.CurPos.RealX = TempXY.X;
						this.CurPos.RealY = TempXY.Y;

						this.CurPos.ContentPos--;
						Item = this.Content[this.CurPos.ContentPos];
						Item.MoveCursorUpToLastRow(this.CurPos.RealX, this.CurPos.RealY, true);
						this.Selection.EndPos = this.CurPos.ContentPos;
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
					this.Selection.Use = false;

				this.CurPos.ContentPos = this.Selection.EndPos;
			}
			else
			{
				var Item = this.Content[this.CurPos.ContentPos];
				if (false === Item.MoveCursorUp(false))
				{
					if (0 != this.CurPos.ContentPos)
					{
						var TempXY        = Item.GetCurPosXY();
						this.CurPos.RealX = TempXY.X;
						this.CurPos.RealY = TempXY.Y;

						this.CurPos.ContentPos--;
						Item = this.Content[this.CurPos.ContentPos];
						Item.MoveCursorUpToLastRow(this.CurPos.RealX, this.CurPos.RealY, false);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}
			}
		}

		return ReturnValue;
	}
};
CDocumentContent.prototype.MoveCursorDown = function(AddToSelect)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
		return this.LogicDocument.DrawingObjects.cursorMoveDown(AddToSelect);
	else if (docpostype_Content === this.CurPos.Type)
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		var ReturnValue = true;
		this.RemoveNumberingSelection();

		if (true === this.Selection.Use)
		{
			if (true === AddToSelect)
			{
				var SelectDirection = this.Selection.StartPos === this.Selection.EndPos ? 0 : this.Selection.StartPos < this.Selection.EndPos ? 1 : -1;

				var Item = this.Content[this.Selection.EndPos];
				if (false === Item.MoveCursorDown(true))
				{
					if (this.Content.length - 1 != this.Selection.EndPos)
					{
						var TempXY        = Item.GetCurPosXY();
						this.CurPos.RealX = TempXY.X;
						this.CurPos.RealY = TempXY.Y;

						if (-1 === SelectDirection)
							Item.RemoveSelection();

						this.Selection.EndPos++;
						Item = this.Content[this.Selection.EndPos];
						Item.MoveCursorDownToFirstRow(this.CurPos.RealX, this.CurPos.RealY, true);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
					this.Selection.Use = false;

				this.CurPos.ContentPos = this.Selection.EndPos;
			}
			else
			{
				// Мы должны переместиться на строку ниже, чем конец селекта
				var End = this.Selection.EndPos;
				if (End < this.Selection.StartPos)
					End = this.Selection.StartPos;

				this.CurPos.ContentPos = End;

				var Item = this.Content[this.CurPos.ContentPos];
				if (false === this.Content[this.CurPos.ContentPos].MoveCursorDown(false))
				{
					if (this.Content.length - 1 != this.CurPos.ContentPos)
					{
						var TempXY        = Item.GetCurPosXY();
						this.CurPos.RealX = TempXY.X;
						this.CurPos.RealY = TempXY.Y;

						this.CurPos.ContentPos++;
						Item = this.Content[this.CurPos.ContentPos];
						Item.MoveCursorDownToFirstRow(this.CurPos.RealX, this.CurPos.RealY, false);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				this.RemoveSelection();
			}
		}
		else
		{
			if (true === AddToSelect)
			{
				this.Selection.Use      = true;
				this.Selection.StartPos = this.CurPos.ContentPos;
				this.Selection.EndPos   = this.CurPos.ContentPos;

				var Item = this.Content[this.CurPos.ContentPos];
				if (false === Item.MoveCursorDown(true))
				{
					if (this.Content.length - 1 != this.CurPos.ContentPos)
					{
						var TempXY        = Item.GetCurPosXY();
						this.CurPos.RealX = TempXY.X;
						this.CurPos.RealY = TempXY.Y;

						this.CurPos.ContentPos++;
						Item = this.Content[this.CurPos.ContentPos];
						Item.MoveCursorDownToFirstRow(this.CurPos.RealX, this.CurPos.RealY, true);
						this.Selection.EndPos = this.CurPos.ContentPos;
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
					this.Selection.Use = false;

				this.CurPos.ContentPos = this.Selection.EndPos;
			}
			else
			{
				var Item = this.Content[this.CurPos.ContentPos];

				if (false === Item.MoveCursorDown(AddToSelect))
				{
					if (this.Content.length - 1 != this.CurPos.ContentPos)
					{
						var TempXY        = Item.GetCurPosXY();
						this.CurPos.RealX = TempXY.X;
						this.CurPos.RealY = TempXY.Y;

						this.CurPos.ContentPos++;
						Item = this.Content[this.CurPos.ContentPos];
						Item.MoveCursorDownToFirstRow(this.CurPos.RealX, this.CurPos.RealY, false);
					}
					else
					{
						// Сообщаем родительскому классу, что надо выйти из данного элемента
						ReturnValue = false;
					}
				}
			}
		}

		return ReturnValue;
	}
};
CDocumentContent.prototype.MoveCursorToEndOfLine = function(AddToSelect)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
		return this.LogicDocument.DrawingObjects.cursorMoveEndOfLine(AddToSelect);
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		this.RemoveNumberingSelection();
		if (true === this.Selection.Use)
		{
			if (true === AddToSelect)
			{
				var Item = this.Content[this.Selection.EndPos];
				Item.MoveCursorToEndOfLine(AddToSelect);

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
				{
					this.Selection.Use     = false;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}
			}
			else
			{
				var Pos                = ( this.Selection.EndPos >= this.Selection.StartPos ? this.Selection.EndPos : this.Selection.StartPos );
				this.CurPos.ContentPos = Pos;

				var Item = this.Content[Pos];
				Item.MoveCursorToEndOfLine(AddToSelect);

				this.RemoveSelection();
			}
		}
		else
		{
			if (true === AddToSelect)
			{
				this.Selection.Use      = true;
				this.Selection.StartPos = this.CurPos.ContentPos;
				this.Selection.EndPos   = this.CurPos.ContentPos;

				var Item = this.Content[this.CurPos.ContentPos];
				Item.MoveCursorToEndOfLine(AddToSelect);

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
				{
					this.Selection.Use     = false;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}
			}
			else
			{
				var Item = this.Content[this.CurPos.ContentPos];
				Item.MoveCursorToEndOfLine(AddToSelect);
			}
		}
	}
};
CDocumentContent.prototype.MoveCursorToStartOfLine = function(AddToSelect)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
		return this.LogicDocument.DrawingObjects.cursorMoveStartOfLine(AddToSelect);
	else // if( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		this.RemoveNumberingSelection();
		if (true === this.Selection.Use)
		{
			if (true === AddToSelect)
			{
				var Item = this.Content[this.Selection.EndPos];
				Item.MoveCursorToStartOfLine(AddToSelect);

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
				{
					this.Selection.Use     = false;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}
			}
			else
			{
				var Pos                = ( this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos );
				this.CurPos.ContentPos = Pos;

				var Item = this.Content[Pos];
				Item.MoveCursorToStartOfLine(AddToSelect);

				this.RemoveSelection();
			}
		}
		else
		{
			if (true === AddToSelect)
			{
				this.Selection.Use      = true;
				this.Selection.StartPos = this.CurPos.ContentPos;
				this.Selection.EndPos   = this.CurPos.ContentPos;

				var Item = this.Content[this.CurPos.ContentPos];
				Item.MoveCursorToStartOfLine(AddToSelect);

				// Проверяем не обнулился ли селект (т.е. ничего не заселекчено)
				if (this.Selection.StartPos == this.Selection.EndPos && false === this.Content[this.Selection.StartPos].IsSelectionUse())
				{
					this.Selection.Use     = false;
					this.CurPos.ContentPos = this.Selection.EndPos;
				}
			}
			else
			{
				var Item = this.Content[this.CurPos.ContentPos];
				Item.MoveCursorToStartOfLine(AddToSelect);
			}
		}
	}
};
CDocumentContent.prototype.MoveCursorToXY = function(X, Y, AddToSelect, bRemoveOldSelection, CurPage)
{
	if (!this.IsRecalculated())
		return;
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;
	
	if (undefined !== CurPage)
	{
		if (CurPage < 0)
		{
			CurPage = 0;
			Y       = 0;
		}
		else if (CurPage >= this.Pages.length)
		{
			CurPage = this.Pages.length - 1;
			Y       = this.Pages[CurPage].YLimit;
		}

		this.CurPage = CurPage;
	}

	if (false != bRemoveOldSelection)
	{
		this.RemoveNumberingSelection();
	}

	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			var oMouseEvent  = new AscCommon.CMouseEventHandler();
			oMouseEvent.Type = AscCommon.g_mouse_event_type_up;
			this.Selection_SetEnd(X, Y, this.CurPage, oMouseEvent);
		}
		else
		{
			this.RemoveSelection();

			var ContentPos         = this.Internal_GetContentPosByXY(X, Y);
			this.CurPos.ContentPos = ContentPos;
			var ElementPageIndex   = this.private_GetElementPageIndexByXY(ContentPos, X, Y, this.CurPage);
			this.Content[ContentPos].MoveCursorToXY(X, Y, false, false, ElementPageIndex);
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			this.StartSelectionFromCurPos();
			var oMouseEvent  = new AscCommon.CMouseEventHandler();
			oMouseEvent.Type = AscCommon.g_mouse_event_type_up;
			this.Selection_SetEnd(X, Y, this.CurPage, oMouseEvent);
		}
		else
		{
			var ContentPos         = this.Internal_GetContentPosByXY(X, Y);
			this.CurPos.ContentPos = ContentPos;
			var ElementPageIndex   = this.private_GetElementPageIndexByXY(ContentPos, X, Y, this.CurPage);
			this.Content[ContentPos].MoveCursorToXY(X, Y, false, false, ElementPageIndex);
		}
	}
};
CDocumentContent.prototype.IsCursorAtBegin = function(bOnlyPara)
{
	if (undefined === bOnlyPara)
		bOnlyPara = false;

	if (true === bOnlyPara && true != this.Is_CurrentElementParagraph())
		return false;

	if (docpostype_DrawingObjects === this.CurPos.Type)
		return false;
	else if (false != this.Selection.Use || 0 != this.CurPos.ContentPos)
		return false;

	var Item = this.Content[0];
	return Item.IsCursorAtBegin();
};
CDocumentContent.prototype.IsCursorAtEnd = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
		return false;
	else if (false != this.Selection.Use || this.Content.length - 1 != this.CurPos.ContentPos)
		return false;

	var Item = this.Content[this.Content.length - 1];
	return Item.IsCursorAtEnd();
};
CDocumentContent.prototype.GetCurPosXY = function()
{
	return {X : this.CurPos.RealX, Y : this.CurPos.RealY};
};
CDocumentContent.prototype.SetCurPosXY = function(X, Y)
{
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;
	
	this.CurPos.RealX = X;
	this.CurPos.RealY = Y;
};
CDocumentContent.prototype.IsSelectionUse = function()
{
	if (docpostype_DrawingObjects === this.GetDocPosType())
		return this.DrawingObjects.isSelectionUse();

	if (true === this.Selection.Use)
		return true;

	return false;
};
CDocumentContent.prototype.IsSelectionToEnd = function()
{
	if (true !== this.Selection.Use)
		return false;

	if ((this.Selection.StartPos === this.Content.length - 1 || this.Selection.EndPos === this.Content.length - 1) && true === this.Content[this.Content.length - 1].IsSelectionToEnd())
		return true;

	return false;
};
CDocumentContent.prototype.IsTextSelectionUse = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
		return this.LogicDocument.DrawingObjects.isTextSelectionUse();

	return this.IsSelectionUse();
};
/**
 * Возвращаем выделенный текст, если в выделении не более 1 параграфа, и там нет картинок, нумерации страниц и т.д.
 * @param bClearText
 * @param oPr
 * @returns {?string}
 */
CDocumentContent.prototype.GetSelectedText = function(bClearText, oPr)
{
	if (true === this.ApplyToAll)
	{
		if (true === bClearText && this.Content.length <= 1)
		{
			this.Content[0].SetApplyToAll(true);
			var ResultText = this.Content[0].GetSelectedText(true, oPr);
			this.Content[0].SetApplyToAll(false);
			return ResultText;
		}
		else if (true != bClearText)
		{
			var ResultText = "";
			var Count      = this.Content.length;
			for (var Index = 0; Index < Count; Index++)
			{
				this.Content[Index].SetApplyToAll(true);
				ResultText += this.Content[Index].GetSelectedText(false, oPr);
				this.Content[Index].SetApplyToAll(false);
			}

			return ResultText;
		}
	}
	else
	{
		if (docpostype_DrawingObjects === this.CurPos.Type)
			return this.LogicDocument.DrawingObjects.GetSelectedText(bClearText, oPr);

		// Либо у нас нет выделения, либо выделение внутри одного элемента
		if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && selectionflag_Common === this.Selection.Flag ) || false === this.Selection.Use ))
		{
			if (true === bClearText && (this.Selection.StartPos === this.Selection.EndPos || false === this.Selection.Use ))
			{
				var Pos = ( true == this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
				return this.Content[Pos].GetSelectedText(true, oPr);
			}
			else if (false === bClearText)
			{
				var StartPos = ( true == this.Selection.Use ? Math.min(this.Selection.StartPos, this.Selection.EndPos) : this.CurPos.ContentPos );
				var EndPos   = ( true == this.Selection.Use ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : this.CurPos.ContentPos );

				var ResultText = "";

				for (var Index = StartPos; Index <= EndPos; Index++)
				{
					ResultText += this.Content[Index].GetSelectedText(false, oPr);
				}

				return ResultText;
			}
		}
	}

	return null;
};
CDocumentContent.prototype.GetSelectedElementsInfo = function(oInfo)
{
	if (!oInfo)
		oInfo = new CSelectedElementsInfo();

	if (true === this.ApplyToAll)
	{
		var nCount = this.Content.length;
		if (nCount > 1)
			oInfo.SetMixedSelection();

		if (oInfo.IsCheckAllSelection() || nCount === 1)
		{
			for (var nPos = 0; nPos < nCount; ++nPos)
			{
				var oElement = this.Content[nPos];
				oElement.SetApplyToAll(true);
				oElement.GetSelectedElementsInfo(oInfo);
				oElement.SetApplyToAll(false);
			}
		}
	}
	else
	{
		if (docpostype_DrawingObjects === this.CurPos.Type)
		{
			let logicDocument = this.GetLogicDocument();
			if (logicDocument)
			{
				if (logicDocument.IsDocumentEditor())
					logicDocument.DrawingsController.GetSelectedElementsInfo(oInfo);
				else if (logicDocument.DrawingObjects)
					logicDocument.DrawingObjects.getSelectedElementsInfo(oInfo);
			}
		}
		else //if ( docpostype_Content == this.CurPos.Type )
		{
			if (selectionflag_Numbering === this.Selection.Flag)
			{
				if (this.Selection.Data && this.Selection.Data.CurPara)
				{
					this.Selection.Data.CurPara.GetSelectedElementsInfo(oInfo);
				}
			}
			else
			{
				if (true === this.Selection.Use)
				{
					if (this.Selection.StartPos != this.Selection.EndPos)
						oInfo.SetMixedSelection();

					if (oInfo.IsCheckAllSelection() || this.Selection.StartPos === this.Selection.EndPos)
					{
						var nStartPos = this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos;
						var nEndPos   = this.Selection.StartPos < this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos;

						for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
						{
							this.Content[nPos].GetSelectedElementsInfo(oInfo);
						}
					}
				}
				else
				{
					this.Content[this.CurPos.ContentPos].GetSelectedElementsInfo(oInfo);
				}
			}
		}
	}

	return oInfo;
};
CDocumentContent.prototype.GetSelectedContent = function(oSelectedContent)
{
	let oSContent = oSelectedContent ? oSelectedContent : new AscCommonWord.CSelectedContent();

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.DrawingObjects.GetSelectedContent(oSContent);
	}
	else if (this.Selection.Use && selectionflag_Common === this.Selection.Flag)
	{
		let nStartPos = this.Selection.StartPos;
		let nEndPos   = this.Selection.EndPos;
		if (nStartPos > nEndPos)
		{
			nStartPos = this.Selection.EndPos;
			nEndPos   = this.Selection.StartPos;
		}

		for (let nPos = nStartPos; nPos <= nEndPos; ++nPos)
		{
			this.Content[nPos].GetSelectedContent(oSContent);
		}
	}

	return oSContent;
};
CDocumentContent.prototype.SetParagraphPr = function(oParaPr)
{
	if (this.IsApplyToAll())
	{
		for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			var oItem = this.Content[nIndex];
			oItem.SetApplyToAll(true);
			oItem.SetParagraphPr(oParaPr);
			oItem.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.GetDocPosType())
	{
		// TODO: Прокинуть в Drawings
	}
	else
	{
		if (this.Selection.Use)
		{
			var nStartPos = this.Selection.StartPos;
			var nEndPos   = this.Selection.EndPos;
			if (nEndPos < nStartPos)
			{
				var nTemp = nStartPos;
				nStartPos = nEndPos;
				nEndPos   = nTemp;
			}

			for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
			{
				this.Content[nIndex].SetParagraphPr(oParaPr);
			}
		}
		else
		{
			this.Content[this.CurPos.ContentPos].SetParagraphPr(oParaPr);
		}
	}
};
CDocumentContent.prototype.SetParagraphBidi = function(bidi)
{
	if (this.IsApplyToAll())
	{
		for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			var oItem = this.Content[nIndex];
			oItem.SetApplyToAll(true);
			oItem.SetParagraphBidi(bidi);
			oItem.SetApplyToAll(false);
		}
		
		return;
	}
	
	if (docpostype_DrawingObjects === this.GetDocPosType())
	{
		//return this.LogicDocument.DrawingObjects.setParagraphBidi(bidi);
	}
	else
	{
		if (this.Selection.Use)
		{
			var nStartPos = this.Selection.StartPos;
			var nEndPos   = this.Selection.EndPos;
			if (nEndPos < nStartPos)
			{
				var nTemp = nStartPos;
				nStartPos = nEndPos;
				nEndPos   = nTemp;
			}
			
			for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
			{
				this.Content[nIndex].SetParagraphBidi(bidi);
			}
		}
		else
		{
			this.Content[this.CurPos.ContentPos].SetParagraphBidi(bidi);
		}
	}
};
CDocumentContent.prototype.SetParagraphAlign = function(Align)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphAlign(Align);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphAlign(Align);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphAlign(Align);
			}
		}
		else
		{
			this.Content[this.CurPos.ContentPos].SetParagraphAlign(Align);
		}
	}
};
CDocumentContent.prototype.SetParagraphDefaultTabSize = function(TabSize)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphDefaultTabSize(TabSize);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphDefaultTabSize(TabSize);
		}
	}
	else
	{
		this.Content[this.CurPos.ContentPos].SetParagraphDefaultTabSize(TabSize);
	}
};
CDocumentContent.prototype.SetParagraphSpacing = function(Spacing)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphSpacing(Spacing);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphSpacing(Spacing);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphSpacing(Spacing);
			}
		}
		else
		{
			this.Content[this.CurPos.ContentPos].SetParagraphSpacing(Spacing);
		}
	}
};
CDocumentContent.prototype.SetParagraphIndent = function(Ind)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphIndent(Ind);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphIndent(Ind);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphIndent(Ind);
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.SetParagraphIndent(Ind);
		}
	}
};
CDocumentContent.prototype.Set_ParagraphPresentationNumbering = function(Bullet)
{
	var Index, StartPos, EndPos;
	if (true === this.ApplyToAll)
    {
		StartPos = 0;
		EndPos = this.Content.length - 1;
    }
    else
	{
		if (this.CurPos.ContentPos < 0)
			return;

		if (true === this.Selection.Use)
		{
			StartPos = this.Selection.StartPos;
			EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}
		}
		else
		{
			StartPos = this.CurPos.ContentPos;
			EndPos = this.CurPos.ContentPos;
		}
	}

	var oLastParagraph = this.Content[EndPos];
	var oLastBullet;
	if(oLastParagraph)
	{
		oLastBullet = oLastParagraph.Get_PresentationNumbering();
	}
	for (Index = StartPos; Index <= EndPos; Index++)
	{
		this.Content[Index].Add_PresentationNumbering(Bullet);
	}
	if(oLastParagraph)
	{
		if(oLastBullet)
		{
			var Level = AscFormat.isRealNumber(oLastParagraph.Pr.Lvl) ? oLastParagraph.Pr.Lvl : 0;

			var oLastBullet_ = oLastParagraph.Get_PresentationNumbering();
			if (oLastBullet_.IsNumbered())
			{
				var Next = oLastParagraph.Next;
				while (null != Next && type_Paragraph === Next.GetType())
				{
					var NextLevel = Next.PresentationPr.Level;
					var NextBullet = Next.Get_PresentationNumbering();
					if (Level < NextLevel)
					{
						Next = Next.Next;
						continue;
					}
					else if (Level > NextLevel)
						break;
					else if (NextBullet.Get_Type() === oLastBullet.Get_Type() && NextBullet.Get_StartAt() === oLastBullet.Get_StartAt() )
					{
						var nNumStartAt = Bullet.getNumStartAt();
						if(AscFormat.isRealNumber(nNumStartAt))
						{
							var oPrBullet = new AscFormat.CBullet();
							oPrBullet.putNumStartAt(nNumStartAt);
							Next.Add_PresentationNumbering(oPrBullet);
						}
						Next = Next.Next;
					}
					else
					{
						break;
					}
				}
			}
		}
	}
};
CDocumentContent.prototype.Can_IncreaseParagraphLevel         = function(bIncrease)
{
    if (true === this.ApplyToAll)
    {
        for (var Index = 0; Index < this.Content.length; Index++)
        {
            if (!this.Content[Index].Can_IncreaseLevel(bIncrease))
                return false;
        }
        return true;
    }

    if (this.CurPos.ContentPos < 0)
        return false;

    if (true === this.Selection.Use)
    {
        var StartPos = this.Selection.StartPos;
        var EndPos   = this.Selection.EndPos;
        if (EndPos < StartPos)
        {
            var Temp = StartPos;
            StartPos = EndPos;
            EndPos   = Temp;
        }

        for (var Index = StartPos; Index <= EndPos; Index++)
        {
            if (!this.Content[Index].Can_IncreaseLevel(bIncrease))
            {
                return false;
            }
        }

        return true;
    }
    return this.Content[this.CurPos.ContentPos].Can_IncreaseLevel(bIncrease);
};
CDocumentContent.prototype.Increase_ParagraphLevel            = function(bIncrease)
{
    if (!this.Can_IncreaseParagraphLevel(bIncrease))
        return;
    if (true === this.ApplyToAll)
    {
        for (var Index = 0; Index < this.Content.length; Index++)
        {
            this.Content[Index].Increase_Level(bIncrease);
        }
        return;
    }

    if (this.CurPos.ContentPos < 0)
        return false;

    if (true === this.Selection.Use)
    {
        var StartPos = this.Selection.StartPos;
        var EndPos   = this.Selection.EndPos;
        if (EndPos < StartPos)
        {
            var Temp = StartPos;
            StartPos = EndPos;
            EndPos   = Temp;
        }

        for (var Index = StartPos; Index <= EndPos; Index++)
        {
            this.Content[Index].Increase_Level(bIncrease);
        }
        this.Recalculate();
        return;
    }
    this.Content[this.CurPos.ContentPos].Increase_Level(bIncrease);
    this.Recalculate();
};
CDocumentContent.prototype.SetParagraphShd = function(Shd)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			// При изменении цвета фона параграфа, не надо ничего пересчитывать
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphShd(Shd);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphShd(Shd);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;

			if (this.LogicDocument && true === this.LogicDocument.UseTextShd && StartPos === EndPos && type_Paragraph === this.Content[StartPos].GetType() && false === this.Content[StartPos].Selection_CheckParaEnd() && selectionflag_Common === this.Selection.Flag)
			{
				this.AddToParagraph(new ParaTextPr({Shd : Shd}));
			}
			else
			{
				if (EndPos < StartPos)
				{
					var Temp = StartPos;
					StartPos = EndPos;
					EndPos   = Temp;
				}

				for (var Index = StartPos; Index <= EndPos; Index++)
				{
					var Item = this.Content[Index];
					Item.SetParagraphShd(Shd);
				}
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.SetParagraphShd(Shd);
		}
	}
};
CDocumentContent.prototype.SetParagraphStyle = function(Name)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphStyle(Name);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphStyle(Name);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			this.RemoveNumberingSelection();

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphStyle(Name);
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.SetParagraphStyle(Name);
		}
	}
};
CDocumentContent.prototype.SetParagraphTabs = function(Tabs)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphTabs(Tabs);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphTabs(Tabs);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphTabs(Tabs);
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.SetParagraphTabs(Tabs);
		}
	}
};
CDocumentContent.prototype.SetParagraphContextualSpacing = function(Value)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphContextualSpacing(Value);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphContextualSpacing(Value);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphContextualSpacing(Value);
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.SetParagraphContextualSpacing(Value);
		}
	}
};
CDocumentContent.prototype.SetParagraphPageBreakBefore = function(Value)
{
	// В таблице или вне самого верхнего документа нет смысла ставить PageBreak
	if (docpostype_Content !== this.GetDocPosType() || this.IsInTable() || this.GetTopDocumentContent() !== this.LogicDocument)
		return;

	if (this.CurPos.ContentPos < 0)
		return false;

	if (true === this.Selection.Use)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (EndPos < StartPos)
		{
			var Temp = StartPos;
			StartPos = EndPos;
			EndPos   = Temp;
		}

		for (var Index = StartPos; Index <= EndPos; Index++)
		{
			var Item = this.Content[Index];
			Item.SetParagraphPageBreakBefore(Value);
		}
	}
	else
	{
		var Item = this.Content[this.CurPos.ContentPos];
		Item.SetParagraphPageBreakBefore(Value);
	}
};
CDocumentContent.prototype.SetParagraphKeepLines = function(Value)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphKeepLines(Value);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphKeepLines(Value);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphKeepLines(Value);
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.SetParagraphKeepLines(Value);
		}
	}
};
CDocumentContent.prototype.SetParagraphKeepNext = function(Value)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphKeepNext(Value);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphKeepNext(Value);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphKeepNext(Value);
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.SetParagraphKeepNext(Value);
		}
	}
};
CDocumentContent.prototype.SetParagraphWidowControl = function(Value)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphWidowControl(Value);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphWidowControl(Value);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphWidowControl(Value);
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.SetParagraphWidowControl(Value);
		}
	}
};
CDocumentContent.prototype.SetParagraphBorders = function(Borders)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.SetParagraphBorders(Borders);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setParagraphBorders(Borders);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				var Item = this.Content[Index];
				Item.SetParagraphBorders(Borders);
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			if (type_Paragraph === Item.GetType())
			{
				// Мы должны выставить границу для всех параграфов, входящих в текущую группу параграфов
				// с одинаковыми границами

				var StartPos = Item.Index;
				var EndPos   = Item.Index;
				var CurBrd   = Item.Get_CompiledPr().ParaPr.Brd;

				while (true != CurBrd.First)
				{
					StartPos--;
					if (StartPos < 0)
					{
						StartPos = 0;
						break;
					}

					var TempItem = this.Content[StartPos];
					if (type_Paragraph !== TempItem.GetType())
					{
						StartPos++;
						break;
					}

					CurBrd = TempItem.Get_CompiledPr().ParaPr.Brd;
				}

				CurBrd = Item.Get_CompiledPr().ParaPr.Brd;
				while (true != CurBrd.Last)
				{
					EndPos++;
					if (EndPos >= this.Content.length)
					{
						EndPos = this.Content.length - 1;
						break;
					}

					var TempItem = this.Content[EndPos];
					if (type_Paragraph !== TempItem.GetType())
					{
						EndPos--;
						break;
					}

					CurBrd = TempItem.Get_CompiledPr().ParaPr.Brd;
				}

				for (var Index = StartPos; Index <= EndPos; Index++)
					this.Content[Index].SetParagraphBorders(Borders);
			}
			else
			{
				Item.SetParagraphBorders(Borders);
			}
		}
	}
};
CDocumentContent.prototype.IncreaseDecreaseFontSize = function(bIncrease)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.IncreaseDecreaseFontSize(bIncrease);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.paragraphIncDecFontSize(bIncrease);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (this.CurPos.ContentPos < 0)
			return false;

		if (true === this.Selection.Use)
		{
			switch (this.Selection.Flag)
			{
				case selectionflag_Common:
				{
					var StartPos = this.Selection.StartPos;
					var EndPos   = this.Selection.EndPos;
					if (EndPos < StartPos)
					{
						var Temp = StartPos;
						StartPos = EndPos;
						EndPos   = Temp;
					}

					for (var Index = StartPos; Index <= EndPos; Index++)
					{
						var Item = this.Content[Index];
						Item.IncreaseDecreaseFontSize(bIncrease);
					}
					break;
				}
				case  selectionflag_Numbering:
				{
					var NewFontSize = this.GetCalculatedTextPr().GetIncDecFontSize(bIncrease);
					var TextPr      = new CTextPr();
					TextPr.FontSize = NewFontSize;
					this.AddToParagraph(new ParaTextPr(TextPr), true);
					break;
				}
			}
		}
		else
		{
			var Item = this.Content[this.CurPos.ContentPos];
			Item.IncreaseDecreaseFontSize(bIncrease);
		}
	}
};
CDocumentContent.prototype.IncreaseDecreaseIndent = function(bIncrease)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.IncreaseDecreaseIndent(bIncrease);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		if (true != this.LogicDocument.DrawingObjects.isSelectedText())
		{
			var ParaDrawing = this.LogicDocument.DrawingObjects.getMajorParaDrawing();
			if (null != ParaDrawing)
			{
				var Paragraph = ParaDrawing.Parent;
				Paragraph.IncreaseDecreaseIndent(bIncrease);
			}
		}
		else
		{
			this.DrawingObjects.paragraphIncDecIndent(bIncrease);
		}
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use && selectionflag_Common === this.Selection.Flag)
		{
			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			for (var Index = StartPos; Index <= EndPos; Index++)
			{
				this.Content[Index].IncreaseDecreaseIndent(bIncrease);
			}
		}
		else
		{
			this.Content[this.CurPos.ContentPos].IncreaseDecreaseIndent(bIncrease);
		}
	}
};
CDocumentContent.prototype.PasteFormatting = function(oData)
{
	if (true === this.ApplyToAll)
	{
		for (var Index = 0; Index < this.Content.length; Index++)
		{
			var Item = this.Content[Index];
			Item.SetApplyToAll(true);
			Item.PasteFormatting(oData);
			Item.SetApplyToAll(false);
		}

		return;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.pasteFormatting(oData);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			switch (this.Selection.Flag)
			{
				case selectionflag_Numbering    :
					return;
				case selectionflag_Common:
				{
					var Start = this.Selection.StartPos;
					var End   = this.Selection.EndPos;
					if (Start > End)
					{
						Start = this.Selection.EndPos;
						End   = this.Selection.StartPos;
					}

					for (var Pos = Start; Pos <= End; Pos++)
					{
						this.Content[Pos].PasteFormatting(oData);
					}
					break;
				}
			}
		}
		else
		{
			this.Content[this.CurPos.ContentPos].PasteFormatting(oData);
		}
	}
};
CDocumentContent.prototype.SetImageProps = function(Props)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.LogicDocument.DrawingObjects.setProps(Props);
		this.Document_UpdateInterfaceState();
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Table == this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Table == this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		if (true == this.Selection.Use)
			this.Content[this.Selection.StartPos].SetImageProps(Props);
		else
			this.Content[this.CurPos.ContentPos].SetImageProps(Props);
	}
};
CDocumentContent.prototype.SetTableProps = function(Props)
{
	if (true === this.ApplyToAll)
		return false;

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.setTableProps(Props);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		var Pos = -1;
		if (true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos)
			Pos = this.Selection.StartPos;
		else if (false === this.Selection.Use)
			Pos = this.CurPos.ContentPos;

		if (-1 !== Pos)
			return this.Content[Pos].SetTableProps(Props);

		return false;
	}
};
CDocumentContent.prototype.GetTableProps = function()
{
	var TablePr = null;
	if (docpostype_Content == this.GetDocPosType() && ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos) || false == this.Selection.Use))
	{
		if (true == this.Selection.Use)
			TablePr = this.Content[this.Selection.StartPos].GetTableProps();
		else
			TablePr = this.Content[this.CurPos.ContentPos].GetTableProps();
	}

	if (null !== TablePr)
		TablePr.CanBeFlow = true;

	return TablePr;
};
CDocumentContent.prototype.GetCalculatedParaPr = function()
{
	var Result_ParaPr = new CParaPr();
	var FirstTextPr, FirstTextPrTmp, oBullet;

	if (true === this.ApplyToAll)
	{
		var StartPr = this.Content[0].GetCalculatedParaPr();
		var Pr      = StartPr.Copy();
		Pr.Locked   = StartPr.Locked;
		if(this.bPresentation)
		{
			if(this.Content[0].GetType() === type_Paragraph)
			{
				FirstTextPr = this.Content[0].Get_FirstTextPr2();
			}
		}
		for (var Index = 1; Index < this.Content.length; Index++)
		{
			var TempPr = this.Content[Index].GetCalculatedParaPr();
			Pr         = Pr.Compare(TempPr);
			if(this.bPresentation)
			{
				if(this.Content[Index].GetType() === type_Paragraph)
				{
					FirstTextPrTmp = this.Content[Index].Get_FirstTextPr2();
					if(!FirstTextPr)
					{
						FirstTextPr = FirstTextPrTmp;
					}
					else
					{
						FirstTextPr = FirstTextPr.Compare(FirstTextPrTmp);
					}
				}
			}
		}

		if (Pr.Ind.Left == UnknownValue)
			Pr.Ind.Left = StartPr.Ind.Left;

		if (Pr.Ind.Right == UnknownValue)
			Pr.Ind.Right = StartPr.Ind.Right;

		if (Pr.Ind.FirstLine == UnknownValue)
			Pr.Ind.FirstLine = StartPr.Ind.FirstLine;

		Result_ParaPr             = Pr;
		Result_ParaPr.CanAddTable = ( true === Pr.Locked ? false : true ) && !(this.bPresentation === true);
		if (Result_ParaPr.Shd && Result_ParaPr.Shd.Unifill)
		{
			Result_ParaPr.Shd.Unifill.check(this.Get_Theme(), this.Get_ColorMap());
		}
		if(this.bPresentation && Result_ParaPr.Bullet)
		{
			oBullet = Result_ParaPr.Bullet;
			if(oBullet)
			{
				if(oBullet.bulletColor &&  oBullet.bulletColor.UniColor)
				{
					oBullet.bulletColor.UniColor.check(this.Get_Theme(), this.Get_ColorMap());
				}
				if(oBullet.bulletType.startAt !== undefined)
				{
					oBullet.bulletType.startAt = this.Content[0].GetBulletNum();
				}
				if(!AscFormat.isRealNumber(Result_ParaPr.Lvl))
				{
					oBullet.bulletType.startAt = undefined;
				}
			}
		}

		return Result_ParaPr;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.getParagraphParaPr();
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		var StartPos, EndPos;
		if (true === this.Selection.Use && selectionflag_Common === this.Selection.Flag)
		{
			StartPos = this.Selection.StartPos;
			EndPos   = this.Selection.EndPos;
			if (EndPos < StartPos)
			{
				var Temp = StartPos;
				StartPos = EndPos;
				EndPos   = Temp;
			}

			var StartPr = this.Content[StartPos].GetCalculatedParaPr();
			var Pr      = StartPr.Copy();
			Pr.Locked   = StartPr.Locked;

			if(this.bPresentation)
			{
				if(this.Content[StartPos].GetType() === type_Paragraph)
				{
					FirstTextPr = this.Content[StartPos].Get_FirstTextPr2();
				}
			}
			for (var Index = StartPos + 1; Index <= EndPos; Index++)
			{
				var TempPr = this.Content[Index].GetCalculatedParaPr();
				Pr         = Pr.Compare(TempPr);
				if(this.bPresentation)
				{
					if(this.Content[Index].GetType() === type_Paragraph)
					{
						FirstTextPrTmp = this.Content[Index].Get_FirstTextPr2();
						if(!FirstTextPr)
						{
							FirstTextPr = FirstTextPrTmp;
						}
						else
						{
							FirstTextPr = FirstTextPr.Compare(FirstTextPrTmp);
						}
					}
				}
			}

			if (undefined === Pr.Ind.Left)
				Pr.Ind.Left = StartPr.Ind.Left;

			if (undefined === Pr.Ind.Right)
				Pr.Ind.Right = StartPr.Ind.Right;

			if (undefined === Pr.Ind.FirstLine)
				Pr.Ind.FirstLine = StartPr.Ind.FirstLine;

			Result_ParaPr             = Pr;
			Result_ParaPr.CanAddTable = (true !== Pr.Locked) && !(this.bPresentation === true);
		}
		else
		{
			StartPos = this.CurPos.ContentPos;
			var Item = this.Content[this.CurPos.ContentPos];
			if (type_Paragraph == Item.GetType())
			{
				Result_ParaPr             = Item.GetCalculatedParaPr().Copy();
				Result_ParaPr.CanAddTable = (true === Result_ParaPr.Locked ? Item.IsCursorAtEnd() : true) && !(this.bPresentation === true);
				if(this.bPresentation)
				{
					FirstTextPr = Item.Get_FirstTextPr2();
				}
			}
			else
			{
				Result_ParaPr = Item.GetCalculatedParaPr();
			}
		}
		if (Result_ParaPr.Shd && Result_ParaPr.Shd.Unifill)
		{
			Result_ParaPr.Shd.Unifill.check(this.Get_Theme(), this.Get_ColorMap());
		}
		if(this.bPresentation && Result_ParaPr.Bullet)
		{
			oBullet = Result_ParaPr.Bullet;
			var oTheme = this.Get_Theme();
			var oColorMap = this.Get_ColorMap();
			if(oBullet)
			{
				if(oBullet.bulletColor &&  oBullet.bulletColor.UniColor)
				{
					oBullet.bulletColor.UniColor.check(oTheme, oColorMap);
				}
				if(oBullet.bulletType.startAt !== undefined)
				{
					oBullet.bulletType.startAt = this.Content[StartPos].GetBulletNum();
				}
				if(!AscFormat.isRealNumber(Result_ParaPr.Lvl))
				{
					oBullet.bulletType.startAt = undefined;
				}
			}
			if(FirstTextPr)
			{
				Result_ParaPr.FirstTextPr = FirstTextPr;

				if(FirstTextPr.Unifill)
				{
					FirstTextPr.Unifill.check(oTheme, oColorMap);
				}
				FirstTextPr.ReplaceThemeFonts(oTheme.themeElements.fontScheme);
			}
		}
		return Result_ParaPr;
	}
};
CDocumentContent.prototype.GetCalculatedTextPr = function(skipFontCalculator)
{
	let textPr = null;
	if (true === this.ApplyToAll)
	{
		var VisTextPr;
		this.Content[0].SetApplyToAll(true);
		VisTextPr = this.Content[0].GetCalculatedTextPr();
		this.Content[0].SetApplyToAll(false);

		var Count = this.Content.length;
		for (var Index = 1; Index < Count; Index++)
		{
			this.Content[Index].SetApplyToAll(true);
			var CurPr = this.Content[Index].GetCalculatedTextPr();
			VisTextPr = VisTextPr.Compare(CurPr);
			this.Content[Index].SetApplyToAll(false);
		}
		
		textPr = VisTextPr;
	}
	else if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		textPr = this.LogicDocument.DrawingObjects.getParagraphTextPr();
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			var VisTextPr;
			switch (this.Selection.Flag)
			{
				case selectionflag_Common:
				{
					var StartPos = this.Selection.StartPos;
					var EndPos   = this.Selection.EndPos;
					if (EndPos < StartPos)
					{
						var Temp = StartPos;
						StartPos = EndPos;
						EndPos   = Temp;
					}

					VisTextPr = this.Content[StartPos].GetCalculatedTextPr();

					for (var Index = StartPos + 1; Index <= EndPos; Index++)
					{
						var CurPr = this.Content[Index].GetCalculatedTextPr();
						VisTextPr = VisTextPr.Compare(CurPr);
					}

					break;
				}
				case selectionflag_Numbering:
				{
					if (!this.Selection.Data || !this.Selection.Data.CurPara)
						break;

					VisTextPr = this.Selection.Data.CurPara.GetNumberingTextPr();
					break;
				}
			}
			
			textPr = VisTextPr;
		}
		else
		{
			textPr = this.Content[this.CurPos.ContentPos].GetCalculatedTextPr();
		}
	}
	
	if (true !== skipFontCalculator && textPr)
	{
		AscWord.FontCalculator.Calculate(this, textPr);
		
		let theme = this.GetTheme();
		if (theme)
			textPr.ReplaceThemeFonts(theme.themeElements.fontScheme);
	}
	
	return textPr;
};
CDocumentContent.prototype.GetDirectTextPr = function()
{
	var Result_TextPr = null;

	if (true === this.ApplyToAll)
	{
		var Item      = this.Content[0];
		Result_TextPr = Item.GetDirectTextPr();
		return Result_TextPr;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.getParagraphTextPrCopy();
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			var VisTextPr;
			switch (this.Selection.Flag)
			{
				case selectionflag_Common:
				{
					var StartPos = this.Selection.StartPos;
					if (this.Selection.EndPos < StartPos)
						StartPos = this.Selection.EndPos;

					var Item  = this.Content[StartPos];
					VisTextPr = Item.GetDirectTextPr();

					break;
				}
				case selectionflag_Numbering:
				{
					if (!this.Selection.Data || !this.Selection.Data.CurPara)
						break;

					var oNumPr = this.Selection.Data.CurPara.GetNumPr();
					VisTextPr  = this.GetNumbering().GetNum(oNumPr.NumId).GetLvl(oNumPr.Lvl).GetTextPr();

					break;
				}
			}

			Result_TextPr = VisTextPr;
		}
		else
		{
			var Item      = this.Content[this.CurPos.ContentPos];
			Result_TextPr = Item.GetDirectTextPr();
		}

		return Result_TextPr;
	}
};
CDocumentContent.prototype.GetDirectParaPr = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.getParagraphParaPrCopy();
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		var Result_ParaPr = null;

		// Inline объекты
		if (docpostype_Content == this.CurPos.Type)
		{
			if (true === this.Selection.Use)
			{
				switch (this.Selection.Flag)
				{
					case selectionflag_Common:
					{
						var StartPos = this.Selection.StartPos;
						if (this.Selection.EndPos < StartPos)
							StartPos = this.Selection.EndPos;

						var Item      = this.Content[StartPos];
						Result_ParaPr = Item.GetDirectParaPr();

						break;
					}
					case selectionflag_Numbering:
					{
						if (!this.Selection.Data || !this.Selection.Data.CurPara)
							break;

						var oNumPr    = this.Selection.Data.CurPara.GetNumPr();
						Result_ParaPr = this.GetNumbering().GetNum(oNumPr.NumId).GetLvl(oNumPr.Lvl).GetParaPr();

						break;
					}
				}
			}
			else
			{
				var Item      = this.Content[this.CurPos.ContentPos];
				Result_ParaPr = Item.GetDirectParaPr();
			}
		}

		return Result_ParaPr;
	}
};
//-----------------------------------------------------------------------------------
// Функции для работы с интерфейсом
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.Interface_Update_DrawingPr = function(Flag)
{
    var ImagePr = {};

    if (docpostype_DrawingObjects === this.CurPos.Type)
        ImagePr = this.LogicDocument.DrawingObjects.getProps();

    if (true === Flag)
        return ImagePr;
    else
        editor.sync_ImgPropCallback(ImagePr);

};
CDocumentContent.prototype.Interface_Update_TablePr   = function(Flag)
{
    var TablePr = null;
    if (docpostype_DrawingObjects == this.CurPos.Type)
    {
        TablePr = this.LogicDocument.DrawingObjects.getTableProps();
    }
    else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Table == this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Table == this.Content[this.CurPos.ContentPos].GetType() ) ))
    {
        if (true == this.Selection.Use)
            TablePr = this.Content[this.Selection.StartPos].Get_Props();
        else
            TablePr = this.Content[this.CurPos.ContentPos].Get_Props();
    }

    if (true === Flag)
        return TablePr;
    else if (null != TablePr)
        editor.sync_TblPropCallback(TablePr);
};
//-----------------------------------------------------------------------------------
// Функции для работы с селектом
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.RemoveSelection = function(bNoCheckDrawing)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		var oParaDrawing = this.LogicDocument.DrawingObjects.getMajorParaDrawing();
		if (oParaDrawing)
			oParaDrawing.GoToText(undefined, false);

		return this.LogicDocument.DrawingObjects.resetSelection(undefined, bNoCheckDrawing);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			switch (this.Selection.Flag)
			{
				case selectionflag_Common:
				{
					var Start = this.Selection.StartPos;
					var End   = this.Selection.EndPos;

					if (Start > End)
					{
						var Temp = Start;
						Start    = End;
						End      = Temp;
					}

					Start = Math.max(0, Start);
					End   = Math.min(this.Content.length - 1, End);

					for (var Index = Start; Index <= End; Index++)
					{
						this.Content[Index].RemoveSelection();
					}
					break;
				}
				case selectionflag_Numbering:
				{
					if (!this.Selection.Data)
						break;

					for (var nIndex = 0, nCount = this.Selection.Data.Paragraphs.length; nIndex < nCount; ++nIndex)
					{
						this.Selection.Data.Paragraphs[nIndex].RemoveSelection();
					}

					break;
				}
			}
		}

		this.Selection.StartPos = 0;
		this.Selection.EndPos   = 0;

		this.Selection.Use   = false;
		this.Selection.Start = false;
		this.Selection.Flag  = selectionflag_Common;
	}
};
CDocumentContent.prototype.DrawSelectionOnPage = function(PageIndex, clipInfo)
{
	if (!this.IsRecalculated())
		return;
	
    var CurPage = PageIndex;
    if (CurPage < 0 || CurPage >= this.Pages.length)
        return;
	
	let drawingDocument = Asc.editor.getDrawingDocument();
	if (this.transform && drawingDocument)
		drawingDocument.MultiplyTargetTransform(this.transform.CreateDublicate());
	
	clipInfo = this.IntersectClip(clipInfo, PageIndex);

    if (docpostype_DrawingObjects === this.CurPos.Type)
    {
        this.DrawingDocument.SetTextSelectionOutline(true);
        var PageAbs = CurPage + this.Get_StartPage_Absolute();
        this.LogicDocument.DrawingObjects.drawSelectionPage(PageAbs, clipInfo);
    }
    else
    {
        var Pos_start = this.Pages[CurPage].Pos;
        var Pos_end   = this.Pages[CurPage].EndPos;

        if (true === this.Selection.Use)
        {
            switch (this.Selection.Flag)
            {
                case selectionflag_Common:
                {
                    var Start = this.Selection.StartPos;
                    var End   = this.Selection.EndPos;

                    if (Start > End)
                    {
                        Start = this.Selection.EndPos;
                        End   = this.Selection.StartPos;
                    }

                    var Start = Math.max(Start, Pos_start);
                    var End   = Math.min(End, Pos_end);

                    for (var Index = Start; Index <= End; Index++)
                    {
                        var ElementPageIndex = this.private_GetElementPageIndex(Index, CurPage, 0, 1);
                        this.Content[Index].DrawSelectionOnPage(ElementPageIndex, clipInfo);
                    }

                    break;
                }
                case selectionflag_Numbering:
                {
                    if (!this.Selection.Data)
                        break;

					var nPageAbs = this.Get_AbsolutePage(CurPage);
                    for (var nIndex = 0, nCount = this.Selection.Data.Paragraphs.length; nIndex < nCount; ++nIndex)
					{
						var oParagraph = this.Selection.Data.Paragraphs[nIndex];
						var nParaPageAbs = oParagraph.GetNumberingPage(true);
						if (nParaPageAbs === nPageAbs)
							oParagraph.DrawSelectionOnPage(oParagraph.GetNumberingPage(false), clipInfo);
					}

                    break;
                }
            }
        }
    }
};
CDocumentContent.prototype.Selection_SetStart = function(X, Y, CurPage, MouseEvent)
{
	if (!this.IsRecalculated())
		return;
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;

	if (CurPage < 0)
	{
		CurPage = 0;
		Y       = 0;
	}
	else if (CurPage >= this.Pages.length)
	{
		CurPage = this.Pages.length - 1;
		Y       = this.Pages[CurPage].YLimit;
	}

	this.CurPage = CurPage;
	var AbsPage  = this.Get_AbsolutePage(this.CurPage);

	var isTopDocumentContent = this === this.GetTopDocumentContent();

	// Сначала проверим, не попали ли мы в один из "плавающих" объектов
	var bInText      = null !== this.IsInText(X, Y, CurPage);
	var bTableBorder = null !== this.IsTableBorder(X, Y, CurPage);
	var nInDrawing   = -1;
	if(this.LogicDocument && this.LogicDocument.IsDocumentEditor())
	{
		nInDrawing = this.LogicDocument && this.LogicDocument.DrawingObjects.IsInDrawingObject(X, Y, AbsPage, this);
	}

	if (this.Parent instanceof CHeaderFooter && ( nInDrawing === DRAWING_ARRAY_TYPE_BEFORE || nInDrawing === DRAWING_ARRAY_TYPE_INLINE || ( false === bTableBorder && false === bInText && nInDrawing >= 0 ) ))
	{
		if (docpostype_DrawingObjects != this.CurPos.Type)
			this.RemoveSelection();

		// Прячем курсор
		this.DrawingDocument.TargetEnd();
		this.DrawingDocument.SetCurrentPage(AbsPage);

		var HdrFtr = this.IsHdrFtr(true);
		if (null === HdrFtr)
		{
			this.LogicDocument.Selection.Use   = true;
			this.LogicDocument.Selection.Start = true;
			this.LogicDocument.Selection.Flag  = selectionflag_Common;
			this.LogicDocument.SetDocPosType(docpostype_DrawingObjects);
		}
		else
		{
			HdrFtr.Content.Selection.Use   = true;
			HdrFtr.Content.Selection.Start = true;
			HdrFtr.Content.Selection.Flag  = selectionflag_Common;
			HdrFtr.Content.SetDocPosType(docpostype_DrawingObjects);
		}

		this.LogicDocument.DrawingObjects.OnMouseDown(MouseEvent, X, Y, AbsPage);
	}
	else
	{
		var bOldSelectionIsCommon = true;

		if (docpostype_DrawingObjects === this.CurPos.Type && true != this.IsInDrawing(X, Y, AbsPage))
		{
			this.LogicDocument.DrawingObjects.resetSelection();
			bOldSelectionIsCommon = false;
		}

		var ContentPos = this.Internal_GetContentPosByXY(X, Y);

		if (docpostype_Content != this.CurPos.Type)
		{
			this.SetDocPosType(docpostype_Content);
			this.CurPos.ContentPos = ContentPos;
			bOldSelectionIsCommon  = false;
		}

		var SelectionUse_old = this.Selection.Use;
		var Item             = this.Content[ContentPos];

		// Убираем селект, кроме случаев либо текущего параграфа, либо при движении границ внутри таблицы
		if (!(true === SelectionUse_old && true === MouseEvent.ShiftKey && true === bOldSelectionIsCommon))
		{
			if ((selectionflag_Common != this.Selection.Flag) || ( true === this.Selection.Use && MouseEvent.ClickCount <= 1 && true != bTableBorder ))
				this.RemoveSelection();
		}

		this.Selection.Use   = true;
		this.Selection.Start = true;
		this.Selection.Flag  = selectionflag_Common;

		if (true === SelectionUse_old && true === MouseEvent.ShiftKey && true === bOldSelectionIsCommon)
		{
			this.Selection_SetEnd(X, Y, this.CurPage, {
				Type           : AscCommon.g_mouse_event_type_up,
				ClickCount     : 1
			});
			this.Selection.Use    = true;
			this.Selection.Start  = true;
			this.Selection.EndPos = ContentPos;
			this.Selection.Data   = null;
		}
		else
		{
			var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, this.CurPage);
			Item.Selection_SetStart(X, Y, ElementPageIndex, MouseEvent);

			if (this.IsNumberingSelection())
			{
				// TODO: Можно сделать передвигание нумерации как в Word
				return;
			}

			if (isTopDocumentContent)
			{
				Item.Selection_SetEnd(X, Y, ElementPageIndex, {
					Type           : AscCommon.g_mouse_event_type_move,
					ClickCount     : 1
				});
			}

			if (true !== bTableBorder)
			{
				this.Selection.Use      = true;
				this.Selection.StartPos = ContentPos;
				this.Selection.EndPos   = ContentPos;
				this.Selection.Data     = null;

				this.CurPos.ContentPos = ContentPos;

				if (type_Paragraph === Item.GetType() && true === MouseEvent.CtrlKey)
				{
					var oHyperlink   = Item.CheckHyperlink(X, Y, ElementPageIndex);
					var oPageRefLink = Item.CheckPageRefLink(X, Y, ElementPageIndex);
					if (null != oHyperlink)
					{
						this.Selection.Data = {
							Hyperlink : oHyperlink
						};
					}
					else if (null !== oPageRefLink)
					{
						this.Selection.Data = {
							PageRef : oPageRefLink
						};
					}
				}
			}
			else
			{
				this.Selection.Data = {
					TableBorder : true,
					Pos         : ContentPos,
					Selection   : SelectionUse_old
				};
			}
		}
	}
};
// Данная функция может использоваться как при движении, так и при окончательном выставлении селекта.
// Если bEnd = true, тогда это конец селекта.
CDocumentContent.prototype.Selection_SetEnd = function(X, Y, CurPage, MouseEvent)
{
	if (!this.IsRecalculated())
		return;
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;
	
	if (CurPage < 0)
	{
		CurPage = 0;
		Y       = 0;
	}
	else if (CurPage >= this.Pages.length)
	{
		CurPage = this.Pages.length - 1;
		Y       = this.Pages[CurPage].YLimit;
	}

	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		var PageAbs = this.Get_StartPage_Absolute(CurPage);
		if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
		{
			this.LogicDocument.DrawingObjects.OnMouseUp(MouseEvent, X, Y, PageAbs);
			this.Selection.Start = false;
			this.Selection.Use   = true;
		}
		else
		{
			this.LogicDocument.DrawingObjects.OnMouseMove(MouseEvent, X, Y, PageAbs);
		}
		return;
	}

	this.CurPage = CurPage;

	if (selectionflag_Numbering === this.Selection.Flag)
		return;

	// Обрабатываем движение границы у таблиц
	if (null != this.Selection.Data && true === this.Selection.Data.TableBorder && type_Table == this.Content[this.Selection.Data.Pos].GetType())
	{
		var nPos = this.Selection.Data.Pos;

		// Убираем селект раньше, чтобы при создании точки в истории не сохранялось состояние передвижения границы таблицы
		if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
		{
			this.Selection.Start = false;
			this.Selection.Use   = this.Selection.Data.Selection;
			this.Selection.Data  = null;
		}

		var Item             = this.Content[nPos];
		var ElementPageIndex = this.private_GetElementPageIndexByXY(nPos, X, Y, this.CurPage);
		Item.Selection_SetEnd(X, Y, ElementPageIndex, MouseEvent);
		return;
	}

	if (false === this.Selection.Use)
		return;

	var ContentPos = this.Internal_GetContentPosByXY(X, Y);

	this.CurPos.ContentPos = ContentPos;
	var OldEndPos          = this.Selection.EndPos;
	this.Selection.EndPos  = ContentPos;

	// Удалим отметки о старом селекте
	if (OldEndPos < this.Selection.StartPos && OldEndPos < this.Selection.EndPos)
	{
		var TempLimit = Math.min(this.Selection.StartPos, this.Selection.EndPos);
		for (var Index = OldEndPos; Index < TempLimit; Index++)
		{
			this.Content[Index].RemoveSelection();
		}
	}
	else if (OldEndPos > this.Selection.StartPos && OldEndPos > this.Selection.EndPos)
	{
		var TempLimit = Math.max(this.Selection.StartPos, this.Selection.EndPos);
		for (var Index = TempLimit + 1; Index <= OldEndPos; Index++)
		{
			this.Content[Index].RemoveSelection();
		}
	}

	// Направление селекта: 1 - прямое, -1 - обратное, 0 - отмечен 1 элемент документа
	var Direction = ( ContentPos > this.Selection.StartPos ? 1 : ( ContentPos < this.Selection.StartPos ? -1 : 0 )  );

	if (AscCommon.g_mouse_event_type_up == MouseEvent.Type)
		this.StopSelection();

	var Start, End;
	if (0 == Direction)
	{
		var Item             = this.Content[this.Selection.StartPos];
		var ElementPageIndex = this.private_GetElementPageIndexByXY(this.Selection.StartPos, X, Y, this.CurPage);
		Item.Selection_SetEnd(X, Y, ElementPageIndex, MouseEvent);

		if (this.IsNumberingSelection())
		{
			// Ничего не делаем
		}
		else if (false === Item.IsSelectionUse())
		{
			this.Selection.Use = false;

			if (this.IsInText(X, Y, this.CurPage))
			{
				if (null != this.Selection.Data && this.Selection.Data.Hyperlink)
				{
					var oHyperlink    = this.Selection.Data.Hyperlink;
					var sBookmarkName = oHyperlink.GetAnchor();
					var sValue        = oHyperlink.GetValue();

					if (oHyperlink.IsTopOfDocument())
					{
						if (this.LogicDocument && this.LogicDocument.MoveCursorToStartOfDocument)
							this.LogicDocument.MoveCursorToStartOfDocument();
					}
					else if (sValue)
					{
						editor && editor.sync_HyperlinkClickCallback(sBookmarkName ? sValue + "#" + sBookmarkName : sValue);
						this.Selection.Data.Hyperlink.SetVisited(true);
						if (this.DrawingDocument.m_oLogicDocument)
						{
							if (editor.isDocumentEditor)
							{
								if (false == editor.isPdfEditor()) {
									for (var PageIdx = Item.Get_AbsolutePage(0); PageIdx < Item.Get_AbsolutePage(0) + Item.Get_PagesCount(); PageIdx++)
										this.DrawingDocument.OnRecalculatePage(PageIdx, this.DrawingDocument.m_oLogicDocument.Pages[PageIdx]);
								}
								else {
									if (this.Parent && this.Parent.parent && this.Parent.parent.IsDrawing()) {
										this.Parent.parent.SetNeedRecalc(true);	
									}
								}
							}
							else
							{
								if(this.DrawingDocument.OnRecalculateSlide)
								{
									this.DrawingDocument.OnRecalculateSlide(this.GetAbsolutePage(0));
								}
							}
							this.DrawingDocument.OnEndRecalculate(false, true);
						}
					}
					else if (sBookmarkName)
					{
						var oBookmarksManagers = this.LogicDocument && this.LogicDocument.GetBookmarksManager ? this.LogicDocument.GetBookmarksManager() : null;
						var oBookmark          = oBookmarksManagers ? oBookmarksManagers.GetBookmarkByName(sBookmarkName) : null;
						if (oBookmark)
							oBookmark[0].GoToBookmark();
					}
				}
				else if (null !== this.Selection.Data && this.Selection.Data.PageRef)
				{
					var oInstruction = this.Selection.Data.PageRef.GetInstruction();
					if (oInstruction && AscWord.fieldtype_PAGEREF === oInstruction.GetType())
					{
						var oBookmarksManagers = this.LogicDocument && this.LogicDocument.GetBookmarksManager ? this.LogicDocument.GetBookmarksManager() : null;
						var oBookmark          = oBookmarksManagers ? oBookmarksManagers.GetBookmarkByName(oInstruction.GetBookmarkName()) : null;
						if (oBookmark)
							oBookmark[0].GoToBookmark();
					}
				}
			}
		}
		else
		{
			this.Selection.Use = true;
		}

		return;
	}
	else if (Direction > 0)
	{
		Start = this.Selection.StartPos;
		End   = this.Selection.EndPos;
	}
	else
	{
		End   = this.Selection.StartPos;
		Start = this.Selection.EndPos;
	}

	// Чтобы не было эффекта, когда ничего не поселекчено, а при удалении соединяются параграфы
	if (Direction > 0 && type_Paragraph === this.Content[Start].GetType() && true === this.Content[Start].IsSelectionEmpty() && this.Content[Start].Selection.StartPos == this.Content[Start].Content.length - 1)
	{
		this.Content[Start].Selection.StartPos = this.Content[Start].Internal_GetEndPos();
		this.Content[Start].Selection.EndPos   = this.Content[Start].Content.length - 1;
	}

	var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, this.CurPage);
	this.Content[ContentPos].Selection_SetEnd(X, Y, ElementPageIndex, MouseEvent);

	for (var Index = Start; Index <= End; Index++)
	{
		var Item = this.Content[Index];
		Item.SetSelectionUse(true);

		switch (Index)
		{
			case Start:

				Item.SetSelectionToBeginEnd(Direction > 0 ? false : true, false);
				break;

			case End:

				Item.SetSelectionToBeginEnd(Direction > 0 ? true : false, true);
				break;

			default:

				Item.SelectAll(Direction);
				break;
		}
	}
};
CDocumentContent.prototype.CheckPosInSelection = function(X, Y, CurPage, NearPos)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.DrawingObjects.selectionCheck(X, Y, this.Get_AbsolutePage(CurPage), NearPos);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		let point = this.TransformPoint(X, Y);
		X = point.x;
		Y = point.y;
		
		if (true === this.Selection.Use || true === this.ApplyToAll)
		{
			switch (this.Selection.Flag)
			{
				case selectionflag_Common:
				{
					var Start = this.Selection.StartPos;
					var End   = this.Selection.EndPos;

					if (Start > End)
					{
						Start = this.Selection.EndPos;
						End   = this.Selection.StartPos;
					}

					if (undefined !== NearPos)
					{
						if (true === this.ApplyToAll)
						{
							Start = 0;
							End   = this.Content.length - 1;
						}

						for (var Index = Start; Index <= End; Index++)
						{
							if (true === this.ApplyToAll)
								this.Content[Index].SetApplyToAll(true);

							if (true === this.Content[Index].CheckPosInSelection(0, 0, 0, NearPos))
							{
								if (true === this.ApplyToAll)
									this.Content[Index].SetApplyToAll(false);

								return true;
							}

							if (true === this.ApplyToAll)
								this.Content[Index].SetApplyToAll(false);
						}

						return false;
					}
					else
					{
						var ContentPos = this.Internal_GetContentPosByXY(X, Y, CurPage);
						if (ContentPos > Start && ContentPos < End)
						{
							return true;
						}
						else if (ContentPos < Start || ContentPos > End)
						{
							return false;
						}
						else
						{
							var ElementPageIndex = this.private_GetElementPageIndexByXY(ContentPos, X, Y, CurPage);
							return this.Content[ContentPos].CheckPosInSelection(X, Y, ElementPageIndex, NearPos);
						}

						return false;
					}
				}
				case selectionflag_Numbering :
					return false;
			}

			return false;
		}

		return false;
	}
};
CDocumentContent.prototype.IsSelectionEmpty = function(bCheckHidden)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.selectionIsEmpty(bCheckHidden);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			// Выделение нумерации
			if (selectionflag_Numbering == this.Selection.Flag)
				return false;
			// Обрабатываем движение границы у таблиц
			else if (null != this.Selection.Data && true === this.Selection.Data.TableBorder && type_Table == this.Content[this.Selection.Data.Pos].GetType())
				return false;
			else
			{
				if (this.Selection.StartPos === this.Selection.EndPos)
					return this.Content[this.Selection.StartPos].IsSelectionEmpty(bCheckHidden);
				else
					return false;
			}
		}

		return true;
	}
};
CDocumentContent.prototype.SelectAll = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type && true === this.DrawingObjects.isSelectedText())
	{
		this.DrawingObjects.selectAll();
	}
	else
	{
		if (true === this.Selection.Use)
			this.RemoveSelection();

		this.SetDocPosType(docpostype_Content);
		this.Selection.Use   = true;
		this.Selection.Start = false;
		this.Selection.Flag  = selectionflag_Common;

		this.Selection.StartPos = 0;
		this.Selection.EndPos   = this.Content.length - 1;

		for (var Index = 0; Index < this.Content.length; Index++)
		{
			this.Content[Index].SelectAll();
		}
	}
};
CDocumentContent.prototype.SetSelectionUse = function(isUse)
{
	if (true === isUse)
		this.Selection.Use = true;
	else
		this.RemoveSelection();
};
CDocumentContent.prototype.SetSelectionToBeginEnd = function(isSelectionStart, isElementStart)
{
	if (this.Content.length <= 0)
		return;

	if (true === isElementStart)
	{
		this.Content[0].SetSelectionUse(true);
		this.Content[0].SetSelectionToBeginEnd(isSelectionStart, true);
		if (isSelectionStart)
			this.Selection.StartPos = 0;
		else
			this.Selection.EndPos = 0;
	}
	else
	{
		this.Content[this.Content.length - 1].SetSelectionUse(true);
		this.Content[this.Content.length - 1].SetSelectionToBeginEnd(isSelectionStart, false);

		if (isSelectionStart)
			this.Selection.StartPos = this.Content.length - 1;
		else
			this.Selection.EndPos = this.Content.length - 1;
	}
};
CDocumentContent.prototype.Select_DrawingObject      = function(Id)
{
	let drawingObject = AscCommon.g_oTableId.GetById(Id);
	if (!drawingObject || !drawingObject.IsUseInDocument())
		return;
	
    this.RemoveSelection();

    this.Parent.Set_CurrentElement(true, this.Get_StartPage_Absolute() + this.CurPage, this);

    // Прячем курсор
    this.DrawingDocument.TargetEnd();
    this.DrawingDocument.SetCurrentPage(this.Get_StartPage_Absolute() + this.CurPage);

    var HdrFtr = this.IsHdrFtr(true);
    if (null != HdrFtr)
    {
        HdrFtr.Content.SetDocPosType(docpostype_DrawingObjects);
        HdrFtr.Content.Selection.Use   = true;
        HdrFtr.Content.Selection.Start = false;

        this.LogicDocument.Selection.Use   = true;
        this.LogicDocument.Selection.Start = false;
    }
    else
    {
        this.LogicDocument.SetDocPosType(docpostype_DrawingObjects);
        this.LogicDocument.Selection.Use   = true;
        this.LogicDocument.Selection.Start = false;
    }

    this.LogicDocument.DrawingObjects.selectById(Id, this.Get_StartPage_Absolute() + this.CurPage);

    // TODO: Пока сделаем так, в будущем надо сделать функцию, которая у родительского класса обновляет Select
    editor.WordControl.m_oLogicDocument.Document_UpdateSelectionState();
    editor.WordControl.m_oLogicDocument.Document_UpdateInterfaceState();
};
//-----------------------------------------------------------------------------------
// Функции для работы с таблицами
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.AddTableRow = function(bBefore, nCount)
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableAddRow(bBefore, nCount);
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		this.Content[Pos].AddTableRow(bBefore, nCount);
		if (false === this.Selection.Use && true === this.Content[Pos].IsSelectionUse())
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = Pos;
			this.Selection.EndPos   = Pos;
		}

		return true;
	}

	return false;
};
CDocumentContent.prototype.AddTableColumn = function(bBefore, nCount)
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableAddCol(bBefore, nCount);
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		this.Content[Pos].AddTableColumn(bBefore, nCount);
		if (false === this.Selection.Use && true === this.Content[Pos].IsSelectionUse())
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = Pos;
			this.Selection.EndPos   = Pos;
		}

		return true;
	}

	return false;
};
CDocumentContent.prototype.RemoveTableRow = function()
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableRemoveRow();
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		if (false === this.Content[Pos].RemoveTableRow())
			this.RemoveTable();

		return true;
	}

	return false;
};
CDocumentContent.prototype.RemoveTableColumn = function()
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableRemoveCol();
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		if (false === this.Content[Pos].RemoveTableColumn())
			this.RemoveTable();

		return true;
	}

	return false;
};
CDocumentContent.prototype.MergeTableCells = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableMergeCells();
	}
	else if (docpostype_Content === this.CurPos.Type)
	{
		var nPos = true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos;
		return this.Content[nPos].MergeTableCells();
	}

	return false;
};
CDocumentContent.prototype.SplitTableCells = function(nCols, nRows)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableSplitCell();
	}
	else if (docpostype_Content === this.CurPos.Type)
	{
		var nPos = true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos;
		return this.Content[nPos].SplitTableCells(nCols, nRows);
	}

	return false;
};
CDocumentContent.prototype.RemoveTableCells = function()
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableRemoveCells();
	}
	else if (docpostype_Content == this.CurPos.Type
		&& ((true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && !this.Content[this.Selection.StartPos].IsParagraph())
		|| (false == this.Selection.Use && !this.Content[this.CurPos.ContentPos].IsParagraph())))
	{
		var nPos = true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos;
		if (false === this.Content[nPos].RemoveTableCells())
			this.RemoveTable();

		return true;
	}

	return false;
};
CDocumentContent.prototype.RemoveTable = function()
{
    if (docpostype_DrawingObjects == this.CurPos.Type)
    {
        return this.LogicDocument.DrawingObjects.tableRemoveTable();
    }
    else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType() ) ))
    {
        var Pos;
        if (true === this.Selection.Use)
            Pos = this.Selection.StartPos;
        else
            Pos = this.CurPos.ContentPos;

        var Table = this.Content[Pos];
        if (type_Table === Table.GetType())
		{
			if (true === Table.IsInnerTable())
			{
				Table.RemoveInnerTable();
			}
			else
			{
				var oLogicDocument    = this.GetLogicDocument();
				var isNeedRemoveTable = true;
				if (oLogicDocument && oLogicDocument.IsTrackRevisions())
				{
					this.Content[Pos].SelectAll();
					isNeedRemoveTable = !this.Content[Pos].RemoveTableRow();
				}

				if (isNeedRemoveTable)
				{
					this.RemoveSelection();
					Table.PreDelete();
					this.Internal_Content_Remove(Pos, 1);

					if (Pos >= this.Content.length - 1)
						Pos--;

					if (Pos < 0)
						Pos = 0;

					this.SetDocPosType(docpostype_Content);
					this.CurPos.ContentPos = Pos;
					this.Content[Pos].MoveCursorToStartPos();
				}
			}

			return true;
		}
		else
		{
			return Table.RemoveTable();
		}
    }
    return false;
};
CDocumentContent.prototype.SelectTable = function(Type)
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableSelect(Type);
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		this.Content[Pos].SelectTable(Type);
		if (false === this.Selection.Use && true === this.Content[Pos].IsSelectionUse())
		{
			this.Selection.Use      = true;
			this.Selection.StartPos = Pos;
			this.Selection.EndPos   = Pos;
		}
		return true;
	}

	return false;
};
CDocumentContent.prototype.CanMergeTableCells = function()
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableCheckMerge();
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		return this.Content[Pos].CanMergeTableCells();
	}

	return false;
};
CDocumentContent.prototype.CanSplitTableCells = function()
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.tableCheckSplit();
	}
	else if (docpostype_Content == this.CurPos.Type)
	{
		var nPos = true === this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos;
		return this.Content[nPos].CanSplitTableCells();
	}

	return false;
};
CDocumentContent.prototype.DistributeTableCells = function(isHorizontally)
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.distributeTableCells(isHorizontally);
	}
	else if (docpostype_Content == this.CurPos.Type && ( ( true === this.Selection.Use && this.Selection.StartPos == this.Selection.EndPos && type_Paragraph !== this.Content[this.Selection.StartPos].GetType() ) || ( false == this.Selection.Use && type_Paragraph !== this.Content[this.CurPos.ContentPos].GetType() ) ))
	{
		var Pos = 0;
		if (true === this.Selection.Use)
			Pos = this.Selection.StartPos;
		else
			Pos = this.CurPos.ContentPos;

		return this.Content[Pos].DistributeTableCells(isHorizontally);
	}

	return false;
};
//-----------------------------------------------------------------------------------
// Вспомогательные(внутренние ) функции
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.Internal_GetContentPosByXY = function(X, Y, PageNum)
{
	if (!this.IsRecalculated())
		return;
	
	let point = this.TransformPoint(X, Y);
	X = point.x;
	Y = point.y;
	
    if (undefined === PageNum || null === PageNum)
        PageNum = this.CurPage;

    PageNum = Math.max(0, Math.min(PageNum, this.Pages.length - 1));

	if(this.LogicDocument && this.LogicDocument.IsDocumentEditor())
	{
		var oFlow    =  this.LogicDocument.DrawingObjects.getTableByXY(X, Y, this.GetAbsolutePage(PageNum), this);
		var nFlowPos = this.private_GetContentIndexByFlowObject(oFlow, X, Y);
		if (-1 !== nFlowPos)
			return nFlowPos;
	}

    // Теперь проверим пустые параграфы с окончанием секций (в нашем случае это пустой параграф послей таблицы внутри таблицы)
    var SectCount = this.Pages[PageNum].EndSectionParas.length;
    for (var Index = 0; Index < SectCount; ++Index)
    {
        var Item   = this.Pages[PageNum].EndSectionParas[Index];
        var Bounds = Item.Pages[0].Bounds;

        if (Y < Bounds.Bottom && Y > Bounds.Top && X > Bounds.Left && X < Bounds.Right)
            return Item.Index;
    }

    var StartPos = Math.min(this.Pages[PageNum].Pos, this.Content.length - 1);
    var EndPos   = this.Content.length - 1;

    if (PageNum < this.Pages.length - 1)
        EndPos = Math.min(this.Pages[PageNum + 1].Pos, EndPos);

    // Сохраним позиции всех Inline элементов на данной странице
    var InlineElements = [];
    for (var Index = StartPos; Index <= EndPos; Index++)
    {
        var Item = this.Content[Index];
        var bEmptySectPara = this.Pages[PageNum].Check_EndSectionPara(Item);

        if (false != Item.Is_Inline() && (type_Paragraph !== Item.GetType() || false === bEmptySectPara))
            InlineElements.push(Index);
    }

    var Count = InlineElements.length;
    if (Count <= 0)
        return StartPos;

    for (var Pos = 0; Pos < Count - 1; Pos++)
    {
        var Item = this.Content[InlineElements[Pos + 1]];

		if (Item.GetPagesCount() <= 0 || Y < Item.GetPageBounds(0).Top)
			return InlineElements[Pos];

        if (Item.GetPagesCount() > 1)
        {
            if (true !== Item.IsStartFromNewPage())
                return InlineElements[Pos + 1];

            return InlineElements[Pos];
        }

        if (Pos === Count - 2)
        {
            // Такое возможно, если страница заканчивается Flow-таблицей
            return InlineElements[Count - 1];
        }
    }

    return InlineElements[0];
};
CDocumentContent.prototype.private_CheckCurPage = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		// TODO: переделать
		this.CurPage = 0;
	}
	else if (docpostype_Content === this.CurPos.Type)
	{
		if (true === this.Selection.Use)
		{
			this.CurPage = this.Content[this.Selection.EndPos].Get_CurrentPage_Relative();
		}
		else if (this.CurPos.ContentPos >= 0)
		{
			this.CurPage = this.Content[this.CurPos.ContentPos].Get_CurrentPage_Relative();
		}
	}
};
CDocumentContent.prototype.Internal_Content_Add = function(Position, NewObject, isCorrectContent)
{
	// Position = this.Content.length  допускается
	if (Position < 0 || Position > this.Content.length)
		return;

	var PrevObj = this.Content[Position - 1] ? this.Content[Position - 1] : null;
	var NextObj = this.Content[Position] ? this.Content[Position] : null;

	this.private_RecalculateNumbering([NewObject]);
	AscCommon.History.Add(new CChangesDocumentContentAddItem(this, Position, [NewObject]));
	this.Content.splice(Position, 0, NewObject);
	this.private_UpdateSelectionPosOnAdd(Position);
	NewObject.Set_Parent(this);
	NewObject.Set_DocumentNext(NextObj);
	NewObject.Set_DocumentPrev(PrevObj);

	if (null != PrevObj)
		PrevObj.Set_DocumentNext(NewObject);

	if (null != NextObj)
		NextObj.Set_DocumentPrev(NewObject);

	if (Position <= this.CurPos.TableMove)
		this.CurPos.TableMove++;

	// Проверим, что последний элемент - параграф или SdtBlockLevel.
	// В самом CSdtBlockLevel такая проверка не нужна
	if (false !== isCorrectContent
		&& !this.Content[this.Content.length - 1].IsParagraph()
		&& !this.Content[this.Content.length - 1].IsBlockLevelSdt()
		&& !this.IsBlockLevelSdtContent())
		this.Internal_Content_Add(this.Content.length, new AscWord.Paragraph(this, this.bPresentation === true));

	this.private_ReindexContent(Position);
	this.OnContentChange();
	this.Recalculated = false;
};
CDocumentContent.prototype.Internal_Content_Remove = function(Position, Count, isCorrectContent)
{
	if (Position < 0 || Position >= this.Content.length || Count <= 0)
		return;

	var PrevObj = this.Content[Position - 1] ? this.Content[Position - 1] : null;
	var NextObj = this.Content[Position + Count] ? this.Content[Position + Count] : null;

	for (var Index = 0; Index < Count; Index++)
		this.Content[Position + Index].PreDelete();

	AscCommon.History.Add(new CChangesDocumentContentRemoveItem(this, Position, this.Content.slice(Position, Position + Count)));
	var Elements = this.Content.splice(Position, Count);
	this.private_RecalculateNumbering(Elements);
	this.private_UpdateSelectionPosOnRemove(Position, Count);

	if (null != PrevObj)
		PrevObj.Set_DocumentNext(NextObj);

	if (null != NextObj)
		NextObj.Set_DocumentPrev(PrevObj);

	// Проверим, что последний элемент - параграф или SdtBlockLevel.
	// В самом CSdtBlockLevel такая проверка не нужна

	if (false !== isCorrectContent
		&& (this.Content.length <= 0
			|| (!this.Content[this.Content.length - 1].IsParagraph()
				&& !this.Content[this.Content.length - 1].IsBlockLevelSdt()
				&& !this.IsBlockLevelSdtContent())))
		this.Internal_Content_Add(this.Content.length, new AscWord.Paragraph(this, this.bPresentation === true));

	this.private_ReindexContent(Position);
	this.OnContentChange();
	this.Recalculated = false;
};
CDocumentContent.prototype.Clear_ContentChanges = function()
{
	this.m_oContentChanges.Clear();
};
CDocumentContent.prototype.Add_ContentChanges = function(Changes)
{
	this.m_oContentChanges.Add(Changes);
};
CDocumentContent.prototype.Refresh_ContentChanges = function()
{
	this.m_oContentChanges.Refresh();
};
CDocumentContent.prototype.Internal_Content_RemoveAll = function()
{
	for (let index = 0, count = this.Content.length; index < count; ++index)
	{
		this.Content[index].PreDelete();
	}

	AscCommon.History.Add(new CChangesDocumentRemoveItem(this, 0, this.Content.slice(0, this.Content.length)));
	this.Content = [];
};
//-----------------------------------------------------------------------------------
// Функции для работы с номерами страниц
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.Get_StartPage_Absolute = function()
{
	return this.Get_AbsolutePage(0);
};
CDocumentContent.prototype.Get_StartPage_Relative = function()
{
	return this.StartPage;
};
CDocumentContent.prototype.Get_StartColumn_Absolute = function()
{
	return this.Get_AbsoluteColumn(0);
};
CDocumentContent.prototype.Set_StartPage = function(StartPage, StartColumn, ColumnsCount)
{
	this.StartPage    = StartPage;
	this.StartColumn  = undefined !== StartColumn ? StartColumn : 0;
	this.ColumnsCount = undefined !== ColumnsCount ? ColumnsCount : 1;
};
CDocumentContent.prototype.Get_ColumnsCount = function()
{
	return this.ColumnsCount;
};
CDocumentContent.prototype.private_GetRelativePageIndex = function(CurPage)
{
	if (!this.ColumnsCount || 0 === this.ColumnsCount)
		return this.StartPage + CurPage;

	return this.StartPage + ((this.StartColumn + CurPage) / this.ColumnsCount | 0);
};
CDocumentContent.prototype.private_GetAbsolutePageIndex = function(CurPage)
{
	if (!this.Parent)
		return CurPage;

	return this.Parent.Get_AbsolutePage(this.private_GetRelativePageIndex(CurPage));
};
CDocumentContent.prototype.Get_StartColumn = function()
{
	return this.StartColumn;
};
CDocumentContent.prototype.Get_AbsolutePage = function(CurPage)
{
	return this.private_GetAbsolutePageIndex(CurPage);
};
CDocumentContent.prototype.Get_AbsoluteColumn = function(CurPage)
{
	return this.private_GetColumnIndex(CurPage);
};
CDocumentContent.prototype.private_GetColumnIndex = function(CurPage)
{
	// TODO: Разобраться здесь нужно ли данное условие. Оно появилось из-за параграфов в таблице в
	//       основной части документа и из-за параграфов в сносках.
	if (this.Parent && 1 === this.ColumnsCount)
    	return this.Parent.Get_AbsoluteColumn(this.private_GetRelativePageIndex(CurPage));

	return (this.StartColumn + CurPage) - (((this.StartColumn + CurPage) / this.ColumnsCount | 0) * this.ColumnsCount);
};
CDocumentContent.prototype.GetAbsolutePage = function(nCurPage)
{
	return this.Get_AbsolutePage(nCurPage);
};
CDocumentContent.prototype.GetAbsoluteColumn = function(nCurPage)
{
	return this.Get_AbsoluteColumn(nCurPage);
};
//-----------------------------------------------------------------------------------
// Undo/Redo функции
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.GetSelectionState = function()
{
	var DocState    = {};
	DocState.CurPos = {
		X          : this.CurPos.X,
		Y          : this.CurPos.Y,
		ContentPos : this.CurPos.ContentPos,
		RealX      : this.CurPos.RealX,
		RealY      : this.CurPos.RealY,
		Type       : this.CurPos.Type
	};

	DocState.Selection = {

		Start    : this.Selection.Start,
		Use      : this.Selection.Use,
		StartPos : this.Selection.StartPos,
		EndPos   : this.Selection.EndPos,
		Flag     : this.Selection.Flag,
		Data     : this.Selection.Data
	};
	
	DocState.ShiftView = {
		X : this.ShiftViewX,
		Y : this.ShiftViewY
	};

	DocState.CurPage = this.CurPage;

	var State = null;

	if (this.LogicDocument && (editor.isStartAddShape || editor.isInkDrawerOn()) && docpostype_DrawingObjects === this.CurPos.Type)
	{
		DocState.CurPos.Type     = docpostype_Content;
		DocState.Selection.Start = false;
		DocState.Selection.Use   = false;

		this.Content[DocState.CurPos.ContentPos].RemoveSelection();
		State = this.Content[this.CurPos.ContentPos].GetSelectionState();
	}
	else
	{
		// Работаем с колонтитулом
		if (docpostype_DrawingObjects === this.CurPos.Type)
		{
			State = this.LogicDocument.DrawingObjects.getSelectionState();
		}
		else// if (docpostype_Content === this.CurPos.Type)
		{
			if (true === this.Selection.Use)
			{
				if (selectionflag_Numbering === this.Selection.Flag)
				{
					State = [this.GetCurrentParagraph()];
				}
				else
				{
					var StartPos = this.Selection.StartPos;
					var EndPos   = this.Selection.EndPos;
					if (StartPos > EndPos)
					{
						var Temp = StartPos;
						StartPos = EndPos;
						EndPos   = Temp;
					}

					State = [];

					var TempState = [];
					for (var Index = StartPos; Index <= EndPos; Index++)
					{
						TempState.push(this.Content[Index].GetSelectionState());
					}

					State.push(TempState);
				}
			}
			else
			{
				State = this.Content[this.CurPos.ContentPos].GetSelectionState();
			}
		}
	}

	if (null != this.Selection.Data && true === this.Selection.Data.TableBorder)
	{
		DocState.Selection.Data = null;
	}

	State.push(DocState);
	return State;
};
CDocumentContent.prototype.SetSelectionState = function(State, StateIndex)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
		this.LogicDocument.DrawingObjects.resetSelection();

	if (State.length <= 0)
		return;
	
	if (undefined === StateIndex)
		StateIndex = State.length - 1;

	var DocState = State[StateIndex];

	this.CurPos = {
		X          : DocState.CurPos.X,
		Y          : DocState.CurPos.Y,
		ContentPos : DocState.CurPos.ContentPos,
		RealX      : DocState.CurPos.RealX,
		RealY      : DocState.CurPos.RealY,
		Type       : DocState.CurPos.Type
	};

	this.SetDocPosType(DocState.CurPos.Type);

	this.Selection = {

		Start    : DocState.Selection.Start,
		Use      : DocState.Selection.Use,
		StartPos : DocState.Selection.StartPos,
		EndPos   : DocState.Selection.EndPos,
		Flag     : DocState.Selection.Flag,
		Data     : DocState.Selection.Data
	};
	
	this.ResetShiftView();
	if (DocState.ShiftView)
		this.ShiftView(DocState.ShiftView.X, DocState.ShiftView.Y);

	this.CurPage = DocState.CurPage;

	var NewStateIndex = StateIndex - 1;

	// Работаем с колонтитулом
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.LogicDocument.DrawingObjects.setSelectionState(State, NewStateIndex);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			if (selectionflag_Numbering === this.Selection.Flag)
			{
				let curPara = State[NewStateIndex];
				if (curPara && curPara.IsParagraph && curPara.IsParagraph())
				{
					let numPr     = curPara.GetNumPr();
					let prevNumPr = curPara.GetPrChangeNumPr();
					
					if (numPr && numPr.IsValid())
						this.SelectNumbering(numPr, curPara);
					else if (prevNumPr && prevNumPr.IsValid())
						this.SelectNumberingSingleParagraph(curPara);
					else
						this.LogicDocument.RemoveSelection();
				}
				else
				{
					this.LogicDocument.RemoveSelection();
				}
			}
			else
			{
				var StartPos = this.Selection.StartPos;
				var EndPos   = this.Selection.EndPos;
				if (StartPos > EndPos)
				{
					var Temp = StartPos;
					StartPos = EndPos;
					EndPos   = Temp;
				}

				var CurState = State[NewStateIndex];

				for (var Index = StartPos; Index <= EndPos; Index++)
				{
					this.Content[Index].SetSelectionState(CurState[Index - StartPos], CurState[Index - StartPos].length - 1);
				}
			}
		}
		else
		{
			this.Content[this.CurPos.ContentPos].SetSelectionState(State, NewStateIndex);
		}
	}
};
CDocumentContent.prototype.Get_ParentObject_or_DocumentPos = function()
{
    return this.Parent.Get_ParentObject_or_DocumentPos();
};
CDocumentContent.prototype.Refresh_RecalcData = function(oData)
{
	var nCurPage = 0;
	switch (oData.Type)
	{
		case AscDFH.historyitem_DocumentContent_AddItem:
		case AscDFH.historyitem_DocumentContent_RemoveItem:
		{
			var nDataPos = 0;
			if (oData instanceof CChangesDocumentContentAddItem || oData instanceof CChangesDocumentContentRemoveItem)
				nDataPos = oData.GetMinPos();
			else if (undefined !== oData.Pos)
				nDataPos = oData.Pos;

			for (nCurPage = this.Pages.length - 1; nCurPage > 0; nCurPage--)
			{
				if (nDataPos > this.Pages[nCurPage].Pos)
					break;
			}

			break;
		}
	}

	this.Refresh_RecalcData2(0, nCurPage);
};
CDocumentContent.prototype.Refresh_RecalcData2 = function(nIndex, nPageRel)
{
	if (-1 === nIndex || !this.Parent)
		return;

	this.Parent.Refresh_RecalcData2(this.StartPage + nPageRel);
};
//-----------------------------------------------------------------------------------
// Функции для работы с гиперссылками
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.AddHyperlink = function(HyperProps)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.hyperlinkAdd(HyperProps);
	}
	else if (docpostype_Content === this.CurPos.Type
		&& (false === this.Selection.Use || this.Selection.StartPos === this.Selection.EndPos))
	{
		var Pos = ( true == this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
		return this.Content[Pos].AddHyperlink(HyperProps);
	}
	
	return null;
};
CDocumentContent.prototype.ModifyHyperlink = function(HyperProps)
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.hyperlinkModify(HyperProps);
	}
	else if (docpostype_Content == this.CurPos.Type
		&& (false === this.Selection.Use || this.Selection.StartPos === this.Selection.EndPos))
	{
		var Pos = ( true == this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
		this.Content[Pos].ModifyHyperlink(HyperProps);
	}
};
CDocumentContent.prototype.RemoveHyperlink = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.hyperlinkRemove();
	}
	else if (docpostype_Content == this.CurPos.Type
		&& (false === this.Selection.Use || this.Selection.StartPos === this.Selection.EndPos))
	{
		var Pos = ( true == this.Selection.Use ? this.Selection.StartPos : this.CurPos.ContentPos );
		this.Content[Pos].RemoveHyperlink();
	}
};
CDocumentContent.prototype.CanAddHyperlink = function(bCheckInHyperlink)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.hyperlinkCanAdd(bCheckInHyperlink);
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			switch (this.Selection.Flag)
			{
				case selectionflag_Numbering:
					return false;
				case selectionflag_Common:
				{
					if (this.Selection.StartPos != this.Selection.EndPos)
						return false;

					return this.Content[this.Selection.StartPos].CanAddHyperlink(bCheckInHyperlink);
				}
			}
		}
		else
			return this.Content[this.CurPos.ContentPos].CanAddHyperlink(bCheckInHyperlink);
	}

	return false;
};
CDocumentContent.prototype.IsCursorInHyperlink = function(bCheckEnd)
{
	if (docpostype_DrawingObjects == this.CurPos.Type)
	{
		return this.LogicDocument.DrawingObjects.hyperlinkCheck(bCheckEnd);
	}
	else //if ( docpostype_Content == this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			switch (this.Selection.Flag)
			{
				case selectionflag_Numbering:
					return null;
				case selectionflag_Common:
				{
					if (this.Selection.StartPos != this.Selection.EndPos)
						return null;

					return this.Content[this.Selection.StartPos].IsCursorInHyperlink(bCheckEnd);
				}
			}
		}
		else
		{
			return this.Content[this.CurPos.ContentPos].IsCursorInHyperlink(bCheckEnd);
		}
	}

	return null;
};
//-----------------------------------------------------------------------------------
// Функции для работы с совместным редактирования
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.Write_ToBinary2 = function(Writer)
{
	Writer.WriteLong(AscDFH.historyitem_type_DocumentContent);

	// String : Id текущего элемента
	// Long   : StartPage
	// String : Id родительского класса
	// Bool   : TurnOffInnerWrap
	// Bool   : Split
	// Long   : Количество элементов в массиве this.Content
	// Array of string : массив Id элементов

	Writer.WriteString2(this.Id);
	Writer.WriteLong(this.StartPage);
	Writer.WriteString2(this.Parent.Get_Id());
	Writer.WriteBool(this.TurnOffInnerWrap);
	Writer.WriteBool(this.Split);
	AscFormat.writeBool(Writer, this.bPresentation);


	var ContentToWrite;
	if (this.StartState)
	{
		ContentToWrite = this.StartState.Content;
	}
	else
	{
		ContentToWrite = this.Content;
	}

	var Count = ContentToWrite.length;
	Writer.WriteLong(Count);
	for (var Index = 0; Index < Count; Index++)
		Writer.WriteString2(ContentToWrite[Index].Get_Id());

};
CDocumentContent.prototype.Read_FromBinary2 = function(Reader)
{
	// String : Id текущего элемента
	// Long   : StartPage
	// String : Id родительского класса
	// Bool   : TurnOffInnerWrap
	// Bool   : Split
	// Long   : Количество элементов в массиве this.Content
	// Array of string : массив Id элементов

	this.Id               = Reader.GetString2();
	this.StartPage        = Reader.GetLong();

	
	// Сам класс не должен проставлять себе родительский класс. Он должен проставляться ТОЛЬКО родительским классом при
	// при добавлении в своего содержимое. Пока оставляю тут эту заглушку, чтобы в таблицах работало
	let parent = g_oTableId.Get_ById(Reader.GetString2());
	if (parent)
	{
		if (parent.SetDocumentContent)
			parent.SetDocumentContent(this);
		else
			this.Parent = parent;
	}
	
	this.TurnOffInnerWrap = Reader.GetBool();
	this.Split            = Reader.GetBool();
	this.bPresentation    = AscFormat.readBool(Reader);

	var Count    = Reader.GetLong();
	this.Content = [];
	for (var Index = 0; Index < Count; Index++)
	{
		var Element = g_oTableId.Get_ById(Reader.GetString2());
		if (null != Element)
		{
			this.Content.push(Element);
			Element.Parent = this;
		}
	}

	var oCellApi = window["Asc"] && window["Asc"]["editor"];
	if (oCellApi && oCellApi.wbModel)
	{
		this.DrawingDocument = oCellApi.wbModel.DrawingDocument;
	}
	else
	{
		var DrawingDocument;
		if (editor && editor.WordControl && editor.WordControl.m_oDrawingDocument)
			DrawingDocument = editor.WordControl.m_oDrawingDocument;
		if (undefined !== DrawingDocument && null !== DrawingDocument)
		{
			this.DrawingDocument = DrawingDocument;

			if (undefined !== editor && true === editor.isDocumentEditor)
			{
				this.LogicDocument  = DrawingDocument.m_oLogicDocument;
				this.DrawingObjects = DrawingDocument.m_oLogicDocument.DrawingObjects; // Массив укзателей на все инлайновые графические объекты
			}
		}
	}
};
CDocumentContent.prototype.Get_SelectionState2 = function()
{
    // Сохраняем Id ближайшего элемента в текущем классе
    var State = new CDocumentSelectionState();

    State.Id   = this.Get_Id();
    State.Type = docpostype_Content;

    var Element = this.Content[this.CurPos.ContentPos];
    State.Data  = Element.Get_SelectionState2();

    return State;
};
CDocumentContent.prototype.Set_SelectionState2 = function(State)
{
    var ElementId = State.Data.Id;

    var CurId = ElementId;

    var bFlag = false;

    var Pos = 0;

    // Найдем элемент с Id = CurId
    var Count = this.Content.length;
    for (Pos = 0; Pos < Count; Pos++)
    {
        if (this.Content[Pos].Get_Id() == CurId)
        {
            bFlag = true;
            break;
        }
    }

    if (true !== bFlag)
    {
        var TempElement = g_oTableId.Get_ById(CurId);
        Pos             = ( null != TempElement ? Math.min(this.Content.length - 1, TempElement.Index) : 0 );
    }

    this.Selection.Start    = false;
    this.Selection.Use      = false;
    this.Selection.StartPos = Pos;
    this.Selection.EndPos   = Pos;
    this.Selection.Flag     = selectionflag_Common;

    this.SetDocPosType(docpostype_Content);
    this.CurPos.ContentPos = Pos;

    if (true !== bFlag)
        this.Content[this.CurPos.ContentPos].MoveCursorToStartPos();
    else
    {
        this.Content[this.CurPos.ContentPos].Set_SelectionState2(State.Data);
    }
};
//-----------------------------------------------------------------------------------
// Функции для работы с комментариями
//-----------------------------------------------------------------------------------
CDocumentContent.prototype.AddComment = function(Comment, bStart, bEnd)
{
	if (true === this.ApplyToAll)
	{
		if (this.Content.length <= 1 && true === bStart && true === bEnd)
		{
			this.Content[0].SetApplyToAll(true);
			this.Content[0].AddComment(Comment, true, true);
			this.Content[0].SetApplyToAll(false);
		}
		else
		{
			if (true === bStart)
			{
				this.Content[0].SetApplyToAll(true);
				this.Content[0].AddComment(Comment, true, false);
				this.Content[0].SetApplyToAll(false);
			}

			if (true === bEnd)
			{
				this.Content[this.Content.length - 1].SetApplyToAll(true);
				this.Content[this.Content.length - 1].AddComment(Comment, false, true);
				this.Content[this.Content.length - 1].SetApplyToAll(true);
			}
		}
	}
	else
	{
		if (docpostype_DrawingObjects === this.CurPos.Type)
		{
			var oLogicDocument  = this.GetLogicDocument();
			var oDrawingObjects = oLogicDocument.DrawingObjects;

			if (!oDrawingObjects.isSelectedText())
			{
				var oParaDrawing = oDrawingObjects.getMajorParaDrawing();
				var oParagraph;
				if (oParaDrawing && (oParagraph = oParaDrawing.GetParagraph()))
					oParagraph.AddCommentToDrawingObject(Comment, oParaDrawing.GetId());
			}
			else
			{
				oDrawingObjects.addComment(Comment);
			}
		}
		else //if ( docpostype_Content === this.CurPos.Type )
		{
			if (selectionflag_Numbering === this.Selection.Flag)
				return;

			if (true === this.Selection.Use)
			{
				var StartPos, EndPos;
				if (this.Selection.StartPos < this.Selection.EndPos)
				{
					StartPos = this.Selection.StartPos;
					EndPos   = this.Selection.EndPos;
				}
				else
				{
					StartPos = this.Selection.EndPos;
					EndPos   = this.Selection.StartPos;
				}

				if (StartPos === EndPos)
					this.Content[StartPos].AddComment(Comment, bStart, bEnd);
				else
				{
					if (true === bStart)
						this.Content[StartPos].AddComment(Comment, true, false);

					if (true === bEnd)
						this.Content[EndPos].AddComment(Comment, false, true);
				}
			}
			else
			{
				this.Content[this.CurPos.ContentPos].AddComment(Comment, bStart, bEnd);
			}
		}
	}
};
CDocumentContent.prototype.CanAddComment = function()
{
	if (true === this.ApplyToAll)
	{
		if (this.Content.length > 1)
		{
			return true;
		}
		else
		{
			var oElement = this.Content[0];
			oElement.SetApplyToAll(true);
			var isCanAdd = oElement.CanAddComment();
			oElement.SetApplyToAll(false);
			return isCanAdd;
		}
	}
	else
	{
		if (docpostype_DrawingObjects === this.CurPos.Type)
		{
			if (true != this.LogicDocument.DrawingObjects.isSelectedText())
				return true;
			else
				return this.LogicDocument.DrawingObjects.canAddComment();
		}
		else //if ( docpostype_Content === this.CurPos.Type )
		{
			switch (this.Selection.Flag)
			{
				case selectionflag_Numbering:
					return false;
				case selectionflag_Common:
				{
					if (true === this.Selection.Use && this.Selection.StartPos != this.Selection.EndPos)
						return true;
					else
					{
						var Pos     = ( this.Selection.Use === true ? this.Selection.StartPos : this.CurPos.ContentPos );
						var Element = this.Content[Pos];
						return Element.CanAddComment();
					}
				}
			}
		}
	}

	return false;
};
CDocumentContent.prototype.GetSelectionBounds = function()
{
	if (true === this.Selection.Use && selectionflag_Common === this.Selection.Flag)
	{
		var Start = this.Selection.StartPos;
		var End   = this.Selection.EndPos;

		if (Start > End)
		{
			Start = this.Selection.EndPos;
			End   = this.Selection.StartPos;
		}

		if (Start === End)
			return this.Content[Start].GetSelectionBounds();
		else
		{
			var Result       = {};
			Result.Start     = this.Content[Start].GetSelectionBounds().Start;
			Result.End       = this.Content[End].GetSelectionBounds().End;
			Result.Direction = (this.Selection.StartPos > this.Selection.EndPos ? -1 : 1);
			return Result;
		}
	}
	else if (this.Content[this.CurPos.ContentPos])
	{
		return this.Content[this.CurPos.ContentPos].GetSelectionBounds();
	}

	return null;
};
CDocumentContent.prototype.GetSelectionAnchorPos = function()
{
	var Pos = ( true === this.Selection.Use ? ( this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos ) : this.CurPos.ContentPos );
	return this.Content[Pos].GetSelectionAnchorPos();
};
CDocumentContent.prototype.GetEndInfo = function()
{
	var ContentLen = this.Content.length;
	if (ContentLen > 0)
		return this.Content[ContentLen - 1].GetEndInfo();
	else
		return null;
};
CDocumentContent.prototype.GetPrevElementEndInfo = function(CurElement)
{
	var PrevElement = CurElement.Get_DocumentPrev();

	if (null !== PrevElement && undefined !== PrevElement)
	{
		return PrevElement.GetEndInfo();
	}
	else if (this.Parent)
	{
		return this.Parent.GetPrevElementEndInfo(this);
	}

	return null;
};
CDocumentContent.prototype.GetTopElement = function()
{
    if (this.Parent)
        return this.Parent.GetTopElement();

    return null;
};
CDocumentContent.prototype.CompareDrawingsLogicPositions = function(CompareObject)
{
    for (var Index = 0, Count = this.Content.length; Index < Count; Index++)
    {
        var Element = this.Content[Index];
        Element.CompareDrawingsLogicPositions(CompareObject);

        if (0 !== CompareObject.Result)
            return;
    }
};
CDocumentContent.prototype.StartSelectionFromCurPos = function()
{
    if (docpostype_DrawingObjects === this.CurPos.Type)
    {
        return this.DrawingObjects.startSelectionFromCurPos();
    }
    else //if (docpostype_Content === this.CurPos.Type)
    {
        this.Selection.Use      = true;
        this.Selection.Start    = false;
        this.Selection.StartPos = this.CurPos.ContentPos;
        this.Selection.EndPos   = this.CurPos.ContentPos;
        this.Content[this.CurPos.ContentPos].StartSelectionFromCurPos();
    }
};
CDocumentContent.prototype.GetStyleFromFormatting = function()
{
    if (docpostype_DrawingObjects === this.CurPos.Type)
    {
        return this.DrawingObjects.GetStyleFromFormatting();
    }
    else //if (docpostype_Content === this.CurPos.Type)
    {
        if (true == this.Selection.Use)
        {
            if (this.Selection.StartPos > this.Selection.EndPos)
                return this.Content[this.Selection.EndPos].GetStyleFromFormatting();
            else
                return this.Content[this.Selection.StartPos].GetStyleFromFormatting();
        }
        else
        {
            return this.Content[this.CurPos.ContentPos].GetStyleFromFormatting();
        }
    }
};
CDocumentContent.prototype.IsTrackRevisions = function()
{
    if (this.LogicDocument)
        return this.LogicDocument.IsTrackRevisions();

    return false;
};
CDocumentContent.prototype.Get_SectPr = function()
{
    if (this.Parent && this.Parent.Get_SectPr)
        return this.Parent.Get_SectPr();

    return null;
};
CDocumentContent.prototype.SetParagraphFramePr = function(FramePr, bDelete)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		// Не добавляем и не работаем с рамками в автофигурах
		return;
	}
	else //if ( docpostype_Content === this.CurPos.Type )
	{
		if (true === this.Selection.Use)
		{
			// Проверим, если у нас все выделенные элементы - параграфы, с одинаковыми настройками
			// FramePr, тогда мы можем применить новую настройку FramePr

			var StartPos = this.Selection.StartPos;
			var EndPos   = this.Selection.EndPos;

			if (StartPos > EndPos)
			{
				StartPos = this.Selection.EndPos;
				EndPos   = this.Selection.StartPos;
			}

			var Element = this.Content[StartPos];

			if (type_Paragraph !== Element.GetType() || undefined === Element.Get_FramePr())
				return;

			var FramePr = Element.Get_FramePr();
			for (var Pos = StartPos + 1; Pos < EndPos; Pos++)
			{
				var TempElement = this.Content[Pos];

				if (type_Paragraph !== TempElement.GetType() || undefined === TempElement.Get_FramePr() || true != FramePr.Compare(TempElement.Get_FramePr()))
					return;
			}

			// Раз дошли до сюда, значит можно у всех выделенных параграфов менять настройку рамки
			var FrameParas = this.Content[StartPos].Internal_Get_FrameParagraphs();
			var FrameCount = FrameParas.length;
			for (var Pos = 0; Pos < FrameCount; Pos++)
			{
				FrameParas[Pos].Set_FramePr(FramePr, bDelete);
			}
		}
		else
		{
			var Element = this.Content[this.CurPos.ContentPos];

			if (type_Paragraph !== Element.GetType())
				return;

			// Возможно, предыдущий элемент является буквицей
			if (undefined === Element.Get_FramePr())
			{
				var PrevElement = Element.Get_DocumentPrev();

				if (type_Paragraph !== PrevElement.GetType() || undefined === PrevElement.Get_FramePr() || undefined === PrevElement.Get_FramePr().DropCap)
					return;

				Element = PrevElement;
			}

			var FrameParas = Element.Internal_Get_FrameParagraphs();
			var FrameCount = FrameParas.length;
			for (var Pos = 0; Pos < FrameCount; Pos++)
			{
				FrameParas[Pos].Set_FramePr(FramePr, bDelete);
			}
		}
	}
};
CDocumentContent.prototype.Add_ToContent = function(Pos, Item, isCorrectContent)
{
    this.Internal_Content_Add(Pos, Item, isCorrectContent);
};
CDocumentContent.prototype.Remove_FromContent = function(Pos, Count, isCorrectContent)
{
    this.Internal_Content_Remove(Pos, Count, isCorrectContent);
};
CDocumentContent.prototype.GetContentPosition = function(bSelection, bStart, PosArray)
{
    if (undefined === PosArray)
        PosArray = [];

    var Pos = (true === bSelection ? (true === bStart ? this.Selection.StartPos : this.Selection.EndPos) : this.CurPos.ContentPos);
    PosArray.push({Class : this, Position : Pos});

    if (undefined !== this.Content[Pos] && this.Content[Pos].GetContentPosition)
        this.Content[Pos].GetContentPosition(bSelection, bStart, PosArray);

    return PosArray;
};
CDocumentContent.prototype.GetDocumentPositionFromObject = function(arrPos)
{
    if (!arrPos)
		arrPos = [];

    if (this.Parent && this.Parent.GetDocumentPositionFromObject)
		this.Parent.GetDocumentPositionFromObject(arrPos);

    return arrPos;
};
CDocumentContent.prototype.SetContentSelection = function(StartDocPos, EndDocPos, Depth, StartFlag, EndFlag)
{
    if ((0 === StartFlag && (!StartDocPos[Depth] || this !== StartDocPos[Depth].Class)) || (0 === EndFlag && (!EndDocPos[Depth] || this !== EndDocPos[Depth].Class)))
        return;

    if (this.Content.length <= 0)
        return;

    var StartPos = 0, EndPos = 0;
    switch (StartFlag)
    {
        case 0 : StartPos = StartDocPos[Depth].Position; break;
        case 1 : StartPos = 0; break;
        case -1: StartPos = this.Content.length - 1; break;
    }

    switch (EndFlag)
    {
        case 0 : EndPos = EndDocPos[Depth].Position; break;
        case 1 : EndPos = 0; break;
        case -1: EndPos = this.Content.length - 1; break;
    }

    var _StartDocPos = StartDocPos, _StartFlag = StartFlag;
    if (null !== StartDocPos && true === StartDocPos[Depth].Deleted)
    {
        if (StartPos < this.Content.length)
        {
            _StartDocPos = null;
            _StartFlag = 1;
        }
        else if (StartPos > 0)
        {
            StartPos--;
            _StartDocPos = null;
            _StartFlag = -1;
        }
        else
        {
            // Такого не должно быть
            return;
        }
    }

    var _EndDocPos = EndDocPos, _EndFlag = EndFlag;
    if (null !== EndDocPos && true === EndDocPos[Depth].Deleted)
    {
        if (EndPos < this.Content.length)
        {
            _EndDocPos = null;
            _EndFlag = 1;
        }
        else if (EndPos > 0)
        {
            EndPos--;
            _EndDocPos = null;
            _EndFlag = -1;
        }
        else
        {
            // Такого не должно быть
            return;
        }
    }

    StartPos = Math.min(this.Content.length - 1, Math.max(0, StartPos));
    EndPos   = Math.min(this.Content.length - 1, Math.max(0, EndPos));

    this.Selection.Use      = true;
    this.Selection.StartPos = StartPos;
    this.Selection.EndPos   = EndPos;

    if (StartPos !== EndPos)
    {
        this.Content[StartPos].SetContentSelection(_StartDocPos, null, Depth + 1, _StartFlag, StartPos > EndPos ? 1 : -1);
        this.Content[EndPos].SetContentSelection(null, _EndDocPos, Depth + 1, StartPos > EndPos ? -1 : 1, _EndFlag);

        var _StartPos = StartPos;
        var _EndPos = EndPos;
        var Direction = 1;

        if (_StartPos > _EndPos)
        {
            _StartPos = EndPos;
            _EndPos = StartPos;
            Direction = -1;
        }

        for (var CurPos = _StartPos + 1; CurPos < _EndPos; CurPos++)
        {
            this.Content[CurPos].SelectAll(Direction);
        }
    }
    else
    {
        this.Content[StartPos].SetContentSelection(_StartDocPos, _EndDocPos, Depth + 1, _StartFlag, _EndFlag);
    }
};
CDocumentContent.prototype.SetContentPosition = function(DocPos, Depth, Flag)
{
    if (0 === Flag && (!DocPos[Depth] || this !== DocPos[Depth].Class))
        return;

    if (this.Content.length <= 0)
        return;

    var Pos = 0;
    switch (Flag)
    {
        case 0 : Pos = DocPos[Depth].Position; break;
        case 1 : Pos = 0; break;
        case -1: Pos = this.Content.length - 1; break;
    }

    var _DocPos = DocPos, _Flag = Flag;
    if (null !== DocPos && true === DocPos[Depth].Deleted)
    {
        if (Pos < this.Content.length)
        {
            _DocPos = null;
            _Flag = 1;
        }
        else if (Pos > 0)
        {
            Pos--;
            _DocPos = null;
            _Flag = -1;
        }
        else
        {
            // Такого не должно быть
            return;
        }
    }

    Pos = Math.min(this.Content.length - 1, Math.max(0, Pos));
    this.CurPos.ContentPos = Pos;
    this.Content[Pos].SetContentPosition(_DocPos, Depth + 1, _Flag);
};
CDocumentContent.prototype.private_GetElementPageIndex = function(ElementPos, PageIndex, ColumnIndex, ColumnsCount)
{
    var Element = this.Content[ElementPos];
    if (!Element || Element.GetPagesCount() <= 0)
        return 0;

    var StartPage   = Element.Get_StartPage_Relative();
    var StartColumn = Element.Get_StartColumn();

    return ColumnIndex - StartColumn + (PageIndex - StartPage) * ColumnsCount;
};
CDocumentContent.prototype.private_GetElementPageIndexByXY = function(ElementPos, X, Y, PageIndex)
{
    return this.private_GetElementPageIndex(ElementPos, PageIndex, 0, 1);
};
/**
 * Получаем относительную страницу по заданным координатам и абсолютной странице
 * @param X
 * @param Y
 * @param nPageAbs
 * @returns {number}
 */
CDocumentContent.prototype.GetPageIndexByXYAndPageAbs = function(X, Y, nPageAbs)
{
	 var nResultPage  = 0;
	 var nMinDistance = null;

	 for (var nCurPage = 0, nPagesCount = this.Pages.length; nCurPage < nPagesCount; ++nCurPage)
	 {
	 	var nTempPageAbs = this.GetAbsolutePage(nCurPage);
	 	if (nTempPageAbs === nPageAbs)
		{
			var oBounds = this.Pages[nCurPage].Bounds;
			if (oBounds.Left < X && X < oBounds.Right && oBounds.Top < Y && Y < oBounds.Bottom)
			{
				return nCurPage;
			}
			else
			{
				var nTempDistance;
				if (oBounds.Left < X && X < oBounds.Right)
					nTempDistance = Math.min(Math.abs(oBounds.Top - Y), Math.abs(oBounds.Bottom - Y));
				else if (oBounds.Top < Y && Y < oBounds.Bottom)
					nTempDistance = Math.min(Math.abs(oBounds.Left - X), Math.abs(oBounds.Right - X));
				else
					nTempDistance = Math.max(Math.min(Math.abs(oBounds.Top - Y), Math.abs(oBounds.Bottom - Y)), Math.min(Math.abs(oBounds.Left - X), Math.abs(oBounds.Right - X)));

				if (null === nMinDistance || nTempDistance < nMinDistance)
				{
					nResultPage  = nCurPage;
					nMinDistance = nTempDistance;
				}
			}
		}
		else if (nTempPageAbs > nPageAbs)
		{
			break;
		}
	 }

	 return nResultPage;
};
CDocumentContent.prototype.GetTopDocumentContent = function(isOneLevel)
{
    var TopDocument = null;
    if (true !== isOneLevel && this.Parent && this.Parent.GetTopDocumentContent)
        TopDocument = this.Parent.GetTopDocumentContent();

    if (null !== TopDocument && undefined !== TopDocument)
        return TopDocument;

    return this;
};
CDocumentContent.prototype.Set_ParaPropsForVerticalTextInCell = function(isVerticalText)
{
    for (var Pos = 0, Count = this.Content.length; Pos < Count; ++Pos)
    {
        var Element = this.Content[Pos];
        if (type_Paragraph === Element.Get_Type())
            Element.Set_ParaPropsForVerticalTextInCell(isVerticalText);
    }
};
CDocumentContent.prototype.Set_LogicDocument = function(oLogicDocument)
{
	this.SetLogicDocument(oLogicDocument);
};
CDocumentContent.prototype.SetLogicDocument = function(oLogicDocument)
{
	this.LogicDocument   = oLogicDocument;
	this.DrawingObjects  = oLogicDocument.DrawingObjects;
};
CDocumentContent.prototype.Get_LogicDocument = function()
{
	return this.LogicDocument;
};
CDocumentContent.prototype.GetLogicDocument = function()
{
	if (!this.LogicDocument && this.Parent && this.Parent.GetLogicDocument)
		this.LogicDocument = this.Parent.GetLogicDocument();
	
	return this.LogicDocument;
};
CDocumentContent.prototype.RemoveTextSelection = function()
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.DrawingObjects.removeTextSelection();
	}
	else
	{
		this.RemoveSelection();
	}
};
CDocumentContent.prototype.CanUpdateTarget = function(CurPage)
{
	if (!this.IsRecalculated())
		return false;

	if (this.Pages.length <= CurPage)
		return true;

	var nPos = (this.Selection.Use ? this.Selection.EndPos : this.CurPos.ContentPos);


	if (this.Pages[CurPage].EndPos > nPos)
		return true;
	else if (this.Pages[CurPage].EndPos < nPos)
		return false;

	var nElementPageIndex = this.private_GetElementPageIndex(nPos, CurPage, 0, 1);
	return this.Content[nPos].CanUpdateTarget(nElementPageIndex);
};
CDocumentContent.prototype.IsStartFromNewPage = function()
{
	if (this.Content.length <= 0)
		return false;

	return this.Content[0].IsStartFromNewPage();
};
CDocumentContent.prototype.PreDelete = function()
{
	let logicDocument = this.GetLogicDocument();
	if (logicDocument && logicDocument.IsDocumentEditor() && logicDocument.isPreventedPreDelete())
		return;
	
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].PreDelete();
	}

	this.RemoveSelection();
};
CDocumentContent.prototype.IsBlockLevelSdtContent = function()
{
	return (this.Parent && this.Parent instanceof CBlockLevelSdt);
};
CDocumentContent.prototype.IsBlockLevelSdtFirstOnNewPage = function()
{
	if (this.Parent && this.Parent instanceof CBlockLevelSdt)
		return this.Parent.IsBlockLevelSdtFirstOnNewPage();

	return false;
};
CDocumentContent.prototype.IsSelectedAll = function()
{
	if (true === this.Selection.Use
		&& ((0 === this.Selection.StartPos && this.Content.length - 1 === this.Selection.EndPos)
		|| (0 === this.Selection.EndPos && this.Content.length - 1 === this.Selection.StartPos))
		&& true === this.Content[0].IsSelectedAll()
		&& true === this.Content[this.Content.length - 1].IsSelectedAll())
		return true;

	return false;
};
CDocumentContent.prototype.AddContentControl = function(nContentControlType)
{
	if (docpostype_DrawingObjects === this.CurPos.Type)
		return this.DrawingObjects.AddContentControl(nContentControlType);
	else
		return this.private_AddContentControl(nContentControlType);
};
CDocumentContent.prototype.GetAllContentControls = function(arrContentControls)
{
	if (!arrContentControls)
		arrContentControls = [];

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetAllContentControls(arrContentControls);
	}

	return arrContentControls;
};
CDocumentContent.prototype.GetMargins = function()
{
	if (this.Parent && this.Parent.GetMargins)
		return this.Parent.GetMargins();

	return {
		Top    : new CTableMeasurement(tblwidth_Mm, 0),
		Left   : new CTableMeasurement(tblwidth_Mm, 0),
		Bottom : new CTableMeasurement(tblwidth_Mm, 0),
		Right  : new CTableMeasurement(tblwidth_Mm, 0)
	};
};
CDocumentContent.prototype.IsEmptyPage = function(nCurPage)
{
	if (nCurPage < 0 || nCurPage >= this.Pages.length)
		return true;

	var nStartPos = this.Pages[nCurPage].Pos;
	var nEndPos   = this.Pages[nCurPage].EndPos;

	if (nStartPos > nEndPos)
		return true;

	if (nStartPos < nEndPos)
		return false;

	var nElementPageIndex = this.private_GetElementPageIndex(nStartPos, nCurPage, 0, 1);
	return this.Content[nStartPos].IsEmptyPage(nElementPageIndex);
};
CDocumentContent.prototype.GetParent = function()
{
	return this.Parent;
};
CDocumentContent.prototype.GetPlaceHolderObject = function()
{
	var nCurPos = this.CurPos.ContentPos;
	if (this.Selection.Use)
	{
		if (this.Selection.StartPos === this.Selection.EndPos)
			nCurPos = this.Selection.StartPos;
		else
			return null;
	}

	return this.Content[nCurPos].GetPlaceHolderObject();
};
CDocumentContent.prototype.GetAllFields = function(isUseSelection, arrFields)
{
	if (!arrFields)
		arrFields = [];

	var nStartPos = isUseSelection ?
		(this.Selection.Use ?
			(this.Selection.StartPos < this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos)
			: this.CurPos.ContentPos)
		: 0;

	var nEndPos = isUseSelection ?
		(this.Selection.Use ?
			(this.Selection.StartPos < this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos)
			: this.CurPos.ContentPos)
		: this.Content.length - 1;

	for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
	{
		this.Content[nIndex].GetAllFields(isUseSelection, arrFields);
	}

	return arrFields;
};
CDocumentContent.prototype.SetIsRecalculated = function(isRecalculated)
{
	if (this.Parent && this.Parent.SetIsRecalculated)
		this.Parent.SetIsRecalculated(isRecalculated);
};
CDocumentContent.prototype.GetPresentationField = function()
{
	var nCurPos = this.CurPos.ContentPos;
	if (this.Selection.Use)
	{
		if (this.Selection.StartPos === this.Selection.EndPos)
			nCurPos = this.Selection.StartPos;
		else
			return null;
	}

	return this.Content[nCurPos].GetPresentationField();
};
CDocumentContent.prototype.IsTableCellSelection = function()
{
	return (this.Selection.Use && this.Selection.StartPos === this.Selection.EndPos && this.Content[this.Selection.StartPos].IsTable() && this.Content[this.Selection.StartPos].IsTableCellSelection());
};
/**
 * Проверяем можно ли редактировать все контейнеры, находящиеся в данном классе
 * @returns {boolean}
 */
CDocumentContent.prototype.CanEditAllContentControls = function()
{
	var arrCC = this.GetAllContentControls();

	for (var nIndex = 0, nCount = arrCC.length; nIndex < nCount; ++nIndex)
	{
		let cc = arrCC[nIndex];
		
		cc.SkipSpecialContentControlLock(true);
		cc.SkipFillingFormModeCheck(true);
		
		let canEdit = cc.CanBeEdited();
		
		cc.SkipFillingFormModeCheck(false);
		cc.SkipSpecialContentControlLock(false);
		
		if (!canEdit)
			return false;
	}

	return true;
};
/**
 * Проверяем можно ли удалять все контейнеры, находящиеся в данном классе
 * @returns {boolean}
 */
CDocumentContent.prototype.CanDeleteAllContentControls = function()
{
	var arrCC = this.GetAllContentControls();

	for (var nIndex = 0, nCount = arrCC.length; nIndex < nCount; ++nIndex)
	{
		if (!arrCC[nIndex].CanBeDeleted())
			return false;
	}

	return true;
};
CDocumentContent.prototype.Document_Is_SelectionLocked = function(CheckType)
{
	if ( true === this.ApplyToAll )
	{
		var Count = this.Content.length;
		for ( var Index = 0; Index < Count; Index++ )
		{
			this.Content[Index].SetApplyToAll( true );
			this.Content[Index].Document_Is_SelectionLocked(CheckType);
			this.Content[Index].SetApplyToAll( false );
		}
		return;
	}
	else
	{
		if ( docpostype_DrawingObjects === this.CurPos.Type )
		{
			this.LogicDocument.DrawingObjects.documentIsSelectionLocked(CheckType);
		}
		else if ( docpostype_Content == this.CurPos.Type )
		{
			switch ( this.Selection.Flag )
			{
				case selectionflag_Common :
				{
					if ( true === this.Selection.Use )
					{
						var StartPos = ( this.Selection.StartPos > this.Selection.EndPos ? this.Selection.EndPos : this.Selection.StartPos );
						var EndPos   = ( this.Selection.StartPos > this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos );

						if ( StartPos != EndPos && AscCommon.changestype_Delete === CheckType )
							CheckType = AscCommon.changestype_Remove;

						for ( var Index = StartPos; Index <= EndPos; Index++ )
							this.Content[Index].Document_Is_SelectionLocked(CheckType);
					}
					else
					{
						var CurElement = this.Content[this.CurPos.ContentPos];

						if (AscCommon.changestype_Document_Content_Add === CheckType && CurElement.IsParagraph() && CurElement.IsCursorAtEnd() && CurElement.Lock.Is_Locked())
							AscCommon.CollaborativeEditing.Add_CheckLock(false);
						else
							this.Content[this.CurPos.ContentPos].Document_Is_SelectionLocked(CheckType);
					}

					break;
				}
				case selectionflag_Numbering:
				{
					var oNumPr = this.Selection.Data.CurPara.GetNumPr();
					if (oNumPr)
					{
						var oNum = this.GetNumbering().GetNum(oNumPr.NumId);
						oNum.IsSelectionLocked(CheckType);
					}

					this.Content[this.CurPos.ContentPos].Document_Is_SelectionLocked(CheckType);

					break;
				}
			}
		}
	}
};
CDocumentContent.prototype.CheckContentControlEditingLock = function()
{
	if (this.Parent && this.Parent.CheckContentControlEditingLock)
		this.Parent.CheckContentControlEditingLock();
};
/**
 * Оставляем один пустой параграф в содержимом
 * @returns {Paragraph}
 */
CDocumentContent.prototype.MakeSingleParagraphContent = function()
{
	if (this.Content.length <= 0)
	{
		this.AddToContent(0, new AscWord.Paragraph());
	}
	else if (this.Content.length > 1 || !this.Content[0].IsParagraph())
	{
		this.RemoveFromContent(0, this.Content.length, true);
		this.AddToContent(0, new AscWord.Paragraph());
	}

	return this.Content[0];
};
CDocumentContent.prototype.AcceptRevisionChanges = function(nType, bAll)
{
	if (docpostype_Content === this.CurPos.Type || true === bAll)
	{
		this.private_AcceptRevisionChanges(nType, bAll);
	}
	else if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.DrawingObjects.AcceptRevisionChanges(nType, bAll);
	}
};
CDocumentContent.prototype.RejectRevisionChanges = function(nType, bAll)
{
	if (docpostype_Content === this.CurPos.Type || true === bAll)
	{
		this.private_RejectRevisionChanges(nType, bAll);
	}
	else if (docpostype_DrawingObjects === this.CurPos.Type)
	{
		this.DrawingObjects.RejectRevisionChanges(nType, bAll);
	}
};
CDocumentContent.prototype.GetRevisionsChangeElement = function(SearchEngine)
{
	if (true === SearchEngine.IsFound())
		return;

	var Direction = SearchEngine.GetDirection();
	var Pos = 0;
	if (true !== SearchEngine.IsCurrentFound())
	{
		Pos = (true === this.Selection.Use ? (this.Selection.StartPos <= this.Selection.EndPos ? this.Selection.StartPos : this.Selection.EndPos) : this.CurPos.ContentPos);
	}
	else
	{
		if (Direction > 0)
		{
			Pos = 0;
		}
		else
		{
			Pos = this.Content.length - 1;
		}
	}

	this.Content[Pos].GetRevisionsChangeElement(SearchEngine);
	while (true !== SearchEngine.IsFound())
	{
		Pos = (Direction > 0 ? Pos + 1 : Pos - 1);
		if (Pos >= this.Content.length || Pos < 0)
			break;

		this.Content[Pos].GetRevisionsChangeElement(SearchEngine);
	}
};
CDocumentContent.prototype.GetAllTablesOnPage = function(nPageAbs, arrTables)
{
	if (!arrTables)
		arrTables = [];

	var nStartPos = -1;
	var nEndPos   = -2;
	for (var nCurPage = 0, nPagesCount = this.Pages.length; nCurPage < nPagesCount; ++nCurPage)
	{
		var nTempPageAbs = this.GetAbsolutePage(nCurPage);

		if (nPageAbs === nTempPageAbs)
		{
			if (-1 === nStartPos)
			{
				nStartPos = this.Pages[nCurPage].Pos;
			}

			nEndPos = this.Pages[nCurPage].EndPos;
		}
		else if (nTempPageAbs > nPageAbs)
		{
			break;
		}
	}

	for (var nCurPos = nStartPos; nCurPos <= nEndPos; ++nCurPos)
	{
		this.Content[nCurPos].GetAllTablesOnPage(nPageAbs, arrTables);
	}

	return arrTables;
};
CDocumentContent.prototype.GetFramePr = function()
{
	var oFramePr = null;
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		var oTempFramePr = this.Content[nPos].GetFramePr();
		if (!oTempFramePr)
			return null;

		if (!oFramePr)
			oFramePr = oTempFramePr;
		else if (!oFramePr.IsEqual(oTempFramePr))
			return null;
	}

	return oFramePr;
};
CDocumentContent.prototype.SetCalculatedFrame = function(oFrame)
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		this.Content[nPos].SetCalculatedFrame(oFrame);
	}
};
CDocumentContent.prototype.IsCursorInSpecialForm = function()
{
	if (this.Content.length <= 0)
		return false;

	if (this.Selection.Use)
	{
		if (0 !== this.Selection.StartPos || 0 !== this.Selection.EndPos)
			return false;
	}
	else
	{
		if (0 !== this.CurPos.ContentPos)
			return false;
	}

	if (!this.Content[0].IsParagraph())
		return false;

	return this.Content[0].IsCursorInSpecialForm();
};
CDocumentContent.prototype.GetInnerForm = function()
{
	if (this.Content.length !== 1 || !this.Content[0].IsParagraph())
		return null;

	return this.Content[0].GetInnerForm();
};
CDocumentContent.prototype.CalculateTextToTable = function(oEngine)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].CalculateTextToTable(oEngine);
	}
};
CDocumentContent.prototype.CollectSelectedReviewChanges = function(oTrackManager)
{
	var isUseSelection = this.Selection.Use;

	var nStartPos = isUseSelection ? this.Selection.StartPos : this.CurPos.ContentPos;
	var nEndPos   = isUseSelection ? this.Selection.EndPos   : this.CurPos.ContentPos;
	if (nStartPos > nEndPos)
	{
		var nTemp = nStartPos;
		nStartPos = nEndPos;
		nEndPos   = nTemp;
	}

	for (var nPos = nStartPos; nPos <= nEndPos; ++nPos)
	{
		this.Content[nPos].CollectSelectedReviewChanges(oTrackManager);
	}
};
CDocumentContent.prototype.RestartSpellCheck = function()
{
	for (let nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].RestartSpellCheck();
	}
};
//----------------------------------------------------------------------------------------------------------------------
// Search
//----------------------------------------------------------------------------------------------------------------------
CDocumentContent.prototype.Search = function(oSearchEngine, nType)
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		this.Content[nPos].Search(oSearchEngine, nType);
	}
};
CDocumentContent.prototype.GetSearchElementId = function(bNext, bCurrent)
{
	// Получим Id найденного элемента
	var Id = null;

	if ( true === bCurrent )
	{
		if ( docpostype_DrawingObjects === this.CurPos.Type )
		{
			var ParaDrawing = this.DrawingObjects.getMajorParaDrawing();

			Id = ParaDrawing.GetSearchElementId( bNext, true );
			if ( null != Id )
				return Id;

			ParaDrawing.GoToText( true !== bNext, false );
		}

		var Pos = this.CurPos.ContentPos;
		if ( true === this.Selection.Use && selectionflag_Common === this.Selection.Flag )
			Pos = ( true === bNext ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : Math.min(this.Selection.StartPos, this.Selection.EndPos) );

		if ( true === bNext )
		{
			Id = this.Content[Pos].GetSearchElementId(true, true);

			if ( null != Id )
				return Id;

			Pos++;

			var Count = this.Content.length;
			while ( Pos < Count )
			{
				Id = this.Content[Pos].GetSearchElementId(true, false);
				if ( null != Id )
					return Id;

				Pos++;
			}
		}
		else
		{
			Id = this.Content[Pos].GetSearchElementId(false, true);

			if ( null != Id )
				return Id;

			Pos--;

			while ( Pos >= 0 )
			{
				Id = this.Content[Pos].GetSearchElementId(false, false);
				if ( null != Id )
					return Id;

				Pos--;
			}
		}
	}
	else
	{
		var Count = this.Content.length;
		if ( true === bNext )
		{
			var Pos = 0;
			while ( Pos < Count )
			{
				Id = this.Content[Pos].GetSearchElementId(true, false);
				if ( null != Id )
					return Id;

				Pos++;
			}
		}
		else
		{
			var Pos = Count - 1;
			while ( Pos >= 0 )
			{
				Id = this.Content[Pos].GetSearchElementId(false, false);
				if ( null != Id )
					return Id;

				Pos--;
			}
		}
	}

	return null;
};
//----------------------------------------------------------------------------------------------------------------------
CDocumentContent.prototype.EnterText = function(value)
{
	if (undefined === value
		|| null === value
		|| (Array.isArray(value) && !value.length))
		return false;
	
	let codePoints = typeof(value) === "string" ? value.codePointsArray() : value;
	
	if (Array.isArray(codePoints))
	{
		for (let index = 0, count = codePoints.length; index < count; ++index)
		{
			let codePoint = codePoints[index];
			this.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
		}
	}
	else
	{
		let codePoint = codePoints;
		this.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
	}
	
	return true;
};
CDocumentContent.prototype.CorrectEnterText = function(oldValue, newValue, checkAsYouTypeFunc)
{
	if (undefined === oldValue
		|| null === oldValue
		|| (Array.isArray(oldValue) && !oldValue.length))
		return this.EnterText(newValue);
	
	let newCodePoints = typeof(newValue) === "string" ? newValue.codePointsArray() : newValue;
	let oldCodePoints = typeof(oldValue) === "string" ? oldValue.codePointsArray() : oldValue;
	
	if (this.IsSelectionUse())
		return false;
	
	if (!Array.isArray(oldCodePoints))
		oldCodePoints = [oldCodePoints];
	
	let paragraph = this.GetCurrentParagraph();
	if (!paragraph)
		return false;
	
	let contentPos = paragraph.GetContentPosition(false, false);
	let run, inRunPos;
	for (let index = contentPos.length - 1; index >= 0; --index)
	{
		if (contentPos[index].Class instanceof AscWord.CRun)
		{
			run      = contentPos[index].Class;
			inRunPos = contentPos[index].Position;
			break;
		}
	}
	
	if (!run)
		return false;
	
	if (!checkAsYouTypeFunc)
		checkAsYouTypeFunc = AscWord.checkAsYouTypeEnterText;
	
	if (!checkAsYouTypeFunc(run, inRunPos, oldCodePoints[oldCodePoints.length - 1]))
		return false;
	
	if (undefined === newCodePoints || null === newCodePoints)
		newCodePoints = [];
	else if (!Array.isArray(newCodePoints))
		newCodePoints = [newCodePoints];
	
	let oldText = "";
	for (let index = 0, count = oldCodePoints.length; index < count; ++index)
	{
		oldText += String.fromCodePoint(oldCodePoints[index]);
	}
	
	let state     = this.GetSelectionState();
	let startPos  = paragraph.getCurrentPos();
	let endPos    = startPos;
	
	let paraSearchPos = new CParagraphSearchPos();
	
	let maxShifts = oldCodePoints.length;
	let selectedText;
	this.StartSelectionFromCurPos();
	while (maxShifts >= 0)
	{
		paraSearchPos.Reset();
		paragraph.Get_LeftPos(paraSearchPos, endPos);
		
		if (!paraSearchPos.IsFound())
			break;
		
		endPos = paraSearchPos.GetPos().Copy();
		
		paragraph.SetSelectionContentPos(startPos, endPos, false);
		selectedText = paragraph.GetSelectedText(true);
		
		if (!selectedText || selectedText === oldText)
			break;
		
		maxShifts--;
	}
	
	if (selectedText !== oldText)
	{
		this.SetSelectionState(state);
		return false;
	}
	
	this.Remove(1, true, false, true);
	
	for (let index = 0, count = newCodePoints.length; index < count; ++index)
	{
		let codePoint = newCodePoints[index];
		this.AddToParagraph(AscCommon.IsSpace(codePoint) ? new AscWord.CRunSpace(codePoint) : new AscWord.CRunText(codePoint));
	}
	
	return true;
};

function CDocumentContentStartState(DocContent)
{
    this.Content = [];
    for(var i = 0; i < DocContent.Content.length; ++i)
    {
        this.Content.push(DocContent.Content[i]);
    }
}

function CDocumentRecalculateObject()
{
    this.StartPage = 0;

    this.Pages    = [];
    this.Content  = [];
    this.ClipInfo = [];
}

CDocumentRecalculateObject.prototype =
{
    Save : function(Doc)
    {
        this.StartPage = Doc.StartPage;
        this.Pages     = Doc.Pages;
        this.ClipInfo  = Doc.ClipInfo;

        var Content = Doc.Content;
        var Count = Content.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            this.Content[Index] = Content[Index].SaveRecalculateObject();
        }
    },

    Load : function(Doc)
    {
        Doc.StartPage = this.StartPage;
        Doc.Pages     = this.Pages;
        Doc.ClipInfo  = this.ClipInfo;

        var Count = Doc.Content.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            Doc.Content[Index].LoadRecalculateObject( this.Content[Index] );
        }
    },

	GetSummaryHeight : function()
    {
        var Height = 0;
        var PagesCount = this.Pages.length;
        for ( var Page = 0; Page < PagesCount; Page++ )
        {
            var Bounds = this.Get_PageBounds( Page );
            Height += Bounds.Bottom - Bounds.Top;
        }

        return Height;
    },

    Get_PageBounds : function(PageNum)
    {
        if ( this.Pages.length <= 0 )
            return { Top : 0, Left : 0, Right : 0, Bottom : 0 };

        if ( PageNum < 0 || PageNum > this.Pages.length )
            return this.Pages[0].Bounds;

        var Bounds = this.Pages[PageNum].Bounds;

        return Bounds;
    },

    Get_DrawingFlowPos : function(FlowPos)
    {
        var Count = this.Content.length;
        for ( var Index = 0; Index < Count; Index++ )
        {
            this.Content[Index].Get_DrawingFlowPos( FlowPos );
        }
    }

};

(function()
{
	/**
	 * Class for holding clipping rectangle
	 * @param x0
	 * @param x1
	 * @param y0
	 * @param y1
	 * @constructor
	 */
	function ClipRect(x0, x1, y0, y1)
	{
		this.X0 = x0;
		this.X1 = x1;
		this.Y0 = y0;
		this.Y1 = y1;
	}
	ClipRect.prototype.intersect = function(clipRect)
	{
		if (!clipRect)
			return;
		
		if (undefined === this.X0 || null === this.X0)
			this.X0 = clipRect.X0;
		else if (undefined !== clipRect.X0 && null !== clipRect.X0)
			this.X0 = Math.max(this.X0, clipRect.X0);
		
		if (undefined === this.X1 || null === this.X1)
			this.X1 = clipRect.X1;
		else if (undefined !== clipRect.X1 && null !== clipRect.X1)
			this.X1 = Math.min(this.X1, clipRect.X1);
		
		if (undefined === this.Y0 || null === this.Y0)
			this.Y0 = clipRect.Y0;
		else if (undefined !== clipRect.Y0 && null !== clipRect.Y0)
			this.Y0 = Math.max(this.Y0, clipRect.Y0);
		
		if (undefined === this.Y1 || null === this.Y1)
			this.Y1 = clipRect.Y1;
		else if (undefined !== clipRect.Y1 && null !== clipRect.Y1)
			this.Y1 = Math.min(this.Y1, clipRect.Y1);
	};
	ClipRect.prototype.shift = function(dx, dy)
	{
		if (undefined !== this.X0)
			this.X0 += dx;
		if (undefined !== this.X1)
			this.X1 += dx;
		if (undefined !== this.Y0)
			this.Y0 += dy;
		if (undefined !== this.Y1)
			this.Y1 += dy;
	};
	ClipRect.prototype.clone = function()
	{
		return new ClipRect(this.X0, this.X1, this.Y0, this.Y1);
	};
	ClipRect.prototype.correctX0 = function(x)
	{
		return (undefined !== this.X0 && null !== this.X0 ? Math.max(this.X0, x) : x);
	};
	ClipRect.prototype.correctX1 = function(x)
	{
		return (undefined !== this.X1 && null !== this.X1 ? Math.min(this.X1, x) : x);
	};
	ClipRect.prototype.correctY0 = function(y)
	{
		return (undefined !== this.Y0 && null !== this.Y0 ? Math.max(this.Y0, y) : y);
	};
	ClipRect.prototype.correctY1 = function(y)
	{
		return (undefined !== this.Y1 && null !== this.Y1 ? Math.min(this.Y1, y) : y);
	};
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.ClipRect = ClipRect;
})();

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CDocumentContent = CDocumentContent;
window['AscWord'].CDocumentContent = CDocumentContent;
window['AscWord'].DocumentContent = CDocumentContent;
