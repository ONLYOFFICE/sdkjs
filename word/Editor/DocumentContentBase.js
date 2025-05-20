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
var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;

/**
 * Базовый класс для работы с содержимым документа (параграфы и таблицы).
 * @constructor
 */
function CDocumentContentBase()
{
	this.Id = null;

	this.StartPage = 0; // Номер стартовой страницы в родительском классе
	this.CurPage   = 0; // Номер текущей страницы

	this.Content = [];

	this.ReindexStartPos = 0;
}
CDocumentContentBase.prototype.Get_Id = function()
{
	return this.GetId();
};
CDocumentContentBase.prototype.GetId = function()
{
	return this.Id;
};
CDocumentContentBase.prototype.GetApi = function()
{
	return Asc.editor;
};
/**
 * Получаем ссылку на основной объект документа
 * @returns {CDocument}
 */
CDocumentContentBase.prototype.GetLogicDocument = function()
{
	if (this instanceof CDocument)
		return this;

	return this.LogicDocument;
};
CDocumentContentBase.prototype.getDrawingDocument = function()
{
	let api = this.GetApi();
	return api && api.getDrawingDocument();
};
/**
 * Получаем тип активной части документа.
 * @returns {(docpostype_Content | docpostype_HdrFtr | docpostype_DrawingObjects | docpostype_Footnotes)}
 */
CDocumentContentBase.prototype.GetDocPosType = function()
{
	return this.CurPos.Type;
};
/**
 * Выставляем тип активной части документа.
 * @param {(docpostype_Content | docpostype_HdrFtr | docpostype_DrawingObjects | docpostype_Footnotes | docpostype_Endnotes)} nType
 */
CDocumentContentBase.prototype.SetDocPosType = function(nType)
{
	this.CurPos.Type = nType;

	if (this.Controller)
		this.Controller = this.getController(nType)
};
CDocumentContentBase.prototype.getController = function(type)
{
	let controller = this.LogicDocumentController;
	if (docpostype_HdrFtr === type)
		controller = this.HeaderFooterController;
	else if (docpostype_DrawingObjects === type)
		controller = this.DrawingsController;
	else if (docpostype_Footnotes === type)
		controller = this.Footnotes;
	else if (docpostype_Endnotes === type)
		controller = this.Endnotes;
	
	return controller;
};
/**
 * Обновляем индексы элементов.
 */
CDocumentContentBase.prototype.Update_ContentIndexing = function()
{
	if (-1 !== this.ReindexStartPos)
	{
		for (var Index = this.ReindexStartPos, Count = this.Content.length; Index < Count; Index++)
		{
			this.Content[Index].Index = Index;
		}

		this.ReindexStartPos = -1;
	}
};
CDocumentContentBase.prototype.UpdateContentIndexing = function()
{
	return this.Update_ContentIndexing();
};
/**
 * Получаем массив всех автофигур.
 * @param {Array} arrDrawings - Если задан массив, тогда мы дополняем данный массив и возвращаем его.
 * @returns {Array}
 */
CDocumentContentBase.prototype.GetAllDrawingObjects = function(arrDrawings)
{
	if (undefined === arrDrawings || null === arrDrawings)
		arrDrawings = [];

	if (this instanceof CDocument)
	{
		this.SectionsInfo.GetAllDrawingObjects(arrDrawings);
		this.Footnotes.GetAllDrawingObjects(arrDrawings);
	}

	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		this.Content[nPos].GetAllDrawingObjects(arrDrawings);
	}

	return arrDrawings;
};
/**
 * Получаем массив URL всех картинок в документе.
 * @param {Array} arrUrls - Если задан массив, тогда мы дополняем данный массив и возвращаем его.
 * @returns {Array}
 */
CDocumentContentBase.prototype.Get_AllImageUrls = function(arrUrls)
{
	if (undefined === arrUrls || null === arrUrls)
		arrUrls = [];

	var arrDrawings = this.GetAllDrawingObjects();
	for (var nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
	{
		var oParaDrawing = arrDrawings[nIndex];
		oParaDrawing.GraphicObj.getAllRasterImages(arrUrls);
	}

	return arrUrls;
};
/**
 * Переназначаем ссылки на картинки.
 * @param {Object} mapUrls - Мап, в котором ключ - это старая ссылка, а значение - новая.
 */
CDocumentContentBase.prototype.Reassign_ImageUrls = function(mapUrls)
{
	var arrDrawings = this.GetAllDrawingObjects();
	for (var nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
	{
		var oDrawing = arrDrawings[nIndex];
		oDrawing.Reassign_ImageUrls(mapUrls);
	}
};

/**
 * Find all SEQ complex fields with specified type
 * @param {String} sType - field type
 * @param {Array} aFields - array which accumulates complex fields
 */
CDocumentContentBase.prototype.GetAllSeqFieldsByType = function(sType, aFields)
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		this.Content[nPos].GetAllSeqFieldsByType(sType, aFields)
	}
};


/**
 * Finds a paragraph which satisfies the specific condition
 * @param {function} fCondition - represents the condition which paragraph should satisfy.
 * @param {boolean} bBackward - whether to search backward or forward.
 * @param {?number} nStartIdx - index of searching start. If it is null searching starts depends on bBackward.
 * @returns {?Paragraph}
 */
CDocumentContentBase.prototype.FindParagraph = function (fCondition, bBackward, nStartIdx)
{
	let nIdx, oElement, oResultPara = null, oContent;
	let nSearchStartIdx;
	if(bBackward)
	{
		if(nStartIdx !== null)
		{
			nSearchStartIdx = Math.min(nStartIdx, this.Content.length - 1);
		}
		else
		{
			nSearchStartIdx = this.Content.length - 1;
		}
		for(nIdx = nSearchStartIdx; nIdx >= 0; --nIdx)
		{
			oElement = this.Content[nIdx];
			if(oElement.GetType() === type_Paragraph)
			{
				if(fCondition(oElement))
				{
					oResultPara = oElement;
				}
			}
			else if(oElement.GetType() === type_Table)
			{
				oResultPara = oElement.FindParagraph(fCondition, bBackward, null);
			}
			else if(oElement.GetType() === type_BlockLevelSdt)
			{
				oContent = oElement.GetContent();
				oResultPara = oContent.FindParagraph(fCondition, bBackward, null);
			}
			if(oResultPara !== null)
			{
				return oResultPara;
			}
		}
	}
	else
	{
		if(nStartIdx !== null)
		{
			nSearchStartIdx = Math.max(nStartIdx, 0);
		}
		else
		{
			nSearchStartIdx = 0;
		}
		for(nIdx = nSearchStartIdx; nIdx < this.Content.length; ++nIdx)
		{
			oElement = this.Content[nIdx];
			if(oElement.GetType() === type_Paragraph)
			{
				if(fCondition(oElement))
				{
					oResultPara = oElement;
				}
			}
			else if(oElement.GetType() === type_Table)
			{
				oResultPara = oElement.FindParagraph(fCondition, bBackward, null);
			}
			else if(oElement.GetType() === type_BlockLevelSdt)
			{
				oContent = oElement.GetContent();
				oResultPara = oContent.FindParagraph(fCondition, bBackward, null);
			}
			if(oResultPara !== null)
			{
				return oResultPara;
			}
		}
	}
	return null;
};

/**
 * Finds a paragraph with a given style
 * @param {string} sStyleId - id of paragraph style
 * @param {boolean} bBackward - whether to search backward or forward
 * @param {?number} nStartIdx - index of searching start. If it is null searching starts depends on bBackward.
 * @returns {?Paragraph}
 */
CDocumentContentBase.prototype.FindParaWithStyle = function (sStyleId, bBackward, nStartIdx)
{
	let fCondition = function (oParagraph)
	{
		return oParagraph.GetParagraphStyle() === sStyleId;
	};
	return this.FindParagraph(fCondition, bBackward, nStartIdx);
};

/**
 * Finds a paragraph with a given outline level
 * @param {number} nOutlineLvl - outline level of paragraph
 * @param {boolean} bBackward - whether to search backward or forward
 * @param {?number} nStartIdx - index of searching start. If it is null searching starts depends on bBackward.
 * @returns {?Paragraph}
 */
CDocumentContentBase.prototype.FindParaWithOutlineLvl = function (nOutlineLvl, bBackward, nStartIdx)
{
	let fCondition = function (oParagraph) {
		return oParagraph.GetOutlineLvl() === nOutlineLvl;
	};
	return this.FindParagraph(fCondition, bBackward, nStartIdx);
};

/**
 * Находим отрезок сносок, заданный между сносками.
 * @param {?CFootEndnote} oFirstFootnote - если null, то иещм с начала документа
 * @param {?CFootEndnote} oLastFootnote - если null, то ищем до конца документа
 * @param {boolean} [isEndnotes=false] - собираем концевые сноски или нет
 */
CDocumentContentBase.prototype.GetFootnotesList = function(oFirstFootnote, oLastFootnote, isEndnotes)
{
	var oEngine = new CDocumentFootnotesRangeEngine();
	oEngine.Init(oFirstFootnote, oLastFootnote, !isEndnotes, isEndnotes);

	var arrParagraphs = this.GetAllParagraphs({OnlyMainDocument : true, All : true});
	for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
	{
		arrParagraphs[nIndex].GetFootnotesList(oEngine);
	}

	return oEngine.GetRange();
};
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Private area
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
/**
 * Сообщаем, что нужно провести переиндексацию элементов начиная с заданного.
 * @param StartPos
 */
CDocumentContentBase.prototype.private_ReindexContent = function(StartPos)
{
	if (-1 === this.ReindexStartPos || this.ReindexStartPos > StartPos)
		this.ReindexStartPos = StartPos;
};
/**
 * Специальная функия для рассчета пустого параграфа с разрывом секции.
 * @param Element
 * @param PrevElement
 * @param PageIndex
 * @param ColumnIndex
 * @param ColumnsCount
 */
CDocumentContentBase.prototype.private_RecalculateEmptySectionParagraph = function(Element, PrevElement, PageIndex, ColumnIndex, ColumnsCount)
{
	var LastVisibleBounds = PrevElement.GetLastRangeVisibleBounds();

	var ___X = LastVisibleBounds.X + LastVisibleBounds.W;
	var ___Y = LastVisibleBounds.Y;

	// Чтобы у нас знак разрыва секции рисовался красиво и где надо делаем небольшую хитрость:
	// перед пересчетом данного параграфа меняем в нем в скомпилированных настройках прилегание и
	// отступы, а после пересчета помечаем, что настройки нужно скомпилировать заново.
	var CompiledPr           = Element.Get_CompiledPr2(false).ParaPr;
	CompiledPr.Jc            = align_Left;
	CompiledPr.Ind.FirstLine = 0;
	CompiledPr.Ind.Left      = 0;
	CompiledPr.Ind.Right     = 0;

	// Делаем предел по X минимум 10 мм, чтобы всегда было видно элемент разрыва секции
	Element.Reset(___X, ___Y, Math.max(___X + 10, LastVisibleBounds.XLimit), 10000, PageIndex, ColumnIndex, ColumnsCount);
	Element.Recalculate_Page(0);

	Element.Recalc_CompiledPr();

	// Меняем насильно размер строки и страницы данного параграфа, чтобы у него границы попадания и
	// селект были ровно такие же как и у последней строки предыдущего элемента.
	Element.Pages[0].Y             = ___Y;
	Element.Lines[0].Top           = 0;
	Element.Lines[0].Y             = LastVisibleBounds.BaseLine;
	Element.Lines[0].Bottom        = LastVisibleBounds.H;
	Element.Pages[0].Bounds.Top    = ___Y;
	Element.Pages[0].Bounds.Bottom = ___Y + LastVisibleBounds.H;
};
/**
 * Передвигаем курсор (от текущего положения) к началу ссылки на сноску
 * @param isNext двигаемся вперед или назад
 * @param isCurrent находимся ли мы в текущем объекте
 * @param isStepFootnote {boolean} - ищем сноски на странице
 * @param isStepEndnote {boolean} - ищем концевые сноски
 * @returns {boolean}
 * @constructor
 */
CDocumentContentBase.prototype.GotoFootnoteRef = function(isNext, isCurrent, isStepFootnote, isStepEndnote)
{
	var nCurPos = 0;

	if (true === isCurrent)
	{
		if (true === this.Selection.Use)
			nCurPos = Math.min(this.Selection.StartPos, this.Selection.EndPos);
		else
			nCurPos = this.CurPos.ContentPos;
	}
	else
	{
		if (isNext)
			nCurPos = 0;
		else
			nCurPos = this.Content.length - 1;
	}

	if (isNext)
	{
		for (var nIndex = nCurPos, nCount = this.Content.length; nIndex < nCount; ++nIndex)
		{
			var oElement = this.Content[nIndex];
			if (oElement.GotoFootnoteRef(true, true === isCurrent && nIndex === nCurPos, isStepFootnote, isStepEndnote))
				return true;
		}
	}
	else
	{
		for (var nIndex = nCurPos; nIndex >= 0; --nIndex)
		{
			var oElement = this.Content[nIndex];
			if (oElement.GotoFootnoteRef(false, true === isCurrent && nIndex === nCurPos, isStepFootnote, isStepEndnote))
				return true;
		}
	}

	return false;
};
CDocumentContentBase.prototype.MoveCursorToNearestPos = function(oNearestPos)
{
	var oPara = oNearestPos.Paragraph;
	var oParent = oPara.Parent;
	if (oParent)
	{
		var oTopDocument = oParent.Is_TopDocument(true);
		if (oTopDocument)
			oTopDocument.RemoveSelection();
		else
			this.RemoveSelection();
	}

	oPara.Set_ParaContentPos(oNearestPos.ContentPos, true, -1, -1);
	oPara.Document_SetThisElementCurrent(true);
};
CDocumentContentBase.prototype.private_CreateNewParagraph = function()
{
	var oPara = new AscWord.Paragraph(this, this.bPresentation === true);
	oPara.Correct_Content();
	oPara.MoveCursorToStartPos(false);
	return oPara;
};
CDocumentContentBase.prototype.StopSelection = function()
{
	if (true !== this.Selection.Use)
		return;

	this.Selection.Start = false;

	if (this.Content[this.Selection.StartPos])
		this.Content[this.Selection.StartPos].StopSelection();
};
CDocumentContentBase.prototype.GetNumberingInfo = function(oNumberingEngine, oPara, oNumPr, isUseReview)
{
	if (undefined === oNumberingEngine || null === oNumberingEngine)
		oNumberingEngine = new CDocumentNumberingInfoEngine(oPara, oNumPr, this.GetNumbering());

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetNumberingInfo(oNumberingEngine);
		if (oNumberingEngine.IsStop())
			break;
	}

	if (true === isUseReview)
		return [oNumberingEngine.GetNumInfo(), oNumberingEngine.GetNumInfo(false)];

	return oNumberingEngine.GetNumInfo();
};
CDocumentContentBase.prototype.private_Remove = function(Count, isRemoveWholeElement, bRemoveOnlySelection, bOnTextAdd, isWord)
{
	let oLogicDocument = this.GetLogicDocument();
	if (this.CurPos.ContentPos < 0)
		return false;

	if (this.IsNumberingSelection())
	{
		var oPara = this.Selection.Data.CurPara;
		this.RemoveNumberingSelection();
		oPara.RemoveSelection();
		oPara.RemoveNumPr();
		oPara.Set_Ind({FirstLine : undefined, Left : undefined, Right : oPara.Pr.Ind.Right}, true);
		oPara.MoveCursorToStartPos();
		oPara.Document_SetThisElementCurrent(true);
		return true;
	}

	this.RemoveNumberingSelection();

	var isRemoveOnDrag = this.GetLogicDocument() ? this.GetLogicDocument().DragAndDropAction : false;

	var bRetValue = true;
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

		// Проверим, пустой ли селект в конечном элементе (для случая, когда конечный элемент параграф, и в нем
		// не заселекчен знак конца параграфа)
		if (StartPos !== EndPos && true === this.Content[EndPos].IsSelectionEmpty(true))
			EndPos--;

		if (true === this.IsTrackRevisions())
		{
			var _nEndPos;
			if (this.Content[EndPos].IsParagraph() && !this.Content[EndPos].IsSelectionToEnd())
				_nEndPos = EndPos - 1;
			else
				_nEndPos = EndPos;

			var oDirectParaPr = null;
			if (this.Content[StartPos].IsParagraph() && !this.Content[StartPos].IsSelectedAll())
				oDirectParaPr = this.Content[StartPos].GetDirectParaPr();

			// TODO: Сделать для таблиц
			if (oDirectParaPr)
			{
				for (var nIndex = StartPos; nIndex <= EndPos; ++nIndex)
				{
					var oElement = this.Content[nIndex + 1];

					if (oElement && oElement.IsParagraph() && (nIndex < EndPos || this.Content[nIndex].IsSelectionToEnd()))
					{
						var oPrChange   = oElement.GetDirectParaPr();
						var oReviewInfo = new AscWord.ReviewInfo();
						oReviewInfo.Update();

						oElement.SetDirectParaPr(oDirectParaPr);
						oElement.SetPrChange(oPrChange, oReviewInfo);
					}
				}
			}

			if (StartPos === EndPos && this.Content[StartPos].IsTable())
			{
				if (!this.Content[StartPos].IsCellSelection() || bOnTextAdd || Count > 0)
				{
					this.Content[StartPos].Remove(1, true, bRemoveOnlySelection, bOnTextAdd);
				}
				else if (this.Content[StartPos].IsRowSelection())
				{
					this.Content[StartPos].RemoveTableRow();
				}
				else
				{
					// В остальных ситуация мы не отслеживаем изменения
					var isLocalTrackRevisions = oLogicDocument.GetLocalTrackRevisions();
					oLogicDocument.SetLocalTrackRevisions(false);
					this.Content[StartPos].RemoveTableCells();
					oLogicDocument.SetLocalTrackRevisions(isLocalTrackRevisions);
				}
			}
			else
			{
				// Сначала проводим обычное удаление по выделению
				for (var nIndex = StartPos; nIndex <= EndPos; ++nIndex)
				{
					var oElement = this.Content[nIndex];
					if (oElement.IsTable())
						oElement.RemoveTableRow();
					else
						oElement.Remove(Count, true, bRemoveOnlySelection, bOnTextAdd);
				}
			}

			this.RemoveSelection();

			// Удаляем параграфы, если они были ранее добавлены в рецензировании этим же пользователем
			for (var nIndex = _nEndPos; nIndex >= StartPos; --nIndex)
			{
				var oElement = this.Content[nIndex];

				var nReviewType = oElement.GetReviewType();
				var oReviewInfo = oElement.GetReviewInfo();

				if (oElement.IsParagraph())
				{
					if ((reviewtype_Add === nReviewType && oReviewInfo.IsCurrentUser())
						|| (reviewtype_Remove === nReviewType && oReviewInfo.IsPrevAddedByCurrentUser()))
					{
						// Если параграф пустой, тогда удаляем параграф, если не пустой, тогда объединяем его со
						// следующим параграф. Если следующий элемент таблица, тогда ничего не делаем.
						if (oElement.IsEmpty())
						{
							this.RemoveFromContent(nIndex, 1);
						}
						else if (nIndex < this.Content.length - 1 && this.Content[nIndex + 1].IsParagraph())
						{
							oElement.Concat(this.Content[nIndex + 1]);
							this.RemoveFromContent(nIndex + 1, 1);
						}
					}
					else
					{
						if (reviewtype_Add === nReviewType)
						{
							var oNewReviewInfo = oReviewInfo.Copy();
							oNewReviewInfo.SavePrev(reviewtype_Add);
							oNewReviewInfo.Update();
							oElement.SetReviewType(reviewtype_Remove, oNewReviewInfo);
						}
						else
						{
							oElement.SetReviewType(reviewtype_Remove);
						}
					}
				}
				else if (oElement.IsTable())
				{
					// После принятия изменений у нас могла остаться пустая таблица, такую мы должны удалить
					if (oElement.GetRowsCount() <= 0)
					{
						this.RemoveFromContent(nIndex, 1, false);
					}
				}
			}
			
			if (StartPos < this.Content.length)
			{
				this.CurPos.ContentPos = StartPos;
			}
			else if (this.Content.length > 0)
			{
				this.CurPos.ContentPos = this.Content.length - 1;
				this.Content[this.CurPos.ContentPos].MoveCursorToEndPos();
			}
			else
			{
				// This shouldn't happen
				this.CurPos.ContentPos = 0;
			}
		}
		else
		{
			this.Selection.Use      = false;
			this.Selection.StartPos = 0;
			this.Selection.EndPos   = 0;

			if (StartPos !== EndPos)
			{
				var StartType = this.Content[StartPos].GetType();
				var EndType   = this.Content[EndPos].GetType();

				var bStartEmpty = false, bEndEmpty = false;
				if (type_Paragraph === StartType || type_BlockLevelSdt === StartType)
				{
					this.Content[StartPos].Remove(1, true);
					bStartEmpty = this.Content[StartPos].IsEmpty()
				}
				else if (type_Table === StartType)
				{
					bStartEmpty = !(this.Content[StartPos].Row_Remove2());
				}

				if (type_Paragraph === EndType || type_BlockLevelSdt === EndType)
				{
					this.Content[EndPos].Remove(1, true);
					bEndEmpty = this.Content[EndPos].IsEmpty()
				}
				else if (type_Table === EndType)
				{
					bEndEmpty = !(this.Content[EndPos].Row_Remove2());
				}

				if (!bStartEmpty && !bEndEmpty)
				{
					this.Internal_Content_Remove(StartPos + 1, EndPos - StartPos - 1);
					this.CurPos.ContentPos = StartPos;

					if (!isRemoveOnDrag)
					{
						if (oLogicDocument
							&& oLogicDocument.ConcatParagraphsOnRemove
							&& StartPos < this.Content.length - 1
							&& this.Content[StartPos].IsParagraph()
							&& this.Content[StartPos + 1].IsParagraph())
						{
							this.CurPos.ContentPos = StartPos + 1;
							this.Content[StartPos + 1].ConcatBefore(this.Content[StartPos], 0);
							this.RemoveFromContent(StartPos, 1);
							this.CurPos.ContentPos = StartPos;
						}
						else if (type_Paragraph === StartType && type_Paragraph === EndType && true === bOnTextAdd)
						{
							// Встаем в конец параграфа и удаляем 1 элемент (чтобы соединить параграфы)
							this.Content[StartPos].MoveCursorToEndPos(false, false);
							this.Remove(1, true);
						}
						else
						{
							if (true === bOnTextAdd && type_Paragraph !== this.Content[StartPos + 1].GetType() && type_Paragraph !== this.Content[StartPos].GetType())
							{
								this.Internal_Content_Add(StartPos + 1, this.private_CreateNewParagraph());
								this.CurPos.ContentPos = StartPos + 1;
								this.Content[StartPos + 1].MoveCursorToStartPos(false);
							}
							else
							{
								this.CurPos.ContentPos = StartPos;
								this.Content[StartPos].MoveCursorToEndPos(false, false);
							}
						}
					}
				}
				else if (!bStartEmpty)
				{
					if (true === bOnTextAdd && type_Table === StartType)
					{
						if (EndType !== type_Paragraph)
							this.Internal_Content_Remove(StartPos + 1, EndPos - StartPos);
						else
							this.Internal_Content_Remove(StartPos + 1, EndPos - StartPos - 1);

						if (type_Table === this.Content[StartPos + 1].GetType() && true === bOnTextAdd)
							this.Internal_Content_Add(StartPos + 1, this.private_CreateNewParagraph());

						this.CurPos.ContentPos = StartPos + 1;
						this.Content[StartPos + 1].MoveCursorToStartPos(false);
					}
					else
					{
						this.Internal_Content_Remove(StartPos + 1, EndPos - StartPos);

						if (type_Table === StartType)
						{
							// У нас обязательно есть элемент после таблицы (либо снова таблица, либо параграф)
							// Встаем в начало следующего элемента.
							this.CurPos.ContentPos = StartPos + 1;
							this.Content[StartPos + 1].MoveCursorToStartPos(false);
						}
						else if (oLogicDocument
							&& oLogicDocument.ConcatParagraphsOnRemove
							&& StartPos < this.Content.length - 1
							&& this.Content[StartPos].IsParagraph()
							&& this.Content[StartPos + 1].IsParagraph())
						{
							this.CurPos.ContentPos = StartPos + 1;
							this.Content[StartPos + 1].ConcatBefore(this.Content[StartPos], 0);
							this.RemoveFromContent(StartPos, 1);
							this.CurPos.ContentPos = StartPos;
						}
						else
						{
							this.CurPos.ContentPos = StartPos;
							this.Content[StartPos].MoveCursorToEndPos(false, false);
						}
					}
				}
				else if (!bEndEmpty)
				{
					this.Internal_Content_Remove(StartPos, EndPos - StartPos);

					if (type_Table === this.Content[StartPos].GetType() && true === bOnTextAdd)
						this.Internal_Content_Add(StartPos, this.private_CreateNewParagraph());

					this.CurPos.ContentPos = StartPos;
					this.Content[StartPos].MoveCursorToStartPos(false);
				}
				else
				{
					if (true === bOnTextAdd && !isRemoveOnDrag)
					{
						// Удаляем весь промежуточный контент, начальный элемент и конечный элемент, если это
						// таблица, поскольку таблица не может быть последним элементом в документе удаляем без проверок.
						if (type_Paragraph !== EndType)
							this.Internal_Content_Remove(StartPos, EndPos - StartPos + 1);
						else
							this.Internal_Content_Remove(StartPos, EndPos - StartPos);

						if (type_Table === this.Content[StartPos].GetType() && true === bOnTextAdd)
							this.Internal_Content_Add(StartPos, this.private_CreateNewParagraph());

						this.CurPos.ContentPos = StartPos;
						this.Content[StartPos].MoveCursorToStartPos();
					}
					else
					{
						// Удаляем весь промежуточный контент, начальный и конечный параграфы
						// При таком удалении надо убедиться, что в документе останется хотя бы один элемент
						if (0 === StartPos && (EndPos - StartPos + 1) >= this.Content.length)
						{
							this.Internal_Content_Add(0, this.private_CreateNewParagraph());
							this.Internal_Content_Remove(1, this.Content.length - 1);
							bRetValue = false;
						}
						else
						{
							this.Internal_Content_Remove(StartPos, EndPos - StartPos + 1);
						}

						// Выставляем текущую позицию
						if (StartPos >= this.Content.length)
						{
							this.CurPos.ContentPos = this.Content.length - 1;
							this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false, false);
						}
						else
						{
							this.CurPos.ContentPos = StartPos;
							this.Content[StartPos].MoveCursorToStartPos(false);
						}
					}
				}
			}
			else
			{
				let isParagraphMarkRemove = this.Content[StartPos].IsParagraph() && this.Content[StartPos].IsSelectedOnlyParagraphMark();

				this.CurPos.ContentPos = StartPos;
				if (this.Content[StartPos].IsTable()
					&& true === this.Content[StartPos].IsCellSelection()
					&& ((!bOnTextAdd && Count < 0) || isRemoveOnDrag))
				{
					this.RemoveTableCells();
				}
				else if (false === this.Content[StartPos].Remove(Count, isRemoveWholeElement, bRemoveOnlySelection, bOnTextAdd))
				{
					if ((!bOnTextAdd || isRemoveOnDrag) && (isParagraphMarkRemove || ((isRemoveOnDrag || Count > 0 || StartPos < this.Content.length - 1) && this.Content[StartPos].IsEmpty())))
					{
						// В ворде параграфы объединяются только когда у них все настройки совпадают.
						// (почему то при изменении и обратном изменении настроек параграфы перестают объединятся)
						// Пока у нас параграфы будут объединяться всегда и настройки будут браться из первого
						// параграфа, кроме случая, когда первый параграф полностью удаляется.

						if (this.Content.length > 1
							&& (true === this.Content[StartPos].IsEmpty()
							|| (type_BlockLevelSdt === this.Content[StartPos].GetType() && this.Content[StartPos].IsPlaceHolder() && (!this.GetLogicDocument() || !this.GetLogicDocument().IsFillingFormMode()))))
						{
							this.Internal_Content_Remove(StartPos, 1);

							// Выставляем текущую позицию
							if (StartPos >= this.Content.length)
							{
								// Документ не должен заканчиваться таблицей, поэтому здесь проверку не делаем
								this.CurPos.ContentPos = this.Content.length - 1;
								this.Content[this.CurPos.ContentPos].MoveCursorToEndPos(false, false);
							}
							else
							{
								this.CurPos.ContentPos = StartPos;
								this.Content[StartPos].MoveCursorToStartPos(false);
							}
						}
						else if (this.CurPos.ContentPos < this.Content.length - 1 && type_Paragraph === this.Content[this.CurPos.ContentPos].GetType() && type_Paragraph === this.Content[this.CurPos.ContentPos + 1].GetType())
						{
							// Соединяем текущий и предыдущий параграфы
							this.Content[StartPos].Concat(this.Content[StartPos + 1]);
							this.Internal_Content_Remove(StartPos + 1, 1);
						}
						else if (StartPos < this.Content.length - 1 && this.Content[StartPos].IsParagraph() && this.Content[StartPos + 1].IsTable())
						{
							var oCurPara        = this.Content[StartPos];
							var oFirstParagraph = this.Content[StartPos + 1].GetFirstParagraph();

							oFirstParagraph.MoveCursorToStartPos();
							oFirstParagraph.ConcatBefore(oCurPara);
							this.RemoveFromContent(StartPos, 1);
							var oState = oFirstParagraph.SaveSelectionState();
							this.Content[StartPos].MoveCursorToStartPos(false);
							oFirstParagraph.LoadSelectionState(oState);
							this.CurPos.ContentPos = StartPos;
						}
						else if (this.Content.length === 1 && true === this.Content[0].IsEmpty())
						{
							if (Count > 0)
							{
								this.Internal_Content_Add(0, this.private_CreateNewParagraph());
								this.Internal_Content_Remove(1, this.Content.length - 1);
							}

							bRetValue = false;
						}
					}
				}
			}
		}
	}
	else
	{
		if (true === bRemoveOnlySelection || true === bOnTextAdd)
			return true;

		var nCurContentPos = this.CurPos.ContentPos;
		if (this.Content[nCurContentPos].IsParagraph())
		{
			if (false === this.Content[nCurContentPos].Remove(Count, isRemoveWholeElement, false, false, isWord))
			{
				if (Count < 0)
				{
					if (nCurContentPos > 0 && this.Content[nCurContentPos - 1].IsParagraph())
					{
						var CurrFramePr = this.Content[nCurContentPos].Get_FramePr();
						var PrevFramePr = this.Content[nCurContentPos - 1].Get_FramePr();

						if ((undefined === CurrFramePr && undefined === PrevFramePr) || (undefined !== CurrFramePr && undefined !== PrevFramePr && true === CurrFramePr.Compare(PrevFramePr)))
						{
							if (true === this.IsTrackRevisions()
								&& (reviewtype_Add !== this.Content[nCurContentPos - 1].GetReviewType() || !this.Content[nCurContentPos - 1].GetReviewInfo().IsCurrentUser())
								&& (reviewtype_Remove !== this.Content[nCurContentPos - 1].GetReviewType() || !this.Content[nCurContentPos - 1].GetReviewInfo().IsPrevAddedByCurrentUser()))
							{
								if (reviewtype_Add === this.Content[nCurContentPos - 1].GetReviewType())
								{
									var oReviewInfo = this.Content[nCurContentPos - 1].GetReviewInfo().Copy();
									oReviewInfo.SavePrev(reviewtype_Add);
									oReviewInfo.Update();
									this.Content[nCurContentPos - 1].SetReviewTypeWithInfo(reviewtype_Remove, oReviewInfo);
								}
								else
								{
									this.Content[nCurContentPos - 1].SetReviewType(reviewtype_Remove);
								}

								nCurContentPos--;
								this.Content[nCurContentPos].MoveCursorToEndPos(false, false);

								if (this.Content[nCurContentPos].IsParagraph())
								{
									var oParaPr   = this.Content[nCurContentPos].GetDirectParaPr();
									var oPrChange = this.Content[nCurContentPos + 1].GetDirectParaPr();
									var oReviewInfo = new AscWord.ReviewInfo();
									oReviewInfo.Update();

									this.Content[nCurContentPos + 1].SetDirectParaPr(oParaPr);
									this.Content[nCurContentPos + 1].SetPrChange(oPrChange, oReviewInfo);
								}
							}
							else
							{
								if (true === this.Content[nCurContentPos - 1].IsEmpty() && undefined === this.Content[nCurContentPos - 1].GetNumPr())
								{
									// Просто удаляем предыдущий параграф
									this.Internal_Content_Remove(nCurContentPos - 1, 1);
									nCurContentPos--;
									this.Content[nCurContentPos].MoveCursorToStartPos(false);
								}
								else
								{
									// Соединяем текущий и предыдущий параграфы
									var Prev = this.Content[nCurContentPos - 1];

									// Смещаемся в конец до объединения параграфов, чтобы курсор стоял в месте
									// соединения.
									Prev.MoveCursorToEndPos(false, false);

									// Запоминаем новую позицию курсора, после совмещения параграфов
									Prev.Concat(this.Content[nCurContentPos]);
									this.Internal_Content_Remove(nCurContentPos, 1);
									nCurContentPos--;
								}
							}
						}
					}
					else if (nCurContentPos > 0 && this.Content[nCurContentPos - 1].IsBlockLevelSdt())
					{
						if (this.Content[nCurContentPos].IsEmpty()
							&& (nCurContentPos < this.Content.length - 1 || !(this instanceof AscWord.CDocument)))
						{
							if (this.IsTrackRevisions())
							{
								let lastElementIndex = this.Content[nCurContentPos - 1].GetElementsCount() - 1;
								let lastElement = lastElementIndex >= 0 ? this.Content[nCurContentPos - 1].GetElement(lastElementIndex) : null;
								if (lastElement && lastElement.IsParagraph())
								{
									lastElement.SetReviewType(reviewtype_Remove);
								}
							}
							else
							{
								this.RemoveFromContent(nCurContentPos);
							}
						}
						
						--nCurContentPos;
						this.Content[nCurContentPos].MoveCursorToEndPos(false);
					}
					else if (0 === nCurContentPos)
					{
						bRetValue = false;
					}
				}
				else if (Count > 0)
				{
					if (nCurContentPos < this.Content.length - 1 && this.Content[nCurContentPos + 1].IsParagraph())
					{
						var CurrFramePr = this.Content[nCurContentPos].Get_FramePr();
						var NextFramePr = this.Content[nCurContentPos + 1].Get_FramePr();

						if ((undefined === CurrFramePr && undefined === NextFramePr) || ( undefined !== CurrFramePr && undefined !== NextFramePr && true === CurrFramePr.Compare(NextFramePr) ))
						{
							if (true === this.IsTrackRevisions()
								&& (reviewtype_Add !== this.Content[nCurContentPos].GetReviewType() || !this.Content[nCurContentPos].GetReviewInfo().IsCurrentUser())
								&& (reviewtype_Remove !== this.Content[nCurContentPos].GetReviewType() || !this.Content[nCurContentPos].GetReviewInfo().IsPrevAddedByCurrentUser()))
							{
								if (reviewtype_Add === this.Content[nCurContentPos].GetReviewType())
								{
									var oReviewInfo = this.Content[nCurContentPos].GetReviewInfo().Copy();
									oReviewInfo.SavePrev(reviewtype_Add);
									oReviewInfo.Update();
									this.Content[nCurContentPos].SetReviewTypeWithInfo(reviewtype_Remove, oReviewInfo);
								}
								else
								{
									this.Content[nCurContentPos].SetReviewType(reviewtype_Remove);
								}


								this.Content[nCurContentPos].SetReviewType(reviewtype_Remove);
								nCurContentPos++;
								this.Content[nCurContentPos].MoveCursorToStartPos(false);

								if (this.Content[nCurContentPos - 1].IsParagraph())
								{
									var oParaPr   = this.Content[nCurContentPos - 1].GetDirectParaPr();
									var oPrChange = this.Content[nCurContentPos].GetDirectParaPr();
									var oReviewInfo = new AscWord.ReviewInfo();
									oReviewInfo.Update();

									this.Content[nCurContentPos].SetDirectParaPr(oParaPr);
									this.Content[nCurContentPos].SetPrChange(oPrChange, oReviewInfo);
								}
							}
							else
							{
								if (true === this.Content[nCurContentPos].IsEmpty())
								{
									// Просто удаляем текущий параграф
									this.Internal_Content_Remove(nCurContentPos, 1);
									this.Content[nCurContentPos].MoveCursorToStartPos(false);
								}
								else
								{
									// Соединяем текущий и следующий параграфы
									var Cur = this.Content[nCurContentPos];
									Cur.Concat(this.Content[nCurContentPos + 1]);
									this.Internal_Content_Remove(nCurContentPos + 1, 1);
								}
							}
						}
					}
					else if (nCurContentPos < this.Content.length - 1 && this.Content[nCurContentPos + 1].IsTable())
					{
						if (this.Content[nCurContentPos].IsEmpty())
						{
							this.RemoveFromContent(nCurContentPos, 1);
							this.CurPos.ContentPos = nCurContentPos;
							this.Content[nCurContentPos].MoveCursorToStartPos(false);
						}
						else
						{
							var oCurPara        = this.Content[nCurContentPos];
							var oFirstParagraph = this.Content[nCurContentPos + 1].GetFirstParagraph();

							oFirstParagraph.MoveCursorToStartPos();
							oFirstParagraph.ConcatBefore(oCurPara);
							this.RemoveFromContent(nCurContentPos, 1);
							var oState = oFirstParagraph.SaveSelectionState();
							this.Content[nCurContentPos].MoveCursorToStartPos(false);
							oFirstParagraph.LoadSelectionState(oState);
						}
					}
					else if (nCurContentPos < this.Content.length - 1 && this.Content[nCurContentPos + 1].IsBlockLevelSdt())
					{
						if (this.Content[nCurContentPos].IsEmpty())
						{
							let reviewType = this.Content[nCurContentPos].GetReviewType();
							let reviewInfo = this.Content[nCurContentPos].GetReviewInfo();
							if (this.IsTrackRevisions()
								&& (reviewtype_Add !== reviewType || !reviewInfo.IsCurrentUser()
								&& (reviewtype_Remove !== reviewInfo || !reviewInfo.IsPrevAddedByCurrentUser())))
							{
								if (reviewtype_Add === reviewType)
								{
									let newReviewInfo = reviewInfo.Copy();
									newReviewInfo.SavePrev(reviewtype_Add);
									newReviewInfo.Update();
									this.Content[nCurContentPos].SetReviewType(reviewtype_Remove, newReviewInfo);
								}
								else
								{
									this.Content[nCurContentPos].SetReviewType(reviewtype_Remove);
								}
								++nCurContentPos;
							}
							else
							{
								this.RemoveFromContent(nCurContentPos);
							}
						}
						else
						{
							++nCurContentPos;
						}
						
						this.Content[nCurContentPos].MoveCursorToStartPos(false);
					}
					else if (true == this.Content[nCurContentPos].IsEmpty() && nCurContentPos == this.Content.length - 1 && nCurContentPos != 0 && type_Paragraph === this.Content[nCurContentPos - 1].GetType())
					{
						if (this.IsTrackRevisions())
						{
							bRetValue = false;
						}
						else
						{
							// Если данный параграф пустой, последний, не единственный и идущий перед
							// ним элемент не таблица, удаляем его
							this.Internal_Content_Remove(nCurContentPos, 1);
							nCurContentPos--;
							this.Content[nCurContentPos].MoveCursorToEndPos(false, false);
						}
					}
					else if (nCurContentPos === this.Content.length - 1)
					{
						bRetValue = false;
					}
				}
			}

			var Item = this.Content[nCurContentPos];
			if (type_Paragraph === Item.GetType())
			{
				Item.CurPos.RealX = Item.CurPos.X;
				Item.CurPos.RealY = Item.CurPos.Y;
			}
		}
		else if (this.Content[nCurContentPos].IsBlockLevelSdt())
		{
			if (false === this.Content[nCurContentPos].Remove(Count, isRemoveWholeElement, false, false, isWord))
			{
				if (oLogicDocument && true === oLogicDocument.IsFillingFormMode())
				{
					if (Count < 0 && nCurContentPos > 0)
						this.Content[nCurContentPos].MoveCursorToStartPos(false);
					else
						this.Content[nCurContentPos].MoveCursorToEndPos(false);
				}
				else
				{
					if (this.Content[nCurContentPos].IsEmpty())
					{
						if (this.Content[nCurContentPos].CanBeDeleted())
						{
							this.RemoveFromContent(nCurContentPos, 1);

							if ((Count < 0 && nCurContentPos > 0) || nCurContentPos >= this.Content.length)
							{
								nCurContentPos--;
								this.Content[nCurContentPos].MoveCursorToEndPos(false);
							}
							else
							{
								this.Content[nCurContentPos].MoveCursorToStartPos(false);
							}
						}
						else
						{
							if (Count < 0)
							{
								if (nCurContentPos > 0)
								{
									nCurContentPos--;
									this.Content[nCurContentPos].MoveCursorToEndPos(false);
								}
							}
							else
							{
								if (nCurContentPos < this.Content.length - 1)
								{
									nCurContentPos++;
									this.Content[nCurContentPos].MoveCursorToStartPos(false);
								}
							}
						}
					}
					else
					{
						if (Count < 0)
						{
							if (nCurContentPos > 0)
							{
								nCurContentPos--;
								this.Content[nCurContentPos].MoveCursorToEndPos(false);
							}
						}
						else
						{
							if (nCurContentPos < this.Content.length - 1)
							{
								nCurContentPos++;
								this.Content[nCurContentPos].MoveCursorToStartPos(false);
							}
						}
					}
				}
			}
		}
		else
		{
			this.Content[nCurContentPos].Remove(Count, isRemoveWholeElement, false, false, isWord);
		}

		this.CurPos.ContentPos = nCurContentPos;
	}

	return bRetValue;
};
CDocumentContentBase.prototype.IsBlockLevelSdtContent = function()
{
	return false;
};
CDocumentContentBase.prototype.IsBlockLevelSdtFirstOnNewPage = function()
{
	return false;
};
CDocumentContentBase.prototype.private_AddContentControl = function(nContentControlType)
{
	var oElement = this.Content[this.CurPos.ContentPos];

	if (c_oAscSdtLevelType.Block === nContentControlType)
	{
		if (true === this.IsSelectionUse())
		{
			if (this.Selection.StartPos === this.Selection.EndPos
				&& ((type_BlockLevelSdt === this.Content[this.Selection.StartPos].GetType()
				&& true !== this.Content[this.Selection.StartPos].IsSelectedAll())
				|| (type_Table === this.Content[this.Selection.StartPos].GetType()
				&& !this.Content[this.Selection.StartPos].IsCellSelection())))
			{
				return this.Content[this.Selection.StartPos].AddContentControl(nContentControlType);
			}
			else
			{
				var oSdt = new CBlockLevelSdt(editor.WordControl.m_oLogicDocument, this);
				oSdt.SetDefaultTextPr(this.GetDirectTextPr());
				oSdt.SetPlaceholder(c_oAscDefaultPlaceholderName.Text);

				var oLogicDocument = this instanceof CDocument ? this : this.LogicDocument;
				let preventPreDelete = oLogicDocument.PreventPreDelete;
				oLogicDocument.RemoveCommentsOnPreDelete = false;
				oLogicDocument.PreventPreDelete = true;

				var nStartPos = this.Selection.StartPos;
				var nEndPos   = this.Selection.EndPos;
				if (nEndPos < nStartPos)
				{
					nEndPos   = this.Selection.StartPos;
					nStartPos = this.Selection.EndPos;
				}
				
				for (let nIndex = nEndPos; nIndex >= nStartPos; --nIndex)
				{
					let oElement = this.Content[nIndex];
					oSdt.Content.Add_ToContent(0, oElement);
					this.Remove_FromContent(nIndex, 1);
					oElement.SelectAll(1);
				}

				oSdt.Content.Remove_FromContent(oSdt.Content.GetElementsCount() - 1, 1);
				oSdt.Content.Selection.Use      = true;
				oSdt.Content.Selection.StartPos = 0;
				oSdt.Content.Selection.EndPos   = oSdt.Content.GetElementsCount() - 1;

				this.Add_ToContent(nStartPos, oSdt);
				this.Selection.StartPos = nStartPos;
				this.Selection.EndPos   = nStartPos;
				this.CurPos.ContentPos  = nStartPos;

				oLogicDocument.RemoveCommentsOnPreDelete = true;
				oLogicDocument.PreventPreDelete = preventPreDelete;
				return oSdt;
			}
		}
		else
		{
			if (type_Paragraph === oElement.GetType())
			{
				var oSdt = new CBlockLevelSdt(editor.WordControl.m_oLogicDocument, this);
				oSdt.SetDefaultTextPr(this.GetDirectTextPr());
				oSdt.SetPlaceholder(c_oAscDefaultPlaceholderName.Text);
				oSdt.ReplaceContentWithPlaceHolder(false);

				var nContentPos = this.CurPos.ContentPos;
				if (oElement.IsCursorAtBegin())
				{
					this.AddToContent(nContentPos, oSdt);
					this.CurPos.ContentPos = nContentPos;
				}
				else if (oElement.IsCursorAtEnd())
				{
					this.AddToContent(nContentPos + 1, oSdt);
					this.CurPos.ContentPos = nContentPos + 1;
				}
				else
				{
					var oNewParagraph = new AscWord.Paragraph();
					oElement.Split(oNewParagraph);

					this.AddToContent(nContentPos + 1, oNewParagraph);
					this.AddToContent(nContentPos + 1, oSdt);
					this.CurPos.ContentPos = nContentPos + 1;
				}
				oSdt.MoveCursorToStartPos(false);
				return oSdt;
			}
			else
			{
				return oElement.AddContentControl(nContentControlType);
			}
		}
	}
	else
	{
		return oElement.AddContentControl(nContentControlType);
	}
};
CDocumentContentBase.prototype.RecalculateAllTables = function()
{
	for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
	{
		var Item = this.Content[nPos];
		Item.RecalculateAllTables();
	}
};
CDocumentContentBase.prototype.GetLastRangeVisibleBounds = function()
{
	if (this.Content.length <= 0)
		return {
			X        : 0,
			Y        : 0,
			W        : 0,
			H        : 0,
			BaseLine : 0,
			XLimit   : 0
		};

	return this.Content[this.Content.length - 1].GetLastRangeVisibleBounds();
};
CDocumentContentBase.prototype.FindNextFillingForm = function(isNext, isCurrent, isStart)
{
	var nCurPos = this.Selection.Use === true ? this.Selection.StartPos : this.CurPos.ContentPos;

	var nStartPos = 0;
	var nEndPos   = this.Content.length - 1;

	if (isCurrent)
	{
		if (isStart)
		{
			nStartPos = nCurPos;
			nEndPos   = isNext ? this.Content.length - 1 : 0;
		}
		else
		{
			nStartPos = isNext ? 0 : this.Content.length - 1;
			nEndPos   = nCurPos;
		}
	}
	else
	{
		if (isNext)
		{
			nStartPos = 0;
			nEndPos   = this.Content.length - 1;
		}
		else
		{
			nStartPos = this.Content.length - 1;
			nEndPos   = 0;
		}
	}

	if (isNext)
	{
		for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
		{
			var oRes = this.Content[nIndex].FindNextFillingForm(true, isCurrent && nIndex === nCurPos, isStart);
			if (oRes)
				return oRes;
		}
	}
	else
	{
		for (var nIndex = nStartPos; nIndex >= nEndPos; --nIndex)
		{
			var oRes = this.Content[nIndex].FindNextFillingForm(false, isCurrent && nIndex === nCurPos, isStart);
			if (oRes)
				return oRes;

		}
	}

	return null;
};
/**
 * Данный запрос может прийти из внутреннего элемента(параграф, таблица), чтобы узнать происходил ли выделение в
 * пределах одного элеменета.
 * @returns {boolean}
 */
CDocumentContentBase.prototype.IsSelectedSingleElement = function()
{
	return (true === this.Selection.Use && docpostype_Content === this.GetDocPosType() && this.Selection.Flag === selectionflag_Common && this.Selection.StartPos === this.Selection.EndPos)
};
CDocumentContentBase.prototype.GetOutlineParagraphs = function(arrOutline, oPr)
{
	if (!arrOutline)
		arrOutline = [];

	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetOutlineParagraphs(arrOutline, oPr);
	}

	return arrOutline;
};
/**
 * Обновляем список закладок
 * @param oManager
 */
CDocumentContentBase.prototype.UpdateBookmarks = function(oManager)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].UpdateBookmarks(oManager);
	}
};
/**
 * @param StartDocPos
 * @param EndDocPos
 */
CDocumentContentBase.prototype.SetSelectionByContentPositions = function(StartDocPos, EndDocPos)
{
	this.RemoveSelection();

	this.Selection.Use   = true;
	this.Selection.Start = false;
	this.Selection.Flag  = selectionflag_Common;

	this.SetContentSelection(StartDocPos, EndDocPos, 0, 0, 0);

	if (this.Parent && this.LogicDocument)
	{
		this.Parent.Set_CurrentElement(false, this.Get_StartPage_Absolute(), this);
		this.LogicDocument.Selection.Use   = true;
		this.LogicDocument.Selection.Start = false;
	}
};
/**
 * Получем количество элементов
 * @return {number}
 */
CDocumentContentBase.prototype.GetElementsCount = function()
{
	return this.Content.length;
};
/**
 * Получаем элемент по заданной позици
 * @param nIndex
 * @returns {?CDocumentContentElementBase}
 */
CDocumentContentBase.prototype.GetElement = function(nIndex)
{
	if (this.Content[nIndex])
		return this.Content[nIndex];

	return null;
};
/**
 * Добавляем новый элемент (с записью в историю)
 * @param nPos
 * @param oItem
 * @param {boolean} [isCorrectContent=true]
 */
CDocumentContentBase.prototype.AddToContent = function(nPos, oItem, isCorrectContent)
{
	this.Add_ToContent(nPos, oItem, isCorrectContent);
};
/**
 * Добавляем элемент в конец
 * @param oItem
 * @param isCorrectContent {boolean}
 */
CDocumentContentBase.prototype.PushToContent = function(oItem, isCorrectContent)
{
	return this.AddToContent(this.GetElementsCount(), oItem, isCorrectContent);
};
/**
 * Удаляем заданное количество элементов (с записью в историю)
 * @param {number} nPos
 * @param {number} [nCount=1]
 * @param {boolean} [isCorrectContent=true]
 */
CDocumentContentBase.prototype.RemoveFromContent = function(nPos, nCount, isCorrectContent)
{
	if (undefined === nCount || null === nCount)
		nCount = 1;

	this.Remove_FromContent(nPos, nCount, isCorrectContent);
};
/**
 * Меняет содержимое (с записью в историю)
 * @param {array} aContent
 */
CDocumentContentBase.prototype.ReplaceContent = function(aContent)
{
	var i = 0;
	for (var i = 0; i < aContent.length; ++i) {
		this.Add_ToContent(i, aContent[i]);
	}
	this.Remove_FromContent(i, this.Content.length - aContent.length, false);
};
/**
 * Получаем текущий TableOfContents, это может быть просто поле или поле вместе с оберткой Sdt
 * @param isUnique ищем с параметром Unique = true
 * @param isCheckFields Проверять ли TableOfContents, заданные через сложные поля
 * @returns {CComplexField | CBlockLevelSdt | null}
 */
CDocumentContentBase.prototype.GetTableOfContents = function(isUnique, isCheckFields)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		var oResult = this.Content[nIndex].GetTableOfContents(isUnique, isCheckFields);
		if (oResult)
			return oResult;
	}

	return null;
};
/**
 * Get all tables of figures inside
 * @param arrComplexFields
 */
CDocumentContentBase.prototype.GetTablesOfFigures = function(arrComplexFields)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetTablesOfFigures(arrComplexFields);
	}
};
/**
 * Добавляем заданный текст в текущей позиции
 * @param {String} sText
 * @param {AscWord.CTextPr} textPr
 */
CDocumentContentBase.prototype.AddText = function(sText, textPr)
{
	if (this.IsSelectionUse())
		this.Remove(1, true, false, true, false);

	var oParagraph = this.GetCurrentParagraph();
	if (!oParagraph)
		return;

	if (!textPr)
		textPr = oParagraph.GetDirectTextPr();
	
	var oRun = new ParaRun(oParagraph);
	
	if (textPr)
		oRun.SetPr(textPr);
	
	oRun.AddText(sText);
	oParagraph.Add(oRun);
};
/**
 * Проверяем находимся ли мы заголовке хоть какой-либо таблицы
 * @returns {boolean}
 */
CDocumentContentBase.prototype.IsTableHeader = function()
{
	return false;
};
/**
 * Получаем родительский класс
 */
CDocumentContentBase.prototype.GetParent = function()
{
	return null;
};
/**
 * Получаем последний параграф в данном контенте, если последний элемент не параграф, то запрашиваем у него
 * @returns {?Paragraph}
 */
CDocumentContentBase.prototype.GetLastParagraph = function()
{
	if (this.Content.length <= 0)
		return null;

	return this.Content[this.Content.length - 1].GetLastParagraph();
};
/**
 * Получаем первый параграф в данном контенте, если первый элемент не параграф, то запрашиваем у него
 * @returns {?Paragraph}
 * @constructor
 */
CDocumentContentBase.prototype.GetFirstParagraph = function()
{
	if (this.Content.length <= 0)
		return null;

	return this.Content[0].GetFirstParagraph();
};
/**
 * Получаем параграф, следующий за данным элементом
 * @returns {?Paragraph}
 */
CDocumentContentBase.prototype.GetNextParagraph = function()
{
	var oParent = this.GetParent();
	if (oParent && oParent.GetNextParagraph)
		return oParent.GetNextParagraph();

	return null;
};
/**
 * Получаем параграф, идущий перед данным элементом
 * @returns {?Paragraph}
 */
CDocumentContentBase.prototype.GetPrevParagraph = function()
{
	var oParent = this.GetParent();
	if (oParent && oParent.GetPrevParagraph)
		return oParent.GetPrevParagraph();

	return null;
};
/**
 * Находимся ли мы в колонтитуле
 * @param {boolean} [bReturnHdrFtr=false]
 * @returns {?CHeaderFooter | boolean}
 */
CDocumentContentBase.prototype.IsHdrFtr = function(bReturnHdrFtr)
{
	if (true === bReturnHdrFtr)
		return null;

	return false;
};
/**
 * Находимся ли мы в сноске
 * @param {boolean} [bReturnFootnote=false]
 * @returns {?CFootEndnote | boolean}
 */
CDocumentContentBase.prototype.IsFootnote = function(bReturnFootnote)
{
	if (bReturnFootnote)
		return null;

	return false;
};
/**
 * Проверяем находимся ли мы в последней ячейке таблицы
 * @param {boolean} isSelection
 * @returns {boolean}
 */
CDocumentContentBase.prototype.IsLastTableCellInRow = function(isSelection)
{
	return false;
};
/**
 * Получаем ссылку на главный класс нумерации
 * @returns {?AscWord.CNumbering}
 */
CDocumentContentBase.prototype.GetNumbering = function()
{
	return null;
};
/**
 * Получаем верхний класс CDocumentContent
 * @returns {CDocumentContentBase}
 */
CDocumentContentBase.prototype.GetTopDocumentContent = function()
{
	return this;
};
/**
 * Получаем массив параграфов по заданному критерию
 * @param oProps
 * @param [arrParagraphs=undefined]
 * @returns {Paragraph[]}
 */
CDocumentContentBase.prototype.GetAllParagraphs = function(oProps, arrParagraphs)
{
	if (!arrParagraphs)
		arrParagraphs = [];

	return arrParagraphs;
};
/**
 * Получаем массив всех параграфов с заданной нумерацией
 * NB: массив НЕ отсортирован по позиции в документе (для сортировки, если нужно, вызывать метод AscWord.sortByDocumentPosition)
 * @param oNumPr {AscWord.NumPr | AscWord.NumPr[]}
 * @returns {Paragraph[]}
 */
CDocumentContentBase.prototype.GetAllParagraphsByNumbering = function(oNumPr)
{
	let logicDocument = this.GetLogicDocument();
	let numberingCollection = logicDocument && logicDocument.IsDocumentEditor() ? logicDocument.GetNumberingCollection() : null;
	if (numberingCollection)
		return numberingCollection.GetAllParagraphsByNumbering(oNumPr);
	
	return this.GetAllParagraphs({Numbering : true, NumPr : oNumPr});
};
/**
 * Получаем массив параграфов из списка заданных стилей
 * @param arrStylesId {string[]}
 * @returns {Paragraph[]}
 */
CDocumentContentBase.prototype.GetAllParagraphsByStyle = function(arrStylesId)
{
	return this.GetAllParagraphs({Style : true, StylesId : arrStylesId});
};
/**
 * Получаем массив таблиц по заданному критерию
 * @param oProps
 * @param [arrTables=undefined]
 * @returns {Paragraph[]}
 */
CDocumentContentBase.prototype.GetAllTables = function(oProps, arrTables)
{
	if (!arrTables)
		arrTables = [];

	return arrTables;
};
/**
 * Выделяем заданную нумерацию
 * @param oNumPr {AscWord.NumPr}
 * @param oPara {Paragraph} - текущий парограф
 */
CDocumentContentBase.prototype.SelectNumbering = function(oNumPr, oPara)
{
	var oTopDocContent = this.GetTopDocumentContent();
	if (oTopDocContent === this)
	{
		this.RemoveSelection();

		oPara.Document_SetThisElementCurrent(false);

		this.Selection.Use      = true;
		this.Selection.Flag     = selectionflag_Numbering;
		this.Selection.StartPos = this.CurPos.ContentPos;
		this.Selection.EndPos   = this.CurPos.ContentPos;

		var arrParagraphs = this.GetAllParagraphsByNumbering(oNumPr);

		this.Selection.Data = {
			Paragraphs : arrParagraphs,
			CurPara    : oPara
		};

		for (var nIndex = 0, nCount = arrParagraphs.length; nIndex < nCount; ++nIndex)
		{
			arrParagraphs[nIndex].SelectNumbering(arrParagraphs[nIndex] === oPara);
		}

		this.DrawingDocument.SelectEnabled(true);

		var oLogicDocument = this.GetLogicDocument();
		oLogicDocument.Document_UpdateSelectionState();
		oLogicDocument.Document_UpdateInterfaceState();
	}
	else
	{
		oTopDocContent.SelectNumbering(oNumPr, oPara);
	}
};
/**
 * Select numbering in a single paragraph
 * @param para {Paragraph} - current paragraph
 */
CDocumentContentBase.prototype.SelectNumberingSingleParagraph = function(para)
{
	let topDocContent = this.GetTopDocumentContent();
	if (topDocContent === this)
	{
		this.RemoveSelection();
		
		para.Document_SetThisElementCurrent(false);
		
		this.Selection.Use      = true;
		this.Selection.Flag     = selectionflag_Numbering;
		this.Selection.StartPos = this.CurPos.ContentPos;
		this.Selection.EndPos   = this.CurPos.ContentPos;
		this.Selection.Data     = {
			Paragraphs : [para],
			CurPara    : para
		};
		
		para.SelectNumbering(true, true);

		this.DrawingDocument.SelectEnabled(true);
		
		let logicDocument = this.GetLogicDocument();
		logicDocument.UpdateSelection();
		logicDocument.UpdateInterface();
	}
	else
	{
		topDocContent.SelectNumberingSingleParagraph(para);
	}
};
/**
 * Проверяем является ли текущее выделение выделением нумерации
 * @returns {boolean}
 */
CDocumentContentBase.prototype.IsNumberingSelection = function()
{
	return !!(this.IsSelectionUse() && this.Selection.Flag === selectionflag_Numbering);
};
/**
 * Убираем селект с нумерации
 */
CDocumentContentBase.prototype.RemoveNumberingSelection = function()
{
	if (this.IsNumberingSelection())
		this.RemoveSelection();
};
/**
 * Рассчитываем значение нумерованного списка для заданной нумерации
 * @param oPara {Paragraph}
 * @param oNumPr {AscWord.NumPr}
 * @param [isUseReview=false] {boolean}
 * @returns {number[]}
 */
CDocumentContentBase.prototype.CalculateNumberingValues = function(oPara, oNumPr, isUseReview)
{
	var oTopDocument = this.GetTopDocumentContent();
	if (oTopDocument instanceof CFootEndnote)
		return oTopDocument.Parent.GetNumberingInfo(oPara, oNumPr, oTopDocument, true === isUseReview);

	return oTopDocument.GetNumberingInfo(null, oPara, oNumPr, true === isUseReview);
};
/**
 * Вплоть до заданного параграфа ищем последнюю похожую нумерацию
 * @param oContinueEngine {CDocumentNumberingContinueEngine}
 */
CDocumentContentBase.prototype.GetSimilarNumbering = function(oContinueEngine)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].GetSimilarNumbering(oContinueEngine);

		if (oContinueEngine.IsFound())
			break;
	}
};
/**
 * Обновляем позиции курсора и селекта во время добавления элементов
 * @param nPosition {number}
 * @param [nCount=1] {number}
 */
CDocumentContentBase.prototype.private_UpdateSelectionPosOnAdd = function(nPosition, nCount)
{
	if (this.Content.length <= 0)
	{
		this.CurPos.ContentPos  = 0;
		this.Selection.StartPos = 0;
		this.Selection.EndPos   = 0;
		return;
	}

	if (undefined === nCount || null === nCount)
		nCount = 1;

	if (this.CurPos.ContentPos >= nPosition)
		this.CurPos.ContentPos += nCount;

	if (this.Selection.StartPos >= nPosition)
		this.Selection.StartPos += nCount;

	if (this.Selection.EndPos >= nPosition)
		this.Selection.EndPos += nCount;

	this.Selection.StartPos = Math.max(0, Math.min(this.Content.length - 1, this.Selection.StartPos));
	this.Selection.EndPos   = Math.max(0, Math.min(this.Content.length - 1, this.Selection.EndPos));
	this.CurPos.ContentPos  = Math.max(0, Math.min(this.Content.length - 1, this.CurPos.ContentPos));
};
/**
 * Обновляем позиции курсора и селекта во время удаления элементов
 * @param nPosition {number}
 * @param nCount {number}
 */
CDocumentContentBase.prototype.private_UpdateSelectionPosOnRemove = function(nPosition, nCount)
{
	if (this.CurPos.ContentPos >= nPosition + nCount)
	{
		this.CurPos.ContentPos -= nCount;
	}
	else if (this.CurPos.ContentPos >= nPosition)
	{
		if (nPosition < this.Content.length)
			this.CurPos.ContentPos = nPosition;
		else if (nPosition > 0)
			this.CurPos.ContentPos = nPosition - 1;
		else
			this.CurPos.ContentPos = 0;
	}

	if (this.Selection.StartPos <= this.Selection.EndPos)
	{
		if (this.Selection.StartPos >= nPosition + nCount)
			this.Selection.StartPos -= nCount;
		else if (this.Selection.StartPos >= nPosition)
			this.Selection.StartPos = nPosition;

		if (this.Selection.EndPos >= nPosition + nCount)
			this.Selection.EndPos -= nCount;
		else if (this.Selection.EndPos >= nPosition)
			this.Selection.StartPos = nPosition - 1;

		if (this.Selection.StartPos > this.Selection.EndPos)
		{
			this.Selection.Use = false;
			this.Selection.StartPos = 0;
			this.Selection.EndPos   = 0;
		}
	}
	else
	{
		if (this.Selection.EndPos >= nPosition + nCount)
			this.Selection.EndPos -= nCount;
		else if (this.Selection.EndPos >= nPosition)
			this.Selection.EndPos = nPosition;

		if (this.Selection.StartPos >= nPosition + nCount)
			this.Selection.StartPos -= nCount;
		else if (this.Selection.StartPos >= nPosition)
			this.Selection.StartPos = nPosition - 1;

		if (this.Selection.EndPos > this.Selection.StartPos)
		{
			this.Selection.Use = false;
			this.Selection.StartPos = 0;
			this.Selection.EndPos   = 0;
		}
	}

	this.Selection.StartPos = Math.max(0, Math.min(this.Content.length - 1, this.Selection.StartPos));
	this.Selection.EndPos   = Math.max(0, Math.min(this.Content.length - 1, this.Selection.EndPos));
	this.CurPos.ContentPos  = Math.max(0, Math.min(this.Content.length - 1, this.CurPos.ContentPos));
};
/**
 * Соединяем параграф со следующим в заданной позиции
 * @param nPosition {number}
 * @param isUseConcatedStyle {boolean} использовать ли стиль присоединяемого параграфа для итогового параграфа
 * @returns {boolean}
 */
CDocumentContentBase.prototype.ConcatParagraphs = function(nPosition, isUseConcatedStyle)
{
	if (nPosition < this.Content.length - 1 && this.Content[nPosition].IsParagraph() && this.Content[nPosition + 1].IsParagraph())
	{
		this.Content[nPosition].Concat(this.Content[nPosition + 1], isUseConcatedStyle);
		this.RemoveFromContent(nPosition + 1, 1);
		this.Content[nPosition].CorrectContent();
		return true;
	}

	return false;
};
/**
 * Специальная функция удаления параграфа во время приема/отклонения изменений
 * @param {number} nPosition
 */
CDocumentContentBase.prototype.RemoveParagraphForReview = function(nPosition)
{
	if (nPosition >= this.Content.length
		|| nPosition < 0
		|| !this.Content[nPosition].IsParagraph())
		return;

	if (nPosition >= this.Content.length - 1)
	{
		if (1 === this.Content.length)
		{
			let parent = this.GetParent();
			if (parent && parent instanceof AscWord.CBlockLevelSdt)
			{
				// Если после принятия других изменений контент не пустой, то не удаляем ничего,
				// но если он пустой и нужно было удалить последний параграф в нем, то удаляем его,
				// чтобы при проверке выше удалялись блочные контролы с пустым содержимым
				if (parent.IsEmpty())
					this.RemoveFromContent(0, 1, false);
			}
			else
				this.RemoveFromContent(0, 1, true);
		}
		else
		{
			this.RemoveFromContent(nPosition, 1);
		}
	}
	else
	{
		if (nPosition < this.Content.length - 1
			&& !this.Content[nPosition + 1].IsParagraph()
			&& this.Content[nPosition].IsEmpty())
		{
			this.RemoveFromContent(nPosition, 1);
		}
		else
		{
			this.ConcatParagraphs(nPosition, true);
		}
	}
};
/**
 * Пробегаемся по все ранам с заданной функцией
 * @param fCheck - функция проверки содержимого рана
 * @returns {boolean}
 */
CDocumentContentBase.prototype.CheckRunContent = function(fCheck)
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		if (this.Content[nIndex].CheckRunContent(fCheck))
		{
			return true;
		}
	}

	return false;
};
CDocumentContentBase.prototype.CheckSelectedRunContent = function(fCheck)
{
	let nStartPos = this.Selection.StartPos;
	let nEndPos   = this.Selection.EndPos;

	if (nStartPos > nEndPos)
	{
		nStartPos = this.Selection.EndPos;
		nEndPos   = this.Selection.StartPos;
	}

	for (var nIndex = nStartPos; nIndex <= nEndPos; ++nIndex)
	{
		if (this.Content[nIndex].CheckSelectedRunContent(fCheck))
			return true;
	}

	return false;
};
CDocumentContentBase.prototype.private_AcceptRevisionChanges = function(nType, bAll)
{
	var oTrackManager = this.GetLogicDocument() ? this.GetLogicDocument().GetTrackRevisionsManager() : null;
	if (!oTrackManager)
		return;

	var oTrackMove = oTrackManager.GetProcessTrackMove();

	if (true === this.Selection.Use || true === bAll)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (StartPos > EndPos)
		{
			StartPos = this.Selection.EndPos;
			EndPos   = this.Selection.StartPos;
		}

		var LastParaEnd;
		if (true === bAll)
		{
			StartPos    = 0;
			EndPos      = this.Content.length - 1;
			LastParaEnd = true;
		}
		else
		{
			var LastElement = this.Content[EndPos];
			LastParaEnd = (!LastElement.IsParagraph() || true === LastElement.Selection_CheckParaEnd());
		}


		if (undefined === nType || c_oAscRevisionsChangeType.ParaPr === nType)
		{
			for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
			{
				var Element = this.Content[CurPos];
				if (type_Paragraph === Element.Get_Type() && (true === bAll || true === Element.IsSelectedAll()) && true === Element.HavePrChange())
				{
					Element.AcceptPrChange();
				}
			}
		}

		for (var nCurPos = StartPos; nCurPos <= EndPos; ++nCurPos)
		{
			var oElement = this.Content[nCurPos];
			oElement.AcceptRevisionChanges(nType, bAll);
		}

		if (undefined === nType
			|| c_oAscRevisionsChangeType.ParaAdd === nType
			|| c_oAscRevisionsChangeType.ParaRem === nType
			|| c_oAscRevisionsChangeType.RowsRem === nType
			|| c_oAscRevisionsChangeType.RowsAdd === nType
			|| c_oAscRevisionsChangeType.MoveMark === nType)
		{
			EndPos = (true === LastParaEnd ? EndPos : EndPos - 1);
			for (var nCurPos = EndPos; nCurPos >= StartPos; --nCurPos)
			{
				var oElement = this.Content[nCurPos];
				if (oElement.IsParagraph())
				{
					var nReviewType = oElement.GetReviewType();
					var oReviewInfo = oElement.GetReviewInfo();

					if (reviewtype_Add === nReviewType
						&& (undefined === nType
							|| c_oAscRevisionsChangeType.ParaAdd === nType
							|| (c_oAscRevisionsChangeType.MoveMark === nType
								&& oReviewInfo.IsMovedTo()
								&& oTrackMove
								&& !oTrackMove.IsFrom()
								&& oTrackMove.GetUserId() === oReviewInfo.GetUserId())))
					{
						oElement.SetReviewType(reviewtype_Common);
					}
					else if (reviewtype_Remove === nReviewType
						&& (undefined === nType
							|| c_oAscRevisionsChangeType.ParaRem === nType
							|| (c_oAscRevisionsChangeType.MoveMark === nType
								&& oReviewInfo.IsMovedFrom()
								&& oTrackMove
								&& oTrackMove.IsFrom()
								&& oTrackMove.GetUserId() === oReviewInfo.GetUserId())))
					{
						oElement.SetReviewType(reviewtype_Common);
						this.RemoveParagraphForReview(nCurPos);
					}
				}
				else if (oElement.IsTable())
				{
					// После принятия изменений у нас могла остаться пустая таблица, такую мы должны удалить
					if (oElement.GetRowsCount() <= 0)
					{
						this.RemoveFromContent(nCurPos, 1, false);
					}
				}
				else if (oElement.IsBlockLevelSdt())
				{
					if (oElement.GetElementsCount() <= 0)
						this.RemoveFromContent(nCurPos, 1, false);
				}
			}
		}
	}
};
CDocumentContentBase.prototype.private_RejectRevisionChanges = function(nType, bAll)
{
	var oTrackManager = this.GetLogicDocument() ? this.GetLogicDocument().GetTrackRevisionsManager() : null;
	if (!oTrackManager)
		return;

	var oTrackMove = oTrackManager.GetProcessTrackMove();

	if (true === this.Selection.Use || true === bAll)
	{
		var StartPos = this.Selection.StartPos;
		var EndPos   = this.Selection.EndPos;
		if (StartPos > EndPos)
		{
			StartPos = this.Selection.EndPos;
			EndPos   = this.Selection.StartPos;
		}

		var LastParaEnd;
		if (true === bAll)
		{
			StartPos    = 0;
			EndPos      = this.Content.length - 1;
			LastParaEnd = true;
		}
		else
		{
			var LastElement = this.Content[EndPos];
			LastParaEnd = (!LastElement.IsParagraph() || true === LastElement.Selection_CheckParaEnd());
		}

		if (undefined === nType || c_oAscRevisionsChangeType.ParaPr === nType)
		{
			for (var CurPos = StartPos; CurPos <= EndPos; CurPos++)
			{
				var Element = this.Content[CurPos];
				if (type_Paragraph === Element.Get_Type() && (true === bAll || true === Element.IsSelectedAll()) && true === Element.HavePrChange())
				{
					Element.RejectPrChange();
				}
			}
		}

		for (var nCurPos = StartPos; nCurPos <= EndPos; ++nCurPos)
		{
			var oElement = this.Content[nCurPos];
			oElement.RejectRevisionChanges(nType, bAll);
		}

		if (undefined === nType
			|| c_oAscRevisionsChangeType.ParaAdd === nType
			|| c_oAscRevisionsChangeType.ParaRem === nType
			|| c_oAscRevisionsChangeType.RowsRem === nType
			|| c_oAscRevisionsChangeType.RowsAdd === nType
			|| c_oAscRevisionsChangeType.MoveMark === nType)
		{
			EndPos = (true === LastParaEnd ? EndPos : EndPos - 1);
			for (var nCurPos = EndPos; nCurPos >= StartPos; --nCurPos)
			{
				var oElement = this.Content[nCurPos];
				if (oElement.IsParagraph())
				{
					var nReviewType = oElement.GetReviewType();
					var oReviewInfo = oElement.GetReviewInfo();

					if (reviewtype_Add === nReviewType
						&& (undefined === nType
							|| c_oAscRevisionsChangeType.ParaAdd === nType
							|| (c_oAscRevisionsChangeType.MoveMark === nType
								&& oReviewInfo.IsMovedTo()
								&& oTrackMove
								&& !oTrackMove.IsFrom()
								&& oTrackMove.GetUserId() === oReviewInfo.GetUserId())))
					{
						oElement.SetReviewType(reviewtype_Common);
						this.RemoveParagraphForReview(nCurPos);
					}
					else if (reviewtype_Remove === nReviewType
						&& (undefined === nType
							|| c_oAscRevisionsChangeType.ParaRem === nType
							|| (c_oAscRevisionsChangeType.MoveMark === nType
								&& oReviewInfo.IsMovedFrom()
								&& oTrackMove
								&& oTrackMove.IsFrom()
								&& oTrackMove.GetUserId() === oReviewInfo.GetUserId())))
					{
						var oReviewInfo = oElement.GetReviewInfo();
						var oPrevInfo   = oReviewInfo.GetPrevAdded();
						if (oPrevInfo && c_oAscRevisionsChangeType.MoveMark !== nType)
						{
							oElement.SetReviewTypeWithInfo(reviewtype_Add, oPrevInfo.Copy());
						}
						else
						{
							oElement.SetReviewType(reviewtype_Common);
						}
					}
				}
				else if (oElement.IsTable())
				{
					// После принятия изменений у нас могла остаться пустая таблица, такую мы должны удалить
					if (oElement.GetRowsCount() <= 0)
					{
						this.RemoveFromContent(nCurPos, 1, false);
					}
				}
				else if (oElement.IsBlockLevelSdt())
				{
					if (oElement.GetElementsCount() <= 0)
						this.RemoveFromContent(nCurPos, 1, false);
				}
			}
		}
	}
};
/**
 * Вычисляем ли мы сейчас нижнюю границу секции для случая, когда следующая секция расположена на той же странице
 * @returns {boolean}
 */
CDocumentContentBase.prototype.IsCalculatingContinuousSectionBottomLine = function()
{
	return false;
};
/**
 * Проверяем, начинается ли заданная страница с заданного элемента
 * @param nCurPage
 * @param nElementIndex
 * @returns {boolean}
 */
CDocumentContentBase.prototype.IsFirstElementOnPage = function(nCurPage, nElementIndex)
{
	if (!this.Pages[nCurPage])
		return false;

	return (this.Pages[nCurPage].Pos === nElementIndex);
};
/**
 * Является ли данный элемент первым на странице, с которой начинается
 * @param {number} nElementPos
 * @returns {boolean}
 */
CDocumentContentBase.prototype.IsElementStartOnNewPage = function(nElementPos)
{
	for (var nCurPage = 0, nPagesCount = this.Pages.length; nCurPage < nPagesCount; ++nCurPage)
	{
		var oPage = this.Pages[nCurPage];
		if (oPage.Pos === nElementPos)
			return true;

		if (oPage.Pos < nElementPos && nElementPos <= oPage.EndPos)
			return false;
	}

	return false;
};
/**
 * Вычисляем EndInfo для всех параграфаов
 */
CDocumentContentBase.prototype.RecalculateEndInfo = function()
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].RecalculateEndInfo();
	}
};
/**
 * Является ли данный контент блочным контролом
 * @returns {boolean}
 */
CDocumentContentBase.prototype.IsBlockLevelSdtContent = function()
{
	return false;
};
/**
 * По заданному объекту получаем индекс данного объекта, либо элменента, внутри которого находится данный объект
 * @param {CFlowTable | CFlowParagraph} oFlow
 * @param {number} X
 * @param {number} Y
 * @returns {number}
 */
CDocumentContentBase.prototype.private_GetContentIndexByFlowObject = function(oFlow, X, Y)
{
	if (!oFlow)
		return -1;

	var oElement = oFlow.GetElement();
	if (oElement.GetParent() === this)
	{
		if (oFlow.IsFlowTable())
		{
			return oElement.GetIndex();
		}
		else
		{
			var nStartPos  = oFlow.StartIndex;
			var nFlowCount = oFlow.FlowCount;
			for (var nPos = nStartPos; nPos < nStartPos + nFlowCount; ++nPos)
			{
				var oBounds = this.Content[nPos].GetPageBounds(0);
				if (Y < oBounds.Bottom)
					return nPos;
			}

			return nStartPos + nFlowCount - 1;
		}
	}
	else
	{
		var arrDocPos = oElement.GetDocumentPositionFromObject();
		for (var nIndex = 0, nCount = arrDocPos.length; nIndex < nCount; ++nIndex)
		{
			if (arrDocPos[nIndex].Class === this)
				return arrDocPos[nIndex].Position;
		}
	}

	return -1;
};
/**
 * Получаем верхний док контент, который используется для пересчета плавающих объектов и различных переносов.
 * Как правило это нужно, если данный класс - это CBlockLevelSdt
 * @returns {?CDocumentContent}
 */
CDocumentContentBase.prototype.GetDocumentContentForRecalcInfo = function()
{
	return this;
};
CDocumentContentBase.prototype.GetPrevParagraphForLineNumbers = function(nIndex, isNewSection)
{
	var _nIndex = nIndex;
	if (-1 === _nIndex || undefined === _nIndex)
		_nIndex = this.Content.length;

	while (_nIndex >= 0)
	{
		if (0 === _nIndex)
		{
			var oParent = this.GetParent();
			if (oParent && oParent.GetPrevParagraphForLineNumbers)
				return oParent.GetPrevParagraphForLineNumbers(true, isNewSection);

			return null;
		}

		_nIndex--;

		if (this.Content[_nIndex].IsParagraph())
		{
			var oSectPr = this.Content[_nIndex].Get_SectionPr();
			if (oSectPr && (isNewSection || !oSectPr.HaveLineNumbers()))
				return null;

			if (!this.Content[_nIndex].IsCountLineNumbers())
				continue;

			return this.Content[_nIndex];
		}
		else if (this.Content[_nIndex].IsBlockLevelSdt())
		{
			return this.Content[_nIndex].GetPrevParagraphForLineNumbers(false, isNewSection);
		}
	}

	return null;
};
CDocumentContentBase.prototype.UpdateLineNumbersInfo = function()
{
	for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].UpdateLineNumbersInfo();
	}
};
CDocumentContentBase.prototype.GetPrevDocumentElement = function()
{
	if (this.Parent && this.Parent.GetPrevDocumentElement)
		return this.Parent.GetPrevDocumentElement();

	return null;
};
CDocumentContentBase.prototype.GetNextDocumentElement = function()
{
	if (this.Parent && this.Parent.GetNextDocumentElement)
		return this.Parent.GetNextDocumentElement();

	return null;
};
CDocumentContentBase.prototype.IsEmptyParagraphAfterTableInTableCell = function(nIndex)
{
	return false;
};
CDocumentContentBase.prototype.SetThisElementCurrent = function(isUpdateStates)
{
};

/**
 * Method for getting all
 * @param {string} sPluginId - id of plugin which can edit those ole-objects, if it's not specified return all ole-objects
 * @param {Array} arrObjects - create array if not specified
 * @returns {Array}
 */
CDocumentContentBase.prototype.GetAllOleObjects = function(sPluginId, arrObjects)
{
	if (!Array.isArray(arrObjects))
	{
		arrObjects = [];
	}
	let arrDrawings = this.GetAllDrawingObjects([]);
	for(let nDrawing = 0; nDrawing < arrDrawings.length; ++nDrawing)
	{
		arrDrawings[nDrawing].GetAllOleObjects(sPluginId, arrObjects)
	}
	return arrObjects;
};
CDocumentContentBase.prototype.ProcessComplexFields = function()
{
	for (let nIndex = 0, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		this.Content[nIndex].ProcessComplexFields();
	}
};
CDocumentContentBase.prototype.UpdateInterfaceTextPr = function()
{
	let oApi = this.GetApi();
	if (!oApi)
		return;

	let textPr = this.GetCalculatedTextPr();
	if (textPr)
		oApi.UpdateTextPr(textPr);
};
CDocumentContentBase.prototype.UpdateInterfaceParaPr = function()
{
	let api = this.GetApi();
	if (!api)
		return;
	
	let paraPr = this.GetCalculatedParaPr();
	if (!paraPr)
		return;
	
	paraPr.CanAddDropCap = this.CanAddDropCap();
	
	let logicDocument = this.GetLogicDocument();
	if (logicDocument && logicDocument.IsDocumentEditor())
	{
		let selectedInfo = logicDocument.GetSelectedElementsInfo({CheckAllSelection : true});

		let math = selectedInfo.GetMath();
		if (math)
			paraPr.CanAddImage = false;
		else if (false !== paraPr.CanAddImage)
			paraPr.CanAddImage = true;
		
		if (math && !math.IsInline())
			paraPr.Jc = math.GetAlign();
		
		paraPr.CanDeleteBlockCC  = selectedInfo.CanDeleteBlockSdts();
		paraPr.CanEditBlockCC    = selectedInfo.CanEditBlockSdts();
		paraPr.CanDeleteInlineCC = selectedInfo.CanDeleteInlineSdts();
		paraPr.CanEditInlineCC   = selectedInfo.CanEditInlineSdts();
	}
	
	if (paraPr.Tabs)
	{
		let defaultTab = paraPr.DefaultTab != null ? paraPr.DefaultTab : AscCommonWord.Default_Tab_Stop;
		api.Update_ParaTab(defaultTab, paraPr.Tabs);
	}
	
	if (paraPr.Shd && paraPr.Shd.Unifill)
		paraPr.Shd.Unifill.check(this.GetTheme(), this.GetColorMap());
	
	// Если мы находимся внутри автофигуры, тогда нам надо проверить лок параграфа, в котором находится автофигура
	if (logicDocument && logicDocument.IsDocumentEditor())
	{
		if(docpostype_DrawingObjects === logicDocument.GetDocPosType() && true !== paraPr.Locked)
		{
			let drawing = logicDocument.GetDrawingObjects().getMajorParaDrawing();
			if (drawing)
				paraPr.Locked = drawing.Lock.Is_Locked();
		}
		paraPr.StyleName = AscWord.DisplayStyleCalculator.CalculateName(this);
		api.sync_ParaStyleName(paraPr.StyleName);
	}

	api.UpdateParagraphProp(paraPr);
};
CDocumentContentBase.prototype.CanAddDropCap = function()
{
	return false;
};
/**
 * Считаем количество элементов в рамке, начиная с заданного
 * @param nStartIndex
 * @returns {number}
 */
CDocumentContentBase.prototype.CountElementsInFrame = function(nStartIndex)
{
	let oFramePr = this.Content[nStartIndex].GetFramePr();
	if (!oFramePr)
		return 0;

	let nFlowsCount = 1;
	for (let nIndex = nStartIndex + 1, nCount = this.Content.length; nIndex < nCount; ++nIndex)
	{
		let oElement = this.Content[nIndex];

		let oTempFramePr = oElement.GetFramePr();
		if (oTempFramePr && oFramePr.IsEqual(oTempFramePr) && (!oElement.IsParagraph() || !oElement.IsInline()))
			nFlowsCount++;
		else
			break;
	}

	return nFlowsCount;
};
CDocumentContentBase.prototype.OnContentChange = function()
{
	if (this.Parent && this.Parent.OnContentChange)
		this.Parent.OnContentChange();
	
	let shape = this.Is_DrawingShape(true);
	if (shape
		&& this.GetLogicDocument()
		&& this.GetLogicDocument().IsDocumentEditor())
	{
		this.GetLogicDocument().CheckShapeAutoFit(shape);
	}
	else if (shape && shape.OnContentChange) {
		shape.OnContentChange();
	}
};
CDocumentContentBase.prototype.OnTextPrChange = function()
{
	if (this.Parent && this.Parent.OnTextPrChange)
		this.Parent.OnTextPrChange();
	
	let shape = this.Is_DrawingShape(true);
	if (shape
		&& this.GetLogicDocument()
		&& this.GetLogicDocument().IsDocumentEditor())
	{
		this.GetLogicDocument().CheckShapeAutoFit(shape);
	}
	else if (shape && shape.OnTextPrChange) {
		shape.OnTextPrChange();
	}
};
CDocumentContentBase.prototype.GetCalculatedTextPr = function()
{
	var oTextPr = new CTextPr();
	oTextPr.InitDefault();
	return oTextPr;
};
CDocumentContentBase.prototype.GetCalculatedParaPr = function()
{
	var oParaPr = new CParaPr();
	oParaPr.InitDefault();
	return oParaPr;
};
CDocumentContentBase.prototype.GetDirectParaPr = function()
{
	return new CParaPr();
};
CDocumentContentBase.prototype.GetDirectTextPr = function()
{
	return new CTextPr();
};
CDocumentContentBase.prototype.GetFormattingPasteData = function(bCalcPr)
{
	if (docpostype_DrawingObjects === this.GetDocPosType())
	{
		return this.DrawingObjects.getFormatPainterData(bCalcPr);
	}
	else
	{
		let oTextPr;
		let oParaPr;
		if (bCalcPr)
		{
			oTextPr = this.GetCalculatedTextPr();
			oParaPr = this.GetCalculatedParaPr();
		}
		else
		{
			oTextPr = this.GetDirectTextPr();
			oParaPr = this.GetDirectParaPr();
		}
		return new AscCommon.CTextFormattingPasteData(oTextPr, oParaPr);
	}
};
CDocumentContentBase.prototype.UpdateNumberingCollection = function(elements)
{
	let logicDocument = this.GetLogicDocument();
	if (!logicDocument || !logicDocument.IsDocumentEditor())
		return;
	
	let numberingCollection = logicDocument.GetNumberingCollection();
	for (let iElement = 0, nElements = elements.length; iElement < nElements; ++iElement)
	{
		if (elements[iElement].IsParagraph())
		{
			numberingCollection.CheckParagraph(elements[iElement]);
		}
		else
		{
			let paragraphs = elements[iElement].GetAllParagraphs();
			for (let iPara = 0, nParas = paragraphs.length; iPara < nParas; ++iPara)
			{
				numberingCollection.CheckParagraph(paragraphs[iPara]);
			}
		}
	}
};
CDocumentContentBase.prototype.private_RecalculateNumbering = function(elements)
{
	if(this.bPresentation)
		return;
	this.UpdateNumberingCollection(elements);
	
	let logicDocument = this.GetLogicDocument();
	if (!logicDocument || !logicDocument.IsDocumentEditor())
		return;
	
	if (true === AscCommon.g_oIdCounter.m_bLoad || true === AscCommon.g_oIdCounter.m_bRead)
		return;
	
	let history = logicDocument.GetHistory();
	for (let iElement = 0, nElements = elements.length; iElement < nElements; ++iElement)
	{
		if (elements[iElement].IsParagraph())
		{
			history.Add_RecalcNumPr(elements[iElement].GetNumPr());
		}
		else if (elements[iElement].GetAllParagraphs)
		{
			let paragraphs = elements[iElement].GetAllParagraphs({All : true});
			for (let iPara = 0, nParas = paragraphs.length; iPara < nParas; ++iPara)
			{
				history.Add_RecalcNumPr(paragraphs[iPara].GetNumPr());
			}
		}
	}
};
CDocumentContentBase.prototype.Get_Theme = function()
{
	if (this.Parent)
		return this.Parent.Get_Theme();
	
	if (this.LogicDocument)
		return this.LogicDocument.GetTheme();
	
	return AscFormat.GetDefaultTheme();
};
CDocumentContentBase.prototype.Get_ColorMap = function()
{
	if (this.Parent)
		return this.Parent.Get_ColorMap();
	
	if (this.LogicDocument)
		return this.LogicDocument.GetColorMap();
	
	return AscFormat.GetDefaultColorMap();
};
CDocumentContentBase.prototype.GetTheme = function()
{
	return this.Get_Theme();
};
CDocumentContentBase.prototype.GetColorMap = function()
{
	return this.Get_ColorMap();
};
CDocumentContentBase.prototype.GetSelectedParagraphs = function()
{
	let logicDocument = this.GetLogicDocument();
	if (!logicDocument)
		return [];
	
	return logicDocument.GetSelectedParagraphs();
};
CDocumentContentBase.prototype.setSelectionStateSilent = function(state)
{
	let logicDocument = this.GetLogicDocument();
	if (logicDocument && !logicDocument.IsDocumentEditor())
		logicDocument = null;

	if (logicDocument)
		logicDocument.Start_SilentMode();
	
	this.SetSelectionState(state);
	
	if (logicDocument)
		logicDocument.End_SilentMode(false);
};
CDocumentContentBase.prototype.getSpeechDescription = function(prevState, action)
{
	if (!prevState)
		return null;
	
	if (action && (action.type !== AscCommon.SpeakerActionType.keyDown || action.event.KeyCode < 33 || action.event.KeyCode > 40))
		return null;
	
	// В данном метод предполагается, что curState равен this.GetSelectionState()
	let curState = this.GetSelectionState();
	if (!prevState.length || !curState.length)
		return null;
	
	let obj  = {};
	let type = AscCommon.SpeechWorkerCommands.Text;
	
	this.setSelectionStateSilent(prevState);
	let prevInfo = this.getSelectionInfo();
	
	this.setSelectionStateSilent(curState);
	let curInfo = this.getSelectionInfo();
	
	let isActionSelectionChange = action && action.type === AscCommon.SpeakerActionType.keyDown && action.event.ShiftKey;
	
	if (curInfo.docPosType === docpostype_DrawingObjects)
	{
		return AscCommon.getSpeechDescription(prevState, curState, action);
	}
	
	if (prevInfo.docPosType !== curInfo.docPosType
		|| !prevInfo.curPos.length
		|| !curInfo.curPos.length
		|| prevInfo.curPos[0].Class !== curInfo.curPos[0].Class
		|| (!(curInfo.curPos[0].Class instanceof CDocument) && !(curInfo.curPos[0].Class instanceof CDocumentContent)))
	{
		switch (curInfo.docPosType)
		{
			case docpostype_Content:  obj.moveToMainPart = true; break;
			case docpostype_Endnotes:
			case docpostype_Footnotes: obj.moveToFootnote = true; break;
			case docpostype_DrawingObjects: obj.moveToDrawing = true; break;
			case docpostype_HdrFtr: obj.moveToHdrFtr = true; break;
		}
		
		let paragraph = this.GetCurrentParagraph();
		obj.text = paragraph.getNextCharacter();
	}
	else
	{
		let mainDC = curInfo.curPos[0].Class;
		if (prevInfo.isSelection && !curInfo.isSelection && isActionSelectionChange)
		{
			obj.cancelSelection = true;
			this.setSelectionStateSilent(prevState);
			type     = AscCommon.SpeechWorkerCommands.TextUnselected;
			obj.text = this.GetSelectedText(false);
			this.setSelectionStateSilent(curState);
		}
		else if (!curInfo.isSelection || 0 === AscWord.CompareDocumentPositions(curInfo.selectionStart, curInfo.selectionEnd))
		{
			if (prevInfo.isSelection && 0 !== AscWord.CompareDocumentPositions(prevInfo.selectionStart, prevInfo.selectionEnd))
				obj.cancelSelection = true;
			
			if (curInfo.isSelection || !action)
			{
				if (prevInfo.isSelection && 0 !== AscWord.CompareDocumentPositions(prevInfo.selectionStart, prevInfo.selectionEnd))
				{
					this.setSelectionStateSilent(prevState);
					type     = AscCommon.SpeechWorkerCommands.TextUnselected;
					obj.text = this.GetSelectedText(false);
					this.setSelectionStateSilent(curState);
				}
				else
				{
					let paragraph = this.GetCurrentParagraph();
					type     = AscCommon.SpeechWorkerCommands.Text;
					obj.text = paragraph.getNextCharacter();
				}
			}
			else
			{
				let paragraph = this.GetCurrentParagraph();
				obj.text = paragraph.getNextCharacter();
				if (action)
				{
					if (AscCommon.SpeakerActionType.keyDown !== action.type)
						return null;
					
					let keyCode = action.event.KeyCode;
					if (36 === keyCode)
					{
						if (action.event.CtrlKey)
							obj.moveToStartOfDocument = true;
						else
							obj.moveToStartOfLine = true;
					}
					else if (35 === keyCode)
					{
						if (action.event.CtrlKey)
							obj.moveToEndOfDocument = true;
						else
							obj.moveToEndOfLine = true;
					}
					else if (33 === keyCode)
					{
						obj.movePageUp = true;
					}
					else if (34 === keyCode)
					{
						obj.movePageDown = true;
					}
					
					if (36 === keyCode || 38 === keyCode || 40 === keyCode)
						obj.text = paragraph.getTextOnLine();
					else if ((37 === keyCode || 39 === keyCode) && action.event.CtrlKey)
						obj.text = paragraph.getCurrentWord(1);
					else if (35 === keyCode)
						obj.text = "";
				}
			}
		}
		else
		{
			if (!prevInfo.isSelection || 0 === AscWord.CompareDocumentPositions(prevInfo.selectionStart, prevInfo.selectionEnd))
			{
				type = AscCommon.SpeechWorkerCommands.TextSelected;
				obj.text = this.GetSelectedText(false);
			}
			else
			{
				if (0 !== AscWord.CompareDocumentPositions(prevInfo.selectionStart, curInfo.selectionStart)
					|| ((AscWord.CompareDocumentPositions(prevInfo.selectionStart, prevInfo.selectionEnd) <= 0
							&& AscWord.CompareDocumentPositions(prevInfo.selectionStart, curInfo.selectionEnd) > 0)
						|| (AscWord.CompareDocumentPositions(prevInfo.selectionStart, prevInfo.selectionEnd) >= 0
							&& AscWord.CompareDocumentPositions(prevInfo.selectionStart, curInfo.selectionEnd) < 0)))
				{
					// TODO: Нужно ли посылать два ивента?
					// this.setSelectionStateSilent(prevState);
					// type     = AscCommon.SpeechWorkerCommands.TextUnselected;
					// obj.text = this.GetSelectedText(false);
					// this.setSelectionStateSilent(curState);
					
					type     = AscCommon.SpeechWorkerCommands.TextSelected;
					obj.text = this.GetSelectedText(false);
				}
				else
				{
					let isAdd = ((AscWord.CompareDocumentPositions(prevInfo.selectionStart, prevInfo.selectionEnd) <= 0
						&& AscWord.CompareDocumentPositions(curInfo.selectionEnd, prevInfo.selectionEnd) >= 0)
					|| (AscWord.CompareDocumentPositions(prevInfo.selectionStart, prevInfo.selectionEnd) >= 0
							&& AscWord.CompareDocumentPositions(curInfo.selectionEnd, prevInfo.selectionEnd) <= 0));
					
					mainDC.RemoveSelection();
					mainDC.SetContentSelection(curInfo.selectionEnd, prevInfo.selectionEnd, 0, 0, 0);
					
					type     = isAdd ? AscCommon.SpeechWorkerCommands.TextSelected : AscCommon.SpeechWorkerCommands.TextUnselected;
					obj.text = this.GetSelectedText(false);
					
					this.setSelectionStateSilent(curState);
				}
			}
		}
	}
	
	return {obj : obj, type : type};
};
CDocumentContentBase.prototype.getSelectionInfo = function()
{
	return {
		docPosType     : this.GetDocPosType(),
		curPos         : this.getDocumentContentPosition(false),
		isSelection    : this.IsSelectionUse(),
		selectionStart : this.getDocumentContentPosition(true, true),
		selectionEnd   : this.getDocumentContentPosition(true, false)
	};
};
/**
 * Данный метод отличается от обычного GetContentPosition, тем что для класса CDocument он возвращает полную позицию
 * с учетом того, где находится содержимое (основная часть, колонтитул, сноска)
 * @param {boolean} isSelection
 * @param {boolean} isStart
 * @returns {Array}
 */
CDocumentContentBase.prototype.getDocumentContentPosition = function(isSelection, isStart)
{
	return this.GetContentPosition(isSelection, isStart);
};
CDocumentContentBase.prototype.GetCurrentRun = function()
{
	let paragraph = this.GetCurrentParagraph(false, false);
	if (!paragraph || !paragraph.IsParagraph())
		return null;
	
	let paraPos = paragraph.Get_ParaContentPos(false);
	let run = paragraph.GetElementByPos(paraPos);
	if (!run || !(run instanceof AscWord.CRun))
		return null;
	
	return run;
};
CDocumentContentBase.prototype.GetText = function(pr)
{
	let text = "";
	for (let i = 0; i < this.Content.length; ++i)
	{
		if (this.Content[i].GetText)
			text += this.Content[i].GetText(pr);
	}
	
	return text;
};
