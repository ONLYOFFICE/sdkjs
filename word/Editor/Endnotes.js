/*
 * (c) Copyright Ascensio System SIA 2010-2020
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
 * Date: 09.04.2020
 * Time: 13:31
 */

/**
 * Класс, работающий с концевыми сносками документа
 * @param {CDocument} oLogicDocument - Ссылка на главный документ.
 * @constructor
 * @extends {CDocumentControllerBase}
 */
function CEndnotesController(oLogicDocument)
{
	CDocumentControllerBase.call(this, oLogicDocument);

	this.Id = oLogicDocument.GetIdCounter().Get_NewId();

	this.EndnotePr = new CFootnotePr(); // Глобальные настройки для сносок
	this.EndnotePr.InitDefaultEndnotePr();

	this.Endnote = {};
	this.Pages   = [];
	this.Sections = {};

	// Специальные сноски
	this.ContinuationNotice    = null;
	this.ContinuationSeparator = null;
	this.Separator             = null;

	this.Selection = {
		Use : false,

		Start : {
			Endnote          : null,
			Index            : 0,
			Section          : 0,
			Page             : 0,
			Column           : 0,
			EndnotePageIndex : 0
		},

		End : {
			Endnote          : null,
			Index            : 0,
			Section          : 0,
			Page             : 0,
			Column           : 0,
			EndnotePageIndex : 0
		},

		Endnotes  : {},
		Direction : 0
	};

	this.CurEndnote = null;

	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
	oLogicDocument.GetTableId().Add(this, this.Id);
}

CEndnotesController.prototype = Object.create(CDocumentControllerBase.prototype);
CEndnotesController.prototype.constructor = CEndnotesController;

/**
 * Получаем Id данного класса
 */
CEndnotesController.prototype.Get_Id = function()
{
	return this.Id;
};
/**
 * Получаем Id данного класса
 */
CEndnotesController.prototype.GetId = function()
{
	return this.Id;
};
/**
 * Начальная инициализация после загрузки
 */
CEndnotesController.prototype.ResetSpecialEndnotes = function()
{
	var oSeparator = new CFootEndnote(this);
	oSeparator.AddToParagraph(new ParaSeparator(), false);
	var oParagraph = oSeparator.GetElement(0);
	oParagraph.Set_Spacing({After : 0, Line : 1, LineRule : Asc.linerule_Auto}, false);
	this.SetSeparator(oSeparator);

	var oContinuationSeparator = new CFootEndnote(this);
	oContinuationSeparator.AddToParagraph(new ParaContinuationSeparator(), false);
	oParagraph = oContinuationSeparator.GetElement(0);
	oParagraph.Set_Spacing({After : 0, Line : 1, LineRule : Asc.linerule_Auto}, false);
	this.SetContinuationSeparator(oContinuationSeparator);

	this.SetContinuationNotice(null);
};
/**
 * Создаем новую сноску
 * @returns {CFootEndnote}
 */
CEndnotesController.prototype.CreateEndnote = function()
{
	var oEndnote = new CFootEndnote(this);

	this.Endnote[oEndnote.GetId()] = oEndnote;

	this.LogicDocument.GetHistory().Add(new CChangesEndnotesAddEndnote(this, oEndnote.GetId()));
	return oEndnote;
};
/**
 * Добавляем сноску (функция для открытия файла)
 * @param {CFootEndnote} oEndnote
 */
CEndnotesController.prototype.AddEndnote = function(oEndnote)
{
	this.Endnote[oEndnote.GetId()] = oEndnote;
	oEndnote.SetParent(this);
	this.LogicDocument.GetHistory().Add(new CChangesEndnotesAddEndnote(this, oEndnote.GetId()));
};
CEndnotesController.prototype.RemoveEndnote = function(oEndnote)
{
	delete this.Endnote[oEndnote.GetId()];
	this.LogicDocument.GetHistory().Add(new CChangesEndnotesRemoveEndnote(this, oEndnote.GetId()));
};
CEndnotesController.prototype.SetSeparator = CFootnotesController.prototype.SetSeparator;
CEndnotesController.prototype.SetContinuationSeparator = CFootnotesController.prototype.SetContinuationSeparator;
CEndnotesController.prototype.SetContinuationNotice = CFootnotesController.prototype.SetContinuationNotice;
CEndnotesController.prototype.IsSpecialEndnote = CFootnotesController.prototype.IsSpecialFootnote;
CEndnotesController.prototype.SetEndnotePrNumFormat = function(nFormatType)
{
	if (undefined !== nFormatType && this.EndnotePr.NumFormat !== nFormatType)
	{
		this.LogicDocument.GetHistory().Add(new CChangesSectionEndnoteNumFormat(this, this.EndnotePr.NumFormat, nFormatType));
		this.EndnotePr.NumFormat = nFormatType;
	}
};
CEndnotesController.prototype.SetEndnotePrPos = function(nPos)
{
	if (undefined !== nPos && this.EndnotePr.Pos !== nPos)
	{
		this.LogicDocument.GetHistory().Add(new CChangesSectionEndnotePos(this, this.EndnotePr.Pos, nPos));
		this.EndnotePr.Pos = nPos;
	}
};
CEndnotesController.prototype.SetEndnotePrNumStart = function(nStart)
{
	if (undefined !== nStart && this.EndnotePr.NumStart !== nStart)
	{
		this.LogicDocument.GetHistory().Add(new CChangesSectionEndnoteNumStart(this, this.EndnotePr.NumStart, nStart));
		this.EndnotePr.NumStart = nStart;
	}
};
CEndnotesController.prototype.SetEndnotePrNumRestart = function(nRestartType)
{
	if (undefined !== nRestartType && this.EndnotePr.NumRestart !== nRestartType)
	{
		this.LogicDocument.GetHistory().Add(new CChangesSectionEndnoteNumRestart(this, this.EndnotePr.NumRestart, nRestartType));
		this.EndnotePr.NumRestart = nRestartType;
	}
};
CEndnotesController.prototype.GetEndnotePrPos = function()
{
	return this.EndnotePr.Pos;
};
/**
 * Проверяем, используется заданная сноска в документе.
 * @param {string} sEndnoteId
 * @param {CFootEndnote.array} arrEndnotesList
 * @returns {boolean}
 */
CEndnotesController.prototype.Is_UseInDocument = function(sEndnoteId, arrEndnotesList)
{
	if (!arrEndnotesList)
		arrEndnotesList = this.private_GetEndnotesLogicRange(null, null);

	var oEndnote = null;
	for (var nIndex = 0, nCount = arrEndnotesList.length; nIndex < nCount; ++nIndex)
	{
		var oTempEndnote = arrEndnotesList[nIndex];
		if (oTempEndnote.GetId() === sEndnoteId)
		{
			oEndnote = oTempEndnote;
			break;
		}
	}

	if (this.Endnote[sEndnoteId] === oEndnote)
		return true;

	return false;
};
/**
 * Проверяем является ли данная сноска текущей.
 * @param oEndnote
 * return {boolean}
 */
CEndnotesController.prototype.Is_ThisElementCurrent = function(oEndnote)
{
	if (oEndnote === this.CurEndnote && docpostype_Endnotes === this.LogicDocument.GetDocPosType())
		return true;

	return false;
};
/**
 * Есть ли сноски на заданной странице
 * @param {number} nPageAbs
 * @returns {boolean}
 */
CEndnotesController.prototype.IsEmptyPage = function(nPageAbs)
{
	var oPage = this.Pages[nPageAbs];
	if (!oPage)
		return true;

	for (var nIndex = 0, nCount = oPage.Sections.length; nIndex < nCount; ++nIndex)
	{
		var oSection = this.Sections[oPage.Sections[nIndex]];
		if (!oSection)
			continue;

		var oSectionPage = oSection.Pages[nPageAbs];
		if (!oSectionPage)
			continue;

		for (var nColumnIndex = 0, nColumnsCount = oSectionPage.Columns.length; nColumnIndex < nColumnsCount; ++nColumnIndex)
		{
			if (!this.IsEmptyPageColumn(nPageAbs, nColumnIndex, oPage.Sections[nIndex]))
				return false;
		}
	}

	return true;
};
CEndnotesController.prototype.IsEmptyPageColumn = function(nPageIndex, nColumnIndex, nSectionIndex)
{
	var oColumn = this.private_GetPageColumn(nPageIndex, nColumnIndex, nSectionIndex);
	if (!oColumn || oColumn.Elements.length <= 0)
		return true;

	return false;
};
CEndnotesController.prototype.GetPageBounds = function(nPageAbs, nColumnAbs, nSectionAbs)
{
	var oColumn = this.private_GetPageColumn(nPageAbs, nColumnAbs, nSectionAbs);
	if (!oColumn)
		return new CDocumentBounds(0, 0, 0, 0);

	return new CDocumentBounds(oColumn.X, oColumn.Y, oColumn.XLimit, oColumn.Y + oColumn.Height);
};
CEndnotesController.prototype.Refresh_RecalcData = function(Data)
{
};
CEndnotesController.prototype.Refresh_RecalcData2 = function(nRelPageIndex)
{
	var nAbsPageIndex = nRelPageIndex;
	if (this.LogicDocument.Pages[nAbsPageIndex])
	{
		var nIndex = this.LogicDocument.Pages[nAbsPageIndex].Pos;

		if (nIndex >= this.LogicDocument.Content.length)
		{
			History.RecalcData_Add({
				Type    : AscDFH.historyitem_recalctype_NotesEnd,
				PageNum : nAbsPageIndex
			});
		}
		else
		{
			this.LogicDocument.Refresh_RecalcData2(nIndex, nAbsPageIndex);
		}
	}
};
CEndnotesController.prototype.Get_PageContentStartPos = function(nPageAbs, nColumnAbs, nSectionAbs)
{
	var oColumn = this.private_GetPageColumn(nPageAbs, nColumnAbs, nSectionAbs);
	if (!oColumn)
		return {X : 0, Y : 0, XLimit : 0, YLimit : 0};

	return {X : oColumn.X, Y : oColumn.Y + oColumn.Height, XLimit : oColumn.XLimit, YLimit : oColumn.YLimit};
};
CEndnotesController.prototype.OnContentReDraw = function(StartPageAbs, EndPageAbs)
{
	this.LogicDocument.OnContentReDraw(StartPageAbs, EndPageAbs);
};
CEndnotesController.prototype.GetEndnoteNumberOnPage = function(nPageAbs, nColumnAbs, oSectPr, oCurEndnote)
{
	var nNumRestart = section_footnote_RestartContinuous;
	var nNumStart   = 1;
	if (oSectPr)
	{
		nNumRestart = oSectPr.GetEndnoteNumRestart();
		nNumStart   = oSectPr.GetEndnoteNumStart();
	}

	// NumStart никак не влияет в случае RestartEachSect. Влияет только на случай RestartContinuous:
	// к общему количеству сносок добавляется данное значение, взятое для текущей секции, этоже значение из предыдущих
	// секций не учитывается.

	if (section_footnote_RestartEachSect === nNumRestart)
	{
		for (var nPageIndex = nPageAbs; nPageIndex >= 0; --nPageIndex)
		{
			var oPage = this.Pages[nPageIndex];
			if (oPage && oPage.Endnotes.length > 0)
			{
				if (oEndnote === oCurEndnote)
					return oEndnote.GetNumber();

				var oEndnote = oPage.Endnotes[oPage.Endnotes.length - 1];
				if (oEndnote.GetReferenceSectPr() !== oSectPr)
					return 1;

				return oPage.Endnotes[oPage.Endnotes.length - 1].GetNumber() + 1;
			}
		}
	}
	else// if (section_footnote_RestartContinuous === nNumRestart)
	{
		// Здесь нам надо считать, сколько сносок всего в документе до данного момента, отталкиваться от предыдущей мы
		// не можем, потому что Word считает общее количество сносок, а не продолжает нумерацию с предыдущей секции,
		// т.е. после последнего номера 4 в старой секции, в новой секции может идти уже, например, 9.
		var nEndnotesCount = 0;
		for (var nPageIndex = nPageAbs; nPageIndex >= 0; --nPageIndex)
		{
			var oPage = this.Pages[nPageIndex];
			if (oPage && oPage.Endnotes.length > 0)
			{
				for (var nEndnoteIndex = 0, nTempCount = oPage.Endnotes.length; nEndnoteIndex < nTempCount; ++nEndnoteIndex)
				{
					var oEndnote = oPage.Endnotes[nEndnoteIndex];

					if (oEndnote === oCurEndnote)
						return oEndnote.GetNumber();

					if (oEndnote && true !== oEndnote.IsCustomMarkFollows())
						nEndnotesCount++;
				}
			}
		}

		return nEndnotesCount + nNumStart;
	}

	return 1;
};
/**
 * Сбрасываем расчетные данные с заданного места
 * @param nPageIndex
 * @param nSectionIndex
 * @param nColumnIndex
 */
CEndnotesController.prototype.Reset = function(nPageIndex, nSectionIndex, nColumnIndex)
{
	if (0 === nSectionIndex && 0 === nColumnIndex)
	{
		this.Pages.length = nPageIndex;
		if (!this.Pages[nPageIndex])
			this.Pages[nPageIndex] = new CEndnotePage();
	}
	else
	{
		this.Pages[nPageIndex].ResetColumn(nSectionIndex, nColumnIndex);
	}
};
/**
 * Регистрируем сноски на заданной странице
 * @param nPageAbs
 * @param arrEndnotes
 */
CEndnotesController.prototype.RegisterEndnotes = function(nPageAbs, arrEndnotes)
{
	if (!this.Pages[nPageAbs])
		return;

	this.Pages[nPageAbs].AddEndnotes(arrEndnotes);
};
/**
 * Проверяем, есть ли сноски, которые нужно пересчитать в конце заданной секции
 * @param oSectPr {CSectionPr} секция, в конце которой мы расчитываем сноски
 * @param isFinal {boolean} последняя ли это секция документа
 * @returns {boolean}
 */
CEndnotesController.prototype.HaveEndnotes = function(oSectPr, isFinal)
{
	var nEndnotesPos = this.GetEndnotePrPos();

	if (isFinal && Asc.c_oAscEndnotePos.DocEnd === nEndnotesPos)
	{
		for (var nCurPage = 0, nPagesCount = this.Pages.length; nCurPage < nPagesCount; ++nCurPage)
		{
			if (this.Pages[nCurPage].Endnotes.length > 0)
				return true;
		}
	}
	else if (Asc.c_oAscEndnotePos.SectEnd === nEndnotesPos)
	{
		// Мы должны найти просто ссылку на самую последнюю сноску, и если она привязана не данной секции, значит
		// в данной секции и не было никаких сносок
		for (var nCurPage = this.Pages.length - 1; nCurPage >= 0; --nCurPage)
		{
			var oPage = this.Pages[nCurPage];
			if (oPage.Endnotes.length > 0)
			{
				return (oSectPr === oPage.Endnotes[oPage.Endnotes.length - 1].GetReferenceSectPr());
			}
		}
	}

	return false;
};
CEndnotesController.prototype.ClearSection = function(nSectionIndex)
{
	this.Sections.length = nSectionIndex;
	this.Sections[nSectionIndex] = new CEndnoteSection();
};
CEndnotesController.prototype.FillSection = function(nPageAbs, nColumnAbs, oSectPr, nSectionIndex, isFinal)
{
	var oSection = this.private_UpdateSection(oSectPr, nSectionIndex, isFinal, nPageAbs);
	if (oSection.Endnotes.length <= 0)
		return recalcresult2_End;

	oSection.StartPage   = nPageAbs;
	oSection.StartColumn = nColumnAbs;
};
CEndnotesController.prototype.Recalculate = function(X, Y, XLimit, YLimit, nPageAbs, nColumnAbs, nColumnsCount, oSectPr, nSectionIndex, isFinal)
{
	var oSection = this.Sections[nSectionIndex];
	if (!oSection)
		return recalcresult2_End;

	if (this.Pages[nPageAbs])
		this.Pages[nPageAbs].AddSection(nSectionIndex);

	var nStartPos = 0;
	var isStart   = true;

	if (nPageAbs < oSection.StartPage || (nPageAbs === oSection.StartPage && nColumnAbs < oSection.StartColumn))
	{
		// Такого не должно быть
		return recalcresult2_End;
	}
	else if (nPageAbs === oSection.StartPage && nColumnAbs === oSection.StartColumn)
	{
		nStartPos = 0;
		isStart   = true;
	}
	else if (0 === nColumnAbs)
	{
		if (!oSection.Pages[nPageAbs - 1] || oSection.Pages[nPageAbs - 1].Columns.length <= 0)
			return recalcresult2_End;

		nStartPos = oSection.Pages[nPageAbs - 1].Columns[oSection.Pages[nPageAbs - 1].Columns.length - 1].EndPos;
		isStart   = false;
	}
	else
	{
		nStartPos = oSection.Pages[nPageAbs].Columns[nColumnAbs - 1].EndPos;
		isStart   = false;
	}

	// Случай, когда на предыдущей странице не убралось ни одной сноски и мы перенеслись сразу на следующую
	if (-1 === nStartPos)
	{
		nStartPos = 0;
		isStart   = true;
	}

	if (!oSection.Pages[nPageAbs])
	{
		oSection.Pages[nPageAbs] = new CEndnoteSectionPage();

		var nColumnSpace = nColumnAbs > 0 ? oSectPr.GetColumnSpace(nColumnAbs - 1) : 0;

		for (var nColumnIndex = 0; nColumnIndex < nColumnAbs; ++nColumnIndex)
		{
			var oTempColumn = new CEndnoteSectionPageColumn();
			oSection.Pages[nPageAbs].Columns[nColumnIndex] = oTempColumn;

			oTempColumn.X      = X - nColumnSpace;
			oTempColumn.Y      = Y;
			oTempColumn.XLimit = X - nColumnSpace;
			oTempColumn.YLimit = YLimit;
		}
	}

	var oColumn = new CEndnoteSectionPageColumn();
	oSection.Pages[nPageAbs].Columns[nColumnAbs] = oColumn;

	oColumn.X      = X;
	oColumn.Y      = Y;
	oColumn.XLimit = XLimit;
	oColumn.YLimit = YLimit;

	oColumn.StartPos = nStartPos;

	var _Y = Y;
	if (isStart && this.Separator)
	{
		this.Separator.PrepareRecalculateObject();
		this.Separator.SetSectionIndex(nSectionIndex);
		this.Separator.Reset(X, _Y, XLimit, YLimit);
		this.Separator.Set_StartPage(nPageAbs, nColumnAbs, nColumnsCount);
		this.Separator.Recalculate_Page(0, true);
		oColumn.SeparatorRecalculateObject = this.Separator.SaveRecalculateObject();
		oColumn.Separator = true;

		var oBounds = this.Separator.GetPageBounds(0);
		_Y += oBounds.Bottom - oBounds.Top;
		oColumn.Height = _Y - Y;
	}
	else if (!isStart && this.ContinuationSeparator)
	{
		this.ContinuationSeparator.PrepareRecalculateObject();
		this.ContinuationSeparator.SetSectionIndex(nSectionIndex);
		this.ContinuationSeparator.Reset(X, _Y, XLimit, YLimit);
		this.ContinuationSeparator.Set_StartPage(nPageAbs, nColumnAbs, nColumnsCount);
		this.ContinuationSeparator.Recalculate_Page(0, true);
		oColumn.SeparatorRecalculateObject = this.ContinuationSeparator.SaveRecalculateObject();
		oColumn.Separator = false;

		var oBounds = this.Separator.GetPageBounds(0);
		_Y += oBounds.Bottom - oBounds.Top;
		oColumn.Height = _Y - Y;
	}

	for (var nPos = nStartPos, nCount = oSection.Endnotes.length; nPos < nCount; ++nPos)
	{
		var oEndnote = oSection.Endnotes[nPos];

		oEndnote.SetSectionIndex(nSectionIndex);
		if (isStart || nPos !== nStartPos)
		{
			oEndnote.Reset(X, _Y, XLimit, YLimit);
			oEndnote.Set_StartPage(nPageAbs, nColumnAbs, nColumnsCount);
		}

		var nRelativePage = oEndnote.GetElementPageIndex(nPageAbs, nColumnAbs);
		var nRecalcResult = oEndnote.Recalculate_Page(nRelativePage, true);

		if (recalcresult2_NextPage === nRecalcResult)
		{
			if (nColumnAbs >= nColumnsCount - 1)
				this.Pages[nPageAbs].SetContinue(true);

			if (0 === nPos && !oEndnote.IsContentOnFirstPage())
			{
				oColumn.EndPos = -1;
				return recalcresult2_NextPage;
			}
			else
			{
				oColumn.EndPos = nPos;
				oColumn.Elements.push(oEndnote);

				var oBounds = oEndnote.GetPageBounds(nRelativePage);
				_Y += oBounds.Bottom - oBounds.Top;
				oColumn.Height = _Y - Y;

				return recalcresult2_NextPage;
			}
		}
		else if (recalcresult2_CurPage === nRecalcResult)
		{
			// Такого не должно быть при расчете сносок
		}

		oColumn.EndPos = nPos;
		oColumn.Elements.push(oEndnote);

		var oBounds = oEndnote.GetPageBounds(nRelativePage);
		_Y += oBounds.Bottom - oBounds.Top;
		oColumn.Height = _Y - Y;

		if (recalcresult2_NextPage === nRecalcResult)
		{
			return recalcresult2_NextPage;
		}
	}

	for (var nColumnIndex = nColumnAbs + 1; nColumnIndex < nColumnsCount; ++nColumnIndex)
	{
		var oTempColumn = new CEndnoteSectionPageColumn();
		oSection.Pages[nPageAbs].Columns[nColumnIndex] = oTempColumn;

		oTempColumn.X      = XLimit + 10;
		oTempColumn.Y      = Y;
		oTempColumn.XLimit = XLimit + 5;
		oTempColumn.YLimit = YLimit;
	}

	return recalcresult2_End;
};
CEndnotesController.prototype.private_UpdateSection = function(oSectPr, nSectionIndex, isFinal, nPageAbs)
{
	var oPos = this.GetEndnotePrPos();

	this.Sections.length = nSectionIndex;
	this.Sections[nSectionIndex] = new CEndnoteSection();

	for (var nCurPage = 0; nCurPage <= nPageAbs; ++nCurPage)
	{
		var oPage = this.Pages[nCurPage];
		if (oPage)
		{
			for (var nEndnoteIndex = 0, nEndnotesCount = oPage.Endnotes.length; nEndnoteIndex < nEndnotesCount; ++nEndnoteIndex)
			{
				if ((oPos === Asc.c_oAscEndnotePos.DocEnd && isFinal) || (oPos === Asc.c_oAscEndnotePos.SectEnd && oPage.Endnotes[nEndnoteIndex].GetReferenceSectPr() === oSectPr))
					this.Sections[nSectionIndex].Endnotes.push(oPage.Endnotes[nEndnoteIndex]);
			}
		}
	}

	return this.Sections[nSectionIndex];
};
CEndnotesController.prototype.GetLastSectionIndexOnPage = function(nPageAbs)
{
	var oPage = this.Pages[nPageAbs];
	if (oPage && oPage.Sections.length)
		return oPage.Sections[oPage.Sections.length - 1];

	return -1;
};
/**
 * Отрисовываем сноски на заданной странице.
 * @param {number} nPageAbs
 * @param {number} nSectionIndex
 * @param {CGraphics} oGraphics
 */
CEndnotesController.prototype.Draw = function(nPageAbs, nSectionIndex, oGraphics)
{
	var oSection = this.Sections[nSectionIndex];
	if (!oSection)
		return;

	var oPage = oSection.Pages[nPageAbs];
	if (!oPage)
		return;

	for (var nColumnIndex = 0, nColumnsCount = oPage.Columns.length; nColumnIndex < nColumnsCount; ++nColumnIndex)
	{
		var oColumn = oPage.Columns[nColumnIndex];
		if (!oColumn || oColumn.Elements.length <= 0)
			continue;

		if (oColumn.Separator && this.Separator && oColumn.SeparatorRecalculateObject)
		{
			this.Separator.LoadRecalculateObject(oColumn.SeparatorRecalculateObject);
			this.Separator.Draw(nPageAbs, oGraphics);
		} else if (!oColumn.Separator && this.ContinuationSeparator && oColumn.SeparatorRecalculateObject)
		{
			this.ContinuationSeparator.LoadRecalculateObject(oColumn.SeparatorRecalculateObject);
			this.ContinuationSeparator.Draw(nPageAbs, oGraphics);
		}

		for (var nEndnoteIndex = 0, nEndnotesCount = oColumn.Elements.length; nEndnoteIndex < nEndnotesCount; ++nEndnoteIndex)
		{
			var oEndnote = oColumn.Elements[nEndnoteIndex];
			var nEndnotePageIndex = oEndnote.GetElementPageIndex(nPageAbs, nColumnIndex);
			oEndnote.Draw(nEndnotePageIndex + oEndnote.StartPage, oGraphics);
		}
	}
};
CEndnotesController.prototype.StartSelection = function(X, Y, nPageAbs, oMouseEvent)
{
	if (true === this.Selection.Use)
		this.RemoveSelection();

	var oResult = this.private_GetEndnoteByXY(X, Y, nPageAbs);
	if (null === oResult)
	{
		// BAD
		this.Selection.Use = false;
		return;
	}

	this.Selection.Use   = true;
	this.Selection.Start = oResult;
	this.Selection.End   = oResult;

	this.Selection.Start.Endnote.Selection_SetStart(X, Y, this.Selection.Start.EndnotePageIndex, oMouseEvent);

	this.CurEndnote = this.Selection.Start.Endnote;

	this.Selection.Endnotes = {};
	this.Selection.Endnotes[this.Selection.Start.Endnote.GetId()] = this.Selection.Start.Endnote;
	this.Selection.Direction = 0;
};
CEndnotesController.prototype.EndSelection = function(X, Y, nPageAbs, oMouseEvent)
{
	if (true === this.IsMovingTableBorder())
	{
		this.CurEndnote.Selection_SetEnd(X, Y, nPageAbs, oMouseEvent);
		return;
	}

	var oResult = this.private_GetEndnoteByXY(X, Y, nPageAbs);
	if (null === oResult)
	{
		// BAD
		this.Selection.Use = false;
		return;
	}

	this.Selection.End = oResult;
	this.CurEndnote    = this.Selection.End.Endnote;

	var sStartId = this.Selection.Start.Endnote.GetId();
	var sEndId   = this.Selection.End.Endnote.GetId();

	// Очищаем старый селект везде кроме начальной сноски
	for (var sEndnoteId in this.Selection.Endnotes)
	{
		if (sEndnoteId !== sStartId)
			this.Selection.Endnotes[sEndnoteId].RemoveSelection();
	}

	// Новый селект
	if (this.Selection.Start.Endnote !== this.Selection.End.Endnote)
	{
		if (this.Selection.Start.Page > this.Selection.End.Page
			|| (this.Selection.Start.Page === this.Selection.End.Page
				&& (this.Selection.Start.Section > this.Selection.End.Section
					|| (this.Selection.Start.Section === this.Selection.End.Section
						&& (this.Selection.Start.Column > this.Selection.End.Column
							|| (this.Selection.Start.Column === this.Selection.End.Column
								&& this.Selection.Start.Index > this.Selection.End.Index))))))
		{
			this.Selection.Start.Endnote.Selection_SetEnd(-MEASUREMENT_MAX_MM_VALUE, -MEASUREMENT_MAX_MM_VALUE, 0, oMouseEvent);
			this.Selection.End.Endnote.Selection_SetStart(MEASUREMENT_MAX_MM_VALUE, MEASUREMENT_MAX_MM_VALUE, this.Selection.End.Endnote.Pages.length - 1, oMouseEvent);
			this.Selection.Direction = -1;
		}
		else
		{
			this.Selection.Start.Endnote.Selection_SetEnd(MEASUREMENT_MAX_MM_VALUE, MEASUREMENT_MAX_MM_VALUE, this.Selection.Start.Endnote.Pages.length - 1, oMouseEvent);
			this.Selection.End.Endnote.Selection_SetStart(-MEASUREMENT_MAX_MM_VALUE, -MEASUREMENT_MAX_MM_VALUE, 0, oMouseEvent);
			this.Selection.Direction = 1;
		}
		this.Selection.End.Endnote.Selection_SetEnd(X, Y, this.Selection.End.EndnotePageIndex, oMouseEvent);

		var oRange = this.private_GetEndnotesRange(this.Selection.Start, this.Selection.End);
		for (var sEndnoteId in oRange)
		{
			if (sEndnoteId !== sStartId && sEndnoteId !== sEndId)
			{
				var oEndnote = oRange[sEndnoteId];
				oEndnote.SelectAll();
			}
		}
		this.Selection.Endnotes = oRange;
	}
	else
	{
		this.Selection.End.Endnote.Selection_SetEnd(X, Y, this.Selection.End.EndnotePageIndex, oMouseEvent);
		this.Selection.Endnotes = {};
		this.Selection.Endnotes[this.Selection.Start.Endnote.GetId()] = this.Selection.Start.Endnote;
		this.Selection.Direction = 0;
	}
};
CEndnotesController.prototype.Set_CurrentElement = function(bUpdateStates, PageAbs, oEndnote)
{
	if (oEndnote instanceof CFootEndnote)
	{
		if (oEndnote.IsSelectionUse())
		{
			this.CurEndnote              = oEndnote;
			this.Selection.Use           = true;
			this.Selection.Direction     = 0;
			this.Selection.Start.Endnote = oEndnote;
			this.Selection.End.Endnote   = oEndnote;
			this.Selection.Endnotes      = {};

			this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
			this.LogicDocument.Selection.Use   = true;
			this.LogicDocument.Selection.Start = false;
		}
		else
		{
			this.private_SetCurrentEndnoteNoSelection(oEndnote);
			this.LogicDocument.Selection.Use   = false;
			this.LogicDocument.Selection.Start = false;
		}

		var bNeedRedraw = this.LogicDocument.GetDocPosType() === docpostype_HdrFtr;
		this.LogicDocument.SetDocPosType(docpostype_Endnotes);

		if (false !== bUpdateStates)
		{
			this.LogicDocument.UpdateInterface();
			this.LogicDocument.UpdateRulers();
			this.LogicDocument.UpdateSelection();
		}

		if (bNeedRedraw)
		{
			this.LogicDocument.DrawingDocument.ClearCachePages();
			this.LogicDocument.DrawingDocument.FirePaint();
		}
	}
};
CEndnotesController.prototype.AddEndnoteRef = function()
{
	if (true !== this.private_IsOneEndnoteSelected() || null === this.CurEndnote)
		return;

	var oEndnote   = this.CurEndnote;
	var oParagraph = oEndnote.GetFirstParagraph();
	if (!oParagraph)
		return;

	var oStyles = this.LogicDocument.GetStyles();

	var oRun = new ParaRun(oParagraph, false);
	oRun.AddToContent(0, new ParaEndnoteRef(oEndnote), false);
	oRun.SetRStyle(oStyles.GetDefaultEndnoteReference());
	oParagraph.Add_ToContent(0, oRun);
};
CEndnotesController.prototype.GetCurEndnote = function()
{
	return this.CurEndnote;
};
CEndnotesController.prototype.IsInDrawing = function(X, Y, nPageAbs)
{
	var oResult = this.private_GetEndnoteByXY(X, Y, nPageAbs);
	if (oResult)
	{
		var oEndnote = oResult.Endnote;
		return oEndnote.IsInDrawing(X, Y, oResult.EndnotePageIndex);
	}

	return false;
};
CEndnotesController.prototype.IsTableBorder = function(X, Y, nPageAbs)
{
	var oResult = this.private_GetEndnoteByXY(X, Y, nPageAbs);
	if (oResult)
	{
		var oEndnote = oResult.Endnote;
		return oEndnote.IsTableBorder(X, Y, oResult.EndnotePageIndex);
	}

	return null;
};
CEndnotesController.prototype.IsInText = function(X, Y, nPageAbs)
{
	var oResult = this.private_GetEndnoteByXY(X, Y, nPageAbs);
	if (oResult)
	{
		var oEndnote = oResult.Endnote;
		return oEndnote.IsInText(X, Y, oResult.EndnotePageIndex);
	}

	return null;
};
CEndnotesController.prototype.GetNearestPos = function(X, Y, nPageAbs, bAnchor, oDrawing)
{
	var oResult = this.private_GetEndnoteByXY(X, Y, nPageAbs);
	if (oResult)
	{
		var oEndnote = oResult.Endnote;
		return oEndnote.Get_NearestPos(oResult.EndnotePageIndex, X, Y, bAnchor, oDrawing);
	}

	return null;
};
/**
 * Проверяем попадание в сноски на заданной странице.
 * @param X
 * @param Y
 * @param nPageAbs
 * @returns {boolean}
 */
CEndnotesController.prototype.CheckHitInEndnote = function(X, Y, nPageAbs)
{
	var isCheckBottom = this.GetEndnotePrPos() === Asc.c_oAscEndnotePos.SectEnd;

	if (true === this.IsEmptyPage(nPageAbs))
		return false;

	var oPage = this.Pages[nPageAbs];
	for (var nIndex = 0, nCount = oPage.Sections.length; nIndex < nCount; ++nIndex)
	{
		var oSection = this.Sections[oPage.Sections[nIndex]];
		if (!oSection)
			continue;

		var _isCheckBottom = isCheckBottom;
		if (!_isCheckBottom && oPage.Sections[nIndex] === this.Sections.length - 1 && nPageAbs === this.Pages.length - 1)
			_isCheckBottom = false;

		var oSectionPage = oSection.Pages[nPageAbs];

		var oColumn = null;
		var nFindedColumnIndex = 0, nColumnsCount = oSectionPage.Columns.length;
		for (var nColumnIndex = 0; nColumnIndex < nColumnsCount; ++nColumnIndex)
		{
			if (nColumnIndex < nColumnsCount - 1)
			{
				if (X < (oSectionPage.Columns[nColumnIndex].XLimit + oSectionPage.Columns[nColumnIndex + 1].X) / 2)
				{
					oColumn            = oSectionPage.Columns[nColumnIndex];
					nFindedColumnIndex = nColumnIndex;
					break;
				}
			}
			else
			{
				oColumn            = oSectionPage.Columns[nColumnIndex];
				nFindedColumnIndex = nColumnIndex;
			}
		}

		if (!oColumn || nFindedColumnIndex >= nColumnsCount)
			return false;

		for (var nElementIndex = 0, nElementsCount = oColumn.Elements.length; nElementIndex < nElementsCount; ++nElementIndex)
		{
			var oEndnote          = oColumn.Elements[nElementIndex];
			var nEndnotePageIndex = oEndnote.GetElementPageIndex(nPageAbs, nFindedColumnIndex);
			var oBounds           = oEndnote.GetPageBounds(nEndnotePageIndex);

			if (oBounds.Top <= Y && (!isCheckBottom || oBounds.Bottom >= Y))
				return true;
		}
	}

	return false;
};
CEndnotesController.prototype.GetAllParagraphs = function(Props, ParaArray)
{
	for (var sId in this.Endnote)
	{
		var oEndnote = this.Endnote[sId];
		oEndnote.GetAllParagraphs(Props, ParaArray);
	}
};
CEndnotesController.prototype.GetAllTables = function(oProps, arrTables)
{
	if (!arrTables)
		arrTables = [];

	for (var sId in this.Endnote)
	{
		var oEndnote = this.Endnote[sId];
		oEndnote.GetAllTables(oProps, ParaArray);
	}

	return arrTables;
};
CEndnotesController.prototype.GetFirstParagraphs = function()
{
	var aParagraphs = [];
	for (var sId in this.Endnote)
	{
		var oEndnote = this.Endnote[sId];
		var oParagrpaph = oEndnote.GetFirstParagraph();
		if(oParagrpaph && oParagrpaph.Is_UseInDocument())
		{
			aParagraphs.push(oParagrpaph);
		}
	}
	return aParagraphs;
};

/**
 * Перенеслись ли сноски с предыдущей страницы, на новую
 * @param nPageAbs
 * @returns {boolean}
 */
CEndnotesController.prototype.IsContinueRecalculateFromPrevPage = function(nPageAbs)
{
	if (nPageAbs <= 0 || !this.Pages[nPageAbs - 1])
		return false;

	return (this.Pages[nPageAbs - 1].Sections.length > 0 && true === this.Pages[nPageAbs - 1].Continue);
};
CEndnotesController.prototype.GotoNextEndnote = function()
{
	var oNextEndnote = this.private_GetNextEndnote(this.CurEndnote);
	if (oNextEndnote)
	{
		oNextEndnote.MoveCursorToStartPos(false);
		this.private_SetCurrentEndnoteNoSelection(oNextEndnote);
	}
};
CEndnotesController.prototype.GotoPrevEndnote = function()
{
	var oPrevEndnote = this.private_GetPrevEndnote(this.CurEndnote);
	if (oPrevEndnote)
	{
		oPrevEndnote.MoveCursorToStartPos(false);
		this.private_SetCurrentEndnoteNoSelection(oPrevEndnote);
	}
};
CEndnotesController.prototype.GetNumberingInfo = function(oPara, oNumPr, oEndnote, isUseReview)
{
	var oNumberingEngine = new CDocumentNumberingInfoEngine(oPara, oNumPr, this.Get_Numbering());

	if (this.IsSpecialEndnote(oEndnote))
	{
		oEndnote.GetNumberingInfo(oNumberingEngine, oPara, oNumPr);
	}
	else
	{
		var arrEndnotes = this.LogicDocument.GetEndnotesList(null, oEndnote);
		for (var nIndex = 0, nCount = arrEndnotes.length; nIndex < nCount; ++nIndex)
		{
			arrEndnotes[nIndex].GetNumberingInfo(oNumberingEngine, oPara, oNumPr);
		}
	}

	if (true === isUseReview)
		return [oNumberingEngine.GetNumInfo(), oNumberingEngine.GetNumInfo(false)];

	return oNumberingEngine.GetNumInfo();
};
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Private area
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CEndnotesController.prototype.private_GetPageColumn = function(nPageAbs, nColumnAbs, nSectionAbs)
{
	var oSection = this.Sections[nSectionAbs];
	if (!oSection)
		return null;

	var oPage = oSection.Pages[nPageAbs];
	if (!oPage)
		return null;

	var oColumn = oPage.Columns[nColumnAbs];
	if (!oColumn)
		return null;

	return oColumn;
};
CEndnotesController.prototype.private_GetEndnoteOnPageByXY = function(X, Y, nPageAbs)
{
	if (true === this.IsEmptyPage(nPageAbs))
		return null;

	var oPage = this.Pages[nPageAbs];
	for (var nSectionIndex = oPage.Sections.length; nSectionIndex >= 0; --nSectionIndex)
	{
		var oSection = this.Sections[oPage.Sections[nSectionIndex]];
		if (!oSection)
			continue;

		var oSectionPage = oSection.Pages[nPageAbs];
		if (!oSectionPage)
			continue;

		var oColumn      = null;
		var nColumnIndex = 0;
		for (var nColumnsCount = oSectionPage.Columns.length; nColumnIndex < nColumnsCount; ++nColumnIndex)
		{
			if (nColumnIndex < nColumnsCount - 1)
			{
				if (X < (oSectionPage.Columns[nColumnIndex].XLimit + oSectionPage.Columns[nColumnIndex + 1].X) / 2)
				{
					oColumn = oSectionPage.Columns[nColumnIndex];
					break;
				}
			}
			else
			{
				oColumn = oSectionPage.Columns[nColumnIndex];
				break;
			}
		}

		if (!oColumn)
			continue;

		if (oColumn.Elements.length <= 0)
		{
			var nCurColumnIndex = nColumnIndex - 1;
			while (nCurColumnIndex >= 0)
			{
				if (oSectionPage.Columns[nCurColumnIndex].Elements.length > 0)
				{
					oColumn      = oSectionPage.Columns[nCurColumnIndex];
					nColumnIndex = nCurColumnIndex;
					break;
				}
				nCurColumnIndex--;
			}

			if (nCurColumnIndex < 0)
			{
				nCurColumnIndex = nColumnIndex + 1;
				while (nCurColumnIndex <= oSectionPage.Columns.length - 1)
				{
					if (oSectionPage.Columns[nCurColumnIndex].Elements.length > 0)
					{
						oColumn      = oSectionPage.Columns[nCurColumnIndex];
						nColumnIndex = nCurColumnIndex;
						break;
					}
					nCurColumnIndex++;
				}
			}
		}

		if (!oColumn || oColumn.Elements.length <= 0)
			continue;

		var nStartPos = oColumn.Elements.length - 1;
		if (nStartPos > 0)
		{
			var oEndnote = oColumn.Elements[nStartPos];
			if (oEndnote.IsEmptyPage(oEndnote.GetElementPageIndex(nPageAbs, nColumnIndex)))
				nStartPos--;
		}

		for (var nIndex = nStartPos; nIndex >= 0; --nIndex)
		{
			var oEndnote = oColumn.Elements[nIndex];

			var nElementPageIndex = oEndnote.GetElementPageIndex(nPageAbs, nColumnIndex);
			var oBounds           = oEndnote.GetPageBounds(nElementPageIndex);

			if (oBounds.Top <= Y || (0 === nIndex && 0 === nSectionIndex))
				return {
					Endnote          : oEndnote,
					Index            : nIndex,
					Section          : oPage.Sections[nSectionIndex],
					Page             : nPageAbs,
					Column           : nColumnIndex,
					EndnotePageIndex : nElementPageIndex
				};
		}
	}

	return null;
};
CEndnotesController.prototype.private_GetEndnoteByXY = function(X, Y, nPageAbs)
{
	var oResult = this.private_GetEndnoteOnPageByXY(X, Y, nPageAbs);
	if (null !== oResult)
		return oResult;

	var nCurPage = nPageAbs - 1;
	while (nCurPage >= 0)
	{
		oResult = this.private_GetEndnoteOnPageByXY(MEASUREMENT_MAX_MM_VALUE, MEASUREMENT_MAX_MM_VALUE, nCurPage);
		if (null !== oResult)
			return oResult;

		nCurPage--;
	}

	nCurPage = nPageAbs + 1;
	while (nCurPage < this.Pages.length)
	{
		oResult = this.private_GetEndnoteOnPageByXY(-MEASUREMENT_MAX_MM_VALUE, -MEASUREMENT_MAX_MM_VALUE, nCurPage);
		if (null !== oResult)
			return oResult;

		nCurPage++;
	}

	return null;
};
CEndnotesController.prototype.private_GetEndnotesRange = function(Start, End)
{
	var oResult = {};

	if (Start.Page > End.Page || (Start.Page === End.Page && (Start.Section > End.Section || (Start.Section === End.Section && (Start.Column > End.Column || (Start.Column === End.Column && Start.Index > End.Index))))))
	{
		var Temp = Start;
		Start    = End;
		End      = Temp;
	}

	if (Start.Page === End.Page)
	{
		this.private_GetEndnotesOnPage(Start.Page, Start.Section, Start.Column, Start.Index, End.Section, End.Column, End.Index, oResult);
	}
	else
	{
		this.private_GetEndnotesOnPage(Start.Page, Start.Section, Start.Column, Start.Index, -1, -1, -1, oResult);

		for (var nCurPage = Start.Page + 1; nCurPage <= End.Page - 1; ++nCurPage)
		{
			this.private_GetEndnotesOnPage(nCurPage, -1, -1, -1, -1, -1, -1, oResult);
		}

		this.private_GetEndnotesOnPage(End.Page, -1, -1, -1, End.Section, End.Column, End.Index, oResult);
	}

	return oResult;
};
CEndnotesController.prototype.private_GetEndnotesOnPage = function(nPageAbs, nSectionStart, nColumnStart, nStartIndex, nSectionEnd, nColumnEnd, nEndIndex, oEndnotes)
{
	var _nSectionStart = nSectionStart;
	var _nSectionEnd   = nSectionEnd;

	if (-1 === nSectionStart || -1 === nSectionEnd)
	{
		var oPage = this.Pages[nPageAbs];
		if (!oPage || oPage.Sections.length <= 0)
			return;

		if (-1 === nSectionStart)
			_nSectionStart = oPage.Sections[0];

		if (-1 === nSectionEnd)
			_nSectionEnd = oPage.Sections[oPage.Sections.length - 1];
	}

	for (var nSectionIndex = _nSectionStart; nSectionIndex <= _nSectionEnd; ++nSectionIndex)
	{
		var oSection = this.Sections[nSectionIndex];
		if (!oSection)
			return;

		var oSectionPage = oSection.Pages[nPageAbs];
		if (!oSectionPage)
			return;

		var _nColumnStart = -1 === nColumnStart || nSectionIndex !== _nSectionStart ? 0 : nColumnStart;
		var _nColumnEnd   = -1 === nColumnEnd || nSectionIndex !== _nSectionEnd ? oSectionPage.Columns.length - 1 : nColumnEnd;

		var _nStartIndex = -1 === nColumnStart || -1 === nStartIndex ? 0 : nStartIndex;
		var _nEndIndex   = -1 === nColumnEnd || -1 === nEndIndex ? oSectionPage.Columns[_nColumnEnd].Elements.length - 1 : nEndIndex;

		for (var nColIndex = _nColumnStart; nColIndex <= _nColumnEnd; ++nColIndex)
		{
			var nSIndex = (nSectionIndex === _nSectionStart && nColIndex === _nColumnStart) ? _nStartIndex : 0;
			var nEIndex = (nSectionIndex === _nSectionEnd && nColIndex === _nColumnEnd) ? _nEndIndex : oSectionPage.Columns[nColIndex].Elements.length - 1;

			this.private_GetEndnotesOnPageColumn(nPageAbs, nColIndex, nSectionIndex, nSIndex, nEIndex, oEndnotes);
		}
	}
};
CEndnotesController.prototype.private_GetEndnotesOnPageColumn = function(nPageAbs, nColumnAbs, nSectionAbs, nStartIndex, nEndIndex, oEndnotes)
{
	var oColumn = this.private_GetPageColumn(nPageAbs, nColumnAbs, nSectionAbs);

	var _StartIndex = -1 === nStartIndex ? 0 : nStartIndex;
	var _EndIndex   = -1 === nEndIndex ? oColumn.Elements.length - 1 : nEndIndex;

	for (var nIndex = _StartIndex; nIndex <= _EndIndex; ++nIndex)
	{
		var oEndnote = oColumn.Elements[nIndex];
		oEndnotes[oEndnote.GetId()] = oEndnote;
	}
};
CEndnotesController.prototype.private_OnNotValidActionForEndnotes = function()
{
	// Пока ничего не делаем, если надо будет выдавать сообщение, то обработать нужно будет здесь
};
CEndnotesController.prototype.private_IsOneEndnoteSelected = function()
{
	return (0 === this.Selection.Direction);
};
CEndnotesController.prototype.private_CheckEndnotesSelectionBeforeAction = function()
{
	if (true !== this.private_IsOneEndnoteSelected() || null === this.CurEndnote)
	{
		this.private_OnNotValidActionForEndnotes();
		return false;
	}

	return true;
};
CEndnotesController.prototype.private_SetCurrentEndnoteNoSelection = function(oEndnote)
{
	this.Selection.Use           = false;
	this.CurEndnote              = oEndnote;
	this.Selection.Start.Endnote = oEndnote;
	this.Selection.End.Endnote   = oEndnote;
	this.Selection.Direction     = 0;

	this.Selection.Endnotes                   = {};
	this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
};
CEndnotesController.prototype.private_GetPrevEndnote = function(oEndnote)
{
	if (!oEndnote)
		return null;

	var arrList = this.LogicDocument.GetEndnotesList(null, oEndnote);
	if (!arrList || arrList.length <= 1 || arrList[arrList.length - 1] !== oEndnote)
		return null;

	return arrList[arrList.length - 2];
};
CEndnotesController.prototype.private_GetNextEndnote = function(oEndnote)
{
	if (!oEndnote)
		return null;

	var arrList = this.LogicDocument.GetEndnotesList(oEndnote, null);
	if (!arrList || arrList.length <= 1 || arrList[0] !== oEndnote)
		return null;

	return arrList[1];
};
CEndnotesController.prototype.private_GetDirection = function(oEndote1, oEndnote2)
{
	// Предполагается, что эти сноски обязательно есть в документе
	if (oEndote1 === oEndnote2)
		return 0;

	var arrList = this.LogicDocument.GetEndnotesList(null, null);

	for (var nPos = 0, nCount = arrList.length; nPos < nCount; ++nPos)
	{
		if (oEndote1 === arrList[nPos])
			return 1;
		else if (oEndnote2 === arrList[nPos])
			return -1;
	}

	return 0;
};
CEndnotesController.prototype.private_GetEndnotesLogicRange = function(oEndote1, oEndnote2)
{
	return this.LogicDocument.GetEndnotesList(oEndote1, oEndnote2);
};
CEndnotesController.prototype.private_GetSelectionArray = function()
{
	if (true !== this.Selection.Use || 0 === this.Selection.Direction)
		return [this.CurEndnote];

	if (1 === this.Selection.Direction)
		return this.private_GetEndnotesLogicRange(this.Selection.Start.Endnote, this.Selection.End.Endnote);
	else
		return this.private_GetEndnotesLogicRange(this.Selection.End.Endnote, this.Selection.Start.Endnote);
};
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Controller area
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
CEndnotesController.prototype.CanUpdateTarget = function()
{
	var oLogicDocument = this.LogicDocument;
	if (!oLogicDocument)
		return false;

	var oCurEndnote = this.CurEndnote;
	if (null !== oLogicDocument.FullRecalc.Id && oCurEndnote)
	{
		var nPageAbs   = oLogicDocument.FullRecalc.PageIndex;

		var nSectionIndex = oCurEndnote.GetSectionIndex();

		var nLastIndex = this.LogicDocument.GetElementsCount() - 1;
		if (Asc.c_oAscEndnotePos.SectEnd === this.GetEndnotePrPos())
			nLastIndex = this.LogicDocument.SectionsInfo.Get(nSectionIndex).Index;

		if (oLogicDocument.FullRecalc.StartIndex < nLastIndex || (oLogicDocument.FullRecalc.StartIndex === nLastIndex && !oLogicDocument.FullRecalc.Endnotes))
			return false;

		if (!oLogicDocument.FullRecalc.Endnotes)
			return true;

		var _nSectionIndex = this.LogicDocument.SectionsInfo.Find(this.LogicDocument.SectionsInfo.Get_SectPr(oLogicDocument.FullRecalc.StartIndex).SectPr);
		if (_nSectionIndex < nSectionIndex)
			return false;
		else if (_nSectionIndex > nSectionIndex)
			return true;

		var oSection = this.Sections[nSectionIndex];
		if (!oSection)
			return false;

		var nStartPos = 0;

		if (nPageAbs - 1 <= oSection.StartPage || !oSection.Pages[nPageAbs - 1] || !oSection.Pages[nPageAbs - 1].Columns.length)
		{
			return false;
		}
		else
		{
			nStartPos = oSection.Pages[nPageAbs - 1].Columns[0].EndPos;
		}

		if (oSection.Endnotes[nStartPos] === this.CurEndnote)
		{
			var nEndnotePageIndex = this.CurEndnote.GetElementPageIndex(oLogicDocument.FullRecalc.PageIndex, oLogicDocument.FullRecalc.ColumnIndex);
			return this.CurEndnote.CanUpdateTarget(nEndnotePageIndex);
		}
		else
		{
			for (var nPos = 0; nPos < nStartPos; ++nPos)
			{
				if (this.CurEndnote === oSection.Endnotes[nPos])
					return true;
			}
		}

		return false;
	}

	return true;
};
CEndnotesController.prototype.RecalculateCurPos = function(bUpdateX, bUpdateY, isUpdateTarget)
{
	if (this.CurEndnote)
		return this.CurEndnote.RecalculateCurPos(bUpdateX, bUpdateY, isUpdateTarget);

	return {X : 0, Y : 0, Height : 0, PageNum : 0, Internal : {Line : 0, Page : 0, Range : 0}, Transform : null};
};
CEndnotesController.prototype.GetCurPage = function()
{
	if (this.CurEndnote)
		return this.CurEndnote.Get_StartPage_Absolute();

	return -1;
};
CEndnotesController.prototype.AddNewParagraph = function(bRecalculate, bForceAdd)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.AddNewParagraph(bRecalculate, bForceAdd);
};
CEndnotesController.prototype.AddInlineImage = function(nW, nH, oImage, oChart, bFlow)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.AddInlineImage(nW, nH, oImage, oChart, bFlow);
};
CEndnotesController.prototype.AddImages = function(aImages)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.AddImages(aImages);
};
CEndnotesController.prototype.AddOleObject = function(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.AddOleObject(W, H, nWidthPix, nHeightPix, Img, Data, sApplicationId, bSelect);
};
CEndnotesController.prototype.AddTextArt = function(nStyle)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.AddTextArt(nStyle);
};
CEndnotesController.prototype.AddSignatureLine = function(oSignatureDrawing)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.AddSignatureLine(oSignatureDrawing);
};
CEndnotesController.prototype.EditChart = function(oChartPr)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.EditChart(oChartPr);
};
CEndnotesController.prototype.AddInlineTable = function(nCols, nRows, nMode)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return null;

	if (this.CurEndnote)
		return this.CurEndnote.AddInlineTable(nCols, nRows, nMode);

	return null;
};
CEndnotesController.prototype.ClearParagraphFormatting = function(isClearParaPr, isClearTextPr)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.ClearParagraphFormatting(isClearParaPr, isClearTextPr);
	}
};
CEndnotesController.prototype.AddToParagraph = function(oItem, bRecalculate)
{
	if (para_NewLine === oItem.Type && true === oItem.IsPageBreak())
		return;

	if (oItem instanceof ParaTextPr)
	{
		for (var sId in this.Selection.Endnotes)
		{
			var oEndnote = this.Selection.Endnotes[sId];
			oEndnote.AddToParagraph(oItem, false);
		}

		if (false !== bRecalculate)
		{
			this.LogicDocument.Recalculate();
		}
	}
	else
	{
		if (false === this.private_CheckEndnotesSelectionBeforeAction())
			return;

		if (null !== this.CurEndnote)
			this.CurEndnote.AddToParagraph(oItem, bRecalculate);
	}
};
CEndnotesController.prototype.Remove = function(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.Remove(Count, bOnlyText, bRemoveOnlySelection, bOnTextAdd, isWord);
};
CEndnotesController.prototype.GetCursorPosXY = function()
{
	// Если есть селект, тогда конец селекта совпадает с CurEndnote
	if (null !== this.CurEndnote)
		return this.CurEndnote.GetCursorPosXY();

	return {X : 0, Y : 0}
};
CEndnotesController.prototype.MoveCursorToStartPos = function(AddToSelect)
{
	if (true !== AddToSelect)
	{
		this.LogicDocument.controller_MoveCursorToStartPos(false);
	}
	else
	{
		var oEndnote = this.CurEndnote;
		if (true === this.Selection.Use)
			oEndnote = this.Selection.Start.Endnote;

		var arrRange = this.LogicDocument.GetEndnotesList(null, oEndnote);
		if (arrRange.length <= 0)
			return;

		if (true !== this.Selection.Use)
			this.LogicDocument.StartSelectionFromCurPos();

		this.Selection.End.Endnote   = arrRange[0];
		this.Selection.Start.Endnote = oEndnote;
		this.Selection.Endnotes      = {};

		oEndnote.MoveCursorToStartPos(true);
		this.Selection.Endnotes = {};
		this.Selection.Endnotes[oEndnote.GetId()]  = oEndnote;
		for (var nIndex = 0, nCount = arrRange.length; nIndex < nCount; ++nIndex)
		{
			var oTempEndnote = arrRange[nIndex];
			if (oTempEndnote !== oEndnote)
			{
				oTempEndnote.SelectAll(-1);
				this.Selection.Endnotes[oTempEndnote.GetId()] = oTempEndnote;
			}
		}

		if (this.Selection.Start.Endnote !== this.Selection.End.Endnote)
			this.Selection.Direction = -1;
		else
			this.Selection.Direction = 0;
	}
};
CEndnotesController.prototype.MoveCursorToEndPos = function(AddToSelect)
{
	if (true !== AddToSelect)
	{
		this.LogicDocument.controller_MoveCursorToEndPos(false);
	}
	else
	{
		var oEndnote = this.CurEndnote;
		if (true === this.Selection.Use)
			oEndnote = this.Selection.Start.Endnote;

		var arrRange = this.LogicDocument.GetEndnotesList(oEndnote, null);
		if (arrRange.length <= 0)
			return;

		if (true !== this.Selection.Use)
			this.LogicDocument.StartSelectionFromCurPos();

		this.Selection.End.Endnote   = arrRange[arrRange.length - 1];
		this.Selection.Start.Endnote = oEndnote;
		this.Selection.Endnotes      = {};

		oEndnote.MoveCursorToEndPos(true);
		this.Selection.Endnotes = {};
		this.Selection.Endnotes[oEndnote.GetId()]  = oEndnote;
		for (var nIndex = 0, nCount = arrRange.length; nIndex < nCount; ++nIndex)
		{
			var oTempEndnote = arrRange[nIndex];
			if (oTempEndnote !== oEndnote)
			{
				oTempEndnote.SelectAll(1);
				this.Selection.Endnotes[oTempEndnote.Get_Id()] = oTempEndnote;
			}
		}

		if (this.Selection.Start.Endnote !== this.Selection.End.Endnote)
			this.Selection.Direction = 1;
		else
			this.Selection.Direction = 0;
	}
};
CEndnotesController.prototype.MoveCursorLeft = function(AddToSelect, Word)
{
	if (true === this.Selection.Use)
	{
		if (true !== AddToSelect)
		{
			var oEndnote = this.CurEndnote;
			if (0 === this.Selection.Direction)
				oEndnote = this.CurEndnote;
			else if (1 === this.Selection.Direction)
				oEndnote = this.Selection.Start.Endnote;
			else
				oEndnote = this.Selection.End.Endnote;

			for (var sId in this.Selection.Endnotes)
			{
				if (oEndnote !== this.Selection.Endnotes[sId])
					this.Selection.Endnotes[sId].RemoveSelection();
			}

			oEndnote.MoveCursorLeft(false, Word);
			oEndnote.RemoveSelection();
			this.private_SetCurrentEndnoteNoSelection(oEndnote);
		}
		else
		{
			var oEndnote = this.Selection.End.Endnote;
			if (false === oEndnote.MoveCursorLeft(true, Word))
			{
				var oPrevEndnote = this.private_GetPrevEndnote(oEndnote);
				if (null === oPrevEndnote)
					return false;

				if (1 !== this.Selection.Direction)
				{
					this.Selection.End.Endnote = oPrevEndnote;
					this.Selection.Direction   = -1;
					this.CurEndnote            = oPrevEndnote;

					this.Selection.Endnotes[oPrevEndnote.GetId()] = oPrevEndnote;

					oPrevEndnote.MoveCursorToEndPos(false, true);
					oPrevEndnote.MoveCursorLeft(true, Word);
				}
				else
				{
					this.Selection.End.Endnote = oPrevEndnote;
					this.CurEndnote            = oPrevEndnote;

					if (oPrevEndnote === this.Selection.Start.Endnote)
						this.Selection.Direction = 0;

					oEndnote.RemoveSelection();
					delete this.Selection.Endnotes[oEndnote.GetId()];

					oPrevEndnote.MoveCursorLeft(true, Word);
				}
			}
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.CurEndnote;

			this.Selection.Use            = true;
			this.Selection.Start.Endnote  = oEndnote;
			this.Selection.End.Endnote    = oEndnote;
			this.Selection.Endnotes       = {};
			this.Selection.Direction      = 0;

			this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
			if (false === oEndnote.MoveCursorLeft(true, Word))
			{
				var oPrevEndnote = this.private_GetPrevEndnote(oEndnote);
				if (null === oPrevEndnote)
					return false;

				this.Selection.End.Endnote = oPrevEndnote;
				this.Selection.Direction   = -1;
				this.CurEndnote            = oPrevEndnote;

				this.Selection.Endnotes[oPrevEndnote.GetId()] = oPrevEndnote;

				oPrevEndnote.MoveCursorToEndPos(false, true);
				oPrevEndnote.MoveCursorLeft(true, Word);
			}
		}
		else
		{
			var oEndnote = this.CurEndnote;
			if (false === oEndnote.MoveCursorLeft(false, Word))
			{
				var oPrevEndnote = this.private_GetPrevEndnote(oEndnote);
				if (null === oPrevEndnote)
					return false;

				this.Selection.Start.Endnote = oPrevEndnote;
				this.Selection.End.Endnote   = oPrevEndnote;
				this.Selection.Direction     = 0;
				this.CurEndnote              = oPrevEndnote;
				this.Selection.Endnotes      = {};

				this.Selection.Endnotes[oPrevEndnote.GetId()] = oPrevEndnote;

				oPrevEndnote.MoveCursorToEndPos(false);
			}
		}
	}

	return true;
};
CEndnotesController.prototype.MoveCursorRight = function(AddToSelect, Word, FromPaste)
{
	if (true === this.Selection.Use)
	{
		if (true !== AddToSelect)
		{
			var oEndnote = this.CurEndnote;
			if (0 === this.Selection.Direction)
				oEndnote = this.CurEndnote;
			else if (1 === this.Selection.Direction)
				oEndnote = this.Selection.End.Endnote;
			else
				oEndnote = this.Selection.Start.Endnote;

			for (var sId in this.Selection.Endnotes)
			{
				if (oEndnote !== this.Selection.Endnotes[sId])
					this.Selection.Endnotes[sId].RemoveSelection();
			}

			oEndnote.MoveCursorRight(false, Word, FromPaste);
			oEndnote.RemoveSelection();
			this.private_SetCurrentEndnoteNoSelection(oEndnote);
		}
		else
		{
			var oEndnote = this.Selection.End.Endnote;
			if (false === oEndnote.MoveCursorRight(true, Word, FromPaste))
			{
				var oNextEndnote = this.private_GetNextEndnote(oEndnote);
				if (null === oNextEndnote)
					return false;

				if (-1 !== this.Selection.Direction)
				{
					this.Selection.End.Endnote = oNextEndnote;
					this.Selection.Direction   = 1;
					this.CurEndnote            = oNextEndnote;

					this.Selection.Endnotes[oNextEndnote.GetId()] = oNextEndnote;

					oNextEndnote.MoveCursorToStartPos(false);
					oNextEndnote.MoveCursorRight(true, Word, FromPaste);
				}
				else
				{
					this.Selection.End.Endnote = oNextEndnote;
					this.CurEndnote            = oNextEndnote;

					if (oNextEndnote === this.Selection.Start.Endnote)
						this.Selection.Direction = 0;

					oEndnote.RemoveSelection();
					delete this.Selection.Endnotes[oEndnote.GetId()];

					oNextEndnote.MoveCursorRight(true, Word, FromPaste);
				}
			}
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.CurEndnote;

			this.Selection.Use           = true;
			this.Selection.Start.Endnote = oEndnote;
			this.Selection.End.Endnote   = oEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Direction     = 0;

			this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
			if (false === oEndnote.MoveCursorRight(true, Word, FromPaste))
			{
				var oNextEndnote = this.private_GetNextEndnote(oEndnote);
				if (null === oNextEndnote)
					return false;

				this.Selection.End.Endnote = oNextEndnote;
				this.Selection.Direction   = 1;
				this.CurEndnote            = oNextEndnote;

				this.Selection.Endnotes[oNextEndnote.GetId()] = oNextEndnote;

				oNextEndnote.MoveCursorToStartPos(false);
				oNextEndnote.MoveCursorRight(true, Word, FromPaste);
			}
		}
		else
		{
			var oEndnote = this.CurEndnote;
			if (false === oEndnote.MoveCursorRight(false, Word, FromPaste))
			{
				var oNextEndnote = this.private_GetNextEndnote(oEndnote);
				if (null === oNextEndnote)
					return false;

				this.Selection.Start.Endnote = oNextEndnote;
				this.Selection.End.Endnote   = oNextEndnote;
				this.Selection.Direction     = 0;
				this.CurEndnote              = oNextEndnote;
				this.Selection.Endnotes      = {};

				this.Selection.Endnotes[oNextEndnote.GetId()] = oNextEndnote;

				oNextEndnote.MoveCursorToStartPos(false);
			}
		}
	}

	return true;
};
CEndnotesController.prototype.MoveCursorUp = function(AddToSelect)
{
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.Selection.End.Endnote;
			var oPos     = oEndnote.GetCursorPosXY();

			if (false === oEndnote.MoveCursorUp(true))
			{
				var oPrevEndnote = this.private_GetPrevEndnote(oEndnote);
				if (null === oPrevEndnote)
					return false;

				oEndnote.MoveCursorToStartPos(true);

				if (1 !== this.Selection.Direction)
				{
					this.Selection.End.Endnote = oPrevEndnote;
					this.Selection.Direction   = -1;
					this.CurEndnote            = oPrevEndnote;

					this.Selection.Endnotes[oPrevEndnote.GetId()] = oPrevEndnote;

					oPrevEndnote.MoveCursorUpToLastRow(oPos.X, oPos.Y, true);
				}
				else
				{
					this.Selection.End.Endnote = oPrevEndnote;
					this.CurEndnote            = oPrevEndnote;

					if (oPrevEndnote === this.Selection.Start.Endnote)
						this.Selection.Direction = 0;

					oEndnote.RemoveSelection();
					delete this.Selection.Endnotes[oEndnote.GetId()];

					oPrevEndnote.MoveCursorUpToLastRow(oPos.X, oPos.Y, true);
				}

			}
		}
		else
		{
			var oEndnote = this.CurEndnote;
			if (0 === this.Selection.Direction)
				oEndnote = this.CurEndnote;
			else if (1 === this.Selection.Direction)
				oEndnote = this.Selection.Start.Endnote;
			else
				oEndnote = this.Selection.End.Endnote;

			for (var sId in this.Selection.Endnotes)
			{
				if (oEndnote !== this.Selection.Endnotes[sId])
					this.Selection.Endnotes[sId].RemoveSelection();
			}

			oEndnote.MoveCursorLeft(false, false);
			oEndnote.RemoveSelection();
			this.private_SetCurrentEndnoteNoSelection(oEndnote);
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.CurEndnote;
			var oPos     = oEndnote.GetCursorPosXY();

			this.Selection.Use           = true;
			this.Selection.Start.Endnote = oEndnote;
			this.Selection.End.Endnote   = oEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Direction     = 0;

			this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
			if (false === oEndnote.MoveCursorUp(true))
			{
				var oPrevEndnote = this.private_GetPrevEndnote(oEndnote);
				if (null === oPrevEndnote)
					return false;

				oEndnote.MoveCursorToStartPos(true);

				this.Selection.End.Endnote = oPrevEndnote;
				this.Selection.Direction   = -1;
				this.CurEndnote            = oPrevEndnote;

				this.Selection.Endnotes[oPrevEndnote.GetId()] = oPrevEndnote;

				oPrevEndnote.MoveCursorUpToLastRow(oPos.X, oPos.Y, true);
			}
		}
		else
		{
			var oEndnote = this.CurEndnote;
			var oPos     = oEndnote.GetCursorPosXY();
			if (false === oEndnote.MoveCursorUp(false))
			{
				var oPrevEndnote = this.private_GetPrevEndnote(oEndnote);
				if (null === oPrevEndnote)
					return false;

				this.Selection.Start.Endnote = oPrevEndnote;
				this.Selection.End.Endnote   = oPrevEndnote;
				this.Selection.Direction     = 0;
				this.CurEndnote              = oPrevEndnote;
				this.Selection.Endnotes      = {};

				this.Selection.Endnotes[oPrevEndnote.GetId()] = oPrevEndnote;

				oPrevEndnote.MoveCursorUpToLastRow(oPos.X, oPos.Y, false);
			}
		}
	}

	return true;
};
CEndnotesController.prototype.MoveCursorDown = function(AddToSelect)
{
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.Selection.End.Endnote;
			var oPos     = oEndnote.GetCursorPosXY();

			if (false === oEndnote.MoveCursorDown(true))
			{
				var oNextEndnote = this.private_GetNextEndnote(oEndnote);
				if (null === oNextEndnote)
					return false;

				oEndnote.MoveCursorToEndPos(true);

				if (-1 !== this.Selection.Direction)
				{
					this.Selection.End.Endnote = oNextEndnote;
					this.Selection.Direction   = 1;
					this.CurEndnote            = oNextEndnote;

					this.Selection.Endnotes[oNextEndnote.GetId()] = oNextEndnote;

					oNextEndnote.MoveCursorDownToFirstRow(oPos.X, oPos.Y, true);
				}
				else
				{
					this.Selection.End.Endnote = oNextEndnote;
					this.CurEndnote            = oNextEndnote;

					if (oNextEndnote === this.Selection.Start.Endnote)
						this.Selection.Direction = 0;

					oEndnote.RemoveSelection();
					delete this.Selection.Endnotes[oEndnote.GetId()];

					oNextEndnote.MoveCursorDownToFirstRow(oPos.X, oPos.Y, true);
				}

			}
		}
		else
		{
			var oEndnote = this.CurEndnote;
			if (0 === this.Selection.Direction)
				oEndnote = this.CurEndnote;
			else if (1 === this.Selection.Direction)
				oEndnote = this.Selection.End.Endnote;
			else
				oEndnote = this.Selection.Start.Endnote;

			for (var sId in this.Selection.Endnotes)
			{
				if (oEndnote !== this.Selection.Endnotes[sId])
					this.Selection.Endnotes[sId].RemoveSelection();
			}

			oEndnote.MoveCursorRight(false, false);
			oEndnote.RemoveSelection();
			this.private_SetCurrentEndnoteNoSelection(oEndnote);
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.CurEndnote;
			var oPos     = oEndnote.GetCursorPosXY();

			this.Selection.Use           = true;
			this.Selection.Start.Endnote = oEndnote;
			this.Selection.End.Endnote   = oEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Direction     = 0;

			this.Selection.Endnotes[oEndnote.Get_Id()] = oEndnote;
			if (false === oEndnote.MoveCursorDown(true))
			{
				var oNextEndnote = this.private_GetNextEndnote(oEndnote);
				if (null === oNextEndnote)
					return false;

				oEndnote.MoveCursorToEndPos(true, false);

				this.Selection.End.Endnote = oNextEndnote;
				this.Selection.Direction   = 1;
				this.CurEndnote            = oNextEndnote;

				this.Selection.Endnotes[oNextEndnote.Get_Id()] = oNextEndnote;

				oNextEndnote.MoveCursorDownToFirstRow(oPos.X, oPos.Y, true);
			}
		}
		else
		{
			var oEndnote = this.CurEndnote;
			var oPos     = oEndnote.GetCursorPosXY();
			if (false === oEndnote.MoveCursorDown(false))
			{
				var oNextEndnote = this.private_GetNextEndnote(oEndnote);
				if (null === oNextEndnote)
					return false;

				this.Selection.Start.Endnote = oNextEndnote;
				this.Selection.End.Endnote   = oNextEndnote;
				this.Selection.Direction     = 0;
				this.CurEndnote              = oNextEndnote;
				this.Selection.Endnotes      = {};

				this.Selection.Endnotes[oNextEndnote.GetId()] = oNextEndnote;

				oNextEndnote.MoveCursorDownToFirstRow(oPos.X, oPos.Y, false);
			}
		}
	}

	return true;
};
CEndnotesController.prototype.MoveCursorToEndOfLine = function(AddToSelect)
{
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.Selection.End.Endnote;
			oEndnote.MoveCursorToEndOfLine(true);
		}
		else
		{
			var oEndnote = null;
			if (0 === this.Selection.Direction)
				oEndnote = this.CurEndnote;
			else if (1 === this.Selection.Direction)
				oEndnote = this.Selection.End.Endnote;
			else
				oEndnote = this.Selection.Start.Endnote;

			for (var sId in this.Selection.Endnotes)
			{
				if (oEndnote !== this.Selection.Endnotes[sId])
					this.Selection.Endnotes[sId].RemoveSelection();
			}

			oEndnote.MoveCursorToEndOfLine(false);
			oEndnote.RemoveSelection();
			this.private_SetCurrentEndnoteNoSelection(oEndnote);
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.CurEndnote;

			this.Selection.Use           = true;
			this.Selection.Start.Endnote = oEndnote;
			this.Selection.End.Endnote   = oEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Direction     = 0;

			this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
			oEndnote.MoveCursorToEndOfLine(true);
		}
		else
		{
			this.CurEndnote.MoveCursorToEndOfLine(false);
		}
	}

	return true;
};
CEndnotesController.prototype.MoveCursorToStartOfLine = function(AddToSelect)
{
	if (true === this.Selection.Use)
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.Selection.End.Endnote;
			oEndnote.MoveCursorToStartOfLine(true);
		}
		else
		{
			var oEndnote = null;
			if (0 === this.Selection.Direction)
				oEndnote = this.CurEndnote;
			else if (1 === this.Selection.Direction)
				oEndnote = this.Selection.Start.Endnote;
			else
				oEndnote = this.Selection.End.Endnote;

			for (var sId in this.Selection.Endnotes)
			{
				if (oEndnote !== this.Selection.Endnotes[sId])
					this.Selection.Endnotes[sId].RemoveSelection();
			}

			oEndnote.MoveCursorToStartOfLine(false);
			oEndnote.RemoveSelection();
			this.private_SetCurrentEndnoteNoSelection(oEndnote);
		}
	}
	else
	{
		if (true === AddToSelect)
		{
			var oEndnote = this.CurEndnote;

			this.Selection.Use            = true;
			this.Selection.Start.Endnote = oEndnote;
			this.Selection.End.Endnote   = oEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Direction      = 0;

			this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
			oEndnote.MoveCursorToStartOfLine(true);
		}
		else
		{
			this.CurEndnote.MoveCursorToStartOfLine(false);
		}
	}

	return true;
};
CEndnotesController.prototype.MoveCursorToXY = function(X, Y, PageAbs, AddToSelect)
{
	var oResult = this.private_GetEndnoteByXY(X, Y, PageAbs);
	if (!oResult || !oResult.Endnote)
		return;

	var oEndnote = oResult.Endnote;
	var PageRel   = oResult.EndnotePageIndex;
	if (true === AddToSelect)
	{
		var StartEndnote = null;
		if (true === this.Selection.Use)
		{
			StartEndnote = this.Selection.Start.Endnote;
			for (var sId in this.Selection.Endnotes)
			{
				if (this.Selection.Endnotes[sId] !== StartEndnote)
				{
					this.Selection.Endnotes[sId].RemoveSelection();
				}
			}
		}
		else
		{
			StartEndnote = this.CurEndnote;
		}

		var nDirection = this.private_GetDirection(StartEndnote, oEndnote);
		if (0 === nDirection)
		{
			this.Selection.Use            = true;
			this.Selection.Start.Endnote = oEndnote;
			this.Selection.End.Endnote   = oEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Direction      = 0;

			this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
			oEndnote.MoveCursorToXY(X, Y, true, true, PageRel);
		}
		else if (nDirection > 0)
		{
			var arrEndnotes = this.private_GetEndnotesLogicRange(StartEndnote, oEndnote);
			if (arrEndnotes.length <= 1)
				return;

			var oStartEndnote = arrEndnotes[0]; // StartEndnote
			var oEndEndnote   = arrEndnotes[arrEndnotes.length - 1]; // oEndnote

			this.Selection.Use            = true;
			this.Selection.Start.Endnote = oStartEndnote;
			this.Selection.End.Endnote   = oEndEndnote;
			this.CurEndnote              = oEndEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Direction      = 1;

			oStartEndnote.MoveCursorToEndPos(true, false);

			for (var nPos = 0, nCount = arrEndnotes.length; nPos < nCount; ++nPos)
			{
				this.Selection.Endnotes[arrEndnotes[nPos].GetId()] = arrEndnotes[nPos];

				if (0 !== nPos && nPos !== nCount - 1)
					arrEndnotes[nPos].SelectAll(1);
			}

			oEndEndnote.MoveCursorToStartPos(false);
			oEndEndnote.MoveCursorToXY(X, Y, true, true, PageRel);
		}
		else if (nDirection < 0)
		{
			var arrEndnotes = this.private_GetEndnotesLogicRange(oEndnote, StartEndnote);
			if (arrEndnotes.length <= 1)
				return;

			var oEndEndnote   = arrEndnotes[0]; // oEndnote
			var oStartEndnote = arrEndnotes[arrEndnotes.length - 1]; // StartEndnote

			this.Selection.Use            = true;
			this.Selection.Start.Endnote = oStartEndnote;
			this.Selection.End.Endnote   = oEndEndnote;
			this.CurEndnote              = oEndEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Direction      = -1;

			oStartEndnote.MoveCursorToStartPos(true);

			for (var nPos = 0, nCount = arrEndnotes.length; nPos < nCount; ++nPos)
			{
				this.Selection.Endnotes[arrEndnotes[nPos].GetId()] = arrEndnotes[nPos];

				if (0 !== nPos && nPos !== nCount - 1)
					arrEndnotes[nPos].SelectAll(-1);
			}

			oEndEndnote.MoveCursorToEndPos(false, true);
			oEndEndnote.MoveCursorToXY(X, Y, true, true, PageRel);
		}
	}
	else
	{
		if (true === this.Selection.Use)
		{
			this.RemoveSelection();
		}

		this.private_SetCurrentEndnoteNoSelection(oEndnote);
		oEndnote.MoveCursorToXY(X, Y, false, true, PageRel);
	}
};
CEndnotesController.prototype.MoveCursorToCell = function(bNext)
{
	if (true !== this.private_IsOneEndnoteSelected() || null === this.CurEndnote)
		return false;

	return this.CurEndnote.MoveCursorToCell(bNext);
};
CEndnotesController.prototype.SetParagraphAlign = function(Align)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphAlign(Align);
	}
};
CEndnotesController.prototype.SetParagraphSpacing = function(Spacing)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphSpacing(Spacing);
	}
};
CEndnotesController.prototype.SetParagraphTabs = function(Tabs)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphTabs(Tabs);
	}
};
CEndnotesController.prototype.SetParagraphIndent = function(Ind)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphIndent(Ind);
	}
};
CEndnotesController.prototype.SetParagraphShd = function(Shd)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphShd(Shd);
	}
};
CEndnotesController.prototype.SetParagraphStyle = function(Name)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphStyle(Name);
	}
};
CEndnotesController.prototype.SetParagraphContextualSpacing = function(Value)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphContextualSpacing(Value);
	}
};
CEndnotesController.prototype.SetParagraphPageBreakBefore = function(Value)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphPageBreakBefore(Value);
	}
};
CEndnotesController.prototype.SetParagraphKeepLines = function(Value)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphKeepLines(Value);
	}
};
CEndnotesController.prototype.SetParagraphKeepNext = function(Value)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphKeepNext(Value);
	}
};
CEndnotesController.prototype.SetParagraphWidowControl = function(Value)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphWidowControl(Value);
	}
};
CEndnotesController.prototype.SetParagraphBorders = function(Borders)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.SetParagraphBorders(Borders);
	}
};
CEndnotesController.prototype.SetParagraphFramePr = function(FramePr, bDelete)
{
	// Не позволяем делать рамки внутри сносок
};
CEndnotesController.prototype.IncreaseDecreaseFontSize = function(bIncrease)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.IncreaseDecreaseFontSize(bIncrease);
	}
};
CEndnotesController.prototype.IncreaseDecreaseIndent = function(bIncrease)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.IncreaseDecreaseIndent(bIncrease);
	}
};
CEndnotesController.prototype.SetImageProps = function(Props)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	return this.CurEndnote.SetImageProps(Props);
};
CEndnotesController.prototype.SetTableProps = function(Props)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	return this.CurEndnote.SetTableProps(Props);
};
CEndnotesController.prototype.GetCalculatedParaPr = function()
{
	var oStartPr = this.CurEndnote.GetCalculatedParaPr();
	var oPr      = oStartPr.Copy();

	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		var oTempPr  = oEndnote.GetCalculatedParaPr();
		oPr          = oPr.Compare(oTempPr);
	}

	if (undefined === oPr.Ind.Left)
		oPr.Ind.Left = oStartPr.Ind.Left;

	if (undefined === oPr.Ind.Right)
		oPr.Ind.Right = oStartPr.Ind.Right;

	if (undefined === oPr.Ind.FirstLine)
		oPr.Ind.FirstLine = oStartPr.Ind.FirstLine;

	if (true !== this.private_IsOneEndnoteSelected())
		oPr.CanAddTable = false;

	return oPr;
};
CEndnotesController.prototype.GetCalculatedTextPr = function()
{
	var oStartPr = this.CurEndnote.GetCalculatedTextPr();
	var oPr      = oStartPr.Copy();

	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		var oTempPr  = oEndnote.GetCalculatedTextPr();
		oPr          = oPr.Compare(oTempPr);
	}

	return oPr;
};
CEndnotesController.prototype.GetDirectParaPr = function()
{
	if (this.CurEndnote)
		return this.CurEndnote.GetDirectParaPr();

	return new CParaPr();
};
CEndnotesController.prototype.GetDirectTextPr = function()
{
	if (this.CurEndnote)
		return this.CurEndnote.GetDirectTextPr();

	return new CTextPr();
};
CEndnotesController.prototype.RemoveSelection = function(bNoCheckDrawing)
{
	if (true === this.Selection.Use)
	{
		for (var sId in this.Selection.Endnotes)
		{
			this.Selection.Endnotes[sId].RemoveSelection(bNoCheckDrawing);
		}

		this.Selection.Use = false;
	}

	this.Selection.Endnotes = {};
	if (this.CurEndnote)
		this.Selection.Endnotes[this.CurEndnote.GetId()] = this.CurEndnote;
};
CEndnotesController.prototype.IsSelectionEmpty = function(bCheckHidden)
{
	if (true !== this.IsSelectionUse())
		return true;

	var oEndnote = null;
	for (var sId in this.Selection.Endnotes)
	{
		if (null === oEndnote)
			oEndnote = this.Selection.Endnotes[sId];
		else if (oEndnote !== this.Selection.Endnotes[sId])
			return false;
	}

	if (null === oEndnote)
		return true;

	return oEndnote.IsSelectionEmpty(bCheckHidden);
};
CEndnotesController.prototype.DrawSelectionOnPage = function(nPageAbs)
{
	if (true !== this.Selection.Use || true === this.IsEmptyPage(nPageAbs))
		return;

	var oPage = this.Pages[nPageAbs];
	if (!oPage)
		return;


	for (var nSectionIndex = 0, nSectionsCount = oPage.Sections.length; nSectionIndex < nSectionsCount; ++nSectionIndex)
	{
		var oSection = this.Sections[oPage.Sections[nSectionIndex]];
		if (!oSection)
			continue;

		var oSectionPage = oSection.Pages[nPageAbs];
		if (!oSectionPage)
			continue;

		for (var nColumnIndex = 0, nColumnsCount = oSectionPage.Columns.length; nColumnIndex < nColumnsCount; ++nColumnIndex)
		{
			var oColumn = oSectionPage.Columns[nColumnIndex];
			for (var nIndex = 0, nCount = oColumn.Elements.length; nIndex < nCount; ++nIndex)
			{
				var oEndnote = oColumn.Elements[nIndex];
				if (oEndnote === this.Selection.Endnotes[oEndnote.GetId()])
				{
					var nEndnotePageIndex = oEndnote.GetElementPageIndex(nPageAbs, nColumnIndex);
					oEndnote.DrawSelectionOnPage(nEndnotePageIndex);
				}
			}
		}
	}
};
CEndnotesController.prototype.GetSelectionBounds = function()
{
	if (true === this.Selection.Use)
	{
		if (0 === this.Selection.Direction)
		{
			return this.CurEndnote.GetSelectionBounds();
		}
		else if (1 === this.Selection.Direction)
		{
			var StartBounds = this.Selection.Start.Endnote.GetSelectionBounds();
			var EndBounds   = this.Selection.End.Endnote.GetSelectionBounds();

			if (!StartBounds && !EndBounds)
				return null;

			var Result       = {};
			Result.Start     = StartBounds ? StartBounds.Start : EndBounds.Start;
			Result.End       = EndBounds ? EndBounds.End : StartBounds.End;
			Result.Direction = 1;
			return Result;
		}
		else
		{
			var StartBounds = this.Selection.End.Endnote.GetSelectionBounds();
			var EndBounds   = this.Selection.Start.Endnote.GetSelectionBounds();

			if (!StartBounds && !EndBounds)
				return null;

			var Result       = {};
			Result.Start     = StartBounds ? StartBounds.Start : EndBounds.Start;
			Result.End       = EndBounds ? EndBounds.End : StartBounds.End;
			Result.Direction = -1;
			return Result;
		}
	}

	return null;
};
CEndnotesController.prototype.IsMovingTableBorder = function()
{
	if (true !== this.private_IsOneEndnoteSelected())
		return false;

	return this.CurEndnote.IsMovingTableBorder();
};
CEndnotesController.prototype.CheckPosInSelection = function(X, Y, PageAbs, NearPos)
{
	var oResult = this.private_GetEndnoteByXY(X, Y, PageAbs);
	if (oResult)
	{
		var oEndnote = oResult.Endnote;
		return oEndnote.CheckPosInSelection(X, Y, oResult.EndnotePageIndex, NearPos);
	}

	return false;
};
CEndnotesController.prototype.GetSelectedContent = function(SelectedContent)
{
	if (true !== this.Selection.Use)
		return;

	if (0 === this.Selection.Direction)
	{
		this.CurEndnote.GetSelectedContent(SelectedContent);
	}
	else
	{
		var arrEndnotes = this.private_GetSelectionArray();
		for (var nPos = 0, nCount = arrEndnotes.length; nPos < nCount; ++nPos)
		{
			arrEndnotes[nPos].GetSelectedContent(SelectedContent);
		}
	}
};
CEndnotesController.prototype.UpdateCursorType = function(X, Y, PageAbs, MouseEvent)
{
	var oResult = this.private_GetEndnoteByXY(X, Y, PageAbs);
	if (oResult)
	{
		var oEndnote = oResult.Endnote;
		oEndnote.UpdateCursorType(X, Y, oResult.EndnotePageIndex, MouseEvent);
	}
};
CEndnotesController.prototype.PasteFormatting = function(TextPr, ParaPr)
{
	for (var sId in this.Selection.Endnotes)
	{
		this.Selection.Endnotes[sId].PasteFormatting(TextPr, ParaPr, true);
	}
};
CEndnotesController.prototype.IsSelectionUse = function()
{
	return this.Selection.Use;
};
CEndnotesController.prototype.IsNumberingSelection = function()
{
	if (this.CurEndnote)
		return this.CurEndnote.IsNumberingSelection();

	return false;
};
CEndnotesController.prototype.IsTextSelectionUse = function()
{
	if (true !== this.Selection.Use)
		return false;

	if (0 === this.Selection.Direction)
		return this.CurEndnote.IsTextSelectionUse();

	return true;
};
CEndnotesController.prototype.GetCurPosXY = function()
{
	if (this.CurEndnote)
		return this.CurEndnote.GetCurPosXY();

	return {X : 0, Y : 0};
};
CEndnotesController.prototype.GetSelectedText = function(bClearText, oPr)
{
	if (true === bClearText)
	{
		if (true !== this.Selection.Use || 0 !== this.Selection.Direction)
			return "";

		return this.CurEndnote.GetSelectedText(true, oPr);
	}
	else
	{
		var sResult = "";
		var arrEndnotes = this.private_GetSelectionArray();
		for (var nPos = 0, nCount = arrEndnotes.length; nPos < nCount; ++nPos)
		{
			var sTempResult = arrEndnotes[nPos].GetSelectedText(false, oPr);
			if (null == sTempResult)
				return null;

			sResult += sTempResult;
		}

		return sResult;
	}
};
CEndnotesController.prototype.GetCurrentParagraph = function(bIgnoreSelection, arrSelectedParagraphs, oPr)
{
	return this.CurEndnote.GetCurrentParagraph(bIgnoreSelection, arrSelectedParagraphs, oPr);
};
CEndnotesController.prototype.GetCurrentTablesStack = function(arrTables)
{
	if (!arrTables)
		arrTables = [];

	return this.CurEndnote.GetCurrentTablesStack(arrTables);
};
CEndnotesController.prototype.GetSelectedElementsInfo = function(oInfo)
{
	if (true !== this.private_IsOneEndnoteSelected() || null === this.CurEndnote)
		oInfo.SetMixedSelection();
	else
		this.CurEndnote.GetSelectedElementsInfo(oInfo);
};
CEndnotesController.prototype.AddTableRow = function(bBefore)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.AddTableRow(bBefore);
};
CEndnotesController.prototype.AddTableColumn = function(bBefore)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.AddTableColumn(bBefore);
};
CEndnotesController.prototype.RemoveTableRow = function()
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.RemoveTableRow();
};
CEndnotesController.prototype.RemoveTableColumn = function()
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.RemoveTableColumn();
};
CEndnotesController.prototype.MergeTableCells = function()
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.MergeTableCells();
};
CEndnotesController.prototype.SplitTableCells = function(Cols, Rows)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.SplitTableCells(Cols, Rows);
};
CEndnotesController.prototype.RemoveTableCells = function()
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.RemoveTableCells();
};
CEndnotesController.prototype.RemoveTable = function()
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.CurEndnote.RemoveTable();
};
CEndnotesController.prototype.SelectTable = function(Type)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return;

	this.RemoveSelection();

	this.CurEndnote.SelectTable(Type);
	if (true === this.CurEndnote.IsSelectionUse())
	{
		this.Selection.Use           = true;
		this.Selection.Start.Endnote = this.CurEndnote;
		this.Selection.End.Endnote   = this.CurEndnote;
		this.Selection.Endnotes      = {};
		this.Selection.Direction     = 0;

		this.Selection.Endnotes[this.CurEndnote.GetId()] = this.CurEndnote;
	}
};
CEndnotesController.prototype.CanMergeTableCells = function()
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.CanMergeTableCells();
};
CEndnotesController.prototype.CanSplitTableCells = function()
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.CanSplitTableCells();
};
CEndnotesController.prototype.DistributeTableCells = function(isHorizontally)
{
	if (false === this.private_CheckEndnotesSelectionBeforeAction())
		return false;

	return this.CurEndnote.DistributeTableCells(isHorizontally);
};
CEndnotesController.prototype.UpdateInterfaceState = function()
{
	if (this.private_IsOneEndnoteSelected())
	{
		this.CurEndnote.Document_UpdateInterfaceState();
	}
	else
	{
		var oApi = this.LogicDocument.GetApi();
		if (!oApi)
			return;

		var oParaPr = this.GetCalculatedParaPr();

		if (oParaPr.Tabs)
			oApi.Update_ParaTab(AscCommonWord.Default_Tab_Stop, oParaPr.Tabs);

		oApi.UpdateParagraphProp(oParaPr);
		oApi.UpdateTextPr(this.GetCalculatedTextPr());
	}
};
CEndnotesController.prototype.UpdateRulersState = function()
{
	var nPageAbs = this.CurEndnote.Get_StartPage_Absolute();
	if (this.LogicDocument.Pages[nPageAbs])
	{
		var nPos    = this.LogicDocument.Pages[nPageAbs].Pos;
		var oSectPr = this.LogicDocument.SectionsInfo.Get_SectPr(nPos).SectPr;
		var oFrame  = oSectPr.GetContentFrame(nPageAbs);

		this.DrawingDocument.Set_RulerState_Paragraph({L : oFrame.Left, T : oFrame.Top, R : oFrame.Right, B : oFrame.Bottom}, true);
	}

	if (true === this.private_IsOneEndnoteSelected())
		this.CurEndnote.Document_UpdateRulersState();
};
CEndnotesController.prototype.UpdateSelectionState = CFootnotesController.prototype.UpdateSelectionState;
CEndnotesController.prototype.GetSelectionState = function()
{
	var arrResult = [];

	var oState = {
		Endnotes   : {},
		Use        : this.Selection.Use,
		Start      : this.Selection.Start.Endnote,
		End        : this.Selection.End.Endnote,
		Direction  : this.Selection.Direction,
		CurEndnote : this.CurEndnote
	};

	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];

		oState.Endnotes[sId] = {
			Endnote : oEndnote,
			State   : oEndnote.GetSelectionState()
		};
	}

	arrResult.push(oState);
	return arrResult;
};
CEndnotesController.prototype.SetSelectionState = function(State, StateIndex)
{
	var oState = State[StateIndex];

	this.Selection.Use           = oState.Use;
	this.Selection.Start.Endnote = oState.Start;
	this.Selection.End.Endnote   = oState.End;
	this.Selection.Direction     = oState.Direction;
	this.CurEndnote              = oState.CurEndnote;
	this.Selection.Endnotes      = {};

	for (var sId in oState.Endnotes)
	{
		this.Selection.Endnotes[sId] = oState.Endnotes[sId].Endnote;
		this.Selection.Endnotes[sId].SetSelectionState(oState.Endnotes[sId].State, oState.Endnotes[sId].State.length - 1);
	}
};
CEndnotesController.prototype.AddHyperlink = function(oProps)
{
	if (true !== this.IsSelectionUse() || true === this.private_IsOneEndnoteSelected())
		this.CurEndnote.AddHyperlink(oProps);
};
CEndnotesController.prototype.ModifyHyperlink = function(oProps)
{
	if (true !== this.IsSelectionUse() || true === this.private_IsOneEndnoteSelected())
		this.CurEndnote.ModifyHyperlink(oProps);
};
CEndnotesController.prototype.RemoveHyperlink = function()
{
	if (true !== this.IsSelectionUse() || true === this.private_IsOneEndnoteSelected())
		this.CurEndnote.RemoveHyperlink();
};
CEndnotesController.prototype.CanAddHyperlink = function(bCheckInHyperlink)
{
	if (true !== this.IsSelectionUse() || true === this.private_IsOneEndnoteSelected())
		return this.CurEndnote.CanAddHyperlink(bCheckInHyperlink);

	return false;
};
CEndnotesController.prototype.IsCursorInHyperlink = function(bCheckEnd)
{
	if (true !== this.IsSelectionUse() || true === this.private_IsOneEndnoteSelected())
		return this.CurEndnote.IsCursorInHyperlink(bCheckEnd);

	return null;
};
CEndnotesController.prototype.AddComment = function(Comment)
{
	if (true !== this.IsSelectionUse() || true === this.private_IsOneEndnoteSelected())
	{
		this.CurEndnote.AddComment(Comment, true, true);
	}
};
CEndnotesController.prototype.CanAddComment = function()
{
	if (true !== this.IsSelectionUse() || true === this.private_IsOneEndnoteSelected())
		return this.CurEndnote.CanAddComment();

	return false;
};
CEndnotesController.prototype.GetSelectionAnchorPos = function()
{
	if (true !== this.Selection.Use || 0 === this.Selection.Direction)
		return this.CurEndnote.GetSelectionAnchorPos();
	else if (1 === this.Selection.Direction)
		return this.Selection.Start.Endnote.GetSelectionAnchorPos();
	else
		return this.Selection.End.Endnote.GetSelectionAnchorPos();
};
CEndnotesController.prototype.StartSelectionFromCurPos = function()
{
	if (true === this.Selection.Use)
		return;

	this.Selection.Use = true;
	this.Selection.Start.Endnote = this.CurEndnote;
	this.Selection.End.Endnote   = this.CurEndnote;
	this.Selection.Endnotes = {};
	this.Selection.Endnotes[this.CurEndnote.GetId()] = this.CurEndnote;
	this.CurEndnote.StartSelectionFromCurPos();
};
CEndnotesController.prototype.SaveDocumentStateBeforeLoadChanges   = function(State)
{
	State.EndnotesSelectDirection = this.Selection.Direction;
	State.EndnotesSelectionUse    = this.Selection.Use;

	if (0 === this.Selection.Direction || false === this.Selection.Use)
	{
		var oEndnote               = this.CurEndnote;
		State.CurEndnote           = oEndnote;
		State.CurEndnoteSelection  = oEndnote.Selection.Use;
		State.CurEndnoteDocPosType = oEndnote.GetDocPosType();

		if (docpostype_Content === oEndnote.GetDocPosType())
		{
			State.Pos      = oEndnote.GetContentPosition(false, false, undefined);
			State.StartPos = oEndnote.GetContentPosition(true, true, undefined);
			State.EndPos   = oEndnote.GetContentPosition(true, false, undefined);
		}
		else if (docpostype_DrawingObjects === oEndnote.GetDocPosType())
		{
			this.LogicDocument.DrawingObjects.Save_DocumentStateBeforeLoadChanges(State);
		}
	}
	else
	{
		State.EndnotesList  = this.private_GetSelectionArray();
		var oEndnote        = State.EndnotesList[0];
		State.EndnotesStart = {
			Pos      : oEndnote.GetContentPosition(false, false, undefined),
			StartPos : oEndnote.GetContentPosition(true, true, undefined),
			EndPos   : oEndnote.GetContentPosition(true, false, undefined)
		};

		oEndnote          = State.EndnotesList[State.EndnotesList.length - 1];
		State.EndnotesEnd = {
			Pos      : oEndnote.GetContentPosition(false, false, undefined),
			StartPos : oEndnote.GetContentPosition(true, true, undefined),
			EndPos   : oEndnote.GetContentPosition(true, false, undefined)
		};
	}
};
CEndnotesController.prototype.RestoreDocumentStateAfterLoadChanges = function(State)
{
	this.RemoveSelection();
	if (0 === State.EndnotesSelectDirection)
	{
		var oEndnote = State.CurEndnote;
		if (oEndnote && true === this.Is_UseInDocument(oEndnote.GetId()))
		{
			this.Selection.Start.Endnote = oEndnote;
			this.Selection.End.Endnote   = oEndnote;
			this.Selection.Direction     = 0;
			this.CurEndnote              = oEndnote;
			this.Selection.Endnotes      = {};
			this.Selection.Use           = State.EndnotesSelectionUse;

			this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;

			if (docpostype_Content === State.CurEndnoteDocPosType)
			{
				oEndnote.SetDocPosType(docpostype_Content);
				oEndnote.Selection.Use = State.CurEndnoteSelection;
				if (true === oEndnote.Selection.Use)
				{
					oEndnote.SetContentPosition(State.StartPos, 0, 0);
					oEndnote.SetContentSelection(State.StartPos, State.EndPos, 0, 0, 0);
				}
				else
				{
					oEndnote.SetContentPosition(State.Pos, 0, 0);
					this.LogicDocument.NeedUpdateTarget = true;
				}
			}
			else if (docpostype_DrawingObjects === State.CurEndnoteDocPosType)
			{
				oEndnote.SetDocPosType(docpostype_DrawingObjects);
				if (true !== this.LogicDocument.DrawingObjects.Load_DocumentStateAfterLoadChanges(State))
				{
					oEndnote.SetDocPosType(docpostype_Content);
					this.LogicDocument.MoveCursorToXY(State.X ? State.X : 0, State.Y ? State.Y : 0, false);
				}
			}
		}
		else
		{
			this.LogicDocument.EndEndnotesEditing();
		}
	}
	else
	{
		var arrEndnotesList = State.EndnotesList;

		var StartEndnote = null;
		var EndEndnote   = null;

		var arrAllEndnotes = this.private_GetEndnotesLogicRange(null, null);

		for (var nIndex = 0, nCount = arrEndnotesList.length; nIndex < nCount; ++nIndex)
		{
			var oEndnote = arrEndnotesList[nIndex];
			if (true === this.Is_UseInDocument(oEndnote.GetId(), arrAllEndnotes))
			{
				if (null === StartEndnote)
					StartEndnote = oEndnote;

				EndEndnote = oEndnote;
			}
		}

		if (null === StartEndnote || null === EndEndnote)
		{
			this.LogicDocument.EndEndnotesEditing();
			return;
		}

		var arrNewEndnotesList = this.private_GetEndnotesLogicRange(StartEndnote, EndEndnote);
		if (arrNewEndnotesList.length < 1)
		{
			if (null !== EndEndnote)
			{
				EndEndnote.RemoveSelection();
				this.private_SetCurrentEndnoteNoSelection(EndEndnote);
			}
			else if (null !== StartEndnote)
			{
				StartEndnote.RemoveSelection();
				this.private_SetCurrentEndnoteNoSelection(StartEndnote);
			}
			else
			{
				this.LogicDocument.EndEndnotesEditing();
			}
		}
		else if (arrNewEndnotesList.length === 1)
		{
			this.Selection.Use           = true;
			this.Selection.Direction     = 0;
			this.Selection.Endnotes      = {};
			this.Selection.Start.Endnote = StartEndnote;
			this.Selection.End.Endnote   = StartEndnote;
			this.CurEndnote              = StartEndnote;

			this.Selection.Endnotes[StartEndnote.GetId()] = StartEndnote;

			if (arrEndnotesList[0] === StartEndnote)
			{
				StartEndnote.SetDocPosType(docpostype_Content);
				StartEndnote.Selection.Use = true;
				StartEndnote.SetContentPosition(State.EndnotesStart.Pos, 0, 0);
				StartEndnote.SetContentSelection(State.EndnotesStart.StartPos, State.EndnotesStart.EndPos, 0, 0, 0);
			}
			else if (arrEndnotesList[arrAllEndnotes.length - 1] === StartEndnote)
			{
				StartEndnote.SetDocPosType(docpostype_Content);
				StartEndnote.Selection.Use = true;
				StartEndnote.SetContentPosition(State.EndnotesEnd.Pos, 0, 0);
				StartEndnote.SetContentSelection(State.EndnotesEnd.StartPos, State.EndnotesEnd.EndPos, 0, 0, 0);
			}
			else
			{
				StartEndnote.SetDocPosType(docpostype_Content);
				StartEndnote.SelectAll(1);
			}
		}
		else
		{
			this.Selection.Use       = true;
			this.Selection.Direction = State.EndnotesSelectDirection;
			this.Selection.Endnotes  = {};

			for (var nIndex = 1, nCount = arrNewEndnotesList.length; nIndex < nCount - 1; ++nIndex)
			{
				var oEndnote = arrNewEndnotesList[nIndex];
				oEndnote.SelectAll(this.Selection.Direction);
				this.Selection.Endnotes[oEndnote.GetId()] = oEndnote;
			}

			this.Selection.Endnotes[StartEndnote.GetId()] = StartEndnote;
			this.Selection.Endnotes[EndEndnote.GetId()]   = EndEndnote;


			if (arrEndnotesList[0] === StartEndnote)
			{
				StartEndnote.SetDocPosType(docpostype_Content);
				StartEndnote.Selection.Use = true;
				StartEndnote.SetContentPosition(State.EndnotesStart.Pos, 0, 0);
				StartEndnote.SetContentSelection(State.EndnotesStart.StartPos, State.EndnotesStart.EndPos, 0, 0, 0);
			}
			else
			{
				StartEndnote.SetDocPosType(docpostype_Content);
				StartEndnote.SelectAll(1);
			}

			if (arrEndnotesList[arrEndnotesList.length - 1] === EndEndnote)
			{
				EndEndnote.SetDocPosType(docpostype_Content);
				EndEndnote.Selection.Use = true;
				EndEndnote.SetContentPosition(State.EndnotesEnd.Pos, 0, 0);
				EndEndnote.SetContentSelection(State.EndnotesEnd.StartPos, State.EndnotesEnd.EndPos, 0, 0, 0);
			}
			else
			{
				EndEndnote.SetDocPosType(docpostype_Content);
				EndEndnote.SelectAll(1);
			}

			if (1 !== this.Selection.Direction)
			{
				var Temp     = StartEndnote;
				StartEndnote = EndEndnote;
				EndEndnote   = Temp;
			}

			this.Selection.Start.Endnote = StartEndnote;
			this.Selection.End.Endnote   = EndEndnote;
		}
	}
};
CEndnotesController.prototype.GetColumnSize = function()
{
	// TODO: Переделать
	var _w = Math.max(1, AscCommon.Page_Width - (AscCommon.X_Left_Margin + AscCommon.X_Right_Margin));
	var _h = Math.max(1, AscCommon.Page_Height - (AscCommon.Y_Top_Margin + AscCommon.Y_Bottom_Margin));

	return {
		W : AscCommon.Page_Width - (AscCommon.X_Left_Margin + AscCommon.X_Right_Margin),
		H : AscCommon.Page_Height - (AscCommon.Y_Top_Margin + AscCommon.Y_Bottom_Margin)
	};
};
CEndnotesController.prototype.GetCurrentSectionPr = function()
{
	return null;
};
CEndnotesController.prototype.GetColumnFields = function(nPageAbs, nColumnAbs, nSectionIndex)
{
	var oColumn = this.private_GetPageColumn(nPageAbs, nColumnAbs, nSectionIndex);
	if (!oColumn)
		return {X : 0, XLimit : 297};

	return {X : oColumn.X, XLimit : oColumn.XLimit};
};
CEndnotesController.prototype.RemoveTextSelection = function()
{
	if (true === this.Selection.Use)
	{
		for (var sId in this.Selection.Endnotes)
		{
			if (this.Selection.Endnotes[sId] !== this.CurEndnote)
				this.Selection.Endnotes[sId].RemoveSelection();
		}

		this.Selection.Use = false;
	}

	this.Selection.Endnotes = {};
	if (this.CurEndnote)
	{
		this.Selection.Endnotes[this.CurEndnote.GetId()] = this.CurEndnote;
		this.CurEndnote.RemoveTextSelection();
	}
};
CEndnotesController.prototype.ResetRecalculateCache = function()
{
	for (var Id in this.Endnote)
	{
		this.Endnote[Id].Reset_RecalculateCache();
	}
};
CEndnotesController.prototype.AddContentControl = function(nContentControlType)
{
	if (this.CurEndnote)
		return this.CurEndnote.AddContentControl(nContentControlType);

	return null;
};
CEndnotesController.prototype.GetStyleFromFormatting = function()
{
	if (this.CurEndnote)
		return this.CurEndnote.GetStyleFromFormatting();

	return null;
};
CEndnotesController.prototype.GetSimilarNumbering = function(oEngine)
{
	if (this.CurEndnote)
		this.CurEndnote.GetSimilarNumbering(oEngine);
};
CEndnotesController.prototype.GetPlaceHolderObject = function()
{
	if (this.CurEndnote)
		return this.CurEndnote.GetPlaceHolderObject();

	return null;
};
CEndnotesController.prototype.GetAllFields = function(isUseSelection, arrFields)
{
	// Поиск по всем сноскам должен происходить не здесь
	if (!isUseSelection || !this.CurEndnote)
		return arrFields ? arrFields : [];

	return this.CurEndnote.GetAllFields(isUseSelection, arrFields);
};
CEndnotesController.prototype.GetAllDrawingObjects = function(arrDrawings)
{
	if (undefined === arrDrawings || null === arrDrawings)
		arrDrawings = [];

	for (var sId in  this.Endnote)
	{
		var oEndnote = this.Endnote[sId];
		oEndnote.GetAllDrawingObjects(arrDrawings);
	}

	return arrDrawings;
};
CEndnotesController.prototype.UpdateBookmarks = function(oBookmarkManager)
{
	for (var sId in  this.Endnote)
	{
		var oEndnote = this.Endnote[sId];
		oEndnote.UpdateBookmarks(oBookmarkManager);
	}
};
CEndnotesController.prototype.IsTableCellSelection = function()
{
	if (this.CurEndnote)
		return this.CurEndnote.IsTableCellSelection();

	return false;
};
CEndnotesController.prototype.AcceptRevisionChanges = function(Type, bAll)
{
	if (null !== this.CurEndnote)
		this.CurEndnote.AcceptRevisionChanges(Type, bAll);
};
CEndnotesController.prototype.RejectRevisionChanges = function(Type, bAll)
{
	if (null !== this.CurEndnote)
		this.CurEndnote.RejectRevisionChanges(Type, bAll);
};
CEndnotesController.prototype.IsSelectionLocked = function(CheckType)
{
	for (var sId in this.Selection.Endnotes)
	{
		var oEndnote = this.Selection.Endnotes[sId];
		oEndnote.Document_Is_SelectionLocked(CheckType);
	}
};
CEndnotesController.prototype.GetAllTablesOnPage = function(nPageAbs, arrTables)
{
	var oPage = this.Pages[nPageAbs];
	if (!oPage || oPage.Sections.length <= 0)
		return arrTables;

	for (var nSectionIndex = 0, nSectionsCount = oPage.Sections.length; nSectionIndex < nSectionsCount; ++nSectionIndex)
	{
		var oSection = this.Sections[oPage.Sections[nSectionIndex]];
		if (!oSection)
			continue;

		var oSectionPage = oSection.Pages[nPageAbs];
		if (!oSectionPage)
			continue;

		var nColumnsCount = oSectionPage.Columns.length;
		for (var nColumnIndex = 0; nColumnIndex < nColumnsCount; ++nColumnIndex)
		{
			var oColumn = oSectionPage.Columns[nColumnIndex];
			if (!oColumn || oColumn.Elements.length <= 0)
				continue;

			for (var nIndex = 0, nCount = oColumn.Elements.length; nIndex < nCount; ++nIndex)
			{
				oColumn.Elements[nIndex].GetAllTablesOnPage(nPageAbs, arrTables);
			}
		}
	}

	return arrTables;
};
CEndnotesController.prototype.FindNextFillingForm = function(isNext, isCurrent)
{
	var oCurEndnote = this.CurEndnote;

	var arrEndnotes = this.LogicDocument.GetEndnotesList(null, null);
	var nCurPos     = -1;
	var nCount      = arrEndnotes.length;

	if (nCount <= 0)
		return null;

	if (isCurrent)
	{
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			if (arrEndnotes[nIndex] === oCurEndnote)
			{
				nCurPos = nIndex;
				break;
			}
		}
	}

	if (-1 === nCurPos)
	{
		nCurPos      = isNext ? 0 : nCount - 1;
		oCurEndnote = arrEndnotes[nCurPos];
		isCurrent    = false;
	}

	var oRes = oCurEndnote.FindNextFillingForm(isNext, isCurrent, isCurrent);
	if (oRes)
		return oRes;

	if (true === isNext)
	{
		for (var nIndex = nCurPos + 1; nIndex < nCount; ++nIndex)
		{
			oRes = arrEndnotes[nIndex].FindNextFillingForm(isNext, false);
			if (oRes)
				return oRes;
		}
	}
	else
	{
		for (var nIndex = nCurPos - 1; nIndex >= 0; --nIndex)
		{
			oRes = arrEndnotes[nIndex].FindNextFillingForm(isNext, false);
			if (oRes)
				return oRes;
		}
	}

	return null;
};
CEndnotesController.prototype.CollectSelectedReviewChanges = function(oTrackManager)
{
	if (this.Selection.Use)
	{
		for (var sId in this.Selection.Endnotes)
		{
			this.Selection.Endnotes[sId].CollectSelectedReviewChanges(oTrackManager);
		}
	}
	else if (this.CurEndnote)
	{
		this.CurEndnote.CollectSelectedReviewChanges(oTrackManager);
	}
};

/**
 * Класс регистрирующий концевые сноски на странице
 * и номера секций, сноски которых были пересчитаны на данной странице
 * @constructor
 */
function CEndnotePage()
{
	this.Endnotes    = [];
	this.Sections    = [];
	this.Continue    = false; // Сноски на данной странице не закончились и переносятся на следующую
	this.EndnotesPos = [];    // Логическая позиция на странице (номер секции на странице и номер колонки)
	this.CurSection  = 0;
	this.CurColumn   = 0;
}
CEndnotePage.prototype.Reset = function()
{
	this.Endnotes    = [];
	this.Sections    = [];
	this.EndnotesPos = [];
	this.CurSection  = 0;
	this.CurColumn   = 0;
};
CEndnotePage.prototype.ResetColumn = function(nSectionIndex, nColumnIndex)
{
	this.CurSection = nSectionIndex;
	this.CurColumn  = nColumnIndex;

	for (var nEndnoteIndex = 0, nEndnotesCount = this.Endnotes.length; nEndnoteIndex < nEndnotesCount; ++nEndnoteIndex)
	{
		if ((this.EndnotesPos[nEndnoteIndex].Section === nSectionIndex && this.EndnotesPos[nEndnoteIndex].Column >= nColumnIndex)
			|| (this.EndnotesPos[nEndnoteIndex].Section > nSectionIndex))
		{
			this.Endnotes.length = nEndnoteIndex;
			this.EndnotesPos.length = nEndnoteIndex;
			return;
		}
	}
};
CEndnotePage.prototype.AddEndnotes = function(arrEndnotes)
{
	// Может прийти добавление одной и той же сноски несколько раз, т.к. мы можем пересчитывать одну и ту же колонку
	// или страницу несколько раз. Но при этом сама последовательность сносок не должна меняться, поэтому
	// точку поиска следующей сноски спокойно сдвигаем, если нашли для предыдущей.

	var nStartPos = 0;
	for (var nAddIndex = 0, nAddCount = arrEndnotes.length; nAddIndex < nAddCount; ++nAddIndex)
	{
		var oEndnote  = arrEndnotes[nAddIndex];
		var isNeedAdd = true;
		for (var nEndnoteIndex = nStartPos, nEndnotesCount = this.Endnotes.length; nEndnoteIndex < nEndnotesCount; ++nEndnoteIndex)
		{
			if (this.Endnotes[nEndnoteIndex] === oEndnote)
			{
				nStartPos = nEndnoteIndex + 1;
				isNeedAdd = false;
				this.EndnotesPos[nEndnoteIndex] = {Section : this.CurSection, Column : this.CurColumn};
			}
		}

		if (isNeedAdd)
		{
			this.Endnotes.push(oEndnote);
			this.EndnotesPos.push({Section : this.CurSection, Column : this.CurColumn});
		}
	}
};
CEndnotePage.prototype.AddSection = function(nSectionIndex)
{
	this.Sections.push(nSectionIndex);
};
CEndnotePage.prototype.SetContinue = function(isContinue)
{
	this.Continue = isContinue;
};

function CEndnoteSection()
{
	this.Endnotes    = [];
	this.StartPage   = 0;
	this.StartColumn = 0;

	this.Pages = [];

	this.SeparatorRecalculateObject = null;
	this.Separator                  = false; // true - Separator
}

function CEndnoteSectionPage()
{
	this.Columns = [];
}

function CEndnoteSectionPageColumn()
{
	this.Elements = [];
	this.StartPos = 0;
	this.EndPos   = -1;

	this.X      = 0;
	this.XLimit = 0;
	this.Y      = 0;
	this.YLimit = 0;
	this.Height = 0;

	this.ContinuationSeparatorRecalculateObject = null;
}
