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

// Функция копирует объект или массив. (Обычное равенство в javascript приравнивает указатели)
function Common_CopyObj(Obj)
{
    if( !Obj || !('object' == typeof(Obj) || 'array' == typeof(Obj)) )
    {
        return Obj;
    }

    var c = 'function' === typeof Obj.pop ? [] : {};
    var p, v;
    for(p in Obj)
    {
        if(Obj.hasOwnProperty(p))
        {
            v = Obj[p];
            if(v && 'object' === typeof v )
            {
                c[p] = Common_CopyObj(v);
            }
            else
            {
                c[p] = v;
            }
        }
    }
    return c;
}

/**
 * Класс для обркботки конвертации текста в таблицу
 * @constructor
 */
function CTextToTableEngine()
{
	this.SeparatorType = Asc.c_oAscTextToTableSeparator.Paragraph;
	this.Separator     = 0;
	this.MaxCols       = 0;

	this.Mode = 0; // Режим работы
	               // 0 - вычисляем размер
				   // 1 - проверяем типы разделителей
	               // 2 - набиваем элементы

	this.Cols    = 0;
	this.Rows    = 0;
	this.CurCols = 0;

	this.Tab           = true;
	this.Semicolon     = true;
	this.ParaTab       = false;
	this.ParaSemicolon = false;

	this.ParaPositions = [];
	this.Rows          = [];
	this.ItemsBuffer   = [];
	this.CurCol        = 0;
	this.CurRow        = 0;
	this.CC            = null;
}
CTextToTableEngine.prototype.Reset = function()
{
	this.Cols    = 0;
	this.Rows    = 0;
	this.CurCols = 0;

	this.Tab           = true;
	this.Semicolon     = true;
	this.ParaTab       = false;
	this.ParaSemicolon = false;
};
CTextToTableEngine.prototype.GetSeparatorType = function()
{
	return this.Type;
};
CTextToTableEngine.prototype.GetSeparator = function()
{
	return this.Separator;
};
CTextToTableEngine.prototype.AddItem = function()
{
	if (this.IsParagraphSeparator())
		return;

	if (0 === this.CurCols)
		this.Rows++;

	if (this.MaxCols)
	{
		if (this.CurCols < this.MaxCols)
		{
			this.CurCols++;
		}
		else
		{
			if (this.Cols < this.CurCols)
				this.Cols = this.CurCols;

			this.Rows++;
			this.CurCols = 1;
		}
	}
	else
	{
		this.CurCols++;
	}
};
CTextToTableEngine.prototype.OnStartParagraph = function()
{
	if (this.IsCalculateTableSizeMode())
	{
		this.AddItem();
	}
	else if (this.IsCheckSeparatorMode())
	{
		this.ParaTab       = false;
		this.ParaSemicolon = false;
	}
	else if (this.IsConvertMode())
	{
		this.ParaPositions = [];
	}
};
CTextToTableEngine.prototype.OnEndParagraph = function(oParagraph)
{
	if (this.IsCalculateTableSizeMode())
	{
		if (this.IsParagraphSeparator())
		{
			if (this.MaxCols)
			{
				if (0 === this.CurCols)
					this.Rows++;

				if (this.CurCols < this.MaxCols)
				{
					this.CurCols++;
				}
				else
				{
					if (this.Cols < this.CurCols)
						this.Cols = this.CurCols;

					this.Rows++;
					this.CurCols = 1;
				}
			}
			else
			{
				this.Rows++;
			}
		}
		else
		{
			if (this.CurCols)
			{
				if (this.Cols < this.CurCols)
					this.Cols = this.CurCols;

				this.CurCols = 0;
			}
		}
	}
	else if (this.IsCheckSeparatorMode())
	{
		this.Tab       = this.Tab && this.ParaTab;
		this.Semicolon = this.Semicolon && this.ParaSemicolon;
	}
	else if (this.IsConvertMode())
	{
		// Если у нас данный параграф не делится на несколько частей и находится в контроле и он единственный
		// элемент в этом контроле, то его надо оставить в том контроле
		// За исключением самого первого контрола, который лежит в верху, его мы используем для обертки
		var oElement = oParagraph;
		var oParent  = oParagraph.GetParent();
		if (this.ItemsBuffer.length <= 0 && oParent
			&& oParent.IsBlockLevelSdtContent()
			&& 1 === oParent.GetElementsCount()
			&& (this.IsParagraphSeparator() || !this.ParaPositions.length)
			&& oParent.Parent !== this.CC)
		{
			oElement = oParent.Parent;
		}

		if (this.IsParagraphSeparator())
		{
			this.CheckBuffer();
			this.Rows[this.CurRow][this.CurCol].push(oElement.Copy());

			this.CurCol++;

			if (this.CurCol >= this.MaxCols)
			{
				this.CurCol = 0;
				this.CurRow++;
			}
		}
		else
		{
			var arrParagraphs = [];
			for (var nIndex = this.ParaPositions.length - 1; nIndex >= 0; --nIndex)
			{
				var oTempParagraph = oParagraph.SplitNoDuplicate(this.ParaPositions[nIndex]);

				var oRunElements = new CParagraphRunElements(oTempParagraph.GetStartPos(), 1, null);
				oRunElements.SetSaveContentPositions(true);
				oRunElements.SetSkipMath(false);
				oTempParagraph.GetNextRunElements(oRunElements);

				if (1 === oRunElements.Elements.length
					&& this.CheckSeparator(oRunElements.Elements[0]))
				{
					var oTempRunPos = oRunElements.GetContentPositions()[0];
					var nInRunPos = oTempRunPos.Get(oTempRunPos.GetDepth());
					oTempRunPos.DecreaseDepth(1);
					var oTempRun  = oTempParagraph.GetClassByPos(oTempRunPos);
					if (oTempRun)
						oTempRun.RemoveFromContent(nInRunPos, 1);
				}

				arrParagraphs.push(oTempParagraph);
			}

			this.CurCol = 0;

			this.CheckBuffer();
			this.Rows[this.CurRow][this.CurCol].push(oElement.Copy());
			this.CurCol++;

			if (this.CurCol >= this.MaxCols)
			{
				this.CurCol = 0;
				if (arrParagraphs.length > 0)
					this.CurRow++;
			}

			for (var nIndex = arrParagraphs.length - 1; nIndex >= 0; --nIndex)
			{
				if (0 === this.CurCol)
					this.Rows[this.CurRow] = [];

				if (!this.Rows[this.CurRow][this.CurCol])
					this.Rows[this.CurRow][this.CurCol] = [];

				this.Rows[this.CurRow][this.CurCol].push(arrParagraphs[nIndex]);
				this.CurCol++;

				if (this.CurCol >= this.MaxCols)
				{
					this.CurCol = 0;
					if (nIndex > 0)
						this.CurRow++;
				}
			}

			this.CurRow++;
		}
	}
};
CTextToTableEngine.prototype.OnTable = function(oTable)
{
	if (this.IsConvertMode())
	{
		this.ItemsBuffer.push(oTable);
	}
};
CTextToTableEngine.prototype.FinalizeConvert = function()
{
	if (this.IsConvertMode() && this.ItemsBuffer.length > 0)
	{
		// Случай, когда последним элементом идет таблица
		this.CheckBuffer();
	}
};
CTextToTableEngine.prototype.CheckBuffer = function()
{
	if (0 === this.CurCol)
		this.Rows[this.CurRow] = [];

	this.Rows[this.CurRow][this.CurCol] = [];
	if (this.ItemsBuffer.length > 0)
	{
		for (var nIndex = 0, nCount = this.ItemsBuffer.length; nIndex < nCount; ++nIndex)
		{
			this.Rows[this.CurRow][this.CurCol].push(this.ItemsBuffer[nIndex].Copy());
		}

		this.ItemsBuffer = [];
	}
};
CTextToTableEngine.prototype.IsParagraphSeparator = function()
{
	return this.SeparatorType === Asc.c_oAscTextToTableSeparator.Paragraph;
};
CTextToTableEngine.prototype.IsSymbolSeparator = function(nCharCode)
{
	return (this.SeparatorType === Asc.c_oAscTextToTableSeparator.Symbol && this.Separator === nCharCode);
};
CTextToTableEngine.prototype.IsTabSeparator = function()
{
	return this.SeparatorType === Asc.c_oAscTextToTableSeparator.Tab;
};
CTextToTableEngine.prototype.CheckSeparator = function(oRunItem)
{
	var nItemType = oRunItem.Type;

	return ((para_Tab === nItemType && this.IsTabSeparator())
		|| (para_Text === nItemType && this.IsSymbolSeparator(oRunItem.Value))
		|| (para_Space === nItemType && this.IsSymbolSeparator(oRunItem.Value))
		|| (para_Math_Text === nItemType && this.IsSymbolSeparator(oRunItem.value)));
};
CTextToTableEngine.prototype.SetCalculateTableSizeMode = function(nSeparatorType, nSeparator, nMaxCols)
{
	this.Mode = 0;

	this.SeparatorType = undefined !== nSeparatorType ? nSeparatorType : Asc.c_oAscTextToTableSeparator.Paragraph;
	this.Separator     = undefined !== nSeparator ? nSeparator : 0;
	this.MaxCols       = undefined !== nMaxCols ? nMaxCols : 0;
};
CTextToTableEngine.prototype.SetCheckSeparatorMode = function()
{
	this.Mode = 1;
};
CTextToTableEngine.prototype.SetConvertMode = function(nSeparatorType, nSeparator, nMaxCols)
{
	this.Mode = 2;

	this.SeparatorType = undefined !== nSeparatorType ? nSeparatorType : Asc.c_oAscTextToTableSeparator.Paragraph;
	this.Separator     = undefined !== nSeparator ? nSeparator : 0;
	this.MaxCols       = undefined !== nMaxCols ? nMaxCols : 1;
};
CTextToTableEngine.prototype.IsCalculateTableSizeMode = function()
{
	return (0 === this.Mode);
};
CTextToTableEngine.prototype.IsCheckSeparatorMode = function()
{
	return (1 === this.Mode);
};
CTextToTableEngine.prototype.IsConvertMode = function()
{
	return (2 === this.Mode);
};
CTextToTableEngine.prototype.AddTab = function()
{
	this.ParaTab = true;
};
CTextToTableEngine.prototype.AddSemicolon = function()
{
	this.ParaSemicolon = true;
};
CTextToTableEngine.prototype.HaveTab = function()
{
	return this.Tab;
};
CTextToTableEngine.prototype.HaveSemicolon = function()
{
	return this.Semicolon;
};
CTextToTableEngine.prototype.AddParaPosition = function(oParaContentPos)
{
	this.ParaPositions.push(oParaContentPos);
};
CTextToTableEngine.prototype.GetRows = function()
{
	return this.Rows;
};
CTextToTableEngine.prototype.SetContentControl = function(oCC)
{
	this.CC = oCC;
};
CTextToTableEngine.prototype.GetContentControl = function()
{
	return this.CC;
};

//--------------------------------------------------------export--------------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CTextToTableEngine = CTextToTableEngine;

(function(window, undefined)
{
	function private_CheckDrawingDocumentPosition(oDocPos)
	{
		var oShape = null;
		if (oDocPos[0].Class instanceof AscCommonWord.CDocumentContent && (oShape = oDocPos[0].Class.Is_DrawingShape(true)) && oShape.parent instanceof AscCommonWord.ParaDrawing)
		{
			var oDrawingDocPos = oShape.parent.GetDocumentPositionFromObject();
			if (!oDrawingDocPos || !oDrawingDocPos.length)
				return oDocPos;

			return oDrawingDocPos.concat(oDocPos);
		}

		return oDocPos;
	}
	function private_GetTopDocumentPosition(oDocPos)
	{
		var oClass = oDocPos[0].Class;
		if (oClass instanceof AscCommonWord.CDocument)
			return 0;
		else if (!(oClass instanceof AscCommonWord.CDocumentContent))
			return 0xFF;
		if (oClass.IsFootnote() && oClass.GetParent() instanceof CFootnotesController)
			return 1;
		else if (oClass.IsFootnote() && oClass.GetParent() instanceof CEndnotesController)
			return 2;
		else if (oClass.IsHdrFtr())
			return 3;

		return 0xFF;
	}
	function private_CompareDocumentPositions(oDocPos1, oDocPos2)
	{
		for (var nIndex = 0, nCount = oDocPos1.length; nIndex < nCount; ++nIndex)
		{
			if (oDocPos2.length <= nIndex)
				return 1;

			if (oDocPos1[nIndex].Position < oDocPos2[nIndex].Position)
				return -1;
			else if (oDocPos1[nIndex].Position > oDocPos2[nIndex].Position)
				return 1;
		}

		return 0;
	}
	function private_GetNoteReferencePosition(oDocPos)
	{
		var oClass = oDocPos[0].Class;
		if (!(oClass instanceof CFootEndnote))
			return oDocPos;

		var oRef = oClass.GetRef();
		if (!oRef)
			return oDocPos;

		var oRun = oRef.GetRun();
		if (!oRun)
			return oDocPos;

		var oRunDocPos = oRun.GetDocumentPositionFromObject();
		var nInRunPos = oRun.GetElementPosition(oRef);
		if (!oRunDocPos || -1 === nInRunPos)
			return oDocPos;

		oRunDocPos.push({Class : oRun, Position : nInRunPos});
		return oRunDocPos;
	}
	function private_GetSectionHeaderIndex(oHeader)
	{
		var oSection = oHeader.GetSectionPr();
		if (!oSection)
			return -1;

		if (oHeader === oSection.Get_Header_First())
			return 0;
		else if (oHeader === oSection.Get_Footer_First())
			return 1;
		else if (oHeader === oSection.Get_Header_Default())
			return 2;
		else if (oHeader === oSection.Get_Footer_Default())
			return 3;
		else if (oHeader === oSection.Get_Header_Even())
			return 4;
		else if (oHeader === oSection.Get_Footer_Even())
			return 5;

		return -1;
	}
	function private_CompareHdrFtrPosition(oDocPos1, oDocPos2)
	{
		var oHeader1 = oDocPos1[0].Class.IsHdrFtr(true);
		var oHeader2 = oDocPos2[0].Class.IsHdrFtr(true);

		if (!oHeader1 || !oHeader2 || oHeader1 === oHeader2)
			return 0;

		var nSectionIndex1 = oHeader1.GetSectionIndex();
		var nSectionIndex2 = oHeader2.GetSectionIndex();

		if (nSectionIndex1 < nSectionIndex2)
			return -1;
		else if (nSectionIndex1 > nSectionIndex2)
			return 1;

		return (private_GetSectionHeaderIndex(oHeader1) < private_GetSectionHeaderIndex(oHeader2) ? -1 : 1);
	}
	function CompareDocumentPositions(oDocPos1, oDocPos2)
	{
		if (!oDocPos1 || !oDocPos2 || !oDocPos1.length || !oDocPos2.length)
			return 0;

		oDocPos1 = private_CheckDrawingDocumentPosition(oDocPos1);
		oDocPos2 = private_CheckDrawingDocumentPosition(oDocPos2);

		var nTopPos1 = private_GetTopDocumentPosition(oDocPos1);
		var nTopPos2 = private_GetTopDocumentPosition(oDocPos2);

		if (nTopPos1 !== nTopPos2)
			return (nTopPos1 < nTopPos2 ? -1 : 1);

		if (oDocPos1[0].Class !== oDocPos2[0].Class)
		{
			if (1 === nTopPos1 || 2 === nTopPos2)
			{
				oDocPos1 = private_GetNoteReferencePosition(oDocPos1);
				oDocPos2 = private_GetNoteReferencePosition(oDocPos2);
			}
			else if (3 === nTopPos1)
			{
				return private_CompareHdrFtrPosition(oDocPos1, oDocPos2);
			}
		}

		if (oDocPos1[0].Class !== oDocPos2[0].Class)
			return 0;

		return private_CompareDocumentPositions(oDocPos1, oDocPos2);
	}

	//--------------------------------------------------------export----------------------------------------------------
	window['AscCommonWord'] = window['AscCommonWord'] || {};
	window['AscCommonWord'].CompareDocumentPositions = CompareDocumentPositions;

})(window);
