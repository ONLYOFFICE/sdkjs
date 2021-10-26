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
(function(window, builder)
{
	/**
	 * Base class
	 * @global
	 * @class
	 * @name Api
	 */
	var Api = window["Asc"]["asc_docs_api"] || window["Asc"]["spreadsheet_api"];
	var c_oAscRevisionsChangeType = Asc.c_oAscRevisionsChangeType;
	var c_oAscSectionBreakType    = Asc.c_oAscSectionBreakType;
	var c_oAscSdtLockType         = Asc.c_oAscSdtLockType;
	var c_oAscAlignH         = Asc.c_oAscAlignH;
	var c_oAscAlignV         = Asc.c_oAscAlignV;

	var arrApiRanges		 = [];
	function private_RemoveEmptyRanges()
	{
		function ckeck_equal(firstDocPos, secondDocPos)
		{
			if (firstDocPos.length === secondDocPos.length)
			{
				for (var nPos = 0; nPos < firstDocPos.length; nPos++)
				{
					if (firstDocPos[nPos].Class === secondDocPos[nPos].Class && firstDocPos[nPos].Position === secondDocPos[nPos].Position)
						continue;
					else 
						return false;
				}
				return true;
			}
			return false;
		};

		var Range = null;
		for (var nRange = 0; nRange < arrApiRanges.length; nRange++)
		{
			Range = arrApiRanges[nRange];
			if (ckeck_equal(Range.StartPos, Range.EndPos))
			{
				Range.isEmpty = true;
				arrApiRanges.splice(nRange, 1);
				nRange--;
			}
		}
	};
	function private_TrackRangesPositions(bClearTrackedPosition)
	{
		var Document  = private_GetLogicDocument();
		var Range     = null;

		if (bClearTrackedPosition)
			Document.CollaborativeEditing.Clear_DocumentPositions();

		for (var nRange = 0; nRange < arrApiRanges.length; nRange++)
		{
			Range = arrApiRanges[nRange];
			Document.CollaborativeEditing.Add_DocumentPosition(Range.StartPos);
			Document.CollaborativeEditing.Add_DocumentPosition(Range.EndPos);
		}
	};
	function private_RefreshRangesPosition()
	{
		var Document  = private_GetLogicDocument();
		var Range     = null;

		for (var nRange = 0; nRange < arrApiRanges.length; nRange++)
		{
			Range = arrApiRanges[nRange];
			Document.RefreshDocumentPositions([Range.StartPos, Range.EndPos]);
		}
	};
	/**
	 * Get the first Run in the array specified.
	 * @typeofeditors ["CDE"]
	 * @param {Array} arrRuns - Array of Runs.
	 * @return {ApiRun | null} - returns null if arrRuns is invalid.
	 */
	function private_GetFirstRunInArray(arrRuns)
	{
		if (!Array.isArray(arrRuns))
			return null;
			
		var min_pos_Index = 0; // Индекс рана в массиве, с которого начнется выделение

		var MinPos = arrRuns[0].Run.GetDocumentPositionFromObject();

		for (var Index = 1; Index < arrRuns.length; Index++)
		{
			var TempPos = arrRuns[Index].Run.GetDocumentPositionFromObject();

			var MinPosLength = MinPos.length;
			var UsedLength1  = 0;


			if (MinPosLength <= TempPos.length)
				UsedLength1 = MinPosLength;
			else 
				UsedLength1 = TempPos.length;

			for (var Pos = 0; Pos < UsedLength1; Pos++)
			{
				if (TempPos[Pos].Position < MinPos[Pos].Position)
				{
					MinPos = TempPos;
					min_pos_Index = Index;
					break;
				}
				else if (TempPos[Pos].Position === MinPos[Pos].Position)
					continue;
				else if (TempPos[Pos].Position > MinPos[Pos].Position)
					break;
			}
		}
		
		return arrRuns[min_pos_Index];
	};
	/**
	 * Get the first Run in the array specified.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {Array} arrRuns - Array of Runs.
	 * @return {ApiRun | null} - returns null if arrRuns is invalid. 
	 */
	function private_GetLastRunInArray(arrRuns)
	{
		if (!Array.isArray(arrRuns))
			return false;
			
		var max_pos_Index = 0; // Индекс рана в массиве, на котором закончится

		var MaxPos = arrRuns[0].Run.GetDocumentPositionFromObject();

		for (var Index = 1; Index < arrRuns.length; Index++)
		{
			var TempPos = arrRuns[Index].Run.GetDocumentPositionFromObject();

			var MaxPosLength = MaxPos.length;
			var UsedLength2  = 0;

			if (MaxPosLength <= TempPos.length)
				UsedLength2 = MaxPosLength;
			else 
				UsedLength2 = TempPos.length;
			
			for (var Pos = 0; Pos < UsedLength2; Pos++)
			{
				if (TempPos[Pos].Position > MaxPos[Pos].Position)
				{
					MaxPos = TempPos;
					max_pos_Index = Index;
					break;
				}
				else if (TempPos[Pos].Position === MaxPos[Pos].Position)
					continue;
				else if (TempPos[Pos].Position < MaxPos[Pos].Position)
					break;
			}
		}
		return arrRuns[max_pos_Index];
	};
	function private_isMonospaceFont(sFontName)
	{
		if (   sFontName === 'Courier New'
		    || sFontName === 'Consolas'
		    || sFontName === 'Inconsolata'
		    || sFontName === 'Roboto Mono'
		    || sFontName === 'Source Code Pro')
			return true;

		return false;
	};

	/**
	 * Class representing a container for paragraphs and tables.
	 * @param Document
	 * @constructor
	 */
	function ApiDocumentContent(Document)
	{
		this.Document = Document;
	}
	/**
	 * Class represents a Markdown conversion processing.
	 * Each Range object is determined by the position of the start and end characters
	 * @constructor
	 */
	function CMarkdownConverter(oConfig)
	{
		this.HtmlTags =
		{
			Bold: '<strong>',
			Italic: '<em>',
			CodeLine: '<pre>',
			Code: '<code>',
			Paragraph: '<p>',
			Headings: {
				h1: '<h1>',
				h2: '<h2>',
				h3: '<h3>',
				h4: '<h4>',
				h5: '<h5>',
				h6: '<h6>'
			},
			Numberring: {
				Bulleted: '<ul>',
				Numbered: '<ol>',
				Item: '<li>'
			},
			Quote: '<blockquote>'
		};
		this.MdSymbols =
		{
			Bold: '**',
			Italic: '*',
			CodeLine: '```',
			Code:  '`',
			Headings: {
				h1: '#',
				h2: '##',
				h3: '###',
				h4: '####',
				h5: '#####',
				h6: '######'
			},
			Quote: '>'
		}

		this.Config             = oConfig;
		this.isNumbering        = false;
		this.isHeading          = false;
		this.isCodeBlock        = false;
		this.isTableCellContent = false;
		this.isQuoteLine        = false;
		this.currNumberingLvl   = -1;
		this.openedListsHtml    = [];
	};
	CMarkdownConverter.prototype.constructor = CMarkdownConverter;

	CMarkdownConverter.prototype.WrapInSymbol = function(sText, sSyblols, sWrapType)
	{
		switch (sWrapType)
		{
			case 'open':
				return sSyblols + sText
			case 'close':
				return sText + sSyblols;
			case 'wholly':
			default:
				return sSyblols + sText + sSyblols;
		}
	};
	CMarkdownConverter.prototype.WrapInTag = function(sText, sHtmlTag, sWrapType)
	{
		switch (sWrapType)
		{
			case 'open':
				return sHtmlTag + sText;
			case 'close':
				return sText + sHtmlTag.replace('<', '</');
			case 'wholly':
			default:
				return sHtmlTag + sText + sHtmlTag.replace('<', '</');
		}
	};
	CMarkdownConverter.prototype.DoMarkdown = function()
	{
		var oApi               = editor;
		var oDocument          = oApi.GetDocument();
		var sOutputText        = '';
		var arrSelectedContent = [];
		var oSelectedContent   = null;
		var oTempElm           = null;

		if (oDocument.Document.IsSelectionUse())
		{
			oSelectedContent = oDocument.Document.GetSelectedContent();
			for (var nElm = 0; nElm < oSelectedContent.Elements.length; nElm ++)
			{
				oTempElm = oSelectedContent.Elements[nElm].Element;
				if (oTempElm instanceof CTable)
					arrSelectedContent.push(new ApiTable(oTempElm));
				else if (oTempElm instanceof Paragraph)
					arrSelectedContent.push(new ApiParagraph(oTempElm));
				else if (oTempElm instanceof ParaRun)
					arrSelectedContent.push(new ApiRun(oTempElm));
				else if (oTempElm instanceof ParaRun)
					arrSelectedContent.push(new ApiRun(oTempElm));
				else
					continue;
			}

			for (var nElm = 0; nElm < arrSelectedContent.length; nElm ++)
			{
				sOutputText += this.HandleChildElement(arrSelectedContent[nElm], 'markdown');
			}
		}
		else
		{
			for (var nElm = 0, nElmsCount = oDocument.GetElementsCount(); nElm < nElmsCount; nElm ++)
			{
				sOutputText += this.HandleChildElement(oDocument.GetElement(nElm), 'markdown');
			}
		}

		return sOutputText;
	};
	CMarkdownConverter.prototype.DoHtml = function()
	{
		var oApi               = editor;
		var oDocument          = oApi.GetDocument();
		var sOutputText        = '';
		var arrSelectedContent = [];
		var oSelectedContent   = null;
		var oTempElm           = null;

		if (oDocument.Document.IsSelectionUse())
		{
			oSelectedContent = oDocument.Document.GetSelectedContent();
			for (var nElm = 0; nElm < oSelectedContent.Elements.length; nElm ++)
			{
				oTempElm = oSelectedContent.Elements[nElm].Element;
				if (oTempElm instanceof CTable)
					arrSelectedContent.push(new ApiTable(oTempElm));
				else if (oTempElm instanceof Paragraph)
					arrSelectedContent.push(new ApiParagraph(oTempElm));
				else if (oTempElm instanceof ParaRun)
					arrSelectedContent.push(new ApiRun(oTempElm));
				else if (oTempElm instanceof ParaRun)
					arrSelectedContent.push(new ApiRun(oTempElm));
				else
					continue;
			}

			for (var nElm = 0; nElm < arrSelectedContent.length; nElm ++)
				sOutputText += this.HandleChildElement(arrSelectedContent[nElm], 'html');
		}
		else
		{
			for (var nElm = 0, nElmsCount = oDocument.GetElementsCount(); nElm < nElmsCount; nElm ++)
				sOutputText += this.HandleChildElement(oDocument.GetElement(nElm), 'html');
		}

		return sOutputText;
	};
	CMarkdownConverter.prototype.HandleChildElement = function(oChild, sType)
	{
		var childType = oChild.GetClassType();
		switch (childType)
		{
			case "paragraph":
				return this.HandleParagraph(oChild, sType);
			case "hyperlink":
				return this.HandleHyperlink(oChild, sType);
			case "run":
				return this.HandleRun(oChild, sType);
			case "table":
				return this.HandleTable(oChild, sType);
			case "tableRow":
				return this.HandleTableRow(oChild, sType);
			case "tableCell":
				return this.HandleTableCell(oChild, sType);
			default:
				return '';
		}
	};
	CMarkdownConverter.prototype.HandleParagraph = function(oPara, sType)
	{
		function GetParaNumberingLvl(oParagraph)
		{
			var oNumberingInfo = null;
			if (oParagraph)
			{
				oNumberingInfo = oParagraph.GetNumPr();
				if (oNumberingInfo)
					return oNumberingInfo.Lvl;
			}
			return -1;
		};
		function SetNumbering(sOutputText)
		{
			var isBulleted = null;
			if (sNumId)
				isBulleted = oDocument.Numbering.GetNum(sNumId).GetLvl().IsBulleted();

			// если markdown, то маркируем список без тегов списка
			if (oCMarkdownConverter.Config.convertType === 'markdown' && !oCMarkdownConverter.isTableCellContent)
			{
				// маркированный/нумерованный получают соответсвующие символы для markdown.
				if (sNumId && isBulleted)
					sOutputText = oCMarkdownConverter.WrapInSymbol(sOutputText, '- ', 'open');
				else
					sOutputText = oCMarkdownConverter.WrapInSymbol(sOutputText, String(oPara.Paragraph.GetNumberingCalculatedValue()) + '. ', 'open');

				if (!oCMarkdownConverter.isTableCellContent)
				{
					// отступы для уровней нумерации
					for (var nLvl = 0; nLvl < oNumPr.Lvl; nLvl++)
						sOutputText = oCMarkdownConverter.WrapInSymbol(sOutputText, '   ', 'open');
				}
			}
			else if (oCMarkdownConverter.Config.convertType === 'html' || oCMarkdownConverter.isTableCellContent)
			{
				// если имеем новый уровень маркерованного/нумерованного списка выше текущего, помечаем это в oCMarkdownConverter.currNumberingLvl
				// и открываем новый список для нового уровня
				if (oCMarkdownConverter.currNumberingLvl < oNumPr.Lvl)
				{
					oCMarkdownConverter.currNumberingLvl = oNumPr.Lvl;

					// маркированный/нумерованный получают соответсвующие теги для html.
					if (isBulleted)
					{
						sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Numberring.Item, 'wholly');
						sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Numberring.Bulleted + '\n', 'open');

						// запоминаем открытые списки
						oCMarkdownConverter.openedListsHtml.push(oCMarkdownConverter.HtmlTags.Numberring.Bulleted);
					}
					else
					{
						sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Numberring.Item, 'wholly');
						sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Numberring.Numbered + '\n', 'open');

						// запоминаем открытые списки
						oCMarkdownConverter.openedListsHtml.push(oCMarkdownConverter.HtmlTags.Numberring.Numbered);
					}
				}
				else if (oCMarkdownConverter.currNumberingLvl >= oNumPr.Lvl)
				{
					oCMarkdownConverter.currNumberingLvl = oNumPr.Lvl;

					sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Numberring.Item, 'wholly');
				}

				// если следующий параграф не содержит нумерованный/маркированный список или уровень списка меньше текущего,
				// то закрываем текущий список
				var nNextParaNumberingLvl = GetParaNumberingLvl(oPara.Paragraph.GetNextParagraph());
				if (nNextParaNumberingLvl < oCMarkdownConverter.currNumberingLvl)
				{
					if (isBulleted)
						sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, '\n' + oCMarkdownConverter.HtmlTags.Numberring.Bulleted, 'close');
					else
						sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, '\n' + oCMarkdownConverter.HtmlTags.Numberring.Numbered, 'close');
					oCMarkdownConverter.openedListsHtml.shift();

					// == -1 означает, что следующего параграфа не существует или нет нумерации,
					// значит нужно закрыть все открытые списки
					if (nNextParaNumberingLvl == -1)
					{
						for (var nList = 0, nCount = oCMarkdownConverter.openedListsHtml.length; nList < nCount; nList++)
						{
							sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, '\n' + oCMarkdownConverter.openedListsHtml.shift(), 'close');
						}
					}
					oCMarkdownConverter.isNumbering      = false;
					oCMarkdownConverter.currNumberingLvl = -1;
				}
			}
			return sOutputText;
		};

		function SetHeading(sOutputText)
		{
			// определяем уровень heading у параграфа для Markdown, Title и Subtitle тоже учитываем
			// далее выставляем # соответсвенно уровню.
			var oStyle = oDocument.Get_Styles().Get(oPara.Paragraph.Pr.PStyle);
			var nHeadingLvl = -1;

			switch(oStyle.Name)
			{
				case 'Title':
					nHeadingLvl = 0;
					break;
				case 'Subtitle':
					nHeadingLvl = 1;
					break;
				default:
					nHeadingLvl = oDocument.Get_Styles().GetHeadingLevelByName(oStyle.Name);
					break;
			}

			if (nHeadingLvl !== -1)
			{
				// понижаем уровень заголовка, если указано в конфиге (h1 -> h2)
				if (oCMarkdownConverter.Config.demoteHeadings && nHeadingLvl === 0)
					nHeadingLvl = 1;

				if (oCMarkdownConverter.Config.convertType === 'html' || oCMarkdownConverter.isTableCellContent || oCMarkdownConverter.Config.htmlHeadings)
					return oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Headings['h' + String(nHeadingLvl + 1)],'wholly');
				else if (oCMarkdownConverter.Config.convertType === 'markdown')
					return oCMarkdownConverter.WrapInSymbol(sOutputText, oCMarkdownConverter.MdSymbols.Headings['h' + String(nHeadingLvl + 1)], 'open');
			}
		};
		function SetQuote()
		{
			if (oCMarkdownConverter.Config.convertType === 'html' || oCMarkdownConverter.isTableCellContent)
				return oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Quote, 'wholly');
			else if (oCMarkdownConverter.Config.convertType === 'markdown')
				return oCMarkdownConverter.WrapInSymbol(sOutputText, oCMarkdownConverter.MdSymbols.Quote, 'open');
		};
		function IsHeading(oParagraph)
		{
			var Styles         = private_GetLogicDocument().Get_Styles();
			var sParaStyleName = '';
			if (oParagraph.Paragraph.Pr.PStyle)
				sParaStyleName = Styles.Get(oParagraph.Paragraph.Pr.PStyle).Name;
			else
				return false;
			if (sParaStyleName.search('Heading') !== -1 || sParaStyleName.search('Title') !== -1 || sParaStyleName.search('Subtitle') !== -1)
				return true;
			return false;
		};
		function IsCodeLine(oParagraph)
		{
			if (!oParagraph)
				return false;

			var sFirstRunFont = null;
			var sLastRunFont  = null;
			var oTempRun      = null;

			// если первый ран с текстом и последний имеют моноширинный шрифт, считаем, что это строка кода.
			// находим
			for (var nElm = 0, nElmsCount = oParagraph.GetElementsCount(); nElm < nElmsCount; nElm++)
			{
				oTempRun = oParagraph.GetElement(nElm);
				if (!oTempRun || oTempRun.GetClassType() !== 'run')
					continue;

				if (oTempRun.Run.GetText() !== '')
				{
					sFirstRunFont = oTempRun.Run.Get_CompiledPr().FontFamily.Name;
					break;
				}
			}
			for (var nElm = nElmsCount - 1, nElmsCount = oParagraph.GetElementsCount(); nElm >= 0; nElm--)
			{
				oTempRun = oParagraph.GetElement(nElm);
				if (!oTempRun || oTempRun.GetClassType() !== 'run')
					continue;
				if (oTempRun.Run.GetText() !== '')
				{
					sLastRunFont = oTempRun.Run.Get_CompiledPr().FontFamily.Name;
					break;
				}
			}

			if (private_isMonospaceFont(sFirstRunFont) && private_isMonospaceFont(sLastRunFont))
				return true;

			return false;
		};
		function IsQuoteLine(oParagraph)
		{
			var Styles         = private_GetLogicDocument().Get_Styles();
			var sParaStyleId   = oParagraph.Paragraph.Get_CompiledPr2().ParaPr.Get_PStyle();
			var sQuoteStyleId1 = Styles.GetStyleIdByName('Quote');
			var sQuoteStyleId2 = Styles.GetStyleIdByName('Intense Quote');

			if (sParaStyleId === sQuoteStyleId1 || sParaStyleId === sQuoteStyleId2)
				return true;
			return false;
		};

		function SetCodeBlock(sOutputText)
		{
			if (oCMarkdownConverter.Config.convertType === 'markdown')
			{
				sOutputText = oCMarkdownConverter.WrapInSymbol(sOutputText, '\n' + oCMarkdownConverter.MdSymbols.CodeLine + '\n', 'open');
				// если следующий параграф не с кодом или имеется нумерация или параграф стилизован, то закрываем блок кода
				if (!IsCodeLine(oPara.GetNext()) || GetParaNumberingLvl(oPara.GetNext().Paragraph) !== -1 || oPara.GetNext().Paragraph.Pr.PStyle != undefined)
					sOutputText = oCMarkdownConverter.WrapInSymbol(sOutputText, '\n' + oCMarkdownConverter.MdSymbols.CodeLine + '\n', 'close');
			}
			else if (oCMarkdownConverter.Config.convertType === 'html')
			{
				sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, '\n' + oCMarkdownConverter.HtmlTags.CodeLine.replace('>', ' class="prettyprint">') + '\n', 'open');
				// если следующий параграф не с кодом или имеется нумерация или параграф стилизован, то закрываем блок кода
				if (!IsCodeLine(oPara.GetNext()) || GetParaNumberingLvl(oPara.GetNext().Paragraph) !== -1 || oPara.GetNext().Paragraph.Pr.PStyle != undefined)
					sOutputText = oCMarkdownConverter.WrapInTag(sOutputText, '\n' + oCMarkdownConverter.HtmlTags.CodeLine + '\n', 'close');
			}

			return sOutputText;
		};

		if (!oPara.Next && oPara.GetText().trim() === '')
			return '';
		if (oPara.Paragraph.IsTableCellContent())
			this.isTableCellContent = true;

		var oDocument  = private_GetLogicDocument();
		var sNumId     = oPara.Paragraph.Numbering.Internal.FinalNumId;
		var oNumPr     = oPara.Paragraph.GetNumPr();
		var oCMarkdownConverter = this;


		// если не будет нумерации, тогда проверим на заголовки (одновременно и то и другое в конвертации не применяется)
		if (oNumPr)
		{
			this.isNumbering = true;
			this.isCodeBlock = false;
			this.isHeading   = false;
			this.isQuoteLine = false;
			if (IsQuoteLine(oPara))
				this.isQuoteLine = true;
		}
		else if (IsHeading(oPara))
		{
			this.isNumbering = false;
			this.isCodeBlock = false;
			this.isHeading   = true;
			this.isQuoteLine = false
		}
		else if (IsCodeLine(oPara))
		{
			this.isNumbering = false;
			this.isHeading   = false;
			this.isCodeBlock = true;
			this.isQuoteLine = false;
		}
		else if (IsQuoteLine(oPara))
		{
			this.isNumbering = false;
			this.isHeading   = false;
			this.isCodeBlock = false;
			this.isQuoteLine = true;
		}

		else if (sType === 'html' || this.isTableCellContent)
		{
			this.isNumbering = false;
			this.isCodeBlock = false;
			this.isHeading   = false;
			this.isQuoteLine = false;
		}

		// обработка дочерних элементов
		var sOutputText = '';
		for (var nElm = 0, nElmsCount = oPara.GetElementsCount(); nElm < nElmsCount; nElm++)
			sOutputText += this.HandleChildElement(oPara.GetElement(nElm), sType);

		// вызываем для закрытия тега нумерованного/маркированного списка после обработки дочерних элементов
		if (this.isNumbering)
		{
			if (this.isQuoteLine)
				sOutputText = SetQuote(sOutputText);
			sOutputText = SetNumbering(sOutputText);
		}
		// вызываем для закрытия заголовка после обработки дочерних элементов
		else if (this.isHeading)
			sOutputText = SetHeading(sOutputText);
		// вызываем для закрытия блока кода после обработки дочерних элементов
		else if (this.isCodeBlock)
			sOutputText = SetCodeBlock(sOutputText);
		else if (this.isQuoteLine)
			sOutputText = SetQuote(sOutputText);
		// закрытие тега парагарфа, тег добавляем только в случае, если это не нумерованный список/заголовок/блок кода.
		else if ((sType === 'html' || this.isTableCellContent) && !this.isNumbering && !this.isHeading && !this.isCodeBlock)
			sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Paragraph, 'wholly');
		return sOutputText + '\n';
	};
	CMarkdownConverter.prototype.HandleHyperlink = function(oHyperlink, sType)
	{
		var sOutputText = '';
		if (sType === 'html')
			sOutputText += '<a href="' + oHyperlink.GetLinkedText() + '">';
		else
			sOutputText += '[';

		for (var nElm = 0, nElmsCount = oHyperlink.GetElementsCount(); nElm < nElmsCount; nElm++)
		{
			sOutputText += this.HandleChildElement(oHyperlink.GetElement(nElm));
		}

		if (sType === 'html')
			sOutputText += '</a>';
		else
			sOutputText += ']' + '(' + oHyperlink.GetLinkedText() + ')';

		return sOutputText;
	};
	CMarkdownConverter.prototype.HandleRun = function(oRun, sType)
	{
		function IsHaveCodeRun(oRun)
		{
			if (!oRun)
				return false;

			var oRunTextPr = oRun.Run.Get_CompiledPr();
			return private_isMonospaceFont(oRunTextPr.FontFamily.Name);
		};
		function IsBold(oRun)
		{
			if (!oRun)
				return false;

			var oRunTextPr = oRun.Run.Get_CompiledPr();
			if (oRunTextPr.Bold)
				return true;

			return false;
		};
		function IsItalic(oRun)
		{
			if (!oRun)
				return false;

			var oRunTextPr = oRun.Run.Get_CompiledPr();
			if (oRunTextPr.Italic)
				return true;

			return false;
		};

		function GetTextWithPicture(oRun)
		{
			var sText = '';

			var ContentLen = oRun.Run.Content.length;

			for (var CurPos = 0; CurPos < ContentLen; CurPos++)
			{
				var Item     = oRun.Run.Content[CurPos];
				var ItemType = Item.Type;

				switch (ItemType)
				{
					case para_Drawing:
					{
						if (Item.IsPicture())
						{
							if (sType === 'markdown')
								sText += oCMarkdownConverter.Config.base64img ? '![](' + Item.GraphicObj.getBase64Img() + ')' : '![](' + Item.GraphicObj.getImageUrl() + ')';
							else if (sType === 'html')
								sText += oCMarkdownConverter.Config.base64img ? '<img src="' + Item.GraphicObj.getBase64Img() + '">' : '<img src="' + Item.GraphicObj.getImageUrl() + '">';
						}
					}
					case para_PageNum:
					case para_PageCount:
					case para_End:
					{
						break;
					}
					case para_Text :
					{
						sText += String.fromCharCode(Item.Value);
						break;
					}
					case para_Space:
					case para_NewLine:
					case para_Tab:
					{
						sText += " ";
						break;
					}
				}
			}

			return sText;
		};

		var oCMarkdownConverter    = this;
		var arrAllDrawings = oRun.Run.GetAllDrawingObjects();
		var hasPicture     = false;
		var sOutputText    = oRun.Run.GetText();
		var oTextPr        = oRun.Run.Get_CompiledPr();

		// находим все картинки и их позиции в строке
		for (var nDrawing = 0; nDrawing < arrAllDrawings.length; nDrawing ++)
		{
			if (arrAllDrawings[nDrawing].IsPicture())
			{
				hasPicture = true;
				break;
			}
		}

		if (sOutputText === '' && hasPicture === false)
			return '';

		// рендер html тагов
		if (!this.Config.renderHTMLTags)
			sOutputText = sOutputText.replace(/</gi, '&lt;');

		if (!this.isCodeBlock)
		{
			// возможно текст в ране представляет собой блок кода, обрабатываем это
			if (private_isMonospaceFont(oTextPr.FontFamily.Name))
			{
				var oRunNext     = oRun.GetNext();
				while (oRunNext && oRunNext.Run.GetText() === '')
					oRunNext = oRunNext.GetNext();

				var oRunPrev     = oRun.GetPrevious();
				while (oRunPrev && oRunPrev.Run.GetText() === '')
					oRunPrev = oRunPrev.GetPrevious();

				var isCodeNextRun = IsHaveCodeRun(oRunNext);
				var isCodePrevRun = IsHaveCodeRun(oRunPrev);

				if (sType === 'html' || this.Config.htmlHeadings)
				{
					if (!isCodePrevRun && !isCodeNextRun)
						sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Code, 'wholly');
					else if (!isCodePrevRun)
						sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Code, 'open');
					else if (!isCodeNextRun)
						sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Code, 'close');
				}
				else if (sType === 'markdown')
				{
					if (!isCodePrevRun && !isCodeNextRun)
						sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Code, 'wholly');
					else if (!isCodePrevRun)
						sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Code, 'open');
					else if (!isCodeNextRun)
						sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Code, 'close');
				}
			}
			else
			{
				var oRunNext     = oRun.GetNext();
				while (oRunNext && oRunNext.Run.GetText() === '')
					oRunNext = oRunNext.GetNext();

				var oRunPrev     = oRun.GetPrevious();
				while (oRunPrev && oRunPrev.Run.GetText() === '')
					oRunPrev = oRunPrev.GetPrevious();

				var isBoldNextRun   = IsBold(oRunNext);
				var isBoldPrevRun   = IsBold(oRunPrev);
				var isItalicNextRun = IsItalic(oRunNext);
				var isItalicPrevRun = IsItalic(oRunPrev);

				if (hasPicture)
					sOutputText = GetTextWithPicture(oRun);

				if (oTextPr.Bold)
				{
					if (sType === 'html' || this.Config.htmlHeadings)
					{
						if (!isBoldPrevRun && !isBoldNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Bold, 'wholly');
						else if (!isBoldPrevRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Bold, 'open');
						else if (!isBoldNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Bold, 'close');
					}
					else if (sType === 'markdown')
					{
						if (!isBoldPrevRun && !isBoldNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Bold, 'wholly');
						else if (!isBoldPrevRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Bold, 'open');
						else if (!isBoldNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Bold, 'close');
					}
				}

				if (oTextPr.Italic && !this.isQuoteLine)
				{
					if (sType === 'html' || this.Config.htmlHeadings)
					{
						if (!isItalicPrevRun && !isItalicNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Italic, 'wholly');
						else if (!isItalicPrevRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Italic, 'open');
						else if (!isItalicNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Italic, 'close');
					}
					else if (sType === 'markdown')
					{
						if (!isItalicPrevRun && !isItalicNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Italic, 'wholly');
						else if (!isItalicPrevRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Italic, 'open');
						else if (!isItalicNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Italic, 'close');
					}
				}
			}
		}

		return sOutputText;
	};
	CMarkdownConverter.prototype.HandleTable = function(oTable)
	{
		var sOutputText = '<table>\n';

		for (var nRow = 0, nRowsCount = oTable.GetRowsCount(); nRow < nRowsCount; nRow++)
		{
			sOutputText += this.HandleChildElement(oTable.GetRow(nRow), 'html');
		}

		sOutputText += '</table>\n';
		return sOutputText;
	};
	CMarkdownConverter.prototype.HandleTableRow = function(oTableRow)
	{
		var sOutputText = '  <tr>\n';

		for (var nCell = 0, nCellsCount = oTableRow.GetCellsCount(); nCell < nCellsCount; nCell++)
		{
			sOutputText += this.HandleChildElement(oTableRow.GetCell(nCell), 'html');
		}

		sOutputText += '  </tr>\n';
		return sOutputText;
	};
	CMarkdownConverter.prototype.HandleTableCell = function(oTableCell)
	{
		var sOutputText = '   <td>\n';
		var apiCellContent = oTableCell.GetContent();

		for (var nElm = 0, nElmsCount = apiCellContent.GetElementsCount(); nElm < nElmsCount; nElm++)
		{
			sOutputText += this.HandleChildElement(apiCellContent.GetElement(nElm), 'html');
		}

		sOutputText += '</td>\n';
		return sOutputText;
	};
	/**
	 * Class represents a continuous region in a document. 
	 * Each Range object is determined by the position of the start and end characters.
	 * @param oElement - May be Document, Table, Paragraph, Run, Hyperlink.
	 * @param {Number} Start - The start element of Range in the current Element.
	 * @param {Number} End - The end element of Range in the current Element.
	 * @constructor
	 */
	function ApiRange(oElement, Start, End)
	{
		this.Element		= oElement;
		this.Controller		= null;
		this.Start			= undefined;
		this.End 		 	= undefined;
		this.isEmpty 		= true;
		this.Paragraphs 	= [];
		this.Text 			= undefined;
		this.oDocument		= editor.GetDocument();
		this.EndPos			= null;
		this.StartPos		= null;
		this.TextPr 		= new CTextPr();

		this.private_SetRangePos(Start, End);
		this.private_CalcDocPos();

		if (this.StartPos === null || this.EndPos === null)
			return false;
		else 
			this.isEmpty = false;

		this.private_SetController();
		
		this.Text 		= this.GetText();
		this.Paragraphs = this.GetAllParagraphs();

		private_RefreshRangesPosition();
		arrApiRanges.push(this);
		this.private_RemoveEqual();
		private_TrackRangesPositions(true);
	};

	ApiRange.prototype.constructor = ApiRange;
	ApiRange.prototype.private_SetRangePos = function(Start, End)
	{
		function calcSumChars(oRun)
		{
			var nRangePos = 0;

			var nCurPos = oRun.Content.length;
			
			for (var nPos = 0; nPos < nCurPos; ++nPos)
			{
				if (para_Text === oRun.Content[nPos].Type || para_Space === oRun.Content[nPos].Type || para_Tab === oRun.Content[nPos].Type)
					nRangePos++;
			}

			if (nRangePos !== 0)
				charsCount += nRangePos;
		};

		if (Start > End)
		{
			var temp	= Start;
			Start		= End;
			End			= temp;
		}
		if (Start === undefined)
			this.Start = 0;
		else if (typeof(Start) === "number")
			this.Start = Start
		else if (Array.isArray(Start) === true)
			this.StartPos = Start;

		if (End === undefined)
		{
			this.End 		= 0;
			var charsCount 	= 0;

			this.Element.CheckRunContent(calcSumChars);
			
			this.End = charsCount;
			if (this.End > 0)
				this.End--;
		}
		else if (typeof(End) === "number")
			this.End = End;
		else if (Array.isArray(End) === true)
			this.EndPos = End;
	};
	ApiRange.prototype.private_CalcDocPos = function()
	{
		if (this.StartPos || this.EndPos)
			return;

		var isStartDocPosFinded = false;
		var isEndDocPosFinded	= false;
		var StartChar			= this.Start;
		var EndChar				= this.End;
		var StartPos			= null;
		var EndPos				= null;
		var charsCount 			= 0;

		function callback(oRun)
		{
			var nRangePos = 0;

			var nCurPos = oRun.Content.length;
			for (var nPos = 0; nPos < nCurPos; ++nPos)
			{
				if (para_Text === oRun.Content[nPos].Type || para_Space === oRun.Content[nPos].Type || para_Tab === oRun.Content[nPos].Type)
					nRangePos++;

				if (StartChar - charsCount === nRangePos - 1 && !isStartDocPosFinded)
				{
					var DocPosInRun = 
					{
						Class : oRun,
						Position : nPos,
					};
		
					var DocPos = oRun.GetDocumentPositionFromObject();
		
					DocPos.push(DocPosInRun);
		
					StartPos = DocPos;

					isStartDocPosFinded = true;
				}
				
				if (EndChar - charsCount === nRangePos - 1 && !isEndDocPosFinded)
				{
					var DocPosInRun = 
					{
						Class : oRun,
						Position : nPos + 1,
					};
		
					var DocPos = oRun.GetDocumentPositionFromObject();
		
					DocPos.push(DocPosInRun);
		
					EndPos = DocPos;

					isEndDocPosFinded = true;
				}
			}

			if (nRangePos !== 0)
				charsCount += nRangePos;
		};

		if (this.Element instanceof CDocument || this.Element instanceof CDocumentContent || this.Element instanceof CTable || this.Element instanceof CBlockLevelSdt)
		{
			var allParagraphs	= this.Element.GetAllParagraphs({OnlyMainDocument : true, All : true});

			for (var paraItem = 0; paraItem < allParagraphs.length; paraItem++)
			{
				if (isStartDocPosFinded && isEndDocPosFinded)
					break;
				else 
					allParagraphs[paraItem].CheckRunContent(callback);

					this.StartPos	= StartPos;
					this.EndPos		= EndPos;
			}
		}
		else if (this.Element instanceof Paragraph || this.Element instanceof ParaHyperlink || this.Element instanceof CInlineLevelSdt || this.Element instanceof ParaRun)
		{
			this.Element.CheckRunContent(callback);
			
			this.StartPos	= StartPos;
			this.EndPos		= EndPos;
		}
	};
	ApiRange.prototype.private_SetController = function()
	{
		if (this.StartPos[0].Class.IsHdrFtr())
		{
			this.Controller = this.oDocument.Document.GetHdrFtr();
		}
		else if (this.StartPos[0].Class.IsFootnote())
		{
			this.Controller = this.oDocument.Document.GetFootnotesController();
		}
		else if (this.StartPos[0].Class.Is_DrawingShape())
		{
			this.Controller = this.oDocument.Document.DrawingsController;
		}
		else 
		{
			this.Controller = this.oDocument.Document.LogicDocumentController;
		}
	};
	ApiRange.prototype.private_RemoveEqual = function()
	{
		function ckeck_equal(firstDocPos, secondDocPos)
		{
			if (firstDocPos.length === secondDocPos.length)
			{
				for (var nPos = 0; nPos < firstDocPos.length; nPos++)
				{
					if (firstDocPos[nPos].Class === secondDocPos[nPos].Class && firstDocPos[nPos].Position === secondDocPos[nPos].Position)
						continue;
					else 
						return false;
				}
				return true;
			}
			return false;
		};

		var Range = null;
		for (var nRange = 0; nRange < arrApiRanges.length - 1; nRange++)
		{
			Range = arrApiRanges[nRange];
			if (ckeck_equal(this.StartPos, Range.StartPos) && ckeck_equal(this.EndPos, Range.EndPos))
			{
				arrApiRanges.splice(nRange, 1);
				nRange--;
			}
		}
	};
	
	/**
	 * Get the type of this class.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @returns {"range"}
	 */
	ApiRange.prototype.GetClassType = function()
	{
		return "range";
	};

	/**
	 * Get a paragraph from all the paragraphs that are in the range.
	 * @param {Number} nPos - The paragraph position in the range.
	 * @return {ApiParagraph | null} - returns null if position is invalid.
	 */	
	ApiRange.prototype.GetParagraph = function(nPos)
	{
		this.GetAllParagraphs();

		if (nPos > this.Paragraphs.length - 1 || nPos < 0)
			return null;
		
		if (this.Paragraphs[nPos])
			return this.Paragraphs[nPos];
		else 
			return null;
	};

	/**
	 * Add text to the specified position.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {String} sText - The text that will be added.
	 * @param {string} [sPosition = "after"] - The position where the text will be added ("before" or "after" the range specified).
	 * @return {bool} - returns false if range is empty or sText isn't string value.
	 */	
	ApiRange.prototype.AddText = function(sText, sPosition)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document = private_GetLogicDocument();
		Document.RemoveSelection();

		if (this.isEmpty || this.isEmpty === undefined || typeof(sText) !== "string")
			return false;

		if (sPosition !== "after" && sPosition !== "before")
			sPosition = "after";
		
		if (sPosition === "after")
		{
			var lastRun				= this.EndPos[this.EndPos.length - 1].Class;
			var lastRunPos			= this.EndPos[this.EndPos.length - 1].Position;
			var lastRunPosInParent	= this.EndPos[this.EndPos.length - 2].Position;
			var lastRunParent		= lastRun.GetParent();
			var newRunPos			= lastRunPos;
			if (lastRunPos === 0)
			{
				if (lastRunPosInParent - 1 >= 0)
				{
					lastRunPosInParent--;
					lastRun		= lastRunParent.GetElement(lastRunPosInParent);
					lastRunPos	= lastRun.Content.length;
				}
			}
			else 
				for (var oIterator = sText.getUnicodeIterator(); oIterator.check(); oIterator.next())
					newRunPos++;

			lastRun.AddText(sText, lastRunPos);
			this.EndPos[this.EndPos.length - 1].Class = lastRun;
			this.EndPos[this.EndPos.length - 1].Position = newRunPos;
			this.EndPos[this.EndPos.length - 2].Position = lastRunPosInParent;
			private_TrackRangesPositions(true);
		}
		else if (sPosition === "before")
		{
			var firstRun		= this.StartPos[this.StartPos.length - 1].Class;
			var firstRunPos		= this.StartPos[this.StartPos.length - 1].Position;
			firstRun.AddText(sText, firstRunPos);
		}

		return true;
	};

	/**
	 * Add the bookmark to the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {String} sName - The bookmark name.
	 * @return {bool} - returns false if range is empty.
	 */	
	ApiRange.prototype.AddBookmark = function(sName)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined || typeof(sName) !== "string")
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return false;
		}
		private_TrackRangesPositions();

		Document.RemoveBookmark(sName);
		Document.AddBookmark(sName);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return true;
	};

	/**
	 * Add a hyperlink to the specified range. 
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {string} sLink - The link address.
	 * @param {string} sScreenTipText - The screen tip text.
	 * @return {ApiHyperlink | null}  - returns null if range contains more than one paragraph or sLink is invalid. 
	 */
	ApiRange.prototype.AddHyperlink = function(sLink, sScreenTipText)
	{
		if (typeof(sLink) !== "string" || sLink === "" || sLink.length > Asc.c_nMaxHyperlinkLength)
			return null;
		if (typeof(sScreenTipText) !== "string")
			sScreenTipText = "";

		this.GetAllParagraphs();
		if (this.Paragraphs.length > 1)
			return null;

		var Document	= editor.private_GetLogicDocument();
		var hyperlinkPr	= new Asc.CHyperlinkProperty();
		var urlType		= AscCommon.getUrlType(sLink);
		var oHyperlink	= null;

		if (!/(((^https?)|(^ftp)):\/\/)|(^mailto:)/i.test(sLink))
			sLink = (urlType === 0) ? null :(( (urlType === 2) ? 'mailto:' : 'http://' ) + sLink);

		sLink = sLink.replace(new RegExp("%20",'g')," ");
		hyperlinkPr.put_Value(sLink);
		hyperlinkPr.put_ToolTip(sScreenTipText);
		hyperlinkPr.put_Bookmark(null);

		this.Select(false);
		oHyperlink = new ApiHyperlink(this.Paragraphs[0].Paragraph.AddHyperlink(hyperlinkPr));
		Document.RemoveSelection();

		return oHyperlink;
	};

	/**
	 * Get text from the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @returns {String} - returns "" if range is empty.
	 */	
	ApiRange.prototype.GetText = function()
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return "";
		}
		private_TrackRangesPositions();

		var Text = this.Controller.GetSelectedText(false); 
		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return Text;
	};

	/**
	 * Get a collection of paragraphs that represents all the paragraphs in the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @return {ApiParagraph[]}
	 */	
	ApiRange.prototype.GetAllParagraphs = function()
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		if (this.isEmpty || this.isEmpty === undefined)
			return false;

		var done = false;

		var AllParagraphsListOfElement = [];
		var RangeParagraphsList = [];

		var startPara = this.StartPos[this.StartPos.length - 1].Class.GetParagraph();
		var endPara   = this.EndPos[this.EndPos.length - 1].Class.GetParagraph();

		if (startPara instanceof ParaHyperlink)
		{
			startPara = startPara.Paragraph;
		}

		if (endPara instanceof ParaHyperlink)
		{
			endPara = endPara.Paragraph;
		}

		if (startPara.Id === endPara.Id)
		{
			RangeParagraphsList.push(new ApiParagraph(startPara));
			return RangeParagraphsList;
		}

		if (this.Element instanceof CDocument || this.Element instanceof CTable || this.Element instanceof CBlockLevelSdt)
		{
			AllParagraphsListOfElement = this.Element.GetAllParagraphs({All : true});

			for (var Index1 = 0; Index1 < AllParagraphsListOfElement.length; Index1++)
			{
				if (done)
					break;

				if (AllParagraphsListOfElement[Index1].Id === startPara.Id)
				{
					RangeParagraphsList.push(new ApiParagraph(AllParagraphsListOfElement[Index1]));

					for (var Index2 = Index1 + 1; Index2 < AllParagraphsListOfElement.length; Index2++)
					{
						if (AllParagraphsListOfElement[Index2].Id !== endPara.Id)
						{
							RangeParagraphsList.push(new ApiParagraph(AllParagraphsListOfElement[Index2]));
						}
						else 
						{
							RangeParagraphsList.push(new ApiParagraph(endPara));

							done = true;
							break;
						}
					}
				}
			}
		}

		this.Paragraphs = RangeParagraphsList;

		return RangeParagraphsList;
	};

	/**
	 * Set the selection to the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 */	
	ApiRange.prototype.Select = function(bUpdate)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document = private_GetLogicDocument();
		
		if (this.isEmpty || this.isEmpty === undefined)
			return false;

		if (bUpdate === undefined)
			bUpdate = true;

		this.StartPos[0].Class.SetContentPosition(this.StartPos, 0, 0);
		this.StartPos[0].Class.SetSelectionByContentPositions(this.StartPos, this.EndPos);

		if (bUpdate)
		{
			var controllerType = null;

			if (this.StartPos[0].Class.IsHdrFtr())
			{
				controllerType = docpostype_HdrFtr;
			}
			else if (this.StartPos[0].Class.IsFootnote())
			{
				controllerType = docpostype_Footnotes;
			}
			else if (this.StartPos[0].Class.Is_DrawingShape())
			{
				controllerType = docpostype_DrawingObjects;
			}
			else 
			{
				controllerType = docpostype_Content;
			}
			Document.SetDocPosType(controllerType);
			Document.UpdateSelection();
		}
	};

	/**
	 * Return a new range that goes beyond the specified range in any direction and spans a different range. The current range has not changed. Throws an error if two ranges do not have a union.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {ApiRange} oRange - The range that will be expanded.
	 * @return {ApiRange | null} - returns null if the specified range can't be expanded. 
	 */	
	ApiRange.prototype.ExpandTo = function(oRange)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		if (!(oRange instanceof ApiRange) || this.isEmpty || this.isEmpty === undefined || oRange.isEmpty || oRange.isEmpty === undefined)
			return null;

		var firstStartPos 		= this.StartPos;
		var firstEndPos			= this.EndPos;
		var secondStartPos		= oRange.StartPos;
		var secondEndPos		= oRange.EndPos;

		if (this.Controller !== oRange.Controller)
			return null;

		function check_pos(firstPos, secondPos)
		{
			for (var nPos = 0, nLen = Math.min(firstPos.length, secondPos.length); nPos < nLen; ++nPos)
			{
				if (!secondPos[nPos] || !firstPos[nPos] || firstPos[nPos].Class !== secondPos[nPos].Class)
					return 1;

				if (firstPos[nPos].Position < secondPos[nPos].Position)
					return 1;
				else if (firstPos[nPos].Position > secondPos[nPos].Position)
					return -1;
			}

			return 1;
		}
		
		var newRangeStartPos	= null;
		var newRangeEndPos		= null;

		if (check_pos(firstStartPos, secondStartPos) === 1)
			newRangeStartPos = firstStartPos;
		else 
			newRangeStartPos = secondStartPos;

		if (check_pos(firstEndPos, secondEndPos) === 1)
			newRangeEndPos = secondEndPos;
		else 
			newRangeEndPos = firstEndPos;

		return new ApiRange(newRangeStartPos[0].Class, newRangeStartPos, newRangeEndPos);
	};

	/**
	 * Return a new range as the intersection of this range with another range. The current range has not changed. Throws an error if two ranges do not overlap or are not adjacent.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {ApiRange} oRange - The range that will be intersected with another range.
	 * @return {ApiRange | null} - returns null if can't intersect.
	 */	
	ApiRange.prototype.IntersectWith = function(oRange)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		if (!(oRange instanceof ApiRange) || this.isEmpty || this.isEmpty === undefined || oRange.isEmpty || oRange.isEmpty === undefined)
			return null;

		var firstStartPos 		= this.StartPos;
		var firstEndPos			= this.EndPos;
		var secondStartPos		= oRange.StartPos;
		var secondEndPos		= oRange.EndPos;

		if (this.Controller !== oRange.Controller)
			return null;

		function check_direction(firstPos, secondPos)
		{
			for (var nPos = 0, nLen = Math.min(firstPos.length, secondPos.length); nPos < nLen; ++nPos)
			{
				if (!secondPos[nPos] || !firstPos[nPos] || firstPos[nPos].Class !== secondPos[nPos].Class)
					return 1;

				if (firstPos[nPos].Position < secondPos[nPos].Position)
					return 1;
				else if (firstPos[nPos].Position > secondPos[nPos].Position)
					return -1;
			}

			return 1;
		}
		
		var newRangeStartPos	= null;
		var newRangeEndPos		= null;

		// Взаимное расположение диапазонов относительно друг друга. A и B - начало и конец первого диапазона, C и D - начало и конец второго диапазона.
		var AC	= check_direction(firstStartPos, secondStartPos);
		var AD	= check_direction(firstStartPos, secondEndPos);
		var BC	= check_direction(firstEndPos, secondStartPos);
		var BD	= check_direction(firstEndPos, secondEndPos);

		if (AC === AD && AC === BC && AC === BD)
			return null;
		else if (AC === BD && AD !== BC)
		{
			if (AC === 1)
			{
				newRangeStartPos	= secondStartPos;
				newRangeEndPos		= firstEndPos;
			}
			else if (AC === - 1)
			{
				newRangeStartPos	= firstStartPos;
				newRangeEndPos		= secondEndPos;
			}
		}
		else if (AC !== BD && AD !== BC)
		{
			if (AC === 1)
			{
				newRangeStartPos	= secondStartPos;
				newRangeEndPos		= secondEndPos;
			}
			else if (AC === - 1)
			{
				newRangeStartPos	= firstStartPos;
				newRangeEndPos		= firstEndPos;
			}
		}

		return new ApiRange(newRangeStartPos[0].Class, newRangeStartPos, newRangeEndPos);
	};

	/**
	 * Set the bold property to the text character.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {bool} isBold - Specifies if the Range contents are displayed bold or not.
	 * @returns {ApiRange | null} - returns null if can't apply bold.
	 */
	ApiRange.prototype.SetBold = function(isBold)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({Bold : isBold});

		this.Controller.AddToParagraph(ParaTextPr);
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify that any lowercase characters in this text Range are formatted for display only as their capital letter character equivalents.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {bool} isCaps - Specifies if the Range contents are displayed capitalized or not.
	 * @returns {ApiRange | null} - returns null if can't apply caps.
	 */
	ApiRange.prototype.SetCaps = function(isCaps)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({Caps : isCaps});
		Document.AddToParagraph(ParaTextPr);
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Set the text color for the current text Range in the RGB format.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - If this parameter is set to "true", then r,g,b parameters will be ignored.
	 * @returns {ApiRange | null} - returns null if can't apply color.
	 */
	ApiRange.prototype.SetColor = function(r, g, b, isAuto)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var color = new Asc.asc_CColor();
		color.r    = r;
		color.g    = g;
		color.b    = b;
		color.Auto = isAuto;

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = null;
		if (true === color.Auto)
		{
			ParaTextPr = new AscCommonWord.ParaTextPr({
				Color      : {
					Auto : true,
					r    : 0,
					g    : 0,
					b    : 0
				}, Unifill : undefined
			});
			Document.AddToParagraph(ParaTextPr);
		}
		else
		{
			var Unifill        = new AscFormat.CUniFill();
			Unifill.fill       = new AscFormat.CSolidFill();
			Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
			ParaTextPr = new AscCommonWord.ParaTextPr({Unifill : Unifill});
			Document.AddToParagraph(ParaTextPr);
		}

		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify that the contents of this Range are displayed with two horizontal lines through each character displayed on the line.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isDoubleStrikeout - Specifies if the contents of the current Range are displayed double struck through or not.
	 * @returns {ApiRange | null} - returns null if can't apply double strikeout.
	 */
	ApiRange.prototype.SetDoubleStrikeout = function(isDoubleStrikeout)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();
		
		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({DStrikeout : isDoubleStrikeout});
		Document.AddToParagraph(ParaTextPr);
		
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify a highlighting color in the RGB format which is applied as a background for the contents of the current Range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {bool} [isNone=false] If this parameter is set to "true", then r,g,b parameters will be ignored.
	 * @returns {ApiRange | null} - returns null if can't apply highlight.
	 */
	ApiRange.prototype.SetHighlight = function(r, g, b, isNone)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var TextPr = null;
		if (true === isNone)
		{
			TextPr = new ParaTextPr({HighLight : highlight_None});
			Document.AddToParagraph(TextPr);
		}
		else
		{
			var color = new CDocumentColor(r, g, b);
			TextPr = new ParaTextPr({HighLight : color});
			Document.AddToParagraph(TextPr);
		}

		this.TextPr.Merge(TextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify the shading applied to the contents of the current text Range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type applied to the contents of the current text Range.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @returns {ApiRange | null} - returns null if can't apply shadow.
	 */
	ApiRange.prototype.SetShd = function(sType, r, g, b)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var color = new Asc.asc_CColor();
		color.r    = r;
		color.g    = g;
		color.b    = b;
		color.Auto = false;

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var Shd = new CDocumentShd();

		if (sType === "nil")
		{
			var _Shd = {Value : Asc.c_oAscShdNil};
			Shd.Set_FromObject(_Shd);
			Document.SetParagraphShd(_Shd);
		}
		else if (sType === "clear")
		{
			var Unifill        = new AscFormat.CUniFill();
			Unifill.fill       = new AscFormat.CSolidFill();
			Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
			var _Shd = {
				Value   : Asc.c_oAscShdClear,
				Color   : {
					r : color.asc_getR(),
					g : color.asc_getG(),
					b : color.asc_getB()
				},
				Unifill : Unifill
			};
			
			Shd.Set_FromObject(_Shd);
			Document.SetParagraphShd(_Shd);
		}

		this.TextPr.Shd = Shd;
		
		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Set the italic property to the text character.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isItalic - Specifies if the contents of the current Range are displayed italicized or not.
	 * @returns {ApiRange | null} - returns null if can't apply italic.
	 */
	ApiRange.prototype.SetItalic = function(isItalic)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}
				
		var ParaTextPr = new AscCommonWord.ParaTextPr({Italic : isItalic});
		Document.AddToParagraph(ParaTextPr);

		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify that the contents of this Range are displayed with a single horizontal line through the range center.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isStrikeout - Specifies if the contents of the current Range are displayed struck through or not.
	 * @returns {ApiRange | null} - returns null if can't apply strikeout.
	 */
	ApiRange.prototype.SetStrikeout = function(isStrikeout)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();
		
		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({
			Strikeout  : isStrikeout,
			DStrikeout : false
			});
		Document.AddToParagraph(ParaTextPr);
		
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify that all the small letter characters in this text Range are formatted for display only as their capital
	 * letter character equivalents which are two points smaller than the actual font size specified for this text.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isSmallCaps - Specifies if the contents of the current Range are displayed capitalized two points smaller or not.
	 * @returns {ApiRange | null} - returns null if can't apply small caps.
	 */
	ApiRange.prototype.SetSmallCaps = function(isSmallCaps)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({
			SmallCaps : isSmallCaps,
			Caps      : false
		});
		Document.AddToParagraph(ParaTextPr);
		
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Set text spacing measured in twentieths of a point.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {twips} nSpacing - The value of the text spacing measured in twentieths of a point (1/1440 of an inch).
	 * @returns {ApiRange | null} - returns null if can't apply spacing.
	 */
	ApiRange.prototype.SetSpacing = function(nSpacing)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({Spacing : nSpacing});
		Document.AddToParagraph(ParaTextPr);
		
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify that the contents of this Range are displayed along with a line appearing directly below the character
	 * (less than all the spacing above and below the characters on the line).
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isUnderline - Specifies if the contents of the current Range are displayed underlined or not.
	 * @returns {ApiRange | null} - returns null if can't apply underline.
	 */
	ApiRange.prototype.SetUnderline = function(isUnderline)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();
		
		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}


		var ParaTextPr = new AscCommonWord.ParaTextPr({Underline : isUnderline});
		Document.AddToParagraph(ParaTextPr);
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify the alignment which will be applied to the Range contents in relation to the default appearance of the Range text:
	 * * <b>"baseline"</b> - the characters in the current text Range will be aligned by the default text baseline.
	 * * <b>"subscript"</b> - the characters in the current text Range will be aligned below the default text baseline.
	 * * <b>"superscript"</b> - the characters in the current text Range will be aligned above the default text baseline.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {("baseline" | "subscript" | "superscript")} sType - The vertical alignment type applied to the text contents.
	 * @returns {ApiRange | null} - returns null if can't apply align.
	 */
	ApiRange.prototype.SetVertAlign = function(sType)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var value = undefined;

		if (sType === "baseline")
			value = 0;
		else if (sType === "subscript")
			value = 2;
		else if (sType === "superscript")
			value = 1;
		else 
			return null;

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({VertAlign : value});
		Document.AddToParagraph(ParaTextPr);
		
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specify the amount by which text is raised or lowered for this Range in relation to the default
	 * baseline of the surrounding non-positioned text.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {hps} nPosition - Specifies a positive (raised text) or negative (lowered text)
	 * measurement in half-points (1/144 of an inch).
	 * @returns {ApiRange | null} - returns null if can't set position.
	 */
	ApiRange.prototype.SetPosition = function(nPosition)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		if (typeof nPosition !== "number")
			return null;

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({Position : nPosition});
		Document.AddToParagraph(ParaTextPr);
		
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Set the font size for the characters of the current text Range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {hps} nSize - The text size value measured in half-points (1/144 of an inch).
	 * @returns {ApiRange | null} - returns null if can't set font size.
	 */
	ApiRange.prototype.SetFontSize = function(FontSize)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr({FontSize : FontSize});
		Document.AddToParagraph(ParaTextPr);
		
		this.TextPr.Merge(ParaTextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Set all 4 font slots with the specified font family.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {string} sFontFamily - The font family or families used for the current text Range.
	 * @returns {ApiRange | null} - returns null if can't set font family.
	 */
	ApiRange.prototype.SetFontFamily = function(sFontFamily)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		if (typeof sFontFamily !== "string")
			return null;

		var loader				= AscCommon.g_font_loader;
		var fontinfo			= g_fontApplication.GetFontInfo(sFontFamily);
		var isasync				= loader.LoadFont(fontinfo);
		var Document			= null;
		var oldSelectionInfo	= undefined;

		if (isasync === false)
		{
			Document			= private_GetLogicDocument();
			oldSelectionInfo	= Document.SaveDocumentState();

			this.Select(false);
			if (this.isEmpty || this.isEmpty === undefined)
			{
				Document.LoadDocumentState(oldSelectionInfo);
				return null;
			}

			private_TrackRangesPositions();

			var FontFamily = {
				Name : sFontFamily,
				Index : -1
			};

			var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
			if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
			{
				Document.LoadDocumentState(oldSelectionInfo);
				Document.UpdateSelection();
	
				return null;
			}
	
			var ParaTextPr = new AscCommonWord.ParaTextPr({FontFamily : FontFamily});
			Document.AddToParagraph(ParaTextPr);
			
			this.TextPr.Merge(ParaTextPr.Value);

			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return this;
		}
	};

	/**
	 * Set the style for the current Range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {ApiStyle} oStyle - The style which must be applied to the text character.
	 * @returns {ApiRange | null} - returns null if can't set style.
	 */
	ApiRange.prototype.SetStyle = function(oStyle)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined || !(oStyle instanceof ApiStyle))
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		Document.SetParagraphStyle(oStyle.GetName(), true);
		
		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Set the text properties for the current Range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The text properties that will be applied to the current range.
	 * @returns {ApiRange | null} - returns null if can't set text properties.
	 */
	ApiRange.prototype.SetTextPr = function(oTextPr)
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined || !(oTextPr instanceof ApiTextPr))
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return null;
		}

		private_TrackRangesPositions();

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		var ParaTextPr = new AscCommonWord.ParaTextPr(oTextPr.TextPr);
		Document.AddToParagraph(ParaTextPr);
		this.TextPr.Set_FromObject(oTextPr.TextPr);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Delete all the contents from the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @returns {bool} - returns false if range is empty.
	 */
	ApiRange.prototype.Delete = function()
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();
		
		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= Document.SaveDocumentState();

		this.Select(false);
		if (this.isEmpty || this.isEmpty === undefined)
		{
			Document.LoadDocumentState(oldSelectionInfo);
			return false;
		}

		private_TrackRangesPositions();

		this.Controller.Remove(1, true, false, false, false);

		this.isEmpty = true;
		
		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();
		
		return true;
	};

	/**
	 * Class representing a document.
	 * @constructor
	 * @extends {ApiDocumentContent}
	 */
	function ApiDocument(Document)
	{
		ApiDocumentContent.call(this, Document);
	}

	ApiDocument.prototype = Object.create(ApiDocumentContent.prototype);
	ApiDocument.prototype.constructor = ApiDocument;

	/**
	 * Class representing the paragraph properties.
	 * @constructor
	 */
	function ApiParaPr(Parent, ParaPr)
	{
		this.Parent = Parent;
		this.ParaPr = ParaPr;
	}
	
	
	/**
	 * Class representing a paragraph bullet.
	 * @constructor
	 */
	function ApiBullet(Bullet)
	{
		this.Bullet = Bullet;
	}

	/**
	 * Class representing a paragraph.
	 * @constructor
	 * @extends {ApiParaPr}
	 */
	function ApiParagraph(Paragraph)
	{
		ApiParaPr.call(this, this, Paragraph.Pr.Copy());
		this.Paragraph = Paragraph;
	}
	ApiParagraph.prototype = Object.create(ApiParaPr.prototype);
	ApiParagraph.prototype.constructor = ApiParagraph;

	/**
	 * Class representing the table properties.
	 * @constructor
	 */
	function ApiTablePr(Parent, TablePr)
	{
		this.Parent  = Parent;
		this.TablePr = TablePr;
	}

	/**
	 * Class representing a table.
	 * @constructor
	 * @extends {ApiTablePr}
	 */
	function ApiTable(Table)
	{
		ApiTablePr.call(this, this, Table.Pr.Copy());
		this.Table = Table;
	}
	ApiTable.prototype = Object.create(ApiTablePr.prototype);
	ApiTable.prototype.constructor = ApiTable;

	/**
	 * Class representing the text properties.
	 * @constructor
	 */
	function ApiTextPr(Parent, TextPr)
	{
		this.Parent = Parent;
		this.TextPr = TextPr;
	}

	/**
	 * Class representing a small text block called 'run'.
	 * @constructor
	 * @extends {ApiTextPr}
	 */
	function ApiRun(Run)
	{
		ApiTextPr.call(this, this, Run.Pr.Copy());
		this.Run = Run;
	}
	ApiRun.prototype = Object.create(ApiTextPr.prototype);
	ApiRun.prototype.constructor = ApiRun;

	/**
	 * Class representing a Paragraph hyperlink.
	 * @constructor
	 */
	function ApiHyperlink(ParaHyperlink)
	{
		this.ParaHyperlink		= ParaHyperlink;
	}
	ApiHyperlink.prototype.constructor = ApiHyperlink;

	/**
	 * Get the type of this class.
	 * @memberof ApiHyperlink
	 * @typeofeditors ["CDE"]
	 * @returns {"hyperlink"}
	 */
	ApiHyperlink.prototype.GetClassType = function()
	{
		return "hyperlink";
	};
	
	/**
	 * Set the hyperlink address.
	 * @typeofeditors ["CDE"]
	 * @param {string} sLink - The hyperlink address.
	 * @returns {bool} 
	 * */
	ApiHyperlink.prototype.SetLink = function(sLink)
	{
		if (typeof(sLink) !== "string" || sLink.length > Asc.c_nMaxHyperlinkLength)
			return false;
		if (sLink == undefined)
			sLink = "";

		var urlType	= undefined;

		if (sLink !== "")
		{
			urlType		= AscCommon.getUrlType(sLink);
			if (!/(((^https?)|(^ftp)):\/\/)|(^mailto:)/i.test(sLink))
				sLink = (urlType === 0) ? null :(( (urlType === 2) ? 'mailto:' : 'http://' ) + sLink);
			sLink = sLink.replace(new RegExp("%20",'g')," ");
		}
		
		this.ParaHyperlink.SetValue(sLink);
		
		return true;
	};
	/**
	 * Set the hyperlink display text.
	 * @typeofeditors ["CDE"]
	 * @param {string} sDisplay - The text to display the hyperlink.
	 * @returns {bool} 
	 * */
	ApiHyperlink.prototype.SetDisplayedText = function(sDisplay)
	{
		if (typeof(sDisplay) !== "string")
			return false;
		if (sDisplay == undefined)
			sDisplay = "";

		var HyperRun = null;
		var Styles = editor.WordControl.m_oLogicDocument.Get_Styles();

		if (this.ParaHyperlink.Content.length === 0)
		{
			HyperRun = editor.CreateRun(); 
			HyperRun.AddText(sDisplay);
			this.ParaHyperlink.Add_ToContent(0, HyperRun.Run, false);
			HyperRun.Run.Set_RStyle(Styles.GetDefaultHyperlink());
		}
		else 
		{
			HyperRun = this.GetElement(0);

			if (this.ParaHyperlink.Content.length > 1)
			{
				this.ParaHyperlink.RemoveFromContent(1, this.ParaHyperlink.Content.length - 1);
			}

			HyperRun.ClearContent();
			HyperRun.AddText(sDisplay);
		}
		
		return true;
	};
	/**
	 * Set the screen tip text of the hyperlink.
	 * @typeofeditors ["CDE"]
	 * @param {string} sScreenTipText - The screen tip text of the hyperlink.
	 * @returns {bool} 
	 * */
	ApiHyperlink.prototype.SetScreenTipText = function(sScreenTipText)
	{
		if (typeof(sScreenTipText) !== "string")
			return false;
		if (sScreenTipText == undefined)
			sScreenTipText = "";

		this.ParaHyperlink.SetToolTip(sScreenTipText);
		
		return true;
	};
	/**
	 * Get the hyperlink address.
	 * @typeofeditors ["CDE"]
	 * @returns {string} 
	 * */
	ApiHyperlink.prototype.GetLinkedText = function()
	{
		var sText = null;

		if (this.ParaHyperlink.Content.length !== 0)
		{
			sText = this.ParaHyperlink.GetValue();
		}

		return sText;
	};
	/**
	 * Get the hyperlink display text.
	 * @typeofeditors ["CDE"]
	 * @returns {string} 
	 * */
	ApiHyperlink.prototype.GetDisplayedText = function()
	{
		var oText = {Text : ""};

		if (this.ParaHyperlink.Content.length !== 0)
		{
			this.ParaHyperlink.Get_Text(oText);
		}

		return oText.Text;
	};
	/**
	 * Get the screen tip text of the hyperlink.
	 * @typeofeditors ["CDE"]
	 * @returns {string} 
	 * */
	ApiHyperlink.prototype.GetScreenTipText = function()
	{
		var sText = null;

		if (this.ParaHyperlink.Content.length !== 0)
		{
			sText = this.ParaHyperlink.GetToolTip();
		}

		return sText;
	};
	/**
	 * Get the hyperlink element using the position specified.
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {number} nPos - The position where the element which content we want to get must be located.
	 * @returns {?ParagraphContent}
	 */
	ApiHyperlink.prototype.GetElement = function(nPos)
	{
		if (nPos < 0 || nPos >= this.ParaHyperlink.Content.length)
			return null;
		
		if (this.ParaHyperlink.Content[nPos] instanceof ParaRun)
		{
			return new ApiRun(this.ParaHyperlink.Content[nPos]);
		}
	};
	/**
	 * Get the number of elements in the current hyperlink.
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {number}
	 */
	ApiHyperlink.prototype.GetElementsCount = function()
	{
		return this.ParaHyperlink.GetElementsCount();
	};
	/**
	 * Set default hyperlink style.
	 * @typeofeditors ["CDE"]
	 * @returns {bool} 
	 * */
	ApiHyperlink.prototype.SetDefaultStyle = function()
	{
		var HyperRun    = null;
		var Styles = editor.WordControl.m_oLogicDocument.Get_Styles();
		
		for (var nRun = 0; nRun < this.ParaHyperlink.Content.length; nRun++)
		{
			HyperRun = this.ParaHyperlink.Content[nRun];
			if (!(HyperRun instanceof ParaRun))
				continue;

			HyperRun.Set_Underline(undefined);
			HyperRun.Set_Color(undefined);
			HyperRun.Set_Unifill(undefined);
			HyperRun.Set_RStyle(Styles.GetDefaultHyperlink());
		}
			
		return true;
	};
	/**
	 * Return a Range object that represents the document part contained in the specified hyperlink.
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character index in the current element.
	 * @param {Number} End - End character index in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiHyperlink.prototype.GetRange = function(Start, End)
	{
		var Range = new ApiRange(this.ParaHyperlink, Start, End);

		return Range;
	};

	/**
	 * Class representing a style.
	 * @constructor
	 */
	function ApiStyle(Style)
	{
		this.Style = Style;
	}

	/**
	 * Class representing a document section.
	 * @constructor
	 */
	function ApiSection(Section)
	{
		this.Section = Section;
	}

	/**
	 * Class representing the table row properties.
	 * @constructor
	 */
	function ApiTableRowPr(Parent, RowPr)
	{
		this.Parent = Parent;
		this.RowPr  = RowPr;
	}

	/**
	 * Class representing a table row.
	 * @constructor
	 * @extends {ApiTableRowPr}
	 */
	function ApiTableRow(Row)
	{
		ApiTableRowPr.call(this, this, Row.Pr.Copy());
		this.Row = Row;
	}
	ApiTableRow.prototype = Object.create(ApiTableRowPr.prototype);
	ApiTableRow.prototype.constructor = ApiTableRow;

	/**
	 * Class representing the table cell properties.
	 * @constructor
	 */
	function ApiTableCellPr(Parent, CellPr)
	{
		this.Parent = Parent;
		this.CellPr = CellPr;
	}
	/**
	 * Class representing a table cell.
	 * @constructor
	 * @extends {ApiTableCellPr}
	 */
	function ApiTableCell(Cell)
	{
		ApiTableCellPr.call(this, this, Cell.Pr.Copy());
		this.Cell = Cell;
	}
	ApiTableCell.prototype = Object.create(ApiTableCellPr.prototype);
	ApiTableCell.prototype.constructor = ApiTableCell;

	/**
	 * Class representing the numbering properties.
	 * @constructor
	 */
	function ApiNumbering(Num)
	{
		this.Num = Num;
	}

	/**
	 * Class representing a reference to a specified level of the numbering.
	 * @constructor
	 */
	function ApiNumberingLevel(Num, Lvl)
	{
		this.Num = Num;
		this.Lvl = Math.max(0, Math.min(8, Lvl));
	}

	/**
	 * Class representing a set of formatting properties which shall be conditionally applied to the parts of a table
	 * which match the requirement specified on the <code>Type</code>.
	 * @constructor
	 */
	function ApiTableStylePr(Type, Parent, TableStylePr)
	{
		this.Type         = Type;
		this.Parent       = Parent;
		this.TableStylePr = TableStylePr;
	}

	/**
	 * Class representing an unsupported element.
	 * @constructor
	 */
	function ApiUnsupported()
	{
	}

	/**
	 * Class representing a graphical object.
	 * @constructor
	 */
	function ApiDrawing(Drawing)
	{
		this.Drawing = Drawing;
	}

	/**
	 * Class representing an image.
	 * @constructor
	 */
	function ApiImage(Image)
	{
		ApiDrawing.call(this, Image.parent);
		this.Image = Image
	}
	ApiImage.prototype = Object.create(ApiDrawing.prototype);
	ApiImage.prototype.constructor = ApiImage;

	/**
	 * Class representing a shape.
	 * @constructor
	 * */
	function ApiShape(Shape)
	{
		ApiDrawing.call(this, Shape.parent);
		this.Shape = Shape;
	}
	ApiShape.prototype = Object.create(ApiDrawing.prototype);
	ApiShape.prototype.constructor = ApiShape;

	/**
	 * Class representing a chart.
	 * @constructor
	 *
	 */
	function ApiChart(Chart)
	{
		ApiDrawing.call(this, Chart.parent);
		this.Chart = Chart;
	}
	ApiChart.prototype = Object.create(ApiDrawing.prototype);
	ApiChart.prototype.constructor = ApiChart;

	/**
	 * Class representing a base class for color types.
	 * @constructor
	 */
	function ApiUniColor(Unicolor)
	{
		this.Unicolor = Unicolor;
	}
	/**
	 * Class representing RGB Color.
	 * @constructor
	 */
	function ApiRGBColor(r, g, b)
	{
		ApiUniColor.call(this, AscFormat.CreateUniColorRGB(r, g, b));
	}
	ApiRGBColor.prototype = Object.create(ApiUniColor.prototype);
	ApiRGBColor.prototype.constructor = ApiRGBColor;

	/**
	 * Class representing a Scheme Color.
	 * @constructor
	 */
	function ApiSchemeColor(sColorId)
	{
		var oUniColor = new AscFormat.CUniColor();
		oUniColor.setColor(new AscFormat.CSchemeColor());
		switch(sColorId)
		{
			case "accent1": {  oUniColor.color.id  = 0; break;}
			case "accent2": {  oUniColor.color.id  = 1; break;}
			case "accent3": {  oUniColor.color.id  = 2; break;}
			case "accent4": {  oUniColor.color.id  = 3; break;}
			case "accent5": {  oUniColor.color.id  = 4; break;}
			case "accent6": {  oUniColor.color.id  = 5; break;}
			case "bg1": {  oUniColor.color.id      = 6; break;}
			case "bg2": {  oUniColor.color.id      = 7; break;}
			case "dk1": {  oUniColor.color.id      = 8; break;}
			case "dk2": {  oUniColor.color.id      = 9; break;}
			case "lt1": {  oUniColor.color.id      = 12; break;}
			case "lt2": {  oUniColor.color.id      = 13; break;}
			case "tx1": {  oUniColor.color.id      = 15; break;}
			case "tx2": {  oUniColor.color.id      = 16; break;}
			default: {  oUniColor.color.id      = 16; break;}
		}
		ApiUniColor.call(this, oUniColor);
	}
	ApiSchemeColor.prototype = Object.create(ApiUniColor.prototype);
	ApiSchemeColor.prototype.constructor = ApiSchemeColor;

	/**
	 * Class representing a Preset Color.
	 * @constructor
	 * */
	function ApiPresetColor(sPresetColor)
	{
		var oUniColor = new AscFormat.CUniColor();
		oUniColor.setColor(new AscFormat.CPrstColor());
		oUniColor.color.id = sPresetColor;
		ApiUniColor.call(this, oUniColor);
	}
	ApiPresetColor.prototype = Object.create(ApiUniColor.prototype);
	ApiPresetColor.prototype.constructor = ApiPresetColor;

	/**
	 * Class representing a base class fill.
	 * @constructor
	 * */
	function ApiFill(UniFill)
	{
		this.UniFill = UniFill;
	}


	/**
	 * Class representing a stroke.
	 * @constructor
	 */
	function ApiStroke(oLn)
	{
		this.Ln = oLn;
	}


	/**
	 * Class representing gradient stop.
	 * @constructor
	 * */
	function ApiGradientStop(oApiUniColor, pos)
	{
		this.Gs = new AscFormat.CGs();
		this.Gs.pos = pos;
		this.Gs.color = oApiUniColor.Unicolor;
	}

	/**
	 * Class representing a container for the paragraph elements.
	 * @constructor
	 */
	function ApiInlineLvlSdt(Sdt)
	{
		this.Sdt = Sdt;
	}

	/**
	 * Class representing a container for the document content.
	 * @constructor
	 */
	function ApiBlockLvlSdt(Sdt)
	{
		this.Sdt = Sdt;
	}

	/**
	 * Twentieths of a point (equivalent to 1/1440th of an inch).
	 * @typedef {number} twips
	 */

	/**
     * Any valid element which can be added to the document structure.
	 * @typedef {(ApiParagraph | ApiTable | ApiBlockLvlSdt)} DocumentElement
	 */

	/**
     * The style type used for the document element.
	 * @typedef {("paragraph" | "table" | "run" | "numbering")} StyleType
	 */

	/**
	 * 240ths of a line.
	 * @typedef {number} line240
	 */

	/**
	 * Half-points (2 half-points = 1 point).
	 * @typedef {number} hps
	 */

	/**
	 * A numeric value from 0 to 255.
	 * @typedef {number} byte
	 */

	/**
	 * 60000th of a degree (5400000 = 90 degrees).
	 * @typedef {number} PositiveFixedAngle
	 * */

	/**
	 * A border type which will be added to the document element.
     * * **"none"** - no border will be added to the created element or the selected element side.
     * * **"single"** - a single border will be added to the created element or the selected element side.
	 * @typedef {("none" | "single")} BorderType
	 */

	/**
	 * A shade type which can be added to the document element.
	 * @typedef {("nil" | "clear")} ShdType
	 */

	/**
	 * Custom tab types.
	 * @typedef {("clear" | "left" | "right" | "center")} TabJc
	 */

	/**
	 * Eighths of a point (24 eighths of a point = 3 points).
	 * @typedef {number} pt_8
	 */

	/**
	 * A point.
	 * @typedef {number} pt
	 */

	/**
	 * Header and footer types which can be applied to the document sections.
     * * **"default"** - a header or footer which can be applied to any default page.
     * * **"title"** - a header or footer which is applied to the title page.
     * * **"even"** - a header or footer which can be applied to even pages to distinguish them from the odd ones (which will be considered default).
	 * @typedef {("default" | "title" | "even")} HdrFtrType
	 */

	/**
	 * The possible values for the units of the width property are defined by a specific table or table cell width property.
     * * **"auto"** - set the table or table cell width to auto width.
     * * **"twips"** - set the table or table cell width to be measured in twentieths of a point.
     * * **"nul"** - set the table or table cell width to be of a zero value.
     * * **"percent"** - set the table or table cell width to be measured in percent to the parent container.
	 * @typedef {("auto" | "twips" | "nul" | "percent")} TableWidth
	 */

	/**
	 * This simple type specifies possible values for the table sections to which the current conditional formatting properties will be applied when this selected table style is used.
	 * * **"topLeftCell"** - specifies that the table formatting is applied to the top left cell.
	 * * **"topRightCell"** - specifies that the table formatting is applied to the top right cell.
	 * * **"bottomLeftCell"** - specifies that the table formatting is applied to the bottom left cell.
	 * * **"bottomRightCell"** - specifies that the table formatting is applied to the bottom right cell.
	 * * **"firstRow"** - specifies that the table formatting is applied to the first row.
	 * * **"lastRow"** - specifies that the table formatting is applied to the last row.
	 * * **"firstColumn"** - specifies that the table formatting is applied to the first column. Any subsequent row which is in *table header* ({@link ApiTableRowPr#SetTableHeader}) will also use this conditional format.
	 * * **"lastColumn"** - specifies that the table formatting is applied to the last column.
	 * * **"bandedColumn"** - specifies that the table formatting is applied to odd numbered groupings of rows.
	 * * **"bandedColumnEven"** - specifies that the table formatting is applied to even numbered groupings of rows.
	 * * **"bandedRow"** - specifies that the table formatting is applied to odd numbered groupings of columns.
	 * * **"bandedRowEven"** - specifies that the table formatting is applied to even numbered groupings of columns.
	 * * **"wholeTable"** - specifies that the conditional formatting is applied to the whole table.
	 * @typedef {("topLeftCell" | "topRightCell" | "bottomLeftCell" | "bottomRightCell" | "firstRow" | "lastRow" |
	 *     "firstColumn" | "lastColumn" | "bandedColumn" | "bandedColumnEven" | "bandedRow" | "bandedRowEven" |
	 *     "wholeTable")} TableStyleOverrideType
	 */

	/**
	 * The types of elements that can be added to the paragraph structure.
	 * @typedef {(ApiUnsupported | ApiRun | ApiInlineLvlSdt)} ParagraphContent
	 */

	/**
	 * The possible values for the base which the relative horizontal positioning of an object will be calculated from.
	 * @typedef {("character" | "column" | "leftMargin" | "rightMargin" | "margin" | "page")} RelFromH
	 */

	/**
	 * The possible values for the base which the relative vertical positioning of an object will be calculated from.
	 * @typedef {("bottomMargin" | "topMargin" | "margin" | "page" | "line" | "paragraph")} RelFromV
	 */

	/**
	 * English measure unit. 1 mm = 36000 EMUs, 1 inch = 914400 EMUs.
	 * @typedef {number} EMU
	 */

	/**
	 * This type specifies the preset shape geometry that will be used for a shape.
	 * @typedef {("accentBorderCallout1" | "accentBorderCallout2" | "accentBorderCallout3" | "accentCallout1" |
	 *     "accentCallout2" | "accentCallout3" | "actionButtonBackPrevious" | "actionButtonBeginning" |
	 *     "actionButtonBlank" | "actionButtonDocument" | "actionButtonEnd" | "actionButtonForwardNext" |
	 *     "actionButtonHelp" | "actionButtonHome" | "actionButtonInformation" | "actionButtonMovie" |
	 *     "actionButtonReturn" | "actionButtonSound" | "arc" | "bentArrow" | "bentConnector2" | "bentConnector3" |
	 *     "bentConnector4" | "bentConnector5" | "bentUpArrow" | "bevel" | "blockArc" | "borderCallout1" |
	 *     "borderCallout2" | "borderCallout3" | "bracePair" | "bracketPair" | "callout1" | "callout2" | "callout3" |
	 *     "can" | "chartPlus" | "chartStar" | "chartX" | "chevron" | "chord" | "circularArrow" | "cloud" |
	 *     "cloudCallout" | "corner" | "cornerTabs" | "cube" | "curvedConnector2" | "curvedConnector3" |
	 *     "curvedConnector4" | "curvedConnector5" | "curvedDownArrow" | "curvedLeftArrow" | "curvedRightArrow" |
	 *     "curvedUpArrow" | "decagon" | "diagStripe" | "diamond" | "dodecagon" | "donut" | "doubleWave" | "downArrow" | "downArrowCallout" | "ellipse" | "ellipseRibbon" | "ellipseRibbon2" | "flowChartAlternateProcess" | "flowChartCollate" | "flowChartConnector" | "flowChartDecision" | "flowChartDelay" | "flowChartDisplay" | "flowChartDocument" | "flowChartExtract" | "flowChartInputOutput" | "flowChartInternalStorage" | "flowChartMagneticDisk" | "flowChartMagneticDrum" | "flowChartMagneticTape" | "flowChartManualInput" | "flowChartManualOperation" | "flowChartMerge" | "flowChartMultidocument" | "flowChartOfflineStorage" | "flowChartOffpageConnector" | "flowChartOnlineStorage" | "flowChartOr" | "flowChartPredefinedProcess" | "flowChartPreparation" | "flowChartProcess" | "flowChartPunchedCard" | "flowChartPunchedTape" | "flowChartSort" | "flowChartSummingJunction" | "flowChartTerminator" | "foldedCorner" | "frame" | "funnel" | "gear6" | "gear9" | "halfFrame" | "heart" | "heptagon" | "hexagon" | "homePlate" | "horizontalScroll" | "irregularSeal1" | "irregularSeal2" | "leftArrow" | "leftArrowCallout" | "leftBrace" | "leftBracket" | "leftCircularArrow" | "leftRightArrow" | "leftRightArrowCallout" | "leftRightCircularArrow" | "leftRightRibbon" | "leftRightUpArrow" | "leftUpArrow" | "lightningBolt" | "line" | "lineInv" | "mathDivide" | "mathEqual" | "mathMinus" | "mathMultiply" | "mathNotEqual" | "mathPlus" | "moon" | "nonIsoscelesTrapezoid" | "noSmoking" | "notchedRightArrow" | "octagon" | "parallelogram" | "pentagon" | "pie" | "pieWedge" | "plaque" | "plaqueTabs" | "plus" | "quadArrow" | "quadArrowCallout" | "rect" | "ribbon" | "ribbon2" | "rightArrow" | "rightArrowCallout" | "rightBrace" | "rightBracket" | "round1Rect" | "round2DiagRect" | "round2SameRect" | "roundRect" | "rtTriangle" | "smileyFace" | "snip1Rect" | "snip2DiagRect" | "snip2SameRect" | "snipRoundRect" | "squareTabs" | "star10" | "star12" | "star16" | "star24" | "star32" | "star4" | "star5" | "star6" | "star7" | "star8" | "straightConnector1" | "stripedRightArrow" | "sun" | "swooshArrow" | "teardrop" | "trapezoid" | "triangle" | "upArrowCallout" | "upDownArrow" | "upDownArrow" | "upDownArrowCallout" | "uturnArrow" | "verticalScroll" | "wave" | "wedgeEllipseCallout" | "wedgeRectCallout" | "wedgeRoundRectCallout")} ShapeType
	 */

	/**
	 * This type specifies the available chart types which can be used to create a new chart.
	 * @typedef {("bar" | "barStacked" | "barStackedPercent" | "bar3D" | "barStacked3D" | "barStackedPercent3D" |
	 *     "barStackedPercent3DPerspective" | "horizontalBar" | "horizontalBarStacked" | "horizontalBarStackedPercent"
	 *     | "horizontalBar3D" | "horizontalBarStacked3D" | "horizontalBarStackedPercent3D" | "lineNormal" |
	 *     "lineStacked" | "lineStackedPercent" | "line3D" | "pie" | "pie3D" | "doughnut" | "scatter" | "stock" |
	 *     "area" | "areaStacked" | "areaStackedPercent")} ChartType
	 */

	/**
     * The available text vertical alignment (used to align text in a shape with a placement for text inside it).
	 * @typedef {("top" | "center" | "bottom")} VerticalTextAlign
	 * */

	/**
     * The available color scheme identifiers.
	 * @typedef {("accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6" | "bg1" | "bg2" | "dk1" | "dk2"
	 *     | "lt1" | "lt2" | "tx1" | "tx2")} SchemeColorId
	 * */

	/**
     * The available preset color names.
	 * @typedef {("aliceBlue" | "antiqueWhite" | "aqua" | "aquamarine" | "azure" | "beige" | "bisque" | "black" |
	 *     "blanchedAlmond" | "blue" | "blueViolet" | "brown" | "burlyWood" | "cadetBlue" | "chartreuse" | "chocolate"
	 *     | "coral" | "cornflowerBlue" | "cornsilk" | "crimson" | "cyan" | "darkBlue" | "darkCyan" | "darkGoldenrod" |
	 *     "darkGray" | "darkGreen" | "darkGrey" | "darkKhaki" | "darkMagenta" | "darkOliveGreen" | "darkOrange" |
	 *     "darkOrchid" | "darkRed" | "darkSalmon" | "darkSeaGreen" | "darkSlateBlue" | "darkSlateGray" |
	 *     "darkSlateGrey" | "darkTurquoise" | "darkViolet" | "deepPink" | "deepSkyBlue" | "dimGray" | "dimGrey" |
	 *     "dkBlue" | "dkCyan" | "dkGoldenrod" | "dkGray" | "dkGreen" | "dkGrey" | "dkKhaki" | "dkMagenta" |
	 *     "dkOliveGreen" | "dkOrange" | "dkOrchid" | "dkRed" | "dkSalmon" | "dkSeaGreen" | "dkSlateBlue" |
	 *     "dkSlateGray" | "dkSlateGrey" | "dkTurquoise" | "dkViolet" | "dodgerBlue" | "firebrick" | "floralWhite" |
	 *     "forestGreen" | "fuchsia" | "gainsboro" | "ghostWhite" | "gold" | "goldenrod" | "gray" | "green" |
	 *     "greenYellow" | "grey" | "honeydew" | "hotPink" | "indianRed" | "indigo" | "ivory" | "khaki" | "lavender" | "lavenderBlush" | "lawnGreen" | "lemonChiffon" | "lightBlue" | "lightCoral" | "lightCyan" | "lightGoldenrodYellow" | "lightGray" | "lightGreen" | "lightGrey" | "lightPink" | "lightSalmon" | "lightSeaGreen" | "lightSkyBlue" | "lightSlateGray" | "lightSlateGrey" | "lightSteelBlue" | "lightYellow" | "lime" | "limeGreen" | "linen" | "ltBlue" | "ltCoral" | "ltCyan" | "ltGoldenrodYellow" | "ltGray" | "ltGreen" | "ltGrey" | "ltPink" | "ltSalmon" | "ltSeaGreen" | "ltSkyBlue" | "ltSlateGray" | "ltSlateGrey" | "ltSteelBlue" | "ltYellow" | "magenta" | "maroon" | "medAquamarine" | "medBlue" | "mediumAquamarine" | "mediumBlue" | "mediumOrchid" | "mediumPurple" | "mediumSeaGreen" | "mediumSlateBlue" | "mediumSpringGreen" | "mediumTurquoise" | "mediumVioletRed" | "medOrchid" | "medPurple" | "medSeaGreen" | "medSlateBlue" | "medSpringGreen" | "medTurquoise" | "medVioletRed" | "midnightBlue" | "mintCream" | "mistyRose" | "moccasin" | "navajoWhite" | "navy" | "oldLace" | "olive" | "oliveDrab" | "orange" | "orangeRed" | "orchid" | "paleGoldenrod" | "paleGreen" | "paleTurquoise" | "paleVioletRed" | "papayaWhip" | "peachPuff" | "peru" | "pink" | "plum" | "powderBlue" | "purple" | "red" | "rosyBrown" | "royalBlue" | "saddleBrown" | "salmon" | "sandyBrown" | "seaGreen" | "seaShell" | "sienna" | "silver" | "skyBlue" | "slateBlue" | "slateGray" | "slateGrey" | "snow" | "springGreen" | "steelBlue" | "tan" | "teal" | "thistle" | "tomato" | "turquoise" | "violet" | "wheat" | "white" | "whiteSmoke" | "yellow" | "yellowGreen")} PresetColor
	 * */


	/**
     * Possible values for the position of chart tick labels (either horizontal or vertical).
     * * **"none"** - not display the selected tick labels.
     * * **"nextTo"** - set the position of the selected tick labels next to the main label.
     * * **"low"** - set the position of the selected tick labels in the part of the chart with lower values.
     * * **"high"** - set the position of the selected tick labels in the part of the chart with higher values.
	 * @typedef {("none" | "nextTo" | "low" | "high")} TickLabelPosition
	 * **/

	/**
     * The type of a fill which uses an image as a background.
     * * **"tile"** - if the image is smaller than the shaped which is filled, the image will be tiled all over the created shape surface.
     * * **"stretch"** - if the image is smaller than the shape which is filled, the image will be stretched to fit the created shape surface.
	 * @typedef {"tile" | "stretch"} BlipFillType
	 * */

	/**
     * The available preset patterns which can be used for the fill.
	 * @typedef {"cross" | "dashDnDiag" | "dashHorz" | "dashUpDiag" | "dashVert" | "diagBrick" | "diagCross" | "divot"
	 *     | "dkDnDiag" | "dkHorz" | "dkUpDiag" | "dkVert" | "dnDiag" | "dotDmnd" | "dotGrid" | "horz" | "horzBrick" |
	 *     "lgCheck" | "lgConfetti" | "lgGrid" | "ltDnDiag" | "ltHorz" | "ltUpDiag" | "ltVert" | "narHorz" | "narVert"
	 *     | "openDmnd" | "pct10" | "pct20" | "pct25" | "pct30" | "pct40" | "pct5" | "pct50" | "pct60" | "pct70" |
	 *     "pct75" | "pct80" | "pct90" | "plaid" | "shingle" | "smCheck" | "smConfetti" | "smGrid" | "solidDmnd" |
	 *     "sphere" | "trellis" | "upDiag" | "vert" | "wave" | "wdDnDiag" | "wdUpDiag" | "weave" | "zigZag"}
	 *     PatternType
	 * */

	/**
	 *
	 * @typedef {"unlocked" | "contentLocked" | "sdtContentLocked" | "sdtLocked"} SdtLock
	 */
	//------------------------------------------------------------------------------------------------------------------
	//
	// Base Api
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
     * The 1000th of a percent (100000 = 100%).
	 * @typedef {number} PositivePercentage
	 * */

	/**
	 * @typedef {("cross" | "in" | "none" | "out")} TickMark
	 * */

	/**
	 * Get the main document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @returns {ApiDocument}
	 */
	Api.prototype.GetDocument = function()
	{
		return new ApiDocument(this.WordControl.m_oLogicDocument);
	};
	/**
	 * Create a new paragraph.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {ApiParagraph}
	 */
	Api.prototype.CreateParagraph = function()
	{
		return new ApiParagraph(new Paragraph(private_GetDrawingDocument(), private_GetLogicDocument()));
	};
	/**
	 * Create an element range.
	 * If you do not specify the start and end positions, the range will be taken from the entire element.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param oElement - The element from which the range will be taken.
	 * @param nStart - Start range position.
	 * @param nEnd - End range position.
	 * @returns {ApiRange | null} - returns null if oElement isn't supported.
	 */
	Api.prototype.CreateRange = function(oElement, nStart, nEnd)
	{
		if (oElement) 
		{
			switch (oElement.GetClassType())
			{
				case 'paragraph':
				case 'hyperlink':
				case 'run':
				case 'table':
				case 'documentContent':
				case 'document':
				case 'inlineLvlSdt':
				case 'blockLvlSdt':
					return oElement.GetRange(nStart, nEnd);
				default:
					return null;
			}
		}
		return null;
	};
	/**
	 * Create a new table with a specified number of rows and columns.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {number} nCols - Number of columns.
	 * @param {number} nRows - Number of rows.
	 * @returns {ApiTable}
	 */
	Api.prototype.CreateTable = function(nCols, nRows)
	{
		if (!nRows || nRows <= 0 || !nCols || nCols <= 0)
			return null;

		var oTable = new CTable(private_GetDrawingDocument(), private_GetLogicDocument(), true, nRows, nCols, [], false);
		oTable.CorrectBadGrid();
		oTable.Set_TableW(undefined);
		oTable.Set_TableStyle2(undefined);
		return new ApiTable(oTable);
	};
	/**
	 * Create a new smaller text block to be inserted to the current paragraph or table.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiRun}
	 */
	Api.prototype.CreateRun = function()
	{
		return new ApiRun(new ParaRun(null, false));
	};
	/**
	 * Create a new hyperlink text block to be inserted to the current paragraph or table.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {string} sLink - The hyperlink address. 
	 * @param {string} sDisplay - The text to display the hyperlink.
	 * @param {string} sScreenTipText - The screen tip text.
	 * @returns {ApiHyperlink}
	 */
	Api.prototype.CreateHyperlink = function(sLink, sDisplay, sScreenTipText)
	{
		// Создаем гиперссылку
		var oHyperlink		= new ParaHyperlink();
		var apiHyperlink	= new ApiHyperlink(oHyperlink);

		apiHyperlink.SetLink(sLink);
		apiHyperlink.SetDisplayedText(sDisplay);
		apiHyperlink.SetScreenTipText(sScreenTipText);
		
		return apiHyperlink;
	};

	/**
	 * Create an image with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {string} sImageSrc - The image source where the image to be inserted should be taken from (currently only internet URL or Base64 encoded images are supported).
	 * @param {EMU} nWidth - The image width in English measure units.
	 * @param {EMU} nHeight - The image height in English measure units.
	 * @returns {ApiImage}
	 */
	Api.prototype.CreateImage = function(sImageSrc, nWidth, nHeight)
	{
		var nW = private_EMU2MM(nWidth);
		var nH = private_EMU2MM(nHeight);

		var oDrawing = new ParaDrawing(nW, nH, null, private_GetDrawingDocument(), private_GetLogicDocument(), null);
		var oImage = private_GetLogicDocument().DrawingObjects.createImage(sImageSrc, 0, 0, nW, nH);
		oImage.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oImage);
		return new ApiImage(oImage);
	};

	/**
	 * Create a shape with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {ShapeType} [sType="rect"] - The shape type which specifies the preset shape geometry.
	 * @param {EMU} [nWidth = 914400] - The shape width in English measure units.
	 * @param {EMU} [nHeight = 914400] - The shape height in English measure units.
	 * @param {ApiFill} [oFill = Api.CreateNoFill()] - The color or pattern used to fill the shape.
	 * @param {ApiStroke} [oStroke = Api.CreateStroke(0, Api.CreateNoFill())] - The stroke used to create the element shadow.
	 * @returns {ApiShape}
	 * */
	Api.prototype.CreateShape = function(sType, nWidth, nHeight, oFill, oStroke)
	{
		var oLogicDocument = private_GetLogicDocument();
		var oDrawingDocuemnt = private_GetDrawingDocument();
		sType   = sType   || "rect";
        nWidth  = nWidth  || 914400;
	    nHeight = nHeight || 914400;
		oFill   = oFill   || editor.CreateNoFill();
		oStroke = oStroke || editor.CreateStroke(0, editor.CreateNoFill());
		var nW = private_EMU2MM(nWidth);
		var nH = private_EMU2MM(nHeight);
		var oDrawing = new ParaDrawing(nW, nH, null, oDrawingDocuemnt, oLogicDocument, null);
		var oShapeTrack = new AscFormat.NewShapeTrack(sType, 0, 0, oLogicDocument.theme, null, null, null, 0);
		oShapeTrack.track({}, nW, nH);
		var oShape = oShapeTrack.getShape(true, oDrawingDocuemnt, null);
		oShape.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oShape);
		oShape.createTextBoxContent();
		oShape.spPr.setFill(oFill.UniFill);
		oShape.spPr.setLn(oStroke.Ln);
		return new ApiShape(oShape);
	};

	/**
	 * Create a chart with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {ChartType} [sType="bar"] - The chart type used for the chart display.
	 * @param {Array} aSeries - The array of the data used to build the chart from.
	 * @param {Array} aSeriesNames - The array of the names (the source table column names) used for the data which the chart will be build from.
	 * @param {Array} aCatNames - The array of the names (the source table row names) used for the data which the chart will be build from.
	 * @param {EMU} nWidth - The chart width in English measure units.
	 * @param {EMU} nHeight - The chart height in English measure units.
	 * @param {number} nStyleIndex - The chart color style index (can be 1 - 48, as described in OOXML specification).
	 * @returns {ApiChart}
	 * */
	Api.prototype.CreateChart = function(sType, aSeries, aSeriesNames, aCatNames, nWidth, nHeight, nStyleIndex)
	{
		var oDrawingDocument = private_GetDrawingDocument();
		var nW = private_EMU2MM(nWidth);
		var nH = private_EMU2MM(nHeight);
		var oDrawing = new ParaDrawing( nW, nH, null, oDrawingDocument, null, null);
		var oChartSpace = AscFormat.builder_CreateChart(nW, nH, sType, aCatNames, aSeriesNames, aSeries, nStyleIndex);
		if(!oChartSpace)
		{
			return null;
		}
		oChartSpace.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oChartSpace);
		oDrawing.setExtent( oChartSpace.spPr.xfrm.extX, oChartSpace.spPr.xfrm.extY );
		return new ApiChart(oChartSpace);
	};

	/**
	 * Create an RGB color setting the appropriate values for the red, green and blue color components.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @returns {ApiRGBColor}
	 */
	Api.prototype.CreateRGBColor = function(r, g, b)
	{
		return new ApiRGBColor(r, g, b);
	};

	/**
	 * Create a complex color scheme selecting from one of the available schemes.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {SchemeColorId} sSchemeColorId - The color scheme identifier.
	 * @returns {ApiSchemeColor}
	 */
	Api.prototype.CreateSchemeColor = function(sSchemeColorId)
	{
		return new ApiSchemeColor(sSchemeColorId);
	};

	/**
	 * Create a color selecting it from one of the available color presets.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {PresetColor} sPresetColor - A preset selected from the list of the available color preset names.
	 * @returns {ApiPresetColor};
	 * */
	Api.prototype.CreatePresetColor = function(sPresetColor)
	{
		return new ApiPresetColor(sPresetColor);
	};

	/**
	 * Create a solid fill to apply to the object using a selected solid color as the object background.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ApiUniColor} oUniColor - The color used for the element fill.
	 * @returns {ApiFill}
	 * */
	Api.prototype.CreateSolidFill = function(oUniColor)
	{
		return new ApiFill(AscFormat.CreateUniFillByUniColorCopy(oUniColor.Unicolor));
	};

	/**
	 * Create a linear gradient fill to apply to the object using a selected linear gradient as the object background.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {Array} aGradientStop - The array of gradient color stops measured in 1000th of percent.
	 * @param {PositiveFixedAngle} Angle - The angle measured in 60000th of a degree that will define the gradient direction.
	 * @returns {ApiFill}
	 */
	Api.prototype.CreateLinearGradientFill = function(aGradientStop, Angle)
	{
		return new ApiFill(AscFormat.builder_CreateLinearGradient(aGradientStop, Angle));
	};


	/**
	 * Create a radial gradient fill to apply to the object using a selected radial gradient as the object background.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {Array} aGradientStop - The array of gradient color stops measured in 1000th of percent.
	 * @returns {ApiFill}
	 */
	Api.prototype.CreateRadialGradientFill = function(aGradientStop)
	{
		return new ApiFill(AscFormat.builder_CreateRadialGradient(aGradientStop));
	};
	/**
	 * Create a pattern fill to apply to fill the object using a selected pattern as the object background.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {PatternType} sPatternType - The pattern type used for the fill selected from one of the available pattern types.
	 * @param {ApiUniColor} BgColor - The background color used for the pattern creation.
	 * @param {ApiUniColor} FgColor - The foreground color used for the pattern creation.
	 * @returns {ApiFill}
	 */
	Api.prototype.CreatePatternFill = function(sPatternType, BgColor, FgColor)
	{
		return new ApiFill(AscFormat.builder_CreatePatternFill(sPatternType, BgColor, FgColor));
	};

	/**
	 * Create a blip fill to apply to the object using a selected image as the object background.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} sImageUrl - The path to the image used for the blip fill (currently only internet URL or Base64 encoded images are supported).
	 * @param {BlipFillType} sBlipFillType - The type of the fill used for the blip fill (tile or stretch).
	 * @returns {ApiFill}
	 * */
	Api.prototype.CreateBlipFill = function(sImageUrl, sBlipFillType)
	{
		return new ApiFill(AscFormat.builder_CreateBlipFill(sImageUrl, sBlipFillType));
	};

	/**
	 * Create no fill and remove the fill from the element.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiFill}
	 * */
	Api.prototype.CreateNoFill = function()
	{
		return new ApiFill(AscFormat.CreateNoFillUniFill());
	};

	/**
	 * Create a stroke adding shadows to the element.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {EMU} nWidth - The width of the shadow measured in English measure units.
	 * @param {ApiFill} oFill - The fill type used to create the shadow.
	 * @returns {ApiStroke}
	 * */
	Api.prototype.CreateStroke = function(nWidth, oFill)
	{
		return new ApiStroke(AscFormat.builder_CreateLine(nWidth, oFill));
	};

	/**
	 * Create a gradient stop used for different types of gradients.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ApiUniColor} oUniColor - The color used for the gradient stop.
	 * @param {PositivePercentage} nPos - The position of the gradient stop measured in 1000th of percent.
	 * @returns {ApiGradientStop}
	 * */
	Api.prototype.CreateGradientStop = function(oUniColor, nPos)
	{
		return new ApiGradientStop(oUniColor, nPos);
	};

	/**
	 * Create a bullet for a paragraph with the character or symbol specified with the sBullet parameter.
	 * @memberof Api
	 * @typeofeditors ["CSE", "CPE"]
	 * @param {string} sSymbol - The character or symbol which will be used to create the bullet for the paragraph.
	 * @returns {ApiBullet}
	 * */
	Api.prototype.CreateBullet = function(sSymbol){
		var oBullet = new AscFormat.CBullet();
		oBullet.bulletType = new AscFormat.CBulletType();
		if(typeof sSymbol === "string" && sSymbol.length > 0){
			oBullet.bulletType.type = AscFormat.BULLET_TYPE_BULLET_CHAR;
			oBullet.bulletType.Char = sSymbol[0];
		}
		else{
			oBullet.bulletType.type = AscFormat.BULLET_TYPE_BULLET_NONE;
		}
		return new ApiBullet(oBullet);
	};

	/**
	 * Create a bullet for a paragraph with the character or symbol specified with the sType parameter.
	 * @memberof Api
	 * @typeofeditors ["CSE", "CPE"]
	 * @param {BulletType} sType - The numbering type the paragraphs will be numbered with.
	 * @param {number} nStartAt - The number the first numbered paragraph will start with.
	 * @returns {ApiBullet}
	 * */

	Api.prototype.CreateNumbering = function(sType, nStartAt){
		var oBullet = new AscFormat.CBullet();
		oBullet.bulletType = new AscFormat.CBulletType();
		oBullet.bulletType.type = AscFormat.BULLET_TYPE_BULLET_AUTONUM;
		switch(sType){
			case "ArabicPeriod" :{
				oBullet.bulletType.AutoNumType = 12;
				break;
			}
			case "ArabicParenR":{
				oBullet.bulletType.AutoNumType = 11;
				break;
			}
			case "RomanUcPeriod":{
				oBullet.bulletType.AutoNumType = 34;
				break;
			}
			case "RomanLcPeriod":{
				oBullet.bulletType.AutoNumType = 31;
				break;
			}
			case "AlphaLcParenR":{
				oBullet.bulletType.AutoNumType = 1;
				break;
			}
			case "AlphaLcPeriod":{
				oBullet.bulletType.AutoNumType = 2;
				break;
			}
			case "AlphaUcParenR":{
				oBullet.bulletType.AutoNumType = 4;
				break;
			}
			case "AlphaUcPeriod":{
				oBullet.bulletType.AutoNumType = 5;
				break;
			}
			case "None":{
				oBullet.bulletType.type = AscFormat.BULLET_TYPE_BULLET_NONE;
				break;
			}
		}
		if( oBullet.bulletType.type === AscFormat.BULLET_TYPE_BULLET_AUTONUM){
			if(AscFormat.isRealNumber(nStartAt)){
				oBullet.bulletType.startAt = nStartAt;
			}
		}
		return new ApiBullet(oBullet);
	};

	/**
	 * Create a new inline container.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @returns {ApiInlineLvlSdt}
	 */
	Api.prototype.CreateInlineLvlSdt = function()
	{
		var oSdt = new CInlineLevelSdt();
		oSdt.Add_ToContent(0, new ParaRun(null, false));
		return new ApiInlineLvlSdt(oSdt);
	};

	/**
	 * Create a new block level container.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @returns {ApiBlockLvlSdt}
	 */
	Api.prototype.CreateBlockLvlSdt = function()
	{
		return new ApiBlockLvlSdt(new CBlockLevelSdt(editor.private_GetLogicDocument()));
	};

	/**
	 * Save changes to the specified document.
	 * @typeofeditors ["CDE"]
	 * @memberof Api
	 */
	Api.prototype.Save = function () {
		this.SaveAfterMacros = true;
	};

	/**
	 * Load data for the mail merge. 
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {String[][]} aList - Mail merge data. The first element of the array is the array with names of the merge fields.
	 * The rest of the array elements are arrays with values for the merge fields.
	 * @typeofeditors ["CDE"]
	 * @return {bool}  
	 */
	Api.prototype.LoadMailMergeData = function(aList)
	{
		if (!aList || aList.length === 0)
			return false;

		editor.asc_StartMailMergeByList(aList);

		return true;
	};

	/**
	 * Get the mail merge template document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @return {ApiDocumentContent}  
	 */
	Api.prototype.GetMailMergeTemplateDocContent = function()
	{
		var oDocument = editor.private_GetLogicDocument();

		AscCommon.History.TurnOff();
		AscCommon.g_oTableId.TurnOff();

		var LogicDocument = new CDocument(undefined, false);
		AscCommon.History.Document = oDocument;

		// Копируем стили, они все одинаковые для всех документов
		LogicDocument.Styles = oDocument.Styles.Copy();

		// Нумерацию придется повторить для каждого отдельного файла
		LogicDocument.Numbering.Clear();

		LogicDocument.DrawingDocument = oDocument.DrawingDocument;

		LogicDocument.theme = oDocument.theme.createDuplicate();
		LogicDocument.clrSchemeMap   = oDocument.clrSchemeMap.createDuplicate();

		var FieldsManager = oDocument.FieldsManager;

		var ContentCount = oDocument.Content.length;
		var OverallIndex = 0;
		oDocument.ForceCopySectPr = true;

		// Подменяем ссылку на менеджер полей, чтобы скопированные поля регистрировались в новом классе
		oDocument.FieldsManager = LogicDocument.FieldsManager;
		var NewNumbering = oDocument.Numbering.CopyAllNums(LogicDocument.Numbering);

		LogicDocument.Numbering.AppendAbstractNums(NewNumbering.AbstractNum);
		LogicDocument.Numbering.AppendNums(NewNumbering.Num);

		oDocument.CopyNumberingMap = NewNumbering.NumMap;

		for (var ContentIndex = 0; ContentIndex < ContentCount; ContentIndex++)
		{
			LogicDocument.Content[OverallIndex++] = oDocument.Content[ContentIndex].Copy(LogicDocument, oDocument.DrawingDocument);

			if (type_Paragraph === oDocument.Content[ContentIndex].Get_Type())
			{
				var ParaSectPr = oDocument.Content[ContentIndex].Get_SectionPr();
				if (ParaSectPr)
				{
					var NewParaSectPr = new CSectionPr();
					NewParaSectPr.Copy(ParaSectPr, true);
					LogicDocument.Content[OverallIndex - 1].Set_SectionPr(NewParaSectPr, false);
				}
			}
		}

		oDocument.CopyNumberingMap = null;
		oDocument.ForceCopySectPr  = false;

		for (var Index = 0, Count = LogicDocument.Content.length; Index < Count; Index++)
		{
			if (0 === Index)
				LogicDocument.Content[Index].Prev = null;
			else
				LogicDocument.Content[Index].Prev = LogicDocument.Content[Index - 1];

			if (Count - 1 === Index)
				LogicDocument.Content[Index].Next = null;
			else
				LogicDocument.Content[Index].Next = LogicDocument.Content[Index + 1];

			LogicDocument.Content[Index].Parent = LogicDocument;
		}

		oDocument.FieldsManager = FieldsManager;
		AscCommon.g_oTableId.TurnOn();
		AscCommon.History.TurnOn();

		return new ApiDocumentContent(LogicDocument);
	};

	/**
	 * Get the mail merge receptions count.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @return {number}  
	 */
	Api.prototype.GetMailMergeReceptionsCount = function()
	{
		var oDocument = editor.private_GetLogicDocument();

		return oDocument.Get_MailMergeReceptionsCount();
	};

	/**
	 * Replace the main document content with another document content.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {ApiDocumentContent} - The document content which the main document content will be replaced with.
	 */
	Api.prototype.ReplaceDocumentContent = function(oApiDocumentContent)
	{
		var oDocument        = editor.private_GetLogicDocument();
		var mailMergeContent = oApiDocumentContent.Document.Content;
		oDocument.Remove_FromContent(0, oDocument.Content.length);

		for (var nElement = 0; nElement < mailMergeContent.length; nElement++)
			oDocument.Add_ToContent(oDocument.Content.length, mailMergeContent[nElement].Copy(oDocument, oDocument.DrawingDocument), false);

		oDocument.Remove_FromContent(0, 1);
	};

	/**
	 * Start the mail merge process.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {number} nStartIndex - The start index of the document for mail merge process.
	 * @param {number} nEndIndex - The end index of the document for mail merge process.
	 * @param {bool} bAll - if true -> be mail merge all recipients 
	 * @returns {bool}
	 */
	Api.prototype.MailMerge = function(nStartIndex, nEndIndex)
	{
		var oDocument        = editor.private_GetLogicDocument();
		var mailMergeDoc     = null;

		var _nStartIndex = (undefined !== nStartIndex ? Math.max(0, nStartIndex) : 0);
		var _nEndIndex   = (undefined !== nEndIndex   ? Math.min(nEndIndex, oDocument.MailMergeMap.length - 1) : oDocument.MailMergeMap.length - 1);

		mailMergeDoc = oDocument.Get_MailMergedDocument(_nStartIndex, _nEndIndex);

		if (!mailMergeDoc)
			return false;
		
		this.ReplaceDocumentContent(new ApiDocumentContent(mailMergeDoc));

		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiUnsupported
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @typeofeditors ["CDE"]
	 * @returns {"unsupported"}
	 */
	ApiUnsupported.prototype.GetClassType = function()
	{
		return "unsupported";
	};
	/**
	 * Add a comment to the desired element or array of Runs.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {ApiRun[] | ApiParagraph | ApiDocument} oElement - May be Document, Paragraph or Run[].
	 * @param {string} Comment - The comment text.
	 * @param {string} Autor - The author's name (not obligatory).
	 * @returns {bool} - returns false if params are invalid.
	 */
	Api.prototype.AddComment = function(oElement, Comment, Autor)
	{
		if (!Comment || typeof(Comment) !== "string")
			return false;
	
		if (typeof(Autor) !== "string")
			Autor = "";
		
		// Если oElement не является массивом, определяем параграф это или документ
		if (!Array.isArray(oElement))
		{
			// Проверка на параграф
			if (oElement instanceof ApiParagraph | oElement instanceof ApiDocument)
				return oElement.AddComment(Comment, Autor);
		}
		// Проверка на массив с ранами
		else if (Array.isArray(oElement))
		{
			// Если хотя бы один элемент массива не является раном, или хотя бы один ран не принадлежит 
			// ни одному параграфу - не добавляем комментарий
			for (var Index = 0; Index < oElement.length; Index++)
			{
				if (!(oElement[Index] instanceof ApiRun))
					return false;					
			}
			
			// Если раны из принципиально разных контекcтов (из тела и хедера(или футера) то комментарий не добавляем)
			for (var Index = 1; Index < oElement.length; Index++)
			{
				if (oElement[0].Run.GetDocumentPositionFromObject()[0].Class !== oElement[Index].Run.GetDocumentPositionFromObject()[0].Class)
					return false;
			}
			
			var CommentData = new AscCommon.CCommentData();
			CommentData.SetText(Comment);
			CommentData.SetUserName(Autor);

			var oDocument = private_GetLogicDocument();
			
			var oStartRun = private_GetFirstRunInArray(oElement); 
			var StartPos = oStartRun.Run.GetDocumentPositionFromObject();

			var oEndRun	= private_GetLastRunInArray(oElement)
			var EndPos	= oEndRun.Run.GetDocumentPositionFromObject();

			StartPos.push({Class: oStartRun.Run, Position: 0});
			EndPos.push({Class: oEndRun.Run, Position: oEndRun.Run.Content.length});

			oDocument.Selection.Use = true;
			oDocument.SetContentSelection(StartPos, EndPos, 0, 0, 0);
			
			var COMENT = oDocument.AddComment(CommentData, false);
			oDocument.RemoveSelection();
			
			if (null != COMENT)
			{
				editor.sync_AddComment(COMENT.Get_Id(), CommentData);
			}

			return true;
		}
	};

	

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDocumentContent
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of the current class. 
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"documentContent"}
	 */
	ApiDocumentContent.prototype.GetClassType = function()
	{
		return "documentContent";
	};
	/**
	 * Get the number of elements in the current document.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {number}
	 */
	ApiDocumentContent.prototype.GetElementsCount = function()
	{
		return this.Document.Content.length;
	};
	/**
	 * Get the element by its position in the document.
	 * @memberof ApiDocumentContent
	 * @param {number} nPos - The element position that will be taken from the document.
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {?DocumentElement}
	 */
	ApiDocumentContent.prototype.GetElement = function(nPos)
	{
		if (!this.Document.Content[nPos])
			return null;

		var Type = this.Document.Content[nPos].GetType();
		if (type_Paragraph === Type)
			return new ApiParagraph(this.Document.Content[nPos]);
		else if (type_Table === Type)
			return new ApiTable(this.Document.Content[nPos]);
		else if (type_BlockLevelSdt === Type)
			return new ApiBlockLvlSdt(this.Document.Content[nPos]);

		return null;
	};
	/**
	 * Add a paragraph or a table or a blockLvl content control using its position in the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {number} nPos - The position where the current element will be added.
	 * @param {DocumentElement} oElement - The document element which will be added at the current position.
	 */
	ApiDocumentContent.prototype.AddElement = function(nPos, oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			this.Document.Internal_Content_Add(nPos, oElement.private_GetImpl());
		}
	};
	/**
	 * Push a paragraph or a table to actually add it to the document.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {DocumentElement} oElement - The element type which will be pushed to the document.
	 * @returns {bool} - returns false if oElement is unsupported.
	 */
	ApiDocumentContent.prototype.Push = function(oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			this.Document.Internal_Content_Add(this.Document.Content.length, oElement.private_GetImpl());
			return true;
		}

		return false;
	};
	/**
	 * Remove all the elements from the current document or from the current document element.
	 * <note>When all elements are removed, a new empty paragraph is automatically created. If you want to add
	 * content to this paragraph, use the {@link ApiDocumentContent#GetElement} method.</note>
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiDocumentContent.prototype.RemoveAllElements = function()
	{
		this.Document.Internal_Content_Remove(0, this.Document.Content.length, true);
	};
	/**
	 * Remove element using the position specified.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {number} nPos - The element number (position) in the document or inside other element.
	 */
	ApiDocumentContent.prototype.RemoveElement = function(nPos)
	{
		if (nPos < 0 || nPos >= this.GetElementsCount())
			return;

		this.Document.Internal_Content_Remove(nPos, 1, true);
	};
	/**
	 * Return a Range object that represents the part of the document contained in the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiDocumentContent.prototype.GetRange = function(Start, End)
	{
		var Range = new ApiRange(this.Document, Start, End);

		return Range;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDocument
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {"document"}
	 */
	ApiDocument.prototype.GetClassType = function()
	{
		return "document";
	};
	/**
	 * Create new history point.
	 * @memberof ApiDocument
	 */
	ApiDocument.prototype.CreateNewHistoryPoint = function()
	{
		this.Document.Create_NewHistoryPoint(AscDFH.historydescription_Document_ApiBuilder);
	};
	/**
	 * Get a style by the style name.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sStyleName - The style name.
	 * @returns {?ApiStyle}
	 */
	ApiDocument.prototype.GetStyle = function(sStyleName)
	{
		var oStyles  = this.Document.Get_Styles();
		var oStyleId = oStyles.GetStyleIdByName(sStyleName, true);
		return new ApiStyle(oStyles.Get(oStyleId));
	};
	/**
	 * Create a new style with the specified type and name. If there is a style with the same name it will be replaced with a new one.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sStyleName - The name of the style which will be created.
	 * @param {StyleType} [sType="paragraph"] - The document element which the style will be applied to.
	 * @returns {ApiStyle}
	 */
	ApiDocument.prototype.CreateStyle = function(sStyleName, sType)
	{
		var nStyleType = styletype_Paragraph;
		if ("paragraph" === sType)
			nStyleType = styletype_Paragraph;
		else if ("table" === sType)
			nStyleType = styletype_Table;
		else if ("run" === sType)
			nStyleType = styletype_Character;
		else if ("numbering" === sType)
			nStyleType = styletype_Numbering;

		var oStyle        = new CStyle(sStyleName, null, null, nStyleType, false);
		oStyle.qFormat    = true;
		oStyle.uiPriority = 1;
		var oStyles       = this.Document.Get_Styles();

		// Если у нас есть стиль с данным именем, тогда мы старый стиль удаляем, а новый добавляем со старым Id,
		// чтобы если были ссылки на старый стиль - теперь они стали на новый.
		var sOldId    = oStyles.GetStyleIdByName(sStyleName);
		var oOldStyle = oStyles.Get(sOldId);
		if (null != sOldId && oOldStyle)
		{
			oStyles.Remove(sOldId);
			oStyles.RemapIdReferences(sOldId, oStyle.Get_Id());
		}

		oStyles.Add(oStyle);
		return new ApiStyle(oStyle);
	};
	/**
	 * Get the default style parameters for the specified document element.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {StyleType} sStyleType - The document element which we want to get the style for.
	 * @returns {?ApiStyle}
	 */
	ApiDocument.prototype.GetDefaultStyle = function(sStyleType)
	{
		var oStyles = this.Document.Get_Styles();

		if ("paragraph" === sStyleType)
			return new ApiStyle(oStyles.Get(oStyles.Get_Default_Paragraph()));
		else if ("table" === sStyleType)
			return new ApiStyle(oStyles.Get(oStyles.Get_Default_Table()));
		else if ("run" === sStyleType)
			return new ApiStyle(oStyles.Get(oStyles.Get_Default_Character()));
		else if ("numbering" === sStyleType)
			return new ApiStyle(oStyles.Get(oStyles.Get_Default_Numbering()));

		return null;
	};
	/**
	 * Get a set of default properties for the text run in the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTextPr}
	 */
	ApiDocument.prototype.GetDefaultTextPr = function()
	{
		var oStyles = this.Document.Get_Styles();
		return new ApiTextPr(this, oStyles.Get_DefaultTextPr().Copy());
	};
	/**
	 * Get a set of default paragraph properties in the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParaPr}
	 */
	ApiDocument.prototype.GetDefaultParaPr = function()
	{
		var oStyles = this.Document.Get_Styles();
		return new ApiParaPr(this, oStyles.Get_DefaultParaPr().Copy());
	};
	/**
	 * Get the document final section.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @return {ApiSection}
	 */
	ApiDocument.prototype.GetFinalSection = function()
	{
		return new ApiSection(this.Document.SectPr);
	};
	/**
	 * Create a new document section which ends at the specified paragraph. Allows to set local parameters for the current
	 * section - page size, footer, header, columns, etc.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {ApiParagraph} oParagraph - The paragraph after which the new document section will be inserted.
	 * @returns {ApiSection}
	 */
	ApiDocument.prototype.CreateSection = function(oParagraph)
	{
		if (!(oParagraph instanceof ApiParagraph))
			return null;

		var oSectPr = new CSectionPr(this.Document);
		oParagraph.private_GetImpl().Set_SectionPr(oSectPr);
		return new ApiSection(oSectPr);
	};

	/**
	 * Specify whether sections in this document will have different headers and footers for even and
	 * odd pages (one header/footer for odd pages and another header/footer for even pages).
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isEvenAndOdd - If true the header/footer will be different for odd and even pages, if false they will be the same.
	 */
	ApiDocument.prototype.SetEvenAndOddHdrFtr = function(isEvenAndOdd)
	{
		this.Document.Set_DocumentEvenAndOddHeaders(isEvenAndOdd);
	};
	/**
	 * Create an abstract multilevel numbering with a specified type.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {("bullet" | "numbered")} [sType="bullet"] - The type of the numbering which will be created.
	 * @returns {ApiNumbering}
	 */
	ApiDocument.prototype.CreateNumbering = function(sType)
	{
		var oGlobalNumbering = this.Document.GetNumbering();
		var oNum             = oGlobalNumbering.CreateNum();

		if ("numbered" === sType)
			oNum.CreateDefault(c_oAscMultiLevelNumbering.Numbered);
		else
			oNum.CreateDefault(c_oAscMultiLevelNumbering.Bullet);

		return new ApiNumbering(oNum);
	};

	/**
	 * Insert an array of elements into the current position of the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {DocumentElement[]} arrContent - An array of elements to insert.
	 * @param {boolean} [isInline=false] - Inline insert or not (works only for the last and the first element and only if it's a paragraph).
	 * @param {object} [oPr=undefined] - Specifies that text and paragraph document properties are preserved for the inserted elements. 
	 * The object should look like this: {"KeepTextOnly": true}. 
	 * @returns {boolean} Success?
	 */
	ApiDocument.prototype.InsertContent = function(arrContent, isInline, oPr)
	{
		var oSelectedContent = new CSelectedContent();
		for (var nIndex = 0, nCount = arrContent.length; nIndex < nCount; ++nIndex)
		{
			var oElement = arrContent[nIndex];
			if (oElement instanceof ApiParagraph || oElement instanceof ApiTable)
			{
				if (true === isInline && oElement instanceof ApiParagraph)
					oSelectedContent.Add(new CSelectedElement(oElement.private_GetImpl(), false));
				else
					oSelectedContent.Add(new CSelectedElement(oElement.private_GetImpl(), true));
			}
		}
		oSelectedContent.On_EndCollectElements(this.Document, true);

		if (this.Document.IsSelectionUse())
		{
			this.Document.Start_SilentMode();
			this.Document.Remove(1, false, false, isInline);
			this.Document.End_SilentMode();
			this.Document.RemoveSelection(true);
		}

		var oParagraph = this.Document.GetCurrentParagraph(undefined, undefined, {CheckDocContent: true});
		if (!oParagraph)
			return;

		var oNearestPos = {
			Paragraph  : oParagraph,
			ContentPos : oParagraph.Get_ParaContentPos(false, false)
		};

		if (oPr)
		{
			if (oPr["KeepTextOnly"])
			{
				var oParaPr = this.Document.GetDirectParaPr();
				var oTextPr = this.Document.GetDirectTextPr();

				for (var nIndex = 0, nCount = oSelectedContent.Elements.length; nIndex < nCount; ++nIndex)
				{
					var oElement = oSelectedContent.Elements[nIndex].Element;
					var arrParagraphs = oElement.GetAllParagraphs();
					for (var nParaIndex = 0, nParasCount = arrParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
					{
						arrParagraphs[nParaIndex].SetDirectParaPr(oParaPr);
						arrParagraphs[nParaIndex].SetDirectTextPr(oTextPr);
					}
				}
			}
		}

		oParagraph.Check_NearestPos(oNearestPos);

		if (!this.Document.Can_InsertContent(oSelectedContent, oNearestPos))
			return false;

		oParagraph.Parent.InsertContent(oSelectedContent, oNearestPos);
		oParagraph.Clear_NearestPosArray();
		// TODO: Выяснить нужно ли снимать выделение с автофигур
		this.Document.MoveCursorRight(false, false, true);
		return true;
	};

	/**
	 * Get a report about all the comments added to the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {object}
	 */
	ApiDocument.prototype.GetCommentsReport = function()
	{
		var oResult = {};
		var oReport = this.Document.Api.asc_GetCommentsReportByAuthors();
		for (var sUserName in oReport)
		{
			var arrUserComments = oReport[sUserName];
			oResult[sUserName] = [];

			for (var nIndex = 0, nCount = arrUserComments.length; nIndex < nCount; ++nIndex)
			{
				var isAnswer     = oReport[sUserName][nIndex].Top ? false : true;
				var oCommentData = oReport[sUserName][nIndex].Data;

				if (isAnswer)
				{
					oResult[sUserName].push({
						"IsAnswer"       : true,
						"CommentMessage" : oCommentData.GetText(),
						"Date"           : oCommentData.GetDateTime()
					});
				}
				else
				{
					var sQuoteText = oCommentData.GetQuoteText();
					oResult[sUserName].push({
						"IsAnswer"       : false,
						"CommentMessage" : oCommentData.GetText(),
						"Date"           : oCommentData.GetDateTime(),
						"QuoteText"      : sQuoteText,
						"IsSolved"       : oCommentData.IsSolved()
					});
				}
			}
		}

		return oResult;
	};

	/**
	 * Get a report about every change which was made to the document in the review mode.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {object}
	 */
	ApiDocument.prototype.GetReviewReport = function()
	{
		var oResult = {};
		var oReport = this.Document.Api.asc_GetTrackRevisionsReportByAuthors();
		for (var sUserName in oReport)
		{
			var arrUsersChanges = oReport[sUserName];
			oResult[sUserName] = [];

			for (var nIndex = 0, nCount = arrUsersChanges.length; nIndex < nCount; ++nIndex)
			{
				var oChange = oReport[sUserName][nIndex];

				var nType = oChange.get_Type();
				var oElement = {};
				// TODO: Посмотреть почем Value приходит массивом.
				if (c_oAscRevisionsChangeType.TextAdd === nType)
				{
					oElement = {
						"Type" : "TextAdd",
						"Value" : oChange.get_Value().length ? oChange.get_Value()[0] : ""
					};
				}
				else if (c_oAscRevisionsChangeType.TextRem == nType)
				{
					oElement = {
						"Type" : "TextRem",
						"Value" : oChange.get_Value().length ? oChange.get_Value()[0] : ""
					};
				}
				else if (c_oAscRevisionsChangeType.ParaAdd === nType)
				{
					oElement = {
						"Type" : "ParaAdd"
					};
				}
				else if (c_oAscRevisionsChangeType.ParaRem === nType)
				{
					oElement = {
						"Type" : "ParaRem"
					};
				}
				else if (c_oAscRevisionsChangeType.TextPr === nType)
				{
					oElement = {
						"Type" : "TextPr"
					};
				}
				else if (c_oAscRevisionsChangeType.ParaPr === nType)
				{
					oElement = {
						"Type" : "ParaPr"
					};
				}
				else
				{
					oElement = {
						"Type" : "Unknown"
					};
				}
				oElement["Date"] = oChange.get_DateTime();
				oResult[sUserName].push(oElement);
			}
		}
		return oResult;
	};
	/**
	 * Find and replace text.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {Object} oProperties The properties for find and replace.
	 * @param {string} oProperties.searchString Search string.
	 * @param {string} oProperties.replaceString Replacement string.
	 * @param {string} [oProperties.matchCase=true] Case sensitive or not.
	 *
	 */
	ApiDocument.prototype.SearchAndReplace = function(oProperties)
	{
		var sSearch     = oProperties["searchString"];
		var sReplace    = oProperties["replaceString"];
		var isMatchCase = undefined !== oProperties["matchCase"] ? oProperties.matchCase : true;

		var oSearchEngine = this.Document.Search(sSearch, {MatchCase : isMatchCase});
		if (!oSearchEngine)
			return;

		this.Document.ReplaceSearchElement(sReplace, true, null, false);
	};
	/**
	 * Get the list of all the content controls in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiBlockLvlSdt[] | ApiInlineLvlSdt[]}
	 */
	ApiDocument.prototype.GetAllContentControls = function()
	{
		var arrResult = [];
		var arrControls = this.Document.GetAllContentControls();
		for (var nIndex = 0, nCount = arrControls.length; nIndex < nCount; ++nIndex)
		{
			var oControl = arrControls[nIndex];

			if (oControl instanceof CBlockLevelSdt)
				arrResult.push(new ApiBlockLvlSdt(oControl));
			else if (oControl instanceof CInlineLevelSdt)
				arrResult.push(new ApiInlineLvlSdt(oControl));
		}

		return arrResult;
	};
	/**
	 * Set change tracking mode.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param isTrack {boolean} - Specifies if the change tracking mode is set or not.
	 */
	ApiDocument.prototype.SetTrackRevisions = function(isTrack)
	{
		this.Document.SetGlobalTrackRevisions(isTrack);
	};
	/**
	 * Check if change tracking mode is enabled or not.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiDocument.prototype.IsTrackRevisions = function()
	{
		return this.Document.IsTrackRevisions();
	};
	/**
	 * Return a Range object that represents the part of the document contained in the specified document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiDocument.prototype.GetRange = function(Start, End)
	{
		var Range = new ApiRange(this.Document, Start, End);

		return Range;
	};
	/**
	 * Get a range object by the current selection.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRange | null} - returns null if selection doesn't exist.
	 * */
	ApiDocument.prototype.GetRangeBySelect = function()
	{
		if (!this.Document.IsSelectionUse())
			return null;

		private_RefreshRangesPosition();
			
		var selectDirection	= this.Document.GetSelectDirection();
		var documentState	= this.Document.SaveDocumentState();
		var StartPos		= null;
		var EndPos			= null;

		private_TrackRangesPositions();

		if (selectDirection === 1)
		{
			StartPos	= documentState.StartPos;
			EndPos		= documentState.EndPos;
		}
		else if (selectDirection === -1)
		{
			StartPos	= documentState.EndPos;
			EndPos		= documentState.StartPos;
		}

		this.Document.LoadDocumentState(documentState);
		private_RemoveEmptyRanges();
		return new ApiRange(StartPos[0].Class, StartPos, EndPos);
	};
	/**
	 * Get the last document element. 
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {?DocumentElement}
	 */
	ApiDocument.prototype.Last = function()
	{
		return this.GetElement(this.GetElementsCount() - 1);
	};
	
	/**
	 * Remove a bookmark from the document, if one exists.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sName - The bookmark name.
	 * @returns {bool} - returns false if param is invalid.
	 */
	ApiDocument.prototype.DeleteBookmark = function(sName)
	{
		if (sName === undefined)
			return false;

		this.Document.RemoveBookmark(sName);

		return true;
	};
	/**
	 * Add a comment to the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} Comment - The comment text.
	 * @param {string} Autor - The author's name (not obligatory).
	 * @returns {bool} - returns false if params are invalid.
	 */
	ApiDocument.prototype.AddComment = function(Comment, Autor)
	{
		if (!Comment || typeof(Comment) !== "string")
			return false;
	
		if (typeof(Autor) !== "string")
			Autor = "";
		
		var CommentData = new AscCommon.CCommentData();
		CommentData.SetText(Comment);
		CommentData.SetUserName(Autor);

		var COMENT = this.Document.AddComment(CommentData, true);

		if (null !== COMENT)
		{
			editor.sync_AddComment(COMENT.Get_Id(), CommentData);
		}

		return true;
	};
	/**
	 * Get a bookmark's range.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sName - The bookmark name.
	 * @return {ApiRange | null} - returns null if sName is invalid.
	 */
	ApiDocument.prototype.GetBookmarkRange = function(sName)
	{
		if (typeof(sName) !== "string")
			return null;
		
		var Document = private_GetLogicDocument();
		private_RefreshRangesPosition();
		var oldSelectionInfo = Document.SaveDocumentState();
		
		private_TrackRangesPositions();

		this.Document.GoToBookmark(sName, true);

		var oRange = this.GetRangeBySelect();

		this.Document.LoadDocumentState(oldSelectionInfo);
		this.Document.UpdateSelection();

		return oRange;
	};
	/**
	 * Get the collection of section objects in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @return {ApiSection[]}  
	 */
	ApiDocument.prototype.GetSections = function()
	{
		var arrApiSections = [];

		for (var Index = 0; Index < this.Document.SectionsInfo.Elements.length; Index++)
			arrApiSections.push(new ApiSection(this.Document.SectionsInfo.Elements[Index].SectPr))

		return arrApiSections;
	};
	/**
	 * Get the collection of tables on a given absolute page.
	 * <note>This method can be a little bit slow, because it runs the document calculation
	 * process to arrange tables on the specified page.</note>
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param nPage - The page number.
	 * @return {ApiTable[]}  
	 */
	ApiDocument.prototype.GetAllTablesOnPage = function(nPage)
	{
		var arrApiAllTables = [];

		this.Document.private_Recalculate(undefined, true, nPage + 1);
		var arrAllTables = this.Document.GetAllTablesOnPage(nPage);

		for (var Index = 0; Index < arrAllTables.length; Index++)
		{
			arrApiAllTables.push(new ApiTable(arrAllTables[Index].Table));
		}

		return arrApiAllTables;
	};
	/**
	 * Remove current selection.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 */
	ApiDocument.prototype.RemoveSelection = function()
	{
		this.Document.RemoveSelection();
	};
	/**
	 * Get the collection of drawing objects in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @return {ApiDrawing[]}  
	 */
	ApiDocument.prototype.GetAllDrawingObjects = function()
	{
		var arrAllDrawing = this.Document.GetAllDrawingObjects();
		var arrApiShapes  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
			arrApiShapes.push(new ApiDrawing(arrAllDrawing[Index]));
		
		return arrApiShapes;
	};
	/**
	 * Get the collection of shape objects in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @return {ApiShape[]}  
	 */
	ApiDocument.prototype.GetAllShapes = function()
	{
		var arrAllDrawing = this.Document.GetAllDrawingObjects();
		var arrApiShapes  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof AscFormat.CShape)
				arrApiShapes.push(new ApiShape(arrAllDrawing[Index].GraphicObj));
		}
		
		return arrApiShapes;
	};
	/**
	 * Get the collection of image objects in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @return {ApiImage[]}  
	 */
	ApiDocument.prototype.GetAllImages = function()
	{
		var arrAllDrawing = this.Document.GetAllDrawingObjects();
		var arrApiImages  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof CImageShape)
				arrApiImages.push(new ApiImage(arrAllDrawing[Index].GraphicObj));
		}
		
		return arrApiImages;
	};
	/**
	 * Get the collection of chart objects in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @return {ApiChart[]}  
	 */
	ApiDocument.prototype.GetAllCharts = function()
	{
		var arrAllDrawing = this.Document.GetAllDrawingObjects();
		var arrApiCharts  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof CChartSpace)
				arrApiCharts.push(new ApiChart(arrAllDrawing[Index].GraphicObj));
		}
		
		return arrApiCharts;
	};
	/**
	 * Search for the scope of a document object. The search results are a collection of ApiRange objects.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {bool} isMatchCase - Case sensitive or not. 
	 * @return {ApiRange[]}  
	 */
	ApiDocument.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase	= false;
		
		var foundItems 		= [];
		var arrApiRanges	= [];
		var docSearchEngine	= this.Document.Search(sText, {MatchCase : isMatchCase});

		var docSearchEngineElementsLenght = 0;
		for (var FoundId in docSearchEngine.Elements)
			docSearchEngineElementsLenght++;

		for (var Index = 1; Index <= docSearchEngineElementsLenght; Index++)
			foundItems.push(docSearchEngine.Elements[Index]);

		for (var Index1 = 0; Index1 < foundItems.length; Index1++)
		{
			for (var Index2 = Index1 + 1; Index2 < foundItems.length; Index2++)
			{
				if (foundItems[Index1].Id === foundItems[Index2].Id)
				{
					foundItems.splice(Index2, 1);
					Index2--;
				}
			}
		}

		for (var para in foundItems)
		{
			var oParagraph			= new ApiParagraph(foundItems[para]);
			var arrOfParaApiRanges	= oParagraph.Search(sText, isMatchCase);

			for (var itemRange = 0; itemRange < arrOfParaApiRanges.length; itemRange++)	
				arrApiRanges.push(arrOfParaApiRanges[itemRange]);
		}

		return arrApiRanges;
	};
	/**
	 * Convert document to markdown.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {bool} [bHtmlHeadings=false] - If you have used multiple Heading 1 headings in your Doc, set this param true to demote all heading levels to conform with the following standard: single H1 as title, H2 as top-level heading in the text body.
	 * @param {bool} [bBase64img=false] - set this param true if you want images to be created in base64 format.
	 * @param {bool} [bDemoteHeadings=false] - Not all Markdown renderers handle Markdown-style IDs. If that is the case for your target platform, set this param true to generate HTML headings and IDs.
	 * @param {bool} [bRenderHTMLTags=false] - By default, angle brackets (<) will be replaced by the &lt; entity. If you really want to embed HTML tags in your Markdown, set this param true to preserve them. 
	 * Or, if you just want to use an occasional HTML tag, you can escape the opening angle bracket like this: \<tag>text\</tag>.
	 * @returns {string}
	 */
	ApiDocument.prototype.ToMarkdown = function(bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags) 
	{
		var oConfig = 
		{
			convertType : "markdown",
			htmlHeadings : bHtmlHeadings || false,
			base64img : bBase64img || false,
			demoteHeadings : bDemoteHeadings || false,
			renderHTMLTags : bRenderHTMLTags || false
		};
		var oMarkdown = new CMarkdownConverter(oConfig);
		return oMarkdown.DoMarkdown();
	};
	/**
	 * Convert document to html.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {bool} [bHtmlHeadings=false] - If you have used multiple Heading 1 headings in your Doc, set this param true to demote all heading levels to conform with the following standard: single H1 as title, H2 as top-level heading in the text body.
	 * @param {bool} [bBase64img=false] - set this param true if you want images to be created in base64 format.
	 * @param {bool} [bDemoteHeadings=false] - Not all Markdown renderers handle Markdown-style IDs. If that is the case for your target platform, set this param true to generate HTML headings and IDs.
	 * @param {bool} [bRenderHTMLTags=false] - By default, angle brackets (<) will be replaced by the &lt; entity. If you really want to embed HTML tags in your Markdown, set this param true to preserve them. 
	 * Or, if you just want to use an occasional HTML tag, you can escape the opening angle bracket like this: \<tag>text\</tag>.
	 * @returns {string}
	 */
	ApiDocument.prototype.ToHtml = function(bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags) 
	{
		var oConfig = 
		{
			convertType : "html",
			htmlHeadings : bHtmlHeadings || false,
			base64img : bBase64img || false,
			demoteHeadings : bDemoteHeadings || false,
			renderHTMLTags : bRenderHTMLTags || false
		};
		var oMarkdown = new CMarkdownConverter(oConfig);
		return oMarkdown.DoHtml();
	};

	/**
	 * Insert a watermark on each document page.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {?string} [sText="WATERMARK"] - Watermark text.
	 * @param {?boolean} [bIsDiagonal=true] - Specifies if the watermark is placed diagonally (true) or horizontally (false).
	 */
	ApiDocument.prototype.InsertWatermark = function(sText, bIsDiagonal){
		var oSectPrMap = {};
		if(this.Document.SectPr){
			oSectPrMap[this.Document.SectPr.Get_Id()] = this.Document.SectPr;
		}
		var oElement;
		for(var i = 0; i < this.Document.Content.length; ++i){
			oElement = this.Document.Content[i];
			if(oElement instanceof Paragraph){
				if(oElement.SectPr){
					oSectPrMap[oElement.SectPr.Get_Id()] = oElement.SectPr;
				}
			}
		}
		var oHeadersMap = {};
		var oApiSection, oHeader;
		for(var sId in oSectPrMap){
			if(oSectPrMap.hasOwnProperty(sId)){
				oApiSection = new ApiSection(oSectPrMap[sId]);
				oHeader = oApiSection.GetHeader("title", false);
				if(oHeader){
					oHeadersMap[oHeader.Document.Get_Id()] = oHeader;
				}
				oHeader = oApiSection.GetHeader("even", false);
				if(oHeader){
					oHeadersMap[oHeader.Document.Get_Id()] = oHeader;
				}
				oHeader = oApiSection.GetHeader("default", true);
				if(oHeader){
					oHeadersMap[oHeader.Document.Get_Id()] = oHeader;
				}
			}
		}
		for(var sId in oHeadersMap){
			if(oHeadersMap.hasOwnProperty(sId)){
				privateInsertWatermarkToContent(this.Document.Api, oHeadersMap[sId], sText, bIsDiagonal);
			}
		}
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiParagraph
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"document"}
	 */
	ApiParagraph.prototype.GetClassType = function()
	{
		return "paragraph";
	};
	/**
	 * Add some text to the element.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} [sText=""] - The text that we want to insert into the current document element.
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddText = function(sText)
	{
		var oRun = new ParaRun(this.Paragraph, false);

		if (!sText || !sText.length)
			return new ApiRun(oRun);

		oRun.AddText(sText);

		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Add page break and start the next element from the next page.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddPageBreak = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new ParaNewLine(break_Page));
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Add line break to the current position and start the next element from a new line.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddLineBreak = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new ParaNewLine(break_Line));
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};

	/**
	 * Add column break to the current position and start the next element from a new column.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddColumnBreak = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new ParaNewLine(break_Column));
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Insert the number of the current document page into the paragraph.
	 * <note>This method works for the paragraphs in the document header/footer only.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddPageNumber = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new ParaPageNum());
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Insert the number of pages in the current document into the paragraph.
	 * <note>This method works for the paragraphs in the document header/footer only.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddPagesCount = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new ParaPageCount());
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Get the text properties of the paragraph mark which is used to mark the paragraph end. The mark can also acquire
	 * common text properties like bold, italic, underline, etc.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTextPr}
	 */
	ApiParagraph.prototype.GetParagraphMarkTextPr = function()
	{
		return new ApiTextPr(this, this.Paragraph.TextPr.Value.Copy());
	};
	/**
	 * Get paragraph properties.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiParaPr}
	 */
	ApiParagraph.prototype.GetParaPr = function()
	{
		return new ApiParaPr(this, this.Paragraph.Pr.Copy());
	};
	/**
	 * Get a numbering definition and numbering level for the numbered list.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {?ApiNumberingLevel}
	 */
	ApiParagraph.prototype.GetNumbering = function()
	{
		var oNumPr = this.Paragraph.GetNumPr();
		if (!oNumPr)
			return null;

		var oLogicDocument   = private_GetLogicDocument();
		var oGlobalNumbering = oLogicDocument.GetNumbering();
		var oNum             = oGlobalNumbering.GetNum(oNumPr.NumId);
		if (!oNum)
			return null;

		return new ApiNumberingLevel(oNum, oNumPr.Lvl);
	};
	/**
	 * Specify that the current paragraph references a numbering definition instance in the current document.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @see Same as {@link ApiParagraph#SetNumPr}
	 * @param {ApiNumberingLevel} oNumberingLevel - The numbering level which will be used for assigning the numbers to the paragraph.
	 */
	ApiParagraph.prototype.SetNumbering = function(oNumberingLevel)
	{
		if (!(oNumberingLevel instanceof ApiNumberingLevel))
			return;

		this.SetNumPr(oNumberingLevel.GetNumbering(), oNumberingLevel.GetLevelIndex());
	};
	/**
	 * Get the number of elements in the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {number}
	 */
	ApiParagraph.prototype.GetElementsCount = function()
	{
		// TODO: ParaEnd
		return this.Paragraph.Content.length - 1;
	};
	/**
	 * Get the paragraph element using the position specified.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {number} nPos - The position where the element which content we want to get must be located.
	 * @returns {?ParagraphContent}
	 */
	ApiParagraph.prototype.GetElement = function(nPos)
	{
		// TODO: ParaEnd
		if (nPos < 0 || nPos >= this.Paragraph.Content.length - 1)
			return null;

		return private_GetSupportedParaElement(this.Paragraph.Content[nPos]);
	};
	/**
	 * Remove the element using the position specified.
	 * <note>If the element you remove is the last paragraph element (i.e. all the elements are removed from the paragraph),
     * a new empty run is automatically created. If you want to add
	 * content to this run, use the {@link ApiParagraph#GetElement} method.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {number} nPos - The element position which we want to remove from the paragraph.
	 */
	ApiParagraph.prototype.RemoveElement = function(nPos)
	{
		if (nPos < 0 || nPos >= this.Paragraph.Content.length - 1)
			return;

		this.Paragraph.RemoveFromContent(nPos, 1);
		this.Paragraph.CorrectContent();
	};
	/**
	 * Remove all the elements from the current paragraph.
	 * <note>When all the elements are removed from the paragraph, a new empty run is automatically created. If you want to add
	 * content to this run, use the {@link ApiParagraph#GetElement} method.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiParagraph.prototype.RemoveAllElements = function()
	{
		if (this.Paragraph.Content.length > 1)
		{
			this.Paragraph.RemoveFromContent(0, this.Paragraph.Content.length - 1);
			this.Paragraph.CorrectContent();
		}
	};
	/**
	 * Delete current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {bool} - returns false if paragraph haven't parent.
	 */
	ApiParagraph.prototype.Delete = function()
	{
		var parentOfElement = this.Paragraph.GetParent();

		var PosInDocument = parentOfElement.Content.indexOf(this.Paragraph);

		if (PosInDocument !== - 1)
		{
			this.Paragraph.PreDelete();
			parentOfElement.Remove_FromContent(PosInDocument, 1, true);

			return true;
		}
		else 
			return false;
	};
	/**
	 * Get the next paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiParagraph | null} - returns null if paragraph is last.
	 */
	ApiParagraph.prototype.GetNext = function()
	{
		var nextPara = this.Paragraph.GetNextParagraph();
        if (nextPara !== null)
            return new ApiParagraph(nextPara);

        return null;
	};
	/**
	 * Get the previous paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiParagraph} - returns null if paragraph is first.
	 */
	ApiParagraph.prototype.GetPrevious = function()
	{
		var prevPara = this.Paragraph.GetPrevParagraph();
        if (prevPara !== null)
            return new ApiParagraph(prevPara);

        return null;
	};
	/**
	 * Create a paragraph copy. Ingnore comments, footnote references, complex fields.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiParagraph}
	 */
	ApiParagraph.prototype.Copy = function()
	{
		var oParagraph = this.Paragraph.Copy(undefined, private_GetDrawingDocument(), {
			SkipComments          : true,
			SkipAnchors           : true,
			SkipFootnoteReference : true,
			SkipComplexFields     : true
		});

		return new ApiParagraph(oParagraph);
	};
	/**
	 * Add an element to the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ParagraphContent} oElement - The document element which will be added at the current position. Returns false if the
	 * type of oElement is not supported by a paragraph.
	 * @param {number} [nPos] The position where the current element will be added. If this value is not
	 * specified then the element will be added at the end of the current paragraph.
	 * @returns {boolean} Returns <code>false</code> if the type of <code>oElement</code> is not supported by paragraph
	 * content.
	 */
	ApiParagraph.prototype.AddElement = function(oElement, nPos)
	{
		// TODO: ParaEnd
		if (!private_IsSupportedParaElement(oElement) || nPos < 0 || nPos > this.Paragraph.Content.length - 1)
			return false;

		var oParaElement = oElement.private_GetImpl();
		if (undefined !== nPos)
		{
			this.Paragraph.Add_ToContent(nPos, oParaElement);
		}
		else
		{
			private_PushElementToParagraph(this.Paragraph, oParaElement);
		}

		return true;
	};
	/**
	 * Add a tab stop to the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddTabStop = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new ParaTab());
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Add a drawing object (image, shape or chart) to the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {ApiDrawing} oDrawing - The object which will be added to the current paragraph.
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddDrawing = function(oDrawing)
	{
		var oRun = new ParaRun(this.Paragraph, false);

		if (!(oDrawing instanceof ApiDrawing))
			return new ApiRun(oRun);

		oRun.Add_ToContent(0, oDrawing.Drawing);
		private_PushElementToParagraph(this.Paragraph, oRun);
		oDrawing.Drawing.Set_Parent(oRun);

		return new ApiRun(oRun);
	};

	/**
	 * Add an inline container.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {ApiInlineLvlSdt?} oSdt - An inline container. If undefined or null, then new class ApiInlineLvlSdt will be created and added to paragraph.
	 * @returns {ApiInlineLvlSdt}
	 */
	ApiParagraph.prototype.AddInlineLvlSdt = function(oSdt)
	{
		if (!oSdt || !(oSdt instanceof ApiInlineLvlSdt))
		{
			var _oSdt = new CInlineLevelSdt();
			_oSdt.Add_ToContent(0, new ParaRun(null, false));
			oSdt = new ApiInlineLvlSdt(_oSdt);
		}

		private_PushElementToParagraph(this.Paragraph, oSdt.Sdt);
		return oSdt;
	};
	/**
	 * Add a comment to the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string} Comment - The comment text.
	 * @param {string} Autor - The author's name (not obligatory).
	 * @returns {bool} - returns false if params are invalid.
	 */
	ApiParagraph.prototype.AddComment = function(Comment, Autor)
	{
		if (!Comment || typeof(Comment) !== "string")
			return false;
	
		if (typeof(Autor) !== "string")
			Autor = "";

		var CommentData = new AscCommon.CCommentData();
		CommentData.SetText(Comment);
		CommentData.SetUserName(Autor);

		var oDocument = private_GetLogicDocument()

		var oFirstun    = this.Paragraph.GetFirstRun();
		var StartPos	= oFirstun.GetDocumentPositionFromObject();
		var oEndRun     = this.Paragraph.Content[this.Paragraph.Content.length - 2];
		var EndPos		= oEndRun.GetDocumentPositionFromObject();
		StartPos.push({Class: oFirstun, Position: 0});
		EndPos.push({Class: oEndRun, Position: oEndRun.Content.length});

		oDocument.Selection.Use = true;
		oDocument.SetContentSelection(StartPos, EndPos, 0, 0, 0);

		var COMENT = oDocument.AddComment(CommentData, false);
		oDocument.RemoveSelection();
		
		if (null != COMENT)
		{
			editor.sync_AddComment(COMENT.Get_Id(), CommentData);
		}

		return true;
	};
	/**
	 * Add a hyperlink to a paragraph. 
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string} sLink - The link address.
	 * @param {string} sScreenTipText - The screen tip text.
	 * @return {ApiHyperlink | null} - returns null if params are invalid.
	 */
	ApiParagraph.prototype.AddHyperlink = function(sLink, sScreenTipText)
	{
		if (typeof(sLink) !== "string" || sLink === "" || sLink.length > Asc.c_nMaxHyperlinkLength)
			return null;
		if (typeof(sScreenTipText) !== "string")
			sScreenTipText = "";
		
		var oDocument	= editor.private_GetLogicDocument();
		var hyperlinkPr	= new Asc.CHyperlinkProperty();
		var urlType		= AscCommon.getUrlType(sLink);
		var oHyperlink	= null;

		this.Paragraph.SelectAll(1);
		if (!/(((^https?)|(^ftp)):\/\/)|(^mailto:)/i.test(sLink))
			sLink = (urlType === 0) ? null :(( (urlType === 2) ? 'mailto:' : 'http://' ) + sLink);

		sLink = sLink.replace(new RegExp("%20",'g')," ");
		hyperlinkPr.put_Value(sLink);
		hyperlinkPr.put_ToolTip(sScreenTipText);
		hyperlinkPr.put_Bookmark(null);
		
		oHyperlink = new ApiHyperlink(this.Paragraph.AddHyperlink(hyperlinkPr));
		oDocument.RemoveSelection();

		return oHyperlink;
	};
	/**
	 * Return a Range object that represents the part of the document contained in the specified paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiParagraph.prototype.GetRange = function(Start, End)
	{
		var Range = new ApiRange(this.Paragraph, Start, End);
	
		return Range;
	};
	/**
	 * Add an element to the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {ParagraphContent} oElement - The document element which will be added at the current position. Returns false if the
	 * type of oElement is not supported by a paragraph.
	 * @returns {boolean} Returns <code>false</code> if the type of <code>oElement</code> is not supported by paragraph
	 * content.
	 */
	ApiParagraph.prototype.Push = function(oElement)
	{
		if (oElement instanceof ApiRun || ApiInlineLvlSdt)
		{
			this.AddElement(oElement);
		}
		else if (typeof oElement === "string")
		{
			var LastTextPrInParagraph = undefined;

			if (this.GetLastRunWithText() !== null)
			{
				LastTextPrInParagraph = this.GetLastRunWithText().GetTextPr().TextPr;
			}
			else 
			{
				LastTextPrInParagraph = this.Paragraph.TextPr.Value;
			}
			
			var oRun = editor.CreateRun();
			oRun.AddText(oElement);
			oRun.Run.Apply_TextPr(LastTextPrInParagraph, undefined, true);
			
			this.AddElement(oRun);
		}
		else 
			return false;
	};
	/**
	 * Get the last Run with text in the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun} Returns <code>false</code> if the paragraph doesn't containt the required run.
	 */
	ApiParagraph.prototype.GetLastRunWithText = function()
	{
		for (var curElement = this.GetElementsCount() - 1; curElement >= 0; curElement--)
		{
			var Element = this.GetElement(curElement);

			if (Element instanceof ApiRun)
			{
				for (var Index = 0; Index < Element.Run.GetElementsCount(); Index++)
				{
					if (Element.Run.GetElement(Index) instanceof ParaText)
					{
						return Element;
					}
				}
			}
		}

		return this.GetElement(this.GetElementsCount() - 1);
	};
	/**
	 * Set the bold property to the text character.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isBold - Specifies that the contents of this paragraph are displayed bold.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetBold = function(isBold)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({Bold : isBold}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify that any lowercase characters in this paragraph are formatted for display only as their capital letter character equivalents.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isCaps - Specifies that the contents of the current paragraph are displayed capitalized.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetCaps = function(isCaps)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({Caps : isCaps}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Set the text color for the current paragraph in the RGB format.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - If this parameter is set to "true", then r,g,b parameters will be ignored.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetColor = function(r, g, b, isAuto)
	{
		var color = new Asc.asc_CColor();
		color.r    = r;
		color.g    = g;
		color.b    = b;
		color.Auto = isAuto;

		this.Paragraph.SetApplyToAll(true);
		if (true === color.Auto)
		{
			this.Paragraph.Add(new AscCommonWord.ParaTextPr({
				Color      : {
					Auto : true,
					r    : 0,
					g    : 0,
					b    : 0
				}, Unifill : undefined
			}));
		}
		else
		{
			var Unifill        = new AscFormat.CUniFill();
			Unifill.fill       = new AscFormat.CSolidFill();
			Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
			this.Paragraph.Add(new AscCommonWord.ParaTextPr({Unifill : Unifill}));
		}
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify that the contents of this paragraph are displayed with two horizontal lines through each character displayed on the line.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isDoubleStrikeout - Specifies that the contents of the current paragraph are displayed double struck through.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetDoubleStrikeout = function(isDoubleStrikeout)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({DStrikeout : isDoubleStrikeout}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Set all 4 font slots with the specified font family.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string} sFontFamily - The font family or families used for the current paragraph.
	 * @returns {ApiParagraph | false} 
	 */
	ApiParagraph.prototype.SetFontFamily = function(sFontFamily)
	{
		if (typeof sFontFamily !== "string")
			return false;

		var loader   = AscCommon.g_font_loader;
		var fontinfo = g_fontApplication.GetFontInfo(sFontFamily);
		var isasync  = loader.LoadFont(fontinfo);

		if (isasync === false)
		{
			var FontFamily = {
				Name : sFontFamily,
				Index : -1
			};

			this.Paragraph.SetApplyToAll(true);
			this.Paragraph.Add(new AscCommonWord.ParaTextPr({FontFamily : FontFamily}));
			this.Paragraph.SetApplyToAll(false);
			
			return this;
		}
		
		return false;
	};
	/**
	 * Set the font size for the characters of the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {hps} nSize - The text size value measured in half-points (1/144 of an inch).
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetFontSize = function(nSize)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({FontSize : nSize/2}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify a highlighting color in the RGB format which is applied as a background for the contents of the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isNone=false] If this parameter is set to "true", then r,g,b parameters will be ignored.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetHighlight = function(r, g, b, isNone)
	{
		this.Paragraph.SetApplyToAll(true);
		if (true === isNone)
		{
			this.Paragraph.Add(new ParaTextPr({HighLight : highlight_None}));
		}
		else
		{
			var color = new CDocumentColor(r, g, b);
			this.Paragraph.Add(new ParaTextPr({HighLight : color}));
		}
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Set the italic property to the text character.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isItalic - Specifies that the contents of the current paragraph are displayed italicized.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetItalic = function(isItalic)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({Italic : isItalic}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify the amount by which text is raised or lowered for this paragraph in relation to the default
	 * baseline of the surrounding non-positioned text.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {hps} nPosition - Specifies a positive (raised text) or negative (lowered text)
	 * measurement in half-points (1/144 of an inch).
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetPosition = function(nPosition)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({Position : nPosition}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify the shading applied to the contents of the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type applied to the contents of the current paragraph.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetShd = function(sType, r, g, b)
	{
		var color = new Asc.asc_CColor();
		color.r    = r;
		color.g    = g;
		color.b    = b;
		color.Auto = false;

		this.Paragraph.SetApplyToAll(true);

		var Shd = new CDocumentShd();

		if (sType === "nil")
		{
			var _Shd = {Value : Asc.c_oAscShdNil};
			Shd.Set_FromObject(_Shd);
			this.Paragraph.SetParagraphShd(_Shd);
		}
		else if (sType === "clear")
		{
			var Unifill        = new AscFormat.CUniFill();
			Unifill.fill       = new AscFormat.CSolidFill();
			Unifill.fill.color = AscFormat.CorrectUniColor(color, Unifill.fill.color, 1);
			var _Shd = {
				Value   : Asc.c_oAscShdClear,
				Color   : {
					r : color.asc_getR(),
					g : color.asc_getG(),
					b : color.asc_getB()
				},
				Unifill : Unifill
			};
			
			Shd.Set_FromObject(_Shd);
			this.Paragraph.SetParagraphShd(_Shd);
		}
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify that all the small letter characters in this paragraph are formatted for display only as their capital
	 * letter character equivalents which are two points smaller than the actual font size specified for this text.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isSmallCaps - Specifies if the contents of the current paragraph are displayed capitalized two points smaller or not.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetSmallCaps = function(isSmallCaps)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({
			SmallCaps : isSmallCaps,
			Caps      : false
		}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Set text spacing measured in twentieths of a point.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {twips} nSpacing - The value of the text spacing measured in twentieths of a point (1/1440 of an inch).
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetSpacing = function(nSpacing)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({Spacing : nSpacing}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify that the contents of this paragraph are displayed with a single horizontal line through the center of the line.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isStrikeout - Specifies that the contents of the current paragraph are displayed struck through.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetStrikeout = function(isStrikeout)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({
			Strikeout  : isStrikeout,
			DStrikeout : false
			}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify that the contents of this paragraph are displayed along with a line appearing directly below the character
	 * (less than all the spacing above and below the characters on the line).
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isUnderline - Specifies that the contents of the current paragraph are displayed underlined.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetUnderline = function(isUnderline)
	{
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({Underline : isUnderline}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Specify the alignment which will be applied to the contents of this paragraph in relation to the default appearance of the paragraph text:
	 * * <b>"baseline"</b> - the characters in the current paragraph will be aligned by the default text baseline.
	 * * <b>"subscript"</b> - the characters in the current paragraph will be aligned below the default text baseline.
	 * * <b>"superscript"</b> - the characters in the current paragraph will be aligned above the default text baseline.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {("baseline" | "subscript" | "superscript")} sType - The vertical alignment type applied to the text contents.
	 * @returns {ApiParagraph | null} - returns null is sType is invalid.
	 */
	ApiParagraph.prototype.SetVertAlign = function(sType)
	{
		var value = undefined;

		if (sType === "baseline")
			value = 0;
		else if (sType === "subscript")
			value = 2;
		else if (sType === "superscript")
			value = 1;
		else 
			return null;

		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({VertAlign : value}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Get the last no empty element of the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {?ParagraphContent}
	 */
	ApiParagraph.prototype.Last = function()
	{
		var LastNoEmptyElement = null;

		for (var Index = this.GetElementsCount() - 1; Index >= 0; Index--)
		{
			LastNoEmptyElement = this.GetElement(Index);
			
			if (!LastNoEmptyElement || LastNoEmptyElement instanceof ApiUnsupported)
				continue;

			return LastNoEmptyElement;
		}

		return null;
	};
	/**
	 * Get the collection of content control objects in the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiInlineLvlSdt[]}   
	 */
	ApiParagraph.prototype.GetAllContentControls = function()
	{
		var arrApiContentControls = [];

		var ContentControls = this.Paragraph.GetAllContentControls();

		for (var Index = 0; Index < ContentControls.length; Index++)
		{
			if (ContentControls[Index] instanceof CInlineLevelSdt)
				arrApiContentControls.push(new ApiInlineLvlSdt(ContentControls[Index]));
		} 

		return arrApiContentControls;
	};
	/**
	 * Get the collection of drawing objects in the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiDrawing[]}  
	 */
	ApiParagraph.prototype.GetAllDrawingObjects = function()
	{
		var arrAllDrawing = this.Paragraph.GetAllDrawingObjects();
		var arrApiShapes  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			arrApiShapes.push(new ApiDrawing(arrAllDrawing[Index]));
		}
		
		return arrApiShapes;
	};
	/**
	 * Get the collection of shape objects in the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiShape[]}  
	 */
	ApiParagraph.prototype.GetAllShapes = function()
	{
		var arrAllDrawing = this.Paragraph.GetAllDrawingObjects();
		var arrApiShapes  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof AscFormat.CShape)
				arrApiShapes.push(new ApiShape(arrAllDrawing[Index].GraphicObj));
		}

		return arrApiShapes;
	};
	/**
	 * Get the collection of image objects in the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiImage[]}  
	 */
	ApiParagraph.prototype.GetAllImages = function()
	{
		var arrAllDrawing = this.Paragraph.GetAllDrawingObjects();
		var arrApiImages  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof CImageShape)
				arrApiImages.push(new ApiImage(arrAllDrawing[Index].GraphicObj));
		}

		return arrApiImages;
	};
	/**
	 * Get the collection of chart objects in the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiChart[]}  
	 */
	ApiParagraph.prototype.GetAllCharts = function()
	{
		var arrAllDrawing = this.Paragraph.GetAllDrawingObjects();
		var arrApiCharts  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof CChartSpace)
				arrApiCharts.push(new ApiChart(arrAllDrawing[Index].GraphicObj));
		}

		return arrApiCharts;
	};
	/**
	 * Get the content control that contains the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiBlockLvlSdt | null} - returns null is parent content control doesn't exist.  
	 */
	ApiParagraph.prototype.GetParentContentControl = function()
	{
		var ParaPosition = this.Paragraph.GetDocumentPositionFromObject();

		for (var Index = ParaPosition.length - 1; Index >= 1; Index--)
		{
			if (ParaPosition[Index].Class)
				if (ParaPosition[Index].Class instanceof CBlockLevelSdt)
					return new ApiBlockLvlSdt(ParaPosition[Index].Class);
		}

		return null;
	};
	/**
	 * Get the table that contains the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiTable | null} - returns null if parent table doesn't exist.  
	 */
	ApiParagraph.prototype.GetParentTable = function()
	{
		var ParaPosition = this.Paragraph.GetDocumentPositionFromObject();

		for (var Index = ParaPosition.length - 1; Index >= 1; Index--)
		{
			if (ParaPosition[Index].Class instanceof CTable)
				return new ApiTable(ParaPosition[Index].Class);
		}

		return null;
	};
	/**
	 * Get the table cell that contains the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiTableCell | null} - returns null if parent cell doesn't exist.  
	 */
	ApiParagraph.prototype.GetParentTableCell = function()
	{
		var ParaPosition = this.Paragraph.GetDocumentPositionFromObject();

		for (var Index = ParaPosition.length - 1; Index >= 1; Index--)
		{
			if (ParaPosition[Index].Class.Parent && this.Paragraph.IsTableCellContent())
				if (ParaPosition[Index].Class.Parent instanceof CTableCell)
					return new ApiTableCell(ParaPosition[Index].Class.Parent);
		}

		return null;
	};
	/**
	 * Get paragraph text.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {string}  
	 */
	ApiParagraph.prototype.GetText = function()
	{
		var ParaText = this.Paragraph.GetText({ParaEndToSpace : false});

		return ParaText;
	};
	/**
	 * Get paragraph text properties.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiTextPr}  
	 */
	ApiParagraph.prototype.GetTextPr = function()
	{
		var TextPr = this.Paragraph.TextPr.Value;

		return new ApiTextPr(this, TextPr);
	};
	/**
	 * Set paragraph text properties.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr
	 * @return {bool} - returns false if param is invalid.
	 */
	ApiParagraph.prototype.SetTextPr = function(oTextPr)
	{
		if (!(oTextPr instanceof ApiTextPr))
			return false;

		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr(oTextPr.TextPr));
		this.Paragraph.SetApplyToAll(false);

		return true;
	};
	/**
	 * Wrap the paragraph object with a rich text content control.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {number} nType - If nType === 1 -> returns ApiBlockLvlSdt, otherwise returns ApiParagraph. 
	 * @return {ApiParagraph | ApiBlockLvlSdt}  
	 */
	ApiParagraph.prototype.InsertInContentControl = function(nType)
	{
		var Document = private_GetLogicDocument();
		var ContentControl = null;

		var paraIndex	= this.Paragraph.Index;
		if (paraIndex >= 0)
		{
			this.Select(false);
			ContentControl = new ApiBlockLvlSdt(Document.AddContentControl(1));
			Document.RemoveSelection();
		}
		else 
		{
			ContentControl = new ApiBlockLvlSdt(new CBlockLevelSdt(Document, Document));
			ContentControl.Sdt.SetDefaultTextPr(Document.GetDirectTextPr());
			ContentControl.Sdt.Content.RemoveFromContent(0, ContentControl.Sdt.Content.GetElementsCount(), false);
			ContentControl.Sdt.Content.AddToContent(0, this.Paragraph);
			ContentControl.Sdt.SetShowingPlcHdr(false);
		}

		if (nType === 1)
			return ContentControl;
		else 
			return this;
	};
	/**
	 * Insert a paragraph at the specified position.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string | ApiParagraph} paragraph - Text or paragraph.
	 * @param {string} sPosition - The position where the text or paragraph will be inserted ("before" or "after" the paragraph specified).
	 * @param {bool} beRNewPara - If "true" -> returns new paragraph, else returns this paragraph.
	 * @return {ApiParagraph | null} - returns null if param paragraph is invalid. 
	 */
	ApiParagraph.prototype.InsertParagraph = function(paragraph, sPosition, beRNewPara)
	{
		var paraParent = this.Paragraph.GetParent();
		var paraIndex  = paraParent.Content.indexOf(this.Paragraph);
		var oNewPara   = null;

		if (sPosition !== "before" && sPosition !== "after")
			sPosition = "after";

		if (paragraph instanceof ApiParagraph)
		{
			oNewPara = paragraph;

			if (sPosition === "before")
				paraParent.Internal_Content_Add(paraIndex, oNewPara.private_GetImpl());
			else if (sPosition === "after")
				paraParent.Internal_Content_Add(paraIndex + 1, oNewPara.private_GetImpl());
		}
		else if (typeof paragraph === "string")
		{
			oNewPara = editor.CreateParagraph();
			oNewPara.AddText(paragraph);

			if (sPosition === "before")
				paraParent.Internal_Content_Add(paraIndex, oNewPara.private_GetImpl());
			else if (sPosition === "after")
				paraParent.Internal_Content_Add(paraIndex + 1, oNewPara.private_GetImpl());
		}
		else 
			return null;

		if (beRNewPara === true)
			return oNewPara;
		else 
			return this;
	};
	/**
	 * Select a paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {bool} 
	 */
	ApiParagraph.prototype.Select = function()
	{
		var Document = private_GetLogicDocument();
		
		var StartRun	= this.Paragraph.GetFirstRun();
		var StartPos	= StartRun.GetDocumentPositionFromObject();
		var EndRun		= this.Paragraph.Content[this.Paragraph.Content.length - 1];
		var EndPos		= EndRun.GetDocumentPositionFromObject();
		
		StartPos.push({Class: StartRun, Position: 0});
		EndPos.push({Class: EndRun, Position: 1});

		if (StartPos[0].Position === - 1)
			return false;

		StartPos[0].Class.SetSelectionByContentPositions(StartPos, EndPos);

		var controllerType = null;

		if (StartPos[0].Class.IsHdrFtr())
		{
			controllerType = docpostype_HdrFtr;
		}
		else if (StartPos[0].Class.IsFootnote())
		{
			controllerType = docpostype_Footnotes;
		}
		else if (StartPos[0].Class.Is_DrawingShape())
		{
			controllerType = docpostype_DrawingObjects;
		}
		else 
		{
			controllerType = docpostype_Content;
		}
		
		Document.SetDocPosType(controllerType);
		Document.UpdateSelection();

		return true;	
	};
	/**
	 * Search for the scope of a paragraph object. The search results are a collection of ApiRange objects.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {bool} isMatchCase - Case sensitive or not. 
	 * @return {ApiRange[]}  
	 */
	ApiParagraph.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase = false;

		var arrApiRanges	= [];
		var Api				= editor; 
		var oDocument		= Api.GetDocument();
		var SearchEngine	= null;

		if (!oDocument.Document.SearchEngine.Compare(sText, {MatchCase: isMatchCase}))
		{
			SearchEngine		= new CDocumentSearch();
			SearchEngine.Set(sText, {MatchCase: isMatchCase});
			this.Paragraph.Search(sText, {MatchCase: isMatchCase}, SearchEngine, 0)
		}
		else 
			SearchEngine = oDocument.Document.SearchEngine;

		var SearchResults	= this.Paragraph.SearchResults;

		for (var FoundId in SearchResults)
		{
			var StartSearchContentPos	= SearchResults[FoundId].StartPos;
			var EndSearchContentPos		= SearchResults[FoundId].EndPos;

			var StartChar	= this.Paragraph.ConvertParaContentPosToRangePos(StartSearchContentPos);
			var EndChar		= this.Paragraph.ConvertParaContentPosToRangePos(EndSearchContentPos);
			if (EndChar > 0)
				EndChar--;

			arrApiRanges.push(this.GetRange(StartChar, EndChar));
		}

		return arrApiRanges;
	};
	/**
	 * Wrap paragraph content in a mail merge field.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 */
	ApiParagraph.prototype.WrapInMailMergeField = function()
	{
		var oDocument = private_GetLogicDocument();
		var fieldName = this.GetText();
		var oField    = new ParaField(fieldtype_MERGEFIELD, [fieldName], []);
		
		var leftQuote  = new ParaRun();
		var rightQuote = new ParaRun();

		leftQuote.AddText("«");
		rightQuote.AddText("»");

		oField.Add_ToContent(0, leftQuote);

		for (var nElement = 0; nElement < this.Paragraph.Content.length; nElement++)
		{
			oField.Add_ToContent(nElement + 1, this.Paragraph.Content[nElement].Copy())
		}
	
		oField.Add_ToContent(oField.Content.length, rightQuote);
		
		this.RemoveAllElements();
		oDocument.Register_Field(oField);
		this.Paragraph.AddToParagraph(oField);
	};
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRun
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"run"}
	 */
	ApiRun.prototype.GetClassType = function()
	{
		return "run";
	};
	/**
	 * Get the text properties of the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.Run.Pr.Copy());
	};
	/**
	 * Clear content from the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.ClearContent = function()
	{
		this.Run.Remove_FromContent(0, this.Run.Content.length);
	};
	/**
	 * Remove all the elements from the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.RemoveAllElements = function()
	{
		this.Run.Remove_FromContent(0, this.Run.Content.length);
	};
	/**
	 * Delete current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.Delete = function()
	{
		var parentOfElement = this.Run.Get_Parent();

		var PosInParent = parentOfElement.Content.indexOf(this.Run);

		if (PosInParent !== - 1)
		{
			this.Run.PreDelete();
			parentOfElement.Remove_FromContent(PosInParent, 1, true);
		}
		else 
			return false;
	};
	/**
	 * Add some text to this run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} sText - The text which will be added to the current run.
	 */
	ApiRun.prototype.AddText = function(sText)
	{
		if (!sText || !sText.length)
			return;

		this.Run.AddText(sText);
	};
	/**
	 * Add a page break and start the next element from a new page.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 */
	ApiRun.prototype.AddPageBreak = function()
	{
		this.Run.Add_ToContent(this.Run.Content.length, new ParaNewLine(break_Page));
	};
	/**
	 * Add a line break to the current run position and start the next element from a new line.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.AddLineBreak = function()
	{
		this.Run.Add_ToContent(this.Run.Content.length, new ParaNewLine(break_Line));
	};
	/**
	 * Add a column break to the current run position and start the next element from a new column.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 */
	ApiRun.prototype.AddColumnBreak = function()
	{
		this.Run.Add_ToContent(this.Run.Content.length, new ParaNewLine(break_Column));
	};
	/**
	 * Add a tab stop to the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.AddTabStop = function()
	{
		this.Run.Add_ToContent(this.Run.Content.length, new ParaTab());
	};
	/**
	 * Add a drawing object (image, shape or chart) to the current text run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @param {ApiDrawing} oDrawing - The object which will be added to the current run.
	 * @returns {bool} - returns false if param is invalid.
	 */ 
	ApiRun.prototype.AddDrawing = function(oDrawing)
	{
		if (!(oDrawing instanceof ApiDrawing))
			return false;

		this.Run.Add_ToContent(this.Run.Content.length, oDrawing.Drawing);
		oDrawing.Drawing.Set_Parent(this.Run);

		return true;
	};
	/**
	 * Select a run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @return {bool} 
	 */
	ApiRun.prototype.Select = function()
	{
		var Document = private_GetLogicDocument();

		var StartPos		= this.Run.GetDocumentPositionFromObject();
		var EndPos			= this.Run.GetDocumentPositionFromObject();
		var parentParagraph	= this.Run.GetParagraph();

		if (!parentParagraph)
			return false;

		StartPos.push({Class: this.Run, Position: 0});
		EndPos.push({Class: this.Run, Position: this.Run.Content.length});

		if (StartPos[0].Position === - 1)
			return false;

		StartPos[0].Class.SetSelectionByContentPositions(StartPos, EndPos);

		var controllerType = null;

		if (StartPos[0].Class.IsHdrFtr())
		{
			controllerType = docpostype_HdrFtr;
		}
		else if (StartPos[0].Class.IsFootnote())
		{
			controllerType = docpostype_Footnotes;
		}
		else if (StartPos[0].Class.Is_DrawingShape())
		{
			controllerType = docpostype_DrawingObjects;
		}
		else 
		{
			controllerType = docpostype_Content;
		}

		Document.SetDocPosType(controllerType);
		Document.UpdateSelection();

		return true;	
	};
	/**
	 * Add a hyperlink to a run. 
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @param {string} sLink - The link address.
	 * @param {string} sScreenTipText - The screen tip text.
	 * @return {ApiHyperlink | null} - returns false if params are invalid.
	 */
	ApiRun.prototype.AddHyperlink = function(sLink, sScreenTipText)
	{
		if (typeof(sLink) !== "string" || sLink === "" || sLink.length > Asc.c_nMaxHyperlinkLength)
			return null;
		if (typeof(sScreenTipText) !== "string")
			sScreenTipText = "";

		var Document	= editor.private_GetLogicDocument();
		var parentPara	= this.Run.GetParagraph();
		if (!parentPara || this.Run.Content.length === 0)
			return null;
		if (this.GetParentContentControl() instanceof ApiInlineLvlSdt)
			return null;

		function find_parentParaDepth(DocPos)
		{
			for (var nPos = 0; nPos < DocPos.length; nPos++)
			{
				if (DocPos[nPos].Class instanceof Paragraph && DocPos[nPos].Class.Id === parentPara.Id)
				{
					return nPos;
				}
			}
		}

		var StartPos		= this.Run.GetDocumentPositionFromObject();
		var EndPos			= this.Run.GetDocumentPositionFromObject();
		StartPos.push({Class: this.Run, Position: 0});
		EndPos.push({Class: this.Run, Position: this.Run.Content.length});
		var parentParaDepth = find_parentParaDepth(StartPos);
		StartPos[parentParaDepth].Class.SetContentSelection(StartPos, EndPos, parentParaDepth, 0, 0);

		var oHyperlink	= null;
		var hyperlinkPr	= new Asc.CHyperlinkProperty();
		var urlType		= AscCommon.getUrlType(sLink);
		if (!/(((^https?)|(^ftp)):\/\/)|(^mailto:)/i.test(sLink))
			sLink = (urlType === 0) ? null :(( (urlType === 2) ? 'mailto:' : 'http://' ) + sLink);
		sLink = sLink.replace(new RegExp("%20",'g')," ");
		hyperlinkPr.put_Value(sLink);
		hyperlinkPr.put_ToolTip(sScreenTipText);
		hyperlinkPr.put_Bookmark(null);

		parentPara.Selection.Use = true;
		oHyperlink = new ApiHyperlink(parentPara.AddHyperlink(hyperlinkPr));
		Document.RemoveSelection();

		return oHyperlink;
	};
	/**
	 * Create a copy of the run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiRun}
	 */
	ApiRun.prototype.Copy = function()
	{
		var oRun = this.Run.Copy(false, {
			SkipComments          : true,
			SkipAnchors           : true,
			SkipFootnoteReference : true,
			SkipComplexFields     : true
		});

		return new ApiRun(oRun);
	};
	/**
	 * Return a Range object that represents the part of the document contained in the specified run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiRun.prototype.GetRange = function(Start, End)
	{
		var Range = new ApiRange(this.Run, Start, End);

		return Range;
	};

	/**
     * Get the content control that contains the run.
     * @memberof ApiRun
	 * @typeofeditors ["CDE"]
     * @return {ApiBlockLvlSdt | ApiInlineLvlSdt | null} - returns null if parent content control doesn't exist.  
     */
    ApiRun.prototype.GetParentContentControl = function()
    {
        var RunPosition = this.Run.GetDocumentPositionFromObject();

        for (var Index = RunPosition.length - 1; Index >= 1; Index--)
        {
            if (RunPosition[Index].Class)
            {
                if (RunPosition[Index].Class instanceof CBlockLevelSdt)
                    return new ApiBlockLvlSdt(RunPosition[Index].Class);
                else if (RunPosition[Index].Class instanceof CInlineLevelSdt)
                    return new ApiInlineLvlSdt(RunPosition[Index].Class);
            }
        }

        return null;
    };
    /**
     * Get the table that contains the run.
     * @memberof ApiRun
	 * @typeofeditors ["CDE"]
     * @return {ApiTable | null} - returns null if parent table doesn't exist.
     */
    ApiRun.prototype.GetParentTable = function()
    {
        var documentPos = this.Run.GetDocumentPositionFromObject();

        for (var Index = documentPos.length - 1; Index >= 1; Index--)
        {
            if (documentPos[Index].Class)
                if (documentPos[Index].Class instanceof CTable)
                    return new ApiTable(documentPos[Index].Class);
        }

        return null;
    };
    /**
     * Get the table cell that contains the run.
     * @memberof ApiRun
	 * @typeofeditors ["CDE"]
     * @return {ApiTableCell | null} - returns null is parent cell doesn't exist.  
     */
    ApiRun.prototype.GetParentTableCell = function()
    {
        var documentPos = this.Run.GetDocumentPositionFromObject();

        for (var Index = documentPos.length - 1; Index >= 1; Index--)
        {
            if (documentPos[Index].Class.Parent)
                if (documentPos[Index].Class.Parent instanceof CTableCell)
                    return new ApiTableCell(documentPos[Index].Class.Parent);
        }

        return null;
    };
	/**
	 * Set the text properties for the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ApiTextPr} oTextPr - The text properties that will be set to the run specified.
	 * @return {ApiTextPr}  
	 */
	ApiRun.prototype.SetTextPr = function(oTextPr)
	{
		var runTextPr = this.GetTextPr();
		runTextPr.TextPr.Merge(oTextPr.TextPr);
		runTextPr.private_OnChange();

		return runTextPr;
	};
	/**
	 * Set the bold property to the text character.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isBold - Specifies that the contents of this run are displayed bold.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetBold = function(isBold)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetBold(isBold);
		
		return oTextPr;
	};
	/**
	 * Specify that any lowercase characters in this text run are formatted for display only as their capital letter character equivalents.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isCaps - Specifies that the contents of the current run are displayed capitalized.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetCaps = function(isCaps)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetCaps(isCaps);
		
		return oTextPr;
	};
	/**
	 * Set the text color for the current text run in the RGB format.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - If this parameter is set to "true", then r,g,b parameters will be ignored.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetColor = function(r, g, b, isAuto)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetColor(r, g, b, isAuto);
		
		return oTextPr;
	};
	/**
	 * Specify that the contents of this run are displayed with two horizontal lines through each character displayed on the line.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isDoubleStrikeout - Specifies that the contents of the current run are displayed double struck through.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetDoubleStrikeout = function(isDoubleStrikeout)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetDoubleStrikeout(isDoubleStrikeout);
		
		return oTextPr;
	};
	/**
	 * Set the text color for the current text run.
	 * @memberof ApiRun
	 * @typeofeditors ["CSE", "CPE"]
	 * @param {ApiFill} oApiFill - The color or pattern used to fill the text color.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetFill = function(oApiFill)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetFill(oApiFill);
		
		return oTextPr;
	};
	/**
	 * Set all 4 font slots with the specified font family.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} sFontFamily - The font family or families used for the current text run.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetFontFamily = function(sFontFamily)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetFontFamily(sFontFamily);
		
		return oTextPr;
	};
	/**
	 * Set the font size for the characters of the current text run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {hps} nSize - The text size value measured in half-points (1/144 of an inch).
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetFontSize = function(nSize)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetFontSize(nSize);
		
		return oTextPr;
	};
	/**
	 * Specify a highlighting color in the RGB format which is applied as a background for the contents of the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isNone=false] If this parameter is set to "true", then r,g,b parameters will be ignored.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetHighlight = function(r, g, b, isNone)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetHighlight(r, g, b, isNone);
		
		return oTextPr;
	};
	/**
	 * Set the italic property to the text character.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isItalic - Specifies that the contents of the current run are displayed italicized.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetItalic = function(isItalic)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetItalic(isItalic);
		
		return oTextPr;
	};
	/**
	 * Specify the languages which will be used to check spelling and grammar (if requested) when processing
	 * the contents of this text run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} sLangId - The possible value for this parameter is a language identifier as defined by
	 * RFC 4646/BCP 47. Example: "en-CA".
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetLanguage = function(sLangId)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetLanguage(sLangId);
		
		return oTextPr;
	};
	/**
	 * Specify the amount by which text is raised or lowered for this run in relation to the default
	 * baseline of the surrounding non-positioned text.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {hps} nPosition - Specifies a positive (raised text) or negative (lowered text)
	 * measurement in half-points (1/144 of an inch).
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetPosition = function(nPosition)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetPosition(nPosition);
		
		return oTextPr;
	};
	/**
	 * Specify the shading applied to the contents of the current text run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ShdType} sType - The shading type applied to the contents of the current text run.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetShd = function(sType, r, g, b)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetShd(sType, r, g, b);
		
		return oTextPr;
	};
	/**
	 * Specify that all the small letter characters in this text run are formatted for display only as their capital
	 * letter character equivalents which are two points smaller than the actual font size specified for this text.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isSmallCaps - Specifies if the contents of the current run are displayed capitalized two points smaller or not.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetSmallCaps = function(isSmallCaps)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetSmallCaps(isSmallCaps);
		
		return oTextPr;
	};
	/**
	 * Set text spacing measured in twentieths of a point.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nSpacing - The value of the text spacing measured in twentieths of a point (1/1440 of an inch).
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetSpacing = function(SetSpacing)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetSpacing(SetSpacing);
		
		return oTextPr;
	};
	/**
	 * Specify that the contents of this run are displayed with a single horizontal line through the center of the line.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isStrikeout - Specifies that the contents of the current run are displayed struck through.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetStrikeout = function(isStrikeout)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetStrikeout(isStrikeout);
		
		return oTextPr;
	};
	/**
	 * Set style for the current Run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ApiStyle} oStyle - The style which must be applied to the text character.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetStyle = function(oStyle)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetStyle(oStyle);
		
		return oTextPr;
	};
	/**
	 * Specify that the contents of this run are displayed along with a line appearing directly below the character
	 * (less than all the spacing above and below the characters on the line).
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isUnderline - Specifies that the contents of the current run are displayed underlined.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetUnderline = function(isUnderline)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetUnderline(isUnderline);
		
		return oTextPr;
	};
	/**
	 * Specify the alignment which will be applied to the contents of this run in relation to the default appearance of the run text:
	 * * <b>"baseline"</b> - the characters in the current text run will be aligned by the default text baseline.
	 * * <b>"subscript"</b> - the characters in the current text run will be aligned below the default text baseline.
	 * * <b>"superscript"</b> - the characters in the current text run will be aligned above the default text baseline.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {("baseline" | "subscript" | "superscript")} sType - The vertical alignment type applied to the text contents.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetVertAlign = function(sType)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetVertAlign(sType);
		
		return oTextPr;
	};
	/**
	 * Wrap run in a mail merge field.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 */
	ApiRun.prototype.WrapInMailMergeField = function()
	{
		var oDocument = private_GetLogicDocument();
		var fieldName = this.Run.GetText();
		var oField    = new ParaField(fieldtype_MERGEFIELD, [fieldName], []);
		var runParent = this.Run.GetParent();

		var leftQuote  = new ParaRun();
		var rightQuote = new ParaRun();

		leftQuote.AddText("«");
		rightQuote.AddText("»");

		oField.Add_ToContent(0, leftQuote);
		oField.Add_ToContent(1, this.Run);
		oField.Add_ToContent(oField.Content.length, rightQuote);

		if (runParent)
		{
			var indexInParent = runParent.Content.indexOf(this.Run);
			runParent.Remove_FromContent(indexInParent, 1);
			runParent.Add_ToContent(indexInParent, oField);
		}
		
		oDocument.Register_Field(oField);
	};
	/**
	 * Gets the next run if exist.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @return {ApiRun | null} - returns null if next run doesn't exist.
	 */
	ApiRun.prototype.GetNext = function()
	{
		var oParent = this.Run.Parent || this.Run.Paragraph;
		var oRunType = this.Run.Get_Type();
		if (!oParent)
			return null;

		var oRunIndex = oParent.Content.indexOf(this.Run);
		if (oParent.Content[oRunIndex + 1] && oParent.Content[oRunIndex + 1].Type === oRunType && oRunIndex + 1 !== oParent.Content.length - 1)
			return new ApiRun(oParent.Content[oRunIndex + 1]);
		return null;
	};
	/**
	 * Gets the previous run if exist.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @return {ApiRun | null} - returns null if previous run doesn't exist.
	 */
	ApiRun.prototype.GetPrevious = function()
	{
		var oParent = this.Run.Parent || this.Run.Paragraph;
		var oRunType = this.Run.Get_Type();
		if (!oParent)
			return null;

		var oRunIndex = oParent.Content.indexOf(this.Run);
		if (oParent.Content[oRunIndex - 1] && oParent.Content[oRunIndex - 1].Type === oRunType && oRunIndex - 1 !== oParent.Content.length - 1)
			return new ApiRun(oParent.Content[oRunIndex - 1]);
		return null;
	};
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiSection
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @returns {"section"}
	 */
	ApiSection.prototype.GetClassType = function()
	{
		return "section";
	};
	/**
	 * Specify the type of the current section. The section type defines how the contents of the current 
	 * section are placed relative to the previous section.<br/>
	 * WordprocessingML supports five distinct types of section breaks:
	 *   * <b>Next page</b> section breaks (the default if type is not specified), which begin the new section on the
	 *   following page.
	 *   * <b>Odd</b> page section breaks, which begin the new section on the next odd-numbered page.
	 *   * <b>Even</b> page section breaks, which begin the new section on the next even-numbered page.
	 *   * <b>Continuous</b> section breaks, which begin the new section on the following paragraph. This means that
	 *   continuous section breaks might not specify certain page-level section properties, since they shall be
	 *   inherited from the following section. These breaks, however, can specify other section properties, such
	 *   as line numbering and footnote/endnote settings.
	 *   * <b>Column</b> section breaks, which begin the new section on the next column on the page.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {("nextPage" | "oddPage" | "evenPage" | "continuous" | "nextColumn")} sType - The section break type.
	 */
	ApiSection.prototype.SetType = function(sType)
	{
		if ("oddPage" === sType)
			this.Section.Set_Type(c_oAscSectionBreakType.OddPage);
		else if ("evenPage" === sType)
			this.Section.Set_Type(c_oAscSectionBreakType.EvenPage);
		else if ("continuous" === sType)
			this.Section.Set_Type(c_oAscSectionBreakType.Continuous);
		else if ("nextColumn" === sType)
			this.Section.Set_Type(c_oAscSectionBreakType.Column);
		else if ("nextPage" === sType)
			this.Section.Set_Type(c_oAscSectionBreakType.NextPage);
	};
	/**
	 * Specify that all the text columns in the current section are of equal width.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {number} nCount - Number of columns.
	 * @param {twips} nSpace - Distance between columns measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiSection.prototype.SetEqualColumns = function(nCount, nSpace)
	{
		this.Section.Set_Columns_EqualWidth(true);
		this.Section.Set_Columns_Num(nCount);
		this.Section.Set_Columns_Space(private_Twips2MM(nSpace));
	};
	/**
	 * Specify that all the columns in the current section are of a different width. Number of columns is equal 
	 * to the length of the aWidth array. The length of the aSpaces array MUST BE equal to (aWidth.length - 1).
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {twips[]} aWidths - An array of column width values measured in twentieths of a point (1/1440 of an inch).
	 * @param {twips[]} aSpaces - An array of distance values between the columns measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiSection.prototype.SetNotEqualColumns = function(aWidths, aSpaces)
	{
		if (!aWidths || !aWidths.length || aWidths.length <= 1 || aSpaces.length !== aWidths.length - 1)
			return;

		this.Section.Set_Columns_EqualWidth(false);
		var aCols = [];
		for (var nPos = 0, nCount = aWidths.length; nPos < nCount; ++nPos)
		{
			var SectionColumn   = new CSectionColumn();
			SectionColumn.W     = private_Twips2MM(aWidths[nPos]);
			SectionColumn.Space = private_Twips2MM(nPos !== nCount - 1 ? aSpaces[nPos] : 0);
			aCols.push(SectionColumn);
		}

		this.Section.Set_Columns_Cols(aCols);
		this.Section.Set_Columns_Num(aCols.length);
	};
	/**
	 * Specify the properties (size and orientation) for all the pages in the current section.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {twips} nWidth - The page width measured in twentieths of a point (1/1440 of an inch).
	 * @param {twips} nHeight - The page height measured in twentieths of a point (1/1440 of an inch).
	 * @param {boolean} [isPortrait=false] - Specifies the orientation of all the pages in this section (if set to true, then the portrait orientation is chosen).
	 */
	ApiSection.prototype.SetPageSize = function(nWidth, nHeight, isPortrait)
	{
		this.Section.SetPageSize(private_Twips2MM(nWidth), private_Twips2MM(nHeight));
		this.Section.SetOrientation(false === isPortrait ? Asc.c_oAscPageOrientation.PageLandscape : Asc.c_oAscPageOrientation.PagePortrait, false);
	};
	/**
	 * Specify the page margins for all the pages in this section.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {twips} nLeft - The left margin width measured in twentieths of a point (1/1440 of an inch).
	 * @param {twips} nTop - The top margin height measured in twentieths of a point (1/1440 of an inch).
	 * @param {twips} nRight - The right margin width measured in twentieths of a point (1/1440 of an inch).
	 * @param {twips} nBottom - The bottom margin height measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiSection.prototype.SetPageMargins = function(nLeft, nTop, nRight, nBottom)
	{
		this.Section.SetPageMargins(private_Twips2MM(nLeft), private_Twips2MM(nTop), private_Twips2MM(nRight), private_Twips2MM(nBottom));
	};
	/**
	 * Specify the distance from the top edge of the page to the top edge of the header.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {twips} nDistance - The distance from the top edge of the page to the top edge of the header measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiSection.prototype.SetHeaderDistance = function(nDistance)
	{
		this.Section.SetPageMarginHeader(private_Twips2MM(nDistance));
	};
	/**
	 * Specify the distance from the bottom edge of the page to the bottom edge of the footer.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {twips} nDistance - The distance from the bottom edge of the page to the bottom edge of the footer measured
	 * in twentieths of a point (1/1440 of an inch).
	 */
	ApiSection.prototype.SetFooterDistance = function(nDistance)
	{
		this.Section.SetPageMarginFooter(private_Twips2MM(nDistance));
	};
	/**
	 * Get the content for the specified header type.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {HdrFtrType} sType - Header type to get the content from.
	 * @param {boolean} [isCreate=false] - Specifies whether to create a new header or not with the specified header type in case
	 * no header with such a type could be found in the current section.
	 * @returns {?ApiDocumentContent}
	 */
	ApiSection.prototype.GetHeader = function(sType, isCreate)
	{
		var oHeader = null;

		if ("title" === sType)
			oHeader = this.Section.Get_Header_First();
		else if ("even" === sType)
			oHeader = this.Section.Get_Header_Even();
		else if ("default" === sType)
			oHeader = this.Section.Get_Header_Default();
		else
			return null;

		if (null === oHeader && true === isCreate)
		{
			var oLogicDocument = private_GetLogicDocument();
			oHeader            = new CHeaderFooter(oLogicDocument.GetHdrFtr(), oLogicDocument, oLogicDocument.Get_DrawingDocument(), hdrftr_Header);
			if ("title" === sType)
				this.Section.Set_Header_First(oHeader);
			else if ("even" === sType)
				this.Section.Set_Header_Even(oHeader);
			else if ("default" === sType)
				this.Section.Set_Header_Default(oHeader);
		}
		if(!oHeader){
			return null;
		}
		return new ApiDocumentContent(oHeader.Get_DocumentContent());
	};
	/**
	 * Remove the header of the specified type from the current section. After removal the header will be inherited from
	 * the previous section or, if this is the first section in the document, no header of the specified type will be present.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {HdrFtrType} sType - Header type to be removed.
	 */
	ApiSection.prototype.RemoveHeader = function(sType)
	{
		if ("title" === sType)
			this.Section.Set_Header_First(null);
		else if ("even" === sType)
			this.Section.Set_Header_Even(null);
		else if ("default" === sType)
			this.Section.Set_Header_Default(null);
	};
	/**
	 * Get the content for the specified footer type.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {HdrFtrType} sType - Footer type to get the content from.
	 * @param {boolean} [isCreate=false] - Specifies whether to create a new footer or not with the specified footer type in case
	 * no footer with such a type could be found in the current section.
	 * @returns {?ApiDocumentContent}
	 */
	ApiSection.prototype.GetFooter = function(sType, isCreate)
	{
		var oFooter = null;

		if ("title" === sType)
			oFooter = this.Section.Get_Footer_First();
		else if ("even" === sType)
			oFooter = this.Section.Get_Footer_Even();
		else if ("default" === sType)
			oFooter = this.Section.Get_Footer_Default();
		else
			return null;

		if (null === oFooter && true === isCreate)
		{
			var oLogicDocument = private_GetLogicDocument();
			oFooter            = new CHeaderFooter(oLogicDocument.GetHdrFtr(), oLogicDocument, oLogicDocument.Get_DrawingDocument(), hdrftr_Footer);
			if ("title" === sType)
				this.Section.Set_Footer_First(oFooter);
			else if ("even" === sType)
				this.Section.Set_Footer_Even(oFooter);
			else if ("default" === sType)
				this.Section.Set_Footer_Default(oFooter);
		}

		return new ApiDocumentContent(oFooter.Get_DocumentContent());
	};
	/**
	 * Remove the footer of the specified type from the current section. After removal the footer will be inherited from 
	 * the previous section or, if this is the first section in the document, no footer of the specified type will be present.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {HdrFtrType} sType - Type of footer.
	 */
	ApiSection.prototype.RemoveFooter = function(sType)
	{
		if ("title" === sType)
			this.Section.Set_Footer_First(null);
		else if ("even" === sType)
			this.Section.Set_Footer_Even(null);
		else if ("default" === sType)
			this.Section.Set_Footer_Default(null);
	};
	/**
	 * Specify whether the current section in this document have different header and footer for the section first page.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isTitlePage - If true the first page of the section will have header and footer that will differ from the other pages of the same section.
	 */
	ApiSection.prototype.SetTitlePage = function(isTitlePage)
	{
		this.Section.Set_TitlePage(private_GetBoolean(isTitlePage));
	};
	/**
	 * Get the next section if exists.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @returns {ApiSection | null} - returns null if section is last.
	 */
	ApiSection.prototype.GetNext = function()
	{
		var oDocument		= editor.GetDocument();
		var arrApiSections	= oDocument.GetSections();
		var sectionIndex	= -1;

		for (var nSection = 0; nSection < arrApiSections.length; nSection++)
		{
			if (arrApiSections[nSection].Section.Id === this.Section.Id) 
			{
				sectionIndex = nSection;
				break;
			}
		}
		
		if (sectionIndex !== - 1 && arrApiSections[sectionIndex + 1])
		{
			return arrApiSections[sectionIndex + 1];
		}

		return null;
	};
	/**
	 * Get the preious section if exists.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @returns {ApiSection | null} - returns null if section is first.
	 */
	ApiSection.prototype.GetPrevious = function()
	{
		var oDocument		= editor.GetDocument();
		var arrApiSections	= oDocument.GetSections();
		var sectionIndex	= -1;

		for (var nSection = 0; nSection < arrApiSections.length; nSection++)
		{
			if (arrApiSections[nSection].Section.Id === this.Section.Id) 
			{
				sectionIndex = nSection;
				break;
			}
		}
		
		if (sectionIndex !== - 1 && arrApiSections[sectionIndex - 1])
		{
			return arrApiSections[sectionIndex - 1];
		}

		return null;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTable
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @returns {"table"}
	 */
	ApiTable.prototype.GetClassType = function()
	{
		return "table";
	};
	/**
	 * Get the number of rows in the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiTable.prototype.GetRowsCount = function()
	{
		return this.Table.Content.length;
	};
	/**
	 * Get the table row by its position in the table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The row position within the table.
	 * @returns {ApiTableRow | null} - returns null if param is invalid.
	 */
	ApiTable.prototype.GetRow = function(nPos)
	{
		if (nPos < 0 || nPos >= this.Table.Content.length)
			return null;

		return new ApiTableRow(this.Table.Content[nPos]);
	};
	/**
	 * Get the cell by its position.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {number} nRow - The row position in the current table.
	 * @param {number} Cell - The cell position in the current table.
	 * @returns {ApiTableCell | null} - returns null if params are invalid.
	 */
	ApiTable.prototype.GetCell = function(nRow, nCell)
	{
		var Row = this.Table.GetRow(nRow);

		if (Row && nCell >= 0 && nCell <= Row.Content.length)
		{
			return new ApiTableCell(Row.GetCell(nCell));
		}
		else 
			return null;
	};
	/**
	 * Merge an array of cells. If the merge is done successfully, it will return the resulting merged cell, otherwise the result will be "null".
	 * <note>The number of cells in any row and the number of rows in the current table may be changed.</note>
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell[]} aCells - The array of cells to be merged.
	 * @returns {?ApiTableCell}
	 */
	ApiTable.prototype.MergeCells = function(aCells)
	{
		private_StartSilentMode();
		this.private_PrepareTableForActions();

		var oTable            = this.Table;
		oTable.Selection.Use  = true;
		oTable.Selection.Type = table_Selection_Cell;
		oTable.Selection.Data = [];

		for (var nPos = 0, nCount = aCells.length; nPos < nCount; ++nPos)
		{
			var oCell = aCells[nPos].Cell;
			var oPos  = {Cell : oCell.Index, Row : oCell.Row.Index};

			var nResultPos    = 0;
			var nResultLength = oTable.Selection.Data.length;
			for (nResultPos = 0; nResultPos < nResultLength; ++nResultPos)
			{
				var oCurPos = oTable.Selection.Data[nResultPos];
				if (oCurPos.Row < oPos.Row)
				{
					continue;
				}
				else if (oCurPos.Row > oPos.Row)
				{
					break;
				}
				else
				{
					if (oCurPos.Cell < oPos.Cell)
						continue;
					else
						break;
				}
			}

			oTable.Selection.Data.splice(nResultPos, 0, oPos);
		}

		var isMerged = this.Table.MergeTableCells(true);
		var oMergedCell = this.Table.CurCell;
		oTable.RemoveSelection();

		private_EndSilentMode();

		if (true === isMerged)
			return new ApiTableCell(oMergedCell);

		return null;
	};
	/**
	 * Set the style for the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiStyle} oStyle - The style which will be applied to the current table.
	 * @returns {bool} - returns false if param is invalid.
	 */
	ApiTable.prototype.SetStyle = function(oStyle)
	{
		if (!oStyle || !(oStyle instanceof ApiStyle) || styletype_Table !== oStyle.Style.Get_Type())
			return false;

		this.Table.Set_TableStyle(oStyle.Style.Get_Id(), true);

		return true;
	};
	/**
	 * Specify the conditional formatting components of the referenced table style (if one exists) 
	 * which will be applied to the set of table rows with the current table-level property exceptions. A table style 
	 * can specify up to six different optional conditional formats, for example, different formatting for first column, 
	 * which then can be applied or omitted from individual table rows in the parent table.
	 * 
	 * The default setting is to apply the row and column band formatting, but not the first row, last row, first 
	 * column, or last column formatting.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isFirstColumn - Specifies that the first column conditional formatting will be applied to the table.
	 * @param {boolean} isFirstRow - Specifies that the first row conditional formatting will be applied to the table.
	 * @param {boolean} isLastColumn - Specifies that the last column conditional formatting will be applied to the table.
	 * @param {boolean} isLastRow - Specifies that the last row conditional formatting will be applied to the table.
	 * @param {boolean} isHorBand - Specifies that the horizontal band conditional formatting will not be applied to the table.
	 * @param {boolean} isVerBand - Specifies that the vertical band conditional formatting will not be applied to the table.
	 */
	ApiTable.prototype.SetTableLook = function(isFirstColumn, isFirstRow, isLastColumn, isLastRow, isHorBand, isVerBand)
	{
		var oTableLook = new CTableLook(private_GetBoolean(isFirstColumn),
			private_GetBoolean(isFirstRow),
			private_GetBoolean(isLastColumn),
			private_GetBoolean(isLastRow),
			private_GetBoolean(isHorBand),
			private_GetBoolean(isVerBand));
		this.Table.Set_TableLook(oTableLook);
	};
	/**
	 * Split the cell into a given number of rows and columns.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell which will be split.
	 * @param {Number} [nRow=1] - Count of rows which the cell will be split.
	 * @param {Number} [nCol=1] - Count of columns which the cell will be split.
	 * @returns {ApiTable | null} - returns null if can't split.
	 */
	ApiTable.prototype.Split = function(oCell, nRow, nCol)
	{
		if (nRow == undefined)
			nRow = 1;
		if (nCol == undefined)
			nCol = 1;

		if (!(oCell instanceof ApiTableCell) || nCol <= 0 || nRow <= 0)
			return null;

		this.Table.RemoveSelection();
		this.Table.Set_CurCell(oCell.Cell);
		this.Table.SelectTable(c_oAscTableSelectionType.Cell);

		if (!this.Table.CanSplitTableCells())
			return null;

		if (!this.Table.IsRecalculated())
		{
			// Reset делаем для случая, когда таблица вообще ни разу не пересчитывалась
			this.Table.Reset(0, 0, 100, 100, 0, 0, 1);
			this.Table.Recalculate_Grid();
		}

		if (!this.Table.SplitTableCells(nCol, nRow, false))
			return null;

		return this;
	};
	/**
	 * Add a new row to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell after which a new row will be added. If not specified, a new row will
	 * be added at the end of the table.
	 * @param {boolean} [isBefore=false] - Add a new row before or after the specified cell. If no cell is specified, then
	 * this parameter will be ignored.
	 * @returns {ApiTableRow}
	 */
	ApiTable.prototype.AddRow = function(oCell, isBefore)
	{
		private_StartSilentMode();
		this.private_PrepareTableForActions();

		var _isBefore = private_GetBoolean(isBefore, false);
		var _oCell = (oCell instanceof ApiTableCell ? oCell.Cell : undefined);
		if (_oCell && this.Table !== _oCell.Row.Table)
			_oCell = undefined;

		if (!_oCell)
		{
			_oCell = this.Table.Content[this.Table.Content.length - 1].Get_Cell(0);
			_isBefore = false;
		}

		var nRowIndex = true === _isBefore ? _oCell.Row.Index : _oCell.Row.Index + 1;

		this.Table.RemoveSelection();
		this.Table.CurCell = _oCell;
		this.Table.AddTableRow(_isBefore);

		private_EndSilentMode();
		return new ApiTableRow(this.Table.Content[nRowIndex]);
	};
	/**
	 * Add the new rows to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell after which the new rows will be added. If not specified, the new rows will
	 * be added at the end of the table.
	 * @param {Number} nCount - Count of rows to be added.
	 * @param {boolean} [isBefore=false] - Add a new rows before or after the specified cell. If no cell is specified then
	 * this parameter will be ignored.
	 * @returns {ApiTable}
	 */
	ApiTable.prototype.AddRows = function(oCell, nCount, isBefore)
	{
		for (var Index = 0; Index < nCount; Index++)
		{
			this.AddRow(oCell, isBefore);
		}

		return this;
	};
	/**
	 * Add a new column to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell after which a new column will be added. If not specified, a new column will be added at the end of the table.
	 * @param {boolean} [isBefore=false] - Add a new column before or after the specified cell. If no cell is specified,
	 * then this parameter will be ignored.
	 */
	ApiTable.prototype.AddColumn = function(oCell, isBefore)
	{
		private_StartSilentMode();
		this.private_PrepareTableForActions();

		var _isBefore = private_GetBoolean(isBefore, false);
		var _oCell = (oCell instanceof ApiTableCell ? oCell.Cell : undefined);
		if (_oCell && this.Table !== _oCell.Row.Table)
			_oCell = undefined;

		if (!_oCell)
		{
			_oCell = this.Table.Content[0].Get_Cell(this.Table.Content[0].Get_CellsCount() - 1);
			_isBefore = false;
		}

		this.Table.RemoveSelection();
		this.Table.CurCell = _oCell;
		this.Table.AddTableColumn(_isBefore);

		private_EndSilentMode();
	};
	/**
	 * Add the new columns to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell after which the new columns will be added. If not specified, the new columns will be added at the end of the table.
	 * @param {Number} nCount - Count of columns to be added.
	 * @param {boolean} [isBefore=false] - Add the new columns before or after the specified cell. If no cell is specified,
	 * then this parameter will be ignored.
	 */
	ApiTable.prototype.AddColumns = function(oCell, nCount, isBefore)
	{
		for (var Index = 0; Index < nCount; Index++)
		{
			this.AddColumn(oCell, isBefore);
		}

		return this;
	};
	/**
	 * Add a paragraph or a table or a blockLvl content control using its position in the cell.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ApiTableCell} oCell - The cell where the current element will be added.
	 * @param {number} nPos - The position in the cell where the current element will be added.
	 * @param {DocumentElement} oElement - The document element which will be added at the current position.
	 */
	ApiTable.prototype.AddElement = function(oCell, nPos, oElement)
	{
		if (!(oCell instanceof ApiTableCell) || this.Table !== oCell.Cell.Row.Table)
			return false;

		var apiCellContent = oCell.GetContent();

		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			apiCellContent.Document.Internal_Content_Add(nPos, oElement.private_GetImpl());

			return true;
		}

		return false;
	};
	/**
	 * Remove a table row with a specified cell.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} oCell - The cell which is present in the row that will be removed.
	 * @returns {boolean} Is the table empty after removing.
	 */
	ApiTable.prototype.RemoveRow = function(oCell)
	{
		if (!(oCell instanceof ApiTableCell) || this.Table !== oCell.Cell.Row.Table)
			return false;

		private_StartSilentMode();
		this.private_PrepareTableForActions();

		this.Table.RemoveSelection();
		this.Table.CurCell = oCell.Cell;
		var isEmpty = !(this.Table.RemoveTableRow());

		private_EndSilentMode();
		return isEmpty;
	};
	/**
	 * Remove a table column with a specified cell.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} oCell - The cell which is present in the column that will be removed.
	 * @returns {boolean} Is the table empty after removing.
	 */
	ApiTable.prototype.RemoveColumn = function(oCell)
	{
		if (!(oCell instanceof ApiTableCell) || this.Table !== oCell.Cell.Row.Table)
			return false;

		private_StartSilentMode();
		this.private_PrepareTableForActions();

		this.Table.RemoveSelection();
		this.Table.CurCell = oCell.Cell;
		var isEmpty = !(this.Table.RemoveTableColumn());

		private_EndSilentMode();
		return isEmpty;
	};
	/**
	 * Create a copy of a table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiTable}
	 */
	ApiTable.prototype.Copy = function()
	{
		var oTable = this.Table.Copy(private_GetLogicDocument(), private_GetDrawingDocument());
		return new ApiTable(oTable);
	};
	/**
	 * Select a table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE", "CPE"]
	 * @returns {bool}
	 */
	ApiTable.prototype.Select = function()
	{
		var Document = private_GetLogicDocument();
		
		var DocPos = this.Table.GetDocumentPositionFromObject();
		
		if (DocPos[0].Position === - 1)
			return false;

		var controllerType = null;

		if (DocPos[0].Class.IsHdrFtr())
		{
			controllerType = docpostype_HdrFtr;
		}
		else if (DocPos[0].Class.IsFootnote())
		{
			controllerType = docpostype_Footnotes;
		}
		else if (DocPos[0].Class.Is_DrawingShape())
		{
			controllerType = docpostype_DrawingObjects;
		}
		else 
		{
			controllerType = docpostype_Content;
		}
		DocPos[0].Class.CurPos.ContentPos = DocPos[0].Position;
		Document.SetDocPosType(controllerType);
		Document.SelectTable(3);

		return true;	
	};
	/**
	 * Return a Range object that represents the part of the document contained in the specified table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiTable.prototype.GetRange = function(Start, End)
	{
		var Range = new ApiRange(this.Table, Start, End)
		return Range;
	};
	/**
     * Set horizontal alignment for a table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @param {String} sType - Horizontal alignment type: may be "left" or "center" or "right".
     * @returns {bool} - returns false if param is invalid.
     * */
    ApiTable.prototype.SetHAlign = function(sType)
    {
		if (this.Table.IsInline())
		{
			if (sType == "left")
           		this.Table.Set_TableAlign(1);
        	else if (sType == "center")
            	this.Table.Set_TableAlign(2);
      			else if (sType == "right")
           	this.Table.Set_TableAlign(0);
      	  		else return false;
		}
		else if (!this.Table.IsInline())
		{
			if (sType == "left")
           		this.Table.Set_PositionH(0, true, 2);
        	else if (sType == "center")
            	this.Table.Set_PositionH(0, true, 0);
      			else if (sType == "right")
           	this.Table.Set_PositionH(0, true, 4);
      	  		else return false;
		}

        return true;
	};
	/**
     * Set vertical alignment for a table.
     * @typeofeditors ["CDE"]
     * @param {String} sType - Vertical alignment type: may be "top" or "center" or "bottom".
     * @returns {bool} - returns false if param is invalid.
     * */
    ApiTable.prototype.SetVAlign = function(sType)
    {
		if (this.Table.IsInline())
			return false;

        if (sType == "top")
            this.Table.Set_PositionV(0, true, 5);
        else if (sType == "center")
            this.Table.Set_PositionV(0, true, 1);
        else if (sType == "bottom")
            this.Table.Set_PositionV(0, true, 0);
        else return false;

        return true;
	};
	/**
     * Set table paddings.
	 * If table is inline -> only left padding applies.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @param {Number} nLeft - Left padding.
	 * @param {Number} nTop - Top padding.
	 * @param {Number} nRight - Right padding.
	 * @param {Number} nBottom - Bottom padding.
     * @returns {bool} - returns true.
     * */
    ApiTable.prototype.SetPaddings = function(nLeft, nTop, nRight, nBottom)
    {
		if (this.Table.IsInline())
			this.Table.Set_TableInd(nLeft);
		else if (!this.Table.IsInline())
    		this.Table.Set_Distance(nLeft, nTop, nRight, nBottom);

        return true;
	};
	/**
     * Set table wrapping style.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @param {bool} isFlow - Specifies if the table is inline or not.
	 * @returns {bool} - returns false if param is invalid. 
     * */
    ApiTable.prototype.SetWrappingStyle = function(isFlow)
    {
		if (isFlow === true)
		{
			this.Table.Set_Inline(isFlow);
			this.Table.Set_PositionH(0,false,0);
			this.Table.Set_PositionV(0,false,0);
		}
		else if (isFlow === false)
		{
			this.Table.Set_Inline(isFlow);
		}
		else 
			return false;

        return true;
	};
    /**
     * Get a content control that contains the table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @return {ApiBlockLvlSdt | null} - return null is parent content control doesn't exist.
     */
    ApiTable.prototype.GetParentContentControl = function()
    {
        var TablePosition = this.Table.GetDocumentPositionFromObject();

        for (var Index = TablePosition.length - 1; Index >= 1; Index--)
        {
            if (TablePosition[Index].Class)
                if (TablePosition[Index].Class instanceof CBlockLevelSdt)
                    return new ApiBlockLvlSdt(TablePosition[Index].Class);
        }

        return null;
	};
	/**
	 * Wrap the table object with a content control.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {number} nType - if nType === 1 -> returns ApiBlockLvlSdt, else return ApiTable.
	 * @return {ApiTable | ApiBlockLvlSdt}  
	 */
	ApiTable.prototype.InsertInContentControl = function(nType)
	{
		var Document = private_GetLogicDocument();

		var ContentControl = null;

		var tableIndex	= this.Table.Index;

		if (tableIndex >= 0)
		{
			this.Select();
			ContentControl = new ApiBlockLvlSdt(Document.AddContentControl(1));
			Document.RemoveSelection();
		}
		else 
		{
			ContentControl = new ApiBlockLvlSdt(new CBlockLevelSdt(Document, Document))
			ContentControl.Sdt.SetDefaultTextPr(Document.GetDirectTextPr());
			ContentControl.Sdt.Content.RemoveFromContent(0, ContentControl.Sdt.Content.GetElementsCount(), false);
			ContentControl.Sdt.Content.AddToContent(0, this.Table);
			ContentControl.Sdt.SetShowingPlcHdr(false);
		}

		if (nType === 1)
			return ContentControl;
		else 
			return this;
	};
    /**
     * Get the table that contains this table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @return {ApiTable | null} - returns null if parent table doesn't exist.  
     */
    ApiTable.prototype.GetParentTable = function()
    {
        var documentPos = this.Table.GetDocumentPositionFromObject();

        for (var Index = documentPos.length - 1; Index >= 1; Index--)
        {
            if (documentPos[Index].Class)
                if (documentPos[Index].Class instanceof CTable)
                    return new ApiTable(documentPos[Index].Class);
        }

        return null;
	};
	/**
     * Get the tables that contain this table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @return {ApiTable[]}  
     */
    ApiTable.prototype.GetTables = function()
    {
        var arrTables = [];

		var viewRow 	= undefined; // будем запоминать последнюю просмотренную строку, т.к. возможны случаи, когда она разбита на несколько Table Pages, такие просматривать повторно не нужно 
		var viewAbsPage = undefined; // будем запоминать последний абсолютный номер страницы, т.к. возможно случаи, когда строка разбита на несколько страниц, такие строки нужно просматривать повторно на каждой новой странице
		for (var nCurPage = 0, nPagesCount = this.Table.Pages.length; nCurPage < nPagesCount; ++nCurPage)
		{
			if (this.Table.Pages[nCurPage].FirtRow < 0 || this.Table.Pages[nCurPage].LastRow < 0)
				continue;

			var nTempPageAbs 	= this.Table.GetAbsolutePage(nCurPage);
			
			for (var nCurRow = this.Table.Pages[nCurPage].FirstRow; nCurRow <= this.Table.Pages[nCurPage].LastRow; ++nCurRow)
			{
				if (nCurRow === viewRow && viewAbsPage === nTempPageAbs)
					continue;

				viewRow = nCurRow;
				var oRow = this.Table.GetRow(nCurRow);

				if (oRow)
				{
					for (var nCurCell = 0, nCellsCount = oRow.GetCellsCount(); nCurCell < nCellsCount; ++nCurCell)
					{
						var oCell = oRow.GetCell(nCurCell);
						if (oCell.IsMergedCell())
							continue;

						oCell.GetContent().GetAllTablesOnPage(nTempPageAbs, arrTables);
					}
				}
			}

			viewAbsPage	= nTempPageAbs;
		}

		for (var Index = 0; Index < arrTables.length; Index++)
		{
			arrTables[Index] = new ApiTable(arrTables[Index].Table);
		}
		return arrTables;
	};
	/**
     * Get the next table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @return {ApiTable | null} - returns null if table is last.  
     */
    ApiTable.prototype.GetNext = function()
    {
		var oDocument = editor.GetDocument();

		var absEndPage = this.Table.GetAbsolutePage(this.Table.Pages.length - 1); // страница, на которой заканчивается таблица
        
		for (var curPage = absEndPage; curPage < oDocument.Document.Pages.length; curPage++)
		{
			var curPageTables = oDocument.Document.GetAllTablesOnPage(curPage); // все таблицы на странице 
			for (var Index = 0; Index < curPageTables.length; Index++)
			{
				if (curPageTables[Index].Table.Id === this.Table.Id)
				{
					if (curPageTables[Index + 1])
					{
						return new ApiTable(curPageTables[Index + 1].Table)
					}
					else 
						continue;
				}
				else 
					return new ApiTable(curPageTables[Index].Table);
			}
		}
		
		return null; 
	};
	/**
     * Get the previous table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @return {ApiTable | null} - returns null if table is first.  
     */
    ApiTable.prototype.GetPrevious = function()
    {
		var oDocument = editor.GetDocument();

		var absEndPage = this.Table.GetAbsolutePage(0); // страница, на которой заканчивается таблица
        
		for (var curPage = absEndPage; curPage >= 0; curPage--)
		{
			var curPageTables = oDocument.Document.GetAllTablesOnPage(curPage); // все таблицы на странице 
			for (var Index = curPageTables.length - 1; Index >= 0; Index--)
			{
				if (curPageTables[Index].Table.Id === this.Table.Id)
				{
					if (curPageTables[Index - 1])
					{
						return new ApiTable(curPageTables[Index - 1].Table)
					}
					else 
						continue;
				}
				else 
					return new ApiTable(curPageTables[Index].Table);
			}
		}
		
		return null; 
    };
    /**
     * Get the table cell that contains the table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @return {ApiTableCell | null} - returns null if parent cell doesn't exist.  
     */
    ApiTable.prototype.GetParentTableCell = function()
    {
        var documentPos = this.Table.GetDocumentPositionFromObject();

        for (var Index = documentPos.length - 1; Index >= 1; Index--)
        {
            if (documentPos[Index].Class.Parent)
                if (documentPos[Index].Class.Parent instanceof CTableCell)
                    return new ApiTableCell(documentPos[Index].Class.Parent);
        }

        return null;
	};
	/**
	 * Delete the table. 
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @return {bool} - returns false if parent of table doesn't exist.
	 */
	ApiTable.prototype.Delete = function()
	{
		var tableParent = this.Table.Parent;

		if (tableParent)
		{
			this.Table.PreDelete();
			tableParent.Remove_FromContent(this.Table.Index, 1, true);

			return true;
		}
		else 	 
			return false;
	};
	/**
	 * Clear the content from the table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @return {bool} - returns true.
	 */
	ApiTable.prototype.Clear = function()
	{
		for (var curRow = 0, rowsCount = this.Table.GetRowsCount(); curRow < rowsCount; curRow++)
		{
			var Row = this.Table.GetRow(curRow);
			for (var curCell = 0, cellsCount = Row.GetCellsCount(); curCell < cellsCount; curCell++)
			{
				Row.GetCell(curCell).GetContent().Clear_Content();
			}
		}

		return true;
	};
	/**
	 * Search for the scope of a table object. The search results are a collection of ApiRange objects.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {bool} isMatchCase - Case sensitive or not. 
	 * @return {ApiRange[]}  
	 */
	ApiTable.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase	= false;
		
		var arrApiRanges	= [];
		var allParagraphs	= [];
		this.Table.GetAllParagraphs({All : true}, allParagraphs);

		for (var para in allParagraphs)
		{
			var oParagraph			= new ApiParagraph(allParagraphs[para]);
			var arrOfParaApiRanges	= oParagraph.Search(sText, isMatchCase);

			for (var itemRange = 0; itemRange < arrOfParaApiRanges.length; itemRange++)	
				arrApiRanges.push(arrOfParaApiRanges[itemRange]);
		}

		return arrApiRanges;
	};
	/**
	 * Apply text settings to the entire contents of the table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The text properties that will be set to the table specified.
	 * @return {bool} - returns true. 
	 */
	ApiTable.prototype.SetTextPr = function(oTextPr)
	{
		var allParagraphs	= [];
		this.Table.GetAllParagraphs({All : true}, allParagraphs);

		for (var curPara = 0; curPara < allParagraphs.length; curPara++)
		{
			allParagraphs[curPara].SetApplyToAll(true);
			allParagraphs[curPara].Add(new AscCommonWord.ParaTextPr(oTextPr.TextPr));
			allParagraphs[curPara].SetApplyToAll(false);
		}
		
		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableRow
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {"tableRow"}
	 */
	ApiTableRow.prototype.GetClassType = function()
	{
		return "tableRow";
	};
	/**
	 * Get the number of cells in the current row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiTableRow.prototype.GetCellsCount = function()
	{
		return this.Row.Content.length;
	};
	/**
	 * Get the cell by its position.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The cell position in the current table.
	 * @returns {ApiTableCell}
	 */
	ApiTableRow.prototype.GetCell = function(nPos)
	{
		if (nPos < 0 || nPos >= this.Row.Content.length)
			return null;

		return new ApiTableCell(this.Row.Content[nPos]);
	};
	/**
	 * Get the row index.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {Number}
	 */
	ApiTableRow.prototype.GetIndex = function()
	{
		return this.Row.GetIndex();
	};
	/**
	 * Get parent table of the row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTable | null} - returns null if parent table doesn't exist.
	 */
	ApiTableRow.prototype.GetParentTable = function()
	{
		var Table = this.Row.GetTable();
		if (!Table)
			return null;

		return new ApiTable(Table);
	};
	/**
	 * Get the next row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableRow | null} - returns null if row is last.
	 */
	ApiTableRow.prototype.GetNext = function()
	{
		var Next = this.Row.Next;
		if (!Next)
			return null;

		return new ApiTableRow(Next);
	};
	/**
	 * Get the previous row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableRow | null} - returns null if row is first.
	 */
	ApiTableRow.prototype.GetPrevious = function()
	{
		var Prev = this.Row.Prev;
		if (!Prev)
			return null;

		return new ApiTableRow(Prev);
	};
	/**
	 * Add the new rows to the current table.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @param {Number} nCount - Count of rows to be added.
	 * @param {boolean} [isBefore=false] - Specifies if the rows will be added before or after the current row. 
	 * @returns {ApiTable | null} - returns null if parent table doesn't exist.
	 */
	ApiTableRow.prototype.AddRows = function(nCount, isBefore)
	{
		var oTable = this.GetParentTable();
		if(!oTable)
			return null;
		var oCell = this.GetCell(0);
		if (!oCell)
			return null;
			
		oTable.AddRows(oCell, nCount, isBefore);

		return oTable;
	};
	/**
	 * Merge cells in the row. 
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableCell | null} - return null if can't merge.
	 */
	ApiTableRow.prototype.MergeCells = function()
	{
		var oTable = this.GetParentTable();
		if(!oTable)
			return null;
		var cellsArr = [];
		var tempCell			= null;
		var tempGridSpan		= undefined;
		var tempStartGridCol	= undefined;
		var tempVMergeCount		= undefined;

		for (var curCell = 0, cellsCount = this.GetCellsCount(); curCell < cellsCount; curCell++)
		{
			tempCell 			= this.GetCell(curCell);
			tempStartGridCol	= this.Row.GetCellInfo(curCell).StartGridCol;
			tempGridSpan		= tempCell.Cell.GetGridSpan();
			tempVMergeCount		= oTable.Table.Internal_GetVertMergeCount2(this.GetIndex(), tempStartGridCol, tempGridSpan);

			if (tempVMergeCount > 1)
			{
				tempCell = new ApiTableCell(oTable.Table.GetCellByStartGridCol(this.GetIndex() - (tempVMergeCount - 1), tempStartGridCol));
			}

			cellsArr.push(tempCell);
		}
			
		return oTable.MergeCells(cellsArr);
	};
	/**
	 * Clear the content from row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {bool} - returns false if parent table doesn't exist.
	 */
	ApiTableRow.prototype.Clear = function()
	{
		var oTable = this.GetParentTable();
		if(!oTable)
			return false;

		var tempCell			= null;
		var tempGridSpan		= undefined;
		var tempStartGridCol	= undefined;
		var tempVMergeCount		= undefined;

		for (var curCell = 0, cellsCount = this.Row.GetCellsCount(); curCell < cellsCount; curCell++)
		{
			tempCell 			= this.Row.GetCell(curCell);
			tempStartGridCol	= this.Row.GetCellInfo(curCell).StartGridCol;
			tempGridSpan		= tempCell.GetGridSpan();
			tempVMergeCount		= oTable.Table.Internal_GetVertMergeCount2(this.GetIndex(), tempStartGridCol, tempGridSpan);

			if (tempVMergeCount > 1)
			{
				tempCell = oTable.Table.GetCellByStartGridCol(this.GetIndex() - (tempVMergeCount - 1), tempStartGridCol);
			}

			tempCell.GetContent().Clear_Content();
		}

		return true;
	};
	/**
	 * Remove a table row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {bool} - return false if parent table doesn't exist.
	 */
	ApiTableRow.prototype.Remove = function()
	{
		var oTable = this.GetParentTable();
		if (!oTable)
			return false;
		
		var oCell = this.GetCell(0);
		oTable.RemoveRow(oCell);

		return true;
	};
	/**
	 * Set text properties for the row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The text properties that will be set to the table specified.
	 * @returns {bool} - returns false if parent table doesn't exist or param is invalid.
	 */
	ApiTableRow.prototype.SetTextPr = function(oTextPr)
	{
		var oTable = this.GetParentTable();
		if(!oTable)
			return false;
		if (!oTextPr || !oTextPr.GetClassType || oTextPr.GetClassType() !== "textPr")
			return false;

		var tempCell			= null;
		var tempGridSpan		= undefined;
		var tempStartGridCol	= undefined;
		var tempVMergeCount		= undefined;

		for (var curCell = 0, cellsCount = this.Row.GetCellsCount(); curCell < cellsCount; curCell++)
		{
			tempCell 			= this.GetCell(curCell);
			tempStartGridCol	= this.Row.GetCellInfo(curCell).StartGridCol;
			tempGridSpan		= tempCell.Cell.GetGridSpan();
			tempVMergeCount		= oTable.Table.Internal_GetVertMergeCount2(this.GetIndex(), tempStartGridCol, tempGridSpan);

			if (tempVMergeCount > 1)
			{
				tempCell = new ApiTableCell(oTable.Table.GetCellByStartGridCol(this.GetIndex() - (tempVMergeCount - 1), tempStartGridCol));
			}

			tempCell.SetTextPr(oTextPr);
		}

		return true;
	};
	/**
	 * Search for the scope of a table row object. The search results are a collection of ApiRange objects.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {bool} isMatchCase - Case sensitive or not. 
	 * @return {ApiRange[]}  
	 */
	ApiTableRow.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase	= false;
		var oTable = this.GetParentTable();
		if (!oTable)
			return false;

		var arrApiRanges		= [];
		var tempResult			= [];
		var tempCell			= null;
		var tempGridSpan		= undefined;
		var tempStartGridCol	= undefined;
		var tempVMergeCount		= undefined;

		for (var curCell = 0, cellsCount = this.GetCellsCount(); curCell < cellsCount; curCell++)
		{
			tempCell 			= this.GetCell(curCell);
			tempStartGridCol	= this.Row.GetCellInfo(curCell).StartGridCol;
			tempGridSpan		= tempCell.Cell.GetGridSpan();
			tempVMergeCount		= oTable.Table.Internal_GetVertMergeCount2(this.GetIndex(), tempStartGridCol, tempGridSpan);

			if (tempVMergeCount > 1)
			{
				tempCell = new ApiTableCell(oTable.Table.GetCellByStartGridCol(this.GetIndex() - (tempVMergeCount - 1), tempStartGridCol));
			}

			tempResult = tempCell.Search(sText, isMatchCase);
			for (var nRange = 0; nRange < tempResult.length; nRange++)
			{
				arrApiRanges.push(tempResult[nRange]);
			}
		}

		return arrApiRanges;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableCell
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {"tableCell"}
	 */
	ApiTableCell.prototype.GetClassType = function()
	{
		return "tableCell";
	};
	/**
	 * Get the cell content.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {ApiDocumentContent}
	 */
	ApiTableCell.prototype.GetContent = function()
	{
		return new ApiDocumentContent(this.Cell.Content);
	};
	/**
	 * Get the cell index.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {Number}
	 */
	ApiTableCell.prototype.GetIndex = function()
	{
		return this.Cell.GetIndex();
	};
	/**
	 * Get the index of the parent row.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiTableCell.prototype.GetRowIndex = function()
	{
		var Row = this.Cell.GetRow();
		if(!Row)
			return null;

		return Row.GetIndex();
	};
	/**
	 * Get the parent row of the cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableRow | null} - returns null if parent row doesn't exist.
	 */
	ApiTableCell.prototype.GetParentRow = function()
	{
		var Row = this.Cell.GetRow();
		if(!Row)
			return null;

		return new ApiTableRow(Row);
	};
	/**
	 * Get the parent table of the cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTable | null} - returns null if parent table doesn't exist.
	 */
	ApiTableCell.prototype.GetParentTable = function()
	{
		var oTable = this.Cell.GetTable();
		if(!oTable)
			return null;

		return new ApiTable(oTable);
	};
	/**
	 * Add the new rows to the current table.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {Number} nCount - Count of rows to be added.
	 * @param {boolean} [isBefore=false] - Specifies if the new rows will be added before or after the cell. 
	 * @returns {ApiTable | null} - returns null if parent table doesn't exist.
	 */
	ApiTableCell.prototype.AddRows = function(nCount, isBefore)
	{
		var oTable = this.GetParentTable();
		if(!oTable)
			return null;

		oTable.AddRows(this, nCount, isBefore);

		return oTable;
	};
	/**
	 * Add the new columns to the current table.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {Number} nCount - Count of columns to be added.
	 * @param {boolean} [isBefore=false] - Specifies if the new columns will be added before or after the cell. 
	 * @returns {ApiTable | null} - returns null if parent table doesn't exist.
	 */
	ApiTableCell.prototype.AddColumns = function(nCount, isBefore)
	{
		var oTable = this.GetParentTable();
		if(!oTable)
			return null;
			
		oTable.AddColumns(this, nCount, isBefore);

		return oTable;
	};
	/**
	 * Remove a column containing a cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {bool | null} Is the table empty after removing. Returns null if parent table doesn't exist.
	 */
	ApiTableCell.prototype.RemoveColumn = function()
	{
		var oTable = this.GetParentTable();
		if (!oTable)
			return null;

		return oTable.RemoveColumn(this);
	};
	/**
	 * Remove a row containing a cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {bool | null} Is the table empty after removing. Returns null if parent table doesn't exist.
	 */
	ApiTableCell.prototype.RemoveRow = function()
	{
		var oTable = this.GetParentTable();
		if (!oTable)
			return false;

		return oTable.RemoveRow(this);
	};
	/**
	 * Search for the scope of a cell object. The search results are a collection of ApiRange objects.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {bool} isMatchCase - Case sensitive or not. 
	 * @return {ApiRange[]}  
	 */
	ApiTableCell.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase	= false;
		
		var arrApiRanges	= [];
		var allParagraphs	= [];
		var cellContent		= this.Cell.GetContent();
		cellContent.GetAllParagraphs({All : true}, allParagraphs);

		for (var para in allParagraphs)
		{
			var oParagraph			= new ApiParagraph(allParagraphs[para]);
			var arrOfParaApiRanges	= oParagraph.Search(sText, isMatchCase);

			for (var itemRange = 0; itemRange < arrOfParaApiRanges.length; itemRange++)	
				arrApiRanges.push(arrOfParaApiRanges[itemRange]);
		}

		return arrApiRanges;
	};
	/**
	 * Get the next cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableCell | null} - returns null if cell is last.
	 */
	ApiTableCell.prototype.GetNext = function()
	{
		var nextCell = this.Cell.Next;
		if(!nextCell)
			return null;
		
		return new ApiTableCell(nextCell);
	};
	/**
	 * Get the previous cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableCell | null} - returns null is cell is first. 
	 */
	ApiTableCell.prototype.GetPrevious = function()
	{
		var prevCell = this.Cell.Prev;
		if(!prevCell)
			return null;
		
		return new ApiTableCell(prevCell);
	};
	/**
	 * Split the cell into a given number of rows and columns.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {Number} [nRow=1] - Count of rows which the cell will be split.
	 * @param {Number} [nCol=1] - Count of columns which the cell will be split.
	 * @returns {ApiTable | null} - returns null if parent table doesn't exist.
	 */
	ApiTableCell.prototype.Split = function(nRow, nCol)
	{
		var oTable = this.GetParentTable();
		if (!oTable)
			return null;

		return oTable.Split(this, nRow, nCol);
	};
	/**
	 * Set the cell properties.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCellPr} oApiTableCellPr - The properties that will be set to the table cell.
	 * @returns {bool} - returns false if param is invalid.
	 */
	ApiTableCell.prototype.SetCellPr = function(oApiTableCellPr)
	{
		if (!oApiTableCellPr || !oApiTableCellPr.GetClassType || oApiTableCellPr.GetClassType() !== "tableCellPr")
			return false;

		this.CellPr.Merge(oApiTableCellPr.CellPr);
		this.private_OnChange();

		return true;
	};
	/**
	 * Apply text settings to the entire contents of the cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The properties that will be set to the table cell text.
	 * @return {bool} - returns false if param is invalid.
	 */
	ApiTableCell.prototype.SetTextPr = function(oTextPr)
	{
		if (!oTextPr || !oTextPr.GetClassType || oTextPr.GetClassType() !== "textPr")
			return false;

		var cellContent		= this.Cell.GetContent();
		var allParagraphs	= [];

		cellContent.GetAllParagraphs({All : true}, allParagraphs);
		for (var curPara = 0; curPara < allParagraphs.length; curPara++)
		{
			allParagraphs[curPara].SetApplyToAll(true);
			allParagraphs[curPara].Add(new AscCommonWord.ParaTextPr(oTextPr.TextPr));
			allParagraphs[curPara].SetApplyToAll(false);
		}
		
		return true;
	};
	/**
	 * Clear the content from the cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @return {bool} - returns false if parent row is invalid.
	 */
	ApiTableCell.prototype.Clear = function()
	{
		var oRow = this.GetParentRow();
		if (!oRow)
			return false;

		for (var curCell = 0, cellsCount = oRow.GetCellsCount(); curCell < cellsCount; curCell++)
		{
			oRow.Row.GetCell(curCell).GetContent().Clear_Content();
		}

		return true;
	};
	/**
	 * Add a paragraph or a table or a blockLvl content control using its position in the cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The position where the current element will be added.
	 * @param {DocumentElement} oElement - The document element which will be added at the current position.
	 * @returns {bool} - returns false if oElement is invalid.
	 */
	ApiTableCell.prototype.AddElement = function(nPos, oElement)
	{
		var apiCellContent = this.GetContent();

		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			apiCellContent.Document.Internal_Content_Add(nPos, oElement.private_GetImpl());

			return true;
		}

		return false;
	};
	
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiStyle
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {"style"}
	 */
	ApiStyle.prototype.GetClassType = function()
	{
		return "style";
	};
	/**
	 * Get the name of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiStyle.prototype.GetName = function()
	{
		return this.Style.Get_Name();
	};
	/**
	 * Set the name of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @param {string} sStyleName - The name which will be used for the current style.
	 */
	ApiStyle.prototype.SetName = function(sStyleName)
	{
		this.Style.Set_Name(sStyleName);
	};
	/**
	 * Get the type of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {StyleType}
	 */
	ApiStyle.prototype.GetType = function()
	{
		var nStyleType = this.Style.Get_Type();

		if (styletype_Paragraph === nStyleType)
			return "paragraph";
		else if (styletype_Table === nStyleType)
			return "table";
		else if (styletype_Character === nStyleType)
			return "run";
		else if (styletype_Numbering === nStyleType)
			return "numbering";

		return "paragraph";
	};
	/**
	 * Get the text properties of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTextPr}
	 */
	ApiStyle.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.Style.TextPr.Copy());
	};
	/**
	 * Get the paragraph properties of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParaPr}
	 */
	ApiStyle.prototype.GetParaPr = function()
	{
		return new ApiParaPr(this, this.Style.ParaPr.Copy());
	};
	/**
	 * Get the table properties of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {?ApiTablePr} If the type of this style is not a <code>"table"</code> then it will return
	 *     <code>null</code>.
	 */
	ApiStyle.prototype.GetTablePr = function()
	{
		if (styletype_Table !== this.Style.Get_Type())
			return null;

		return new ApiTablePr(this, this.Style.TablePr.Copy());
	};
	/**
	 * Get the table row properties of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {?ApiTableRowPr} If the type of this style is not a <code>"table"</code> then it will return
	 *     <code>null</code>.
	 */
	ApiStyle.prototype.GetTableRowPr = function()
	{
		if (styletype_Table !== this.Style.Get_Type())
			return null;

		return new ApiTableRowPr(this, this.Style.TableRowPr.Copy());
	};
	/**
	 * Get the table cell properties of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {?ApiTableCellPr}
	 */
	ApiStyle.prototype.GetTableCellPr = function()
	{
		if (styletype_Table !== this.Style.Get_Type())
			return null;

		return new ApiTableCellPr(this, this.Style.TableCellPr.Copy());
	};
	/**
	 * Specify the reference to the parent style which this style inherits from in the style hierarchy.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @param {ApiStyle} oStyle - The parent style which the style inherits properties from.
	 */
	ApiStyle.prototype.SetBasedOn = function(oStyle)
	{
		if (!(oStyle instanceof ApiStyle) || this.Style.Get_Type() !== oStyle.Style.Get_Type())
			return;

		this.Style.Set_BasedOn(oStyle.Style.Get_Id());
	};
	/**
	 * Get a set of formatting properties which will be conditionally applied to the parts of a table that match the 
	 * requirement specified in the sType parameter.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @param {TableStyleOverrideType} [sType="wholeTable"] - The table part which the formatting properties must be applied to.
	 * @returns {ApiTableStylePr}
	 */
	ApiStyle.prototype.GetConditionalTableStyle = function(sType)
	{
		if ("topLeftCell" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableTLCell.Copy());
		else if ("topRightCell" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableTRCell.Copy());
		else if ("bottomLeftCell" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableBLCell.Copy());
		else if ("bottomRightCell" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableBRCell.Copy());
		else if ("firstRow" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableFirstRow.Copy());
		else if ("lastRow" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableLastRow.Copy());
		else if ("firstColumn" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableFirstCol.Copy());
		else if ("lastColumn" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableLastCol.Copy());
		else if ("bandedColumn" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableBand1Vert.Copy());
		else if("bandedColumnEven" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableBand2Vert.Copy());
		else if ("bandedRow" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableBand1Horz.Copy());
		else if ("bandedRowEven" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableBand2Horz.Copy());
		else if ("wholeTable" === sType)
			return new ApiTableStylePr(sType, this, this.Style.TableWholeTable.Copy());

		return new ApiTableStylePr(sType, this, this.Style.TableWholeTable.Copy());
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTextPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"textPr"}
	 */
	ApiTextPr.prototype.GetClassType = function()
	{
		return "textPr";
	};
	/**
	 * The text style base method.
	 * <note>This method is not used by itself, as it only forms the basis for the {@link ApiRun#SetStyle} method which sets
	 * the selected or created style for the text.</note>
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {ApiStyle} oStyle - The style which must be applied to the text character.
	 */
	ApiTextPr.prototype.SetStyle = function(oStyle)
	{
		if (!(oStyle instanceof ApiStyle))
			return;

		this.TextPr.RStyle = oStyle.Style.Get_Id();
		this.private_OnChange();
	};
	/**
	 * Set the bold property to the text character.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isBold - Specifies that the contents of the run are displayed bold.
	 */
	ApiTextPr.prototype.SetBold = function(isBold)
	{
		this.TextPr.Bold = isBold;
		this.private_OnChange();
	};
	/**
	 * Set the italic property to the text character.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isItalic - Specifies that the contents of the current run are displayed italicized.
	 */
	ApiTextPr.prototype.SetItalic = function(isItalic)
	{
		this.TextPr.Italic = isItalic;
		this.private_OnChange();
	};
	/**
	 * Specify that the contents of the run are displayed with a single horizontal line through the center of the line.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isStrikeout - Specifies that the contents of the current run are displayed struck through.
	 */
	ApiTextPr.prototype.SetStrikeout = function(isStrikeout)
	{
		this.TextPr.Strikeout = isStrikeout;
		this.private_OnChange();
	};
	/**
	 * Specify that the contents of the run are displayed along with a line appearing directly below the character
	 * (less than all the spacing above and below the characters on the line).
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isUnderline - Specifies that the contents of the current run are displayed underlined.
	 */
	ApiTextPr.prototype.SetUnderline = function(isUnderline)
	{
		this.TextPr.Underline = isUnderline;
		this.private_OnChange();
	};
	/**
	 * Set all 4 font slots with the specified font family.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} sFontFamily - The font family or families used for the current text run.
	 */
	ApiTextPr.prototype.SetFontFamily = function(sFontFamily)
	{
		this.TextPr.RFonts.SetAll(sFontFamily, -1);
		this.private_OnChange();
	};
	/**
	 * Set the font size for the characters of the current text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {hps} nSize - The text size value measured in half-points (1/144 of an inch).
	 */
	ApiTextPr.prototype.SetFontSize = function(nSize)
	{
		this.TextPr.FontSize = private_GetHps(nSize);
		this.private_OnChange();
	};
	/**
	 * Set the text color for the current text run in the RGB format.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - If this parameter is set to "true", then r,g,b parameters will be ignored.
	 */
	ApiTextPr.prototype.SetColor = function(r, g, b, isAuto)
	{
		this.TextPr.Color = private_GetColor(r, g, b, isAuto);
		this.private_OnChange();
	};
	/**
	 * Specify the alignment which will be applied to the contents of this run in relation to the default appearance of the run text:
	 * * <b>"baseline"</b> - the characters in the current text run will be aligned by the default text baseline.
	 * * <b>"subscript"</b> - the characters in the current text run will be aligned below the default text baseline.
	 * * <b>"superscript"</b> - the characters in the current text run will be aligned above the default text baseline.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {("baseline" | "subscript" | "superscript")} sType - The vertical alignment type applied to the text contents.
	 */
	ApiTextPr.prototype.SetVertAlign = function(sType)
	{
		if ("baseline" === sType)
			this.TextPr.VertAlign = AscCommon.vertalign_Baseline;
		else if ("subscript" === sType)
			this.TextPr.VertAlign = AscCommon.vertalign_SubScript;
		else if ("superscript" === sType)
			this.TextPr.VertAlign = AscCommon.vertalign_SuperScript;

		this.private_OnChange();
	};
	/**
	 * Specify a highlighting color in the RGB format which is applied as a background for the contents of the current run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isNone=false] If this parameter is set to "true", then r,g,b parameters will be ignored.
	 */
	ApiTextPr.prototype.SetHighlight = function(r, g, b, isNone)
	{
		if (undefined === isNone)
			isNone = false;

		if (true === isNone)
			this.TextPr.HighLight = AscCommonWord.highlight_None;
		else
			this.TextPr.HighLight = new AscCommonWord.CDocumentColor(r, g, b, false);

		this.private_OnChange();
	};
	/**
	 * Set text spacing measured in twentieths of a point.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nSpacing - The value of the text spacing measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiTextPr.prototype.SetSpacing = function(nSpacing)
	{
		this.TextPr.Spacing = private_Twips2MM(nSpacing);
		this.private_OnChange();
	};
	/**
	 * Specify that the contents of this run are displayed with two horizontal lines through each character displayed on the line.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isDoubleStrikeout - Specifies that the contents of the current run are displayed double struck through.
	 */
	ApiTextPr.prototype.SetDoubleStrikeout = function(isDoubleStrikeout)
	{
		this.TextPr.DStrikeout = isDoubleStrikeout;
		this.private_OnChange();
	};
	/**
	 * Specify that any lowercase characters in this text run are formatted for display only as their capital letter character equivalents.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isCaps - Specifies that the contents of the current run are displayed capitalized.
	 */
	ApiTextPr.prototype.SetCaps = function(isCaps)
	{
		this.TextPr.Caps = isCaps;
		this.private_OnChange();
	};
	/**
	 * Specify that all the small letter characters in the text run are formatted for display only as their capital
	 * letter character equivalents which are two points smaller than the actual font size specified for this text.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isSmallCaps - Specifies if the contents of the current run are displayed capitalized two points smaller or not.
	 */
	ApiTextPr.prototype.SetSmallCaps = function(isSmallCaps)
	{
		this.TextPr.SmallCaps = isSmallCaps;
		this.private_OnChange();
	};
	/**
	 * Specify the amount by which text is raised or lowered for this run in relation to the default
	 * baseline of the surrounding non-positioned text.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {hps} nPosition - Specifies a positive (raised text) or negative (lowered text)
	 * measurement in half-points (1/144 of an inch).
	 */
	ApiTextPr.prototype.SetPosition = function(nPosition)
	{
		this.TextPr.Position = private_PtToMM(private_GetHps(nPosition));
		this.private_OnChange();
	};
	/**
	 * Specify the languages which will be used to check spelling and grammar (if requested) when processing
	 * the contents of this text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {string} sLangId - The possible value for this parameter is a language identifier as defined by
	 * RFC 4646/BCP 47. Example: "en-CA".
	 */
	ApiTextPr.prototype.SetLanguage = function(sLangId)
	{
		var nLcid = Asc.g_oLcidNameToIdMap[sLangId];
		if (undefined !== nLcid)
		{
			this.TextPr.Lang.Val = nLcid;
			this.private_OnChange();
		}
	};
	/**
	 * Specify the shading applied to the contents of the current text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type applied to the contents of the current text run.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTextPr.prototype.SetShd = function(sType, r, g, b)
	{
		this.TextPr.Shd = private_GetShd(sType, r, g, b, false);
		this.private_OnChange();
	};


	/**
	 * Set the text color for the current text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CSE", "CPE"]
	 * @param {ApiFill} oApiFill - The color or pattern used to fill the text color.
	 */
	ApiTextPr.prototype.SetFill = function(oApiFill)
	{
		this.TextPr.Unifill = oApiFill.UniFill;
		this.private_OnChange();
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiParaPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"paraPr"}
	 */
	ApiParaPr.prototype.GetClassType = function()
	{
		return "paraPr";
	};
	/**
	 * The paragraph style base method.
	 * <note>This method is not used by itself, as it only forms the basis for the {@link ApiParagraph#SetStyle} method which sets the selected or created style for the paragraph.</note>
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {ApiStyle} oStyle - The style of the paragraph to be set.
	 */
	ApiParaPr.prototype.SetStyle = function(oStyle)
	{
		if (!oStyle || !(oStyle instanceof ApiStyle))
			return;

		this.ParaPr.PStyle = oStyle.Style.Get_Id();
		this.private_OnChange();
	};
	/**
	 * Get the paragraph style method.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @return {?ApiStyle} - The style of the paragraph.
	 */
	ApiParaPr.prototype.GetStyle = function()
	{
		var oDocument = private_GetLogicDocument();
		var oStyles   = oDocument.GetStyles();

		var styleId;
		if (!this.Parent)
		{
			styleId = this.ParaPr.PStyle;
			if (styleId)
				return new ApiStyle(oStyles.Get(styleId));

			return null;
		}

		styleId = this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.PStyle;
		if (styleId)
			return new ApiStyle(oStyles.Get(styleId));

		return null;
	};
	/**
	 * Specify that any space before or after this paragraph set using the 
	 * {@link ApiParaPr#SetSpacingBefore} or {@link ApiParaPr#SetSpacingAfter} spacing element, should not be applied when the preceding and 
	 * following paragraphs are of the same paragraph style, affecting the top and bottom spacing respectively.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isContextualSpacing - The true value will enable the paragraph contextual spacing.
	 */
	ApiParaPr.prototype.SetContextualSpacing = function(isContextualSpacing)
	{
		this.ParaPr.ContextualSpacing = private_GetBoolean(isContextualSpacing);
		this.private_OnChange();
	};
	/**
	 * Set the paragraph left side indentation.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nValue - The paragraph left side indentation value measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiParaPr.prototype.SetIndLeft = function(nValue)
	{
		this.ParaPr.Ind.Left = private_Twips2MM(nValue);
		this.private_OnChange();
	};
	/**
	 * Get the paragraph left side indentation.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {twips | undefined} - The paragraph left side indentation value measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiParaPr.prototype.GetIndLeft = function()
	{
		if (!this.Parent)
		{
			if (this.ParaPr.Ind.Left !== undefined)
				return AscCommon.MMToTwips(this.ParaPr.Ind.Left);
			return undefined;
		}
		return AscCommon.MMToTwips(this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Ind.Left);
	};
	/**
	 * Set the paragraph right side indentation.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nValue - The paragraph right side indentation value measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiParaPr.prototype.SetIndRight = function(nValue)
	{
		this.ParaPr.Ind.Right = private_Twips2MM(nValue);
		this.private_OnChange();
	};
	/**
	 * Get the paragraph right side indentation.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {twips | undefined} - The paragraph right side indentation value measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiParaPr.prototype.GetIndRight = function()
	{
		if (!this.Parent)
		{
			if (this.ParaPr.Ind.Right !== undefined)
				return AscCommon.MMToTwips(this.ParaPr.Ind.Right);

			return undefined;
		}
		return AscCommon.MMToTwips(this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Ind.Right);
	};
	/**
	 * Set the paragraph first line indentation.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nValue - The paragraph first line indentation value measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiParaPr.prototype.SetIndFirstLine = function(nValue)
	{
		this.ParaPr.Ind.FirstLine = private_Twips2MM(nValue);
		this.private_OnChange();
	};
	/**
	 * Get the paragraph first line indentation.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {twips | undefined} - The paragraph first line indentation value measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiParaPr.prototype.GetIndFirstLine = function()
	{
		if (!this.Parent)
		{
			if (this.ParaPr.Ind.FirstLine !== undefined)
				return AscCommon.MMToTwips(this.ParaPr.Ind.FirstLine);

			return undefined;
		}

		return AscCommon.MMToTwips(this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Ind.FirstLine);
	};
	/**
	 * Set paragraph contents justification.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {("left" | "right" | "both" | "center")} sJc - The justification type that
	 * will be applied to the paragraph contents.
	 */
	ApiParaPr.prototype.SetJc = function(sJc)
	{
		this.ParaPr.Jc = private_GetParaAlign(sJc);
		this.private_OnChange();
	};
	/**
	 * Get paragraph contents justification.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {("left" | "right" | "both" | "center" | undefined)} 
	 */
	ApiParaPr.prototype.GetJc = function()
	{
		function GetJC(nType) 
		{
			switch (nType)
			{
				case align_Right :
					return "right";
				case align_Left :
					return "left";
				case align_Center :
					return "center";
				case align_Justify : 
					return "both";
			}

			return "left";
		}

		if (!this.Parent)
		{
			if (this.ParaPr.Jc !== undefined)
				return GetJC(this.ParaPr.Jc);

			return undefined;
		}

		return GetJC(this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Jc);
	};
	/**
	 * Specify that when rendering this document using a page view, all lines of this paragraph are maintained on a single page whenever possible.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isKeepLines - The true value will enable the option to keep lines of the paragraph on a single page.
	 */
	ApiParaPr.prototype.SetKeepLines = function(isKeepLines)
	{
		this.ParaPr.KeepLines = isKeepLines;
		this.private_OnChange();
	};
	/**
	 * Specify that when rendering this document using a paginated view, the contents of this paragraph are at least
	 * partly rendered on the same page as the following paragraph whenever possible.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isKeepNext - The true value will enable the option to keep lines of the paragraph on the same
	 * page as the following paragraph.
	 */
	ApiParaPr.prototype.SetKeepNext = function(isKeepNext)
	{
		this.ParaPr.KeepNext = isKeepNext;
		this.private_OnChange();
	};
	/**
	 * Specify that when rendering this document using a paginated view, the contents of this paragraph are rendered at
	 * the beginning of a new page in the document.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isPageBreakBefore - The true value will enable the option to render the contents of the paragraph
	 * at the beginning of a new page in the document.
	 */
	ApiParaPr.prototype.SetPageBreakBefore = function(isPageBreakBefore)
	{
		this.ParaPr.PageBreakBefore = isPageBreakBefore;
		this.private_OnChange();
	};
	/**
	 * Set the paragraph line spacing. If the value of the sLineRule parameter is either 
	 * "atLeast" or "exact", then the value of nLine will be interpreted as twentieths of a point. If 
	 * the value of the sLineRule parameter is "auto", then the value of the 
	 * nLine parameter will be interpreted as 240ths of a line.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {(twips | line240)} nLine - The line spacing value measured either in twentieths of a point (1/1440 of an inch) or in 240ths of a line.
	 * @param {("auto" | "atLeast" | "exact")} sLineRule - The rule that determines the measuring units of the nLine parameter.
	 */
	ApiParaPr.prototype.SetSpacingLine = function(nLine, sLineRule)
	{
		if (undefined !== nLine && undefined !== sLineRule)
		{
			if ("auto" === sLineRule)
			{
				this.ParaPr.Spacing.LineRule = Asc.linerule_Auto;
				this.ParaPr.Spacing.Line     = nLine / 240.0;
			}
			else if ("atLeast" === sLineRule)
			{
				this.ParaPr.Spacing.LineRule = Asc.linerule_AtLeast;
				this.ParaPr.Spacing.Line     = private_Twips2MM(nLine);

			}
			else if ("exact" === sLineRule)
			{
				this.ParaPr.Spacing.LineRule = Asc.linerule_Exact;
				this.ParaPr.Spacing.Line     = private_Twips2MM(nLine);
			}
		}

		this.private_OnChange();
	};
	/**
	 * Get the paragraph line spacing value.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {twips | line240 | undefined} - to know is twips or line240 use ApiParaPr.prototype.GetSpacingLineRule().
	 */
	ApiParaPr.prototype.GetSpacingLineValue = function()
	{
		function GetValue(oSpacing)
		{
			switch (oSpacing.LineRule)
			{
				case Asc.linerule_Auto:
					return oSpacing.Line * 240.0;
				case Asc.linerule_AtLeast:
				case Asc.linerule_Exact:
					return AscCommon.MMToTwips(oSpacing.Line);
			}

			return undefined;
		}

		if (!this.Parent)
		{
			if (this.ParaPr.Spacing)
				return GetValue(this.ParaPr.Spacing);

			return undefined;
		}

		return GetValue(this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Spacing);
	};
	/**
	 * Get the paragraph line spacing rule.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"auto" | "atLeast" | "exact" | undefined} 
	 */
	ApiParaPr.prototype.GetSpacingLineRule = function()
	{
		function GetRule(nLineRule)
		{
			switch (nLineRule)
			{
				case Asc.linerule_Auto:
					return "auto";
				case Asc.linerule_AtLeast:
					return "atLeast";
				case Asc.linerule_Exact:
					return "exact";
			}

			return "atLeast";
		}

		if (!this.Parent)
		{
			if (this.ParaPr.Spacing)
				return GetRule(this.ParaPr.Spacing.LineRule);

			return undefined;
		}

		return GetRule(this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Spacing.LineRule);

	};
	/**
	 * Set the spacing before the current paragraph. If the value of the isBeforeAuto parameter is true, then 
	 * any value of the nBefore is ignored. If isBeforeAuto parameter is not specified, then 
	 * it will be interpreted as false.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nBefore - The value of the spacing before the current paragraph measured in twentieths of a point (1/1440 of an inch).
	 * @param {boolean} [isBeforeAuto=false] - The true value will disable the nBefore parameter.
	 */
	ApiParaPr.prototype.SetSpacingBefore = function(nBefore, isBeforeAuto)
	{
		if (undefined !== nBefore)
			this.ParaPr.Spacing.Before = private_Twips2MM(nBefore);

		if (undefined !== isBeforeAuto)
			this.ParaPr.Spacing.BeforeAutoSpacing = isBeforeAuto;

		this.private_OnChange();
	};
	/**
	 * Get the spacing before value of the current paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {twips} - The value of the spacing before the current paragraph measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiParaPr.prototype.GetSpacingBefore = function()
	{
		if (!this.Parent)
		{
			if (this.ParaPr.Spacing.Before !== undefined)
				return AscCommon.MMToTwips(this.ParaPr.Spacing.Before);

			return undefined;
		}

		return AscCommon.MMToTwips(this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Spacing.Before);
	};
	/**
	 * Set the spacing after the current paragraph. If the value of the isAfterAuto parameter is true, then 
	 * any value of the nAfter is ignored. If isAfterAuto parameter is not specified, then it 
	 * will be interpreted as false.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nAfter - The value of the spacing after the current paragraph measured in twentieths of a point (1/1440 of an inch).
	 * @param {boolean} [isAfterAuto=false] - The true value will disable the nAfter parameter.
	 */
	ApiParaPr.prototype.SetSpacingAfter = function(nAfter, isAfterAuto)
	{
		if (undefined !== nAfter)
			this.ParaPr.Spacing.After = private_Twips2MM(nAfter);

		if (undefined !== isAfterAuto)
			this.ParaPr.Spacing.AfterAutoSpacing = isAfterAuto;

		this.private_OnChange();
	};
	/**
	 * Get the spacing after value of the current paragraph. 
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {twips} - The value of the spacing after the current paragraph measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiParaPr.prototype.GetSpacingAfter = function()
	{
		if (!this.Parent)
		{
			if (this.ParaPr.Spacing.After !== undefined)
				return AscCommon.MMToTwips(this.ParaPr.Spacing.After);

			return undefined;
		}

		return AscCommon.MMToTwips(this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Spacing.After);
	};
	/**
	 * Specify the shading applied to the contents of the paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type which will be applied to the contents of the current paragraph.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - The true value will disable paragraph contents shading.
	 */
	ApiParaPr.prototype.SetShd = function(sType, r, g, b, isAuto)
	{
		this.ParaPr.Shd = private_GetShd(sType, r, g, b, isAuto);
		this.private_OnChange();
	};
	/**
	 * Get the shading applied to the contents of the paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @returns {?ApiRGBColor}
	 */
	ApiParaPr.prototype.GetShd = function()
	{
		var oColor = null;
		var oShd   = null;
		if (!this.Parent)
		{
			oShd = this.ParaPr.Shd;
			if (!oShd)
				return null;

			oColor = this.ParaPr.Shd.Color;
			if (oColor)
				return new ApiRGBColor(oColor.r, oColor.g, oColor.b);
			
			return null;
		}

		oShd = this.ParaPr.Shd;
		if (!oShd)
			return null;

		oColor = this.Parent.private_GetImpl().Get_CompiledPr2().ParaPr.Shd.Color;
		if (oColor)
			return new ApiRGBColor(oColor.r, oColor.g, oColor.b);

		return null;
	};
	/**
	 * Specify the border which will be displayed below a set of paragraphs which have the same paragraph border settings.
	 * <note>The paragraphs of the same style going one by one are considered as a single block, so the border is added
	 * to the whole block rather than to every paragraph in this block.</note>
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The border style.
	 * @param {pt_8} nSize - The width of the current bottom border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset below the paragraph measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiParaPr.prototype.SetBottomBorder = function(sType, nSize, nSpace, r, g, b)
	{
		this.ParaPr.Brd.Bottom = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Specify the border which will be displayed at the left side of the page around the specified paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The border style.
	 * @param {pt_8} nSize - The width of the current left border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset to the left of the paragraph measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiParaPr.prototype.SetLeftBorder = function(sType, nSize, nSpace, r, g, b)
	{
		this.ParaPr.Brd.Left = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Specify the border which will be displayed at the right side of the page around the specified paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The border style.
	 * @param {pt_8} nSize - The width of the current right border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset to the right of the paragraph measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiParaPr.prototype.SetRightBorder = function(sType, nSize, nSpace, r, g, b)
	{
		this.ParaPr.Brd.Right = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Specify the border which will be displayed above a set of paragraphs which have the same set of paragraph border settings.
	 * <note>The paragraphs of the same style going one by one are considered as a single block, so the border is added to the whole block rather than to every paragraph in this block.</note>
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The border style.
	 * @param {pt_8} nSize - The width of the current top border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset above the paragraph measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiParaPr.prototype.SetTopBorder = function(sType, nSize, nSpace, r, g, b)
	{
		this.ParaPr.Brd.Top = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Specify the border which will be displayed between each paragraph in a set of paragraphs which have the same set of paragraph border settings.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The border style.
	 * @param {pt_8} nSize - The width of the current border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset between the paragraphs measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiParaPr.prototype.SetBetweenBorder = function(sType, nSize, nSpace, r, g, b)
	{
		this.ParaPr.Brd.Between = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Specify whether a single line of this paragraph will be prevented from being displayed on a separate page from the remaining content at display time by moving the line onto the following page.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isWidowControl - The true value will enable the SetWidowControl method use.
	 */
	ApiParaPr.prototype.SetWidowControl = function(isWidowControl)
	{
		this.ParaPr.WidowControl = isWidowControl;
		this.private_OnChange();
	};
	/**
	 * Specify a sequence of custom tab stops which will be used for any tab characters in the current paragraph.
	 * <b>Warning</b>: The lengths of aPos array and aVal array <b>MUST BE</b> equal to each other.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips[]} aPos - An array of the positions of custom tab stops with respect to the current page margins
	 * measured in twentieths of a point (1/1440 of an inch).
	 * @param {TabJc[]} aVal - An array of the styles of custom tab stops, which determines the behavior of the tab
	 * stop and the alignment which will be applied to text entered at the current custom tab stop.
	 */
	ApiParaPr.prototype.SetTabs = function(aPos, aVal)
	{
		if (!(aPos instanceof Array) || !(aVal instanceof Array) || aPos.length !== aVal.length)
			return;

		var oTabs = new CParaTabs();
		for (var nIndex = 0, nCount = aPos.length; nIndex < nCount; ++nIndex)
		{
			oTabs.Add(private_GetTabStop(aPos[nIndex], aVal[nIndex]));
		}
		this.ParaPr.Tabs = oTabs;
		this.private_OnChange();
	};
	/**
	 * Specify that the current paragraph references a numbering definition instance in the current document.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {ApiNumbering} oNumPr - Specifies a numbering definition.
	 * @param {number} [nLvl=0] - Specifies a numbering level reference. If the current instance of the ApiParaPr class is direct
	 * formatting of a paragraph, then this parameter MUST BE specified. Otherwise if the current instance of the ApiParaPr class
	 * is the part of ApiStyle properties, this parameter will be ignored.
	 */
	ApiParaPr.prototype.SetNumPr = function(oNumPr, nLvl)
	{
		if (!(oNumPr instanceof ApiNumbering))
			return;

		this.ParaPr.NumPr       = new CNumPr();
		this.ParaPr.NumPr.NumId = oNumPr.Num.GetId();
		this.ParaPr.NumPr.Lvl   = undefined;

		if (this.Parent instanceof ApiParagraph)
		{
			this.ParaPr.NumPr.Lvl = Math.min(8, Math.max(0, (nLvl ? nLvl : 0)));
		}
		this.private_OnChange();
	};
	/**
	 * Set the bullet or numbering to the current paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CSE", "CPE"]
	 * @param {?ApiBullet} oBullet - The bullet object created using either the {@link Api#CreateBullet} or {@link Api#CreateNumbering} method.
	 */
	ApiParaPr.prototype.SetBullet = function(oBullet){
		if(oBullet){
			this.ParaPr.Bullet = oBullet.Bullet;
		}
		else{
			this.ParaPr.Bullet = null;
		}
		this.private_OnChange();
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiNumbering
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiNumbering
	 * @typeofeditors ["CDE"]
	 * @returns {"numbering"}
	 */
	ApiNumbering.prototype.GetClassType = function()
	{
		return "numbering";
	};
	/**
	 * Get the specified level of the current numbering.
	 * @memberof ApiNumbering
	 * @typeofeditors ["CDE"]
	 * @param {number} nLevel - The numbering level index. This value MUST BE from 0 to 8.
	 * @returns {ApiNumberingLevel}
	 */
	ApiNumbering.prototype.GetLevel = function(nLevel)
	{
		return new ApiNumberingLevel(this.Num, nLevel);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiNumberingLevel
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {"numberingLevel"}
	 */
	ApiNumberingLevel.prototype.GetClassType = function()
	{
		return "numberingLevel";
	};
	/**
	 * Get the numbering definition.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {ApiNumbering}
	 */
	ApiNumberingLevel.prototype.GetNumbering = function()
	{
		return new ApiNumbering(this.Num);
	};
	/**
	 * Get the level index.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiNumberingLevel.prototype.GetLevelIndex = function()
	{
		return this.Lvl;
	};
	/**
	 * Specify the text properties which will be applied to the text in the current numbering level itself, not to the text in the subsequent paragraph.
	 * <note>To change the text style for the paragraph, a style must be applied to it using the {@link ApiRun#SetStyle} method.</note>
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTextPr}
	 */
	ApiNumberingLevel.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.Num.GetLvl(this.Lvl).TextPr.Copy());
	};
	/**
	 * The paragraph properties which are applied to any numbered paragraph that references the given numbering definition and numbering level.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParaPr}
	 */
	ApiNumberingLevel.prototype.GetParaPr = function()
	{
		return new ApiParaPr(this, this.Num.GetLvl(this.Lvl).ParaPr.Copy());
	};
	/**
	 * Set one of the existing predefined numbering templates.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @param {("none" | "bullet" | "1)" | "1." | "I." | "A." | "a)" | "a." | "i." )} sType - The predefined numbering template.
	 * @param {string} [sSymbol=""] - The symbol used for the list numbering. This parameter have a meaning only if the sType="bullet" property is selected.
	 */
	ApiNumberingLevel.prototype.SetTemplateType = function(sType, sSymbol)
	{
		switch (sType)
		{
			case "none"  :
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.None);
				break;
			case "bullet":
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.Bullet, sSymbol, new CTextPr());
				break;
			case "1)"    :
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.DecimalBracket_Right);
				break;
			case "1."    :
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.DecimalDot_Right);
				break;
			case "I."    :
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.UpperRomanDot_Right);
				break;
			case "A."    :
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.UpperLetterDot_Left);
				break;
			case "a)"    :
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.LowerLetterBracket_Left);
				break;
			case "a."    :
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.LowerLetterDot_Left);
				break;
			case "i."    :
				this.Num.SetLvlByType(this.Lvl, c_oAscNumberingLevel.LowerRomanDot_Right);
				break;
		}
	};
	/**
	 * Set your own customized numbering type.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @param {("none" | "bullet" | "decimal" | "lowerRoman" | "upperRoman" | "lowerLetter" | "upperLetter" |
	 *     "decimalZero")} sType - The custom numbering type used for the current numbering definition.
	 * @param {string} sTextFormatString - Any text in this parameter will be taken as literal text to be repeated in each instance of this numbering level, except for any use of the percent symbol (%) followed by a number, which will be used to indicate the one-based index of the number to be used at this level. Any number of a level higher than this level will be ignored.
	 * @param {("left" | "right" | "center")} sAlign - Type of justification applied to the text run in the current numbering level.
	 */
	ApiNumberingLevel.prototype.SetCustomType = function(sType, sTextFormatString, sAlign)
	{
		var nType = Asc.c_oAscNumberingFormat.None;
		if ("none" === sType)
			nType = Asc.c_oAscNumberingFormat.None;
		else if ("bullet" === sType)
			nType = Asc.c_oAscNumberingFormat.Bullet;
		else if ("decimal" === sType)
			nType = Asc.c_oAscNumberingFormat.Decimal;
		else if ("lowerRoman" === sType)
			nType = Asc.c_oAscNumberingFormat.LowerRoman;
		else if ("upperRoman" === sType)
			nType = Asc.c_oAscNumberingFormat.UpperRoman;
		else if ("lowerLetter" === sType)
			nType = Asc.c_oAscNumberingFormat.LowerLetter;
		else if ("upperLetter" === sType)
			nType = Asc.c_oAscNumberingFormat.UpperLetter;
		else if ("decimalZero" === sType)
			nType = Asc.c_oAscNumberingFormat.DecimalZero;

		var nAlign = align_Left;
		if ("left" === sAlign)
			nAlign = align_Left;
		else if ("right" === sAlign)
			nAlign = align_Right;
		else if ("center" === sAlign)
			nAlign = align_Center;

		this.Num.SetLvlByFormat(this.Lvl, nType, sTextFormatString, nAlign);
	};
	/**
	 * Specify a one-based index which determines when a numbering level should restart to its starting value. A numbering level restarts when an instance of the specified numbering level, which will be higher (earlier than the this level) is used in the given document contents. By default this value is true.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isRestart - The true value will enable the SetRestart method use.
	 */
	ApiNumberingLevel.prototype.SetRestart = function(isRestart)
	{
		this.Num.SetLvlRestart(this.Lvl, private_GetBoolean(isRestart, true));
	};
	/**
	 * Specify the starting value for the numbering used by the parent numbering level within a given numbering level definition. By default this value is 1.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @param {number} nStart - The starting value for the numbering used by the parent numbering level.
	 */
	ApiNumberingLevel.prototype.SetStart = function(nStart)
	{
		this.Num.SetLvlStart(this.Lvl, private_GetInt(nStart));
	};
	/**
	 * Specify the content which will be added between a given numbering level text and the text of every numbered paragraph which references that numbering level. By default this value is "tab".
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @param {("space" | "tab" | "none")} sType - The content added between the numbering level text and the text in the numbered paragraph.
	 */
	ApiNumberingLevel.prototype.SetSuff = function(sType)
	{
		if ("space" === sType)
			this.Num.SetLvlSuff(this.Lvl, Asc.c_oAscNumberingSuff.Space);
		else if ("tab" === sType)
			this.Num.SetLvlSuff(this.Lvl, Asc.c_oAscNumberingSuff.Tab);
		else if ("none" === sType)
			this.Num.SetLvlSuff(this.Lvl, Asc.c_oAscNumberingSuff.None);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTablePr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @returns {"tablePr"}
	 */
	ApiTablePr.prototype.GetClassType = function()
	{
		return "tablePr";
	};
	/**
	 * Specify the number of columns which will comprise each table column band for this table style.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {number} nCount - The number of columns measured in positive integers.
	 */
	ApiTablePr.prototype.SetStyleColBandSize = function(nCount)
	{
		this.TablePr.TableStyleColBandSize = private_GetInt(nCount, 1, null);
		this.private_OnChange();
	};
	/**
	 * Specify the number of rows which will comprise each table row band for this table style.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {number} nCount - The number of rows measured in positive integers.
	 */
	ApiTablePr.prototype.SetStyleRowBandSize = function(nCount)
	{
		this.TablePr.TableStyleRowBandSize = private_GetInt(nCount, 1, null);
		this.private_OnChange();
	};
	/**
	 * Specify the alignment of the current table with respect to the text margins in the current section.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {("left" | "right" | "center")} sJcType - The alignment type used for the current table placement.
	 */
	ApiTablePr.prototype.SetJc = function(sJcType)
	{
		if ("left" === sJcType)
			this.TablePr.Jc = align_Left;
		else if ("right" === sJcType)
			this.TablePr.Jc = align_Right;
		else if ("center" === sJcType)
			this.TablePr.Jc = align_Center;
		this.private_OnChange();
	};
	/**
	 * Specify the shading which is applied to the extents of the current table.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type applied to the extents of the current table.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - The true value will disable the SetShd method use.
	 */
	ApiTablePr.prototype.SetShd = function(sType, r, g, b, isAuto)
	{
		this.TablePr.Shd = private_GetShd(sType, r, g, b, isAuto);
		this.private_OnChange();
	};
	/**
	 * Set the border which will be displayed at the top of the current table.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The top border style.
	 * @param {pt_8} nSize - The width of the current top border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the top part of the table measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTablePr.prototype.SetTableBorderTop = function(sType, nSize, nSpace, r, g, b)
	{
		this.TablePr.TableBorders.Top = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Set the border which will be displayed at the bottom of the current table.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The bottom border style.
	 * @param {pt_8} nSize - The width of the current bottom border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the bottom part of the table measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTablePr.prototype.SetTableBorderBottom = function(sType, nSize, nSpace, r, g, b)
	{
		this.TablePr.TableBorders.Bottom = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Set the border which will be displayed on the left of the current table.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The left border style.
	 * @param {pt_8} nSize - The width of the current left border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the left part of the table measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTablePr.prototype.SetTableBorderLeft = function(sType, nSize, nSpace, r, g, b)
	{
		this.TablePr.TableBorders.Left = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Set the border which will be displayed on the right of the current table.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The right border style.
	 * @param {pt_8} nSize - The width of the current right border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the right part of the table measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTablePr.prototype.SetTableBorderRight = function(sType, nSize, nSpace, r, g, b)
	{
		this.TablePr.TableBorders.Right = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Specify the border which will be displayed on all horizontal table cell borders which are not on an outmost edge
	 * of the parent table (all horizontal borders which are not the topmost or bottommost border).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The horizontal table cell border style.
	 * @param {pt_8} nSize - The width of the current border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the horizontal table cells of the table measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTablePr.prototype.SetTableBorderInsideH = function(sType, nSize, nSpace, r, g, b)
	{
		this.TablePr.TableBorders.InsideH = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Specify the border which will be displayed on all vertical table cell borders which are not on an outmost edge
	 * of the parent table (all vertical borders which are not the leftmost or rightmost border).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The vertical table cell border style.
	 * @param {pt_8} nSize - The width of the current border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the vertical table cells of the table measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTablePr.prototype.SetTableBorderInsideV = function(sType, nSize, nSpace, r, g, b)
	{
		this.TablePr.TableBorders.InsideV = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};

	/**
	 * Specify the amount of space which will be left between the bottom extent of the cell contents and the border
	 * of all table cells within the parent table (or table row).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {twips} nValue - The value for the amount of space below the bottom extent of the cell measured in
	 * twentieths of a point (1/1440 of an inch).
	 */
	ApiTablePr.prototype.SetTableCellMarginBottom = function(nValue)
	{
		this.TablePr.TableCellMar.Bottom = private_GetTableMeasure("twips", nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the amount of space which will be present between the left extent of the cell contents and the left
	 * border of all table cells within the parent table (or table row).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {twips} nValue - The value for the amount of space to the left extent of the cell measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiTablePr.prototype.SetTableCellMarginLeft = function(nValue)
	{
		this.TablePr.TableCellMar.Left = private_GetTableMeasure("twips", nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the amount of space which will be present between the right extent of the cell contents and the right
	 * border of all table cells within the parent table (or table row).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {twips} nValue - The value for the amount of space to the right extent of the cell measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiTablePr.prototype.SetTableCellMarginRight = function(nValue)
	{
		this.TablePr.TableCellMar.Right = private_GetTableMeasure("twips", nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the amount of space which will be present between the top extent of the cell contents and the top border
	 * of all table cells within the parent table (or table row).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {twips} nValue - The value for the amount of space above the top extent of the cell measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiTablePr.prototype.SetTableCellMarginTop = function(nValue)
	{
		this.TablePr.TableCellMar.Top = private_GetTableMeasure("twips", nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the default table cell spacing (the spacing between adjacent cells and the edges of the table).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {?twips} nValue - Spacing value measured in twentieths of a point (1/1440 of an inch). <code>"Null"</code> means no spacing will be applied.
	 */
	ApiTablePr.prototype.SetCellSpacing = function(nValue)
	{
		if (null === nValue)
			this.TablePr.TableCellSpacing = null;
		else
			this.TablePr.TableCellSpacing = private_Twips2MM(nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the indentation which will be added before the leading edge of the current table in the document
	 * (the left edge in a left-to-right table, and the right edge in a right-to-left table).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {twips} nValue - The indentation value measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiTablePr.prototype.SetTableInd = function(nValue)
	{
		this.TablePr.TableInd = private_Twips2MM(nValue);
		this.private_OnChange();
	};
	/**
	 * Set the preferred width for this table.
	 * <note>Tables are created with the {@link ApiTable#SetWidth} method properties set by default, which always override the {@link ApiTablePr#SetWidth} method properties. That is why there is no use to try and apply {@link ApiTablePr#SetWidth}, we recommend that you use the  {@link ApiTablePr#SetWidth}  method instead.</note>
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {TableWidth} sType - Type of the width value from one of the available width values types.
	 * @param {number} [nValue] - The table width value measured in positive integers.
	 */
	ApiTablePr.prototype.SetWidth = function(sType, nValue)
	{
		this.TablePr.TableW = private_GetTableMeasure(sType, nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the algorithm which will be used to lay out the contents of this table within the document.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {("autofit" | "fixed")} sType - The type of the table layout in the document.
	 */
	ApiTablePr.prototype.SetTableLayout = function(sType)
	{
		if ("autofit" === sType)
			this.TablePr.TableLayout = tbllayout_AutoFit;
		else if ("fixed" === sType)
			this.TablePr.TableLayout = tbllayout_Fixed;

		this.private_OnChange();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableRowPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiTableRowPr
	 * @typeofeditors ["CDE"]
	 * @returns {"tableRowPr"}
	 */
	ApiTableRowPr.prototype.GetClassType = function()
	{
		return "tableRowPr";
	};
	/**
     * Set the height of the current table row within the current table.
	 * @memberof ApiTableRowPr
	 * @typeofeditors ["CDE"]
	 * @param {("auto" | "atLeast")} sHRule - The rule to either apply or ignore the height value to the current table row. Use the <code>"atLeast"</code> value to enable the <code>SetHeight</code> method use.
	 * @param {twips} [nValue] - The height for the current table row measured in twentieths of a point (1/1440 of an inch). This value will be ignored if <code>sHRule="auto"<code>.
	 */
	ApiTableRowPr.prototype.SetHeight = function(sHRule, nValue)
	{
		if ("auto" === sHRule)
			this.RowPr.Height = new CTableRowHeight(0, Asc.linerule_Auto);
		else if ("atLeast" === sHRule)
			this.RowPr.Height = new CTableRowHeight(private_Twips2MM(nValue), Asc.linerule_AtLeast);

		this.private_OnChange();
	};
	/**
	 * Specify that the current table row will be repeated at the top of each new page 
     * wherever this table is displayed. This gives this table row the behavior of a 'header' row on 
     * each of these pages. This element can be applied to any number of rows at the top of the 
     * table structure in order to generate multi-row table headers.
	 * @memberof ApiTableRowPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isHeader - The true value will enable the SetTableHeader method use.
	 */
	ApiTableRowPr.prototype.SetTableHeader = function(isHeader)
	{
		this.RowPr.TableHeader = private_GetBoolean(isHeader);
		this.private_OnChange();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableCellPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @returns {"tableCellPr"}
	 */
	ApiTableCellPr.prototype.GetClassType = function()
	{
		return "tableCellPr";
	};
	/**
	 * Specify the shading applied to the contents of the table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type which will be applied to the contents of the current table cell.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - The true value will disable table cell contents shading.
	 */
	ApiTableCellPr.prototype.SetShd = function(sType, r, g, b, isAuto)
	{
		this.CellPr.Shd = private_GetShd(sType, r, g, b, isAuto);
		this.private_OnChange();
	};
	/**
	 * Specify the amount of space which will be left between the bottom extent of the cell contents and the border
	 * of a specific table cell within a table.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {?twips} nValue - The value for the amount of space below the bottom extent of the cell measured in twentieths
	 * of a point (1/1440 of an inch). If this value is <code>null</code>, then default table cell bottom margin will be used, otherwise
	 * the table cell bottom margin will be overridden with the specified value for the current cell.
	 */
	ApiTableCellPr.prototype.SetCellMarginBottom = function(nValue)
	{
		if (!this.CellPr.TableCellMar)
		{
			this.CellPr.TableCellMar =
			{
				Bottom : undefined,
				Left   : undefined,
				Right  : undefined,
				Top    : undefined
			};
		}

		if (null === nValue)
			this.CellPr.TableCellMar.Bottom = undefined;
		else
			this.CellPr.TableCellMar.Bottom = private_GetTableMeasure("twips", nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the amount of space which will be left between the left extent of the cell contents and 
	 * the border of a specific table cell within a table.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {?twips} nValue - The value for the amount of space to the left extent of the cell measured in twentieths
	 * of a point (1/1440 of an inch). If this value is <code>null<c/ode>, then default table cell left margin will be used, otherwise
	 * the table cell left margin will be overridden with the specified value for the current cell.
	 */
	ApiTableCellPr.prototype.SetCellMarginLeft = function(nValue)
	{
		if (!this.CellPr.TableCellMar)
		{
			this.CellPr.TableCellMar =
			{
				Bottom : undefined,
				Left   : undefined,
				Right  : undefined,
				Top    : undefined
			};
		}

		if (null === nValue)
			this.CellPr.TableCellMar.Left = undefined;
		else
			this.CellPr.TableCellMar.Left = private_GetTableMeasure("twips", nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the amount of space which will be left between the right extent of the cell contents and the border of a specific table cell within a table.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {?twips} nValue - The value for the amount of space to the right extent of the cell measured in twentieths
	 * of a point (1/1440 of an inch). If this value is <code>null</code>, then default table cell right margin will be used, otherwise
	 * the table cell right margin will be overridden with the specified value for the current cell.
	 */
	ApiTableCellPr.prototype.SetCellMarginRight = function(nValue)
	{
		if (!this.CellPr.TableCellMar)
		{
			this.CellPr.TableCellMar =
			{
				Bottom : undefined,
				Left   : undefined,
				Right  : undefined,
				Top    : undefined
			};
		}

		if (null === nValue)
			this.CellPr.TableCellMar.Right = undefined;
		else
			this.CellPr.TableCellMar.Right = private_GetTableMeasure("twips", nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the amount of space which will be left between the upper extent of the cell contents
	 * and the border of a specific table cell within a table.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {?twips} nValue - The value for the amount of space above the upper extent of the cell measured in twentieths
	 * of a point (1/1440 of an inch). If this value is <code>null</code>, then default table cell top margin will be used, otherwise
	 * the table cell top margin will be overridden with the specified value for the current cell.
	 */
	ApiTableCellPr.prototype.SetCellMarginTop = function(nValue)
	{
		if (!this.CellPr.TableCellMar)
		{
			this.CellPr.TableCellMar =
			{
				Bottom : undefined,
				Left   : undefined,
				Right  : undefined,
				Top    : undefined
			};
		}

		if (null === nValue)
			this.CellPr.TableCellMar.Top = undefined;
		else
			this.CellPr.TableCellMar.Top = private_GetTableMeasure("twips", nValue);
		this.private_OnChange();
	};
	/**
	 * Set the border which will be displayed at the bottom of the current table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The cell bottom border style.
	 * @param {pt_8} nSize - The width of the current cell bottom border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the bottom part of the table cell measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTableCellPr.prototype.SetCellBorderBottom = function(sType, nSize, nSpace, r, g, b)
	{
		this.CellPr.TableCellBorders.Bottom = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Set the border which will be displayed to the left of the current table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The cell left border style.
	 * @param {pt_8} nSize - The width of the current cell left border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the left part of the table cell measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTableCellPr.prototype.SetCellBorderLeft = function(sType, nSize, nSpace, r, g, b)
	{
		this.CellPr.TableCellBorders.Left = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Set the border which will be displayed to the right of the current table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The cell right border style.
	 * @param {pt_8} nSize - The width of the current cell right border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the right part of the table cell measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTableCellPr.prototype.SetCellBorderRight = function(sType, nSize, nSpace, r, g, b)
	{
		this.CellPr.TableCellBorders.Right = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Set the border which will be displayed at the top of the current table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {BorderType} sType - The cell top border style.
	 * @param {pt_8} nSize - The width of the current cell top border measured in eighths of a point.
	 * @param {pt} nSpace - The spacing offset in the top part of the table cell measured in points used to place this border.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 */
	ApiTableCellPr.prototype.SetCellBorderTop = function(sType, nSize, nSpace, r, g, b)
	{
		this.CellPr.TableCellBorders.Top = private_GetTableBorder(sType, nSize, nSpace, r, g, b);
		this.private_OnChange();
	};
	/**
	 * Set the preferred width for the current table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {TableWidth} sType - Type of the width value from one of the available width values types.
	 * @param {number} [nValue] - The table cell width value measured in positive integers.
	 */
	ApiTableCellPr.prototype.SetWidth = function(sType, nValue)
	{
		this.CellPr.TableCellW = private_GetTableMeasure(sType, nValue);
		this.private_OnChange();
	};
	/**
	 * Specify the vertical alignment for text contents within the current table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {("top" | "center" | "bottom")} sType - The available types of the vertical alignment for the text contents of the current table cell.
	 */
	ApiTableCellPr.prototype.SetVerticalAlign = function(sType)
	{
		if ("top" === sType)
			this.CellPr.VAlign = vertalignjc_Top;
		else if ("bottom" === sType)
			this.CellPr.VAlign = vertalignjc_Bottom;
		else if ("center" === sType)
			this.CellPr.VAlign = vertalignjc_Center;

		this.private_OnChange();
	};
	/**
	 * Specify the direction of the text flow for this table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {("lrtb" | "tbrl" | "btlr")} sType - The available types of the text direction in the table cell: <code>"lrtb"</code>
	 * - text direction left-to-right moving from top to bottom, <code>"tbrl"</code> - text direction top-to-bottom moving from right
	 * to left, <code>"btlr"</code> - text direction bottom-to-top moving from left to right.
	 */
	ApiTableCellPr.prototype.SetTextDirection = function(sType)
	{
		if ("lrtb" === sType)
			this.CellPr.TextDirection = textdirection_LRTB;
		else if ("tbrl" === sType)
			this.CellPr.TextDirection = textdirection_TBRL;
		else if ("btlr" === sType)
			this.CellPr.TextDirection = textdirection_BTLR;

		this.private_OnChange();
	};
	/**
	 * Specify how this table cell is laid out when the parent table is displayed in a document. This setting
	 * only affects the behavior of the cell when the {@link ApiTablePr#SetTableLayout} table layout for this table is set to use the <code>"autofit"</code> algorithm.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isNoWrap - The true value will enable the <code>SetNoWrap</code> method use.
	 */
	ApiTableCellPr.prototype.SetNoWrap = function(isNoWrap)
	{
		this.CellPr.NoWrap = private_GetBoolean(isNoWrap);
		this.private_OnChange();
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableStylePr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {"tableStylePr"}
	 */
	ApiTableStylePr.prototype.GetClassType = function()
	{
		return "tableStylePr";
	};
	/**
	 * Get the type of the current table conditional style.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {TableStyleOverrideType}
	 */
	ApiTableStylePr.prototype.GetType = function()
	{
		return this.Type;
	};
	/**
	 * Get the set of the text run properties which will be applied to all the text runs within the table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTextPr}
	 */
	ApiTableStylePr.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.TableStylePr.TextPr);
	};
	/**
	 * Get the set of the paragraph properties which will be applied to all the paragraphs within a table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParaPr}
	 */
	ApiTableStylePr.prototype.GetParaPr = function()
	{
		return new ApiParaPr(this, this.TableStylePr.ParaPr);
	};
	/**
	 * Get the set of the table properties which will be applied to all the regions within a table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTablePr}
	 */
	ApiTableStylePr.prototype.GetTablePr = function()
	{
		return new ApiTablePr(this, this.TableStylePr.TablePr);
	};
	/**
	 * Get the set of the table row properties which will be applied to all the rows within a table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableRowPr}
	 */
	ApiTableStylePr.prototype.GetTableRowPr = function()
	{
		return new ApiTableRowPr(this, this.TableStylePr.TableRowPr);
	};
	/**
	 * Get the set of the table cell properties which will be applied to all the cells within a table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableCellPr}
	 */
	ApiTableStylePr.prototype.GetTableCellPr = function()
	{
		return new ApiTableCellPr(this, this.TableStylePr.TableCellPr);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDrawing
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of the class based on this base class.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE", "CPE"]
	 * @returns {"drawing"}
	 */
	ApiDrawing.prototype.GetClassType = function()
	{
		return "drawing";
	};
	/**
	 * Set the size of the object (image, shape, chart) bounding box.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {EMU} nWidth - The object width measured in English measure units.
	 * @param {EMU} nHeight - The object height measured in English measure units.
	 */
	ApiDrawing.prototype.SetSize = function(nWidth, nHeight)
	{
		var fWidth = private_EMU2MM(nWidth);
		var fHeight = private_EMU2MM(nHeight);
		this.Drawing.setExtent(fWidth, fHeight);
		if(this.Drawing.GraphicObj && this.Drawing.GraphicObj.spPr && this.Drawing.GraphicObj.spPr.xfrm)
		{
			this.Drawing.GraphicObj.spPr.xfrm.setExtX(fWidth);
			this.Drawing.GraphicObj.spPr.xfrm.setExtY(fHeight);
		}
	};
	/**
	 * Set the wrapping type of this object (image, shape, chart). One of the following wrapping style types can be set:
	 * * <b>"inline"</b> - the object is considered to be a part of the text, like a character, so when the text moves, the object moves as well. In this case the positioning options are inaccessible.
	 * If one of the following styles is selected, the object can be moved independently of the text and positioned on the page exactly:
	 * * <b>"square"</b> - the text wraps the rectangular box that bounds the object.
	 * * <b>"tight"</b> - the text wraps the actual object edges.
	 * * <b>"through"</b> - the text wraps around the object edges and fills in the open white space within the object.
	 * * <b>"topAndBottom"</b> - the text is only above and below the object.
	 * * <b>"behind"</b> - the text overlaps the object.
	 * * <b>"inFront"</b> - the object overlaps the text.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {"inline" | "square" | "tight" | "through" | "topAndBottom" | "behind" | "inFront"} sType - The wrapping style type available for the object.
	 */
	ApiDrawing.prototype.SetWrappingStyle = function(sType)
	{
		if(this.Drawing)
		{
			if ("inline" === sType)
			{
				this.Drawing.Set_DrawingType(drawing_Inline);
				this.Drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
				this.Drawing.Set_BehindDoc(false);
			}
			else if ("square" === sType)
			{
				this.Drawing.Set_DrawingType(drawing_Anchor);
				this.Drawing.Set_WrappingType(WRAPPING_TYPE_SQUARE);
				this.Drawing.Set_BehindDoc(false);
			}
			else if ("tight" === sType)
			{
				this.Drawing.Set_DrawingType(drawing_Anchor);
				this.Drawing.Set_WrappingType(WRAPPING_TYPE_TIGHT);
				this.Drawing.Set_BehindDoc(true);
			}
			else if ("through" === sType)
			{
				this.Drawing.Set_DrawingType(drawing_Anchor);
				this.Drawing.Set_WrappingType(WRAPPING_TYPE_THROUGH);
				this.Drawing.Set_BehindDoc(true);
			}
			else if ("topAndBottom" === sType)
			{
				this.Drawing.Set_DrawingType(drawing_Anchor);
				this.Drawing.Set_WrappingType(WRAPPING_TYPE_TOP_AND_BOTTOM);
				this.Drawing.Set_BehindDoc(false);
			}
			else if ("behind" === sType)
			{
				this.Drawing.Set_DrawingType(drawing_Anchor);
				this.Drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
				this.Drawing.Set_BehindDoc(true);
			}
			else if ("inFront" === sType)
			{
				this.Drawing.Set_DrawingType(drawing_Anchor);
				this.Drawing.Set_WrappingType(WRAPPING_TYPE_NONE);
				this.Drawing.Set_BehindDoc(false);
			}
			this.Drawing.Check_WrapPolygon();
			if(this.Drawing.GraphicObj && this.Drawing.GraphicObj.setRecalculateInfo)
			{
				this.Drawing.GraphicObj.setRecalculateInfo();
			}
		}
	};
	/**
	 * Specify how the floating object will be horizontally aligned.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {RelFromH} [sRelativeFrom="page"] - The document element which will be taken as a countdown point for the object horizontal alignment.
	 * @param {("left" | "right" | "center")} [sAlign="left"] - The alingment type which will be used for the object horizontal alignment.
	 */
	ApiDrawing.prototype.SetHorAlign = function(sRelativeFrom, sAlign)
	{
		var nAlign        = private_GetAlignH(sAlign);
		var nRelativeFrom = private_GetRelativeFromH(sRelativeFrom);
		this.Drawing.Set_PositionH(nRelativeFrom, true, nAlign, false);
	};
	/**
	 * Specify how the floating object will be vertically aligned.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {RelFromV} [sRelativeFrom="page"] - The document element which will be taken as a countdown point for the object vertical alignment.
	 * @param {("top" | "bottom" | "center")} [sAlign="top"] - The alingment type which will be used for the object vertical alignment.
	 */
	ApiDrawing.prototype.SetVerAlign = function(sRelativeFrom, sAlign)
	{
		var nAlign        = private_GetAlignV(sAlign);
		var nRelativeFrom = private_GetRelativeFromV(sRelativeFrom);
		this.Drawing.Set_PositionV(nRelativeFrom, true, nAlign, false);
	};
	/**
	 * Set an absolute measurement for the horizontal positioning of the floating object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {RelFromH} sRelativeFrom - The document element which will be taken as a countdown point for the object horizontal alignment.
	 * @param {EMU} nDistance - The distance from the right side of the document element to the floating object measured in English measure units.
	 */
	ApiDrawing.prototype.SetHorPosition = function(sRelativeFrom, nDistance)
	{
		var nValue        = private_EMU2MM(nDistance);
		var nRelativeFrom = private_GetRelativeFromH(sRelativeFrom);
		this.Drawing.Set_PositionH(nRelativeFrom, false, nValue, false);
	};
	/**
	 * Set an absolute measurement for the vertical positioning of the floating object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {RelFromV} sRelativeFrom - The document element which will be taken as a countdown point for the object vertical alignment.
	 * @param {EMU} nDistance - The distance from the bottom part of the document element to the floating object measured in English measure units.
	 */
	ApiDrawing.prototype.SetVerPosition = function(sRelativeFrom, nDistance)
	{
		var nValue        = private_EMU2MM(nDistance);
		var nRelativeFrom = private_GetRelativeFromV(sRelativeFrom);
		this.Drawing.Set_PositionV(nRelativeFrom, false, nValue, false);
	};
	/**
	 * Specify the minimum distance which will be maintained between the edges of this drawing object and any
	 * subsequent text.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {EMU} nLeft - The distance from the left side of the current object and the subsequent text run measured in English measure units.
	 * @param {EMU} nTop - The distance from the top side of the current object and the preceding text run measured in English measure units.
	 * @param {EMU} nRight - The distance from the right side of the current object and the subsequent text run measured in English measure units.
	 * @param {EMU} nBottom - The distance from the bottom side of the current object and the subsequent text run measured in English measure units.
	 */
	ApiDrawing.prototype.SetDistances = function(nLeft, nTop, nRight, nBottom)
	{
		this.Drawing.Set_Distance(private_EMU2MM(nLeft), private_EMU2MM(nTop), private_EMU2MM(nRight), private_EMU2MM(nBottom));
	};
	/**
	 * Get the parent paragraph that contains the graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @return {ApiParagraph | null} - returns null if parent paragraph doesn't exist.
	 */
	ApiDrawing.prototype.GetParentParagraph = function()
	{
		var Paragraph = this.Drawing.GetParagraph();

		if (Paragraph)
			return new ApiParagraph(this.Drawing.GetParagraph());
		else 
			return null;
	};
	/**
	 * Get the parent content control that contains the graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @return {ApiBlockLvlSdt | null} - returns null if parent content control doesn't exist.
	 */
	ApiDrawing.prototype.GetParentContentControl = function()
	{
		var ParaParent = this.GetParentParagraph();

		if (ParaParent)
			return ParaParent.GetParentContentControl();
		return 	null;
	};
	/**
	 * Get the parent table that contains the graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @return {ApiTable | null} - returns null if parent table doesn't exist.
	 */
	ApiDrawing.prototype.GetParentTable = function()
	{
		var ParaParent = this.GetParentParagraph();

		if (ParaParent)
			return ParaParent.GetParentTable();
		return null;
	};
	/**
	 * Get the parent table cell that contains the graphic object.
	 * @typeofeditors ["CDE"]
	 * @return {ApiTableCell | null} - returns null if parent cell doesn't exist.
	 */
	ApiDrawing.prototype.GetParentTableCell = function()
	{
		var ParaParent = this.GetParentParagraph();

		if (ParaParent)
			return ParaParent.GetParentTableCell();
		return null;
	};
	/**
	 * Delete the graphic object. 
	 * @typeofeditors ["CDE"]
	 * @return {bool} - returns false if drawing object haven't parent.
	 */
	ApiDrawing.prototype.Delete = function()
	{
		var ParaParent = this.GetParentParagraph();

		if (ParaParent)
		{
			this.Drawing.PreDelete();
			var ApiParentRun = new ApiRun(this.Drawing.GetRun());
			ApiParentRun.Run.RemoveElement(this.Drawing);

			return true;
		}
		else 	 
			return false;
	};
	/**
	 * Copy the graphic object. 
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @return {ApiDrawing}
	 */
	ApiDrawing.prototype.Copy = function()
	{
		var oDrawing = this.Drawing.copy();

		if (this instanceof ApiShape
			|| this instanceof ApiChart
			|| this instanceof ApiImage)
			return new this.constructor(oDrawing.GraphicObj);

		return new ApiDrawing(oDrawing);
	};
	/**
	 * Wrap the graphic object with a rich text content control.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {number} nType - If nType === 1 -> returns ApiBlockLvlSdt, else -> return ApiDrawing.
	 * @return {ApiDrawing | ApiBlockLvlSdt}  
	 */
	ApiDrawing.prototype.InsertInContentControl = function(nType)
	{
		var Document			= editor.private_GetLogicDocument();
		var ContentControl		= null;
		var paragraphInControl	= null;
		var parentParagraph		= this.Drawing.GetParagraph();
		var paraIndex 			= -1;
		if (parentParagraph)
			paraIndex = parentParagraph.Index;

		if (paraIndex >= 0)
		{
			this.Select();
			ContentControl = new ApiBlockLvlSdt(Document.AddContentControl(1));
			Document.RemoveSelection();
		}
		else 
		{
			ContentControl		= new ApiBlockLvlSdt(new CBlockLevelSdt(Document, Document))
			ContentControl.Sdt.SetDefaultTextPr(Document.GetDirectTextPr());
			paragraphInControl	= ContentControl.Sdt.GetFirstParagraph();
			if (paragraphInControl.Content.length > 1)
			{
				paragraphInControl.RemoveFromContent(0, paragraphInControl.Content.length - 1);
				paragraphInControl.CorrectContent();
			}
			paragraphInControl.Add(this.Drawing);
			ContentControl.Sdt.SetShowingPlcHdr(false);
		}

		if (nType === 1)
			return ContentControl;
		else
			return this;
	};
	/**
	 * Insert a paragraph at the specified position.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {string | ApiParagraph} paragraph - Text or paragraph.
	 * @param {string} sPosition - The position where the text or paragraph will be inserted ("before" or "after" the drawing specified).
	 * @param {bool} beRNewPara - If "true" -> returns new paragraph, else -> returns this ApiDrawing.
	 * @return {ApiParagraph | ApiDrawing} - returns null if parent paragraph doesn't exist.
	 */
	ApiDrawing.prototype.InsertParagraph = function(paragraph, sPosition, beRNewPara)
	{
		var parentParagraph = this.GetParentParagraph();

		if (parentParagraph)
			if (beRNewPara)
				return parentParagraph.InsertParagraph(paragraph, sPosition, true)
			else 
			{
				parentParagraph.InsertParagraph(paragraph, sPosition, true);
				return this;
			}
		else 
			return null;
	};
	/**
	 * Select the graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 */	
	ApiDrawing.prototype.Select = function()
	{
		var Api = editor;
		var oDocument = Api.GetDocument();
		this.Drawing.SelectAsText();
		oDocument.Document.UpdateSelection();
	};
	/**
	 * Insert a break at the specified location in the main document.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {number}	breakType - If 0 -> page break, 1 -> line break.
	 * @param {string}	position  - The position where the page or line break will be inserted ("before" or "after" the drawing specified).
	 * @returns {bool}  - returns false if drawing object haven't parent run or params are invalid.
	 */	
	ApiDrawing.prototype.AddBreak = function(breakType, position)
	{
		var ParentRun	= (new ApiRun(this.Drawing.GetRun()));

		if (!ParentRun || position !== "before" && position !== "after" || breakType !== 1 && breakType !== 0)
			return false;

		if (breakType === 0)
		{
			if (position === "before")
				ParentRun.Run.Add_ToContent(ParentRun.Run.Content.indexOf(this.Drawing), new ParaNewLine(break_Page));
			else if (position === "after")
				ParentRun.Run.Add_ToContent(ParentRun.Run.Content.indexOf(this.Drawing) + 1, new ParaNewLine(break_Page));
		}
		else if (breakType === 1)
		{
			if (position === "before")
				ParentRun.Run.Add_ToContent(ParentRun.Run.Content.indexOf(this.Drawing), new ParaNewLine(break_Line));
			else if (position === "after")
				ParentRun.Run.Add_ToContent(ParentRun.Run.Content.indexOf(this.Drawing) + 1, new ParaNewLine(break_Line));
		}

		return true;
	};
	/**
	 * Horizontal reflection.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {bool} bFlip - Specifies if the figure will be flipped or not.
	 */	
	ApiDrawing.prototype.SetHorFlip = function(bFlip)
	{
		if (this.Drawing.GraphicObj && this.Drawing.GraphicObj.spPr && this.Drawing.GraphicObj.spPr.xfrm)
			this.Drawing.GraphicObj.spPr.xfrm.setFlipH(bFlip);
	};
	/**
	 * Vertical reflection.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {bool} bFlip - Specifies if the figure will be flipped or not.
	 * @returns {bool} - returns false if param is invalid.
	 */	
	ApiDrawing.prototype.SetVertFlip = function(bFlip)
	{
		if (typeof(bFlip) !== "boolean")
			return false;

		if (this.Drawing.GraphicObj && this.Drawing.GraphicObj.spPr && this.Drawing.GraphicObj.spPr.xfrm)
			this.Drawing.GraphicObj.spPr.xfrm.setFlipV(bFlip);
		
		return true;
	};
	/**
	 * Scale the height of the figure using the specified coefficient.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {number} coefficient - The coefficient by which the figure height will be scaled.
	 * @returns {bool} - return false if param is invalid.
	 */	
	ApiDrawing.prototype.ScaleHeight = function(coefficient)
	{
		if (typeof(coefficient) !== "number")
			return false;

		var currentHeight = this.Drawing.getXfrmExtY();
		var currentWidth  = this.Drawing.getXfrmExtX();

		this.Drawing.setExtent(currentWidth, currentHeight * coefficient);
		if(this.Drawing.GraphicObj && this.Drawing.GraphicObj.spPr && this.Drawing.GraphicObj.spPr.xfrm)
		{
			this.Drawing.GraphicObj.spPr.xfrm.setExtY(currentHeight * coefficient);
		}

		return true;
	};
	/**
	 * Scale the width of the graphic object using the specified coefficient.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {number} coefficient - The coefficient by which the figure height will be scaled.
	 * @returns {bool} - return false if param is invali.
	 */	
	ApiDrawing.prototype.ScaleWidth = function(coefficient)
	{
		if (typeof(coefficient) !== "number")
			return false;

		var currentHeight = this.Drawing.getXfrmExtY();
		var currentWidth  = this.Drawing.getXfrmExtX();

		this.Drawing.setExtent(currentWidth * coefficient, currentHeight);
		if(this.Drawing.GraphicObj && this.Drawing.GraphicObj.spPr && this.Drawing.GraphicObj.spPr.xfrm)
		{
			this.Drawing.GraphicObj.spPr.xfrm.setExtX(currentWidth * coefficient);
		}

		return true;
	};
	/**
	 * Set the fill formatting properties for the specified graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {ApiFill} oFill - The fill type used to fill the graphic object.
	 * @returns {bool} - returns false if param is invalid.
	 */	
	ApiDrawing.prototype.Fill = function(oFill)
	{
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		this.Drawing.GraphicObj.spPr.setFill(oFill.UniFill);
		return true;
	};
	/**
	 * Set the outline properties for the specified graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the graphic object outline.
	 * @returns {bool} - returns false if param is invalid.
	 */	
	ApiDrawing.prototype.SetOutLine = function(oStroke)
	{
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		this.Drawing.GraphicObj.spPr.setLn(oStroke.Ln);;
		return true;
	};
	/**
	 * Get the next inline drawing object. 
	 *  @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @returns {ApiDrawing | null} - returns null if drawing object is last.
	 */
	ApiDrawing.prototype.GetNextDrawing = function()
	{
		var oDocument				= editor.GetDocument();
		var GetAllDrawingObjects	= oDocument.GetAllDrawingObjects();
		var drawingIndex			= null;

		for (var Index = 0; Index < GetAllDrawingObjects.length; Index++)
		{
			if (GetAllDrawingObjects[Index].Drawing.Id === this.Drawing.Id)
			{
				drawingIndex = Index;
				break;
			}
		}

		if (drawingIndex !== null && GetAllDrawingObjects[drawingIndex + 1])
			return GetAllDrawingObjects[drawingIndex + 1];

		return null;
	};
	/**
	 * Get the previous inline drawing object. 
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @returns {ApiDrawing | null} - returns null if drawing object is first.
	 */
	ApiDrawing.prototype.GetPrevDrawing = function()
	{
		var oDocument				= editor.GetDocument();
		var GetAllDrawingObjects	= oDocument.GetAllDrawingObjects();
		var drawingIndex			= null;

		for (var Index = 0; Index < GetAllDrawingObjects.length; Index++)
		{
			if (GetAllDrawingObjects[Index].Drawing.Id === this.Drawing.Id)
			{
				drawingIndex = Index;
				break;
			}
		}

		if (drawingIndex !== null && GetAllDrawingObjects[drawingIndex - 1])
			return GetAllDrawingObjects[drawingIndex - 1];

		return null;
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiImage
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiImage
	 * @typeofeditors ["CDE", "CPE"]
	 * @returns {"image"}
	 */
	ApiImage.prototype.GetClassType = function()
	{
		return "image";
	};
	/**
	 * Get the next inline image. 
	 * @memberof ApiImage
	 * @typeofeditors ["CDE"]
	 * @returns {ApiImage | null} - returns null if image is last.
	 */
	ApiImage.prototype.GetNextImage	= function()
	{
		var oDocument	= editor.GetDocument();
		var AllImages	= oDocument.GetAllImages();
		var imageIndex	= null;

		for (var Index = 0; Index < AllImages.length; Index++)
		{
			if (AllImages[Index].Image.Id === this.Image.Id)
			{
				imageIndex = Index;
				break;
			}
		}

		if (imageIndex !== null && AllImages[imageIndex + 1])
			return AllImages[imageIndex + 1];

		return null;
	};
	/**
	 * Get the previous inline image. 
	 * @memberof ApiImage
	 * @typeofeditors ["CDE"]
	 * @returns {ApiImage | null} - returns null if image is first.
	 */
	ApiImage.prototype.GetPrevImage	= function()
	{
		var oDocument	= editor.GetDocument();
		var AllImages	= oDocument.GetAllImages();
		var imageIndex	= null;

		for (var Index = 0; Index < AllImages.length; Index++)
		{
			if (AllImages[Index].Image.Id === this.Image.Id)
			{
				imageIndex = Index;
				break;
			}
		}

		if (imageIndex !== null && AllImages[imageIndex - 1])
			return AllImages[imageIndex - 1];

		return null;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiShape
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiShape
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {"shape"}
	 */
	ApiShape.prototype.GetClassType = function()
	{
		return "shape";
	};
	/**
	 * Get the shape inner contents where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {?ApiDocumentContent}
	 */
	ApiShape.prototype.GetDocContent = function()
	{
		if(this.Shape && this.Shape.textBoxContent)
		{
			return new ApiDocumentContent(this.Shape.textBoxContent);
		}
		return null;
	};
	/**
	 * Get the shape inner contents where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {?ApiDocumentContent}
	 */
	ApiShape.prototype.GetContent = function()
	{
		if(this.Shape && this.Shape.textBoxContent)
		{
			return new ApiDocumentContent(this.Shape.textBoxContent);
		}
		return null;
	};
	
	/**
	 * Set the vertical alignment for the shape content where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CDE", "CSE"]
	 * @param {VerticalTextAlign} VerticalAlign - The type of the vertical alignment for the shape inner contents.
	 */
	ApiShape.prototype.SetVerticalTextAlign = function(VerticalAlign)
	{
		if(this.Shape)
		{
			switch(VerticalAlign)
			{
				case "top":
				{
					this.Shape.setVerticalAlign(4);
					break;
				}
				case "center":
				{
					this.Shape.setVerticalAlign(1);
					break;
				}
				case "bottom":
				{
					this.Shape.setVerticalAlign(0);
					break;
				}
			}
		}
	};
	/**
	 * Set text paddings
	 * @memberof ApiShape
	 * @typeofeditors ["CDE", "CSE"]
	 * @param {?EMU} nLeft - Left padding.
	 * @param {?EMU} nTop - Top padding.
	 * @param {?EMU} nRight - Right padding.
	 * @param {?EMU} nBottom - Bottom padding.
	 */
	ApiShape.prototype.SetPaddings = function(nLeft, nTop, nRight, nBottom)
	{
		if(this.Shape)
		{
			this.Shape.setPaddings({
				Left: AscFormat.isRealNumber(nLeft) ? private_EMU2MM(nLeft) : null,
				Top: AscFormat.isRealNumber(nTop) ? private_EMU2MM(nTop) : null,
				Right: AscFormat.isRealNumber(nRight) ? private_EMU2MM(nRight) : null,
				Bottom: AscFormat.isRealNumber(nBottom) ? private_EMU2MM(nBottom) : null
			});
		}
	};
	/**
	 * Get the next inline shape. 
	 * @memberof ApiShape
	 * @typeofeditors ["CDE"]
	 * @returns {ApiShape | null} - returns null if shape is last.
	 */
	ApiShape.prototype.GetNextShape = function()
	{
		var oDocument	= editor.GetDocument();
		var AllShapes	= oDocument.GetAllShapes();
		var shapeIndex	= null;

		for (var Index = 0; Index < AllShapes.length; Index++)
		{
			if (AllShapes[Index].Shape.Id === this.Shape.Id)
			{
				shapeIndex = Index;
				break;
			}
		}

		if (shapeIndex !== null && AllShapes[shapeIndex + 1])
			return AllShapes[shapeIndex + 1];

		return null;
	};
	/**
	 * Get the previous inline shape. 
	 * @memberof ApiShape
	 * @typeofeditors ["CDE"]
	 * @returns {ApiShape | null} - returns null is shape is first.
	 */
	ApiShape.prototype.GetPrevShape	= function()
	{
		var oDocument	= editor.GetDocument();
		var AllShapes	= oDocument.GetAllShapes();
		var shapeIndex	= null;

		for (var Index = 0; Index < AllShapes.length; Index++)
		{
			if (AllShapes[Index].Shape.Id === this.Shape.Id)
			{
				shapeIndex = Index;
				break;
			}
		}

		if (shapeIndex !== null && AllShapes[shapeIndex - 1])
			return AllShapes[shapeIndex - 1];

		return null;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiChart
	//
	//------------------------------------------------------------------------------------------------------------------
	/**
	 * Get the type of this class.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @returns {"chart"}
	 */
	ApiChart.prototype.GetClassType = function()
	{
		return "chart";
	};

	ApiChart.prototype.CreateTitle = function(sTitle, nFontSize){
		if(!this.Chart)
		{
			return null;
		}
		if(typeof sTitle === "string" && sTitle.length > 0){
			var oTitle = new AscFormat.CTitle();
			oTitle.setOverlay(false);
			oTitle.setTx(new AscFormat.CChartText());
			var oTextBody = AscFormat.CreateTextBodyFromString(sTitle, this.Chart.getDrawingDocument(), oTitle.tx);
			if(AscFormat.isRealNumber(nFontSize)){
				oTextBody.content.SetApplyToAll(true);
				oTextBody.content.AddToParagraph(new ParaTextPr({ FontSize : nFontSize}));
				oTextBody.content.SetApplyToAll(false);
			}
			oTitle.tx.setRich(oTextBody);
			return oTitle;
		}
		return null;
	};


	/**
	 *  Specify the chart title.
	 *  @memberof ApiChart
	 *  @typeofeditors ["CDE"]
	 *  @param {string} sTitle - The title which will be displayed for the current chart.
	 *  @param {pt} nFontSize - The text size value measured in points.
	 *  @param {?bool} bIsBold - Specifies if the chart title is written in bold font or not.
	 */
	ApiChart.prototype.SetTitle = function (sTitle, nFontSize, bIsBold)
	{
		AscFormat.builder_SetChartTitle(this.Chart, sTitle, nFontSize, bIsBold);
	};

	/**
	 *  Specify the chart horizontal axis title.
	 *  @memberof ApiChart
	 *  @typeofeditors ["CDE"]
	 *  @param {string} sTitle - The title which will be displayed for the horizontal axis of the current chart.
	 *  @param {pt} nFontSize - The text size value measured in points.
	 *  @param {?bool} bIsBold - Specifies if the horizontal axis title is written in bold font or not.
	 * */
	ApiChart.prototype.SetHorAxisTitle = function (sTitle, nFontSize, bIsBold)
	{
		AscFormat.builder_SetChartHorAxisTitle(this.Chart, sTitle, nFontSize, bIsBold);
	};

	/**
	 *  Specify the chart vertical axis title.
	 *  @memberof ApiChart
	 *  @typeofeditors ["CDE"]
	 *  @param {string} sTitle - The title which will be displayed for the vertical axis of the current chart.
	 *  @param {pt} nFontSize - The text size value measured in points.
	 *  @param {?bool} bIsBold - Specifies if the vertical axis title is written in bold font or not.
	 * */
	ApiChart.prototype.SetVerAxisTitle = function (sTitle, nFontSize, bIsBold)
	{
		AscFormat.builder_SetChartVertAxisTitle(this.Chart, sTitle, nFontSize, bIsBold);
	};

	/**
	 * Specify a vertical axis orientation.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {bool} bIsMinMax - The <code>true</code> value will set the normal data direction for the vertical axis (from minimum to maximum).
	 * */
	ApiChart.prototype.SetVerAxisOrientation = function(bIsMinMax){
		AscFormat.builder_SetChartVertAxisOrientation(this.Chart, bIsMinMax);
	};

	/**
	 * Specify a horizontal axis orientation.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {bool} bIsMinMax - The <code>true</code> value will set the normal data direction for the horizontal axis (from minimum to maximum).
	 * */
	ApiChart.prototype.SetHorAxisOrientation = function(bIsMinMax){
		AscFormat.builder_SetChartHorAxisOrientation(this.Chart, bIsMinMax);
	};

	/**
	 * Specify the chart legend position.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {"left" | "top" | "right" | "bottom" | "none"} sLegendPos - The position of the chart legend inside the chart window.
	 * */
	ApiChart.prototype.SetLegendPos = function(sLegendPos)
	{
		if(this.Chart && this.Chart.chart)
		{
			if(sLegendPos === "none")
			{
				if(this.Chart.chart.legend)
				{
					this.Chart.chart.setLegend(null);
				}
			}
			else
			{
				var nLegendPos = null;
				switch(sLegendPos)
				{
					case "left":
					{
						nLegendPos = Asc.c_oAscChartLegendShowSettings.left;
						break;
					}
					case "top":
					{
						nLegendPos = Asc.c_oAscChartLegendShowSettings.top;
						break;
					}
					case "right":
					{
						nLegendPos = Asc.c_oAscChartLegendShowSettings.right;
						break;
					}
					case "bottom":
					{
						nLegendPos = Asc.c_oAscChartLegendShowSettings.bottom;
						break;
					}
				}
				if(null !== nLegendPos)
				{
					if(!this.Chart.chart.legend)
					{
						this.Chart.chart.setLegend(new AscFormat.CLegend());
					}
					if(this.Chart.chart.legend.legendPos !== nLegendPos)
						this.Chart.chart.legend.setLegendPos(nLegendPos);
					if(this.Chart.chart.legend.overlay !== false)
					{
						this.Chart.chart.legend.setOverlay(false);
					}
				}
			}
		}
	};

	/**
	 * Specify a legend font size.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 * */
	ApiChart.prototype.SetLegendFontSize = function(nFontSize)
	{
		AscFormat.builder_SetLegendFontSize(this.Chart, nFontSize);
	};

	/**
	 * Specify which chart data labels are shown for the chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bShowSerName - Whether to show or hide the source table column names used for the data which the chart will be build from.
	 * @param {boolean} bShowCatName - Whether to show or hide the source table row names used for the data which the chart will be build from.
	 * @param {boolean} bShowVal - Whether to show or hide the chart data values.
	 * @param {boolean} bShowPercent - Whether to show or hide the percent for the data values (works with stacked chart types).
	 * */
	ApiChart.prototype.SetShowDataLabels = function(bShowSerName, bShowCatName, bShowVal, bShowPercent)
	{
		AscFormat.builder_SetShowDataLabels(this.Chart, bShowSerName, bShowCatName, bShowVal, bShowPercent);
	};


	/**
	 * Spicify the show options for data labels
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {number} nSeriesIndex - The series index from the array of the data used to build the chart from.
	 * @param {number} nPointIndex - The point index from this series.
	 * @param {boolean} bShowSerName - Whether to show or hide the source table column names used for the data which the chart will be build from.
	 * @param {boolean} bShowCatName - Whether to show or hide the source table row names used for the data which the chart will be build from.
	 * @param {boolean} bShowVal - Whether to show or hide the chart data values.
	 * @param {boolean} bShowPercent - Whether to show or hide the percent for the data values (works with stacked chart types).
	 * */
	ApiChart.prototype.SetShowPointDataLabel = function(nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPercent)
	{
		AscFormat.builder_SetShowPointDataLabel(this.Chart, nSeriesIndex, nPointIndex, bShowSerName, bShowCatName, bShowVal, bShowPercent);
	};

	/**
	 * Spicify tick labels position for vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart vertical tick labels.
	 * */
	ApiChart.prototype.SetVertAxisTickLabelPosition = function(sTickLabelPosition)
	{
		AscFormat.builder_SetChartVertAxisTickLablePosition(this.Chart, sTickLabelPosition);
	};

	/**
	 * Spicify tick labels position for horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart horizontal tick labels.
	 * */
	ApiChart.prototype.SetHorAxisTickLabelPosition = function(sTickLabelPosition)
	{
		AscFormat.builder_SetChartHorAxisTickLablePosition(this.Chart, sTickLabelPosition);
	};

	/**
	 * Specify major tick mark for horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetHorAxisMajorTickMark = function(sTickMark){
		AscFormat.builder_SetChartHorAxisMajorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specify minor tick mark for horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetHorAxisMinorTickMark = function(sTickMark){
		AscFormat.builder_SetChartHorAxisMinorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specify major tick mark for vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */

	ApiChart.prototype.SetVertAxisMajorTickMark = function(sTickMark){
		AscFormat.builder_SetChartVerAxisMajorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specify minor tick mark for vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetVertAxisMinorTickMark = function(sTickMark){
		AscFormat.builder_SetChartVerAxisMinorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specify major vertical gridline's visual properties.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMajorVerticalGridlines = function(oStroke)
	{
		AscFormat.builder_SetVerAxisMajorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};

	/**
	 * Specify minor vertical gridline's visual properties.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMinorVerticalGridlines = function(oStroke)
	{
		AscFormat.builder_SetVerAxisMinorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};


	/**
	 * Specify major horizontal gridline's visual properties.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMajorHorizontalGridlines = function(oStroke)
	{
		AscFormat.builder_SetHorAxisMajorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};

	/**
	 * Specify minor vertical gridline's visual properties.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMinorHorizontalGridlines = function(oStroke)
	{
		AscFormat.builder_SetHorAxisMinorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};


	/**
	 * Specify font size for labels of horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 */
	ApiChart.prototype.SetHorAxisLablesFontSize = function(nFontSize){
		AscFormat.builder_SetHorAxisFontSize(this.Chart, nFontSize);
	};

	/**
	 * Specify font size for labels of vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 */
	ApiChart.prototype.SetVertAxisLablesFontSize = function(nFontSize){
		AscFormat.builder_SetVerAxisFontSize(this.Chart, nFontSize);
	};

	/**
	 * Get the next inline chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @returns {ApiChart | null} - returns null if chart is last.
	 */
	ApiChart.prototype.GetNextChart = function()
	{
		var oDocument	= editor.GetDocument();
		var AllCharts	= oDocument.GetAllCharts();
		var chartIndex	= null;

		for (var Index = 0; Index < AllCharts.length; Index++)
		{
			if (AllCharts[Index].Chart.Id === this.Chart.Id)
			{
				chartIndex = Index;
				break;
			}
		}

		if (chartIndex !== null && AllCharts[chartIndex + 1])
			return AllCharts[chartIndex + 1];

		return null;
	};

	/**
	 * Get the previous inline chart. 
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @returns {ApiChart | null} - return null if char if first.
	 */
	ApiChart.prototype.GetPrevChart	= function()
	{
		var oDocument	= editor.GetDocument();
		var AllCharts	= oDocument.GetAllCharts();
		var chartIndex	= null;

		for (var Index = 0; Index < AllCharts.length; Index++)
		{
			if (AllCharts[Index].Chart.Id === this.Chart.Id)
			{
				chartIndex = Index;
				break;
			}
		}

		if (chartIndex !== null && AllCharts[chartIndex - 1])
			return AllCharts[chartIndex -1];

		return null;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiFill
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiFill
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"fill"}
	 */
	ApiFill.prototype.GetClassType = function()
	{
		return "fill";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiStroke
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiStroke
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"stroke"}
	 */
	ApiStroke.prototype.GetClassType = function()
	{
		return "stroke";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiGradientStop
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiGradientStop
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"gradientStop"}
	 */
	ApiGradientStop.prototype.GetClassType = function ()
	{
		return "gradientStop"
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiUniColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of the class based on this base class.
	 * @memberof ApiUniColor
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"uniColor"}
	 */
	ApiUniColor.prototype.GetClassType = function ()
	{
		return "uniColor"
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRGBColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiRGBColor
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"rgbColor"}
	 */
	ApiRGBColor.prototype.GetClassType = function ()
	{
		return "rgbColor"
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiSchemeColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiSchemeColor
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"schemeColor"}
	 */
	ApiSchemeColor.prototype.GetClassType = function ()
	{
		return "schemeColor"
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPresetColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiPresetColor
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"presetColor"}
	 */
	ApiPresetColor.prototype.GetClassType = function ()
	{
		return "presetColor"
	};

	/**
	 * Get the type of this class.
	 * @memberof ApiBullet
	 * @typeofeditors ["CSE", "CPE"]
	 * @returns {"bullet"}
	 */
	ApiBullet.prototype.GetClassType = function()
	{
		return "bullet";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiInlineLvlSdt
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {"inlineLvlSdt"}
	 */
	ApiInlineLvlSdt.prototype.GetClassType = function()
	{
		return "inlineLvlSdt";
	};

	/**
	 * Set the lock to the current inline text content control.
	 * <b>"contentLocked"</b> - content cannot be edited.
	 * <b>"sdtContentLocked"</b> - content cannot be edited and BlockLvlSdt cannot be deleted.
	 * <b>"sdtLocked"</b> - BlockLvlSdt cannot be deleted.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {"contentLocked" | "sdtContentLocked" | "sdtLocked"} sLockType - The lock type applied to the inline text content control.
	 */
	ApiInlineLvlSdt.prototype.SetLock = function(sLockType)
	{
		var nLock = c_oAscSdtLockType.Unlocked;
		if ("contentLocked" === sLockType)
			nLock = c_oAscSdtLockType.ContentLocked;
		else if ("sdtContentLocked" === sLockType)
			nLock = c_oAscSdtLockType.SdtContentLocked;
		else if ("sdtLocked" === sLockType)
			nLock = c_oAscSdtLockType.SdtLocked;

		this.Sdt.SetContentControlLock(nLock);
	};

	/**
	 * Get the lock type of this container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {SdtLock}
	 */
	ApiInlineLvlSdt.prototype.GetLock = function()
	{
		var nLock = this.Sdt.GetContentControlLock();

		var sResult = "unlocked";

		if (c_oAscSdtLockType.ContentLocked === nLock)
			sResult = "contentLocked";
		else if (c_oAscSdtLockType.SdtContentLocked === nLock)
			sResult = "sdtContentLocked";
		else if (c_oAscSdtLockType.SdtLocked === nLock)
			sResult = "sdtLocked";

		return sResult;
	};

	/**
	 * Add a string tag to the current inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sTag - The tag which will be added to the current inline text content control.
	 */
	ApiInlineLvlSdt.prototype.SetTag = function(sTag)
	{
		this.Sdt.SetTag(sTag);
	};

	/**
	 * Get the tag attribute for this container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiInlineLvlSdt.prototype.GetTag = function()
	{
		return this.Sdt.GetTag();
	};

	/**
	 * Add a string label to the current inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sLabel - The label which will be added to the current inline text content control. Can be a positive or negative integer from **-2147483647** to **2147483647**.
	 */
	ApiInlineLvlSdt.prototype.SetLabel = function(sLabel)
	{
		this.Sdt.SetLabel(sLabel);
	};

	/**
	 * Get the label attribute for this container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiInlineLvlSdt.prototype.GetLabel = function()
	{
		return this.Sdt.GetLabel();
	};

	/**
	 * Set the alias attribute for this container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sAlias The alias which will be added to the current inline text content control.
	 */
	ApiInlineLvlSdt.prototype.SetAlias = function(sAlias)
	{
		this.Sdt.SetAlias(sAlias);
	};

	/**
	 * Get the alias attribute for this container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiInlineLvlSdt.prototype.GetAlias = function()
	{
		return this.Sdt.GetAlias();
	};

	/**
	 * Get the number of elements in the current inline text content control. The text content 
     * control is created with one text run present in it by default, so even without any 
     * element added this method will return the value of '1'.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiInlineLvlSdt.prototype.GetElementsCount = function()
	{
		return this.Sdt.Content.length;
	};

	/**
	 * Get the element of the current inline text content control using the position specified.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The position where the element which content we want to get must be located.
	 * @returns {?ParagraphContent}
	 */
	ApiInlineLvlSdt.prototype.GetElement = function(nPos)
	{
		if (nPos < 0 || nPos >= this.Sdt.Content.length)
			return null;

		return private_GetSupportedParaElement(this.Sdt.Content[nPos]);
	};

	/**
	 * Remove the element using the position specified from the current inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The position of the element which we want to remove from the current inline text content control.
	 */
	ApiInlineLvlSdt.prototype.RemoveElement = function(nPos)
	{
		if (nPos < 0 || nPos >= this.Sdt.Content.length)
			return;

		this.Sdt.RemoveFromContent(nPos, 1);
		this.Sdt.CorrectContent();
	};

	/**
	 * Remove all the elements from the current inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {bool} - returns false if control haven't elements.
	 */
	ApiInlineLvlSdt.prototype.RemoveAllElements = function()
	{
		if (this.Sdt.Content.length > 0)
		{
			this.Sdt.RemoveFromContent(0, this.Sdt.Content.length);
			this.Sdt.CorrectContent();

			return true;
		}

		return false;
	};

	/**
	 * Add an element to the inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {ParagraphContent} oElement - The document element which will be added at the position specified. Returns **false** if the type of *oElement* is not supported by an inline text content control.
	 * @param {number} [nPos] - The position of the element where it will be added to the current inline text content control. If this value is not specified, then the element will be added to the end of the current inline text content control.
	 * @returns {bool} - returns false if oElement unsupported.
	 */
	ApiInlineLvlSdt.prototype.AddElement = function(oElement, nPos)
	{
		if (!private_IsSupportedParaElement(oElement) || nPos < 0 || nPos > this.Sdt.Content.length)
			return false;

		var oParaElement = oElement.private_GetImpl();
		if (undefined !== nPos)
		{
			this.Sdt.AddToContent(nPos, oParaElement);
		}
		else
		{
			private_PushElementToParagraph(this.Sdt, oParaElement);
		}

		return true;
	};

	/**
	 * Add an element to the end of inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The position where the current element will be added.
	 * @param {DocumentElement} oElement - The document element which will be added at the current position.
	 * @returns {bool} - returns false if oElement unsupported.
	 */
	ApiInlineLvlSdt.prototype.Push = function(oElement)
	{
		if (!private_IsSupportedParaElement(oElement))
			return false;

		if (this.Sdt.IsShowingPlcHdr())
		{
			this.Sdt.RemoveFromContent(0, this.Sdt.GetElementsCount(), false);
			this.Sdt.SetShowingPlcHdr(false);
		}

		var oParaElement = oElement.private_GetImpl();

		this.Sdt.AddToContent(this.Sdt.GetElementsCount(), oParaElement);

		return true;
	};

	/**
	 * Add text to the current content control. 
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {String} sText - The text which will be added to the content control.
	 * @returns {bool} - returns false if param is invalid.
	 */
	ApiInlineLvlSdt.prototype.AddText = function(sText)
	{
		if (typeof sText === "string")
		{
			var newRun = editor.CreateRun();
			newRun.AddText(sText);
			this.AddElement(newRun, this.GetElementsCount())

			return true;
		}

		return false;
	};

	/**
	 * Remove content control and content. If keepContent is true, the content is not deleted.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {bool} keepContent - Specifies if the content will be deleted or not.
	 * @returns {bool} - returns false if control haven't parent paragraph.
	 */
	ApiInlineLvlSdt.prototype.Delete = function(keepContent)
	{
		if (this.Sdt.Paragraph)
		{
			if (keepContent)
			{
				this.Sdt.RemoveContentControlWrapper();
			}
			else 
			{
				this.Sdt.PreDelete();
				var controlIndex = this.Sdt.Paragraph.Content.indexOf(this.Sdt);
				this.Sdt.Paragraph.RemoveFromContent(controlIndex, 1);
			}

			return true;
		}

		return false;
	};

	/**
	 * Apply text settings to content of content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The properties that will be set to the content of content control.
	 * @returns {ApiInlineLvlSdt} this.
	 */
	ApiInlineLvlSdt.prototype.SetTextPr = function(oTextPr)
	{
		for (var Index = 0; Index < this.Sdt.Content.length; Index++)
		{
			var Run = new ApiRun(this.Sdt.Content[Index]);
			var runTextPr = Run.GetTextPr();
			runTextPr.TextPr.Merge(oTextPr.TextPr);
			runTextPr.private_OnChange();
		}

		return this;
	};

	/**
	 * Get the content control that contains the current content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiBlockLvlSdt | null} - returns null if parent content control doesn't exist.
	 */
	ApiInlineLvlSdt.prototype.GetParentContentControl = function()
	{
		var parentContentControls = this.Sdt.GetParentContentControls();

		if (parentContentControls[parentContentControls.length - 2])
		{
			var ContentControl = parentContentControls[parentContentControls.length - 2];

			if (ContentControl instanceof CBlockLevelSdt)
				return new ApiBlockLvlSdt(ContentControl);
			else if (ContentControl instanceof CInlineLevelSdt)
				return new ApiInlineLvlSdt(ContentControl);
		}

		return null; 
	};

	/**
	 * Get the table that contains the content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiTable | null} - returns null if parent table doesn't exist.  
	 */
	ApiInlineLvlSdt.prototype.GetParentTable = function()
	{
		var documentPos = this.Sdt.GetDocumentPositionFromObject();

		for (var Index = documentPos.length - 1; Index >= 1; Index--)
		{
			if (documentPos[Index].Class)
				if (documentPos[Index].Class instanceof CTable)
					return new ApiTable(documentPos[Index].Class);
		}

		return null;
	};

	/**
	 * Get the table cell that contains the content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiTableCell | null} - return null if parent cell doesn't exist.  
	 */
	ApiInlineLvlSdt.prototype.GetParentTableCell = function()
	{
		var documentPos = this.Sdt.GetDocumentPositionFromObject();

		for (var Index = documentPos.length - 1; Index >= 1; Index--)
		{
			if (documentPos[Index].Class.Parent)
				if (documentPos[Index].Class.Parent instanceof CTableCell)
					return new ApiTableCell(documentPos[Index].Class.Parent);
		}

		return null;
	};

	/**
	 * Return a Range object that represents the part of the document contained in the specified content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiInlineLvlSdt.prototype.GetRange = function(Start, End)
	{
		var Range = new ApiRange(this.Sdt, Start, End);

		return Range;
	};

	/**
	 * Create a copy of an inline content control. Ingonore comments, footnote references, complex fields.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {ApiInlineLvlSdt}
	 */
	ApiInlineLvlSdt.prototype.Copy = function()
	{
		var oInlineSdt = this.Sdt.Copy(false, {
			SkipComments          : true,
			SkipAnchors           : true,
			SkipFootnoteReference : true,
			SkipComplexFields     : true
		});

		return new ApiInlineLvlSdt(oInlineSdt);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiBlockLvlSdt
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Get the type of this class.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {"blockLvlSdt"}
	 */
	ApiBlockLvlSdt.prototype.GetClassType = function()
	{
		return "blockLvlSdt";
	};

	/**
	 * Set the lock to the current inline text content control.
	 * <b>"contentLocked"</b> - content cannot be edited.
	 * <b>"sdtContentLocked"</b> - content cannot be edited and BlockLvlSdt cannot be deleted.
	 * <b>"sdtLocked"</b> - BlockLvlSdt cannot be deleted.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {"contentLocked" | "sdtContentLocked" | "sdtLocked"} sLockType - The type of the lock applied to the inline text content control.
	 */
	ApiBlockLvlSdt.prototype.SetLock = function(sLockType)
	{
		var nLock = c_oAscSdtLockType.Unlocked;
		if ("contentLocked" === sLockType)
			nLock = c_oAscSdtLockType.ContentLocked;
		else if ("sdtContentLocked" === sLockType)
			nLock = c_oAscSdtLockType.SdtContentLocked;
		else if ("sdtLocked" === sLockType)
			nLock = c_oAscSdtLockType.SdtLocked;

		this.Sdt.SetContentControlLock(nLock);
	};

	/**
	 * Get the lock type of this container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {SdtLock}
	 */
	ApiBlockLvlSdt.prototype.GetLock = function()
	{
		var nLock = this.Sdt.GetContentControlLock();

		var sResult = "unlocked";

		if (c_oAscSdtLockType.ContentLocked === nLock)
			sResult = "contentLocked";
		else if (c_oAscSdtLockType.SdtContentLocked === nLock)
			sResult = "sdtContentLocked";
		else if (c_oAscSdtLockType.SdtLocked === nLock)
			sResult = "sdtLocked";

		return sResult;
	};

	/**
	 * Set the tag attribute for this container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sTag - The tag which will be added to the current container.
	 */
	ApiBlockLvlSdt.prototype.SetTag = function(sTag)
	{
		this.Sdt.SetTag(sTag);
	};

	/**
	 * Get the tag attribute for this container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiBlockLvlSdt.prototype.GetTag = function()
	{
		return this.Sdt.GetTag();
	};

	/**
	 * Set the label attribute for this container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sLabel - The label which will be added to the current container. Can be a positive or negative integer from **-2147483647** to **2147483647**.
	 */
	ApiBlockLvlSdt.prototype.SetLabel = function(sLabel)
	{
		this.Sdt.SetLabel(sLabel);
	};

	/**
	 * Get the label attribute for this container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiBlockLvlSdt.prototype.GetLabel = function()
	{
		return this.Sdt.GetLabel();
	};

	/**
	 * Set the alias attribute for this container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sAlias - The alias which will be added to the current container.
	 */
	ApiBlockLvlSdt.prototype.SetAlias = function(sAlias)
	{
		this.Sdt.SetAlias(sAlias);
	};

	/**
	 * Get the alias attribute for this container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiBlockLvlSdt.prototype.GetAlias = function()
	{
		return this.Sdt.GetAlias();
	};

	/**
	 * Get the content of this container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {ApiDocumentContent}
	 */
	ApiBlockLvlSdt.prototype.GetContent = function()
	{
		return new ApiDocumentContent(this.Sdt.GetContent());
	};

	/**
	 * Get the collection of content control objects in the content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {ApiBlockLvlSdt[] | ApiInlineLvlSdt[]}
	 */
	ApiBlockLvlSdt.prototype.GetAllContentControls = function()
	{
		var arrContentControls    = [];
		var arrApiContentControls = [];
		this.Sdt.Content.GetAllContentControls(arrContentControls);

		for (var Index = 0, nCount = arrContentControls.length; Index < nCount; Index++)
		{
			var oControl = arrContentControls[Index];

			if (oControl instanceof CBlockLevelSdt)
				arrApiContentControls.push(new ApiBlockLvlSdt(oControl));
			else if (oControl instanceof CInlineLevelSdt)
				arrApiContentControls.push(new ApiInlineLvlSdt(oControl));
		}

		return arrApiContentControls;
	};

	/**
	 * Get a collection of paragraph objects in a content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParagraph[]}
	 */
	ApiBlockLvlSdt.prototype.GetAllParagraphs = function()
	{
		var arrParagraphs		= [];
		var arrApiParagraphs	= [];

		this.Sdt.GetAllParagraphs({All : true}, arrParagraphs);

		for (var Index = 0, nCount = arrParagraphs.length; Index < nCount; Index++)
		{
			arrApiParagraphs.push(new ApiParagraph(arrParagraphs[Index]));
		}

		return arrApiParagraphs;

	};

	/**
	 * Get a collection of tables on a given absolute page.
	 * <note>This method can be a little bit slow, because it runs the document calculation
	 * process to arrange tables on the specified page.</note>
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param nPage - Page number, if not specified, will return an empty array
	 * @return {ApiTable[]}  
	 */
	ApiBlockLvlSdt.prototype.GetAllTablesOnPage = function(nPage)
	{
		var oLogicDocument = this.Sdt.GetLogicDocument();
		if (oLogicDocument)
			oLogicDocument.private_Recalculate(undefined, true, nPage + 1);

		var arrTables		= this.Sdt.GetAllTablesOnPage(nPage);
		var arrApiTables	= [];

		for (var Index = 0, nCount = arrTables.length; Index < nCount; Index++)
		{
			arrApiTables.push(new ApiTable(arrTables[Index].Table));
		}

		return arrApiTables;
	};

	/**
	 * Clear the contents from a content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {bool} - returns true.
	 */
	ApiBlockLvlSdt.prototype.RemoveAllElements = function()
	{
		this.Sdt.Content.ClearContent(true);

		return true;
	};

	/**
	 * Remove content control and content. If keepContent is true, the content is not deleted.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {bool} keepContent - Specifies if the content will be deleted or not.
	 * @returns {bool} - returns false if content control haven't parent.
	 */
	ApiBlockLvlSdt.prototype.Delete = function(keepContent)
	{
		if (this.Sdt.Index >= 0)
		{
			if (keepContent)
			{
				this.Sdt.RemoveContentControlWrapper();
			}
			else 
			{
				this.Sdt.PreDelete();
				this.Sdt.Parent.RemoveFromContent(this.Sdt.Index, 1, true);
			}

			return true;
		}

		return false;
	};

	/**
	 * Apply text settings to content of content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The properties that will be set to the content of content control.
	 */
	ApiBlockLvlSdt.prototype.SetTextPr = function(oTextPr)
	{
		var ParaTextPr = new AscCommonWord.ParaTextPr(oTextPr.TextPr);
		this.Sdt.Content.SetApplyToAll(true);
		this.Sdt.Add(ParaTextPr);
		this.Sdt.Content.SetApplyToAll(false);
	};

	/**
	 * Get the collection of drawing objects in the document.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiDrawing[]}  
	 */
	ApiBlockLvlSdt.prototype.GetAllDrawingObjects = function()
	{
		var arrAllDrawing = this.Sdt.GetAllDrawingObjects();
		var arrApiDrawings  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
			arrApiDrawings.push(new ApiDrawing(arrAllDrawing[Index]));

		return arrApiDrawings;
	};

	/**
	 * Get the content control that contains the current content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiBlockLvlSdt | null} - returns null if parent content control doesn't exist.  
	 */
	ApiBlockLvlSdt.prototype.GetParentContentControl = function()
	{
		var documentPos = this.Sdt.GetDocumentPositionFromObject();

		for (var Index = documentPos.length - 1; Index >= 1; Index--)
		{
			if (documentPos[Index].Class)
				if (documentPos[Index].Class instanceof CBlockLevelSdt)
					return new ApiBlockLvlSdt(documentPos[Index].Class);
		}

		return null;
	};

	/**
	 * Get the table that contains the content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiTable | null} - returns null is parent table does'n exist.  
	 */
	ApiBlockLvlSdt.prototype.GetParentTable = function()
	{
		var documentPos = this.Sdt.GetDocumentPositionFromObject();

		for (var Index = documentPos.length - 1; Index >= 1; Index--)
		{
			if (documentPos[Index].Class)
				if (documentPos[Index].Class instanceof CTable)
					return new ApiTable(documentPos[Index].Class);
		}

		return null;
	};

	/**
	 * Get the table cell that contains the content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiTableCell | null} - returns null if parent cell doesn't exist.  
	 */
	ApiBlockLvlSdt.prototype.GetParentTableCell = function()
	{
		var documentPos = this.Sdt.GetDocumentPositionFromObject();

		for (var Index = documentPos.length - 1; Index >= 1; Index--)
		{
			if (documentPos[Index].Class.Parent)
				if (documentPos[Index].Class.Parent instanceof CTableCell)
					return new ApiTableCell(documentPos[Index].Class.Parent);
		}

		return null;
	};

	/**
	 * Push a paragraph or a table or BlockLvl content control to actually add it to the document.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {DocumentElement} oElement - The type of the element which will be pushed to the document.
	 * @return {bool} - returns false if oElement unsupported.
	 */
	ApiBlockLvlSdt.prototype.Push = function(oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || ApiBlockLvlSdt)
		{
			if (this.Sdt.IsShowingPlcHdr())
			{
				this.Sdt.Content.RemoveFromContent(0, this.Sdt.Content.GetElementsCount(), false);
				this.Sdt.SetShowingPlcHdr(false);
			}
			
			this.Sdt.Content.Internal_Content_Add(this.Sdt.Content.Content.length, oElement.private_GetImpl());
			return true;
		}

		return false;
	};

	/**
	 * Add a paragraph or a table or a blocklvl content control to the document.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {DocumentElement} oElement - The type of the element which will be added to the document.
	 * @param {Number} nPos - The specified position.
	 * @return {bool} - returns false if oElement unsupported.
	 */
	ApiBlockLvlSdt.prototype.AddElement = function(oElement, nPos)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || ApiBlockLvlSdt)
		{
			if (this.Sdt.IsShowingPlcHdr())
			{
				this.Sdt.Content.RemoveFromContent(0, this.Sdt.Content.GetElementsCount(), false);
				this.Sdt.SetShowingPlcHdr(false);
			}
			
			this.Sdt.Content.Internal_Content_Add(nPos, oElement.private_GetImpl());
			return true;
		}

		return false;
	};

	/**
	 * Add a text to the current content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {String} sText - The text which will be added to the content control.
	 * @return {bool} - returns false if param is invalid.
	 */
	ApiBlockLvlSdt.prototype.AddText = function(sText)
	{
		if (typeof sText === "string")
		{
			var oParagraph = editor.CreateParagraph();
			oParagraph.AddText(sText);
			this.Sdt.Content.Internal_Content_Add(this.Sdt.Content.Content.length, oParagraph.private_GetImpl());

			return true;
		}

		return false;
	};

	/**
	 * Return a Range object that represents the part of the document contained in the specified content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiBlockLvlSdt.prototype.GetRange = function(Start, End)
	{
		var Range = new ApiRange(this.Sdt, Start, End);

		return Range;
	};

	/**
	 * Search for the scope of a content control object. The search results are a collection of ApiRange objects.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {bool} isMatchCase - Case sensitive or not. 
	 * @return {ApiRange[]}  
	 */
	ApiBlockLvlSdt.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase	= false;

		var arrApiRanges	= [];
		var allParagraphs	= [];
		this.Sdt.GetAllParagraphs({All : true}, allParagraphs);

		for (var para in allParagraphs)
		{
			var oParagraph			= new ApiParagraph(allParagraphs[para]);
			var arrOfParaApiRanges	= oParagraph.Search(sText, isMatchCase);

			for (var itemRange = 0; itemRange < arrOfParaApiRanges.length; itemRange++)	
				arrApiRanges.push(arrOfParaApiRanges[itemRange]);
		}

		return arrApiRanges;
	};

	/**
	 * Select a content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 */
	ApiBlockLvlSdt.prototype.Select = function()
	{
		var Document = private_GetLogicDocument();

		this.Sdt.SelectContentControl();
		Document.UpdateSelection();
	};

	/**
	 * Replaces each paragraph(or text in cell) in the select with the corresponding text from an array of strings.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {Array} arrString - represents an array of strings. 
	 */
	Api.prototype.ReplaceTextSmart = function(arrString)
	{
		var allRunsInfo      = null;
		var textDelta        = null;
		var arrSelectedParas = null;

		function GetRunInfo(oRun)
		{
			var StartPos = 0;
			var EndPos   = 0;

			if (oRun.IsSelectionUse() && oRun.State.Selection.StartPos !== oRun.State.Selection.EndPos)
			{
				var runInfo = {
					Run : oRun,
					StartPos : null,
					GlobStartPos : null,
					GlobEndPos : null,
					StringCount : 0,
					String : ""
				};

				if ( true === oRun.Selection.Use )
				{
					StartPos = oRun.State.Selection.StartPos;
					EndPos   = oRun.State.Selection.EndPos;

					if ( StartPos > EndPos )
					{
						var Temp = EndPos;
						EndPos   = StartPos;
						StartPos = Temp;
					}

					runInfo.StartPos     = StartPos;
					runInfo.GlobStartPos = StartPos;
					runInfo.GlobEndPos   = EndPos;
				}

				var posToSplit = [StartPos];

				for ( var Pos = StartPos; Pos < EndPos; Pos++ )
				{
					var Item = oRun.Content[Pos];
					var ItemType = Item.Type;

					switch ( ItemType )
					{
						case para_Numbering:
						case para_PresentationNumbering:
						case para_PageNum:
						case para_PageCount:
						case para_End:
						case para_Drawing:
						case para_NewLine:
						{
							if (posToSplit.indexOf(Pos) === -1)
								posToSplit.push(Pos);
							
							break;
						}
					}
				}

				var noTextCount = 0;
				for (var Index = 0; Index < posToSplit.length; Index++)
				{
					var oInfo = {
						Run : oRun,
						StartPos : null,
						GlobStartPos : null,
						GlobEndPos : null,
						StringCount : 0,
						String : ""
					};

					var nEndPos = EndPos;
					if (posToSplit[Index + 1])
						nEndPos = posToSplit[Index + 1]

					for (var nPos = posToSplit[Index]; nPos < nEndPos; nPos++)
					{
						var Item = oRun.Content[nPos];
						var ItemType = Item.Type;

						switch ( ItemType )
						{
							case para_Numbering:
							case para_PresentationNumbering:
							case para_PageNum:
							case para_PageCount:
							case para_End:
							case para_Drawing:
							case para_NewLine:
							{
								noTextCount++;
								break;
							}
							case para_Text :
							{
								oInfo.String += AscCommon.encodeSurrogateChar(Item.Value);
								oInfo.StringCount++;				
								break;
							}
							case para_Space:
							case para_Tab  : 
							{
								oInfo.String += " ";
								oInfo.StringCount++; 
								break;
							}
						}
					}
					
					if (oInfo.String === "")
						break;
					
					oInfo.StartPos = posToSplit[Index] + noTextCount;

					if (allRunsInfo[allRunsInfo.length - 1])
					{
						oInfo.GlobStartPos = allRunsInfo[allRunsInfo.length - 1].GlobStartPos + allRunsInfo[allRunsInfo.length - 1].StringCount;
						oInfo.GlobEndPos = oInfo.GlobStartPos + Math.max(0, oInfo.StringCount - 1);
					}
					else 
					{
						oInfo.GlobStartPos = 0
						oInfo.GlobEndPos = oInfo.GlobStartPos + Math.max(0, oInfo.StringCount - 1);
					}

					allRunsInfo.push(oInfo);
				}
			}
		}

		function DelInsertChars()
		{
			for (var nChange = textDelta.length - 1; nChange >= 0; nChange--)
			{
				var oChange = textDelta[nChange];
				var DelCount = oChange.deleteCount;
				var infoToAdd = null;
				for (var nInfo = 0; nInfo < allRunsInfo.length; nInfo++)
				{
					var oInfo = allRunsInfo[nInfo];
					if (oChange.pos >= oInfo.GlobStartPos || oChange.pos + DelCount > oInfo.GlobStartPos)
					{
						var nPosToDel   = Math.max(0, oChange.pos - oInfo.GlobStartPos + oInfo.StartPos);
						var nPosToAdd   = nPosToDel
						var nCharsToDel = Math.min(oChange.deleteCount, oInfo.StringCount);
						
						if ((nPosToDel >= oInfo.Run.Content.length && nCharsToDel !== 0) || (nCharsToDel === 0 && oChange.deleteCount !== 0)
							|| nPosToAdd > oInfo.Run.Content.length)
							continue;

						for (var nChar = 0; nChar < nCharsToDel; nChar++)
						{
							if (!oInfo.Run.Content[nPosToDel])
								break;
								
							if (para_Text === oInfo.Run.Content[nPosToDel].Type || para_Space === oInfo.Run.Content[nPosToDel].Type || para_Tab === oInfo.Run.Content[nPosToDel].Type)
							{
								oInfo.Run.RemoveFromContent(nPosToDel, 1);
								nChar--;
								oChange.deleteCount--;
								nCharsToDel--;
							}
							else
							{
								nPosToDel++;
								nChar--;
							}
						}
						
						if (oChange.insert.length === 0)
							continue;

						if (oChange.deleteCount !== 0)
						{
							infoToAdd = 
							{
								Run: oInfo.Run,
								Pos: nPosToAdd
							};
							continue;
						}
						
						for (var nChar = 0; nChar < oChange.insert.length; nChar++)
						{
							var itemText = null;
							
							if (AscCommon.IsSpace(oChange.insert[nChar]))
								itemText = new AscCommonWord.ParaSpace(oChange.insert[nChar]);
							else
								itemText = new AscCommonWord.ParaText(oChange.insert[nChar]);

							itemText.Parent = oInfo.Run.GetParagraph();
							if (oInfo.Run.Content.length === 0 && infoToAdd)
							{
								infoToAdd.Run.AddToContent(infoToAdd.Pos, itemText);
								infoToAdd.Pos++;
							}
							else
								oInfo.Run.AddToContent(nPosToAdd, itemText);

							oChange.insert.shift();
							nChar--;
							nPosToAdd++;
						}
					}
				}
			}
		};

		function ReplaceInParas(arrBasicParas) 
		{
			allRunsInfo = [];

			for (var Index = 0; Index < arrBasicParas.length; Index++)
			{
				var oPara = arrBasicParas[Index];
				var oParaText = "";
				
				if (oPara.Selection.Use)
					oPara.CheckRunContent(GetRunInfo);
					
				for (var nRun = 0; nRun < allRunsInfo.length; nRun++)
					oParaText += allRunsInfo[nRun].String;

				if (oParaText == "")
				{
					allRunsInfo = [];
					continue;
				}
					
				textDelta = AscCommon.getTextDelta(oParaText, arrString[Index]);

				DelInsertChars();
				allRunsInfo = [];
			}
		}

		if (this.editorId === AscCommon.c_oEditorId.Spreadsheet) 
		{
			var oWorksheet        = this.GetActiveSheet();
			var oRange            = oWorksheet.GetSelection();
			var tempRange         = null;
			var nCountLinesInCell = null;
			var resultText        = null;
			var nTextToReplace    = 0;
			var ws                = this.wb.getWorksheet();
			var oContent = ws.objectRender.controller.getTargetDocContent();

			if (oContent) 
			{
				arrSelectedParas = [];
				oContent.GetCurrentParagraph(false, arrSelectedParas, {});
				ReplaceInParas(arrSelectedParas);
				Asc.editor.wb.recalculateDrawingObjects();
				return;
			}

			for (var nRow = oRange.range.bbox.r1; nRow <= oRange.range.bbox.r2; nRow++)
			{
				for (var nCol = oRange.range.bbox.c1; nCol <= oRange.range.bbox.c2; nCol++)
				{
					if (oWorksheet.worksheet.getRowHidden(nRow))
						continue;

					resultText        = '';
					tempRange         = oWorksheet.GetRangeByNumber(nRow, nCol);
					nCountLinesInCell = tempRange.GetValue().split('\n').length;

					for (var nText = nTextToReplace; nText < nTextToReplace + nCountLinesInCell; nText++) 
					{
						if (!arrString[nText])
							continue;
							
						resultText += arrString[nText];
						if (nText !== nTextToReplace + nCountLinesInCell - 1)
							resultText += '\n';

					}
					nTextToReplace += nCountLinesInCell;

					if (resultText !== '')
						if (!this.wb.getCellEditMode())
							tempRange.SetValue(resultText);
						else
							this.wb.cellEditor.pasteText(resultText);
				}
			}
		}
		else 
		{
			var oDocument = this.GetDocument();
			arrSelectedParas = oDocument.Document.GetSelectedParagraphs();
			
			ReplaceInParas(arrSelectedParas);
			
			if (arrSelectedParas[0] && arrSelectedParas[0].Parent)
				arrSelectedParas[0].Parent.RemoveSelection();
			else 
				oDocument.Document.RemoveSelection();
		}
	};
	Api.prototype.CoAuthoringChatSendMessage = function(sString)
	{
		if (typeof sString !== 'string' || sString === '')
			return false;

		this.asc_coAuthoringChatSendMessage(sString);
		return true;
	};
	/**
	 * Convert document to markdown or html text.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {"markdown" | "html"} [sConvertType="markdown"] - type of converting.
	 * @param {bool} [bHtmlHeadings=false] - If you have used multiple Heading 1 headings in your Doc, set this param true to demote all heading levels to conform with the following standard: single H1 as title, H2 as top-level heading in the text body.
	 * @param {bool} [bBase64img=false] - set this param true if you want images to be created in base64 format.
	 * @param {bool} [bDemoteHeadings=false] - Not all Markdown renderers handle Markdown-style IDs. If that is the case for your target platform, set this param true to generate HTML headings and IDs.
	 * @param {bool} [bRenderHTMLTags=false] - By default, angle brackets (<) will be replaced by the &lt; entity. If you really want to embed HTML tags in your Markdown, set this param true to preserve them. 
	 * Or, if you just want to use an occasional HTML tag, you can escape the opening angle bracket like this: \<tag>text\</tag>.
	 * @returns {string}
	 */
	Api.prototype.ConvertDocument = function(sConvertType, bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags) 
	{
		var oDocument = this.GetDocument();
		if (sConvertType === "html")
			return oDocument.ToHtml(bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags);
		else
			return oDocument.ToMarkdown(bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags);
	};

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Export
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	Api.prototype["GetDocument"]                     = Api.prototype.GetDocument;
	Api.prototype["CreateParagraph"]                 = Api.prototype.CreateParagraph;
	Api.prototype["CreateTable"]                     = Api.prototype.CreateTable;
	Api.prototype["AddComment"]                      = Api.prototype.AddComment;
	Api.prototype["CreateRun"]                       = Api.prototype.CreateRun;
	Api.prototype["CreateHyperlink"]                 = Api.prototype.CreateHyperlink;
	Api.prototype["CreateImage"]                     = Api.prototype.CreateImage;
	Api.prototype["CreateShape"]                     = Api.prototype.CreateShape;
	Api.prototype["CreateChart"]                     = Api.prototype.CreateChart;
	Api.prototype["CreateRGBColor"]                  = Api.prototype.CreateRGBColor;
	Api.prototype["CreateSchemeColor"]               = Api.prototype.CreateSchemeColor;
	Api.prototype["CreatePresetColor"]               = Api.prototype.CreatePresetColor;
	Api.prototype["CreateSolidFill"]                 = Api.prototype.CreateSolidFill;
	Api.prototype["CreateLinearGradientFill"]        = Api.prototype.CreateLinearGradientFill;
	Api.prototype["CreateRadialGradientFill"]        = Api.prototype.CreateRadialGradientFill;
	Api.prototype["CreatePatternFill"]               = Api.prototype.CreatePatternFill;
	Api.prototype["CreateBlipFill"]                  = Api.prototype.CreateBlipFill;
	Api.prototype["CreateNoFill"]                    = Api.prototype.CreateNoFill;
	Api.prototype["CreateStroke"]                    = Api.prototype.CreateStroke;
	Api.prototype["CreateGradientStop"]              = Api.prototype.CreateGradientStop;
	Api.prototype["CreateBullet"]                    = Api.prototype.CreateBullet;
	Api.prototype["CreateNumbering"]                 = Api.prototype.CreateNumbering;
	Api.prototype["CreateInlineLvlSdt"]              = Api.prototype.CreateInlineLvlSdt;
	Api.prototype["CreateBlockLvlSdt"]               = Api.prototype.CreateBlockLvlSdt;
	Api.prototype["Save"]               			 = Api.prototype.Save;
	Api.prototype["LoadMailMergeData"]               = Api.prototype.LoadMailMergeData;
	Api.prototype["GetMailMergeTemplateDocContent"]  = Api.prototype.GetMailMergeTemplateDocContent;
	Api.prototype["GetMailMergeReceptionsCount"]     = Api.prototype.GetMailMergeReceptionsCount;
	Api.prototype["ReplaceDocumentContent"]          = Api.prototype.ReplaceDocumentContent;
	Api.prototype["MailMerge"]                       = Api.prototype.MailMerge;
	Api.prototype["ReplaceTextSmart"]				 = Api.prototype.ReplaceTextSmart;
	Api.prototype["CoAuthoringChatSendMessage"]		 = Api.prototype.CoAuthoringChatSendMessage;
	Api.prototype["ConvertDocument"]		         = Api.prototype.ConvertDocument;
	
	ApiUnsupported.prototype["GetClassType"]         = ApiUnsupported.prototype.GetClassType;

	ApiDocumentContent.prototype["GetClassType"]     = ApiDocumentContent.prototype.GetClassType;
	ApiDocumentContent.prototype["GetElementsCount"] = ApiDocumentContent.prototype.GetElementsCount;
	ApiDocumentContent.prototype["GetElement"]       = ApiDocumentContent.prototype.GetElement;
	ApiDocumentContent.prototype["AddElement"]       = ApiDocumentContent.prototype.AddElement;
	ApiDocumentContent.prototype["Push"]             = ApiDocumentContent.prototype.Push;
	ApiDocumentContent.prototype["RemoveAllElements"]= ApiDocumentContent.prototype.RemoveAllElements;
	ApiDocumentContent.prototype["RemoveElement"]    = ApiDocumentContent.prototype.RemoveElement;
	ApiDocumentContent.prototype["GetRange"]         = ApiDocumentContent.prototype.GetRange;

	ApiRange.prototype["GetClassType"]               = ApiRange.prototype.GetClassType
	ApiRange.prototype["GetParagraph"]               = ApiRange.prototype.GetParagraph;
	ApiRange.prototype["AddText"]                    = ApiRange.prototype.AddText;
	ApiRange.prototype["AddBookmark"]                = ApiRange.prototype.AddBookmark;
	ApiRange.prototype["AddHyperlink"]               = ApiRange.prototype.AddHyperlink;
	ApiRange.prototype["GetText"]                    = ApiRange.prototype.GetText;
	ApiRange.prototype["GetAllParagraphs"]           = ApiRange.prototype.GetAllParagraphs;
	ApiRange.prototype["Select"]                     = ApiRange.prototype.Select;
	ApiRange.prototype["ExpandTo"]                   = ApiRange.prototype.ExpandTo;
	ApiRange.prototype["IntersectWith"]              = ApiRange.prototype.IntersectWith;
	ApiRange.prototype["SetBold"]                    = ApiRange.prototype.SetBold;
	ApiRange.prototype["SetCaps"]                    = ApiRange.prototype.SetCaps;
	ApiRange.prototype["SetColor"]                   = ApiRange.prototype.SetColor;
	ApiRange.prototype["SetDoubleStrikeout"]         = ApiRange.prototype.SetDoubleStrikeout;
	ApiRange.prototype["SetHighlight"]               = ApiRange.prototype.SetHighlight;
	ApiRange.prototype["SetShd"]                     = ApiRange.prototype.SetShd;
	ApiRange.prototype["SetItalic"]                  = ApiRange.prototype.SetItalic;
	ApiRange.prototype["SetStrikeout"]               = ApiRange.prototype.SetStrikeout;
	ApiRange.prototype["SetSmallCaps"]               = ApiRange.prototype.SetSmallCaps;
	ApiRange.prototype["SetSpacing"]                 = ApiRange.prototype.SetSpacing;
	ApiRange.prototype["SetUnderline"]               = ApiRange.prototype.SetUnderline;
	ApiRange.prototype["SetVertAlign"]               = ApiRange.prototype.SetVertAlign;
	ApiRange.prototype["SetPosition"]                = ApiRange.prototype.SetPosition;
	ApiRange.prototype["SetFontSize"]                = ApiRange.prototype.SetFontSize;
	ApiRange.prototype["SetFontFamily"]              = ApiRange.prototype.SetFontFamily;
	ApiRange.prototype["SetStyle"]                   = ApiRange.prototype.SetStyle;
	ApiRange.prototype["SetTextPr"]                  = ApiRange.prototype.SetTextPr;
	ApiRange.prototype["Delete"]                     = ApiRange.prototype.Delete;

	ApiDocument.prototype["GetClassType"]            = ApiDocument.prototype.GetClassType;
	ApiDocument.prototype["CreateNewHistoryPoint"]   = ApiDocument.prototype.CreateNewHistoryPoint;
	ApiDocument.prototype["GetDefaultTextPr"]        = ApiDocument.prototype.GetDefaultTextPr;
	ApiDocument.prototype["GetDefaultParaPr"]        = ApiDocument.prototype.GetDefaultParaPr;
	ApiDocument.prototype["GetStyle"]                = ApiDocument.prototype.GetStyle;
	ApiDocument.prototype["CreateStyle"]             = ApiDocument.prototype.CreateStyle;
	ApiDocument.prototype["GetDefaultStyle"]         = ApiDocument.prototype.GetDefaultStyle;
	ApiDocument.prototype["GetFinalSection"]         = ApiDocument.prototype.GetFinalSection;
	ApiDocument.prototype["CreateSection"]           = ApiDocument.prototype.CreateSection;
	ApiDocument.prototype["SetEvenAndOddHdrFtr"]     = ApiDocument.prototype.SetEvenAndOddHdrFtr;
	ApiDocument.prototype["CreateNumbering"]         = ApiDocument.prototype.CreateNumbering;
	ApiDocument.prototype["InsertContent"]           = ApiDocument.prototype.InsertContent;
	ApiDocument.prototype["GetCommentsReport"]       = ApiDocument.prototype.GetCommentsReport;
	ApiDocument.prototype["GetReviewReport"]         = ApiDocument.prototype.GetReviewReport;
	ApiDocument.prototype["InsertWatermark"]         = ApiDocument.prototype.InsertWatermark;
	ApiDocument.prototype["SearchAndReplace"]        = ApiDocument.prototype.SearchAndReplace;
	ApiDocument.prototype["GetAllContentControls"]   = ApiDocument.prototype.GetAllContentControls;
	ApiDocument.prototype["SetTrackRevisions"]       = ApiDocument.prototype.SetTrackRevisions;
	ApiDocument.prototype["IsTrackRevisions"]        = ApiDocument.prototype.IsTrackRevisions;
	ApiDocument.prototype["GetRange"]                = ApiDocument.prototype.GetRange;
	ApiDocument.prototype["GetRangeBySelect"]        = ApiDocument.prototype.GetRangeBySelect;
	ApiDocument.prototype["Last"]                    = ApiDocument.prototype.Last;
	ApiDocument.prototype["Push"]                    = ApiDocument.prototype.Push;
	ApiDocument.prototype["DeleteBookmark"]          = ApiDocument.prototype.DeleteBookmark;
	ApiDocument.prototype["AddComment"]              = ApiDocument.prototype.AddComment;
	ApiDocument.prototype["GetBookmarkRange"]        = ApiDocument.prototype.GetBookmarkRange;
	ApiDocument.prototype["GetSections"]             = ApiDocument.prototype.GetSections;
	ApiDocument.prototype["GetAllTablesOnPage"]      = ApiDocument.prototype.GetAllTablesOnPage;
	ApiDocument.prototype["RemoveSelection"]         = ApiDocument.prototype.RemoveSelection;
	ApiDocument.prototype["GetAllDrawingObjects"]    = ApiDocument.prototype.GetAllDrawingObjects;
	ApiDocument.prototype["GetAllShapes"]            = ApiDocument.prototype.GetAllShapes;
	ApiDocument.prototype["GetAllImages"]            = ApiDocument.prototype.GetAllImages;
	ApiDocument.prototype["GetAllCharts"]            = ApiDocument.prototype.GetAllCharts;
	ApiDocument.prototype["Search"]                  = ApiDocument.prototype.Search;
	ApiDocument.prototype["ToMarkdown"]              = ApiDocument.prototype.ToMarkdown;
	ApiDocument.prototype["ToHtml"]                  = ApiDocument.prototype.ToHtml;

	ApiParagraph.prototype["GetClassType"]           = ApiParagraph.prototype.GetClassType;
	ApiParagraph.prototype["AddText"]                = ApiParagraph.prototype.AddText;
	ApiParagraph.prototype["AddPageBreak"]           = ApiParagraph.prototype.AddPageBreak;
	ApiParagraph.prototype["AddLineBreak"]           = ApiParagraph.prototype.AddLineBreak;
	ApiParagraph.prototype["AddColumnBreak"]         = ApiParagraph.prototype.AddColumnBreak;
	ApiParagraph.prototype["AddPageNumber"]          = ApiParagraph.prototype.AddPageNumber;
	ApiParagraph.prototype["AddPagesCount"]          = ApiParagraph.prototype.AddPagesCount;
	ApiParagraph.prototype["GetParagraphMarkTextPr"] = ApiParagraph.prototype.GetParagraphMarkTextPr;
	ApiParagraph.prototype["GetParaPr"]              = ApiParagraph.prototype.GetParaPr;
	ApiParagraph.prototype["GetNumbering"]           = ApiParagraph.prototype.GetNumbering;
	ApiParagraph.prototype["SetNumbering"]           = ApiParagraph.prototype.SetNumbering;
	ApiParagraph.prototype["GetElementsCount"]       = ApiParagraph.prototype.GetElementsCount;
	ApiParagraph.prototype["GetElement"]             = ApiParagraph.prototype.GetElement;
	ApiParagraph.prototype["RemoveElement"]          = ApiParagraph.prototype.RemoveElement;
	ApiParagraph.prototype["Delete"]                 = ApiParagraph.prototype.Delete;
	ApiParagraph.prototype["GetNext"]                = ApiParagraph.prototype.GetNext;
	ApiParagraph.prototype["GetPrevious"]            = ApiParagraph.prototype.GetPrevious;
	ApiParagraph.prototype["RemoveAllElements"]      = ApiParagraph.prototype.RemoveAllElements;
	ApiParagraph.prototype["AddElement"]             = ApiParagraph.prototype.AddElement;
	ApiParagraph.prototype["AddTabStop"]             = ApiParagraph.prototype.AddTabStop;
	ApiParagraph.prototype["AddDrawing"]             = ApiParagraph.prototype.AddDrawing;
	ApiParagraph.prototype["AddInlineLvlSdt"]        = ApiParagraph.prototype.AddInlineLvlSdt;
	ApiParagraph.prototype["Copy"]                   = ApiParagraph.prototype.Copy;
	ApiParagraph.prototype["AddComment"]             = ApiParagraph.prototype.AddComment;
	ApiParagraph.prototype["AddHyperlink"]           = ApiParagraph.prototype.AddHyperlink;
	ApiParagraph.prototype["GetRange"]               = ApiParagraph.prototype.GetRange;
	ApiParagraph.prototype["Push"]                   = ApiParagraph.prototype.Push;
	ApiParagraph.prototype["GetLastRunWithText"]     = ApiParagraph.prototype.GetLastRunWithText;
	ApiParagraph.prototype["SetBold"]                = ApiParagraph.prototype.SetBold;
	ApiParagraph.prototype["SetCaps"]                = ApiParagraph.prototype.SetCaps;
	ApiParagraph.prototype["SetColor"]               = ApiParagraph.prototype.SetColor;
	ApiParagraph.prototype["SetDoubleStrikeout"]     = ApiParagraph.prototype.SetDoubleStrikeout;
	ApiParagraph.prototype["SetFontFamily"]          = ApiParagraph.prototype.SetFontFamily;
	ApiParagraph.prototype["SetFontSize"]            = ApiParagraph.prototype.SetFontSize;
	ApiParagraph.prototype["SetHighlight"]           = ApiParagraph.prototype.SetHighlight;
	ApiParagraph.prototype["SetItalic"]              = ApiParagraph.prototype.SetItalic;
	ApiParagraph.prototype["SetPosition"]            = ApiParagraph.prototype.SetPosition;
	ApiParagraph.prototype["SetShd"]                 = ApiParagraph.prototype.SetShd;
	ApiParagraph.prototype["SetSmallCaps"]           = ApiParagraph.prototype.SetSmallCaps;
	ApiParagraph.prototype["SetSpacing"]             = ApiParagraph.prototype.SetSpacing;
	ApiParagraph.prototype["SetStrikeout"]           = ApiParagraph.prototype.SetStrikeout;
	ApiParagraph.prototype["SetUnderline"]           = ApiParagraph.prototype.SetUnderline;
	ApiParagraph.prototype["SetVertAlign"]           = ApiParagraph.prototype.SetVertAlign;
	ApiParagraph.prototype["Last"]                   = ApiParagraph.prototype.Last;
	ApiParagraph.prototype["GetAllContentControls"]  = ApiParagraph.prototype.GetAllContentControls;
	ApiParagraph.prototype["GetAllDrawingObjects"]   = ApiParagraph.prototype.GetAllDrawingObjects;
	ApiParagraph.prototype["GetAllShapes"]           = ApiParagraph.prototype.GetAllShapes;
	ApiParagraph.prototype["GetAllImages"]           = ApiParagraph.prototype.GetAllImages;
	ApiParagraph.prototype["GetAllCharts"]           = ApiParagraph.prototype.GetAllCharts;
	ApiParagraph.prototype["GetParentContentControl"]= ApiParagraph.prototype.GetParentContentControl;
	ApiParagraph.prototype["GetParentTable"]         = ApiParagraph.prototype.GetParentTable;
	ApiParagraph.prototype["GetParentTableCell"]     = ApiParagraph.prototype.GetParentTableCell;
	ApiParagraph.prototype["GetText"]                = ApiParagraph.prototype.GetText;
	ApiParagraph.prototype["GetTextPr"]              = ApiParagraph.prototype.GetTextPr;
	ApiParagraph.prototype["SetTextPr"]              = ApiParagraph.prototype.SetTextPr;
	ApiParagraph.prototype["InsertInContentControl"] = ApiParagraph.prototype.InsertInContentControl;
	ApiParagraph.prototype["InsertParagraph"]        = ApiParagraph.prototype.InsertParagraph;
	ApiParagraph.prototype["Select"]                 = ApiParagraph.prototype.Select;
	ApiParagraph.prototype["Search"]                 = ApiParagraph.prototype.Search;
	ApiParagraph.prototype["WrapInMailMergeField"]   = ApiParagraph.prototype.WrapInMailMergeField;


	ApiRun.prototype["GetClassType"]                 = ApiRun.prototype.GetClassType;
	ApiRun.prototype["GetTextPr"]                    = ApiRun.prototype.GetTextPr;
	ApiRun.prototype["ClearContent"]                 = ApiRun.prototype.ClearContent;
	ApiRun.prototype["AddText"]                      = ApiRun.prototype.AddText;
	ApiRun.prototype["AddPageBreak"]                 = ApiRun.prototype.AddPageBreak;
	ApiRun.prototype["AddLineBreak"]                 = ApiRun.prototype.AddLineBreak;
	ApiRun.prototype["AddColumnBreak"]               = ApiRun.prototype.AddColumnBreak;
	ApiRun.prototype["AddTabStop"]                   = ApiRun.prototype.AddTabStop;
	ApiRun.prototype["AddDrawing"]                   = ApiRun.prototype.AddDrawing;
	ApiRun.prototype["Select"]                       = ApiRun.prototype.Select;
	ApiRun.prototype["AddHyperlink"]                 = ApiRun.prototype.AddHyperlink;
	ApiRun.prototype["Copy"]                         = ApiRun.prototype.Copy;
	ApiRun.prototype["RemoveAllElements"]            = ApiRun.prototype.RemoveAllElements;
	ApiRun.prototype["Delete"]                       = ApiRun.prototype.Delete;
	ApiRun.prototype["GetRange"]                     = ApiRun.prototype.GetRange;
	ApiRun.prototype["GetParentContentControl"]      = ApiRun.prototype.GetParentContentControl;
	ApiRun.prototype["GetParentTable"]               = ApiRun.prototype.GetParentTable;
	ApiRun.prototype["GetParentTableCell"]           = ApiRun.prototype.GetParentTableCell;
	ApiRun.prototype["SetTextPr"]                    = ApiRun.prototype.SetTextPr;
	ApiRun.prototype["SetBold"]                      = ApiRun.prototype.SetBold;
	ApiRun.prototype["SetCaps"]                      = ApiRun.prototype.SetCaps;
	ApiRun.prototype["SetColor"]                     = ApiRun.prototype.SetColor;
	ApiRun.prototype["SetDoubleStrikeout"]           = ApiRun.prototype.SetDoubleStrikeout;
	ApiRun.prototype["SetFill"]                      = ApiRun.prototype.SetFill;
	ApiRun.prototype["SetFontFamily"]                = ApiRun.prototype.SetFontFamily;
	ApiRun.prototype["SetFontSize"]                  = ApiRun.prototype.SetFontSize;
	ApiRun.prototype["SetHighlight"]                 = ApiRun.prototype.SetHighlight;
	ApiRun.prototype["SetItalic"]                    = ApiRun.prototype.SetItalic;
	ApiRun.prototype["SetLanguage"]                  = ApiRun.prototype.SetLanguage;
	ApiRun.prototype["SetPosition"]                  = ApiRun.prototype.SetPosition;
	ApiRun.prototype["SetShd"]                       = ApiRun.prototype.SetShd;
	ApiRun.prototype["SetSmallCaps"]                 = ApiRun.prototype.SetSmallCaps;
	ApiRun.prototype["SetSpacing"]                   = ApiRun.prototype.SetSpacing;
	ApiRun.prototype["SetStrikeout"]                 = ApiRun.prototype.SetStrikeout;
	ApiRun.prototype["SetUnderline"]                 = ApiRun.prototype.SetUnderline;
	ApiRun.prototype["SetVertAlign"]                 = ApiRun.prototype.SetVertAlign;
	ApiRun.prototype["WrapInMailMergeField"]         = ApiRun.prototype.WrapInMailMergeField;

	ApiHyperlink.prototype["GetClassType"]           = ApiHyperlink.prototype.GetClassType;
	ApiHyperlink.prototype["SetLink"]                = ApiHyperlink.prototype.SetLink;
	ApiHyperlink.prototype["SetDisplayedText"]       = ApiHyperlink.prototype.SetDisplayedText;
	ApiHyperlink.prototype["SetScreenTipText"]       = ApiHyperlink.prototype.SetScreenTipText;
	ApiHyperlink.prototype["GetLinkedText"]          = ApiHyperlink.prototype.GetLinkedText;
	ApiHyperlink.prototype["GetDisplayedText"]       = ApiHyperlink.prototype.GetDisplayedText;
	ApiHyperlink.prototype["GetScreenTipText"]       = ApiHyperlink.prototype.GetScreenTipText;
	ApiHyperlink.prototype["GetElement"]             = ApiHyperlink.prototype.GetElement;
	ApiHyperlink.prototype["GetElementsCount"]       = ApiHyperlink.prototype.GetElementsCount;
	ApiHyperlink.prototype["SetDefaultStyle"]        = ApiHyperlink.prototype.SetDefaultStyle;
	ApiHyperlink.prototype["GetRange"]               = ApiHyperlink.prototype.GetRange;

	ApiSection.prototype["GetClassType"]             = ApiSection.prototype.GetClassType;
	ApiSection.prototype["SetType"]                  = ApiSection.prototype.SetType;
	ApiSection.prototype["SetEqualColumns"]          = ApiSection.prototype.SetEqualColumns;
	ApiSection.prototype["SetNotEqualColumns"]       = ApiSection.prototype.SetNotEqualColumns;
	ApiSection.prototype["SetPageSize"]              = ApiSection.prototype.SetPageSize;
	ApiSection.prototype["SetPageMargins"]           = ApiSection.prototype.SetPageMargins;
	ApiSection.prototype["SetHeaderDistance"]        = ApiSection.prototype.SetHeaderDistance;
	ApiSection.prototype["SetFooterDistance"]        = ApiSection.prototype.SetFooterDistance;
	ApiSection.prototype["GetHeader"]                = ApiSection.prototype.GetHeader;
	ApiSection.prototype["RemoveHeader"]             = ApiSection.prototype.RemoveHeader;
	ApiSection.prototype["GetFooter"]                = ApiSection.prototype.GetFooter;
	ApiSection.prototype["RemoveFooter"]             = ApiSection.prototype.RemoveFooter;
	ApiSection.prototype["SetTitlePage"]             = ApiSection.prototype.SetTitlePage;
	ApiSection.prototype["GetNext"]                  = ApiSection.prototype.GetNext;
	ApiSection.prototype["GetPrevious"]              = ApiSection.prototype.GetPrevious;

	ApiTable.prototype["GetClassType"]               = ApiTable.prototype.GetClassType;
	ApiTable.prototype["SetJc"]                      = ApiTable.prototype.SetJc;
	ApiTable.prototype["GetRowsCount"]               = ApiTable.prototype.GetRowsCount;
	ApiTable.prototype["GetRow"]                     = ApiTable.prototype.GetRow;
	ApiTable.prototype["MergeCells"]                 = ApiTable.prototype.MergeCells;
	ApiTable.prototype["SetStyle"]                   = ApiTable.prototype.SetStyle;
	ApiTable.prototype["SetTableLook"]               = ApiTable.prototype.SetTableLook;
	ApiTable.prototype["AddRow"]                     = ApiTable.prototype.AddRow;
	ApiTable.prototype["AddRows"]                    = ApiTable.prototype.AddRows;
	ApiTable.prototype["AddColumn"]                  = ApiTable.prototype.AddColumn;
	ApiTable.prototype["AddColumns"]                 = ApiTable.prototype.AddColumns;
	ApiTable.prototype["AddElement"]                 = ApiTable.prototype.AddElement;
	ApiTable.prototype["RemoveRow"]                  = ApiTable.prototype.RemoveRow;
	ApiTable.prototype["RemoveColumn"]               = ApiTable.prototype.RemoveColumn;
	ApiTable.prototype["Copy"]                       = ApiTable.prototype.Copy;
	ApiTable.prototype["GetCell"]    				 = ApiTable.prototype.GetCell;
	ApiTable.prototype["Split"]    					 = ApiTable.prototype.Split;
	ApiTable.prototype["AddRows"]    				 = ApiTable.prototype.AddRows;
	ApiTable.prototype["AddColumns"]   				 = ApiTable.prototype.AddColumns;
	ApiTable.prototype["Select"]    			     = ApiTable.prototype.Select;
	ApiTable.prototype["GetRange"]    				 = ApiTable.prototype.GetRange;
	ApiTable.prototype["SetHAlign"]    				 = ApiTable.prototype.SetHAlign;
	ApiTable.prototype["SetVAlign"]    				 = ApiTable.prototype.SetVAlign;
	ApiTable.prototype["SetPaddings"]    			 = ApiTable.prototype.SetPaddings;
	ApiTable.prototype["SetWrappingStyle"]    		 = ApiTable.prototype.SetWrappingStyle;
	ApiTable.prototype["GetParentContentControl"]    = ApiTable.prototype.GetParentContentControl;
	ApiTable.prototype["InsertInContentControl"]     = ApiTable.prototype.InsertInContentControl;
	ApiTable.prototype["GetParentTable"]    		 = ApiTable.prototype.GetParentTable;
	ApiTable.prototype["GetTables"]     			 = ApiTable.prototype.GetTables;
	ApiTable.prototype["GetNext"]    				 = ApiTable.prototype.GetNext;
	ApiTable.prototype["GetPrevious"]    			 = ApiTable.prototype.GetPrevious;
	ApiTable.prototype["GetParentTableCell"]   	 	 = ApiTable.prototype.GetParentTableCell;
	ApiTable.prototype["Delete"]    				 = ApiTable.prototype.Delete;
	ApiTable.prototype["Clear"]    					 = ApiTable.prototype.Clear;
	ApiTable.prototype["Search"]    				 = ApiTable.prototype.Search;
	ApiTable.prototype["SetTextPr"]    				 = ApiTable.prototype.SetTextPr;

	ApiTableRow.prototype["GetClassType"]            = ApiTableRow.prototype.GetClassType;
	ApiTableRow.prototype["GetCellsCount"]           = ApiTableRow.prototype.GetCellsCount;
	ApiTableRow.prototype["GetCell"]                 = ApiTableRow.prototype.GetCell;
	ApiTableRow.prototype["GetIndex"]           	 = ApiTableRow.prototype.GetIndex;
	ApiTableRow.prototype["GetParentTable"]          = ApiTableRow.prototype.GetParentTable;
	ApiTableRow.prototype["GetNext"]           		 = ApiTableRow.prototype.GetNext;
	ApiTableRow.prototype["GetPrevious"]             = ApiTableRow.prototype.GetPrevious;
	ApiTableRow.prototype["AddRows"]           		 = ApiTableRow.prototype.AddRows;
	ApiTableRow.prototype["MergeCells"]          	 = ApiTableRow.prototype.MergeCells;
	ApiTableRow.prototype["Clear"]           		 = ApiTableRow.prototype.Clear;
	ApiTableRow.prototype["Remove"]           		 = ApiTableRow.prototype.Remove;
	ApiTableRow.prototype["SetTextPr"]          	 = ApiTableRow.prototype.SetTextPr;
	ApiTableRow.prototype["Search"]          		 = ApiTableRow.prototype.Search;

	ApiTableCell.prototype["GetClassType"]           = ApiTableCell.prototype.GetClassType;
	ApiTableCell.prototype["GetContent"]             = ApiTableCell.prototype.GetContent;
	ApiTableCell.prototype["GetIndex"]    			 = ApiTableCell.prototype.GetIndex;
	ApiTableCell.prototype["GetRowIndex"]    		 = ApiTableCell.prototype.GetRowIndex;
	ApiTableCell.prototype["GetParentRow"]    		 = ApiTableCell.prototype.GetParentRow;
	ApiTableCell.prototype["GetParentTable"]    	 = ApiTableCell.prototype.GetParentTable;
	ApiTableCell.prototype["AddRows"]    			 = ApiTableCell.prototype.AddRows;
	ApiTableCell.prototype["AddColumns"]    		 = ApiTableCell.prototype.AddColumns;
	ApiTableCell.prototype["RemoveColumn"]    		 = ApiTableCell.prototype.RemoveColumn;
	ApiTableCell.prototype["RemoveRow"]    			 = ApiTableCell.prototype.RemoveRow;
	ApiTableCell.prototype["Search"]    			 = ApiTableCell.prototype.Search;
	ApiTableCell.prototype["GetNext"]    			 = ApiTableCell.prototype.GetNext;
	ApiTableCell.prototype["GetPrevious"]    		 = ApiTableCell.prototype.GetPrevious;
	ApiTableCell.prototype["Split"]    				 = ApiTableCell.prototype.Split;
	ApiTableCell.prototype["SetCellPr"]    			 = ApiTableCell.prototype.SetCellPr;
	ApiTableCell.prototype["SetTextPr"]    			 = ApiTableCell.prototype.SetTextPr;
	ApiTableCell.prototype["Clear"]    		         = ApiTableCell.prototype.Clear;
	ApiTableCell.prototype["AddElement"]    		 = ApiTableCell.prototype.AddElement;

	ApiStyle.prototype["GetClassType"]               = ApiStyle.prototype.GetClassType;
	ApiStyle.prototype["GetName"]                    = ApiStyle.prototype.GetName;
	ApiStyle.prototype["SetName"]                    = ApiStyle.prototype.SetName;
	ApiStyle.prototype["GetType"]                    = ApiStyle.prototype.GetType;
	ApiStyle.prototype["GetTextPr"]                  = ApiStyle.prototype.GetTextPr;
	ApiStyle.prototype["GetParaPr"]                  = ApiStyle.prototype.GetParaPr;
	ApiStyle.prototype["GetTablePr"]                 = ApiStyle.prototype.GetTablePr;
	ApiStyle.prototype["GetTableRowPr"]              = ApiStyle.prototype.GetTableRowPr;
	ApiStyle.prototype["GetTableCellPr"]             = ApiStyle.prototype.GetTableCellPr;
	ApiStyle.prototype["SetBasedOn"]                 = ApiStyle.prototype.SetBasedOn;
	ApiStyle.prototype["GetConditionalTableStyle"]   = ApiStyle.prototype.GetConditionalTableStyle;

	ApiNumbering.prototype["GetClassType"]           = ApiNumbering.prototype.GetClassType;
	ApiNumbering.prototype["GetLevel"]               = ApiNumbering.prototype.GetLevel;

	ApiNumberingLevel.prototype["GetClassType"]      = ApiNumberingLevel.prototype.GetClassType;
	ApiNumberingLevel.prototype["GetNumbering"]      = ApiNumberingLevel.prototype.GetNumbering;
	ApiNumberingLevel.prototype["GetLevelIndex"]     = ApiNumberingLevel.prototype.GetLevelIndex;
	ApiNumberingLevel.prototype["GetTextPr"]         = ApiNumberingLevel.prototype.GetTextPr;
	ApiNumberingLevel.prototype["GetParaPr"]         = ApiNumberingLevel.prototype.GetParaPr;
	ApiNumberingLevel.prototype["SetTemplateType"]   = ApiNumberingLevel.prototype.SetTemplateType;
	ApiNumberingLevel.prototype["SetCustomType"]     = ApiNumberingLevel.prototype.SetCustomType;
	ApiNumberingLevel.prototype["SetRestart"]        = ApiNumberingLevel.prototype.SetRestart;
	ApiNumberingLevel.prototype["SetStart"]          = ApiNumberingLevel.prototype.SetStart;
	ApiNumberingLevel.prototype["SetSuff"]           = ApiNumberingLevel.prototype.SetSuff;

	ApiTextPr.prototype["GetClassType"]              = ApiTextPr.prototype.GetClassType;
	ApiTextPr.prototype["SetStyle"]                  = ApiTextPr.prototype.SetStyle;
	ApiTextPr.prototype["SetBold"]                   = ApiTextPr.prototype.SetBold;
	ApiTextPr.prototype["SetItalic"]                 = ApiTextPr.prototype.SetItalic;
	ApiTextPr.prototype["SetStrikeout"]              = ApiTextPr.prototype.SetStrikeout;
	ApiTextPr.prototype["SetUnderline"]              = ApiTextPr.prototype.SetUnderline;
	ApiTextPr.prototype["SetFontFamily"]             = ApiTextPr.prototype.SetFontFamily;
	ApiTextPr.prototype["SetFontSize"]               = ApiTextPr.prototype.SetFontSize;
	ApiTextPr.prototype["SetColor"]                  = ApiTextPr.prototype.SetColor;
	ApiTextPr.prototype["SetVertAlign"]              = ApiTextPr.prototype.SetVertAlign;
	ApiTextPr.prototype["SetHighlight"]              = ApiTextPr.prototype.SetHighlight;
	ApiTextPr.prototype["SetSpacing"]                = ApiTextPr.prototype.SetSpacing;
	ApiTextPr.prototype["SetDoubleStrikeout"]        = ApiTextPr.prototype.SetDoubleStrikeout;
	ApiTextPr.prototype["SetCaps"]                   = ApiTextPr.prototype.SetCaps;
	ApiTextPr.prototype["SetSmallCaps"]              = ApiTextPr.prototype.SetSmallCaps;
	ApiTextPr.prototype["SetPosition"]               = ApiTextPr.prototype.SetPosition;
	ApiTextPr.prototype["SetLanguage"]               = ApiTextPr.prototype.SetLanguage;
	ApiTextPr.prototype["SetShd"]                    = ApiTextPr.prototype.SetShd;
	ApiTextPr.prototype["SetFill"]                   = ApiTextPr.prototype.SetFill;

	ApiParaPr.prototype["GetClassType"]              = ApiParaPr.prototype.GetClassType;
	ApiParaPr.prototype["SetStyle"]                  = ApiParaPr.prototype.SetStyle;
	ApiParaPr.prototype["SetContextualSpacing"]      = ApiParaPr.prototype.SetContextualSpacing;
	ApiParaPr.prototype["SetIndLeft"]                = ApiParaPr.prototype.SetIndLeft;
	ApiParaPr.prototype["SetIndRight"]               = ApiParaPr.prototype.SetIndRight;
	ApiParaPr.prototype["SetIndFirstLine"]           = ApiParaPr.prototype.SetIndFirstLine;
	ApiParaPr.prototype["SetJc"]                     = ApiParaPr.prototype.SetJc;
	ApiParaPr.prototype["SetKeepLines"]              = ApiParaPr.prototype.SetKeepLines;
	ApiParaPr.prototype["SetKeepNext"]               = ApiParaPr.prototype.SetKeepNext;
	ApiParaPr.prototype["SetPageBreakBefore"]        = ApiParaPr.prototype.SetPageBreakBefore;
	ApiParaPr.prototype["SetSpacingLine"]            = ApiParaPr.prototype.SetSpacingLine;
	ApiParaPr.prototype["SetSpacingBefore"]          = ApiParaPr.prototype.SetSpacingBefore;
	ApiParaPr.prototype["SetSpacingAfter"]           = ApiParaPr.prototype.SetSpacingAfter;
	ApiParaPr.prototype["SetShd"]                    = ApiParaPr.prototype.SetShd;
	ApiParaPr.prototype["SetBottomBorder"]           = ApiParaPr.prototype.SetBottomBorder;
	ApiParaPr.prototype["SetLeftBorder"]             = ApiParaPr.prototype.SetLeftBorder;
	ApiParaPr.prototype["SetRightBorder"]            = ApiParaPr.prototype.SetRightBorder;
	ApiParaPr.prototype["SetTopBorder"]              = ApiParaPr.prototype.SetTopBorder;
	ApiParaPr.prototype["SetBetweenBorder"]          = ApiParaPr.prototype.SetBetweenBorder;
	ApiParaPr.prototype["SetWidowControl"]           = ApiParaPr.prototype.SetWidowControl;
	ApiParaPr.prototype["SetTabs"]                   = ApiParaPr.prototype.SetTabs;
	ApiParaPr.prototype["SetNumPr"]                  = ApiParaPr.prototype.SetNumPr;
	ApiParaPr.prototype["SetBullet"]                 = ApiParaPr.prototype.SetBullet;
	ApiParaPr.prototype["GetStyle"]                  = ApiParaPr.prototype.GetStyle;
	ApiParaPr.prototype["GetSpacingLineValue"]       = ApiParaPr.prototype.GetSpacingLineValue;
	ApiParaPr.prototype["GetSpacingLineRule"]        = ApiParaPr.prototype.GetSpacingLineRule;
	ApiParaPr.prototype["GetSpacingBefore"]          = ApiParaPr.prototype.GetSpacingBefore;
	ApiParaPr.prototype["GetSpacingAfter"]           = ApiParaPr.prototype.GetSpacingAfter;
	ApiParaPr.prototype["GetShd"]                    = ApiParaPr.prototype.GetShd;
	ApiParaPr.prototype["GetJc"]                     = ApiParaPr.prototype.GetJc;
	ApiParaPr.prototype["GetIndRight"]               = ApiParaPr.prototype.GetIndRight;
	ApiParaPr.prototype["GetIndLeft"]                = ApiParaPr.prototype.GetIndLeft;
	ApiParaPr.prototype["GetIndFirstLine"]           = ApiParaPr.prototype.GetIndFirstLine;

	ApiTablePr.prototype["GetClassType"]             = ApiTablePr.prototype.GetClassType;
	ApiTablePr.prototype["SetStyleColBandSize"]      = ApiTablePr.prototype.SetStyleColBandSize;
	ApiTablePr.prototype["SetStyleRowBandSize"]      = ApiTablePr.prototype.SetStyleRowBandSize;
	ApiTablePr.prototype["SetJc"]                    = ApiTablePr.prototype.SetJc;
	ApiTablePr.prototype["SetShd"]                   = ApiTablePr.prototype.SetShd;
	ApiTablePr.prototype["SetTableBorderTop"]        = ApiTablePr.prototype.SetTableBorderTop;
	ApiTablePr.prototype["SetTableBorderBottom"]     = ApiTablePr.prototype.SetTableBorderBottom;
	ApiTablePr.prototype["SetTableBorderLeft"]       = ApiTablePr.prototype.SetTableBorderLeft;
	ApiTablePr.prototype["SetTableBorderRight"]      = ApiTablePr.prototype.SetTableBorderRight;
	ApiTablePr.prototype["SetTableBorderInsideH"]    = ApiTablePr.prototype.SetTableBorderInsideH;
	ApiTablePr.prototype["SetTableBorderInsideV"]    = ApiTablePr.prototype.SetTableBorderInsideV;
	ApiTablePr.prototype["SetTableCellMarginBottom"] = ApiTablePr.prototype.SetTableCellMarginBottom;
	ApiTablePr.prototype["SetTableCellMarginLeft"]   = ApiTablePr.prototype.SetTableCellMarginLeft;
	ApiTablePr.prototype["SetTableCellMarginRight"]  = ApiTablePr.prototype.SetTableCellMarginRight;
	ApiTablePr.prototype["SetTableCellMarginTop"]    = ApiTablePr.prototype.SetTableCellMarginTop;
	ApiTablePr.prototype["SetCellSpacing"]           = ApiTablePr.prototype.SetCellSpacing;
	ApiTablePr.prototype["SetTableInd"]              = ApiTablePr.prototype.SetTableInd;
	ApiTablePr.prototype["SetWidth"]                 = ApiTablePr.prototype.SetWidth;
	ApiTablePr.prototype["SetTableLayout"]           = ApiTablePr.prototype.SetTableLayout;

	ApiTableRowPr.prototype["GetClassType"]          = ApiTableRowPr.prototype.GetClassType;
	ApiTableRowPr.prototype["SetHeight"]             = ApiTableRowPr.prototype.SetHeight;
	ApiTableRowPr.prototype["SetTableHeader"]        = ApiTableRowPr.prototype.SetTableHeader;

	ApiTableCellPr.prototype["GetClassType"]         = ApiTableCellPr.prototype.GetClassType;
	ApiTableCellPr.prototype["SetShd"]               = ApiTableCellPr.prototype.SetShd;
	ApiTableCellPr.prototype["SetCellMarginBottom"]  = ApiTableCellPr.prototype.SetCellMarginBottom;
	ApiTableCellPr.prototype["SetCellMarginLeft"]    = ApiTableCellPr.prototype.SetCellMarginLeft;
	ApiTableCellPr.prototype["SetCellMarginRight"]   = ApiTableCellPr.prototype.SetCellMarginRight;
	ApiTableCellPr.prototype["SetCellMarginTop"]     = ApiTableCellPr.prototype.SetCellMarginTop;
	ApiTableCellPr.prototype["SetCellBorderBottom"]  = ApiTableCellPr.prototype.SetCellBorderBottom;
	ApiTableCellPr.prototype["SetCellBorderLeft"]    = ApiTableCellPr.prototype.SetCellBorderLeft;
	ApiTableCellPr.prototype["SetCellBorderRight"]   = ApiTableCellPr.prototype.SetCellBorderRight;
	ApiTableCellPr.prototype["SetCellBorderTop"]     = ApiTableCellPr.prototype.SetCellBorderTop;
	ApiTableCellPr.prototype["SetWidth"]             = ApiTableCellPr.prototype.SetWidth;
	ApiTableCellPr.prototype["SetVerticalAlign"]     = ApiTableCellPr.prototype.SetVerticalAlign;
	ApiTableCellPr.prototype["SetTextDirection"]     = ApiTableCellPr.prototype.SetTextDirection;
	ApiTableCellPr.prototype["SetNoWrap"]            = ApiTableCellPr.prototype.SetNoWrap;

	ApiTableStylePr.prototype["GetClassType"]        = ApiTableStylePr.prototype.GetClassType;
	ApiTableStylePr.prototype["GetType"]             = ApiTableStylePr.prototype.GetType;
	ApiTableStylePr.prototype["GetTextPr"]           = ApiTableStylePr.prototype.GetTextPr;
	ApiTableStylePr.prototype["GetParaPr"]           = ApiTableStylePr.prototype.GetParaPr;
	ApiTableStylePr.prototype["GetTablePr"]          = ApiTableStylePr.prototype.GetTablePr;
	ApiTableStylePr.prototype["GetTableRowPr"]       = ApiTableStylePr.prototype.GetTableRowPr;
	ApiTableStylePr.prototype["GetTableCellPr"]      = ApiTableStylePr.prototype.GetTableCellPr;

	ApiDrawing.prototype["GetClassType"]             = ApiDrawing.prototype.GetClassType;
	ApiDrawing.prototype["SetSize"]                  = ApiDrawing.prototype.SetSize;
	ApiDrawing.prototype["SetWrappingStyle"]         = ApiDrawing.prototype.SetWrappingStyle;
	ApiDrawing.prototype["SetHorAlign"]              = ApiDrawing.prototype.SetHorAlign;
	ApiDrawing.prototype["SetVerAlign"]              = ApiDrawing.prototype.SetVerAlign;
	ApiDrawing.prototype["SetHorPosition"]           = ApiDrawing.prototype.SetHorPosition;
	ApiDrawing.prototype["SetVerPosition"]           = ApiDrawing.prototype.SetVerPosition;
	ApiDrawing.prototype["SetDistances"]             = ApiDrawing.prototype.SetDistances;
	ApiDrawing.prototype["GetParentParagraph"]       = ApiDrawing.prototype.GetParentParagraph;
	ApiDrawing.prototype["GetParentContentControl"]  = ApiDrawing.prototype.GetParentContentControl;
	ApiDrawing.prototype["GetParentTable"]           = ApiDrawing.prototype.GetParentTable;
	ApiDrawing.prototype["GetParentTableCell"]       = ApiDrawing.prototype.GetParentTableCell;
	ApiDrawing.prototype["Delete"]                   = ApiDrawing.prototype.Delete;
	ApiDrawing.prototype["Copy"]                     = ApiDrawing.prototype.Copy;
	ApiDrawing.prototype["InsertInContentControl"]   = ApiDrawing.prototype.InsertInContentControl;
	ApiDrawing.prototype["InsertParagraph"]          = ApiDrawing.prototype.InsertParagraph;
	ApiDrawing.prototype["Select"]                   = ApiDrawing.prototype.Select;
	ApiDrawing.prototype["AddBreak"]                 = ApiDrawing.prototype.AddBreak;
	ApiDrawing.prototype["SetHorFlip"]               = ApiDrawing.prototype.SetHorFlip;
	ApiDrawing.prototype["SetVertFlip"]              = ApiDrawing.prototype.SetVertFlip;
	ApiDrawing.prototype["ScaleHeight"]              = ApiDrawing.prototype.ScaleHeight;
	ApiDrawing.prototype["ScaleWidth"]               = ApiDrawing.prototype.ScaleWidth;
	ApiDrawing.prototype["Fill"]                     = ApiDrawing.prototype.Fill;
	ApiDrawing.prototype["SetOutLine"]               = ApiDrawing.prototype.SetOutLine;
	ApiDrawing.prototype["GetNextDrawing"]           = ApiDrawing.prototype.GetNextDrawing;
	ApiDrawing.prototype["GetPrevDrawing"]           = ApiDrawing.prototype.GetPrevDrawing;

	ApiImage.prototype["GetClassType"]               = ApiImage.prototype.GetClassType;
	ApiImage.prototype["GetNextImage"]               = ApiImage.prototype.GetNextImage;
	ApiImage.prototype["GetPrevImage"]               = ApiImage.prototype.GetPrevImage;

	ApiShape.prototype["GetClassType"]               = ApiShape.prototype.GetClassType;
	ApiShape.prototype["GetDocContent"]              = ApiShape.prototype.GetDocContent;
	ApiShape.prototype["GetContent"]                 = ApiShape.prototype.GetContent;
	ApiShape.prototype["SetVerticalTextAlign"]       = ApiShape.prototype.SetVerticalTextAlign;
	ApiShape.prototype["SetPaddings"]                = ApiShape.prototype.SetPaddings;
	ApiShape.prototype["GetNextShape"]               = ApiShape.prototype.GetNextShape;
	ApiShape.prototype["GetPrevShape"]               = ApiShape.prototype.GetPrevShape;

	ApiChart.prototype["GetClassType"]                 = ApiChart.prototype.GetClassType;
	ApiChart.prototype["SetTitle"]                     = ApiChart.prototype.SetTitle;
	ApiChart.prototype["SetHorAxisTitle"]              = ApiChart.prototype.SetHorAxisTitle;
	ApiChart.prototype["SetVerAxisTitle"]              = ApiChart.prototype.SetVerAxisTitle;
	ApiChart.prototype["SetVerAxisOrientation"]        = ApiChart.prototype.SetVerAxisOrientation;
	ApiChart.prototype["SetHorAxisOrientation"]        = ApiChart.prototype.SetHorAxisOrientation;
	ApiChart.prototype["SetLegendPos"]                 = ApiChart.prototype.SetLegendPos;
	ApiChart.prototype["SetLegendFontSize"]            = ApiChart.prototype.SetLegendFontSize;
	ApiChart.prototype["SetShowDataLabels"]            = ApiChart.prototype.SetShowDataLabels;
	ApiChart.prototype["SetShowPointDataLabel"]        = ApiChart.prototype.SetShowPointDataLabel;
	ApiChart.prototype["SetVertAxisTickLabelPosition"] = ApiChart.prototype.SetVertAxisTickLabelPosition;
	ApiChart.prototype["SetHorAxisTickLabelPosition"]  = ApiChart.prototype.SetHorAxisTickLabelPosition;

	ApiChart.prototype["SetHorAxisMajorTickMark"]      =  ApiChart.prototype.SetHorAxisMajorTickMark;
	ApiChart.prototype["SetHorAxisMinorTickMark"]      =  ApiChart.prototype.SetHorAxisMinorTickMark;
	ApiChart.prototype["SetVertAxisMajorTickMark"]     =  ApiChart.prototype.SetVertAxisMajorTickMark;
	ApiChart.prototype["SetVertAxisMinorTickMark"]     =  ApiChart.prototype.SetVertAxisMinorTickMark;
	ApiChart.prototype["SetMajorVerticalGridlines"]    =  ApiChart.prototype.SetMajorVerticalGridlines;
	ApiChart.prototype["SetMinorVerticalGridlines"]    =  ApiChart.prototype.SetMinorVerticalGridlines;
	ApiChart.prototype["SetMajorHorizontalGridlines"]  =  ApiChart.prototype.SetMajorHorizontalGridlines;
	ApiChart.prototype["SetMinorHorizontalGridlines"]  =  ApiChart.prototype.SetMinorHorizontalGridlines;
	ApiChart.prototype["SetHorAxisLablesFontSize"]     =  ApiChart.prototype.SetHorAxisLablesFontSize;
	ApiChart.prototype["SetVertAxisLablesFontSize"]    =  ApiChart.prototype.SetVertAxisLablesFontSize;
	ApiChart.prototype["GetNextChart"]                 = ApiChart.prototype.GetNextChart;
    ApiChart.prototype["GetPrevChart"]                 = ApiChart.prototype.GetPrevChart;

	ApiFill.prototype["GetClassType"]                = ApiFill.prototype.GetClassType;

	ApiStroke.prototype["GetClassType"]              = ApiStroke.prototype.GetClassType;

	ApiGradientStop.prototype["GetClassType"]        = ApiGradientStop.prototype.GetClassType;

	ApiUniColor.prototype["GetClassType"]            = ApiUniColor.prototype.GetClassType;

	ApiRGBColor.prototype["GetClassType"]            = ApiRGBColor.prototype.GetClassType;

	ApiSchemeColor.prototype["GetClassType"]         = ApiSchemeColor.prototype.GetClassType;

	ApiPresetColor.prototype["GetClassType"]         = ApiPresetColor.prototype.GetClassType;

	ApiBullet.prototype["GetClassType"]              = ApiBullet.prototype.GetClassType;

	ApiInlineLvlSdt.prototype["GetClassType"]           = ApiInlineLvlSdt.prototype.GetClassType;
	ApiInlineLvlSdt.prototype["SetLock"]                = ApiInlineLvlSdt.prototype.SetLock;
	ApiInlineLvlSdt.prototype["GetLock"]                = ApiInlineLvlSdt.prototype.GetLock;
	ApiInlineLvlSdt.prototype["SetTag"]                 = ApiInlineLvlSdt.prototype.SetTag;
	ApiInlineLvlSdt.prototype["GetTag"]                 = ApiInlineLvlSdt.prototype.GetTag;
	ApiInlineLvlSdt.prototype["SetLabel"]               = ApiInlineLvlSdt.prototype.SetLabel;
	ApiInlineLvlSdt.prototype["GetLabel"]               = ApiInlineLvlSdt.prototype.GetLabel;
	ApiInlineLvlSdt.prototype["SetAlias"]               = ApiInlineLvlSdt.prototype.SetAlias;
	ApiInlineLvlSdt.prototype["GetAlias"]               = ApiInlineLvlSdt.prototype.GetAlias;
	ApiInlineLvlSdt.prototype["GetElementsCount"]       = ApiInlineLvlSdt.prototype.GetElementsCount;
	ApiInlineLvlSdt.prototype["GetElement"]             = ApiInlineLvlSdt.prototype.GetElement;
	ApiInlineLvlSdt.prototype["RemoveElement"]          = ApiInlineLvlSdt.prototype.RemoveElement;
	ApiInlineLvlSdt.prototype["RemoveAllElements"]      = ApiInlineLvlSdt.prototype.RemoveAllElements;
	ApiInlineLvlSdt.prototype["AddElement"]             = ApiInlineLvlSdt.prototype.AddElement;
	ApiInlineLvlSdt.prototype["Push"]                   = ApiInlineLvlSdt.prototype.Push;
	ApiInlineLvlSdt.prototype["AddText"]                = ApiInlineLvlSdt.prototype.AddText;
	ApiInlineLvlSdt.prototype["Delete"]                 = ApiInlineLvlSdt.prototype.Delete;
	ApiInlineLvlSdt.prototype["SetTextPr"]              = ApiInlineLvlSdt.prototype.SetTextPr;
	ApiInlineLvlSdt.prototype["GetParentContentControl"]= ApiInlineLvlSdt.prototype.GetParentContentControl;
	ApiInlineLvlSdt.prototype["GetParentTable"]         = ApiInlineLvlSdt.prototype.GetParentTable;
	ApiInlineLvlSdt.prototype["GetParentTableCell"]     = ApiInlineLvlSdt.prototype.GetParentTableCell;
	ApiInlineLvlSdt.prototype["GetRange"]               = ApiInlineLvlSdt.prototype.GetRange;
	ApiInlineLvlSdt.prototype["Copy"]                   = ApiInlineLvlSdt.prototype.Copy;

	ApiBlockLvlSdt.prototype["GetClassType"]            = ApiBlockLvlSdt.prototype.GetClassType;
	ApiBlockLvlSdt.prototype["SetLock"]                 = ApiBlockLvlSdt.prototype.SetLock;
	ApiBlockLvlSdt.prototype["GetLock"]                 = ApiBlockLvlSdt.prototype.GetLock;
	ApiBlockLvlSdt.prototype["SetTag"]                  = ApiBlockLvlSdt.prototype.SetTag;
	ApiBlockLvlSdt.prototype["GetTag"]                  = ApiBlockLvlSdt.prototype.GetTag;
	ApiBlockLvlSdt.prototype["SetLabel"]                = ApiBlockLvlSdt.prototype.SetLabel;
	ApiBlockLvlSdt.prototype["GetLabel"]                = ApiBlockLvlSdt.prototype.GetLabel;
	ApiBlockLvlSdt.prototype["SetAlias"]                = ApiBlockLvlSdt.prototype.SetAlias;
	ApiBlockLvlSdt.prototype["GetAlias"]                = ApiBlockLvlSdt.prototype.GetAlias;
	ApiBlockLvlSdt.prototype["GetContent"]              = ApiBlockLvlSdt.prototype.GetContent;
	ApiBlockLvlSdt.prototype["GetAllContentControls"]   = ApiBlockLvlSdt.prototype.GetAllContentControls;
	ApiBlockLvlSdt.prototype["GetAllParagraphs"]        = ApiBlockLvlSdt.prototype.GetAllParagraphs;
	ApiBlockLvlSdt.prototype["GetAllTablesOnPage"]      = ApiBlockLvlSdt.prototype.GetAllTablesOnPage;
	ApiBlockLvlSdt.prototype["RemoveAllElements"]       = ApiBlockLvlSdt.prototype.RemoveAllElements;
	ApiBlockLvlSdt.prototype["Delete"]                  = ApiBlockLvlSdt.prototype.Delete;
	ApiBlockLvlSdt.prototype["SetTextPr"]               = ApiBlockLvlSdt.prototype.SetTextPr;
	ApiBlockLvlSdt.prototype["GetAllDrawingObjects"]    = ApiBlockLvlSdt.prototype.GetAllDrawingObjects;
	ApiBlockLvlSdt.prototype["GetParentContentControl"] = ApiBlockLvlSdt.prototype.GetParentContentControl;
	ApiBlockLvlSdt.prototype["GetParentTable"]          = ApiBlockLvlSdt.prototype.GetParentTable;
	ApiBlockLvlSdt.prototype["GetParentTableCell"]      = ApiBlockLvlSdt.prototype.GetParentTableCell;
	ApiBlockLvlSdt.prototype["Push"]                    = ApiBlockLvlSdt.prototype.Push;
	ApiBlockLvlSdt.prototype["AddElement"]              = ApiBlockLvlSdt.prototype.AddElement;
	ApiBlockLvlSdt.prototype["AddText"]                 = ApiBlockLvlSdt.prototype.AddText;
	ApiBlockLvlSdt.prototype["GetRange"]                = ApiBlockLvlSdt.prototype.GetRange;
	ApiBlockLvlSdt.prototype["Search"]                  = ApiBlockLvlSdt.prototype.Search;
	ApiBlockLvlSdt.prototype["Select"]                  = ApiBlockLvlSdt.prototype.Select;

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function private_GetDrawingDocument()
	{
		return editor.WordControl.m_oLogicDocument.DrawingDocument;
	}

	function private_PushElementToParagraph(oPara, oElement)
	{
		// Добавляем не в конец из-за рана с символом конца параграфа TODO: ParaEnd
		oPara.Add_ToContent(oPara.Content.length - 1, oElement);
	}

	function private_IsSupportedParaElement(oElement)
	{
		if (oElement instanceof ApiRun
			|| oElement instanceof ApiInlineLvlSdt 
			|| oElement instanceof ApiHyperlink)
			return true;

		return false;
	}
	
	function private_GetSupportedParaElement(oElement)
	{
		if (oElement instanceof ParaRun)
			return new ApiRun(oElement);
		else if (oElement instanceof CInlineLevelSdt)
			return new ApiInlineLvlSdt(oElement);
		else if (oElement instanceof ParaHyperlink)
			return new ApiHyperlink(oElement);
		else
			return new ApiUnsupported();
	}

	function private_GetLogicDocument()
	{
		return editor.WordControl.m_oLogicDocument;
	}

	function private_Twips2MM(twips)
	{
		return 25.4 / 72.0 / 20 * twips;
	}

	function private_EMU2MM(EMU)
	{
		return EMU / 36000.0;
	}

	function private_GetHps(hps)
	{
		return Math.ceil(hps) / 2.0;
	}

	function private_GetColor(r, g, b, Auto)
	{
		return new AscCommonWord.CDocumentColor(r, g, b, Auto ? Auto : false);
	}

	function private_GetTabStop(nPos, sValue)
	{
		var nType = tab_Left;
		if ("left" === sValue)
			nType = tab_Left;
		else if ("right" === sValue)
			nType = tab_Right;
		else if ("clear" === sValue)
			nType = tab_Clear;
		else if ("center" === sValue)
			nType = tab_Center;

		return new CParaTab(nType, private_Twips2MM(nPos));
	}

	function private_GetParaAlign(sJc)
	{
		if ("left" === sJc)
			return align_Left;
		else if ("right" === sJc)
			return align_Right;
		else if ("both" === sJc)
			return align_Justify;
		else if ("center" === sJc)
			return align_Center;

		return undefined;
	}

	function private_GetTableBorder(sType, nSize, nSpace, r, g, b)
	{
		var oBorder = new CDocumentBorder();

		if ("none" === sType)
		{
			oBorder.Value = border_None;
			oBorder.Size  = 0;
			oBorder.Space = 0;
			oBorder.Color.Set(0, 0, 0, true);
		}
		else
		{
			if ("single" === sType)
				oBorder.Value = border_Single;

			oBorder.Size  = private_Pt_8ToMM(nSize);
			oBorder.Space = private_PtToMM(nSpace);
			oBorder.Color.Set(r, g, b);
		}

		return oBorder;
	}

	function private_GetTableMeasure(sType, nValue)
	{
		var nType = tblwidth_Auto;
		var nW    = 0;
		if ("auto" === sType)
		{
			nType = tblwidth_Auto;
			nW    = 0;
		}
		else if ("nil" === sType)
		{
			nType = tblwidth_Nil;
			nW    = 0;
		}
		else if ("percent" === sType)
		{
			nType = tblwidth_Pct;
			nW    = private_GetInt(nValue, null, null);
		}
		else if ("twips" === sType)
		{
			nType = tblwidth_Mm;
			nW    = private_Twips2MM(nValue);
		}

		return new CTableMeasurement(nType, nW);
	}

	function private_GetShd(sType, r, g, b, isAuto)
	{
		var oShd = new CDocumentShd();

		if ("nil" === sType)
			oShd.Value = Asc.c_oAscShdNil;
		else if ("clear" === sType)
			oShd.Value = Asc.c_oAscShdClear;

		oShd.Color.Set(r, g, b, isAuto);
		oShd.Fill = new CDocumentColor(r, g, b, isAuto);
		return oShd;
	}

	function private_GetBoolean(bValue, bDefValue)
	{
		if (true === bValue)
			return true;
		else if (false === bValue)
			return false;
		else
			return (undefined !== bDefValue ? bDefValue : false);
	}

	function private_GetInt(nValue, nMin, nMax)
	{
		var nResult = nValue | 0;

		if (undefined !== nMin && null !== nMin)
			nResult = Math.max(nMin, nResult);

		if (undefined !== nMax && null !== nMax)
			nResult = Math.min(nMax, nResult);

		return nResult;
	}

	function private_PtToMM(pt)
	{
		return 25.4 / 72.0 * pt;
	}

	function private_Pt_8ToMM(pt)
	{
		return 25.4 / 72.0 / 8 * pt;
	}

	function private_StartSilentMode()
	{
		private_GetLogicDocument().Start_SilentMode();
	}
	function private_EndSilentMode()
	{
		private_GetLogicDocument().End_SilentMode(false);
	}
	function private_GetAlignH(sAlign)
	{
		if ("left" === sAlign)
			return c_oAscAlignH.Left;
		else if ("right" === sAlign)
			return c_oAscAlignH.Right;
		else if ("center" === sAlign)
			return c_oAscAlignH.Center;

		return c_oAscAlignH.Left;
	}

	function private_GetAlignV(sAlign)
	{
		if ("top" === sAlign)
			return c_oAscAlignV.Top;
		else if ("bottom" === sAlign)
			return c_oAscAlignV.Bottom;
		else if ("center" === sAlign)
			return c_oAscAlignV.Center;

		return c_oAscAlignV.Center;
	}
	function private_GetRelativeFromH(sRel)
	{
		if ("character" === sRel)
			return Asc.c_oAscRelativeFromH.Character;
		else if ("column" === sRel)
			return Asc.c_oAscRelativeFromH.Column;
		else if ("leftMargin" === sRel)
			return Asc.c_oAscRelativeFromH.LeftMargin;
		else if ("rightMargin" === sRel)
			return Asc.c_oAscRelativeFromH.RightMargin;
		else if ("margin" === sRel)
			return Asc.c_oAscRelativeFromH.Margin;
		else if ("page" === sRel)
			return Asc.c_oAscRelativeFromH.Page;

		return Asc.c_oAscRelativeFromH.Page;
	}

	function private_GetRelativeFromV(sRel)
	{
		if ("bottomMargin" === sRel)
			return Asc.c_oAscRelativeFromV.BottomMargin;
		else if ("topMargin" === sRel)
			return Asc.c_oAscRelativeFromV.TopMargin;
		else if ("margin" === sRel)
			return Asc.c_oAscRelativeFromV.Margin;
		else if ("page" === sRel)
			return Asc.c_oAscRelativeFromV.Page;
		else if ("line" === sRel)
			return Asc.c_oAscRelativeFromV.Line;
		else if ("paragraph" === sRel)
			return Asc.c_oAscRelativeFromV.Paragraph;

		return Asc.c_oAscRelativeFromV.Page;
	}

	function private_CreateWatermark(sText, bDiagonal){
		var oLogicDocument = private_GetLogicDocument();
		var oProps = new Asc.CAscWatermarkProperties();
		oProps.put_Type(Asc.c_oAscWatermarkType.Text);
		oProps.put_IsDiagonal(bDiagonal === true);
		oProps.put_Text(sText);
		oProps.put_Opacity(127);
		var oTextPr = new Asc.CTextProp();
		oTextPr.put_FontSize(-1);
		oTextPr.put_FontFamily(new AscCommon.asc_CTextFontFamily({Name : "Arial", Index : -1}));
		oTextPr.put_Color(AscCommon.CreateAscColorCustom(192, 192, 192));
		oProps.put_TextPr(oTextPr);
		var oDrawing = oLogicDocument.DrawingObjects.createWatermark(oProps);
		var oApiShape = new ApiShape(oDrawing.GraphicObj);
		return oApiShape;
	}


	function privateInsertWatermarkToContent(oApi, oContent, sText, bIsDiagonal){
		if(oContent){
			var nElementsCount = oContent.GetElementsCount();
			for(var i = 0; i < nElementsCount; ++i){
				var oElement = oContent.GetElement(i);
				if(oElement.GetClassType() === "paragraph"){
					oElement.AddDrawing(private_CreateWatermark(sText, bIsDiagonal));
					break;
				}
			}
			if(i === nElementsCount){
				oElement = oApi.CreateParagraph();
				oElement.AddDrawing(private_CreateWatermark(sText, bIsDiagonal));
				oContent.Push(oElement);
			}
		}
	}

	ApiDocument.prototype.OnChangeParaPr = function(oApiParaPr)
	{
		var oStyles = this.Document.Get_Styles();
		oStyles.Set_DefaultParaPr(oApiParaPr.ParaPr);
		oApiParaPr.ParaPr = oStyles.Get_DefaultParaPr().Copy();
	};
	ApiDocument.prototype.OnChangeTextPr = function(oApiTextPr)
	{
		var oStyles = this.Document.Get_Styles();
		oStyles.Set_DefaultTextPr(oApiTextPr.TextPr);
		oApiTextPr.TextPr = oStyles.Get_DefaultTextPr().Copy();
	};
	ApiParagraph.prototype.private_GetImpl = function()
	{
		return this.Paragraph;
	};
	ApiParagraph.prototype.OnChangeParaPr = function(oApiParaPr)
	{
		this.Paragraph.Set_Pr(oApiParaPr.ParaPr);
		oApiParaPr.ParaPr = this.Paragraph.Pr.Copy();
	};
	ApiParagraph.prototype.OnChangeTextPr = function(oApiTextPr)
	{
		this.Paragraph.TextPr.Set_Value(oApiTextPr.TextPr);
		oApiTextPr.TextPr = this.Paragraph.TextPr.Value.Copy();
	};
	ApiRun.prototype.private_GetImpl = function()
	{
		return this.Run;
	};
	ApiHyperlink.prototype.private_GetImpl = function()
	{
		return this.ParaHyperlink;
	};
	ApiRun.prototype.OnChangeTextPr = function(oApiTextPr)
	{
		this.Run.Set_Pr(oApiTextPr.TextPr);
		oApiTextPr.TextPr = this.Run.Pr.Copy();
	};
	ApiTable.prototype.private_GetImpl = function()
	{
		return this.Table;
	};
	ApiTable.prototype.OnChangeTablePr = function(oApiTablePr)
	{
		this.Table.Set_Pr(oApiTablePr.TablePr);
		oApiTablePr.TablePr = this.Table.Pr.Copy();
	};
	ApiTable.prototype.private_PrepareTableForActions = function()
	{
		this.Table.private_RecalculateGrid();
		this.Table.private_UpdateCellsGrid();
	};
	ApiStyle.prototype.OnChangeTextPr = function(oApiTextPr)
	{
		this.Style.Set_TextPr(oApiTextPr.TextPr);
		oApiTextPr.TextPr = this.Style.TextPr.Copy();
	};
	ApiStyle.prototype.OnChangeParaPr = function(oApiParaPr)
	{
		this.Style.Set_ParaPr(oApiParaPr.ParaPr);
		oApiParaPr.ParaPr = this.Style.ParaPr.Copy();
	};
	ApiStyle.prototype.OnChangeTablePr = function(oApiTablePr)
	{
		this.Style.Set_TablePr(oApiTablePr.TablePr);
		oApiTablePr.TablePr = this.Style.TablePr.Copy();
	};
	ApiStyle.prototype.OnChangeTableRowPr = function(oApiTableRowPr)
	{
		this.Style.Set_TableRowPr(oApiTableRowPr.RowPr);
		oApiTableRowPr.RowPr = this.Style.TableRowPr.Copy();
	};
	ApiStyle.prototype.OnChangeTableCellPr = function(oApiTableCellPr)
	{
		this.Style.Set_TableCellPr(oApiTableCellPr.CellPr);
		oApiTableCellPr.CellPr = this.Style.TableCellPr.Copy();
	};
	ApiStyle.prototype.OnChangeTableStylePr = function(oApiTableStylePr)
	{
		var sType = oApiTableStylePr.GetType();
		switch(sType)
		{
			case "topLeftCell":
			{
				this.Style.Set_TableTLCell(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableTLCell.Copy();
				break;
			}
			case "topRightCell":
			{
				this.Style.Set_TableTRCell(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableTRCell.Copy();
				break;
			}
			case "bottomLeftCell":
			{
				this.Style.Set_TableBLCell(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableBLCell.Copy();
				break;
			}
			case "bottomRightCell":
			{
				this.Style.Set_TableBRCell(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableBRCell.Copy();
				break;
			}
			case "firstRow":
			{
				this.Style.Set_TableFirstRow(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableFirstRow.Copy();
				break;
			}
			case "lastRow":
			{
				this.Style.Set_TableLastRow(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableLastRow.Copy();
				break;
			}
			case "firstColumn":
			{
				this.Style.Set_TableFirstCol(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableFirstCol.Copy();
				break;
			}
			case "lastColumn":
			{
				this.Style.Set_TableLastCol(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableLastCol.Copy();
				break;
			}
			case "bandedColumn":
			{
				this.Style.Set_TableBand1Vert(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableBand1Vert.Copy();
				break;
			}
			case "bandedColumnEven":
			{
				this.Style.Set_TableBand2Vert(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableBand2Vert.Copy();
				break;
			}
			case "bandedRow":
			{
				this.Style.Set_TableBand1Horz(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableBand1Horz.Copy();
				break;
			}
			case "bandedRowEven":
			{
				this.Style.Set_TableBand2Horz(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableBand2Horz.Copy();
				break;
			}
			case "wholeTable":
			{
				this.Style.Set_TableWholeTable(oApiTableStylePr.TableStylePr);
				oApiTableStylePr.TableStylePr = this.Style.TableWholeTable.Copy();
				break;
			}
		}
	};
	ApiNumberingLevel.prototype.OnChangeTextPr = function(oApiTextPr)
	{
		this.Num.SetTextPr(this.Lvl, oApiTextPr.TextPr);
		oApiTextPr.TextPr = this.Num.GetLvl(this.Lvl).GetTextPr().Copy();
	};
	ApiNumberingLevel.prototype.OnChangeParaPr = function(oApiParaPr)
	{
		this.Num.SetParaPr(this.Lvl, oApiParaPr.ParaPr);
		oApiParaPr.ParaPr = this.Num.GetLvl(this.Lvl).GetParaPr().Copy();
	};
	ApiTableRow.prototype.OnChangeTableRowPr = function(oApiTableRowPr)
	{
		this.Row.Set_Pr(oApiTableRowPr.RowPr);
		oApiTableRowPr.RowPr = this.Row.Pr.Copy();
	};
	ApiTableCell.prototype.OnChangeTableCellPr = function(oApiTableCellPr)
	{
		this.Cell.Set_Pr(oApiTableCellPr.CellPr);
		oApiTableCellPr.CellPr = this.Cell.Pr.Copy();
	};
	ApiTextPr.prototype.private_OnChange = function()
	{
		this.Parent.OnChangeTextPr(this);
	};
	ApiParaPr.prototype.private_OnChange = function()
	{
		this.Parent.OnChangeParaPr(this);
	};
	ApiTablePr.prototype.private_OnChange = function()
	{
		this.Parent.OnChangeTablePr(this);
	};
	ApiTableRowPr.prototype.private_OnChange = function()
	{
		this.Parent.OnChangeTableRowPr(this);
	};
	ApiTableCellPr.prototype.private_OnChange = function()
	{
		this.Parent.OnChangeTableCellPr(this);
	};
	ApiTableStylePr.prototype.private_OnChange = function()
	{
		this.Parent.OnChangeTableStylePr(this);
	};
	ApiTableStylePr.prototype.OnChangeTextPr = function()
	{
		this.private_OnChange();
	};
	ApiTableStylePr.prototype.OnChangeParaPr = function()
	{
		this.private_OnChange();
	};
	ApiTableStylePr.prototype.OnChangeTablePr = function()
	{
		this.private_OnChange();
	};
	ApiTableStylePr.prototype.OnChangeTableRowPr = function()
	{
		this.private_OnChange();
	};
	ApiTableStylePr.prototype.OnChangeTableCellPr = function()
	{
		this.private_OnChange();
	};
	ApiInlineLvlSdt.prototype.private_GetImpl = function()
	{
		return this.Sdt;
	};
	ApiBlockLvlSdt.prototype.private_GetImpl = function()
	{
		return this.Sdt;
	};

	Api.prototype.private_CreateApiParagraph = function(oParagraph){
		return new ApiParagraph(oParagraph);
	};

	Api.prototype.private_CreateApiDocContent = function(oDocContent){
		return new ApiDocumentContent(oDocContent);
	};
}(window, null));



