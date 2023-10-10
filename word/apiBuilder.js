/*
 * (c) Copyright Ascensio System SIA 2010-2023
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
					if (firstDocPos[nPos].Class !== secondDocPos[nPos].Class || firstDocPos[nPos].Position !== secondDocPos[nPos].Position)
						return false;
				}
				return true;
			}
			return false;
		}

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
	}
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
	}
	function private_RefreshRangesPosition()
	{
		var Document  = private_GetLogicDocument();
		var Range     = null;

		for (var nRange = 0; nRange < arrApiRanges.length; nRange++)
		{
			Range = arrApiRanges[nRange];
			Document.RefreshDocumentPositions([Range.StartPos, Range.EndPos]);
		}
	}
	/**
	 * Returns the first Run in the array specified.
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
				else if (TempPos[Pos].Position > MinPos[Pos].Position)
					break;
			}
		}
		
		return arrRuns[min_pos_Index];
	}
	/**
	 * Returns the last Run in the array specified.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {Array} arrRuns - Array of Runs.
	 * @return {ApiRun | null} - returns null if arrRuns is invalid. 
	 */
	function private_GetLastRunInArray(arrRuns)
	{
		if (!Array.isArray(arrRuns))
			return null;
			
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
				else if (TempPos[Pos].Position < MaxPos[Pos].Position)
					break;
			}
		}
		return arrRuns[max_pos_Index];
	}
	function private_isMonospaceFont(sFontName)
	{
		return sFontName === 'Courier New'
			|| sFontName === 'Consolas'
			|| sFontName === 'Inconsolata'
			|| sFontName === 'Roboto Mono'
			|| sFontName === 'Source Code Pro';
	}

	function private_CheckDrawingOnAdd(oApiDrawing) {
		if(window.g_asc_plugins)
		{
			let oDrawing = oApiDrawing.Drawing;
			if(oDrawing)
			{
				if(oDrawing.GraphicObj &&
					oDrawing.GraphicObj.getObjectType() === AscDFH.historyitem_type_OleObject)
				{
					let oData = oDrawing.GraphicObj.getDataObject();
					window.g_asc_plugins.onPluginEvent("onInsertOleObjects", [oData]);
				}
			}
		}
	}

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
	 * Class representing the Markdown conversion processing.
	 * Each Range object is determined by the position of the start and end characters.
	 * @constructor
	 */
	function CMarkdownConverter(oConfig)
	{
		this.HtmlTags =
		{
			Bold: '<strong>',
			Italic: '<em>',
			Span: '<span>',
			CodeLine: '<pre>',
			SubScript: '<sub>',
			SupScript: '<sup>',
			Strikeout: '<del>',
			Code: '<code>',
			Paragraph: '<p>',
			Headings: ['<h1>', '<h2>', '<h3>', '<h4>', '<h5>', '<h6>'],
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
			Strikeout: '~~',
			Code:  '`',
			Headings: ['#', '##', '###', '####', '#####', '######'],
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
	}
	CMarkdownConverter.prototype.constructor = CMarkdownConverter;

	CMarkdownConverter.prototype.WrapInSymbol = function(sText, sSyblols, sWrapType)
	{
		var nFirstNonSpaceChar = sText.search(/\S/);
		var nSpaceCharsCountOnEnd = 0;
		if (nFirstNonSpaceChar !== - 1)
			nSpaceCharsCountOnEnd = sText.slice(nFirstNonSpaceChar).length - sText.slice(nFirstNonSpaceChar).trim().length;
		else
			nSpaceCharsCountOnEnd = sText.length - sText.slice(nFirstNonSpaceChar).trim().length;

		switch (sWrapType)
		{
			case 'open':
				// пробелов нет в начале
				if (nFirstNonSpaceChar === 0)
					return sSyblols + sText
				// строка из пробелов
				else if (nFirstNonSpaceChar === -1)
					return sText + sSyblols
				// в начале строки есть пробелы
				else if (nFirstNonSpaceChar !== -1)
					return sText.slice(0, nFirstNonSpaceChar) + sSyblols + sText.slice(nFirstNonSpaceChar);
			case 'close':
				// пробелов нет в конце
				if (nSpaceCharsCountOnEnd === 0)
					return sText + sSyblols;
				// строка из пробелов
				else if (nFirstNonSpaceChar === -1)
					return sSyblols + sText;
				// в конце строки есть пробелы
				else if (nSpaceCharsCountOnEnd !== 0)
					return sText.slice(0, sText.length - nSpaceCharsCountOnEnd) + sSyblols + sText.slice(sText.length - nSpaceCharsCountOnEnd);
			case 'wholly':
			default:
				// пробелов нет в начале и нет в конце
				if (nFirstNonSpaceChar === 0 && nSpaceCharsCountOnEnd === 0)
					return sSyblols + sText + sSyblols;
				// пробелов нет в начале и есть в конце
				else if (nFirstNonSpaceChar === 0 && nSpaceCharsCountOnEnd !== 0)
					return sSyblols + sText.slice(0, sText.length - nSpaceCharsCountOnEnd) + sSyblols + sText.slice(sText.length - nSpaceCharsCountOnEnd);
				// пробелы есть в начале и нет в конце
				else if (nFirstNonSpaceChar !== 0 && nSpaceCharsCountOnEnd === 0)
					return sText.slice(0, nFirstNonSpaceChar) + sSyblols + sText.slice(nFirstNonSpaceChar) + sSyblols;
				// пробелы есть в начале и есть в конце
				else if (nFirstNonSpaceChar !== 0 && nSpaceCharsCountOnEnd !== 0)
					return sText.slice(0, nFirstNonSpaceChar) + sSyblols + sText.slice(nFirstNonSpaceChar, sText.length - nSpaceCharsCountOnEnd) + sSyblols + sText.slice(sText.length - nSpaceCharsCountOnEnd);
		}
	};
	CMarkdownConverter.prototype.WrapInTag = function(sText, sHtmlTag, sWrapType, sStyle)
	{
		switch (sWrapType)
		{
			case 'open':
				if (!sStyle)
					return sHtmlTag + sText;
				return sHtmlTag.replace('>', ' style="' + sStyle + ';">');
			case 'close':
				return sText + sHtmlTag.replace('<', '</');
			case 'wholly':
			default:
				if (!sStyle)
					return sHtmlTag + sText + sHtmlTag.replace('<', '</');
				return sHtmlTag.replace('>', ' style="' + sStyle + ';">') + sText + sHtmlTag.replace('<', '</');
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
			oSelectedContent = oDocument.Document.GetSelectedContent(false, {SaveNumberingValues: true});
			for (var nElm = 0; nElm < oSelectedContent.Elements.length; nElm ++)
			{
				oTempElm = oSelectedContent.Elements[nElm].Element;
				if (oTempElm instanceof CTable)
					arrSelectedContent.push(new ApiTable(oTempElm));
				else if (oTempElm instanceof Paragraph)
					arrSelectedContent.push(new ApiParagraph(oTempElm));
				else if (oTempElm instanceof ParaRun)
					arrSelectedContent.push(new ApiRun(oTempElm));
			}

			for (nElm = 0; nElm < arrSelectedContent.length; nElm ++)
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
			}

			for (nElm = 0; nElm < arrSelectedContent.length; nElm ++)
				sOutputText += this.HandleChildElement(arrSelectedContent[nElm], 'html');
		}
		else
		{
			for (var nElm = 0, nElmsCount = oDocument.GetElementsCount(); nElm < nElmsCount; nElm ++)
				sOutputText += this.HandleChildElement(oDocument.GetElement(nElm), 'html');
		}

		// рендер html тагов
		if (!this.Config.renderHTMLTags) {
			sOutputText = sOutputText.replace(/</gi, '&lt;').replace(/>/gi, '&gt;');
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
		}
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
					sOutputText = oCMarkdownConverter.WrapInSymbol(sOutputText, '* ', 'open');
				else
				{
					var oNumInfo = oPara.Paragraph.SavedNumberingValues ? oPara.Paragraph.SavedNumberingValues.NumInfo : null;
					if (!oNumInfo)
					{
						oNumInfo = oNumPr ? oPara.Paragraph.GetParent().CalculateNumberingValues(oPara.Paragraph, oNumPr, true) : null;
					}
					
					sOutputText = oCMarkdownConverter.WrapInSymbol(sOutputText, String(oNumInfo[0][0]) + '. ', 'open');
				}
					
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
		}

		function SetHeading(sOutputText)
		{
			// определяем уровень heading у параграфа для Markdown, Title и Subtitle тоже учитываем
			// далее выставляем # соответсвенно уровню.
			var oStyle = oDocument.Get_Styles().Get(oPara.Paragraph.Pr.PStyle);
			var nHeadingLvl;

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
					return oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Headings[Math.min(nHeadingLvl, oCMarkdownConverter.HtmlTags.Headings.length - 1)],'wholly');
				else if (oCMarkdownConverter.Config.convertType === 'markdown')
					return oCMarkdownConverter.WrapInSymbol(sOutputText, oCMarkdownConverter.MdSymbols.Headings[Math.min(nHeadingLvl, oCMarkdownConverter.MdSymbols.Headings.length - 1)] + ' ', 'open');
			}
		}
		function SetQuote()
		{
			if (oCMarkdownConverter.Config.convertType === 'html' || oCMarkdownConverter.isTableCellContent)
				return oCMarkdownConverter.WrapInTag(sOutputText, oCMarkdownConverter.HtmlTags.Quote, 'wholly');
			else if (oCMarkdownConverter.Config.convertType === 'markdown')
				return oCMarkdownConverter.WrapInSymbol(sOutputText, oCMarkdownConverter.MdSymbols.Quote, 'open');
		}
		function IsHeading(oParagraph)
		{
			var Styles         = private_GetLogicDocument().Get_Styles();
			var sParaStyleName = '';
			if (oParagraph.Paragraph.Pr.PStyle)
				sParaStyleName = Styles.Get(oParagraph.Paragraph.Pr.PStyle).Name;
			else
				return false;
			return sParaStyleName.search('Heading') !== -1 || sParaStyleName.search('Title') !== -1 || sParaStyleName.search('Subtitle') !== -1;

		}
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
			for (nElm = nElmsCount - 1, nElmsCount = oParagraph.GetElementsCount(); nElm >= 0; nElm--)
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

			return private_isMonospaceFont(sFirstRunFont) && private_isMonospaceFont(sLastRunFont);
		}
		function IsQuoteLine(oParagraph)
		{
			var Styles         = private_GetLogicDocument().Get_Styles();
			var sParaStyleId   = oParagraph.Paragraph.Get_CompiledPr2().ParaPr.GetPStyle();
			var sQuoteStyleId1 = Styles.GetStyleIdByName('Quote');
			var sQuoteStyleId2 = Styles.GetStyleIdByName('Intense Quote');

			return sParaStyleId === sQuoteStyleId1 || sParaStyleId === sQuoteStyleId2;
		}
		function HaveSepLine(oParagraph)
		{
			return oParagraph.Pr.Brd.Bottom && !oParagraph.Pr.Brd.Top && !oParagraph.Pr.Brd.Left && !oParagraph.Pr.Brd.Right
		}
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
		}

		if (!oPara.Next && oPara.Paragraph.IsEmpty())
		{
			if (HaveSepLine(oPara.Paragraph) && this.Config.convertType === "html")
				return "<hr>";
			return '';
		}
			
		this.isTableCellContent = oPara.Paragraph.IsTableCellContent();

		var oDocument  = private_GetLogicDocument();
		var sNumId     = null;
		var oNumPr     = null;
		if (!(oPara.Paragraph.Parent instanceof AscFormat.CDrawingDocContent) && oDocument instanceof AscCommonWord.CDocument)
			oNumPr           = oPara.Paragraph.GetNumPr();
		var oCMarkdownConverter = this;

		// если не будет нумерации, тогда проверим на заголовки (одновременно и то и другое в конвертации не применяется)
		if (oNumPr)
		{
			sNumId = oNumPr.NumId;
			this.isNumbering = true;
			this.isCodeBlock = false;
			this.isHeading   = false;
			this.isQuoteLine = IsQuoteLine(oPara);
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
		else
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
		else if ((this.Config.convertType === "html" && !this.isTableCellContent) && !this.isNumbering && !this.isHeading && !this.isCodeBlock)
			sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Paragraph, 'wholly');
		if (HaveSepLine(oPara.Paragraph) && this.Config.convertType === "html")
			sOutputText += "\n<hr>"

		// Add \n\n for correct parsing
		return sOutputText + '\n\n';
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
		}
		function IsBold(oRun)
		{
			if (!oRun)
				return false;

			var oRunTextPr = oRun.Run.Get_CompiledPr();

			return !!oRunTextPr.Bold
		}
		function IsItalic(oRun)
		{
			if (!oRun)
				return false;

			var oRunTextPr = oRun.Run.Get_CompiledPr();

			return !!oRunTextPr.Italic;
		}
		function isUnderline(oRun)
		{
			if (!oRun)
				return false;

			var oRunTextPr = oRun.Run.Get_CompiledPr();

			return !!oRunTextPr.Underline;
		}
		function isEqualTxPr(oRun1, oRun2)
		{
			if (!oRun2)
				return false;

			var oTextPr1 = oRun1.Run.Get_CompiledPr();
			var oTextPr2 = oRun2.Run.Get_CompiledPr();
			var sVertAlg1 = GetVertAlign(oRun1);
			var sVertAlg2 = GetVertAlign(oRun2);

			if (oTextPr1.Bold === oTextPr2.Bold && oTextPr1.Italic === oTextPr2.Italic && oTextPr1.Strikeout === oTextPr2.Strikeout)
				return !(this.Config.convertType === "html" && (oTextPr1.Underline !== oTextPr2.Underline || sVertAlg1 !== sVertAlg2));

			return false;
		}
		function isStrikeout(oRun)
		{
			if (!oRun)
				return false;

			var oRunTextPr = oRun.Run.Get_CompiledPr();

			return !!oRunTextPr.Strikeout;
		}
		function GetVertAlign(oRun)
		{
			if (!oRun)
				return "";

			var oRunTextPr = oRun.Run.Get_CompiledPr();
			if (oRunTextPr.VertAlign === 1)
				return "sup";
			else if (oRunTextPr.VertAlign === 2)
				return "sub";

			return "";
		}
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
								sText += oCMarkdownConverter.Config.base64img ? '<img src="' + Item.GraphicObj.getBase64Img() + '" alt="img">' : '<img src="' + Item.GraphicObj.getImageUrl() + '" alt="img">';
						}
						break;
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
					case para_NewLine:
						if (!this.isHeading)
						{
							if (this.Config.convertType === "html")
								sText += "<br>"
							else
								sText += " \\\n";
						}
						break;
					case para_Space:
					{
						sText += " ";
						break;
					}
					case para_Tab:
					{
						sText += "	";
						break;
					}
				}
			}

			return sText;
		}
		function GetText(oRun)
		{
			var sText = "";
			var ContentLen = oRun.Content.length;

			for (var CurPos = 0; CurPos < ContentLen; CurPos++)
			{
				var Item     = oRun.Content[CurPos];
				var ItemType = Item.Type;

				switch (ItemType)
				{
					case para_Drawing:
					case para_PageNum:
					case para_PageCount:
					case para_End:
						break;
					case para_Text :
					{
						sText += String.fromCharCode(Item.Value);
						break;
					}
					case para_NewLine:
						if (!this.isHeading)
						{
							if (this.Config.convertType === "html")
								sText += "<br>"
							else
								sText += " \\\n";
						}
						break;
					case para_Space:
					{
						sText += " ";
						break;
					}
					case para_Tab:
					{
						sText += "	";
						break;
					}
				}

			}

			return sText;
		}

		var oCMarkdownConverter    = this;
		var arrAllDrawings = oRun.Run.GetAllDrawingObjects();
		var hasPicture     = false;
		var sOutputText    = GetText.call(this, oRun.Run);
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

		if (sOutputText.trim() === "" && hasPicture === false)
			return sOutputText;

		if (!this.isCodeBlock)
		{
			var oRunNext, oRunPrev;
			// возможно текст в ране представляет собой блок кода, обрабатываем это
			if (private_isMonospaceFont(oTextPr.FontFamily.Name))
			{
				oRunNext     = oRun.private_GetNextInParent();
				while (oRunNext && oRunNext.Run.GetText().trim() === '')
					oRunNext = oRunNext.private_GetNextInParent();

				oRunPrev     = oRun.private_GetPreviousInParent();
				while (oRunPrev && oRunPrev.Run.GetText().trim() === '')
					oRunPrev = oRunPrev.private_GetPreviousInParent();

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
				oRunNext     = oRun.private_GetNextInParent();
				while (oRunNext && oRunNext.Run.GetText().trim() === '')
					oRunNext = oRunNext.private_GetNextInParent();

				oRunPrev     = oRun.private_GetPreviousInParent();
				while (oRunPrev && oRunPrev.Run.GetText().trim() === '')
					oRunPrev = oRunPrev.private_GetPreviousInParent();

				var isBoldNextRun   = IsBold(oRunNext);
				var isBoldPrevRun   = IsBold(oRunPrev);
				var isItalicNextRun = IsItalic(oRunNext);
				var isItalicPrevRun = IsItalic(oRunPrev);
				var isUnderlineNextRun = isUnderline(oRunNext);
				var isUnderlinePrevRun = isUnderline(oRunPrev);
				var isStrikeoutNextRun = isStrikeout(oRunNext);
				var isStrikeoutPrevRun = isStrikeout(oRunPrev);
				var sVertAlgnNextRun = GetVertAlign(oRunNext);
				var sVertAlgnPrevRun = GetVertAlign(oRunPrev);
				var sVertAlgn = GetVertAlign(oRun);
				

				if (hasPicture)
					sOutputText = GetTextWithPicture.call(this, oRun);

				if (sVertAlgn)
					sType = 'html';
					
				if (oTextPr.Strikeout)
				{
					if (sType === 'html' || this.Config.htmlHeadings)
					{
						if (!isStrikeoutPrevRun && !isStrikeoutNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Strikeout, 'wholly');
						else if (!isStrikeoutPrevRun || !isEqualTxPr.call(this, oRun, oRunPrev))
						{
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Strikeout, 'open');
							if (!isStrikeoutNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
								sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Strikeout, 'close');
						}
						else if (!isStrikeoutNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Strikeout, 'close');
					}
					else if (sType === 'markdown')
					{
						if (!isStrikeoutPrevRun && !isStrikeoutNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Strikeout, 'wholly');
						else if (!isStrikeoutPrevRun || !isEqualTxPr.call(this, oRun, oRunPrev))
						{
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Strikeout, 'open');
							if (!isStrikeoutNextRun || !isEqualTxPr.call(this, oRun, oRunNext) || sVertAlgnNextRun)
								sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Strikeout, 'close');
						}
						else if (!isStrikeoutNextRun || !isEqualTxPr.call(this, oRun, oRunNext) || sVertAlgnNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Strikeout, 'close');
					}
				}
				if (oTextPr.Bold)
				{
					if (sType === 'html' || this.Config.htmlHeadings)
					{
						if (!isBoldPrevRun && !isBoldNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Bold, 'wholly');
						else if (!isBoldPrevRun || !isEqualTxPr.call(this, oRun, oRunPrev))
						{
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Bold, 'open');
							if (!isBoldNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
								sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Bold, 'close');
						}
						else if (!isBoldNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Bold, 'close');
					}
					else if (sType === 'markdown')
					{
						if (!isBoldPrevRun && !isBoldNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Bold, 'wholly');
						else if (!isBoldPrevRun || !isEqualTxPr.call(this, oRun, oRunPrev))
						{
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Bold, 'open');
							if (!isBoldNextRun || !isEqualTxPr.call(this, oRun, oRunNext) || sVertAlgnNextRun)
								sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Bold, 'close');
						}
						else if (!isBoldNextRun || !isEqualTxPr.call(this, oRun, oRunNext) || sVertAlgnNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Bold, 'close');
					}
				}

				if (oTextPr.Italic && !this.isQuoteLine)
				{
					if (sType === 'html' || this.Config.htmlHeadings)
					{
						if (!isItalicPrevRun && !isItalicNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Italic, 'wholly');
						else if (!isItalicPrevRun || !isEqualTxPr.call(this, oRun, oRunPrev))
						{
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Italic, 'open');
							if (!isItalicNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
								sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Italic, 'close');
						}
						else if (!isItalicNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Italic, 'close');
					}
					else if (sType === 'markdown')
					{
						if (!isItalicPrevRun && !isItalicNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Italic, 'wholly');
						else if (!isItalicPrevRun || !isEqualTxPr.call(this, oRun, oRunPrev))
						{
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Italic, 'open');
							if (!isItalicNextRun || !isEqualTxPr.call(this, oRun, oRunNext) || sVertAlgnNextRun)
								sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Italic, 'close');
						}
						else if (!isItalicNextRun || !isEqualTxPr.call(this, oRun, oRunNext) || sVertAlgnNextRun)
							sOutputText = this.WrapInSymbol(sOutputText, this.MdSymbols.Italic, 'close');
					}
				}
				if (oTextPr.Underline && !this.isQuoteLine)
				{
					if (sType === 'html' || this.Config.htmlHeadings)
					{
						if (!isUnderlinePrevRun && !isUnderlineNextRun)
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Span, 'wholly', 'text-decoration:underline');
						else if (!isUnderlinePrevRun || !isEqualTxPr.call(this, oRun, oRunPrev))
						{
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Span, 'open', 'text-decoration:underline');
							if (!isUnderlineNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
								sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Span, 'close');
						}
						else if (!isUnderlineNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
							sOutputText = this.WrapInTag(sOutputText, this.HtmlTags.Span, 'close');
					}
				}
				
				
				if (sVertAlgn && !this.isQuoteLine)
				{
					if (!sVertAlgnNextRun && !sVertAlgnPrevRun)
						sOutputText = this.WrapInTag(sOutputText, sVertAlgn === "sup" ? this.HtmlTags.SupScript : this.HtmlTags.SubScript, 'wholly');
					else if (!sVertAlgnPrevRun || !isEqualTxPr.call(this, oRun, oRunNext))
					{
						sOutputText = this.WrapInTag(sOutputText, sVertAlgn === "sup" ? this.HtmlTags.SupScript : this.HtmlTags.SubScript, 'open');
						if (!isUnderlineNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
						{
							sOutputText = this.WrapInTag(sOutputText, sVertAlgn === "sup" ? this.HtmlTags.SupScript : this.HtmlTags.SubScript, 'close');
						}
					}
					else if (!sVertAlgnNextRun || !isEqualTxPr.call(this, oRun, oRunNext))
						sOutputText = this.WrapInTag(sOutputText, sVertAlgn === "sup" ? this.HtmlTags.SupScript : this.HtmlTags.SubScript, 'close');
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
			sOutputText += this.HandleChildElement(oTable.GetRow(nRow), this.Config.convertType);
		}

		// Add \n\n for correct parsing
		sOutputText += '</table>\n\n';
		return sOutputText;
	};
	CMarkdownConverter.prototype.HandleTableRow = function(oTableRow)
	{
		var sOutputText = '  <tr>\n';

		for (var nCell = 0, nCellsCount = oTableRow.GetCellsCount(); nCell < nCellsCount; nCell++)
		{
			sOutputText += this.HandleChildElement(oTableRow.GetCell(nCell), this.Config.convertType);
		}

		sOutputText += '  </tr>\n';
		return sOutputText;
	};
	CMarkdownConverter.prototype.HandleTableCell = function(oTableCell)
	{
		// Add 'th' for the firs row
		let symbol = (oTableCell.GetRowIndex() === 0) ? 'td' : 'td';
		// Add \n\n for correct parsing
		var sOutputText = '   <' + symbol + '>\n\n';
		var apiCellContent = oTableCell.GetContent();

		for (var nElm = 0, nElmsCount = apiCellContent.GetElementsCount(); nElm < nElmsCount; nElm++)
		{
			sOutputText += this.HandleChildElement(apiCellContent.GetElement(nElm), this.Config.convertType);
		}

		sOutputText += '</' + symbol + '>\n';
		return sOutputText;
	};
	/**
	 * Class representing a continuous region in a document. 
	 * Each Range object is determined by the position of the start and end characters.
	 * @param oElement - The document element that may be Document, Table, Paragraph, Run or Hyperlink.
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
	}

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
		}

		if (typeof(Start) == "number" && typeof(End) == "number" && Start > End)
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
		var DocPos, DocPosInRun;

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
					DocPosInRun =
					{
						Class : oRun,
						Position : nPos,
					};
		
					DocPos = oRun.GetDocumentPositionFromObject();
		
					DocPos.push(DocPosInRun);
		
					StartPos = DocPos;

					isStartDocPosFinded = true;
				}
				
				if (EndChar - charsCount === nRangePos - 1 && !isEndDocPosFinded)
				{
					DocPosInRun =
					{
						Class : oRun,
						Position : nPos + 1,
					};
		
					DocPos = oRun.GetDocumentPositionFromObject();
		
					DocPos.push(DocPosInRun);
		
					EndPos = DocPos;

					isEndDocPosFinded = true;
				}
			}

			if (nRangePos !== 0)
				charsCount += nRangePos;
		}

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
					if (firstDocPos[nPos].Class !== secondDocPos[nPos].Class || firstDocPos[nPos].Position !== secondDocPos[nPos].Position)
						return false;
				}
				return true;
			}
			return false;
		}

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
	 * Returns a type of the ApiRange class.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @returns {"range"}
	 */
	ApiRange.prototype.GetClassType = function()
	{
		return "range";
	};

	/**
	 * Returns a paragraph from all the paragraphs that are in the range.
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
	 * Adds a text to the specified position.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {String} sText - The text that will be added.
	 * @param {string} [sPosition = "after"] - The position where the text will be added ("before" or "after" the range specified).
	 * @return {boolean} - returns false if range is empty or sText isn't string value.
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
	 * Adds a bookmark to the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {String} sName - The bookmark name.
	 * @return {boolean} - returns false if range is empty.
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
	 * Adds a hyperlink to the specified range. 
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

		if (!AscCommon.rx_allowedProtocols.test(sLink))
			sLink = (urlType === 0) ? null :(( (urlType === 2) ? 'mailto:' : 'http://' ) + sLink);

		sLink = sLink && sLink.replace(new RegExp("%20",'g')," ");
		hyperlinkPr.put_Value(sLink);
		hyperlinkPr.put_ToolTip(sScreenTipText);
		hyperlinkPr.put_Bookmark(null);

		this.Select(false);
		var oHyperlink = new ApiHyperlink(this.Paragraphs[0].Paragraph.AddHyperlink(hyperlinkPr));
		Document.RemoveSelection();

		return oHyperlink;
	};

	/**
	 * Returns a text from the specified range.
	 * @memberof ApiRange
	 * @param {object} oPr - The resulting string display properties.
     * @param {boolean} [oPr.NewLineParagraph=false] - Defines if the resulting string will include paragraph line boundaries or not.
     * @param {boolean} [oPr.Numbering=false] - Defines if the resulting string will include numbering or not.
     * @param {boolean} [oPr.Math=false] - Defines if the resulting string will include mathematical expressions or not.
	 * @param {string} [oPr.NewLineSeparator='\r'] - Defines how the line separator will be specified in the resulting string.
     * @param {string} [oPr.TableCellSeparator='\t'] - Defines how the table cell separator will be specified in the resulting string.
     * @param {string} [oPr.TableRowSeparator='\r\n'] - Defines how the table row separator will be specified in the resulting string.
     * @param {string} [oPr.ParaSeparator='\r\n'] - Defines how the paragraph separator will be specified in the resulting string.
	 * @param {string} [oPr.TabSymbol='\t'] - Defines how the tab will be specified in the resulting string (does not apply to numbering)
	 * @typeofeditors ["CDE"]
	 * @returns {String} - returns "" if range is empty.
	 */
	ApiRange.prototype.GetText = function(oPr)
	{
		if (!oPr) {
			oPr = {};
		}
		
		let oProp = {
			NewLineSeparator:	(oPr.hasOwnProperty("NewLineSeparator")) ? oPr["NewLineSeparator"] : "\r",
			NewLineParagraph:	(oPr.hasOwnProperty("NewLineParagraph")) ? oPr["NewLineParagraph"] : true,
			Numbering:			(oPr.hasOwnProperty("Numbering")) ? oPr["Numbering"] : true,
			Math:				(oPr.hasOwnProperty("Math")) ? oPr["Math"] : true,
			TableCellSeparator:	oPr["TableCellSeparator"],
			TableRowSeparator:	oPr["TableRowSeparator"],
			ParaSeparator:		oPr["ParaSeparator"],
			TabSymbol:			oPr["TabSymbol"]
		}

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

		var Text = this.Controller.GetSelectedText(false, oProp); 
		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return Text;
	};

	/**
	 * Returns a collection of paragraphs that represents all the paragraphs in the specified range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @return {ApiParagraph[]}
	 */	
	ApiRange.prototype.GetAllParagraphs = function()
	{
		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		if (this.isEmpty || this.isEmpty === undefined)
			return [];

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
	 * Sets the selection to the specified range.
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
			var controllerType;

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
	 * Returns a new range that goes beyond the specified range in any direction and spans a different range. The current range has not changed. Throws an error if two ranges do not have a union.
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
		
		var newRangeStartPos;
		var newRangeEndPos;

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
	 * Returns a new range as the intersection of the current range with another range. The current range has not changed. Throws an error if two ranges do not overlap or are not adjacent.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {ApiRange} oRange - The range that will be intersected with the current range.
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
	 * Sets the bold property to the text character.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isBold - Specifies if the Range contents are displayed bold or not.
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
	 * Specifies that any lowercase characters in the current text Range are formatted for display only as their capital letter character equivalents.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isCaps - Specifies if the Range contents are displayed capitalized or not.
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
	 * Sets the text color to the current text Range in the RGB format.
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

		var ParaTextPr;
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
	 * Specifies that the contents of the current Range are displayed with two horizontal lines through each character displayed on the line.
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
	 * Specifies a highlighting color which is applied as a background to the contents of the current Range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {highlightColor} sColor - Available highlight color.
	 * @returns {ApiRange | null} - returns null if can't apply highlight.
	 */
	ApiRange.prototype.SetHighlight = function(sColor)
	{
		var color = private_getHighlightColorByName(sColor);
		if (!color && sColor !== "none")
			return null;

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

		var TextPr;
		if ("none" === sColor)
		{
			TextPr = new ParaTextPr({HighLight : highlight_None});
			Document.AddToParagraph(TextPr);
		}
		else
		{
			color = new CDocumentColor(color.r, color.g, color.b);
			TextPr = new ParaTextPr({HighLight : color});
			Document.AddToParagraph(TextPr);
		}

		this.TextPr.Merge(TextPr.Value);

		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Specifies the shading applied to the contents of the current text Range.
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

		let oShd = private_GetShd(sType, r, g, b, false);

		var SelectedContent = Document.GetSelectedElementsInfo({CheckAllSelection : true});
		if (!SelectedContent.CanEditBlockSdts() || !SelectedContent.CanDeleteInlineSdts())
		{
			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();

			return null;
		}

		Document.SetParagraphShd(oShd);
		this.TextPr.Shd = oShd;
		
		Document.LoadDocumentState(oldSelectionInfo);
		Document.UpdateSelection();

		return this;
	};

	/**
	 * Sets the italic property to the text character.
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
	 * Specifies that the contents of the current Range are displayed with a single horizontal line through the range center.
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
	 * Specifies that all the lowercase letter characters in the current text Range are formatted for display only as their capital
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
	 * Sets the text spacing measured in twentieths of a point.
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
	 * Specifies that the contents of the current Range are displayed along with a line appearing directly below the character
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
	 * Specifies the alignment which will be applied to the Range contents in relation to the default appearance of the Range text:
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
			value = AscCommon.vertalign_Baseline;
		else if (sType === "subscript")
			value = AscCommon.vertalign_SubScript;
		else if (sType === "superscript")
			value = AscCommon.vertalign_SuperScript;
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
	 * Specifies the amount by which text is raised or lowered for the current Range in relation to the default
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
	 * Sets the font size to the characters of the current text Range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {hps} FontSize - The text size value measured in half-points (1/144 of an inch).
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
	 * Sets all 4 font slots with the specified font family.
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

		var oThis               = this;
		var loader				= AscCommon.g_font_loader;
		var fontinfo			= g_fontApplication.GetFontInfo(sFontFamily);
		var isasync				= loader.LoadFont(fontinfo, setFontFamily);
		var Document			= private_GetLogicDocument();
		var oldSelectionInfo	= undefined;

		if (isasync === false)
			setFontFamily()

		function setFontFamily()
		{
			oldSelectionInfo = Document.SaveDocumentState();

			oThis.Select(false);
			if (oThis.isEmpty || oThis.isEmpty === undefined)
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
			
			oThis.TextPr.Merge(ParaTextPr.Value);

			Document.LoadDocumentState(oldSelectionInfo);
			Document.UpdateSelection();
		}

		return this;
	};

	/**
	 * Sets the style to the current Range.
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
	 * Sets the text properties to the current Range.
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
	 * Deletes all the contents from the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @returns {boolean} - returns false if range is empty.
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
	 * Converts the ApiRange object into the JSON object.
	 * @memberof ApiRange
	 * @param {boolean} bWriteNumberings - Specifies if the used numberings will be written to the JSON object or not.
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiRange.prototype.ToJSON = function(bWriteNumberings, bWriteStyles)
	{
		var oDocument = private_GetLogicDocument();
		var aContent = [];
		var oJSON = {
			"type":    "document",
			"content": []
		}

		private_RefreshRangesPosition();
		private_RemoveEmptyRanges();

		var oldSelectionInfo = oDocument.SaveDocumentState();
		this.Select(false);

		var oSelectedContent = oDocument.GetSelectedContent();

		if (this.isEmpty || this.isEmpty === undefined)
		{
			oDocument.LoadDocumentState(oldSelectionInfo);
			return "";
		}
		private_TrackRangesPositions();

		oDocument.LoadDocumentState(oldSelectionInfo);
		oDocument.UpdateSelection();

		for (var nElm = 0; nElm < oSelectedContent.Elements.length; nElm++)
			aContent.push(oSelectedContent.Elements[nElm].Element)

		if (aContent.length > 0)
		{
			var oWriter = new AscJsonConverter.WriterToJSON();
			oJSON["content"] = oWriter.SerContent(aContent);
		}
		else
			return "";

		// numbering и styles в конце, потому что сначала нужно обойти все параграфы
		if (bWriteNumberings)
			oJSON["numbering"] = oWriter.jsonWordNumberings;
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();

		return JSON.stringify(oJSON);
	};

	/**
	 * Adds a comment to the current range.
	 * @memberof ApiRange
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment text (required).
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {?ApiComment} - Returns null if the comment was not added.
	 */
	ApiRange.prototype.AddComment = function(sText, sAuthor)
	{
		let oDocument = private_GetLogicDocument();

		if (typeof(sText) !== "string" || sText.trim() === "")
			return null;
		if (typeof(sAuthor) !== "string")
			sAuthor = "";
		
		var CommentData = new AscCommon.CCommentData();
		CommentData.SetText(sText);
		if (sAuthor !== "")
			CommentData.SetUserName(sAuthor);

		var documentState = oDocument.SaveDocumentState();
		this.Select();

		let comment = AddCommentToDocument(oDocument, CommentData);
		oDocument.LoadDocumentState(documentState);
		oDocument.UpdateSelection();
		return comment;
	};

	/**
     * Returns a Range object that represents the document part contained in the specified range.
     * @typeofeditors ["CDE"]
     * @param {Number} [Start=0] - Start character index in the current range.
     * @param {Number} [End=-1] - End character index in the current range (if <= 0, then the range is taken to the end).
     * @returns {ApiRange}
     * */
	ApiRange.prototype.GetRange = function(nStart, nEnd)
	{
		if (typeof(nStart) != "number" || nStart < 0)
			nStart = 0;
		if (typeof(nEnd) != "number" || nEnd < 0)
			nEnd = -1;

		if (nEnd > 0 && nStart > nEnd)
			[nStart, nEnd] = [nEnd, nStart];

		let curStartPos	= this.StartPos;
		let curEndPos	= this.EndPos;

		let newStartPos	= null;
		let newEndPos	= nEnd == -1 ? curEndPos : nEnd;

		let isStartDocPosFinded	= false;
		let isEndDocPosFinded	= nEnd == -1 ? true : false;
		
		let nCharsNewRange = nEnd === -1 ? -1 : nEnd - nStart + 1; // кол-во символов в новом Range
		let tempCharsCount = 0; // все символы по которым прошлись

		let isSearchBegin = false; // нашли стартовый ран
		let isStartRun = false;	// начинаем считать символы со Start позиции текущего Range (только в стартовом Run).

		let isOutOfRange = false;

		function checkRelateDocPos(firstPos, secondPos)
		{
			for (var nPos = 0, nLen = Math.min(firstPos.length, secondPos.length); nPos < nLen; ++nPos)
			{
				if (!secondPos[nPos] || !firstPos[nPos] || firstPos[nPos].Class !== secondPos[nPos].Class)
					return -1;

				if (firstPos[nPos].Position < secondPos[nPos].Position)
					return -1;
				else if (firstPos[nPos].Position > secondPos[nPos].Position)
					return 1;
			}

			return 0;
		}

		function processRun(oRun)
		{
			if (isOutOfRange)
				return;

			var nRangePos = 0;

			if (isSearchBegin == false && checkRelateDocPos(oRun.GetDocumentPositionFromObject(), curStartPos) === 0) {
				isStartRun = true;
				isSearchBegin = true;
			}
			else if (isSearchBegin == false)
				return;
			
			var nLenght = oRun.Content.length;
			var nPos = isStartRun == true ? curStartPos[curStartPos.length - 1].Position : 0;

			for (; nPos < nLenght; ++nPos)
			{
				if (para_Text === oRun.Content[nPos].Type || para_Space === oRun.Content[nPos].Type || para_Tab === oRun.Content[nPos].Type) {
					nRangePos++;
					tempCharsCount++;
				}
					
				if (isStartDocPosFinded == false && nStart === tempCharsCount - 1)
				{
					let DocPosInRun =
					{
						Class : oRun,
						Position : nPos,
					};
		
					let DocPos = oRun.GetDocumentPositionFromObject();
		
					DocPos.push(DocPosInRun);
		
					newStartPos = DocPos;

					isStartDocPosFinded = true;
					tempCharsCount = 1; // теперь считаем кол-во символов от стартового в Range. Конец Range будет через nCharsNewRange символов.
				}
				
				if (isStartDocPosFinded == true && isEndDocPosFinded == false && nCharsNewRange == tempCharsCount)
				{
					let DocPosInRun =
					{
						Class : oRun,
						Position : nPos + 1,
					};
		
					let DocPos = oRun.GetDocumentPositionFromObject();
					DocPos.push(DocPosInRun);
		
					newEndPos = DocPos;
					isEndDocPosFinded = true;
				}
			}

			isStartRun = false;
		}
		
		if (this.Element instanceof CDocument || this.Element instanceof CDocumentContent || this.Element instanceof CTable || this.Element instanceof CBlockLevelSdt)
		{
			var allParagraphs = this.Element.GetAllParagraphs({OnlyMainDocument : true, All : true});

			for (var paraItem = 0; paraItem < allParagraphs.length; paraItem++)
			{
				if (isOutOfRange)
					return null;

				if (isStartDocPosFinded && isEndDocPosFinded)
					break;
				else
					allParagraphs[paraItem].CheckRunContent(processRun);
			}
		}
		else if (this.Element instanceof Paragraph || this.Element instanceof ParaHyperlink || this.Element instanceof CInlineLevelSdt || this.Element instanceof ParaRun)
		{
			this.Element.CheckRunContent(processRun);
		}
		
		let oRange = new ApiRange(this.Element, newStartPos, newEndPos);
		if (oRange.isEmpty)
			return null;
		
		return oRange;
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
	 * Class representing a comment.
	 * @constructor
	 */
	function ApiComment(oComment)
	{
		this.Comment = oComment;
	}

	/**
	 * Class representing a comment reply.
	 * @constructor
	 */
	function ApiCommentReply(oParentComm, oCommentReply)
	{
		this.Comment = oParentComm;
		this.Data = oCommentReply;
	}

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
	 * Returns a type of the ApiHyperlink class.
	 * @memberof ApiHyperlink
	 * @typeofeditors ["CDE"]
	 * @returns {"hyperlink"}
	 */
	ApiHyperlink.prototype.GetClassType = function()
	{
		return "hyperlink";
	};
	
	/**
	 * Class representing a document form base.
	 * @constructor
	 * @typeofeditors ["CDE", "CFE"]
	 * @property {string} key - Form key.
	 * @property {string} tip - Form tip text.
	 * @property {boolean} required - Specifies if the form is required or not.
	 * @property {string} placeholder - Form placeholder text.
	 */
	function ApiFormBase(oSdt)
	{
		this.Sdt = oSdt;
	}
	
	/**
	 * Class representing a document text field.
	 * @constructor
	 * @typeofeditors ["CDE", "CFE"]
	 * @property {boolean} comb - Specifies if the text field should be a comb of characters with the same cell width. The maximum number of characters must be set to a positive value.
	 * @property {number} maxCharacters - The maximum number of characters in the text field.
	 * @property {number} cellWidth - The cell width for each character measured in millimeters. If this parameter is not specified or equal to 0 or less, then the width will be set automatically.
	 * @property {boolean} multiLine - Specifies if the current fixed size text field is multiline or not.
	 * @property {boolean} autoFit - Specifies if the text field content should be autofit, i.e. whether the font size adjusts to the size of the fixed size form.
	 * @extends {ApiFormBase}
	 */
	function ApiTextForm(oSdt)
	{
		ApiFormBase.call(this, oSdt);
	}

	ApiTextForm.prototype = Object.create(ApiFormBase.prototype);
	ApiTextForm.prototype.constructor = ApiTextForm;

	/**
	 * Class representing a document combo box / dropdown list.
	 * @constructor
	 * @typeofeditors ["CDE", "CFE"]
	 * @property {boolean} editable - Specifies if the combo box text can be edited.
	 * @property {boolean} autoFit - Specifies if the combo box form content should be autofit, i.e. whether the font size adjusts to the size of the fixed size form.
	 * @property {Array.<string | Array.<string>>} items - The combo box items.
     * This array consists of strings or arrays of two strings where the first string is the displayed value and the second one is its meaning.
     * If the array consists of single strings, then the displayed value and its meaning are the same.
     * Example: ["First", ["Second", "2"], ["Third", "3"], "Fourth"].
	 * @extends {ApiFormBase}
	 */
	function ApiComboBoxForm(oSdt)
	{
		ApiFormBase.call(this, oSdt);
	}

	ApiComboBoxForm.prototype = Object.create(ApiFormBase.prototype);
	ApiComboBoxForm.prototype.constructor = ApiComboBoxForm;

	/**
	 * Class representing a document checkbox / radio button.
	 * @constructor
	 * @typeofeditors ["CDE", "CFE"]
	 * @property {boolean} radio - Specifies if the current checkbox is a radio button. In this case, the key parameter is considered as an identifier for the group of radio buttons.
	 * @extends {ApiFormBase}
	 */
	function ApiCheckBoxForm(oSdt)
	{
		ApiFormBase.call(this, oSdt);
	}

	ApiCheckBoxForm.prototype = Object.create(ApiFormBase.prototype);
	ApiCheckBoxForm.prototype.constructor = ApiCheckBoxForm;

	/**
	 * Class representing a document picture form.
	 * @constructor
	 * @typeofeditors ["CDE", "CFE"]
	 * @property {ScaleFlag} scaleFlag - The condition to scale an image in the picture form: "always", "never", "tooBig" or "tooSmall".
	 * @property {boolean} lockAspectRatio - Specifies if the aspect ratio of the picture form is locked or not.
	 * @property {boolean} respectBorders - Specifies if the form border width is respected or not when scaling the image.
	 * @property {percentage} shiftX - Horizontal picture position inside the picture form measured in percent:
	 * * <b>0</b> - the picture is placed on the left;
	 * * <b>50</b> - the picture is placed in the center;
	 * * <b>100</b> - the picture is placed on the right.
	 * @property {percentage} shiftY - Vertical picture position inside the picture form measured in percent:
	 * * <b>0</b> - the picture is placed on top;
	 * * <b>50</b> - the picture is placed in the center;
	 * * <b>100</b> - the picture is placed on the bottom.
	 * @extends {ApiFormBase}
	 */
	function ApiPictureForm(oSdt)
	{
		ApiFormBase.call(this, oSdt);
	}

	ApiPictureForm.prototype = Object.create(ApiFormBase.prototype);
	ApiPictureForm.prototype.constructor = ApiPictureForm;

	/**
	 * Class representing a complex field.
	 * @param oSdt
	 * @constructor
	 * @typeofeditors ["CDE", "CFE"]
	 * @extends {ApiFormBase}
	 */
	function ApiComplexForm(oSdt)
	{
		ApiFormBase.call(this, oSdt);
	}
	ApiComplexForm.prototype = Object.create(ApiFormBase.prototype);
	ApiComplexForm.prototype.constructor = ApiComplexForm;

	/**
	 * Sets the hyperlink address.
	 * @typeofeditors ["CDE"]
	 * @param {string} sLink - The hyperlink address.
	 * @returns {boolean}
	 * */
	ApiHyperlink.prototype.SetLink = function(sLink)
	{
		if (typeof(sLink) !== "string" || sLink.length > Asc.c_nMaxHyperlinkLength)
			return false;

		var urlType	= undefined;

		if (sLink !== "")
		{
			urlType		= AscCommon.getUrlType(sLink);
			if (!AscCommon.rx_allowedProtocols.test(sLink))
				sLink = (urlType === 0) ? null :(( (urlType === 2) ? 'mailto:' : 'http://' ) + sLink);
			sLink = sLink && sLink.replace(new RegExp("%20",'g')," ");
		}
		
		this.ParaHyperlink.SetValue(sLink);
		
		return true;
	};
	/**
	 * Sets the hyperlink display text.
	 * @typeofeditors ["CDE"]
	 * @param {string} sDisplay - The text to display the hyperlink.
	 * @returns {boolean}
	 * */
	ApiHyperlink.prototype.SetDisplayedText = function(sDisplay)
	{
		if (typeof(sDisplay) !== "string")
			return false;

		var HyperRun;
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
	 * Sets the screen tip text of the hyperlink.
	 * @typeofeditors ["CDE"]
	 * @param {string} sScreenTipText - The screen tip text of the hyperlink.
	 * @returns {boolean}
	 * */
	ApiHyperlink.prototype.SetScreenTipText = function(sScreenTipText)
	{
		if (typeof(sScreenTipText) !== "string")
			return false;

		this.ParaHyperlink.SetToolTip(sScreenTipText);
		
		return true;
	};
	/**
	 * Returns the hyperlink address.
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
	 * Returns the hyperlink display text.
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
	 * Returns the screen tip text of the hyperlink.
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
	 * Returns the hyperlink element using the position specified.
	 * @typeofeditors ["CDE"]
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
	 * Returns a number of elements in the current hyperlink.
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiHyperlink.prototype.GetElementsCount = function()
	{
		return this.ParaHyperlink.GetElementsCount();
	};
	/**
	 * Sets the default hyperlink style.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
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

			HyperRun.SetUnderline(undefined);
			HyperRun.Set_Color(undefined);
			HyperRun.Set_Unifill(undefined);
			HyperRun.Set_RStyle(Styles.GetDefaultHyperlink());
		}
			
		return true;
	};
	/**
	 * Returns a Range object that represents the document part contained in the specified hyperlink.
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character index in the current element.
	 * @param {Number} End - End character index in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiHyperlink.prototype.GetRange = function(Start, End)
	{
		return new ApiRange(this.ParaHyperlink, Start, End);
	};
	/**
	 * Converts the ApiHyperlink object into the JSON object.
	 * @memberof ApiHyperlink
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiHyperlink.prototype.ToJSON = function(bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerHyperlink(this.ParaHyperlink);
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();

		return JSON.stringify(oJSON);
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
	 * Class representing an Ole-object.
	 * @constructor
	 */
	function ApiOleObject(OleObject)
	{
		ApiDrawing.call(this, OleObject.parent);
		this.OleObject = OleObject
	}
	ApiOleObject.prototype = Object.create(ApiDrawing.prototype);
	ApiOleObject.prototype.constructor = ApiOleObject;

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
	 * Class representing an RGB Color.
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
	 * Class representing a base class for fill.
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
	 * Class representing a list values of combobox/dropdown list content control.
	 * @constructor
	 */
	function ApiContentControlList(Parent)
	{
		this.Sdt    = Parent.Sdt;
		this.Parent = Parent;
	}

	/**
	 * Class representing an entry of a list values of combobox/dropdown list content control.
	 * @constructor
	 */
	function ApiContentControlListEntry(Sdt, Parent, Text, Value)
	{
		this.Sdt    = Sdt;
		this.Parent = Parent;
		this.Text   = Text;
		this.Value  = Value;
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
     * * <b>"none"</b> - no border will be added to the created element or the selected element side.
     * * <b>"single"</b> - a single border will be added to the created element or the selected element side.
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
     * * <b>"default"</b> - a header or footer which can be applied to any default page.
     * * <b>"title"</b> - a header or footer which is applied to the title page.
     * * <b>"even"</b> - a header or footer which can be applied to even pages to distinguish them from the odd ones (which will be considered default).
	 * @typedef {("default" | "title" | "even")} HdrFtrType
	 */

	/**
	 * The possible values for the units of the width property are defined by a specific table or table cell width property.
     * * <b>"auto"</b> - sets the table or table cell width to auto width.
     * * <b>"twips"</b> - sets the table or table cell width to be measured in twentieths of a point.
     * * <b>"nul"</b> - sets the table or table cell width to be of a zero value.
     * * <b>"percent"</b> - sets the table or table cell width to be measured in percent to the parent container.
	 * @typedef {("auto" | "twips" | "nul" | "percent")} TableWidth
	 */

	/**
	 * This simple type specifies possible values for the table sections to which the current conditional formatting properties will be applied when this selected table style is used.
	 * * <b>"topLeftCell"</b> - specifies that the table formatting is applied to the top left cell.
	 * * <b>"topRightCell"</b> - specifies that the table formatting is applied to the top right cell.
	 * * <b>"bottomLeftCell"</b> - specifies that the table formatting is applied to the bottom left cell.
	 * * <b>"bottomRightCell"</b> - specifies that the table formatting is applied to the bottom right cell.
	 * * <b>"firstRow"</b> - specifies that the table formatting is applied to the first row.
	 * * <b>"lastRow"</b> - specifies that the table formatting is applied to the last row.
	 * * <b>"firstColumn"</b> - specifies that the table formatting is applied to the first column. Any subsequent row which is in *table header* ({@link ApiTableRowPr#SetTableHeader}) will also use this conditional format.
	 * * <b>"lastColumn"</b> - specifies that the table formatting is applied to the last column.
	 * * <b>"bandedColumn"</b> - specifies that the table formatting is applied to odd numbered groupings of rows.
	 * * <b>"bandedColumnEven"</b> - specifies that the table formatting is applied to even numbered groupings of rows.
	 * * <b>"bandedRow"</b> - specifies that the table formatting is applied to odd numbered groupings of columns.
	 * * <b>"bandedRowEven"</b> - specifies that the table formatting is applied to even numbered groupings of columns.
	 * * <b>"wholeTable"</b> - specifies that the conditional formatting is applied to the whole table.
	 * @typedef {("topLeftCell" | "topRightCell" | "bottomLeftCell" | "bottomRightCell" | "firstRow" | "lastRow" |
	 *     "firstColumn" | "lastColumn" | "bandedColumn" | "bandedColumnEven" | "bandedRow" | "bandedRowEven" |
	 *     "wholeTable")} TableStyleOverrideType
	 */

	/**
	 * The types of elements that can be added to the paragraph structure.
	 * @typedef {(ApiUnsupported | ApiRun | ApiInlineLvlSdt | ApiHyperlink | ApiFormBase)} ParagraphContent
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
     * * <b>"none"</b> - not display the selected tick labels.
     * * <b>"nextTo"</b> - sets the position of the selected tick labels next to the main label.
     * * <b>"low"</b> - sets the position of the selected tick labels in the part of the chart with lower values.
     * * <b>"high"</b> - sets the position of the selected tick labels in the part of the chart with higher values.
	 * @typedef {("none" | "nextTo" | "low" | "high")} TickLabelPosition
	 * **/

	/**
     * The type of a fill which uses an image as a background.
     * * <b>"tile"</b> - if the image is smaller than the shape which is filled, the image will be tiled all over the created shape surface.
     * * <b>"stretch"</b> - if the image is smaller than the shape which is filled, the image will be stretched to fit the created shape surface.
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
	 * The lock type of the content control.
	 * @typedef {"unlocked" | "contentLocked" | "sdtContentLocked" | "sdtLocked"} SdtLock
	 */

	/**
     * Text transform type.
	 * @typedef {("textArchDown" | "textArchDownPour" | "textArchUp" | "textArchUpPour" | "textButton" | "textButtonPour" | "textCanDown"
	 * | "textCanUp" | "textCascadeDown" | "textCascadeUp" | "textChevron" | "textChevronInverted" | "textCircle" | "textCirclePour"
	 * | "textCurveDown" | "textCurveUp" | "textDeflate" | "textDeflateBottom" | "textDeflateInflate" | "textDeflateInflateDeflate" | "textDeflateTop"
	 * | "textDoubleWave1" | "textFadeDown" | "textFadeLeft" | "textFadeRight" | "textFadeUp" | "textInflate" | "textInflateBottom" | "textInflateTop"
	 * | "textPlain" | "textRingInside" | "textRingOutside" | "textSlantDown" | "textSlantUp" | "textStop" | "textTriangle" | "textTriangleInverted"
	 * | "textWave1" | "textWave2" | "textWave4" | "textNoShape")} TextTransform
	 * */

	/**
	 * Form type.
	 * The available form types.
	 * @typedef {"textForm" | "comboBoxForm" | "dropDownForm" | "checkBoxForm" | "radioButtonForm" | "pictureForm"} FormType
	 */

	/**
	 * 1 millimetre equals 1/10th of a centimetre.
	 * @typedef {number} mm
	 */

	/**
	 * The condition to scale an image in the picture form.
	 * @typedef {"always" | "never" | "tooBig" | "tooSmall"} ScaleFlag
	 */

	/**
	 * Value from 0 to 100.
	 * @typedef {number} percentage
	 */

	/**
	 * Available highlight colors.
	 * @typedef {"black" | "blue" | "cyan" | "green" | "magenta" | "red" | "yellow" | "white" | "darkBlue" |
	 * "darkCyan" | "darkGreen" | "darkMagenta" | "darkRed" | "darkYellow" | "darkGray" | "lightGray" | "none"} highlightColor
	 */

	//------------------------------------------------------------------------------------------------------------------
	//
	// Cross-reference
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Available values of the "numbered" reference type:
	 * * <b>"pageNum"</b> - the numbered item page number;
     * * <b>"paraNum"</b> - the numbered item paragraph number;
	 * * <b>"noCtxParaNum"</b> - the abbreviated paragraph number (the specific item only, e.g. instead of "4.1.1" you refer to "1" only);
     * * <b>"fullCtxParaNum"</b> - the full paragraph number, e.g. "4.1.1";
	 * * <b>"text"</b> - the paragraph text value, e.g. if you have "4.1.1. Terms and Conditions", you refer to "Terms and Conditions" only;
     * * <b>"aboveBelow"</b> - the words "above" or "below" depending on the item position.
	 * @typedef {"pageNum" | "paraNum" | "noCtxParaNum" | "fullCtxParaNum" | "text" | "aboveBelow"} numberedRefTo
	 */

	/**
	 * Available values of the "heading" reference type:
	 * * <b>"text"</b> - the entire heading text;
	 * * <b>"pageNum"</b> - the heading page number;
     * * <b>"headingNum"</b> - the heading sequence number;
	 * * <b>"noCtxHeadingNum"</b> - the abbreviated heading number. Make sure the cursor pointer is in the section you are referencing to, e.g. you are in section 4 and you wish to refer to heading 4.B, so instead of "4.B" you receive "B" only;
     * * <b>"fullCtxHeadingNum"</b> - the full heading number even if the cursor pointer is in the same section;
     * * <b>"aboveBelow"</b> - the words "above" or "below" depending on the item position.
	 * @typedef {"text" | "pageNum" | "headingNum" | "noCtxHeadingNum" | "fullCtxHeadingNum" | "aboveBelow"} headingRefTo
	 */

	/**
	 * Available values of the "bookmark" reference type:
	 * * <b>"text"</b> - the entire bookmark text;
	 * * <b>"pageNum"</b> - the bookmark page number;
     * * <b>"paraNum"</b> - the bookmark paragraph number;
	 * * <b>"noCtxParaNum"</b> - the abbreviated paragraph number (the specific item only, e.g. instead of "4.1.1" you refer to "1" only);
     * * <b>"fullCtxParaNum</b> - the full paragraph number, e.g. "4.1.1";
     * * <b>"aboveBelow"</b> - the words "above" or "below" depending on the item position.
	 * @typedef {"text" | "pageNum" | "paraNum" | "noCtxParaNum" | "fullCtxParaNum" | "aboveBelow"} bookmarkRefTo
	 */

	/**
	 * Available values of the "footnote" reference type:
	 * * <b>"footnoteNum"</b> - the footnote number;
	 * * <b>"pageNum"</b> - the page number of the footnote;
     * * <b>"aboveBelow"</b> - the words "above" or "below" depending on the position of the item;
	 * * <b>"formFootnoteNum"</b> - the form number formatted as a footnote. The numbering of the actual footnotes is not affected.
	 * @typedef {"footnoteNum" | "pageNum" | "aboveBelow" | "formFootnoteNum"} footnoteRefTo
	 */

	/**
	 * Available values of the "endnote" reference type:
	 * * <b>"endnoteNum"</b> - the endnote number;
	 * * <b>"pageNum"</b> - the endnote page number;
     * * <b>"aboveBelow"</b> - the words "above" or "below" depending on the item position;
	 * * <b>"formEndnoteNum"</b> - the form number formatted as an endnote. The numbering of the actual endnotes is not affected.
	 * @typedef {"endnoteNum" | "pageNum" | "aboveBelow" | "formEndnoteNum"} endnoteRefTo
	 */

	/**
	 * Available values of the "equation"/"figure"/"table" reference type:
	 * * <b>"entireCaption"</b>- the entire caption text;
	 * * <b>"labelNumber"</b> - the label and object number only, e.g. "Table 1.1";
     * * <b>"captionText"</b> - the caption text only;
	 * * <b>"pageNum"</b> - the page number containing the referenced object;
	 * * <b>"aboveBelow"</b> - the words "above" or "below" depending on the item position.
	 * @typedef {"entireCaption" | "labelNumber" | "captionText" | "pageNum" | "aboveBelow"} captionRefTo
	 */

	//------------------------------------------------------End Cross-reference types--------------------------------------------------

	/**
	 * Axis position in the chart.
	 * @typedef {("top" | "bottom" | "right" | "left")} AxisPos
	 */

	/**
	 * Standard numeric format.
	 * @typedef {("General" | "0" | "0.00" | "#,##0" | "#,##0.00" | "0%" | "0.00%" |
	 * "0.00E+00" | "# ?/?" | "# ??/??" | "m/d/yyyy" | "d-mmm-yy" | "d-mmm" | "mmm-yy" | "h:mm AM/PM" |
	 * "h:mm:ss AM/PM" | "h:mm" | "h:mm:ss" | "m/d/yyyy h:mm" | "#,##0_);(#,##0)" | "#,##0_);[Red](#,##0)" | 
	 * "#,##0.00_);(#,##0.00)" | "#,##0.00_);[Red](#,##0.00)" | "mm:ss" | "[h]:mm:ss" | "mm:ss.0" | "##0.0E+0" | "@")} NumFormat
	 */

	/**
	 * Types of all supported forms.
	 * @typedef {ApiTextForm | ApiComboBoxForm | ApiCheckBoxForm | ApiPictureForm | ApiComplexForm} ApiForm
	 */

	/**
     * Possible values for the caption numbering format.
     * * <b>"ALPHABETIC"</b> - upper letter.
     * * <b>"alphabetic"</b> - lower letter.
     * * <b>"Roman"</b> - upper Roman.
     * * <b>"roman"</b> - lower Roman.
	 * * <b>"Arabic"</b> - arabic.
	 * @typedef {("ALPHABETIC" | "alphabetic" | "Roman" | "roman" | "Arabic")} CaptionNumberingFormat
	 * **/

	/**
     * Possible values for the caption separator.
     * * <b>"hyphen"</b> - the "-" punctuation mark.
     * * <b>"period"</b> - the "." punctuation mark.
     * * <b>"colon"</b> - the ":" punctuation mark.
     * * <b>"longDash"</b> - the "—" punctuation mark.
	 * * <b>"dash"</b> - the "-" punctuation mark.
	 * @typedef {("hyphen" | "period" | "colon" | "longDash" | "dash")} CaptionSep
	 * **/

	/**
     * Possible values for the caption label.
     * @typedef {("Table" | "Equation" | "Figure")} CaptionLabel
	 * **/

	/**
	 * Table of contents properties.
	 * @typedef {Object} TocPr
	 * @property {boolean} [ShowPageNums=true] - Specifies whether to show page numbers in the table of contents.
	 * @property {boolean} [RightAlgn=true] - Specifies whether to right-align page numbers in the table of contents.
	 * @property {TocLeader} [LeaderType="dot"] - The leader type in the table of contents.
	 * @property {boolean} [FormatAsLinks=true] - Specifies whether to format the table of contents as links.
	 * @property {TocBuildFromPr} [BuildFrom={OutlineLvls=9}] - Specifies whether to generate the table of contents from the outline levels or the specified styles.
	 * @property {TocStyle} [TocStyle="standard"] - The table of contents style type.
	 */

	/**
	 * Table of figures properties.
	 * @typedef {Object} TofPr
	 * @property {boolean} [ShowPageNums=true] - Specifies whether to show page numbers in the table of figures.
	 * @property {boolean} [RightAlgn=true] - Specifies whether to right-align page numbers in the table of figures.
	 * @property {TocLeader} [LeaderType="dot"] - The leader type in the table of figures.
	 * @property {boolean} [FormatAsLinks=true] - Specifies whether to format the table of figures as links.
	 * @property {CaptionLabel | string} [BuildFrom="Figure"] - Specifies whether to generate the table of figures based on the specified caption label or the paragraph style name used (for example, "Heading 1").
	 * @property {boolean} [LabelNumber=true] - Specifies whether to include the label and number in the table of figures.
	 * @property {TofStyle} [TofStyle="distinctive"] - The table of figures style type.
	 */

	/**
	 * Table of contents properties which specify whether to generate the table of contents from the outline levels or the specified styles.
	 * @typedef {Object} TocBuildFromPr
	 * @property {number} [OutlineLvls=9] - Maximum number of levels in the table of contents.
	 * @property {TocStyleLvl[]} StylesLvls - Style levels (for example, [{Name: "Heading 1", Lvl: 2}, {Name: "Heading 2", Lvl: 3}]).
	 * <note>If StylesLvls.length > 0, then the OutlineLvls property will be ignored.</note>
	 */

	/**
	 * Table of contents style levels.
	 * @typedef {Object} TocStyleLvl
	 * @property {string} Name - Style name (for example, "Heading 1").
	 * @property {number} Lvl - Level which will be applied to the specified style in the table of contents.
	 */

	/**
	 * Possible values for the table of contents leader:
	 * * <b>"dot"</b> - "......."
	 * * <b>"dash"</b> - "-------"
	 * * <b>"underline"</b> - "_______"
     * @typedef {("dot" | "dash" | "underline" | "none")} TocLeader
	 * **/

	/**
     * Possible values for the table of contents style.
     * @typedef {("simple" | "online" | "standard" | "modern" | "classic")} TocStyle
	 * **/

	/**
     * Possible values for the table of figures style.
     * @typedef {("simple" | "online" | "classic" | "distinctive" | "centered" | "formal")} TofStyle
	 * **/

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
	 * The type of tick mark appearance.
	 * @typedef {("cross" | "in" | "none" | "out")} TickMark
	 * */

	/**
	 * Returns the main document.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @returns {ApiDocument}
	 */
	Api.prototype.GetDocument = function()
	{
		return new ApiDocument(this.WordControl.m_oLogicDocument);
	};
	/**
	 * Creates a new paragraph.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {ApiParagraph}
	 */
	Api.prototype.CreateParagraph = function()
	{
		return new ApiParagraph(new Paragraph(private_GetDrawingDocument(), private_GetLogicDocument()));
	};
	/**
	 * Creates an element range.
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
	 * Creates a new table with a specified number of rows and columns.
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
	 * Creates a new smaller text block to be inserted to the current paragraph or table.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiRun}
	 */
	Api.prototype.CreateRun = function()
	{
		return new ApiRun(new ParaRun(null, false));
	};
	/**
	 * Creates a new hyperlink text block to be inserted to the current paragraph or table.
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
	 * Creates an image with the parameters specified.
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
	 * Creates a shape with the parameters specified.
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
	 * Creates a chart with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {ChartType} [sType="bar"] - The chart type used for the chart display.
	 * @param {Array} aSeries - The array of the data used to build the chart from.
	 * @param {Array} aSeriesNames - The array of the names (the source table column names) used for the data which the chart will be build from.
	 * @param {Array} aCatNames - The array of the names (the source table row names) used for the data which the chart will be build from.
	 * @param {EMU} nWidth - The chart width in English measure units.
	 * @param {EMU} nHeight - The chart height in English measure units.
	 * @param {number} nStyleIndex - The chart color style index (can be 1 - 48, as described in OOXML specification).
	 * @param {NumFormat[] | String[]} aNumFormats - Numeric formats which will be applied to the series (can be custom formats).
     * The default numeric format is "General".
	 * @returns {ApiChart}
	 * */
	Api.prototype.CreateChart = function(sType, aSeries, aSeriesNames, aCatNames, nWidth, nHeight, nStyleIndex, aNumFormats)
	{
		var oDrawingDocument = private_GetDrawingDocument();
		var nW = private_EMU2MM(nWidth);
		var nH = private_EMU2MM(nHeight);
		var oDrawing = new ParaDrawing( nW, nH, null, oDrawingDocument, null, null);
		var oChartSpace = AscFormat.builder_CreateChart(nW, nH, sType, aCatNames, aSeriesNames, aSeries, nStyleIndex, aNumFormats);
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
	 * Creates an OLE object with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {string} sImageSrc - The image source where the image to be inserted should be taken from (currently, only internet URL or Base64 encoded images are supported).
	 * @param {EMU} nWidth - The OLE object width in English measure units.
	 * @param {EMU} nHeight - The OLE object height in English measure units.
	 * @param {string} sData - The OLE object string data.
	 * @param {string} sAppId - The application ID associated with the current OLE object.
	 * @returns {ApiOleObject}
	 */
	Api.prototype.CreateOleObject = function(sImageSrc, nWidth, nHeight, sData, sAppId)
	{
		if (typeof sImageSrc === "string" && sImageSrc.length > 0 && typeof sData === "string"
			&& typeof sAppId === "string" && sAppId.length > 0
			&& AscFormat.isRealNumber(nWidth) && AscFormat.isRealNumber(nHeight)
		)

		var nW = private_EMU2MM(nWidth);
		var nH = private_EMU2MM(nHeight);

		var oDrawing = new ParaDrawing(nW, nH, null, private_GetDrawingDocument(), private_GetLogicDocument(), null);
		var oImage = private_GetLogicDocument().DrawingObjects.createOleObject(sData, sAppId, sImageSrc, 0, 0, nW, nH);
		oImage.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oImage);
		return new ApiOleObject(oImage);
	};

	/**
	 * Creates an RGB color setting the appropriate values for the red, green and blue color components.
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
	 * Creates a complex color scheme selecting from one of the available schemes.
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
	 * Creates a color selecting it from one of the available color presets.
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
	 * Creates a solid fill to apply to the object using a selected solid color as the object background.
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
	 * Creates a linear gradient fill to apply to the object using the selected linear gradient as the object background.
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
	 * Creates a radial gradient fill to apply to the object using the selected radial gradient as the object background.
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
	 * Creates a pattern fill to apply to the object using the selected pattern as the object background.
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
	 * Creates a blip fill to apply to the object using the selected image as the object background.
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
	 * Creates no fill and removes the fill from the element.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiFill}
	 * */
	Api.prototype.CreateNoFill = function()
	{
		return new ApiFill(AscFormat.CreateNoFillUniFill());
	};

	/**
	 * Creates a stroke adding shadows to the element.
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
	 * Creates a gradient stop used for different types of gradients.
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
	 * Creates a bullet for a paragraph with the character or symbol specified with the sSymbol parameter.
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
	 * Creates a bullet for a paragraph with the numbering character or symbol specified with the sType parameter.
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
	 * Creates a new inline container.
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
	 * Creates a new block level container.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @returns {ApiBlockLvlSdt}
	 */
	Api.prototype.CreateBlockLvlSdt = function()
	{
		return new ApiBlockLvlSdt(new CBlockLevelSdt(editor.private_GetLogicDocument(), private_GetLogicDocument()));
	};

	/**
	 * Saves changes to the specified document.
	 * @typeofeditors ["CDE"]
	 * @memberof Api
	 */
	Api.prototype.Save = function () {
		this.SaveAfterMacros = true;
	};

	/**
	 * Loads data for the mail merge. 
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {String[][]} aList - Mail merge data. The first element of the array is the array with names of the merge fields.
	 * The rest of the array elements are arrays with values for the merge fields.
	 * @typeofeditors ["CDE"]
	 * @return {boolean}
	 */
	Api.prototype.LoadMailMergeData = function(aList)
	{
		if (!aList || aList.length === 0)
			return false;

		editor.asc_StartMailMergeByList(aList);

		return true;
	};

	/**
	 * Returns the mail merge template document.
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
	 * Returns the mail merge receptions count.
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
	 * Replaces the main document content with another document content.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {ApiDocumentContent} oApiDocumentContent - The document content which the main document content will be replaced with.
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
	 * Starts the mail merge process.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {number} [nStartIndex=0] - The start index of the document for mail merge process.
	 * @param {number} [nEndIndex=Api.GetMailMergeReceptionsCount() - 1] - The end index of the document for mail merge process.
	 * @returns {boolean}
	 */
	Api.prototype.MailMerge = function(nStartIndex, nEndIndex)
	{
		var oDocument = editor.private_GetLogicDocument();

		var _nStartIndex = (undefined !== nStartIndex ? Math.max(0, nStartIndex) : 0);
		var _nEndIndex   = (undefined !== nEndIndex   ? Math.min(nEndIndex, oDocument.MailMergeMap.length - 1) : oDocument.MailMergeMap.length - 1);

		var mailMergeDoc = oDocument.Get_MailMergedDocument(_nStartIndex, _nEndIndex);

		if (!mailMergeDoc)
			return false;
		
		this.ReplaceDocumentContent(new ApiDocumentContent(mailMergeDoc));

		return true;
	};

	/**
	 * Converts the specified JSON object into the Document Builder object of the corresponding type.
	 * @memberof Api
	 * @param {JSON} sMessage - The JSON object to convert.
	 * @typeofeditors ["CDE"]
	 */
	Api.prototype.FromJSON = function(sMessage)
	{
		var oReader = new AscJsonConverter.ReaderFromJSON();
		AscJsonConverter.ActiveReader = oReader;

		var oDocument = this.GetDocument();
		var oParsedObj  = JSON.parse(sMessage);
		var oResult = null;
		if (oParsedObj["styles"])
			oReader.StylesFromJSON(oParsedObj["styles"]);

		switch (oParsedObj["type"])
		{
			case "document":
				if (oParsedObj["numbering"])
				{
					oReader.parsedNumbering = oParsedObj["numbering"];
				}
				if (oParsedObj["textPr"])
				{
					var oNewTextPr = oReader.TextPrFromJSON(oParsedObj["textPr"]);
					var oDefaultTextPr = oDocument.GetDefaultTextPr();

					oDefaultTextPr.TextPr = oNewTextPr;
					oDefaultTextPr.private_OnChange();
				}
				if (oParsedObj["paraPr"])
				{
					var oNewParaPr = oReader.ParaPrFromJSON(oParsedObj["paraPr"]);
					var oDefaultParaPr = oDocument.GetDefaultParaPr();

					oDefaultParaPr.ParaPr.Set_FromObject(oNewParaPr);
					oDefaultParaPr.private_OnChange();
				}
				if (oParsedObj["theme"])
				{
					var oDefaultTheme = oDocument.Document.GetTheme();

					oDefaultTheme.setColorScheme(oReader.ClrSchemeFromJSON(oParsedObj["theme"]["themeElements"]["clrScheme"]));
					oDefaultTheme.setFontScheme(oReader.FontSchemeFromJSON(oParsedObj["theme"]["themeElements"]["fontScheme"]));
					oDefaultTheme.setFormatScheme(oReader.FmtSchemeFromJSON(oParsedObj["theme"]["themeElements"]["fmtScheme"]));

					oDefaultTheme.setLnDef(oParsedObj["theme"]["objectDefaults"]["lnDef"] ? oReader.DefSpDefinitionFromJSON(oParsedObj["theme"]["objectDefaults"]["lnDef"]) : null);
					oDefaultTheme.setSpDef(oParsedObj["theme"]["objectDefaults"]["spDef"] ? oReader.DefSpDefinitionFromJSON(oParsedObj["theme"]["objectDefaults"]["spDef"]) : null);
					oDefaultTheme.setTxDef(oParsedObj["theme"]["objectDefaults"]["txDef"] ? oReader.DefSpDefinitionFromJSON(oParsedObj["theme"]["objectDefaults"]["txDef"]) : null);

					oDefaultTheme.setName(oParsedObj["theme"]);
					oDefaultTheme.setIsThemeOverride(oParsedObj["theme"]["isThemeOverride"]);
				}
				if (oParsedObj["sectPr"])
				{
					var oNewSectPr = oReader.SectPrFromJSON(oParsedObj["sectPr"]);
					var oCurSectPr = oDocument.Document.SectPr;

					// borders
					oCurSectPr.SetBordersOffsetFrom(oNewSectPr.Borders.OffsetFrom);
					oCurSectPr.Set_Borders_Display(oNewSectPr.Borders.Display);
					oCurSectPr.Set_Borders_ZOrder(oNewSectPr.Borders.ZOrder);

					oCurSectPr.Set_Borders_Bottom(oNewSectPr.Borders.Bottom ? oNewSectPr.Borders.Bottom.Copy() : null);
					oCurSectPr.Set_Borders_Left(oNewSectPr.Borders.Left ? oNewSectPr.Borders.Left.Copy() : null);
					oCurSectPr.Set_Borders_Right(oNewSectPr.Borders.Right ? oNewSectPr.Borders.Right.Copy() : null);
					oCurSectPr.Set_Borders_Top(oNewSectPr.Borders.Top ? oNewSectPr.Borders.Top.Copy() : null);

					// columns
					oCurSectPr.Set_Columns_Cols(oNewSectPr.Columns.Cols);
					oCurSectPr.Set_Columns_EqualWidth(oNewSectPr.Columns.EqualWidth);
					oCurSectPr.Set_Columns_Num(oNewSectPr.Columns.Num);
					oCurSectPr.Set_Columns_Sep(oNewSectPr.Columns.Sep);
					oCurSectPr.Set_Columns_Space(oNewSectPr.Columns.Space);

					// endnotePr
					oCurSectPr.SetEndnoteNumFormat(oNewSectPr.EndnotePr.NumFormat);
					oCurSectPr.SetEndnoteNumRestart(oNewSectPr.EndnotePr.NumRestart);
					oCurSectPr.SetEndnoteNumStart(oNewSectPr.EndnotePr.NumStart);
					oCurSectPr.SetEndnotePos(oNewSectPr.EndnotePr.Pos);

					// footnotePr
					oCurSectPr.SetFootnoteNumFormat(oNewSectPr.FootnotePr.NumFormat);
					oCurSectPr.SetFootnoteNumRestart(oNewSectPr.FootnotePr.NumRestart);
					oCurSectPr.SetFootnoteNumStart(oNewSectPr.FootnotePr.NumStart);
					oCurSectPr.SetFootnotePos(oNewSectPr.FootnotePr.Pos);

					// pageMargins
					oCurSectPr.SetGutter(oNewSectPr.PageMargins.Gutter);
					oCurSectPr.SetPageMargins(oNewSectPr.PageMargins.Left, oNewSectPr.PageMargins.Top, oNewSectPr.PageMargins.Right, oNewSectPr.PageMargins.Bottom);
					oCurSectPr.SetPageMarginFooter(oNewSectPr.PageMargins.Footer);
					oCurSectPr.SetPageMarginHeader(oNewSectPr.PageMargins.Header);

					// pageSize
					oCurSectPr.SetPageSize(oNewSectPr.PageSize.W, oNewSectPr.PageSize.H);
					oCurSectPr.SetOrientation(oNewSectPr.PageSize.Orient);

					// footer
					oCurSectPr.Set_Footer_Default(oNewSectPr.FooterDefault);
					oCurSectPr.Set_Footer_Even(oNewSectPr.FooterEven);
					oCurSectPr.Set_Footer_First(oNewSectPr.FooterFirst);

					// header
					oCurSectPr.Set_Header_Default(oNewSectPr.HeaderDefault);
					oCurSectPr.Set_Header_Even(oNewSectPr.HeaderEven);
					oCurSectPr.Set_Header_First(oNewSectPr.HeaderFirst);

					// other
					oCurSectPr.Set_TitlePage(oNewSectPr.TitlePage);
					oCurSectPr.Set_Type(oNewSectPr.Type);
					oCurSectPr.Set_PageNum_Start(oNewSectPr.PageNumType.Start);

				}

				var aContent = oReader.ContentFromJSON(oParsedObj["content"]);
				for (var nElm = 0; nElm < aContent.length; nElm++)
				{
					if (aContent[nElm] instanceof AscCommonWord.Paragraph)
						aContent[nElm] = new ApiParagraph(aContent[nElm]);
					else if (aContent[nElm] instanceof AscCommonWord.CTable)
						aContent[nElm] = new ApiTable(aContent[nElm]);
					else if (aContent[nElm] instanceof AscCommonWord.CBlockLevelSdt)
						aContent[nElm] = new ApiBlockLvlSdt(aContent[nElm]);
				}

				oResult = aContent;
				break;
			case "docContent":
				oResult = new ApiDocumentContent(oReader.DocContentFromJSON(oParsedObj));
				break;
			case "drawingDocContent":
				oResult = new ApiDocumentContent(oReader.DrawingDocContentFromJSON(oParsedObj));
				break;
			case "paragraph":
				oResult = new ApiParagraph(oReader.ParagraphFromJSON(oParsedObj));
				break;
			case "run":
			case "mathRun":
			case "endRun":
				oResult = new ApiRun(oReader.ParaRunFromJSON(oParsedObj));
				break;
			case "hyperlink":
				oResult = new ApiHyperlink(oReader.HyperlinkFromJSON(oParsedObj));
				break;
			case "inlineLvlSdt":
				oResult = new ApiInlineLvlSdt(oReader.InlineLvlSdtFromJSON(oParsedObj));
				break;
			case "blockLvlSdt":
				oResult = new ApiBlockLvlSdt(oReader.BlockLvlSdtFromJSON(oParsedObj));
				break;
			case "table":
				oResult = new ApiTable(oReader.TableFromJSON(oParsedObj));
				break;
			case "paraDrawing":
				oResult = new ApiDrawing(oReader.ParaDrawingFromJSON(oParsedObj));
				break;
			case "nextPage":
			case "oddPage":
			case "evenPage":
			case "continuous":
			case "nextColumn":
				oResult = new ApiSection(oReader.SectPrFromJSON(oParsedObj));
				break;
			case "numbering":
				oReader.parsedNumbering = oParsedObj;
				let sNumId		= Object.keys(oParsedObj["num"])[0];
				let sAddedNumId	= oReader.NumberingFromJSON(sNumId);
				let oNum		= oDocument.Document.GetNumbering().GetNum(sAddedNumId);

				oResult = new ApiNumbering(oNum);
				break;
			case "textPr":
				oResult = oParsedObj["bFromDocument"] ? new ApiTextPr(null, oReader.TextPrFromJSON(oParsedObj)) : new ApiTextPr(null, oReader.TextPrDrawingFromJSON(oParsedObj));
				break;
			case "paraPr":
				oResult = oParsedObj["bFromDocument"] ? new ApiParaPr(null, oReader.ParaPrFromJSON(oParsedObj)) : new ApiParaPr(null, oReader.ParaPrDrawingFromJSON(oParsedObj));
				break;
			case "tablePr":
				oResult = new ApiTablePr(null, oReader.TablePrFromJSON(null, oParsedObj));
				break;
			case "tableRowPr":
				oResult = new ApiTableRowPr(null, oReader.TableRowPrFromJSON(oParsedObj));
				break;
			case "tableCellPr":
				oResult = new ApiTableCellPr(null, oReader.TableCellPrFromJSON(oParsedObj));
				break;
			case "fill":
				oResult = new ApiFill(oReader.FillFromJSON(oParsedObj));
				break;
			case "stroke":
				oResult = new ApiStroke(oReader.LnFromJSON(oParsedObj));
				break;
			case "gradStop":
				var oGs = oReader.GradStopFromJSON(oParsedObj);
				oResult = new ApiGradientStop(new ApiUniColor(oGs.color), oGs.pos);
				break;
			case "uniColor":
				oResult = new ApiUniColor(oReader.ColorFromJSON(oParsedObj));
				break;
			case "style":
				var oStyle = oReader.StyleFromJSON(oParsedObj);
				var oDocument = private_GetLogicDocument();
				if (oStyle)
				{
					var nExistingStyle = oDocument.Styles.GetStyleIdByName(oStyle.Name);
					// если такого стиля нет - добавляем новый
					if (nExistingStyle === null)
						oDocument.Styles.Add(oStyle);
					else
					{
						var oExistingStyle = oDocument.Styles.Get(nExistingStyle);
						// если стили идентичны, стиль не добавляем
						if (!oStyle.IsEqual(oExistingStyle))
						{
							oStyle.Set_Name("Custom_Style " + AscCommon.g_oIdCounter.Get_NewId());
							oDocument.Styles.Add(oStyle);
						}
						else
							oStyle = oExistingStyle;
					}
				}
				oResult = new ApiStyle(oStyle);
				break;
			case "tableStyle":
				oResult = new ApiTableStylePr(oParsedObj.styleType, null, oReader.TableStylePrFromJSON(oParsedObj));
				break;
		}

		oReader.AssignConnectedObjects();
		return oResult;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiUnsupported
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiUnsupported class.
	 * @typeofeditors ["CDE"]
	 * @returns {"unsupported"}
	 */
	ApiUnsupported.prototype.GetClassType = function()
	{
		return "unsupported";
	};
	/**
	 * Adds a comment to the specifed document element or array of Runs.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {ApiRun[] | DocumentElement} oElement - The element where the comment will be added. It may be applied to any element which has the *AddComment* method.
	 * @param {string} sText - The comment text (required).
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {ApiComment?} - Returns null if the comment was not added.
	 */
	Api.prototype.AddComment = function(oElement, sText, sAuthor)
	{
		if (!sText || typeof(sText) !== "string")
			return null;
	
		if (typeof(sAuthor) !== "string")
			sAuthor = "";
		
		// Если oElement не является массивом, определяем параграф это или документ
		if (!Array.isArray(oElement))
		{
			if (oElement instanceof ApiParagraph || oElement instanceof ApiDocument || oElement instanceof ApiRange ||
				oElement instanceof ApiBlockLvlSdt || oElement instanceof ApiInlineLvlSdt || oElement instanceof ApiTable ||
				oElement instanceof ApiRun)
				{
					return oElement.AddComment(sText, sAuthor);
				}
		}
		// Проверка на массив с ранами
		else if (Array.isArray(oElement))
		{
			// Если хотя бы один элемент массива не является раном, или хотя бы один ран не принадлежит 
			// ни одному параграфу - не добавляем комментарий
			for (var Index = 0; Index < oElement.length; Index++)
			{
				if (!(oElement[Index] instanceof ApiRun))
					return null;					
			}
			
			// Если раны из принципиально разных контекcтов (из тела и хедера(или футера) то комментарий не добавляем)
			for (Index = 1; Index < oElement.length; Index++)
			{
				if (oElement[0].Run.GetDocumentPositionFromObject()[0].Class !== oElement[Index].Run.GetDocumentPositionFromObject()[0].Class)
					return null;
			}
			
			var oDocument = private_GetLogicDocument();

			var CommentData = new AscCommon.CCommentData();
			CommentData.SetText(sText);
			CommentData.SetUserName(sAuthor);

			var oStartRun = private_GetFirstRunInArray(oElement); 
			var oStartPos = oStartRun.Run.GetDocumentPositionFromObject();
			var oEndRun	= private_GetLastRunInArray(oElement)
			var oEndPos	= oEndRun.Run.GetDocumentPositionFromObject();

			oStartPos.push({Class: oStartRun.Run, Position: 0});
			oEndPos.push({Class: oEndRun.Run, Position: oEndRun.Run.Content.length});

			oDocument.Selection.Use = true;
			oDocument.SetContentSelection(oStartPos, oEndPos, 0, 0, 0);

			var sQuotedText = oDocument.GetSelectedText(false);
			CommentData.Set_QuoteText(sQuotedText);

			var oComment = new AscCommon.CComment(oDocument.Comments, CommentData);
			oComment.GenerateDurableId();
			oDocument.Comments.Add(oComment);
			oDocument.RemoveSelection();

			var oCommentEnd = new AscCommon.ParaComment(false, oComment.Get_Id());
			oComment.SetRangeEnd(oCommentEnd.GetId());
			var oEndRunParent = oEndRun.Run.GetParent();
			if (!oEndRunParent)
				return null;
			var nEndRunPosInParent = oEndRun.Run.GetPosInParent();
			oEndRunParent.Internal_Content_Add(nEndRunPosInParent + 1, oCommentEnd);

			var oCommentStart = new AscCommon.ParaComment(true, oComment.Get_Id());
			oComment.SetRangeStart(oCommentStart.GetId());
			var oStartRunParent = oStartRun.Run.GetParent();
			if (!oStartRunParent)
				return null;
			var nStartRunPosInParent = oStartRun.Run.GetPosInParent();
			oStartRunParent.Internal_Content_Add(nStartRunPosInParent, oCommentStart);

			if (null != oComment)
				editor.sync_AddComment(oComment.Get_Id(), CommentData);

			return new ApiComment(oComment);
		}

		return null;
	};

	/**
	 * Subscribes to the specified event and calls the callback function when the event fires.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {string} eventName - The event name.
	 * @param {function} callback - Function to be called when the event fires.
	 */
	Api.prototype["attachEvent"] = Api.prototype.attachEvent;

	/**
	 * Unsubscribes from the specified event.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {string} eventName - The event name.
	 */
	Api.prototype["detachEvent"] = Api.prototype.detachEvent;

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDocumentContent
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiDocumentContent class. 
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"documentContent"}
	 */
	ApiDocumentContent.prototype.GetClassType = function()
	{
		return "documentContent";
	};
	/**
	 * Returns a number of elements in the current document.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {number}
	 */
	ApiDocumentContent.prototype.GetElementsCount = function()
	{
		return this.Document.Content.length;
	};
	/**
	 * Returns an element by its position in the document.
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
	 * Adds a paragraph or a table or a blockLvl content control using its position in the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {number} nPos - The position where the current element will be added.
	 * @param {DocumentElement} oElement - The document element which will be added at the current position.
	 */
	ApiDocumentContent.prototype.AddElement = function(nPos, oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;
			this.Document.Internal_Content_Add(nPos, oElm);
		}
	};
	/**
	 * Pushes a paragraph or a table to actually add it to the document.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {DocumentElement} oElement - The element type which will be pushed to the document.
	 * @returns {boolean} - returns false if oElement is unsupported.
	 */
	ApiDocumentContent.prototype.Push = function(oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;

			this.Document.Internal_Content_Add(this.Document.Content.length, oElm);
			return true;
		}

		return false;
	};
	/**
	 * Removes all the elements from the current document or from the current document element.
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
	 * Removes an element using the position specified.
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
	 * Returns a Range object that represents the part of the document contained in the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiDocumentContent.prototype.GetRange = function(Start, End)
	{
		return new ApiRange(this.Document, Start, End);
	};
	/**
	 * Converts the ApiDocumentContent object into the JSON object.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteNumberings - Specifies if the used numberings will be written to the JSON object or not.
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiDocumentContent.prototype.ToJSON = function(bWriteNumberings, bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerDocContent(this.Document);
		if (bWriteNumberings)
			oJSON["numbering"] = oWriter.jsonWordNumberings;
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();

		return JSON.stringify(oJSON);
	};
	/**
	 * Returns an array of document elements from the current ApiDocumentContent object.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bGetCopies - Specifies if the copies of the document elements will be returned or not.
	 * @returns {Array}
	 */
	ApiDocumentContent.prototype.GetContent = function(bGetCopies)
	{
		var aContent = [];
		var oTempElm = null;

		for (var nElm = 0; nElm < this.Document.Content.length; nElm++)
		{
			oTempElm = this.Document.Content[nElm];

			if (bGetCopies)
				oTempElm = oTempElm.Copy();

			if (oTempElm instanceof AscCommonWord.CTable)
				aContent.push(new ApiTable(oTempElm));
			else if (oTempElm instanceof AscCommonWord.Paragraph)
				aContent.push(new ApiParagraph(oTempElm));
			else if (oTempElm instanceof AscCommonWord.CBlockLevelSdt)
				aContent.push(new ApiBlockLvlSdt(oTempElm));
		}

		return aContent;
	};
	/**
	 * Returns a collection of drawing objects from the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @return {ApiDrawing[]}  
	 */
	ApiDocumentContent.prototype.GetAllDrawingObjects = function()
	{
		var arrAllDrawing = this.Document.GetAllDrawingObjects();
		var arrApiShapes  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
			arrApiShapes.push(new ApiDrawing(arrAllDrawing[Index]));
		
		return arrApiShapes;
	};
	/**
	 * Returns a collection of shape objects from the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @return {ApiShape[]}  
	 */
	ApiDocumentContent.prototype.GetAllShapes = function()
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
	 * Returns a collection of image objects from the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @return {ApiImage[]}  
	 */
	ApiDocumentContent.prototype.GetAllImages = function()
	{
		var arrAllDrawing = this.Document.GetAllDrawingObjects();
		var arrApiImages  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof AscFormat.CImageShape)
				arrApiImages.push(new ApiImage(arrAllDrawing[Index].GraphicObj));
		}
		
		return arrApiImages;
	};
	/**
	 * Returns a collection of chart objects from the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @return {ApiChart[]}  
	 */
	ApiDocumentContent.prototype.GetAllCharts = function()
	{
		var arrAllDrawing = this.Document.GetAllDrawingObjects();
		var arrApiCharts  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof AscFormat.CChartSpace)
				arrApiCharts.push(new ApiChart(arrAllDrawing[Index].GraphicObj));
		}
		
		return arrApiCharts;
	};
	/**
	 * Returns a collection of OLE objects from the document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @return {ApiOleObject[]}  
	 */
	ApiDocumentContent.prototype.GetAllOleObjects = function()
	{
		var arrAllDrawing = this.Document.GetAllDrawingObjects();
		var arrApiOleObjects  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof AscFormat.COleObject)
				arrApiOleObjects.push(new ApiOleObject(arrAllDrawing[Index].GraphicObj));
		}
		
		return arrApiOleObjects;
	};
	/**
	 * Returns an array of all paragraphs from the current document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @return {ApiParagraph[]}
	 */
	ApiDocumentContent.prototype.GetAllParagraphs = function()
	{
		let result = [];
		this.Document.GetAllParagraphs().forEach(function(paragraph)
		{
			result.push(new ApiParagraph(paragraph));
		});
		return result;
	};
	/**
	 * Returns an array of all tables from the current document content.
	 * @memberof ApiDocumentContent
	 * @typeofeditors ["CDE"]
	 * @return {ApiParagraph[]}
	 */
	ApiDocumentContent.prototype.GetAllTables = function()
	{
		let result = [];
		this.Document.GetAllTables().forEach(function(table)
		{
			result.push(new ApiTable(table));
		});
		return result;
	};
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDocument
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiDocument class.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {"document"}
	 */
	ApiDocument.prototype.GetClassType = function()
	{
		return "document";
	};
	/**
	 * Creates a new history point.
	 * @memberof ApiDocument
	 */
	ApiDocument.prototype.CreateNewHistoryPoint = function()
	{
		this.Document.Create_NewHistoryPoint(AscDFH.historydescription_Document_ApiBuilder);
	};
	/**
	 * Returns a style by its name.
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
	 * Creates a new style with the specified type and name. If there is a style with the same name it will be replaced with a new one.
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
	 * Returns the default style parameters for the specified document element.
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
			return new ApiStyle(oStyles.Get(oStyles.GetDefaultCharacter()));
		else if ("numbering" === sStyleType)
			return new ApiStyle(oStyles.Get(oStyles.Get_Default_Numbering()));

		return null;
	};
	/**
	 * Returns a set of default properties for the text run in the current document.
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
	 * Returns a set of default paragraph properties in the current document.
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
	 * Returns the document final section.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @return {ApiSection}
	 */
	ApiDocument.prototype.GetFinalSection = function()
	{
		return new ApiSection(this.Document.SectPr);
	};
	/**
	 * Creates a new document section which ends at the specified paragraph. Allows to set local parameters to the current
	 * section - page size, footer, header, columns, etc.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {ApiParagraph} oParagraph - The paragraph after which a new document section will be inserted.
	 * Paragraph must be in a document.
	 * @returns {ApiSection | null}
	 */
	ApiDocument.prototype.CreateSection = function(oParagraph)
	{
		if (!(oParagraph instanceof ApiParagraph)) {
			console.error(new Error('Parameter is invalid.'));
			return null;
		}
		if (!oParagraph.Paragraph.CanAddSectionPr()) {
			console.error(new Error('Paragraph must be in a document.'));
			return null;
		}

		var oSectPr = new CSectionPr(this.Document);

		var nContentPos = oParagraph.Paragraph.GetIndex();
		var oCurSectPr = this.Document.SectionsInfo.Get_SectPr(nContentPos).SectPr;

		oSectPr.Copy(oCurSectPr);
		oCurSectPr.Set_Type(oSectPr.Type);
		oCurSectPr.Set_PageNum_Start(-1);
		oCurSectPr.Clear_AllHdrFtr();

		oParagraph.private_GetImpl().Set_SectionPr(oSectPr);
		return new ApiSection(oSectPr);
	};

	/**
	 * Specifies whether sections in this document will have different headers and footers for even and
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
	 * Creates an abstract multilevel numbering with a specified type.
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
	 * Inserts an array of elements into the current position of the document.
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
		var oSelectedContent = new AscCommonWord.CSelectedContent();
		var oElement;
		for (var nIndex = 0, nCount = arrContent.length; nIndex < nCount; ++nIndex)
		{
			oElement = arrContent[nIndex];
			var oElm;
			if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
			{
				oElm = oElement.private_GetImpl();
				if (oElm.IsUseInDocument())
					continue;

				if (oElm.Parent != null) {
					oElm.SetParent(private_GetLogicDocument());
				}

				oSelectedContent.Add(new AscCommonWord.CSelectedElement(oElm, true));
			}
		}
		oSelectedContent.EndCollect(this.Document);

		if (isInline)
			oSelectedContent.ForceInlineInsert();

		if (this.Document.IsSelectionUse())
		{
			this.Document.Start_SilentMode();
			this.Document.Remove(1, false, false, true);
			this.Document.End_SilentMode();
			this.Document.RemoveSelection(true);
		}

		var oParagraph = this.Document.GetCurrentParagraph(undefined, undefined, {CheckDocContent: true});
		if (!oParagraph)
			return false;

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

				for (nIndex = 0, nCount = oSelectedContent.Elements.length; nIndex < nCount; ++nIndex)
				{
					oElement = oSelectedContent.Elements[nIndex].Element;
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
		oSelectedContent.Insert(oNearestPos);
		oParagraph.Clear_NearestPosArray();
		return true;
	};
	
	/**
	 * Record of one comment.
	 * @typedef {Object} CommentReportRecord
	 * @property {boolean} [IsAnswer=false] - Specifies whether this is an initial comment or a reply to another comment.
	 * @property {string} CommentMessage - The text of the current comment.
	 * @property {number} Date - The time when this change was made in local time.
	 * @property {number} DateUTC - The time when this change was made in UTC.
	 * @property {string} [QuoteText=undefined] - The text to which this comment is related.
	 */
	
	/**
	 * Report on all comments.
	 * This is a dictionary where the keys are usernames.
	 * @typedef {Object.<string, Array.<CommentReportRecord>>} CommentReport
	 * @example
	 *  {
	 *    "John Smith" : [{IsAnswer: false, CommentMessage: 'Good text', Date: 1688588002698, DateUTC: 1688570002698, QuoteText: 'Some text'},
	 *      {IsAnswer: true, CommentMessage: "I don't think so", Date: 1688588012661, DateUTC: 1688570012661}],
	 *
	 *    "Mark Pottato" : [{IsAnswer: false, CommentMessage: 'Need to change this part', Date: 1688587967245, DateUTC: 1688569967245, QuoteText: 'The quick brown fox jumps over the lazy dog'},
	 *      {IsAnswer: false, CommentMessage: 'We need to add a link', Date: 1688587967245, DateUTC: 1688569967245, QuoteText: 'OnlyOffice'}]
	 *  }
	 */
	
	
	/**
	 * Returns a report about all the comments added to the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {CommentReport}
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
				var isAnswer     = !oReport[sUserName][nIndex].Top;
				var oCommentData = oReport[sUserName][nIndex].Data;
				var nDateUTC     = parseInt(oCommentData.m_sOOTime);
				if (isNaN(nDateUTC))
					nDateUTC = 0;

				if (isAnswer)
				{
					oResult[sUserName].push({
						"IsAnswer"       : true,
						"CommentMessage" : oCommentData.GetText(),
						"Date"           : oCommentData.GetDateTime(),
						"DateUTC"        : nDateUTC
					});
				}
				else
				{
					var sQuoteText = oCommentData.GetQuoteText();
					oResult[sUserName].push({
						"IsAnswer"       : false,
						"CommentMessage" : oCommentData.GetText(),
						"Date"           : oCommentData.GetDateTime(),
						"DateUTC"        : nDateUTC,
						"QuoteText"      : sQuoteText,
						"IsSolved"       : oCommentData.IsSolved()
					});
				}
			}
		}

		return oResult;
	};
	
	/**
	 * Review record type.
	 * @typedef {("TextAdd" | "TextRem" | "ParaAdd" | "ParaRem" | "TextPr" | "ParaPr" | "Unknown")} ReviewReportRecordType
	 */
	
	/**
	 * Record of one review change.
	 * @typedef {Object} ReviewReportRecord
	 * @property {ReviewReportRecordType} Type - Review record type.
	 * @property {string} [Value=undefined] - Review change value that is set for the "TextAdd" and "TextRem" types only.
	 * @property {number} Date - The time when this change was made.
	 */
	
	/**
	 * Report on all review changes.
	 * This is a dictionary where the keys are usernames.
	 * @typedef {Object.<string, Array.<ReviewReportRecord>>} ReviewReport
	 * @example
	 * 	{
	 * 	  "John Smith" : [{Type: 'TextRem', Value: 'Hello, Mark!', Date: 1679941734161},
	 * 	                {Type: 'TextAdd', Value: 'Dear Mr. Pottato.', Date: 1679941736189}],
	 * 	  "Mark Pottato" : [{Type: 'ParaRem', Date: 1679941755942},
	 * 	                  {Type: 'TextPr', Date: 1679941757832}]
	 * 	}
	 */
	
	/**
	 * Returns a report about every change which was made to the document in the review mode.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ReviewReport}
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
	 * Finds and replaces the text.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {Object} oProperties - The properties to find and replace.
	 * @param {string} oProperties.searchString - Search string.
	 * @param {string} oProperties.replaceString - Replacement string.
	 * @param {string} [oProperties.matchCase=true] - Case sensitive or not.
	 *
	 */
	ApiDocument.prototype.SearchAndReplace = function(oProperties)
	{
		let oProps = new AscCommon.CSearchSettings();
		oProps.SetText(oProperties["searchString"]);
		oProps.SetMatchCase(undefined !== oProperties["matchCase"] ? oProperties.matchCase : true);

		var oSearchEngine = this.Document.Search(oProps);
		if (!oSearchEngine)
			return;

		var sReplace = oProperties["replaceString"];
		
		sReplace = sReplace.replaceAll('\t', "^t");
		sReplace = sReplace.replaceAll('\v', "^l");
		sReplace = sReplace.replaceAll('\f', "^m");
		// TODO: ^p пока не поддерживается для замены
		//sReplace = sReplace.replaceAll('\r', "^p");
		sReplace = sReplace.replaceAll('\x0e', "^n");
		sReplace = sReplace.replaceAll('\x1e', "^~");
		// TODO: Сделать поддержку softHyphen
		//sReplace = sReplace.replaceAll('\x1f', "");
		
		this.Document.ReplaceSearchElement(sReplace, true, null, false);
	};
	/**
	 * Returns a list of all the content controls from the document.
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
			let oControl = ToApiContentControl(arrControls[nIndex]);
			if (oControl)
				arrResult.push(oControl);
		}

		return arrResult;
	};
	/**
	 * Returns a list of all tags that are used for all content controls in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {String[]}
	 */
	ApiDocument.prototype.GetTagsOfAllContentControls = function()
	{
		let oTags       = {};
		let arrResult   = [];
		let arrControls = this.Document.GetAllContentControls();
		for (let nIndex = 0, nCount = arrControls.length; nIndex < nCount; ++nIndex)
		{
			let oControl = arrControls[nIndex];
			let sTag = oControl.GetTag();

			if (sTag && !oTags[sTag])
			{
				oTags[sTag] = 1;
				arrResult.push(sTag);
			}
		}

		return arrResult;
	};
	/**
	 * Returns a list of all tags that are used for all forms in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {String[]}
	 */
	ApiDocument.prototype.GetTagsOfAllForms = function()
	{
		let oTags     = {};
		let arrResult = [];
		let arrForms  = this.Document.GetFormsManager().GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oForm = arrForms[nIndex];
			let sTag  = oForm.GetTag();

			if (sTag && !oTags[sTag])
			{
				oTags[sTag] = 1;
				arrResult.push(sTag);
			}
		}

		return arrResult;
	};
	/**
	 * Returns a list of all content controls in the document with the specified tag name.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param sTag {string} - Content control tag.
	 * @returns {ApiBlockLvlSdt[] | ApiInlineLvlSdt[]}
	 */
	ApiDocument.prototype.GetContentControlsByTag = function(sTag)
	{
		let _sTag = GetStringParameter(sTag, "");
		if (!_sTag)
			return [];

		let arrResult   = [];
		let arrControls = this.Document.GetAllContentControls();
		for (let nIndex = 0, nCount = arrControls.length; nIndex < nCount; ++nIndex)
		{
			let oControl    = arrControls[nIndex];
			let oApiControl = ToApiContentControl(oControl);
			if (_sTag === oControl.GetTag() && oApiControl)
				arrResult.push(oApiControl);
		}

		return arrResult;
	};
	/**
	 * Returns a list of all forms in the document with the specified tag name.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param sTag {string} - Form tag.
	 * @returns {ApiBlockLvlSdt[] | ApiInlineLvlSdt[]}
	 */
	ApiDocument.prototype.GetFormsByTag = function(sTag)
	{
		let _sTag = GetStringParameter(sTag, "");
		if (!_sTag)
			return [];

		let arrResult = [];
		let arrForms  = this.Document.GetFormsManager().GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oControl = arrForms[nIndex];
			let oForm    = ToApiForm(oControl);
			if (oControl.IsForm() && _sTag === oControl.GetTag() && oForm)
				arrResult.push(oForm);
		}

		return arrResult;
	};

	/**
	 * Sets the change tracking mode.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param isTrack {boolean} - Specifies if the change tracking mode is set or not.
	 */
	ApiDocument.prototype.SetTrackRevisions = function(isTrack)
	{
		this.Document.SetGlobalTrackRevisions(isTrack);
	};
	/**
	 * Checks if change tracking mode is enabled or not.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiDocument.prototype.IsTrackRevisions = function()
	{
		return this.Document.IsTrackRevisions();
	};
	/**
	 * Returns a Range object that represents the part of the document contained in the specified document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiDocument.prototype.GetRange = function(Start, End)
	{
		return new ApiRange(this.Document, Start, End);
	};
	/**
	 * Returns a range object by the current selection.
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
	 * Returns the last document element. 
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {?DocumentElement}
	 */
	ApiDocument.prototype.Last = function()
	{
		return this.GetElement(this.GetElementsCount() - 1);
	};
	
	/**
	 * Removes a bookmark from the document, if one exists.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sName - The bookmark name.
	 * @returns {boolean} - returns false if param is invalid.
	 */
	ApiDocument.prototype.DeleteBookmark = function(sName)
	{
		if (sName === undefined)
			return false;

		this.Document.RemoveBookmark(sName);

		return true;
	};
	/**
	 * Adds a comment to the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment text (required).
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {ApiComment?} - Returns null if the comment was not added.
	 */
	ApiDocument.prototype.AddComment = function(sText, sAuthor)
	{
		if (!sText || typeof(sText) !== "string")
			return null;
	
		if (typeof(sAuthor) !== "string")
			sAuthor = "";
		
		var CommentData = new AscCommon.CCommentData();
		CommentData.SetText(sText);
		CommentData.SetUserName(sAuthor);

		return AddGlobalCommentToDocument(this.Document, CommentData);
	};
	/**
	 * Returns a bookmark range.
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
	 * Returns a collection of section objects in the document.
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
	 * Returns a collection of tables on a given absolute page.
	 * <note>This method can be a little bit slow, because it runs the document calculation
	 * process to arrange tables on the specified page.</note>
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {number} nPage - The page number.
	 * @return {ApiTable[]}
	 */
	ApiDocument.prototype.GetAllTablesOnPage = function(nPage)
	{
		var arrApiAllTables = [];

		this.ForceRecalculate(nPage + 1);

		var arrAllTables = this.Document.GetAllTablesOnPage(nPage);
		for (var Index = 0; Index < arrAllTables.length; Index++)
		{
			arrApiAllTables.push(new ApiTable(arrAllTables[Index].Table));
		}

		return arrApiAllTables;
	};
	/**
	 * Removes the current selection.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 */
	ApiDocument.prototype.RemoveSelection = function()
	{
		this.Document.RemoveSelection();
	};
	/**
	 * Searches for a scope of a document object. The search results are a collection of ApiRange objects.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {boolean} isMatchCase - Case sensitive or not.
	 * @return {ApiRange[]}  
	 */
	ApiDocument.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase	= false;

		let oProps = new AscCommon.CSearchSettings();
		oProps.SetText(sText);
		oProps.SetMatchCase(isMatchCase);

		let searchEngine = this.Document.Search(oProps);
		let mapElements  = searchEngine.GetElementsMap();

		let apiRanges = [];
		for (let paraId in mapElements)
		{
			let paragraph = new ApiParagraph(mapElements[paraId]);
			apiRanges = apiRanges.concat(paragraph.Search(sText, isMatchCase));
		}

		return apiRanges;
	};
	/**
	 * Converts a document to Markdown.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {boolean} [bHtmlHeadings=false] - Defines if the HTML headings and IDs will be generated when the Markdown renderer of your target platform does not handle Markdown-style IDs.
	 * @param {boolean} [bBase64img=false] - Defines if the images will be created in the base64 format.
	 * @param {boolean} [bDemoteHeadings=false] - Defines if all heading levels in your document will be demoted to conform with the following standard: single H1 as title, H2 as top-level heading in the text body.
	 * @param {boolean} [bRenderHTMLTags=false] - Defines if HTML tags will be preserved in your Markdown. If you just want to use an occasional HTML tag, you can avoid using the opening angle bracket
	 * in the following way: \<tag&gt;text\</tag&gt;. By default, the opening angle brackets will be replaced with the special characters.
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
	 * Converts a document to HTML.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {boolean} [bHtmlHeadings=false] - Defines if the HTML headings and IDs will be generated when the Markdown renderer of your target platform does not handle Markdown-style IDs.
	 * @param {boolean} [bBase64img=false] - Defines if the images will be created in the base64 format.
	 * @param {boolean} [bDemoteHeadings=false] - Defines if all heading levels in your document will be demoted to conform with the following standard: single H1 as title, H2 as top-level heading in the text body.
	 * @param {boolean} [bRenderHTMLTags=false] - Defines if HTML tags will be preserved in your Markdown. If you just want to use an occasional HTML tag, you can avoid using the opening angle bracket
	 * in the following way: \<tag&gt;text\</tag&gt;. By default, the opening angle brackets will be replaced with the special characters.
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
	 * Inserts a watermark on each document page.
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
		for(sId in oHeadersMap){
			if(oHeadersMap.hasOwnProperty(sId)){
				privateInsertWatermarkToContent(this.Document.Api, oHeadersMap[sId], sText, bIsDiagonal);
			}
		}
	};

	/**
	 * Updates all tables of contents in the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {boolean} [bOnlyPageNumbers=false] - Specifies that only page numbers will be updated.
	 */
	ApiDocument.prototype.UpdateAllTOC = function(bOnlyPageNumbers)
	{
		if (typeof(bOnlyPageNumbers) !== "boolean")
			bOnlyPageNumbers = false;

		var oDocument = private_GetLogicDocument();
		var allTOC = oDocument.GetAllTablesOfContentsInDoc();

		for (var nItem = 0; nItem < allTOC.length; nItem++)
		{
			var oTOC = allTOC[nItem];
			if (!oTOC)
			{
				oTOC = oDocument.GetTableOfContents();
				if (!oTOC)
					return;
			}

			if (oTOC instanceof AscCommonWord.CBlockLevelSdt)
				oTOC = oTOC.GetInnerTableOfContents();

			if (!oTOC)
				return;

			var oState = oDocument.SaveDocumentState();

			oTOC.SelectField();

			if (bOnlyPageNumbers)
			{
				var arrParagraphs = oDocument.GetCurrentParagraph(false, true);

					for (var nParaIndex = 0, nParasCount = arrParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
					{
						var arrPageNumbers = arrParagraphs[nParaIndex].GetComplexFieldsArrayByType(AscCommonWord.fieldtype_PAGEREF);
						for (var nRefIndex = 0, nRefsCount = arrPageNumbers.length; nRefIndex < nRefsCount; ++nRefIndex)
						{
							arrPageNumbers[nRefIndex].Update();
						}
					}

					oDocument.LoadDocumentState(oState);
			}
			else
			{
				oTOC.Update();
				oDocument.LoadDocumentState(oState);
			}
		}
	};
	/**
	 * Updates all tables of figures in the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {boolean} [bOnlyPageNumbers=false] - Specifies that only page numbers will be updated.
	 */
	ApiDocument.prototype.UpdateAllTOF = function(bOnlyPageNumbers)
	{
		if (typeof(bOnlyPageNumbers) !== "boolean")
			bOnlyPageNumbers = false;

		var oDocument = private_GetLogicDocument();
		var allTOC = oDocument.GetAllTablesOfFiguresInDoc();

		for (var nItem = 0; nItem < allTOC.length; nItem++)
		{
			var oTOC = allTOC[nItem];
			if (!oTOC)
			{
				oTOC = oDocument.GetTableOfContents();
				if (!oTOC)
					return;
			}

			if (oTOC instanceof AscCommonWord.CBlockLevelSdt)
				oTOC = oTOC.GetInnerTableOfContents();

			if (!oTOC)
				return;

			var oState = oDocument.SaveDocumentState();

			oTOC.SelectField();

			if (bOnlyPageNumbers)
			{
				var arrParagraphs = oDocument.GetCurrentParagraph(false, true);

				for (var nParaIndex = 0, nParasCount = arrParagraphs.length; nParaIndex < nParasCount; ++nParaIndex)
				{
					var arrPageNumbers = arrParagraphs[nParaIndex].GetComplexFieldsArrayByType(AscCommonWord.fieldtype_PAGEREF);
					for (var nRefIndex = 0, nRefsCount = arrPageNumbers.length; nRefIndex < nRefsCount; ++nRefIndex)
					{
						arrPageNumbers[nRefIndex].Update();
					}
				}

				oDocument.LoadDocumentState(oState);
			}
			else
			{
				oTOC.Update();
				oDocument.LoadDocumentState(oState);
			}
		}
	};
	/**
	 * Converts the ApiDocument object into the JSON object.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteDefaultTextPr - Specifies if the default text properties will be written to the JSON object or not.
	 * @param {boolean} bWriteDefaultParaPr - Specifies if the default paragraph properties will be written to the JSON object or not.
	 * @param {boolean} bWriteTheme - Specifies if the document theme will be written to the JSON object or not.
	 * @param {boolean} bWriteSectionPr - Specifies if the document section properties will be written to the JSON object or not.
	 * @param {boolean} bWriteNumberings - Specifies if the document numberings will be written to the JSON object or not.
	 * @param {boolean} bWriteStyles - Specifies if the document styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiDocument.prototype.ToJSON = function(bWriteDefaultTextPr, bWriteDefaultParaPr, bWriteTheme, bWriteSectionPr, bWriteNumberings, bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();

		var oResult = {
			"type":      "document",
			"textPr":    bWriteDefaultTextPr ? oWriter.SerTextPr(this.GetDefaultTextPr().TextPr) : undefined,
			"paraPr":    bWriteDefaultParaPr ? oWriter.SerParaPr(this.GetDefaultParaPr().ParaPr) : undefined,
			"theme":     bWriteTheme ? oWriter.SerTheme(this.Document.GetTheme()) : undefined,
			"sectPr":    bWriteSectionPr ? oWriter.SerSectionPr(this.Document.SectPr) : undefined,
			"content":   oWriter.SerContent(this.Document.Content, undefined, undefined, undefined, true),
			"numbering": bWriteNumberings ? oWriter.jsonWordNumberings : undefined,
			"styles":    bWriteStyles ? oWriter.SerWordStylesForWrite() : undefined
		}

		return JSON.stringify(oResult);
	};
	/**
	 * Returns all existing forms in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiForm[]}
	 */
	ApiDocument.prototype.GetAllForms = function()
	{
		let arrApiForms = [];
		let arrForms    = this.Document.GetFormsManager().GetAllForms();
		for (let nIndex = 0, nCount = arrForms.length; nIndex < nCount; ++nIndex)
		{
			let oForm    = arrForms[nIndex];
			let oApiForm = ToApiForm(oForm);
			if (oApiForm)
				arrApiForms.push(oApiForm);
		}

		return arrApiForms;
	};

	/**
	 * Clears all forms in the document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 */
	ApiDocument.prototype.ClearAllFields = function()
	{
		this.Document.ClearAllSpecialForms(false);
	};

	/**
	 * Sets the highlight to the forms in the document.
	 * @memberof ApiDocument
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [bNone=false] - Defines that highlight will not be set.
	 * @typeofeditors ["CDE"]
	 */
	ApiDocument.prototype.SetFormsHighlight = function(r, g, b, bNone)
	{
		if (bNone === true)
			this.Document.SetSpecialFormsHighlight(null, null, null);
		else
			this.Document.SetSpecialFormsHighlight(r, g, b);
	};

	/**
	 * Returns all comments from the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiComment[]}
	 */
	ApiDocument.prototype.GetAllComments = function() 
	{
		let oCommManager = this.Document.GetCommentsManager();

		let aComments = Object.values(oCommManager.GetAllComments());
		let aApiComments = aComments.map(function(oComment) {
			return new ApiComment(oComment);
		});

		return aApiComments;
	};

	/**
	 * Returns a comment from the current document by its ID.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {string} sId - The comment ID.
	 * @returns {?ApiComment}
	 */
	ApiDocument.prototype.GetCommentById = function(sId) 
	{
		let manager = this.Document.GetCommentsManager();
		let comment = manager.GetByDurableId(sId);
		if (!comment)
			comment = manager.GetById(sId);

		return comment ? new ApiComment(comment) : null;
	};

	/**
	 * Returns all numbered paragraphs from the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParagraph[]}
	 */
	ApiDocument.prototype.GetAllNumberedParagraphs = function() 
	{
		var allParas = this.Document.GetAllNumberedParagraphs();
		var aResult = [];
		for (var nPara = 0; nPara < allParas.length; nPara++)
			aResult.push(new ApiParagraph(allParas[nPara]));

		return aResult;
	};

	/**
	 * Returns all heading paragraphs from the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParagraph[]}
	 */
	ApiDocument.prototype.GetAllHeadingParagraphs = function() 
	{
		var allParas = editor.asc_GetAllHeadingParagraphs();
		var aResult = [];
		for (var nPara = 0; nPara < allParas.length; nPara++)
			aResult.push(new ApiParagraph(allParas[nPara]));

		return aResult;
	};

	/**
	 * Returns the first paragraphs from all footnotes in the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParagraph[]}
	 */
	ApiDocument.prototype.GetFootnotesFirstParagraphs = function() 
	{
		var allParas = this.Document.GetFootNotesFirstParagraphs();
		var aResult = [];
		for (var nPara = 0; nPara < allParas.length; nPara++)
			aResult.push(new ApiParagraph(allParas[nPara]));

		return aResult;
	};

	/**
	 * Returns the first paragraphs from all endnotes in the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParagraph[]}
	 */
	ApiDocument.prototype.GetEndNotesFirstParagraphs = function() 
	{
		var allParas = this.Document.GetEndNotesFirstParagraphs();
		var aResult = [];
		for (var nPara = 0; nPara < allParas.length; nPara++)
			aResult.push(new ApiParagraph(allParas[nPara]));

		return aResult;
	};

	/**
	 * Returns all caption paragraphs of the specified type from the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {CaptionLabel | string} sCaption - The caption label ("Equation", "Figure", "Table", or another caption label).
	 * @returns {ApiParagraph[]}
	 */
	ApiDocument.prototype.GetAllCaptionParagraphs = function(sCaption) 
	{
		if (typeof(sCaption) !== "string" || sCaption.length === 0)
			return [];

		var allParas = this.Document.GetAllCaptionParagraphs(sCaption);
		var aResult = [];
		for (var nPara = 0; nPara < allParas.length; nPara++)
			aResult.push(new ApiParagraph(allParas[nPara]));

		return aResult;
	};
	
	/**
	 * Accepts all changes made in review mode.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 */
	ApiDocument.prototype.AcceptAllRevisionChanges = function()
	{
		this.Document.AcceptAllRevisionChanges();
	};

	/**
	 * Rejects all changes made in review mode.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 */
	ApiDocument.prototype.RejectAllRevisionChanges = function()
	{
		this.Document.RejectAllRevisionChanges();
	};

	/**
	 * Returns an array with names of all bookmarks in the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {string[]}
	 */
	ApiDocument.prototype.GetAllBookmarksNames = function() 
	{
		var aNames = [];
		var oManager = this.Document.GetBookmarksManager();
		oManager.Update();

		for (var i = 0, nCount = oManager.GetCount(); i < nCount; i++)
		{
			var sName = oManager.GetName(i);
			if (!oManager.IsInternalUseBookmark(sName) && !oManager.IsHiddenBookmark(sName))
				aNames.push(sName);
		}

		return aNames;
	};

	/**
     * Returns all the selected drawings in the current document.
     * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
     * @returns {ApiShape[] | ApiImage[] | ApiChart[] | ApiDrawing[]}
     */
	ApiDocument.prototype.GetSelectedDrawings = function() 
	{
		var aSelected = this.Document.DrawingObjects.selectedObjects;
		var aResult = [];
		for (var nDrawing = 0; nDrawing < aSelected.length; nDrawing++)
		{
			if (aSelected[nDrawing].isImage())
				aResult.push(new ApiImage(aSelected[nDrawing]));
			else if (aSelected[nDrawing].isChart())
				aResult.push(new ApiChart(aSelected[nDrawing]));
			else if (aSelected[nDrawing].isShape())
				aResult.push(new ApiShape(aSelected[nDrawing]));
			else
				aResult.push(new ApiDrawing(aSelected[nDrawing].parent));
		}

		var aSelectedInText = this.Document.GetSelectedDrawingObjectsInText();
		for (nDrawing = 0; nDrawing < aSelectedInText.length; nDrawing++)
		{
			if (aSelectedInText[nDrawing].GraphicObj.isImage())
				aResult.push(new ApiImage(aSelectedInText[nDrawing].GraphicObj));
			else if (aSelectedInText[nDrawing].GraphicObj.isChart())
				aResult.push(new ApiChart(aSelectedInText[nDrawing].GraphicObj));
			else if (aSelectedInText[nDrawing].GraphicObj.isShape())
				aResult.push(new ApiShape(aSelectedInText[nDrawing].GraphicObj));
			else
				aResult.push(new ApiDrawing(aSelected[nDrawing]));
		}

		return aResult;
	};

	/**
	 * Replaces the current image with an image specified.
	 * @typeofeditors ["CDE"]
	 * @memberof ApiDocument
	 * @param {string} sImageUrl - The image source where the image to be inserted should be taken from (currently, only internet URL or Base64 encoded images are supported).
	 * @param {EMU} Width - The image width in English measure units.
	 * @param {EMU} Height - The image height in English measure units.
	 */
	ApiDocument.prototype.ReplaceCurrentImage = function(sImageUrl, Width, Height)
	{
		let oLogicDocument = private_GetLogicDocument();
		let oDrawingObjects = oLogicDocument.DrawingObjects;

		let dK = 1 / 36000 / AscCommon.g_dKoef_pix_to_mm;
		oDrawingObjects.putImageToSelection(sImageUrl, Width * dK, Height * dK );
	};

	/**
	 * Replaces a drawing with a new drawing.
	 * @memberof ApiDocument
	 * @param {ApiDrawing} oOldDrawing - A drawing which will be replaced.
	 * @param {ApiDrawing} oNewDrawing - A drawing to replace the old drawing.
	 * @param {boolean} [bSaveOldDrawingPr=false] - Specifies if the old drawing settings will be saved.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiDocument.prototype.ReplaceDrawing = function(oOldDrawing, oNewDrawing, bSaveOldDrawingPr)
	{
		if (!(oOldDrawing instanceof ApiDrawing) || !oOldDrawing.Drawing.Parent || !(oNewDrawing instanceof ApiDrawing))
			return false;

		var oDrawing = oNewDrawing.Copy();

		if (bSaveOldDrawingPr === true)
			oDrawing.SetDrawingPrFromDrawing(oOldDrawing);

		var oDocument = private_GetLogicDocument();
		var oRun = oOldDrawing.Drawing.Parent.Get_DrawingObjectRun(oOldDrawing.Drawing.Id);
		if (oRun)
		{
			var oldSelectionInfo = oDocument.SaveDocumentState();

			for ( var CurPos = 0; CurPos < oRun.Content.length; CurPos++ )
			{
				var Element = oRun.Content[CurPos];

				if ( para_Drawing === Element.Type && oOldDrawing.Drawing.Id === Element.Get_Id() )
				{
					oRun.Remove_FromContent(CurPos, 1);
					oRun.Add_ToContent(CurPos, oDrawing.Drawing);
					break;
				}
			}

			oDocument.LoadDocumentState(oldSelectionInfo);
			return true;
		}

		return false;
	};

	/**
	 * Adds a footnote for the selected text (or the current position if the selection doesn't exist).
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {?ApiDocumentContent}
	 */
	ApiDocument.prototype.AddFootnote = function()
	{
		let oResult = this.Document.AddFootnote();
		if (!oResult)
			return null;

		return new ApiDocumentContent(oResult);
	};

	/**
	 * Adds an endnote for the selected text (or the current position if the selection doesn't exist).
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {?ApiDocumentContent}
	 */
	ApiDocument.prototype.AddEndnote = function()
	{
		let oResult = this.Document.AddEndnote();
		if (!oResult)
			return null;

		return new ApiDocumentContent(oResult);
	};

	/**
	 * Sets the highlight to the content controls from the current document.
	 * @memberof ApiDocument
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [bNone=false] - Defines that highlight will not be set.
	 * @typeofeditors ["CDE"]
	 */
	ApiDocument.prototype.SetControlsHighlight = function(r, g, b, bNone)
	{
		if (bNone === true)
			this.Document.SetSdtGlobalShowHighlight(false);
		else
		{
			this.Document.SetSdtGlobalShowHighlight(true);
			this.Document.SetSdtGlobalColor(r, g, b);
		}
	};

	/**
	 * Adds a table of content to the current document.
	 * <note>Please note that the new table of contents replaces the existing table of contents.</note>
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {TocPr} [oTocPr={}] - Table of contents properties.
	 */
	ApiDocument.prototype.AddTableOfContents = function(oTocPr)
	{
		if (typeof(oTocPr) !== "object")
			oTocPr = {};

		if (!oTocPr["BuildFrom"])
			oTocPr["BuildFrom"] = {"OutlineLvls": 9};

		let oTargetPr = new Asc.CTableOfContentsPr();
		oTargetPr.PageNumbers = typeof(oTocPr["ShowPageNums"]) == "boolean" ? oTocPr["ShowPageNums"] : true;
		oTargetPr.RightTab = typeof(oTocPr["RightAlgn"]) == "boolean" ? oTocPr["RightAlgn"] : true;
		oTargetPr.Hyperlink = typeof(oTocPr["FormatAsLinks"]) == "boolean" ? oTocPr["FormatAsLinks"] : true;

		if (oTocPr["LeaderType"] == null)
			oTocPr["LeaderType"] = "dot";
		switch (oTocPr["LeaderType"])
		{
			case "dot":
				oTargetPr.TabLeader = 0;
				break;
			case "dash":
				oTargetPr.TabLeader = 2;
				break;
			case "underline":
				oTargetPr.TabLeader = 5;
				break;
			case "none":
				oTargetPr.TabLeader = 4;
				break;
		}

		if (Array.isArray(oTocPr["BuildFrom"]["StylesLvls"]) && oTocPr["BuildFrom"]["StylesLvls"].length > 0)
		{
			oTargetPr.Styles = oTocPr["BuildFrom"]["StylesLvls"];
			oTargetPr.OutlineEnd = -1;
			oTargetPr.OutlineStart = -1;
		}
		else if (oTocPr["BuildFrom"]["OutlineLvls"] >= 1 && oTocPr["BuildFrom"]["OutlineLvls"] <= 9)
		{
			oTargetPr.OutlineEnd = oTocPr["BuildFrom"]["OutlineLvls"];
		}

		if (oTocPr["TocStyle"] == null)
			oTocPr["TocStyle"] = "standard";
		switch (oTocPr["TocStyle"])
		{
			case "simple":
				oTargetPr.StylesType = 1;
				break;
			case "online":
				oTargetPr.StylesType = 5;
				break;
			case "standard":
				oTargetPr.StylesType = 2;
				break;
			case "modern":
				oTargetPr.StylesType = 3;
				break;
			case "classic":
				oTargetPr.StylesType = 4;
				break;
		}

		var oTOC = this.Document.GetTableOfContents();
		if (oTOC instanceof AscCommonWord.CBlockLevelSdt && oTOC.IsBuiltInUnique())
		{
			var oInnerTOC = oTOC.GetInnerTableOfContents();
			oTOC = oInnerTOC;
			if (!oTOC)
				return;

			var oStyles = this.Document.GetStyles();
			var nStylesType = oTargetPr.get_StylesType();
			var isNeedChangeStyles = (Asc.c_oAscTOCStylesType.Current !== nStylesType && nStylesType !== oStyles.GetTOCStylesType());

			oTOC.SelectField();
			if (isNeedChangeStyles)
				oStyles.SetTOCStylesType(nStylesType);

			oTOC.SetPr(oTargetPr);
			oTOC.Update();
			return;
		}

		this.Document.AddTableOfContents(null, oTargetPr, undefined);
	};

	/**
	 * Adds a table of figures to the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @param {TofPr} [oTofPr={}] - Table of figures properties.
	 * <note>Please note that the table of figures properties will be filled with the default properties if they are undefined.</note>
	 * @param {boolean} [bReplace=true] - Specifies whether to replace the selected table of figures instead of adding a new one.
	 * @returns {boolean}
	 */
	ApiDocument.prototype.AddTableOfFigures = function(oTofPr, bReplace)
	{
		if (typeof(oTofPr) !== "object")
			oTofPr = {};
		if (typeof(bReplace) !== "boolean")
			bReplace = true;

		let oTargetPr = new Asc.CTableOfContentsPr();
		oTargetPr.PageNumbers = typeof(oTofPr["ShowPageNums"]) == "boolean" ? oTofPr["ShowPageNums"] : true;
		oTargetPr.RightTab = typeof(oTofPr["RightAlgn"]) == "boolean" ? oTofPr["RightAlgn"] : true;
		oTargetPr.Hyperlink = typeof(oTofPr["FormatAsLinks"]) == "boolean" ? oTofPr["FormatAsLinks"] : true;
		oTargetPr.IsIncludeLabelAndNumber = typeof(oTofPr["LabelNumber"]) == "boolean" ? oTofPr["LabelNumber"] : true;
		oTargetPr.OutlineEnd = -1;
		oTargetPr.OutlineStart = -1;


		if (oTofPr["LeaderType"] == null)
			oTofPr["LeaderType"] = "dot";
		switch (oTofPr["LeaderType"])
		{
			case "dot":
				oTargetPr.TabLeader = 0;
				break;
			case "dash":
				oTargetPr.TabLeader = 2;
				break;
			case "underline":
				oTargetPr.TabLeader = 5;
				break;
			case "none":
				oTargetPr.TabLeader = 4;
				break;
		}

		if (typeof(oTofPr["BuildFrom"]) !== "string")
			oTofPr["BuildFrom"] = "Figure";

		let aStyles = this.Document.GetAllUsedParagraphStyles();
		let isUsedStyle = false;
		for (var i = 0; i < aStyles.length; i++)
		{
			if (aStyles[i].Name == oTofPr["BuildFrom"])
			{
				isUsedStyle = true;
				break;
			}
		}

		if (isUsedStyle)
		{
			oTargetPr.Styles = [{ Name: oTofPr["BuildFrom"], Lvl: undefined }];
			oTargetPr.Caption = null;
		}
		else
			oTargetPr.Caption = oTofPr["BuildFrom"];

		if (oTofPr["TofStyle"] == null)
			oTofPr["TofStyle"] = "distinctive";
		switch (oTofPr["TofStyle"])
		{
			case "simple":
				oTargetPr.StylesType = 5;
				break;
			case "online":
				oTargetPr.StylesType = 6;
				break;
			case "classic":
				oTargetPr.StylesType = 1;
				break;
			case "distinctive":
				oTargetPr.StylesType = 2;
				break;
			case "centered":
				oTargetPr.StylesType = 3;
				break;
			case "formal":
				oTargetPr.StyleType = 4;
				break;
		}

		let aTOF = this.Document.GetAllTablesOfFigures(true);
		let oTOFToReplace = null;
		if(aTOF.length > 0)
		{
			oTOFToReplace = aTOF[0];
		}

		if (oTOFToReplace && bReplace)
		{
			oTOFToReplace.SelectFieldValue();
			oTargetPr.ComplexField = oTOFToReplace;

			var oStyles     = this.Document.GetStyles();
			var nStylesType = oTargetPr.get_StylesType();
			var isNeedChangeStyles = (Asc.c_oAscTOFStylesType.Current !== nStylesType && nStylesType !== oStyles.GetTOFStyleType());
			if (isNeedChangeStyles)
				oStyles.SetTOFStyleType(nStylesType);

			oTOFToReplace.SetPr(oTargetPr);
			oTOFToReplace.Update();
		}
		else
			this.Document.AddTableOfFigures(oTargetPr);

		return true;
	};

	/**
	 * Returns the document statistics represented as an object with the following parameters:
	 * * <b>PageCount</b> - number of pages;
	 * * <b>WordsCount</b> - number of words;
	 * * <b>ParagraphCount</b> - number of paragraphs;
	 * * <b>SymbolsCount</b> - number of symbols;
	 * * <b>SymbolsWSCount</b> - number of symbols with spaces.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {object}
	 */
	ApiDocument.prototype.GetStatistics = function()
	{
		let oLogicDocument = this.Document;
		
		oLogicDocument.Statistics.StartPos = 0;
        oLogicDocument.Statistics.CurPage  = 0;

        oLogicDocument.Statistics.Pages           = 0;
        oLogicDocument.Statistics.Words           = 0;
        oLogicDocument.Statistics.Paragraphs      = 0;
        oLogicDocument.Statistics.SymbolsWOSpaces = 0;
        oLogicDocument.Statistics.SymbolsWhSpaces = 0;

		oLogicDocument.Statistics.Update_Pages(oLogicDocument.Pages.length);
		for (let CurPage = 0, PagesCount = oLogicDocument.Pages.length; CurPage < PagesCount; ++CurPage)
		{
			oLogicDocument.DrawingObjects.documentStatistics(CurPage, oLogicDocument.Statistics);
		}

		let Count = oLogicDocument.Content.length;
		for (let Index = oLogicDocument.Statistics.StartPos; Index < Count; ++Index)
		{
			let Element = oLogicDocument.Content[Index];
			Element.CollectDocumentStatistics(oLogicDocument.Statistics);
		}

		return {
            "PageCount"      : oLogicDocument.Statistics.Pages,
            "WordsCount"     : oLogicDocument.Statistics.Words,
            "ParagraphCount" : oLogicDocument.Statistics.Paragraphs,
            "SymbolsCount"   : oLogicDocument.Statistics.SymbolsWOSpaces,
            "SymbolsWSCount" : oLogicDocument.Statistics.SymbolsWhSpaces
        }
	};
	/**
	 * Returns a number of pages in the current document.
	 * <note>This method can be slow for large documents because it runs the document calculation
	 * process before the full recalculation.</note>
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @return {number}
	 */
	ApiDocument.prototype.GetPageCount = function()
	{
		this.ForceRecalculate();
		return this.Document.GetPagesCount();
	};
	/**
	 * Returns all styles of the current document.
	 * @memberof ApiDocument
	 * @typeofeditors ["CDE"]
	 * @returns {ApiStyle[]}
	 */
	ApiDocument.prototype.GetAllStyles = function()
	{
		let aApiStyles = [];
		var aStyles  = this.Document.Get_Styles().Style;
		
		aStyles.forEach(function(style) {
			aApiStyles.push(new ApiStyle(style));
		});

		return aApiStyles;
	};
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiParagraph
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiParagraph class.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"paragraph"}
	 */
	ApiParagraph.prototype.GetClassType = function()
	{
		return "paragraph";
	};
	/**
	 * Adds some text to the current paragraph.
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
	 * Adds a page break and starts the next element from the next page.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddPageBreak = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new AscWord.CRunBreak(AscWord.break_Page));
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Adds a line break to the current position and starts the next element from a new line.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddLineBreak = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new AscWord.CRunBreak(AscWord.break_Line));
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};

	/**
	 * Adds a column break to the current position and starts the next element from a new column.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddColumnBreak = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new AscWord.CRunBreak(AscWord.break_Column));
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Inserts a number of the current document page into the paragraph.
	 * <note>This method works for the paragraphs in the document header/footer only.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddPageNumber = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new AscWord.CRunPageNum());
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Inserts a number of pages in the current document into the paragraph.
	 * <note>This method works for the paragraphs in the document header/footer only.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddPagesCount = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new AscWord.CRunPagesCount());
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Returns the text properties of the paragraph mark which is used to mark the paragraph end. The mark can also acquire
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
	 * Returns the paragraph properties.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiParaPr}
	 */
	ApiParagraph.prototype.GetParaPr = function()
	{
		return new ApiParaPr(this, this.Paragraph.Pr.Copy());
	};
	/**
	 * Returns the numbering definition and numbering level for the numbered list.
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
	 * Specifies that the current paragraph references the numbering definition instance in the current document.
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
	 * Returns a number of elements in the current paragraph.
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
	 * Returns a paragraph element using the position specified.
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
	 * Removes an element using the position specified.
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
	 * Removes all the elements from the current paragraph.
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
	 * Deletes the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {boolean} - returns false if paragraph haven't parent.
	 */
	ApiParagraph.prototype.Delete = function()
	{
		var oParent = this.Paragraph.GetParent();
		var nPosInParent = this.Paragraph.GetIndex();

		if (nPosInParent !== - 1)
		{
			this.Paragraph.PreDelete();
			oParent.Remove_FromContent(nPosInParent, 1, true);

			return true;
		}
		else 
			return false;
	};
	/**
	 * Returns the next paragraph.
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
	 * Returns the previous paragraph.
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
	 * Creates a paragraph copy. Ingnore comments, footnote references, complex fields.
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
	 * Adds an element to the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ParagraphContent} oElement - The document element which will be added at the current position. Returns false if the
	 * oElement type is not supported by a paragraph.
	 * @param {number} [nPos] - The position where the current element will be added. If this value is not
	 * specified, then the element will be added at the end of the current paragraph.
	 * @returns {boolean} Returns <code>false</code> if the type of <code>oElement</code> is not supported by paragraph
	 * content.
	 */
	ApiParagraph.prototype.AddElement = function(oElement, nPos)
	{
		// TODO: ParaEnd
		if (!private_IsSupportedParaElement(oElement) || nPos < 0 || nPos > this.Paragraph.Content.length - 1)
			return false;

		var oParaElement = oElement.private_GetImpl();
		if (oParaElement.IsUseInDocument())
			return false;

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
	 * Adds a tab stop to the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiRun}
	 */
	ApiParagraph.prototype.AddTabStop = function()
	{
		var oRun = new ParaRun(this.Paragraph, false);
		oRun.Add_ToContent(0, new AscWord.CRunTab());
		private_PushElementToParagraph(this.Paragraph, oRun);
		return new ApiRun(oRun);
	};
	/**
	 * Adds a drawing object (image, shape or chart) to the current paragraph.
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
		private_CheckDrawingOnAdd(oDrawing);
		return new ApiRun(oRun);
	};

	/**
	 * Adds an inline container.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {ApiInlineLvlSdt} oSdt - An inline container. If undefined or null, then new class ApiInlineLvlSdt will be created and added to the paragraph.
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
	 * Adds a comment to the current paragraph.
	 * <note>Please note that this paragraph must be in the document.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment text (required).
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {ApiComment?} - Returns null if the comment was not added.
	 */
	ApiParagraph.prototype.AddComment = function(sText, sAuthor)
	{
		if (!sText || typeof(sText) !== "string")
			return null;
		if (typeof(sAuthor) !== "string")
			sAuthor = "";

		if (!this.Paragraph.IsUseInDocument())
			return null;

		var oDocument = private_GetLogicDocument();

		var sQuotedText = this.GetText();
		var CommentData = new AscCommon.CCommentData();
		CommentData.Set_QuoteText(sQuotedText);
		CommentData.SetText(sText);
		CommentData.SetUserName(sAuthor);

		var oComment = new AscCommon.CComment(oDocument.Comments, CommentData);
		oComment.GenerateDurableId();
		oDocument.Comments.Add(oComment);
		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.AddComment(oComment, true, true);
		this.Paragraph.SetApplyToAll(false);

		if (null != oComment)
			editor.sync_AddComment(oComment.Get_Id(), CommentData);
		else
			return null;

		return new ApiComment(oComment)
	};
	/**
	 * Adds a hyperlink to a paragraph. 
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
		var oHyperlink;

		this.Paragraph.SelectAll(1);
		if (!AscCommon.rx_allowedProtocols.test(sLink))
			sLink = (urlType === 0) ? null :(( (urlType === 2) ? 'mailto:' : 'http://' ) + sLink);

		sLink = sLink && sLink.replace(new RegExp("%20",'g')," ");
		hyperlinkPr.put_Value(sLink);
		hyperlinkPr.put_ToolTip(sScreenTipText);
		hyperlinkPr.put_Bookmark(null);
		
		oHyperlink = new ApiHyperlink(this.Paragraph.AddHyperlink(hyperlinkPr));
		oDocument.RemoveSelection();

		return oHyperlink;
	};
	/**
	 * Returns a Range object that represents the part of the document contained in the specified paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiParagraph.prototype.GetRange = function(Start, End)
	{
		return new ApiRange(this.Paragraph, Start, End);
	};
	/**
	 * Adds an element to the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {ParagraphContent} oElement - The document element which will be added at the current position. Returns false if the
	 * oElement type is not supported by a paragraph.
	 * @returns {boolean} Returns <code>false</code> if the type of <code>oElement</code> is not supported by paragraph
	 * content.
	 */
	ApiParagraph.prototype.Push = function(oElement)
	{
		if (oElement.private_GetImpl().IsUseInDocument())
			return false;

		if (private_IsSupportedParaElement(oElement))
		{
			this.AddElement(oElement);
		}
		else if (typeof oElement === "string")
		{
			var LastTextPrInParagraph;

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
	 * Returns the last Run with text in the current paragraph.
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
					if (Element.Run.GetElement(Index).IsText())
					{
						return Element;
					}
				}
			}
		}

		return this.GetElement(this.GetElementsCount() - 1);
	};
	/**
	 * Sets the bold property to the text character.
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
	 * Specifies that any lowercase characters in this paragraph are formatted for display only as their capital letter character equivalents.
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
	 * Sets the text color to the current paragraph in the RGB format.
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
	 * Specifies that the contents of this paragraph are displayed with two horizontal lines through each character displayed on the line.
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
	 * Sets all 4 font slots with the specified font family.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string} sFontFamily - The font family or families used for the current paragraph.
	 * @returns {?ApiParagraph} this
	 */
	ApiParagraph.prototype.SetFontFamily = function(sFontFamily)
	{
		if (typeof sFontFamily !== "string")
			return null;

		var oThis    = this;
		var loader   = AscCommon.g_font_loader;
		var fontinfo = g_fontApplication.GetFontInfo(sFontFamily);
		var isasync  = loader.LoadFont(fontinfo, setFontFamily);

		if (isasync === false)
			setFontFamily()

		function setFontFamily()
		{
			var FontFamily = {
				Name : sFontFamily,
				Index : -1
			};

			oThis.Paragraph.SetApplyToAll(true);
			oThis.Paragraph.Add(new AscCommonWord.ParaTextPr({FontFamily : FontFamily}));
			oThis.Paragraph.SetApplyToAll(false);
		}

		return this;
	};
	/**
	 * Returns all font names from all elements inside the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @returns {string[]} - The font names used for the current paragraph.
	 */
	ApiParagraph.prototype.GetFontNames = function()
	{
		let fontMap = {};
		let arrFonts = [];
		// this.Paragraph.Document_Get_AllFontNames(fontMap);
		this.Paragraph.Get_CompiledPr().TextPr.Document_Get_AllFontNames(fontMap);
		for (let i = 0; i < this.Paragraph.Content.length; i++)
		{
			let element = this.Paragraph.Content[i];
			element.Get_CompiledTextPr().Document_Get_AllFontNames(fontMap);
		}
		for (let key in fontMap)
		{
			arrFonts.push(key);
		}
		return arrFonts;
	};
	/**
	 * Sets the font size to the characters of the current paragraph.
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
	 * Specifies a highlighting color which is applied as a background to the contents of the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {highlightColor} sColor - Available highlight color.
	 * @returns {ApiParagraph} this
	 */
	ApiParagraph.prototype.SetHighlight = function(sColor)
	{
		if (!editor && Asc.editor)
		 	return this;

		this.Paragraph.SetApplyToAll(true);
		if ("none" === sColor)
		{
			if (editor.editorId === AscCommon.c_oEditorId.Word)
				this.Paragraph.Add(new ParaTextPr({HighLight : highlight_None}));
			else if (editor.editorId === AscCommon.c_oEditorId.Presentation)
				this.Paragraph.Add(new ParaTextPr({HighlightColor : null}));
		}
		else
		{
			var color = private_getHighlightColorByName(sColor);
			if (color && editor.editorId === AscCommon.c_oEditorId.Word)
			{
				color = new CDocumentColor(color.r, color.g, color.b);
				this.Paragraph.Add(new ParaTextPr({HighLight : color}));
			}
			else if (color && editor.editorId === AscCommon.c_oEditorId.Presentation)
			{
				color = AscFormat.CreateUniColorRGB(color.r, color.g, color.b);
				this.Paragraph.Add(new ParaTextPr({HighlightColor : color}));
			}
		}
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Sets the italic property to the text character.
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
	 * Specifies an amount by which text is raised or lowered for this paragraph in relation to the default
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
	 * Specifies that all the small letter characters in this paragraph are formatted for display only as their capital
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
	 * Sets the text spacing measured in twentieths of a point.
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
	 * Specifies that the contents of this paragraph are displayed with a single horizontal line through the center of the line.
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
	 * Specifies that the contents of this paragraph are displayed along with a line appearing directly below the character
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
	 * Specifies the alignment which will be applied to the contents of this paragraph in relation to the default appearance of the paragraph text:
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
			value = AscCommon.vertalign_Baseline;
		else if (sType === "subscript")
			value = AscCommon.vertalign_SubScript;
		else if (sType === "superscript")
			value = AscCommon.vertalign_SuperScript;
		else 
			return null;

		this.Paragraph.SetApplyToAll(true);
		this.Paragraph.Add(new AscCommonWord.ParaTextPr({VertAlign : value}));
		this.Paragraph.SetApplyToAll(false);
		
		return this;
	};
	/**
	 * Returns the last element of the paragraph which is not empty.
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
	 * Returns a collection of content control objects in the paragraph.
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
			let oControl = ToApiContentControl(ContentControls[Index]);
			if (oControl)
				arrApiContentControls.push(oControl);
		} 

		return arrApiContentControls;
	};
	/**
	 * Returns a collection of drawing objects in the paragraph.
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
	 * Returns a collection of shape objects in the paragraph.
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
	 * Returns a collection of image objects in the paragraph.
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
			if (arrAllDrawing[Index].GraphicObj instanceof AscFormat.CImageShape)
				arrApiImages.push(new ApiImage(arrAllDrawing[Index].GraphicObj));
		}

		return arrApiImages;
	};
	/**
	 * Returns a collection of chart objects in the paragraph.
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
			if (arrAllDrawing[Index].GraphicObj instanceof AscFormat.CChartSpace)
				arrApiCharts.push(new ApiChart(arrAllDrawing[Index].GraphicObj));
		}

		return arrApiCharts;
	};
	/**
	 * Returns a collection of OLE objects in the paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiOleObject[]}  
	 */
	ApiParagraph.prototype.GetAllOleObjects = function()
	{
		var arrAllDrawing = this.Paragraph.GetAllDrawingObjects();
		var arrApiOleObjects  = [];

		for (var Index = 0; Index < arrAllDrawing.length; Index++)
		{
			if (arrAllDrawing[Index].GraphicObj instanceof AscFormat.COleObject)
				arrApiOleObjects.push(new ApiOleObject(arrAllDrawing[Index].GraphicObj));
		}

		return arrApiOleObjects;
	};
	/**
	 * Returns a content control that contains the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {ApiBlockLvlSdt | null} - returns null is parent content control doesn't exist.  
	 */
	ApiParagraph.prototype.GetParentContentControl = function()
	{
		var ParaPosition = this.Paragraph.GetDocumentPositionFromObject();

		for (var Index = ParaPosition.length - 1; Index >= 1; Index--)
		{
			if (ParaPosition[Index].Class && ParaPosition[Index].Class.Parent && ParaPosition[Index].Class.Parent instanceof CBlockLevelSdt)
			{
				return new ApiBlockLvlSdt(ParaPosition[Index].Class.Parent);
			}
		}

		return null;
	};
	/**
	 * Returns a table that contains the current paragraph.
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
	 * Returns a table cell that contains the current paragraph.
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
	 * Returns the paragraph text.
	 * @memberof ApiParagraph
	 * @param {object} oPr - The resulting string display properties.
     * @param {boolean} [oPr.Numbering=false] - Defines if the resulting string will include numbering or not.
     * @param {boolean} [oPr.Math=false] - Defines if the resulting string will include mathematical expressions or not.
	 * @param {string} [oPr.NewLineSeparator='\r'] - Defines how the line separator will be specified in the resulting string.
	 * @param {string} [oPr.TabSymbol='\t'] - Defines how the tab will be specified in the resulting string (does not apply to numbering).
	 * @typeofeditors ["CDE"]
	 * @return {string}
	 */
	ApiParagraph.prototype.GetText = function(oPr)
	{
		if (!oPr) {
			oPr = {};
		}

		let oProp =	{
			NewLineSeparator:	(oPr.hasOwnProperty("NewLineSeparator")) ? oPr["NewLineSeparator"] : "\r",
			Numbering:			(oPr.hasOwnProperty("Numbering")) ? oPr["Numbering"] : true,
			Math:				(oPr.hasOwnProperty("Math")) ? oPr["Math"] : true,
			TabSymbol:			oPr["TabSymbol"],
			ParaEndToSpace:		false
		}

		return this.Paragraph.GetText(oProp);
	};
	/**
	 * Returns the paragraph text properties.
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
	 * Sets the paragraph text properties.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The paragraph text properties.
	 * @return {boolean} - returns false if param is invalid.
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
	 * Wraps the paragraph object with a rich text content control.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {number} nType - Defines if this method returns the ApiBlockLvlSdt (nType === 1) or ApiParagraph (any value except 1) object.
	 * @return {ApiParagraph | ApiBlockLvlSdt}  
	 */
	ApiParagraph.prototype.InsertInContentControl = function(nType)
	{
		var Document = private_GetLogicDocument();
		var ContentControl;

		var paraIndex	= this.Paragraph.Index;
		if (paraIndex >= 0)
		{
			this.Select();
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
	 * Inserts a paragraph at the specified position.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string | ApiParagraph} paragraph - Text or paragraph.
	 * @param {string} sPosition - The position where the text or paragraph will be inserted ("before" or "after" the paragraph specified).
	 * @param {boolean} beRNewPara - Defines if this method returns a new paragraph (true) or the current paragraph (false).
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
	 * Selects the current paragraph.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @return {boolean}
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

		var controllerType;

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
	 * Searches for a scope of a paragraph object. The search results are a collection of ApiRange objects.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {boolean} isMatchCase - Case sensitive or not.
	 * @return {ApiRange[]}  
	 */
	ApiParagraph.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase = false;

		let oDocument = private_GetLogicDocument();
		let oProps    = new AscCommon.CSearchSettings();
		oProps.SetText(sText);
		oProps.SetMatchCase(!!isMatchCase);
		oDocument.Search(oProps);

		var SearchResults = this.Paragraph.SearchResults;
		let arrApiRanges  = [];
		for (var FoundId in SearchResults)
		{
			var StartSearchContentPos = SearchResults[FoundId].StartPos;
			var EndSearchContentPos   = SearchResults[FoundId].EndPos;

			var StartChar	= this.Paragraph.ConvertParaContentPosToRangePos(StartSearchContentPos);
			var EndChar		= this.Paragraph.ConvertParaContentPosToRangePos(EndSearchContentPos);
			if (EndChar > 0)
				EndChar--;

			arrApiRanges.push(this.GetRange(StartChar, EndChar));
		}

		return arrApiRanges;
	};
	/**
	 * Wraps the paragraph content in a mail merge field.
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

	/**
	 * Adds a numbered cross-reference to the current paragraph.
	 * <note>Please note that this paragraph must be in the document.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {numberedRefTo} sRefType - The text or numeric value of a numbered reference you want to insert.
	 * @param {ApiParagraph} oParaTo - The numbered paragraph to be referred to (must be in the document).
	 * @param {boolean} [bLink=true] - Specifies if the reference will be inserted as a hyperlink.
	 * @param {boolean} [bAboveBelow=false] - Specifies if the above/below words indicating the position of the reference should be included (don't used with the "text" and "aboveBelow" sRefType).
	 * @param {string} [sSepWith=""] - A number separator (used only with the "fullCtxParaNum" sRefType).
	 * @returns {boolean}
	 */
	ApiParagraph.prototype.AddNumberedCrossRef = function(sRefTo, oParaTo, bLink, bAboveBelow, sSepWith) 
	{
		var nRefTo = -1;
		switch (sRefTo)
		{
			case "pageNum":
				nRefTo = 1;
				break;
			case "paraNum":
				nRefTo = 2;
				break;
			case "noCtxParaNum":
				nRefTo = 3;
				break;
			case "fullCtxParaNum":
				nRefTo = 4;
				break;
			case "text":
				nRefTo = 0;
				break;
			case "aboveBelow":
				nRefTo = 5;
				break;
		}
		if (nRefTo === -1 || !(oParaTo instanceof ApiParagraph) || !this.Paragraph.IsUseInDocument())
			return false;
		if (typeof(bLink) !== "boolean")
			bLink = true;
		if (typeof(bAboveBelow) !== "boolean" || nRefTo === 5 || nRefTo === 0)
			bAboveBelow = false;
		if (typeof(sSepWith) !== "string" || nRefTo !== 4)
			sSepWith = "";

		var oDocument = private_GetLogicDocument();
		var oldSelectionInfo;

		var allNumberedParas = oDocument.GetAllNumberedParagraphs();
		for (var nPara = 0; nPara < allNumberedParas.length; nPara++)
		{
			if (allNumberedParas[nPara].Id === oParaTo.Paragraph.Id)
			{
				oldSelectionInfo = oDocument.SaveDocumentState();
				oDocument.RemoveSelection();
				this.Paragraph.Document_SetThisElementCurrent();
				this.Paragraph.SetCurrentPos(this.Paragraph.Content.length - 1);
				oDocument.AddRefToParagraph(oParaTo.private_GetImpl(), nRefTo, bLink, bAboveBelow, sSepWith);
				oDocument.LoadDocumentState(oldSelectionInfo);
				return true;
			}
		}
		return false;
	};

	/**
	 * Adds a heading cross-reference to the current paragraph.
	 * <note>Please note that this paragraph must be in the document.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {headingRefTo} sRefType - The text or numeric value of a heading reference you want to insert.
	 * @param {ApiParagraph} oParaTo - The heading paragraph to be referred to (must be in the document).
	 * @param {boolean} [bLink=true] - Specifies if the reference will be inserted as a hyperlink.
	 * @param {boolean} [bAboveBelow=false] - Specifies if the above/below words indicating the position of the reference should be included (don't used with the "text" and "aboveBelow" sRefType).
	 * @returns {boolean}
	 */
	ApiParagraph.prototype.AddHeadingCrossRef = function(sRefTo, oParaTo, bLink, bAboveBelow) 
	{
		var nRefTo = -1;
		switch (sRefTo)
		{
			case "text":
				nRefTo = 0;
				break;
			case "pageNum":
				nRefTo = 1;
				break;
			case "headingNum":
				nRefTo = 2;
				break;
			case "noCtxHeadingNum":
				nRefTo = 3;
				break;
			case "fullCtxHeadingNum":
				nRefTo = 4;
				break;
			case "aboveBelow":
				nRefTo = 5;
				break;
		}
		if (nRefTo === -1 || !(oParaTo instanceof ApiParagraph) || !oParaTo.Paragraph.IsUseInDocument() || !this.Paragraph.IsUseInDocument())
			return false;
		if (typeof(bLink) !== "boolean")
			bLink = true;
		if (typeof(bAboveBelow) !== "boolean" || nRefTo === 5 || nRefTo === 0)
			bAboveBelow = false;

		var nOutlineLvl = oParaTo.Paragraph.GetOutlineLvl();
		if (nOutlineLvl === undefined || nOutlineLvl < 0 || nOutlineLvl > 8)
			return false;

		var oDocument = private_GetLogicDocument();
		var oldSelectionInfo = oDocument.SaveDocumentState();
		oDocument.RemoveSelection();
		this.Paragraph.Document_SetThisElementCurrent();
		this.Paragraph.SetCurrentPos(this.Paragraph.Content.length - 1);
		oDocument.AddRefToParagraph(oParaTo.private_GetImpl(), nRefTo, bLink, bAboveBelow, undefined);
		oDocument.LoadDocumentState(oldSelectionInfo);
		return true;
	};

	/**
	 * Adds a bookmark cross-reference to the current paragraph.
	 * <note>Please note that this paragraph must be in the document.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {bookmarkRefTo} sRefType - The text or numeric value of a bookmark reference you want to insert.
	 * @param {string} sBookmarkName - The name of the bookmark to be referred to (must be in the document).
	 * @param {boolean} [bLink=true] - Specifies if the reference will be inserted as a hyperlink.
	 * @param {boolean} [bAboveBelow=false] - Specifies if the above/below words indicating the position of the reference should be included (don't used with the "text" and "aboveBelow" sRefType).
	 * @param {string} [sSepWith=""] - A number separator (used only with the "fullCtxParaNum" sRefType).
	 * @returns {boolean}
	 */
	ApiParagraph.prototype.AddBookmarkCrossRef = function(sRefTo, sBookmarkName, bLink, bAboveBelow, sSepWith) 
	{
		var nRefTo = -1;
		switch (sRefTo)
		{
			case "text":
				nRefTo = 0;
				break;
			case "pageNum":
				nRefTo = 1;
				break;
			case "paraNum":
				nRefTo = 2;
				break;
			case "noCtxParaNum":
				nRefTo = 3;
				break;
			case "fullCtxParaNum":
				nRefTo = 4;
				break;
			case "aboveBelow":
				nRefTo = 5;
				break;
		}
		if (nRefTo === -1 || typeof(sBookmarkName) !== "string" || sBookmarkName.length === 0 || !this.Paragraph.IsUseInDocument())
			return false;
		if (typeof(bLink) !== "boolean")
			bLink = true;
		if (typeof(bAboveBelow) !== "boolean" || nRefTo === 5 || nRefTo === 0)
			bAboveBelow = false;
		if (typeof(sSepWith) !== "string" || nRefTo !== 4)
			sSepWith = "";

		var oDocument = private_GetLogicDocument();
		var oManager = oDocument.GetBookmarksManager();
		var sName, oldSelectionInfo;
		for (var nBookmark = 0, nCount = oManager.GetCount(); nBookmark < nCount; nBookmark++)
		{
			sName = oManager.GetName(nBookmark);
			if (!oManager.IsInternalUseBookmark(sName) && !oManager.IsHiddenBookmark(sName) && sName === sBookmarkName)
			{
				oldSelectionInfo = oDocument.SaveDocumentState();
				oDocument.RemoveSelection();
				this.Paragraph.Document_SetThisElementCurrent();
				this.Paragraph.SetCurrentPos(this.Paragraph.Content.length - 1);
				oDocument.AddRefToBookmark(sBookmarkName, nRefTo, bLink, bAboveBelow, sSepWith);
				oDocument.LoadDocumentState(oldSelectionInfo);
				return true;
			}
		}
		
		return false;
	};

	/**
	 * Adds a footnote cross-reference to the current paragraph.
	 * <note>Please note that this paragraph must be in the document.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {footnoteRefTo} sRefType - The text or numeric value of a footnote reference you want to insert.
	 * @param {ApiParagraph} oParaTo - The first paragraph from a footnote to be referred to (must be in the document).
	 * @param {boolean} [bLink=true] - Specifies if the reference will be inserted as a hyperlink.
	 * @param {boolean} [bAboveBelow=false] - Specifies if the above/below words indicating the position of the reference should be included (don't used with the "aboveBelow" sRefType).
	 * @returns {boolean}
	 */
	ApiParagraph.prototype.AddFootnoteCrossRef = function(sRefTo, oParaTo, bLink, bAboveBelow) 
	{
		var nRefTo = -1;
		switch (sRefTo)
		{
			case "footnoteNum":
				nRefTo = 8;
				break;
			case "pageNum":
				nRefTo = 1;
				break;
			case "aboveBelow":
				nRefTo = 5;
				break;
			case "formFootnoteNum":
				nRefTo = 9;
				break;
		}
		if (nRefTo === -1 || !(oParaTo instanceof ApiParagraph) || !this.Paragraph.IsUseInDocument())
			return false;
		if (typeof(bLink) !== "boolean")
			bLink = true;
		if (typeof(bAboveBelow) !== "boolean" || nRefTo === 5)
			bAboveBelow = false;

		var oDocument = private_GetLogicDocument();
		var oldSelectionInfo;

		var aFirstFootnoteParas = oDocument.GetFootNotesFirstParagraphs();
		for (var nPara = 0; nPara < aFirstFootnoteParas.length; nPara++)
		{
			if (aFirstFootnoteParas[nPara].Id === oParaTo.Paragraph.Id)
			{
				oldSelectionInfo = oDocument.SaveDocumentState();
				oDocument.RemoveSelection();
				this.Paragraph.Document_SetThisElementCurrent();
				this.Paragraph.SetCurrentPos(this.Paragraph.Content.length - 1);
				oDocument.AddNoteRefToParagraph(oParaTo.private_GetImpl(), nRefTo, bLink, bAboveBelow);
				oDocument.LoadDocumentState(oldSelectionInfo);
				return true;
			}
		}
		return false;
	};

	/**
	 * Adds an endnote cross-reference to the current paragraph.
	 * <note>Please note that this paragraph must be in the document.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {endnoteRefTo} sRefType - The text or numeric value of an endnote reference you want to insert.
	 * @param {ApiParagraph} oParaTo - The first paragraph from an endnote to be referred to (must be in the document).
	 * @param {boolean} [bLink=true] - Specifies if the reference will be inserted as a hyperlink.
	 * @param {boolean} [bAboveBelow=false] - Specifies if the above/below words indicating the position of the reference should be included (don't used with the "aboveBelow" sRefType).
	 * @returns {boolean}
	 */
	ApiParagraph.prototype.AddEndnoteCrossRef = function(sRefTo, oParaTo, bLink, bAboveBelow) 
	{
		var nRefTo = -1;
		switch (sRefTo)
		{
			case "endnoteNum":
				nRefTo = 8;
				break;
			case "pageNum":
				nRefTo = 1;
				break;
			case "aboveBelow":
				nRefTo = 5;
				break;
			case "formEndnoteNum":
				nRefTo = 9;
				break;
		}
		if (nRefTo === -1 || !(oParaTo instanceof ApiParagraph) || !this.Paragraph.IsUseInDocument())
			return false;
		if (typeof(bLink) !== "boolean")
			bLink = true;
		if (typeof(bAboveBelow) !== "boolean" || nRefTo === 5)
			bAboveBelow = false;

		var oDocument = private_GetLogicDocument();
		var aFirstEndnoteParas = oDocument.GetEndNotesFirstParagraphs();
		var oldSelectionInfo;
		
		for (var nPara = 0; nPara < aFirstEndnoteParas.length; nPara++)
		{
			if (aFirstEndnoteParas[nPara].Id === oParaTo.Paragraph.Id)
			{
				oldSelectionInfo = oDocument.SaveDocumentState();
				oDocument.RemoveSelection();
				this.Paragraph.Document_SetThisElementCurrent();
				this.Paragraph.SetCurrentPos(this.Paragraph.Content.length - 1);
				oDocument.AddNoteRefToParagraph(oParaTo.private_GetImpl(), nRefTo, bLink, bAboveBelow);
				oDocument.LoadDocumentState(oldSelectionInfo);
				return true;
			}
		}
		return false;
	};

	/**
	 * Adds a caption cross-reference to the current paragraph.
	 * <note>Please note that this paragraph must be in the document.</note>
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {CaptionLabel | string} sCaption - The caption label ("Equation", "Figure", "Table", or another caption label).
	 * @param {captionRefTo} sRefType - The text or numeric value of a caption reference you want to insert.
	 * @param {ApiParagraph} oParaTo - The caption paragraph to be referred to (must be in the document).
	 * @param {boolean} [bLink=true] - Specifies if the reference will be inserted as a hyperlink.
	 * @param {boolean} [bAboveBelow=false] - Specifies if the above/below words indicating the position of the reference should be included (used only with the "pageNum" sRefType).
	 * @returns {boolean}
	 */
	ApiParagraph.prototype.AddCaptionCrossRef = function(sCaption, sRefTo, oParaTo, bLink, bAboveBelow) 
	{
		var nRefTo = -1;
		if (typeof(sCaption) !== "string" || sCaption.length === 0)
			return false;

		switch (sRefTo)
		{
			case "entireCaption":
				nRefTo = 0;
				break;
			case "labelNumber":
				nRefTo = 6;
				break;
			case "captionText":
				nRefTo = 7;
				break;
			case "pageNum":
				nRefTo = 1;
				break;
			case "aboveBelow":
				nRefTo = 5;
				break;
		}
		if (nRefTo === -1 || !(oParaTo instanceof ApiParagraph) || !this.Paragraph.IsUseInDocument())
			return false;
		if (typeof(bLink) !== "boolean")
			bLink = true;
		if (typeof(bAboveBelow) !== "boolean" || nRefTo !== 1)
			bAboveBelow = false;

		var aTempCompFlds = [];
		oParaTo = oParaTo.private_GetImpl();
		
		oParaTo.GetAllSeqFieldsByType(sCaption, aTempCompFlds);
		if (aTempCompFlds.length === 0)
			return false;
		if (nRefTo === 7 && oParaTo.CanAddRefAfterSEQ(sCaption) === false)
		{
			console.log("The request reference is empty.");
			return false;
		}

		var oDocument = private_GetLogicDocument();
		var oldSelectionInfo = oDocument.SaveDocumentState();
		oDocument.RemoveSelection();
		this.Paragraph.Document_SetThisElementCurrent();
		this.Paragraph.SetCurrentPos(this.Paragraph.Content.length - 1);
		oDocument.AddRefToCaption(sCaption, oParaTo, nRefTo, bLink, bAboveBelow);
		oDocument.LoadDocumentState(oldSelectionInfo);
		return true;
	};

	/**
	 * Converts the ApiParagraph object into the JSON object.
	 * @memberof ApiParagraph
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteNumberings - Specifies if the used numberings will be written to the JSON object or not.
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiParagraph.prototype.ToJSON = function(bWriteNumberings, bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerParagraph(this.Paragraph);
		if (bWriteNumberings)
			oJSON["numbering"] = oWriter.jsonWordNumberings;
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};

	/**
     * Returns the paragraph position within its parent element.
     * @memberof ApiParagraph
     * @typeofeditors ["CDE"]
     * @returns {Number} - returns -1 if the paragraph parent doesn't exist. 
     */
	ApiParagraph.prototype.GetPosInParent = function()
	{
		return this.Paragraph.GetIndex();
	};

	/**
     * Replaces the current paragraph with a new element.
     * @memberof ApiParagraph
     * @typeofeditors ["CDE"]
     * @param {DocumentElement} oElement - The element to replace the current paragraph with.
     * @returns {boolean}
     */
	ApiParagraph.prototype.ReplaceByElement = function(oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;

			var oParent = this.Paragraph.GetParent();
			var nParaPos = this.Paragraph.GetIndex();
			if (oParent && nParaPos !== -1)
			{
				this.Delete();
				oParent.Internal_Content_Add(nParaPos, oElm);
				return true;
			}
		}

		return false;
	};

	/**
     * Adds a caption paragraph after (or before) the current paragraph.
	 * <note>Please note that the current paragraph must be in the document (not in the footer/header).
	 * And if the current paragraph is placed in a shape, then a caption is added after (or before) the parent shape.</note>
     * @memberof ApiParagraph
     * @typeofeditors ["CDE"]
     * @param {string} sAdditional - The additional text.
	 * @param {CaptionLabel | String} [sLabel="Table"] - The caption label.
	 * @param {boolean} [bExludeLabel=false] - Specifies whether to exclude the label from the caption.
	 * @param {CaptionNumberingFormat} [sNumberingFormat="Arabic"] - The possible caption numbering format.
	 * @param {boolean} [bBefore=false] - Specifies whether to insert the caption before the current paragraph (true) or after (false) (after/before the shape if it is placed in the shape).
	 * @param {Number} [nHeadingLvl=undefined] - The heading level (used if you want to specify the chapter number).
	 * <note>If you want to specify "Heading 1", then nHeadingLvl === 0 and etc.</note>
	 * @param {CaptionSep} [sCaptionSep="hyphen"] - The caption separator (used if you want to specify the chapter number).
     * @returns {boolean}
     */
	ApiParagraph.prototype.AddCaption = function(sAdditional, sLabel, bExludeLabel, sNumberingFormat, bBefore, nHeadingLvl, sCaptionSep)
	{
		var oParaParent = this.Paragraph.GetParent();
		if (this.Paragraph.IsUseInDocument() === false || !oParaParent || oParaParent.Is_TopDocument(true) !== private_GetLogicDocument())
			return false;
		if (typeof(sAdditional) !== "string" || sAdditional.trim() === "")
			sAdditional = "";
		if (typeof(bExludeLabel) !== "boolean")
			bExludeLabel = false;
		if (typeof(bBefore) !== "boolean")
			bBefore = false;
		if (typeof(sLabel) !== "string" || sLabel.trim() === "")
			sLabel = "Table";
		
		let oCapPr = new Asc.CAscCaptionProperties();
		let oDoc = private_GetLogicDocument();

		let nNumFormat;
		switch (sNumberingFormat)
		{
			case "ALPHABETIC":
				nNumFormat = Asc.c_oAscNumberingFormat.UpperLetter;
				break;
			case "alphabetic":
				nNumFormat = Asc.c_oAscNumberingFormat.LowerLetter;
				break;
			case "Roman":
				nNumFormat = Asc.c_oAscNumberingFormat.UpperRoman;
				break;
			case "roman":
				nNumFormat = Asc.c_oAscNumberingFormat.LowerRoman;
				break;
			default:
				nNumFormat = Asc.c_oAscNumberingFormat.Decimal;
				break;
		}
		switch (sCaptionSep)
		{
			case "hyphen":
				sCaptionSep = "-";
				break;
			case "period":
				sCaptionSep = ".";
				break;
			case "colon":
				sCaptionSep = ":";
				break;
			case "longDash":
				sCaptionSep = "—";
				break;
			case "dash":
				sCaptionSep = "-";
				break;
			default:
				sCaptionSep = "-";
				break;
		}

		oCapPr.Label = sLabel;
		oCapPr.Before = bBefore;
		oCapPr.ExcludeLabel = bExludeLabel;
		oCapPr.NumFormat = nNumFormat;
		oCapPr.Separator = sCaptionSep;
		oCapPr.Additional = sAdditional;

		if (nHeadingLvl >= 0 && nHeadingLvl <= 8)
		{
			oCapPr.HeadingLvl = nHeadingLvl;
			oCapPr.IncludeChapterNumber = true;
		}
		else oCapPr.HeadingLvl = 0;

		this.Paragraph.SetThisElementCurrent();

		oDoc.AddCaption(oCapPr);
		return true;
	};

	/**
     * Returns the paragraph section.
     * @memberof ApiParagraph
     * @typeofeditors ["CDE"]
     * @returns {ApiSection}
     */
	ApiParagraph.prototype.GetSection = function()
	{
		let oSectPr = this.private_GetImpl().Get_SectPr();
		if (!oSectPr)
			return null;

		return new ApiSection(oSectPr);
	};
	/**
     * Sets the specified section to the current paragraph.
     * @memberof ApiParagraph
     * @typeofeditors ["CDE"]
     * @param {ApiSection} oSection - The section which will be set to the paragraph.
     * @returns {boolean}
     */
	ApiParagraph.prototype.SetSection = function(oSection)
	{
		if (typeof(oSection) != "object" || !(oSection instanceof ApiSection))
			return false;

		if (!this.Paragraph.CanAddSectionPr()) {
			console.error(new Error('Paragraph must be in a document.'));
			return false;
		}

		let oDoc = private_GetLogicDocument();
		if (!oDoc)
			return false;

		let oSectPr = oSection.Section;
		let nContentPos = this.Paragraph.GetIndex();
		let oCurSectPr = oDoc.SectionsInfo.Get_SectPr(nContentPos).SectPr;

		oCurSectPr.Set_Type(oSectPr.Type);
		oCurSectPr.Set_PageNum_Start(-1);
		oCurSectPr.Clear_AllHdrFtr();

		this.private_GetImpl().Set_SectionPr(oSectPr);
		return true;
	};
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRun
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiRun class.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"run"}
	 */
	ApiRun.prototype.GetClassType = function()
	{
		return "run";
	};
	/**
	 * Returns the text properties of the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.TextPr);
	};
	/**
	 * Clears the content from the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.ClearContent = function()
	{
		this.Run.Remove_FromContent(0, this.Run.Content.length);
	};
	/**
	 * Removes all the elements from the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.RemoveAllElements = function()
	{
		this.Run.Remove_FromContent(0, this.Run.Content.length);
	};
	/**
	 * Deletes the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.Delete = function()
	{
		var oParent = this.Run.Get_Parent();
		var nPosInParent = this.Run.GetPosInParent();

		if (nPosInParent !== - 1)
		{
			this.Run.PreDelete();
			oParent.Remove_FromContent(nPosInParent, 1, true);
		}
		else 
			return false;
	};
	/**
	 * Adds some text to the current run.
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
	 * Adds a page break and starts the next element from a new page.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 */
	ApiRun.prototype.AddPageBreak = function()
	{
		this.Run.Add_ToContent(this.Run.Content.length, new AscWord.CRunBreak(AscWord.break_Page));
	};
	/**
	 * Adds a line break to the current run position and starts the next element from a new line.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.AddLineBreak = function()
	{
		this.Run.Add_ToContent(this.Run.Content.length, new AscWord.CRunBreak(AscWord.break_Line));
	};
	/**
	 * Adds a column break to the current run position and starts the next element from a new column.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 */
	ApiRun.prototype.AddColumnBreak = function()
	{
		this.Run.Add_ToContent(this.Run.Content.length, new AscWord.CRunBreak(AscWord.break_Column));
	};
	/**
	 * Adds a tab stop to the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 */
	ApiRun.prototype.AddTabStop = function()
	{
		this.Run.Add_ToContent(this.Run.Content.length, new AscWord.CRunTab());
	};
	/**
	 * Adds a drawing object (image, shape or chart) to the current text run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @param {ApiDrawing} oDrawing - The object which will be added to the current run.
	 * @returns {boolean} - returns false if param is invalid.
	 */ 
	ApiRun.prototype.AddDrawing = function(oDrawing)
	{
		if (!(oDrawing instanceof ApiDrawing))
			return false;

		this.Run.Add_ToContent(this.Run.Content.length, oDrawing.Drawing);
		oDrawing.Drawing.Set_Parent(this.Run);
		private_CheckDrawingOnAdd(oDrawing);
		return true;
	};
	/**
	 * Selects the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @return {boolean}
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

		var controllerType;

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
	 * Adds a hyperlink to the current run. 
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

		var oHyperlink;
		var hyperlinkPr	= new Asc.CHyperlinkProperty();
		var urlType		= AscCommon.getUrlType(sLink);
		if (!AscCommon.rx_allowedProtocols.test(sLink))
			sLink = (urlType === 0) ? null :(( (urlType === 2) ? 'mailto:' : 'http://' ) + sLink);
		sLink = sLink && sLink.replace(new RegExp("%20",'g')," ");
		hyperlinkPr.put_Value(sLink);
		hyperlinkPr.put_ToolTip(sScreenTipText);
		hyperlinkPr.put_Bookmark(null);

		parentPara.Selection.Use = true;
		oHyperlink = new ApiHyperlink(parentPara.AddHyperlink(hyperlinkPr));
		Document.RemoveSelection();

		return oHyperlink;
	};
	/**
	 * Creates a copy of the current run.
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
	 * Returns a Range object that represents the part of the document contained in the specified run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiRun.prototype.GetRange = function(Start, End)
	{
		return new ApiRange(this.Run, Start, End);
	};

	/**
     * Returns a content control that contains the current run.
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
                if (RunPosition[Index].Class.Parent && RunPosition[Index].Class.Parent instanceof CBlockLevelSdt)
                    return new ApiBlockLvlSdt(RunPosition[Index].Class);
                else if (RunPosition[Index].Class instanceof CInlineLevelSdt)
                    return new ApiInlineLvlSdt(RunPosition[Index].Class);
            }
        }

        return null;
    };
    /**
     * Returns a table that contains the current run.
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
     * Returns a table cell that contains the current run.
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
	 * Sets the text properties to the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ApiTextPr} oTextPr - The text properties that will be set to the current run.
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
	 * Sets the bold property to the text character.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isBold - Specifies that the contents of the current run are displayed bold.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetBold = function(isBold)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetBold(isBold);
		
		return oTextPr;
	};
	/**
	 * Specifies that any lowercase characters in the current text run are formatted for display only as their capital letter character equivalents.
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
	 * Sets the text color for the current text run in the RGB format.
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
	 * Specifies that the contents of the current run are displayed with two horizontal lines through each character displayed on the line.
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
	 * Sets the text color to the current text run.
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
	 * Sets all 4 font slots with the specified font family.
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
	 * Returns all font names from all elements inside the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {string[]} - The font names used for the current run.
	 */
	ApiRun.prototype.GetFontNames = function()
	{
		let fontMap = {};
		let arrFonts = [];
		// this.Run.Get_AllFontNames(fontMap);
		this.Run.Get_CompiledTextPr().Document_Get_AllFontNames(fontMap);
		for (let key in fontMap)
		{
			arrFonts.push(key);
		}
		return arrFonts;
	};
	/**
	 * Sets the font size to the characters of the current text run.
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
	 * Specifies a highlighting color which is applied as a background to the contents of the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {highlightColor} sColor - Available highlight color.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetHighlight = function(sColor)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetHighlight(sColor);
		
		return oTextPr;
	};
	/**
	 * Sets the italic property to the text character.
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
	 * Specifies the languages which will be used to check spelling and grammar (if requested) when processing
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
	 * Specifies an amount by which text is raised or lowered for this run in relation to the default
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
	 * Specifies the shading applied to the contents of the current text run.
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
	 * Specifies that all the small letter characters in this text run are formatted for display only as their capital
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
	 * Sets the text spacing measured in twentieths of a point.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nSpacing - The value of the text spacing measured in twentieths of a point (1/1440 of an inch).
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetSpacing = function(nSpacing)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetSpacing(nSpacing);
		
		return oTextPr;
	};
	/**
	 * Specifies that the contents of the current run are displayed with a single horizontal line through the center of the line.
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
	 * Sets a style to the current run.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {ApiStyle} oStyle - The style which must be applied to the text run.
	 * @returns {ApiTextPr}
	 */
	ApiRun.prototype.SetStyle = function(oStyle)
	{
		var oTextPr = this.GetTextPr();
		oTextPr.SetStyle(oStyle);
		
		return oTextPr;
	};
	/**
	 * Specifies that the contents of the current run are displayed along with a line appearing directly below the character
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
	 * Specifies the alignment which will be applied to the contents of the current run in relation to the default appearance of the text run:
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
	 * Wraps a run in a mail merge field.
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
	 * Converts the ApiRun object into the JSON object.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiRun.prototype.ToJSON = function(bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerParaRun(this.Run);
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};

	/**
	 * Adds a comment to the current run.
	 * <note>Please note that this run must be in the document.</note>
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment text (required).
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {ApiComment?} - Returns null if the comment was not added.
	 */
	ApiRun.prototype.AddComment = function(sText, sAuthor)
	{
		if (!sText || typeof(sText) !== "string")
			return null;
		if (typeof(sAuthor) !== "string")
			sAuthor = "";

		if (!this.Run.IsUseInDocument())
			return null;

		var oDocument = private_GetLogicDocument();
		var CommentData = new AscCommon.CCommentData();

		CommentData.SetText(sText);
		if (sAuthor !== "")
			CommentData.SetUserName(sAuthor);

		var oDocumentState = oDocument.SaveDocumentState();
		this.Run.SelectThisElement();

		let comment = AddCommentToDocument(oDocument, CommentData);
		oDocument.LoadDocumentState(oDocumentState);
		oDocument.UpdateSelection();
		return comment;
	};

	/**
	 * Returns a text from the text run.
	 * @memberof ApiRun
	 * @param {object} oPr - The resulting string display properties.
     * @param {string} [oPr.NewLineSeparator='\r'] - Defines how the line separator will be specified in the resulting string.
	 * @param {string} [oPr.TabSymbol='\t'] - Defines how the tab will be specified in the resulting string.
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */	
	ApiRun.prototype.GetText = function(oPr)
	{
		if (!oPr) {
			oPr = {};
		}

		let oProp = {
			Text: "",
			NewLineSeparator:	(oPr.hasOwnProperty("NewLineSeparator")) ? oPr["NewLineSeparator"] : "\r",
			TabSymbol:			oPr["TabSymbol"],
			ParaSeparator:		oPr["ParaSeparator"]
		}

		return this.Run.GetText(oProp);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiSection
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiSection class.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @returns {"section"}
	 */
	ApiSection.prototype.GetClassType = function()
	{
		return "section";
	};
	/**
	 * Specifies a type of the current section. The section type defines how the contents of the current 
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
	 * Specifies that all the text columns in the current section are of equal width.
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
	 * Specifies that all the columns in the current section have the different widths. Number of columns is equal 
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
	 * Specifies the properties (size and orientation) for all the pages in the current section.
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
	 * Specifies the page margins for all the pages in this section.
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
	 * Specifies the distance from the top edge of the page to the top edge of the header.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {twips} nDistance - The distance from the top edge of the page to the top edge of the header measured in twentieths of a point (1/1440 of an inch).
	 */
	ApiSection.prototype.SetHeaderDistance = function(nDistance)
	{
		this.Section.SetPageMarginHeader(private_Twips2MM(nDistance));
	};
	/**
	 * Specifies the distance from the bottom edge of the page to the bottom edge of the footer.
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
	 * Returns the content for the specified header type.
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
	 * Removes the header of the specified type from the current section. After removal, the header will be inherited from
	 * the previous section, or if this is the first section in the document, no header of the specified type will be presented.
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
	 * Returns the content for the specified footer type.
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

		if (oFooter)
			return new ApiDocumentContent(oFooter.Get_DocumentContent());
		
		return null;
	};
	/**
	 * Removes the footer of the specified type from the current section. After removal, the footer will be inherited from 
	 * the previous section, or if this is the first section in the document, no footer of the specified type will be presented.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {HdrFtrType} sType - Footer type to be removed.
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
	 * Specifies whether the current section in this document has the different header and footer for the section first page.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isTitlePage - If true, the first page of the section will have header and footer that will differ from the other pages of the same section.
	 */
	ApiSection.prototype.SetTitlePage = function(isTitlePage)
	{
		this.Section.Set_TitlePage(private_GetBoolean(isTitlePage));
	};
	/**
	 * Returns the next section if exists.
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
	 * Returns the previous section if exists.
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
	/**
	 * Converts the ApiSection object into the JSON object.
	 * @memberof ApiSection
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteNumberings - Specifies if the used numberings will be written to the JSON object or not.
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiSection.prototype.ToJSON = function(bWriteNumberings, bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerSectionPr(this.Section);
		if (bWriteNumberings)
			oJSON["numbering"] = oWriter.jsonWordNumberings;
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTable
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiTable class.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @returns {"table"}
	 */
	ApiTable.prototype.GetClassType = function()
	{
		return "table";
	};
	/**
	 * Returns a number of rows in the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiTable.prototype.GetRowsCount = function()
	{
		return this.Table.Content.length;
	};
	/**
	 * Returns a table row by its position in the table.
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
	 * Returns a cell by its position.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {number} nRow - The row position in the current table where the specified cell is placed.
	 * @param {number} nCell - The cell position in the current table.
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
	 * Merges an array of cells. If the merge is done successfully, it will return the resulting merged cell, otherwise the result will be "null".
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
					if (oCurPos.Cell >= oPos.Cell)
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
	 * Sets a style to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiStyle} oStyle - The style which will be applied to the current table.
	 * @returns {boolean} - returns false if param is invalid.
	 */
	ApiTable.prototype.SetStyle = function(oStyle)
	{
		if (!oStyle || !(oStyle instanceof ApiStyle) || styletype_Table !== oStyle.Style.Get_Type())
			return false;

		this.Table.Set_TableStyle(oStyle.Style.Get_Id(), true);

		return true;
	};
	/**
	 * Specifies the conditional formatting components of the referenced table style (if one exists) 
	 * which will be applied to the set of table rows with the current table-level property exceptions. A table style 
	 * can specify up to six different optional conditional formats, for example, different formatting for the first column, 
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
		var oTableLook = new AscCommon.CTableLook(private_GetBoolean(isFirstColumn),
			private_GetBoolean(isFirstRow),
			private_GetBoolean(isLastColumn),
			private_GetBoolean(isLastRow),
			private_GetBoolean(isHorBand),
			private_GetBoolean(isVerBand));
		this.Table.Set_TableLook(oTableLook);
	};
	/**
	 * Splits the cell into a given number of rows and columns.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell which will be split.
	 * @param {Number} [nRow=1] - Count of rows into which the cell will be split.
	 * @param {Number} [nCol=1] - Count of columns into which the cell will be split.
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
	 * Adds a new row to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell after which a new row will be added. If not specified, a new row will
	 * be added at the end of the table.
	 * @param {boolean} [isBefore=false] - Adds a new row before (false) or after (true) the specified cell. If no cell is specified, then
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
	 * Adds the new rows to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell after which the new rows will be added. If not specified, the new rows will
	 * be added at the end of the table.
	 * @param {Number} nCount - Count of rows to be added.
	 * @param {boolean} [isBefore=false] - Adds the new rows before (false) or after (true) the specified cell. If no cell is specified, then
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
	 * Adds a new column to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell after which a new column will be added. If not specified, a new column will be added at the end of the table.
	 * @param {boolean} [isBefore=false] - Adds a new column before (false) or after (true) the specified cell. If no cell is specified,
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
	 * Adds the new columns to the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} [oCell] - The cell after which the new columns will be added. If not specified, the new columns will be added at the end of the table.
	 * @param {Number} nCount - Count of columns to be added.
	 * @param {boolean} [isBefore=false] - Adds the new columns before (false) or after (true) the specified cell. If no cell is specified,
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
	 * Adds a paragraph or a table or a blockLvl content control using its position in the cell.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {ApiTableCell} oCell - The cell where the specified element will be added.
	 * @param {number} nPos - The position in the cell where the specified element will be added.
	 * @param {DocumentElement} oElement - The document element which will be added at the current position.
	 */
	ApiTable.prototype.AddElement = function(oCell, nPos, oElement)
	{
		if (!(oCell instanceof ApiTableCell) || this.Table !== oCell.Cell.Row.Table)
			return false;

		var apiCellContent = oCell.GetContent();

		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;
			apiCellContent.Document.Internal_Content_Add(nPos, oElm);

			return true;
		}

		return false;
	};
	/**
	 * Removes a table row with a specified cell.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} oCell - The cell which is placed in the row that will be removed.
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
	 * Removes a table column with a specified cell.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCell} oCell - The cell which is placed in the column that will be removed.
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
	 * Creates a copy of the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE", "CPE"]
	 * @returns {ApiTable}
	 */
	ApiTable.prototype.Copy = function()
	{
		var oTable = this.Table.Copy(private_GetLogicDocument(), private_GetDrawingDocument());
		return new ApiTable(oTable);
	};
	/**
	 * Selects the current table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE", "CPE"]
	 * @returns {boolean}
	 */
	ApiTable.prototype.Select = function()
	{
		var Document = private_GetLogicDocument();
		
		var DocPos = this.Table.GetDocumentPositionFromObject();
		
		if (DocPos[0].Position === - 1)
			return false;

		var controllerType;

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
	 * Returns a Range object that represents the part of the document contained in the specified table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiTable.prototype.GetRange = function(Start, End)
	{
		return new ApiRange(this.Table, Start, End);
	};
	/**
     * Sets the horizontal alignment to the table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @param {String} sType - Horizontal alignment type: may be "left" or "center" or "right".
     * @returns {boolean} - returns false if param is invalid.
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
     * Sets the vertical alignment to the table.
     * @typeofeditors ["CDE"]
     * @param {String} sType - Vertical alignment type: may be "top" or "center" or "bottom".
     * @returns {boolean} - returns false if param is invalid.
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
     * Sets the table paddings.
	 * If table is inline, then only left padding is applied.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @param {Number} nLeft - Left padding.
	 * @param {Number} nTop - Top padding.
	 * @param {Number} nRight - Right padding.
	 * @param {Number} nBottom - Bottom padding.
     * @returns {boolean} - returns true.
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
     * Sets the table wrapping style.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @param {boolean} isFlow - Specifies if the table is inline or not.
	 * @returns {boolean} - returns false if param is invalid.
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
     * Returns a content control that contains the current table.
     * @memberof ApiTable
	 * @typeofeditors ["CDE"]
     * @return {ApiBlockLvlSdt | null} - return null is parent content control doesn't exist.
     */
    ApiTable.prototype.GetParentContentControl = function()
    {
        var TablePosition = this.Table.GetDocumentPositionFromObject();

        for (var Index = TablePosition.length - 1; Index >= 1; Index--)
        {
            if (TablePosition[Index].Class && TablePosition[Index].Class.Parent && TablePosition[Index].Class.Parent instanceof CBlockLevelSdt)
				return new ApiBlockLvlSdt(TablePosition[Index].Class.Parent);
        }

        return null;
	};
	/**
	 * Wraps the current table object with a content control.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {number} nType - Defines if this method returns the ApiBlockLvlSdt (nType === 1) or ApiTable (any value except 1) object.
	 * @return {ApiTable | ApiBlockLvlSdt}  
	 */
	ApiTable.prototype.InsertInContentControl = function(nType)
	{
		var Document = private_GetLogicDocument();

		var ContentControl;

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
     * Returns a table that contains the current table.
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
     * Returns the tables that contain the current table.
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
			if (this.Table.Pages[nCurPage].FirstRow < 0 || this.Table.Pages[nCurPage].LastRow < 0)
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
     * Returns the next table if exists.
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
				}
				else 
					return new ApiTable(curPageTables[Index].Table);
			}
		}
		
		return null; 
	};
	/**
     * Returns the previous table if exists.
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
						return new ApiTable(curPageTables[Index - 1].Table)
				}
				else 
					return new ApiTable(curPageTables[Index].Table);
			}
		}
		
		return null; 
    };
    /**
     * Returns a table cell that contains the current table.
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
	 * Deletes the current table. 
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @return {boolean} - returns false if parent of table doesn't exist.
	 */
	ApiTable.prototype.Delete = function()
	{
		var oParent = this.Table.Parent;
		var nTablePos = this.Table.GetIndex();
		if (nTablePos !== -1)
		{
			this.Table.PreDelete();
			oParent.Remove_FromContent(nTablePos, 1, true);

			return true;
		}
		else 	 
			return false;
	};
	/**
	 * Clears the content from the table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @return {boolean} - returns true.
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
	 * Searches for a scope of a table object. The search results are a collection of ApiRange objects.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {boolean} isMatchCase - Case sensitive or not.
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
	 * Applies the text settings to the entire contents of the table.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The text properties that will be set to the current table.
	 * @return {boolean} - returns true.
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
	/**
	 * Sets the background color to all cells in the current table.
	 * @memberof ApiTable
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} bNone - Defines that background color will not be set.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiTable.prototype.SetBackgroundColor = function(r, g, b, bNone)
	{
		if ((typeof(r) == "number" && typeof(g) == "number" && typeof(b) == "number" && !bNone) || bNone)
		{
			var oRow;
			for (var nRow = 0, nCount = this.GetRowsCount(); nRow < nCount; nRow++)
			{
				oRow = this.GetRow(nRow);
				oRow.SetBackgroundColor(r, g, b, bNone);
			}
			return true;
		}
		else
			return false;
	};
	
	/**
	 * Converts the ApiTable object into the JSON object.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteNumberings - Specifies if the used numberings will be written to the JSON object or not.
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiTable.prototype.ToJSON = function(bWriteNumberings, bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerTable(this.Table);
		if (bWriteNumberings)
			oJSON["numbering"] = oWriter.jsonWordNumberings;
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};

	/**
     * Returns the table position within its parent element.
     * @memberof ApiTable
     * @typeofeditors ["CDE"]
     * @returns {Number} - returns -1 if the table parent doesn't exist. 
     */
	ApiTable.prototype.GetPosInParent = function()
	{
		return this.Table.GetIndex();
	};
 
	/**
	 * Replaces the current table with a new element.
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {DocumentElement} oElement - The element to replace the current table with.
	 * @returns {boolean}
	 */
	ApiTable.prototype.ReplaceByElement = function(oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;

			var oParent = this.Table.GetParent();
			var nTablePos = this.Table.GetIndex();
			if (oParent && nTablePos !== -1)
			{
				this.Delete();
				oParent.Internal_Content_Add(nTablePos, oElm);
				return true;
			}
		}

		return false;
	};

	/**
	 * Adds a comment to all contents of the current table.
	 * <note>Please note that this table must be in the document.</note>
	 * @memberof ApiTable
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment text (required).
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {ApiComment?} - Returns null if the comment was not added.
	 */
	ApiTable.prototype.AddComment = function(sText, sAuthor)
	{
		if (!sText || typeof(sText) !== "string")
			return null;
		if (typeof(sAuthor) !== "string")
			sAuthor = "";

		if (!this.Table.IsUseInDocument())
			return null;

		var oDocument = private_GetLogicDocument();
		var CommentData = new AscCommon.CCommentData();

		CommentData.SetText(sText);
		if (sAuthor !== "")
			CommentData.SetUserName(sAuthor);

		var oDocumentState = oDocument.SaveDocumentState();
		this.Table.SelectAll();
		this.Table.Document_SetThisElementCurrent(true);

		let comment = AddCommentToDocument(oDocument, CommentData);
		oDocument.LoadDocumentState(oDocumentState);
		oDocument.UpdateSelection();
		return new ApiComment(comment)
	};

	/**
     * Adds a caption paragraph after (or before) the current table.
	 * <note>Please note that the current table must be in the document (not in the footer/header).
	 * And if the current table is placed in a shape, then a caption is added after (or before) the parent shape.</note>
     * @memberof ApiTable
     * @typeofeditors ["CDE"]
     * @param {string} sAdditional - The additional text.
	 * @param {CaptionLabel | String} [sLabel="Table"] - The caption label.
	 * @param {boolean} [bExludeLabel=false] - Specifies whether to exclude the label from the caption.
	 * @param {CaptionNumberingFormat} [sNumberingFormat="Arabic"] - The possible caption numbering format.
	 * @param {boolean} [bBefore=false] - Specifies whether to insert the caption before the current table (true) or after (false) (after/before the shape if it is placed in the shape).
	 * @param {Number} [nHeadingLvl=undefined] - The heading level (used if you want to specify the chapter number).
	 * <note>If you want to specify "Heading 1", then nHeadingLvl === 0 and etc.</note>
	 * @param {CaptionSep} [sCaptionSep="hyphen"] - The caption separator (used if you want to specify the chapter number).
     * @returns {boolean}
     */
	ApiTable.prototype.AddCaption = function(sAdditional, sLabel, bExludeLabel, sNumberingFormat, bBefore, nHeadingLvl, sCaptionSep)
	{
		var oTableParent = this.Table.GetParent();
		if (this.Table.IsUseInDocument() === false || !oTableParent || oTableParent.Is_TopDocument(true) !== private_GetLogicDocument())
			return false;
		if (typeof(sAdditional) !== "string" || sAdditional.trim() === "")
			sAdditional = "";
		if (typeof(bExludeLabel) !== "boolean")
			bExludeLabel = false;
		if (typeof(bBefore) !== "boolean")
			bBefore = false;
		if (typeof(sLabel) !== "string" || sLabel.trim() === "")
			sLabel = "Table";
		
		let oCapPr = new Asc.CAscCaptionProperties();
		let oDoc = private_GetLogicDocument();

		let nNumFormat;
		switch (sNumberingFormat)
		{
			case "ALPHABETIC":
				nNumFormat = Asc.c_oAscNumberingFormat.UpperLetter;
				break;
			case "alphabetic":
				nNumFormat = Asc.c_oAscNumberingFormat.LowerLetter;
				break;
			case "Roman":
				nNumFormat = Asc.c_oAscNumberingFormat.UpperRoman;
				break;
			case "roman":
				nNumFormat = Asc.c_oAscNumberingFormat.LowerRoman;
				break;
			default:
				nNumFormat = Asc.c_oAscNumberingFormat.Decimal;
				break;
		}
		switch (sCaptionSep)
		{
			case "hyphen":
				sCaptionSep = "-";
				break;
			case "period":
				sCaptionSep = ".";
				break;
			case "colon":
				sCaptionSep = ":";
				break;
			case "longDash":
				sCaptionSep = "—";
				break;
			case "dash":
				sCaptionSep = "-";
				break;
			default:
				sCaptionSep = "-";
				break;
		}

		oCapPr.Label = sLabel;
		oCapPr.Before = bBefore;
		oCapPr.ExcludeLabel = bExludeLabel;
		oCapPr.NumFormat = nNumFormat;
		oCapPr.Separator = sCaptionSep;
		oCapPr.Additional = sAdditional;

		if (nHeadingLvl >= 0 && nHeadingLvl <= 8)
		{
			oCapPr.HeadingLvl = nHeadingLvl;
			oCapPr.IncludeChapterNumber = true;
		}
		else oCapPr.HeadingLvl = 0;

		this.Table.Document_SetThisElementCurrent();

		oDoc.AddCaption(oCapPr);
		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableRow
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiTableRow class.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {"tableRow"}
	 */
	ApiTableRow.prototype.GetClassType = function()
	{
		return "tableRow";
	};
	/**
	 * Returns a number of cells in the current row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiTableRow.prototype.GetCellsCount = function()
	{
		return this.Row.Content.length;
	};
	/**
	 * Returns a cell by its position.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The cell position in the current row.
	 * @returns {ApiTableCell}
	 */
	ApiTableRow.prototype.GetCell = function(nPos)
	{
		if (nPos < 0 || nPos >= this.Row.Content.length)
			return null;

		return new ApiTableCell(this.Row.Content[nPos]);
	};
	/**
	 * Returns the current row index.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {Number}
	 */
	ApiTableRow.prototype.GetIndex = function()
	{
		return this.Row.GetIndex();
	};
	/**
	 * Returns the parent table of the current row.
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
	 * Returns the next row if exists.
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
	 * Returns the previous row if exists.
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
	 * Adds the new rows to the current table.
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
	 * Merges the cells in the current row. 
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
			tempVMergeCount		= oTable.Table.Internal_GetVertMergeCountUp(this.GetIndex(), tempStartGridCol, tempGridSpan);

			if (tempVMergeCount > 1)
			{
				tempCell = new ApiTableCell(oTable.Table.GetCellByStartGridCol(this.GetIndex() - (tempVMergeCount - 1), tempStartGridCol));
			}

			cellsArr.push(tempCell);
		}
			
		return oTable.MergeCells(cellsArr);
	};
	/**
	 * Clears the content from the current row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {boolean} - returns false if parent table doesn't exist.
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
			tempVMergeCount		= oTable.Table.Internal_GetVertMergeCountUp(this.GetIndex(), tempStartGridCol, tempGridSpan);

			if (tempVMergeCount > 1)
			{
				tempCell = oTable.Table.GetCellByStartGridCol(this.GetIndex() - (tempVMergeCount - 1), tempStartGridCol);
			}

			tempCell.GetContent().Clear_Content();
		}

		return true;
	};
	/**
	 * Removes the current table row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @returns {boolean} - return false if parent table doesn't exist.
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
	 * Sets the text properties to the current row.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The text properties that will be set to the current row.
	 * @returns {boolean} - returns false if parent table doesn't exist or param is invalid.
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
			tempVMergeCount		= oTable.Table.Internal_GetVertMergeCountUp(this.GetIndex(), tempStartGridCol, tempGridSpan);

			if (tempVMergeCount > 1)
			{
				tempCell = new ApiTableCell(oTable.Table.GetCellByStartGridCol(this.GetIndex() - (tempVMergeCount - 1), tempStartGridCol));
			}

			tempCell.SetTextPr(oTextPr);
		}

		return true;
	};
	/**
	 * Searches for a scope of a table row object. The search results are a collection of ApiRange objects.
	 * @memberof ApiTableRow
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {boolean} isMatchCase - Case sensitive or not.
	 * @return {ApiRange[]}
	 */
	ApiTableRow.prototype.Search = function(sText, isMatchCase)
	{
		if (isMatchCase === undefined)
			isMatchCase	= false;
		var oTable = this.GetParentTable();
		if (!oTable)
			return [];

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
			tempVMergeCount		= oTable.Table.Internal_GetVertMergeCountUp(this.GetIndex(), tempStartGridCol, tempGridSpan);

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
	/**
	 * Sets the background color to all cells in the current table row.
	 * @memberof ApiTableRow
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} bNone - Defines that background color will not be set.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiTableRow.prototype.SetBackgroundColor = function(r, g, b, bNone)
	{
		if ((typeof(r) == "number" && typeof(g) == "number" && typeof(b) == "number" && !bNone) || bNone)
		{
			var oCell;
			for (var nCell = 0, nCount = this.GetCellsCount(); nCell < nCount; nCell++)
			{
				oCell = this.GetCell(nCell);
				oCell.SetBackgroundColor(r, g, b, bNone);
			}
			return true;
		}
		else
			return false;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableCell
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiTableCell class.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {"tableCell"}
	 */
	ApiTableCell.prototype.GetClassType = function()
	{
		return "tableCell";
	};
	/**
	 * Returns the current cell content.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {ApiDocumentContent}
	 */
	ApiTableCell.prototype.GetContent = function()
	{
		return new ApiDocumentContent(this.Cell.Content);
	};
	/**
	 * Returns the current cell index.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {Number}
	 */
	ApiTableCell.prototype.GetIndex = function()
	{
		return this.Cell.GetIndex();
	};
	/**
	 * Returns an index of the parent row.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiTableCell.prototype.GetRowIndex = function()
	{
		var Row = this.Cell.GetRow();
		if(!Row)
			return -1;

		return Row.GetIndex();
	};
	/**
	 * Returns a parent row of the current cell.
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
	 * Returns a parent table of the current cell.
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
	 * Adds the new rows to the current table.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {Number} nCount - Count of rows to be added.
	 * @param {boolean} [isBefore=false] - Specifies if the new rows will be added before or after the current cell. 
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
	 * Adds the new columns to the current table.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {Number} nCount - Count of columns to be added.
	 * @param {boolean} [isBefore=false] - Specifies if the new columns will be added before or after the current cell. 
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
	 * Removes a column containing the current cell.
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
	 * Removes a row containing the current cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @returns {boolean} Is the table empty after removing.
	 */
	ApiTableCell.prototype.RemoveRow = function()
	{
		var oTable = this.GetParentTable();
		if (!oTable)
			return false;

		return oTable.RemoveRow(this);
	};
	/**
	 * Searches for a scope of a table cell object. The search results are a collection of ApiRange objects.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {boolean} isMatchCase - Case sensitive or not.
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
	 * Returns the next cell if exists.
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
	 * Returns the previous cell if exists.
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
	 * Splits the cell into a given number of rows and columns.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {Number} [nRow=1] - Count of rows into which the cell will be split.
	 * @param {Number} [nCol=1] - Count of columns into which the cell will be split.
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
	 * Sets the cell properties to the current cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {ApiTableCellPr} oApiTableCellPr - The properties that will be set to the current table cell.
	 * @returns {boolean} - returns false if param is invalid.
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
	 * Applies the text settings to the entire contents of the current cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The properties that will be set to the current table cell text.
	 * @return {boolean} - returns false if param is invalid.
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
	 * Clears the content from the current cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @return {boolean} - returns false if parent row is invalid.
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
	 * Adds a paragraph or a table or a blockLvl content control using its position in the cell.
	 * @memberof ApiTableCell
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The position where the current element will be added.
	 * @param {DocumentElement} oElement - The document element which will be added at the current position.
	 * @returns {boolean} - returns false if oElement is invalid.
	 */
	ApiTableCell.prototype.AddElement = function(nPos, oElement)
	{
		var apiCellContent = this.GetContent();

		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;
			apiCellContent.Document.Internal_Content_Add(nPos, oElm);

			return true;
		}

		return false;
	};
	/**
	 * Sets the background color to the current table cell.
	 * @memberof ApiTableCell
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} bNone - Defines that background color will not be set.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiTableCell.prototype.SetBackgroundColor = function(r, g, b, bNone)
	{
		let oUnifill = new AscFormat.CUniFill();
		oUnifill.setFill(new AscFormat.CSolidFill());
		oUnifill.fill.setColor(new AscFormat.CUniColor());
		oUnifill.fill.color.setColor(new AscFormat.CRGBColor());

		if (r >=0 && g >=0 && b >=0)
			oUnifill.fill.color.color.setColor(r, g, b);
		else
			return false;

		var oNewShd = {
			Value : bNone ? Asc.c_oAscShd.Nil : Asc.c_oAscShd.Clear,
			Color : {
				r    : r,
				g    : g,
				b    : b,
				Auto : false
			},

			Fill    : {
				r    : r,
				g    : g,
				b    : b,
				Auto : false
			},
			Unifill   : oUnifill.createDuplicate(),
			ThemeFill : oUnifill.createDuplicate()
		}

		this.Cell.Set_Shd(oNewShd);
		return true;
	};
	/**
	 * Sets the background color to all cells in the column containing the current cell.
	 * @memberof ApiTableCell
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} bNone - Defines that background color will not be set.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiTableCell.prototype.SetColumnBackgroundColor = function(r, g, b, bNone)
	{
		if ((typeof(r) == "number" && typeof(g) == "number" && typeof(b) == "number" && !bNone) || bNone)
		{
			var oTable = this.GetParentTable();
			var aColumnCells = oTable.Table.GetColumn(this.GetIndex(), this.GetParentRow().GetIndex());
			var aCellsToFill = [];

			for (var nCell = 0; nCell < aColumnCells.length; nCell++)
				aCellsToFill[nCell] = new ApiTableCell(aColumnCells[nCell]);

			if (aCellsToFill.length > 0)
			{
				for (nCell = 0; nCell < aCellsToFill.length; nCell++)
				{
					aCellsToFill[nCell].SetBackgroundColor(r, g, b, bNone);
				}
				return true;
			}
			return false;
		}
		else
			return false;
	};
	
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiStyle
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiStyle class.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {"style"}
	 */
	ApiStyle.prototype.GetClassType = function()
	{
		return "style";
	};
	/**
	 * Returns a name of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiStyle.prototype.GetName = function()
	{
		return this.Style.Get_Name();
	};
	/**
	 * Sets a name of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @param {string} sStyleName - The name which will be used for the current style.
	 */
	ApiStyle.prototype.SetName = function(sStyleName)
	{
		this.Style.Set_Name(sStyleName);
	};
	/**
	 * Returns a type of the current style.
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
	 * Returns the text properties of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTextPr}
	 */
	ApiStyle.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.Style.TextPr.Copy());
	};
	/**
	 * Returns the paragraph properties of the current style.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParaPr}
	 */
	ApiStyle.prototype.GetParaPr = function()
	{
		return new ApiParaPr(this, this.Style.ParaPr.Copy());
	};
	/**
	 * Returns the table properties of the current style.
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
	 * Returns the table row properties of the current style.
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
	 * Returns the table cell properties of the current style.
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
	 * Specifies the reference to the parent style which this style inherits from in the style hierarchy.
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
	 * Returns a set of formatting properties which will be conditionally applied to the parts of a table that match the 
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
	/**
	 * Converts the ApiStyle object into the JSON object.
	 * @memberof ApiStyle
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteNumberings - Specifies if the used numberings will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiStyle.prototype.ToJSON = function(bWriteNumberings)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerWordStyle(this.Style);
		if (bWriteNumberings)
			oJSON["numbering"] = oWriter.jsonWordNumberings;
		return JSON.stringify(oJSON);
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTextPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiTextPr class.
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
	 * the selected or created style to the text.</note>
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {ApiStyle} oStyle - The style which must be applied to the text character.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetStyle = function(oStyle)
	{
		if (!(oStyle instanceof ApiStyle))
			return;

		this.TextPr.RStyle = oStyle.Style.Get_Id();
		this.private_OnChange();
		return this;
	};
	/**
	 * Sets the bold property to the text character.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isBold - Specifies that the contents of the run are displayed bold.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetBold = function(isBold)
	{
		this.TextPr.Bold = isBold;
		this.private_OnChange();
		return this;
	};
	/**
	 * Sets the italic property to the text character.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isItalic - Specifies that the contents of the current run are displayed italicized.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetItalic = function(isItalic)
	{
		this.TextPr.Italic = isItalic;
		this.private_OnChange();
		return this;
	};
	/**
	 * Specifies that the contents of the run are displayed with a single horizontal line through the center of the line.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isStrikeout - Specifies that the contents of the current run are displayed struck through.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetStrikeout = function(isStrikeout)
	{
		this.TextPr.Strikeout = isStrikeout;
		this.private_OnChange();
		return this;
	};
	/**
	 * Specifies that the contents of the run are displayed along with a line appearing directly below the character
	 * (less than all the spacing above and below the characters on the line).
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isUnderline - Specifies that the contents of the current run are displayed underlined.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetUnderline = function(isUnderline)
	{
		this.TextPr.Underline = isUnderline;
		this.private_OnChange();
		return this;
	};
	/**
	 * Sets all 4 font slots with the specified font family.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {string} sFontFamily - The font family or families used for the current text run.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetFontFamily = function(sFontFamily)
	{
		this.TextPr.RFonts.SetAll(sFontFamily, -1);
		this.private_OnChange();
		return this;
	};
	/**
	 * Sets the font size to the characters of the current text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {hps} nSize - The text size value measured in half-points (1/144 of an inch).
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetFontSize = function(nSize)
	{
		this.TextPr.FontSize = private_GetHps(nSize);
		this.private_OnChange();
		return this;
	};
	/**
	 * Sets the text color to the current text run in the RGB format.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - If this parameter is set to "true", then r,g,b parameters will be ignored.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetColor = function(r, g, b, isAuto)
	{
		this.TextPr.Color = private_GetColor(r, g, b, isAuto);
		this.private_OnChange();
		return this;
	};
	/**
	 * Specifies the alignment which will be applied to the contents of the run in relation to the default appearance of the run text:
	 * * <b>"baseline"</b> - the characters in the current text run will be aligned by the default text baseline.
	 * * <b>"subscript"</b> - the characters in the current text run will be aligned below the default text baseline.
	 * * <b>"superscript"</b> - the characters in the current text run will be aligned above the default text baseline.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {("baseline" | "subscript" | "superscript")} sType - The vertical alignment type applied to the text contents.
	 * @return {ApiTextPr} - this text properties.
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
		return this;
	};
	/**
	 * Specifies a highlighting color which is added to the text properties and applied as a background to the contents of the current run/range/paragraph.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {highlightColor} sColor - Available highlight color.
	 * @returns {ApiTextPr}
	 */
	ApiTextPr.prototype.SetHighlight = function(sColor)
	{
		if (!editor && Asc.editor)
		 	return this;

		if ("none" === sColor)
		{
			if (editor.editorId === AscCommon.c_oEditorId.Word)
				this.TextPr.HighLight = AscCommonWord.highlight_None;
			else if (editor.editorId === AscCommon.c_oEditorId.Presentation)
				this.TextPr.HighlightColor = null;
		}
		else
		{
			var color = private_getHighlightColorByName(sColor);
			if (color && editor.editorId === AscCommon.c_oEditorId.Word)
				this.TextPr.HighLight = new CDocumentColor(color.r, color.g, color.b)
			else if (color && editor.editorId === AscCommon.c_oEditorId.Presentation)
				this.TextPr.HighlightColor = AscFormat.CreateUniColorRGB(color.r, color.g, color.b)
		}
		this.private_OnChange();

		return this;
	};
	/**
	 * Sets the text spacing measured in twentieths of a point.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nSpacing - The value of the text spacing measured in twentieths of a point (1/1440 of an inch).
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetSpacing = function(nSpacing)
	{
		this.TextPr.Spacing = private_Twips2MM(nSpacing);
		this.private_OnChange();
		return this;
	};
	/**
	 * Specifies that the contents of the run are displayed with two horizontal lines through each character displayed on the line.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isDoubleStrikeout - Specifies that the contents of the current run are displayed double struck through.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetDoubleStrikeout = function(isDoubleStrikeout)
	{
		this.TextPr.DStrikeout = isDoubleStrikeout;
		this.private_OnChange();
		return this;
	};
	/**
	 * Specifies that any lowercase characters in the text run are formatted for display only as their capital letter character equivalents.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isCaps - Specifies that the contents of the current run are displayed capitalized.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetCaps = function(isCaps)
	{
		this.TextPr.Caps = isCaps;
		this.private_OnChange();
		return this;
	};
	/**
	 * Specifies that all the small letter characters in the text run are formatted for display only as their capital
	 * letter character equivalents which are two points smaller than the actual font size specified for this text.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {boolean} isSmallCaps - Specifies if the contents of the current run are displayed capitalized two points smaller or not.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetSmallCaps = function(isSmallCaps)
	{
		this.TextPr.SmallCaps = isSmallCaps;
		this.private_OnChange();
		return this;
	};
	/**
	 * Specifies an amount by which text is raised or lowered for this run in relation to the default
	 * baseline of the surrounding non-positioned text.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {hps} nPosition - Specifies a positive (raised text) or negative (lowered text)
	 * measurement in half-points (1/144 of an inch).
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetPosition = function(nPosition)
	{
		this.TextPr.Position = private_PtToMM(private_GetHps(nPosition));
		this.private_OnChange();
		return this;
	};
	/**
	 * Specifies the languages which will be used to check spelling and grammar (if requested) when processing
	 * the contents of the text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {string} sLangId - The possible value for this parameter is a language identifier as defined by
	 * RFC 4646/BCP 47. Example: "en-CA".
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetLanguage = function(sLangId)
	{
		var nLcid = Asc.g_oLcidNameToIdMap[sLangId];
		if (undefined !== nLcid)
		{
			this.TextPr.Lang.Val = nLcid;
			this.private_OnChange();
			return this;
		}
	};
	/**
	 * Specifies the shading applied to the contents of the current text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type applied to the contents of the current text run.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetShd = function(sType, r, g, b)
	{
		this.TextPr.Shd = private_GetShd(sType, r, g, b, false);
		this.private_OnChange();
		return this;
	};
	/**
	 * Sets the text color to the current text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CSE", "CPE"]
	 * @param {ApiFill} oApiFill - The color or pattern used to fill the text color.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetFill = function(oApiFill)
	{
		this.TextPr.Unifill = oApiFill.UniFill;
		this.private_OnChange();
		return this;
	};
	/**
	 * Sets the text fill to the current text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CSE", "CPE", "CSE"]
	 * @param {ApiFill} oApiFill - The color or pattern used to fill the text color.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetTextFill = function(oApiFill)
	{
		this.TextPr.TextFill = oApiFill.UniFill;
		this.private_OnChange();
		return this;
	};
	/**
	 * Sets the text outline to the current text run.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CSE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the text outline.
	 * @return {ApiTextPr} - this text properties.
	 */
	ApiTextPr.prototype.SetOutLine = function(oStroke)
	{
		this.TextPr.TextOutline = oStroke.Ln;
		this.private_OnChange();
		return this;
	};

	/**
	 * Converts the ApiTextPr object into the JSON object.
	 * @memberof ApiTextPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiTextPr.prototype.ToJSON = function(bWriteStyles)
	{
		let oWriter = new AscJsonConverter.WriterToJSON();
		let bFromDocument = true;
		let oParentRun = this.Parent;
		let oParentPara = oParentRun instanceof ParaRun ? oParentRun.GetParagraph() : null;

		if (oWriter.isWord === false || (oParentPara instanceof Paragraph && oParentPara.bFromDocument !== true))
			bFromDocument = false;

		let oJSON = bFromDocument ? oWriter.SerTextPr(this.TextPr) : oWriter.SerTextPrDrawing(this.TextPr);
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiParaPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiParaPr class.
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
	 * Returns the paragraph style method.
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
	 * Specifies that any space before or after this paragraph set using the 
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
	 * Sets the paragraph left side indentation.
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
	 * Returns the paragraph left side indentation.
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
	 * Sets the paragraph right side indentation.
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
	 * Returns the paragraph right side indentation.
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
	 * Sets the paragraph first line indentation.
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
	 * Returns the paragraph first line indentation.
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
	 * Sets the paragraph contents justification.
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
	 * Returns the paragraph contents justification.
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
	 * Specifies that when rendering the document using a page view, all lines of the current paragraph are maintained on a single page whenever possible.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isKeepLines - The true value enables the option to keep lines of the paragraph on a single page.
	 */
	ApiParaPr.prototype.SetKeepLines = function(isKeepLines)
	{
		this.ParaPr.KeepLines = isKeepLines;
		this.private_OnChange();
	};
	/**
	 * Specifies that when rendering the document using a paginated view, the contents of the current paragraph are at least
	 * partly rendered on the same page as the following paragraph whenever possible.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isKeepNext - The true value enables the option to keep lines of the paragraph on the same
	 * page as the following paragraph.
	 */
	ApiParaPr.prototype.SetKeepNext = function(isKeepNext)
	{
		this.ParaPr.KeepNext = isKeepNext;
		this.private_OnChange();
	};
	/**
	 * Specifies that when rendering the document using a paginated view, the contents of the current paragraph are rendered at
	 * the beginning of a new page in the document.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isPageBreakBefore - The true value enables the option to render the contents of the paragraph
	 * at the beginning of a new page in the document.
	 */
	ApiParaPr.prototype.SetPageBreakBefore = function(isPageBreakBefore)
	{
		this.ParaPr.PageBreakBefore = isPageBreakBefore;
		this.private_OnChange();
	};
	/**
	 * Sets the paragraph line spacing. If the value of the sLineRule parameter is either 
	 * "atLeast" or "exact", then the value of nLine will be interpreted as twentieths of a point. If 
	 * the value of the sLineRule parameter is "auto", then the value of the 
	 * nLine parameter will be interpreted as 240ths of a line.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {(twips | line240)} nLine - The line spacing value measured either in twentieths of a point (1/1440 of an inch) or in 240ths of a line.
	 * @param {("auto" | "atLeast" | "exact")} sLineRule - The rule that determines the measuring units of the line spacing.
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
	 * Returns the paragraph line spacing value.
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
	 * Returns the paragraph line spacing rule.
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
	 * Sets the spacing before the current paragraph. If the value of the isBeforeAuto parameter is true, then 
	 * any value of the nBefore is ignored. If isBeforeAuto parameter is not specified, then 
	 * it will be interpreted as false.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nBefore - The value of the spacing before the current paragraph measured in twentieths of a point (1/1440 of an inch).
	 * @param {boolean} [isBeforeAuto=false] - The true value disables the spacing before the current paragraph.
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
	 * Returns the spacing before value of the current paragraph.
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
	 * Sets the spacing after the current paragraph. If the value of the isAfterAuto parameter is true, then 
	 * any value of the nAfter is ignored. If isAfterAuto parameter is not specified, then it 
	 * will be interpreted as false.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {twips} nAfter - The value of the spacing after the current paragraph measured in twentieths of a point (1/1440 of an inch).
	 * @param {boolean} [isAfterAuto=false] - The true value disables the spacing after the current paragraph.
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
	 * Returns the spacing after value of the current paragraph. 
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
	 * Specifies the shading applied to the contents of the paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type which will be applied to the contents of the current paragraph.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - The true value disables paragraph contents shading.
	 */
	ApiParaPr.prototype.SetShd = function(sType, r, g, b, isAuto)
	{
		this.ParaPr.Shd = private_GetShd(sType, r, g, b, isAuto);
		this.private_OnChange();
	};
	/**
	 * Returns the shading applied to the contents of the paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @returns {?ApiRGBColor}
	 */
	ApiParaPr.prototype.GetShd = function()
	{
		var oColor = null;
		var oShd;
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
	 * Specifies the border which will be displayed below a set of paragraphs which have the same paragraph border settings.
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
	 * Specifies the border which will be displayed at the left side of the page around the specified paragraph.
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
	 * Specifies the border which will be displayed at the right side of the page around the specified paragraph.
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
	 * Specifies the border which will be displayed above a set of paragraphs which have the same set of paragraph border settings.
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
	 * Specifies the border which will be displayed between each paragraph in a set of paragraphs which have the same set of paragraph border settings.
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
	 * Specifies whether a single line of the current paragraph will be displayed on a separate page from the remaining content at display time by moving the line onto the following page.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isWidowControl - The true value means that a single line of the current paragraph will be displayed on a separate page from the remaining content at display time by moving the line onto the following page.
	 */
	ApiParaPr.prototype.SetWidowControl = function(isWidowControl)
	{
		this.ParaPr.WidowControl = isWidowControl;
		this.private_OnChange();
	};
	/**
	 * Specifies a sequence of custom tab stops which will be used for any tab characters in the current paragraph.
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
	 * Specifies that the current paragraph references a numbering definition instance in the current document.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {ApiNumbering} oNumPr - Specifies a numbering definition.
	 * @param {number} [nLvl=0] - Specifies a numbering level reference. If the current instance of the ApiParaPr class is direct
	 * formatting of a paragraph, then this parameter MUST BE specified. Otherwise, if the current instance of the ApiParaPr class
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
	 * Sets the bullet or numbering to the current paragraph.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CSE", "CPE"]
	 * @param {?ApiBullet} oBullet - The bullet object created with the {@link Api#CreateBullet} or {@link Api#CreateNumbering} method.
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
	/**
	 * Converts the ApiParaPr object into the JSON object.
	 * @memberof ApiParaPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiParaPr.prototype.ToJSON = function(bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = this.Parent != null && this.Parent instanceof Paragraph && this.Parent.bFromDocument !== true ? oWriter.SerParaPrDrawing(this.ParaPr) : oWriter.SerParaPr(this.ParaPr);
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};


	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiNumbering
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiNumbering class.
	 * @memberof ApiNumbering
	 * @typeofeditors ["CDE"]
	 * @returns {"numbering"}
	 */
	ApiNumbering.prototype.GetClassType = function()
	{
		return "numbering";
	};
	/**
	 * Returns the specified level of the current numbering.
	 * @memberof ApiNumbering
	 * @typeofeditors ["CDE"]
	 * @param {number} nLevel - The numbering level index. This value MUST BE from 0 to 8.
	 * @returns {ApiNumberingLevel}
	 */
	ApiNumbering.prototype.GetLevel = function(nLevel)
	{
		return new ApiNumberingLevel(this.Num, nLevel);
	};
	/**
	 * Converts the ApiNumbering object into the JSON object.
	 * @memberof ApiNumbering
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiNumbering.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		oWriter.SerNumbering(this.Num);

		return JSON.stringify(oWriter.jsonWordNumberings);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiNumberingLevel
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiNumberingLevel class.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {"numberingLevel"}
	 */
	ApiNumberingLevel.prototype.GetClassType = function()
	{
		return "numberingLevel";
	};
	/**
	 * Returns the numbering definition.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {ApiNumbering}
	 */
	ApiNumberingLevel.prototype.GetNumbering = function()
	{
		return new ApiNumbering(this.Num);
	};
	/**
	 * Returns the level index.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiNumberingLevel.prototype.GetLevelIndex = function()
	{
		return this.Lvl;
	};
	/**
	 * Specifies the text properties which will be applied to the text in the current numbering level itself, not to the text in the subsequent paragraph.
	 * <note>To change the text style of the paragraph, a style must be applied to it using the {@link ApiRun#SetStyle} method.</note>
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTextPr}
	 */
	ApiNumberingLevel.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.Num.GetLvl(this.Lvl).TextPr.Copy());
	};
	/**
	 * Returns the paragraph properties which are applied to any numbered paragraph that references the given numbering definition and numbering level.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParaPr}
	 */
	ApiNumberingLevel.prototype.GetParaPr = function()
	{
		return new ApiParaPr(this, this.Num.GetLvl(this.Lvl).ParaPr.Copy());
	};
	/**
	 * Sets one of the existing predefined numbering templates.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @param {("none" | "bullet" | "1)" | "1." | "I." | "A." | "a)" | "a." | "i." )} sType - The predefined numbering template.
	 * @param {string} [sSymbol=""] - The symbol used for the list numbering. This parameter has the meaning only if the predefined numbering template is "bullet".
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
	 * Sets your own customized numbering type.
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
	 * Specifies a one-based index which determines when a numbering level should restart to its starting value. A numbering level restarts when an instance of the specified numbering level which is higher (earlier than this level) is used in the given document contents. By default this value is true.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isRestart - The true value means that a numbering level will be restarted to its starting value.
	 */
	ApiNumberingLevel.prototype.SetRestart = function(isRestart)
	{
		this.Num.SetLvlRestart(this.Lvl, private_GetBoolean(isRestart, true));
	};
	/**
	 * Specifies the starting value for the numbering used by the parent numbering level within a given numbering level definition. By default this value is 1.
	 * @memberof ApiNumberingLevel
	 * @typeofeditors ["CDE"]
	 * @param {number} nStart - The starting value for the numbering used by the parent numbering level.
	 */
	ApiNumberingLevel.prototype.SetStart = function(nStart)
	{
		this.Num.SetLvlStart(this.Lvl, private_GetInt(nStart));
	};
	/**
	 * Specifies the content which will be added between the given numbering level text and the text of every numbered paragraph which references that numbering level. By default this value is "tab".
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
	 * Returns a type of the ApiTablePr class.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @returns {"tablePr"}
	 */
	ApiTablePr.prototype.GetClassType = function()
	{
		return "tablePr";
	};
	/**
	 * Specifies a number of columns which will comprise each table column band for this table style.
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
	 * Specifies a number of rows which will comprise each table row band for this table style.
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
	 * Specifies the alignment of the current table with respect to the text margins in the current section.
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
	 * Specifies the shading which is applied to the extents of the current table.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type applied to the extents of the current table.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - The true value disables the SetShd method use.
	 */
	ApiTablePr.prototype.SetShd = function(sType, r, g, b, isAuto)
	{
		this.TablePr.Shd = private_GetShd(sType, r, g, b, isAuto);
		this.private_OnChange();
	};
	/**
	 * Sets the border which will be displayed at the top of the current table.
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
	 * Sets the border which will be displayed at the bottom of the current table.
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
	 * Sets the border which will be displayed on the left of the current table.
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
	 * Sets the border which will be displayed on the right of the current table.
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
	 * Specifies the border which will be displayed on all horizontal table cell borders which are not on the outmost edge
	 * of the parent table (all horizontal borders which are not the topmost or bottommost borders).
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
	 * Specifies the border which will be displayed on all vertical table cell borders which are not on the outmost edge
	 * of the parent table (all vertical borders which are not the leftmost or rightmost borders).
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
	 * Specifies an amount of space which will be left between the bottom extent of the cell contents and the border
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
	 * Specifies an amount of space which will be left between the left extent of the cell contents and the left
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
	 * Specifies an amount of space which will be left between the right extent of the cell contents and the right
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
	 * Specifies an amount of space which will be left between the top extent of the cell contents and the top border
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
	 * Specifies the default table cell spacing (the spacing between adjacent cells and the edges of the table).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {?twips} nValue - Spacing value measured in twentieths of a point (1/1440 of an inch). <code>"Null"</code> means that no spacing will be applied.
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
	 * Specifies the indentation which will be added before the leading edge of the current table in the document
	 * (the left edge in the left-to-right table, and the right edge in the right-to-left table).
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
	 * Sets the preferred width to the current table.
	 * <note>Tables are created with the {@link ApiTable#SetWidth} method properties set by default, which always override the {@link ApiTablePr#SetWidth} method properties. That is why there is no use to try and apply {@link ApiTablePr#SetWidth}. We recommend you to use the  {@link ApiTablePr#SetWidth} method instead.</note>
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
	 * Specifies the algorithm which will be used to lay out the contents of the current table within the document.
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
	/**
	 * Sets the table title (caption).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {string} sTitle - The table title to be set.
	 * @return {boolean}
	 */
	ApiTablePr.prototype.SetTableTitle = function(sTitle)
	{
		if (typeof(sTitle) !== "string" || sTitle === "")
			return false;

		this.TablePr.TableCaption = sTitle;
		this.private_OnChange();
		return true;
	};
	/**
	 * Returns the table title (caption).
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @return {string}
	 */
	ApiTablePr.prototype.GetTableTitle = function()
	{
		if (this.TablePr.TableCaption)
			return this.TablePr.TableCaption;

		return "";
	};
	/**
	 * Sets the table description.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @param {string} sDescr - The table description to be set.
	 * @return {boolean}
	 */
	ApiTablePr.prototype.SetTableDescription = function(sDescr)
	{
		if (typeof(sDescr) !== "string" || sDescr === "")
			return false;

		this.TablePr.TableDescription = sDescr;
		this.private_OnChange();
		return true;
	};
	/**
	 * Returns the table description.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @return {string}
	 */
	ApiTablePr.prototype.GetTableDescription = function()
	{
		if (this.TablePr.TableDescription)
			return this.TablePr.TableDescription;

		return "";
	};
	/**
	 * Converts the ApiTablePr object into the JSON object.
	 * @memberof ApiTablePr
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiTablePr.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerTablePr(this.TablePr));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableRowPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiTableRowPr class.
	 * @memberof ApiTableRowPr
	 * @typeofeditors ["CDE"]
	 * @returns {"tableRowPr"}
	 */
	ApiTableRowPr.prototype.GetClassType = function()
	{
		return "tableRowPr";
	};
	/**
     * Sets the height to the current table row within the current table.
	 * @memberof ApiTableRowPr
	 * @typeofeditors ["CDE"]
	 * @param {("auto" | "atLeast")} sHRule - The rule to apply the height value to the current table row or ignore it. Use the <code>"atLeast"</code> value to enable the <code>SetHeight</code> method use.
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
	 * Specifies that the current table row will be repeated at the top of each new page 
     * wherever this table is displayed. This gives this table row the behavior of a 'header' row on 
     * each of these pages. This element can be applied to any number of rows at the top of the 
     * table structure in order to generate multi-row table headers.
	 * @memberof ApiTableRowPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isHeader - The true value means that the current table row will be repeated at the top of each new page.
	 */
	ApiTableRowPr.prototype.SetTableHeader = function(isHeader)
	{
		this.RowPr.TableHeader = private_GetBoolean(isHeader);
		this.private_OnChange();
	};
	/**
	 * Converts the ApiTableRowPr object into the JSON object.
	 * @memberof ApiTableRowPr
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiTableRowPr.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerTableRowPr(this.RowPr));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableCellPr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiTableCellPr class.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @returns {"tableCellPr"}
	 */
	ApiTableCellPr.prototype.GetClassType = function()
	{
		return "tableCellPr";
	};
	/**
	 * Specifies the shading applied to the contents of the table cell.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {ShdType} sType - The shading type which will be applied to the contents of the current table cell.
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} [isAuto=false] - The true value disables the table cell contents shading.
	 */
	ApiTableCellPr.prototype.SetShd = function(sType, r, g, b, isAuto)
	{
		this.CellPr.Shd = private_GetShd(sType, r, g, b, isAuto);
		this.private_OnChange();
	};
	/**
	 * Specifies an amount of space which will be left between the bottom extent of the cell contents and the border
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
	 * Specifies an amount of space which will be left between the left extent of the cell contents and 
	 * the border of a specific table cell within a table.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {?twips} nValue - The value for the amount of space to the left extent of the cell measured in twentieths
	 * of a point (1/1440 of an inch). If this value is <code>null</code>, then default table cell left margin will be used, otherwise
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
	 * Specifies an amount of space which will be left between the right extent of the cell contents and the border of a specific table cell within a table.
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
	 * Specifies an amount of space which will be left between the upper extent of the cell contents
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
	 * Sets the border which will be displayed at the bottom of the current table cell.
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
	 * Sets the border which will be displayed to the left of the current table cell.
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
	 * Sets the border which will be displayed to the right of the current table cell.
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
	 * Sets the border which will be displayed at the top of the current table cell.
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
	 * Sets the preferred width to the current table cell.
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
	 * Specifies the vertical alignment for the text contents within the current table cell.
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
	 * Specifies the direction of the text flow for this table cell.
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
	 * Specifies how the current table cell is laid out when the parent table is displayed in a document. This setting
	 * only affects the behavior of the cell when the {@link ApiTablePr#SetTableLayout} table layout for this table is set to use the <code>"autofit"</code> algorithm.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @param {boolean} isNoWrap - The true value means that the current table cell will not be wrapped in the parent table.
	 */
	ApiTableCellPr.prototype.SetNoWrap = function(isNoWrap)
	{
		this.CellPr.NoWrap = private_GetBoolean(isNoWrap);
		this.private_OnChange();
	};
	/**
	 * Converts the ApiTableCellPr object into the JSON object.
	 * @memberof ApiTableCellPr
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiTableCellPr.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerTableCellPr(this.CellPr));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTableStylePr
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiTableStylePr class.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {"tableStylePr"}
	 */
	ApiTableStylePr.prototype.GetClassType = function()
	{
		return "tableStylePr";
	};
	/**
	 * Returns a type of the current table conditional style.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {TableStyleOverrideType}
	 */
	ApiTableStylePr.prototype.GetType = function()
	{
		return this.Type;
	};
	/**
	 * Returns a set of the text run properties which will be applied to all the text runs within the table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTextPr}
	 */
	ApiTableStylePr.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.TableStylePr.TextPr);
	};
	/**
	 * Returns a set of the paragraph properties which will be applied to all the paragraphs within a table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiParaPr}
	 */
	ApiTableStylePr.prototype.GetParaPr = function()
	{
		return new ApiParaPr(this, this.TableStylePr.ParaPr);
	};
	/**
	 * Returns a set of the table properties which will be applied to all the regions within a table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTablePr}
	 */
	ApiTableStylePr.prototype.GetTablePr = function()
	{
		return new ApiTablePr(this, this.TableStylePr.TablePr);
	};
	/**
	 * Returns a set of the table row properties which will be applied to all the rows within a table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableRowPr}
	 */
	ApiTableStylePr.prototype.GetTableRowPr = function()
	{
		return new ApiTableRowPr(this, this.TableStylePr.TableRowPr);
	};
	/**
	 * Returns a set of the table cell properties which will be applied to all the cells within a table which match the conditional formatting type.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {ApiTableCellPr}
	 */
	ApiTableStylePr.prototype.GetTableCellPr = function()
	{
		return new ApiTableCellPr(this, this.TableStylePr.TableCellPr);
	};
	/**
	 * Converts the ApiTableStylePr object into the JSON object.
	 * @memberof ApiTableStylePr
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiTableStylePr.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerTableStylePr(this.TableStylePr, this.Type));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiDrawing
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiDrawing class.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE", "CPE"]
	 * @returns {"drawing"}
	 */
	ApiDrawing.prototype.GetClassType = function()
	{
		return "drawing";
	};
	/**
	 * Sets the size of the object (image, shape, chart) bounding box.
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
	 * Sets the wrapping type of the current object (image, shape, chart). One of the following wrapping style types can be set:
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
	 * Specifies how the floating object will be horizontally aligned.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {RelFromH} [sRelativeFrom="page"] - The document element which will be taken as a countdown point for the object horizontal alignment.
	 * @param {("left" | "right" | "center")} [sAlign="left"] - The alignment type which will be used for the object horizontal alignment.
	 */
	ApiDrawing.prototype.SetHorAlign = function(sRelativeFrom, sAlign)
	{
		var nAlign        = private_GetAlignH(sAlign);
		var nRelativeFrom = private_GetRelativeFromH(sRelativeFrom);
		this.Drawing.Set_PositionH(nRelativeFrom, true, nAlign, false);
	};
	/**
	 * Specifies how the floating object will be vertically aligned.
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
	 * Sets the absolute measurement for the horizontal positioning of the floating object.
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
	 * Sets the absolute measurement for the vertical positioning of the floating object.
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
	 * Specifies the minimum distance which will be maintained between the edges of the current drawing object and any
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
	 * Returns a parent paragraph that contains the graphic object.
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
	 * Returns a parent content control that contains the graphic object.
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
	 * Returns a parent table that contains the graphic object.
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
	 * Returns a parent table cell that contains the graphic object.
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
	 * Deletes the current graphic object. 
	 * @typeofeditors ["CDE"]
	 * @return {boolean} - returns false if drawing object haven't parent.
	 */
	ApiDrawing.prototype.Delete = function()
	{
		var ParaParent = this.GetParentParagraph();

		if (ParaParent)
		{
			this.Drawing.PreDelete();
			var oParentRun = this.Drawing.GetRun();
			oParentRun.RemoveElement(this.Drawing);

			return true;
		}
		else 	 
			return false;
	};
	/**
	 * Copies the current graphic object. 
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
	 * Wraps the graphic object with a rich text content control.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {number} nType - Defines if this method returns the ApiBlockLvlSdt (nType === 1) or ApiDrawing (any value except 1) object.
	 * @return {ApiDrawing | ApiBlockLvlSdt}  
	 */
	ApiDrawing.prototype.InsertInContentControl = function(nType)
	{
		var Document			= editor.private_GetLogicDocument();
		var ContentControl;
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
	 * Inserts a paragraph at the specified position.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {string | ApiParagraph} paragraph - Text or paragraph.
	 * @param {string} sPosition - The position where the text or paragraph will be inserted ("before" or "after" the drawing specified).
	 * @param {boolean} beRNewPara - Defines if this method returns a new paragraph (true) or the current ApiDrawing (false).
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
	 * Selects the current graphic object.
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
	 * Inserts a break at the specified location in the main document.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {number}	breakType - The break type: page break (0) or line break (1).
	 * @param {string}	position  - The position where the page or line break will be inserted ("before" or "after" the current drawing).
	 * @returns {boolean}  - returns false if drawing object haven't parent run or params are invalid.
	 */	
	ApiDrawing.prototype.AddBreak = function(breakType, position)
	{
		var ParentRun	= (new ApiRun(this.Drawing.GetRun()));

		if (!ParentRun || position !== "before" && position !== "after" || breakType !== 1 && breakType !== 0)
			return false;

		if (breakType === 0)
		{
			if (position === "before")
				ParentRun.Run.Add_ToContent(ParentRun.Run.Content.indexOf(this.Drawing), new AscWord.CRunBreak(AscWord.break_Page));
			else if (position === "after")
				ParentRun.Run.Add_ToContent(ParentRun.Run.Content.indexOf(this.Drawing) + 1, new AscWord.CRunBreak(AscWord.break_Page));
		}
		else if (breakType === 1)
		{
			if (position === "before")
				ParentRun.Run.Add_ToContent(ParentRun.Run.Content.indexOf(this.Drawing), new AscWord.CRunBreak(AscWord.break_Line));
			else if (position === "after")
				ParentRun.Run.Add_ToContent(ParentRun.Run.Content.indexOf(this.Drawing) + 1, new AscWord.CRunBreak(AscWord.break_Line));
		}

		return true;
	};
	/**
	 * Flips the current drawing horizontally.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bFlip - Specifies if the figure will be flipped horizontally or not.
	 */	
	ApiDrawing.prototype.SetHorFlip = function(bFlip)
	{
		if (this.Drawing.GraphicObj && this.Drawing.GraphicObj.spPr && this.Drawing.GraphicObj.spPr.xfrm)
			this.Drawing.GraphicObj.spPr.xfrm.setFlipH(bFlip);
	};
	/**
	 * Flips the current drawing vertically.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bFlip - Specifies if the figure will be flipped vertically or not.
	 * @returns {boolean} - returns false if param is invalid.
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
	 * Scales the height of the figure using the specified coefficient.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {number} coefficient - The coefficient by which the figure height will be scaled.
	 * @returns {boolean} - return false if param is invalid.
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
	 * Scales the width of the figure using the specified coefficient.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {number} coefficient - The coefficient by which the figure width will be scaled.
	 * @returns {boolean} - return false if param is invali.
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
	 * Sets the fill formatting properties to the current graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {ApiFill} oFill - The fill type used to fill the graphic object.
	 * @returns {boolean} - returns false if param is invalid.
	 */	
	ApiDrawing.prototype.Fill = function(oFill)
	{
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		this.Drawing.GraphicObj.spPr.setFill(oFill.UniFill);
		return true;
	};
	/**
	 * Sets the outline properties to the specified graphic object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the graphic object outline.
	 * @returns {boolean} - returns false if param is invalid.
	 */	
	ApiDrawing.prototype.SetOutLine = function(oStroke)
	{
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		this.Drawing.GraphicObj.spPr.setLn(oStroke.Ln);
		return true;
	};
	/**
	 * Returns the next inline drawing object if exists. 
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
	 * Returns the previous inline drawing object if exists. 
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
	/**
	 * Converts the ApiDrawing object into the JSON object.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteNumberings - Specifies if the used numberings will be written to the JSON object or not.
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @returns {JSON}
	 */
	ApiDrawing.prototype.ToJSON = function(bWriteNumberings, bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerParaDrawing(this.Drawing);
		if (bWriteNumberings)
			oJSON["numbering"] = oWriter.jsonWordNumberings;
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};

	/**
	 * Returns the width of the current drawing.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {EMU}
	 */
	ApiDrawing.prototype.GetWidth = function()
	{
		return private_MM2EMU(this.Drawing.getXfrmExtX());
	};
	/**
	 * Returns the height of the current drawing.
	 * @memberof ApiDrawing
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {EMU}
	 */
	ApiDrawing.prototype.GetHeight = function()
	{
		return private_MM2EMU(this.Drawing.getXfrmExtY());
	};
	/**
     * Returns the lock value for the specified lock type of the current drawing.
     * @typeofeditors ["CPE"]
	 * @param {"noGrp" | "noUngrp" | "noSelect" | "noRot" | "noChangeAspect" | "noMove" | "noResize" | "noEditPoints" | "noAdjustHandles"
	 * 	| "noChangeArrowheads" | "noChangeShapeType" | "noDrilldown" | "noTextEdit" | "noCrop" | "txBox"} sType - Lock type in the string format.
     * @returns {bool}
     */
	ApiDrawing.prototype.GetLockValue = function(sType)
	{
		var nLockType = private_GetDrawingLockType(sType);

		if (nLockType === -1)
			return false;

		if (this.Drawing && this.Drawing.GraphicObj)
			return this.Drawing.GraphicObj.getLockValue(nLockType);

		return false;
	};

	/**
     * Sets the lock value to the specified lock type of the current drawing.
     * @typeofeditors ["CPE"]
	 * @param {"noGrp" | "noUngrp" | "noSelect" | "noRot" | "noChangeAspect" | "noMove" | "noResize" | "noEditPoints" | "noAdjustHandles"
	 * 	| "noChangeArrowheads" | "noChangeShapeType" | "noDrilldown" | "noTextEdit" | "noCrop" | "txBox"} sType - Lock type in the string format.
     * @param {bool} bValue - Specifies if the specified lock is applied to the current drawing.
	 * @returns {bool}
     */
	ApiDrawing.prototype.SetLockValue = function(sType, bValue)
	{
		var nLockType = private_GetDrawingLockType(sType);

		if (nLockType === -1)
			return false;

		if (this.Drawing && this.Drawing.GraphicObj)
		{
			this.Drawing.GraphicObj.setLockValue(nLockType, bValue);
			return true;
		}

		return false;
	};

	/**
     * Sets the properties from another drawing to the current drawing.
	 * The following properties will be copied: horizontal and vertical alignment, distance between the edges of the current drawing object and any subsequent text, wrapping style, drawing name, title and description.
     * @memberof ApiDrawing
     * @param {ApiDrawing} oAnotherDrawing - The drawing which properties will be set to the current drawing.
     * @typeofeditors ["CDE"]
     * @returns {boolean}
     */
	ApiDrawing.prototype.SetDrawingPrFromDrawing = function(oAnotherDrawing)
	{
		if (!(oAnotherDrawing instanceof ApiDrawing))
			return false;

		this.Drawing.SetDrawingPrFromDrawing(oAnotherDrawing.Drawing);
		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiImage
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiImage class.
	 * @memberof ApiImage
	 * @typeofeditors ["CDE", "CPE"]
	 * @returns {"image"}
	 */
	ApiImage.prototype.GetClassType = function()
	{
		return "image";
	};
	/**
	 * Returns the next inline image if exists. 
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
	 * Returns the previous inline image if exists. 
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
	// ApiOleObject
	//
	//------------------------------------------------------------------------------------------------------------------
	
	/**
	 * Returns a type of the ApiOleObject class.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {"oleObject"}
	 */
	ApiOleObject.prototype.GetClassType = function()
	{
		return "oleObject";
	};

	/**
	 * Sets the data to the current OLE object.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {string} sData - The OLE object string data.
	 * @returns {boolean}
	 */
	ApiOleObject.prototype.SetData = function(sData)
	{
		if (typeof(sData) !== "string" || sData === "")
			return false;

		this.OleObject.setData(sData);
		return true;
	};

	/**
	 * Returns the string data from the current OLE object.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {string}
	 */
	ApiOleObject.prototype.GetData = function()
	{
		if (typeof(this.OleObject.m_sData) === "string")
			return this.OleObject.m_sData;
		
		return "";
	};

	/**
	 * Sets the application ID to the current OLE object.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {string} sAppId - The application ID associated with the curent OLE object.
	 * @returns {boolean}
	 */
	ApiOleObject.prototype.SetApplicationId = function(sAppId)
	{
		if (typeof(sAppId) !== "string" || sAppId === "")
			return false;

		this.OleObject.setApplicationId(sAppId);
		return true;
	};

	/**
	 * Returns the application ID from the current OLE object.
	 * @memberof ApiOleObject
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {string}
	 */
	ApiOleObject.prototype.GetApplicationId = function()
	{
		if (typeof(this.OleObject.m_sApplicationId) === "string")
			return this.OleObject.m_sApplicationId;
		
		return "";
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiShape
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiShape class.
	 * @memberof ApiShape
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {"shape"}
	 */
	ApiShape.prototype.GetClassType = function()
	{
		return "shape";
	};
	/**
	 * Returns the shape inner contents where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {?ApiDocumentContent}
	 */
	ApiShape.prototype.GetDocContent = function()
	{
		if(this.Shape && this.Shape.textBoxContent && !this.Shape.isForm())
		{
			return new ApiDocumentContent(this.Shape.textBoxContent);
		}
		return null;
	};
	/**
	 * Returns the shape inner contents where a paragraph or text runs can be inserted.
	 * @memberof ApiShape
	 * @typeofeditors ["CDE", "CSE"]
	 * @returns {?ApiDocumentContent}
	 */
	ApiShape.prototype.GetContent = function()
	{
		if(this.Shape && this.Shape.textBoxContent && !this.Shape.isForm())
		{
			return new ApiDocumentContent(this.Shape.textBoxContent);
		}
		return null;
	};
	
	/**
	 * Sets the vertical alignment to the shape content where a paragraph or text runs can be inserted.
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
	 * Sets the text paddings to the current shape.
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
	 * Returns the next inline shape if exists. 
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
	 * Returns the previous inline shape if exists. 
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
	 * Returns a type of the ApiChart class.
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
	 *  Specifies the chart title.
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
	 *  Specifies the chart horizontal axis title.
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
	 *  Specifies the chart vertical axis title.
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
	 * Specifies the vertical axis orientation.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bIsMinMax - The <code>true</code> value will set the normal data direction for the vertical axis (from minimum to maximum).
	 * */
	ApiChart.prototype.SetVerAxisOrientation = function(bIsMinMax){
		AscFormat.builder_SetChartVertAxisOrientation(this.Chart, bIsMinMax);
	};

	/**
	 * Specifies the horizontal axis orientation.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bIsMinMax - The <code>true</code> value will set the normal data direction for the horizontal axis (from minimum to maximum).
	 * */
	ApiChart.prototype.SetHorAxisOrientation = function(bIsMinMax){
		AscFormat.builder_SetChartHorAxisOrientation(this.Chart, bIsMinMax);
	};

	/**
	 * Specifies the chart legend position.
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
	 * Specifies the legend font size.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 * */
	ApiChart.prototype.SetLegendFontSize = function(nFontSize)
	{
		AscFormat.builder_SetLegendFontSize(this.Chart, nFontSize);
	};

	/**
	 * Specifies which chart data labels are shown for the chart.
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
	 * Spicifies the show options for data labels.
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
	 * Spicifies tick labels position for the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart vertical tick labels.
	 * */
	ApiChart.prototype.SetVertAxisTickLabelPosition = function(sTickLabelPosition)
	{
		AscFormat.builder_SetChartVertAxisTickLablePosition(this.Chart, sTickLabelPosition);
	};

	/**
	 * Spicifies tick labels position for the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickLabelPosition} sTickLabelPosition - The type for the position of chart horizontal tick labels.
	 * */
	ApiChart.prototype.SetHorAxisTickLabelPosition = function(sTickLabelPosition)
	{
		AscFormat.builder_SetChartHorAxisTickLablePosition(this.Chart, sTickLabelPosition);
	};

	/**
	 * Specifies major tick mark for the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetHorAxisMajorTickMark = function(sTickMark){
		AscFormat.builder_SetChartHorAxisMajorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies minor tick mark for the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetHorAxisMinorTickMark = function(sTickMark){
		AscFormat.builder_SetChartHorAxisMinorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies major tick mark for the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */

	ApiChart.prototype.SetVertAxisMajorTickMark = function(sTickMark){
		AscFormat.builder_SetChartVerAxisMajorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies minor tick mark for the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {TickMark} sTickMark - The type of tick mark appearance.
	 * */
	ApiChart.prototype.SetVertAxisMinorTickMark = function(sTickMark){
		AscFormat.builder_SetChartVerAxisMinorTickMark(this.Chart, sTickMark);
	};

	/**
	 * Specifies major vertical gridline visual properties.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMajorVerticalGridlines = function(oStroke)
	{
		AscFormat.builder_SetVerAxisMajorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};

	/**
	 * Specifies minor vertical gridline visual properties.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMinorVerticalGridlines = function(oStroke)
	{
		AscFormat.builder_SetVerAxisMinorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};


	/**
	 * Specifies major horizontal gridline visual properties.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMajorHorizontalGridlines = function(oStroke)
	{
		AscFormat.builder_SetHorAxisMajorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};

	/**
	 * Specifies minor horizontal gridline visual properties.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {?ApiStroke} oStroke - The stroke used to create the element shadow.
	 * */
	ApiChart.prototype.SetMinorHorizontalGridlines = function(oStroke)
	{
		AscFormat.builder_SetHorAxisMinorGridlines(this.Chart, oStroke ?  oStroke.Ln : null);
	};


	/**
	 * Specifies font size for labels of the horizontal axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 */
	ApiChart.prototype.SetHorAxisLablesFontSize = function(nFontSize){
		AscFormat.builder_SetHorAxisFontSize(this.Chart, nFontSize);
	};

	/**
	 * Specifies font size for labels of the vertical axis.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE"]
	 * @param {pt} nFontSize - The text size value measured in points.
	 */
	ApiChart.prototype.SetVertAxisLablesFontSize = function(nFontSize){
		AscFormat.builder_SetVerAxisFontSize(this.Chart, nFontSize);
	};

	/**
	 * Returns the next inline chart if exists.
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
	 * Returns the previous inline chart if exists. 
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

	/**
	 * Removes the specified series from the current chart.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.RemoveSeria = function(nSeria)
	{
		return this.Chart.RemoveSeria(nSeria);
	};

	/**
	 * Sets values to the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {number[]} aValues - The array of the data which will be set to the specified chart series.
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriaValues = function(aValues, nSeria)
	{
		return this.Chart.SetValuesToDataPoints(aValues, nSeria);
	};

	/**
	 * Sets the x-axis values to all chart series. It is used with the scatter charts only.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {string[]} aValues - The array of the data which will be set to the x-axis data points.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetXValues = function(aValues)
	{
		if (this.Chart.isScatterChartType())
			return this.Chart.SetXValuesToDataPoints(aValues);
		return false;
	};

	/**
	 * Sets a name to the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {string} sName - The name which will be set to the specified chart series.
	 * @param {number} nSeria - The index of the chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriaName = function(sName, nSeria)
	{
		return this.Chart.SetSeriaName(sName, nSeria);
	};

	/**
	 * Sets a name to the specified chart category.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {string} sName - The name which will be set to the specified chart category.
	 * @param {number} nCategory - The index of the chart category.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetCategoryName = function(sName, nCategory)
	{
		return this.Chart.SetCatName(sName, nCategory);
	};

	/**
	 * Sets a style to the current chart by style ID.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param nStyleId - One of the styles available in the editor.
	 * @returns {boolean}
	 */
	ApiChart.prototype.ApplyChartStyle = function(nStyleId)
	{
		if (typeof(nStyleId) !== "number" || nStyleId < 0)
			return false;

		var nChartType = this.Chart.getChartType();
		var aStyle = AscCommon.g_oChartStyles[nChartType] && AscCommon.g_oChartStyles[nChartType][nStyleId];

		if (aStyle)
		{
			this.Chart.applyChartStyleByIds(aStyle);
			return true;
		}

		return false;
	};

	/**
	 * Sets the fill to the chart plot area.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the plot area.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetPlotAreaFill = function(oFill)
	{
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		this.Chart.SetPlotAreaFill(oFill.UniFill);
		return true;
	};

	/**
	 * Sets the outline to the chart plot area.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the plot area outline.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetPlotAreaOutLine = function(oStroke)
	{
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		this.Chart.SetPlotAreaOutLine(oStroke.Ln);
		return true;
	};

	/**
	 * Sets the fill to the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the series.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {boolean} [bAll=false] - Specifies if the fill will be applied to all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriesFill = function(oFill, nSeries, bAll)
	{
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetSeriesFill(oFill.UniFill, nSeries, bAll);
	};

	/**
	 * Sets the outline to the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the series outline.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {boolean} [bAll=false] - Specifies if the outline will be applied to all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriesOutLine = function(oStroke, nSeries, bAll)
	{
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetSeriesOutLine(oStroke.Ln, nSeries, bAll);
	};

	/**
	 * Sets the fill to the data point in the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the data point.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {number} nDataPoint - The index of the data point in the specified chart series.
	 * @param {boolean} [bAllSeries=false] - Specifies if the fill will be applied to the specified data point in all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetDataPointFill = function(oFill, nSeries, nDataPoint, bAllSeries)
	{
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetDataPointFill(oFill.UniFill, nSeries, nDataPoint, bAllSeries);
	};

	/**
	 * Sets the outline to the data point in the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the data point outline.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {number} nDataPoint - The index of the data point in the specified chart series.
	 * @param {boolean} bAllSeries - Specifies if the outline will be applied to the specified data point in all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetDataPointOutLine = function(oStroke, nSeries, nDataPoint, bAllSeries)
	{
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetDataPointOutLine(oStroke.Ln, nSeries, nDataPoint, bAllSeries);
	};

	/**
	 * Sets the fill to the marker in the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the marker.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {number} nMarker - The index of the marker in the specified chart series.
	 * @param {boolean} [bAllMarkers=false] - Specifies if the fill will be applied to all markers in the specified chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetMarkerFill = function(oFill, nSeries, nMarker, bAllMarkers)
	{
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetMarkerFill(oFill.UniFill, nSeries, nMarker, bAllMarkers);
	};

	/**
	 * Sets the outline to the marker in the specified chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the marker outline.
	 * @param {number} nSeries - The index of the chart series.
	 * @param {number} nMarker - The index of the marker in the specified chart series.
	 * @param {boolean} [bAllMarkers=false] - Specifies if the outline will be applied to all markers in the specified chart series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetMarkerOutLine = function(oStroke, nSeries, nMarker, bAllMarkers)
	{
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetMarkerOutLine(oStroke.Ln, nSeries, nMarker, bAllMarkers);
	};

	/**
	 * Sets the fill to the chart title.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the title.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetTitleFill = function(oFill)
	{
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetTitleFill(oFill.UniFill);
	};

	/**
	 * Sets the outline to the chart title.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the title outline.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetTitleOutLine = function(oStroke)
	{
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetTitleOutLine(oStroke.Ln);
	};

	/**
	 * Sets the fill to the chart legend.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiFill} oFill - The fill type used to fill the legend.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetLegendFill = function(oFill)
	{
		if (!oFill || !oFill.GetClassType || oFill.GetClassType() !== "fill")
			return false;

		return this.Chart.SetLegendFill(oFill.UniFill);
	};

	/**
	 * Sets the outline to the chart legend.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {ApiStroke} oStroke - The stroke used to create the legend outline.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetLegendOutLine = function(oStroke)
	{
		if (!oStroke || !oStroke.GetClassType || oStroke.GetClassType() !== "stroke")
			return false;

		return this.Chart.SetLegendOutLine(oStroke.Ln);
	};

	/**
	 * Sets the specified numeric format to the axis values.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @param {NumFormat | String} sFormat - Numeric format (can be custom format).
	 * @param {AxisPos} - Axis position in the chart.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetAxieNumFormat = function(sFormat, sAxiePos)
	{
		var nAxiePos = -1;
		switch (sAxiePos)
		{
			case "bottom":
				nAxiePos = AscFormat.AX_POS_B;
				break;
			case "left":
				nAxiePos = AscFormat.AX_POS_L;
				break;
			case "right":
				nAxiePos = AscFormat.AX_POS_R;
				break;
			case "top":
				nAxiePos = AscFormat.AX_POS_T;
				break;
			default:
				return false;
		}

		return this.Chart.SetAxieNumFormat(sFormat, nAxiePos);
	};
 
	/**
	 * Sets the specified numeric format to the chart series.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {NumFormat | String} sFormat - Numeric format (can be custom format).
	 * @param {Number} nSeria - Series index.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetSeriaNumFormat = function(sFormat, nSeria)
	{
		return this.Chart.SetSeriaNumFormat(sFormat, nSeria);
	};

	/**
	 * Sets the specified numeric format to the chart data point.
	 * @memberof ApiChart
	 * @typeofeditors ["CDE", "CPE"]
	 * @param {NumFormat | String} sFormat - Numeric format (can be custom format).
	 * @param {Number} nSeria - Series index.
	 * @param {number} nDataPoint - The index of the data point in the specified chart series.
	 * @param {boolean} bAllSeries - Specifies if the numeric format will be applied to the specified data point in all series.
	 * @returns {boolean}
	 */
	ApiChart.prototype.SetDataPointNumFormat = function(sFormat, nSeria, nDataPoint, bAllSeries)
	{
		return this.Chart.SetDataPointNumFormat(sFormat, nSeria, nDataPoint, bAllSeries);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiFill
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiFill class.
	 * @memberof ApiFill
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"fill"}
	 */
	ApiFill.prototype.GetClassType = function()
	{
		return "fill";
	};
	/**
	 * Converts the ApiFill object into the JSON object.
	 * @memberof ApiFill
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiFill.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerFill(this.UniFill));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiStroke
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiStroke class.
	 * @memberof ApiStroke
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"stroke"}
	 */
	ApiStroke.prototype.GetClassType = function()
	{
		return "stroke";
	};
	/**
	 * Converts the ApiStroke object into the JSON object.
	 * @memberof ApiStroke
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiStroke.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerLn(this.Ln));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiGradientStop
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiGradientStop class.
	 * @memberof ApiGradientStop
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"gradientStop"}
	 */
	ApiGradientStop.prototype.GetClassType = function ()
	{
		return "gradientStop"
	};
	/**
	 * Converts the ApiGradientStop object into the JSON object.
	 * @memberof ApiGradientStop
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiGradientStop.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerGradStop(this.Gs));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiUniColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiUniColor class.
	 * @memberof ApiUniColor
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"uniColor"}
	 */
	ApiUniColor.prototype.GetClassType = function ()
	{
		return "uniColor"
	};
	/**
	 * Converts the ApiUniColor object into the JSON object.
	 * @memberof ApiUniColor
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiUniColor.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerColor(this.Unicolor));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiRGBColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiRGBColor class.
	 * @memberof ApiRGBColor
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"rgbColor"}
	 */
	ApiRGBColor.prototype.GetClassType = function ()
	{
		return "rgbColor"
	};
	/**
	 * Converts the ApiRGBColor object into the JSON object.
	 * @memberof ApiRGBColor
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiRGBColor.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerColor(this.Unicolor));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiSchemeColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiSchemeColor class.
	 * @memberof ApiSchemeColor
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"schemeColor"}
	 */
	ApiSchemeColor.prototype.GetClassType = function ()
	{
		return "schemeColor"
	};
	/**
	 * Converts the ApiSchemeColor object into the JSON object.
	 * @memberof ApiSchemeColor
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiSchemeColor.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerColor(this.Unicolor));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPresetColor
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiPresetColor class.
	 * @memberof ApiPresetColor
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {"presetColor"}
	 */
	ApiPresetColor.prototype.GetClassType = function ()
	{
		return "presetColor"
	};
	/**
	 * Converts the ApiPresetColor object into the JSON object.
	 * @memberof ApiPresetColor
	 * @typeofeditors ["CDE"]
	 * @returns {JSON}
	 */
	ApiPresetColor.prototype.ToJSON = function()
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		return JSON.stringify(oWriter.SerColor(this.Unicolor));
	};

	/**
	 * Returns a type of the ApiBullet class.
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
	 * Returns a type of the ApiInlineLvlSdt class.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {"inlineLvlSdt"}
	 */
	ApiInlineLvlSdt.prototype.GetClassType = function()
	{
		return "inlineLvlSdt";
	};

	/**
	 * Sets the lock to the current inline text content control:
	 * <b>"contentLocked"</b> - content cannot be edited.
	 * <b>"sdtContentLocked"</b> - content cannot be edited and the container cannot be deleted.
	 * <b>"sdtLocked"</b> - the container cannot be deleted.
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
	 * Returns the lock type of the current container.
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
	 * Adds a string tag to the current inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sTag - The tag which will be added to the current inline text content control.
	 */
	ApiInlineLvlSdt.prototype.SetTag = function(sTag)
	{
		this.Sdt.SetTag(sTag);
	};

	/**
	 * Returns the tag attribute for the current container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiInlineLvlSdt.prototype.GetTag = function()
	{
		return this.Sdt.GetTag();
	};

	/**
	 * Adds a string label to the current inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sLabel - The label which will be added to the current inline text content control. Can be a positive or negative integer from <b>-2147483647</b> to <b>2147483647</b>.
	 */
	ApiInlineLvlSdt.prototype.SetLabel = function(sLabel)
	{
		this.Sdt.SetLabel(sLabel);
	};

	/**
	 * Returns the label attribute for the current container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiInlineLvlSdt.prototype.GetLabel = function()
	{
		return this.Sdt.GetLabel();
	};

	/**
	 * Sets the alias attribute to the current container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sAlias - The alias which will be added to the current inline text content control.
	 */
	ApiInlineLvlSdt.prototype.SetAlias = function(sAlias)
	{
		this.Sdt.SetAlias(sAlias);
	};

	/**
	 * Returns the alias attribute for the current container.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiInlineLvlSdt.prototype.GetAlias = function()
	{
		return this.Sdt.GetAlias();
	};

	/**
	 * Returns a number of elements in the current inline text content control. The text content 
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
	 * Returns an element of the current inline text content control using the position specified.
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
	 * Removes an element using the position specified from the current inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {number} nPos - The position of the element which we want to remove from the current inline text content control.
	 * @returns {boolean}
	 */
	ApiInlineLvlSdt.prototype.RemoveElement = function(nPos)
	{
		if (nPos < 0 || nPos >= this.Sdt.Content.length)
			return false;

		this.Sdt.RemoveFromContent(nPos, 1);
		if (this.Sdt.Content.length === 0)
		{
			this.Sdt.SetShowingPlcHdr(true);
			this.Sdt.private_FillPlaceholderContent();
		}

		return true;
	};

	/**
	 * Removes all the elements from the current inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {boolean} - returns false if control has not elements.
	 */
	ApiInlineLvlSdt.prototype.RemoveAllElements = function()
	{
		if (this.Sdt.Content.length > 0)
		{
			this.Sdt.RemoveFromContent(0, this.Sdt.Content.length);
			if (this.Sdt.Content.length === 0)
			{
				this.Sdt.SetShowingPlcHdr(true);
				this.Sdt.private_FillPlaceholderContent();
			}

			return true;
		}

		return false;
	};

	/**
	 * Adds an element to the inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {ParagraphContent} oElement - The document element which will be added at the position specified. Returns <b>false</b> if the type of *oElement* is not supported by an inline text content control.
	 * @param {number} [nPos] - The position of the element where it will be added to the current inline text content control. If this value is not specified, then the element will be added to the end of the current inline text content control.
	 * @returns {boolean} - returns false if oElement unsupported.
	 */
	ApiInlineLvlSdt.prototype.AddElement = function(oElement, nPos)
	{
		if (!private_IsSupportedParaElement(oElement) || nPos < 0 || nPos > this.Sdt.Content.length)
			return false;

		var oParaElement = oElement.private_GetImpl();
		if (oParaElement.IsUseInDocument())
			return false;

		if (this.Sdt.IsShowingPlcHdr())
		{
			this.Sdt.RemoveFromContent(0, this.Sdt.GetElementsCount(), false);
			this.Sdt.SetShowingPlcHdr(false);
		}

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
	 * Adds an element to the end of inline text content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {DocumentElement} oElement - The document element which will be added to the end of the container.
	 * @returns {boolean} - returns false if oElement unsupported.
	 */
	ApiInlineLvlSdt.prototype.Push = function(oElement)
	{
		if (!private_IsSupportedParaElement(oElement))
			return false;

		var oParaElement = oElement.private_GetImpl();
		if (oParaElement.IsUseInDocument())
			return false;

		if (this.Sdt.IsShowingPlcHdr())
		{
			this.Sdt.RemoveFromContent(0, this.Sdt.GetElementsCount(), false);
			this.Sdt.SetShowingPlcHdr(false);
		}

		this.Sdt.AddToContent(this.Sdt.GetElementsCount(), oParaElement);

		return true;
	};

	/**
	 * Adds text to the current content control. 
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {String} sText - The text which will be added to the content control.
	 * @returns {boolean} - returns false if param is invalid.
	 */
	ApiInlineLvlSdt.prototype.AddText = function(sText)
	{
		if (typeof sText === "string")
		{
			if (this.Sdt.IsShowingPlcHdr())
			{
				this.Sdt.RemoveFromContent(0, this.Sdt.GetElementsCount(), false);
				this.Sdt.SetShowingPlcHdr(false);
			}

			var newRun = editor.CreateRun();
			newRun.AddText(sText);
			this.AddElement(newRun, this.GetElementsCount())

			return true;
		}

		return false;
	};

	/**
	 * Removes a content control and its content. If keepContent is true, the content is not deleted.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {boolean} keepContent - Specifies if the content will be deleted or not.
	 * @returns {boolean} - returns false if control haven't parent paragraph.
	 */
	ApiInlineLvlSdt.prototype.Delete = function(keepContent)
	{
		var oParentPara = this.Sdt.GetParagraph();
		if (oParentPara)
		{
			if (keepContent)
			{
				this.Sdt.RemoveContentControlWrapper();
			}
			else 
			{
				this.Sdt.PreDelete();
				var nPosInPara = this.Sdt.GetPosInParent();
				oParentPara.RemoveFromContent(nPosInPara, 1);
			}

			return true;
		}

		return false;
	};

	/**
	 * Applies text settings to the content of the content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The properties that will be set to the content of the content control.
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
	 * Returns a paragraph that contains the current content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiParagraph | null} - returns null if parent paragraph doesn't exist.
	 */
	ApiInlineLvlSdt.prototype.GetParentParagraph = function()
	{
		var oPara = this.Sdt.GetParagraph();

		if (oPara)
			return new ApiParagraph(oPara);

		return null; 
	};

	/**
	 * Returns a content control that contains the current content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiBlockLvlSdt | ApiInlineLvlSdt | null} - returns null if parent content control doesn't exist.
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
	 * Returns a table that contains the current content control.
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
	 * Returns a table cell that contains the current content control.
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
	 * Returns a Range object that represents the part of the document contained in the specified content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiInlineLvlSdt.prototype.GetRange = function(Start, End)
	{
		return new ApiRange(this.Sdt, Start, End);
	};

	/**
	 * Creates a copy of an inline content control. Ignores comments, footnote references, complex fields.
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

	/**
	 * Converts the ApiInlineLvlSdt object into the JSON object.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @return {JSON}
	 */
	ApiInlineLvlSdt.prototype.ToJSON = function(bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerInlineLvlSdt(this.Sdt);
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};

	/**
	 * Returns the placeholder text from the current inline content control.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiInlineLvlSdt.prototype.GetPlaceholderText = function()
	{
		return this.Sdt.GetPlaceholderText();
	};

	/**
	 * Sets the placeholder text to the current inline content control.
	 * *Can't be set to checkbox or radio button*
	 * @memberof ApiInlineLvlSdt
	 * @param {string} sText - The text that will be set to the current inline content control.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiInlineLvlSdt.prototype.SetPlaceholderText = function(sText)
	{
		if (typeof(sText) !== "string" || sText === "")
			return false;
		if (this.Sdt.IsCheckBox() || this.Sdt.IsRadioButton())
			return false;

		this.Sdt.SetPlaceholderText(sText);
		if (this.Sdt.IsEmpty())
			this.Sdt.private_ReplaceContentWithPlaceHolder();
		return true;
	};
	/**
	 * Checks if the content control is a form.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiInlineLvlSdt.prototype.IsForm = function()
	{
		return this.Sdt.IsForm();
	};

	/**
	 * Adds a comment to the current inline content control.
	 * <note>Please note that this inline content control must be in the document.</note>
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment text (required).
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {ApiComment?} - Returns null if the comment was not added.
	 */
	ApiInlineLvlSdt.prototype.AddComment = function(sText, sAuthor)
	{
		if (!sText || typeof(sText) !== "string")
			return null;
		if (typeof(sAuthor) !== "string")
			sAuthor = "";

		if (!this.Sdt.IsUseInDocument())
			return null;

		var oDocument = private_GetLogicDocument();
		var CommentData = new AscCommon.CCommentData();

		CommentData.SetText(sText);
		if (sAuthor !== "")
			CommentData.SetUserName(sAuthor);

		var oDocumentState = oDocument.SaveDocumentState();
		this.Sdt.SelectContentControl();

		let comment = AddCommentToDocument(oDocument, CommentData);
		oDocument.LoadDocumentState(oDocumentState);
		oDocument.UpdateSelection();
		return new ApiComment(comment)
	};

	/**
	 * Checks if the content control is a form.
	 * @memberof ApiInlineLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {ApiContentControlList}
	 */
	ApiInlineLvlSdt.prototype.GetDropdownList = function()
	{
		if (!this.Sdt.IsComboBox() && !this.Sdt.IsDropDownList())
			throw "Not a drop down content control";
		
		return new ApiContentControlList(this);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiContentControlList
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiContentControlList class.
	 * @memberof ApiContentControlList
	 * @typeofeditors ["CDE"]
	 * @returns {"contentControlList"}
	 */
	ApiContentControlList.prototype.GetClassType = function()
	{
		return "contentControlList";
	};

	/**
	 * Returns a collection of the ApiContentControlListEntry of a list.
	 * @memberof ApiContentControlList
	 * @typeofeditors ["CDE"]
	 * @returns {ApiContentControlListEntry[]}
	 */
	ApiContentControlList.prototype.GetAllItems = function()
	{
		let nCount	= this.GetElementsCount();
		let aResult	= [];
		
		for (let i = 0; i < nCount; ++i)
		{
			aResult.push(this.GetItem(i));
		}
		
		return aResult;
	};

	/**
	 * Gets items count of the list of values of combobox/listbox.
	 * @memberof ApiContentControlList
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiContentControlList.prototype.GetElementsCount = function()
	{
		return this.GetListPr().GetItemsCount();
	};

	/**
	 * Gets parent content control.
	 * @memberof ApiContentControlList
	 * @typeofeditors ["CDE"]
	 * @returns {ApiInlineLvlSdt | ApiBlockLvlSdt}
	 */
	ApiContentControlList.prototype.GetParent = function()
	{
		return this.Parent;
	};

	/**
	 * Adds the new value to list of values of combobox/dropdown list.
	 * @memberof ApiContentControlList
	 * @param {string} sText - Specifies the display text for the list item..
	 * @param {string} [sValue=sText] - Specifies the value of the list item.
	 * @param {number} [nIndex=this.GetElementsCount()] - position to add.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiContentControlList.prototype.Add = function(sText, sValue, nIndex)
	{
		let nItemsCount = this.GetElementsCount();
		
		sText = AscBuilder.GetStringParameter(sText, null);
		if (!sText)
			return false;
		
		sValue = AscBuilder.GetStringParameter(sValue, sText);
		nIndex = AscBuilder.GetNumberParameter(nIndex, nItemsCount);
		if (nIndex < 0 || nIndex > nItemsCount)
			nIndex = nItemsCount;

		let listPr = this.GetListPr().Copy();
		if (!listPr.AddItem(sText, sValue, nIndex))
			return false;
		
		this.SetListPr(listPr);
		return true;
	};

	/**
	 * Clears the list of values of combobox/dropdown list.
	 * @memberof ApiContentControlList
	 * @typeofeditors ["CDE"]
	 */
	ApiContentControlList.prototype.Clear = function()
	{
		let listPr = this.GetListPr().Copy();
		listPr.Clear();
		this.SetListPr(listPr)
		this.Sdt.SelectListItem("");
	};

	/**
	 * Gets the item of values of combobox/dropdown list.
	 * @memberof ApiContentControlList
	 * @param {number} nIndex
	 * @typeofeditors ["CDE"]
	 * @returns {ApiContentControlListEntry}
	 */
	ApiContentControlList.prototype.GetItem = function(nIndex)
	{
		let listPr = this.GetListPr();
		
		nIndex = AscBuilder.GetNumberParameter(nIndex, null);
		if (null === nIndex)
			throw "Index must be a number";
		else if (nIndex < 0 || nIndex >= listPr.GetItemsCount())
			throw "Index out of list range";
		
		return new ApiContentControlListEntry(this.Sdt, this, listPr.GetItemDisplayText(nIndex), listPr.GetItemValue(nIndex));
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiContentControlListEntry
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiContentControlListEntry class.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {"contentControlList"}
	 */
	ApiContentControlListEntry.prototype.GetClassType = function()
	{
		return "contentControlListEntry";
	};

	/**
	 * Returns the parent of a content control list item in the collection of list items.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {ApiContentControlList}
	 */
	ApiContentControlListEntry.prototype.GetParent = function()
	{
		if (-1 === this.GetIndex())
			return null;
		
		return this.Parent;
	};

	/**
	 * Selects the list entry in a drop-down list or combo box content control and sets the text of the content control to the value of the item.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiContentControlListEntry.prototype.Select = function()
	{
		if (-1 === this.GetIndex())
			return false;
		
		this.Sdt.SelectListItem(this.GetValue());
		return true;
	};

	/**
	 * Moves the current item in the parent drop-down list or combo box content control up one element.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiContentControlListEntry.prototype.MoveUp = function()
	{
		let nCurIndex = this.GetIndex();
		if (-1 === nCurIndex || 0 === nCurIndex)
			return false;
		
		return this.SetIndex(nCurIndex - 1);
	};

	/**
	 * Moves an item in a drop-down list or combo box content control down one item, so that it is after the item that originally followed it.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiContentControlListEntry.prototype.MoveDown = function()
	{
		let nCurIndex = this.GetIndex();
		let nItemsCount = this.Parent.GetListPr().GetItemsCount();
		if (-1 === nCurIndex || nCurIndex >= nItemsCount - 1)
			return false;
		
		return this.SetIndex(nCurIndex + 1);
	};

	/**
	 * Returns the index of a content control list item in the collection of list items.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {number}
	 */
	ApiContentControlListEntry.prototype.GetIndex = function()
	{
		let listPr = this.Parent.GetListPr();
		return listPr.GetIndex(this.Value);
	};

	/**
	 * Sets the index of a content control list item in the collection of list items.
	 * @memberof ApiContentControlListEntry
	 * @param {number} nIndex
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiContentControlListEntry.prototype.SetIndex = function(nIndex)
	{
		let nCurIndex = this.GetIndex();
		if (-1 === nCurIndex)
			return false;
		else if (nIndex === nCurIndex)
			return true;
		
		let listPr = this.Parent.GetListPr().Copy();
		listPr.RemoveItem(nCurIndex);
		listPr.AddItem(this.Text, this.Value, nIndex);
		this.Parent.SetListPr(listPr);
		
		return true;
	};

	/**
	 * Deletes the specified item in a combo box or drop-down list content control..
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiContentControlListEntry.prototype.Delete = function()
	{
		let nCurIndex = this.GetIndex();
		if (nCurIndex === -1)
			return false;
		
		let listPr = this.Parent.GetListPr().Copy();
		listPr.RemoveItem(nCurIndex);
		this.Parent.SetListPr(listPr);
		
		if (listPr.GetItemsCount())
		{
			this.Parent.GetItem(0).Select();
		}
		else
		{
			this.Sdt.SetShowingPlcHdr(true)
			this.Sdt.ReplaceContentWithPlaceHolder();
		}
		
		return true;
	};

	/**
	 * Returns a String that represents the display text of a list item for a drop-down list or combo box content control.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiContentControlListEntry.prototype.GetText = function()
	{
		return this.Text;
	};

	/**
	 * Sets a String that represents the display text of a list item for a drop-down list or combo box content control.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @param {string} sText
	 * @returns {boolean}
	 */
	ApiContentControlListEntry.prototype.SetText = function(sText)
	{
		let nCurIndex = this.GetIndex();
		if (this.GetValue() === "" || -1 === nCurIndex)
			return false;
		
		sText = AscBuilder.GetStringParameter(sText, null);
		if (null === sText)
			return false;
		
		this.Text = sText;
		
		let listPr = this.Parent.GetListPr().Copy();
		listPr.RemoveItem(nCurIndex);
		listPr.AddItem(this.Text, this.Value, nCurIndex);
		this.Parent.SetListPr(listPr);
		this.Select();
		return true;
	};

	/**
	 * Returns a String that represents the value of a list item for a drop-down list or combo box content control.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiContentControlListEntry.prototype.GetValue = function()
	{
		return this.Value;
	};

	/**
	 * Sets a String that represents the value of a list item for a drop-down list or combo box content control.
	 * @memberof ApiContentControlListEntry
	 * @typeofeditors ["CDE"]
	 * @param {string} sValue
	 * @returns {boolean}
	 */
	ApiContentControlListEntry.prototype.SetValue = function(sValue)
	{
		let nCurIndex = this.GetIndex();
		sValue = AscBuilder.GetStringParameter(sValue, null);
		
		if (null === sValue
			|| this.GetValue() === ""
			|| this.GetValue() === sValue
			|| -1 === nCurIndex
			|| -1 !== this.Parent.GetListPr().GetIndex(sValue))
			return false;
		
		this.Value = sValue;
		
		let listPr = this.Parent.GetListPr().Copy();
		listPr.RemoveItem(nCurIndex);
		listPr.AddItem(this.Text, this.Value, nCurIndex);
		this.Parent.SetListPr(listPr);
		this.Select();
		return true;
	};
	
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiBlockLvlSdt
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiBlockLvlSdt class.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {"blockLvlSdt"}
	 */
	ApiBlockLvlSdt.prototype.GetClassType = function()
	{
		return "blockLvlSdt";
	};

	/**
	 * Sets the lock to the current block text content control:
	 * <b>"contentLocked"</b> - content cannot be edited.
	 * <b>"sdtContentLocked"</b> - content cannot be edited and the container cannot be deleted.
	 * <b>"sdtLocked"</b> - the container cannot be deleted.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {"contentLocked" | "sdtContentLocked" | "sdtLocked"} sLockType - The type of the lock applied to the block text content control.
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
	 * Returns the lock type of the current container.
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
	 * Sets the tag attribute to the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sTag - The tag which will be added to the current container.
	 */
	ApiBlockLvlSdt.prototype.SetTag = function(sTag)
	{
		this.Sdt.SetTag(sTag);
	};

	/**
	 * Returns the tag attribute for the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiBlockLvlSdt.prototype.GetTag = function()
	{
		return this.Sdt.GetTag();
	};

	/**
	 * Sets the label attribute to the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sLabel - The label which will be added to the current container. Can be a positive or negative integer from <b>-2147483647</b> to <b>2147483647</b>.
	 */
	ApiBlockLvlSdt.prototype.SetLabel = function(sLabel)
	{
		this.Sdt.SetLabel(sLabel);
	};

	/**
	 * Returns the label attribute for the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiBlockLvlSdt.prototype.GetLabel = function()
	{
		return this.Sdt.GetLabel();
	};

	/**
	 * Sets the alias attribute to the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sAlias - The alias which will be added to the current container.
	 */
	ApiBlockLvlSdt.prototype.SetAlias = function(sAlias)
	{
		this.Sdt.SetAlias(sAlias);
	};

	/**
	 * Returns the alias attribute for the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiBlockLvlSdt.prototype.GetAlias = function()
	{
		return this.Sdt.GetAlias();
	};

	/**
	 * Returns the content of the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {ApiDocumentContent}
	 */
	ApiBlockLvlSdt.prototype.GetContent = function()
	{
		return new ApiDocumentContent(this.Sdt.GetContent());
	};

	/**
	 * Returns a collection of content control objects in the current content control.
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
			let oControl = ToApiContentControl(arrContentControls[Index]);
			if (oControl)
				arrApiContentControls.push(oControl);
		}

		return arrApiContentControls;
	};

	/**
	 * Returns a collection of paragraph objects in the current content control.
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
	 * Returns a collection of tables on a given absolute page.
	 * <note>This method can be a little bit slow, because it runs the document calculation
	 * process to arrange tables on the specified page.</note>
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param nPage - Page number. If it is not specified, an empty array will be returned.
	 * @return {ApiTable[]}  
	 */
	ApiBlockLvlSdt.prototype.GetAllTablesOnPage = function(nPage)
	{
		var oLogicDocument = this.Sdt.GetLogicDocument();
		if (oLogicDocument)
			(new ApiDocument(oLogicDocument)).ForceRecalculate(nPage + 1);

		var arrTables		= this.Sdt.GetAllTablesOnPage(nPage);
		var arrApiTables	= [];

		for (var Index = 0, nCount = arrTables.length; Index < nCount; Index++)
		{
			arrApiTables.push(new ApiTable(arrTables[Index].Table));
		}

		return arrApiTables;
	};

	/**
	 * Clears the contents from the current content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {boolean} - returns true.
	 */
	ApiBlockLvlSdt.prototype.RemoveAllElements = function()
	{
		this.Sdt.ReplaceContentWithPlaceHolder(false);
		return true;
	};

	/**
	 * Removes a content control and its content. If keepContent is true, the content is not deleted.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {boolean} keepContent - Specifies if the content will be deleted or not.
	 * @returns {boolean} - returns false if content control haven't parent.
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
	 * Applies text settings to the content of the content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} oTextPr - The properties that will be set to the content of the content control.
	 */
	ApiBlockLvlSdt.prototype.SetTextPr = function(oTextPr)
	{
		var ParaTextPr = new AscCommonWord.ParaTextPr(oTextPr.TextPr);
		this.Sdt.Content.SetApplyToAll(true);
		this.Sdt.Add(ParaTextPr);
		this.Sdt.Content.SetApplyToAll(false);
	};

	/**
	 * Returns a collection of drawing objects in the current content control.
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
	 * Returns a content control that contains the current content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @return {ApiBlockLvlSdt | null} - returns null if parent content control doesn't exist.  
	 */
	ApiBlockLvlSdt.prototype.GetParentContentControl = function()
	{
		var documentPos = this.Sdt.GetDocumentPositionFromObject();

		for (var Index = documentPos.length - 1; Index >= 1; Index--)
		{
			if (documentPos[Index].Class && documentPos[Index].Class.Parent && documentPos[Index].Class.Parent instanceof CBlockLevelSdt)
				return new ApiBlockLvlSdt(documentPos[Index].Class.Parent);
		}

		return null;
	};

	/**
	 * Returns a table that contains the current content control.
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
	 * Returns a table cell that contains the current content control.
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
	 * Pushes a paragraph or a table or a block content control to actually add it to the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {DocumentElement} oElement - The type of the element which will be pushed to the current container.
	 * @return {boolean} - returns false if oElement unsupported.
	 */
	ApiBlockLvlSdt.prototype.Push = function(oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;

			if (this.Sdt.IsShowingPlcHdr())
			{
				this.Sdt.Content.RemoveFromContent(0, this.Sdt.Content.GetElementsCount(), false);
				this.Sdt.SetShowingPlcHdr(false);
			}
			
			this.Sdt.Content.Internal_Content_Add(this.Sdt.Content.Content.length, oElm);
			return true;
		}

		return false;
	};

	/**
	 * Adds a paragraph or a table or a block content control to the current container.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {DocumentElement} oElement - The type of the element which will be added to the current container.
	 * @param {Number} nPos - The specified position.
	 * @return {boolean} - returns false if oElement unsupported.
	 */
	ApiBlockLvlSdt.prototype.AddElement = function(oElement, nPos)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;

			if (this.Sdt.IsShowingPlcHdr())
			{
				this.Sdt.Content.RemoveFromContent(0, this.Sdt.Content.GetElementsCount(), false);
				this.Sdt.SetShowingPlcHdr(false);
			}
			
			this.Sdt.Content.Internal_Content_Add(nPos, oElm);
			return true;
		}

		return false;
	};

	/**
	 * Adds a text to the current content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {String} sText - The text which will be added to the content control.
	 * @return {boolean} - returns false if param is invalid.
	 */
	ApiBlockLvlSdt.prototype.AddText = function(sText)
	{
		let _sText = GetStringParameter(sText, null);
		if (null === _sText)
			return false;

		let oParagraph;
		if (this.Sdt.IsPlaceHolder())
		{
			this.Sdt.ReplacePlaceHolderWithContent();
			let oDocContent = this.GetContent();
			if (oDocContent.GetElementsCount() && oDocContent.GetElement(0) instanceof ApiParagraph)
				oParagraph = oDocContent.GetElement(0);
		}

		if (!oParagraph)
		{
			oParagraph = Api.prototype.CreateParagraph();
			this.GetContent().Push(oParagraph);
		}

		oParagraph.AddText(_sText);
		return true;
	};

	/**
	 * Returns a Range object that represents the part of the document contained in the specified content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {Number} Start - Start character in the current element.
	 * @param {Number} End - End character in the current element.
	 * @returns {ApiRange} 
	 * */
	ApiBlockLvlSdt.prototype.GetRange = function(Start, End)
	{
		return new ApiRange(this.Sdt, Start, End);
	};

	/**
	 * Searches for a scope of a content control object. The search results are a collection of ApiRange objects.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - Search string.
	 * @param {boolean} isMatchCase - Case sensitive or not.
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
	 * Selects the current content control.
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
	 * Returns the placeholder text from the current content control.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiBlockLvlSdt.prototype.GetPlaceholderText = function()
	{
		return this.Sdt.GetPlaceholderText();
	};

	/**
	 * Sets the placeholder text to the current content control.
	 * @memberof ApiBlockLvlSdt
	 * @param {string} sText - The text that will be set to the current content control.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiBlockLvlSdt.prototype.SetPlaceholderText = function(sText)
	{
		if (typeof(sText) !== "string" || sText === "")
			return false;

		this.Sdt.SetPlaceholderText(sText);
		if (this.Sdt.IsEmpty())
			this.Sdt.private_ReplaceContentWithPlaceHolder();

		return true;
	};

	/**
     * Returns the content control position within its parent element.
     * @memberof ApiBlockLvlSdt
     * @typeofeditors ["CDE"]
     * @returns {Number} - returns -1 if the content control parent doesn't exist. 
     */
	ApiBlockLvlSdt.prototype.GetPosInParent = function()
	{
		return this.Sdt.GetIndex();
	};

	/**
	 * Replaces the current content control with a new element.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {DocumentElement} oElement - The element to replace the current content control with.
	 * @returns {boolean}
	 */
	ApiBlockLvlSdt.prototype.ReplaceByElement = function(oElement)
	{
		if (oElement instanceof ApiParagraph || oElement instanceof ApiTable || oElement instanceof ApiBlockLvlSdt)
		{
			var oElm = oElement.private_GetImpl();
			if (oElm.IsUseInDocument())
				return false;

			var oParent = this.Sdt.GetParent();
			var nCCPos = this.Sdt.GetIndex();
			if (oParent && nCCPos !== -1)
			{
				this.Delete();
				oParent.Internal_Content_Add(nCCPos, oElm);
				return true;
			}
		}

		return false;
	};

	/**
	 * Adds a comment to the current block content control.
	 * <note>Please note that the current block content control must be in the document.</note>
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment text (required).
	 * @param {string} sAuthor - The author's name (optional).
	 * @returns {?ApiComment} - Returns null if the comment was not added.
	 */
	ApiBlockLvlSdt.prototype.AddComment = function(sText, sAuthor)
	{
		if (!sText || typeof(sText) !== "string")
			return null;
		if (typeof(sAuthor) !== "string")
			sAuthor = "";

		if (!this.Sdt.IsUseInDocument())
			return null;

		var oDocument = private_GetLogicDocument();
		var CommentData = new AscCommon.CCommentData();

		CommentData.SetText(sText);
		if (sAuthor !== "")
			CommentData.SetUserName(sAuthor);

		var oDocumentState = oDocument.SaveDocumentState();
		this.Sdt.SelectContentControl();

		let comment = AddCommentToDocument(oDocument, CommentData);
		oDocument.LoadDocumentState(oDocumentState);
		oDocument.UpdateSelection();
		
		return comment;
	};

	/**
	 * Sets the background color to the current block content control.
	 * @memberof ApiBlockLvlSdt
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} bNone - Defines that background color will not be set.
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiBlockLvlSdt.prototype.SetBackgroundColor = function(r, g, b, bNone)
	{
		var oFormPr = this.Sdt.GetFormPr().Copy();
		
		let oUnifill = new AscFormat.CUniFill();
		oUnifill.setFill(new AscFormat.CSolidFill());
		oUnifill.fill.setColor(new AscFormat.CUniColor());
		oUnifill.fill.color.setColor(new AscFormat.CRGBColor());

		if (r >=0 && g >=0 && b >=0)
			oUnifill.fill.color.color.setColor(r, g, b);
		else
			return false;

		oFormPr.Shd = new CDocumentShd();
		oFormPr.Shd.Set_FromObject({
			Value: bNone ? Asc.c_oAscShd.Clear : Asc.c_oAscShd.Clear,
			Color: {
				r: r,
				g: g,
				b: b,
				Auto: false
			},
			Fill: {
				r: r,
				g: g,
				b: b,
				Auto: false
			},
			Unifill: oUnifill
		});

		this.Sdt.SetFormPr(oFormPr);
		return true;
	};

	/**
     * Adds a caption paragraph after (or before) the current content control.
	 * <note>Please note that the current content control must be in the document (not in the footer/header).
	 * And if the current content control is placed in a shape, then a caption is added after (or before) the parent shape.</note>
     * @memberof ApiBlockLvlSdt
     * @typeofeditors ["CDE"]
     * @param {string} sAdditional - The additional text.
	 * @param {CaptionLabel | String} [sLabel="Table"] - The caption label.
	 * @param {boolean} [bExludeLabel=false] - Specifies whether to exclude the label from the caption.
	 * @param {CaptionNumberingFormat} [sNumberingFormat="Arabic"] - The possible caption numbering format.
	 * @param {boolean} [bBefore=false] - Specifies whether to insert the caption before the current content control (true) or after (false) (after/before the shape if it is placed in the shape).
	 * @param {Number} [nHeadingLvl=undefined] - The heading level (used if you want to specify the chapter number).
	 * <note>If you want to specify "Heading 1", then nHeadingLvl === 0 and etc.</note>
	 * @param {CaptionSep} [sCaptionSep="hyphen"] - The caption separator (used if you want to specify the chapter number).
     * @returns {boolean}
     */
	ApiBlockLvlSdt.prototype.AddCaption = function(sAdditional, sLabel, bExludeLabel, sNumberingFormat, bBefore, nHeadingLvl, sCaptionSep)
	{
		var oSdtParent = this.Sdt.GetParent();
		if (this.Sdt.IsUseInDocument() === false || !oSdtParent || oSdtParent.Is_TopDocument(true) !== private_GetLogicDocument())
			return false;
		if (typeof(sAdditional) !== "string" || sAdditional.trim() === "")
			sAdditional = "";
		if (typeof(bExludeLabel) !== "boolean")
			bExludeLabel = false;
		if (typeof(bBefore) !== "boolean")
			bBefore = false;
		if (typeof(sLabel) !== "string" || sLabel.trim() === "")
			sLabel = "Table";
		
		let oCapPr = new Asc.CAscCaptionProperties();
		let oDoc = private_GetLogicDocument();

		let nNumFormat;
		switch (sNumberingFormat)
		{
			case "ALPHABETIC":
				nNumFormat = Asc.c_oAscNumberingFormat.UpperLetter;
				break;
			case "alphabetic":
				nNumFormat = Asc.c_oAscNumberingFormat.LowerLetter;
				break;
			case "Roman":
				nNumFormat = Asc.c_oAscNumberingFormat.UpperRoman;
				break;
			case "roman":
				nNumFormat = Asc.c_oAscNumberingFormat.LowerRoman;
				break;
			default:
				nNumFormat = Asc.c_oAscNumberingFormat.Decimal;
				break;
		}
		switch (sCaptionSep)
		{
			case "hyphen":
				sCaptionSep = "-";
				break;
			case "period":
				sCaptionSep = ".";
				break;
			case "colon":
				sCaptionSep = ":";
				break;
			case "longDash":
				sCaptionSep = "—";
				break;
			case "dash":
				sCaptionSep = "-";
				break;
			default:
				sCaptionSep = "-";
				break;
		}

		oCapPr.Label = sLabel;
		oCapPr.Before = bBefore;
		oCapPr.ExcludeLabel = bExludeLabel;
		oCapPr.NumFormat = nNumFormat;
		oCapPr.Separator = sCaptionSep;
		oCapPr.Additional = sAdditional;

		if (nHeadingLvl >= 0 && nHeadingLvl <= 8)
		{
			oCapPr.HeadingLvl = nHeadingLvl;
			oCapPr.IncludeChapterNumber = true;
		}
		else oCapPr.HeadingLvl = 0;

		this.Sdt.SetThisElementCurrent();

		oDoc.AddCaption(oCapPr);
		return true;
	};
	
	/**
	 * Checks if the content control is a form.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @returns {ApiContentControlList}
	 */
	ApiBlockLvlSdt.prototype.GetDropdownList = function()
	{
		if (!this.Sdt.IsComboBox() && !this.Sdt.IsDropDownList())
			throw "Not a drop down content control";
		
		return new ApiContentControlList(this);
	};
	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiFormBase
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiFormBase class.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {"form"}
	 */
	ApiFormBase.prototype.GetClassType = function()
	{
		return "form";
	};
	/**
	 * Returns a type of the current form.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {FormType}
	 */
	ApiFormBase.prototype.GetFormType = function()
	{
		if (this.Sdt.IsTextForm())
			return "textForm";
		if (this.Sdt.IsComboBox())
			return "comboBoxForm";
		if (this.Sdt.IsDropDownList())
			return "dropDownForm";
		if (this.Sdt.IsRadioButton())
			return "radioButtonForm";
		if (this.Sdt.IsCheckBox())
			return "checkBoxForm";
		if (this.Sdt.IsPictureForm())
			return "pictureForm";
	};
	/**
	 * Returns the current form key.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {string}
	 */
	ApiFormBase.prototype.GetFormKey = function()
	{
		var sKey;
		if (this.GetFormType() === "radioButtonForm")
			sKey = this.Sdt.GetRadioButtonGroupKey();
		else
			sKey = this.Sdt.GetFormKey();

		if (typeof(sKey) !== "string")
			sKey = "";
		return sKey;
	};
	/**
	 * Sets a key to the current form.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @param {string} sKey - Form key.
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.SetFormKey = function(sKey)
	{
		if (typeof(sKey) !== "string")
			return false;

		if (this.GetFormType() === "radioButtonForm")
		{
			sKey = sKey === "" ? "Group 1" : sKey;
			this.Sdt.GetCheckBoxPr().SetGroupKey(sKey);
		}
		else
		{
			sKey = sKey === "" ? undefined : sKey;
			var oFormPr = this.Sdt.GetFormPr().Copy();
			oFormPr && oFormPr.SetKey(sKey);
			this.Sdt.SetFormPr(oFormPr);
		}
		
		return true;
	};
	/**
	 * Returns the tip text of the current form.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {string}
	 */
	ApiFormBase.prototype.GetTipText = function()
	{
		var oFormPr = this.Sdt.GetFormPr();
		var sTip = oFormPr.HelpText;
		if (typeof(sTip) !== "string")
			sTip = "";

		return sTip;
	};
	/**
	 * Sets the tip text to the current form.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @param {string} sText - Tip text.
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.SetTipText = function(sText)
	{
		if (typeof(sText) !== "string")
			return false;

		var oFormPr = this.Sdt.GetFormPr().Copy();
		oFormPr && oFormPr.SetHelpText(sText);

		this.Sdt.SetFormPr(oFormPr);
		return true;
	};
	/**
	 * Checks if the current form is required.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.IsRequired = function()
	{
		return this.Sdt.IsFormRequired();
	};
	/**
	 * Specifies if the current form should be required.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @param {boolean} bRequired - Defines if the current form is required (true) or not (false).
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.SetRequired = function(bRequired)
	{
		if (typeof(bRequired) !== "boolean")
			return false;
		if (bRequired === this.IsRequired())
			return true;

		var oFormPr = this.Sdt.GetFormPr().Copy();
		oFormPr && oFormPr.SetRequired(bRequired);

		this.Sdt.SetFormPr(oFormPr);
		return true;
	};
	/**
	 * Checks if the current form is fixed size.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.IsFixed = function()
	{
		return (this.GetFormType() === "pictureForm" || this.Sdt.IsFixedForm());
	};
	/**
	 * Converts the current form to a fixed size form.
	 * @memberof ApiFormBase
	 * @param {twips} nWidth - The wrapper shape width measured in twentieths of a point (1/1440 of an inch).
	 * @param {twips} nHeight - The wrapper shape height measured in twentieths of a point (1/1440 of an inch).
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.ToFixed = function(nWidth, nHeight)
	{
		if (this.IsFixed())
			return false;
		
		this.Sdt.ConvertFormToFixed(private_Twips2MM(nWidth), private_Twips2MM(nHeight));
		return true;
	};
	/**
	 * Converts the current form to an inline form.
	 * *Picture form can't be converted to an inline form, it's always a fixed size object.*
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.ToInline = function()
	{
		if (!this.IsFixed())
			return false;

		this.Sdt.ConvertFormToInline();
		return true;
	};
	/**
	 * Sets the border color to the current form.
	 * @memberof ApiFormBase
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} bNone - Defines that border color will not be set.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.SetBorderColor = function(r, g, b, bNone)
	{
		var oFormPr = this.Sdt.GetFormPr().Copy();
		var oBorder;
		if (typeof(r) == "number" && typeof(g) == "number" && typeof(b) == "number" && !bNone)
		{
			oBorder = new CDocumentBorder();
			oBorder.Color = new CDocumentColor(r, g, b);
		}
		else if (bNone)
			oBorder = undefined;
		else
			return false;

		oFormPr.Border = oBorder;

		this.Sdt.SetFormPr(oFormPr);
		return true;
	};
	/**
	 * Sets the background color to the current form.
	 * @memberof ApiFormBase
	 * @param {byte} r - Red color component value.
	 * @param {byte} g - Green color component value.
	 * @param {byte} b - Blue color component value.
	 * @param {boolean} bNone - Defines that background color will not be set.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.SetBackgroundColor = function(r, g, b, bNone)
	{
		var oFormPr = this.Sdt.GetFormPr().Copy();
		
		let oUnifill = new AscFormat.CUniFill();
		oUnifill.setFill(new AscFormat.CSolidFill());
		oUnifill.fill.setColor(new AscFormat.CUniColor());
		oUnifill.fill.color.setColor(new AscFormat.CRGBColor());

		if (r >=0 && g >=0 && b >=0)
			oUnifill.fill.color.color.setColor(r, g, b);
		else
			return false;

		oFormPr.Shd = new CDocumentShd();
		oFormPr.Shd.Set_FromObject({
			Value: bNone ? Asc.c_oAscShd.Clear : Asc.c_oAscShd.Clear,
			Color: {
				r: r,
				g: g,
				b: b,
				Auto: false
			},
			Fill: {
				r: r,
				g: g,
				b: b,
				Auto: false
			},
			Unifill: oUnifill
		});

		this.Sdt.SetFormPr(oFormPr);
		return true;
	};
	/**
	 * Returns the text from the current form.
	 * *This method is used only for text and combo box forms.*
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {string}
	 */
	ApiFormBase.prototype.GetText = function()
	{
		return this.Sdt.GetInnerText();
	};
	/**
	 * Clears the current form.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 */
	ApiFormBase.prototype.Clear = function()
	{
		this.Sdt.ClearContentControlExt();
	};
	/**
	 * Returns a shape in which the form is placed to control the position and size of the fixed size form frame.
	 * The null value will be returned for the inline forms.
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {?ApiShape} - returns the shape in which the form is placed.
	 */
	ApiFormBase.prototype.GetWrapperShape = function()
    {
        var oParagraph = this.Sdt.GetParagraph();
        var oShape     = oParagraph ? oParagraph.Parent.Is_DrawingShape(true) : null;
        if (!oShape || !oShape.parent)
            return null;
        
        return new ApiShape(oShape);
    };
	/**
	 * Sets the placeholder text to the current form.
	 * *Can't be set to checkbox or radio button.*
	 * @memberof ApiFormBase
	 * @param {string} sText - The text that will be set to the current form.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiFormBase.prototype.SetPlaceholderText = function(sText)
	{
		if (typeof(sText) !== "string" || sText === "")
			return false;
		if (this.Sdt.IsCheckBox() || this.Sdt.IsRadioButton())
			return false;
		
		this.Sdt.SetPlaceholderText(sText);
		return true;
	};
	/**
	 * Sets the text properties to the current form.
	 * *This method is used only for text and combo box forms.*
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @param {ApiTextPr} oTextPr - The text properties that will be set to the current form.
	 * @return {boolean}  
	 */
	ApiFormBase.prototype.SetTextPr = function(oTextPr)
	{
		if (oTextPr && oTextPr.GetClassType && oTextPr.GetClassType() === "textPr")
		{
			this.Sdt.Apply_TextPr(oTextPr.TextPr);
			return true;
		}

		return false;
	};
	/**
	 * Returns the text properties from the current form.
	 * *This method is used only for text and combo box forms.*
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @return {ApiTextPr}  
	 */
	ApiFormBase.prototype.GetTextPr = function()
	{
		return new ApiTextPr(this, this.Sdt.Pr.TextPr.Copy());
	};
	/**
	 * Copies the current form (copies with the shape if it exists).
	 * @memberof ApiFormBase
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {?ApiForm}
	 */
	ApiFormBase.prototype.Copy = function()
	{
		let oSdt;
		if (this.IsFixed())
		{
			var oParagraph = this.Sdt.GetParagraph();
			var oShape     = oParagraph.Parent.Is_DrawingShape(true);
			if (!oShape || !oShape.parent || !oShape.isForm())
				return null;

			var oDrawing = oShape.parent.Copy({
				SkipComments          : true,
				SkipAnchors           : true,
				SkipFootnoteReference : true,
				SkipComplexFields     : true
			});
			oSdt = oDrawing.GraphicObj.getInnerForm();
		}
		else
		{
			oSdt = this.Sdt.Copy(false, {
				SkipComments          : true,
				SkipAnchors           : true,
				SkipFootnoteReference : true,
				SkipComplexFields     : true
			});
		}

		if (!oSdt)
			return null;

		return new this.constructor(oSdt);
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiTextForm
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Checks if the text field content is autofit, i.e. whether the font size adjusts to the size of the fixed size form.
	 * @memberof ApiTextForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiTextForm.prototype.IsAutoFit = function()
	{
		return this.Sdt.IsAutoFitContent();
	};
	/**
	 * Specifies if the text field content should be autofit, i.e. whether the font size adjusts to the size of the fixed size form.
	 * @memberof ApiTextForm
	 * @param {boolean} bAutoFit - Defines if the text field content is autofit (true) or not (false).
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiTextForm.prototype.SetAutoFit = function(bAutoFit)
	{
		if (typeof(bAutoFit) !== "boolean" || !this.IsFixed())
			return false;
		if (bAutoFit === this.IsAutoFit())
			return true;

		var oPr = this.Sdt.GetTextFormPr().Copy();
		oPr.SetAutoFit(bAutoFit);

		this.Sdt.SetTextFormPr(oPr);
		return true;
	};
	/**
	 * Checks if the current text field is multiline.
	 * @memberof ApiTextForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiTextForm.prototype.IsMultiline = function()
	{
		return this.Sdt.IsMultiLineForm();
	};
	/**
	 * Specifies if the current text field should be miltiline.
	 * @memberof ApiTextForm
	 * @param {boolean} bMultiline - Defines if the current text field is multiline (true) or not (false).
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean} - return false, if the text field is not fixed size.
	 */
	ApiTextForm.prototype.SetMultiline = function(bMultiline)
	{
		if (typeof(bMultiline) !== "boolean" || !this.IsFixed())
			return false;
		if (!this.IsFixed())
			return false;
		if (bMultiline === this.IsMultiline())
			return true;

		var oPr = this.Sdt.GetTextFormPr().Copy();
		oPr.SetMultiLine(bMultiline);
		this.Sdt.SetTextFormPr(oPr);

		return true;
	};
	/**
	 * Returns a limit of the text field characters.
	 * @memberof ApiTextForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {number} - if this method returns -1 -> the form has no limit for characters
	 */
	ApiTextForm.prototype.GetCharactersLimit = function()
	{
		var oPr = this.Sdt.GetTextFormPr();
		if (!oPr)
			return -1;

		return oPr.GetMaxCharacters();
	};
	/**
	 * Sets a limit to the text field characters.
	 * @memberof ApiTextForm
	 * @param {number} nChars - The maximum number of characters in the text field. If this parameter is equal to -1, no limit will be set.
	 * A limit is required to be set if a comb of characters is applied.
	 * Maximum value for this parameter is 1000000.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiTextForm.prototype.SetCharactersLimit = function(nChars)
	{
		if (typeof(nChars) !== "number")
			return false;

		const nMax = 1000000;
		nChars = nChars > nMax ? nMax : Math.floor(nChars);

		if (nChars <= 0)
			nChars = -1;

		let oPr = this.Sdt.GetTextFormPr();
		if (!oPr || (-1 === nChars && this.IsComb()))
			return false;

		oPr = oPr.Copy();
		oPr.SetMaxCharacters(nChars);

		this.Sdt.SetTextFormPr(oPr);
		return true;
	};
	/**
	 * Checks if the text field is a comb of characters with the same cell width.
	 * @memberof ApiTextForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiTextForm.prototype.IsComb = function()
	{
		let oPr = this.Sdt.GetTextFormPr();
		return oPr ? oPr.IsComb() : false;
	};
	/**
	 * Specifies if the text field should be a comb of characters with the same cell width.
	 * The maximum number of characters must be set to a positive value.
	 * @memberof ApiTextForm
	 * @param {boolean} bComb - Defines if the text field is a comb of characters (true) or not (false).
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiTextForm.prototype.SetComb = function(bComb)
	{
		if (typeof(bComb) !== "boolean")
			return false;

		let oPr = this.Sdt.GetTextFormPr();
		if (!oPr)
			return false;

		if (oPr.IsComb() === bComb)
			return true;

		oPr = oPr.Copy();
		oPr.SetComb(bComb);
		if (oPr.GetMaxCharacters() === -1)
			oPr.SetMaxCharacters(10);
		oPr.SetWidth(0);

		this.Sdt.SetTextFormPr(oPr);
		return true;
	};
	/**
	 * Sets the cell width to the applied comb of characters.
	 * @memberof ApiTextForm
	 * @param {mm} [nCellWidth=0] - The cell width measured in millimeters.
	 * If this parameter is not specified or equal to 0 or less, then the width will be set automatically. Must be >= 1 and <= 558.8.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiTextForm.prototype.SetCellWidth = function(nCellWidth)
	{
		if (typeof(nCellWidth) !== "number" || !this.IsComb())
			return false;

		var nWidthMax = 558.8;
		nCellWidth = nCellWidth < 1 ? 1 : Math.floor(nCellWidth * 100) / 100;
		nCellWidth = nCellWidth > nWidthMax ? nWidthMax : nCellWidth;

		var oPr = this.Sdt.GetTextFormPr().Copy();
		oPr.SetWidth(Math.floor(nCellWidth * 72 * 20 / 25.4 + 0.5));

		this.Sdt.SetTextFormPr(oPr);
		return true;
	};
	/**
	 * Sets the text to the current text field.
	 * @memberof ApiTextForm
	 * @param {string} sText - The text that will be set to the current text field.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiTextForm.prototype.SetText = function(sText)
	{
		let _sText = GetStringParameter(sText, null);
		if (!_sText)
			return false;

		this.Sdt.SetInnerText(_sText);
		this.OnChangeValue();
		return true;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiPictureForm
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns the current scaling condition of the picture form.
	 * @memberof ApiPictureForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {ScaleFlag}
	 */
	ApiPictureForm.prototype.GetScaleFlag = function()
	{
		let sScaleFlag = "always";
		let oPr = this.Sdt.GetPictureFormPr();
		switch (oPr.GetScaleFlag())
		{
			case Asc.c_oAscPictureFormScaleFlag.Always:
				sScaleFlag = "always";
				break;
			case Asc.c_oAscPictureFormScaleFlag.Never:
				sScaleFlag = "never";
				break;
			case Asc.c_oAscPictureFormScaleFlag.Bigger:
				sScaleFlag = "tooBig";
				break;
			case Asc.c_oAscPictureFormScaleFlag.Small:
				sScaleFlag = "tooSmall";
				break;
		}

		return sScaleFlag;
	};
	/**
	 * Sets the scaling condition to the current picture form.
	 * @memberof ApiPictureForm
	 * @param {ScaleFlag} sScaleFlag - Picture scaling condition: "always", "never", "tooBig" or "tooSmall".
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiPictureForm.prototype.SetScaleFlag = function(sScaleFlag)
	{
		let nScaleFlag;
		switch (sScaleFlag)
		{
			case "always":
				nScaleFlag = Asc.c_oAscPictureFormScaleFlag.Always;
				break;
			case "never":
				nScaleFlag = Asc.c_oAscPictureFormScaleFlag.Never;
				break;
			case "tooBig":
				nScaleFlag = Asc.c_oAscPictureFormScaleFlag.Bigger;
				break;
			case "tooSmall":
				nScaleFlag = Asc.c_oAscPictureFormScaleFlag.Small;
				break;
			default:
				return false;
		}

		var oPr = this.Sdt.GetPictureFormPr().Copy();
		oPr.SetScaleFlag(nScaleFlag);
		this.Sdt.SetPictureFormPr(oPr);
		this.Sdt.UpdatePictureFormLayout();
		return true;
	};
	/**
	 * Locks the aspect ratio of the current picture form.
	 * @memberof ApiPictureForm
	 * @param {boolean} [isLock=true] - Specifies if the aspect ratio of the current picture form will be locked (true) or not (false).
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiPictureForm.prototype.SetLockAspectRatio = function(isLock)
	{
		let oPr = this.Sdt.GetPictureFormPr().Copy();
		oPr.SetConstantProportions(GetBoolParameter(isLock, false));
		this.Sdt.SetPictureFormPr(oPr);
		this.Sdt.UpdatePictureFormLayout();
		return true;
	};
	/**
	 * Checks if the aspect ratio of the current picture form is locked or not.
	 * @memberof ApiPictureForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiPictureForm.prototype.IsLockAspectRatio = function()
	{
		return this.Sdt.GetPictureFormPr().IsConstantProportions();
	};
	/**
	 * Sets the picture position inside the current form:
	 * * <b>0</b> - the picture is placed on the left/top;
	 * * <b>50</b> - the picture is placed in the center;
	 * * <b>100</b> - the picture is placed on the right/bottom.
	 * @memberof ApiPictureForm
	 * @param {percentage} nShiftX - Horizontal position measured in percent.
	 * @param {percentage} nShiftY - Vertical position measured in percent.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiPictureForm.prototype.SetPicturePosition = function(nShiftX, nShiftY)
	{
		let oPr = this.Sdt.GetPictureFormPr().Copy();
		oPr.SetShiftX(Math.max(0, Math.min(100, GetNumberParameter(nShiftX, 50))) / 100);
		oPr.SetShiftY(Math.max(0, Math.min(100, GetNumberParameter(nShiftY, 50))) / 100);
		this.Sdt.SetPictureFormPr(oPr);
		this.Sdt.UpdatePictureFormLayout();
		return true;
	};
	/**
	 * Returns the picture position inside the current form.
	 * @memberof ApiPictureForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {Array.<percentage>} Array of two numbers [shiftX, shiftY]
	 */
	ApiPictureForm.prototype.GetPicturePosition = function()
	{
		let oPr = this.Sdt.GetPictureFormPr();
		return [(oPr.GetShiftX() * 100) | 0, (oPr.GetShiftY() * 100) | 0];
	};
	/**
	 * Respects the form border width when scaling the image.
	 * @memberof ApiPictureForm
	 * @param {boolean} [isRespect=true] - Specifies if the form border width will be respected (true) or not (false).
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiPictureForm.prototype.SetRespectBorders = function(isRespect)
	{
		let oPr = this.Sdt.GetPictureFormPr().Copy();
		oPr.SetRespectBorders(GetBoolParameter(isRespect, true));
		this.Sdt.SetPictureFormPr(oPr);
		this.Sdt.UpdatePictureFormLayout();
		return true;
	};
	/**
	 * Checks if the form border width is respected or not.
	 * @memberof ApiPictureForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiPictureForm.prototype.IsRespectBorders = function()
	{
		return this.Sdt.GetPictureFormPr().IsRespectBorders();
	};
	/**
	 * Returns an image in the base64 format from the current picture form.
	 * @memberof ApiPictureForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {base64img}
	 */
	ApiPictureForm.prototype.GetImage = function()
	{
		var oImg;
		var allDrawings = this.Sdt.GetAllDrawingObjects();
		for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++)
		{
			if (allDrawings[nDrawing].IsPicture())
			{
				oImg = allDrawings[nDrawing].GraphicObj;
				break;
			}
		}
		if (oImg)
			return oImg.getBase64Img();

		return "";
	};
	/**
	 * Sets an image to the current picture form.
	 * @memberof ApiPictureForm
	 * @param {string} sImageSrc - The image source where the image to be inserted should be taken from (currently, only internet URL or base64 encoded images are supported).
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiPictureForm.prototype.SetImage = function(sImageSrc)
	{
		if (typeof(sImageSrc) !== "string" || sImageSrc === "")
			return false;

		var oImg;
		var allDrawings = this.Sdt.GetAllDrawingObjects();
		for (var nDrawing = 0; nDrawing < allDrawings.length; nDrawing++)
		{
			if (allDrawings[nDrawing].IsPicture())
			{
				oImg = allDrawings[nDrawing].GraphicObj;
				break;
			}
		}

		if (oImg)
		{
			oImg.setBlipFill(AscFormat.CreateBlipFillRasterImageId(sImageSrc));
			this.OnChangeValue();
			return true;
		}

		return false;
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiComboBoxForm
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns the list values from the current combo box.
	 * @memberof ApiComboBoxForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {string[]}
	 */
	ApiComboBoxForm.prototype.GetListValues = function()
	{
		var aValues = [];

		var oSpecProps = this.Sdt.IsComboBox() ? this.Sdt.GetComboBoxPr() : this.Sdt.GetDropDownListPr();
		if (!oSpecProps)
			return [];

		for (var nItem = 0, nCount = oSpecProps.GetItemsCount(); nItem < nCount; nItem++)
			aValues.push(oSpecProps.GetItemDisplayText(nItem));

		return aValues;
	};
	/**
	 * Sets the list values to the current combo box.
	 * @memberof ApiComboBoxForm
	 * @param {string[]} aListString - The combo box list values.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiComboBoxForm.prototype.SetListValues = function(aListString)
	{
		if (!Array.isArray(aListString))
			return false;

		let isComboBox = this.Sdt.IsComboBox();
		let oSpecProps = isComboBox ? this.Sdt.GetComboBoxPr() : this.Sdt.GetDropDownListPr();
		if (!oSpecProps)
			return false;

		oSpecProps = oSpecProps.Copy();
		oSpecProps.Clear();
		for (let nValue = 0; nValue < aListString.length; nValue++)
		{
			if (typeof(aListString[nValue]) === "string" && aListString[nValue] !== "")
				oSpecProps.AddItem(aListString[nValue], aListString[nValue]);
		}

		if (isComboBox)
			this.Sdt.SetComboBoxPr(oSpecProps);
		else
			this.Sdt.SetDropDownListPr(oSpecProps);

		return true;
	};
	/**
	 * Selects the specified value from the combo box list values. 
	 * @memberof ApiComboBoxForm
	 * @param {string} sValue - The combo box list value that will be selected.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiComboBoxForm.prototype.SelectListValue = function(sValue)
	{
		if (typeof(sValue) !== "string" || sValue === "")
			return false;

		var oSpecProps = this.Sdt.IsComboBox() ? this.Sdt.GetComboBoxPr() : this.Sdt.GetDropDownListPr();
		if (!oSpecProps)
			return false;

		if (null == oSpecProps.GetTextByValue(sValue))
			return false;

		this.Sdt.SelectListItem(sValue);
		this.OnChangeValue();
		return true;
	};
	/**
	 * Sets the text to the current combo box.
	 * *Available only for editable combo box forms.*
	 * @memberof ApiComboBoxForm
	 * @param {string} sText - The combo box text.
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiComboBoxForm.prototype.SetText = function(sText)
	{
		if (typeof (sText) !== "string" || sText === "")
			return false;

		if (!this.Sdt.IsComboBox())
			return false;

		if (this.Sdt.IsPlaceHolder())
			this.Sdt.ReplacePlaceHolderWithContent();

		let oRun = this.Sdt.MakeSingleRunElement();
		oRun.ClearContent();
		oRun.AddText(sText);
		
		this.OnChangeValue();
		return true;
	};
	/**
	 * Checks if the combo box text can be edited. If it is not editable, then this form is a dropdown list.
	 * @memberof ApiComboBoxForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiComboBoxForm.prototype.IsEditable = function()
	{
		return (this.Sdt.IsComboBox());
	};

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiCheckBoxForm
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Checks the current checkbox.
	 * @memberof ApiCheckBoxForm
	 * @param {boolean} isChecked - Specifies if the current checkbox will be checked (true) or not (false).
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiCheckBoxForm.prototype.SetChecked = function(isChecked)
	{
		if (typeof(isChecked) !== "boolean")
			return false;

		this.Sdt.ToggleCheckBox(isChecked);
		this.OnChangeValue();
		return true;
	};
	/**
	 * Returns the state of the current checkbox (checked or not).
	 * @memberof ApiCheckBoxForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiCheckBoxForm.prototype.IsChecked = function()
	{
		return this.Sdt.IsCheckBoxChecked();
	};
	/**
	 * Checks if the current checkbox is a radio button. 
	 * @memberof ApiCheckBoxForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {boolean}
	 */
	ApiCheckBoxForm.prototype.IsRadioButton = function()
	{
		return this.Sdt.IsRadioButton();
	};
	/**
	 * Returns the radio group key if the current checkbox is a radio button.
	 * @memberof ApiCheckBoxForm
	 * @typeofeditors ["CDE", "CFE"]
	 * @returns {string}
	 */
	ApiCheckBoxForm.prototype.GetRadioGroup = function()
	{
		let sRadioGroup = this.Sdt.GetRadioButtonGroupKey();
		return (sRadioGroup ? sRadioGroup : "");
	};
	/**
	 * Sets the radio group key to the current radio button.
	 * @memberof ApiCheckBoxForm
	 * @param {string} sKey - Radio group key.
	 * @typeofeditors ["CDE", "CFE"]
	 */
	ApiCheckBoxForm.prototype.SetRadioGroup = function(sKey)
	{
		let oPr = this.Sdt.GetCheckBoxPr();
		if (!oPr)
			return;

		oPr = oPr.Copy();
		oPr.SetGroupKey(sKey);
		this.Sdt.SetCheckBoxPr(oPr);
	};


	/**
	 * Converts the ApiBlockLvlSdt object into the JSON object.
	 * @memberof ApiBlockLvlSdt
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bWriteNumberings - Specifies if the used numberings will be written to the JSON object or not.
	 * @param {boolean} bWriteStyles - Specifies if the used styles will be written to the JSON object or not.
	 * @return {JSON}
	 */
	ApiBlockLvlSdt.prototype.ToJSON = function(bWriteNumberings, bWriteStyles)
	{
		var oWriter = new AscJsonConverter.WriterToJSON();
		var oJSON = oWriter.SerBlockLvlSdt(this.Sdt);
		if (bWriteNumberings)
			oJSON["numbering"] = oWriter.jsonWordNumberings;
		if (bWriteStyles)
			oJSON["styles"] = oWriter.SerWordStylesForWrite();
		return JSON.stringify(oJSON);
	};

	/**
	 * Replaces each paragraph (or text in cell) in the select with the corresponding text from an array of strings.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @param {Array} arrString - An array of replacement strings.
	 * @param {string} [sParaTab=" "] - A character which is used to specify the tab in the source text.
	 * @param {string} [sParaNewLine=" "] - A character which is used to specify the line break character in the source text.
	 */
	Api.prototype.ReplaceTextSmart = function(arrString, sParaTab, sParaNewLine)
	{
		if (typeof(sParaTab) !== "string")
			sParaTab = String.fromCharCode(32);
		if (typeof(sParaNewLine) !== "string")
			sParaNewLine = String.fromCharCode(32);

		var allRunsInfo      = null;
		var textDelta        = null;
		var arrSelectedParas = null;

		function GetRunInfo(oRun)
		{
			var StartPos = 0;
			var EndPos   = 0;
			var Item;
			var ItemType;

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
					Item = oRun.Content[Pos];
					ItemType = Item.Type;

					switch ( ItemType )
					{
						case para_Numbering:
						case para_PresentationNumbering:
						case para_PageNum:
						case para_PageCount:
						case para_End:
						case para_Drawing:
						case para_FieldChar:
						case para_InstrText:
						case para_RevisionMove:
						case para_FootnoteReference:
						case para_FootnoteRef:
						case para_EndnoteReference:
						case para_EndnoteRef:
						{
							if (posToSplit.indexOf(Pos) === -1)
								posToSplit.push(Pos);
							
							break;
						}
						case para_NewLine:
							if (sParaNewLine === "")
							{
								if (posToSplit.indexOf(Pos) === -1)
									posToSplit.push(Pos);
								
								break;
							}
							break;
						case para_Tab:
							if (sParaTab === "")
							{
								if (posToSplit.indexOf(Pos) === -1)
									posToSplit.push(Pos);
								
								break;
							}
							break;
					}
				}
				
				for (var Index = 0; Index < posToSplit.length; Index++)
				{
					var noTextCount = 0;

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
						Item = oRun.Content[nPos];
						ItemType = Item.Type;

						switch ( ItemType )
						{
							case para_Numbering:
							case para_PresentationNumbering:
							case para_PageNum:
							case para_PageCount:
							case para_End:
							case para_FieldChar:
							case para_InstrText:
							case para_Drawing:
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
							case para_Tab:
							{
								if (sParaTab !== "")
								{
									oInfo.String += sParaTab;
									oInfo.StringCount++; 
									break;
								}
								else
								{
									noTextCount++;
									break;
								}
							}
							case para_NewLine:
							{
								if (sParaNewLine !== "")
								{
									oInfo.String += sParaNewLine;
									oInfo.StringCount++; 
									break;
								}
								else 
								{
									noTextCount++;
									break;
								}
							}
							case para_Space:
							{
								oInfo.String += " ";
								oInfo.StringCount++; 
								break;
							}
						}
					}
					
					if (oInfo.String === "")
						continue;
					
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
						
						if ((nPosToDel >= oInfo.StartPos + oInfo.StringCount && nCharsToDel !== 0) || (nCharsToDel === 0 && oChange.deleteCount !== 0)
							|| nPosToAdd > oInfo.StartPos + oInfo.StringCount)
							continue;

						for (var nChar = 0; nChar < nCharsToDel; nChar++)
						{
							if (!oInfo.Run.Content[nPosToDel])
								break;
								
							if (para_Text === oInfo.Run.Content[nPosToDel].Type || para_Space === oInfo.Run.Content[nPosToDel].Type || para_Tab === oInfo.Run.Content[nPosToDel].Type || para_NewLine === oInfo.Run.Content[nPosToDel].Type)
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
						
						for (nChar = 0; nChar < oChange.insert.length; nChar++)
						{
							var itemText = null;
							if (oChange.insert[nChar] === 160)
								oChange.insert[nChar] = 32;

							if (AscCommon.IsSpace(oChange.insert[nChar]))
								itemText = new AscWord.CRunSpace(oChange.insert[nChar]);
							else if (oChange.insert[nChar] === '\t')
								itemText = new AscWord.CRunTab();
							else
								itemText = new AscWord.CRunText(oChange.insert[nChar]);

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
		}

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
			var oContent          = ws.objectRender.controller != null ? ws.objectRender.controller.getTargetDocContent() : null;
			var isPasteLocked     = false;
			var isLockedRange     = oWorksheet.worksheet.isLockedRange(oRange.range.bbox);

			if (oContent) 
			{
				arrSelectedParas = [];
				oContent.GetCurrentParagraph(false, arrSelectedParas, {});
				ReplaceInParas(arrSelectedParas);
				if (arrSelectedParas[0] && arrSelectedParas[0].Parent)
					arrSelectedParas[0].Parent.RemoveSelection();
				Asc.editor.wb.recalculateDrawingObjects();
				return;
			}

			if (oWorksheet.worksheet.getSheetProtection()) {
				let aProtRanges = oWorksheet.worksheet.protectedRangesContainsRange(oRange.range.bbox) || [];
				if (aProtRanges.length == 0 && isLockedRange)
					isPasteLocked = true;
				else if (aProtRanges.length > 0) {
					for (let Index = 0; Index < aProtRanges.length; Index++) {
						if (aProtRanges[Index].asc_isPassword() && !aProtRanges[Index].isUserEnteredPassword())
							isPasteLocked = true;
					}
				}
			}
			
			if (isPasteLocked) {
				if (oWorksheet.worksheet.workbook.handlers)
					oWorksheet.worksheet.workbook.handlers.trigger("asc_onError", c_oAscError.ID.ChangeOnProtectedSheet, c_oAscError.Level.NoCritical);
				return false;
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
			if(arrSelectedParas.length <= 0 )
			{
				return;
			}
			ReplaceInParas(arrSelectedParas);
			
			if (arrSelectedParas[0] && arrSelectedParas[0].Parent)
				arrSelectedParas[0].Parent.RemoveSelection();
			else 
				oDocument.Document.RemoveSelection();

			// вставка оставшихся параграфов из arrString
			var oParaParent   = arrSelectedParas[0].Parent;
			var nIndexToPaste = arrSelectedParas[arrSelectedParas.length - 1].Index + 1;
			var isPres        = !arrSelectedParas[0].bFromDocument;
			if (!oParaParent)
				return;

			for (var nPara = arrSelectedParas.length; nPara < arrString.length; nPara++)
			{
				var oPara = new AscCommonWord.Paragraph(private_GetDrawingDocument(), oParaParent, isPres);
				var oRun = new ParaRun(oPara, false);
				oRun.AddText(arrString[nPara]);
				private_PushElementToParagraph(oPara, oRun);
				oParaParent.AddToContent(nIndexToPaste, oPara);

				nIndexToPaste++;
			}
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
	 * Converts a document to Markdown or HTML text.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {"markdown" | "html"} [sConvertType="markdown"] - Conversion type.
	 * @param {boolean} [bHtmlHeadings=false] - Defines if the HTML headings and IDs will be generated when the Markdown renderer of your target platform does not handle Markdown-style IDs.
	 * @param {boolean} [bBase64img=false] - Defines if the images will be created in the base64 format.
	 * @param {boolean} [bDemoteHeadings=false] - Defines if all heading levels in your document will be demoted to conform with the following standard: single H1 as title, H2 as top-level heading in the text body.
	 * @param {boolean} [bRenderHTMLTags=false] - Defines if HTML tags will be preserved in your Markdown. If you just want to use an occasional HTML tag, you can avoid using the opening angle bracket
	 * in the following way: \<tag&gt;text\</tag&gt;. By default, the opening angle brackets will be replaced with the special characters.
	 * @returns {string}
	 */
	Api.prototype.ConvertDocument = function(sConvertType, bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags) 
	{
		var oDocument = this.GetDocument();
		if (!oDocument.Document)
			return "Please, use this plugin with the Word document editor";

		if (sConvertType === "html")
			return oDocument.ToHtml(bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags);
		else
			return oDocument.ToMarkdown(bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags);
	};

	/**
	 * Creates the empty text properties.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @returns {ApiTextPr}
	 */
	Api.prototype.CreateTextPr = function () {
		return this.private_CreateTextPr(null, new AscCommonWord.CTextPr());
	};

	/**
	 * Creates a Text Art object with the parameters specified.
	 * @memberof Api
	 * @typeofeditors ["CDE"]
	 * @param {ApiTextPr} [oTextPr=Api.CreateTextPr()] - The text properties.
	 * @param {string} [sText="Your text here"] - The text for the Text Art object.
	 * @param {TextTransform} [sTransform="textNoShape"] - Text transform type.
	 * @param {ApiFill}   [oFill=Api.CreateNoFill()] - The color or pattern used to fill the Text Art object.
	 * @param {ApiStroke} [oStroke=Api.CreateStroke(0, Api.CreateNoFill())] - The stroke used to create the Text Art object shadow.
	 * @param {number} [nRotAngle=0] - Rotation angle.
	 * @param {EMU} [nWidth=1828800] - The Text Art width measured in English measure units.
	 * @param {EMU} [nHeight=1828800] - The Text Art heigth measured in English measure units.
	 * @returns {ApiDrawing}
	 */
	Api.prototype.CreateWordArt = function(oTextPr, sText, sTransform, oFill, oStroke, nRotAngle, nWidth, nHeight) {
		oTextPr   = oTextPr && oTextPr.TextPr ? oTextPr.TextPr : null;
		nRotAngle = typeof(nRotAngle) === "number" && nRotAngle > 0 ? nRotAngle : 0;
		nWidth    = typeof(nWidth) === "number" && nWidth > 0 ? nWidth : 1828800;
		nHeight   = typeof(nHeight) === "number" && nHeight > 0 ? nHeight : 1828800;
		oFill     = oFill && oFill.UniFill ? oFill.UniFill : this.CreateNoFill().UniFill;
		oStroke   = oStroke && oStroke.Ln ? oStroke.Ln : this.CreateStroke(0, this.CreateNoFill()).Ln;
		sTransform = typeof(sTransform) === "string" && sTransform !== "" ? sTransform : "textNoShape";

		var oDrawing = new ParaDrawing(private_EMU2MM(nWidth), private_EMU2MM(nHeight), null, private_GetDrawingDocument(), private_GetLogicDocument(), null);
		var oArt = this.private_createWordArt(oTextPr, sText, sTransform, oFill, oStroke, nRotAngle, nWidth, nHeight);
		oArt.setParent(oDrawing);
		oDrawing.Set_GraphicObject(oArt);

		return new ApiDrawing(oDrawing);
	};

	/**
	 * Returns the full name of the currently opened file.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CPE", "CSE"]
	 * @returns {string}
	 */
	Api.prototype.GetFullName = function () {
		return this.DocInfo.Title;
	};
	Object.defineProperty(Api.prototype, "FullName", {
		get: function () {
			return this.GetFullName();
		}
	});

	//------------------------------------------------------------------------------------------------------------------
	//
	// ApiComment
	//
	//------------------------------------------------------------------------------------------------------------------

	/**
	 * Returns a type of the ApiComment class.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {"comment"}
	 */
	ApiComment.prototype.GetClassType = function ()
	{
		return "comment";
	};
	
	/**
	 * Returns the current comment ID. If the comment doesn't have an ID, null is returned.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {?string}
	 */
	ApiComment.prototype.GetId = function ()
	{
		let durableId = this.Comment.GetDurableId();
		if (-1 === durableId || null === durableId)
			return null;
		
		return ("" + durableId);
	};
	
	/**
	 * Returns the comment text.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiComment.prototype.GetText = function () {
		return this.Comment.GetData().Get_Text();
	};

	/**
	 * Sets the comment text.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment text.
	 * @returns {ApiComment} - this
	 */
	ApiComment.prototype.SetText = function (sText) {
		this.Comment.GetData().Set_Text(sText);
		this.private_OnChange();
		return this;
	};

	/**
	 * Returns the comment author's name.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiComment.prototype.GetAuthorName = function () {
		return this.Comment.GetData().Get_Name();
	};

	/**
	 * Sets the comment author's name.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {string} sAuthorName - The comment author's name.
	 * @returns {ApiComment} - this
	 */
	ApiComment.prototype.SetAuthorName = function (sAuthorName) {
		this.Comment.GetData().Set_Name(sAuthorName);
		this.private_OnChange();
		return this;
	};

	/**
	 * Returns the user ID of the comment author.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiComment.prototype.GetUserId = function () {
		return this.Comment.GetData().m_sUserId;
	};

	/**
	 * Sets the user ID to the comment author.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {string} sUserId - The user ID of the comment author.
	 * @returns {ApiComment} - this
	 */
	ApiComment.prototype.SetUserId = function (sUserId) {
		this.Comment.GetData().m_sUserId = sUserId;
		this.private_OnChange();
		return this;
	};
	
	/**
	 * Checks if a comment is solved or not.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiComment.prototype.IsSolved = function () {
		return this.Comment.GetData().IsSolved();
	};

	/**
	 * Marks a comment as solved.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {boolean} bSolved - Specifies if a comment is solved or not.
	 * @returns {ApiComment} - this
	 */
	ApiComment.prototype.SetSolved = function (bSolved) {
		this.Comment.GetData().SetSolved(bSolved);
		this.private_OnChange();
		return this;
	};

	/**
	 * Returns the timestamp of the comment creation in UTC format.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {Number}
	 */
	ApiComment.prototype.GetTimeUTC = function () {
		let nTime = parseInt(this.Comment.GetData().m_sOOTime);
		if (isNaN(nTime))
			return 0;
		return nTime;
	};

	/**
	 * Sets the timestamp of the comment creation in UTC format.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {Number | String} nTimeStamp - The timestamp of the comment creation in UTC format.
	 * @returns {ApiComment} - this
	 */
	ApiComment.prototype.SetTimeUTC = function (timeStamp) {
		let nTime = parseInt(timeStamp);
		if (isNaN(nTime))
			this.Comment.GetData().m_sOOTime = "0";
		else
			this.Comment.GetData().m_sOOTime = String(nTime);
		
		this.private_OnChange();
		return this;
	};

	/**
	 * Returns the timestamp of the comment creation in the current time zone format.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {Number}
	 */
	 ApiComment.prototype.GetTime = function () {
		return this.Comment.GetData().GetDateTime();
	};

	/**
	 * Sets the timestamp of the comment creation in the current time zone format.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {Number | String} nTimeStamp - The timestamp of the comment creation in the current time zone format.
	 * @returns {ApiComment} - this
	 */
	ApiComment.prototype.SetTime = function (timeStamp) {
		let nTime = parseInt(timeStamp);
		if (isNaN(nTime))
			this.Comment.GetData().m_sTime = "0";
		else
			this.Comment.GetData().m_sTime = String(nTime);
		
		this.private_OnChange();
		return this;
	};

	/**
	 * Returns the quote text of the current comment.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {Number?}
	 */
	ApiComment.prototype.GetQuoteText = function () {
		return this.Comment.GetData().GetQuoteText();
	};

	/**
	 * Returns a number of the comment replies.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {Number?}
	 */
	ApiComment.prototype.GetRepliesCount = function () {
		return this.Comment.GetData().Get_RepliesCount();
	};

	/**
	 * Returns the specified comment reply.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {Number} [nIndex = 0] - The comment reply index.
	 * @returns {ApiCommentReply?}
	 */
	ApiComment.prototype.GetReply = function (nIndex) {
		if (typeof(nIndex) != "number" || nIndex < 0 || nIndex >= this.GetRepliesCount())
			nIndex = 0;
			
		let oReply = this.Comment.GetData().GetReply(nIndex);
		if (!oReply)
			return null;

		return new ApiCommentReply(this.Comment, oReply);
	};

	/**
	 * Adds a reply to a comment.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {String} sText - The comment reply text (required).
	 * @param {String} sAuthorName - The name of the comment reply author (optional).
	 * @param {String} sUserId - The user ID of the comment reply author (optional).
	 * @param {Number} [nPos=this.GetRepliesCount()] - The comment reply position.
	 * @returns {ApiComment?} - this
	 */
	ApiComment.prototype.AddReply = function (sText, sAuthorName, sUserId, nPos) {
		if (typeof(sText) !== "string" || sText === "")
			return null;
		
		if (typeof(nPos) !== "number" || nPos < 0 || nPos > this.GetRepliesCount())
			nPos = this.GetRepliesCount();

		var oReply = new AscCommon.CCommentData();

		oReply.SetText(sText);
		if (typeof(sAuthorName) === "string" && sAuthorName !== "")
			oReply.SetUserName(sAuthorName);
		if (sUserId != undefined && typeof(sUserId) === "string" && sUserId !== "")
			oReply.m_sUserId = sUserId;

		this.Comment.Data.m_aReplies.splice(nPos, 0, oReply);
		this.private_OnChange();
		return this;
	};

	/**
	 * Removes the specified comment replies.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @param {Number} [nPos = 0] - The position of the first comment reply to remove.
	 * @param {Number} [nCount = 1] - A number of comment replies to remove.
	 * @param {boolean} [bRemoveAll = false] - Specifies whether to remove all comment replies or not.
	 * @returns {ApiComment?} - this
	 */
	ApiComment.prototype.RemoveReplies = function (nPos, nCount, bRemoveAll) {
		if (typeof(nPos) !== "number" || nPos < 0 || nPos > this.GetRepliesCount())
			nPos = 0;
		if (typeof(nCount) !== "number" || nCount < 0)
			nCount = 1;
		if (typeof(bRemoveAll) !== "boolean")
			bRemoveAll = false;

		if (bRemoveAll)
			nCount = this.GetRepliesCount();

		this.Comment.Data.m_aReplies.splice(nPos, nCount);
		this.private_OnChange();
		return this;
	};

	/**
	 * Deletes the current comment from the document.
	 * @memberof ApiComment
	 * @typeofeditors ["CDE"]
	 * @returns {boolean}
	 */
	ApiComment.prototype.Delete = function ()
	{
		let logicDocument = private_GetLogicDocument();
		if (!logicDocument)
			return false;
		
		return logicDocument.RemoveComment(this.Comment.GetId(), true);
	};

	/**
	 * Returns a type of the ApiCommentReply class.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CDE"]
	 * @returns {"commentReply"}
	 */
	ApiCommentReply.prototype.GetClassType = function () {
		return "commentReply";
	};

	/**
	 * Returns the comment reply text.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiCommentReply.prototype.GetText = function () {
		return this.Data.Get_Text();
	};

	/**
	 * Sets the comment reply text.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CDE"]
	 * @param {string} sText - The comment reply text.
	 * @returns {ApiCommentReply} - this
	 */
	ApiCommentReply.prototype.SetText = function (sText) {
		this.Data.Set_Text(sText);
		this.private_OnChange();
		return this;
	};
	
	/**
	 * Returns the comment reply author's name.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiCommentReply.prototype.GetAuthorName = function () {
		return this.Data.Get_Name();
	};

	/**
	 * Sets the comment reply author's name.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CDE"]
	 * @param {string} sAuthorName - The comment reply author's name.
	 * @returns {ApiCommentReply} - this
	 */
	ApiCommentReply.prototype.SetAuthorName = function (sAuthorName) {
		this.Data.Set_Name(sAuthorName);
		this.private_OnChange();
		return this;
	};

	/**
	 * Returns the user ID of the comment reply author.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CDE"]
	 * @returns {string}
	 */
	ApiCommentReply.prototype.GetUserId = function () {
		return this.Data.m_sUserId;
	};

	/**
	 * Sets the user ID to the comment reply author.
	 * @memberof ApiCommentReply
	 * @typeofeditors ["CDE"]
	 * @param {string} sUserId - The user ID of the comment reply author.
	 * @returns {ApiCommentReply} - this
	 */
	ApiCommentReply.prototype.SetUserId = function (sUserId) {
		this.Data.m_sUserId = sUserId;
		this.private_OnChange();
		return this;
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
	Api.prototype["CreateTextPr"]		             = Api.prototype.CreateTextPr;
	Api.prototype["CreateWordArt"]		             = Api.prototype.CreateWordArt;
	Api.prototype["CreateOleObject"]		         = Api.prototype.CreateOleObject;
	Api.prototype["GetFullName"]		             = Api.prototype.GetFullName;

	Api.prototype["ConvertDocument"]		         = Api.prototype.ConvertDocument;
	Api.prototype["FromJSON"]		                 = Api.prototype.FromJSON;
	Api.prototype["CreateRange"]		             = Api.prototype.CreateRange;

	ApiUnsupported.prototype["GetClassType"]         = ApiUnsupported.prototype.GetClassType;

	ApiDocumentContent.prototype["GetClassType"]         = ApiDocumentContent.prototype.GetClassType;
	ApiDocumentContent.prototype["GetElementsCount"]     = ApiDocumentContent.prototype.GetElementsCount;
	ApiDocumentContent.prototype["GetElement"]           = ApiDocumentContent.prototype.GetElement;
	ApiDocumentContent.prototype["AddElement"]           = ApiDocumentContent.prototype.AddElement;
	ApiDocumentContent.prototype["Push"]                 = ApiDocumentContent.prototype.Push;
	ApiDocumentContent.prototype["RemoveAllElements"]    = ApiDocumentContent.prototype.RemoveAllElements;
	ApiDocumentContent.prototype["RemoveElement"]        = ApiDocumentContent.prototype.RemoveElement;
	ApiDocumentContent.prototype["GetRange"]             = ApiDocumentContent.prototype.GetRange;
	ApiDocumentContent.prototype["ToJSON"]               = ApiDocumentContent.prototype.ToJSON;
	ApiDocumentContent.prototype["GetContent"]           = ApiDocumentContent.prototype.GetContent;
	ApiDocumentContent.prototype["GetAllDrawingObjects"] = ApiDocumentContent.prototype.GetAllDrawingObjects;
	ApiDocumentContent.prototype["GetAllShapes"]         = ApiDocumentContent.prototype.GetAllShapes;
	ApiDocumentContent.prototype["GetAllImages"]         = ApiDocumentContent.prototype.GetAllImages;
	ApiDocumentContent.prototype["GetAllCharts"]         = ApiDocumentContent.prototype.GetAllCharts;
	ApiDocumentContent.prototype["GetAllOleObjects"]     = ApiDocumentContent.prototype.GetAllOleObjects;
	ApiDocumentContent.prototype["GetAllParagraphs"]     = ApiDocumentContent.prototype.GetAllParagraphs;
	ApiDocumentContent.prototype["GetAllTables"]         = ApiDocumentContent.prototype.GetAllTables;

	ApiRange.prototype["GetClassType"]               = ApiRange.prototype.GetClassType;
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
	ApiRange.prototype["ToJSON"]                     = ApiRange.prototype.ToJSON;
	ApiRange.prototype["AddComment"]                 = ApiRange.prototype.AddComment;
	ApiRange.prototype["GetRange"]                   = ApiRange.prototype.GetRange;

	ApiDocument.prototype["GetClassType"]                = ApiDocument.prototype.GetClassType;
	ApiDocument.prototype["CreateNewHistoryPoint"]       = ApiDocument.prototype.CreateNewHistoryPoint;
	ApiDocument.prototype["GetDefaultTextPr"]            = ApiDocument.prototype.GetDefaultTextPr;
	ApiDocument.prototype["GetDefaultParaPr"]            = ApiDocument.prototype.GetDefaultParaPr;
	ApiDocument.prototype["GetStyle"]                    = ApiDocument.prototype.GetStyle;
	ApiDocument.prototype["CreateStyle"]                 = ApiDocument.prototype.CreateStyle;
	ApiDocument.prototype["GetDefaultStyle"]             = ApiDocument.prototype.GetDefaultStyle;
	ApiDocument.prototype["GetFinalSection"]             = ApiDocument.prototype.GetFinalSection;
	ApiDocument.prototype["CreateSection"]               = ApiDocument.prototype.CreateSection;
	ApiDocument.prototype["SetEvenAndOddHdrFtr"]         = ApiDocument.prototype.SetEvenAndOddHdrFtr;
	ApiDocument.prototype["CreateNumbering"]             = ApiDocument.prototype.CreateNumbering;
	ApiDocument.prototype["InsertContent"]               = ApiDocument.prototype.InsertContent;
	ApiDocument.prototype["GetCommentsReport"]           = ApiDocument.prototype.GetCommentsReport;
	ApiDocument.prototype["GetReviewReport"]             = ApiDocument.prototype.GetReviewReport;
	ApiDocument.prototype["InsertWatermark"]             = ApiDocument.prototype.InsertWatermark;
	ApiDocument.prototype["SearchAndReplace"]            = ApiDocument.prototype.SearchAndReplace;
	ApiDocument.prototype["GetAllContentControls"]       = ApiDocument.prototype.GetAllContentControls;
	ApiDocument.prototype["GetTagsOfAllContentControls"] = ApiDocument.prototype.GetTagsOfAllContentControls;
	ApiDocument.prototype["GetTagsOfAllForms"]           = ApiDocument.prototype.GetTagsOfAllForms;
	ApiDocument.prototype["GetContentControlsByTag"]     = ApiDocument.prototype.GetContentControlsByTag;
	ApiDocument.prototype["GetFormsByTag"]               = ApiDocument.prototype.GetFormsByTag;
	ApiDocument.prototype["SetTrackRevisions"]           = ApiDocument.prototype.SetTrackRevisions;
	ApiDocument.prototype["IsTrackRevisions"]            = ApiDocument.prototype.IsTrackRevisions;
	ApiDocument.prototype["GetRange"]                    = ApiDocument.prototype.GetRange;
	ApiDocument.prototype["GetRangeBySelect"]            = ApiDocument.prototype.GetRangeBySelect;
	ApiDocument.prototype["Last"]                        = ApiDocument.prototype.Last;
	ApiDocument.prototype["Push"]                        = ApiDocument.prototype.Push;
	ApiDocument.prototype["DeleteBookmark"]              = ApiDocument.prototype.DeleteBookmark;
	ApiDocument.prototype["AddComment"]                  = ApiDocument.prototype.AddComment;
	ApiDocument.prototype["GetBookmarkRange"]            = ApiDocument.prototype.GetBookmarkRange;
	ApiDocument.prototype["GetSections"]                 = ApiDocument.prototype.GetSections;
	ApiDocument.prototype["GetAllTablesOnPage"]          = ApiDocument.prototype.GetAllTablesOnPage;
	ApiDocument.prototype["RemoveSelection"]             = ApiDocument.prototype.RemoveSelection;
	ApiDocument.prototype["Search"]                      = ApiDocument.prototype.Search;
	ApiDocument.prototype["ToMarkdown"]                  = ApiDocument.prototype.ToMarkdown;
	ApiDocument.prototype["ToHtml"]                      = ApiDocument.prototype.ToHtml;
	ApiDocument.prototype["GetAllNumberedParagraphs"]    = ApiDocument.prototype.GetAllNumberedParagraphs;
	ApiDocument.prototype["GetAllHeadingParagraphs"]     = ApiDocument.prototype.GetAllHeadingParagraphs;
	ApiDocument.prototype["GetFootnotesFirstParagraphs"] = ApiDocument.prototype.GetFootnotesFirstParagraphs;
	ApiDocument.prototype["GetEndNotesFirstParagraphs"]  = ApiDocument.prototype.GetEndNotesFirstParagraphs;
	ApiDocument.prototype["GetAllCaptionParagraphs"]     = ApiDocument.prototype.GetAllCaptionParagraphs;
	ApiDocument.prototype["GetAllBookmarksNames"]        = ApiDocument.prototype.GetAllBookmarksNames;
	ApiDocument.prototype["AddFootnote"]                 = ApiDocument.prototype.AddFootnote;
	ApiDocument.prototype["AddEndnote"]                  = ApiDocument.prototype.AddEndnote;
	ApiDocument.prototype["SetControlsHighlight"]        = ApiDocument.prototype.SetControlsHighlight;
	ApiDocument.prototype["GetAllComments"]              = ApiDocument.prototype.GetAllComments;
	ApiDocument.prototype["GetCommentById"]              = ApiDocument.prototype.GetCommentById;
	ApiDocument.prototype["GetStatistics"]               = ApiDocument.prototype.GetStatistics;
	ApiDocument.prototype["GetPageCount"]                = ApiDocument.prototype.GetPageCount;
	ApiDocument.prototype["GetAllStyles"]                = ApiDocument.prototype.GetAllStyles;
	
	ApiDocument.prototype["GetSelectedDrawings"]         = ApiDocument.prototype.GetSelectedDrawings;
	ApiDocument.prototype["ReplaceCurrentImage"]         = ApiDocument.prototype.ReplaceCurrentImage;
	ApiDocument.prototype["ReplaceDrawing"]              = ApiDocument.prototype.ReplaceDrawing;
	ApiDocument.prototype["AcceptAllRevisionChanges"]    = ApiDocument.prototype.AcceptAllRevisionChanges;
	ApiDocument.prototype["RejectAllRevisionChanges"]    = ApiDocument.prototype.RejectAllRevisionChanges;

	ApiDocument.prototype["ToJSON"]                  = ApiDocument.prototype.ToJSON;
	ApiDocument.prototype["UpdateAllTOC"]		     = ApiDocument.prototype.UpdateAllTOC;
	ApiDocument.prototype["UpdateAllTOF"]		     = ApiDocument.prototype.UpdateAllTOF;
	ApiDocument.prototype["AddTableOfContents"]		 = ApiDocument.prototype.AddTableOfContents;
	ApiDocument.prototype["AddTableOfFigures"]		 = ApiDocument.prototype.AddTableOfFigures;

	ApiDocument.prototype["GetAllForms"]             = ApiDocument.prototype.GetAllForms;
	ApiDocument.prototype["ClearAllFields"]          = ApiDocument.prototype.ClearAllFields;
	ApiDocument.prototype["SetFormsHighlight"]       = ApiDocument.prototype.SetFormsHighlight;

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
	ApiParagraph.prototype["GetFontNames"]           = ApiParagraph.prototype.GetFontNames;
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
	ApiParagraph.prototype["GetAllOleObjects"]       = ApiParagraph.prototype.GetAllOleObjects;
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
	ApiParagraph.prototype["AddNumberedCrossRef"]    = ApiParagraph.prototype.AddNumberedCrossRef;
	ApiParagraph.prototype["AddHeadingCrossRef"]     = ApiParagraph.prototype.AddHeadingCrossRef;
	ApiParagraph.prototype["AddBookmarkCrossRef"]    = ApiParagraph.prototype.AddBookmarkCrossRef;
	ApiParagraph.prototype["AddFootnoteCrossRef"]    = ApiParagraph.prototype.AddFootnoteCrossRef;
	ApiParagraph.prototype["AddEndnoteCrossRef"]     = ApiParagraph.prototype.AddEndnoteCrossRef;
	ApiParagraph.prototype["AddCaptionCrossRef"]     = ApiParagraph.prototype.AddCaptionCrossRef;
	ApiParagraph.prototype["GetPosInParent"]         = ApiParagraph.prototype.GetPosInParent;
	ApiParagraph.prototype["ReplaceByElement"]       = ApiParagraph.prototype.ReplaceByElement;
	ApiParagraph.prototype["AddCaption"]             = ApiParagraph.prototype.AddCaption;
	ApiParagraph.prototype["SetSection"]             = ApiParagraph.prototype.SetSection;
	ApiParagraph.prototype["GetSection"]             = ApiParagraph.prototype.GetSection;

	ApiParagraph.prototype["ToJSON"]                 = ApiParagraph.prototype.ToJSON;


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
	ApiRun.prototype["GetFontNames"]                 = ApiRun.prototype.GetFontNames;
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
	ApiRun.prototype["ToJSON"]                       = ApiRun.prototype.ToJSON;
	ApiRun.prototype["AddComment"]                   = ApiRun.prototype.AddComment;
	ApiRun.prototype["GetText"]                      = ApiRun.prototype.GetText;


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
	ApiHyperlink.prototype["ToJSON"]                 = ApiHyperlink.prototype.ToJSON;

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
	ApiSection.prototype["ToJSON"]                   = ApiSection.prototype.ToJSON;

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
	ApiTable.prototype["SetBackgroundColor"]    	 = ApiTable.prototype.SetBackgroundColor;
	ApiTable.prototype["ToJSON"]    				 = ApiTable.prototype.ToJSON;
	ApiTable.prototype["GetPosInParent"]    	     = ApiTable.prototype.GetPosInParent;
	ApiTable.prototype["ReplaceByElement"]    		 = ApiTable.prototype.ReplaceByElement;
	ApiTable.prototype["AddComment"]  		  		 = ApiTable.prototype.AddComment;
	ApiTable.prototype["AddCaption"]  		  		 = ApiTable.prototype.AddCaption;

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
	ApiTableRow.prototype["SetBackgroundColor"]      = ApiTableRow.prototype.SetBackgroundColor;

	ApiTableCell.prototype["GetClassType"]             = ApiTableCell.prototype.GetClassType;
	ApiTableCell.prototype["GetContent"]               = ApiTableCell.prototype.GetContent;
	ApiTableCell.prototype["GetIndex"]    			   = ApiTableCell.prototype.GetIndex;
	ApiTableCell.prototype["GetRowIndex"]    		   = ApiTableCell.prototype.GetRowIndex;
	ApiTableCell.prototype["GetParentRow"]    		   = ApiTableCell.prototype.GetParentRow;
	ApiTableCell.prototype["GetParentTable"]    	   = ApiTableCell.prototype.GetParentTable;
	ApiTableCell.prototype["AddRows"]    			   = ApiTableCell.prototype.AddRows;
	ApiTableCell.prototype["AddColumns"]    		   = ApiTableCell.prototype.AddColumns;
	ApiTableCell.prototype["RemoveColumn"]    		   = ApiTableCell.prototype.RemoveColumn;
	ApiTableCell.prototype["RemoveRow"]    			   = ApiTableCell.prototype.RemoveRow;
	ApiTableCell.prototype["Search"]    			   = ApiTableCell.prototype.Search;
	ApiTableCell.prototype["GetNext"]    			   = ApiTableCell.prototype.GetNext;
	ApiTableCell.prototype["GetPrevious"]    		   = ApiTableCell.prototype.GetPrevious;
	ApiTableCell.prototype["Split"]    				   = ApiTableCell.prototype.Split;
	ApiTableCell.prototype["SetCellPr"]    			   = ApiTableCell.prototype.SetCellPr;
	ApiTableCell.prototype["SetTextPr"]    			   = ApiTableCell.prototype.SetTextPr;
	ApiTableCell.prototype["Clear"]    		           = ApiTableCell.prototype.Clear;
	ApiTableCell.prototype["AddElement"]    		   = ApiTableCell.prototype.AddElement;
	ApiTableCell.prototype["SetBackgroundColor"]       = ApiTableCell.prototype.SetBackgroundColor;
	ApiTableCell.prototype["SetColumnBackgroundColor"] = ApiTableCell.prototype.SetColumnBackgroundColor;

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
	ApiStyle.prototype["ToJSON"]                     = ApiStyle.prototype.ToJSON;

	ApiNumbering.prototype["GetClassType"]           = ApiNumbering.prototype.GetClassType;
	ApiNumbering.prototype["GetLevel"]               = ApiNumbering.prototype.GetLevel;
	ApiNumbering.prototype["ToJSON"]                 = ApiNumbering.prototype.ToJSON;

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
	ApiTextPr.prototype["SetTextFill"]               = ApiTextPr.prototype.SetTextFill;
	ApiTextPr.prototype["SetOutLine"]                = ApiTextPr.prototype.SetOutLine;
	ApiTextPr.prototype["ToJSON"]                    = ApiTextPr.prototype.ToJSON;

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
	ApiParaPr.prototype["ToJSON"]                    = ApiParaPr.prototype.ToJSON;

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
	ApiTablePr.prototype["SetTableTitle"]            = ApiTablePr.prototype.SetTableTitle;
	ApiTablePr.prototype["GetTableTitle"]            = ApiTablePr.prototype.GetTableTitle;
	ApiTablePr.prototype["SetTableDescription"]      = ApiTablePr.prototype.SetTableDescription;
	ApiTablePr.prototype["GetTableDescription"]      = ApiTablePr.prototype.GetTableDescription;

	ApiTablePr.prototype["ToJSON"]                   = ApiTablePr.prototype.ToJSON;

	ApiTableRowPr.prototype["GetClassType"]          = ApiTableRowPr.prototype.GetClassType;
	ApiTableRowPr.prototype["SetHeight"]             = ApiTableRowPr.prototype.SetHeight;
	ApiTableRowPr.prototype["SetTableHeader"]        = ApiTableRowPr.prototype.SetTableHeader;
	ApiTableRowPr.prototype["ToJSON"]                = ApiTableRowPr.prototype.ToJSON;

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
	ApiTableCellPr.prototype["ToJSON"]               = ApiTableCellPr.prototype.ToJSON;

	ApiTableStylePr.prototype["GetClassType"]        = ApiTableStylePr.prototype.GetClassType;
	ApiTableStylePr.prototype["GetType"]             = ApiTableStylePr.prototype.GetType;
	ApiTableStylePr.prototype["GetTextPr"]           = ApiTableStylePr.prototype.GetTextPr;
	ApiTableStylePr.prototype["GetParaPr"]           = ApiTableStylePr.prototype.GetParaPr;
	ApiTableStylePr.prototype["GetTablePr"]          = ApiTableStylePr.prototype.GetTablePr;
	ApiTableStylePr.prototype["GetTableRowPr"]       = ApiTableStylePr.prototype.GetTableRowPr;
	ApiTableStylePr.prototype["GetTableCellPr"]      = ApiTableStylePr.prototype.GetTableCellPr;
	ApiTableStylePr.prototype["ToJSON"]              = ApiTableStylePr.prototype.ToJSON;

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
	ApiDrawing.prototype["GetWidth"]                 = ApiDrawing.prototype.GetWidth;
	ApiDrawing.prototype["GetHeight"]                = ApiDrawing.prototype.GetHeight;
	ApiDrawing.prototype["GetLockValue"]             = ApiDrawing.prototype.GetLockValue;
	ApiDrawing.prototype["SetLockValue"]             = ApiDrawing.prototype.SetLockValue;
	ApiDrawing.prototype["SetDrawingPrFromDrawing"]  = ApiDrawing.prototype.SetDrawingPrFromDrawing;

	ApiDrawing.prototype["ToJSON"]                   = ApiDrawing.prototype.ToJSON;

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
	ApiChart.prototype["GetNextChart"]                 =  ApiChart.prototype.GetNextChart;
    ApiChart.prototype["GetPrevChart"]                 =  ApiChart.prototype.GetPrevChart;
    ApiChart.prototype["RemoveSeria"]                  =  ApiChart.prototype.RemoveSeria;
    ApiChart.prototype["SetSeriaValues"]               =  ApiChart.prototype.SetSeriaValues;
    ApiChart.prototype["SetXValues"]                   =  ApiChart.prototype.SetXValues;
    ApiChart.prototype["SetSeriaName"]                 =  ApiChart.prototype.SetSeriaName;
    ApiChart.prototype["SetCategoryName"]              =  ApiChart.prototype.SetCategoryName;
	ApiChart.prototype["ApplyChartStyle"]              =  ApiChart.prototype.ApplyChartStyle;
	ApiChart.prototype["SetPlotAreaFill"]              =  ApiChart.prototype.SetPlotAreaFill;
	ApiChart.prototype["SetPlotAreaOutLine"]           =  ApiChart.prototype.SetPlotAreaOutLine;
	ApiChart.prototype["SetSeriesFill"]                =  ApiChart.prototype.SetSeriesFill;
	ApiChart.prototype["SetSeriesOutLine"]             =  ApiChart.prototype.SetSeriesOutLine;
	ApiChart.prototype["SetDataPointFill"]             =  ApiChart.prototype.SetDataPointFill;
	ApiChart.prototype["SetDataPointOutLine"]          =  ApiChart.prototype.SetDataPointOutLine;
	ApiChart.prototype["SetMarkerFill"]                =  ApiChart.prototype.SetMarkerFill;
	ApiChart.prototype["SetMarkerOutLine"]             =  ApiChart.prototype.SetMarkerOutLine;
	ApiChart.prototype["SetTitleFill"]                 =  ApiChart.prototype.SetTitleFill;
	ApiChart.prototype["SetTitleOutLine"]              =  ApiChart.prototype.SetTitleOutLine;
	ApiChart.prototype["SetLegendFill"]                =  ApiChart.prototype.SetLegendFill;
	ApiChart.prototype["SetLegendOutLine"]             =  ApiChart.prototype.SetLegendOutLine;
	ApiChart.prototype["SetAxieNumFormat"]             =  ApiChart.prototype.SetAxieNumFormat;
	ApiChart.prototype["SetSeriaNumFormat"]            =  ApiChart.prototype.SetSeriaNumFormat;
	ApiChart.prototype["SetDataPointNumFormat"]        =  ApiChart.prototype.SetDataPointNumFormat;

	ApiOleObject.prototype["GetClassType"]             = ApiOleObject.prototype.GetClassType;
	ApiOleObject.prototype["SetData"]               = ApiOleObject.prototype.SetData;
	ApiOleObject.prototype["GetData"]               = ApiOleObject.prototype.GetData;
	ApiOleObject.prototype["SetApplicationId"]         = ApiOleObject.prototype.SetApplicationId;
	ApiOleObject.prototype["GetApplicationId"]         = ApiOleObject.prototype.GetApplicationId;

	ApiFill.prototype["GetClassType"]                = ApiFill.prototype.GetClassType;
	ApiFill.prototype["ToJSON"]                      = ApiFill.prototype.ToJSON;

	ApiStroke.prototype["GetClassType"]              = ApiStroke.prototype.GetClassType;
	ApiStroke.prototype["ToJSON"]                    = ApiStroke.prototype.ToJSON;

	ApiGradientStop.prototype["GetClassType"]        = ApiGradientStop.prototype.GetClassType;
	ApiGradientStop.prototype["ToJSON"]              = ApiGradientStop.prototype.ToJSON;

	ApiUniColor.prototype["GetClassType"]            = ApiUniColor.prototype.GetClassType;
	ApiUniColor.prototype["ToJSON"]                  = ApiUniColor.prototype.ToJSON;

	ApiRGBColor.prototype["GetClassType"]            = ApiRGBColor.prototype.GetClassType;
	ApiRGBColor.prototype["ToJSON"]                  = ApiRGBColor.prototype.ToJSON;

	ApiSchemeColor.prototype["GetClassType"]         = ApiSchemeColor.prototype.GetClassType;
	ApiSchemeColor.prototype["ToJSON"]               = ApiSchemeColor.prototype.ToJSON;

	ApiPresetColor.prototype["GetClassType"]         = ApiPresetColor.prototype.GetClassType;
	ApiPresetColor.prototype["ToJSON"]               = ApiPresetColor.prototype.ToJSON;

	ApiBullet.prototype["GetClassType"]              = ApiBullet.prototype.GetClassType;
	ApiBullet.prototype["ToJSON"]                    = ApiBullet.prototype.ToJSON;

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
	ApiInlineLvlSdt.prototype["GetParentParagraph"]     = ApiInlineLvlSdt.prototype.GetParentParagraph;
	ApiInlineLvlSdt.prototype["GetParentContentControl"]= ApiInlineLvlSdt.prototype.GetParentContentControl;
	ApiInlineLvlSdt.prototype["GetParentTable"]         = ApiInlineLvlSdt.prototype.GetParentTable;
	ApiInlineLvlSdt.prototype["GetParentTableCell"]     = ApiInlineLvlSdt.prototype.GetParentTableCell;
	ApiInlineLvlSdt.prototype["GetRange"]               = ApiInlineLvlSdt.prototype.GetRange;
	ApiInlineLvlSdt.prototype["Copy"]                   = ApiInlineLvlSdt.prototype.Copy;
	ApiInlineLvlSdt.prototype["ToJSON"]                 = ApiInlineLvlSdt.prototype.ToJSON;
	ApiInlineLvlSdt.prototype["AddComment"]             = ApiInlineLvlSdt.prototype.AddComment;

	ApiInlineLvlSdt.prototype["GetPlaceholderText"]     = ApiInlineLvlSdt.prototype.GetPlaceholderText;
	ApiInlineLvlSdt.prototype["SetPlaceholderText"]     = ApiInlineLvlSdt.prototype.SetPlaceholderText;
	ApiInlineLvlSdt.prototype["IsForm"]                 = ApiInlineLvlSdt.prototype.IsForm;
	ApiInlineLvlSdt.prototype["GetForm"]                = ApiInlineLvlSdt.prototype.GetForm;
	ApiInlineLvlSdt.prototype["GetDropdownList"]        = ApiInlineLvlSdt.prototype.GetDropdownList;

	ApiContentControlList.prototype["GetClassType"]		= ApiContentControlList.prototype.GetClassType;
	ApiContentControlList.prototype["GetAllItems"]		= ApiContentControlList.prototype.GetAllItems;
	ApiContentControlList.prototype["GetElementsCount"]	= ApiContentControlList.prototype.GetElementsCount;
	ApiContentControlList.prototype["GetParent"]		= ApiContentControlList.prototype.GetParent;
	ApiContentControlList.prototype["Add"]				= ApiContentControlList.prototype.Add;
	ApiContentControlList.prototype["Clear"]			= ApiContentControlList.prototype.Clear;
	ApiContentControlList.prototype["GetItem"]			= ApiContentControlList.prototype.GetItem;

	ApiContentControlListEntry.prototype["GetClassType"]	= ApiContentControlListEntry.prototype.GetClassType;
	ApiContentControlListEntry.prototype["GetParent"]		= ApiContentControlListEntry.prototype.GetParent;
	ApiContentControlListEntry.prototype["Select"]			= ApiContentControlListEntry.prototype.Select;
	ApiContentControlListEntry.prototype["MoveUp"]			= ApiContentControlListEntry.prototype.MoveUp;
	ApiContentControlListEntry.prototype["MoveDown"]		= ApiContentControlListEntry.prototype.MoveDown;
	ApiContentControlListEntry.prototype["GetIndex"]		= ApiContentControlListEntry.prototype.GetIndex;
	ApiContentControlListEntry.prototype["SetIndex"]		= ApiContentControlListEntry.prototype.SetIndex;
	ApiContentControlListEntry.prototype["Delete"]			= ApiContentControlListEntry.prototype.Delete;
	ApiContentControlListEntry.prototype["GetText"]			= ApiContentControlListEntry.prototype.GetText;
	ApiContentControlListEntry.prototype["SetText"]			= ApiContentControlListEntry.prototype.SetText;
	ApiContentControlListEntry.prototype["GetValue"]		= ApiContentControlListEntry.prototype.GetValue;
	ApiContentControlListEntry.prototype["SetValue"]		= ApiContentControlListEntry.prototype.SetValue;

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
	ApiBlockLvlSdt.prototype["ToJSON"]                  = ApiBlockLvlSdt.prototype.ToJSON;
	ApiBlockLvlSdt.prototype["GetPlaceholderText"]      = ApiBlockLvlSdt.prototype.GetPlaceholderText;
	ApiBlockLvlSdt.prototype["SetPlaceholderText"]      = ApiBlockLvlSdt.prototype.SetPlaceholderText;
	ApiBlockLvlSdt.prototype["GetPosInParent"]          = ApiBlockLvlSdt.prototype.GetPosInParent;
	ApiBlockLvlSdt.prototype["ReplaceByElement"]        = ApiBlockLvlSdt.prototype.ReplaceByElement;
	ApiBlockLvlSdt.prototype["AddComment"]              = ApiBlockLvlSdt.prototype.AddComment;
	ApiBlockLvlSdt.prototype["SetBackgroundColor"]      = ApiBlockLvlSdt.prototype.SetBackgroundColor;
	ApiBlockLvlSdt.prototype["AddCaption"]              = ApiBlockLvlSdt.prototype.AddCaption;
	ApiBlockLvlSdt.prototype["GetDropdownList"]         = ApiBlockLvlSdt.prototype.GetDropdownList;

	ApiFormBase.prototype["GetClassType"]        = ApiFormBase.prototype.GetClassType;
	ApiFormBase.prototype["GetFormType"]         = ApiFormBase.prototype.GetFormType;
	ApiFormBase.prototype["GetFormKey"]          = ApiFormBase.prototype.GetFormKey;
	ApiFormBase.prototype["SetFormKey"]          = ApiFormBase.prototype.SetFormKey;
	ApiFormBase.prototype["GetTipText"]          = ApiFormBase.prototype.GetTipText;
	ApiFormBase.prototype["SetTipText"]          = ApiFormBase.prototype.SetTipText;
	ApiFormBase.prototype["IsRequired"]          = ApiFormBase.prototype.IsRequired;
	ApiFormBase.prototype["SetRequired"]         = ApiFormBase.prototype.SetRequired;
	ApiFormBase.prototype["IsFixed"]             = ApiFormBase.prototype.IsFixed;
	ApiFormBase.prototype["ToFixed"]             = ApiFormBase.prototype.ToFixed;
	ApiFormBase.prototype["ToInline"]            = ApiFormBase.prototype.ToInline;
	ApiFormBase.prototype["SetBorderColor"]      = ApiFormBase.prototype.SetBorderColor;
	ApiFormBase.prototype["SetBackgroundColor"]  = ApiFormBase.prototype.SetBackgroundColor;
	ApiFormBase.prototype["GetText"]             = ApiFormBase.prototype.GetText;
	ApiFormBase.prototype["Clear"]               = ApiFormBase.prototype.Clear;
	ApiFormBase.prototype["GetWrapperShape"]     = ApiFormBase.prototype.GetWrapperShape;
	ApiFormBase.prototype["SetPlaceholderText"]  = ApiFormBase.prototype.SetPlaceholderText;
	ApiFormBase.prototype["SetTextPr"]           = ApiFormBase.prototype.SetTextPr;
	ApiFormBase.prototype["GetTextPr"]           = ApiFormBase.prototype.GetTextPr;

	ApiTextForm.prototype["IsAutoFit"]           = ApiTextForm.prototype.IsAutoFit;
	ApiTextForm.prototype["SetAutoFit"]          = ApiTextForm.prototype.SetAutoFit;
	ApiTextForm.prototype["IsMultiline"]         = ApiTextForm.prototype.IsMultiline;
	ApiTextForm.prototype["SetMultiline"]        = ApiTextForm.prototype.SetMultiline;
	ApiTextForm.prototype["GetCharactersLimit"]  = ApiTextForm.prototype.GetCharactersLimit;
	ApiTextForm.prototype["SetCharactersLimit"]  = ApiTextForm.prototype.SetCharactersLimit;
	ApiTextForm.prototype["IsComb"]              = ApiTextForm.prototype.IsComb;
	ApiTextForm.prototype["SetComb"]             = ApiTextForm.prototype.SetComb;
	ApiTextForm.prototype["SetCellWidth"]        = ApiTextForm.prototype.SetCellWidth;
	ApiTextForm.prototype["SetText"]             = ApiTextForm.prototype.SetText;
	ApiTextForm.prototype["Copy"]                = ApiTextForm.prototype.Copy;

	ApiPictureForm.prototype["GetScaleFlag"]       = ApiPictureForm.prototype.GetScaleFlag;
	ApiPictureForm.prototype["SetScaleFlag"]       = ApiPictureForm.prototype.SetScaleFlag;
	ApiPictureForm.prototype["SetLockAspectRatio"] = ApiPictureForm.prototype.SetLockAspectRatio;
	ApiPictureForm.prototype["IsLockAspectRatio"] = ApiPictureForm.prototype.IsLockAspectRatio;
	ApiPictureForm.prototype["SetPicturePosition"] = ApiPictureForm.prototype.SetPicturePosition;
	ApiPictureForm.prototype["GetPicturePosition"] = ApiPictureForm.prototype.GetPicturePosition;
	ApiPictureForm.prototype["SetRespectBorders"] = ApiPictureForm.prototype.SetRespectBorders;
	ApiPictureForm.prototype["IsRespectBorders"] = ApiPictureForm.prototype.IsRespectBorders;
	ApiPictureForm.prototype["GetImage"]           = ApiPictureForm.prototype.GetImage;
	ApiPictureForm.prototype["SetImage"]           = ApiPictureForm.prototype.SetImage;
	ApiPictureForm.prototype["Copy"]               = ApiPictureForm.prototype.Copy;

	ApiComboBoxForm.prototype["GetListValues"]       = ApiComboBoxForm.prototype.GetListValues;
	ApiComboBoxForm.prototype["SetListValues"]       = ApiComboBoxForm.prototype.SetListValues;
	ApiComboBoxForm.prototype["SelectListValue"]     = ApiComboBoxForm.prototype.SelectListValue;
	ApiComboBoxForm.prototype["SetText"]             = ApiComboBoxForm.prototype.SetText;
	ApiComboBoxForm.prototype["IsEditable"]          = ApiComboBoxForm.prototype.IsEditable;
	ApiComboBoxForm.prototype["Copy"]                = ApiComboBoxForm.prototype.Copy;

	ApiCheckBoxForm.prototype["SetChecked"]    = ApiCheckBoxForm.prototype.SetChecked;
	ApiCheckBoxForm.prototype["IsChecked"]     = ApiCheckBoxForm.prototype.IsChecked;
	ApiCheckBoxForm.prototype["IsRadioButton"] = ApiCheckBoxForm.prototype.IsRadioButton;
	ApiCheckBoxForm.prototype["GetRadioGroup"] = ApiCheckBoxForm.prototype.GetRadioGroup;
	ApiCheckBoxForm.prototype["SetRadioGroup"] = ApiCheckBoxForm.prototype.SetRadioGroup;
	ApiCheckBoxForm.prototype["Copy"]          = ApiCheckBoxForm.prototype.Copy;

	ApiComment.prototype["GetClassType"]	= ApiComment.prototype.GetClassType;
	ApiComment.prototype["GetId"]           = ApiComment.prototype.GetId;
	ApiComment.prototype["GetText"]			= ApiComment.prototype.GetText;
	ApiComment.prototype["SetText"]			= ApiComment.prototype.SetText;
	ApiComment.prototype["GetAuthorName"]	= ApiComment.prototype.GetAuthorName;
	ApiComment.prototype["SetAuthorName"]	= ApiComment.prototype.SetAuthorName;
	ApiComment.prototype["GetAutorName"]	= ApiComment.prototype.GetAuthorName; // compatibility with a typo in the old name
	ApiComment.prototype["SetAutorName"]	= ApiComment.prototype.SetAuthorName;
	ApiComment.prototype["GetUserId"]		= ApiComment.prototype.GetUserId;
	ApiComment.prototype["SetUserId"]		= ApiComment.prototype.SetUserId;
	ApiComment.prototype["IsSolved"]		= ApiComment.prototype.IsSolved;
	ApiComment.prototype["SetSolved"]		= ApiComment.prototype.SetSolved;
	ApiComment.prototype["GetTimeUTC"]		= ApiComment.prototype.GetTimeUTC;
	ApiComment.prototype["SetTimeUTC"]		= ApiComment.prototype.SetTimeUTC;
	ApiComment.prototype["GetTime"]			= ApiComment.prototype.GetTime;
	ApiComment.prototype["SetTime"]			= ApiComment.prototype.SetTime;
	ApiComment.prototype["GetQuoteText"]	= ApiComment.prototype.GetQuoteText;
	ApiComment.prototype["GetRepliesCount"]	= ApiComment.prototype.GetRepliesCount;
	ApiComment.prototype["GetReply"]		= ApiComment.prototype.GetReply;
	ApiComment.prototype["AddReply"]		= ApiComment.prototype.AddReply;
	ApiComment.prototype["RemoveReplies"]	= ApiComment.prototype.RemoveReplies;
	ApiComment.prototype["Delete"]			= ApiComment.prototype.Delete;

	ApiCommentReply.prototype["GetClassType"]	= ApiCommentReply.prototype.GetClassType;
	ApiCommentReply.prototype["GetText"]		= ApiCommentReply.prototype.GetText;
	ApiCommentReply.prototype["SetText"]		= ApiCommentReply.prototype.SetText;
	ApiCommentReply.prototype["GetAuthorName"]	= ApiCommentReply.prototype.GetAuthorName;
	ApiCommentReply.prototype["SetAuthorName"]	= ApiCommentReply.prototype.SetAuthorName;
	ApiCommentReply.prototype["GetAutorName"]	= ApiCommentReply.prototype.GetAuthorName; // compatibility with a typo in the old name
	ApiCommentReply.prototype["SetAutorName"]	= ApiCommentReply.prototype.SetAuthorName;
	ApiCommentReply.prototype["GetUserId"]		= ApiCommentReply.prototype.GetUserId;
	ApiCommentReply.prototype["SetUserId"]		= ApiCommentReply.prototype.SetUserId;

	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Export for internal usage
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	window['AscBuilder'] = window['AscBuilder'] || {};
	window['AscBuilder'].Api                = Api;
	window['AscBuilder'].ApiDocumentContent = ApiDocumentContent;
	window['AscBuilder'].ApiRange           = ApiRange;
	window['AscBuilder'].ApiDocument        = ApiDocument;
	window['AscBuilder'].ApiParagraph       = ApiParagraph;
	window['AscBuilder'].ApiRun             = ApiRun;
	window['AscBuilder'].ApiHyperlink       = ApiHyperlink;
	window['AscBuilder'].ApiSection         = ApiSection;
	window['AscBuilder'].ApiTable           = ApiTable;
	window['AscBuilder'].ApiTableRow        = ApiTableRow;
	window['AscBuilder'].ApiTableCell       = ApiTableCell;
	window['AscBuilder'].ApiStyle           = ApiStyle;
	window['AscBuilder'].ApiNumbering       = ApiNumbering;
	window['AscBuilder'].ApiNumberingLevel  = ApiNumberingLevel;
	window['AscBuilder'].ApiTextPr          = ApiTextPr;
	window['AscBuilder'].ApiParaPr          = ApiParaPr;
	window['AscBuilder'].ApiTablePr         = ApiTablePr;
	window['AscBuilder'].ApiTableRowPr      = ApiTableRowPr;
	window['AscBuilder'].ApiTableCellPr     = ApiTableCellPr;
	window['AscBuilder'].ApiTableStylePr    = ApiTableStylePr;
	window['AscBuilder'].ApiDrawing         = ApiDrawing;
	window['AscBuilder'].ApiImage           = ApiImage;
	window['AscBuilder'].ApiShape           = ApiShape;
	window['AscBuilder'].ApiChart           = ApiChart;
	window['AscBuilder'].ApiInlineLvlSdt    = ApiInlineLvlSdt;
	window['AscBuilder'].ApiBlockLvlSdt     = ApiBlockLvlSdt;
	window['AscBuilder'].ApiFormBase        = ApiFormBase;
	window['AscBuilder'].ApiTextForm        = ApiTextForm;
	window['AscBuilder'].ApiPictureForm     = ApiPictureForm;
	window['AscBuilder'].ApiComboBoxForm    = ApiComboBoxForm;
	window['AscBuilder'].ApiCheckBoxForm    = ApiCheckBoxForm;
	window['AscBuilder'].ApiComplexForm     = ApiComplexForm;
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Area for internal usage
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function GetStringParameter(parameter, defaultValue)
	{
		if (undefined !== parameter && typeof(parameter) === "string" && "" !== parameter)
			return parameter;

		return defaultValue;
	}
	function GetBoolParameter(parameter, defaultValue)
	{
		if (undefined !== parameter && typeof(parameter) === "boolean")
			return parameter;

		return defaultValue;
	}
	function GetNumberParameter(parameter, defaultValue)
	{
		if (undefined !== parameter && typeof(parameter) === "number")
			return parameter;

		return defaultValue;
	}
	function GetArrayParameter(parameter, defaultValue)
	{
		if (undefined !== parameter && Array.isArray(parameter))
			return parameter;

		return defaultValue;
	}
	window['AscBuilder'].GetStringParameter = GetStringParameter;
	window['AscBuilder'].GetBoolParameter   = GetBoolParameter;
	window['AscBuilder'].GetNumberParameter = GetNumberParameter;
	window['AscBuilder'].GetArrayParameter  = GetArrayParameter;
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	// Private area
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function ToApiForm(oForm)
	{
		if (!oForm)
			return null;

		if (oForm.IsComplexForm())
			return new ApiComplexForm(oForm);
		else if (oForm.IsTextForm())
			return new ApiTextForm(oForm);
		else if (oForm.IsComboBox() || oForm.IsDropDownList())
			return new ApiComboBoxForm(oForm);
		else if (oForm.IsRadioButton() || oForm.IsCheckBox())
			return new ApiCheckBoxForm(oForm);
		else if (oForm.IsPictureForm())
			return new ApiPictureForm(oForm);

		return null;
	}
	function ToApiContentControl(oControl)
	{
		if (!oControl)
			return null;

		if (oControl instanceof CBlockLevelSdt)
			return (new ApiBlockLvlSdt(oControl));
		else if (oControl instanceof CInlineLevelSdt)
			return (new ApiInlineLvlSdt(oControl));

		return null;
	}
	function AddGlobalCommentToDocument(logicDocument, commentData)
	{
		return AddCommentToDocument(logicDocument, commentData, true);
	}
	function AddCommentToDocument(logicDocument, commentData, forceGlobal)
	{
		let comment = logicDocument.AddComment(commentData, !!forceGlobal);
		if (!comment)
			return null;

		comment.GenerateDurableId();
		logicDocument.GetApi().sync_AddComment(comment.GetId(), comment.GetData());
		return new ApiComment(comment);
	}

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
			|| oElement instanceof ApiHyperlink
			|| oElement instanceof ApiFormBase)
			return true;

		return false;
	}

	function private_GetSupportedParaElement(oElement)
	{
		if (oElement instanceof ParaRun)
		{
			let arrDrawings = oElement.GetAllDrawingObjects();
			for (let nIndex = 0, nCount = arrDrawings.length; nIndex < nCount; ++nIndex)
			{
				if (arrDrawings[nIndex].IsForm())
					return private_CheckForm(arrDrawings[nIndex].GetInnerForm());
			}

			return new ApiRun(oElement);
		}
		else if (oElement instanceof CInlineLevelSdt)
			return private_CheckForm(oElement);
		else if (oElement instanceof ParaHyperlink)
			return new ApiHyperlink(oElement);
		else if (oElement instanceof ApiFormBase)
			return (new ApiInlineLvlSdt(oElement)).GetForm();
		else
			return new ApiUnsupported();
	}

	function private_CheckForm(oSdt)
	{
		if (!oSdt)
			return new ApiUnsupported();

		if (oSdt.IsComplexForm())
			return new ApiComplexForm(oSdt);
		else if (oSdt.IsTextForm())
			return new ApiTextForm(oSdt);
		else if (oSdt.IsComboBox() || oSdt.IsDropDownList())
			return new ApiComboBoxForm(oSdt);
		else if (oSdt.IsCheckBox() || oSdt.IsRadioButton())
			return new ApiCheckBoxForm(oSdt);
		else if (oSdt.IsPictureForm())
			return new ApiPictureForm(oSdt)

		return new ApiInlineLvlSdt(oSdt);
	}

	function private_GetLogicDocument()
	{
		return editor.WordControl.m_oLogicDocument;
	}

	function private_Twips2MM(twips)
	{
		return 25.4 / 72.0 / 20 * twips;
	}

	function private_MM2Twips(mm)
	{
		return mm / (25.4 / 72.0 / 20);
	}

	function private_EMU2MM(EMU)
	{
		return EMU / 36000.0;
	}
	
	function private_MM2EMU(mm)
	{
		return mm * 36000.0;
	}

	function private_GetHps(hps)
	{
		if (hps < 0) {
			return - Math.ceil(Math.abs(hps)) / 2.0
		}
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

	function private_Pt2EMU(pt)
	{
		return 12700.0 * pt;
	}

	function private_MM2Pt(mm)
	{
		return mm / (25.4 / 72.0);
	};

	function private_EMU2Pt(emu)
	{
		return emu / 12700.0
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

	function private_GetDrawingLockType(sType)
	{
		var nLockType = -1;
		switch (sType)
		{
			case "noGrp":
				nLockType = AscFormat.LOCKS_MASKS.noGrp;
				break;
			case "noUngrp":
				nLockType = AscFormat.LOCKS_MASKS.noUngrp;
				break;
			case "noSelect":
				nLockType = AscFormat.LOCKS_MASKS.noSelect;
				break;
			case "noRot":
				nLockType = AscFormat.LOCKS_MASKS.noRot;
				break;
			case "noChangeAspect":
				nLockType = AscFormat.LOCKS_MASKS.noChangeAspect;
				break;
			case "noMove":
				nLockType = AscFormat.LOCKS_MASKS.noMove;
				break;
			case "noResize":
				nLockType = AscFormat.LOCKS_MASKS.noResize;
				break;
			case "noEditPoints":
				nLockType = AscFormat.LOCKS_MASKS.noEditPoints;
				break;
			case "noAdjustHandles":
				nLockType = AscFormat.LOCKS_MASKS.noAdjustHandles;
				break;
			case "noChangeArrowheads":
				nLockType = AscFormat.LOCKS_MASKS.noChangeArrowheads;
				break;
			case "noChangeShapeType":
				nLockType = AscFormat.LOCKS_MASKS.noChangeShapeType;
				break;
			case "noDrilldown":
				nLockType = AscFormat.LOCKS_MASKS.noDrilldown;
				break;
			case "noTextEdit":
				nLockType = AscFormat.LOCKS_MASKS.noTextEdit;
				break;
			case "noCrop":
				nLockType = AscFormat.LOCKS_MASKS.noCrop;
				break;
			case "txBox":
				nLockType = AscFormat.LOCKS_MASKS.txBox;
				break;
		}

		return nLockType;
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
		return new ApiShape(oDrawing.GraphicObj)
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

	/**
	 * Gets a document color object by color name.
	 * @param {highlightColor} - available highlight color
	 * @returns {object}
	 */
	function private_getHighlightColorByName(sColor)
	{
		var oColor;
		switch (sColor)
		{
			case "black":
				oColor = {r: 0, g: 0, b: 0};
				break;
			case "blue":
				oColor = {r: 0, g: 0, b: 255};
				break;
			case "cyan":
				oColor = {r: 0, g: 255, b: 255};
				break;
			case "green":
				oColor = {r: 0, g: 255, b: 0};
				break;
			case "magenta":
				oColor = {r: 255, g: 0, b: 255};
				break;
			case "red":
				oColor = {r: 255, g: 0, b: 0};
				break;
			case "yellow":
				oColor = {r: 255, g: 255, b: 0};
				break;
			case "white":
				oColor = {r: 255, g: 255, b: 255};
				break;
			case "darkBlue":
				oColor = {r: 0, g: 0, b: 139};
				break;
			case "darkCyan":
				oColor = {r: 0, g: 139, b: 139};
				break;
			case "darkGreen":
				oColor = {r: 0, g: 100, b: 0};
				break;
			case "darkMagenta":
				oColor = {r: 128, g: 0, b: 128};
				break;
			case "darkRed":
				oColor = {r: 139, g: 0, b: 0};
				break;
			case "darkYellow":
				oColor = {r: 128, g: 128, b: 0};
				break;
			case "darkGray":
				oColor = {r: 169, g: 169, b: 169};
				break;
			case "lightGray":
				oColor = {r: 211, g: 211, b: 211};
				break;
		}

		return oColor;
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
	ApiDocument.prototype.ForceRecalculate = function(nPage)
	{
		let oDocument = this.Document;
		let nOffCount = 0;
		while (!oDocument.Is_OnRecalculate())
		{
			nOffCount++;
			oDocument.TurnOn_Recalculate(false);
		}

		// oDocument.RecalculateAllAtOnce(false, nPage);
		// oDocument.FinalizeAction();
		// oDocument.GetHistory().TurnOn();
		// oDocument.StartAction();

		// TODO: В билдере не создаются точки истории, которые нужны для контролирования того, что нужно пересчитать.
		//       Поэтому пока мы все время вынуждены вести расчет с начала документа. Чтобы этого избежать можно
		//       создавать точки после расчета и финализировать их перед следующим (и включить саму историю)
		oDocument.RecalculateAllAtOnce(true, nPage);

		while (nOffCount > 0)
		{
			nOffCount--;
			oDocument.TurnOff_Recalculate();
		}
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
	/**
	 * Returns the next run in parent if exists.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @return {?ApiRun} - returns null if next run doesn't exist.
	 */
	ApiRun.prototype.private_GetNextInParent = function()
	{
		var oParent = this.Run.Parent || this.Run.Paragraph;
		if (!oParent)
			return null;

		var nRunIndex = this.Run.GetPosInParent();
		for (var nElm = nRunIndex + 1; nElm < oParent.Content.length - 1; nElm++)
		{
			if (oParent.Content[nElm] && oParent.Content[nElm].IsRun && oParent.Content[nElm].IsRun() === true)
				return new ApiRun(oParent.Content[nElm]);
		}

		return null;
	};
	/**
	 * Returns the previous run in parent if exists.
	 * @memberof ApiRun
	 * @typeofeditors ["CDE"]
	 * @return {?ApiRun} - returns null if previous run doesn't exist.
	 */
	ApiRun.prototype.private_GetPreviousInParent = function()
	{
		var oParent = this.Run.Parent || this.Run.Paragraph;
		if (!oParent)
			return null;

		var nRunIndex = this.Run.GetPosInParent();

		for (var nElm = nRunIndex - 1; nElm > - 1; nElm--)
		{
			if (oParent.Content[nElm] && oParent.Content[nElm].IsRun && oParent.Content[nElm].IsRun() === true)
				return new ApiRun(oParent.Content[nElm]);
		}

		return null;
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
		if (this.Parent)
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
	ApiContentControlList.prototype.GetListPr = function()
	{
		if (this.Sdt.IsComboBox())
			return this.Sdt.GetComboBoxPr();
		else
			return this.Sdt.GetDropDownListPr();
	};
	ApiContentControlList.prototype.SetListPr = function(newPr)
	{
		if (this.Sdt.IsComboBox())
			return this.Sdt.SetComboBoxPr(newPr);
		else
			return this.Sdt.SetDropDownListPr(newPr);
	};

	ApiFormBase.prototype.private_GetImpl = function()
	{
		let oShape;
		if (this.IsFixed() && (oShape = this.GetWrapperShape()))
		{
			let oRun = new ParaRun(null, false);
			oRun.AddToContent(0, oShape.Drawing);
			return oRun;
		}

		return this.Sdt;
	};
	ApiFormBase.prototype.OnChangeTextPr  = function(oApiTextPr)
	{
		this.Sdt.Apply_TextPr(oApiTextPr.TextPr);
		oApiTextPr.TextPr = this.Sdt.Pr.TextPr;
	};
	ApiFormBase.prototype.OnChangeValue = function()
	{
		let logicDocument = private_GetLogicDocument();
		if (!logicDocument || !logicDocument.IsDocumentEditor())
			return;
		
		logicDocument.ClearActionOnChangeForm();
		logicDocument.GetFormsManager().OnChange(this.Sdt);
	};
	
	ApiComment.prototype.private_OnChange = function()
	{
		let oLogicDocument = private_GetLogicDocument();
		oLogicDocument.EditComment(this.Comment.GetId(), this.Comment.GetData());
		editor.sync_ChangeCommentData(this.Comment.GetId(), this.Comment.GetData());
	};
	ApiCommentReply.prototype.private_OnChange = function()
	{
		let oComment = new ApiComment(this.Comment);
		oComment.private_OnChange();
	};

	Api.prototype.private_CreateApiParagraph = function(oParagraph){
		return new ApiParagraph(oParagraph);
	};
	Api.prototype.private_CreateTextPr = function(oParent, oTextPr){
		return new ApiTextPr(oParent, oTextPr);
	};

	Api.prototype.private_CreateApiDocContent = function(oDocContent){
		return new ApiDocumentContent(oDocContent);
	};

	Api.prototype.private_CreateCheckBoxForm = function(oCC){
		return new ApiCheckBoxForm(oCC);
	};
	Api.prototype.private_CreateTextForm = function(oCC){
		return new ApiTextForm(oCC);
	};
	Api.prototype.private_CreateComboBoxForm = function(oCC){
		return new ApiComboBoxForm(oCC);
	};
	Api.prototype.private_CreatePictureForm = function(oCC){
		return new ApiPictureForm(oCC);
	};
	Api.prototype.private_CreateComplexForm = function(oCC)
	{
		return new ApiComplexForm(oCC);
	};
	

	Api.prototype.private_createWordArt = function(oTextPr, sText, sTransform, oFill, oStroke, nRotAngle, nWidth, nHeight) {
		var oWorksheet, bWord, nFontSize;
		if (this.editorId === AscCommon.c_oEditorId.Spreadsheet)
			oWorksheet = this.GetActiveSheet().worksheet;
		else if (this.editorId === AscCommon.c_oEditorId.Presentation)
			bWord = false;
		else if (this.editorId === AscCommon.c_oEditorId.Word)
			bWord = true;

		var dAngle = nRotAngle !== 0 ? (Math.PI / 180) * nRotAngle : 0;
		var oArt = new AscFormat.CShape();

		oArt.setWordShape(bWord === true);
		oArt.setBDeleted(false);
		if (oWorksheet)
			oArt.setWorksheet(oWorksheet);

		if(bWord)
        {
            nFontSize = 36;
            oArt.createTextBoxContent();
        }
        else
        {
            nFontSize = 54;
            oArt.createTextBody();
        }

		var sText = typeof(sText) == "string" && sText !== "" ? sText : "Your text here";

		if (!oTextPr)
		{
			oTextPr = new CTextPr();
            oTextPr.FontSize = nFontSize;
            oTextPr.RFonts.Ascii = {Name: "Cambria Math", Index: -1};
            oTextPr.RFonts.HAnsi = {Name: "Cambria Math", Index: -1};
            oTextPr.RFonts.CS = {Name: "Cambria Math", Index: -1};
            oTextPr.RFonts.EastAsia = {Name: "Cambria Math", Index: -1};
		}
		if (!oTextPr.FontSize)
			oTextPr.FontSize = nFontSize;

		var oSpPr = new AscFormat.CSpPr();
        var oXfrm = new AscFormat.CXfrm();
		oXfrm.setOffX(0);
        oXfrm.setOffY(0);
        oXfrm.setExtX(nWidth/36000);
        oXfrm.setExtY(nHeight/36000);
        oSpPr.setXfrm(oXfrm);
        oXfrm.setParent(oSpPr);
		if (dAngle !== 0)
		{
			var dRot = AscFormat.normalizeRotate(dAngle);
            oXfrm.setRot(dRot);
		}

		oSpPr.setFill(oFill);
        oSpPr.setLn(oStroke);
        oSpPr.setGeometry(AscFormat.CreateGeometry("rect"));
        oArt.setSpPr(oSpPr);
        oSpPr.setParent(oArt);

		var oContent = oArt.getDocContent();
		AscFormat.AddToContentFromString(oContent, sText);

		oContent.SetApplyToAll(true);
        oContent.AddToParagraph(new ParaTextPr(oTextPr));
        oContent.SetParagraphAlign(AscCommon.align_Center);
        oContent.SetApplyToAll(false);
        var oBodyPr = oArt.getBodyPr().createDuplicate();
        oBodyPr.rot = 0;
        oBodyPr.spcFirstLastPara = false;
        oBodyPr.vertOverflow = AscFormat.nVOTOverflow;
        oBodyPr.horzOverflow = AscFormat.nHOTOverflow;
        oBodyPr.vert = AscFormat.nVertTThorz;
        oBodyPr.wrap = AscFormat.nTWTNone;
        oBodyPr.setDefaultInsets();
        oBodyPr.numCol = 1;
        oBodyPr.spcCol = 0;
        oBodyPr.rtlCol = 0;
        oBodyPr.fromWordArt = false;
        oBodyPr.anchor = 4;
        oBodyPr.anchorCtr = false;
        oBodyPr.forceAA = false;
        oBodyPr.compatLnSpc = true;
        oBodyPr.prstTxWarp = AscFormat.CreatePrstTxWarpGeometry(sTransform);
        oBodyPr.textFit = new AscFormat.CTextFit();
        oBodyPr.textFit.type = AscFormat.text_fit_Auto;
        if(bWord)
        {
            oArt.setBodyPr(oBodyPr);
        }
        else
        {
            oArt.txBody.setBodyPr(oBodyPr);
        }

		return oArt;
	};

	Api.prototype.private_CreateApiRun = function(oRun){
		return new ApiRun(oRun);
	};
	Api.prototype.private_CreateApiHyperlink = function(oHyperlink){
		return new ApiRun(oHyperlink);
	};
	Api.prototype.private_CreateApiTextPr = function(oTextPr){
		return new ApiTextPr(null, oTextPr);
	};
	Api.prototype.private_CreateApiParaPr = function(oParaPr){
		return new ApiParaPr(null, oParaPr);
	};
	Api.prototype.private_CreateApiFill = function(oFill){
		return new ApiFill(oFill);
	};
	Api.prototype.private_CreateApiStroke = function(oStroke){
		return new ApiStroke(oStroke);
	};
	Api.prototype.private_CreateApiGradStop = function(oApiUniColor, pos){
		return new ApiGradientStop(oApiUniColor, pos);
	};
	Api.prototype.private_CreateApiUniColor = function(oUniColor){
		return new ApiUniColor(oUniColor);
	};

}(window, null));



