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

(function(window)
{
	/**
	 * Данный класс рассчитывает отображаемый стиль в панели стилей
	 * @constructor
	 */
	function CDisplayStyleCalculator()
	{
		this.PStyle = null;
		this.RStyle = null;
	}
	CDisplayStyleCalculator.prototype.CalculateName = function(docContent)
	{
		this.Reset();
		
		if (docContent.IsNumberingSelection())
			this.HandleNumberingSelection(docContent);
		else if (docContent.IsTextSelectionUse() && !docContent.IsSelectionEmpty())
			this.HandleRegularSelection(docContent);
		else
			this.HandleNoSelection(docContent);
		
		let logicDocument = docContent.GetLogicDocument();
		if (!logicDocument)
			return "";
		
		let styleManager = logicDocument.GetStyleManager();
		let styleId = this.RStyle ? this.RStyle : this.PStyle;
		
		if (!styleId)
			return "";
		else if (!styleManager.Get(styleId))
			return styleManager.Get(styleManager.GetDefaultParagraph()).GetName();
		else
			return styleManager.Get(styleId).GetName();
	};
	//------------------------------------------------------------------------------------------------------------------
	CDisplayStyleCalculator.prototype.Reset = function()
	{
		this.PStyle = null;
		this.RStyle = null;
	};
	CDisplayStyleCalculator.prototype.HandleNumberingSelection = function(docContent)
	{
		let paragraph = docContent.GetCurrentParagraph();
		if (!paragraph)
			return;
		
		this.RStyle = null;
		this.PStyle = paragraph.GetParagraphStyle();
	};
	CDisplayStyleCalculator.prototype.HandleRegularSelection = function(docContent)
	{
		let paragraphs = docContent.GetSelectedParagraphs();
		if (!paragraphs.length)
			return;
		
		let pStyle = paragraphs[0].GetParagraphStyle();
		for (let iPara = 1, nPara = paragraphs.length; iPara < nPara; ++iPara)
		{
			if (pStyle !== paragraphs[iPara].GetParagraphStyle())
			{
				pStyle = null;
				break;
			}
		}
		
		let rStyle       = null;
		let overAllCount = 0;
		let overAllMax   = 10000;  // Не учитываем более 10000 ранов
		docContent.CheckSelectedRunContent(function(run, startPos, endPos)
		{
			if (undefined === rStyle || overAllCount >= overAllMax)
				return true;
			
			let check = false;
			for (let pos = startPos; pos < endPos; ++pos)
			{
				let item = run.GetElement(pos);
				if (item.IsText() || item.IsSpace())
				{
					check = true;
					++overAllCount;
				}
			}
			
			if (check)
			{
				if (null === rStyle)
					rStyle = run.GetRStyle();
				else if (rStyle !== run.GetRStyle())
					rStyle = undefined;
			}
			
			return (undefined === rStyle);
		});
		
		let logicDocument = docContent.GetLogicDocument();
		if (logicDocument && rStyle === logicDocument.GetStyleManager().Get_Default_Character())
			rStyle = null;
		
		this.PStyle = pStyle;
		this.RStyle = rStyle;
	};
	CDisplayStyleCalculator.prototype.HandleNoSelection = function(docContent)
	{
		let paragraph = docContent.GetCurrentParagraph();
		if (!paragraph)
			return;
		
		let paraContentPos = paragraph.GetParaContentPos(false);
		if (!paraContentPos)
			return;
		
		let run = paragraph.GetClassByPos(paraContentPos);
		if (!run || !(run instanceof AscWord.CRun))
			return;
		
		this.RStyle = run.GetRStyle();
		this.PStyle = paragraph.GetParagraphStyle();
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'] = window['AscWord'] || {};
	window['AscWord'].DisplayStyleCalculator = new CDisplayStyleCalculator();
	
})(window);
