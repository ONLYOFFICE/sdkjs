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

(function(window)
{
	/**
	 * Класс для работы с документом в режиме читалки
	 * @param oLogicDocument
	 * @extends {AscWord.CDocumentLayoutBase}
	 * @constructor
	 */
	function CDocumentReadView(oLogicDocument)
	{
		AscWord.CDocumentLayoutBase.call(this, oLogicDocument);
		
		this.W        = 297;
		this.H        = 210;
		this.Scale    = 1;
		this.SectPr   = null;
		this.SectInfo = null;
		
		let oThis = this;
		AscCommon.ExecuteNoHistory(function()
		{
			oThis.SectPr   = new CSectionPr(oLogicDocument);
			oThis.SectInfo = new CDocumentSectionsInfoElement(oThis.SectPr, 0);
		}, oLogicDocument);
		
		this.OriginalSectPr = false;
	}
	
	CDocumentReadView.prototype = Object.create(AscWord.CDocumentLayoutBase.prototype);
	CDocumentReadView.prototype.constructor = CDocumentReadView;
	CDocumentReadView.prototype.IsReadMode = function()
	{
		return true;
	};
	CDocumentReadView.prototype.Set = function(nW, nH, nScale)
	{
		this.W     = nW;
		this.H     = nH;
		this.Scale = nScale;
		
		let oSectPr = this.SectPr;
		
		AscCommon.ExecuteNoHistory(function()
		{
			oSectPr.SetPageSize(nW, nH);
			oSectPr.SetPageMargins(5, 5, 5, 5);
		}, this.LogicDocument);
	};
	CDocumentReadView.prototype.IsHeaders = function()
	{
		return false;
	};
	CDocumentReadView.prototype.GetFontScale = function()
	{
		return this.Scale;
	};
	CDocumentReadView.prototype.GetSectionHdrFtr = function(nPageAbs, isFirst, isEven)
	{
		if (this.OriginalSectPr)
			return AscWord.CDocumentLayoutBase.prototype.GetSectionHdrFtr.apply(this, arguments);
		
		return {
			Header : null,
			Footer : null,
			SectPr : this.SectPr
		};
	};
	CDocumentReadView.prototype.GetPageContentFrame = function(nPageAbs, oSectPr)
	{
		let oFrame = this.SectPr.GetContentFrame(0);
		return {
			X      : oFrame.Left,
			Y      : oFrame.Top,
			XLimit : oFrame.Right,
			YLimit : oFrame.Bottom
		};
	};
	CDocumentReadView.prototype.GetColumnContentFrame = function(nPageAbs, nColumnAbs, oSectPr)
	{
		let oFrame = oSectPr.GetContentFrame(nPageAbs);
		return {
			X                 : oFrame.Left,
			Y                 : oFrame.Top,
			XLimit            : oFrame.Right,
			YLimit            : oFrame.Bottom,
			ColumnSpaceBefore : 0,
			ColumnSpaceAfter  : 0
		};
	};
	CDocumentReadView.prototype.GetSection = function(nPageAbs, nContentIndex)
	{
		if (this.OriginalSectPr)
			return AscWord.CDocumentLayoutBase.prototype.GetSection.apply(this, arguments);
		
		return this.SectPr;
	};
	CDocumentReadView.prototype.GetSectionByPos = function(nContentIndex)
	{
		if (this.OriginalSectPr)
			return AscWord.CDocumentLayoutBase.prototype.GetSectionByPos.apply(this, arguments);
		
		return this.SectPr;
	};
	CDocumentReadView.prototype.GetSectionInfo = function(nContentIndex)
	{
		if (this.OriginalSectPr)
			return AscWord.CDocumentLayoutBase.prototype.GetSectionInfo.apply(this, arguments);
		
		return this.SectInfo;
	};
	CDocumentReadView.prototype.GetSectionIndex = function(oSectPr)
	{
		if (this.OriginalSectPr)
			return AscWord.CDocumentLayoutBase.prototype.GetSectionIndex.apply(this, arguments);
		
		return 0;
	};
	CDocumentReadView.prototype.GetCalculateTimeLimit = function()
	{
		return 100;
	};
	CDocumentReadView.prototype.GetScaleBySection = function(oSectPr)
	{
		let nW = oSectPr.GetPageWidth();
		let nH = oSectPr.GetPageHeight();
		
		let nCoef = 1;
		if (this.W < nW)
			nCoef = this.W / nW;
		
		if (this.H < nH)
			nCoef = Math.min(this.H / nH, nCoef);
		
		return nCoef;
	};
	CDocumentReadView.prototype.calculateIndent = function(ind, element)
	{
		if (!element || !element.Get_SectPr)
			return ind;
		
		this.OriginalSectPr = true;
		let sectPr = element.Get_SectPr();
		this.OriginalSectPr = false;
		
		if (ind > 0 && sectPr)
			return ind * this.W / sectPr.GetPageWidth();
		
		return Math.max(ind, -2);
	};
	
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CDocumentReadView = CDocumentReadView;

})(window);
