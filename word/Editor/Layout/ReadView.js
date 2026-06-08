/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
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
			oThis.SectPr   = new AscWord.SectPr(oLogicDocument);
			oThis.SectInfo = new AscWord.DocumentSection(oThis.SectPr, 0);
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
	CDocumentReadView.prototype.GetSectionByElement = function(element)
	{
		if (this.OriginalSectPr)
			return AscWord.CDocumentLayoutBase.prototype.GetSectionByElement.apply(this, arguments);
		
		return this.SectPr;
	};
	CDocumentReadView.prototype.CheckSectPr = function(sectPr)
	{
		if (this.OriginalSectPr)
			return AscWord.CDocumentLayoutBase.prototype.CheckSectPr.apply(this, arguments);
		
		return this.SectPr;
	};
	CDocumentReadView.prototype.GetFinalSectPr = function()
	{
		return this.SectPr;
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
