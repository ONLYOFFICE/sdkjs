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
	 * Класс для работы с документом в режиме для печати (стандартный вариант)
	 * @param oLogicDocument {AscWord.CDocument}
	 * @extends {AscWord.CDocumentLayoutBase}
	 * @constructor
	 */
	function CDocumentPrintView(oLogicDocument)
	{
		AscWord.CDocumentLayoutBase.call(this, oLogicDocument);

		this.SectionsInfo = oLogicDocument.GetSections();
	}
	CDocumentPrintView.prototype = Object.create(AscWord.CDocumentLayoutBase.prototype);
	CDocumentPrintView.prototype.constructor = CDocumentPrintView;
	CDocumentPrintView.prototype.IsPrintMode = function()
	{
		return true;
	};
	CDocumentPrintView.prototype.GetPageContentFrame = function(nPageAbs, oSectPr)
	{
		let oFrame = oSectPr.GetContentFrame(nPageAbs);
		let Y      = oFrame.Top;
		let YLimit = oFrame.Bottom;
		let X      = oFrame.Left;
		let XLimit = oFrame.Right;

		let HdrFtrLine = this.LogicDocument.HdrFtr.GetHdrFtrLines(nPageAbs);

		let YHeader = HdrFtrLine.Top;
		if (null !== YHeader && YHeader > Y && oSectPr.GetPageMarginTop() >= 0)
			Y = YHeader;

		let YFooter = HdrFtrLine.Bottom;
		if (null !== YFooter && YFooter < YLimit && oSectPr.GetPageMarginBottom() >= 0)
			YLimit = YFooter;

		return {
			X      : X,
			Y      : Y,
			XLimit : XLimit,
			YLimit : YLimit
		};
	};
	CDocumentPrintView.prototype.GetColumnContentFrame = function(nPageAbs, nColumnAbs, oSectPr)
	{
		var oFrame = oSectPr.GetContentFrame(nPageAbs);

		var Y      = oFrame.Top;
		var YLimit = oFrame.Bottom;
		var X      = oFrame.Left;
		var XLimit = oFrame.Right;

		let nSectionIndex = this.LogicDocument.FullRecalc.SectionIndex;
		let oPage         = this.LogicDocument.GetPage(nPageAbs);
		let oPageSection  = oPage ? oPage.GetSection(nSectionIndex) : null;
		if (oPageSection)
		{
			Y      = oPageSection.GetY();
			YLimit = oPageSection.GetYLimit();
		}

		let nColumnsCount = oSectPr.GetColumnCount();
		for (let nColumnIndex = 0; nColumnIndex < nColumnAbs; ++nColumnIndex)
		{
			X += oSectPr.GetColumnWidth(nColumnIndex);
			X += oSectPr.GetColumnSpace(nColumnIndex);
		}

		if (nColumnsCount - 1 !== nColumnAbs)
			XLimit = X + oSectPr.GetColumnWidth(nColumnAbs);

		let HdrFtrLine = this.LogicDocument.HdrFtr.GetHdrFtrLines(nPageAbs);

		let YHeader = HdrFtrLine.Top;
		if (null !== YHeader && YHeader > Y && oSectPr.GetPageMarginTop() >= 0)
			Y = YHeader;

		let YFooter = HdrFtrLine.Bottom;
		if (null !== YFooter && YFooter < YLimit && oSectPr.GetPageMarginBottom() >= 0)
			YLimit = YFooter;

		let nColumnSpaceBefore = (nColumnAbs > 0 ? oSectPr.GetColumnSpace(nColumnAbs - 1) : 0);
		let nColumnSpaceAfter  = (nColumnAbs < nColumnsCount - 1 ? oSectPr.GetColumnSpace(nColumnAbs) : 0);

		return {
			X                 : X,
			Y                 : Y,
			XLimit            : XLimit,
			YLimit            : YLimit,
			ColumnSpaceBefore : nColumnSpaceBefore,
			ColumnSpaceAfter  : nColumnSpaceAfter
		};
	};
	CDocumentPrintView.prototype.GetCalculateTimeLimit = function()
	{
		return 10;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['AscWord'].CDocumentPrintView = CDocumentPrintView;

})(window);
