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
	 * Класс для работы с документом в режиме для печати (стандартный вариант)
	 * @param oLogicDocument {AscWord.CDocument}
	 * @extends {AscWord.CDocumentLayoutBase}
	 * @constructor
	 */
	function CDocumentPrintView(oLogicDocument)
	{
		AscWord.CDocumentLayoutBase.call(this, oLogicDocument);

		this.SectionsInfo = oLogicDocument.GetSectionsInfo();
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

		let nColumnsCount = oSectPr.GetColumnsCount();
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
