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
	 * @constructor
	 */
	function CDocumentSectionProps(oSectPr, oLogicDocument)
	{
		if (oSectPr && oLogicDocument)
		{
			this.W      = oSectPr.GetPageWidth();
			this.H      = oSectPr.GetPageHeight();
			this.Orient = oSectPr.GetOrientation();
			
			this.Left   = oSectPr.GetPageMarginLeft();
			this.Top    = oSectPr.GetPageMarginTop();
			this.Right  = oSectPr.GetPageMarginRight();
			this.Bottom = oSectPr.GetPageMarginBottom();
			
			this.Header = oSectPr.GetPageMarginHeader();
			this.Footer = oSectPr.GetPageMarginFooter();
			
			this.Gutter        = oSectPr.GetGutter();
			this.GutterRTL     = oSectPr.IsGutterRTL();
			this.GutterAtTop   = oLogicDocument.IsGutterAtTop();
			this.MirrorMargins = oLogicDocument.IsMirrorMargins();
		}
		else
		{
			this.W      = undefined;
			this.H      = undefined;
			this.Orient = undefined;
			
			this.Left   = undefined;
			this.Top    = undefined;
			this.Right  = undefined;
			this.Bottom = undefined;
			
			this.Header = undefined;
			this.Footer = undefined;
			
			this.Gutter        = undefined;
			this.GutterRTL     = undefined;
			this.GutterAtTop   = undefined;
			this.MirrorMargins = undefined;
		}
	}
	CDocumentSectionProps.prototype.get_W = function()
	{
		return this.W;
	};
	CDocumentSectionProps.prototype.put_W = function(W)
	{
		this.W = W;
	};
	CDocumentSectionProps.prototype.get_H = function()
	{
		return this.H;
	};
	CDocumentSectionProps.prototype.put_H = function(H)
	{
		this.H = H;
	};
	CDocumentSectionProps.prototype.get_Orientation = function()
	{
		return this.Orient;
	};
	CDocumentSectionProps.prototype.put_Orientation = function(Orient)
	{
		this.Orient = Orient;
	};
	CDocumentSectionProps.prototype.get_LeftMargin = function()
	{
		return this.Left;
	};
	CDocumentSectionProps.prototype.put_LeftMargin = function(Left)
	{
		this.Left = Left;
	};
	CDocumentSectionProps.prototype.get_TopMargin = function()
	{
		return this.Top;
	};
	CDocumentSectionProps.prototype.put_TopMargin = function(Top)
	{
		this.Top = Top;
	};
	CDocumentSectionProps.prototype.get_RightMargin = function()
	{
		return this.Right;
	};
	CDocumentSectionProps.prototype.put_RightMargin = function(Right)
	{
		this.Right = Right;
	};
	CDocumentSectionProps.prototype.get_BottomMargin = function()
	{
		return this.Bottom;
	};
	CDocumentSectionProps.prototype.put_BottomMargin = function(Bottom)
	{
		this.Bottom = Bottom;
	};
	CDocumentSectionProps.prototype.get_HeaderDistance = function()
	{
		return this.Header;
	};
	CDocumentSectionProps.prototype.put_HeaderDistance = function(Header)
	{
		this.Header = Header;
	};
	CDocumentSectionProps.prototype.get_FooterDistance = function()
	{
		return this.Footer;
	};
	CDocumentSectionProps.prototype.put_FooterDistance = function(Footer)
	{
		this.Footer = Footer;
	};
	CDocumentSectionProps.prototype.get_Gutter = function()
	{
		return this.Gutter;
	};
	CDocumentSectionProps.prototype.put_Gutter = function(nGutter)
	{
		this.Gutter = nGutter;
	};
	CDocumentSectionProps.prototype.get_GutterRTL = function()
	{
		return this.GutterRTL;
	};
	CDocumentSectionProps.prototype.put_GutterRTL = function(isRTL)
	{
		this.GutterRTL = isRTL;
	};
	CDocumentSectionProps.prototype.get_GutterAtTop = function()
	{
		return this.GutterAtTop;
	};
	CDocumentSectionProps.prototype.put_GutterAtTop = function(isAtTop)
	{
		this.GutterAtTop = isAtTop;
	};
	CDocumentSectionProps.prototype.get_MirrorMargins = function()
	{
		return this.MirrorMargins;
	};
	CDocumentSectionProps.prototype.put_MirrorMargins = function(isMirrorMargins)
	{
		this.MirrorMargins = isMirrorMargins;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['Asc']['CDocumentSectionProps'] = window['Asc'].CDocumentSectionProps = CDocumentSectionProps;
	CDocumentSectionProps.prototype["get_W"]              = CDocumentSectionProps.prototype.get_W;
	CDocumentSectionProps.prototype["put_W"]              = CDocumentSectionProps.prototype.put_W;
	CDocumentSectionProps.prototype["get_H"]              = CDocumentSectionProps.prototype.get_H;
	CDocumentSectionProps.prototype["put_H"]              = CDocumentSectionProps.prototype.put_H;
	CDocumentSectionProps.prototype["get_Orientation"]    = CDocumentSectionProps.prototype.get_Orientation;
	CDocumentSectionProps.prototype["put_Orientation"]    = CDocumentSectionProps.prototype.put_Orientation;
	CDocumentSectionProps.prototype["get_LeftMargin"]     = CDocumentSectionProps.prototype.get_LeftMargin;
	CDocumentSectionProps.prototype["put_LeftMargin"]     = CDocumentSectionProps.prototype.put_LeftMargin;
	CDocumentSectionProps.prototype["get_TopMargin"]      = CDocumentSectionProps.prototype.get_TopMargin;
	CDocumentSectionProps.prototype["put_TopMargin"]      = CDocumentSectionProps.prototype.put_TopMargin;
	CDocumentSectionProps.prototype["get_RightMargin"]    = CDocumentSectionProps.prototype.get_RightMargin;
	CDocumentSectionProps.prototype["put_RightMargin"]    = CDocumentSectionProps.prototype.put_RightMargin;
	CDocumentSectionProps.prototype["get_BottomMargin"]   = CDocumentSectionProps.prototype.get_BottomMargin;
	CDocumentSectionProps.prototype["put_BottomMargin"]   = CDocumentSectionProps.prototype.put_BottomMargin;
	CDocumentSectionProps.prototype["get_HeaderDistance"] = CDocumentSectionProps.prototype.get_HeaderDistance;
	CDocumentSectionProps.prototype["put_HeaderDistance"] = CDocumentSectionProps.prototype.put_HeaderDistance;
	CDocumentSectionProps.prototype["get_FooterDistance"] = CDocumentSectionProps.prototype.get_FooterDistance;
	CDocumentSectionProps.prototype["put_FooterDistance"] = CDocumentSectionProps.prototype.put_FooterDistance;
	CDocumentSectionProps.prototype["get_Gutter"]         = CDocumentSectionProps.prototype.get_Gutter;
	CDocumentSectionProps.prototype["put_Gutter"]         = CDocumentSectionProps.prototype.put_Gutter;
	CDocumentSectionProps.prototype["get_GutterRTL"]      = CDocumentSectionProps.prototype.get_GutterRTL;
	CDocumentSectionProps.prototype["put_GutterRTL"]      = CDocumentSectionProps.prototype.put_GutterRTL;
	CDocumentSectionProps.prototype["get_GutterAtTop"]    = CDocumentSectionProps.prototype.get_GutterAtTop;
	CDocumentSectionProps.prototype["put_GutterAtTop"]    = CDocumentSectionProps.prototype.put_GutterAtTop;
	CDocumentSectionProps.prototype["get_MirrorMargins"]  = CDocumentSectionProps.prototype.get_MirrorMargins;
	CDocumentSectionProps.prototype["put_MirrorMargins"]  = CDocumentSectionProps.prototype.put_MirrorMargins;
	
	/**
	 * @constructor
	 */
	function SectionPageNumProps()
	{
		this.Start     = undefined;
		this.Format    = undefined;
		this.ChapStyle = undefined;
		this.ChapSep   = undefined;
	}
	SectionPageNumProps.prototype.get_Start = function()
	{
		return this.Start;
	};
	SectionPageNumProps.prototype.put_Start = function(v)
	{
		this.Start = v;
	};
	SectionPageNumProps.prototype.get_Format = function()
	{
		return this.Format;
	};
	SectionPageNumProps.prototype.put_Format = function(v)
	{
		this.Format = v;
	};
	SectionPageNumProps.prototype.get_ChapStyle = function()
	{
		return this.ChapStyle;
	};
	SectionPageNumProps.prototype.put_ChapStyle = function(v)
	{
		this.ChapStyle = v;
	};
	SectionPageNumProps.prototype.get_ChapSep = function()
	{
		return this.ChapSep;
	};
	SectionPageNumProps.prototype.put_ChapSep = function(v)
	{
		this.ChapSep = v;
	};
	//--------------------------------------------------------export----------------------------------------------------
	window['Asc']['SectionPageNumProps'] = window['Asc'].SectionPageNumProps = SectionPageNumProps;
	SectionPageNumProps.prototype["get_Start"]     = SectionPageNumProps.prototype.get_Start;
	SectionPageNumProps.prototype["put_Start"]     = SectionPageNumProps.prototype.put_Start;
	SectionPageNumProps.prototype["get_Format"]    = SectionPageNumProps.prototype.get_Format;
	SectionPageNumProps.prototype["put_Format"]    = SectionPageNumProps.prototype.put_Format;
	SectionPageNumProps.prototype["get_ChapStyle"] = SectionPageNumProps.prototype.get_ChapStyle;
	SectionPageNumProps.prototype["put_ChapStyle"] = SectionPageNumProps.prototype.put_ChapStyle;
	SectionPageNumProps.prototype["get_ChapSep"]   = SectionPageNumProps.prototype.get_ChapSep;
	SectionPageNumProps.prototype["put_ChapSep"]   = SectionPageNumProps.prototype.put_ChapSep;
})(window);
