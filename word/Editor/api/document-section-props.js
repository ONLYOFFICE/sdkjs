/*
 * (c) Copyright Ascensio System SIA 2010-2025
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
	
})(window);
