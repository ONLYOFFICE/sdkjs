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
	AscDFH.historyitem_Section_PageSize_Orient       = AscDFH.historyitem_type_Section | 1;
	AscDFH.historyitem_Section_PageSize_Size         = AscDFH.historyitem_type_Section | 2;
	AscDFH.historyitem_Section_PageMargins           = AscDFH.historyitem_type_Section | 3;
	AscDFH.historyitem_Section_Type                  = AscDFH.historyitem_type_Section | 4;
	AscDFH.historyitem_Section_Borders_Left          = AscDFH.historyitem_type_Section | 5;
	AscDFH.historyitem_Section_Borders_Top           = AscDFH.historyitem_type_Section | 6;
	AscDFH.historyitem_Section_Borders_Right         = AscDFH.historyitem_type_Section | 7;
	AscDFH.historyitem_Section_Borders_Bottom        = AscDFH.historyitem_type_Section | 8;
	AscDFH.historyitem_Section_Borders_Display       = AscDFH.historyitem_type_Section | 9;
	AscDFH.historyitem_Section_Borders_OffsetFrom    = AscDFH.historyitem_type_Section | 10;
	AscDFH.historyitem_Section_Borders_ZOrder        = AscDFH.historyitem_type_Section | 11;
	AscDFH.historyitem_Section_Header_First          = AscDFH.historyitem_type_Section | 12;
	AscDFH.historyitem_Section_Header_Even           = AscDFH.historyitem_type_Section | 13;
	AscDFH.historyitem_Section_Header_Default        = AscDFH.historyitem_type_Section | 14;
	AscDFH.historyitem_Section_Footer_First          = AscDFH.historyitem_type_Section | 15;
	AscDFH.historyitem_Section_Footer_Even           = AscDFH.historyitem_type_Section | 16;
	AscDFH.historyitem_Section_Footer_Default        = AscDFH.historyitem_type_Section | 17;
	AscDFH.historyitem_Section_TitlePage             = AscDFH.historyitem_type_Section | 18;
	AscDFH.historyitem_Section_PageMargins_Header    = AscDFH.historyitem_type_Section | 19;
	AscDFH.historyitem_Section_PageMargins_Footer    = AscDFH.historyitem_type_Section | 20;
	AscDFH.historyitem_Section_PageNumType_Start     = AscDFH.historyitem_type_Section | 21;
	AscDFH.historyitem_Section_Columns_EqualWidth    = AscDFH.historyitem_type_Section | 22;
	AscDFH.historyitem_Section_Columns_Space         = AscDFH.historyitem_type_Section | 23;
	AscDFH.historyitem_Section_Columns_Num           = AscDFH.historyitem_type_Section | 24;
	AscDFH.historyitem_Section_Columns_Sep           = AscDFH.historyitem_type_Section | 25;
	AscDFH.historyitem_Section_Columns_Col           = AscDFH.historyitem_type_Section | 26;
	AscDFH.historyitem_Section_Columns_SetCols       = AscDFH.historyitem_type_Section | 27;
	AscDFH.historyitem_Section_Footnote_Pos          = AscDFH.historyitem_type_Section | 28;
	AscDFH.historyitem_Section_Footnote_NumStart     = AscDFH.historyitem_type_Section | 29;
	AscDFH.historyitem_Section_Footnote_NumRestart   = AscDFH.historyitem_type_Section | 30;
	AscDFH.historyitem_Section_Footnote_NumFormat    = AscDFH.historyitem_type_Section | 31;
	AscDFH.historyitem_Section_PageMargins_Gutter    = AscDFH.historyitem_type_Section | 32;
	AscDFH.historyitem_Section_Gutter_RTL            = AscDFH.historyitem_type_Section | 33;
	AscDFH.historyitem_Section_Endnote_Pos           = AscDFH.historyitem_type_Section | 34;
	AscDFH.historyitem_Section_Endnote_NumStart      = AscDFH.historyitem_type_Section | 35;
	AscDFH.historyitem_Section_Endnote_NumRestart    = AscDFH.historyitem_type_Section | 36;
	AscDFH.historyitem_Section_Endnote_NumFormat     = AscDFH.historyitem_type_Section | 37;
	AscDFH.historyitem_Section_LnNumType             = AscDFH.historyitem_type_Section | 38;
	AscDFH.historyitem_Section_PageNumType_Format    = AscDFH.historyitem_type_Section | 39;
	AscDFH.historyitem_Section_PageNumType_ChapStyle = AscDFH.historyitem_type_Section | 40;
	AscDFH.historyitem_Section_PageNumType_ChapSep   = AscDFH.historyitem_type_Section | 41;
	
	//------------------------------------------------------------------------------------------------------------------
	// Карта зависимости изменений
	//------------------------------------------------------------------------------------------------------------------
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageSize_Orient]       = [AscDFH.historyitem_Section_PageSize_Orient];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageSize_Size]         = [AscDFH.historyitem_Section_PageSize_Size];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageMargins]           = [AscDFH.historyitem_Section_PageMargins];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Type]                  = [AscDFH.historyitem_Section_Type];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Borders_Left]          = [AscDFH.historyitem_Section_Borders_Left];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Borders_Top]           = [AscDFH.historyitem_Section_Borders_Top];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Borders_Right]         = [AscDFH.historyitem_Section_Borders_Right];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Borders_Bottom]        = [AscDFH.historyitem_Section_Borders_Bottom];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Borders_Display]       = [AscDFH.historyitem_Section_Borders_Display];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Borders_OffsetFrom]    = [AscDFH.historyitem_Section_Borders_OffsetFrom];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Borders_ZOrder]        = [AscDFH.historyitem_Section_Borders_ZOrder];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Header_First]          = [AscDFH.historyitem_Section_Header_First];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Header_Even]           = [AscDFH.historyitem_Section_Header_Even];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Header_Default]        = [AscDFH.historyitem_Section_Header_Default];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Footer_First]          = [AscDFH.historyitem_Section_Footer_First];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Footer_Even]           = [AscDFH.historyitem_Section_Footer_Even];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Footer_Default]        = [AscDFH.historyitem_Section_Footer_Default];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_TitlePage]             = [AscDFH.historyitem_Section_TitlePage];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageMargins_Header]    = [AscDFH.historyitem_Section_PageMargins_Header];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageMargins_Footer]    = [AscDFH.historyitem_Section_PageMargins_Footer];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageNumType_Start]     = [AscDFH.historyitem_Section_PageNumType_Start];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Columns_EqualWidth]    = [AscDFH.historyitem_Section_Columns_EqualWidth];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Columns_Space]         = [AscDFH.historyitem_Section_Columns_Space];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Columns_Num]           = [AscDFH.historyitem_Section_Columns_Num];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Columns_Sep]           = [AscDFH.historyitem_Section_Columns_Sep];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Columns_Col]           = [AscDFH.historyitem_Section_Columns_Col, AscDFH.historyitem_Section_Columns_SetCols];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Columns_SetCols]       = [AscDFH.historyitem_Section_Columns_Col, AscDFH.historyitem_Section_Columns_SetCols];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Footnote_Pos]          = [AscDFH.historyitem_Section_Footnote_Pos];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Footnote_NumStart]     = [AscDFH.historyitem_Section_Footnote_NumStart];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Footnote_NumRestart]   = [AscDFH.historyitem_Section_Footnote_NumRestart];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Footnote_NumFormat]    = [AscDFH.historyitem_Section_Footnote_NumFormat];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageMargins_Gutter]    = [AscDFH.historyitem_Section_PageMargins_Gutter];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Gutter_RTL]            = [AscDFH.historyitem_Section_Gutter_RTL];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Endnote_Pos]           = [AscDFH.historyitem_Section_Endnote_Pos];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Endnote_NumStart]      = [AscDFH.historyitem_Section_Endnote_NumStart];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Endnote_NumRestart]    = [AscDFH.historyitem_Section_Endnote_NumRestart];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_Endnote_NumFormat]     = [AscDFH.historyitem_Section_Endnote_NumFormat];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_LnNumType]             = [AscDFH.historyitem_Section_LnNumType];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageNumType_Format]    = [AscDFH.historyitem_Section_PageNumType_Format];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageNumType_ChapStyle] = [AscDFH.historyitem_Section_PageNumType_ChapStyle];
	AscDFH.changesRelationMap[AscDFH.historyitem_Section_PageNumType_ChapSep]   = [AscDFH.historyitem_Section_PageNumType_ChapSep];

	//------------------------------------------------------------------------------------------------------------------
	
	
	/**
	 * Базовый класс для изменения колонтитулов
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesSectionBaseHeaderFooter(Class, Old, New)
	{
		AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
	}
	
	CChangesSectionBaseHeaderFooter.prototype                = Object.create(AscDFH.CChangesBaseProperty.prototype);
	CChangesSectionBaseHeaderFooter.prototype.constructor    = CChangesSectionBaseHeaderFooter;
	CChangesSectionBaseHeaderFooter.prototype.WriteToBinary  = function(Writer)
	{
		// Long : Flags
		// 1bit : is New null?
		// 2bit : is Old null?
		// 1bit == 0: String : New id
		// 2bit == 0: String : Old id
		
		var nFlags = 0;
		if (null === this.New)
			nFlags |= 1;
		if (null === this.Old)
			nFlags |= 2;
		
		Writer.WriteLong(nFlags);
		if (null !== this.New)
			Writer.WriteString2(this.New.Get_Id());
		if (null !== this.Old)
			Writer.WriteString2(this.Old.Get_Id());
	};
	CChangesSectionBaseHeaderFooter.prototype.ReadFromBinary = function(Reader)
	{
		// Long : Flags
		// 1bit : is New null?
		// 2bit : is Old null?
		// 1bit == 0: String : New id
		// 2bit == 0: String : Old id
		
		var nFlags = Reader.GetLong();
		
		if (nFlags & 1)
			this.New = null;
		else
			this.New = AscCommon.g_oTableId.Get_ById(Reader.GetString2());
		
		if (nFlags & 2)
			this.Old = null;
		else
			this.Old = AscCommon.g_oTableId.Get_ById(Reader.GetString2());
	};
	
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseByteValue}
	 */
	function CChangesSectionPageOrient(Class, Old, New)
	{
		AscDFH.CChangesBaseByteValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageOrient,
		AscDFH.CChangesBaseByteValue,
		AscDFH.historyitem_Section_PageSize_Orient,
		function(Value)
		{
			this.Class.PageSize.Orient = Value;
		}
	);
	AscDFH.CChangesSectionPageOrient = CChangesSectionPageOrient;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesSectionPageSize(Class, Old, New)
	{
		AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageSize,
		AscDFH.CChangesBaseProperty,
		AscDFH.historyitem_Section_PageSize_Size,
		function(Value)
		{
			this.Class.PageSize.W = Value.W;
			this.Class.PageSize.H = Value.H;
		}
	);
	CChangesSectionPageSize.prototype.WriteToBinary = function(Writer)
	{
		Writer.WriteDouble(this.New.W);
		Writer.WriteDouble(this.New.H);
		Writer.WriteDouble(this.Old.W);
		Writer.WriteDouble(this.Old.H);
	};
	CChangesSectionPageSize.prototype.ReadFromBinary = function(Reader)
	{
		this.New = {
			W : Reader.GetDouble(),
			H : Reader.GetDouble()
		};
		
		this.Old = {
			W : Reader.GetDouble(),
			H : Reader.GetDouble()
		};
	};
	AscDFH.CChangesSectionPageSize = CChangesSectionPageSize;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesSectionPageMargins(Class, Old, New)
	{
		AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageMargins,
		AscDFH.CChangesBaseProperty,
		AscDFH.historyitem_Section_PageMargins,
		function(Value)
		{
			this.Class.PageMargins.Left   = Value.L;
			this.Class.PageMargins.Top    = Value.T;
			this.Class.PageMargins.Right  = Value.R;
			this.Class.PageMargins.Bottom = Value.B;
		}
	);
	CChangesSectionPageMargins.prototype.WriteToBinary = function(Writer)
	{
		Writer.WriteDouble(this.New.L);
		Writer.WriteDouble(this.New.T);
		Writer.WriteDouble(this.New.R);
		Writer.WriteDouble(this.New.B);
		
		Writer.WriteDouble(this.Old.L);
		Writer.WriteDouble(this.Old.T);
		Writer.WriteDouble(this.Old.R);
		Writer.WriteDouble(this.Old.B);
	};
	CChangesSectionPageMargins.prototype.ReadFromBinary = function(Reader)
	{
		this.New = {
			L : Reader.GetDouble(),
			T : Reader.GetDouble(),
			R : Reader.GetDouble(),
			B : Reader.GetDouble()
		};
		
		this.Old = {
			L : Reader.GetDouble(),
			T : Reader.GetDouble(),
			R : Reader.GetDouble(),
			B : Reader.GetDouble()
		};
	};
	AscDFH.CChangesSectionPageMargins = CChangesSectionPageMargins;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseByteValue}
	 */
	function CChangesSectionType(Class, Old, New)
	{
		AscDFH.CChangesBaseByteValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionType,
		AscDFH.CChangesBaseByteValue,
		AscDFH.historyitem_Section_Type,
		function(Value)
		{
			this.Class.Type = Value;
		}
	);
	AscDFH.CChangesSectionType = CChangesSectionType;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseObjectValue}
	 */
	function CChangesSectionBordersLeft(Class, Old, New)
	{
		AscDFH.CChangesBaseObjectValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyObjectChange(
		CChangesSectionBordersLeft,
		AscDFH.CChangesBaseObjectValue,
		AscDFH.historyitem_Section_Borders_Left,
		function()
		{
			return new CDocumentBorder();
		},
		function(Value)
		{
			this.Class.Borders.Left = Value;
		}
	);
	AscDFH.CChangesSectionBordersLeft = CChangesSectionBordersLeft;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseObjectValue}
	 */
	function CChangesSectionBordersTop(Class, Old, New)
	{
		AscDFH.CChangesBaseObjectValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyObjectChange(
		CChangesSectionBordersTop,
		AscDFH.CChangesBaseObjectValue,
		AscDFH.historyitem_Section_Borders_Top,
		function()
		{
			return new CDocumentBorder();
		},
		function(Value)
		{
			this.Class.Borders.Top = Value;
		}
	);
	AscDFH.CChangesSectionBordersTop = CChangesSectionBordersTop;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseObjectValue}
	 */
	function CChangesSectionBordersRight(Class, Old, New)
	{
		AscDFH.CChangesBaseObjectValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyObjectChange(
		CChangesSectionBordersRight,
		AscDFH.CChangesBaseObjectValue,
		AscDFH.historyitem_Section_Borders_Right,
		function()
		{
			return new CDocumentBorder();
		},
		function(Value)
		{
			this.Class.Borders.Right = Value;
		}
	);

	AscDFH.CChangesSectionBordersRight = CChangesSectionBordersRight;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseObjectValue}
	 */
	function CChangesSectionBordersBottom(Class, Old, New)
	{
		AscDFH.CChangesBaseObjectValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyObjectChange(
		CChangesSectionBordersBottom,
		AscDFH.CChangesBaseObjectValue,
		AscDFH.historyitem_Section_Borders_Bottom,
		function()
		{
			return new CDocumentBorder();
		},
		function(Value)
		{
			this.Class.Borders.Bottom = Value;
		}
	);
	AscDFH.CChangesSectionBordersBottom = CChangesSectionBordersBottom;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseByteValue}
	 */
	function CChangesSectionBordersDisplay(Class, Old, New)
	{
		AscDFH.CChangesBaseByteValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionBordersDisplay,
		AscDFH.CChangesBaseByteValue,
		AscDFH.historyitem_Section_Borders_Display,
		function(Value)
		{
			this.Class.Borders.Display = Value;
		}
	);
	AscDFH.CChangesSectionBordersDisplay = CChangesSectionBordersDisplay;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseByteValue}
	 */
	function CChangesSectionBordersOffsetFrom(Class, Old, New)
	{
		AscDFH.CChangesBaseByteValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionBordersOffsetFrom,
		AscDFH.CChangesBaseByteValue,
		AscDFH.historyitem_Section_Borders_OffsetFrom,
		function(Value)
		{
			this.Class.Borders.OffsetFrom = Value;
		}
	);
	AscDFH.CChangesSectionBordersOffsetFrom = CChangesSectionBordersOffsetFrom;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseByteValue}
	 */
	function CChangesSectionBordersZOrder(Class, Old, New)
	{
		AscDFH.CChangesBaseByteValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionBordersZOrder,
		AscDFH.CChangesBaseByteValue,
		AscDFH.historyitem_Section_Borders_ZOrder,
		function(Value)
		{
			this.Class.Borders.ZOrder = Value;
		}
	);
	AscDFH.CChangesSectionBordersZOrder = CChangesSectionBordersZOrder;
	
	/**
	 * @constructor
	 * @extends {CChangesSectionBaseHeaderFooter}
	 */
	function CChangesSectionHeaderFirst(Class, Old, New)
	{
		CChangesSectionBaseHeaderFooter.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionHeaderFirst,
		CChangesSectionBaseHeaderFooter,
		AscDFH.historyitem_Section_Header_First,
		function(Value)
		{
			this.Class.HeaderFirst = Value;
		}
	);
	AscDFH.CChangesSectionHeaderFirst = CChangesSectionHeaderFirst;
	
	/**
	 * @constructor
	 * @extends {CChangesSectionBaseHeaderFooter}
	 */
	function CChangesSectionHeaderEven(Class, Old, New)
	{
		CChangesSectionBaseHeaderFooter.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionHeaderEven,
		CChangesSectionBaseHeaderFooter,
		AscDFH.historyitem_Section_Header_Even,
		function(Value)
		{
			this.Class.HeaderEven = Value;
		}
	);
	AscDFH.CChangesSectionHeaderEven = CChangesSectionHeaderEven;
	
	/**
	 * @constructor
	 * @extends {CChangesSectionBaseHeaderFooter}
	 */
	function CChangesSectionHeaderDefault(Class, Old, New)
	{
		CChangesSectionBaseHeaderFooter.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionHeaderDefault,
		CChangesSectionBaseHeaderFooter,
		AscDFH.historyitem_Section_Header_Default,
		function(Value)
		{
			this.Class.HeaderDefault = Value;
		}
	);
	AscDFH.CChangesSectionHeaderDefault = CChangesSectionHeaderDefault;
	
	/**
	 * @constructor
	 * @extends {CChangesSectionBaseHeaderFooter}
	 */
	function CChangesSectionFooterFirst(Class, Old, New)
	{
		CChangesSectionBaseHeaderFooter.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionFooterFirst,
		CChangesSectionBaseHeaderFooter,
		AscDFH.historyitem_Section_Footer_First,
		function(Value)
		{
			this.Class.FooterFirst = Value;
		}
	);
	AscDFH.CChangesSectionFooterFirst = CChangesSectionFooterFirst;
	
	/**
	 * @constructor
	 * @extends {CChangesSectionBaseHeaderFooter}
	 */
	function CChangesSectionFooterEven(Class, Old, New)
	{
		CChangesSectionBaseHeaderFooter.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionFooterEven,
		CChangesSectionBaseHeaderFooter,
		AscDFH.historyitem_Section_Footer_Even,
		function(Value)
		{
			this.Class.FooterEven = Value;
		}
	);
	AscDFH.CChangesSectionFooterEven = CChangesSectionFooterEven;
	
	/**
	 * @constructor
	 * @extends {CChangesSectionBaseHeaderFooter}
	 */
	function CChangesSectionFooterDefault(Class, Old, New)
	{
		CChangesSectionBaseHeaderFooter.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionFooterDefault,
		CChangesSectionBaseHeaderFooter,
		AscDFH.historyitem_Section_Footer_Default,
		function(Value)
		{
			this.Class.FooterDefault = Value;
		}
	);
	AscDFH.CChangesSectionFooterDefault = CChangesSectionFooterDefault;
	
/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseBoolValue}
	 */
	function CChangesSectionTitlePage(Class, Old, New)
	{
		AscDFH.CChangesBaseBoolValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionTitlePage,
		AscDFH.CChangesBaseBoolValue,
		AscDFH.historyitem_Section_TitlePage,
		function(Value)
		{
			this.Class.TitlePage = Value;
		}
	);
	AscDFH.CChangesSectionTitlePage = CChangesSectionTitlePage;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseDoubleValue}
	 */
	function CChangesSectionPageMarginsHeader(Class, Old, New)
	{
		AscDFH.CChangesBaseDoubleValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageMarginsHeader,
		AscDFH.CChangesBaseDoubleValue,
		AscDFH.historyitem_Section_PageMargins_Header,
		function(Value)
		{
			this.Class.PageMargins.Header = Value;
		}
	);
	AscDFH.CChangesSectionPageMarginsHeader = CChangesSectionPageMarginsHeader;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseDoubleValue}
	 */
	function CChangesSectionPageMarginsFooter(Class, Old, New)
	{
		AscDFH.CChangesBaseDoubleValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageMarginsFooter,
		AscDFH.CChangesBaseDoubleValue,
		AscDFH.historyitem_Section_PageMargins_Footer,
		function(Value)
		{
			this.Class.PageMargins.Footer = Value;
		}
	);
	AscDFH.CChangesSectionPageMarginsFooter = CChangesSectionPageMarginsFooter;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongValue}
	 */
	function CChangesSectionPageNumTypeStart(Class, Old, New)
	{
		AscDFH.CChangesBaseLongValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageNumTypeStart,
		AscDFH.CChangesBaseLongValue,
		AscDFH.historyitem_Section_PageNumType_Start,
		function(Value)
		{
			this.Class.PageNumType.Start = Value;
		}
	);
	AscDFH.CChangesSectionPageNumTypeStart = CChangesSectionPageNumTypeStart;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseBoolValue}
	 */
	function CChangesSectionColumnsEqualWidth(Class, Old, New)
	{
		AscDFH.CChangesBaseBoolValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionColumnsEqualWidth,
		AscDFH.CChangesBaseBoolValue,
		AscDFH.historyitem_Section_Columns_EqualWidth,
		function(Value)
		{
			this.Class.Columns.EqualWidth = Value;
		}
	);
	AscDFH.CChangesSectionColumnsEqualWidth = CChangesSectionColumnsEqualWidth;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseDoubleValue}
	 */
	function CChangesSectionColumnsSpace(Class, Old, New)
	{
		AscDFH.CChangesBaseDoubleValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionColumnsSpace,
		AscDFH.CChangesBaseDoubleValue,
		AscDFH.historyitem_Section_Columns_Space,
		function(Value)
		{
			this.Class.Columns.Space = Value;
		}
	);
	AscDFH.CChangesSectionColumnsSpace = CChangesSectionColumnsSpace;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongValue}
	 */
	function CChangesSectionColumnsNum(Class, Old, New)
	{
		AscDFH.CChangesBaseLongValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionColumnsNum,
		AscDFH.CChangesBaseLongValue,
		AscDFH.historyitem_Section_Columns_Num,
		function(Value)
		{
			this.Class.Columns.Num = Value;
		}
	);
	AscDFH.CChangesSectionColumnsNum = CChangesSectionColumnsNum;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseBoolValue}
	 */
	function CChangesSectionColumnsSep(Class, Old, New)
	{
		AscDFH.CChangesBaseBoolValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionColumnsSep,
		AscDFH.CChangesBaseBoolValue,
		AscDFH.historyitem_Section_Columns_Sep,
		function(Value)
		{
			this.Class.Columns.Sep = Value;
		}
	);
	AscDFH.CChangesSectionColumnsSep = CChangesSectionColumnsSep;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesSectionColumnsCol(Class, Old, New, Index)
	{
		AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
		
		this.Index = Index;
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionColumnsCol,
		AscDFH.CChangesBaseProperty,
		AscDFH.historyitem_Section_Columns_Col,
		function(Value)
		{
			this.Class.Columns.Cols[this.Index] = Value;
		}
	);
	CChangesSectionColumnsCol.prototype.WriteToBinary = function(Writer)
	{
		// Long : ColumnIndex
		// Long : flags
		// 1bit : is new undefined?
		// 2bit : is old undefined?
		// 1bit == 0 : AscWord.SectionColumn : New
		// 2bit == 0 : AscWord.SectionColumn : Old
		
		Writer.WriteLong(this.Index);
		
		var nFlags = 0;
		if (undefined === this.New)
			nFlags |= 1;
		
		if (undefined === this.Old)
			nFlags |= 2;
		
		Writer.WriteLong(nFlags);
		
		if (undefined !== this.New)
			this.New.Write_ToBinary(Writer);
		
		if (undefined !== this.Old)
			this.Old.Write_ToBinary(Writer);
	};
	CChangesSectionColumnsCol.prototype.ReadFromBinary = function(Reader)
	{
		// Long : ColumnIndex
		// Long : flags
		// 1bit : is new undefined?
		// 2bit : is old undefined?
		// 1bit == 0 : AscWord.SectionColumn : New
		// 2bit == 0 : AscWord.SectionColumn : Old
		
		this.Index = Reader.GetLong();
		
		var nFlags = Reader.GetLong();
		if (nFlags & 1)
		{
			this.New = undefined;
		}
		else
		{
			this.New = new AscWord.SectionColumn();
			this.New.Read_FromBinary(Reader);
		}
		
		if (nFlags & 2)
		{
			this.Old = undefined;
		}
		else
		{
			this.Old = new AscWord.SectionColumn();
			this.Old.Read_FromBinary(Reader);
		}
	};
	CChangesSectionColumnsCol.prototype.CreateReverseChange = function()
	{
		return new CChangesSectionColumnsCol(this.Class, this.New, this.Old, this.Index);
	};
	CChangesSectionColumnsCol.prototype.Merge = function(oChange)
	{
		if (this.Class !== oChange.Class)
			return true;
		
		if (this.Type === oChange.Type)
		{
			if (this.Index !== oChange.Index)
				return true;
			else
				return false;
		}
		else if (AscDFH.historyitem_Section_Columns_SetCols === oChange.Type)
		{
			return false
		}
		
		return true;
	};
	AscDFH.CChangesSectionColumnsCol = CChangesSectionColumnsCol;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseProperty}
	 */
	function CChangesSectionColumnsSetCols(Class, Old, New)
	{
		AscDFH.CChangesBaseProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionColumnsSetCols,
		AscDFH.CChangesBaseProperty,
		AscDFH.historyitem_Section_Columns_SetCols,
		function(Value)
		{
			this.Class.Columns.Cols = Value;
		}
	);
	CChangesSectionColumnsSetCols.prototype.WriteToBinary = function(Writer)
	{
		// Long : Count of new column
		// Array of AscWord.SectionColumn: new columns
		// Long : Count of old column
		// Array of AscWord.SectionColumn: old columns
		
		
		var nCount = this.New.length;
		Writer.WriteLong(nCount);
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			this.New[nIndex].Write_ToBinary(Writer);
		}
		
		nCount = this.Old.length;
		Writer.WriteLong(nCount);
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			this.Old[nIndex].Write_ToBinary(Writer);
		}
	};
	CChangesSectionColumnsSetCols.prototype.ReadFromBinary = function(Reader)
	{
		// Long : Count of new column
		// Array of AscWord.SectionColumn: new columns
		// Long : Count of old column
		// Array of AscWord.SectionColumn: old columns
		
		var nCount = Reader.GetLong();
		this.New   = [];
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			this.New[nIndex] = new AscWord.SectionColumn();
			this.New[nIndex].Read_FromBinary(Reader);
		}
		
		nCount   = Reader.GetLong();
		this.Old = [];
		for (var nIndex = 0; nIndex < nCount; ++nIndex)
		{
			this.Old[nIndex] = new AscWord.SectionColumn();
			this.Old[nIndex].Read_FromBinary(Reader);
		}
	};
	CChangesSectionColumnsSetCols.prototype.Merge = function(oChange)
	{
		if (this.Class !== oChange.Class)
			return true;
		
		if (this.Type === oChange.Type)
			return false;
		
		if (AscDFH.historyitem_Section_Columns_Col === oChange.Type)
		{
			if (!this.New)
				this.New = [];
			
			this.New[oChange.Index] = oChange.New;
		}
		
		return true;
	};
	AscDFH.CChangesSectionColumnsSetCols = CChangesSectionColumnsSetCols;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionFootnotePos(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionFootnotePos,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_Footnote_Pos,
		function(Value)
		{
			this.Class.FootnotePr.Pos = Value;
		}
	);
	AscDFH.CChangesSectionFootnotePos = CChangesSectionFootnotePos;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionFootnoteNumStart(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionFootnoteNumStart,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_Footnote_NumStart,
		function(Value)
		{
			this.Class.FootnotePr.NumStart = Value;
		}
	);
	AscDFH.CChangesSectionFootnoteNumStart = CChangesSectionFootnoteNumStart;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionFootnoteNumRestart(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionFootnoteNumRestart,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_Footnote_NumRestart,
		function(Value)
		{
			this.Class.FootnotePr.NumRestart = Value;
		}
	);
	AscDFH.CChangesSectionFootnoteNumRestart = CChangesSectionFootnoteNumRestart;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionFootnoteNumFormat(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionFootnoteNumFormat,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_Footnote_NumFormat,
		function(Value)
		{
			this.Class.FootnotePr.NumFormat = Value;
		}
	);
	AscDFH.CChangesSectionFootnoteNumFormat = CChangesSectionFootnoteNumFormat;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseDoubleValue}
	 */
	function CChangesSectionPageMarginsGutter(Class, Old, New)
	{
		AscDFH.CChangesBaseDoubleValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageMarginsGutter,
		AscDFH.CChangesBaseDoubleValue,
		AscDFH.historyitem_Section_PageMargins_Gutter,
		function(Value)
		{
			this.Class.PageMargins.Gutter = Value;
		}
	);
	AscDFH.CChangesSectionPageMarginsGutter = CChangesSectionPageMarginsGutter;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseBoolValue}
	 */
	function CChangesSectionGutterRTL(Class, Old, New)
	{
		AscDFH.CChangesBaseBoolValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionGutterRTL,
		AscDFH.CChangesBaseBoolValue,
		AscDFH.historyitem_Section_Gutter_RTL,
		function(Value)
		{
			this.Class.GutterRTL = Value;
		}
	);
	AscDFH.CChangesSectionGutterRTL = CChangesSectionGutterRTL;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionEndnotePos(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionEndnotePos,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_Endnote_Pos,
		function(Value)
		{
			this.Class.EndnotePr.Pos = Value;
		}
	);
	AscDFH.CChangesSectionEndnotePos = CChangesSectionEndnotePos;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionEndnoteNumStart(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionEndnoteNumStart,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_Endnote_NumStart,
		function(Value)
		{
			this.Class.EndnotePr.NumStart = Value;
		}
	);
	AscDFH.CChangesSectionEndnoteNumStart = CChangesSectionEndnoteNumStart;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionEndnoteNumRestart(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionEndnoteNumRestart,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_Endnote_NumRestart,
		function(Value)
		{
			this.Class.EndnotePr.NumRestart = Value;
		}
	);
	AscDFH.CChangesSectionEndnoteNumRestart = CChangesSectionEndnoteNumRestart;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionEndnoteNumFormat(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionEndnoteNumFormat,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_Endnote_NumFormat,
		function(Value)
		{
			this.Class.EndnotePr.NumFormat = Value;
		}
	);
	AscDFH.CChangesSectionEndnoteNumFormat = CChangesSectionEndnoteNumFormat;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseObjectProperty}
	 */
	function CChangesSectionLnNumType(Class, Old, New)
	{
		AscDFH.CChangesBaseObjectProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyObjectChange(
		CChangesSectionLnNumType,
		AscDFH.CChangesBaseObjectProperty,
		AscDFH.historyitem_Section_LnNumType,
		function()
		{
			return new AscWord.SectionLnNumType();
		},
		function(Value)
		{
			this.Class.LnNumType = Value;
		}
	);
	AscDFH.CChangesSectionLnNumType = CChangesSectionLnNumType;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongValue}
	 */
	function CChangesSectionPageNumTypeFormat(Class, Old, New)
	{
		AscDFH.CChangesBaseLongValue.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageNumTypeFormat,
		AscDFH.CChangesBaseLongValue,
		AscDFH.historyitem_Section_PageNumType_Format,
		function(Value)
		{
			this.Class.PageNumType.Format = Value;
		}
	);
	AscDFH.CChangesSectionPageNumTypeFormat = CChangesSectionPageNumTypeFormat;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseLongProperty}
	 */
	function CChangesSectionPageNumTypeChapStyle(Class, Old, New)
	{
		AscDFH.CChangesBaseLongProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageNumTypeChapStyle,
		AscDFH.CChangesBaseLongProperty,
		AscDFH.historyitem_Section_PageNumType_ChapStyle,
		function(Value)
		{
			this.Class.PageNumType.ChapStyle = Value;
		}
	);
	AscDFH.CChangesSectionPageNumTypeChapStyle = CChangesSectionPageNumTypeChapStyle;
	
	/**
	 * @constructor
	 * @extends {AscDFH.CChangesBaseByteProperty}
	 */
	function CChangesSectionPageNumTypeChapSep(Class, Old, New)
	{
		AscDFH.CChangesBaseByteProperty.call(this, Class, Old, New);
	}
	AscDFH.InheritPropertyChange(
		CChangesSectionPageNumTypeChapSep,
		AscDFH.CChangesBaseByteProperty,
		AscDFH.historyitem_Section_PageNumType_ChapSep,
		function(Value)
		{
			this.Class.PageNumType.ChapSep = Value;
		}
	);
	AscDFH.CChangesSectionPageNumTypeChapSep = CChangesSectionPageNumTypeChapSep;
})();
