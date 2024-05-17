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

// Сортируем объекты {X0, X1} по X0
function Sort_Ranges_X0(A,B)
{
    if ( !A.X0 || !B.X0  )
        return 0;

    if( A.X0 < B.X0 )
        return -1;
    else if( A.X0 > B.X0 )
        return 1;

    return 0;
}

// Сравниваем, вложено ли множество первых отрезков во второе
// (множество отрезков здесь - это дизъюнктное объединение отрезков)
function FlowObjects_CheckInjection(Range1, Range2)
{
    for ( var Index = 0; Index < Range1.length; Index++ )
    {
        var R1 = Range1[Index];
        var bInject = false;
        for ( var Index2 = 0; Index2 < Range2.length; Index2++ )
        {
            var R2 = Range2[Index2];

            if ( R1.X0 >= R2.X0 && R1.X0 <= R2.X1 && R1.X1 >= R2.X0 && R1.X1 <= R2.X1 )
                bInject = true;
        }

        if ( !bInject )
            return false;
    }

    return true;
}

// Сравниваем, совпали ли множества отрезков
function FlowObjects_CompareRanges(Range1, Range2)
{
    if ( Range1.length < Range2.length )
        return -1;
    else if ( Range1.length > Range2.length )
        return -1;

    for ( var Index = 0; Index < Range1.length; Index++ )
    {
        if ( Math.abs( Range1[Index].X0 - Range2[Index].X0 ) > 0.001 || Math.abs( Range1[Index].X1 - Range2[Index].X1 ) )
            return -1;
    }

    return 0;
}

/**
 * Базовый класс для плавающих объектов (таблиц и параграфов)
 * @constructor
 */
function CFlowBase()
{

}
CFlowBase.prototype.IsPointIn = function(X, Y)
{
	return (X <= this.X + this.W && X >= this.X && Y >= this.Y && Y <= this.Y + this.H);
};
CFlowBase.prototype.UpdateCursorType = function(X, Y, PageIndex)
{
};
CFlowBase.prototype.Get_Distance = function()
{
	var oDist = this.Distance;
	return new AscFormat.CDistance(
		AscFormat.getValOrDefault(oDist.L, AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT),
		AscFormat.getValOrDefault(oDist.T, 0),
		AscFormat.getValOrDefault(oDist.R, AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT),
		AscFormat.getValOrDefault(oDist.B, 0));
};
CFlowBase.prototype.getArrayWrapIntervals = function(x0, y0, x1, y1, Y0Sp, Y1Sp, LeftField, RightField, ret, bMathWrap)
{
	if (this.WrappingType === WRAPPING_TYPE_THROUGH || this.WrappingType === WRAPPING_TYPE_TIGHT)
	{
		y0 = Y0Sp;
		y1 = Y1Sp;
	}
	var top    = this.Y - AscFormat.getValOrDefault(this.Distance.T, 0);
	var bottom = this.Y + this.H + AscFormat.getValOrDefault(this.Distance.B, 0);
	if (y1 < top || y0 > bottom)
		return ret;

	var b_check = false, X0, X1, Y1, WrapType = (bMathWrap === true) ? WRAPPING_TYPE_SQUARE : this.WrappingType;
	switch (WrapType)
	{
		case WRAPPING_TYPE_NONE:
		{
			return ret;
		}
		case WRAPPING_TYPE_SQUARE:
		case WRAPPING_TYPE_THROUGH:
		case WRAPPING_TYPE_TIGHT:
		{
			X0      = this.X - AscFormat.getValOrDefault(this.Distance.L, AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT);
			X1      = this.X + this.W + AscFormat.getValOrDefault(this.Distance.R, AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT);
			Y1      = bottom;
			b_check = true;
			break;
		}
		case WRAPPING_TYPE_TOP_AND_BOTTOM:
		{
			var L = this.X - AscFormat.getValOrDefault(this.Distance.L, AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT);
			var R = this.X + this.W + AscFormat.getValOrDefault(this.Distance.R, AscFormat.DISTANCE_TO_TEXT_LEFTRIGHT);

			if (R < LeftField || L > RightField)
				return ret;

			X0 = x0;
			X1 = x1;
			Y1 = bottom;
			break;
		}
	}
	if (b_check)
	{
		var dx = this.WrappingType === WRAPPING_TYPE_SQUARE ? 6.35 : 3.175;
		if (X0 < LeftField + dx)
		{
			X0 = x0;
		}
		if (X1 > RightField - dx)
		{
			X1 = x1;
		}
	}
	ret.push({X0 : X0, X1 : X1, Y1 : Y1, typeLeft : this.WrappingType, typeRight : this.WrappingType, bFlowObject: true});
	return ret;
};
CFlowBase.prototype.GetElement = function()
{
	return this.Table;
};
CFlowBase.prototype.GetPage = function()
{
	return this.PageNum;
};
CFlowBase.prototype.IsFlowTable = function()
{
	return false;
};
CFlowBase.prototype.CheckDocumentContent = function(oDocContent)
{
	if (!this.Table.Parent)
		return false;

	return (oDocContent === this.Table.Parent || oDocContent.GetDocumentContentForRecalcInfo() === this.Table.Parent.GetDocumentContentForRecalcInfo());
};

function CFlowTable(Table, PageIndex)
{
    CFlowBase.call(this);

    this.Table          = Table;
    this.Id             = Table.Get_Id();
    this.PageNum        = Table.Get_StartPage_Absolute();
    this.PageController = PageIndex - Table.PageNum;
    this.Distance       = Table.Distance;
	this.WrappingType   = WRAPPING_TYPE_SQUARE;

	var oBounds = this.Table.GetPageBounds(this.PageController);

	this.X = oBounds.Left;
	this.Y = AscCommon.CorrectMMToTwips(oBounds.Top) + AscCommon.TwipsToMM(1); // Сдвигаемся на 1 twips вниз, чтобы не было пересечения с предыдущей строкой
	this.W = oBounds.Right - oBounds.Left;
	this.H = AscCommon.CorrectMMToTwips(oBounds.Bottom - oBounds.Top);
}

CFlowTable.prototype = Object.create(CFlowBase.prototype);
CFlowTable.prototype.constructor = CFlowTable;

CFlowTable.prototype.IsFlowTable = function()
{
    return true;
};

function CFlowParagraph(Paragraph, X, Y, W, H, Dx, Dy, StartIndex, FlowCount, Wrap)
{
    CFlowBase.call(this);

    this.Table     = Paragraph;
    this.Paragraph = Paragraph;
    this.Id        = Paragraph.Get_Id();

    this.PageNum   = Paragraph.PageNum + Paragraph.GetPagesCount() - 1;
    this.PageController = 0;

    this.StartIndex = StartIndex;
    this.FlowCount  = FlowCount;

    this.Distance =
    {
        T : Dy,
        B : Dy,
        L : Dx,
        R : Dx
    };

    this.X = X;
    this.Y = Y;
    this.W = W;
    this.H = H;

    this.WrappingType = WRAPPING_TYPE_SQUARE;

    switch (Wrap)
    {
        case undefined:
        case wrap_Around:
        case wrap_Auto:      this.WrappingType = WRAPPING_TYPE_SQUARE;         break;
        case wrap_None:      this.WrappingType = WRAPPING_TYPE_NONE;           break;
        case wrap_NotBeside: this.WrappingType = WRAPPING_TYPE_TOP_AND_BOTTOM; break;
        case wrap_Through:   this.WrappingType = WRAPPING_TYPE_THROUGH;        break;
        case wrap_Tight:     this.WrappingType = WRAPPING_TYPE_TIGHT;          break;
    }
}

CFlowParagraph.prototype = Object.create(CFlowBase.prototype);
CFlowParagraph.prototype.constructor = CFlowParagraph;

