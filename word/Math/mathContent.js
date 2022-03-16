/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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

// Import
var History = AscCommon.History;


/** @enum {number} */
var c_oAscMathType = {
    //----------------------------------------------------------------------------------------------------------------------
    Symbol_pm                               : 0x00000000,
    Symbol_infinity                         : 0x00000001,
    Symbol_equals                           : 0x00000002,
    Symbol_neq                              : 0x00000003,
    Symbol_about                            : 0x00000004,
    Symbol_times                            : 0x00000005,
    Symbol_div                              : 0x00000006,
    Symbol_factorial                        : 0x00000007,
    Symbol_propto                           : 0x00000008,
    Symbol_less                             : 0x00000009,
    Symbol_ll                               : 0x0000000a,
    Symbol_greater                          : 0x0000000b,
    Symbol_gg                               : 0x0000000c,
    Symbol_leq                              : 0x0000000d,
    Symbol_geq                              : 0x0000000e,
    Symbol_mp                               : 0x0000000f,
    Symbol_cong                             : 0x00000010,
    Symbol_approx                           : 0x00000011,
    Symbol_equiv                            : 0x00000012,
    Symbol_forall                           : 0x00000013,
    Symbol_additional                       : 0x00000014,
    Symbol_partial                          : 0x00000015,
    Symbol_sqrt                             : 0x00000016,
    Symbol_cbrt                             : 0x00000017,
    Symbol_qdrt                             : 0x00000018,
    Symbol_cup                              : 0x00000019,
    Symbol_cap                              : 0x0000001a,
    Symbol_emptyset                         : 0x0000001b,
    Symbol_percent                          : 0x0000001c,
    Symbol_degree                           : 0x0000001d,
    Symbol_fahrenheit                       : 0x0000001e,
    Symbol_celsius                          : 0x0000001f,
    Symbol_inc                              : 0x00000020,
    Symbol_nabla                            : 0x00000021,
    Symbol_exists                           : 0x00000022,
    Symbol_notexists                        : 0x00000023,
    Symbol_in                               : 0x00000024,
    Symbol_ni                               : 0x00000025,
    Symbol_leftarrow                        : 0x00000026,
    Symbol_uparrow                          : 0x00000027,
    Symbol_rightarrow                       : 0x00000028,
    Symbol_downarrow                        : 0x00000029,
    Symbol_leftrightarrow                   : 0x0000002a,
    Symbol_therefore                        : 0x0000002b,
    Symbol_plus                             : 0x0000002c,
    Symbol_minus                            : 0x0000002d,
    Symbol_not                              : 0x0000002e,
    Symbol_ast                              : 0x0000002f,
    Symbol_bullet                           : 0x00000030,
    Symbol_vdots                            : 0x00000031,
    Symbol_cdots                            : 0x00000032,
    Symbol_rddots                           : 0x00000033,
    Symbol_ddots                            : 0x00000034,
    Symbol_aleph                            : 0x00000035,
    Symbol_beth                             : 0x00000036,
    Symbol_QED                              : 0x00000037,
    Symbol_alpha                            : 0x00010000,
    Symbol_beta                             : 0x00010001,
    Symbol_gamma                            : 0x00010002,
    Symbol_delta                            : 0x00010003,
    Symbol_varepsilon                       : 0x00010004,
    Symbol_epsilon                          : 0x00010005,
    Symbol_zeta                             : 0x00010006,
    Symbol_eta                              : 0x00010007,
    Symbol_theta                            : 0x00010008,
    Symbol_vartheta                         : 0x00010009,
    Symbol_iota                             : 0x0001000a,
    Symbol_kappa                            : 0x0001000b,
    Symbol_lambda                           : 0x0001000c,
    Symbol_mu                               : 0x0001000d,
    Symbol_nu                               : 0x0001000e,
    Symbol_xsi                              : 0x0001000f,
    Symbol_o                                : 0x00010010,
    Symbol_pi                               : 0x00010011,
    Symbol_varpi                            : 0x00010012,
    Symbol_rho                              : 0x00010013,
    Symbol_varrho                           : 0x00010014,
    Symbol_sigma                            : 0x00010015,
    Symbol_varsigma                         : 0x00010016,
    Symbol_tau                              : 0x00010017,
    Symbol_upsilon                          : 0x00010018,
    Symbol_varphi                           : 0x00010019,
    Symbol_phi                              : 0x0001001a,
    Symbol_chi                              : 0x0001001b,
    Symbol_psi                              : 0x0001001c,
    Symbol_omega                            : 0x0001001d,
    Symbol_Alpha                            : 0x00020000,
    Symbol_Beta                             : 0x00020001,
    Symbol_Gamma                            : 0x00020002,
    Symbol_Delta                            : 0x00020003,
    Symbol_Epsilon                          : 0x00020004,
    Symbol_Zeta                             : 0x00020005,
    Symbol_Eta                              : 0x00020006,
    Symbol_Theta                            : 0x00020007,
    Symbol_Iota                             : 0x00020008,
    Symbol_Kappa                            : 0x00020009,
    Symbol_Lambda                           : 0x0002000a,
    Symbol_Mu                               : 0x0002000b,
    Symbol_Nu                               : 0x0002000c,
    Symbol_Xsi                              : 0x0002000d,
    Symbol_O                                : 0x0002000e,
    Symbol_Pi                               : 0x0002000f,
    Symbol_Rho                              : 0x00020010,
    Symbol_Sigma                            : 0x00020011,
    Symbol_Tau                              : 0x00020012,
    Symbol_Upsilon                          : 0x00020013,
    Symbol_Phi                              : 0x00020014,
    Symbol_Chi                              : 0x00020015,
    Symbol_Psi                              : 0x00020016,
    Symbol_Omega                            : 0x00020017,
    //----------------------------------------------------------------------------------------------------------------------
    FractionVertical                        : 0x01000000,
    FractionDiagonal                        : 0x01000001,
    FractionHorizontal                      : 0x01000002,
    FractionSmall                           : 0x01000003,
    FractionDifferential_1                  : 0x01010000,
    FractionDifferential_2                  : 0x01010001,
    FractionDifferential_3                  : 0x01010002,
    FractionDifferential_4                  : 0x01010003,
    FractionPi_2                            : 0x01010004,
    //----------------------------------------------------------------------------------------------------------------------
    ScriptSup                               : 0x02000000,
    ScriptSub                               : 0x02000001,
    ScriptSubSup                            : 0x02000002,
    ScriptSubSupLeft                        : 0x02000003,
    ScriptCustom_1                          : 0x02010000,
    ScriptCustom_2                          : 0x02010001,
    ScriptCustom_3                          : 0x02010002,
    ScriptCustom_4                          : 0x02010003,
    //----------------------------------------------------------------------------------------------------------------------
    RadicalSqrt                             : 0x03000000,
    RadicalRoot_n                           : 0x03000001,
    RadicalRoot_2                           : 0x03000002,
    RadicalRoot_3                           : 0x03000003,
    RadicalCustom_1                         : 0x03010000,
    RadicalCustom_2                         : 0x03010001,
    //----------------------------------------------------------------------------------------------------------------------
    Integral                                : 0x04000000,
    IntegralSubSup                          : 0x04000001,
    IntegralCenterSubSup                    : 0x04000002,
    IntegralDouble                          : 0x04000003,
    IntegralDoubleSubSup                    : 0x04000004,
    IntegralDoubleCenterSubSup              : 0x04000005,
    IntegralTriple                          : 0x04000006,
    IntegralTripleSubSup                    : 0x04000007,
    IntegralTripleCenterSubSup              : 0x04000008,
    IntegralOriented                        : 0x04010000,
    IntegralOrientedSubSup                  : 0x04010001,
    IntegralOrientedCenterSubSup            : 0x04010002,
    IntegralOrientedDouble                  : 0x04010003,
    IntegralOrientedDoubleSubSup            : 0x04010004,
    IntegralOrientedDoubleCenterSubSup      : 0x04010005,
    IntegralOrientedTriple                  : 0x04010006,
    IntegralOrientedTripleSubSup            : 0x04010007,
    IntegralOrientedTripleCenterSubSup      : 0x04010008,
    Integral_dx                             : 0x04020000,
    Integral_dy                             : 0x04020001,
    Integral_dtheta                         : 0x04020002,
    //----------------------------------------------------------------------------------------------------------------------
    LargeOperator_Sum                       : 0x05000000,
    LargeOperator_Sum_CenterSubSup          : 0x05000001,
    LargeOperator_Sum_SubSup                : 0x05000002,
    LargeOperator_Sum_CenterSub             : 0x05000003,
    LargeOperator_Sum_Sub                   : 0x05000004,
    LargeOperator_Prod                      : 0x05010000,
    LargeOperator_Prod_CenterSubSup         : 0x05010001,
    LargeOperator_Prod_SubSup               : 0x05010002,
    LargeOperator_Prod_CenterSub            : 0x05010003,
    LargeOperator_Prod_Sub                  : 0x05010004,
    LargeOperator_CoProd                    : 0x05010005,
    LargeOperator_CoProd_CenterSubSup       : 0x05010006,
    LargeOperator_CoProd_SubSup             : 0x05010007,
    LargeOperator_CoProd_CenterSub          : 0x05010008,
    LargeOperator_CoProd_Sub                : 0x05010009,
    LargeOperator_Union                     : 0x05020000,
    LargeOperator_Union_CenterSubSup        : 0x05020001,
    LargeOperator_Union_SubSup              : 0x05020002,
    LargeOperator_Union_CenterSub           : 0x05020003,
    LargeOperator_Union_Sub                 : 0x05020004,
    LargeOperator_Intersection              : 0x05020005,
    LargeOperator_Intersection_CenterSubSup : 0x05020006,
    LargeOperator_Intersection_SubSup       : 0x05020007,
    LargeOperator_Intersection_CenterSub    : 0x05020008,
    LargeOperator_Intersection_Sub          : 0x05020009,
    LargeOperator_Disjunction               : 0x05030000,
    LargeOperator_Disjunction_CenterSubSup  : 0x05030001,
    LargeOperator_Disjunction_SubSup        : 0x05030002,
    LargeOperator_Disjunction_CenterSub     : 0x05030003,
    LargeOperator_Disjunction_Sub           : 0x05030004,
    LargeOperator_Conjunction               : 0x05030005,
    LargeOperator_Conjunction_CenterSubSup  : 0x05030006,
    LargeOperator_Conjunction_SubSup        : 0x05030007,
    LargeOperator_Conjunction_CenterSub     : 0x05030008,
    LargeOperator_Conjunction_Sub           : 0x05030009,
    LargeOperator_Custom_1                  : 0x05040000,
    LargeOperator_Custom_2                  : 0x05040001,
    LargeOperator_Custom_3                  : 0x05040002,
    LargeOperator_Custom_4                  : 0x05040003,
    LargeOperator_Custom_5                  : 0x05040004,
    //----------------------------------------------------------------------------------------------------------------------
    Bracket_Round                           : 0x06000000,
    Bracket_Square                          : 0x06000001,
    Bracket_Curve                           : 0x06000002,
    Bracket_Angle                           : 0x06000003,
    Bracket_LowLim                          : 0x06000004,
    Bracket_UppLim                          : 0x06000005,
    Bracket_Line                            : 0x06000006,
    Bracket_LineDouble                      : 0x06000007,
    Bracket_Square_OpenOpen                 : 0x06000008,
    Bracket_Square_CloseClose               : 0x06000009,
    Bracket_Square_CloseOpen                : 0x0600000a,
    Bracket_SquareDouble                    : 0x0600000b,
    Bracket_Round_Delimiter_2               : 0x06010000,
    Bracket_Curve_Delimiter_2               : 0x06010001,
    Bracket_Angle_Delimiter_2               : 0x06010002,
    Bracket_Angle_Delimiter_3               : 0x06010003,
    Bracket_Round_OpenNone                  : 0x06020000,
    Bracket_Round_NoneOpen                  : 0x06020001,
    Bracket_Square_OpenNone                 : 0x06020002,
    Bracket_Square_NoneOpen                 : 0x06020003,
    Bracket_Curve_OpenNone                  : 0x06020004,
    Bracket_Curve_NoneOpen                  : 0x06020005,
    Bracket_Angle_OpenNone                  : 0x06020006,
    Bracket_Angle_NoneOpen                  : 0x06020007,
    Bracket_LowLim_OpenNone                 : 0x06020008,
    Bracket_LowLim_NoneNone                 : 0x06020009,
    Bracket_UppLim_OpenNone                 : 0x0602000a,
    Bracket_UppLim_NoneOpen                 : 0x0602000b,
    Bracket_Line_OpenNone                   : 0x0602000c,
    Bracket_Line_NoneOpen                   : 0x0602000d,
    Bracket_LineDouble_OpenNone             : 0x0602000e,
    Bracket_LineDouble_NoneOpen             : 0x0602000f,
    Bracket_SquareDouble_OpenNone           : 0x06020010,
    Bracket_SquareDouble_NoneOpen           : 0x06020011,
    Bracket_Custom_1                        : 0x06030000,
    Bracket_Custom_2                        : 0x06030001,
    Bracket_Custom_3                        : 0x06030002,
    Bracket_Custom_4                        : 0x06030003,
    Bracket_Custom_5                        : 0x06040000,
    Bracket_Custom_6                        : 0x06040001,
    Bracket_Custom_7                        : 0x06040002,
    //----------------------------------------------------------------------------------------------------------------------
    Function_Sin                            : 0x07000000,
    Function_Cos                            : 0x07000001,
    Function_Tan                            : 0x07000002,
    Function_Csc                            : 0x07000003,
    Function_Sec                            : 0x07000004,
    Function_Cot                            : 0x07000005,
    Function_1_Sin                          : 0x07010000,
    Function_1_Cos                          : 0x07010001,
    Function_1_Tan                          : 0x07010002,
    Function_1_Csc                          : 0x07010003,
    Function_1_Sec                          : 0x07010004,
    Function_1_Cot                          : 0x07010005,
    Function_Sinh                           : 0x07020000,
    Function_Cosh                           : 0x07020001,
    Function_Tanh                           : 0x07020002,
    Function_Csch                           : 0x07020003,
    Function_Sech                           : 0x07020004,
    Function_Coth                           : 0x07020005,
    Function_1_Sinh                         : 0x07030000,
    Function_1_Cosh                         : 0x07030001,
    Function_1_Tanh                         : 0x07030002,
    Function_1_Csch                         : 0x07030003,
    Function_1_Sech                         : 0x07030004,
    Function_1_Coth                         : 0x07030005,
    Function_Custom_1                       : 0x07040000,
    Function_Custom_2                       : 0x07040001,
    Function_Custom_3                       : 0x07040002,
    //----------------------------------------------------------------------------------------------------------------------
    Accent_Dot                              : 0x08000000,
    Accent_DDot                             : 0x08000001,
    Accent_DDDot                            : 0x08000002,
    Accent_Hat                              : 0x08000003,
    Accent_Check                            : 0x08000004,
    Accent_Accent                           : 0x08000005,
    Accent_Grave                            : 0x08000006,
    Accent_Smile                            : 0x08000007,
    Accent_Tilde                            : 0x08000008,
    Accent_Bar                              : 0x08000009,
    Accent_DoubleBar                        : 0x0800000a,
    Accent_CurveBracketTop                  : 0x0800000b,
    Accent_CurveBracketBot                  : 0x0800000c,
    Accent_GroupTop                         : 0x0800000d,
    Accent_GroupBot                         : 0x0800000e,
    Accent_ArrowL                           : 0x0800000f,
    Accent_ArrowR                           : 0x08000010,
    Accent_ArrowD                           : 0x08000011,
    Accent_HarpoonL                         : 0x08000012,
    Accent_HarpoonR                         : 0x08000013,
    Accent_BorderBox                        : 0x08010000,
    Accent_BorderBoxCustom                  : 0x08010001,
    Accent_BarTop                           : 0x08020000,
    Accent_BarBot                           : 0x08020001,
    Accent_Custom_1                         : 0x08030000,
    Accent_Custom_2                         : 0x08030001,
    Accent_Custom_3                         : 0x08030002,
    //----------------------------------------------------------------------------------------------------------------------
    LimitLog_LogBase                        : 0x09000000,
    LimitLog_Log                            : 0x09000001,
    LimitLog_Lim                            : 0x09000002,
    LimitLog_Min                            : 0x09000003,
    LimitLog_Max                            : 0x09000004,
    LimitLog_Ln                             : 0x09000005,
    LimitLog_Custom_1                       : 0x09010000,
    LimitLog_Custom_2                       : 0x09010001,
    //----------------------------------------------------------------------------------------------------------------------
    Operator_ColonEquals                    : 0x0a000000,
    Operator_EqualsEquals                   : 0x0a000001,
    Operator_PlusEquals                     : 0x0a000002,
    Operator_MinusEquals                    : 0x0a000003,
    Operator_Definition                     : 0x0a000004,
    Operator_UnitOfMeasure                  : 0x0a000005,
    Operator_DeltaEquals                    : 0x0a000006,
    Operator_ArrowL_Top                     : 0x0a010000,
    Operator_ArrowR_Top                     : 0x0a010001,
    Operator_ArrowL_Bot                     : 0x0a010002,
    Operator_ArrowR_Bot                     : 0x0a010003,
    Operator_DoubleArrowL_Top               : 0x0a010004,
    Operator_DoubleArrowR_Top               : 0x0a010005,
    Operator_DoubleArrowL_Bot               : 0x0a010006,
    Operator_DoubleArrowR_Bot               : 0x0a010007,
    Operator_ArrowD_Top                     : 0x0a010008,
    Operator_ArrowD_Bot                     : 0x0a010009,
    Operator_DoubleArrowD_Top               : 0x0a01000a,
    Operator_DoubleArrowD_Bot               : 0x0a01000b,
    Operator_Custom_1                       : 0x0a020000,
    Operator_Custom_2                       : 0x0a020001,
    //----------------------------------------------------------------------------------------------------------------------
    Matrix_1_2                              : 0x0b000000,
    Matrix_2_1                              : 0x0b000001,
    Matrix_1_3                              : 0x0b000002,
    Matrix_3_1                              : 0x0b000003,
    Matrix_2_2                              : 0x0b000004,
    Matrix_2_3                              : 0x0b000005,
    Matrix_3_2                              : 0x0b000006,
    Matrix_3_3                              : 0x0b000007,
    Matrix_Dots_Center                      : 0x0b010000,
    Matrix_Dots_Baseline                    : 0x0b010001,
    Matrix_Dots_Vertical                    : 0x0b010002,
    Matrix_Dots_Diagonal                    : 0x0b010003,
    Matrix_Identity_2                       : 0x0b020000,
    Matrix_Identity_2_NoZeros               : 0x0b020001,
    Matrix_Identity_3                       : 0x0b020002,
    Matrix_Identity_3_NoZeros               : 0x0b020003,
    Matrix_2_2_RoundBracket                 : 0x0b030000,
    Matrix_2_2_SquareBracket                : 0x0b030001,
    Matrix_2_2_LineBracket                  : 0x0b030002,
    Matrix_2_2_DLineBracket                 : 0x0b030003,
    Matrix_Flat_Round                       : 0x0b040000,
    Matrix_Flat_Square                      : 0x0b040001,
    //----------------------------------------------------------------------------------------------------------------------
    Default_Text                            : 0x0c000000
};

function CRPI()
{
    this.bDecreasedComp            = false;
    this.bInline                   = false;
    this.bChangeInline             = false;
    this.bNaryInline               = false; /*для CDegreeSupSub внутри N-арного оператора, этот флаг необходим, чтобы итераторы максимально близко друг к другу расположить*/
    this.bEqArray                  = false; /*для амперсанда*/
    this.bMathFunc                 = false;
    this.bRecalcCtrPrp             = false; // пересчет ctrPrp нужен, когда на Undo и тп изменился размер первого Run, а ctrPrp уже для мат объектов пересчитались
    this.bCorrect_ConvertFontSize  = false;
    this.bSmallFraction            = false;
}
CRPI.prototype.MergeMathInfo = function(MathInfo)
{
    this.bInline                   = MathInfo.bInline || (MathInfo.bInternalRanges == true && MathInfo.bStartRanges == false);
    this.bRecalcCtrPrp             = MathInfo.bRecalcCtrPrp;
    this.bChangeInline             = MathInfo.bChangeInline;
    this.bCorrect_ConvertFontSize  = MathInfo.bCorrect_ConvertFontSize;
};

function CMathPointInfo()
{
    this.x    = 0;
    this.y    = 0;

    this.bEven      = true;
    this.CurrPoint  = 0;

    this.InfoPoints = {};
}
CMathPointInfo.prototype.SetInfoPoints = function(InfoPoints)
{
    this.InfoPoints.GWidths       = InfoPoints.GWidths;
    this.InfoPoints.GPoints       = InfoPoints.GPoints;
    this.InfoPoints.ContentPoints = InfoPoints.ContentPoints.Widths;
    this.InfoPoints.GMaxDimWidths = InfoPoints.GMaxDimWidths;
};
CMathPointInfo.prototype.NextAlignRange = function()
{
    if(this.bEven)
        this.bEven = false;
    else
    {
        this.CurrPoint++;
        this.bEven = true;
    }
};
CMathPointInfo.prototype.GetAlign = function()
{
    var align = 0;

    if(this.bEven)
    {
        var alignEven, alignGeneral, alignOdd;

        var Len   = this.InfoPoints.ContentPoints.length,
            Point = this.InfoPoints.ContentPoints[this.CurrPoint];

        var GWidth = this.InfoPoints.GWidths[this.CurrPoint],
            GPoint = this.InfoPoints.GPoints[this.CurrPoint];

        if(this.CurrPoint == Len - 1 && Point.odd == -1) // то есть последняя точка четная, выравнивание по центру
        {
            var GMaxDimWidth = this.InfoPoints.GMaxDimWidths[this.CurrPoint];

            alignGeneral = (GMaxDimWidth - Point.even)/2;
            alignEven = 0;
        }
        else
        {
            alignGeneral = (GWidth - GPoint.even - GPoint.odd)/2;
            alignEven = GPoint.even - Point.even;
        }

        if(this.CurrPoint > 0)
        {
            var PrevGenPoint = this.InfoPoints.GPoints[this.CurrPoint-1],
                PrevGenWidth = this.InfoPoints.GWidths[this.CurrPoint-1],
                PrevPoint    = this.InfoPoints.ContentPoints[this.CurrPoint-1];

            var alignPrevGen = (PrevGenWidth - PrevGenPoint.even - PrevGenPoint.odd)/2;
            alignOdd = alignPrevGen +  PrevGenPoint.odd - PrevPoint.odd;
        }
        else
            alignOdd = 0;

        align = alignGeneral + alignEven + alignOdd;
    }

    return align;
};

function CInfoPoints()
{
    this.GWidths       = null;
    this.GPoints       = null;
    this.GMaxDimWidths = null;
    this.ContentPoints = new AmperWidths();
}
CInfoPoints.prototype.SetDefault = function()
{
    this.GWidths       = null;
    this.GPoints       = null;
    this.GMaxDimWidths = null;
    this.ContentPoints.SetDefault();
};

function CMathPosInfo()
{
    this.CurRange = -1;
    this.CurLine  = -1;

    this.DispositionOpers = null;
}

function CMathPosition()
{
    this.x  = 0;
    this.y  = 0;
}
CMathPosition.prototype.Set = function(Pos)
{
    this.x = Pos.x;
    this.y = Pos.y;
};

function AmperWidths()
{
    this.bEven     = true; // является ли текущая точка нечетной
    this.Widths    = [];
}
AmperWidths.prototype.UpdatePoint = function(value)
{
    var len = this.Widths.length;

    if(len == 0)
    {
        // дефолтное значение bEven true, для случая если первый элемент в контенте будет Ampersand
        var NewPoint = new CMathPoint();
        NewPoint.even = value;
        this.Widths.push(NewPoint);
        this.bEven = true;
    }
    else
    {
        if(this.bEven)
            this.Widths[len-1].even += value;
        else
            this.Widths[len-1].odd += value;
    }

};
AmperWidths.prototype.AddNewAlignRange = function()
{
    var len = this.Widths.length;

    if(!this.bEven || len == 0)
    {
        var NewPoint = new CMathPoint();
        NewPoint.even = 0;
        this.Widths.push(NewPoint);
    }

    if(this.bEven)
    {
        len = this.Widths.length;
        this.Widths[len-1].odd = 0;
    }


    this.bEven = !this.bEven;

};
AmperWidths.prototype.SetDefault = function()
{
    this.bEven         = true;
    this.Widths.length = 0;
};


function CGeneralObjectGaps(Left, Right)
{
    this.left  = Left;
    this.right = Right;
}

function CGaps(oSign, oEqual, oZeroOper, oLett)
{
    this.sign = oSign;
    this.equal = oEqual;
    this.zeroOper = oZeroOper;
    this.letters = oLett;
}

function CCoeffGaps()
{
    var LeftSign  = new CGaps(0.52, 0.26, 0, 0.52),
        RightSign = new CGaps(0.49, 0, 0, 0.49);

    this.Sign = new CGeneralObjectGaps(LeftSign, RightSign);

    var LeftMult  = new CGaps(0, 0, 0, 0.46),
        RightMult = new CGaps(0, 0, 0, 0.49);

    this.Mult = new CGeneralObjectGaps(LeftMult, RightMult);

    var LeftEqual  = new CGaps(0, 0, 0, 0.7),
        RightEqual = new CGaps(0, 0, 0, 0.5);

    this.Equal = new CGeneralObjectGaps(LeftEqual, RightEqual);

    var LeftDefault  = new CGaps(0, 0, 0, 0),
        RightDefault = new CGaps(0, 0, 0, 0);

    this.Default = new CGeneralObjectGaps(LeftDefault, RightDefault);
}
CCoeffGaps.prototype =
{
    getCoeff: function(codeCurr, codeLR , direct) // obj - либо codeChar, либо мат объект
    {
        var operator = null;

        if(this.checkEqualSign(codeCurr))
            operator = this.Equal;
        else if(this.checkOperSign(codeCurr))
            operator = this.Sign;
        else if(codeCurr == 0x2A)
            operator = this.Mult;
        else
            operator = this.Default;

        var part = direct == -1 ? operator.left : operator.right;

        var coeff = 0;
        if(codeLR == -1) // мат объект
            coeff = part.letters;
        else if(this.checkOperSign(codeLR))
            coeff = part.sign;
        else if(this.checkEqualSign(codeLR))
            coeff = part.equal;
        else if(this.checkZeroSign(codeLR, direct))
            coeff = part.zeroOper;
        else
            coeff = part.letters;

        return coeff;
    },
    checkOperSign: function(code) // "+", "-", "±", "∓", "×", "÷"
    {
        var PLUS            = 0x2B,
            MINUS           = 0x2D,
            PLUS_MINUS      = 0xB1,
            MINUS_PLUS      = 0x2213,
            MULTIPLICATION  = 0xD7,
            DIVISION        = 0xF7;

        return code == PLUS || code == MINUS || code == PLUS_MINUS || code == MINUS_PLUS || code == MULTIPLICATION || code == DIVISION;
    },
    checkEqualSign: function(code)
    {
        var COMPARE       = code == 0x3C || code == 0x3E; // LESS, GREATER
        var ARROWS        = (code >= 0x2190 && code <= 0x21B3) || (code == 0x21B6) || (code == 0x21B7) || (code >= 0x21BA && code <= 0x21E9) || (code >=0x21F4 && code <= 0x21FF);
        var INTERSECTION  = code >= 0x2223 && code <= 0x222A;
        var EQUALS        = code == 0x3D || (code >= 0x2234 && code <= 0x22BD) || (code >= 0x22C4 && code <= 0x22FF);
        var ARR_FISHES    = (code >= 0x27DA && code <= 0x27E5) || (code >= 0x27EC && code <= 0x297F);
        var TRIANGLE_SYMB = code >= 0x29CE && code <= 0x29D7;
        var OTH_SYMB      = code == 0x29DF || (code >= 0x29E1 && code <= 0x29E7) || (code >= 0x29F4 && code <= 0x29F8) || (code >= 0x2A22 && code <= 0x2AF0) || (code >= 0x2AF2 && code <= 0x2AFB) || code == 0x2AFD || code == 0x2AFE;


        return COMPARE || ARROWS || INTERSECTION || EQUALS || ARR_FISHES || TRIANGLE_SYMB || OTH_SYMB;
    },
    checkZeroSign: function(code, direct) // "*", "/", "\"
    {
        var MULT     = 0x2A,
            DIVISION = 0x2F,
            B_SLASH  = 0x5C;

        var bOper = code == MULT || code == DIVISION || code == B_SLASH;
        var bLeftBracket = direct == -1 && (code == 0x28 || code == 0x5B || code == 0x7B);
        var bRightBracket = direct == 1 && (code == 0x29 || code == 0x5D || code == 0x7D);


        return bOper || bLeftBracket || bRightBracket;
    }
};

var COEFF_GAPS = new CCoeffGaps();

function CMathArgSize()
{
    this.value       = undefined;
}
CMathArgSize.prototype =
{
    Decrease: function()
    {
        if(this.value == undefined)
            this.value = 0;

        if( this.value > -2 )
            this.value--;

        return this.value;
    },
    Increase: function()
    {
        if(this.value == undefined)
            this.value = 0;

        if(this.value < 2)
            this.value++;

        return this.value;
    },
    Set: function(ArgSize)
    {
        this.value = ArgSize.value;
    },
    GetValue: function()
    {
        return this.value;
    },
    SetValue: function(val)
    {
        if(val === null || val === undefined)
            this.value = undefined;
        else if(val < - 2)
            this.value = -2;
        else if(val > 2)
            this.value = 2;
        else
            this.value = val;
    },
    Copy: function()
    {
        var ArgSize = new CMathArgSize();
        ArgSize.value = this.value;

        return ArgSize;
    },
    Merge: function(ArgSize)
    {
        if(this.value == undefined)
            this.value = 0;

        if(ArgSize.value == undefined)
            ArgSize.value = 0;

        this.SetValue(this.value + ArgSize.value);
    },
    Can_Decrease: function()
    {
        return this.value !== -2;
    },
    Can_Increase: function()
    {
        return this.value == -1 || this.value == -2;
    },
    Can_SimpleIncrease: function()
    {
        return this.value !== 2;
    },
    Write_ToBinary: function(Writer)
    {
        if(this.value == undefined)
        {
            Writer.WriteBool(true);
        }
        else
        {
            Writer.WriteBool(false);
            Writer.WriteLong(this.value);
        }
    },
    Read_FromBinary: function(Reader)
    {
        if(Reader.GetBool() == false)
        {
            this.value = Reader.GetLong();
        }
        else
        {
            this.value = undefined;
        }
    }
};

function CMathGapsInfo(argSize)
{
    this.argSize = argSize; // argSize выставляем один раз для всего контента

    this.Left    = null;    // элемент слева
    this.Current = null;    // текущий элемент

    this.LeftFontSize    = null;
    this.CurrentFontSize = null;
    this.bUpdate         = false;
}
CMathGapsInfo.prototype =
{
    setGaps: function(Current, CurrentFontSize)
    {
        this.updateCurrentObject(Current, CurrentFontSize);
        this.updateGaps();
    },
    updateCurrentObject: function(Current, CurrentFontSize)
    {
        this.Left = this.Current;
        this.LeftFontSize = this.CurrentFontSize;

        this.Current = Current;
        this.CurrentFontSize = CurrentFontSize;
    },
    updateGaps: function()
    {
        if(this.argSize < 0)
        {
            this.Current.GapLeft = 0;

            if(this.Left !== null)
                this.Left.GapRight = 0;
        }
        else
        {
            var leftCoeff = 0,  /// for Current Object
                rightCoeff = 0; /// for Left Object

            var leftCode;


            if(this.Current.IsText())
            {
                var currCode = this.Current.getCodeChr();

                if(this.Left !== null)
                {
                    if(this.Left.Type == para_Math_Composition)
                    {
                        rightCoeff = this.getGapsMComp(this.Left, 1);
                        leftCoeff = COEFF_GAPS.getCoeff(currCode, -1, -1);

                        if(leftCoeff > rightCoeff)
                            leftCoeff -= rightCoeff;
                    }
                    else if(this.Left.IsText())
                    {
                        leftCode = this.Left.getCodeChr();
                        leftCoeff = COEFF_GAPS.getCoeff(currCode, leftCode, -1);
                        rightCoeff = COEFF_GAPS.getCoeff(leftCode, currCode, 1);
                    }

                }
                else
                    this.Current.GapLeft = 0;
            }
            else if(this.Current.Type == para_Math_Composition)
            {
                leftCoeff = this.getGapsMComp(this.Current, -1);

                if(this.Left !== null)
                {
                    if(this.Left.Type == para_Math_Composition)
                    {
                        rightCoeff = this.getGapsMComp(this.Left, 1);

                        if(rightCoeff/2 > leftCoeff)
                            rightCoeff -= leftCoeff;
                        else
                            rightCoeff /= 2;

                        if(leftCoeff < rightCoeff/2)
                            leftCoeff = rightCoeff/2;
                        else
                            leftCoeff -= rightCoeff/2;
                    }
                    else if(this.Left.IsText())
                    {
                        leftCode = this.Left.getCodeChr();
                        rightCoeff = COEFF_GAPS.getCoeff(leftCode, -1, 1);
                        if(rightCoeff > leftCoeff)
                            rightCoeff -= leftCoeff;
                    }
                }
                else
                    leftCoeff = 0;
            }

            var LGapSign = 0.1513*this.CurrentFontSize;
            this.Current.GapLeft = (leftCoeff*LGapSign*100 | 0)/100; // если ни один случай не выполнился, выставляем "нулевые" gaps (default): необходимо, если что-то удалили и объект стал первый или последним в контенте

            if(this.Left !== null)
            {
                var RGapSign = 0.1513*this.LeftFontSize;
                this.Left.GapRight = (rightCoeff*RGapSign*100 | 0)/100;
            }
        }
    },
    getGapsMComp: function(MComp, direct)
    {
        var kind = MComp.kind;
        var checkGap = this.checkGapKind(MComp);

        var bNeedGap = !checkGap.bEmptyGaps && !checkGap.bChildGaps;

        var coeffLeft  = 0.001,
            coeffRight = 0; // for checkGap.bEmptyGaps

        //var bDegree = kind == MATH_DEGREE || kind == MATH_DEGREESubSup;
        var bDegree = kind == MATH_DEGREE;

        if(checkGap.bChildGaps)
        {
            if(bDegree)
            {
                coeffLeft  = 0.03;

                if(MComp.IsPlhIterator())
                    coeffRight = 0.12;
                else
                    coeffRight = 0.16;
            }

            var gapsChild = MComp.getGapsInside(this);

            coeffLeft  = coeffLeft  < gapsChild.left  ? gapsChild.left  : coeffLeft;
            coeffRight = coeffRight < gapsChild.right ? gapsChild.right : coeffRight;
        }
        else if(bNeedGap)
        {
            coeffLeft = 0.4;
            coeffRight = 0.3;
        }


        return direct == -1 ? coeffLeft : coeffRight;
    },
    checkGapKind: function(Comp)
    {
        var kind       = Comp.kind;

        var bEmptyGaps = kind == MATH_MATRIX /*|| (kind == MATH_DELIMITER && Comp.Is_EmptyGaps())*/,
            bChildGaps = kind == MATH_DEGREE || kind == MATH_DEGREESubSup || kind == MATH_ACCENT || kind == MATH_RADICAL || kind == MATH_LIMIT || kind == MATH_BORDER_BOX;


        return  {bEmptyGaps: bEmptyGaps, bChildGaps: bChildGaps};
    }
};

function CMPrp()
{
    this.sty      = undefined;
    this.scr      = undefined;
    this.nor      = undefined;

    this.aln      = undefined;
    this.brk      = undefined;
    this.lit      = undefined;

    // Default
    /*this.sty    = STY_ITALIC;
    this.scr      = TXT_ROMAN;

    this.nor      = false;

    this.aln      = false;
    this.brk      = false;
    this.lit      = false;*/

    // TXT_NORMAL
    // если normal == false, то берем TextPrp отсюда (в wRunPrp bold/italic не учитываем, выставляем отсюда)
    // если normal == true, то их Word не учитывает и берет TextPr из wRunPrp

    // TXT_PLAIN
    // если plain == true
    // буквы берутся обычные, не специальные для Cambria Math : то есть как для TXT_NORMAL
    // отличие от TXT_NORMAL w:rPrp в этом случае не учитываются !

}
CMPrp.prototype =
{
    Set_Pr: function(Pr)
    {
        if(Pr.sty !== undefined)
            this.sty = Pr.sty;

        if(Pr.scr !== undefined)
            this.scr = Pr.scr;

        if(Pr.nor !== undefined)
            this.nor = Pr.nor;

        if(Pr.aln !== undefined)
            this.aln = Pr.aln;

        if(Pr.lit !== undefined)
            this.lit = Pr.lit;

        if(Pr.brk !== undefined)
        {
            this.brk = new CMathBreak();
            this.brk.Set_FromObject(Pr.brk);
        }
    },
    GetTxtPrp: function()
    {
        var textPrp = new CTextPr();

        if(this.sty == undefined)
        {
            textPrp.Italic = true;
            textPrp.Bold   = false;
        }
        else
        {
            textPrp.Italic = this.sty == STY_BI || this.sty == STY_ITALIC;
            textPrp.Bold   = this.sty == STY_BI || this.sty == STY_BOLD;
        }


        return textPrp;
    },
    Copy: function()
    {
        var NewMPrp = new CMPrp();
        
        NewMPrp.aln      = this.aln;
        NewMPrp.lit      = this.lit;
        NewMPrp.nor      = this.nor;
        NewMPrp.sty      = this.sty;
        NewMPrp.scr      = this.scr;

        if(this.brk !== undefined)
            NewMPrp.brk = this.brk.Copy();
        
        return NewMPrp;
    },
    IsBreak: function()
    {
        return this.brk !== undefined;
    },
    Get_AlignBrk: function()
    {
        return this.brk !== undefined ? this.brk.Get_AlignBrk() : null;
    },
    Get_AlnAt: function()
    {
        return this.brk != undefined ? this.brk.Get_AlnAt() : undefined;
    },
    GetCompiled_ScrStyles : function()
    {
        var nor = this.nor == undefined ? false : this.nor;
        var scr = this.scr == undefined ? TXT_ROMAN : this.scr;
        var sty = this.sty == undefined ? STY_ITALIC : this.sty;

        return {nor: nor, scr: scr, sty: sty};
    },
    SetStyle: function(Bold, Italic) /// из ctrPrp получить style для MathPrp
    {
        if(Bold == true && Italic == true)
            this.sty = STY_BI;
        else if(Italic == true)
            this.sty = STY_ITALIC;
        else if(Bold == true)
            this.sty = STY_BOLD;
        else if(Bold == false && Italic == false)
            this.sty = STY_PLAIN;
        else
            this.sty = undefined;
    },
    GetBoldItalic: function()
    {
        var Object =
        {
            Italic: undefined,
            Bold:   undefined
        };

        if(this.sty == STY_BI)
        {
            Object.Bold   = true;

        }
        else if(this.sty == STY_BOLD)
        {
            Object.Bold   = true;
            Object.Italic = false;
        }

        return Object;
    },
    Displace_Break: function(isForward)
    {
        if(this.brk !== undefined)
        {
            this.brk.Displace(isForward);
        }
    },
    Apply_AlnAt: function(alnAt)
    {
        if(this.brk !== undefined)
        {
            this.brk.Apply_AlnAt(alnAt);
        }
    },
    Insert_ForcedBreak: function(AlnAt)
    {
        if(this.brk == undefined)
            this.brk = new CMathBreak();

        this.brk.Apply_AlnAt(AlnAt);
    },
    Delete_ForcedBreak: function()
    {
        this.brk = undefined;
    }
};
CMPrp.prototype.Write_ToBinary = function(Writer)
{
	var StartPos = Writer.GetCurPosition();
	Writer.Skip(4);

	var Flags = 0;
	if (undefined != this.aln)
	{
		Writer.WriteBool(this.aln);
		Flags |= 1;
	}
	if (undefined != this.brk)
	{
		this.brk.Write_ToBinary(Writer);
		Flags |= 2;
	}
	if (undefined != this.lit)
	{
		Writer.WriteBool(this.lit);
		Flags |= 4;
	}
	if (undefined != this.nor)
	{
		Writer.WriteBool(this.nor);
		Flags |= 8;
	}
	if (undefined != this.scr)
	{
		Writer.WriteLong(this.scr);
		Flags |= 16;
	}
	if (undefined != this.sty)
	{
		Writer.WriteLong(this.sty);
		Flags |= 32;
	}

	var EndPos = Writer.GetCurPosition();
	Writer.Seek(StartPos);
	Writer.WriteLong(Flags);
	Writer.Seek(EndPos);
};
CMPrp.prototype.Read_FromBinary = function(Reader)
{
	var Flags = Reader.GetLong();

	if (Flags & 1)
		this.aln = Reader.GetBool();

	if (Flags & 2)
	{
		this.brk = new CMathBreak();
		this.brk.Read_FromBinary(Reader);
	}

	if (Flags & 4)
		this.lit = Reader.GetBool();

	if (Flags & 8)
		this.nor = Reader.GetBool();

	if (Flags & 16)
		this.scr = Reader.GetLong();

	if (Flags & 32)
		this.sty = Reader.GetLong();
};

/**
 *
 * @constructor
 * @extends {CParagraphContentWithParagraphLikeContent}
 */
function CMathContent()
{
	CParagraphContentWithParagraphLikeContent.call(this);

	this.Id = AscCommon.g_oIdCounter.Get_NewId();		

    this.Content = []; // array of mathElem

    this.Type = para_Math_Content;
    this.CurPos = 0;
    this.pos    = new CMathPosition();   // относительная позиция

    //  Properties
    this.ParaMath       = null;
    this.ArgSize        = new CMathArgSize();
    this.Compiled_ArgSz = new CMathArgSize();

    // for EqArray
    this.InfoPoints = new CInfoPoints();
    ///////////////

    this.Bounds = new CMathBounds();

    this.plhHide        = false;
    this.bRoot          = false;
    this.bOneLine  = false;

    //////////////////

    this.Selection =
    {
        StartPos:  0,
        EndPos:    0,
        Use:    false
    };

    this.RecalcInfo =
    {
        TextPr:             true,
        bEqArray:          false,
        bChangeInfoPoints:  false,
        Measure:            true
    };

    this.NearPosArray = [];
    this.ParentElement = null;

    this.size = new CMathSize();

	this.m_oContentChanges = new AscCommon.CContentChanges(); // список изменений(добавление/удаление элементов)
	
	// Добавляем данный класс в таблицу Id (обязательно в конце конструктора)
    AscCommon.g_oTableId.Add( this, this.Id );
}
CMathContent.prototype = Object.create(CParagraphContentWithParagraphLikeContent.prototype);
CMathContent.prototype.constructor = CMathContent;
CMathContent.prototype.init = function()
{

};
CMathContent.prototype.addElementToContent = function(obj)
{
    this.Internal_Content_Add(this.Content.length, obj, false);
    this.CurPos = this.Content.length-1;
};
CMathContent.prototype.SetPlaceholder = function()
{
	this.Remove_FromContent(0, this.Content.length);

    var oRun = new ParaRun(null, true);
    oRun.AddMathPlaceholder();

    this.addElementToContent(oRun);

    this.RemoveSelection();
    this.MoveCursorToEndPos();
};
//////////////////////////////////////
CMathContent.prototype.PreRecalc = function(Parent, ParaMath, ArgSize, RPI)
{
    this.ParaMath = ParaMath;
    if(Parent !== null)
    {
        this.bRoot = false;
        this.Parent = Parent;
    }

    if(ArgSize !== null && ArgSize !== undefined)
    {
        this.Compiled_ArgSz.value = this.ArgSize.value;
        this.Compiled_ArgSz.Merge(ArgSize);
    }

    var lng = this.Content.length;

    var GapsInfo = new CMathGapsInfo(this.Compiled_ArgSz.value);

    if(!this.bRoot)
        this.RecalcInfo.bEqArray = this.Parent.IsEqArray();

    for(var pos = 0; pos < lng; pos++)
    {
        if(this.Content[pos].Type == para_Math_Composition)
        {
            this.Content[pos].PreRecalc(this, ParaMath, this.Compiled_ArgSz, RPI, GapsInfo);
        }
        else if(this.Content[pos].Type == para_Math_Run)
            this.Content[pos].Math_PreRecalc(this, ParaMath, this.Compiled_ArgSz, RPI, GapsInfo);
    }

    if(GapsInfo.Current !== null)
        GapsInfo.Current.GapRight = 0;
};
CMathContent.prototype.Math_UpdateGaps = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var GapsInfo = new CMathGapsInfo(this.Compiled_ArgSz.value);

    if(StartPos !== undefined && EndPos !== undefined && CurLine < this.protected_GetLinesCount())
    {
        if(CurLine > 0 && StartPos !== EndPos) // выставим объект, который будет Left для первого элемента в текущей строке
        {
            var EndPosPrev   = this.protected_GetRangeEndPos(CurLine - 1, CurRange);
            this.Content[EndPosPrev].UpdLastElementForGaps(_CurLine - 1, _CurRange, GapsInfo);
        }


        for(var Pos = StartPos; Pos <= EndPos; Pos++)
        {
            this.Content[Pos].Math_UpdateGaps(_CurLine, _CurRange, GapsInfo);
        }
    }
};
CMathContent.prototype.IsEqArray = function()
{
    return this.RecalcInfo.bEqArray;
};
CMathContent.prototype.Get_WidthPoints = function()
{
    return this.InfoPoints.ContentPoints;
};
CMathContent.prototype.ShiftPage = function(Dx)
{
    this.Bounds.ShiftPage(Dx);

    for(var Pos = 0; Pos < this.Content.length; Pos++)
    {
        if(this.Content[Pos].Type === para_Math_Composition)
            this.Content[Pos].ShiftPage(Dx);
    }
};
CMathContent.prototype.Get_CompiledArgSize = function()
{
    return this.Compiled_ArgSz;
};
CMathContent.prototype.getGapsInside = function(GapsInfo) // учитываем gaps внутренних объектов
{
    var gaps = {left: 0, right: 0};
    var bFirstComp = false,
        bLastComp = false;

    var len = this.Content.length;

    if(len > 1)
    {
        var bFRunEmpty = this.Content[0].Is_Empty();
        bFirstComp = bFRunEmpty && this.Content[1].Type == para_Math_Composition; // первый всегда идет Run

        var bLastRunEmpty = this.Content[len - 1].Is_Empty(); // т.к. после мат. объекта стоит пустой Run
        bLastComp = bLastRunEmpty && this.Content[len - 2].Type == para_Math_Composition;
    }

    var checkGap;

    if(bFirstComp)
    {
        checkGap = GapsInfo.checkGapKind(this.Content[1]);

        if(!checkGap.bChildGaps)
        {
            gaps.left = GapsInfo.getGapsMComp(this.Content[1], -1);
        }
    }

    if(bLastComp)
    {
        checkGap = GapsInfo.checkGapKind(this.Content[len - 1]);

        if(!checkGap.bChildGaps)
        {
            gaps.right = GapsInfo.getGapsMComp(this.Content[len - 2], 1);
        }
    }

    return gaps;
};
CMathContent.prototype.draw = function(x, y, pGraphics, PDSE)
{
    var StartPos, EndPos;

    if(this.bRoot)
    {
        var CurLine  = PDSE.Line - this.StartLine;
        var CurRange = ( 0 === CurLine ? PDSE.Range - this.StartRange : PDSE.Range );

        StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
        EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);
    }
    else
    {
        StartPos = 0;
        EndPos   = this.Content.length - 1;
    }

    var bHidePlh = this.plhHide && this.IsPlaceholder();

    if( !bHidePlh )
    {
        for(var i = StartPos; i <= EndPos;i++)
        {
            if(this.Content[i].Type == para_Math_Composition)
            {
                this.Content[i].draw(x, y, pGraphics, PDSE);
            }
            else
                this.Content[i].Draw_Elements(PDSE);
        }
    }
};
CMathContent.prototype.Draw_Elements = function(PDSE)
{
    var StartPos, EndPos;

    if(this.protected_GetLinesCount() > 0)
    {
        var CurLine  = PDSE.Line - this.StartLine;
        var CurRange = ( 0 === CurLine ? PDSE.Range - this.StartRange : PDSE.Range );

        StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
        EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);
    }
    else
    {
        StartPos = 0;
        EndPos   = this.Content.length - 1;
    }

    var bHidePlh = this.plhHide && this.IsPlaceholder();

    if( !bHidePlh )
    {
        for(var CurPos = StartPos; CurPos <= EndPos;CurPos++)
            this.Content[CurPos].Draw_Elements(PDSE);
    }

};
CMathContent.prototype.setCtrPrp = function()
{

};
CMathContent.prototype.Is_InclineLetter = function()
{
    var result = false;

    if(this.Content.length == 1)
        result = this.Content[0].Math_Is_InclineLetter();

    return result;
};
CMathContent.prototype.IsPlaceholder = function()
{
    var bPlh = false;
    if(this.Content.length == 1)
        bPlh  = this.Content[0].IsPlaceholder();

    return bPlh;
};
CMathContent.prototype.Can_GetSelection = function()
{
    var bPlh = false;

    if(this.Content.length == 1)
        bPlh  = this.Content[0].IsPlaceholder();

    return bPlh || this.bRoot == false;
};
CMathContent.prototype.IsJustDraw = function()
{
    return false;
};
CMathContent.prototype.ApplyPoints = function(WidthsPoints, Points, MaxDimWidths)
{
    this.InfoPoints.GWidths       = WidthsPoints;
    this.InfoPoints.GPoints       = Points;
    this.InfoPoints.GMaxDimWidths = MaxDimWidths;
    // точки выравнивания данного контента содержатся в ContentPoints

    var PosInfo = new CMathPointInfo();
    PosInfo.SetInfoPoints(this.InfoPoints);

    this.size.width = 0;

    for(var i = 0 ; i < this.Content.length; i++)
    {
        if(this.Content[i].Type === para_Math_Run)
        {
            this.Content[i].ApplyPoints(PosInfo);
        }

        this.size.width += this.Content[i].size.width;
    }

    this.Bounds.SetWidth(0, 0, this.size.width);
};
CMathContent.prototype.UpdateBoundsPosInfo = function(PRSA, _CurLine, _CurRange, _CurPage)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    this.Bounds.SetGenPos(CurLine, CurRange, PRSA);
    this.Bounds.SetPage(CurLine, CurRange, _CurPage);

    for(var Pos = StartPos; Pos <= EndPos; Pos++)
    {
        if(this.Content[Pos].Type == para_Math_Composition)
            this.Content[Pos].UpdateBoundsPosInfo(PRSA, _CurLine, _CurRange, _CurPage);
    }

};
CMathContent.prototype.setPosition = function(pos, PosInfo)
{
    var Line  = PosInfo.CurLine,
        Range = PosInfo.CurRange;

    var CurLine  = Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? Range - this.StartRange : Range );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    if(this.RecalcInfo.bEqArray)
    {
        var PosInfoEqq = new CMathPointInfo();
        PosInfoEqq.SetInfoPoints(this.InfoPoints);

        pos.x += PosInfoEqq.GetAlign();
    }

    this.pos.x = pos.x;
    this.pos.y = pos.y;

    this.Bounds.SetPos(CurLine, CurRange, this.pos);

    for(var Pos = StartPos; Pos <= EndPos; Pos++)
    {
        if(this.Content[Pos].Type == para_Math_Run)
            this.Content[Pos].Math_SetPosition(pos, PosInfo);
        else
            this.Content[Pos].setPosition(pos, PosInfo);
    }
};
CMathContent.prototype.Shift_Range = function(Dx, Dy, _CurLine, _CurRange, _CurPage)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    this.Bounds.ShiftPos(CurLine, CurRange, Dx, Dy);

	CParagraphContentWithParagraphLikeContent.prototype.Shift_Range.call(this, Dx, Dy, _CurLine, _CurRange, _CurPage);
};
CMathContent.prototype.SetParent = function(Parent, ParaMath)
{
    this.Parent   = Parent;
    this.ParaMath = ParaMath;
};
CMathContent.prototype.CheckRunContent = function(fCheck)
{
	for (var i = 0; i < this.Content.length; ++i)
	{
		if (para_Math_Run === this.Content[i].Type)
			fCheck(this.Content[i]);
	}
};


///// properties /////
CMathContent.prototype.hidePlaceholder = function(flag)
{
    this.plhHide = flag;
};
///////// RunPrp, CtrPrp

CMathContent.prototype.Get_FirstTextPr = function()
{
    return this.Content[0].Get_FirstTextPr();
};
CMathContent.prototype.getFirstRPrp  = function()
{
    return this.Content[0].Get_CompiledPr(true);
};
CMathContent.prototype.GetCtrPrp = function()       // for placeholder
{
    var ctrPrp = new CTextPr();
    if(!this.bRoot)
        ctrPrp.Merge( this.Parent.Get_CompiledCtrPrp_2() );

    return ctrPrp;
};
CMathContent.prototype.IsAccent = function()
{
    var result = false;

    if(!this.bRoot)
        result = this.Parent.IsAccent();

    return result;
};
////////////////////////
/// For Para Math
CMathContent.prototype.GetParent = function()
{
    return this.Parent.GetParent();
};
CMathContent.prototype.SetArgSize = function(val)
{
	History.Add(new CChangesMathContentArgSize(this, this.ArgSize.GetValue(), val));
	this.ArgSize.SetValue(val);
};
CMathContent.prototype.GetArgSize = function()
{
    return this.ArgSize.value;
};
/////////   Перемещение     ////////////

// Поиск позиции, селект
CMathContent.prototype.IsSelectedAll = function(Props)
{
    var bFirst = false, bEnd = false;

    if(this.Selection.StartPos == 0 && this.Selection.EndPos == this.Content.length - 1)
    {
        if(this.Content[this.Selection.StartPos].Type == para_Math_Run)
            bFirst = this.Content[this.Selection.StartPos].IsSelectedAll(Props);
        else
            bFirst = true;

        if(this.Content[this.Selection.EndPos].Type == para_Math_Run)
            bEnd = this.Content[this.Selection.EndPos].IsSelectedAll(Props);
        else
            bEnd = true;
    }

    return bFirst && bEnd;
};

///////////////////////

CMathContent.prototype.Get_Id = function()
{
    return this.GetId();
};
CMathContent.prototype.GetId = function()
{
    return this.Id;
};
CMathContent.prototype.private_CorrectContent = function()
{
    var len = this.Content.length;

    var EmptyRun = null;
    var RPr      = null;
    var CurrPos = 0;

    while(CurrPos < len)
    {
        var Current = this.Content[CurrPos];

        var bLeftRun  = CurrPos > 0 ? this.Content[CurrPos-1].Type == para_Math_Run : false,
            bRightRun = CurrPos < len - 1 ? this.Content[CurrPos + 1].Type === para_Math_Run : false;

        var bCurrComp       = Current.Type == para_Math_Composition,
            bCurrEmptyRun   = Current.Type == para_Math_Run && Current.Is_Empty();

        var bDeleteEmptyRun = bCurrEmptyRun && (bLeftRun || bRightRun);

        if(bCurrComp && !bLeftRun)
        {
            EmptyRun = new ParaRun(null, true);
            EmptyRun.Set_RFont_ForMathRun();

            RPr = Current.Get_CtrPrp(false);
            EmptyRun.Apply_Pr(RPr);
            this.Internal_Content_Add(CurrPos, EmptyRun);
            CurrPos += 2;

        }
        else if(bDeleteEmptyRun && false == Current.Is_CheckingNearestPos()) // если NearPosArray не нулевой длины, то это вызов происходит на InsertContent, не удаляем пустые Run
        {
            var tempPos = this.CurPos;
            this.Remove_FromContent(CurrPos, 1);

            if (tempPos === CurrPos)
            {
                if (bLeftRun)
                {
                    this.CurPos = CurrPos - 1;
                    this.Content[this.CurPos].MoveCursorToEndPos(false);
                }
                else
                {
                    this.Content[this.CurPos].MoveCursorToStartPos();
                }
            }
        }
        else
        {
            CurrPos++;
        }

        len = this.Content.length;
    }

    if(len > 1)
    {
        if(this.Content[len - 1].Type == para_Math_Composition)
        {
            EmptyRun = new ParaRun(null, true);
            EmptyRun.Set_RFont_ForMathRun();

            RPr = this.Content[len - 1].Get_CtrPrp(false);
            EmptyRun.Apply_Pr(RPr);
            this.Internal_Content_Add(CurrPos, EmptyRun);
        }
    }

};
CMathContent.prototype.Correct_Content = function(bInnerCorrection)
{
    if (true === bInnerCorrection)
    {
        for (var nPos = 0, nCount = this.Content.length; nPos < nCount; nPos++)
        {
            if (para_Math_Composition === this.Content[nPos].Type)
                this.Content[nPos].Correct_Content(true);
        }
    }

    this.private_CorrectContent();

    // Если в контенте ничего нет, тогда добавляем пустой ран
    if (this.Content.length < 1)
    {
        var NewMathRun = new ParaRun(null, true);
        NewMathRun.Set_RFont_ForMathRun();
        this.Add_ToContent(0, NewMathRun);
    }

	// Если единственный элемент данного контента ран и он пустой, заполняем его плейсхолдером
	if (1 === this.Content.length && para_Math_Run === this.Content[0].Type && true === this.Content[0].Is_Empty())
	{
		this.Content[0].AddMathPlaceholder();
	}

	// Возможна ситуация, когда у нас остались лишние плейсхолдеры (либо их слишком много, либо они вообще не нужны)
	if (true !== this.IsPlaceholder())
	{
		var isEmptyContent = true;
		for (var nPos = 0, nCount = this.Content.length; nPos < nCount; ++nPos)
		{
			if (para_Math_Run === this.Content[nPos].Type)
			{
				this.Content[nPos].RemoveMathPlaceholder();
				if (false === this.Content[nPos].Is_Empty())
					isEmptyContent = false;
			}
			else
			{
				isEmptyContent = false;
			}
		}

		if (isEmptyContent)
			this.SetPlaceholder();
	}

	if (this.CurPos >= this.Content.length)
		this.CurPos = this.Content.length - 1;
    if (this.CurPos < 0)
    	this.CurPos = 0;
};
CMathContent.prototype.Correct_ContentPos = function(nDirection)
{
    var nCurPos = this.CurPos;

    if (nCurPos < 0)
    {
        this.CurPos = 0;
        this.Content[0].MoveCursorToStartPos();
    }
    else if (nCurPos > this.Content.length - 1)
    {
        this.CurPos = this.Content.length - 1;
        this.Content[this.CurPos].MoveCursorToEndPos();
    }
    else if (para_Math_Run !== this.Content[nCurPos].Type)
    {
        if (nDirection > 0)
        {
            this.CurPos = nCurPos + 1;
            this.Content[this.CurPos].MoveCursorToStartPos();
        }
        else
        {
            this.CurPos = nCurPos - 1;
            this.Content[this.CurPos].MoveCursorToEndPos();
        }
    }
};

/// функции для работы с курсором
CMathContent.prototype.Cursor_Is_Start = function()
{
    var result = false;

    if( !this.Is_Empty() )
    {
        if(this.CurPos == 0)
            result = this.Content[0].Cursor_Is_Start();
    }

    return result;
};
CMathContent.prototype.Cursor_Is_End = function()
{
    var result = false;

    if(!this.Is_Empty())
    {
        var len = this.Content.length - 1;
        if(this.CurPos == len)
        {
            result = this.Content[len].Cursor_Is_End();
        }
    }

    return result;
};
//////////////////////////////////////

/////////////////////////
//  Text Properties
///////////////

CMathContent.prototype.Get_TextPr = function(ContentPos, Depth)
{
    var pos = ContentPos.Get(Depth);

    var TextPr;

    if(true !== this.bRoot && this.IsPlaceholder())
        TextPr = this.Parent.Get_CtrPrp(true);
    else
        TextPr = this.Content[pos].Get_TextPr(ContentPos, Depth + 1);

    return TextPr;
};
CMathContent.prototype.Get_ParentCtrRunPr = function(bCopy)
{
    return this.Parent.Get_CtrPrp(bCopy);
};
CMathContent.prototype.Get_CompiledTextPr = function(Copy, bAll)
{
    var TextPr = null;

    if(true !== this.bRoot && this.IsPlaceholder())
    {
        TextPr = this.Parent.Get_CompiledCtrPrp_2();
    }
    else if (this.Selection.Use || bAll == true)
    {
        var StartPos, EndPos;

        if(bAll == true)
        {
            StartPos = 0;
            EndPos = this.Content.length - 1;
        }
        else
        {
            StartPos = this.Selection.StartPos;
            EndPos   = this.Selection.EndPos;

            if ( StartPos > EndPos )
            {
                StartPos = this.Selection.EndPos;
                EndPos   = this.Selection.StartPos;
            }
        }


        // пропускаем пустые рана только для случая, когда есть селект

        while ( null === TextPr && StartPos <= EndPos )
        {
            var bComp = this.Content[StartPos].Type == para_Math_Composition,
                bEmptyRun = this.Content[StartPos].Type == para_Math_Run && true === this.Content[StartPos].IsSelectionEmpty();

            if(bComp || !bEmptyRun || bAll)    //пропускаем пустые Run
                TextPr = this.Content[StartPos].Get_CompiledTextPr(true);

            StartPos++;
        }

        while(this.Content[EndPos].Type == para_Math_Run && true === this.Content[EndPos].IsSelectionEmpty() && StartPos < EndPos + 1 && bAll == false) //пропускаем пустые Run
        {
            EndPos--;
        }


        for ( var CurPos = StartPos; CurPos < EndPos + 1; CurPos++ )
        {
            //var CurTextPr = this.Content[CurPos].Get_CompiledPr(false);
            var CurTextPr = this.Content[CurPos].Get_CompiledTextPr(false);

            if ( null !== CurTextPr )
                TextPr = TextPr.Compare( CurTextPr );
        }
    }
    else
    {
        var CurPos = this.CurPos;

        if ( CurPos >= 0 && CurPos < this.Content.length )
            TextPr = this.Content[CurPos].Get_CompiledTextPr(Copy);
    }

    return TextPr;
};
CMathContent.prototype.GetMathTextPrForMenu = function(ContentPos, Depth)
{
    var pos = ContentPos.Get(Depth);

    return this.Content[pos].GetMathTextPrForMenu(ContentPos, Depth + 1);
};
CMathContent.prototype.Apply_TextPr = function(TextPr, IncFontSize, ApplyToAll, StartPos, EndPos)
{
    if ( true === ApplyToAll )
    {
        for ( var i = 0; i < this.Content.length; i++ )
            this.Content[i].Apply_TextPr( TextPr, IncFontSize, true );
    }
    else
    {
        var StartPos, EndPos, bMenu = false;

        if(StartPos !== undefined && EndPos !== undefined)
        {
            StartPos = StartPos;
            EndPos   = EndPos;

            bMenu = true;
        }
        else
        {
            StartPos = this.Selection.StartPos;
            EndPos   = this.Selection.EndPos;
        }

        var NewRuns;
        var LRun, CRun, RRun;

        var bSelectOneElement = this.Selection.Use && StartPos == EndPos;

        var FirstPos = this.Selection.Use ? Math.min(StartPos, EndPos) : this.CurPos;

        if(FirstPos == 0 && this.bRoot)
            this.ParaMath.SetRecalcCtrPrp(this.Content[0]);

        if( ( !this.Selection.Use && !bMenu ) || (bSelectOneElement && this.Content[StartPos].Type == para_Math_Run) ) // TextPr меняем только в одном Run
        {
            var Pos = !this.Selection.Use ? this.CurPos :  StartPos;

            NewRuns = this.Content[Pos].Apply_TextPr(TextPr, IncFontSize, false);

            LRun = NewRuns[0];
            CRun = NewRuns[1];
            RRun = NewRuns[2];

            var CRunPos = Pos;

            if(LRun !== null)
            {
                this.Internal_Content_Add(Pos+1, CRun);
                CRunPos = Pos + 1;
            }

            if(RRun !== null)
            {
                this.Internal_Content_Add(CRunPos+1, RRun);
            }

            this.CurPos             = CRunPos;
            this.Selection.StartPos = CRunPos;
            this.Selection.EndPos   = CRunPos;

        }
        else if(bSelectOneElement && this.Content[StartPos].Type == para_Math_Composition)  // заселекчен только один мат. объект
        {
            this.Content[StartPos].Apply_TextPr(TextPr, IncFontSize, true);
        }
        else
        {

            if(StartPos > EndPos)
            {
                var temp = StartPos;
                StartPos = EndPos;
                EndPos = temp;
            }


            for(var i = StartPos + 1; i < EndPos; i++)
                this.Content[i].Apply_TextPr(TextPr, IncFontSize, true );


            if(this.Content[EndPos].Type == para_Math_Run)
            {
                NewRuns = this.Content[EndPos].Apply_TextPr(TextPr, IncFontSize, false);

                // LRun - null
                CRun = NewRuns[1];
                RRun = NewRuns[2];

                if(RRun !== null)
                {
                    this.Internal_Content_Add(EndPos+1, RRun);
                }

            }
            else
                this.Content[EndPos].Apply_TextPr(TextPr, IncFontSize, true);


            if(this.Content[StartPos].Type == para_Math_Run)
            {
                NewRuns = this.Content[StartPos].Apply_TextPr(TextPr, IncFontSize, false);

                LRun = NewRuns[0];
                CRun = NewRuns[1];
                // RRun - null


                if(LRun !== null)
                {
                    this.Internal_Content_Add(StartPos+1, CRun);
                }

            }
            else
                this.Content[StartPos].Apply_TextPr(TextPr, IncFontSize, true);


            var bStartComposition = this.Content[StartPos].Type == para_Math_Composition || (this.Content[StartPos].Is_Empty() && this.Content[StartPos + 1].Type == para_Math_Composition);
            var bEndCompostion    = this.Content[EndPos].Type == para_Math_Composition || (this.Content[EndPos].Is_Empty()   && this.Content[EndPos - 1].Type == para_Math_Composition);

            if(!bStartComposition)
            {
                if(this.Selection.StartPos < this.Selection.EndPos && true === this.Content[this.Selection.StartPos].IsSelectionEmpty(true) )
                    this.Selection.StartPos++;
                else if (this.Selection.EndPos < this.Selection.StartPos && true === this.Content[this.Selection.EndPos].IsSelectionEmpty(true) )
                    this.Selection.EndPos++;
            }


            if(!bEndCompostion)
            {
                if(this.Selection.StartPos < this.Selection.EndPos && true === this.Content[this.Selection.EndPos].IsSelectionEmpty(true) )
                    this.Selection.EndPos--;
                else if (this.Selection.EndPos < this.Selection.StartPos && true === this.Content[this.Selection.StartPos].IsSelectionEmpty(true) )
                    this.Selection.StartPos--;
            }

        }
    }

};
CMathContent.prototype.Set_MathTextPr2 = function(TextPr, MathPr, bAll, StartPos, Count)
{
    if(bAll)
    {
        StartPos = 0;
        Count = this.Content.length - 1;
    }

    if(Count < 0 || StartPos + Count > this.Content.length - 1)
        return;

    for(var pos = StartPos; pos <= StartPos + Count; pos++)
        this.Content[pos].Set_MathTextPr2(TextPr, MathPr, true);

};
CMathContent.prototype.IsNormalTextInRuns = function()
{
    var flag = true;

    if(this.Selection.Use)
    {
        var StartPos = this.Selection.StartPos,
            EndPos   = this.Selection.EndPos;

        if ( StartPos > EndPos )
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        for(var i = StartPos; i < EndPos+1; i++)
        {
            var curr = this.Content[i],
                currType = curr.Type;
            if(currType == para_Math_Composition || (currType == para_Math_Run && false == curr.IsNormalText()))
            {
                flag = false;
                break;
            }
        }
    }
    else
        flag = false;

    return flag;
};
CMathContent.prototype.Internal_Content_Add = function(Pos, Item, bUpdatePosition)
{
	Item.Set_ParaMath(this.ParaMath);

	if (Item.SetParagraph)
		Item.SetParagraph(this.Paragraph);

	Item.Parent = this;
	Item.Recalc_RunsCompiledPr();

	History.Add(new CChangesMathContentAddItem(this, Pos, [Item]));
	this.Content.splice(Pos, 0, Item);

	this.private_UpdatePosOnAdd(Pos, bUpdatePosition);
};
CMathContent.prototype.private_UpdatePosOnAdd = function(Pos, bUpdatePosition)
{
    if(bUpdatePosition !== false)
    {
        if ( this.CurPos >= Pos )
            this.CurPos++;

        if ( this.Selection.StartPos >= Pos )
            this.Selection.StartPos++;

        if ( this.Selection.EndPos >= Pos )
            this.Selection.EndPos++;

        this.private_CorrectSelectionPos();
        this.private_CorrectCurPos();
    }

    // Обновляем позиции в NearestPos
    var NearPosLen = this.NearPosArray.length;
    for ( var Index = 0; Index < NearPosLen; Index++ )
    {
        var HyperNearPos = this.NearPosArray[Index];
        var ContentPos = HyperNearPos.NearPos.ContentPos;
        var Depth      = HyperNearPos.Depth;

        if (ContentPos.Data[Depth] >= Pos)
            ContentPos.Data[Depth]++;
    }
};
CMathContent.prototype.private_CorrectSelectionPos = function()
{
    this.Selection.StartPos = Math.max(0, Math.min(this.Content.length - 1, this.Selection.StartPos));
    this.Selection.EndPos   = Math.max(0, Math.min(this.Content.length - 1, this.Selection.EndPos));
};
CMathContent.prototype.private_CorrectCurPos = function()
{
    if (this.Content.length <= 0)
    {
        this.CurPos = 0;
        return;
    }

    if (this.CurPos > this.Content.length - 1)
    {
        this.CurPos = this.Content.length - 1;

        if (para_Math_Run === this.Content[this.CurPos].Type)
            this.Content[this.CurPos].MoveCursorToEndPos(false);
    }

    if (this.CurPos < 0)
    {
        this.CurPos = this.Content.length - 1;

        if (para_Math_Run === this.Content[this.CurPos].Type)
            this.Content[this.CurPos].MoveCursorToStartPos();
    }
};
CMathContent.prototype.Correct_ContentCurPos = function()
{
    this.private_CorrectCurPos();

    for(var Pos = 0; Pos < this.Content.length; Pos++)
    {
        if(this.Content[Pos].Type == para_Math_Composition)
            this.Content[Pos].Correct_ContentCurPos();
    }
};
CMathContent.prototype.SplitContent = function(NewContent, ContentPos, Depth)
{
    var Pos = ContentPos.Get(Depth);

    if(para_Math_Run === this.Content[Pos].Type)
    {
        var NewRun = this.Content[Pos].Split(ContentPos, Depth+1);
        NewContent.Add_ToContent(0, NewRun);

        var len = this.Content.length;
        if(Pos < len - 1)
        {
            NewContent.Concat_ToEnd( this.Content.slice(Pos + 1) );
            this.Remove_FromContent(Pos+1, len - Pos - 1);
        }
    }
    else
	{
		NewContent.Add_ToContent(0, new ParaRun(this.Paragraph, true));
	}
};
CMathContent.prototype.Add_ToContent = function(Pos, Item)
{
    if (Item && para_Run === Item.Type)
    {
        var MathRun = new ParaRun(Item.GetParagraph(), true);
        this.Internal_Content_Add(Pos, MathRun);
    }
    else
    {
        this.Internal_Content_Add(Pos, Item);
    }
};
CMathContent.prototype.Concat_ToEnd = function(NewItems)
{
    this.ConcatToContent(2, NewItems);
};
CMathContent.prototype.ConcatToContent = function(Pos, NewItems)
{
	if (NewItems != undefined && NewItems.length > 0)
	{
		var Count = NewItems.length;

		for (var i = 0; i < Count; i++)
		{
			NewItems[i].Set_ParaMath(this.ParaMath);
			NewItems[i].Parent = this;
			NewItems[i].Recalc_RunsCompiledPr();
		}

		History.Add(new CChangesMathContentAddItem(this, Pos, NewItems));

		var Array_start = this.Content.slice(0, Pos);
		var Array_end   = this.Content.slice(Pos);

		this.Content = Array_start.concat(NewItems, Array_end);
	}
};
CMathContent.prototype.Remove_FromContent = function(Pos, Count)
{
	var DeletedItems = this.Content.splice(Pos, Count);
	History.Add(new CChangesMathContentRemoveItem(this, Pos, DeletedItems));

	// Обновим текущую позицию
	if (this.CurPos > Pos + Count)
		this.CurPos -= Count;
	else if (this.CurPos > Pos)
		this.CurPos = Pos;

	this.private_CorrectCurPos();
	this.private_UpdatePosOnRemove(Pos, Count);
};
CMathContent.prototype.Remove_Content = function(Pos, Count)
{
	var DeletedItems = this.Content.splice(Pos, Count);
	History.Add(new CChangesMathContentRemoveItem(this, Pos, DeletedItems));

	// Обновим текущую позицию
	if (this.CurPos > Pos + Count) {
		this.CurPos -= Count;
    } else if (this.CurPos > Pos) {
        this.CurPos = Pos;
    } else if (this.CurPos > (this.Content.length - 1)) { // подправить условие
        this.CurPos = this.Content.length - 1;
    }
};
CMathContent.prototype.private_UpdatePosOnRemove = function(Pos, Count)
{
    // Обновим начало и конец селекта
    if (true === this.Selection.Use)
    {
        if (this.Selection.StartPos <= this.Selection.EndPos)
        {
            if (this.Selection.StartPos > Pos + Count)
                this.Selection.StartPos -= Count;
            else if (this.Selection.StartPos > Pos)
                this.Selection.StartPos = Pos;

            if (this.Selection.EndPos >= Pos + Count)
                this.Selection.EndPos -= Count;
            else if (this.Selection.EndPos >= Pos)
                this.Selection.EndPos = Math.max(0, Pos - 1);
        }
        else
        {
            if (this.Selection.StartPos >= Pos + Count)
                this.Selection.StartPos -= Count;
            else if (this.Selection.StartPos >= Pos)
                this.Selection.StartPos = Math.max(0, Pos - 1);

            if (this.Selection.EndPos > Pos + Count)
                this.Selection.EndPos -= Count;
            else if (this.Selection.EndPos > Pos)
                this.Selection.EndPos = Pos;
        }

        this.Selection.StartPos = Math.min(this.Content.length - 1, Math.max(0, this.Selection.StartPos));
        this.Selection.EndPos   = Math.min(this.Content.length - 1, Math.max(0, this.Selection.EndPos));
    }

    // Обновляем позиции в NearestPos
    var NearPosLen = this.NearPosArray.length;
    for (var Index = 0; Index < NearPosLen; Index++)
    {
        var HyperNearPos = this.NearPosArray[Index];
        var ContentPos = HyperNearPos.NearPos.ContentPos;
        var Depth      = HyperNearPos.Depth;

        if (ContentPos.Data[Depth] > Pos + Count)
            ContentPos.Data[Depth] -= Count;
        else if (ContentPos.Data[Depth] > Pos)
            ContentPos.Data[Depth] = Math.max(0 , Pos);
    }
};
CMathContent.prototype.Get_Default_TPrp = function()
{
    return this.ParaMath.Get_Default_TPrp();
};
/////////////////////////
CMathContent.prototype.Is_Empty = function()
{
    return this.Content.length == 0;
};
CMathContent.prototype.Copy = function(Selected, oPr)
{
    var NewContent = new CMathContent();
    this.CopyTo(NewContent, Selected, oPr);
    return NewContent;
};
CMathContent.prototype.CopyTo = function(OtherContent, Selected, oPr)
{
    var nStartPos, nEndPos;

    if(true === Selected)
    {
        if(this.Selection.StartPos < this.Selection.EndPos)
        {
            nStartPos = this.Selection.StartPos;
            nEndPos   = this.Selection.EndPos;
        }
        else
        {
            nStartPos = this.Selection.EndPos;
            nEndPos   = this.Selection.StartPos;
        }
    }
    else
    {
        nStartPos = 0;
        nEndPos   = this.Content.length - 1;
    }

    OtherContent.plHid = this.plhHide;
    OtherContent.SetArgSize(this.ArgSize.GetValue());

    for(var nPos = nStartPos; nPos <= nEndPos; nPos++)
    {
        var oElement;
        if(this.Content[nPos].Type == para_Math_Run)
            oElement = this.Content[nPos].Copy(Selected, oPr);
        else
            oElement = this.Content[nPos].Copy(false, oPr);

        OtherContent.Internal_Content_Add(OtherContent.Content.length, oElement);
    }
    if(oPr && oPr.Comparison)
    {
        oPr.Comparison.updateReviewInfo(OtherContent, reviewtype_Add);
    }
};
CMathContent.prototype.getElem = function(nNum)
{
    return this.Content[nNum];
};
CMathContent.prototype.GetLastElement = function()
{
    var pos = this.Content.length - 1;

    while(this.Content[pos].Type == para_Math_Run && this.Content[pos].Is_Empty() && pos > 0)
    {
        pos--;
    }

    var last = this.Content[pos].Type == para_Math_Run ? this.Content[pos] : this.Content[pos].GetLastElement();

    return last;
};
CMathContent.prototype.GetFirstElement = function()
{
    var pos = 0;

    while(this.Content[pos].Type == para_Math_Run && this.Content[pos].Is_Empty() && pos < this.Content.length - 1)
    {
        pos++;
    }

    var first = this.Content[pos].Type == para_Math_Run ? this.Content[pos] : this.Content[pos].GetFirstElement();

    return first;
};
////////////////////////////////////////////////////////////////
CMathContent.prototype.Write_ToBinary2 = function(Writer)
{
    Writer.WriteLong(AscDFH.historyitem_type_MathContent);

    // Long : Id
    Writer.WriteString2(this.Id);
};
CMathContent.prototype.Read_FromBinary2 = function(Reader)
{
    // Long : Id
    this.Id = Reader.GetString2();
};
CMathContent.prototype.Refresh_RecalcData = function()
{
    if(this.ParaMath !== null)
        this.ParaMath.Refresh_RecalcData(); // Refresh_RecalcData сообщает родительскому классу, что у него произошли изменения, нужно пересчитать
};
CMathContent.prototype.Insert_MathContent = function(oMathContent, Pos, bSelect)
{
    if (null === this.ParaMath || null === this.ParaMath.Paragraph)
        bSelect = false;

    if (undefined === Pos)
        Pos = this.CurPos;

    var nCount = oMathContent.Content.length;
    for (var nIndex = 0; nIndex < nCount; nIndex++)
    {
        this.Internal_Content_Add(Pos + nIndex, oMathContent.Content[nIndex], false);

        if (true === bSelect)
        {
            oMathContent.Content[nIndex].SelectAll();
        }
    }

    this.CurPos = Pos + nCount;

    if (true === bSelect)
    {
        this.Selection.Use = true;
        this.Selection.StartPos = Pos;
        this.Selection.EndPos   = Pos + nCount - 1;

        if (!this.bRoot)
            this.ParentElement.Select_MathContent(this);
        else
            this.ParaMath.bSelectionUse = true;

        this.ParaMath.Paragraph.Select_Math(this.ParaMath);
    }

    this.Correct_Content(true);
    this.Correct_ContentPos(-1);
};
CMathContent.prototype.Set_ParaMath = function(ParaMath, Parent)
{
    this.Parent   = Parent;
    this.ParaMath = ParaMath;

    for (var Index = 0, Count = this.Content.length; Index < Count; Index++)
    {
        this.Content[Index].Set_ParaMath(ParaMath, this);
    }
};
CMathContent.prototype.Load_FromMenu = function(Type, Paragraph, TextPr, oSelectedContent)
{
    this.Paragraph = Paragraph;

    var Pr = {
    	ctrPrp : TextPr ? TextPr.Copy() : new CTextPr()
    };

    Pr.ctrPrp.Italic = true;
    Pr.ctrPrp.RFonts.SetAll("Cambria Math", -1);

    var MainType = Type >> 24;

    if (MainType === c_oAscMathMainType.Symbol)
		this.private_LoadFromMenuSymbol(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Fraction)
		this.private_LoadFromMenuFraction(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Script)
        this.private_LoadFromMenuScript(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Radical)
        this.private_LoadFromMenuRadical(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Integral)
        this.private_LoadFromMenuIntegral(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.LargeOperator)
        this.private_LoadFromMenuLargeOperator(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Bracket)
        this.private_LoadFromMenuBracket(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Function)
        this.private_LoadFromMenuFunction(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Accent)
        this.private_LoadFromMenuAccent(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.LimitLog)
        this.private_LoadFromMenuLimitLog(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Operator)
        this.private_LoadFromMenuOperator(Type, Pr, oSelectedContent);
    else if (MainType === c_oAscMathMainType.Matrix)
        this.private_LoadFromMenuMatrix(Type, Pr, oSelectedContent);
    else if(MainType == c_oAscMathMainType.Empty_Content)
		this.private_LoadFromMenuDefaultText(Type, Pr, oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuSymbol = function(Type, Pr)
{
    var Code = -1;

    switch (Type)
    {
        case c_oAscMathType.Symbol_pm            : Code = 0x00B1; break;
        case c_oAscMathType.Symbol_infinity      : Code = 0x221E; break;
        case c_oAscMathType.Symbol_equals        : Code = 0x003D; break;
        case c_oAscMathType.Symbol_neq           : Code = 0x2260; break;
        case c_oAscMathType.Symbol_about         : Code = 0x007E; break;
        case c_oAscMathType.Symbol_times         : Code = 0x00D7; break;
        case c_oAscMathType.Symbol_div           : Code = 0x00F7; break;
        case c_oAscMathType.Symbol_factorial     : Code = 0x0021; break;
        case c_oAscMathType.Symbol_propto        : Code = 0x221D; break;
        case c_oAscMathType.Symbol_less          : Code = 0x003C; break;
        case c_oAscMathType.Symbol_ll            : Code = 0x226A; break;
        case c_oAscMathType.Symbol_greater       : Code = 0x003E; break;
        case c_oAscMathType.Symbol_gg            : Code = 0x226B; break;
        case c_oAscMathType.Symbol_leq           : Code = 0x2264; break;
        case c_oAscMathType.Symbol_geq           : Code = 0x2265; break;
        case c_oAscMathType.Symbol_mp            : Code = 0x2213; break;
        case c_oAscMathType.Symbol_cong          : Code = 0x2245; break;
        case c_oAscMathType.Symbol_approx        : Code = 0x2248; break;
        case c_oAscMathType.Symbol_equiv         : Code = 0x2261; break;
        case c_oAscMathType.Symbol_forall        : Code = 0x2200; break;
        case c_oAscMathType.Symbol_additional    : Code = 0x2201; break;
        case c_oAscMathType.Symbol_partial       : Code = 0x1D715; break;
        case c_oAscMathType.Symbol_sqrt          : this.Add_Radical(Pr, null, null); break;
        case c_oAscMathType.Symbol_cbrt          : this.Add_Radical({ctrPrp : Pr.ctrPrp, type : DEGREE_RADICAL}, null, "3"); break;
        case c_oAscMathType.Symbol_qdrt          : this.Add_Radical({ctrPrp : Pr.ctrPrp, type : DEGREE_RADICAL}, null, "4"); break;
        case c_oAscMathType.Symbol_cup           : Code = 0x222A; break;
        case c_oAscMathType.Symbol_cap           : Code = 0x2229; break;
        case c_oAscMathType.Symbol_emptyset      : Code = 0x2205; break;
        case c_oAscMathType.Symbol_percent       : Code = 0x0025; break;
        case c_oAscMathType.Symbol_degree        : Code = 0x00B0; break;
        case c_oAscMathType.Symbol_fahrenheit    : Code = 0x2109; break;
        case c_oAscMathType.Symbol_celsius       : Code = 0x2103; break;
        case c_oAscMathType.Symbol_inc           : Code = 0x2206; break;
        case c_oAscMathType.Symbol_nabla         : Code = 0x2207; break;
        case c_oAscMathType.Symbol_exists        : Code = 0x2203; break;
        case c_oAscMathType.Symbol_notexists     : Code = 0x2204; break;
        case c_oAscMathType.Symbol_in            : Code = 0x2208; break;
        case c_oAscMathType.Symbol_ni            : Code = 0x220B; break;
        case c_oAscMathType.Symbol_leftarrow     : Code = 0x2190; break;
        case c_oAscMathType.Symbol_uparrow       : Code = 0x2191; break;
        case c_oAscMathType.Symbol_rightarrow    : Code = 0x2192; break;
        case c_oAscMathType.Symbol_downarrow     : Code = 0x2193; break;
        case c_oAscMathType.Symbol_leftrightarrow: Code = 0x2194; break;
        case c_oAscMathType.Symbol_therefore     : Code = 0x2234; break;
        case c_oAscMathType.Symbol_plus          : Code = 0x002B; break;
        case c_oAscMathType.Symbol_minus         : Code = 0x2212; break;
        case c_oAscMathType.Symbol_not           : Code = 0x00AC; break;
        case c_oAscMathType.Symbol_ast           : Code = 0x2217; break;
        case c_oAscMathType.Symbol_bullet        : Code = 0x2219; break;
        case c_oAscMathType.Symbol_vdots         : Code = 0x22EE; break;
        case c_oAscMathType.Symbol_cdots         : Code = 0x22EF; break;
        case c_oAscMathType.Symbol_rddots        : Code = 0x22F0; break;
        case c_oAscMathType.Symbol_ddots         : Code = 0x22F1; break;
        case c_oAscMathType.Symbol_aleph         : Code = 0x2135; break;
        case c_oAscMathType.Symbol_beth          : Code = 0x2136; break;
        case c_oAscMathType.Symbol_QED           : Code = 0x220E; break;
        case c_oAscMathType.Symbol_alpha         : Code = 0x03B1; break;
        case c_oAscMathType.Symbol_beta          : Code = 0x03B2; break;
        case c_oAscMathType.Symbol_gamma         : Code = 0x03B3; break;
        case c_oAscMathType.Symbol_delta         : Code = 0x03B4; break;
        case c_oAscMathType.Symbol_varepsilon    : Code = 0x03B5; break;
        case c_oAscMathType.Symbol_epsilon       : Code = 0x03F5; break;
        case c_oAscMathType.Symbol_zeta          : Code = 0x03B6; break;
        case c_oAscMathType.Symbol_eta           : Code = 0x03B7; break;
        case c_oAscMathType.Symbol_theta         : Code = 0x03B8; break;
        case c_oAscMathType.Symbol_vartheta      : Code = 0x03D1; break;
        case c_oAscMathType.Symbol_iota          : Code = 0x03B9; break;
        case c_oAscMathType.Symbol_kappa         : Code = 0x03BA; break;
        case c_oAscMathType.Symbol_lambda        : Code = 0x03BB; break;
        case c_oAscMathType.Symbol_mu            : Code = 0x03BC; break;
        case c_oAscMathType.Symbol_nu            : Code = 0x03BD; break;
        case c_oAscMathType.Symbol_xsi           : Code = 0x03BE; break;
        case c_oAscMathType.Symbol_o             : Code = 0x03BF; break;
        case c_oAscMathType.Symbol_pi            : Code = 0x03C0; break;
        case c_oAscMathType.Symbol_varpi         : Code = 0x03D6; break;
        case c_oAscMathType.Symbol_rho           : Code = 0x03C1; break;
        case c_oAscMathType.Symbol_varrho        : Code = 0x03F1; break;
        case c_oAscMathType.Symbol_sigma         : Code = 0x03C3; break;
        case c_oAscMathType.Symbol_varsigma      : Code = 0x03C2; break;
        case c_oAscMathType.Symbol_tau           : Code = 0x03C4; break;
        case c_oAscMathType.Symbol_upsilon       : Code = 0x03C5; break;
        case c_oAscMathType.Symbol_varphi        : Code = 0x03C6; break;
        case c_oAscMathType.Symbol_phi           : Code = 0x03D5; break;
        case c_oAscMathType.Symbol_chi           : Code = 0x03C7; break;
        case c_oAscMathType.Symbol_psi           : Code = 0x03C8; break;
        case c_oAscMathType.Symbol_omega         : Code = 0x03C9; break;
        case c_oAscMathType.Symbol_Alpha         : Code = 0x0391; break;
        case c_oAscMathType.Symbol_Beta          : Code = 0x0392; break;
        case c_oAscMathType.Symbol_Gamma         : Code = 0x0393; break;
        case c_oAscMathType.Symbol_Delta         : Code = 0x0394; break;
        case c_oAscMathType.Symbol_Epsilon       : Code = 0x0395; break;
        case c_oAscMathType.Symbol_Zeta          : Code = 0x0396; break;
        case c_oAscMathType.Symbol_Eta           : Code = 0x0397; break;
        case c_oAscMathType.Symbol_Theta         : Code = 0x0398; break;
        case c_oAscMathType.Symbol_Iota          : Code = 0x0399; break;
        case c_oAscMathType.Symbol_Kappa         : Code = 0x039A; break;
        case c_oAscMathType.Symbol_Lambda        : Code = 0x039B; break;
        case c_oAscMathType.Symbol_Mu            : Code = 0x039C; break;
        case c_oAscMathType.Symbol_Nu            : Code = 0x039D; break;
        case c_oAscMathType.Symbol_Xsi           : Code = 0x039E; break;
        case c_oAscMathType.Symbol_O             : Code = 0x039F; break;
        case c_oAscMathType.Symbol_Pi            : Code = 0x03A0; break;
        case c_oAscMathType.Symbol_Rho           : Code = 0x03A1; break;
        case c_oAscMathType.Symbol_Sigma         : Code = 0x03A3; break;
        case c_oAscMathType.Symbol_Tau           : Code = 0x03A4; break;
        case c_oAscMathType.Symbol_Upsilon       : Code = 0x03A5; break;
        case c_oAscMathType.Symbol_Phi           : Code = 0x03A6; break;
        case c_oAscMathType.Symbol_Chi           : Code = 0x03A7; break;
        case c_oAscMathType.Symbol_Psi           : Code = 0x03A8; break;
        case c_oAscMathType.Symbol_Omega         : Code = 0x03A9; break;
    }

    if (-1 !== Code)
    {
        var TextPr, MathPr;

        if (this.Content.length <= 0)
			this.Correct_Content();

        if(this.Content.length > 0 && this.Content[this.CurPos].Type == para_Math_Run && this.IsSelectionEmpty() == true) // находимся в Run, селект отсутствует
        {
            TextPr = this.Content[this.CurPos].Get_TextPr();
            TextPr.RFonts.SetAll("Cambria Math", -1);          //  на данный момент добавляются символы исключительно из Cambria Math
            MathPr = this.Content[this.CurPos].GetMathPr();
        }

        this.Add_Symbol(Code, TextPr, MathPr);
    }
};
CMathContent.prototype.private_LoadFromMenuFraction = function(Type, Pr, oSelectedContent)
{
	var oFraction = null;
    switch (Type)
    {
        case c_oAscMathType.FractionVertical   : oFraction = this.Add_Fraction(Pr, null, null); break;
        case c_oAscMathType.FractionDiagonal   : oFraction = this.Add_Fraction({ctrPrp : Pr.ctrPrp, type : SKEWED_FRACTION}, null, null); break;
        case c_oAscMathType.FractionHorizontal : oFraction = this.Add_Fraction({ctrPrp : Pr.ctrPrp, type : LINEAR_FRACTION}, null, null); break;
        case c_oAscMathType.FractionSmall:
            var oBox = new CBox(Pr);
            this.Add_Element(oBox);
            var BoxMathContent = oBox.getBase();
            BoxMathContent.SetArgSize(-1);
            oFraction = BoxMathContent.Add_Fraction(Pr, null, null);
            break;

        case c_oAscMathType.FractionDifferential_1: this.Add_Fraction(Pr, "dx", "dy"); break;
        case c_oAscMathType.FractionDifferential_2: this.Add_Fraction(Pr, String.fromCharCode(916) + "y", String.fromCharCode(916) + "x"); break;
        case c_oAscMathType.FractionDifferential_3: this.Add_Fraction(Pr, String.fromCharCode(8706) + "y", String.fromCharCode(8706) + "x"); break;
        case c_oAscMathType.FractionDifferential_4: this.Add_Fraction(Pr, String.fromCharCode(948) + "y", String.fromCharCode(948) + "x"); break;
        case c_oAscMathType.FractionPi_2          : this.Add_Fraction(Pr, String.fromCharCode(960), "2"); break;
    }

    if (oFraction && oSelectedContent)
		oFraction.getNumeratorMathContent().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuScript = function(Type, Pr, oSelectedContent)
{
	var oScript = null;
    switch (Type)
    {
        case c_oAscMathType.ScriptSup: oScript = this.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, null, null, null); break;
        case c_oAscMathType.ScriptSub: oScript = this.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, null, null, null); break;
        case c_oAscMathType.ScriptSubSup: oScript = this.Add_Script(true, {ctrPrp : Pr.ctrPrp, type : DEGREE_SubSup}, null, null, null); break;
        case c_oAscMathType.ScriptSubSupLeft: oScript = this.Add_Script(true, {ctrPrp : Pr.ctrPrp, type : DEGREE_PreSubSup}, null, null, null); break;
        case c_oAscMathType.ScriptCustom_1:
            Pr.type = DEGREE_SUBSCRIPT;
            var Script = this.Add_Script(false, Pr, "x", null, null);
            var SubMathContent = Script.getLowerIterator();
            Pr.type = DEGREE_SUPERSCRIPT;
            SubMathContent.Add_Script(false, Pr, "y", "2", null);
            break;

        case c_oAscMathType.ScriptCustom_2: this.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "e", "-i" + String.fromCharCode(969) + "t", null); break;
        case c_oAscMathType.ScriptCustom_3: this.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "x", "2", null); break;
        case c_oAscMathType.ScriptCustom_4: this.Add_Script(true, {ctrPrp : Pr.ctrPrp, type : DEGREE_PreSubSup}, "Y", "n", "1"); break;
    }

	if (oScript && oSelectedContent)
		oScript.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuRadical = function(Type, Pr, oSelectedContent)
{
	var oRadical = null;
    switch (Type)
    {
        case c_oAscMathType.RadicalSqrt:
            Pr.type    = SQUARE_RADICAL;
            Pr.degHide = true;
            oRadical = this.Add_Radical(Pr, null, null);
            break;

        case c_oAscMathType.RadicalRoot_n:
            Pr.type = DEGREE_RADICAL;
			oRadical = this.Add_Radical(Pr, null, null);
            break;

        case c_oAscMathType.RadicalRoot_2:
            Pr.type = DEGREE_RADICAL;
			oRadical = this.Add_Radical(Pr, null, "2");
            break;

        case c_oAscMathType.RadicalRoot_3:
            Pr.type = DEGREE_RADICAL;
			oRadical = this.Add_Radical(Pr, null, "3");
            break;

        case c_oAscMathType.RadicalCustom_1:
            var Fraction = this.Add_Fraction(Pr, null, null);
            var NumMathContent = Fraction.getNumeratorMathContent();
            var DenMathContent = Fraction.getDenominatorMathContent();

            NumMathContent.Add_Text("-b" + String.fromCharCode(177), this.Paragraph);
            Pr.type    = SQUARE_RADICAL;
            Pr.degHide = true;
            var Radical = NumMathContent.Add_Radical(Pr, null, null);
            var RadicalBaseMathContent = Radical.getBase();
            RadicalBaseMathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "b", "2", null);
            RadicalBaseMathContent.Add_Text("-4ac", this.Paragraph);

            DenMathContent.Add_Text("2a", this.Paragraph);

            break;
        case c_oAscMathType.RadicalCustom_2:
            Pr.type    = SQUARE_RADICAL;
            Pr.degHide = true;
            var Radical = this.Add_Radical(Pr, null, null);
            var BaseMathContent = Radical.getBase();

            var ScriptPr = {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT};
            BaseMathContent.Add_Script(false, ScriptPr, "a", "2", null);
            BaseMathContent.Add_Text("+", this.Paragraph);
            BaseMathContent.Add_Script(false, ScriptPr, "b", "2", null);
            break;
    }

	if (oRadical && oSelectedContent)
		oRadical.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuIntegral = function(Type, Pr, oSelectedContent)
{
	var oIntegral = null;
    switch(Type)
    {
        case c_oAscMathType.Integral:                           oIntegral = this.Add_Integral(1, false, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralSubSup:                     oIntegral = this.Add_Integral(1, false, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralCenterSubSup:               oIntegral = this.Add_Integral(1, false, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralDouble:                     oIntegral = this.Add_Integral(2, false, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralDoubleSubSup:               oIntegral = this.Add_Integral(2, false, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralDoubleCenterSubSup:         oIntegral = this.Add_Integral(2, false, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralTriple:                     oIntegral = this.Add_Integral(3, false, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralTripleSubSup:               oIntegral = this.Add_Integral(3, false, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralTripleCenterSubSup:         oIntegral = this.Add_Integral(3, false, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOriented:                   oIntegral = this.Add_Integral(1,  true, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedSubSup:             oIntegral = this.Add_Integral(1,  true, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedCenterSubSup:       oIntegral = this.Add_Integral(1,  true, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedDouble:             oIntegral = this.Add_Integral(2,  true, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedDoubleSubSup:       oIntegral = this.Add_Integral(2,  true, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedDoubleCenterSubSup: oIntegral = this.Add_Integral(2,  true, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedTriple:             oIntegral = this.Add_Integral(3,  true, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedTripleSubSup:       oIntegral = this.Add_Integral(3,  true, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.IntegralOrientedTripleCenterSubSup: oIntegral = this.Add_Integral(3,  true, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.Integral_dx:     Pr.diff = 1; this.Add_Box(Pr, "dx"); break;
        case c_oAscMathType.Integral_dy:     Pr.diff = 1; this.Add_Box(Pr, "dy"); break;
        case c_oAscMathType.Integral_dtheta: Pr.diff = 1; this.Add_Box(Pr, "d" + String.fromCharCode(952)); break;
    }

	if (oIntegral && oSelectedContent)
		oIntegral.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuLargeOperator = function(Type, Pr, oSelectedContent)
{
	var oOperator = null;
    switch(Type)
    {
        case c_oAscMathType.LargeOperator_Sum:              oOperator = this.Add_LargeOperator(1, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Sum_CenterSubSup: oOperator = this.Add_LargeOperator(1, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Sum_SubSup:       oOperator = this.Add_LargeOperator(1, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Sum_CenterSub:    oOperator = this.Add_LargeOperator(1, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Sum_Sub:          oOperator = this.Add_LargeOperator(1, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Prod:              oOperator = this.Add_LargeOperator(2, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Prod_CenterSubSup: oOperator = this.Add_LargeOperator(2, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Prod_SubSup:       oOperator = this.Add_LargeOperator(2, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Prod_CenterSub:    oOperator = this.Add_LargeOperator(2, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Prod_Sub:          oOperator = this.Add_LargeOperator(2, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_CoProd:              oOperator = this.Add_LargeOperator(3, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_CoProd_CenterSubSup: oOperator = this.Add_LargeOperator(3, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_CoProd_SubSup:       oOperator = this.Add_LargeOperator(3, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_CoProd_CenterSub:    oOperator = this.Add_LargeOperator(3, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_CoProd_Sub:          oOperator = this.Add_LargeOperator(3, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Union:              oOperator = this.Add_LargeOperator(4, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Union_CenterSubSup: oOperator = this.Add_LargeOperator(4, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Union_SubSup:       oOperator = this.Add_LargeOperator(4, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Union_CenterSub:    oOperator = this.Add_LargeOperator(4, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Union_Sub:          oOperator = this.Add_LargeOperator(4, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Intersection:              oOperator = this.Add_LargeOperator(5, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Intersection_CenterSubSup: oOperator = this.Add_LargeOperator(5, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Intersection_SubSup:       oOperator = this.Add_LargeOperator(5, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Intersection_CenterSub:    oOperator = this.Add_LargeOperator(5, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Intersection_Sub:          oOperator = this.Add_LargeOperator(5, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Disjunction:              oOperator = this.Add_LargeOperator(6, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Disjunction_CenterSubSup: oOperator = this.Add_LargeOperator(6, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Disjunction_SubSup:       oOperator = this.Add_LargeOperator(6, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Disjunction_CenterSub:    oOperator = this.Add_LargeOperator(6, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Disjunction_Sub:          oOperator = this.Add_LargeOperator(6, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Conjunction:              oOperator = this.Add_LargeOperator(7, NARY_UndOvr,  true,  true, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Conjunction_CenterSubSup: oOperator = this.Add_LargeOperator(7, NARY_UndOvr, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Conjunction_SubSup:       oOperator = this.Add_LargeOperator(7, NARY_SubSup, false, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Conjunction_CenterSub:    oOperator = this.Add_LargeOperator(7, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null); break;
        case c_oAscMathType.LargeOperator_Conjunction_Sub:          oOperator = this.Add_LargeOperator(7, NARY_SubSup,  true, false, Pr.ctrPrp, null, null, null); break;

        case c_oAscMathType.LargeOperator_Custom_1:
            var Sum = this.Add_LargeOperator(1, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, "k");
            var BaseMathContent = Sum.getBaseMathContent();
            var Delimiter = BaseMathContent.Add_Delimiter({ctrPrp : Pr.ctrPrp, column : 1}, 1, [null]);
            var DelimiterMathContent = Delimiter.getElementMathContent(0);
            DelimiterMathContent.Add_Fraction({ctrPrp: Pr.ctrPrp, type : NO_BAR_FRACTION}, "n", "k");
            break;

        case c_oAscMathType.LargeOperator_Custom_2:
            this.Add_LargeOperator(1, NARY_UndOvr, false, false, Pr.ctrPrp, null, "n", "i=0");
            break;

        case c_oAscMathType.LargeOperator_Custom_3:
            var Sum = this.Add_LargeOperator(1, NARY_UndOvr,  true, false, Pr.ctrPrp, null, null, null);
            var SubMathContent = Sum.getSubMathContent();
            SubMathContent.Add_EqArray({ctrPrp: Pr.ctrPrp, row : 2}, 2, ["0≤ i ≤ m", "0< j < n"]);
            var BaseMathContent = Sum.getBaseMathContent();
            BaseMathContent.Add_Text("P", this.Paragraph);
            BaseMathContent.Add_Delimiter({ctrPrp : Pr.ctrPrp, column : 1}, 1, ["i, j"]);
            break;

        case c_oAscMathType.LargeOperator_Custom_4:
            var Prod = this.Add_LargeOperator(2, NARY_UndOvr, false, false, Pr.ctrPrp, null, "n", "k=1");
            var BaseMathContent = Prod.getBaseMathContent();
            BaseMathContent.Add_Script(false, {ctrPrp: Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, "A", null, "k");
            break;

        case c_oAscMathType.LargeOperator_Custom_5:
            var Union = this.Add_LargeOperator(4, NARY_UndOvr, false, false, Pr.ctrPrp, null, "m", "n=1");
            var BaseMathContent = Union.getBaseMathContent();
            var Delimiter = BaseMathContent.Add_Delimiter({ctrPrp : Pr.ctrPrp, column : 1}, 1, [null]);
            BaseMathContent = Delimiter.getElementMathContent(0);
            BaseMathContent.Add_Script(false, {ctrPrp: Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, "X", null, "n");
            BaseMathContent.Add_Text(String.fromCharCode(8745), this.Paragraph);
            BaseMathContent.Add_Script(false, {ctrPrp: Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, "Y", null, "n");
            break;
    }

	if (oOperator && oSelectedContent)
		oOperator.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuBracket = function(Type, Pr, oSelectedContent)
{
	var oBracket  = null;
	var oFraction = null;

    switch(Type)
    {
        case c_oAscMathType.Bracket_Round                 : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  null,  null); break;
        case c_oAscMathType.Bracket_Square                : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    91,    93); break;
        case c_oAscMathType.Bracket_Curve                 : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,   125); break;
        case c_oAscMathType.Bracket_Angle                 : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 10216, 10217); break;
        case c_oAscMathType.Bracket_LowLim                : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 0x230A, 0x230B); break;
        case c_oAscMathType.Bracket_UppLim                : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 0x2308, 0x2309); break;
        case c_oAscMathType.Bracket_Line                  : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   124,   124); break;
        case c_oAscMathType.Bracket_LineDouble            : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  8214,  8214); break;
        case c_oAscMathType.Bracket_Square_OpenOpen       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    91,    91); break;
        case c_oAscMathType.Bracket_Square_CloseClose     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    93,    93); break;
        case c_oAscMathType.Bracket_Square_CloseOpen      : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    93,    91); break;
        case c_oAscMathType.Bracket_SquareDouble          : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 10214, 10215); break;

        case c_oAscMathType.Bracket_Round_Delimiter_2     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 2, [null, null], null, null); break;
        case c_oAscMathType.Bracket_Curve_Delimiter_2     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 2, [null, null], 123, 125); break;
        case c_oAscMathType.Bracket_Angle_Delimiter_2     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 2, [null, null], 10216, 10217); break;
        case c_oAscMathType.Bracket_Angle_Delimiter_3     : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 3, [null, null, null], 10216, 10217); break;

        case c_oAscMathType.Bracket_Round_OpenNone        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  null,    -1); break;
        case c_oAscMathType.Bracket_Round_NoneOpen        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,  null); break;
        case c_oAscMathType.Bracket_Square_OpenNone       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    91,    -1); break;
        case c_oAscMathType.Bracket_Square_NoneOpen       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,    93); break;
        case c_oAscMathType.Bracket_Curve_OpenNone        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,    -1); break;
        case c_oAscMathType.Bracket_Curve_NoneOpen        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,   125); break;
        case c_oAscMathType.Bracket_Angle_OpenNone        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 10216,    -1); break;
        case c_oAscMathType.Bracket_Angle_NoneOpen        : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1, 10217); break;
        case c_oAscMathType.Bracket_LowLim_OpenNone       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 0x230A,    -1); break;
        case c_oAscMathType.Bracket_LowLim_NoneNone       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,  0x230B); break;
        case c_oAscMathType.Bracket_UppLim_OpenNone       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 0x2308,    -1); break;
        case c_oAscMathType.Bracket_UppLim_NoneOpen       : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,  0x2309); break;
        case c_oAscMathType.Bracket_Line_OpenNone         : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   124,    -1); break;
        case c_oAscMathType.Bracket_Line_NoneOpen         : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,   124); break;
        case c_oAscMathType.Bracket_LineDouble_OpenNone   : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  8214,    -1); break;
        case c_oAscMathType.Bracket_LineDouble_NoneOpen   : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1,  8214); break;
        case c_oAscMathType.Bracket_SquareDouble_OpenNone : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null], 10214,    -1); break;
        case c_oAscMathType.Bracket_SquareDouble_NoneOpen : oBracket = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],    -1, 10215); break;

        case c_oAscMathType.Bracket_Custom_1:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,    -1);
            var BaseMathContent = Delimiter.getElementMathContent(0);
			oBracket = BaseMathContent.Add_EqArray({ctrPrp : Pr.ctrPrp, row : 2}, 2, [null, null]);
            break;

        case c_oAscMathType.Bracket_Custom_2:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,    -1);
            var BaseMathContent = Delimiter.getElementMathContent(0);
			oBracket = BaseMathContent.Add_EqArray({ctrPrp : Pr.ctrPrp, row : 3}, 3, [null, null, null]);
            break;

        case c_oAscMathType.Bracket_Custom_3:
			oFraction = this.Add_Fraction({ctrPrp : Pr.ctrPrp, type : NO_BAR_FRACTION}, null, null);
            break;

        case c_oAscMathType.Bracket_Custom_4:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  null,  null);
            var BaseMathContent = Delimiter.getElementMathContent(0);
			oFraction = BaseMathContent.Add_Fraction({ctrPrp : Pr.ctrPrp, type : NO_BAR_FRACTION}, null, null);
            break;

        case c_oAscMathType.Bracket_Custom_5:
            this.Add_Text("f", this.Paragraph);
            this.Add_DelimiterEx(Pr.ctrPrp, 1, ["x"],  null,  null);
            this.Add_Text("=", this.Paragraph);
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],   123,    -1);
            var BaseMathContent = Delimiter.getElementMathContent(0);
			oBracket = BaseMathContent.Add_EqArray({ctrPrp : Pr.ctrPrp, row : 2}, 2, ["-x,  &x<0", "x,  &x" + String.fromCharCode(8805) + "0"]);
            break;

        case c_oAscMathType.Bracket_Custom_6:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  null,  null);
            var BaseMathContent = Delimiter.getElementMathContent(0);
            BaseMathContent.Add_Fraction({ctrPrp : Pr.ctrPrp, type : NO_BAR_FRACTION}, "n", "k");
            break;

        case c_oAscMathType.Bracket_Custom_7:
            var Delimiter = this.Add_DelimiterEx(Pr.ctrPrp, 1, [null],  10216,  10217);
            var BaseMathContent = Delimiter.getElementMathContent(0);
            BaseMathContent.Add_Fraction({ctrPrp : Pr.ctrPrp, type : NO_BAR_FRACTION}, "n", "k");
            break;
    }

	if (oBracket && oSelectedContent)
		oBracket.getElementMathContent(0).private_FillSelectedContent(oSelectedContent);
    else if (oFraction && oSelectedContent)
		oFraction.getNumeratorMathContent().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuFunction = function(Type, Pr, oSelectedContent)
{
	var oFunction = null;
    switch(Type)
    {
        case c_oAscMathType.Function_Sin : oFunction = this.Add_Function(Pr, "sin", null); break;
        case c_oAscMathType.Function_Cos : oFunction = this.Add_Function(Pr, "cos", null); break;
        case c_oAscMathType.Function_Tan : oFunction = this.Add_Function(Pr, "tan", null); break;
        case c_oAscMathType.Function_Csc : oFunction = this.Add_Function(Pr, "csc", null); break;
        case c_oAscMathType.Function_Sec : oFunction = this.Add_Function(Pr, "sec", null); break;
        case c_oAscMathType.Function_Cot : oFunction = this.Add_Function(Pr, "cot", null); break;

        case c_oAscMathType.Function_1_Sin : oFunction = this.Add_Function_1(Pr, "sin", null); break;
        case c_oAscMathType.Function_1_Cos : oFunction = this.Add_Function_1(Pr, "cos", null); break;
        case c_oAscMathType.Function_1_Tan : oFunction = this.Add_Function_1(Pr, "tan", null); break;
        case c_oAscMathType.Function_1_Csc : oFunction = this.Add_Function_1(Pr, "csc", null); break;
        case c_oAscMathType.Function_1_Sec : oFunction = this.Add_Function_1(Pr, "sec", null); break;
        case c_oAscMathType.Function_1_Cot : oFunction = this.Add_Function_1(Pr, "cot", null); break;

        case c_oAscMathType.Function_Sinh : oFunction = this.Add_Function(Pr, "sinh", null); break;
        case c_oAscMathType.Function_Cosh : oFunction = this.Add_Function(Pr, "cosh", null); break;
        case c_oAscMathType.Function_Tanh : oFunction = this.Add_Function(Pr, "tanh", null); break;
        case c_oAscMathType.Function_Csch : oFunction = this.Add_Function(Pr, "csch", null); break;
        case c_oAscMathType.Function_Sech : oFunction = this.Add_Function(Pr, "sech", null); break;
        case c_oAscMathType.Function_Coth : oFunction = this.Add_Function(Pr, "coth", null); break;

        case c_oAscMathType.Function_1_Sinh : oFunction = this.Add_Function_1(Pr, "sinh", null); break;
        case c_oAscMathType.Function_1_Cosh : oFunction = this.Add_Function_1(Pr, "cosh", null); break;
        case c_oAscMathType.Function_1_Tanh : oFunction = this.Add_Function_1(Pr, "tanh", null); break;
        case c_oAscMathType.Function_1_Csch : oFunction = this.Add_Function_1(Pr, "csch", null); break;
        case c_oAscMathType.Function_1_Sech : oFunction = this.Add_Function_1(Pr, "sech", null); break;
        case c_oAscMathType.Function_1_Coth : oFunction = this.Add_Function_1(Pr, "coth", null); break;

        case c_oAscMathType.Function_Custom_1 : this.Add_Function(Pr, "sin", String.fromCharCode(952)); break;
        case c_oAscMathType.Function_Custom_2 : this.Add_Function(Pr, "cos", "2x"); break;
        case c_oAscMathType.Function_Custom_3 :
            var Theta = String.fromCharCode(952);
            this.Add_Function(Pr, "tan", Theta);
            this.Add_Text("=", this.Paragraph);
            var Fraction = this.Add_Fraction(Pr, null, null);
            var NumMathContent = Fraction.getNumeratorMathContent();
            var DenMathContent = Fraction.getDenominatorMathContent();
            NumMathContent.Add_Function(Pr, "sin", Theta);
            DenMathContent.Add_Function(Pr, "cos", Theta);
            break;
    }

	if (oFunction && oSelectedContent)
		oFunction.getArgument().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuAccent = function(Type, Pr, oSelectedContent)
{
	var oAccent = null;
    switch(Type)
    {
        case c_oAscMathType.Accent_Dot       : oAccent = this.Add_Accent(Pr.ctrPrp, 775, null); break;
        case c_oAscMathType.Accent_DDot      : oAccent = this.Add_Accent(Pr.ctrPrp, 776, null); break;
        case c_oAscMathType.Accent_DDDot     : oAccent = this.Add_Accent(Pr.ctrPrp, 8411, null); break;
        case c_oAscMathType.Accent_Hat       : oAccent = this.Add_Accent(Pr.ctrPrp, null, null); break;
        case c_oAscMathType.Accent_Check     : oAccent = this.Add_Accent(Pr.ctrPrp, 780, null); break;
        case c_oAscMathType.Accent_Accent    : oAccent = this.Add_Accent(Pr.ctrPrp, 769, null); break;
        case c_oAscMathType.Accent_Grave     : oAccent = this.Add_Accent(Pr.ctrPrp, 768, null); break;
        case c_oAscMathType.Accent_Smile     : oAccent = this.Add_Accent(Pr.ctrPrp, 774, null); break;
        case c_oAscMathType.Accent_Tilde     : oAccent = this.Add_Accent(Pr.ctrPrp, 771, null); break;
        case c_oAscMathType.Accent_Bar       : oAccent = this.Add_Accent(Pr.ctrPrp, 773, null); break;
        case c_oAscMathType.Accent_DoubleBar : oAccent = this.Add_Accent(Pr.ctrPrp, 831, null); break;

        case c_oAscMathType.Accent_CurveBracketTop : oAccent = this.Add_GroupCharacter({ctrPrp : Pr.ctrPrp, chr : 9182, pos : VJUST_TOP, vertJc : VJUST_BOT}, null ); break;
        case c_oAscMathType.Accent_CurveBracketBot : oAccent = this.Add_GroupCharacter({ctrPrp : Pr.ctrPrp}, null ); break;
        case c_oAscMathType.Accent_GroupTop:
            var Limit = this.Add_Limit({ctrPrp : Pr.ctrPrp, type : LIMIT_UP}, null, null);
            var MathContent = Limit.getFName();
			oAccent = MathContent.Add_GroupCharacter({ctrPrp : Pr.ctrPrp, chr : 9182, pos : VJUST_TOP, vertJc : VJUST_BOT}, null );
            break;

        case c_oAscMathType.Accent_GroupBot:
            var Limit = this.Add_Limit({ctrPrp : Pr.ctrPrp, type : LIMIT_LOW}, null, null);
            var MathContent = Limit.getFName();
			oAccent = MathContent.Add_GroupCharacter({ctrPrp : Pr.ctrPrp}, null );
            break;

        case c_oAscMathType.Accent_ArrowL  : oAccent = this.Add_Accent(Pr.ctrPrp, 8406, null); break;
        case c_oAscMathType.Accent_ArrowR  : oAccent = this.Add_Accent(Pr.ctrPrp, 8407, null); break;
        case c_oAscMathType.Accent_ArrowD  : oAccent = this.Add_Accent(Pr.ctrPrp, 8417, null); break;
        case c_oAscMathType.Accent_HarpoonL: oAccent = this.Add_Accent(Pr.ctrPrp, 8400, null); break;
        case c_oAscMathType.Accent_HarpoonR: oAccent = this.Add_Accent(Pr.ctrPrp, 8401, null); break;

        case c_oAscMathType.Accent_BorderBox :
			oAccent = this.Add_BorderBox(Pr, null);
            break;

        case c_oAscMathType.Accent_BorderBoxCustom :
            var BorderBox = this.Add_BorderBox(Pr, null);
            var MathContent = BorderBox.getBase();
            MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "a", "2", null);
            MathContent.Add_Text("=", this.Paragraph);
            MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "b", "2", null);
            MathContent.Add_Text("+", this.Paragraph);
            MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "c", "2", null);
            break;

        case c_oAscMathType.Accent_BarTop : oAccent = this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_TOP}, null); break;
        case c_oAscMathType.Accent_BarBot : oAccent = this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_BOT}, null); break;

        case c_oAscMathType.Accent_Custom_1 :
            this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_TOP}, "A");
            break;

        case c_oAscMathType.Accent_Custom_2 :
            this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_TOP}, "ABC");
            break;

        case c_oAscMathType.Accent_Custom_3 :
            this.Add_Bar({ctrPrp : Pr.ctrPrp, pos : LOCATION_TOP}, "x" + String.fromCharCode(8853) + "y");
            break;
    }

	if (oAccent && oSelectedContent)
		oAccent.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuLimitLog = function(Type, Pr, oSelectedContent)
{
	var oFunction = null;
    switch(Type)
    {
        case c_oAscMathType.LimitLog_LogBase:
			oFunction = this.Add_Function(Pr, null, null);
            var MathContent = oFunction.getFName();
            var Script = MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUBSCRIPT}, null, null, null);
            MathContent = Script.getBase();
            MathContent.Add_Text("log", this.Paragraph, STY_PLAIN);
            break;

        case c_oAscMathType.LimitLog_Log: oFunction = this.Add_Function(Pr, "log", null); break;
        case c_oAscMathType.LimitLog_Lim: oFunction = this.Add_FunctionWithLimit(Pr, "lim", null, null); break;
        case c_oAscMathType.LimitLog_Min: oFunction = this.Add_FunctionWithLimit(Pr, "min", null, null); break;
        case c_oAscMathType.LimitLog_Max: oFunction = this.Add_FunctionWithLimit(Pr, "max", null, null); break;
        case c_oAscMathType.LimitLog_Ln : oFunction = this.Add_Function(Pr, "ln", null); break;

        case c_oAscMathType.LimitLog_Custom_1:
            var Function = this.Add_FunctionWithLimit(Pr, "lim", "n" + String.fromCharCode(8594,8734), null);
            var MathContent = Function.getArgument();
            var Script = MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, null, "n", null);
            MathContent = Script.getBase();
            var Delimiter = MathContent.Add_Delimiter({ctrPrp : Pr.ctrPrp, column : 1}, 1, [null]);
            MathContent = Delimiter.getElementMathContent(0);
            MathContent.Add_Text("1+", this.Paragraph);
            MathContent.Add_Fraction({ctrPrp : Pr.ctrPrp}, "1", "n");
            break;

        case c_oAscMathType.LimitLog_Custom_2:
            var Function = this.Add_FunctionWithLimit(Pr, "max", "0" + String.fromCharCode(8804) + "x" + String.fromCharCode(8804) + "1", null);
            var MathContent = Function.getArgument();
            MathContent.Add_Text("x", this.Paragraph);
            var Script = MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "e", null, null);
            MathContent = Script.getUpperIterator();
            MathContent.Add_Text("-", this.Paragraph);
            MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, "x", "2", null);
            break;
    }

	if (oFunction && oSelectedContent)
		oFunction.getArgument().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuOperator = function(Type, Pr, oSelectedContent)
{
	var oAccent = null;
    switch(Type)
    {
        case c_oAscMathType.Operator_ColonEquals     : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, String.fromCharCode(0x2254)); break;
        case c_oAscMathType.Operator_EqualsEquals    : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, "=="); break;
        case c_oAscMathType.Operator_PlusEquals      : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, "+="); break;
        case c_oAscMathType.Operator_MinusEquals     : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, "-="); break;
        case c_oAscMathType.Operator_Definition      : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, String.fromCharCode(8797)); break;
        case c_oAscMathType.Operator_UnitOfMeasure   : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, String.fromCharCode(8798)); break;
        case c_oAscMathType.Operator_DeltaEquals     : this.Add_Box({ctrPrp : Pr.ctrPrp, opEmu : 1}, String.fromCharCode(8796)); break;
        case c_oAscMathType.Operator_ArrowL_Top      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8592, null); break;
        case c_oAscMathType.Operator_ArrowR_Top      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8594, null); break;
        case c_oAscMathType.Operator_ArrowL_Bot      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8592, null); break;
        case c_oAscMathType.Operator_ArrowR_Bot      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8594, null); break;
        case c_oAscMathType.Operator_DoubleArrowL_Top: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8656, null); break;
        case c_oAscMathType.Operator_DoubleArrowR_Top: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8658, null); break;
        case c_oAscMathType.Operator_DoubleArrowL_Bot: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8656, null); break;
        case c_oAscMathType.Operator_DoubleArrowR_Bot: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8658, null); break;
        case c_oAscMathType.Operator_ArrowD_Top      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8596, null); break;
        case c_oAscMathType.Operator_ArrowD_Bot      : oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8596, null); break;
        case c_oAscMathType.Operator_DoubleArrowD_Top: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_TOP, 8660, null); break;
        case c_oAscMathType.Operator_DoubleArrowD_Bot: oAccent = this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8660, null); break;
        case c_oAscMathType.Operator_Custom_1        : this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8594, "yields"); break;
        case c_oAscMathType.Operator_Custom_2        : this.Add_BoxWithGroupChar({ctrPrp : Pr.ctrPrp, opEmu : 1}, VJUST_BOT, 8594, String.fromCharCode(8710)); break;
    }

	if (oAccent && oSelectedContent)
		oAccent.getBase().private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuMatrix = function(Type, Pr, oSelectedContent)
{
	var oMatrix = null;
    switch(Type)
    {
        case c_oAscMathType.Matrix_1_2: oMatrix = this.Add_Matrix(Pr.ctrPrp, 1, 2, false, []); break;
        case c_oAscMathType.Matrix_2_1: oMatrix = this.Add_Matrix(Pr.ctrPrp, 2, 1, false, []); break;
        case c_oAscMathType.Matrix_1_3: oMatrix = this.Add_Matrix(Pr.ctrPrp, 1, 3, false, []); break;
        case c_oAscMathType.Matrix_3_1: oMatrix = this.Add_Matrix(Pr.ctrPrp, 3, 1, false, []); break;
        case c_oAscMathType.Matrix_2_2: oMatrix = this.Add_Matrix(Pr.ctrPrp, 2, 2, false, []); break;
        case c_oAscMathType.Matrix_2_3: oMatrix = this.Add_Matrix(Pr.ctrPrp, 2, 3, false, []); break;
        case c_oAscMathType.Matrix_3_2: oMatrix = this.Add_Matrix(Pr.ctrPrp, 3, 2, false, []); break;
        case c_oAscMathType.Matrix_3_3: oMatrix = this.Add_Matrix(Pr.ctrPrp, 3, 3, false, []); break;

        case c_oAscMathType.Matrix_Dots_Center   : this.Add_Text(String.fromCharCode(8943), this.Paragraph); break;
        case c_oAscMathType.Matrix_Dots_Baseline : this.Add_Text(String.fromCharCode(8230), this.Paragraph); break;
        case c_oAscMathType.Matrix_Dots_Vertical : this.Add_Text(String.fromCharCode(8942), this.Paragraph); break;
        case c_oAscMathType.Matrix_Dots_Diagonal : this.Add_Text(String.fromCharCode(8945), this.Paragraph); break;

        case c_oAscMathType.Matrix_Identity_2         : this.Add_Matrix(Pr.ctrPrp, 2, 2, false, ["1", "0", "0", "1"]); break;
        case c_oAscMathType.Matrix_Identity_2_NoZeros : this.Add_Matrix(Pr.ctrPrp, 2, 2, true, ["1", null, null, "1"]); break;
        case c_oAscMathType.Matrix_Identity_3         : this.Add_Matrix(Pr.ctrPrp, 3, 3, false, ["1", "0", "0", "0", "1", "0", "0", "0", "1"]); break;
        case c_oAscMathType.Matrix_Identity_3_NoZeros : this.Add_Matrix(Pr.ctrPrp, 3, 3, true, ["1", null, null, null, "1", null, null, null, "1"]); break;

        case c_oAscMathType.Matrix_2_2_RoundBracket  : oMatrix = this.Add_MatrixWithBrackets(null, null, Pr.ctrPrp, 2, 2, false, []); break;
        case c_oAscMathType.Matrix_2_2_SquareBracket : oMatrix = this.Add_MatrixWithBrackets(  91,   93, Pr.ctrPrp, 2, 2, false, []); break;
        case c_oAscMathType.Matrix_2_2_LineBracket   : oMatrix = this.Add_MatrixWithBrackets( 124,  124, Pr.ctrPrp, 2, 2, false, []); break;
        case c_oAscMathType.Matrix_2_2_DLineBracket  : oMatrix = this.Add_MatrixWithBrackets(8214, 8214, Pr.ctrPrp, 2, 2, false, []); break;

        case c_oAscMathType.Matrix_Flat_Round  : this.Add_MatrixWithBrackets(null, null, Pr.ctrPrp, 3, 3, false, [null, String.fromCharCode(8943), null, String.fromCharCode(8942), String.fromCharCode(8945), String.fromCharCode(8942), null, String.fromCharCode(8943), null]); break;
        case c_oAscMathType.Matrix_Flat_Square : this.Add_MatrixWithBrackets(  91,   93, Pr.ctrPrp, 3, 3, false, [null, String.fromCharCode(8943), null, String.fromCharCode(8942), String.fromCharCode(8945), String.fromCharCode(8942), null, String.fromCharCode(8943), null]); break;
    }

    if (oMatrix)
		oMatrix.getContentElement(0, 0).private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_LoadFromMenuDefaultText = function(Type, Pr, oSelectedContent)
{
	if (oSelectedContent)
		this.private_FillSelectedContent(oSelectedContent);
};
CMathContent.prototype.private_FillSelectedContent = function(oSelectedContent)
{
	if (oSelectedContent instanceof ParaMath)
	{
		oSelectedContent.Root.CopyTo(this, false);
	}
	else if (oSelectedContent)
	{
		this.Add_Text(oSelectedContent, this.Paragraph);
	}
};
CMathContent.prototype.Add_Element = function(Element)
{
    this.Internal_Content_Add(this.CurPos, Element, false);
    this.CurPos++;
};
CMathContent.prototype.Add_Text = function(sText, Paragraph, MathStyle)
{
    this.Paragraph = Paragraph;

    if (sText)
    {
        var MathRun = new ParaRun(this.Paragraph, true);

        for (var nCharPos = 0, nTextLen = sText.length; nCharPos < nTextLen; nCharPos++)
        {
            var oText = null;
            if (0x0026 == sText.charCodeAt(nCharPos))
                oText = new CMathAmp();
            else
            {
                oText = new CMathText(false);
                oText.addTxt(sText[nCharPos]);
            }
            MathRun.Add(oText, true);
        }

        MathRun.Set_RFont_ForMathRun();

        if (undefined !== MathStyle && null !== MathStyle)
            MathRun.Math_Apply_Style(MathStyle);

        this.Internal_Content_Add(this.CurPos, MathRun, false);
        this.CurPos++;
    }
};
CMathContent.prototype.Add_Symbol = function(Code, TextPr, MathPr)
{
    var MathRun = new ParaRun(this.Paragraph, true);

    var Symbol = new CMathText(false);
    Symbol.add(Code);
    MathRun.Add(Symbol, true);

    if(TextPr !== undefined)
        MathRun.Apply_Pr(TextPr);

    if(MathPr !== undefined)
        MathRun.Set_MathPr(MathPr);

    this.Internal_Content_Add(this.CurPos, MathRun, false);
    this.CurPos++;
};
CMathContent.prototype.Add_Fraction = function(Pr, NumText, DenText)
{
    var Fraction = new CFraction(Pr);
    this.Add_Element(Fraction);

    var DenMathContent = Fraction.getDenominatorMathContent();
    DenMathContent.Add_Text(DenText, this.Paragraph);

    var NumMathContent = Fraction.getNumeratorMathContent();
    NumMathContent.Add_Text(NumText, this.Paragraph);

    return Fraction;
};
CMathContent.prototype.Add_Script = function(bSubSup, Pr, BaseText, SupText, SubText)
{
    var Script = null;

    if (bSubSup)
        Script = new CDegreeSubSup(Pr);
    else
        Script = new CDegree(Pr);

    this.Add_Element(Script);

    var MathContent = Script.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    MathContent = Script.getUpperIterator();
    MathContent.Add_Text(SupText, this.Paragraph);

    MathContent = Script.getLowerIterator();
    MathContent.Add_Text(SubText, this.Paragraph);

    return Script;
};
CMathContent.prototype.Add_Radical = function(Pr, BaseText, DegreeText)
{
    var Radical = new CRadical(Pr);
    this.Add_Element(Radical);

    var MathContent = Radical.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    MathContent = Radical.getDegree();
    MathContent.Add_Text(DegreeText, this.Paragraph);

    return Radical;
};
CMathContent.prototype.Add_NAry = function(Pr, BaseText, SupText, SubText)
{
    var NAry = new CNary(Pr);
    this.Add_Element(NAry);

    var MathContent = NAry.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    MathContent = NAry.getSubMathContent();
    MathContent.Add_Text(SubText, this.Paragraph);

    MathContent = NAry.getSupMathContent();
    MathContent.Add_Text(SupText, this.Paragraph);

    return NAry;
};
CMathContent.prototype.Add_Integral = function(Dim, bOriented, limLoc, supHide, subHide, ctrPr, BaseText, SupText, SubText)
{
    var Pr = {ctrPrp : ctrPr};

    if (null !== limLoc)
        Pr.limLoc = limLoc;

    if (null !== supHide)
        Pr.supHide = supHide;

    if (null !== subHide)
        Pr.subHide = subHide;

    var chr = null;
    switch(Dim)
    {
        case 3: chr = (bOriented ? 8752 : 8749); break;
        case 2: chr = (bOriented ? 8751 : 8748); break;
        default:
        case 1: chr = (bOriented ? 8750 : null); break;
    }

    if (null !== chr)
        Pr.chr = chr;

    return this.Add_NAry(Pr, BaseText, SupText, SubText);
};
CMathContent.prototype.Add_LargeOperator = function(Type, limLoc, supHide, subHide, ctrPr, BaseText, SupText, SubText)
{
    var Pr = {ctrPrp : ctrPr};

    if (null !== limLoc)
        Pr.limLoc = limLoc;

    if (null !== supHide)
        Pr.supHide = supHide;

    if (null !== subHide)
        Pr.subHide = subHide;

    var chr = null;
    switch(Type)
    {
        default:
        case 1: chr = 8721; break;
        case 2: chr = 8719; break;
        case 3: chr = 8720; break;
        case 4: chr = 8899; break;
        case 5: chr = 8898; break;
        case 6: chr = 8897; break;
        case 7: chr = 8896; break;
    }

    if (null !== chr)
        Pr.chr = chr;

    return this.Add_NAry(Pr, BaseText, SupText, SubText);
};
CMathContent.prototype.Add_Delimiter = function(Pr, Count, aText)
{
    var Del = new CDelimiter(Pr);

    this.Add_Element(Del);

    for (var Index = 0; Index < Count; Index++)
    {
        var MathContent = Del.getElementMathContent(Index);
        MathContent.Add_Text(aText[Index], this.Paragraph);
    }

    return Del;
};
CMathContent.prototype.Add_DelimiterEx = function(ctrPr, Count, aText, begChr, endChr)
{
    var Pr =
    {
        ctrPrp : ctrPr,
        column : Count,
        begChr : begChr,
        endChr : endChr
    };

    return this.Add_Delimiter(Pr, Count, aText);
};
CMathContent.prototype.Add_EqArray = function(Pr, Count, aText)
{
    var EqArray = new CEqArray(Pr);

    this.Add_Element(EqArray);

    for (var Index = 0; Index < Count; Index++)
    {
        var MathContent = EqArray.getElementMathContent(Index);
        MathContent.Add_Text(aText[Index], this.Paragraph);
    }

    return EqArray;
};
CMathContent.prototype.Add_Box = function(Pr, BaseText)
{
    var Box = new CBox(Pr);
    this.Add_Element(Box);

    var MathContent = Box.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Box;
};
CMathContent.prototype.Add_BoxWithGroupChar = function(BoxPr, GroupPos, GroupChr, BaseText)
{
    var Box = this.Add_Box(BoxPr, null);
    var MathContent = Box.getBase();

    var oGroup = null;
    if (GroupPos === VJUST_TOP)
		oGroup = MathContent.Add_GroupCharacter({ctrPrp : BoxPr.ctrPrp, pos : GroupPos, chr : GroupChr}, BaseText);
    else
		oGroup = MathContent.Add_GroupCharacter({ctrPrp : BoxPr.ctrPrp, vertJc : GroupPos, chr : GroupChr}, BaseText);

    return oGroup;
};
CMathContent.prototype.Add_BorderBox = function(Pr, BaseText)
{
    var Box = new CBorderBox(Pr);
    this.Add_Element(Box);

    var MathContent = Box.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Box;
};
CMathContent.prototype.Add_Bar = function(Pr, BaseText)
{
    var Bar = new CBar(Pr);
    this.Add_Element(Bar);

    var MathContent = Bar.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Bar;
};
CMathContent.prototype.Add_Function = function(Pr, FName, BaseText)
{
    var MathFunc = new CMathFunc(Pr);
    this.Add_Element(MathFunc);

    var MathContent = MathFunc.getFName();
    MathContent.Add_Text(FName, this.Paragraph, STY_PLAIN);

    MathContent = MathFunc.getArgument();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return MathFunc;
};
CMathContent.prototype.Add_Function_1 = function(Pr, FName, BaseText)
{
    var MathFunc = new CMathFunc(Pr);
    this.Add_Element(MathFunc);

    MathFunc.SetParagraph(this.Paragraph);

    var MathContent = MathFunc.getFName();
    var Script = MathContent.Add_Script(false, {ctrPrp : Pr.ctrPrp, type : DEGREE_SUPERSCRIPT}, null, "-1", null);
    MathContent = Script.getBase();
    MathContent.Add_Text(FName, this.Paragraph, STY_PLAIN);

    MathContent = MathFunc.getArgument();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return MathFunc;
};
CMathContent.prototype.Add_FunctionWithLimit = function(Pr, FName, LimitText, BaseText)
{
    var MathFunc = new CMathFunc(Pr);
    this.Add_Element(MathFunc);

    MathFunc.SetParagraph(this.Paragraph);

    var MathContent = MathFunc.getFName();
    var Limit = MathContent.Add_Limit({ctrPrp : Pr.ctrPrp, type : LIMIT_LOW}, null, LimitText);
    MathContent = Limit.getFName();
    MathContent.Add_Text(FName, this.Paragraph, STY_PLAIN);

    MathContent = MathFunc.getArgument();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return MathFunc;
};
CMathContent.prototype.Add_Accent = function(ctrPr, chr, BaseText)
{
    var Pr =
    {
        ctrPrp : ctrPr,
        chr    : chr
    };
    var Accent = new CAccent(Pr);
    this.Add_Element(Accent);

    var MathContent = Accent.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Accent;
};
CMathContent.prototype.Add_GroupCharacter = function(Pr, BaseText)
{
    var Group = new CGroupCharacter(Pr);
    this.Add_Element(Group);

    var MathContent = Group.getBase();
    MathContent.Add_Text(BaseText, this.Paragraph);

    return Group;
};
CMathContent.prototype.Add_Limit = function(Pr, BaseText, LimitText)
{
    var Limit = new CLimit(Pr);
    this.Add_Element(Limit);

    var MathContent = Limit.getFName();
    MathContent.Add_Text(BaseText, this.Paragraph);

    MathContent = Limit.getIterator();
    MathContent.Add_Text(LimitText, this.Paragraph);

    return Limit;
};
CMathContent.prototype.Add_Matrix = function(ctrPr, RowsCount, ColsCount, plcHide, aText)
{
    var Pr =
    {
        ctrPrp  : ctrPr,
        row     : RowsCount,
        mcs     : [{count : ColsCount, mcJc : MCJC_CENTER}],
        plcHide : plcHide
    };

    var Matrix = new CMathMatrix(Pr);
    this.Add_Element(Matrix);

    for (var RowIndex = 0; RowIndex < RowsCount; RowIndex++)
    {
        for (var ColIndex = 0; ColIndex < ColsCount; ColIndex++)
        {
            var MathContent = Matrix.getContentElement(RowIndex, ColIndex);
            MathContent.Add_Text(aText[RowIndex * ColsCount + ColIndex], this.Paragraph);
        }
    }

    return Matrix;
};
CMathContent.prototype.Add_MatrixWithBrackets = function(begChr, endChr, ctrPr, RowsCount, ColsCount, plcHide, aText)
{
    var Delimiter = this.Add_DelimiterEx(ctrPr, 1, [null], begChr, endChr);
    var MathContent = Delimiter.getElementMathContent(0);
    return MathContent.Add_Matrix(ctrPr, RowsCount, ColsCount, plcHide, aText);
};
CMathContent.prototype.Recalculate_CurPos = function(_X, _Y, CurrentRun, _CurRange, _CurLine, _CurPage, UpdateCurPos, UpdateTarget, ReturnTarget)
{
	if (-1 === this.StartLine)
	{
		return {
			X         : 0,
			Y         : 0,
			Height    : 0,
			PageNum   : 0,
			Internal  : {Line : 0, Page : 0, Range : 0},
			Transform : null
		};
	}

    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var _EndPos = ( true === CurrentRun ? Math.min( EndPos, this.CurPos ) : EndPos );

    return this.Content[_EndPos].Recalculate_CurPos(_X, _Y, CurrentRun, _CurRange, _CurLine, _CurPage, UpdateCurPos, UpdateTarget, ReturnTarget);
};
CMathContent.prototype.GetCurrentParaPos = function()
{
	if (this.CurPos >= 0 && this.CurPos < this.Content.length)
		return this.Content[this.CurPos].GetCurrentParaPos();

	return new CParaPos(this.StartRange, this.StartLine, 0, 0);
};
CMathContent.prototype.Get_ParaContentPos = function(bSelection, bStart, ContentPos, bUseCorrection)
{
    if (true === bUseCorrection && true === bSelection)
	{
		var nPos = false !== bStart ? this.Selection.StartPos : this.Selection.EndPos;

		if (para_Math_Run !== this.Content[nPos].Type
			&& (this.Selection.StartPos !== this.Selection.EndPos || true !== this.Content[nPos].Is_InnerSelection())
			&& ((true === bStart && nPos > 0) || (true !== bStart && nPos < this.Content.length - 1)))
		{
			if (true === bStart && nPos > 0)
			{
				ContentPos.Add(nPos - 1);
				this.Content[nPos - 1].Get_EndPos(false, ContentPos, ContentPos.Get_Depth() + 1);
			}
			else
			{
				ContentPos.Add(nPos + 1);
				this.Content[nPos + 1].Get_StartPos(ContentPos, ContentPos.Get_Depth() + 1);
			}
		}
		else
		{
			ContentPos.Add(nPos);

			if (undefined !== this.Content[nPos])
				this.Content[nPos].Get_ParaContentPos(bSelection, bStart, ContentPos, bUseCorrection);
		}
	}
	else
	{
		var nPos = (true !== bSelection ? this.CurPos : (false !== bStart ? this.Selection.StartPos : this.Selection.EndPos));
		nPos = Math.max(0, Math.min(nPos, this.Content.length - 1));

		ContentPos.Add(nPos);

		if (undefined !== this.Content[nPos])
			this.Content[nPos].Get_ParaContentPos(bSelection, bStart, ContentPos, bUseCorrection);
	}
};
CMathContent.prototype.Set_ParaContentPos = function(ContentPos, Depth)
{
    var CurPos = ContentPos.Get(Depth);

    // Делаем такие проверки, потому что после совместного редактирования, позиция может остаться старой, а
    // контент измениться.
    if (undefined === CurPos || CurPos < 0)
    {
        this.CurPos = 0;
        if (this.Content[this.CurPos])
        	this.Content[this.CurPos].MoveCursorToStartPos();
    }
    else if (CurPos > this.Content.length - 1)
    {
        this.CurPos = this.Content.length - 1;
        if (this.Content[this.CurPos])
        	this.Content[this.CurPos].MoveCursorToEndPos(false);
    }
    else
    {
        this.CurPos = CurPos;
        if (this.Content[this.CurPos])
        	this.Content[this.CurPos].Set_ParaContentPos(ContentPos, Depth + 1);
    }
};
CMathContent.prototype.IsSelectionEmpty = function()
{
    if (true !== this.Selection.Use)
        return true;

    if (this.Selection.StartPos === this.Selection.EndPos)
        return this.Content[this.Selection.StartPos].IsSelectionEmpty();

    return false;
};
CMathContent.prototype.GetSelectContent = function(isAll)
{
	if (true === isAll)
	{
		return {Content : this, Start : 0, End : this.Content.length - 1};
	}
    else if (false === this.Selection.Use)
    {
        if (para_Math_Composition === this.Content[this.CurPos].Type)
            return this.Content[this.CurPos].GetSelectContent();
        else
            return {Content : this, Start : this.CurPos, End : this.CurPos};
    }
    else
    {
        var StartPos = this.Selection.StartPos;
        var EndPos   = this.Selection.EndPos;

        if (StartPos > EndPos)
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        if (StartPos === EndPos && para_Math_Composition === this.Content[StartPos].Type && true === this.Content[StartPos].Is_InnerSelection())
            return this.Content[StartPos].GetSelectContent();

        return {Content : this, Start : StartPos, End : EndPos};
    }
};
CMathContent.prototype.Get_LeftPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
    if (true !== this.ParentElement.Is_ContentUse(this))
        return false;

    if (false === UseContentPos && para_Math_Run === this.Content[this.Content.length - 1].Type)
    {
        // При переходе в новый контент встаем в его конец
        var CurPos = this.Content.length - 1;
        this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found = true;

        return true;
    }

    var CurPos = UseContentPos ? ContentPos.Get(Depth) : this.Content.length - 1;

    var bStepStart = false;
    if (CurPos > 0 || !this.Content[0].Cursor_Is_Start())
        bStepStart = true;

    this.Content[CurPos].Get_LeftPos(SearchPos, ContentPos, Depth + 1, UseContentPos);
    SearchPos.Pos.Update(CurPos, Depth);

    if (true === SearchPos.Found)
        return true;

    CurPos--;

    if (true === UseContentPos && para_Math_Composition === this.Content[CurPos + 1].Type)
    {
        // При выходе из формулы встаем в конец рана
        this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found = true;
        return true;
    }

    while (CurPos >= 0)
    {
        this.Content[CurPos].Get_LeftPos(SearchPos, ContentPos, Depth + 1, false);
        SearchPos.Pos.Update( CurPos, Depth );

        if (true === SearchPos.Found)
            return true;

        CurPos--;
    }

    if (true === bStepStart)
    {
        // Перед выходом из контента встаем в его начало
        this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(0, Depth);
        SearchPos.Found = true;

        return true;
    }

    return false;
};
CMathContent.prototype.Get_RightPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
    if (true !== this.ParentElement.Is_ContentUse(this))
        return false;

    if (false === UseContentPos && para_Math_Run === this.Content[0].Type)
    {
        // При переходе в новый контент встаем в его начало
        this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(0, Depth);
        SearchPos.Found = true;

        return true;
    }

    var CurPos = true === UseContentPos ? ContentPos.Get(Depth) : 0;

    var Count = this.Content.length;
    var bStepEnd = false;
    if (CurPos < Count - 1 || !this.Content[Count - 1].Cursor_Is_End())
        bStepEnd = true;

    this.Content[CurPos].Get_RightPos(SearchPos, ContentPos, Depth + 1, UseContentPos, StepEnd);
    SearchPos.Pos.Update( CurPos, Depth );

    if (true === SearchPos.Found)
        return true;

    CurPos++;

    if (true === UseContentPos && para_Math_Composition === this.Content[CurPos - 1].Type)
    {
        // При выходе из формулы встаем в начало рана
        this.Content[CurPos].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found = true;
        return true;
    }

    while (CurPos < Count)
    {
        this.Content[CurPos].Get_RightPos(SearchPos, ContentPos, Depth + 1, false, StepEnd);
        SearchPos.Pos.Update(CurPos, Depth);

        if (true === SearchPos.Found)
            return true;

        CurPos++;
    }

    if (true === bStepEnd)
    {
        // Перед выходом из контента встаем в его конец
        this.Content[Count - 1].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(Count - 1, Depth);
        SearchPos.Found = true;

        return true;
    }

    return false;
};
CMathContent.prototype.Get_WordStartPos = function(SearchPos, ContentPos, Depth, UseContentPos)
{
    if (true !== this.ParentElement.Is_ContentUse(this))
        return false;

    if (false === UseContentPos && para_Math_Run === this.Content[this.Content.length - 1].Type)
    {
        // При переходе в новый контент встаем в его конец
        var CurPos = this.Content.length - 1;
        this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found     = true;
        SearchPos.UpdatePos = true;
        return true;
    }

    var CurPos = true === UseContentPos ? ContentPos.Get(Depth) : this.Content.length - 1;

    var bStepStart = false;
    if (CurPos > 0 || !this.Content[0].Cursor_Is_Start())
        bStepStart = true;

    this.Content[CurPos].Get_WordStartPos(SearchPos, ContentPos, Depth + 1, UseContentPos);

    if (true === SearchPos.UpdatePos)
        SearchPos.Pos.Update( CurPos, Depth );

    if (true === SearchPos.Found)
        return;

    CurPos--;

    var bStepStartRun = false;
    if (true === UseContentPos && para_Math_Composition === this.Content[CurPos + 1].Type)
    {
        // При выходе из формулы встаем в конец рана
        this.Content[CurPos].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found     = true;
        SearchPos.UpdatePos = true;
        return true;
    }
    else if (para_Math_Run === this.Content[CurPos + 1].Type && true === SearchPos.Shift)
        bStepStartRun = true;

    while (CurPos >= 0)
    {
        if (true !== bStepStartRun || para_Math_Run === this.Content[CurPos].Type)
        {
            var OldUpdatePos = SearchPos.UpdatePos;

            this.Content[CurPos].Get_WordStartPos(SearchPos, ContentPos, Depth + 1, false);

            if (true === SearchPos.UpdatePos)
                SearchPos.Pos.Update(CurPos, Depth);
            else
                SearchPos.UpdatePos = OldUpdatePos;

            if (true === SearchPos.Found)
                return;

            if (true === SearchPos.Shift)
                bStepStartRun = true;
        }
        else
        {
            // Встаем в начало рана перед формулой
            this.Content[CurPos + 1].Get_StartPos(SearchPos.Pos, Depth + 1);
            SearchPos.Pos.Update(CurPos + 1, Depth);
            SearchPos.Found     = true;
            SearchPos.UpdatePos = true;
            return true;
        }
        CurPos--;
    }

    if (true === bStepStart)
    {
        // Перед выходом из контента встаем в его начало
        this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(0, Depth);
        SearchPos.Found     = true;
        SearchPos.UpdatePos = true;
        return true;
    }
};
CMathContent.prototype.Get_WordEndPos = function(SearchPos, ContentPos, Depth, UseContentPos, StepEnd)
{
    if (true !== this.ParentElement.Is_ContentUse(this))
        return false;

    if (false === UseContentPos && para_Math_Run === this.Content[0].Type)
    {
        // При переходе в новый контент встаем в его начало
        this.Content[0].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(0, Depth);
        SearchPos.Found     = true;
        SearchPos.UpdatePos = true;
        return true;
    }

    var CurPos = true === UseContentPos ? ContentPos.Get(Depth) : 0;

    var Count = this.Content.length;
    var bStepEnd = false;
    if (CurPos < Count - 1 || !this.Content[Count - 1].Cursor_Is_End())
        bStepEnd = true;

    this.Content[CurPos].Get_WordEndPos(SearchPos, ContentPos, Depth + 1, UseContentPos, StepEnd);

    if (true === SearchPos.UpdatePos)
        SearchPos.Pos.Update( CurPos, Depth);

    if (true === SearchPos.Found)
        return;

    CurPos++;

    var bStepEndRun = false;
    if (true === UseContentPos && para_Math_Composition === this.Content[CurPos - 1].Type)
    {
        // При выходе из формулы встаем в начало рана
        this.Content[CurPos].Get_StartPos(SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(CurPos, Depth);
        SearchPos.Found     = true;
        SearchPos.UpdatePos = true;
        return true;
    }
    else if (para_Math_Run === this.Content[CurPos - 1].Type && true === SearchPos.Shift)
        bStepEndRun = true;

    while (CurPos < Count)
    {
        if (true !== bStepEndRun || para_Math_Run === this.Content[CurPos].Type)
        {
            var OldUpdatePos = SearchPos.UpdatePos;

            this.Content[CurPos].Get_WordEndPos(SearchPos, ContentPos, Depth + 1, false, StepEnd);

            if (true === SearchPos.UpdatePos)
                SearchPos.Pos.Update(CurPos, Depth);
            else
                SearchPos.UpdatePos = OldUpdatePos;

            if (true === SearchPos.Found)
                return;

            if (true === SearchPos.Shift)
                bStepEndRun = true;
        }
        else
        {
            // Встаем в конец рана перед формулой
            this.Content[CurPos - 1].Get_EndPos(false, SearchPos.Pos, Depth + 1);
            SearchPos.Pos.Update(CurPos - 1, Depth);
            SearchPos.Found     = true;
            SearchPos.UpdatePos = true;
            return true;
        }

        CurPos++;
    }

    if (true === bStepEnd)
    {
        // Перед выходом из контента встаем в его конец
        this.Content[Count - 1].Get_EndPos(false, SearchPos.Pos, Depth + 1);
        SearchPos.Pos.Update(Count - 1, Depth);
        SearchPos.Found     = true;
        SearchPos.UpdatePos = true;
        return true;
    }
};
CMathContent.prototype.Get_StartPos = function(ContentPos, Depth)
{
    ContentPos.Update(0, Depth);
    this.Content[0].Get_StartPos(ContentPos, Depth + 1);
};
CMathContent.prototype.Get_EndPos = function(BehindEnd, ContentPos, Depth)
{
    var nLastPos = this.Content.length - 1;

    ContentPos.Update(nLastPos, Depth);
    if(undefined !== this.Content[nLastPos])
        this.Content[nLastPos].Get_EndPos(BehindEnd, ContentPos, Depth + 1);
};
CMathContent.prototype.Draw_HighLights = function(PDSH, bAll)
{
	if (!this.bRoot && this.Parent && true !== this.ParentElement.Is_ContentUse(this))
		return;

    var Bound = this.Get_LineBound(PDSH.Line, PDSH.Range);
    PDSH.X    = Bound.X;

    var CurLine  = PDSH.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PDSH.Range - this.StartRange : PDSH.Range );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var Y0 = PDSH.Y0,
        Y1 = PDSH.Y1;

    var FirstRunInRootNotShd = this.bRoot && this.Content.length > 0 && this.Content[StartPos].IsShade() == false;

    if(FirstRunInRootNotShd || this.bRoot == false)
    {
        Y0 = Bound.Y;
        Y1 = Bound.Y + Bound.H;
    }

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        PDSH.Y0 = Y0;
        PDSH.Y1 = Y1;

        if(bAll && this.Content[CurPos].Type == para_Math_Run)
            this.Content[CurPos].SelectAll();

        this.Content[CurPos].Draw_HighLights(PDSH, bAll);
    }
};
CMathContent.prototype.Draw_Lines = function(PDSL)
{
    var CurLine  = PDSL.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PDSL.Range - this.StartRange : PDSL.Range );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var Bound = this.Get_LineBound(PDSL.Line, PDSL.Range);

    var Baseline = Bound.Y + Bound.Asc;

    PDSL.Baseline = Baseline;
    PDSL.X        = Bound.X;

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Draw_Lines(PDSL);
        PDSL.Baseline = Baseline;
    }
};
CMathContent.prototype.RemoveSelection = function()
{
    var StartPos = this.Selection.StartPos;
    var EndPos   = this.Selection.EndPos;

    if (StartPos > EndPos)
    {
        StartPos = this.Selection.EndPos;
        EndPos   = this.Selection.StartPos;
    }

    StartPos = Math.max(0, StartPos);
    EndPos   = Math.min(this.Content.length - 1, EndPos);

    for (var nPos = StartPos; nPos <= EndPos; nPos++)
    {
        this.Content[nPos].RemoveSelection();
    }

    this.Selection.Use   = false;
    this.Selection.StartPos = 0;
    this.Selection.EndPos   = 0;
};
CMathContent.prototype.SelectAll = function(Direction)
{
    this.Selection.Use   = true;
    this.Selection.StartPos = 0;
    this.Selection.EndPos   = this.Content.length - 1;

    for (var nPos = 0, nCount = this.Content.length; nPos < nCount; nPos++)
    {
        this.Content[nPos].SelectAll(Direction);
    }
};
CMathContent.prototype.Selection_DrawRange = function(_CurLine, _CurRange, SelectionDraw)
{
    var SelectionStartPos = this.Selection.StartPos;
    var SelectionEndPos   = this.Selection.EndPos;

    if(SelectionStartPos > SelectionEndPos)
    {
        SelectionStartPos = this.Selection.EndPos;
        SelectionEndPos   = this.Selection.StartPos;
    }

    var SelectionUse = this.Selection.Use;

    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    if(this.bRoot == false)
    {
        var Bound = this.Get_LineBound(_CurLine, _CurRange);
        SelectionDraw.StartY   = Bound.Y;
        SelectionDraw.H        = Bound.H;
    }

    for(var CurPos = StartPos; CurPos <= EndPos; CurPos++)
    {
        var Item = this.Content[CurPos];

        var bSelectAll = SelectionUse && SelectionStartPos <= CurPos && CurPos <= SelectionEndPos && SelectionStartPos !== SelectionEndPos;

        if(Item.Type == para_Math_Composition && bSelectAll)
        {
            SelectionDraw.FindStart = false;
            SelectionDraw.W += Item.Get_Width(_CurLine, _CurRange);
        }
        else
        {
            Item.Selection_DrawRange( _CurLine, _CurRange, SelectionDraw );
        }
    }
};
CMathContent.prototype.SelectElementByPos = function(nPos)
{
	this.Selection.Use      = true;
	this.Selection.StartPos = nPos;
	this.Selection.EndPos   = nPos;

	this.Content[nPos].SelectAll();

	this.Correct_Selection();

	if (!this.bRoot)
		this.ParentElement.Select_MathContent(this);
	else
		this.ParaMath.bSelectionUse = true;
};
CMathContent.prototype.Select_Element = function(Element, bWhole)
{
    var nPos = -1;
    for(var nCurPos = 0, nCount = this.Content.length; nCurPos < nCount; nCurPos++)
    {
        if (this.Content[nCurPos] === Element)
        {
            nPos = nCurPos;
            break;
        }
    }

    if (-1 !== nPos)
    {
        this.Selection.Use   = true;
        this.Selection.StartPos = nPos;
        this.Selection.EndPos   = nPos;

        if (bWhole)
            this.Correct_Selection();

        if (!this.bRoot)
            this.ParentElement.Select_MathContent(this);
        else
            this.ParaMath.bSelectionUse = true;
    }
};
CMathContent.prototype.Correct_Selection = function()
{
    if (true !== this.Selection.Use)
        return;

    // Здесь мы делаем так, чтобы селект всегда начинался и заканчивался в ране.
    // Предполагается, что контент скорректирован верно до выполнения данной функции.

    var nContentLen = this.Content.length;
    var nStartPos = Math.max(0, Math.min(this.Selection.StartPos, nContentLen - 1));
    var nEndPos   = Math.max(0, Math.min(this.Selection.EndPos,   nContentLen - 1));

    if (nStartPos > nEndPos)
    {
        var nTemp = nStartPos;
        nStartPos = nEndPos;
        nEndPos   = nTemp;
    }

    var oStartElement = this.Content[nStartPos];
    if (para_Math_Run !== oStartElement.Type)
    {
        // Предыдущий элемент должен быть раном
        this.Selection.StartPos = nStartPos - 1;
        this.Content[this.Selection.StartPos].Set_SelectionAtEndPos();
    }

    var oEndElement = this.Content[nEndPos];
    if (para_Math_Run !== oEndElement.Type)
    {
        // Следующий элемент должен быть раном
        this.Selection.EndPos = nEndPos + 1;
        this.Content[this.Selection.EndPos].Set_SelectionAtStartPos();
    }
};
CMathContent.prototype.Create_FontMap = function(Map)
{
    // ArgSize компилируется только тогда, когда выставлены все ссылки на родительские классы
    for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; nIndex++)
        this.Content[nIndex].Create_FontMap(Map);
};
CMathContent.prototype.Get_AllFontNames = function(AllFonts)
{
    for (var nIndex = 0, nCount = this.Content.length; nIndex < nCount; nIndex++)
        this.Content[nIndex].Get_AllFontNames(AllFonts);
};
CMathContent.prototype.Selection_CheckParaContentPos = function(ContentPos, Depth, bStart, bEnd)
{
    var CurPos = ContentPos.Get(Depth);

    var bStartPos = this.Selection.StartPos,
        bEndPos   = this.Selection.EndPos;

    if(bStartPos > bEndPos)
    {
        var temp = bStartPos;
        bStartPos = bEndPos;
        bEndPos = temp;
    }

    if(bStartPos < CurPos)
        bStart = false;

    if(CurPos < bEndPos)
        bEnd = false;


    if(bStart === false && bEnd === false)
        return true;
    else if((bStartPos <= CurPos || bStart === false) && (CurPos <= bEndPos || bEnd === false))
        return this.Content[CurPos].Selection_CheckParaContentPos(ContentPos, Depth + 1, bStart, bEnd);

    return false;
};
CMathContent.prototype.Check_NearestPos = function(ParaNearPos, Depth)
{
    var HyperNearPos = new CParagraphElementNearPos();
    HyperNearPos.NearPos = ParaNearPos.NearPos;
    HyperNearPos.Depth   = Depth;

    this.NearPosArray.push(HyperNearPos);
    ParaNearPos.Classes.push(this);

    var CurPos = ParaNearPos.NearPos.ContentPos.Get(Depth);
    this.Content[CurPos].Check_NearestPos(ParaNearPos, Depth + 1);
};
CMathContent.prototype.Recalculate_Range = function(PRS, ParaPr, Depth)
{
    this.bOneLine = PRS.bMath_OneLine;

    // для неинлайн формул:
    // у операторов, находяхщихся на этом уровне (в Run) приоритет выше, чем у внутренних операторов (внутри мат объектов)
    // возможен только принудительный разрыв
    var bOnlyForcedBreak = PRS.bOnlyForcedBreak;
    var bNoOneBreakOperator = PRS.bNoOneBreakOperator;

    var CurLine  = PRS.Line - this.StartLine;
    var CurRange = ( 0 === CurLine ? PRS.Range - this.StartRange : PRS.Range );

    var ContentLen = this.Content.length;

    var RangeStartPos = this.protected_AddRange(CurLine, CurRange);
    var RangeEndPos   = ContentLen - 1;

    if(this.RecalcInfo.bEqArray)
        this.InfoPoints.SetDefault();

    var ascent = 0, descent = 0;
    this.size.width = 0;
    var bInline = this.ParaMath.Is_Inline();
    var bOperBefore = this.ParaMath.Is_BrkBinBefore(); // true - оператор находится в начале строки, false - оператор находится в конце строки

    // для внутристроковой формулы : начало формулы - начало нового слова
    if(this.bRoot && bInline && true == this.IsStartRange(PRS.Line, PRS.Range) && PRS.Ranges.length == 0)
    {
        PRS.Update_CurPos(0, Depth);
        PRS.Update_CurPos(0, Depth+1); // нулевой элемент всегда Run

        PRS.Set_LineBreakPos(0);

        if(PRS.Word == true)
        {
            PRS.FirstItemOnLine = false;
            PRS.X += PRS.SpaceLen + PRS.WordLen;

			PRS.Word               = false;
			PRS.EmptyLine          = false;
			PRS.EmptyLineWithBreak = false;
			PRS.SpaceLen           = 0;
			PRS.WordLen            = 0;
			PRS.TextOnLine         = true;

            PRS.XRange = PRS.X;
        }
    }

    var bCurInsideOper   = false;

    for(var Pos = RangeStartPos; Pos < ContentLen; Pos++)
    {
        var Item = this.Content[Pos],
            Type = Item.Type;

        PRS.bEndRunToContent = Pos == ContentLen - 1;

        // для однострочных мат объектов обновляем CurLine и CurRange, Run в этом случае не могут разбиваться на несколько строк
        if (this.bOneLine || (0 === Pos && 0 === CurLine && 0 === CurRange ) || Pos !== RangeStartPos)
            Item.Recalculate_Reset( PRS.Range, PRS.Line, PRS );

        PRS.Update_CurPos( Pos, Depth );

        if(this.bOneLine == true) // контент занимает всегда(!) одну строку
        {
            Item.Recalculate_Range_OneLine(PRS, ParaPr, Depth + 1);

            if(this.RecalcInfo.bEqArray && Type == para_Math_Composition)
                this.InfoPoints.ContentPoints.UpdatePoint(this.Content[Pos].size.width);

            this.size.width += Item.size.width;

            if(ascent < Item.size.ascent)
                ascent = Item.size.ascent;

            if(descent < Item.size.height - Item.size.ascent)
                descent = Item.size.height - Item.size.ascent;

            this.size.ascent = ascent;
            this.size.height = ascent + descent;
        }
        else if(bOnlyForcedBreak == true)
        {
            if(Type == para_Math_Run)
            {
                if(true === Item.Is_StartForcedBreakOperator())
                {
                    Item.Recalculate_Range(PRS, ParaPr, Depth + 1);
                }
                else
                {
                    Item.Recalculate_Range_OneLine(PRS, ParaPr, Depth + 1);
                    PRS.WordLen += Item.size.width;
                }
            }
            else
            {
                Item.Recalculate_Range(PRS, ParaPr, Depth + 1);

                if(Item.kind == MATH_BOX && true == Item.IsForcedBreak())
                {
                    this.private_ForceBreakBox(PRS, Item.size.width);
                }
            }

            if(PRS.NewRange == false && PRS.X + PRS.WordLen + PRS.SpaceLen > PRS.XEnd)
            {
                PRS.NewRange = true;
                PRS.MoveToLBP = true;
            }

        }
        else // контент может занимать несколько строк
        {
            // запомним позицию конца Run перед тем как зайдем во внутр мат объект (конечная позиция Run может измениться при пересчете внутр мат объекта)
            var _Depth      = PRS.PosEndRun.Depth;
            var PrevLastPos = PRS.PosEndRun.Get(_Depth-1),
                LastPos     = PRS.PosEndRun.Get(_Depth);

            var PrevWord = PRS.Word,
                MathFirstItem = PRS.MathFirstItem;

            PRS.bInsideOper = false;

            Item.Recalculate_Range(PRS, ParaPr, Depth + 1);

            if(Type == para_Math_Composition)
            {
                // перед мат объектом идет box break_operator и он не является первым элементом в строке
                if(Item.kind == MATH_BOX)
                {
                    PRS.MathFirstItem = MathFirstItem;
                    if(true == Item.IsForcedBreak())
                    {
                        this.private_ForceBreakBox(PRS, Item, _Depth, PrevLastPos, LastPos); // _Depth, PrevLastPos, LastPos запоминаем до пересчета, поэтому передаем эти параметры в функцию
                        PRS.bBreakBox = true;
                    }
                    else if(true == Item.IsOperatorEmulator())
                    {
                        this.private_BoxOperEmulator(PRS, Item, _Depth, PrevLastPos, LastPos); // _Depth, PrevLastPos, LastPos запоминаем до пересчета, поэтому передаем эти параметры в функцию
                        PRS.bBreakBox = true;
                    }
                    else
                    {
                        PRS.WordLen += Item.size.width;
                        PRS.Word = true;

                        if(PRS.X + PRS.SpaceLen + PRS.WordLen > PRS.XEnd)
                        {
                            if (PRS.FirstItemOnLine == false)
                            {
                                PRS.MoveToLBP = true;
                                PRS.NewRange = true;

                                this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);
                            }
                            else
                            {
                                PRS.bMathWordLarge = true;
                            }
                        }
                    }

                    PRS.MathFirstItem = false;
                }
                else
                {
                    // Слово не убирается в отрезке. Переносим слово в следующий отрезок
                    // FirstItemOnLine == false - слово оказалось не единственным элементом в промежутке, делаем перенос
                    if(PRS.X + PRS.SpaceLen + PRS.WordLen > PRS.XEnd)
                    {
                        if (PRS.FirstItemOnLine == false)
                        {
                            PRS.MoveToLBP = true;
                            PRS.NewRange = true;

                            this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);
                        }
                        else
                        {
                            PRS.bMathWordLarge = true;
                        }
                    }


                    // обновляем BreakPos на конец Run, т.к. внутри мат объекта BreakPos  может измениться на  if(true !== Word)
                    // обновляем только в том случае, если Word = false, иначе можем здесь перебить корректный LineBreakPos
                    if(bCurInsideOper == true && PrevWord == false && bOperBefore == false && bNoOneBreakOperator == false && PRS.bInsideOper == false)
                    {
                        // обновим : начало нового слова - конец предыдущего Run

                        PRS.Update_CurPos(PrevLastPos, _Depth-1);
                        PRS.Set_LineBreakPos(LastPos);

                        if(PRS.NewRange == true) // делаем возврат к предыдущему оператору
                            PRS.MoveToLBP = true;
                    }

                    PRS.Word = true;

                }

                if(PRS.NewRange == false && 0 === CurRange && 0 === CurLine)
                {
                    var PrevRecalcInfo = PRS.RunRecalcInfoLast,
                        NumberingAdd;

                    if(null === PrevRecalcInfo)
                        NumberingAdd = true;
                    else
                        NumberingAdd = PrevRecalcInfo.NumberingAdd;

                    if(NumberingAdd)
                    {
                        PRS.X = PRS.Recalculate_Numbering(Item, this, ParaPr, PRS.X);
                        PRS.RunRecalcInfoLast.NumberingAdd = false;
                        PRS.RunRecalcInfoLast.NumberingUse = true;
                        PRS.RunRecalcInfoLast.NumberingItem = PRS.Paragraph.Numbering;
                    }
                }
            }
            else
            {
                if(PRS.MathFirstItem == true && false == Item.IsEmptyRange(PRS.Line, PRS.Range))
                {
                    PRS.MathFirstItem = false;
                }

                var CheckWrapIndent = PRS.bFirstLine == true ? PRS.X - PRS.XRange > PRS.WrapIndent : true;

                if(PRS.bInsideOper == true && CheckWrapIndent == true)
                {
                    PRS.bOnlyForcedBreak = true;
                }
            }

            bCurInsideOper = bCurInsideOper || PRS.bInsideOper;
        }

        if ( true === PRS.NewRange )
        {
            RangeEndPos = Pos;
            break;
        }
    }

    PRS.bInsideOper = PRS.bInsideOper || bCurInsideOper;
    PRS.bOnlyForcedBreak = bOnlyForcedBreak;

    if ( Pos >= ContentLen )
    {
        RangeEndPos = Pos - 1;
    }

    var bSingleBarFraction = false;

    for(var Pos = 0; Pos < ContentLen; Pos++)
    {

        if(this.Content[Pos].kind == MATH_FRACTION && this.Content[Pos].Pr.type == BAR_FRACTION)
        {
            if(bSingleBarFraction)
            {
                bSingleBarFraction = false;
                break;
            }

            bSingleBarFraction = true;
        }
        else if( !(this.Content[Pos].Type == para_Math_Run && true == this.Content[Pos].Is_Empty()) ) // не пустой Run, другой мат объект
        {
            bSingleBarFraction = false;
            break;
        }
    }

    PRS.bSingleBarFraction = bSingleBarFraction;

    this.protected_FillRange(CurLine, CurRange, RangeStartPos, RangeEndPos);
};
CMathContent.prototype.private_ForceBreakBox = function(PRS, Box, _Depth, PrevLastPos, LastPos)
{
    var BoxLen = Box.size.width;

    if(true == PRS.MathFirstItem)
    {
        PRS.WordLen += BoxLen;
    }
    else
    {
        if(true === this.ParaMath.Is_BrkBinBefore())
        {
            PRS.X += PRS.SpaceLen + PRS.WordLen;
            PRS.Update_CurPos(PrevLastPos, _Depth-1);
            PRS.Set_LineBreakPos(LastPos);

            this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);

            PRS.MoveToLBP = true;
            PRS.NewRange = true;
        }
        else
        {
            PRS.Word         = false;
            PRS.bForcedBreak = true;
            PRS.WordLen     += BoxLen;
        }
    }
};
CMathContent.prototype.private_BoxOperEmulator = function(PRS, Box, _Depth, PrevLastPos, LastPos)
{
    var BoxLen      = Box.size.width,
        BoxGapRight = Box.GapRight;

    var CheckWrapIndent = PRS.bFirstLine == true ? PRS.X - PRS.XRange > PRS.WrapIndent : true;
    var bOperBefore = this.ParaMath.Is_BrkBinBefore() == true;

    var bOnlyForcedBreakBefore = bOperBefore == true && PRS.MathFirstItem == false,
        bOnlyforcedBreakAfter  = bOperBefore == false;

    if(CheckWrapIndent == true && (bOnlyForcedBreakBefore || bOnlyforcedBreakAfter))
    {
        PRS.bOnlyForcedBreak = true;
    }

    var bOverXEnd;

    if(bOperBefore) // оператор находится в начале строки
    {
        bOverXEnd = PRS.X + PRS.WordLen + PRS.SpaceLen > PRS.XEnd;

        if(true == PRS.MathFirstItem)
        {
            PRS.WordLen += PRS.SpaceLen + PRS.WordLen + BoxLen;
        }
        else if(PRS.FirstItemOnLine == false && bOverXEnd)
        {
            PRS.MoveToLBP = true;
            PRS.NewRange = true;

            this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);
        }
        else
        {
            PRS.X += PRS.SpaceLen + PRS.WordLen;
            // обновим : начало нового слова - конец предыдущего Run
            PRS.bInsideOper = true;
            PRS.FirstItemOnLine = false;

            PRS.Update_CurPos(PrevLastPos, _Depth-1);
            PRS.Set_LineBreakPos(LastPos);

            PRS.SpaceLen = BoxLen;

            PRS.WordLen = 0;
            PRS.Word = true;
        }
    }
    else
    {
        bOverXEnd = PRS.X + PRS.SpaceLen + PRS.WordLen + BoxLen - BoxGapRight > PRS.XEnd;

        PRS.OperGapRight = BoxGapRight;

        // Слово не убирается в отрезке. Переносим слово в следующий отрезок
        // FirstItemOnLine == false - слово оказалось не единственным элементом в промежутке, делаем перенос
        if(PRS.FirstItemOnLine == false && bOverXEnd)
        {
            PRS.MoveToLBP = true;
            PRS.NewRange = true;

            this.ParaMath.UpdateWidthLine(PRS, PRS.X - PRS.XRange);
        }
        else
        {
            PRS.bInsideOper = true;
        }

        PRS.X += PRS.SpaceLen + PRS.WordLen + BoxLen;
        PRS.SpaceLen = 0;
        PRS.WordLen = 0;

        PRS.Word = false;
        PRS.FirstItemOnLine = false;
    }
};
CMathContent.prototype.Math_Set_EmptyRange = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var RangeStartPos = this.protected_AddRange(CurLine, CurRange);
    var RangeEndPos   = RangeStartPos;

    this.protected_FillRange(CurLine, CurRange, RangeStartPos, RangeEndPos);

    this.Content[RangeStartPos].Math_Set_EmptyRange(_CurLine, _CurRange);
};
CMathContent.prototype.Recalculate_Reset = function(StartRange, StartLine, PRS)
{
    var bNotUpdate = PRS !== null && PRS!== undefined && PRS.bFastRecalculate == true;
    if(bNotUpdate == false)
    {
        this.StartLine   = StartLine;
        this.StartRange  = StartRange;

        if(this.Content.length > 0)
            this.Content[0].Recalculate_Reset(StartRange, StartLine, PRS);

        this.protected_ClearLines();
    }
};
CMathContent.prototype.IsEmptyRange = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    var bEmpty = true;

    if(StartPos == EndPos)
    {
        bEmpty = this.Content[StartPos].IsEmptyRange(_CurLine, _CurRange);
    }
    else
    {
        var Pos = StartPos;

        while(Pos <= EndPos)
        {
            if(false == this.Content[Pos].IsEmptyRange(_CurLine, _CurRange))
            {
                bEmpty = false;
                break;
            }
            Pos++;
        }
    }

    return bEmpty;
};
CMathContent.prototype.Displace_BreakOperator = function(isForward, bBrkBefore, CountOperators)
{
	var Pos = this.CurPos;

	if (this.Content[Pos].Type == para_Math_Run)
	{
		var bApplyBreak = this.Content[Pos].Displace_BreakOperator(isForward, bBrkBefore, CountOperators);
		var NewPos      = bBrkBefore ? Pos + 1 : Pos - 1;

		if (this.Content[NewPos] && bApplyBreak == false && (this.Content[NewPos].Type == para_Math_Run || this.Content[NewPos].kind == MATH_BOX))
		{
			this.Content[NewPos].Displace_BreakOperator(isForward, bBrkBefore, CountOperators);
		}
	}
	else
	{
		this.Content[Pos].Displace_BreakOperator(isForward, bBrkBefore, CountOperators);
	}
};
CMathContent.prototype.Recalculate_Range_Width = function(PRSC, _CurLine, _CurRange)
{
    var RangeW = PRSC.Range.W;
    var CurLine = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for ( var CurPos = StartPos; CurPos <= EndPos; CurPos++ )
    {
        this.Content[CurPos].Recalculate_Range_Width( PRSC, _CurLine, _CurRange );
    }

    this.Bounds.SetWidth(CurLine, CurRange, PRSC.Range.W - RangeW);
};
CMathContent.prototype.RecalculateMinMaxContentWidth = function(MinMax)
{
    if(this.RecalcInfo.bEqArray)
        this.InfoPoints.SetDefault();

    var ascent = 0, descent = 0;
    this.size.width = 0;

    var Lng = this.Content.length;

    for(var Pos = 0; Pos < Lng; Pos++)
    {
        var Item = this.Content[Pos],
            Type = Item.Type;

        if(MinMax.bMath_OneLine)
        {
            if(Type == para_Math_Run)
                Item.Math_RecalculateContent();
            else
                Item.RecalculateMinMaxContentWidth(MinMax);

            if(this.RecalcInfo.bEqArray && Type == para_Math_Composition)
                this.InfoPoints.ContentPoints.UpdatePoint(this.Content[Pos].size.width);

            this.size.width += Item.size.width;

            if(ascent < Item.size.ascent)
                ascent = Item.size.ascent;

            if(descent < Item.size.height - Item.size.ascent)
                descent = Item.size.height - Item.size.ascent;

            this.size.ascent = ascent;
            this.size.height = ascent + descent;
        }
        else
        {
            Item.RecalculateMinMaxContentWidth(MinMax);
        }
    }

};
CMathContent.prototype.Recalculate_LineMetrics = function(PRS, ParaPr, _CurLine, _CurRange, ContentMetrics)
{
    var CurLine = _CurLine - this.StartLine;
    var CurRange = (0 === CurLine ? _CurRange - this.StartRange : _CurRange);

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    if(PRS.bFastRecalculate === false)
    {
        this.Bounds.Reset(CurLine, CurRange);
    }

    var NewContentMetrics = new CMathBoundsMeasures();

    for(var Pos = StartPos; Pos <= EndPos; Pos++)
    {
        var Item = this.Content[Pos];
        Item.Recalculate_LineMetrics(PRS, ParaPr, _CurLine, _CurRange, NewContentMetrics);
    }

    this.Bounds.UpdateMetrics(CurLine, CurRange, NewContentMetrics);

    ContentMetrics.UpdateMetrics(NewContentMetrics);
};
CMathContent.prototype.Math_UpdateLineMetrics = function(PRS, ParaPr)
{
    this.Content[0].Math_UpdateLineMetrics(PRS, ParaPr); // this.Content[0] - Run
};
CMathContent.prototype.UpdateOperators = function(_CurLine, _CurRange, bEmptyGapLeft, bEmptyGapRight)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var EndPos   = this.protected_GetRangeEndPos(CurLine, CurRange);

    for(var Pos = StartPos; Pos <= EndPos; Pos++)
    {
        var _bEmptyGapLeft  = bEmptyGapLeft  && Pos == StartPos,
            _bEmptyGapRight = bEmptyGapRight && Pos == EndPos;

        this.Content[Pos].UpdateOperators(_CurLine, _CurRange, _bEmptyGapLeft, _bEmptyGapRight);
    }
};
CMathContent.prototype.Get_Bounds = function()
{
    return this.Bounds.Get_Bounds();
};
CMathContent.prototype.Get_LineBound = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.Get_LineBound(CurLine, CurRange);
};
CMathContent.prototype.GetPos = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.GetPos(CurLine, CurRange);
};
CMathContent.prototype.Get_Width = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.Get_Width(CurLine, CurRange);
};
CMathContent.prototype.GetAscent = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.GetAscent(CurLine, CurRange);
};
CMathContent.prototype.GetDescent = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine,
        CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    return this.Bounds.GetDescent(CurLine, CurRange);
};
CMathContent.prototype.Get_StartRangePos = function(_CurLine, _CurRange, SearchPos, Depth, bStartPos)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var StartPos = this.protected_GetRangeStartPos(CurLine, CurRange);
    var CurPos = this.CurPos;
    var Result;

    var bStart = this.bRoot ? CurRange == 0 : bStartPos;

    if(this.Content[CurPos].Type == para_Math_Composition && bStartPos !== true)
    {
        bStart = bStart && CurPos == StartPos;
        Result = this.Content[CurPos].Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bStart); // пытаемся встать в начало внутреннего контента

        if ( true === Result )
        {
            SearchPos.Pos.Update( CurPos, Depth );
        }
        else if(this.bRoot && CurPos !== StartPos)
        {
            if(this.Content[StartPos].Type == para_Math_Composition)
                Result = this.Content[StartPos].Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, true);
            else
                Result = this.Content[StartPos].Math_Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, true);

            if ( true === Result )
                SearchPos.Pos.Update(StartPos, Depth);
        }
    }
    else
    {
        if(this.bRoot && CurLine == 0)
        {
            bStart = bStart && StartPos < CurPos;
        }

        if(this.Content[StartPos].Type == para_Math_Composition)
            Result = this.Content[StartPos].Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bStart); // может произойти, если мат объект разбивается на строки
        else
            Result = this.Content[StartPos].Math_Get_StartRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bStart);

        if ( true === Result )
            SearchPos.Pos.Update(StartPos, Depth );
    }

    return Result;
};
CMathContent.prototype.Get_EndRangePos = function(_CurLine, _CurRange, SearchPos, Depth, bEndPos)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var EndPos = this.protected_GetRangeEndPos(CurLine, CurRange);
    var CurPos = this.CurPos;
    var Result;

    var bLastRange = CurRange == this.protected_GetRangesCount(CurLine) - 1;

    var bEnd = this.bRoot ? bLastRange : bEndPos;

    if(this.Content[CurPos].Type == para_Math_Composition && bEndPos !== true)
    {
        Result = this.Content[CurPos].Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1); // пытаемся встать в конец внутреннего контента

        if ( true === Result )
        {
            SearchPos.Pos.Update( CurPos, Depth );
        }
        else if(this.bRoot && CurPos !== EndPos)
        {
            if(this.Content[EndPos].Type == para_Math_Composition)
                Result = this.Content[EndPos].Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, true);
            else
                Result = this.Content[EndPos].Math_Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, true);

            if ( true === Result )
                SearchPos.Pos.Update(EndPos, Depth );
        }
    }
    else
    {
        bEnd = this.bRoot && bLastRange ? true /*иначе после того как встанем после формулы перед знаком параграфа, на следующем End встанем перед формулой*/ : bEnd || CurPos < EndPos;

        if(this.Content[EndPos].Type == para_Math_Composition)
            Result = this.Content[EndPos].Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bEnd); // может произойти, если мат объект разбивается на строки
        else
            Result = this.Content[EndPos].Math_Get_EndRangePos(_CurLine, _CurRange, SearchPos, Depth + 1, bEnd);

        if ( true === Result )
            SearchPos.Pos.Update(EndPos, Depth );

    }

    return Result;
};
CMathContent.prototype.Math_Is_End = function(_CurLine, _CurRange)
{
    var CurLine  = _CurLine - this.StartLine;
    var CurRange = ( 0 === CurLine ? _CurRange - this.StartRange : _CurRange );

    var EndPos = this.protected_GetRangeEndPos(CurLine, CurRange);

    var Len = this.Content.length;

    var result = false;

    if(EndPos == Len - 1)
    {
        result = this.Content[Len - 1].Math_Is_End(_CurLine, _CurRange);
    }

    return result;
};
CMathContent.prototype.Get_AlignBrk = function(_CurLine, bBrkBefore)
{
    var AlnAt = null;

    var CurLine  = _CurLine - this.StartLine;
    var RangesCount = this.protected_GetRangesCount(CurLine - 1);
    var EndPos = this.protected_GetRangeEndPos(CurLine - 1, RangesCount - 1);

    if(CurLine !== 0) // получаем смещение до расчета Recalculate_Range
    {
        var bEndRun  = this.Content[EndPos].Type == para_Math_Run && true == this.Content[EndPos].Math_Is_End(_CurLine - 1, RangesCount - 1),
            bNextBox = EndPos < this.Content.length - 1 && this.Content[EndPos + 1].kind == MATH_BOX;

        var bCheckNextBox = bEndRun == true && bNextBox == true && bBrkBefore == true;

        var bRunEmptyRange = this.Content[EndPos].Type == para_Math_Run && this.Content[EndPos].IsEmptyRange(_CurLine - 1, RangesCount - 1),
            bPrevBox = EndPos > 0 && this.Content[EndPos - 1].kind == MATH_BOX;

        var bCheckPrevNextBox = bRunEmptyRange == true && bPrevBox == true && bBrkBefore == false;

        if(bCheckPrevNextBox)
        {
            AlnAt = this.Content[EndPos - 1].Get_AlignBrk(_CurLine, bBrkBefore);
        }
        else if(bCheckNextBox)
        {
            AlnAt = this.Content[EndPos + 1].Get_AlignBrk(_CurLine, bBrkBefore);
        }
        else
        {
            AlnAt = this.Content[EndPos].Get_AlignBrk(_CurLine, bBrkBefore);
        }
    }

    return AlnAt;
};
CMathContent.prototype.IsStartRange = function(Line, Range)
{
    return Line - this.StartLine == 0 && Range - this.StartRange == 0;
};
CMathContent.prototype.IsStartLine = function(Line)
{
    return Line == this.StartLine;
};
CMathContent.prototype.GetSelectDirection = function()
{
    if (true !== this.Selection.Use)
        return 0;

    if (this.Selection.StartPos < this.Selection.EndPos)
        return 1;
    else if (this.Selection.StartPos > this.Selection.EndPos)
        return -1;

    return this.Content[this.Selection.StartPos].GetSelectDirection();
};
CMathContent.prototype.MoveCursorToStartPos = function()
{
    this.CurPos = 0;
    this.Content[0].MoveCursorToStartPos();
};
CMathContent.prototype.MoveCursorToEndPos = function(SelectFromEnd)
{
    this.CurPos = this.Content.length - 1;
    this.Content[this.CurPos].MoveCursorToEndPos(SelectFromEnd);
};
CMathContent.prototype.Check_Composition = function()
{
    var Pos = this.private_FindCurrentPosInContent();

    return Pos !== null && this.Content[Pos].Type == para_Math_Composition;
};
CMathContent.prototype.Can_ModifyForcedBreak = function(Pr)
{
    var Pos = this.private_GetPosRunForForcedBreak();

    if(Pos !== null && this.bOneLine == false)
    {
        var bBreakOperator      = this.Content[Pos].Check_ForcedBreak();
        var CurrentRun          = this.Content[Pos];
        var bCanCheckNearsRun   = bBreakOperator == false && false == CurrentRun.IsSelectionUse();
        var bPrevItem           = bCanCheckNearsRun && Pos > 0 && true == CurrentRun.Cursor_Is_Start(),
            bNextItem           = bCanCheckNearsRun && Pos < this.Content.length - 1 && true == CurrentRun.Cursor_Is_End();

        var bPrevRun = bPrevItem &&  this.Content[Pos - 1].Type == para_Math_Run,
            bNextRun = bNextItem &&  this.Content[Pos + 1].Type == para_Math_Run;

        if(bBreakOperator)
        {
            this.Content[Pos].Math_Can_ModidyForcedBreak(Pr);
        }
        else if(bPrevRun)
        {
            this.Content[Pos - 1].Math_Can_ModidyForcedBreak(Pr, true, false);
        }
        else if(bNextRun)
        {
            this.Content[Pos + 1].Math_Can_ModidyForcedBreak(Pr, false, true);
        }
    }
};
CMathContent.prototype.private_GetPosRunForForcedBreak = function()
{
    var Pos = null;

    if(true === this.Selection.Use)
    {
        var StartPos = this.Selection.StartPos,
            EndPos   = this.Selection.EndPos;

        if ( StartPos > EndPos )
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }


        var bHaveSelectedItem = false;
        for(var CurPos = StartPos; CurPos <= EndPos; CurPos++)
        {
            var Item = this.Content[CurPos];
            var bSelect     = true !== Item.IsSelectionEmpty(),
                bSelectRun  = bSelect == true && Item.Type == para_Math_Run,
                bSelectComp = bSelect == true && Item.Type == para_Math_Composition;
            var bSelectManyRuns = bSelectRun && bHaveSelectedItem;

            if(bSelectComp || bSelectManyRuns)
            {
                Pos = null;
                break;
            }

            if(bSelectRun)
            {
                bHaveSelectedItem = true;
                Pos = CurPos;
            }
        }
    }
    else
    {
        Pos = this.CurPos;
    }

    return Pos;
};
CMathContent.prototype.private_FindCurrentPosInContent = function()
{
    var Pos = null;

    if(true === this.Selection.Use)
    {
        var StartPos = this.Selection.StartPos,
            EndPos   = this.Selection.EndPos;

        if ( StartPos > EndPos )
        {
            StartPos = this.Selection.EndPos;
            EndPos   = this.Selection.StartPos;
        }

        var bComposition = false;

        for(var CurPos = StartPos; CurPos <= EndPos; CurPos++)
        {
            var Item = this.Content[CurPos];

            if(Item.Type == para_Math_Run && true !== Item.IsSelectionEmpty())
            {
                Pos = bComposition == true ? null : CurPos;
                break;
            }
            else if(Item.Type ==  para_Math_Composition)
            {
                if(bComposition == true)
                {
                    Pos = null;
                    break;
                }
                else
                {
                    Pos = CurPos;
                    bComposition = true;
                }
            }
        }
    }
    else
    {
        Pos = this.CurPos;
    }

    return Pos;
};
CMathContent.prototype.Is_CurrentContent = function()
{
    var Pos = this.private_FindCurrentPosInContent();

    return Pos == null || this.Content[Pos].Type == para_Math_Run;
};
CMathContent.prototype.Set_MenuProps = function(Props)
{
    var Pos = this.private_FindCurrentPosInContent();

    if(true == this.Is_CurrentContent())
    {
        this.Apply_MenuProps(Props, Pos);
    }
    else if(false == this.private_IsMenuPropsForContent(Props.Action) &&  true == this.Content[Pos].Can_ApplyMenuPropsToObject())
    {
        // не нужно проходиться по вложенным элементам
        // 1. уже применили изменения, продожать нет необходимости
        // 2. потому что могут совпать типы текущего элемента и вложенного и тогда изменения применятся к обоим элементам
        if(false === this.Delete_ItemToContentThroughInterface(Props, Pos)) // try to delete
        {
            this.Content[Pos].Apply_MenuProps(Props);
        }
    }
    else
    {
        this.Content[Pos].Set_MenuProps(Props);
    }
};
CMathContent.prototype.Apply_MenuProps = function(Props, Pos)
{
	var ArgSize, NewArgSize;

	if (Props.Action & c_oMathMenuAction.IncreaseArgumentSize)
	{
		if (true === this.Parent.Can_ModifyArgSize() && true == this.Compiled_ArgSz.Can_Increase() && true == this.ArgSize.Can_SimpleIncrease())
		{
			ArgSize    = this.ArgSize.GetValue();
			NewArgSize = this.ArgSize.Increase();

			History.Add(new CChangesMathContentArgSize(this, ArgSize, NewArgSize));
			this.Recalc_RunsCompiledPr();
		}
	}
	else if (Props.Action & c_oMathMenuAction.DecreaseArgumentSize)
	{
		if (true === this.Parent.Can_ModifyArgSize() && true == this.Compiled_ArgSz.Can_Decrease() && true == this.ArgSize.Can_Decrease())
		{
			ArgSize    = this.ArgSize.GetValue();
			NewArgSize = this.ArgSize.Decrease();

			History.Add(new CChangesMathContentArgSize(this, ArgSize, NewArgSize));
			this.Recalc_RunsCompiledPr();
		}
	}

	var Run;

	if (Pos !== null && Props.Action & c_oMathMenuAction.InsertForcedBreak)
	{
		Run = this.private_Get_RunForForcedBreak(Pos);
		Run.Set_MathForcedBreak(true);

	}
	else if (Pos !== null && Props.Action & c_oMathMenuAction.DeleteForcedBreak)
	{
		Run = this.private_Get_RunForForcedBreak(Pos);
		Run.Set_MathForcedBreak(false);
	}
};
CMathContent.prototype.private_Get_RunForForcedBreak = function(Pos)
{
    var CurrentRun          = this.Content[Pos];
    var bCurrentForcedBreak = this.Content[Pos].Type == para_Math_Run && true == CurrentRun.Check_ForcedBreak(),
        bPrevForcedBreak    = Pos > 0 && true == CurrentRun.Cursor_Is_Start(),
        bNextForcedBreak    = Pos < this.Content.length && true == CurrentRun.Cursor_Is_End();

    var Run = null;

    if(bCurrentForcedBreak)
    {
        Run = this.Content[Pos];

        var NewRun = Run.Math_SplitRunForcedBreak();

        if(NewRun !== null)
        {
            this.Internal_Content_Add(Pos+1, NewRun, true);
            Run = NewRun;
        }
    }
    else if(bPrevForcedBreak)
    {
        Run = this.Content[Pos - 1];
    }
    else if(bNextForcedBreak)
    {
        Run = this.Content[Pos + 1];
    }

    return Run;
};
CMathContent.prototype.Delete_ItemToContentThroughInterface = function(Props, Pos)
{
    var bDelete = false;
    var Item = this.Content[Pos];

    if(Item.kind == MATH_DEGREESubSup && Item.Pr.type == DEGREE_SubSup && Props.Type == Asc.c_oAscMathInterfaceType.Script)
    {
        if(Props.ScriptType == Asc.c_oAscMathInterfaceScript.Sup)
        {
            this.private_AddModifiedDegree(Pos, DEGREE_SUPERSCRIPT);
            bDelete = true;
        }
        else if(Props.ScriptType == Asc.c_oAscMathInterfaceScript.Sub)
        {
            this.private_AddModifiedDegree(Pos, DEGREE_SUBSCRIPT);
            bDelete = true;
        }
    }

    var RemoveChar      = Props.Action & c_oMathMenuAction.RemoveAccentCharacter && Item.kind == MATH_ACCENT,
        RemoveBar       = Props.Action & c_oMathMenuAction.RemoveBar && Item.kind == MATH_BAR,
        RemoveScript    = Props.Type == Asc.c_oAscMathInterfaceType.Script && Props.ScriptType == Asc.c_oAscMathInterfaceScript.None && (Item.kind == MATH_DEGREESubSup || Item.kind == MATH_DEGREE),
        RemoveLimit     = Props.Type == Asc.c_oAscMathInterfaceType.Limit && Props.Pos == Asc.c_oAscMathInterfaceLimitPos.None && Item.kind === MATH_LIMIT,
        RemoveMatrix    = Props.Type == Asc.c_oAscMathInterfaceType.Matrix && this.Content[Pos].Is_DeletedItem(Props.Action),
        RemoveEqArray   = Props.Type == Asc.c_oAscMathInterfaceType.EqArray && this.Content[Pos].Is_DeletedItem(Props.Action),
        RemoveDelimiter = Props.Action & c_oMathMenuAction.RemoveDelimiter && Item.kind == MATH_DELIMITER,
        RemoveGroupChar = Props.Type == Asc.c_oAscMathInterfaceType.GroupChar && Props.Pos == Asc.c_oAscMathInterfaceGroupCharPos.None && Item.kind == MATH_GROUP_CHARACTER,
        RemoveRadical   = Props.Action & c_oMathMenuAction.RemoveRadical && Item.kind == MATH_RADICAL;


    if(RemoveChar || RemoveBar || RemoveScript || RemoveLimit || RemoveMatrix || RemoveEqArray || RemoveDelimiter || RemoveGroupChar || RemoveRadical)
    {
        var Items = this.Content[Pos].Get_DeletedItemsThroughInterface();

        if(Items == null) // такого не должно произойти
            return;

        this.Remove_FromContent(Pos, 1);

        this.ConcatToContent(Pos, Items);

        this.Correct_Content();

        bDelete = true;
    }

    return bDelete;
};
CMathContent.prototype.private_AddModifiedDegree = function(Pos, Type)
{
    var DegreeSubSup = this.Content[Pos];
    var Base     = DegreeSubSup.getBase();
    var Iterator = Type == DEGREE_SUBSCRIPT ? DegreeSubSup.getLowerIterator() : DegreeSubSup.getUpperIterator();

    var Degree = new CDegree({type: Type}, false);

    this.Remove_FromContent(Pos, 1);
    this.Add_ToContent(Pos, Degree);

    var oBase     = Degree.getBase(),
        oIterator = Degree.getIterator();

    oBase.Concat_ToEnd(Base.Content);
    oIterator.Concat_ToEnd(Iterator.Content);
};
CMathContent.prototype.Get_MenuProps = function()
{
    var Pr = new CMathMenuBase();

    var Pos = this.private_FindCurrentPosInContent();

    if(Pos !== null && this.Content[Pos].Type == para_Math_Composition)
    {
        Pr = this.Content[Pos].Get_MenuProps();
    }
    else
    {
        this.Can_ModifyForcedBreak(Pr);
    }

    return Pr;
};
CMathContent.prototype.private_IsMenuPropsForContent = function(Action)
{
    // данные изменения могут прийти для любого типа изменений
    var bInsertForcedBreak = Action & c_oMathMenuAction.InsertForcedBreak,
        bDeleteForcedBreak = Action & c_oMathMenuAction.DeleteForcedBreak,
        bIncreaseArgSize   = Action & c_oMathMenuAction.IncreaseArgumentSize,
        bDecreaseArgSize   = Action & c_oMathMenuAction.DecreaseArgumentSize;

    return bDecreaseArgSize || bIncreaseArgSize || bInsertForcedBreak || bDeleteForcedBreak;
};
CMathContent.prototype.Process_AutoCorrect = function(ActionElement) {
    var bNeedAutoCorrect = this.private_NeedAutoCorrect(ActionElement);
    var AutoCorrectEngine = new CMathAutoCorrectEngine(ActionElement, this.CurPos, this.GetParagraph());
    //добавить все элементы
    AutoCorrectEngine.private_Add_Element(this.Content);
    if (null == AutoCorrectEngine.TextPr) {
        AutoCorrectEngine.TextPr = new CTextPr();
    }
    if (null == AutoCorrectEngine.MathPr) {
        AutoCorrectEngine.MathPr = new CMPrp();
    }
    var oParagraph = this.GetParagraph();
    var oLogicDocument = oParagraph ? oParagraph.LogicDocument : null;
    var CanMakeAutoCorrectEquation = false;
    var CanMakeAutoCorrectFunc     = false;
    var CanMakeAutoCorrect         = false;
    if (false === bNeedAutoCorrect && ActionElement.Type === para_Math_Text) {
        return false;
    } else {
        this.private_UpdateAutoCorrectMathSymbols();
        // Смотрим возможно ли выполнить автозамену, если нет, тогда пробуем произвести автозамену пропуская последний символ
        if (AutoCorrectEngine.IntFlag) {
            if (g_aMathAutoCorrectTriggerCharCodes[ActionElement.value]) {
                CanMakeAutoCorrect = this.private_CanAutoCorrectText(AutoCorrectEngine, true);
            } else {
                CanMakeAutoCorrect = this.private_CanAutoCorrectText(AutoCorrectEngine, false);
            }
        }
        if (!CanMakeAutoCorrect && g_aMathAutoCorrectTextFunc[ActionElement.value]) { 
            CanMakeAutoCorrectFunc = this.private_CanAutoCorrectTextFunc(AutoCorrectEngine);
        }
        AutoCorrectEngine.CurPos = AutoCorrectEngine.Elements.length - AutoCorrectEngine.Remove.total - 1;
        // Пробуем сделать формульную автозамену
        if (!CanMakeAutoCorrectFunc && !CanMakeAutoCorrect) {
            CanMakeAutoCorrectEquation = AutoCorrectEngine.private_CanAutoCorrectEquation(CanMakeAutoCorrect);
        }
    }
    if (CanMakeAutoCorrect || CanMakeAutoCorrectEquation || CanMakeAutoCorrectFunc) {
        AscFonts.FontPickerByCharacter.checkText(AutoCorrectEngine.RepCharsCode, this, function() {
            if (AscCommon.g_fontManager) {
                AscCommon.g_fontManager.ClearFontsRasterCache();
                AscCommon.g_fontManager.m_pFont = null;
            }
            if (AscCommon.g_fontManager2) {
                AscCommon.g_fontManager2.ClearFontsRasterCache();
                AscCommon.g_fontManager2.m_pFont = null;
            }
            if(oLogicDocument) {
                this.private_ReplaceAutoCorrect(AutoCorrectEngine);
                if (oLogicDocument.Api.WordControl.EditorType == "presentations") {
                    if (this.Paragraph.Parent.Parent.parent)
                        this.Paragraph.Parent.Parent.parent.checkExtentsByDocContent();
                }
            } else {
                var shape = this.Paragraph.Parent.DrawingDocument.drawingObjects.controller.selectedObjects[0];
                var wb = shape.worksheet.workbook.oApi.wb;
                for (var i = 0, length = wb.fmgrGraphics.length; i < length; ++i) 
                    wb.fmgrGraphics[i].ClearFontsRasterCache();
                
                this.private_ReplaceAutoCorrect(AutoCorrectEngine);
                shape.checkExtentsByDocContent();
                this.Paragraph.Parent.DrawingDocument.drawingObjects.controller.startRecalculate();
            }
            if (AutoCorrectEngine.StartHystory) {
                if (oLogicDocument) {
                     oLogicDocument.FinalizeAction();
                }
                // History.Remove_LastPoint();
                AutoCorrectEngine.StartHystory = false;
            }
        }, true, false, true); 
    }
};
CMathContent.prototype.private_NeedAutoCorrect = function(ActionElement) {
    var CharCode;
    if (para_Math_Ampersand == ActionElement.Type) {
        CharCode = 0x26;
    } else {
        CharCode = ActionElement.value;
    }
    if (1 === g_aMathAutoCorrectTriggerCharCodes[CharCode]) {
        return true;
    }
    return false;
};
CMathContent.prototype.private_UpdateAutoCorrectMathSymbols = function() {
    g_AutoCorrectMathSymbols = window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathSymbols;
    g_AutoCorrectMathFuncs = window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathFuncs;

};
CMathContent.prototype.private_CanAutoCorrectText = function(AutoCorrectEngine, bSkipLast) {
    var IndexAdd = (g_aMathAutoCorrectTriggerCharCodes[AutoCorrectEngine.ActionElement.value]) ? 1 : 0;
    var skip = IndexAdd - (AutoCorrectEngine.ActionElement.value == 0x20) ? 1 : 0;
    var ElCount = AutoCorrectEngine.Elements.length;
    if (ElCount < 1 + IndexAdd) {
        return false;
    }
    var Result = false;
    var FlagEnd = false;
    var RemoveCount = 0;
    var Start = 0;
    var ReplaceChars = [0x0020];
    var AutoCorrectCount = g_AutoCorrectMathSymbols.length;
    for (var nIndex = 0; nIndex < AutoCorrectCount; nIndex++) {
        var AutoCorrectElement = g_AutoCorrectMathSymbols[nIndex];
        var Ind = g_aMathAutoCorrectSpecSymb.findIndex(function(el, ind) {
            if (el == AutoCorrectElement[0])
                return el;
        });
        var IndexSkip = (Ind === -1) ? 0 : 1;
        var CheckString = AutoCorrectElement[0];
        var CheckStringLen = CheckString.length;

        if (ElCount < CheckStringLen + IndexAdd - IndexSkip) {
            continue;
        }
        var Found = true;
        for (var nStringPos = 0; nStringPos < CheckStringLen; nStringPos++) {
            var LastEl = AutoCorrectEngine.Elements[ElCount - nStringPos - 1 - IndexAdd + IndexSkip];
            if (!LastEl.Element.IsText()) {
                FlagEnd = true;
                Found = false;
                break;
            }
            if (String.fromCharCode(LastEl.Element.value) !== CheckString[CheckStringLen - nStringPos - 1]) {
                Found = false;
                break;
            }
        }
        if (true === Found) {
            RemoveCount = CheckStringLen + IndexAdd - skip;
            Start = ElCount - RemoveCount - skip + IndexSkip;

            if (undefined === AutoCorrectElement[1].length) {
                ReplaceChars[0] = AutoCorrectElement[1];
            } else {
                for (var i = 0, Len = AutoCorrectElement[1].length; i < Len; i++) {
                    ReplaceChars[i] = AutoCorrectElement[1][i];
                }
            }
            FlagEnd = true;
        }
        if (FlagEnd) {
            break;
        }
    }
    if (RemoveCount > 0) {
        AutoCorrectEngine.StartHystory = true;
        var oParagraph = this.GetParagraph();
        var oLogicDocument = oParagraph ? oParagraph.LogicDocument : null;
        if(oLogicDocument) {
            oLogicDocument.StartAction(AscDFH.historydescription_Document_MathAutoCorrect);
        } else {
            History.Create_NewPoint(AscDFH.historydescription_Document_MathAutoCorrect);
        }
        var MathRun = new ParaRun(this.Paragraph, true);
        MathRun.Set_Pr(AutoCorrectEngine.TextPr.Copy());
        MathRun.Set_MathPr(AutoCorrectEngine.MathPr);

        for (var i = 0, Count = ReplaceChars.length; i < Count; i++) {
            var ReplaceText = new CMathText();
            ReplaceText.add(ReplaceChars[i]);
            MathRun.Add(ReplaceText, true);
            AutoCorrectEngine.RepCharsCode.push(ReplaceChars[i]);
        }
        AutoCorrectEngine.Remove.push({Count:RemoveCount, Start:Start});
        AutoCorrectEngine.Remove.total += RemoveCount;
        AutoCorrectEngine.ReplaceContent.push(MathRun);
        Result = true;
    }
    return Result;
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectText = function(Elements) {
    var ElCount = Elements.length || 0;
    if (ElCount < 2) {
        return;
    }
    var AutoCorrectCount = g_AutoCorrectMathSymbols.length;
    for (var nIndex = 0; nIndex < AutoCorrectCount; nIndex++) {
        var AutoCorrectElement = g_AutoCorrectMathSymbols[nIndex];
        var CheckString = AutoCorrectElement[0];
        var CheckStringLen = CheckString.length;
        if (ElCount < CheckStringLen) {
            continue;
        }
        var Found = true;
        for (var nStringPos = 0; nStringPos < CheckStringLen; nStringPos++) {
            var LastEl = Elements[ElCount - nStringPos - 1];
            if (!LastEl.IsText()) {
                Found = false;
                break;
            }
            if (String.fromCharCode(LastEl.value) !== CheckString[CheckStringLen - nStringPos - 1]) {
                Found = false;
                break;
            }
        }
        if (true === Found) {
            var Start = ElCount - nStringPos;
            Elements.splice(Start + 1, CheckStringLen);
            var ReplaceText = new CMathText();
            ReplaceText.add(AutoCorrectElement[1]);
            Elements[Start] = ReplaceText;
            return;
        }
    }    
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectTextFunc = function(Elements) {
    var ElCount = Elements.length || 0;
    if (ElCount < 2 ) {
        return;
    }
    var AutoCorrectCount = g_AutoCorrectMathFuncs.length;
    for (var nIndex = 0; nIndex < AutoCorrectCount; nIndex++) {
        var AutoCorrectElement = g_AutoCorrectMathFuncs[nIndex];
        var CheckStringLen = AutoCorrectElement.length;
        if (ElCount < CheckStringLen) {
            continue;
        }
        var Found = true;
        for (var nStringPos = 0; nStringPos < CheckStringLen; nStringPos++) {
            var LastEl = Elements[ElCount - nStringPos - 1];
            if (!LastEl.IsText()) {
                Found = false;
                break;
            }
            if (String.fromCharCode(LastEl.value) !== AutoCorrectElement[CheckStringLen - nStringPos - 1]) {
                Found = false;
                break;
            }
            if (nStringPos === (CheckStringLen - 1) && (ElCount - CheckStringLen) > 0) {
                LastEl = Elements[ElCount - nStringPos - 2];
                if ((LastEl.IsText() && LastEl.value !== 32)) {
                    Found = false;
                }
            }
        }
        if (Found === true) {
            var Start = ElCount - nStringPos;
            Elements.splice(Start + 1, CheckStringLen);
            var Pr = {ctrPrp: new CTextPr()};
            var MathFunc = new CMathFunc(Pr);
            var MathContent = MathFunc.getFName();
            var MathRun = new ParaRun(this.Paragraph, true);
            for (var nCharPos = 0, nTextLen = AutoCorrectElement.length; nCharPos < nTextLen; nCharPos++) {
                var oText = null;
                if (0x0026 == AutoCorrectElement.charCodeAt(nCharPos)) {
                    oText = new CMathAmp();
                } else {
                    oText = new CMathText(false);
                    oText.addTxt(AutoCorrectElement[nCharPos]);
                }
                MathRun.Add(oText, true);
            }
            MathRun.Math_Apply_Style(STY_PLAIN);
            MathContent.Internal_Content_Add(0, MathRun);
            Elements[Start] = MathFunc;
            return;
        }
    }
};
CMathAutoCorrectEngine.prototype.private_FindBracketsSkip = function(Elements) {
    var Pos = Elements.length - 1;
    var BracketsR = [];
    var BracketsL = [];
    while (Pos >= 0) {
        var Elem = Elements[Pos].value;
        if (g_MathRightBracketAutoCorrectCharCodes[Elem]) {
            BracketsR.push(Pos);
            Pos--;
            continue;
        } else if (g_MathLeftBracketAutoCorrectCharCodes[Elem]) {
            if (BracketsR.length < BracketsL.length) {
                return [];
            }
            if (Pos - 1 >= 0 && g_aMathAutoCorrectSkipBrackets[Elements[Pos-1].value]) {
                BracketsL.push(Pos);
            } else {
                BracketsR.pop();
            }
            Pos--;
            continue;
        }
        Pos--;
    }
    return [BracketsL,BracketsR];
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectEquation = function(Elements) {
    //пока закоментил вызов этих функции, так как в ворде они не вызываются
    // this.private_AutoCorrectText(Elements);
    // this.private_AutoCorrectTextFunc(Elements);
    var Brackets = [];
    var CurPos = Elements.length - 1;
    var Param = {
        Type : null,
        Kind : null,
        Props : null,
        Bracket : this.private_FindBracketsSkip(Elements),
        bOff : false
    };
    var ElPos = [];
    var End = CurPos;
    while (CurPos >= 0) {
        var Elem = Elements[CurPos];
        if (Elem.value === undefined) {
            CurPos--;
            continue;
        } else if (Elem.value === 0x002F) { // /
            if (Param.Type !== null) {
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos+1,ElPos[0]+1)];
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                }
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (CurPos - 1 > 0) {
                var tmp = Elements[CurPos-1];
                if (tmp.Type === para_Math_BreakOperator && tmp.value === 0x005C) { // \
                    ElPos[0] = CurPos;
                    if (!Brackets[0]) {
                        Param.Type = MATH_FRACTION;
                        Param.Props = {type: LINEAR_FRACTION};
                        CurPos -= 2;
                        continue;
                    }
                }
            }
            if (!Brackets[0]) {
                Param.Type = MATH_FRACTION;
                Param.Props = {};
            }
            ElPos[0] = CurPos;
            CurPos--;
            continue;
        } else if (Elem.value === 0x2044) { // fraction
            if (Param.Type !== null) {
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos+1,ElPos[0]+1)];
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                }
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                Param.Type = MATH_FRACTION;
                Param.Props = {type: SKEWED_FRACTION};
            }
            ElPos[0] = CurPos;
            CurPos--;
            continue;
        } else if (0x00A6 === Elem.value) { // fraction
            if (Param.Type !== null) {
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos+1,ElPos[0]+1)];
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                }
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                Param.Type = MATH_FRACTION;
                Param.Props = {type: NO_BAR_FRACTION};
            }
            ElPos[0] = CurPos;
            CurPos--;
            continue;
        } else if (g_MathRightBracketAutoCorrectCharCodes[Elem.value]) { // right bracket
            //если справа от скобки не "^","_" и тип степень, то сделать автозамену степени до этой скобки
            if (Param.Type == MATH_DEGREE || Param.Type == MATH_DEGREESubSup) {
                if (Elements[CurPos+1] && (Elements[CurPos+1].value != 0x005E && Elements[CurPos+1].value != 0x005F)) {
                    var tmp = null;
                    if (Param.Type == MATH_DEGREE) {
                        tmp =  [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos+1)];
                    } else {
                        tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos+1,ElPos[0]-CurPos+1)];
                    }
                    this.private_CorrectEquation(Param,tmp);
                    Param.Type = null;
                    Param.Props = {};
                    Param.Kind = null;
                    ElPos = [];
                    End = CurPos + 1;
                    Elements.splice(End, 0, tmp[0]);
                }
            }
            Brackets.splice(0,0,CurPos);
            CurPos--;
            continue;
        } else if (g_MathLeftBracketAutoCorrectCharCodes[Elem.value]) { // left bracket
            if (!Brackets[0]){
                break;
            }
            var fSkip = false;
            for (var i = Brackets[0] + 1; i < Elements.length;i++) {
                var tmpElem = Elements[i].value;
                if(g_MathRightBracketAutoCorrectCharCodes[tmpElem]) {
                    continue;
                } else if (q_aMathAutoCorrectAccentCharCodes[tmpElem]) {
                    fSkip = true;
                }
                break;
            }
            for (var i = 0; i < Param.Bracket[0].length; i++) {
                if (Param.Bracket[0][i] <= CurPos && CurPos <= Param.Bracket[1][i]) {
                    fSkip = true;
                    break;
                }
            }
            if (fSkip || Param.Type == MATH_FRACTION) {
                Brackets.splice(0,1);
                CurPos--;
                continue;
            }
            if (Param.Type !== null) {
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos,Brackets[0]-CurPos+2)];
                    fSkip = true;
                } else if (Param.Type === MATH_DEGREE) {
                    tmp =  [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos,Brackets[0]-CurPos+2)];
                    fSkip = true;
                } else {
                    tmp =  [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(Brackets[0]+1,ElPos[0]-Brackets[0])];
                }
                this.private_CorrectEquation(Param,tmp);
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
                ElPos = [];
                if (fSkip) {
                    End = CurPos;
                    Elements.splice(End, 0, tmp[0]);
                    Brackets.splice(0,1);
                    CurPos--;
                    continue;
                }
                End = Brackets[0] + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            var count = Brackets[0] - CurPos + 1;
            var tmp = Elements.splice(CurPos,count);
            this.private_CorrectEquation({Type:MATH_DELIMITER,Props:null,Kind:null},tmp);
            Elements.splice(CurPos,0,tmp[0]);
            End = CurPos;
            Brackets.splice(0,1);
            for (var i = 0; i < Brackets.length; i++) {
                Brackets[i] -= count - 1;
            }
            CurPos--;
            continue;
        } else if  (Elem.value === 0x005E) { // ^
            var bSkip = false;
            if (CurPos - 1 >= 0 && Param.Type != MATH_DEGREESubSup) {
                var tmp = Elements[CurPos-1];
                if (tmp.value === 0x2061) {
                    bSkip = true;
                    if (Param.Type == MATH_DEGREE && Param.Kind != DEGREE_SUPERSCRIPT) {
                        Param.Type = null;
                        ElPos.unshift(CurPos);
                    } else {
                        Param.Kind = DEGREE_SUPERSCRIPT;
                        ElPos[0] = CurPos;
                    }
                }
            }
            if (Param.Type == MATH_FRACTION || bSkip) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null && (Param.Type !== MATH_DEGREE)) {
                var tmp = null;
                if (Param.Type == MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos+1,ElPos[0]-CurPos)];                    
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,(ElPos[0]-CurPos))];
                }
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;+
                Elements.splice(End, 0, tmp[0]);
                Param.Type = null;
                Param.Kind = null;
                Param.Props = {};
                ElPos = [];
            }
            if (Param.Type == MATH_DEGREE && Param.Kind == DEGREE_SUBSCRIPT) {
                Param.Kind = DEGREE_SubSup;
                Param.Type = MATH_DEGREESubSup;
                ElPos[1] = ElPos[0];
            } else {
                if (!Brackets[0]) {
                    Param.Kind = DEGREE_SUPERSCRIPT;
                    Param.Type = MATH_DEGREE;
                }
            }
            ElPos[0] = CurPos;
            CurPos--;
            continue;
        } else if (Elem.value === 0x005F) { // _
            var bSkip = false;
            if (CurPos - 1 >= 0 && Param.Type != MATH_DEGREESubSup) {
                var tmp = Elements[CurPos-1];
                if (tmp.value === 0x2061) {
                    bSkip = true;
                    if (Param.Type == MATH_DEGREE && Param.Kind != DEGREE_SUBSCRIPT) {
                        Param.Type = null;
                        ElPos.unshift(CurPos);
                    } else {
                        Param.Kind = DEGREE_SUBSCRIPT;
                        ElPos[0] = CurPos;
                    }
                }
            }
            if (Param.Type == MATH_FRACTION || bSkip) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null && (Param.Type !== MATH_DEGREE)) {
                var tmp = null;
                if (Param.Type == MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos+1,ElPos[0]-CurPos)];                    
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,(ElPos[0]-CurPos))];
                }
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
                Param.Type = null;
                Param.Kind = null;
                Param.Props = {};
                ElPos = [];
            }
            if (Param.Type == MATH_DEGREE && Param.Kind == DEGREE_SUPERSCRIPT) {
                Param.Kind = DEGREE_SubSup;
                Param.Type = MATH_DEGREESubSup;
                ElPos[1] = ElPos[0];
            } else {
                if (!Brackets[0]) {
                    Param.Kind = DEGREE_SUBSCRIPT;
                    Param.Type = MATH_DEGREE;
                }
            }
            ElPos[0] = CurPos;
            CurPos--;
            continue;
        } else if (q_aMathAutoCorrectControlAggregationCodes[Elem.value]) { // cNary
            if (Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null && Param.Type !== MATH_DEGREESubSup && Param.Type !== MATH_DEGREE) {
                var tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                var tmp = null;
                if (Param.Type == MATH_DEGREE) {
                    tmp = [Elements.splice(ElPos[0],End-ElPos[0]+1),Elements.splice(CurPos,ElPos[0]-CurPos)];
                } else if (Param.Type == MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1],End-ElPos[1]+1),Elements.splice(ElPos[0],ElPos[1]-ElPos[0]),Elements.splice(CurPos,ElPos[0]-CurPos)];
                } else {
                    tmp = [Elements.splice(CurPos+1,(End-CurPos)),Elements.splice(CurPos,ElPos[0]-CurPos)];
                }
                Param.Type = MATH_NARY;
                this.private_CorrectEquation(Param,tmp);
                End = CurPos;
                Elements.splice(End,0,tmp[0]);  
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
                Param.bOff = false;
            }
            CurPos--;
            continue;
        } else if (g_aMathAutoCorrectRadicalCharCode[Elem.value]) { // sqrt
            if (Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null && Param.Type !== MATH_DEGREE && Param.Type !== MATH_DEGREESubSup) {
                var tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                var tmp = [Elements.splice(CurPos+1,(End-CurPos)),Elements.splice(CurPos,1)];
                Param.Type = MATH_RADICAL;
                Param.Bracket[0][0] -= (CurPos + 1);
                Param.Bracket[1][0] -= (CurPos + 1);
                this.private_CorrectEquation(Param,tmp);
                End = CurPos;
                Elements.splice(End,0,tmp[0]);  
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
            }
            CurPos--;
            continue;
        } else if (Elem.value == 0x25A1 || Elem.value == 0x25AD) { // box OR border box
            if (Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null && Param.Type !== MATH_DEGREE && Param.Type !== MATH_DEGREESubSup) {
                var tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                var tmp = [Elements.splice(CurPos+1,(End-CurPos)),Elements.splice(CurPos,1)];
                Param.Type = MATH_BOX;
                Param.Bracket[0][0] -= (CurPos + 1);
                Param.Bracket[1][0] -= (CurPos + 1);
                this.private_CorrectEquation(Param,tmp);
                End = CurPos;
                Elements.splice(End,0,tmp[0]);  
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
            }
            CurPos--;
            continue;
        } else if (g_aMathAutoCorrectGroupChar[Elem.value]) { // group character
            if (Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null /*&& Param.Type !== MATH_DEGREE && Param.Type !== MATH_DEGREESubSup*/) {
                var fSkip = false;
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos,ElPos[0]-CurPos+1)];
                    fSkip = true;
                } 
                else if (Param.Type === MATH_DEGREE) {
                    tmp = [Elements.splice(ElPos[0],(End-ElPos[0]+1)),Elements.splice(CurPos+1,ElPos[0]-CurPos-1),Elements.splice(CurPos,1)];
                    Param.Type = MATH_GROUP_CHARACTER;
                    Param.Bracket[0][0] -= (CurPos + 1);
                    Param.Bracket[1][0] -= (CurPos + 1);
                    fSkip = true;
                }
                else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                }
                this.private_CorrectEquation(Param,tmp);
                if (fSkip) {
                    End = CurPos;
                    Elements.splice(End, 0, tmp[0]);
                    CurPos--;
                    Param.Type = null;
                    Param.Props = {};
                    Param.Kind = null;
                    continue;
                }
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                var tmp = [Elements.splice(CurPos+1,(End-CurPos)),Elements.splice(CurPos,1)];
                Param.Type = MATH_GROUP_CHARACTER;
                Param.Bracket[0][0] -= (CurPos + 1);
                Param.Bracket[1][0] -= (CurPos + 1);
                this.private_CorrectEquation(Param,tmp);
                End = CurPos;
                Elements.splice(End,0,tmp[0]);  
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
            }
            CurPos--;
            continue;
        } else if (Elem.value === 0x00AF || Elem.value === 0x2581) { // bar
            if (Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null /*&& Param.Type !== MATH_DEGREE && Param.Type !== MATH_DEGREESubSup*/) {
                var fSkip = false;
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos,ElPos[0]-CurPos+1)];
                    fSkip = true;
                } else if (Param.Type === MATH_DEGREE) {
                    tmp =  [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos,ElPos[0]-CurPos+1)];
                    fSkip = true;
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                }
                this.private_CorrectEquation(Param,tmp);
                if (fSkip) {
                    End = CurPos;
                    Elements.splice(End, 0, tmp[0]);
                    CurPos--;
                    Param.Type = null;
                    Param.Props = {};
                    Param.Kind = null;
                    continue;
                }
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                var tmp = [Elements.splice(CurPos+1,(End-CurPos)),Elements.splice(CurPos,1)];
                Param.Type = MATH_BAR;
                Param.Bracket[0][0] -= (CurPos + 1);
                Param.Bracket[1][0] -= (CurPos + 1);
                this.private_CorrectEquation(Param,tmp);
                End = CurPos;
                Elements.splice(End,0,tmp[0]);  
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
            }
            CurPos--;
            continue;
        } else if (g_aMathAutoCorrectEqArrayMatrix[Elem.value]) { // eq array and matrix
            if (Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null) {
                var fSkip = false;
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos,ElPos[0]-CurPos+1)];
                    fSkip = true;
                } else if (Param.Type === MATH_DEGREE) {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos,ElPos[0]-CurPos+1)];
                    fSkip = true;
                } else if (Param.Type == MATH_LIMIT) {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos,ElPos[0]-CurPos+1)];
                    fSkip = true;
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                }
                this.private_CorrectEquation(Param,tmp);
                if (fSkip) {
                    End = CurPos;
                    Elements.splice(End, 0, tmp[0]);
                    CurPos--;
                    Param.Type = null;
                    Param.Props = {};
                    Param.Kind = null;
                    continue;
                }
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                tmp = [Elements.splice(CurPos+1,(Param.Bracket[1][0]-CurPos)),Elements.splice(CurPos,1)];
                Param.Type = MATH_MATRIX;
                Param.Bracket[0][0] -= (CurPos + 1);
                Param.Bracket[1][0] -= (CurPos + 1);
                this.private_CorrectEquation(Param,tmp);
                End = CurPos;
                Elements.splice(End,0,tmp[0]);  
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
            }
            CurPos--;
            continue;
        } else if (q_aMathAutoCorrectAccentCharCodes[Elem.value]) { // accent
            if (Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null && Param.Type !== MATH_DEGREE && Param.Type !== MATH_DEGREESubSup) {
                var tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                if (CurPos >= 1) {
                    var tmpPos = CurPos;
                    var tmpEl = Elements[CurPos-1];
                    if (g_MathRightBracketAutoCorrectCharCodes[tmpEl.value]) {
                        var countR = 1;
                        tmpPos--;
                        for (var i = CurPos - 2; i >= 0; i--) {
                            tmpEl = Elements[i];
                            tmpPos--;
                            if (g_MathRightBracketAutoCorrectCharCodes[tmpEl.value]) {
                                countR++;
                            } if (g_MathLeftBracketAutoCorrectCharCodes[tmpEl.value]) {
                                countR--;
                                if (!countR) {
                                    break;
                                }
                            }
                        }
                    } else {
                        tmpPos--;
                    }
                    if (tmpEl !== Elements[CurPos-1] && countR) {
                        CurPos--;
                        continue;
                    }
                }
                var fSkip = false;
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(tmpPos,ElPos[0]-tmpPos+1)];
                    fSkip = true;
                } else if (Param.Type === MATH_DEGREE) {
                    tmp =  [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(tmpPos,ElPos[0]-tmpPos+1)];
                    fSkip = true;
                } else {
                    tmp = [Elements.splice(CurPos,1),Elements.splice(tmpPos,CurPos-tmpPos)];
                }
                if (fSkip) {
                    this.private_CorrectEquation(Param,tmp);
                    CurPos = End = tmpPos;
                    Elements.splice(End, 0, tmp[0]);
                    CurPos --;
                    Param.Type = null;
                    Param.Props = {};
                    Param.Kind = null;
                    continue;
                }
                Param.Type = MATH_ACCENT;
                this.private_CorrectEquation(Param,tmp);
                End = tmpPos;
                Elements.splice(End,0,tmp[0]);
                CurPos = End; 
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
            }
            CurPos--;
            continue;
        } else if (Elem.value === 0x2534 || Elem.value === 0x252C) { // above and below
            if (Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null && Param.Type !== MATH_DEGREESubSup && Param.Type !== MATH_DEGREE) {
                var tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
            }
            if (!Brackets[0]) {
                Param.Type = MATH_LIMIT;
                ElPos[0] = CurPos;
                Param.Props = (Elem.value === 0x2534) ? {type : LIMIT_UP} : {type : LIMIT_LOW};
            }
            CurPos--;
            continue;
        } else if (Elem.value === 0x2061) { // 0x2061
            if (Param.Type !== MATH_FRACTION && !Brackets[0]) {
                if (Param.Type == MATH_DELIMITER) {
                    Param.Type = null;
                }
                if (Param.Type !== null  && Param.Type !== MATH_DEGREE) {
                    var tmp = null;
                    if (Param.Type === MATH_DEGREESubSup) {
                        tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos,ElPos[0]-CurPos+1)];
                    } else {
                        tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                    }
                    this.private_CorrectEquation(Param,tmp);
                    End = CurPos + 1;
                    Elements.splice(End, 0, tmp[0]);
                }
                var tmpPos = CurPos;
                CurPos--;
                while (CurPos >= 0) {
                    Elem = Elements[CurPos];
                    if (g_aMathAutoCorrectLatinAlph[Elem.value]) {
                        CurPos--;
                    } else {
                        break;
                    }
                }
                var tmp = {
                    arrName : null,
                    arrSup : null,
                    arrSub : null,
                    arrArg : null
                };
                if (ElPos.length > 1) {
                    if (Param.Kind == 1) {
                        tmp.arrSup = Elements.splice(ElPos[1], End - ElPos[1] + 1);
                        tmp.arrSub = Elements.splice(ElPos[0], ElPos[1] - ElPos[0]);
                    } else {
                        tmp.arrSub = Elements.splice(ElPos[1], End - ElPos[1] + 1);
                        tmp.arrSup = Elements.splice(ElPos[0], ElPos[1] - ElPos[0]);
                    }
                } else if (ElPos.length) {
                    if (Param.Kind == 1) {
                        tmp.arrSup = Elements.splice(ElPos[0], End - ElPos[0] + 1);
                    } else {
                        tmp.arrSub = Elements.splice(ElPos[0], End - ElPos[0] + 1);
                    }
                } else {
                    tmp.arrArg = Elements.splice(tmpPos + 1, End - tmpPos);
                }
                tmp.arrName = Elements.splice(CurPos + 1, tmpPos - CurPos);
                Param.Type = MATH_FUNCTION;
                this.private_CorrectEquation(Param,tmp);
                End = (CurPos < 0) ? 0 : CurPos + 1;
                Elements.splice(End,0,tmp.arrName);  
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
            } else {
                CurPos--;
            }
            continue;
        } else if (Elem.value === 0x2592) { // naryand (off)
            if (Param.Type !== null) {
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos+1,ElPos[0]+1)];
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                }
                this.private_CorrectEquation(Param,tmp);
                End = CurPos + 1;
                Elements.splice(End, 0, tmp[0]);
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
            }
            Param.bOff = true;
            CurPos--;
            continue;
        } else if (g_aMathAutoCorrectFracCharCodes[Elem.value] && !Brackets[0]) { // space + - # @ and another
            if(Elem.value == 0x5C && Param.Type == MATH_FRACTION) {
                CurPos--;
                continue;
            }
            if (Param.Type !== null) {
                var tmp = null;
                if (Param.Type === MATH_DEGREESubSup) {
                    tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                } else {
                    tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(CurPos+1,ElPos[0]-CurPos)];
                }
                this.private_CorrectEquation(Param,tmp);
                Elements.splice(CurPos + 1, 0, tmp[0]);
                
                Param.Type = null;
                Param.Props = {};
                Param.Kind = null;
                ElPos = [];
            }
            End = CurPos - 1;
        }
        CurPos--;
    }
    if (Brackets[0]) {
        return;
    }
    if (Param.Type !== null) {
        var tmp = null;
        if (Param.Type === MATH_DEGREESubSup) {
            tmp = [Elements.splice(ElPos[1]+1,End-ElPos[1]),Elements.splice(ElPos[0]+1,ElPos[1]-ElPos[0]),Elements.splice(0,ElPos[0]+1)];
        } else {
            tmp = [Elements.splice(ElPos[0]+1,(End-ElPos[0])),Elements.splice(0,ElPos[0]-CurPos)];
        }
        this.private_CorrectEquation(Param,tmp);
        Elements.splice(0,0,tmp[0]);
    }
};
CMathAutoCorrectEngine.prototype.private_CorrectEquation = function(Param, Elements) {
    switch(Param.Type) {
        case MATH_FRACTION:
            this.private_CorrectBuffForFrac(Elements, Param.Props);
            var props = new CMathFractionPr();
            props.Set_FromObject(Param.Props);
            props.ctrPrp = this.TextPr.Copy();
            var Fraction = new CFraction(props);
            var DenMathContent = Fraction.getDenominatorMathContent();
            var NumMathContent = Fraction.getNumeratorMathContent();
            this.private_PackTextToContent(DenMathContent, Elements[0], true);
            this.private_PackTextToContent(NumMathContent, Elements[1], true);
            Elements.splice(0,Elements.length, Fraction);
            break;
        case MATH_DEGREE:
            Elements[1].splice(Elements[1].length-1,1);
            var props = new CMathDegreePr();
            props.ctrPrp = this.TextPr.Copy();
            props.type = Param.Kind;
            var oDegree = new CDegree(props)
            var BaseContent = oDegree.Content[0];
            var IterContent = oDegree.Content[1];
            this.private_PackTextToContent(BaseContent, Elements[1], false);
            this.private_PackTextToContent(IterContent, Elements[0], true);
            Elements.splice(0,Elements.length, oDegree);
            break;
        case MATH_DEGREESubSup:
            var flag = (Elements[1][Elements[1].length-1].value == 0x005E) ? true : false;
            Elements[1].splice(Elements[1].length-1,1);
            Elements[2].splice(Elements[2].length-1,1);
            var props = new CMathDegreePr();
            props.ctrPrp = this.TextPr.Copy();
            props.type = Param.Kind;
            var oDegree = new CDegreeSubSup(props);
            var BaseContent = oDegree.Content[0];
            var IterDnContent = oDegree.Content[1];
            var IterUpContent = oDegree.Content[2];
            if (flag) {
                this.private_PackTextToContent(IterUpContent, Elements[1], true);
                this.private_PackTextToContent(IterDnContent, Elements[0], true);
            } else {
                this.private_PackTextToContent(IterUpContent, Elements[0], true);
                this.private_PackTextToContent(IterDnContent, Elements[1], true);
            }
            this.private_PackTextToContent(BaseContent, Elements[2], false);
            Elements.splice(0,Elements.length, oDegree);
            break;
        case MATH_DELIMITER:
            var props = new CMathDelimiterPr();
            props.column = 1;
            props.begChr = Elements.splice(0,1)[0].value;
            props.endChr = Elements.splice(Elements.length-1,1)[0].value;
            var oDelimiter = new CDelimiter(props);
            var oBase = oDelimiter.getBase();
            this.private_PackTextToContent(oBase, Elements, false);
            Elements.splice(0,Elements.length, oDelimiter);
            break;
        case MATH_NARY:
            var props = {};
            props.supHide = true;
            props.subHide = true;
            props.chr = (Elements.length == 1) ? Elements[0][0].value : Elements[Elements.length-1][0].value;
            this.private_CorrectBuffForNary(Elements);

            var arrBase = [];
            if (Param.bOff) {
                for (var i = Elements[0].length - 1; i >= 0 ; i--) {
                    if (Elements[0][i].value == 0x2592) {
                        Elements[0].pop();
                        break;
                    }
                    arrBase.unshift(Elements[0].pop());
                }
            }
            if (Elements[0].Type) {
                if (Elements[0].Type == DEGREE_SUPERSCRIPT && Elements[0].length) {
                props.supHide = false;
                }
                if (Elements[0].Type == DEGREE_SUBSCRIPT && Elements[0].length) {
                props.subHide = false;
                }
            }
            if (Elements[1]) {
                if (Elements[1].Type == DEGREE_SUPERSCRIPT && Elements[1].length) {
                    props.supHide = false;
                }
                if (Elements[1].Type == DEGREE_SUBSCRIPT && Elements[1].length) {
                    props.subHide = false;
                }
            }
            props.ctrPrp = this.TextPr.Copy();
            var oNary = new CNary(props);
            var oSub = oNary.getLowerIterator();
            var oSup = oNary.getUpperIterator();
            var oBase = oNary.getBase();
            if (!props.supHide) {
                if (Elements[0].Type == DEGREE_SUPERSCRIPT) {
                    this.private_PackTextToContent(oSup, Elements[0], true);
                } else if (Elements[1].Type && Elements[1].Type == DEGREE_SUPERSCRIPT) {
                    this.private_PackTextToContent(oSup, Elements[1], true);
                }
            }
            if (!props.subHide) {
                if (Elements[0].Type == DEGREE_SUBSCRIPT) {
                    this.private_PackTextToContent(oSub, Elements[0], true);
                } else if (Elements[1].Type && Elements[1].Type == DEGREE_SUBSCRIPT) {
                    this.private_PackTextToContent(oSub, Elements[1], true);
                }
            }
            this.private_PackTextToContent(oBase, arrBase, true);
            Elements.splice(0,Elements.length, oNary);
            break;
        case MATH_RADICAL:
            Elements.splice(0,2,Elements[1].concat(Elements[0]));
            var props = new CMathRadicalPr();
            props.degHide = this.private_CorrectBuffForRadical(Elements, true, [Param.Bracket[0][0],Param.Bracket[1][0]]);
            props.ctrPrp = this.TextPr.Copy();
            var Radical = new CRadical(props);
            var Base = Radical.getBase();
            var Degree = Radical.getDegree();
            if (!props.degHide) {
                this.private_PackTextToContent(Degree, Elements[0][1], false);
                this.private_PackTextToContent(Base, Elements[0][0], true);
            } else {
                this.private_PackTextToContent(Base, Elements[0], true);
            }
            Elements.splice(0,Elements.length, Radical);
            Param.Bracket[0].splice(0,1);
            Param.Bracket[1].splice(0,1);
            break;
        case MATH_BOX:
            var symbol = Elements.splice(1,1)[0][0].value;
            var props = {};
            props.ctrPrp = this.TextPr.Copy();
            var Box = (symbol == 0x25A1) ? new CBox(props) : new CBorderBox(props);
            var Base = Box.getBase();
            this.private_PackTextToContent(Base, Elements[0], true);
            Elements.splice(0,Elements.length, Box);
            Param.Bracket[0].splice(0,1);
            Param.Bracket[1].splice(0,1); 
            break;
        case MATH_GROUP_CHARACTER:
            var degree = [];
            if (Param.Kind) {
                degree = Elements.splice(0,1)[0];
                degree.splice(0,1);
            }
            var base = Elements.splice(0,1)[0];
            var symbol = Elements.splice(0,1)[0][0].value;
            var props = {};
            switch (symbol) {
                case 0x23E0:
                case 0x23DC:
                case 0x23DE:
                    props = {chr:symbol, pos:LOCATION_TOP, vertJc:VJUST_BOT};
                    break;
                case 0x23DD:
                    props = {chr:symbol};
                    break;
            }
            props.ctrPrp = this.TextPr.Copy();
            var oGroupChr = new CGroupCharacter(props);
            var oBase = oGroupChr.getBase();
            this.private_PackTextToContent(oBase, base, true);
            if (degree.length) {
                props = {
                    ctrPrp : this.TextPr.Copy(),
                    type : (Param.Kind == 1) ? 1 : 0
                };
                var Limit = new CLimit(props);
                var MathContent = Limit.getFName();
                MathContent.Add_Element(oGroupChr);
                MathContent = Limit.getIterator();
                this.private_PackTextToContent(MathContent, degree, true);
                Elements.splice(0,Elements.length, Limit);
            } else {
                Elements.splice(0,Elements.length, oGroupChr);
            }
            Param.Bracket[0].splice(0,1);
            Param.Bracket[1].splice(0,1);
            break;
        case MATH_BAR:
            var symbol = Elements.splice(1,1)[0][0].value;
            var props = (symbol === 0x00AF) ? {pos:LOCATION_TOP} : {};
            props.ctrPrp = this.TextPr.Copy();
            var oBar = new CBar(props);
            var oBase = oBar.getBase();
            this.private_PackTextToContent(oBase, Elements[0], true);
            Elements.splice(0,Elements.length, oBar);
            Param.Bracket[0].splice(0,1);
            Param.Bracket[1].splice(0,1);
            break;
        case MATH_MATRIX:
            var bEqArray = false;
            var symbol = Elements.splice(1,1)[0][0];
            if (symbol.value == 0x24B8 || symbol.value == 0x2588) {
                bEqArray = true;
            }
            var Del = null;
            switch (symbol.value) {
                case 0x24B8:
                    var props = new CMathDelimiterPr();
                    props.begChr = 123;
                    props.endChr = -1;
                    props.column = 1;
                    Del = new CDelimiter(props);
                    break;
                case 0x24A8:
                    var props = new CMathDelimiterPr();
                    props.column = 1;
                    Del = new CDelimiter(props);
                    break;
                case 0x24A9:
                    var props = new CMathDelimiterPr();
                    props.column = 1;
                    props.begChr = 0x2016;
                    props.endChr = 0x2016;
                    Del = new CDelimiter(props);
                    break;
            }
            var Element = null;
            if (bEqArray) {
                var arrContent = [];
                var row = 0;
                arrContent[row] = [];
                if (Param.Bracket[0].length) {
                    for (var i = Param.Bracket[0][0]+1; i < Param.Bracket[1][0]; i++) {
                        var oCurElem = Elements[0][i];
                        if (oCurElem.value == 0x40) {  // @
                            row++;
                            arrContent[row] = [];
                        } else if (oCurElem.IsText()) {
                            arrContent[row].push(oCurElem);
                        } else {
                            arrContent[row].push(oCurElem);
                        }
                    }
                }
                if (row < 1 && Del) {
                    Elements.splice(0,0, symbol);
                    Param.Bracket[0].splice(0,1);
                    Param.Bracket[1].splice(0,1);
                    break;
                }
                var props = new CMathEqArrPr();
                props.row = row + 1;
                props.ctrPrp = this.TextPr.Copy();
                Element = new CEqArray(props);
                for (var i = 0; i < arrContent.length; i++) {
                    var El = Element.getElement(i);
                    //getElementMathContent
                    this.private_PackTextToContent(El, arrContent[i], false);
                }
            } else {
                var arrContent = [];
                var col = 0;
                var row = 0;
                var mcs = [];
                arrContent[row] = [];
                arrContent[row][col] = [];
                mcs[0] = {count: 1, mcJc: 0};
                if (Param.Bracket[0].length) {
                    for (var i = Param.Bracket[0][0]+1; i < Param.Bracket[1][0]; i++) {
                        var oCurElem = Elements[0][i];
                        if (oCurElem.value == 0x26) {  // &
                            col++;
                            if (col+1 > mcs[0].count) {
                                mcs[0] = {count: col+1, mcJc: 0};
                            }
                            arrContent[row][col] = [];
                        } else if (oCurElem.value == 0x40) { // @
                            row++;
                            col = 0;
                            arrContent[row] = [];
                            arrContent[row][col] = [];
                        } else {
                            arrContent[row][col].push(oCurElem);
                        }
                    }
                }
                if ((row < 1 && Del && bEqArray) || (row < 1 && !bEqArray)) {
                    Elements.splice(0,0, symbol);  
                    Param.Bracket[0].splice(0,1);
                    Param.Bracket[1].splice(0,1);
                    break;
                }
                var props = new CMathMatrixPr();
                props.row = row + 1;
                props.mcs = mcs;
                props.ctrPrp = this.TextPr.Copy();
                Element = new CMathMatrix(props);
                for (var i = 0; i < arrContent.length; i++) {
                    for (var j = 0; j < arrContent[i].length; j++) {
                        var El = Element.getElement(i,j);
                        this.private_PackTextToContent(El, arrContent[i][j], false)
                    }
                }
            }
            if (Del) {
                var oDelElem = Del.getBase(0);
                oDelElem.addElementToContent(Element);
                Elements.splice(0,Elements.length, Del);
            } else {
                Elements.splice(0,Elements.length, Element);
            }
            Param.Bracket[0].splice(0,1);
            Param.Bracket[1].splice(0,1);
            break;
        case MATH_ACCENT: 
            var props = new CMathAccentPr();
            props.chr = Elements.splice(0,1)[0][0].value;
            var oAccent = new CAccent(props);
            var oBase = oAccent.getBase();
            this.private_PackTextToContent(oBase, Elements[0], true);
            Elements.splice(0,Elements.length, oAccent);
            break;
        case MATH_LIMIT:
            var props = {
                ctrPrp : this.TextPr.Copy(),
                type : Param.Props.type
            };
            if (Elements[1].length) {
                Elements[1].splice(Elements[1].length-1,1);
            }
            var Limit = new CLimit(props);
            var MathContent = Limit.getFName();
            this.private_PackTextToContent(MathContent, Elements[1], true);
            MathContent = Limit.getIterator();
            this.private_PackTextToContent(MathContent, Elements[0], true);
            Elements.splice(0,Elements.length, Limit);
            break;
        case MATH_FUNCTION:
            var Pr = {ctrPrp: this.TextPr.Copy()};
            var MathFunc = new CMathFunc(Pr);
            MathFunc.SetParagraph(this.Paragraph);
            var MathContent = MathFunc.getFName();
            Elements.arrName.pop();
            var funcName = "";
            for (var i = 0; i < Elements.arrName.length; i++) {
                funcName += String.fromCharCode(Elements.arrName[i].value);
            }
            var arrFunc = ["lim", "inf", "det", "gcd", "Pr", "min", "max", "sup"];
            //maybe add flag in private_PackTextToContent for skip private_AutoCorrectEquation inside function
            if (Elements.arrSup && Elements.arrSub) {
                Elements.arrName = Elements.arrName.concat(Elements.arrSup, Elements.arrSub);
                this.private_PackTextToContent(MathContent, Elements.arrName, true);
            } else if (Elements.arrSup || Elements.arrSub && !arrFunc.indexOf(funcName)) {
                var type = (Elements.arrSup) ? 1 : 0;
                var Limit = MathContent.Add_Limit({ctrPrp : Pr.ctrPrp, type : type}, funcName, null);
                MathContent = Limit.getIterator();
                if (type)
                Elements.arrSup.shift();
                else
                Elements.arrSub.shift();
                this.private_PackTextToContent(MathContent, (Elements.arrSup || Elements.arrSub), true);
            } else {
                this.private_PackTextToContent(MathContent, Elements.arrName, true);
            }
            if (Elements.arrArg) {
                MathContent = MathFunc.getArgument();
                this.private_PackTextToContent(MathContent, Elements.arrArg, true);
            }
            Elements.arrName = MathFunc;
            break;
    }
};
CMathContent.prototype.private_CanAutoCorrectTextFunc = function(AutoCorrectEngine) {
    var bActionIsSpace = (AutoCorrectEngine.ActionElement.value == 0x20) ? true : false;
    var foundedEl = null;
    var ElCount = AutoCorrectEngine.Elements.length;
    var Result = false;
    var FlagEnd = false;
    var Start = ElCount - 1;
    var RemoveCount = (bActionIsSpace) ? 1 : 0;
    var AutoCorrectCount = g_AutoCorrectMathFuncs.length;
    for (var nIndex = 0; nIndex < AutoCorrectCount; nIndex++) {
        var AutoCorrectElement = g_AutoCorrectMathFuncs[nIndex];
        var CheckStringLen = AutoCorrectElement.length;
        if (ElCount < CheckStringLen + 1) {
            continue;
        }
        var Found = true;
        // Начинаем проверять с конца строки
        for (var nStringPos = 0; nStringPos < CheckStringLen; nStringPos++) {
            var LastEl = AutoCorrectEngine.Elements[ElCount - nStringPos - 1 - 1];
            if (!LastEl.Element.IsText()) {
                FlagEnd = true;
                Found = false;
                break;
            }
            if (String.fromCharCode(LastEl.Element.value) !== AutoCorrectElement[CheckStringLen - nStringPos - 1]) {
                Found = false;
                break;
            }
        }
        if (Found === true) {
            foundedEl = AutoCorrectElement;
            FlagEnd = true;
        }
        if (FlagEnd) {
            break;
        }
    }
    if (Found) {
        if (!AutoCorrectEngine.StartHystory) {
            AutoCorrectEngine.StartHystory = true;
            var oParagraph = this.GetParagraph();
            var oLogicDocument = oParagraph ? oParagraph.LogicDocument : null;
            if(oLogicDocument) {
                oLogicDocument.StartAction(AscDFH.historydescription_Document_MathAutoCorrect);
            } else {
                History.Create_NewPoint(AscDFH.historydescription_Document_MathAutoCorrect);
            }
        }
        if (bActionIsSpace && foundedEl !== "lim") {
            Start = ElCount - foundedEl.length - 1;
            RemoveCount = foundedEl.length + 1;
            var Pr = {ctrPrp: new CTextPr()};	
            var MathFunc = new CMathFunc(Pr);	
            var MathContent = MathFunc.getFName();	
            var MathRun = new ParaRun(this.Paragraph, true);
            for (var nCharPos = 0, nTextLen = foundedEl.length; nCharPos < nTextLen; nCharPos++) {
                var oText = null;
                if (0x0026 == foundedEl.charCodeAt(nCharPos)) {
                    oText = new CMathAmp();
                } else {
                    oText = new CMathText(false);
                    oText.addTxt(foundedEl[nCharPos]);
                }
                MathRun.Add(oText, true);
            }	
            MathRun.Math_Apply_Style(STY_PLAIN);	
            MathContent.Internal_Content_Add(0, MathRun);	
            AutoCorrectEngine.Remove.push({Count:RemoveCount, Start:Start});	
            AutoCorrectEngine.Remove.total += RemoveCount;	
            AutoCorrectEngine.ReplaceContent.push(MathFunc);
        } else {
            var MathRun = new ParaRun(this.Paragraph, true);
            var Symbol = new CMathText(false);
            Symbol.add(0x2061);
            MathRun.Add(Symbol, true);
            MathRun.Apply_Pr(AutoCorrectEngine.TextPr);
            MathRun.Set_MathPr(AutoCorrectEngine.MathPr);
            AutoCorrectEngine.Remove.unshift({Count:RemoveCount, Start:Start});
            AutoCorrectEngine.ReplaceContent.unshift(MathRun);
        }
        Result = true;
    }
    return Result;
};
CMathContent.prototype.Clear_ContentChanges = function()
{
	this.m_oContentChanges.Clear();
};
CMathContent.prototype.Add_ContentChanges = function(Changes)
{
	this.m_oContentChanges.Add(Changes);
};
CMathContent.prototype.Refresh_ContentChanges = function()
{
	this.m_oContentChanges.Refresh();
};
CMathAutoCorrectEngine.prototype.private_PackTextToContent = function(Element, TempElements, bReplaceBrackets) {
    if (TempElements.length === undefined) {
        Element.Internal_Content_Add(0, TempElements);        
    } else {
        var len = TempElements.length - 1;
        if (len >= 2 && bReplaceBrackets) {
            if (g_MathLeftBracketAutoCorrectCharCodes[TempElements[0].value] && g_MathRightBracketAutoCorrectCharCodes[TempElements[len].value]) {
                TempElements.splice(0,1);
                TempElements.length--;
            }
        }
        this.private_AutoCorrectEquation(TempElements);
        var bNewRun = true;
        var MathRun = null;
        var PosElemnt = 0;
        var PosInRun = 0;
        for (var nPos = 0; nPos < TempElements.length; nPos++) {
            if (TempElements[nPos].value === undefined) {
                Element.Internal_Content_Add(PosElemnt, TempElements[nPos]);
                bNewRun = true;
                PosElemnt++;
                PosInRun = 0;
            } else {
                if (bNewRun) {
                    MathRun = new ParaRun(this.Paragraph, true);
                    MathRun.Set_Pr(this.TextPr.Copy());
                    MathRun.Set_MathPr(this.MathPr);
                    Element.Internal_Content_Add(PosElemnt, MathRun);
                    bNewRun = false;
                    PosElemnt++;
                    PosInRun = 0;
                }
                MathRun.Add_ToContent(PosInRun, TempElements[nPos]);
                PosInRun++;
            }
        }
    }  
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectDelimiter = function(CanMakeAutoCorrect) {
    var Elements = this.Elements.slice();
    //пока выключил, так как все скобки глубже первого уровня обработаются в private_AutoCorrectEquation
    while (this.Brackets[2]) {
        var lv = this.Brackets.length - 1;
        this.private_AutoCorrectPackDelimiter(this.Brackets[lv], Elements);
        this.Brackets.splice(lv,1);
    }
    for(var j = 0; j < this.Brackets[1]['left'].length; j++) {
        var props = new CMathDelimiterPr();
        props.column = 1;
        props.begChr = this.Brackets[1]['left'][j].bracket.value;
        props.endChr = this.Brackets[1]['right'][j].bracket.value;
        var oDelimiter = new CDelimiter(props);

        var oBase = oDelimiter.getBase();
        var TempElements = [];
        for (var i = this.Brackets[1]['right'][j].pos - 1; i > this.Brackets[1]['left'][j].pos; i--) {
            if (Elements[i]) {
                TempElements.splice(0, 0, Elements[i].Element);
            }
        }
        this.private_PackTextToContent(oBase, TempElements, false);
        this.Shift = Elements.length - 1 - this.Brackets[1]['right'][j].pos;
        if (this.ActionElement.value == 0x20) {
            this.Shift--;
        }
        var nRemoveCount = this.Brackets[1]['right'][j].pos - this.Brackets[1]['left'][j].pos + 1;
        if (CanMakeAutoCorrect) {
            nRemoveCount += this.Remove[0].Count;
        } else if (this.ActionElement.value == 0x20 && !this.Remove[0] && this.Brackets[1]['right'][j].pos >= (Elements.length - 2)) {
            nRemoveCount++;
        }
        this.Remove['total'] += nRemoveCount;
        var Start = this.Brackets[1]['left'][j].pos;
        this.Remove.unshift({Count:nRemoveCount, Start:Start});
        this.ReplaceContent.unshift(oDelimiter);
    }
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectPackDelimiter = function(data, Elements) {
    for(var j = 0; j < data['left'].length; j++) {
        var Pos = data['left'][j].pos - 1;
        var fSkip = false;
        while (Pos >= 0) {
            var Elem = Elements[Pos].Element.value;
            if (g_aMathAutoCorrectSkipBrackets[Elem]) {
                fSkip = true;
            } else if (g_MathLeftBracketAutoCorrectCharCodes[Elem]) {
                Pos--;
                continue;
            }
            break;
        }
        Pos = data['right'][j].pos + 1;
        var len = Elements.length;
        while (Pos < len) {
            var Elem = Elements[Pos].Element.value;
            if (q_aMathAutoCorrectAccentCharCodes[Elem] || Elem == 0x5E || Elem == 0x5F) {
                fSkip = true;
            } else if (g_MathRightBracketAutoCorrectCharCodes[Elem]) {
                Pos++;
                continue;
            }
            break;
        }
        if (fSkip) {
            continue;
        }
        var props = new CMathDelimiterPr();
        props.column = 1;
        props.begChr = data['left'][j].bracket.value;
        props.endChr = data['right'][j].bracket.value;
        var oDelimiter = new CDelimiter(props);

        var oBase = oDelimiter.getBase();
        var TempElements = [];
        for (var i = data['right'][j].pos - 1; i > data['left'][j].pos; i--) {
            if (Elements[i]) {
                TempElements.splice(0, 0, Elements[i].Element);
                Elements[i] = null;
            }
        }
        Elements[data['right'][j].pos] = null;
        Elements[data['left'][j].pos] = null;
        this.private_PackTextToContent(oBase, TempElements, false);
        Elements[data['left'][j].pos] = {Element : oDelimiter};
    }
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectFraction = function(buff) {
    this.private_CorrectBuffForFrac(buff);
    var Fraction = buff[0];
    var RemoveCount = buff[0].length;
    for (var i = 1; i < buff.length; i++) {
        var props = new CMathFractionPr();
        props.Set_FromObject(this.props);
        props.ctrPrp = this.TextPr.Copy();
        var tmp = new CFraction(props);

        var DenMathContent = tmp.getDenominatorMathContent();
        var NumMathContent = tmp.getNumeratorMathContent();
        RemoveCount += buff[i].length + ((this.props.type === 2) ? 2 : 1);

        this.private_PackTextToContent(DenMathContent, Fraction, true);
        this.private_PackTextToContent(NumMathContent, buff[i], true);
        Fraction = tmp;
    }
    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.unshift({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(Fraction);
};
CMathAutoCorrectEngine.prototype.private_CorrectBuffForFrac = function(buff, Props) {
    var props = Props || this.props;
    for (var i = 0; i < buff.length; i++) {
        var end = ((buff[i].length - 1) >= 0) ? buff[i].length - 1 : 0;
        if (!buff[i][end]) {
            continue;
        }
        if (buff[i][end].value === 0x002F || (buff[i][end].value === 0x2044 && props.type === 1) || (buff[i][end].value === 0x00A6 && props.type === 3)) {
            if (buff[i][end-1] && (buff[i][end-1].value === 0x005C && props.type === 2)) { 
                buff[i].splice(end-1,2);
            } else {
                buff[i].splice(end,1);
            }
        }
    }
};
CMathAutoCorrectEngine.prototype.private_CorrectBuffForNary = function(buff) {
    for (var i = 0; i < buff.length; i++) {
        if (!buff[i][0]) {
            continue;
        }
        if (buff[i][0].value == 0x005E) {
            buff[i].Type = DEGREE_SUPERSCRIPT;
            buff[i].splice(0,1);
        } else if (buff[i][0].value == 0x005F) {
            buff[i].Type = DEGREE_SUBSCRIPT;
            buff[i].splice(0,1);
        }
    }
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectDegree = function(buff) {
    var oDegree = buff[0];
    var RemoveCount = buff[0].length;
    for (var i = 1; i < buff.length; i++) {
        var props = new CMathDegreePr();
        props.ctrPrp = this.TextPr.Copy();
        props.type = this.Kind;
        var tmp = new CDegree(props)
        var BaseContent = tmp.Content[0];
        var IterContent = tmp.Content[1];
        RemoveCount += buff[i].length + 1;
        //разделить контент если нет скокобок и сохранить оставшуюся часть
        this.private_PackTextToContent(BaseContent, buff[i], false);
        this.private_PackTextToContent(IterContent, oDegree, true);
        oDegree = tmp;
    }
    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(oDegree);
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectDegreeSubSup = function(buff) {
    var props = new CMathDegreePr();
    props.ctrPrp = this.TextPr.Copy();
    props.type = this.Kind;
    var oDegree = new CDegreeSubSup(props)

    var BaseContent = oDegree.Content[0];
    var IterDnContent = oDegree.Content[1];
    var IterUpContent = oDegree.Content[2];
    var RemoveCount = buff[0].length + buff[1].length + buff[2].length + 2;

    if (buff[0].Type == DEGREE_SUPERSCRIPT) {
        this.private_PackTextToContent(IterUpContent, buff[1], true);
        this.private_PackTextToContent(IterDnContent, buff[0], true);
    } else if (buff[0].Type == DEGREE_SUBSCRIPT) {
        this.private_PackTextToContent(IterUpContent, buff[0], true);
        this.private_PackTextToContent(IterDnContent, buff[1], true);
    }
    var BaseElems = buff[2];
    this.private_PackTextToContent(BaseContent, BaseElems, false);

    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(oDegree);
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectCNary = function(buff, bOff) {
    var RemoveCount = buff[0].length;
    RemoveCount += (buff[1])? buff[1].length + 1 : 0;
    RemoveCount += (buff[2])? buff[2].length + 1 : 0;
    var arrBase = [];
    if (bOff) {
        for (var i = buff[0].length - 1; i >= 0 ; i--) {
            if (buff[0][i].value == 0x2592) {
                buff[0].pop();
                break;
            }
            arrBase.unshift(buff[0].pop());
        }
    }
    var props = {};
    props.supHide = true;
    props.subHide = true;
    props.chr = (buff.length == 1) ? buff[0][0].value : buff[buff.length-1][0].value;

    if (buff[0].Type) {
        if (buff[0].Type == DEGREE_SUPERSCRIPT && buff[0].length) {
           props.supHide = false;
        }
        if (buff[0].Type == DEGREE_SUBSCRIPT && buff[0].length) {
           props.subHide = false;
        }
    }
    if (buff[1]) {
        if (buff[1].Type == DEGREE_SUPERSCRIPT && buff[1].length) {
            props.supHide = false;
        }
        if (buff[1].Type == DEGREE_SUBSCRIPT && buff[1].length) {
            props.subHide = false;
        }
    }

    props.ctrPrp = this.TextPr.Copy();
    var oNary = new CNary(props);

    var oSub = oNary.getLowerIterator();
    var oSup = oNary.getUpperIterator();
    var oBase = oNary.getBase();

    if (!props.supHide) {
        if (buff[0].Type == DEGREE_SUPERSCRIPT) {
            this.private_PackTextToContent(oSup, buff[0], true);
        } else if (buff[1].Type && buff[1].Type == DEGREE_SUPERSCRIPT) {
            this.private_PackTextToContent(oSup, buff[1], true);
        }
    }
    if (!props.subHide) {
        if (buff[0].Type == DEGREE_SUBSCRIPT) {
            this.private_PackTextToContent(oSub, buff[0], true);
        } else if (buff[1].Type && buff[1].Type == DEGREE_SUBSCRIPT) {
            this.private_PackTextToContent(oSub, buff[1], true);
        }
    }
    this.private_PackTextToContent(oBase, arrBase, true);

    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(oNary);
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectCRadical = function(buff) {
    var props = new CMathRadicalPr();
    var RemoveCount = buff[0].length;
    props.degHide = this.private_CorrectBuffForRadical(buff, false);
    props.ctrPrp = this.TextPr.Copy();
    var Radical = new CRadical(props);

    var Base = Radical.getBase();
    var Degree = Radical.getDegree();
    buff = buff[0];
    if (!props.degHide) {
        this.private_PackTextToContent(Degree, buff[1], false);
        this.private_PackTextToContent(Base, buff[0], true);
    } else {
        this.private_PackTextToContent(Base, buff, true);
    }
    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(Radical);
};
CMathAutoCorrectEngine.prototype.private_CorrectBuffForRadical = function(buff, inside, bracket) {
    var radical = buff[0].splice(0,1)[0].value;
    switch (radical) {
        case 0x221A:
            if (inside) {
                if (!bracket.length) {
                    return true;
                }
                var bOpenBrk = 0;
                var ampPos = null;
                for (var i = bracket[0] + 1; i < bracket[1]; i++) {
                    if (g_MathLeftBracketAutoCorrectCharCodes[buff[0][i].value]) {
                        bOpenBrk++;
                    } else if (g_MathRightBracketAutoCorrectCharCodes[buff[0][i].value]) {
                        bOpenBrk--;
                    } else if (buff[0][i].value === 0x26 && !bOpenBrk) {
                        if (ampPos === null) {
                            ampPos = i;
                        } else {
                            ampPos = null;
                            break;
                        }
                    }
                }
                if (ampPos) {
                    buff[0].splice(bracket[1],1);
                    buff[0].splice(0,1);
                    var tmp1 = buff[0].splice(0,ampPos-1);
                    var tmp2 = buff[0].splice(1,buff[0].length);
                    buff[0].splice(0,1,tmp2,tmp1);
                    return false;
                } else {
                    return true;
                }
            } else {
                if (!this.Brackets.length) {
                    return true;
                }
                var len = this.Brackets[1]['left'].length - 1;
                if (this.Brackets[1]['left'][len].pos - this.CurPos !== 1) {
                    return true;
                }
                var bOpenBrk = 0;
                var ampPos = null;
                for (var i = 1; i < this.Brackets[1]['right'][len].pos - this.CurPos - 1; i++) {
                    if (g_MathLeftBracketAutoCorrectCharCodes[buff[0][i].value]) {
                        bOpenBrk++;
                    } else if (g_MathRightBracketAutoCorrectCharCodes[buff[0][i].value]) {
                        bOpenBrk--;
                    } else if (buff[0][i].value === 0x26 && !bOpenBrk) {
                        if (ampPos === null) {
                            ampPos = i;
                        } else {
                            ampPos = null;
                            break;
                        }
                    }
                }
                if (ampPos) {
                    buff[0].splice(this.Brackets[1]['right'][len].pos - this.CurPos - 1,1);
                    buff[0].splice(0,1);
                    var tmp1 = buff[0].splice(0,ampPos-1);
                    var tmp2 = buff[0].splice(1,buff[0].length);
                    buff[0].splice(0,1,tmp2,tmp1);
                    return false;
                } else {
                    return true;
                }
            }
        case 0x221B:
            var MathText = new CMathText();
            MathText.addTxt("3");
            var tmp = buff[0].splice(0, buff[0].length);
            buff[0].push(tmp,[MathText]);
            return false;
        case 0x221C:
            var MathText = new CMathText();
            MathText.addTxt("4");
            var tmp = buff[0].splice(0, buff[0].length);
            buff[0].push(tmp,[MathText]);
            return false;
    }
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectBox = function(buff) {
    buff = buff[0];
    var RemoveCount = buff.length;
    var symbol = buff.splice(0,1)[0].value;
    var props = {};
    props.ctrPrp = this.TextPr.Copy();
    var Box = (symbol == 0x25A1) ? new CBox(props) : new CBorderBox(props);
    var Base = Box.getBase();
    this.private_PackTextToContent(Base, buff, true);
    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(Box);
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectGroupCharacter = function(buff) {
    var base = (buff[0].Type) ? buff[1] : buff [0];
    var degree = (buff[0].Type) ? buff[0] : [];
    if (!degree.length && (this.ActionElement.value == 0x5E || this.ActionElement.value == 0x5F)) {
        for (var i = 0; i < base.length; i++) {
            if (base[i].value == 0x5E || base[i].value == 0x5F) {
                return;
            }
        }
    }
    var RemoveCount = base.length + degree.length;
    var symbol = base.splice(0,1)[0].value;
    var props = {};
    switch (symbol) {
        case 0x23E0:
        case 0x23DC:
        case 0x23DE:
            props = {chr:symbol, pos:LOCATION_TOP, vertJc:VJUST_BOT};
            break;
        case 0x23DD:
            props = {chr:symbol};
            break;
    }
    props.ctrPrp = this.TextPr.Copy();
    var oGroupChr = new CGroupCharacter(props);
    var oBase = oGroupChr.getBase();
    this.private_PackTextToContent(oBase, base, true);
    if (degree.length) {
        props = {
            ctrPrp : this.TextPr.Copy(),
            type : (degree.Type == 1) ? 1 : 0
        };
        var Limit = new CLimit(props);
        var MathContent = Limit.getFName();
        MathContent.Add_Element(oGroupChr);
        MathContent = Limit.getIterator();
        this.private_PackTextToContent(MathContent, degree, true);
        this.ReplaceContent.unshift(Limit);
        RemoveCount++;
    } else {
        this.ReplaceContent.unshift(oGroupChr);
    }
    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectBar = function(buff) {
    buff = buff[0];
    var RemoveCount = buff.length;
    var symbol = buff.splice(0,1)[0].value;
    var props = (symbol === 0x00AF) ? {pos:LOCATION_TOP} : {};
    props.ctrPrp = this.TextPr.Copy();
    var oBar = new CBar(props);
    var oBase = oBar.getBase();
    this.private_PackTextToContent(oBase, buff, true);
    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(oBar);
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectMatrix = function(buff) {
    var Shift = 0;
    for (var k = 0; k < buff.length - 1; k++) 
    {
        var bEqArray = false;
        var buffer = buff[k];
        var RemoveCount = null;
        var symbol = null;
        for (var i = 0; i < buffer.length; i++) {
            if (!symbol && (g_aMathAutoCorrectEqArrayMatrix[buffer[i].value])) {
                symbol = i;
                if (buffer[i+1] && buffer[i+1].value != 0x28) {
                    break;
                }
            } else if (symbol !== null) {
                if (buffer[i].value == 0x29) {
                    RemoveCount = i;
                    break;
                }
            }
        }
        if (!RemoveCount) {
            RemoveCount = 1;
        } else {
            RemoveCount -= symbol;
        }
        symbol = buffer.shift().value;
        if (symbol == 0x24B8 || symbol == 0x2588) {
            bEqArray = true;
        }
        var Del = null;
        switch (symbol) {
            case 0x24B8:
                var props = new CMathDelimiterPr();
                props.begChr = 123;
                props.endChr = -1;
                props.column = 1;
                Del = new CDelimiter(props);
                break;
            case 0x24A8:
                var props = new CMathDelimiterPr();
                props.column = 1;
                Del = new CDelimiter(props);
                break;
            case 0x24A9:
                var props = new CMathDelimiterPr();
                props.column = 1;
                props.begChr = 0x2016;
                props.endChr = 0x2016;
                Del = new CDelimiter(props);
                break;
        }
        var Element = null;
        if (bEqArray) {
            var arrContent = [];
            var row = 0;
            arrContent[row] = [];
            if (RemoveCount > 1) {
                for (var i = 1; i < RemoveCount - 1; i++) {
                    var oCurElem = buffer[i];
                    if (oCurElem.value == 0x40) {  // @
                        row++;
                        arrContent[row] = [];
                    } else {
                        arrContent[row].push(oCurElem);
                    }
                }
                RemoveCount++;
            }
            if ((row < 1 && Del && bEqArray) || (row < 1 && !bEqArray)) {
                return;
            }
            var props = new CMathEqArrPr();
            props.row = row + 1;
            props.ctrPrp = this.TextPr.Copy();
            Element = new CEqArray(props);
            for (var i = 0; i < arrContent.length; i++) {
                var El = Element.getElement(i);
                //getElementMathContent
                this.private_PackTextToContent(El, arrContent[i], false);
            }
        } else {
            var arrContent = [];
            var col = 0;
            var row = 0;
            var mcs = [];
            arrContent[row] = [];
            arrContent[row][col] = [];
            mcs[0] = {count: 1, mcJc: 0};
            if (RemoveCount > 1) {
                for (var i = 1; i < RemoveCount - 1; i++) {
                    var oCurElem = buffer[i];
                    if (oCurElem.value == 0x26) {  // &
                        col++;
                        if (col+1 > mcs[0].count) {
                            mcs[0] = {count: col+1, mcJc: 0};
                        }
                        arrContent[row][col] = [];
                    } else if (oCurElem.value == 0x40) { // @
                        row++;
                        col = 0;
                        arrContent[row] = [];
                        arrContent[row][col] = [];
                    } else {
                        arrContent[row][col].push(oCurElem);
                    }
                }
                RemoveCount++;
            }
            if (row < 1) {
                return;
            }
            var props = new CMathMatrixPr();
            props.row = row + 1;
            props.mcs = mcs;
            props.ctrPrp = this.TextPr.Copy();
            Element = new CMathMatrix(props);
            for (var i = 0; i < arrContent.length; i++) {
                for (var j = 0; j < arrContent[i].length; j++) {
                    var El = Element.getElement(i,j);
                    this.private_PackTextToContent(El, arrContent[i][j], false)
                }
            } 
        }
        var Start = this.Elements.length - buffer.length - 2 - this.Remove.total - Shift;
        Shift += buffer.length - RemoveCount + 1;
        this.Remove.unshift({Count:RemoveCount, Start:Start});
        this.Remove.total += RemoveCount;
        if (Del) {
            var oDelElem = Del.getBase(0);
            oDelElem.addElementToContent(Element);
            this.ReplaceContent.unshift(Del);
        } else {
            this.ReplaceContent.unshift(Element);
        }
        if (this.ActionElement.value == 0x20 && k == buff.length - 2) {
            this.Remove.push({Count:1, Start:(this.Elements.length-1)});
            this.ReplaceContent.push(null);
        }
    }
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectAccent = function(buff) {
    buff = buff[0];
    var RemoveCount = buff.length;
    var CurPos = buff.length - this.props.skip + 1;
    var props = new CMathAccentPr();
    props.chr = buff[CurPos].value;
    var oAccent = new CAccent(props);
    var oBase = oAccent.getBase();
    var TempElements = buff.splice(0,CurPos);
    var Start = this.Elements.length - RemoveCount - 1;
    RemoveCount = TempElements.length + 1;
    this.private_PackTextToContent(oBase, TempElements, true);
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(oAccent);
    if (this.ActionElement.value == 0x20) {
        this.Remove.push({Count:1, Start:(this.Elements.length-1)});
        this.ReplaceContent.push(null);
    }
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectAboveBelow = function(buff) {
    var RemoveCount = buff[0].length + buff[1].length + 1;
    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var props = {
        ctrPrp : this.TextPr.Copy(),
        type : this.props.type
    };
    var Limit = new CLimit(props);
    var MathContent = Limit.getFName();
    this.private_PackTextToContent(MathContent, buff[1], true);
    MathContent = Limit.getIterator();
    this.private_PackTextToContent(MathContent, buff[0], true);
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(Limit);
};
CMathAutoCorrectEngine.prototype.private_AutoCorrectFunction = function(buff) {
    var Pr = {ctrPrp: this.TextPr.Copy()};
    var MathFunc = new CMathFunc(Pr);
    MathFunc.SetParagraph(this.Paragraph);
    var MathContent = MathFunc.getFName();
    var arrName = buff[buff.length-1],
        arrSup = null,
        arrSub = null,
        RemoveCount = arrName.length + 1,
        arrArg = null;
    for (var i = 0; i < buff.length - 1; i++) {
        RemoveCount += buff[i].length;
        if (buff[i].Type == 1) {
            RemoveCount ++;
            arrSup = buff[i];
        } else if (buff[i].Type == -1) {
            RemoveCount ++;
            arrSub = buff[i];
        } else {
            arrArg = buff[i];
        }
    }
    if (this.ActionElement.value == 0x20) {
        RemoveCount++;
    }
    var funcName = "";
    for (var i = 0; i < arrName.length; i++) {
        funcName += String.fromCharCode(arrName[i].value);
    }
    var arrFunc = ["lim", "inf", "det", "gcd", "Pr", "min", "max", "sup"];
    //maybe add flag in private_PackTextToContent for skip private_AutoCorrectEquation inside function
    if (arrSup && arrSub) {
        var tmpText = new CMathText(false);
        tmpText.add(0x5E);
        arrSup.unshift(tmpText);
        tmpText = new CMathText(false);
        tmpText.add(0x5F);
        arrSub.unshift(tmpText);
        arrName = arrName.concat(arrSup, arrSub);
        this.private_PackTextToContent(MathContent, arrName, true);
    } else if (arrSup || arrSub && !arrFunc.indexOf(funcName)) {
        var type = (arrSup) ? 1 : 0;
        var Limit = MathContent.Add_Limit({ctrPrp : Pr.ctrPrp, type : type}, funcName, null);
        MathContent = Limit.getIterator();
        this.private_PackTextToContent(MathContent, (arrSup || arrSub), true);
    } else {
        this.private_PackTextToContent(MathContent, arrName, true);
    }
    if (arrArg) {
        MathContent = MathFunc.getArgument();
        this.private_PackTextToContent(MathContent, arrArg, true);
    }
    var Start = this.Elements.length - RemoveCount - this.Shift;
    this.Remove.push({Count:RemoveCount, Start:Start});
    this.ReplaceContent.unshift(MathFunc);
};

// CMathAutoCorrectEngine.prototype.private_AutoCorrectPhantom = function(buff) {
//     var RemoveCount = buff[0].length + 1;
//     if (this.ActionElement.value == 0x20) {
//         RemoveCount++;
//     }

//     var props = new CMathPhantomPr();
//     props.Set_FromObject(this.props);
//     var oPhantom = new CPhantom(props);

//     var oBase = oPhantom.getBase();

//     this.private_PackTextToContent(oBase, buff[0], true);
    
//     var Start = this.Elements.length - RemoveCount - this.Shift;

//     this.Remove.push({Count:RemoveCount, Start:Start});
//     this.ReplaceContent.unshift(oPhantom);

// };
CMathAutoCorrectEngine.prototype.private_CanAutoCorrectEquation = function(CanMakeAutoCorrect) {
    var buffer = [];
    var CurLvBuf = 0;
    buffer[CurLvBuf] = [];
    var lvBrackets = 1;
    var bBrackOpen = false;
    var bOff = false;
    var bSecSpace = false;
    while (this.CurPos >= 0) {
        var Elem = this.Elements[this.CurPos].Element;
        if (Elem.value === undefined) {
            buffer[CurLvBuf].splice(0, 0, Elem);
        } else if (Elem.value === 0x002F) { // /
            if (this.Type == MATH_LIMIT || bOff) {
                break;
            }
            if (Elem == this.ActionElement) {
                this.CurPos--;
                this.Shift = 1;
                continue;
            }
            if (g_aMathAutoCorrectNotDoFraction[this.ActionElement.value]) {
                //дробь не закончена и рано еще делать автозамену
                this.CurPos--;
                buffer[CurLvBuf].splice(0, 0, Elem);
                continue;
            }
            if ((this.Type == MATH_DEGREE || this.Type == MATH_DEGREESubSup) && !bBrackOpen) {
                break;
            }
            if (this.CurPos - 1 > 0) {
                var tmp = this.Elements[this.CurPos-1].Element;
                if (tmp.Type === para_Math_BreakOperator && tmp.value === 0x005C && !bBrackOpen) { // \
                    this.CurPos--;
                    CurLvBuf++;
                    buffer[CurLvBuf] = [];
                    buffer[CurLvBuf].splice(0, 0, Elem);
                    if (this.Type !== MATH_FRACTION || (this.Type === MATH_FRACTION && this.props.type === LINEAR_FRACTION) ) {
                        this.Type = MATH_FRACTION;
                        this.props = {type: LINEAR_FRACTION};
                        this.CurPos --;
                        buffer[CurLvBuf].splice(0, 0, tmp);
                    } else {
                        this.props = {};
                    }
                    continue;
                }
            }
            if (!bBrackOpen) {
                this.Type = MATH_FRACTION;
                this.props = {};
            }
            this.CurPos--;
            if (!bBrackOpen) {
                CurLvBuf++;
                buffer[CurLvBuf] = [];
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            continue;
        } else if (Elem.value === 0x2044) { // fraction
            if (this.Type == MATH_LIMIT || bOff) {
                break;
            }
            if (g_aMathAutoCorrectNotDoFraction[this.ActionElement.value]) {
                //дробь не закончена и рано еще делать автозамену
                this.CurPos--;
                buffer[CurLvBuf].splice(0, 0, Elem);
                continue;
            }
            if (!bBrackOpen) {
                this.Type = MATH_FRACTION;
                this.props = {type: SKEWED_FRACTION};
            }
            this.CurPos--;
            CurLvBuf++;
            buffer[CurLvBuf] = [];
            buffer[CurLvBuf].splice(0, 0, Elem);
            continue;
        } else if (0x00A6 === Elem.value) { // fraction
            if (this.Type == MATH_LIMIT || bOff) {
                break;
            }
            if (g_aMathAutoCorrectNotDoFraction[this.ActionElement.value]) {
                //дробь не закончена и рано еще делать автозамену
                this.CurPos--;
                buffer[CurLvBuf].splice(0, 0, Elem);
                continue;
            }
            if (!bBrackOpen) {
                this.Type = MATH_FRACTION;
                this.props = {type: NO_BAR_FRACTION};
                CurLvBuf++;
                buffer[CurLvBuf] = [];
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            this.CurPos--;
            continue;
        } else if (g_MathRightBracketAutoCorrectCharCodes[Elem.value]) { // right bracket
            //если справа от скобки не "^","_" и тип степень, то сделать автозамену степени до этой скобки
            if (this.Type == MATH_DEGREE || this.Type == MATH_DEGREESubSup) {
                var tmp = this.Elements[this.CurPos+1].Element;
                if (tmp && tmp.value != 0x005E && tmp.value != 0x005F) {
                    break;
                }
            }
            if (!this.Brackets[lvBrackets]) {
                this.Brackets[lvBrackets] = {};	
                this.Brackets[lvBrackets]['left'] = [];	
                this.Brackets[lvBrackets]['right'] = [];
            }
            this.Brackets[lvBrackets]['right'].push({bracket : Elem, pos: this.CurPos});
            lvBrackets++;
            buffer[CurLvBuf].splice(0, 0, Elem);
            bBrackOpen = true;
            if (this.Type === null && !g_aMathAutoCorrectDoNotDelimiter[this.ActionElement.value]) {
                if (bOff) {
                    break;
                }
                this.Type = MATH_DELIMITER;
            }
        }  else if (g_MathLeftBracketAutoCorrectCharCodes[Elem.value]) { // left bracket
            if (lvBrackets > 1) {
                lvBrackets--;
            }
            if (!this.Brackets[lvBrackets] || lvBrackets == 1 && this.Brackets[1]['left'].length == this.Brackets[1]['right'].length) { // if firs brackesk is left
                break;
            }
            if (this.Brackets[lvBrackets]['left'].length === this.Brackets[lvBrackets]['right'].length && this.Brackets[lvBrackets]['right'].length) {
                buffer[CurLvBuf].splice(0, 0, Elem);
                this.CurPos--;
            } else if (this.Brackets[lvBrackets]['left'].length < this.Brackets[lvBrackets]['right'].length) {
                this.Brackets[lvBrackets]['left'].push({bracket : Elem, pos: this.CurPos});
                if (lvBrackets == 1) {
                    bBrackOpen = false;
                }
                buffer[CurLvBuf].splice(0, 0, Elem);
                this.CurPos--;
            }
            continue;  
        } else if  (Elem.value === 0x005E) { // ^
            if ((this.Type == MATH_DEGREE && (buffer[CurLvBuf-1] && buffer[CurLvBuf-1].Type == DEGREE_SUPERSCRIPT)) ) {
                break;
                //remove this if you need work with ^^^...
            }
            if (Elem === this.ActionElement) {
                this.Shift = 1;
                this.CurPos--;
                continue;
            }
            var bSkip = false;
            if (this.CurPos - 1 >= 0 && this.Type != MATH_DEGREESubSup) {
                var tmp = this.Elements[this.CurPos-1].Element;
                if (tmp.value === 0x2061)
                    bSkip = true;
                if (this.Type == MATH_DEGREE) 
                    this.Type = null;
                buffer[CurLvBuf].Type = DEGREE_SUPERSCRIPT;
            }
            if (this.Type == MATH_FRACTION || g_aMathAutoCorrectDoNotDegree[this.ActionElement.value] || bSkip) {
                if (!bSkip)
                    buffer[CurLvBuf].splice(0,0,Elem);
                this.CurPos--;
                continue;
            }
            //если скобки одинаковые - то степень, разные - delimiter
            if (this.Type == MATH_RADICAL || this.Type == MATH_NARY || this.Type == MATH_DELIMITER || this.Type == MATH_EQ_ARRAY || this.Type == MATH_MATRIX || !this.Type) {
                if (!bBrackOpen) {
                    this.Kind = DEGREE_SUPERSCRIPT;
                    this.Type = MATH_DEGREE;
                }
                
            }
            if (this.Type == MATH_DEGREE && (buffer[CurLvBuf-1] && buffer[CurLvBuf-1].Type == DEGREE_SUBSCRIPT)) {
                if (!bBrackOpen) {
                    this.Kind = DEGREE_SubSup;
                    this.Type = MATH_DEGREESubSup;
                }
            }
            buffer[CurLvBuf].Type = DEGREE_SUPERSCRIPT;
            this.CurPos--;
            if (!bBrackOpen) {
                CurLvBuf++;
                buffer[CurLvBuf] = [];
            } else {
                buffer[CurLvBuf].splice(0,0,Elem);
            }
            continue;
        } else if (Elem.value === 0x005F) { // _
            if ((this.Type == MATH_DEGREE && (buffer[CurLvBuf-1] && buffer[CurLvBuf-1].Type == DEGREE_SUBSCRIPT))) {
                //remove this if you need work with ___...
            }
            if (Elem === this.ActionElement) {
                this.Shift = 1;
                this.CurPos--;
                continue;
            }
            var bSkip = false;
            if (this.CurPos - 1 >= 0 && this.Type != MATH_DEGREESubSup) {
                var tmp = this.Elements[this.CurPos-1].Element;
                if (tmp.value === 0x2061)
                    bSkip = true;
                if (this.Type == MATH_DEGREE) {
                    this.Type = null;
                }
                buffer[CurLvBuf].Type = DEGREE_SUBSCRIPT;
            }
            if (this.Type == MATH_FRACTION || g_aMathAutoCorrectDoNotDegree[this.ActionElement.value] || bSkip) {
                if (!bSkip)
                    buffer[CurLvBuf].splice(0,0,Elem);
                this.CurPos--;
                continue;
            }
            //если скобки одинаковые - то степень, разные - delimiter
            if (this.Type == MATH_RADICAL || this.Type == MATH_NARY || this.Type == MATH_DELIMITER || this.Type == MATH_EQ_ARRAY || this.Type == MATH_MATRIX || !this.Type) {
                if (!bBrackOpen) {
                    this.Kind = DEGREE_SUBSCRIPT;
                    this.Type = MATH_DEGREE;
                }
            }
            if (this.Type == MATH_DEGREE && (buffer[CurLvBuf-1] && buffer[CurLvBuf-1].Type == DEGREE_SUPERSCRIPT)) {
                if (!bBrackOpen) {
                    this.Kind = DEGREE_SubSup;
                    this.Type = MATH_DEGREESubSup;
                }
            }
            buffer[CurLvBuf].Type = DEGREE_SUBSCRIPT;
            this.CurPos--;
            if (!bBrackOpen) {
                CurLvBuf++;
                buffer[CurLvBuf] = [];
            } else {
                buffer[CurLvBuf].splice(0,0,Elem);
            }
            continue;
        } else if (q_aMathAutoCorrectControlAggregationCodes[Elem.value]) { // cNary
            if (this.Type == MATH_LIMIT) {
                break;
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            if (g_aMathAutoCorrectNotDoCNary[this.ActionElement.value]) {
                this.CurPos--;
                continue;
            }
            if (!bBrackOpen && this.Type !== MATH_FRACTION) {
                this.Type = MATH_NARY;
                break;
            } else {
                this.CurPos--;
                continue;
            }
        } else if (g_aMathAutoCorrectRadicalCharCode[Elem.value]) { // sqrt
            // если на 1 уровне больше одной пары скобок, то делать скобки, а не корень
            if (((this.Type == MATH_DEGREE || this.Type == MATH_DEGREESubSup || this.Type == MATH_LIMIT) && !bBrackOpen) || bOff) {
                break;
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            if (g_aMathAutoCorrectDoNotRadical[this.ActionElement.value]) {
                this.CurPos--;
                continue;
            }
            if (!bBrackOpen && this.Type !== MATH_FRACTION) {
                this.Type = MATH_RADICAL;
                break;
            }
            this.CurPos--;
            continue;
        } else if (Elem.value == 0x25A1 || Elem.value == 0x25AD) { // box OR border box
            if (((this.Type == MATH_DEGREE || this.Type == MATH_DEGREESubSup || this.Type == MATH_LIMIT) && !bBrackOpen) || bOff) {
                break;
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            if (this.Type == MATH_FRACTION && !bBrackOpen) {
                break;
            }
            if (g_aMathAutoCorrectDoNotBox[this.ActionElement.value]) {
                this.CurPos--;
                continue;
            }
            if (!bBrackOpen && this.Type !== MATH_FRACTION) {
                this.Type = MATH_BOX;
                break;
            }
            this.CurPos--;
            continue;
        } else if (g_aMathAutoCorrectGroupChar[Elem.value]) { // group character
            if ((this.Type == MATH_GROUP_CHARACTER && !bBrackOpen) || bOff) {
                break;
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            if (this.Type == MATH_LIMIT) {
                break;
            }
            if ((this.Type == MATH_DEGREE || this.Type == MATH_DEGREESubSup) && !bBrackOpen) {
                if (this.Type == MATH_DEGREE) {
                    this.Type = MATH_GROUP_CHARACTER;
                }
                break;
            }
            if (g_aMathAutoCorrectDoNotGroupChar[this.ActionElement.value]) {
                this.CurPos--;
                continue;
            }
            if (!bBrackOpen && this.Type !== MATH_FRACTION) {

                this.Type = MATH_GROUP_CHARACTER;
                break;
            }
            this.CurPos--;
            continue;
        } else if (Elem.value === 0x00AF || Elem.value === 0x2581) { // bar
            if((this.Type == MATH_BAR && !bBrackOpen) || bOff) {
                break;
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            if (this.Type == MATH_LIMIT) {
                break;
            }
            if ((this.Type == MATH_DEGREE || this.Type == MATH_DEGREESubSup) && !bBrackOpen) {
                break;
            }
            if (g_aMathAutoCorrectDoNotGroupChar[this.ActionElement.value]) {
                this.CurPos--;
                continue;
            }
            if (!bBrackOpen && this.Type !== MATH_FRACTION) {
                this.Type = MATH_BAR;
                break;
            }
            this.CurPos--;
            continue;
        } else if (g_aMathAutoCorrectEqArrayMatrix[Elem.value]) { // eq array and matrix
            buffer[CurLvBuf].splice(0, 0, Elem);
            if (this.Type == MATH_LIMIT || bOff) {
                break;
            }
            if((this.Type == MATH_MATRIX || this.Type === null || this.Type == MATH_DELIMITER) && !bBrackOpen) {
                CurLvBuf++;
                buffer[CurLvBuf] = [];
            }
            
            if (g_aMathAutoCorrectDoNotMatrix[this.ActionElement.value]) {
                this.CurPos--;
                continue;
            }
            if (!bBrackOpen && (this.Type === null || this.Type == MATH_DELIMITER)) {
                this.Type = MATH_MATRIX;
            }
            this.CurPos--;
            continue;
        } else if (q_aMathAutoCorrectAccentCharCodes[Elem.value]) { // accent
            if (((this.Type == MATH_DEGREE || this.Type == MATH_DEGREESubSup) && !bBrackOpen) || bOff) {
                break;
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            if (this.Type == MATH_LIMIT) {
                this.CurPos--;
                continue;
            }
            if (g_aMathAutoCorrectDoNotAccentNotClose[this.ActionElement.value]) {
                this.CurPos--;
                continue;
            }
            if (!bBrackOpen && this.Type !== MATH_FRACTION) {
                if (this.CurPos >= 1) {
                    var tmp = this.Elements[this.CurPos-1].Element;
                    if (g_MathRightBracketAutoCorrectCharCodes[tmp.value]) {
                        var countR = 1;
                        buffer[CurLvBuf].splice(0,0,tmp);
                        for (var i = this.CurPos - 2; i >= 0; i--) {
                            tmp = this.Elements[i].Element;
                            buffer[CurLvBuf].splice(0,0,tmp);
                            if (g_MathRightBracketAutoCorrectCharCodes[tmp.value]) {
                                countR++;
                            } if (g_MathLeftBracketAutoCorrectCharCodes[tmp.value]) {
                                countR--;
                                if (!countR) {
                                    break;
                                }
                            }
                        }
                    } else {
                        buffer[CurLvBuf].splice(0,0,tmp);
                    }
                    if (tmp !== this.Elements[this.CurPos-1].Element && countR) {
                        break;
                    }
                }
                
                this.props = {skip:this.Elements.length - this.CurPos};
                this.Type = MATH_ACCENT;
                break;
            }
            this.CurPos--;
            continue;
        } else if (Elem.value === 0x2534 || Elem.value === 0x252C) { // above and below
            if (this.Type == MATH_FRACTION || this.Type == MATH_MATRIX || this.Type == MATH_ACCENT || g_aMathAutoCorrectDoNotAbove[this.ActionElement.value]) {
                buffer[CurLvBuf].splice(0, 0, Elem);
                this.CurPos--;
                continue;
            }
            if((this.Type !== null && !bBrackOpen) || bOff) {
                break;
            }
            if (!bBrackOpen){
                CurLvBuf++;
                buffer[CurLvBuf] = [];
                this.props = (Elem.value === 0x2534) ? {type : LIMIT_UP} : {type : LIMIT_LOW};
                this.Type = MATH_LIMIT;
            }
            this.CurPos--;
            continue;
        } else if (Elem.value === 0x2061) { // 0x2061
            if (this.Type == MATH_FRACTION || bBrackOpen || g_aMathAutoCorrectDoNotMathFunc[this.ActionElement.value]) {
                buffer[CurLvBuf].splice(0, 0, Elem);
                this.CurPos--;
                continue;
            } else {
                if (this.Type == MATH_DELIMITER) {
                    this.Type = null;
                }
                if (this.Type !== null || bOff) {
                    break;
                }
                this.Type = MATH_FUNCTION;
                CurLvBuf++;
                buffer[CurLvBuf] = [];
                this.CurPos--;
                while (this.CurPos >= 0) {
                    Elem = this.Elements[this.CurPos].Element;
                    if (g_aMathAutoCorrectLatinAlph[Elem.value]) {
                        buffer[CurLvBuf].splice(0, 0, Elem);
                    } else {
                        break;
                    }
                    this.CurPos--;
                }
                break;
            }
        } else if (Elem.value === 0x2592) { // naryand (off)
            this.CurPos--;
            var skip = false;
            if (this.CurPos >= 0) {
                var tmp = this.Elements[this.CurPos].Element;
                if (tmp.value == 0x5E || tmp.value == 0x5F || tmp.value == 0x2534 || tmp.value == 0x252C) {
                    skip = true;
                }
            }
            if (!bBrackOpen && !skip) {
                if (this.Type !== null) {
                    if (this.Type == MATH_DEGREE || this.Type == MATH_DEGREESubSup || this.Type == MATH_LIMIT)
                        buffer[CurLvBuf].splice(0, 0, Elem);
                    break;
                }
                bOff = true;
            }
            buffer[CurLvBuf].splice(0, 0, Elem);
            continue;

        // } else if (Elem.value === 0x27E1) { // phantom
        //     if (this.Type == MATH_FRACTION) {
        //         buffer[CurLvBuf].splice(0, 0, Elem);
        //         this.CurPos--;
        //         continue;
        //     }
        //     if(this.Type !== null && !bBrackOpen) {
        //         break;
        //     }
        //     if (!bBrackOpen){
        //         this.props = {
        //             show : false,
        //             // transp: true,
        //             // zeroAsc : true,
        //             // zeroDesc : true,
        //             // zeroWid : true
        //         };
        //         // buffer[CurLvBuf].splice(0, 0, Elem);
        //         this.Type = MATH_PHANTOM;
        //         break;
        //     }
        //     this.CurPos--;
        //     continue;
        // } 
        } else if (g_aMathAutoCorrectFracCharCodes[Elem.value]) { // space + - # @ and another
            if (Elem === this.ActionElement) {
                if (Elem.value !== 0x0020) {
                    this.Shift = 1;
                } else if ((this.CurPos -1 > 0) && this.Elements[this.CurPos-1].Element.value == 0x20) {
                    break;
                }
                this.CurPos--;
                continue;
            }
            if (bBrackOpen || (!bSecSpace && this.Type == null)) {
                buffer[CurLvBuf].splice(0, 0, Elem);
                this.CurPos--;
                if (!bBrackOpen) {
                    this.Shift += buffer[CurLvBuf].length;
                    bSecSpace = true;
                    buffer = [];
                    CurLvBuf = 0;
                    buffer[CurLvBuf] = [];
                }
                continue;
            }
            break;
        }
        else {
            buffer[CurLvBuf].splice(0, 0, Elem);
        }
        this.CurPos--;
    }
    if (bBrackOpen || this.Type === null){
        return false;
    } else if (!this.StartHystory && ((this.Type !== MATH_DEGREE || this.Type !== MATH_DEGREESubSup) && !bOff)) {
        this.StartHystory = true;
        var oLogicDocument = this.Paragraph ? this.Paragraph.LogicDocument : null;
        if(oLogicDocument) {
            oLogicDocument.StartAction(AscDFH.historydescription_Document_MathAutoCorrect);
        } else {
            History.Create_NewPoint(AscDFH.historydescription_Document_MathAutoCorrect);
        }
    }
    var result = false;
    switch(this.Type) {
        case MATH_DELIMITER :
            if(!bBrackOpen && lvBrackets === 1) {
                this.private_AutoCorrectDelimiter(CanMakeAutoCorrect);
                result = true;
            } else {
                result = false;
            }
            break;
        case MATH_FRACTION:
            this.private_AutoCorrectFraction(buffer);
            result = true;
            break;
        case MATH_DEGREE:
            if (bOff) {
                result = false;
            } else {
                this.private_AutoCorrectDegree(buffer);
                result = true;
            }
            break;
        case MATH_DEGREESubSup:
            if (bOff) {
                result = false;
            } else {
                result = true;
                this.private_AutoCorrectDegreeSubSup(buffer);
            }
            break;
        case MATH_NARY:
            this.private_AutoCorrectCNary(buffer, bOff);
            result = true;
            break;
        case MATH_RADICAL:
            this.private_AutoCorrectCRadical(buffer);
            result = true;
            break;
        case MATH_BOX:
            this.private_AutoCorrectBox(buffer);
            result = true;
            break;
        case MATH_GROUP_CHARACTER:
            this.private_AutoCorrectGroupCharacter(buffer);
            result = true;
            break;
        case MATH_BAR:
            this.private_AutoCorrectBar(buffer);
            result = true;
            break;
        case MATH_MATRIX:
            this.private_AutoCorrectMatrix(buffer);
            result = true;
            break;
        case MATH_ACCENT:
            this.private_AutoCorrectAccent(buffer);
            result = true;
            break;
        case MATH_LIMIT:
            this.private_AutoCorrectAboveBelow(buffer);
            result = true;
            break;
        case MATH_FUNCTION:
            this.private_AutoCorrectFunction(buffer);
            result = true;
            break;
        // case MATH_PHANTOM:
        //     this.private_AutoCorrectPhantom(buffer);
        //     result = true;
        //     break;
        default :
            result = false;
    }
    return result;
};
CMathContent.prototype.private_ReplaceAutoCorrect = function(AutoCorrectEngine) {
    for (var i = AutoCorrectEngine.Remove.length - 1; i >= 0; i--) {
        var LastEl = null;
        var Start = AutoCorrectEngine.Remove[i].Start;
        var End = AutoCorrectEngine.Remove[i].Start + AutoCorrectEngine.Remove[i].Count - 1;
        var FirstEl = AutoCorrectEngine.Elements[Start];
        var bDelete = false;
        var counter = 0;
        var contPos = null;
        if (i == AutoCorrectEngine.Remove.length - 1 && this.Content[FirstEl.ElPos].Type == para_Math_Run) {
            contPos = this.Content[FirstEl.ElPos].State.ContentPos;
        }
        for (var nPos = End; nPos >= Start; nPos--) {
            LastEl = AutoCorrectEngine.Elements[nPos];
            if (undefined !== LastEl.Element.Parent && LastEl.Element.kind === undefined) {
                this.Content[LastEl.ElPos].Remove_FromContent(LastEl.ContPos, 1, true);
                counter++;
                if (this.Content[LastEl.ElPos].Is_Empty()) {
                    this.Remove_FromContent(LastEl.ElPos, 1);
                    bDelete = true;
                    counter = 0;
                    if (FirstEl.ElPos == LastEl.ElPos) {
                        FirstEl.ElPos--;
                    }
                }
            } else {
                this.Remove_FromContent(LastEl.ElPos, 1);
                bDelete = true;
                if (FirstEl.ElPos == LastEl.ElPos) {
                    FirstEl.ElPos--;
                }
            }
        }
        if (AutoCorrectEngine.ReplaceContent[i]) {
            if (FirstEl.Element.Type != para_Math_Composition) {
                if (!bDelete) {
                    var NewRun = this.Content[FirstEl.ElPos].Split2(FirstEl.ContPos);
                    if (!NewRun.Is_Empty()) {
                        this.Internal_Content_Add(FirstEl.ElPos + 1, NewRun, false);
                        NewRun.State.ContentPos = (contPos - counter - FirstEl.ContPos) >= 0 ? contPos - counter - FirstEl.ContPos : 0;
                        this.CurPos++;
                    }
                    if (this.Content[FirstEl.ElPos].Is_Empty()) {
                        this.Remove_FromContent(FirstEl.ElPos, 1);
                        FirstEl.ElPos--;
                    }
                }
                this.Internal_Content_Add(FirstEl.ElPos + 1, AutoCorrectEngine.ReplaceContent[i], false);
                this.CurPos++;
            } else {
                this.Internal_Content_Add(FirstEl.ElPos, AutoCorrectEngine.ReplaceContent[i], false);
                this.CurPos++;
            }
            if (AutoCorrectEngine.ReplaceContent[i].kind && i == AutoCorrectEngine.Remove.length - 1 && (AutoCorrectEngine.ReplaceContent[i].kind == MATH_NARY || AutoCorrectEngine.ReplaceContent[i].kind == MATH_FUNCTION)) {
                    var oContentElem = this.Content[FirstEl.ElPos + 1];
                    this.Correct_Content(true);
                    for (var k = this.Content.length - 1; k >= 0 ; k--) {
                        if (AutoCorrectEngine.ReplaceContent[i] === this.Content[k]) {
                            this.CurPos = k;
                            break;
                        }
                    }
                    var index = (AutoCorrectEngine.ReplaceContent[i].kind == 4) ? 2 : 1;
                    oContentElem.CurPos = index;
                    oContentElem.Content[index].MoveCursorToEndPos();
            } else {
                this.Correct_Content(true);
                if (this.Content[this.CurPos].kind !== undefined && this.Content[this.CurPos + 1] && this.Content[this.CurPos + 1].Type == para_Math_Run) {
                     /*&& i == AutoCorrectEngine.Remove.length - 1*/
                    this.CurPos++;
                    this.Content[this.CurPos].MoveCursorToStartPos();
                }
            }
        }
    }
};
CMathContent.prototype.GetTextContent = function(bSelectedText) {
	var arr = [], str = "", bIsContainsOperator = false, paraRunArr = [];
	var addText = function(value, bIsAddParenthesis, paraRun) {
		if (bIsAddParenthesis) {
			arr.push("(");
			str += "(";
			paraRunArr.push("(");
		}
		arr.push(value);
		str += value;
		if(undefined === paraRun) {
			paraRunArr.push(value);
		} else {
			paraRunArr.push(paraRun);
		}
		if (bIsAddParenthesis) {
			arr.push(")");
			str += ")";
			paraRunArr.push(")");
		}
	};

	var getAndPushTextContent = function(elem, bAddBrackets) {
		var tempStr = elem.GetTextContent();
		if (tempStr.str) {
			addText(tempStr.str, tempStr.bIsContainsOperator || bAddBrackets, tempStr.paraRunArr);
		}
    };
    
    var checkBracket = function(str) {
        var result = false;
        var MathLeftRighBracketsCharCodes = 
        [
            0x28, 0x29, 0x5B, 0x5D, 0x7B, 0x7D, 0x27E8, 0x27E9, 0x2329, 0x232A, 0x27E6, 0x27E7,
            0x27EA, 0x27EB, 0x2308, 0x2309, 0x230A, 0x230B, 0x3016, 0x3017, 0x2524, 0x251C
        ];
		for (var i = 0 ; i < MathLeftRighBracketsCharCodes.length; i++) {
            var templ = String.fromCharCode(MathLeftRighBracketsCharCodes[i]);
            if (str.indexOf(templ) !== -1) {
                result = true;
                break;
            }
        }
        return result;
	};

	var parseMathComposition = function(elem) {
        var tempStr;
        if(elem instanceof CDegree) {//степень
			//основание
            getAndPushTextContent(elem.getBase(), false);
            //знак
			addText((elem.Pr.type === DEGREE_SUPERSCRIPT) ? "^" : "_");
            //показатель
            getAndPushTextContent(elem.getIterator(), checkBracket(elem.getIterator().GetTextContent().str));
        } else if (elem instanceof CDegreeSubSup) {//^ и _
            var bAddBrackets = false;
            var base = elem.getBase();
            var lower = elem.getLowerIterator();
            var upper = elem.getUpperIterator();
            if (elem.Pr.type == 1) {
                if (checkBracket(elem.getBase().GetTextContent().str)) {
                    addText("〖");
                    bAddBrackets = true;
                }
                //основание
                getAndPushTextContent(base);
                if (bAddBrackets) {
                    addText("〗");
                }
                addText("^");
                getAndPushTextContent(upper, checkBracket(upper.GetTextContent().str));
                addText("_");
                //показатель _
                getAndPushTextContent(lower, checkBracket(lower.GetTextContent().str));
            } else {
                //показатель _
                getAndPushTextContent(lower, checkBracket(lower.GetTextContent().str));
                addText("_");
                //показатель ^
                getAndPushTextContent(upper, checkBracket(upper.GetTextContent().str));
                addText("^");
                if (checkBracket(base.GetTextContent().str)) {
                    addText("〖");
                    bAddBrackets = true;
                }
                //основание
                getAndPushTextContent(base);
                if (bAddBrackets) {
                    addText("〗");
                }
            }   
        } else if (elem instanceof CBar) {
            //символ Bar
            addText(String.fromCharCode((elem.Pr.pos) ? 9601 : 175));        
            //содержимое
            getAndPushTextContent(elem.getBase(), true);
        } else if (elem instanceof CBox || elem instanceof CBorderBox) {
            addText(String.fromCharCode((elem instanceof CBox) ? 9633 : 9645));
            //содержимое
            getAndPushTextContent(elem.getBase(), true);
        } else if (elem instanceof CEqArray) {
            addText(String.fromCharCode(9608));
            addText("(");
            for(var j = 0; j < elem.Content.length; j++) {
                if(para_Math_Content === elem.Content[j].Type) {
                    getAndPushTextContent(elem.Content[j]);
                    if (j !== elem.Content.length - 1)
                        addText(String.fromCharCode(64));
                }
            }
			addText(")");
        } else if (elem instanceof CDelimiter && elem.Pr.endChr == -1) {
            //символ Eq array
            addText(String.fromCharCode(9400));
            addText("(");
            var tempEl = null;
            for (var i = 0; i < elem.Content[0].Content.length; i++) {
                if (elem.Content[0].Content[i] instanceof CEqArray) {
                    tempEl = elem.Content[0].Content[i];
                    break;
                }
            }
            if (tempEl) {
                for(var j = 0; j < tempEl.Content.length; j++) {
                    if(para_Math_Content === tempEl.Content[j].Type) {
                        getAndPushTextContent(tempEl.Content[j]);
                        if (j !== tempEl.Content.length - 1)
                            addText(String.fromCharCode(64));
                    }
                }
            }
			addText(")");
        } else if (elem instanceof CGroupCharacter) {
            addText(String.fromCharCode(elem.Pr.chr || elem.operator.Get_CodeChr()));
            getAndPushTextContent(elem.getBase(), true);
        } else if (elem instanceof CAccent) {
            getAndPushTextContent(elem.getBase(), true);
            addText(String.fromCharCode(elem.Pr.chr || elem.operator.Get_CodeChr()));
        } else if (elem instanceof CDelimiter) {
            addText(String.fromCharCode(elem.Pr.begChr || elem.begOper.code || 40));
            for (var i = 0; i < elem.Content.length; i++) {
                getAndPushTextContent(elem.Content[i]);
                if (i != elem.Content.length - 1)
                    addText(String.fromCharCode(elem.sepOper.code))
            }
            addText(String.fromCharCode(elem.Pr.endChr || elem.endOper.code || 41));            
		} else if (elem instanceof CNary) {
            addText(String.fromCharCode(elem.Pr.chr || elem.getSign().chrCode));
            //верхняя граница суммирования
            if (!elem.Pr.supHide) {
				addText("^");             
                getAndPushTextContent(elem.getSupMathContent(), checkBracket(elem.getSupMathContent().GetTextContent().str));
            }
            //нижняя граница суммирования
			if (!elem.Pr.subHide) {
				addText("_");
                getAndPushTextContent(elem.getSubMathContent(), checkBracket(elem.getSubMathContent().GetTextContent().str));
			}
            addText(String.fromCharCode(9618));
            var bAddBrackets = false;
            if (checkBracket(elem.getBase().GetTextContent().str)) {
                addText("〖");
                bAddBrackets = true;
            }
            getAndPushTextContent(elem.getBase(), false);
            if (bAddBrackets) {
                addText("〗");
            }
        } else if (elem instanceof CFraction) {//дробь
            //числитель
			getAndPushTextContent(elem.getNumerator(), checkBracket(elem.getNumerator().GetTextContent().str));
            switch (elem.Pr.type) {
                case 0:
                    addText(String.fromCharCode(47));
                    break;
                case 1:
                    addText(String.fromCharCode(8260));
                    break;
                case 2:
                    addText(String.fromCharCode(92));
                    addText(String.fromCharCode(47));
                    break;
                case 3:
                    addText(String.fromCharCode(166));
                    break;
            }
			//знаменатель
            getAndPushTextContent(elem.getDenominator(), checkBracket(elem.getDenominator().GetTextContent().str));       
		} else if (elem instanceof CRadical) {//корень
			addText(String.fromCharCode(8730));
			//степень корня
			tempStr = elem.getDegree().GetTextContent(bSelectedText);
			var bOpenBr = false;
			if (tempStr.str) {
				addText("(");
				addText(tempStr.str, tempStr.bIsContainsOperator, tempStr.paraRunArr);
				bOpenBr = true;
			}
			if (tempStr.str) {
				addText(String.fromCharCode(38));
            }
            tempStr = elem.getBase().GetTextContent();
            if (tempStr.str !== "") {
                if (!bOpenBr && checkBracket(tempStr.str)) {
                    addText("(");
                    bOpenBr = true;
                }
                //подкоренное выражение
			    getAndPushTextContent(elem.getBase());
            } 
            if (bOpenBr) {
                addText(")");
            }
        } else if (elem instanceof CMathMatrix) {//matrix
            var symbol = [];
            for (var row = 0; row < elem.nRow; row++) {
				for(var col = 1; col < elem.nCol; col++) {
					symbol.push(String.fromCharCode(38));
                }
                if (row !== elem.nRow - 1)
                    symbol.push(String.fromCharCode(64));
			}
			addText(String.fromCharCode(9632));
			addText("(");
			for (var j = 0; j < elem.Content.length; j++) {
				if (para_Math_Content === elem.Content[j].Type) {
					getAndPushTextContent(elem.Content[j]);
					if (symbol[j]) {
						addText(symbol[j]);
					}
				}
			}
			addText(")");
		} else if(elem instanceof CMathFunc) {
			//функция
            getAndPushTextContent(elem.getFName());
            addText(String.fromCharCode(0x2061));
            var tmp = elem.getArgument();
            var bAddBrackets = false;
            if (checkBracket(tmp.GetTextContent().str)) {
                addText("〖");
                bAddBrackets = true;
            }
			//аргумент
            getAndPushTextContent(tmp, false);
            if (bAddBrackets) {
                addText("〗");
            }
		} else if(elem instanceof CLimit) {
			//функция
            getAndPushTextContent(elem.getFName(), false);
            addText((elem.Pr.type == 1) ? "┴" : "┬");
			//аргумент
            getAndPushTextContent(elem.getIterator(), true);
		}

	};

	var StartPos = 0, EndPos = this.Content.length;
	if (bSelectedText) {
		StartPos = (this.Selection.Use == true ? Math.min(this.Selection.StartPos, this.Selection.EndPos) : this.CurPos.ContentPos);
		EndPos   = (this.Selection.Use == true ? Math.max(this.Selection.StartPos, this.Selection.EndPos) : this.CurPos.ContentPos);
	}

	for (var i = StartPos; i <= EndPos; i++) {
		if (!this.Content[i]) {
			continue;
		}

		switch (this.Content[i].Type) {
			case para_Math_Run:
				if (this.Content[i].Content.length) {
					var StartContentPos = 0, EndContentPos = this.Content[i].Content.length;
					if (bSelectedText) {
						StartContentPos = (this.Content[i].Selection.Use == true ? Math.min(this.Content[i].Selection.StartPos, this.Content[i].Selection.EndPos) : this.Content[i].CurPos.ContentPos);
						EndContentPos   = (this.Content[i].Selection.Use == true ? Math.max(this.Content[i].Selection.StartPos, this.Content[i].Selection.EndPos) : this.Content[i].CurPos.ContentPos);
					}
					var string = "";
					for (var j = StartContentPos; j < EndContentPos; j++) {
						if (!this.Content[i].Content[j]) {
							continue;
						}
						if (this.Content[i].Content[j].Type === para_Math_Text) {
							string += String.fromCharCode(this.Content[i].Content[j].value);
						} else if (this.Content[i].Content[j].Type === para_Math_BreakOperator) {
							string += String.fromCharCode(this.Content[i].Content[j].value);
							// bIsContainsOperator = true;
						}
					}
					addText(string, null, this.Content[i]);
				}
				break;
			case para_Math_Composition:
				parseMathComposition(this.Content[i]);
				break;
		}
	}
	return {str: str, bIsContainsOperator: bIsContainsOperator, paraRunArr: paraRunArr};
};
function CMathAutoCorrectEngine(Elem, CurPos, Paragraph) {

    this.ActionElement    = Elem;                                               // элемент на которотом срабатывает автодополнение
    this.CurElement       = CurPos;                                             // индекс текущего элемента, где стоит курсор
    this.CurPos           = null;                                               // индекс элемента с которого будет начинаться автозамена
    this.Elements         = [];                                                 // елементы для обработки
    this.Brackets         = [];                                                 // скобки по уровням
    this.Type			  = null;                                               // тип автозамены
    this.Kind			  = null;                                               // для автозамены со степенью
    this.props            = {};                                                 // для автозамены с дробью
    this.Paragraph        = Paragraph;                                          // параграф в котором находится формула
    this.Remove           = [];                                                 // массив индексов элементов для удаления из основного контента
    this.Remove['total']  = 0;                                                  // общее число элементов для удаления
    this.ReplaceContent   = [];                                                 // элементы для вставки в основной контент после автозамены
    this.Shift 			  = 0;                                                  // отступ
    this.StartHystory     = false;                                              // флаг, обозначающий была ли уже создана точка в истории автозаменой
    this.IntFlag          = window['AscCommonWord'].b_DoAutoCorrectMathSymbols; // флаг из интерфейса делать ли автозамену  символов из списка
    this.RepCharsCode     = [];                                                 // массив символов, добавленных в документ при автозамене
    this.TextPr           = null;
    this.MathPr           = null;
};
CMathAutoCorrectEngine.prototype.private_Add_Element = function(Content) {
    var nCount = this.CurElement;
    for (var i = nCount; i >= 0; i--) {
        if (Content[i].Type === 49) {
            var kStart = (i === nCount) ? Content[i].State.ContentPos : Content[i].Content.length;
            for (var k = kStart - 1; k >= 0; k--) {
                this.Elements.unshift({Element: Content[i].Content[k], ElPos: i, ContPos: k});
            }
        } else {
            this.Elements.unshift({Element: Content[i], ElPos: i});
        }
    }
};

var g_DefaultAutoCorrectMathFuncs =
[
    'arcsin', 'asin', 'sin', 'arcsinh', 'asinh', 'sinh',
    'arcsec', 'sec', 'asec', 'arcsech', 'asech','sech',
    'arccos', 'acos', 'cos', 'arccosh','acosh', 'cosh',
    'arccsc', 'acsc', 'csc', 'arccsch', 'acsch', 'csch',
    'arctan', 'atan', 'tan', 'arctanh', 'atanh', 'tanh',
    'arccot', 'acot', 'cot', 'arccoth', 'acoth', 'coth',
    'arg', 'det', 'exp', 'inf', 'lim', 'min',
    'def', 'dim', 'gcd', 'ker', 'log', 'Pr',
    'deg', 'erf', 'hom', 'lg', 'ln', 'max', 'sup'
];
var g_DefaultAutoCorrectMathSymbolsList =
[
    ['!!', 0x203C],
    ['...', 0x2026],
    ['::', 0x2237],
    [':=', 0x2254],
    ['/<', 0x226E],
    ['/>', 0x226F],
    ['/=', 0x2260],
    ['\\above', 0x2534],
    ['\\acute', 0x0301],
    ['\\aleph', 0x2135],
    ['\\alpha', 0x03B1],
    ['\\Alpha', 0x0391],
    ['\\amalg', 0x2210],
    ['\\angle', 0x2220],
    ['\\aoint', 0x222E],
    ['\\approx', 0x2248],
    ['\\asmash', 0x2B06],
    ['\\ast', 0x2217],
    ['\\asymp', 0x224D],
    ['\\atop', 0x00A6],
    ['\\bar', 0x0305],
    ['\\Bar', 0x033F],
    ['\\because', 0x2235],
    ['\\begin', 0x3016],
    ['\\below', 0x252C],
    ['\\bet', 0x2136],
    ['\\beta', 0x03B2],
    ['\\Beta', 0x0392],
    ['\\beth', 0x2136],
    ['\\bigcap', 0x22C2],
    ['\\bigcup', 0x22C3],
    ['\\bigodot', 0x2A00],
    ['\\bigoplus', 0x2A01],
    ['\\bigotimes', 0x2A02],
    ['\\bigsqcup', 0x2A06],
    ['\\biguplus', 0x2A04],
    ['\\bigvee', 0x22C1],
    ['\\bigwedge', 0x22C0],
    ['\\binomial',[0x0028, 0x0061, 0x002B, 0x0062, 0x0029, 0x005E, 0x005E, 0x003D, 0x2211, 0x005F, 0x0028, 0x006B, 0x003D, 0x0030, 0x0029, 0x005E, 0x006E, 0x0020, 0x2592, 0x0028, 0x006E, 0x00A6, 0x006B, 0x0029, 0x0061, 0x005E, 0x006B, 0x0020, 0x0062, 0x005E, 0x0028, 0x006E, 0x002D, 0x006B, 0x0029]],
    ['\\bot', 0x22A5],
    ['\\bowtie', 0x22C8],
    ['\\box', 0x25A1],
    ['\\boxdot', 0x22A1],
    ['\\boxminus', 0x229F],
    ['\\boxplus', 0x229E],
    ['\\bra', 0x27E8],
    ['\\break', 0x2936],
    ['\\breve', 0x0306],
    ['\\bullet', 0x2219],
    ['\\cap', 0x2229],
    ['\\cbrt', 0x221B],
    ['\\cases', 0x24B8],
    ['\\cdot', 0x22C5],
    ['\\cdots', 0x22EF],
    ['\\check', 0x030C],
    ['\\chi', 0x03C7],
    ['\\Chi', 0x03A7],
    ['\\circ', 0x2218],
    ['\\close', 0x2524],
    ['\\clubsuit', 0x2663],
    ['\\coint', 0x2232],
    ['\\cong', 0x2245],
    ['\\coprod', 0x2210],
    ['\\cup', 0x222A],
    ['\\dalet', 0x2138],
    ['\\daleth', 0x2138],
    ['\\dashv', 0x22A3],
    ['\\dd', 0x2146],
    ['\\Dd', 0x2145],
    ['\\ddddot', 0x20DC],
    ['\\dddot', 0x20DB],
    ['\\ddot', 0x0308],
    ['\\ddots', 0x22F1],
    ['\\defeq', 0x225D],
    ['\\degc', 0x2103],
    ['\\degf', 0x2109],
    ['\\degree', 0x00B0],
    ['\\delta', 0x03B4],
    ['\\Delta', 0x0394],
    ['\\Deltaeq', 0x225C],
    ['\\diamond', 0x22C4],
    ['\\diamondsuit', 0x2662],
    ['\\div', 0x00F7],
    ['\\dot', 0x0307],
    ['\\doteq', 0x2250],
    ['\\dots', 0x2026],
    ['\\doublea', 0x1D552],
    ['\\doubleA', 0x1D538],
    ['\\doubleb', 0x1D553],
    ['\\doubleB', 0x1D539],
    ['\\doublec', 0x1D554],
    ['\\doubleC', 0x2102],
    ['\\doubled', 0x1D555],
    ['\\doubleD', 0x1D53B],
    ['\\doublee', 0x1D556],
    ['\\doubleE', 0x1D53C],
    ['\\doublef', 0x1D557],
    ['\\doubleF', 0x1D53D],
    ['\\doubleg', 0x1D558],
    ['\\doubleG', 0x1D53E],
    ['\\doubleh', 0x1D559],
    ['\\doubleH', 0x210D],
    ['\\doublei', 0x1D55A],
    ['\\doubleI', 0x1D540],
    ['\\doublej', 0x1D55B],
    ['\\doubleJ', 0x1D541],
    ['\\doublek', 0x1D55C],
    ['\\doubleK', 0x1D542],
    ['\\doublel', 0x1D55D],
    ['\\doubleL', 0x1D543],
    ['\\doublem', 0x1D55E],
    ['\\doubleM', 0x1D544],
    ['\\doublen', 0x1D55F],
    ['\\doubleN', 0x2115],
    ['\\doubleo', 0x1D560],
    ['\\doubleO', 0x1D546],
    ['\\doublep', 0x1D561],
    ['\\doubleP', 0x2119],
    ['\\doubleq', 0x1D562],
    ['\\doubleQ', 0x211A],
    ['\\doubler', 0x1D563],
    ['\\doubleR', 0x211D],
    ['\\doubles', 0x1D564],
    ['\\doubleS', 0x1D54A],
    ['\\doublet', 0x1D565],
    ['\\doubleT', 0x1D54B],
    ['\\doubleu', 0x1D566],
    ['\\doubleU', 0x1D54C],
    ['\\doublev', 0x1D567],
    ['\\doubleV', 0x1D54D],
    ['\\doublew', 0x1D568],
    ['\\doubleW', 0x1D54E],
    ['\\doublex', 0x1D569],
    ['\\doubleX', 0x1D54F],
    ['\\doubley', 0x1D56A],
    ['\\doubleY', 0x1D550],
    ['\\doublez', 0x1D56B],
    ['\\doubleZ', 0x2124],
    ['\\downarrow', 0x2193],
    ['\\Downarrow', 0x21D3],
    ['\\dsmash', 0x2B07],
    ['\\ee', 0x2147],
    ['\\ell', 0x2113],
    ['\\emptyset', 0x2205],
    ['\\emsp', 0x2003],
    ['\\end', 0x3017],
    ['\\ensp', 0x2002],
    ['\\epsilon', 0x03F5],
    ['\\Epsilon', 0x0395],
    ['\\eqarray', 0x2588],
    ['\\equiv', 0x2261],
    ['\\eta', 0x03B7],
    ['\\Eta', 0x0397],
    ['\\exists', 0x2203],
    ['\\forall', 0x2200],
    ['\\fraktura', 0x1D51E],
    ['\\frakturA', 0x1D504],
    ['\\frakturb', 0x1D51F],
    ['\\frakturB', 0x1D505],
    ['\\frakturc', 0x1D520],
    ['\\frakturC', 0x212D],
    ['\\frakturd', 0x1D521],
    ['\\frakturD', 0x1D507],
    ['\\frakture', 0x1D522],
    ['\\frakturE', 0x1D508],
    ['\\frakturf', 0x1D523],
    ['\\frakturF', 0x1D509],
    ['\\frakturg', 0x1D524],
    ['\\frakturG', 0x1D50A],
    ['\\frakturh', 0x1D525],
    ['\\frakturH', 0x210C],
    ['\\frakturi', 0x1D526],
    ['\\frakturI', 0x2111],
    ['\\frakturj', 0x1D527],
    ['\\frakturJ', 0x1D50D],
    ['\\frakturk', 0x1D528],
    ['\\frakturK', 0x1D50E],
    ['\\frakturl', 0x1D529],
    ['\\frakturL', 0x1D50F],
    ['\\frakturm', 0x1D52A],
    ['\\frakturM', 0x1D510],
    ['\\frakturn', 0x1D52B],
    ['\\frakturN', 0x1D511],
    ['\\frakturo', 0x1D52C],
    ['\\frakturO', 0x1D512],
    ['\\frakturp', 0x1D52D],
    ['\\frakturP', 0x1D513],
    ['\\frakturq', 0x1D52E],
    ['\\frakturQ', 0x1D514],
    ['\\frakturr', 0x1D52F],
    ['\\frakturR', 0x211C],
    ['\\frakturs', 0x1D530],
    ['\\frakturS', 0x1D516],
    ['\\frakturt', 0x1D531],
    ['\\frakturT', 0x1D517],
    ['\\frakturu', 0x1D532],
    ['\\frakturU', 0x1D518],
    ['\\frakturv', 0x1D533],
    ['\\frakturV', 0x1D519],
    ['\\frakturw', 0x1D534],
    ['\\frakturW', 0x1D51A],
    ['\\frakturx', 0x1D535],
    ['\\frakturX', 0x1D51B],
    ['\\fraktury', 0x1D536],
    ['\\frakturY', 0x1D51C],
    ['\\frakturz', 0x1D537],
    ['\\frakturZ', 0x2128],
    ['\\frown', 0x2311],
    ['\\funcapply', 0x2061],
    ['\\G', 0x0393],
    ['\\gamma', 0x03B3],
    ['\\Gamma', 0x0393],
    ['\\ge', 0x2265],
    ['\\geq', 0x2265],
    ['\\gets', 0x2190],
    ['\\gg', 0x226B],
    ['\\gimel', 0x2137],
    ['\\grave', 0x0300],
    ['\\hairsp', 0x200A],
    ['\\hat', 0x0302],
    ['\\hbar', 0x210F],
    ['\\heartsuit', 0x2661],
    ['\\hookleftarrow', 0x21A9],
    ['\\hookrightarrow', 0x21AA],
    ['\\hphantom', 0x2B04],
    ['\\hsmash', 0x2B0C],
    ['\\hvec', 0x20D1],
    ['\\identitymatrix', [0x0028, 0x25A0, 0x0028, 0x0031, 0x0026, 0x0030, 0x0026, 0x0030, 0x0040, 0x0030, 0x0026,0x0031, 0x0026, 0x0030, 0x0040, 0x0030, 0x0026, 0x0030, 0x0026, 0x0031, 0x0029, 0x0029]],
    ['\\ii', 0x2148],
    ['\\iiint', 0x222D],
    ['\\iint', 0x222C],
    ['\\iiiint', 0x2A0C],
    ['\\Im', 0x2111],
    ['\\imath', 0x0131],
    // ['/\\in', 0x2209],
    ['\\in', 0x2208],
    ['\\inc', 0x2206],
    ['\\infty', 0x221E],
    ['\\int', 0x222B],
    ['\\integral', [0x0031, 0x002F, 0x0032, 0x03C0, 0x222B, 0x005F, 0x0030, 0x005E, 0x0032, 0x03C0, 0x2592, 0x2146, 0x03B8, 0x0020, 0x0028, 0x0061, 0x002B, 0x0062, 0x0073, 0x0069, 0x006E, 0x0020, 0x03B8, 0x0029, 0x003D, 0x0031, 0x002F, 0x221A, 0x0028, 0x0061, 0x005E, 0x0032, 0x002D, 0x0062, 0x005E, 0x0032, 0x0029]],
    ['\\iota', 0x03B9],
    ['\\Iota', 0x0399],
    ['\\itimes', 0x2062],
    ['\\j', [0x004A, 0x0061, 0x0079]],
    ['\\jj', 0x2149],
    ['\\jmath', 0x0237],
    ['\\kappa', 0x03BA],
    ['\\Kappa', 0x039A],
    ['\\ket', 0x27E9],
    ['\\lambda', 0x03BB],
    ['\\Lambda', 0x039B],
    ['\\langle', 0x2329],
    ['\\\lbbrack', 0x27E6],
    ['\\lbrace', 0x007B],
    ['\\lbrack', 0x005B],
    ['\\lceil', 0x2308],
    ['\\ldiv', 0x2215],
    ['\\ldivide', 0x2215],
    ['\\ldots', 0x2026],
    ['\\le', 0x2264],
    ['\\left', 0x251C],
    ['\\leftarrow', 0x2190],
    ['\\Leftarrow', 0x21D0],
    ['\\leftharpoondown', 0x21BD],
    ['\\leftharpoonup', 0x21BC],
    ['\\leftrightarrow', 0x2194],
    ['\\Leftrightarrow', 0x21D4],
    ['\\leq', 0x2264],
    ['\\lfloor', 0x230A],
    ['\\lhvec', 0x20D0],
    ['\\limit', [0x006C, 0x0069, 0x006D, 0x005F, 0x0028, 0x006E, 0x2192, 0x221E, 0x0029, 0x2061, 0x3016, 0x0028, 0x0031, 0x002B, 0x0031, 0x002F, 0x006E, 0x0029, 0x005E, 0x006E, 0x3017, 0x003D, 0x0065]],
    ['\\ll', 0x226A],
    ['\\lmoust', 0x23B0],
    ['\\Longleftarrow', 0x27F8],
    ['\\Longleftrightarrow', 0x27FA],
    ['\\Longrightarrow', 0x27F9],
    ['\\lrhar', 0x21CB],
    ['\\lvec', 0x20D6],
    ['\\mapsto', 0x21A6],
    ['\\matrix', 0x25A0],
    ['\\medsp', 0x205F],
    ['\\mid', 0x2223],
    ['\\middle', 0x24DC],
    ['\\models', 0x22A8],
    ['\\mp', 0x2213],
    ['\\mu', 0x03BC],
    ['\\Mu', 0x039C],
    ['\\nabla', 0x2207],
    ['\\naryand', 0x2592],
    ['\\nbsp', 0x00A0],
    ['\\ne', 0x2260],
    ['\\nearrow', 0x2197],
    ['\\neq', 0x2260],
    ['\\ni', 0x220B],
    ['\\norm', 0x2016],
    ['\\notcontain', 0x220C],
    ['\\notelement', 0x2209],
    ['\\notin', 0x2209],
    ['\\nu', 0x03BD],
    ['\\Nu', 0x039D],
    ['\\nwarrow', 0x2196],
    ['\\o', 0x03BF],
    ['\\O', 0x039F],
    ['\\odot', 0x2299],
    ['\\of', 0x2592],
    ['\\oiiint', 0x2230],
    ['\\oiint', 0x222F],
    ['\\oint', 0x222E],
    ['\\omega', 0x03C9],
    ['\\Omega', 0x03A9],
    ['\\ominus', 0x2296],
    ['\\open', 0x251C],
    ['\\oplus', 0x2295],
    ['\\otimes', 0x2297],
    ['\\over', 0x002F],
    ['\\overbar', 0x00AF],
    ['\\overbrace', 0x23DE],
    ['\\overbracket', 0x23B4],
    ['\\overline', 0x00AF],
    ['\\overparen', 0x23DC],
    ['\\overshell', 0x23E0],
    ['\\parallel', 0x2225],
    ['\\partial', 0x2202],
    ['\\pmatrix', 0x24A8],
    ['\\perp', 0x22A5],
    ['\\phantom', 0x27E1],
    ['\\phi', 0x03D5],
    ['\\Phi', 0x03A6],
    ['\\pi', 0x03C0],
    ['\\Pi', 0x03A0],
    ['\\pm', 0x00B1],
    ['\\pppprime', 0x2057],
    ['\\ppprime', 0x2034],
    ['\\pprime', 0x2033],
    ['\\prec', 0x227A],
    ['\\preceq', 0x227C],
    ['\\prime', 0x2032],
    ['\\prod', 0x220F],
    ['\\propto', 0x221D],
    ['\\psi', 0x03C8],
    ['\\Psi', 0x03A8],
    ['\\qdrt', 0x221C],
    ['\\quadratic', [0x0078, 0x003d, 0x0028, 0x002d, 0x0062, 0x00B1, 0x221A, 0x0028, 0x0062, 0x005e, 0x0032, 0x002d, 0x0034, 0x0061, 0x0063, 0x0029, 0x0029, 0x002f, 0x0032, 0x0061]],
    ['\\rangle', 0x232A],
    ['\\Rangle', 0x27EB],
    ['\\ratio', 0x2236],
    ['\\rbrace', 0x007D],
    ['\\rbrack', 0x005D],
    ['\\Rbrack', 0x27E7],
    ['\\rceil', 0x2309],
    ['\\rddots', 0x22F0],
    ['\\Re', 0x211C],
    ['\\rect', 0x25AD],
    ['\\rfloor', 0x230B],
    ['\\rho', 0x03C1],
    ['\\Rho', 0x03A1],
    ['\\rhvec', 0x20D1],
    ['\\right', 0x2524],
    ['\\rightarrow', 0x2192],
    ['\\Rightarrow', 0x21D2],
    ['\\rightharpoondown', 0x21C1],
    ['\\rightharpoonup', 0x21C0],
    ['\\rmoust', 0x23B1],
    ['\\root', 0x24AD],
    ['\\scripta', 0x1D4B6],
    ['\\scriptA', 0x1D49C],
    ['\\scriptb', 0x1D4B7],
    ['\\scriptB', 0x212C],
    ['\\scriptc', 0x1D4B8],
    ['\\scriptC', 0x1D49E],
    ['\\scriptd', 0x1D4B9],
    ['\\scriptD', 0x1D49F],
    ['\\scripte', 0x212F],
    ['\\scriptE', 0x2130],
    ['\\scriptf', 0x1D4BB],
    ['\\scriptF', 0x2131],
    ['\\scriptg', 0x210A],
    ['\\scriptG', 0x1D4A2],
    ['\\scripth', 0x1D4BD],
    ['\\scriptH', 0x210B],
    ['\\scripti', 0x1D4BE],
    ['\\scriptI', 0x2110],
    ['\\scriptj', 0x1D4BF],
    ['\\scriptJ', 0x1D4A5],
    ['\\scriptk', 0x1D4C0],
    ['\\scriptK', 0x1D4A6],
    ['\\scriptl', 0x2113],
    ['\\scriptL', 0x2112],
    ['\\scriptm', 0x1D4C2],
    ['\\scriptM', 0x2133],
    ['\\scriptn', 0x1D4C3],
    ['\\scriptN', 0x1D4A9],
    ['\\scripto', 0x2134],
    ['\\scriptO', 0x1D4AA],
    ['\\scriptp', 0x1D4C5],
    ['\\scriptP', 0x1D4AB],
    ['\\scriptq', 0x1D4C6],
    ['\\scriptQ', 0x1D4AC],
    ['\\scriptr', 0x1D4C7],
    ['\\scriptR', 0x211B],
    ['\\scripts', 0x1D4C8],
    ['\\scriptS', 0x1D4AE],
    ['\\scriptt', 0x1D4C9],
    ['\\scriptT', 0x1D4AF],
    ['\\scriptu', 0x1D4CA],
    ['\\scriptU', 0x1D4B0],
    ['\\scriptv', 0x1D4CB],
    ['\\scriptV', 0x1D4B1],
    ['\\scriptw', 0x1D4CC],
    ['\\scriptW', 0x1D4B2],
    ['\\scriptx', 0x1D4CD],
    ['\\scriptX', 0x1D4B3],
    ['\\scripty', 0x1D4CE],
    ['\\scriptY', 0x1D4B4],
    ['\\scriptz', 0x1D4CF],
    ['\\scriptZ', 0x1D4B5],
    ['\\sdiv', 0x2044],
    ['\\sdivide', 0x2044],
    ['\\searrow', 0x2198],
    ['\\setminus', 0x2216],
    ['\\sigma', 0x03C3],
    ['\\Sigma', 0x03A3],
    ['\\sim', 0x223C],
    ['\\simeq', 0x2243],
    ['\\smash', 0x2B0D],
    ['\\smile', 0x2323],
    ['\\spadesuit', 0x2660],
    ['\\sqcap', 0x2293],
    ['\\sqcup', 0x2294],
    ['\\sqrt', 0x221A],
    ['\\sqsubseteq', 0x2291],
    ['\\sqsuperseteq', 0x2292],
    ['\\star', 0x22C6],
    ['\\subset', 0x2282],
    // ['/\\subseteq', 0x2288],
    ['\\subseteq', 0x2286],
    ['\\succ', 0x227B],
    ['\\succeq', 0x227D],
    ['\\sum', 0x2211],
    ['\\superset', 0x2283],
    ['\\superseteq', 0x2287],
    ['\\swarrow', 0x2199],
    ['\\tau', 0x03C4],
    ['\\Tau', 0x03A4],
    ['\\therefore', 0x2234],
    ['\\theta', 0x03B8],
    ['\\Theta', 0x0398],
    ['\\thicksp', 0x2005],
    ['\\thinsp', 0x2006],
    ['\\tilde', 0x0303],
    ['\\times', 0x00D7],
    ['\\to', 0x2192],
    ['\\top', 0x22A4],
    ['\\tvec', 0x20E1],
    ['\\ubar', 0x0332],
    ['\\Ubar', 0x0333],
    ['\\underbar', 0x2581],
    ['\\underbrace', 0x23DF],
    ['\\underbracket', 0x23B5],
    ['\\underline', 0x25B1],
    ['\\underparen', 0x23DD],
    ['\\uparrow', 0x2191],
    ['\\Uparrow', 0x21D1],
    ['\\updownarrow', 0x2195],
    ['\\Updownarrow', 0x21D5],
    ['\\uplus', 0x228E],
    ['\\upsilon', 0x03C5],
    ['\\Upsilon', 0x03A5],
    ['\\varepsilon', 0x03B5],
    ['\\varphi', 0x03C6],
    ['\\varpi', 0x03D6],
    ['\\varrho', 0x03F1],
    ['\\varsigma', 0x03C2],
    ['\\vartheta', 0x03D1],
    ['\\vbar', 0x2502],
    ['\\vdash', 0x22A2],
    ['\\vdots', 0x22EE],
    ['\\vec', 0x20D7],
    ['\\vee', 0x2228],
    ['\\vert', 0x007C],
    ['\\Vert', 0x2016],
    ['\\Vmatrix', 0x24A9],
    ['\\vphantom', 0x21F3],
    ['\\vthicksp', 0x2004],
    ['\\wedge', 0x2227],
    ['\\wp', 0x2118],
    ['\\wr', 0x2240],
    ['\\xi', 0x03BE],
    ['\\Xi', 0x039E],
    ['\\zeta', 0x03B6],
    ['\\Zeta', 0x0396],
    ['\\zwnj', 0x200C],
    ['\\zwsp', 0x200B],
    ['~=', 0x2245],
    ['-+', 0x2213],
    ['+-', 0x00B1],
    ['<<', 0x226A],
    ['<=', 0x2264],
    ['->', 0x2192],
    ['>=', 0x2265],
    ['>>', 0x226B]
];
var g_AutoCorrectMathSymbols = JSON.parse(JSON.stringify(g_DefaultAutoCorrectMathSymbolsList));
var g_AutoCorrectMathFuncs = JSON.parse(JSON.stringify(g_DefaultAutoCorrectMathFuncs));
var g_AutoCorrectMathsList = {
    DefaultAutoCorrectMathSymbolsList : g_DefaultAutoCorrectMathSymbolsList,
    AutoCorrectMathSymbols : g_AutoCorrectMathSymbols,
    DefaultAutoCorrectMathFuncs : g_DefaultAutoCorrectMathFuncs,
    AutoCorrectMathFuncs : g_AutoCorrectMathFuncs
};
//символы для mathfunc (интеграл, сумма...)
var q_aMathAutoCorrectControlAggregationCodes =
{
    0x2211 : 1, 0x220F : 1, 0x2210 : 1, 0x22C0 : 1, 0x222E : 1,
    0x22C1 : 1, 0x22C2 : 1, 0x22C3 : 1, 0x2A06 : 1, 0x2A04 : 1,
    0x2A00 : 1, 0x2A01 : 1, 0x2A02 : 1, 0x222B : 1, 0x222C : 1,
    0x222D : 1, 0x2A0C : 1, 0x222F : 1, 0x2230 : 1, 0x2232 : 1
};
//символы accent
var q_aMathAutoCorrectAccentCharCodes =
{
    0x0305 : 1, 0x033F : 1, 0x0332 : 1, 0x0333 : 1, 0x0301 : 1,
    0x0300 : 1, 0x20D7 : 1, 0x0306 : 1, 0x0302 : 1, 0x20E1 : 1,
    0x20D1 : 1, 0x030C : 1, 0x0303 : 1, 0x20D6 : 1, 0x20D0 : 1,
    0x0307 : 1, 0x0308 : 1, 0x20DB : 1, 0x20DC : 1, 0x2032 : 1,
    0x2033 : 1, 0x2034 : 1, 0x2057 : 1
};
//left brackets
var g_MathLeftBracketAutoCorrectCharCodes =
{
    0x28 : 1, 0x5B : 1, 0x7B : 1, /*0x7C : 1, 0x2016 : 1,*/ 0x27E8 : 1,
    0x2329 : 1, 0x27E6 : 1, 0x27EA : 1, 0x2308 : 1, 0x230A : 1,
    0x3016 : 1, 0x251C : 1
};
//righ brackets
var g_MathRightBracketAutoCorrectCharCodes =
{
    0x29 : 1, 0x5D : 1, 0x7D : 1,/* 0x7C : 1, 0x2016 : 1,*/ 0x27E9 : 1,
    0x232A : 1, 0x27E7 : 1, 0x27EB : 1, 0x2309 : 1, 0x230B : 1,   
    0x3017 : 1, 0x2524 : 1
};
//знаки (минус, сумма...)
var g_aMathAutoCorrectFracCharCodes =
{
    0x20 : 1, 0x21 : 1, 0x22 : 1, 0x23 : 1,	0x24 : 1, 0x25 : 1, 0x26 : 1,
    0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2A : 1, 0x2B : 1, 0x2C : 1, 0x2D : 1,
    0x2E : 1, 0x2F : 1, 0x3A : 1, 0x3B : 1, 0x3C : 1, 0x3D : 1, 0x3E : 1,
    0x3F : 1, 0x40 : 1, 0x5B : 1, /*0x5C : 1,*/ 0x5D : 1, 0x5E : 1, 0x5F : 1,
    0x60 : 1, 0x7B : 1, 0x7C : 1, 0x7D : 1, 0x7E : 1, /*0x2592 : 1,*/ 0xD7 : 1
};
//символы для определения необходимости автозамены
var g_aMathAutoCorrectTriggerCharCodes =
{
    0x20 : 1, 0x21 : 1, 0x22 : 1, 0x23 : 1, 0x24 : 1, 0x25 : 1, 0x26 : 1,
    0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2A : 1, 0x2B : 1, 0x2C : 1, 0x2D : 1,
    0x2E : 1, 0x2F : 1, 0x3A : 1, 0x3B : 1, 0x3C : 1, 0x3D : 1, 0x3E : 1,
    0x3F : 1, 0x40 : 1, 0x5B : 1, 0x5D : 1, 0x5C : 1, 0x5E : 1, 0x5F : 1,
    0x60 : 1, 0x7B : 1, 0x7D : 1, 0x7C : 1, 0x7E : 1 /*,0x2592 : 1*/
};
//символы при которых производится автозамена текстовой функции (lim, max, sin...)
var g_aMathAutoCorrectTextFunc = {
    0x20 : 1, 0x28 : 1, 0x29 : 1, 0x5B : 1, 0x5E : 1, 0x5F : 1, 0x7B : 1
};
//символы, при которых не производится автозамена дроби без скобок
var g_aMathAutoCorrectNotDoFraction = {
    0x21 : 1, 0x22 : 1, 0x27 : 1, 0x24 : 1, 0x28 : 1, 0x29 : 1, 0x5B : 1,
    0x5D : 1, 0x7B : 1, 0x7D : 1, 0x5C : 1, 0x5E : 1, 0x5F : 1, 0x7C : 1
};
//символы, при которых не производится автозамена степени без скобок
var g_aMathAutoCorrectDoNotDegree = {
    0x21 : 1, 0x22 : 1, 0x27 : 1, 0x24 : 1, 0x28 : 1, 0x29 : 1, 0x5B : 1, 0x5D : 1,
    0x7B : 1, 0x7D : 1, 0x2C : 1, 0x2E : 1, 0x5C : 1, 0x5E : 1, 0x5F : 1, 0x7C : 1
};
//символы, при которых не производится автозамена nary
var g_aMathAutoCorrectNotDoCNary = {
   0x22 : 1, 0x24 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2E : 1, 0x5B : 1,
   0x5C : 1, 0x5D : 1, 0x5E : 1, 0x5F : 1, 0x7B : 1, 0x7D : 1, 0x7C : 1 
};
//символы, при которых не производится автозамена скобок
var g_aMathAutoCorrectDoNotDelimiter = {
    0x22 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2F : 1,
    0x5B : 1, 0x5C : 1, 0x5E : 1, 0x5F : 1, 0x7B : 1
 };
// символы, при которых не производится автозамена корня
var g_aMathAutoCorrectDoNotRadical = {
    0x22 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x5C : 1,
    0x5E : 1, 0x5F : 1, 0x5D : 1, 0x7D : 1
};
// символы, при которых не производится автозамена Box
var g_aMathAutoCorrectDoNotBox = {
    0x21 : 1, 0x22 : 1, 0x24 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1,
    0x5B : 1, 0x5D : 1, 0x7B : 1, 0x7D : 1, 0x5C : 1, 0x5E : 1,
    0x5F : 1, 0x7C : 1,
};
// символы, при которых не производится автозамена Group character
var g_aMathAutoCorrectDoNotGroupChar = {
    0x21 : 1, 0x22 : 1, 0x24 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x2C : 1,
    0x2E : 1, 0x5B : 1, 0x5D : 1, 0x7B : 1, 0x7D : 1, 0x5C : 1, 0x7C : 1
};
// символы, при которых не производится автозамена EqArray и Matrix
var g_aMathAutoCorrectDoNotMatrix = {
    0x22 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x5C : 1,
    0x5D : 1, 0x5E : 1, 0x5F : 1, 0x7D : 1
};
// символы, при которых не производится автозамена Accent не вплотную
var g_aMathAutoCorrectDoNotAccentNotClose = {
    0x22 : 1, 0x27 : 1, 0x28 : 1, 0x29 : 1, 0x5C : 1,
    0x5D : 1, 0x5E : 1, 0x5F : 1, 0x7D : 1
};
// символы, при которых не производится автозамена Accent вплотную
var g_aMathAutoCorrectDoNotAccentClose = {
    0x22 : 1, 0x29 : 1, 0x5D : 1, 0x7D : 1
};
// символы при которых не производится автозамена Above
var g_aMathAutoCorrectDoNotAbove = {
    0x21 : 1, 0x22 : 1, 0x24 : 1, 0x27 : 1, 0x28 : 1,
    0x29 : 1, 0x2C : 1, 0x2E : 1, 0x5B : 1, 0x5C : 1,
    0x5D : 1, 0x5E : 1, 0x5F : 1, 0x7B : 1, 0x7C : 1, 0x7D : 1
};
// символы при которых не производится автозамена mathfunc 0x2061
var g_aMathAutoCorrectDoNotMathFunc = {
    0x21 : 1, 0x22 : 1, 0x24 : 1, 0x27 : 1,
    0x28 : 1, 0x29 : 1, 0x2C : 1, 0x2E : 1,
    0x5B : 1, 0x5C : 1, 0x5D : 1, 0x5E : 1,
    0x5F : 1, 0x7B : 1, 0x7C : 1, 0x7D : 1
};
// символы корней
var g_aMathAutoCorrectRadicalCharCode = {
    0x221A : 1, 0x221B : 1, 0x221C : 1
};
//символы, при которых надо пропускать обработку скобок
//во вложенном контенте
var g_aMathAutoCorrectSkipBrackets = {
    0x221A : 1, 0x25A1 : 1, 0x25AD : 1, 0x23DE : 1, 0x23DC : 1,
    0x23E0 : 1, 0x23DF : 1, 0x23DD : 1, 0x00AF : 1, 0x2581 : 1,
    0x24B8 : 1, 0x2588 : 1, 0x25A0 : 1, 0x24A8 : 1, 0x24A9 : 1
};
//символы MATH_GROUP_CHARACTER
var g_aMathAutoCorrectGroupChar = {
    0x23DE : 1, 0x23DC : 1, 0x23E0 : 1, 0x23DF : 1, 0x23DD : 1
};
//symbols matrix and eqarray
var g_aMathAutoCorrectEqArrayMatrix = {
    0x24B8 : 1, 0x2588 : 1, 0x25A0 : 1, 0x24A8 : 1, 0x24A9 : 1
};
// latin alphabet ascii code
var g_aMathAutoCorrectLatinAlph = {
    0x41 : 1, 0x42 : 1, 0x43 : 1, 0x44 : 1, 0x45 : 1, 0x46 : 1,
    0x47 : 1, 0x48 : 1, 0x49 : 1, 0x4A : 1, 0x4B : 1, 0x4C : 1,
    0x4D : 1, 0x4E : 1, 0x4F : 1, 0x50 : 1, 0x51 : 1, 0x52 : 1,
    0x53 : 1, 0x54 : 1, 0x55 : 1, 0x56 : 1, 0x57 : 1, 0x58 : 1,
    0x59 : 1, 0x5A : 1, 0x60 : 1, 0x61 : 1, 0x62 : 1, 0x63 : 1,
    0x64 : 1, 0x65 : 1, 0x66 : 1, 0x67 : 1, 0x68 : 1, 0x69 : 1,
    0x6A : 1, 0x6B : 1, 0x6C : 1, 0x6D : 1, 0x6E : 1, 0x6F : 1,
    0x70 : 1, 0x71 : 1, 0x72 : 1, 0x73 : 1, 0x74 : 1, 0x75 : 1,
    0x76 : 1, 0x77 : 1, 0x78 : 1, 0x79 : 1, 0x7A : 1
};
// special symbols for autocorrect text
var g_aMathAutoCorrectSpecSymb = [
    '!!', '...', '::', ':=', '/<', '/>', '/=',
    '~=', '-+', '+-', '<<', '<=', '->', '>=', '>>'
];

//--------------------------------------------------------export----------------------------------------------------
window['AscCommonWord'] = window['AscCommonWord'] || {};
window['AscCommonWord'].CMathContent = CMathContent;
window['AscCommonWord'].g_AutoCorrectMathsList = g_AutoCorrectMathsList;

