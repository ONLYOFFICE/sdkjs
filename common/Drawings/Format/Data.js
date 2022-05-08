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
(
  /**
   * @param {Window} window
   * @param {undefined} undefined
   */
  function (window, undefined) {
/*
The current module is designed to implement SmartArt support.
At the moment, there is partial support for the format, its saving and editing.
At the moment, there is support for the drawing.xml file - this should be abandoned, smart arts are built with information from the data.xml file, drawn by bypassing the layout.xml file.
Need to support:
1. The connection must be built data -> drawing, at the moment the opposite is happening.

2. Rendering should take place according to the layout.xml file.

3. Synchronous filling of a paragraph in data and drawing, at the moment this is not done correctly - available paragraphs are divided evenly and filled in contentpoints.
CShape.prototype.copyTextInfoFromShapeToPoint = function (paddings) {
Because of this, the display is sometimes not correct.

4. Support placeholders for individual paragraphs. At the moment, there are two contents that replace each other when in focus and out of focus.

5. Support changing the smartart tree to add new nodes.
*/
    // imports

    var InitClass = AscFormat.InitClass;
    var InitClassWithoutType = AscFormat.InitClassWithoutType;
    var CBaseFormatObject = AscFormat.CBaseFormatObject;
    var oHistory = AscCommon.History;
    var CChangeBool = AscDFH.CChangesDrawingsBool;
    var CChangeLong = AscDFH.CChangesDrawingsLong;
    var CChangeDouble = AscDFH.CChangesDrawingsDouble;
    var CChangeString = AscDFH.CChangesDrawingsString;
    var CChangeObjectNoId = AscDFH.CChangesDrawingsObjectNoId;
    var CChangeObject = AscDFH.CChangesDrawingsObject;
    var CChangeContent = AscDFH.CChangesDrawingsContent;
    var CChangeDouble2 = AscDFH.CChangesDrawingsDouble2;
    var CChangesContentNoId = AscDFH.CChangesDrawingsContentNoId;

    var drawingsChangesMap = AscDFH.drawingsChangesMap;
    var drawingContentChanges = AscDFH.drawingContentChanges;
    var changesFactory = AscDFH.changesFactory;
    var drawingConstructorsMap = window['AscDFH'].drawingsConstructorsMap;
    var CUniColor = AscFormat.CUniColor;
    var SchemeClr = AscFormat.CSchemeColor;
    var ColorMod = AscFormat.CColorMod;
    var ColorModLst = AscFormat.CColorModifiers;
    var StyleRef = AscFormat.StyleRef;
    var RGBClr = AscFormat.CRGBColor;
    var ShapeStyle = AscFormat.CShapeStyle;
    var FontRef = AscFormat.FontRef;
    var CGraphicObjectBase = AscFormat.CGraphicObjectBase;
    var CGroupShape = AscFormat.CGroupShape;

    // consts
    var Point_type_asst = 1;
    var Point_type_doc = 2;
    var Point_type_node = 0;
    var Point_type_parTrans = 4;
    var Point_type_pres = 3;
    var Point_type_sibTrans = 5;

    var Cxn_type_parOf = 0;
    var Cxn_type_presOf = 1;
    var Cxn_type_presParOf = 2;
    var Cxn_type_unknownRelationShip = 3;

    var LayoutNode_type_b = 0;
    var LayoutNode_type_t = 1;

    var Alg_type_composite = 0;
    var Alg_type_conn = 1;
    var Alg_type_cycle = 2;
    var Alg_type_hierChild = 3;
    var Alg_type_hierRoot = 4;
    var Alg_type_lin = 6;
    var Alg_type_pyra = 5;
    var Alg_type_snake = 9;
    var Alg_type_sp = 7;
    var Alg_type_tx = 8;

    var Param_type_alignTx = 0;
    var Param_type_ar = 1;
    var Param_type_autoTxRot = 2;
    var Param_type_begPts = 3;
    var Param_type_begSty = 4;
    var Param_type_bendPt = 5;
    var Param_type_bkpt = 6;
    var Param_type_bkPtFixedVal = 7;
    var Param_type_chAlign = 8;
    var Param_type_chDir = 9;
    var Param_type_connRout = 10;
    var Param_type_contDir = 11;
    var Param_type_ctrShpMap = 12;
    var Param_type_dim = 13;
    var Param_type_dstNode = 14;
    var Param_type_endPts = 15;
    var Param_type_endSty = 16;
    var Param_type_fallback = 17;
    var Param_type_flowDir = 18;
    var Param_type_grDir = 19;
    var Param_type_hierAlign = 20;
    var Param_type_horzAlign = 21;
    var Param_type_linDir = 22;
    var Param_type_lnSpAfChP = 23;
    var Param_type_lnSpAfParP = 24;
    var Param_type_lnSpCh = 25;
    var Param_type_lnSpPar = 26;
    var Param_type_nodeHorzAlign = 27;
    var Param_type_nodeVertAlign = 28;
    var Param_type_off = 29;
    var Param_type_parTxLTRAlign = 30;
    var Param_type_parTxRTLAlign = 31;
    var Param_type_pyraAcctBkgdNode = 32;
    var Param_type_pyraAcctPos = 33;
    var Param_type_pyraAcctTxMar = 34;
    var Param_type_pyraAcctTxNode = 35;
    var Param_type_pyraLvlNode = 36;
    var Param_type_rotPath = 37;
    var Param_type_rtShortDist = 38;
    var Param_type_secChAlign = 39;
    var Param_type_secLinDir = 40;
    var Param_type_shpTxLTRAlignCh = 41;
    var Param_type_shpTxRTLAlignCh = 42;
    var Param_type_spanAng = 43;
    var Param_type_srcNode = 44;
    var Param_type_stAng = 45;
    var Param_type_stBulletLvl = 46;
    var Param_type_stElem = 47;
    var Param_type_txAnchorHorz = 48;
    var Param_type_txAnchorHorzCh = 49;
    var Param_type_txAnchorVert = 50;
    var Param_type_txAnchorVertCh = 51;
    var Param_type_txBlDir = 52;
    var Param_type_txDir = 53;
    var Param_type_vertAlign = 54;

    var AxisType_value_ancst = 6;
    var AxisType_value_ancstOrSelf = 7;
    var AxisType_value_ch = 2;
    var AxisType_value_des = 3;
    var AxisType_value_desOrSelf = 4;
    var AxisType_value_follow = 10;
    var AxisType_value_followSib = 8;
    var AxisType_value_none = 0;
    var AxisType_value_par = 5;
    var AxisType_value_preced = 11;
    var AxisType_value_precedSib = 9;
    var AxisType_value_root = 12;
    var AxisType_value_self = 1;

    var ElementType_value_all = 0;
    var ElementType_value_asst = 5;
    var ElementType_value_doc = 1;
    var ElementType_value_node = 2;
    var ElementType_value_nonAsst = 6;
    var ElementType_value_nonNorm = 4;
    var ElementType_value_norm = 3;
    var ElementType_value_parTrans = 7;
    var ElementType_value_pres = 8;
    var ElementType_value_sibTrans = 9;

    var If_op_equ = 0;
    var If_op_gt = 2;
    var If_op_gte = 4;
    var If_op_lt = 3;
    var If_op_lte = 5;
    var If_op_neq = 1;

    var If_func_cnt = 0;
    var If_func_depth = 6;
    var If_func_maxDepth = 7;
    var If_func_pos = 1;
    var If_func_posEven = 3;
    var If_func_posOdd = 4;
    var If_func_revPos = 2;
    var If_func_var = 5;

    var If_arg_animLvl = 0;
    var If_arg_animOne = 1;
    var If_arg_bulEnabled = 2;
    var If_arg_chMax = 3;
    var If_arg_chPref = 4;
    var If_arg_dir = 5;
    var If_arg_hierBranch = 6;
    var If_arg_none = 7;
    var If_arg_orgChart = 8;
    var If_arg_resizeHandles = 9;

    var Constr_for_ch = 1;
    var Constr_for_des = 2;
    var Constr_for_self = 0;

    var Constr_op_equ = 1;
    var Constr_op_gte = 2;
    var Constr_op_lte = 3;
    var Constr_op_none = 0;

    var Constr_type_alignOff = 1;
    var Constr_type_b = 5;
    var Constr_type_begMarg = 2;
    var Constr_type_begPad = 4;
    var Constr_type_bendDist = 3;
    var Constr_type_bMarg = 6;
    var Constr_type_bOff = 7;
    var Constr_type_connDist = 12;
    var Constr_type_ctrX = 8;
    var Constr_type_ctrXOff = 9;
    var Constr_type_ctrY = 10;
    var Constr_type_ctrYOff = 11;
    var Constr_type_diam = 13;
    var Constr_type_endMarg = 14;
    var Constr_type_endPad = 15;
    var Constr_type_h = 16;
    var Constr_type_hArH = 17;
    var Constr_type_hOff = 63; // TODO: add to constr type in x2t
    var Constr_type_l = 18;
    var Constr_type_lMarg = 19;
    var Constr_type_lOff = 20;
    var Constr_type_none = 0;
    var Constr_type_primFontSz = 24;
    var Constr_type_pyraAcctRatio = 25;
    var Constr_type_r = 21;
    var Constr_type_rMarg = 22;
    var Constr_type_rOff = 23;
    var Constr_type_secFontSz = 26;
    var Constr_type_secSibSp = 28;
    var Constr_type_sibSp = 27;
    var Constr_type_sp = 29;
    var Constr_type_stemThick = 30;
    var Constr_type_t = 31;
    var Constr_type_tMarg = 32;
    var Constr_type_tOff = 33;
    var Constr_type_userA = 34;
    var Constr_type_userB = 35;
    var Constr_type_userC = 36;
    var Constr_type_userD = 37;
    var Constr_type_userE = 38;
    var Constr_type_userF = 39;
    var Constr_type_userG = 40;
    var Constr_type_userH = 41;
    var Constr_type_userI = 42;
    var Constr_type_userJ = 43;
    var Constr_type_userK = 44;
    var Constr_type_userL = 45;
    var Constr_type_userM = 46;
    var Constr_type_userN = 47;
    var Constr_type_userO = 48;
    var Constr_type_userP = 49;
    var Constr_type_userQ = 50;
    var Constr_type_userR = 51;
    var Constr_type_userS = 52;
    var Constr_type_userT = 53;
    var Constr_type_userU = 54;
    var Constr_type_userV = 55;
    var Constr_type_userW = 56;
    var Constr_type_userX = 57;
    var Constr_type_userY = 58;
    var Constr_type_userZ = 59;
    var Constr_type_w = 60;
    var Constr_type_wArH = 61;
    var Constr_type_wOff = 62;

    var kForInsFitFontSize = 71.12 / 360;

    var LayoutShapeType_outputShapeType_conn = 0;
    var LayoutShapeType_outputShapeType_none = 1;
    var LayoutShapeType_shapeType_accentBorderCallout1 = 2;
    var LayoutShapeType_shapeType_accentBorderCallout2 = 3;
    var LayoutShapeType_shapeType_accentBorderCallout3 = 4;
    var LayoutShapeType_shapeType_accentCallout1 = 5;
    var LayoutShapeType_shapeType_accentCallout2 = 6;
    var LayoutShapeType_shapeType_accentCallout3 = 7;
    var LayoutShapeType_shapeType_actionButtonBackPrevious = 8;
    var LayoutShapeType_shapeType_actionButtonBeginning = 9;
    var LayoutShapeType_shapeType_actionButtonBlank = 10;
    var LayoutShapeType_shapeType_actionButtonDocument = 11;
    var LayoutShapeType_shapeType_actionButtonEnd = 12;
    var LayoutShapeType_shapeType_actionButtonForwardNext = 13;
    var LayoutShapeType_shapeType_actionButtonHelp = 14;
    var LayoutShapeType_shapeType_actionButtonHome = 15;
    var LayoutShapeType_shapeType_actionButtonInformation = 16;
    var LayoutShapeType_shapeType_actionButtonMovie = 17;
    var LayoutShapeType_shapeType_actionButtonReturn = 18;
    var LayoutShapeType_shapeType_actionButtonSound = 19;
    var LayoutShapeType_shapeType_arc = 20;
    var LayoutShapeType_shapeType_bentArrow = 21;
    var LayoutShapeType_shapeType_bentConnector2 = 22;
    var LayoutShapeType_shapeType_bentConnector3 = 23;
    var LayoutShapeType_shapeType_bentConnector4 = 24;
    var LayoutShapeType_shapeType_bentConnector5 = 25;
    var LayoutShapeType_shapeType_bentUpArrow = 26;
    var LayoutShapeType_shapeType_bevel = 27;
    var LayoutShapeType_shapeType_blockArc = 28;
    var LayoutShapeType_shapeType_borderCallout1 = 29;
    var LayoutShapeType_shapeType_borderCallout2 = 30;
    var LayoutShapeType_shapeType_borderCallout3 = 31;
    var LayoutShapeType_shapeType_bracePair = 32;
    var LayoutShapeType_shapeType_bracketPair = 33;
    var LayoutShapeType_shapeType_callout1 = 34;
    var LayoutShapeType_shapeType_callout2 = 35;
    var LayoutShapeType_shapeType_callout3 = 36;
    var LayoutShapeType_shapeType_can = 37;
    var LayoutShapeType_shapeType_chartPlus = 38;
    var LayoutShapeType_shapeType_chartStar = 39;
    var LayoutShapeType_shapeType_chartX = 40;
    var LayoutShapeType_shapeType_chevron = 41;
    var LayoutShapeType_shapeType_chord = 42;
    var LayoutShapeType_shapeType_circularArrow = 43;
    var LayoutShapeType_shapeType_cloud = 44;
    var LayoutShapeType_shapeType_cloudCallout = 45;
    var LayoutShapeType_shapeType_corner = 46;
    var LayoutShapeType_shapeType_cornerTabs = 47;
    var LayoutShapeType_shapeType_cube = 48;
    var LayoutShapeType_shapeType_curvedConnector2 = 49;
    var LayoutShapeType_shapeType_curvedConnector3 = 50;
    var LayoutShapeType_shapeType_curvedConnector4 = 51;
    var LayoutShapeType_shapeType_curvedConnector5 = 52;
    var LayoutShapeType_shapeType_curvedDownArrow = 53;
    var LayoutShapeType_shapeType_curvedLeftArrow = 54;
    var LayoutShapeType_shapeType_curvedRightArrow = 55;
    var LayoutShapeType_shapeType_curvedUpArrow = 56;
    var LayoutShapeType_shapeType_decagon = 57;
    var LayoutShapeType_shapeType_diagStripe = 58;
    var LayoutShapeType_shapeType_diamond = 59;
    var LayoutShapeType_shapeType_dodecagon = 60;
    var LayoutShapeType_shapeType_donut = 61;
    var LayoutShapeType_shapeType_doubleWave = 62;
    var LayoutShapeType_shapeType_downArrow = 63;
    var LayoutShapeType_shapeType_downArrowCallout = 64;
    var LayoutShapeType_shapeType_ellipse = 65;
    var LayoutShapeType_shapeType_ellipseRibbon = 66;
    var LayoutShapeType_shapeType_ellipseRibbon2 = 67;
    var LayoutShapeType_shapeType_flowChartAlternateProcess = 68;
    var LayoutShapeType_shapeType_flowChartCollate = 69;
    var LayoutShapeType_shapeType_flowChartConnector = 70;
    var LayoutShapeType_shapeType_flowChartDecision = 71;
    var LayoutShapeType_shapeType_flowChartDelay = 72;
    var LayoutShapeType_shapeType_flowChartDisplay = 73;
    var LayoutShapeType_shapeType_flowChartDocument = 74;
    var LayoutShapeType_shapeType_flowChartExtract = 75;
    var LayoutShapeType_shapeType_flowChartInputOutput = 76;
    var LayoutShapeType_shapeType_flowChartInternalStorage = 77;
    var LayoutShapeType_shapeType_flowChartMagneticDisk = 78;
    var LayoutShapeType_shapeType_flowChartMagneticDrum = 79;
    var LayoutShapeType_shapeType_flowChartMagneticTape = 80;
    var LayoutShapeType_shapeType_flowChartManualInput = 81;
    var LayoutShapeType_shapeType_flowChartManualOperation = 82;
    var LayoutShapeType_shapeType_flowChartMerge = 83;
    var LayoutShapeType_shapeType_flowChartMultidocument = 84;
    var LayoutShapeType_shapeType_flowChartOfflineStorage = 85;
    var LayoutShapeType_shapeType_flowChartOffpageConnector = 86;
    var LayoutShapeType_shapeType_flowChartOnlineStorage = 87;
    var LayoutShapeType_shapeType_flowChartOr = 88;
    var LayoutShapeType_shapeType_flowChartPredefinedProcess = 89;
    var LayoutShapeType_shapeType_flowChartPreparation = 90;
    var LayoutShapeType_shapeType_flowChartProcess = 91;
    var LayoutShapeType_shapeType_flowChartPunchedCard = 92;
    var LayoutShapeType_shapeType_flowChartPunchedTape = 93;
    var LayoutShapeType_shapeType_flowChartSort = 94;
    var LayoutShapeType_shapeType_flowChartSummingJunction = 95;
    var LayoutShapeType_shapeType_flowChartTerminator = 96;
    var LayoutShapeType_shapeType_foldedCorner = 97;
    var LayoutShapeType_shapeType_frame = 98;
    var LayoutShapeType_shapeType_funnel = 99;
    var LayoutShapeType_shapeType_gear6 = 100;
    var LayoutShapeType_shapeType_gear9 = 101;
    var LayoutShapeType_shapeType_halfFrame = 102;
    var LayoutShapeType_shapeType_heart = 103;
    var LayoutShapeType_shapeType_heptagon = 104;
    var LayoutShapeType_shapeType_hexagon = 105;
    var LayoutShapeType_shapeType_homePlate = 106;
    var LayoutShapeType_shapeType_horizontalScroll = 107;
    var LayoutShapeType_shapeType_irregularSeal1 = 108;
    var LayoutShapeType_shapeType_irregularSeal2 = 109;
    var LayoutShapeType_shapeType_leftArrow = 110;
    var LayoutShapeType_shapeType_leftArrowCallout = 111;
    var LayoutShapeType_shapeType_leftBrace = 112;
    var LayoutShapeType_shapeType_leftBracket = 113;
    var LayoutShapeType_shapeType_leftCircularArrow = 114;
    var LayoutShapeType_shapeType_leftRightArrow = 115;
    var LayoutShapeType_shapeType_leftRightArrowCallout = 116;
    var LayoutShapeType_shapeType_leftRightCircularArrow = 117;
    var LayoutShapeType_shapeType_leftRightRibbon = 118;
    var LayoutShapeType_shapeType_leftRightUpArrow = 119;
    var LayoutShapeType_shapeType_leftUpArrow = 120;
    var LayoutShapeType_shapeType_lightningBolt = 121;
    var LayoutShapeType_shapeType_line = 122;
    var LayoutShapeType_shapeType_lineInv = 123;
    var LayoutShapeType_shapeType_mathDivide = 124;
    var LayoutShapeType_shapeType_mathEqual = 125;
    var LayoutShapeType_shapeType_mathMinus = 126;
    var LayoutShapeType_shapeType_mathMultiply = 127;
    var LayoutShapeType_shapeType_mathNotEqual = 128;
    var LayoutShapeType_shapeType_mathPlus = 129;
    var LayoutShapeType_shapeType_moon = 130;
    var LayoutShapeType_shapeType_nonIsoscelesTrapezoid = 131;
    var LayoutShapeType_shapeType_noSmoking = 132;
    var LayoutShapeType_shapeType_notchedRightArrow = 133;
    var LayoutShapeType_shapeType_octagon = 134;
    var LayoutShapeType_shapeType_parallelogram = 135;
    var LayoutShapeType_shapeType_pentagon = 136;
    var LayoutShapeType_shapeType_pie = 137;
    var LayoutShapeType_shapeType_pieWedge = 138;
    var LayoutShapeType_shapeType_plaque = 139;
    var LayoutShapeType_shapeType_plaqueTabs = 140;
    var LayoutShapeType_shapeType_plus = 141;
    var LayoutShapeType_shapeType_quadArrow = 142;
    var LayoutShapeType_shapeType_quadArrowCallout = 143;
    var LayoutShapeType_shapeType_rect = 144;
    var LayoutShapeType_shapeType_ribbon = 145;
    var LayoutShapeType_shapeType_ribbon2 = 146;
    var LayoutShapeType_shapeType_rightArrow = 147;
    var LayoutShapeType_shapeType_rightArrowCallout = 148;
    var LayoutShapeType_shapeType_rightBrace = 149;
    var LayoutShapeType_shapeType_rightBracket = 150;
    var LayoutShapeType_shapeType_round1Rect = 151;
    var LayoutShapeType_shapeType_round2DiagRect = 152;
    var LayoutShapeType_shapeType_round2SameRect = 153;
    var LayoutShapeType_shapeType_roundRect = 154;
    var LayoutShapeType_shapeType_rtTriangle = 155;
    var LayoutShapeType_shapeType_smileyFace = 156;
    var LayoutShapeType_shapeType_snip1Rect = 157;
    var LayoutShapeType_shapeType_snip2DiagRect = 158;
    var LayoutShapeType_shapeType_snip2SameRect = 159;
    var LayoutShapeType_shapeType_snipRoundRect = 160;
    var LayoutShapeType_shapeType_squareTabs = 161;
    var LayoutShapeType_shapeType_star10 = 162;
    var LayoutShapeType_shapeType_star12 = 163;
    var LayoutShapeType_shapeType_star16 = 164;
    var LayoutShapeType_shapeType_star24 = 165;
    var LayoutShapeType_shapeType_star32 = 166;
    var LayoutShapeType_shapeType_star4 = 167;
    var LayoutShapeType_shapeType_star5 = 168;
    var LayoutShapeType_shapeType_star6 = 169;
    var LayoutShapeType_shapeType_star7 = 170;
    var LayoutShapeType_shapeType_star8 = 171;
    var LayoutShapeType_shapeType_straightConnector1 = 172;
    var LayoutShapeType_shapeType_stripedRightArrow = 173;
    var LayoutShapeType_shapeType_sun = 174;
    var LayoutShapeType_shapeType_swooshArrow = 175;
    var LayoutShapeType_shapeType_teardrop = 176;
    var LayoutShapeType_shapeType_trapezoid = 177;
    var LayoutShapeType_shapeType_triangle = 178;
    var LayoutShapeType_shapeType_upArrow = 179;
    var LayoutShapeType_shapeType_upArrowCallout = 180;
    var LayoutShapeType_shapeType_upDownArrow = 181;
    var LayoutShapeType_shapeType_upDownArrowCallout = 182;
    var LayoutShapeType_shapeType_uturnArrow = 183;
    var LayoutShapeType_shapeType_verticalScroll = 184;
    var LayoutShapeType_shapeType_wave = 185;
    var LayoutShapeType_shapeType_wedgeEllipseCallout = 186;
    var LayoutShapeType_shapeType_wedgeRectCallout = 187;
    var LayoutShapeType_shapeType_wedgeRoundRectCallout = 188;


    var AnimLvl_val_ctr = 1;
    var AnimLvl_val_lvl = 2;
    var AnimLvl_val_none = 0;

    var AnimOne_val_branch = 1;
    var AnimOne_val_none = 0;
    var AnimOne_val_one = 2;

    var DiagramDirection_val_norm = 0;
    var DiagramDirection_val_rev = 1;

    var HierBranch_val_hang = 0;
    var HierBranch_val_init = 1;
    var HierBranch_val_l = 2;
    var HierBranch_val_r = 3;
    var HierBranch_val_std = 4;

    var ResizeHandles_val_exact = 0;
    var ResizeHandles_val_rel = 1;

    var ClrLst_hueDir_ccw = 0;
    var ClrLst_hueDir_cw = 1;
    var ClrLst_meth_cycle = 0;
    var ClrLst_meth_repeat = 1;
    var ClrLst_meth_span = 2;

    var Camera_prst_isometricBottomDown = 0;
    var Camera_prst_isometricBottomUp = 1;
    var Camera_prst_isometricLeftDown = 2;
    var Camera_prst_isometricLeftUp = 3;
    var Camera_prst_isometricOffAxis1Left = 4;
    var Camera_prst_isometricOffAxis1Right = 5;
    var Camera_prst_isometricOffAxis1Top = 6;
    var Camera_prst_isometricOffAxis2Left = 7;
    var Camera_prst_isometricOffAxis2Right = 8;
    var Camera_prst_isometricOffAxis2Top = 9;
    var Camera_prst_isometricOffAxis3Bottom = 10;
    var Camera_prst_isometricOffAxis3Left = 11;
    var Camera_prst_isometricOffAxis3Right = 12;
    var Camera_prst_isometricOffAxis4Bottom = 13;
    var Camera_prst_isometricOffAxis4Left = 14;
    var Camera_prst_isometricOffAxis4Right = 15;
    var Camera_prst_isometricRightDown = 16;
    var Camera_prst_isometricRightUp = 17;
    var Camera_prst_isometricTopDown = 18;
    var Camera_prst_isometricTopUp = 19;
    var Camera_prst_legacyObliqueBottom = 20;
    var Camera_prst_legacyObliqueBottomLeft = 21;
    var Camera_prst_legacyObliqueBottomRight = 22;
    var Camera_prst_legacyObliqueFront = 23;
    var Camera_prst_legacyObliqueLeft = 24;
    var Camera_prst_legacyObliqueRight = 25;
    var Camera_prst_legacyObliqueTop = 26;
    var Camera_prst_legacyObliqueTopLeft = 27;
    var Camera_prst_legacyObliqueTopRight = 28;
    var Camera_prst_legacyPerspectiveBottom = 29;
    var Camera_prst_legacyPerspectiveBottomLeft = 30;
    var Camera_prst_legacyPerspectiveBottomRight = 31;
    var Camera_prst_legacyPerspectiveFront = 32;
    var Camera_prst_legacyPerspectiveLeft = 33;
    var Camera_prst_legacyPerspectiveRight = 34;
    var Camera_prst_legacyPerspectiveTop = 35;
    var Camera_prst_legacyPerspectiveTopLeft = 36;
    var Camera_prst_legacyPerspectiveTopRight = 37;
    var Camera_prst_obliqueBottom = 38;
    var Camera_prst_obliqueBottomLeft = 39;
    var Camera_prst_obliqueBottomRight = 40;
    var Camera_prst_obliqueLeft = 41;
    var Camera_prst_obliqueRight = 42;
    var Camera_prst_obliqueTop = 43;
    var Camera_prst_obliqueTopLeft = 44;
    var Camera_prst_obliqueTopRight = 45;
    var Camera_prst_orthographicFront = 46;
    var Camera_prst_perspectiveAbove = 47;
    var Camera_prst_perspectiveAboveLeftFacing = 48;
    var Camera_prst_perspectiveAboveRightFacing = 49;
    var Camera_prst_perspectiveBelow = 50;
    var Camera_prst_perspectiveContrastingLeftFacing = 51;
    var Camera_prst_perspectiveContrastingRightFacing = 52;
    var Camera_prst_perspectiveFront = 53;
    var Camera_prst_perspectiveHeroicExtremeLeftFacing = 54;
    var Camera_prst_perspectiveHeroicExtremeRightFacing = 55;
    var Camera_prst_perspectiveHeroicLeftFacing = 56;
    var Camera_prst_perspectiveHeroicRightFacing = 57;
    var Camera_prst_perspectiveLeft = 58;
    var Camera_prst_perspectiveRelaxed = 59;
    var Camera_prst_perspectiveRelaxedModerately = 60;
    var Camera_prst_perspectiveRight = 61;

    var Sp3d_prstMaterial_clear = 0;
    var Sp3d_prstMaterial_dkEdge = 1;
    var Sp3d_prstMaterial_flat = 2;
    var Sp3d_prstMaterial_legacyMatte = 3;
    var Sp3d_prstMaterial_legacyMetal = 4;
    var Sp3d_prstMaterial_legacyPlastic = 5;
    var Sp3d_prstMaterial_legacyWireframe = 6;
    var Sp3d_prstMaterial_matte = 7;
    var Sp3d_prstMaterial_metal = 8;
    var Sp3d_prstMaterial_plastic = 9;
    var Sp3d_prstMaterial_powder = 10;
    var Sp3d_prstMaterial_softEdge = 11;
    var Sp3d_prstMaterial_softmetal = 12;
    var Sp3d_prstMaterial_translucentPowder = 13;
    var Sp3d_prstMaterial_warmMatte = 14;

    var LightRig_dir_b = 0;
    var LightRig_dir_bl = 1;
    var LightRig_dir_br = 2;
    var LightRig_dir_l = 3;
    var LightRig_dir_r = 4;
    var LightRig_dir_t = 5;
    var LightRig_dir_tl = 6;
    var LightRig_dir_tr = 7;

    var LightRig_rig_balanced = 0;
    var LightRig_rig_brightRoom = 1;
    var LightRig_rig_chilly = 2;
    var LightRig_rig_contrasting = 3;
    var LightRig_rig_flat = 4;
    var LightRig_rig_flood = 5;
    var LightRig_rig_freezing = 6;
    var LightRig_rig_glow = 7;
    var LightRig_rig_harsh = 8;
    var LightRig_rig_legacyFlat1 = 9;
    var LightRig_rig_legacyFlat2 = 10;
    var LightRig_rig_legacyFlat3 = 11;
    var LightRig_rig_legacyFlat4 = 12;
    var LightRig_rig_legacyHarsh1 = 13;
    var LightRig_rig_legacyHarsh2 = 14;
    var LightRig_rig_legacyHarsh3 = 15;
    var LightRig_rig_legacyHarsh4 = 16;
    var LightRig_rig_legacyNormal1 = 17;
    var LightRig_rig_legacyNormal2 = 18;
    var LightRig_rig_legacyNormal3 = 19;
    var LightRig_rig_legacyNormal4 = 20;
    var LightRig_rig_morning = 21;
    var LightRig_rig_soft = 22;
    var LightRig_rig_sunrise = 23;
    var LightRig_rig_sunset = 24;
    var LightRig_rig_threePt = 25;
    var LightRig_rig_twoPt = 26;

    var Bevel_prst_angle = 0;
    var Bevel_prst_artDeco = 1;
    var Bevel_prst_circle = 2;
    var Bevel_prst_convex = 3;
    var Bevel_prst_coolSlant = 4;
    var Bevel_prst_cross = 5;
    var Bevel_prst_divot = 6;
    var Bevel_prst_hardEdge = 7;
    var Bevel_prst_relaxedInset = 8;
    var Bevel_prst_riblet = 9;
    var Bevel_prst_slope = 10;
    var Bevel_prst_softRound = 11;

    var ParameterVal_arrowheadStyle_arr = 0;
    var ParameterVal_arrowheadStyle_auto = 1;
    var ParameterVal_arrowheadStyle_noArr = 2;
    var ParameterVal_autoTextRotation_grav = 0;
    var ParameterVal_autoTextRotation_none = 1;
    var ParameterVal_autoTextRotation_upr = 2;
    var ParameterVal_bendPoint_beg = 0;
    var ParameterVal_bendPoint_def = 1;
    var ParameterVal_bendPoint_end = 2;
    var ParameterVal_breakpoint_bal = 0;
    var ParameterVal_breakpoint_endCnv = 1;
    var ParameterVal_breakpoint_fixed = 2;
    var ParameterVal_centerShapeMapping_fNode = 0;
    var ParameterVal_centerShapeMapping_none = 1;
    var ParameterVal_childAlignment_b = 0;
    var ParameterVal_childAlignment_l = 1;
    var ParameterVal_childAlignment_r = 2;
    var ParameterVal_childAlignment_t = 3;
    var ParameterVal_childDirection_horz = 0;
    var ParameterVal_childDirection_vert = 1;
    var ParameterVal_connectorDimension_1D = 0;
    var ParameterVal_connectorDimension_2D = 1;
    var ParameterVal_connectorDimension_cust = 2;
    var ParameterVal_connectorPoint_auto = 0;
    var ParameterVal_connectorPoint_bCtr = 1;
    var ParameterVal_connectorPoint_bL = 2;
    var ParameterVal_connectorPoint_bR = 3;
    var ParameterVal_connectorPoint_ctr = 4;
    var ParameterVal_connectorPoint_midL = 5;
    var ParameterVal_connectorPoint_midR = 6;
    var ParameterVal_connectorPoint_radial = 7;
    var ParameterVal_connectorPoint_tCtr = 8;
    var ParameterVal_connectorPoint_tL = 9;
    var ParameterVal_connectorPoint_tR = 10;
    var ParameterVal_connectorRouting_bend = 0;
    var ParameterVal_connectorRouting_curve = 1;
    var ParameterVal_connectorRouting_longCurve = 2;
    var ParameterVal_connectorRouting_stra = 3;
    var ParameterVal_continueDirection_revDir = 0;
    var ParameterVal_continueDirection_sameDir = 1;
    var ParameterVal_diagramHorizontalAlignment_ctr = 0;
    var ParameterVal_diagramHorizontalAlignment_l = 1;
    var ParameterVal_diagramHorizontalAlignment_none = 2;
    var ParameterVal_diagramHorizontalAlignment_r = 3;
    var ParameterVal_diagramTextAlignment_ctr = 0;
    var ParameterVal_diagramTextAlignment_l = 1;
    var ParameterVal_diagramTextAlignment_r = 2;
    var ParameterVal_fallbackDimension_1D = 0;
    var ParameterVal_fallbackDimension_2D = 1;
    var ParameterVal_flowDirection_col = 0;
    var ParameterVal_flowDirection_row = 1;
    var ParameterVal_growDirection_bL = 0;
    var ParameterVal_growDirection_bR = 1;
    var ParameterVal_growDirection_tL = 2;
    var ParameterVal_growDirection_tR = 3;
    var ParameterVal_hierarchyAlignment_bCtrCh = 0;
    var ParameterVal_hierarchyAlignment_bCtrDes = 1;
    var ParameterVal_hierarchyAlignment_bL = 2;
    var ParameterVal_hierarchyAlignment_bR = 3;
    var ParameterVal_hierarchyAlignment_lB = 4;
    var ParameterVal_hierarchyAlignment_lCtrCh = 5;
    var ParameterVal_hierarchyAlignment_lCtrDes = 6;
    var ParameterVal_hierarchyAlignment_lT = 7;
    var ParameterVal_hierarchyAlignment_rB = 8;
    var ParameterVal_hierarchyAlignment_rCtrCh = 9;
    var ParameterVal_hierarchyAlignment_rCtrDes = 10;
    var ParameterVal_hierarchyAlignment_rT = 11;
    var ParameterVal_hierarchyAlignment_tCtrCh = 12;
    var ParameterVal_hierarchyAlignment_tCtrDes = 13;
    var ParameterVal_hierarchyAlignment_tL = 14;
    var ParameterVal_hierarchyAlignment_tR = 15;
    var ParameterVal_linearDirection_fromB = 0;
    var ParameterVal_linearDirection_fromL = 1;
    var ParameterVal_linearDirection_fromR = 2;
    var ParameterVal_linearDirection_fromT = 3;
    var ParameterVal_nodeHorizontalAlignment_ctr = 0;
    var ParameterVal_nodeHorizontalAlignment_l = 1;
    var ParameterVal_nodeHorizontalAlignment_r = 2;
    var ParameterVal_nodeVerticalAlignment_b = 0;
    var ParameterVal_nodeVerticalAlignment_mid = 1;
    var ParameterVal_nodeVerticalAlignment_t = 2;
    var ParameterVal_offset_ctr = 0;
    var ParameterVal_offset_off = 1;
    var ParameterVal_pyramidAccentPosition_aft = 0;
    var ParameterVal_pyramidAccentPosition_bef = 1;
    var ParameterVal_pyramidAccentTextMargin_stack = 0;
    var ParameterVal_pyramidAccentTextMargin_step = 1;
    var ParameterVal_rotationPath_alongPath = 0;
    var ParameterVal_rotationPath_none = 1;
    var ParameterVal_secondaryChildAlignment_b = 0;
    var ParameterVal_secondaryChildAlignment_l = 1;
    var ParameterVal_secondaryChildAlignment_none = 2;
    var ParameterVal_secondaryChildAlignment_r = 3;
    var ParameterVal_secondaryChildAlignment_t = 4;
    var ParameterVal_secondaryLinearDirection_fromB = 0;
    var ParameterVal_secondaryLinearDirection_fromL = 1;
    var ParameterVal_secondaryLinearDirection_fromR = 2;
    var ParameterVal_secondaryLinearDirection_fromT = 3;
    var ParameterVal_secondaryLinearDirection_none = 4;
    var ParameterVal_startingElement_node = 0;
    var ParameterVal_startingElement_trans = 1;
    var ParameterVal_textAnchorHorizontal_ctr = 0;
    var ParameterVal_textAnchorHorizontal_none = 1;
    var ParameterVal_textAnchorVertical_b = 0;
    var ParameterVal_textAnchorVertical_mid = 1;
    var ParameterVal_textAnchorVertical_top = 2;
    var ParameterVal_textBlockDirection_horz = 0;
    var ParameterVal_textBlockDirection_vert = 1;
    var ParameterVal_textDirection_fromB = 0;
    var ParameterVal_textDirection_fromT = 1;
    var ParameterVal_verticalAlignment_b = 0;
    var ParameterVal_verticalAlignment_mid = 1;
    var ParameterVal_verticalAlignment_none = 2;
    var ParameterVal_verticalAlignment_t = 3;

    var FunctionValue_animLvlStr_ctr = 0;
    var FunctionValue_animLvlStr_lvl = 1;
    var FunctionValue_animLvlStr_none = 2;
    var FunctionValue_animOneStr_branch = 0;
    var FunctionValue_animOneStr_none = 1;
    var FunctionValue_animOneStr_one = 2;
    var FunctionValue_direction_norm = 0;
    var FunctionValue_direction_rev = 1;
    var FunctionValue_hierBranchStyle_hang = 0;
    var FunctionValue_hierBranchStyle_init = 1;
    var FunctionValue_hierBranchStyle_l = 2;
    var FunctionValue_hierBranchStyle_r = 3;
    var FunctionValue_hierBranchStyle_std = 4;
    var FunctionValue_resizeHandlesStr_exact = 0;
    var FunctionValue_resizeHandlesStr_rel = 1;

    var Coordinate_universalMeasure_cm = 0;
    var Coordinate_universalMeasure_mm = 1;
    var Coordinate_universalMeasure_in = 2;
    var Coordinate_universalMeasure_pt = 3;
    var Coordinate_universalMeasure_pc = 4;
    var Coordinate_universalMeasure_pi = 5;

    var Constr_font_scale = 360;

    if (!String.prototype.includes) {
      String.prototype.includes = function(search, start) {
        if (typeof start !== 'number') {
          start = 0;
        }

        if (start + search.length > this.length) {
          return false;
        } else {
          return this.indexOf(search, start) !== -1;
        }
      };
    }

    changesFactory[AscDFH.historyitem_DiagramDataDataModel] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_DiagramDataDataModel] = function (oClass, value) {
      oClass.dataModel = value;
    };
    function DiagramData() {
      CBaseFormatObject.call(this);
      this.dataModel = null;
    }

    InitClass(DiagramData, CBaseFormatObject, AscDFH.historyitem_type_DiagramData);

    DiagramData.prototype.setDataModel = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DiagramDataDataModel, this.getDataModel(), oPr));
      this.dataModel = oPr;
      this.setParentToChild(oPr);
    }

    DiagramData.prototype.getDataModel = function () {
      return this.dataModel;
    }

    DiagramData.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.dataModel) {
        oCopy.setDataModel(this.dataModel.createDuplicate(oIdMap));
      }
    }

    DiagramData.prototype.privateWriteAttributes = null;
    DiagramData.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.dataModel);
    };
    DiagramData.prototype.readAttribute = null;
    DiagramData.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setDataModel(new DataModel());
          this.dataModel.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    DiagramData.prototype.getChildren = function() {
      return [this.dataModel];
    };


    changesFactory[AscDFH.historyitem_DataModelBg] = CChangeObject;
    changesFactory[AscDFH.historyitem_DataModelCxnLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_DataModelExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_DataModelPtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_DataModelWhole] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_DataModelBg] = function (oClass, value) {
      oClass.bg = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DataModelCxnLst] = function (oClass, value) {
      oClass.cxnLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DataModelExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DataModelPtLst] = function (oClass, value) {
      oClass.ptLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DataModelWhole] = function (oClass, value) {
      oClass.whole = value;
    };

    function DataModel() {
      CBaseFormatObject.call(this);
      this.bg = null;
      this.cxnLst = null;
      this.extLst = null;
      this.ptLst = null;
      this.whole = null;
    }

    InitClass(DataModel, CBaseFormatObject, AscDFH.historyitem_type_DataModel);

    DataModel.prototype.setBg = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelBg, this.getBg(), oPr));
      this.bg = oPr;
      this.setParentToChild(oPr);
    }

    DataModel.prototype.setCxnLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelCxnLst, this.getCxnLst(), oPr));
      this.cxnLst = oPr;
      this.setParentToChild(oPr);
    }

    DataModel.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    DataModel.prototype.setPtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelPtLst, this.getPtLst(), oPr));
      this.ptLst = oPr;
      this.setParentToChild(oPr);
    }

    DataModel.prototype.setWhole = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelWhole, this.getWhole(), oPr));
      this.whole = oPr;
      this.setParentToChild(oPr);
    }

    DataModel.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.bg) {
        oCopy.setBg(this.bg.createDuplicate(oIdMap));
      }
      if (this.cxnLst) {
        oCopy.setCxnLst(this.cxnLst.createDuplicate(oIdMap));
      }
      if (this.extLst) {
        oCopy.setExtLst(this.extLst.createDuplicate(oIdMap));
      }
      if (this.ptLst) {
        oCopy.setPtLst(this.ptLst.createDuplicate(oIdMap));
      }
      if (this.whole) {
        oCopy.setWhole(this.whole.createDuplicate(oIdMap));
      }
    }

    DataModel.prototype.getBg = function () {
      return this.bg;
    }

    DataModel.prototype.getCxnLst = function () {
      return this.cxnLst;
    }

    DataModel.prototype.getExtLst = function () {
      return this.extLst;
    }

    DataModel.prototype.getPtLst = function () {
      return this.ptLst;
    }

    DataModel.prototype.getWhole = function () {
      return this.whole;
    }

    DataModel.prototype.privateWriteAttributes = null;
    DataModel.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.ptLst);
      this.writeRecord2(pWriter, 1, this.cxnLst);
      this.writeRecord2(pWriter, 2, this.whole);
      this.writeRecord2(pWriter, 3, this.bg);
    };
    DataModel.prototype.readAttribute = null;
    DataModel.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setPtLst(new PtLst());
          this.ptLst.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setCxnLst(new CxnLst());
          this.cxnLst.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setWhole(new Whole());
          this.whole.fromPPTY(pReader);
          break;
        }
        case 3: {
          this.setBg(new BgFormat());
          this.bg.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    DataModel.prototype.getChildren = function() {
      return [this.ptLst, this.cxnLst, this.whole, this.bg];
    };


    changesFactory[AscDFH.historyitem_CCommonDataListAdd] = CChangeContent;
    changesFactory[AscDFH.historyitem_CCommonDataListRemove] = CChangeContent;
    drawingContentChanges[AscDFH.historyitem_CCommonDataListAdd] = function (oClass) {
      return oClass.list;
    };
    drawingContentChanges[AscDFH.historyitem_CCommonDataListRemove] = function (oClass) {
      return oClass.list;
    };

    function CCommonDataList() {
      CBaseFormatObject.call(this);
      this.list = [];
    }

    InitClass(CCommonDataList, CBaseFormatObject, AscDFH.historyitem_type_CCommonDataList);

    CCommonDataList.prototype.addToLst = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_CCommonDataListAdd, nInsertIdx, [oPr], true));
      this.list.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    CCommonDataList.prototype.removeFromLst = function (nIdx) {
      if (nIdx > -1 && nIdx < this.list.length) {
        this.list[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_CCommonDataListRemove, nIdx, [this.list[nIdx]], false));
        this.list.splice(nIdx, 1);
      }
    };

    CCommonDataList.prototype.fillObject = function (oCopy, oIdMap) {
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        oCopy.addToLst(nIdx, this.list[nIdx].createDuplicate(oIdMap));
      }
    };

    CCommonDataList.prototype.getChildren = function() {
      return [].concat(this.list);
    };
    CCommonDataList.prototype.privateWriteAttributes = null;
    CCommonDataList.prototype.writeChildren = function(pWriter) {
      for (var i = 0; i < this.list.length; i += 1) {
        this.writeRecord2(pWriter, 0, this.list[i]);
      }
    };
    CCommonDataList.prototype.readAttribute = null;


    function PtLst() {
      CCommonDataList.call(this);
    }

    InitClass(PtLst, CCommonDataList, AscDFH.historyitem_type_PtLst);

    PtLst.prototype.privateWriteAttributes = null;
    PtLst.prototype.writeChildren = function(pWriter) {
      for (var i = 0; i < this.list.length; i += 1) {
        this.writeRecord2(pWriter, 0, this.list[i]);
      }
    };
    PtLst.prototype.readAttribute = null;
    PtLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          var oChild = new Point();
          oChild.fromPPTY(pReader);
          this.addToLst(this.list.length, oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };


    function CxnLst() {
      CCommonDataList.call(this);
    }

    InitClass(CxnLst, CCommonDataList, AscDFH.historyitem_type_CxnLst);

    CxnLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          var oChild = new Cxn();
          oChild.fromPPTY(pReader);
          this.addToLst(this.list.length, oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    changesFactory[AscDFH.historyitem_CxnDestId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnDestOrd] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnModelId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnParTransId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnPresId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnSibTransId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnSrcId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnSrcOrd] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnType] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnExtLst] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_CxnDestId] = function (oClass, value) {
      oClass.destId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnDestOrd] = function (oClass, value) {
      oClass.destOrd = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnModelId] = function (oClass, value) {
      oClass.modelId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnParTransId] = function (oClass, value) {
      oClass.parTransId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnPresId] = function (oClass, value) {
      oClass.presId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnSibTransId] = function (oClass, value) {
      oClass.sibTransId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnSrcId] = function (oClass, value) {
      oClass.srcId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnSrcOrd] = function (oClass, value) {
      oClass.srcOrd = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnType] = function (oClass, value) {
      oClass.type = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CxnExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };

    function Cxn() {
      CBaseFormatObject.call(this);
      this.destId = null;
      this.destOrd = null;
      this.modelId = null;
      this.parTransId = null;
      this.presId = null;
      this.sibTransId = null;
      this.srcId = null;
      this.srcOrd = null;
      this.type = null;

      this.extLst = null;
    }

    InitClass(Cxn, CBaseFormatObject, AscDFH.historyitem_type_Cxn);

    Cxn.prototype.setDestId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnDestId, this.getDestId(), pr));
      this.destId = pr;
    }

    Cxn.prototype.setDestOrd = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnDestOrd, this.getDestOrd(), pr));
      this.destOrd = pr;
    }

    Cxn.prototype.setModelId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnModelId, this.getModelId(), pr));
      this.modelId = pr;
    }

    Cxn.prototype.setParTransId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnParTransId, this.getParTransId(), pr));
      this.parTransId = pr;
    }

    Cxn.prototype.setPresId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnPresId, this.getPresId(), pr));
      this.presId = pr;
    }

    Cxn.prototype.setSibTransId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnSibTransId, this.getSibTransId(), pr));
      this.sibTransId = pr;
    }

    Cxn.prototype.setSrcId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnSrcId, this.getSrcId(), pr));
      this.srcId = pr;
    }

    Cxn.prototype.setSrcOrd = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnSrcOrd, this.getSrcOrd(), pr)); // TODO: srcord, type is long maybe
      this.srcOrd = pr;
    }

    Cxn.prototype.setType = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnType, this.getType(), pr));
      this.type = pr;
    }

    Cxn.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CxnExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    Cxn.prototype.getDestId = function () {
      return this.destId;
    }

    Cxn.prototype.getDestOrd = function () {
      return this.destOrd;
    }

    Cxn.prototype.getModelId = function () {
      return this.modelId;
    }

    Cxn.prototype.getParTransId = function () {
      return this.parTransId;
    }

    Cxn.prototype.getPresId = function () {
      return this.presId;
    }

    Cxn.prototype.getSibTransId = function () {
      return this.sibTransId;
    }

    Cxn.prototype.getSrcId = function () {
      return this.srcId;
    }

    Cxn.prototype.getSrcOrd = function () {
      return this.srcOrd;
    }

    Cxn.prototype.getType = function () {
      return this.type;
    }

    Cxn.prototype.getExtLst = function () {
      return this.extLst;
    }

    Cxn.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setDestId(this.getDestId());
      oCopy.setDestOrd(this.getDestOrd());
      oCopy.setModelId(this.getModelId());
      oCopy.setParTransId(this.getParTransId());
      oCopy.setPresId(this.getPresId());
      oCopy.setSibTransId(this.getSibTransId());
      oCopy.setSrcId(this.getSrcId());
      oCopy.setSrcOrd(this.getSrcOrd());
      oCopy.setType(this.getType());
      if (this.getExtLst()) {
        oCopy.setExtLst(this.extLst.createDuplicate(oIdMap));
      }
    }

    Cxn.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.modelId);
      pWriter._WriteString2(1, this.type);
      pWriter._WriteString2(2, this.destId);
      pWriter._WriteString2(3, this.destOrd);
      pWriter._WriteString2(4, this.srcId);
      pWriter._WriteString2(5, this.srcOrd);
      pWriter._WriteString2(6, this.parTransId);
      pWriter._WriteString2(7, this.sibTransId);
      pWriter._WriteString2(8, this.presId);
    };
    Cxn.prototype.writeChildren = function(pWriter) {
    };
    Cxn.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setModelId(oStream.GetString2());
      else if (1 === nType) this.setType(oStream.GetString2());
      else if (2 === nType) this.setDestId(oStream.GetString2());
      else if (3 === nType) this.setDestOrd(oStream.GetString2());
      else if (4 === nType) this.setSrcId(oStream.GetString2());
      else if (5 === nType) this.setSrcOrd(oStream.GetString2());
      else if (6 === nType) this.setParTransId(oStream.GetString2());
      else if (7 === nType) this.setSibTransId(oStream.GetString2());
      else if (8 === nType) this.setPresId(oStream.GetString2());
    };
    Cxn.prototype.readChild = function(nType, pReader) {

    };


    function ExtLst() {
      CCommonDataList.call(this);
    }

    InitClass(ExtLst, CCommonDataList, AscDFH.historyitem_type_ExtLst);

    ExtLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          var oChild = new Ext();
          oChild.fromPPTY(pReader);
          this.addToLst(this.list.length, oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    changesFactory[AscDFH.historyitem_ExtUri] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_ExtUri] = function (oClass, value) {
      oClass.uri = value;
    };

    function Ext() {
      CBaseFormatObject.call(this);
      this.uri = null;
    }

    InitClass(Ext, CBaseFormatObject, AscDFH.historyitem_type_Ext);

    Ext.prototype.setUri = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ExtUri, this.getUri(), pr));
      this.uri = pr;
    }

    Ext.prototype.getUri = function () {
      return this.uri;
    }

    Ext.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setUri(this.getUri());
    }

    changesFactory[AscDFH.historyitem_BgFormatFill] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_BgFormatEffect] = CChangeObjectNoId;
    drawingsChangesMap[AscDFH.historyitem_BgFormatFill] = function (oClass, value) {
      oClass.fill = value;
      oClass.handleUpdateFill();
    };
    drawingsChangesMap[AscDFH.historyitem_BgFormatEffect] = function (oClass, value) {
      oClass.effect = value;
    };

    drawingConstructorsMap[AscDFH.historyitem_BgFormatFill] = AscFormat.CUniFill;
    drawingConstructorsMap[AscDFH.historyitem_BgFormatEffect] = AscFormat.CEffectProperties;

    function BgFormat() {
      CBaseFormatObject.call(this);
      this.fill = null;
      this.effect = null;
    }

    InitClass(BgFormat, CBaseFormatObject, AscDFH.historyitem_type_BgFormat);


    BgFormat.prototype.setFill = function (oPr) {
      oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_BgFormatFill, this.getFill(), oPr));
      this.fill = oPr;
      this.handleUpdateFill();
    }

    BgFormat.prototype.setEffect = function (oPr) {
      oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_BgFormatEffect, this.getEffect(), oPr));
      this.effect = oPr;
      this.setParentToChild(oPr);
    }

    BgFormat.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getFill()) {
        oCopy.setFill(this.getFill().createDuplicate(oIdMap));
      }
      if (this.getEffect()) {
        oCopy.setEffect(this.getEffect().createDuplicate(oIdMap));
      }
    }

    BgFormat.prototype.getFill = function () {
      return this.fill;
    }

    BgFormat.prototype.getEffect = function () {
      return this.effect;
    }

    BgFormat.prototype.privateWriteAttributes = null;
    BgFormat.prototype.writeChildren = function(pWriter) {
      pWriter.WriteRecord1(0, this.fill, pWriter.WriteUniFill);
      var oEffectPr = this.effect;
      if(oEffectPr)
      {
        if(oEffectPr.EffectLst)
        {
          pWriter.WriteRecord1(1, oEffectPr.EffectLst, pWriter.WriteEffectLst);
        }
        else if(oEffectPr.EffectDag)
        {
          pWriter.WriteRecord1(1, oEffectPr.EffectDag, pWriter.WriteEffectDag)
        }
      }
    };
    BgFormat.prototype.readAttribute = null;
    BgFormat.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setFill(pReader.ReadUniFill());
          break;
        }
        case 1: {
          this.setEffect(pReader.ReadEffectProperties());
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    BgFormat.prototype.getChildren = function() {
      return [this.fill, this.effect];
    };

    BgFormat.prototype.getSmartArt = function() {
      var oCurParent = this.parent;
      while (oCurParent) {
        if(oCurParent instanceof SmartArt) {
          break;
        }
        oCurParent = oCurParent.parent;
      }
      return oCurParent;
    };
    BgFormat.prototype.handleUpdateFill = function() {
      var oSmartArt = this.getSmartArt();
      if(oSmartArt) {
        oSmartArt.handleUpdateFill();
      }
    };
    BgFormat.prototype.Refresh_RecalcData = function(data)
    {
      switch(data.Type)
      {
        case AscDFH.historyitem_BgFormatFill:
        {
          this.handleUpdateFill();
          break;
        }
      }
    };

    BgFormat.prototype.Refresh_RecalcData2 = function(data)
    {
    };



    changesFactory[AscDFH.historyitem_WholeEffect] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_WholeLn] = CChangeObjectNoId;
    drawingConstructorsMap[AscDFH.historyitem_WholeLn] = AscFormat.CLn;
    drawingConstructorsMap[AscDFH.historyitem_WholeEffect] = AscFormat.CEffectProperties;
    drawingsChangesMap[AscDFH.historyitem_WholeEffect] = function (oClass, value) {
      oClass.effect = value;
    };
    drawingsChangesMap[AscDFH.historyitem_WholeLn] = function (oClass, value) {
      oClass.ln = value;
    };

    function Whole() {
      CBaseFormatObject.call(this);
      this.effect = null;
      this.ln = null;
    }

    InitClass(Whole, CBaseFormatObject, AscDFH.historyitem_type_Whole);

    Whole.prototype.setEffect = function (oPr) {
      oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_WholeEffect, this.getEffect(), oPr));
      this.effect = oPr;
      this.setParentToChild(oPr);
    }

    Whole.prototype.setLn = function (oPr) {
      oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_WholeLn, this.getLn(), oPr));
      this.ln = oPr;
      this.setParentToChild(oPr);
    }

    Whole.prototype.getEffect = function () {
      return this.effect;
    }

    Whole.prototype.getLn = function () {
      return this.ln;
    }

    Whole.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getEffect()) {
        oCopy.setEffect(this.effect.createDuplicate(oIdMap));
      }
      if (this.getLn()) {
        oCopy.setLn(this.ln.createDuplicate(oIdMap));
      }
    }

    Whole.prototype.privateWriteAttributes = null;
    Whole.prototype.writeChildren = function(pWriter) {
      pWriter.WriteRecord2(0, this.ln, pWriter.WriteLn);
      var oEffectPr = this.effect;
      if(oEffectPr)
      {
        if(oEffectPr.EffectLst)
        {
          pWriter.WriteRecord1(1, oEffectPr.EffectLst, pWriter.WriteEffectLst);
        }
        else if(oEffectPr.EffectDag)
        {
          pWriter.WriteRecord1(1, oEffectPr.EffectDag, pWriter.WriteEffectDag)
        }
      }
    };
    Whole.prototype.readAttribute = null;
    Whole.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setLn(pReader.ReadLn());
          break;
        }
        case 1: {
          this.setEffect(pReader.ReadEffectProperties());
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    Whole.prototype.getChildren = function() {
      return [this.ln, this.effect];
    };

    Whole.prototype.getSmartArt = function() {
      var oCurParent = this.parent;
      while (oCurParent) {
        if(oCurParent instanceof SmartArt) {
          break;
        }
        oCurParent = oCurParent.parent;
      }
      return oCurParent;
    };
    Whole.prototype.handleUpdateLn = function() {
      var oSmartArt = this.getSmartArt();
      if(oSmartArt) {
        oSmartArt.handleUpdateLn();
      }
    };
    Whole.prototype.Refresh_RecalcData = function(data)
    {
      switch(data.Type)
      {
        case AscDFH.historyitem_WholeLn:
        {
          this.handleUpdateLn();
          break;
        }
      }
    };

    Whole.prototype.Refresh_RecalcData2 = function(data)
    {
    };


    changesFactory[AscDFH.historyitem_PointInfoPoint] = CChangeObject;
    changesFactory[AscDFH.historyitem_PointInfoAssociation] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_PointInfoPoint] = function (oClass, value) {
      oClass.point = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PointInfoAssociation] = function (oClass, value) {
      oClass.association = value;
    };

    changesFactory[AscDFH.historyitem_PointCxnId] = CChangeString;
    changesFactory[AscDFH.historyitem_PointModelId] = CChangeString;
    changesFactory[AscDFH.historyitem_PointType] = CChangeLong;
    changesFactory[AscDFH.historyitem_PointExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_PointPrSet] = CChangeObject;
    changesFactory[AscDFH.historyitem_PointSpPr] = CChangeObject;
    changesFactory[AscDFH.historyitem_PointT] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_PointCxnId] = function (oClass, value) {
      oClass.cxnId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PointModelId] = function (oClass, value) {
      oClass.modelId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PointType] = function (oClass, value) {
      oClass.type = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PointExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PointPrSet] = function (oClass, value) {
      oClass.prSet = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PointSpPr] = function (oClass, value) {
      oClass.spPr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PointT] = function (oClass, value) {
      oClass.t = value;
    };

    function Point() {
      CBaseFormatObject.call(this);
      this.cxnId = null;
      this.modelId = null;
      this.type = null;

      this.extLst = null;
      this.prSet = null;
      this.spPr = null;
      this.t = null;
    }

    InitClass(Point, CBaseFormatObject, AscDFH.historyitem_type_Point);

    Point.prototype.getDrawingDocument = function () {
    }

    Point.prototype.isForm = function () {
      return false;
    }

    Point.prototype.Get_Theme = function () {
      return null;
    }

    Point.prototype.Get_ColorMap = function() {
      return null;
    };

    Point.prototype.setCxnId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PointCxnId, this.getCxnId(), pr));
      this.cxnId = pr;
    }

    Point.prototype.setModelId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PointModelId, this.getModelId(), pr));
      this.modelId = pr;

    }

    Point.prototype.getShape = function () {
      if (this.parent && this.parent.parent instanceof CShape) {
        return this.parent.parent;
      }
    }

    Point.prototype.isRecalculateInsets = function () {
      const insets = {Top: true, Bottom: true, Left: true, Right: true};
      if (this.t) {
        var bodyPr = this.t.bodyPr;
        if (bodyPr) {
          if (AscFormat.isRealNumber(bodyPr.tIns)) {
            insets.Top = false;
          }
          if (AscFormat.isRealNumber(bodyPr.bIns)) {
            insets.Bottom = false;
          }
          if (AscFormat.isRealNumber(bodyPr.lIns)) {
            insets.Left = false;
          }
          if (AscFormat.isRealNumber(bodyPr.rIns)) {
            insets.Right = false;
          }
        }
      }
      return insets;
    }

    Point.prototype.setType = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PointType, this.getType(), pr));
      this.type = pr;
    }

    Point.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PointExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    Point.prototype.setPrSet = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PointPrSet, this.getPrSet(), oPr));
      this.prSet = oPr;
      this.setParentToChild(oPr);
    }

    Point.prototype.setSpPr = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PointSpPr, this.getSpPr(), oPr));
      this.spPr = oPr;
      this.setParentToChild(oPr);
    }

    Point.prototype.setT = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PointT, this.getT(), oPr));
      this.t = oPr;
      this.setParentToChild(oPr);
    }

    Point.prototype.setPhldrT = function(pr) {
      var prSet = this.getPrSet();
      prSet && prSet.setPhldrT(pr);
    }

    Point.prototype.getCxnId = function () {
      return this.cxnId;
    }

    Point.prototype.getModelId = function () {
      return this.modelId;

    }

    Point.prototype.getType = function () {
      return this.type;
    }

    Point.prototype.getExtLst = function () {
      return this.extLst;
    }

    Point.prototype.getPrSet = function () {
      return this.prSet;
    }

    Point.prototype.getSpPr = function () {
      return this.spPr;
    }

    Point.prototype.getT = function () {
      return this.t;
    }

    Point.prototype.initSpPr = function () {
      if (!this.spPr) {
        this.setSpPr(new AscFormat.CSpPr());
      }
    }

    Point.prototype.changeFlipH = function (bFlipH) {
      var prSet = this.getPrSet();
      prSet && prSet.setCustFlipHor(bFlipH);
    }

    Point.prototype.changeFlipV = function (bFlipV) {
      var prSet = this.getPrSet();
      prSet && prSet.setCustFlipVert(bFlipV);
    }

    Point.prototype.resetUniFill = function () {
      this.spPr && this.spPr.setFill(null);
    }

    Point.prototype.setUniFill = function (unifill) {
      this.initSpPr();
      this.spPr.setFill(unifill);
    }

    Point.prototype.changeShadow = function (shadow) {
      this.initSpPr(shadow);
      this.spPr.changeShadow(shadow);
    }

    Point.prototype.setLine = function (line) {
      this.initSpPr();
      this.spPr.setLn(line);
    }

    Point.prototype.setGeometry = function (geometry) {
      this.initSpPr();
      this.spPr.setGeometry(geometry);
    }

    Point.prototype.isBlipFillPlaceholder = function () {
      var imagePlaceholderArrStylelbl = ['alignImgPlace1', 'bgImgPlace1', 'fgImgPlace1'];
      var imagePlaceholderArrName = ['Image', 'image', 'imageRepeatNode', 'pictRect'];
      var pointAssociationPrSet = this.prSet;
      return pointAssociationPrSet && ((imagePlaceholderArrStylelbl.indexOf(pointAssociationPrSet.presStyleLbl) !== -1) ||
        (imagePlaceholderArrName.indexOf(pointAssociationPrSet.presName) !== -1) ||
        (pointAssociationPrSet.presName === 'rect1' && pointAssociationPrSet.presStyleLbl === 'bgShp') ||
        (pointAssociationPrSet.presName === 'rect1' && pointAssociationPrSet.presStyleLbl === 'lnNode1') ||
        (pointAssociationPrSet.presName === 'adorn' && pointAssociationPrSet.presStyleLbl === 'fgAccFollowNode1'))
    }

    Point.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setCxnId(this.getCxnId());
      oCopy.setModelId(this.getModelId());
      oCopy.setType(this.getType());
      if (this.extLst) {
        oCopy.setExtLst(this.getExtLst().createDuplicate());
      }
      if (this.prSet) {
        oCopy.setPrSet(this.getPrSet().createDuplicate());
      }
      if (this.spPr) {
        oCopy.setSpPr(this.getSpPr().createDuplicate());
      }
      if (this.t) {
        oCopy.setT(this.getT().createDuplicate());
      }
    }
    Point.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.modelId);
      pWriter._WriteUChar2(1, this.type);
      pWriter._WriteString2(2, this.cxnId);
    };
    Point.prototype.writeChildren = function(pWriter) {
      pWriter.WriteRecord2(0, this.spPr, pWriter.WriteSpPr);
      pWriter.WriteRecord2(1, this.t, pWriter.WriteTxBody);
      this.writeRecord2(pWriter, 2, this.prSet);
    };
    Point.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setModelId(oStream.GetString2());
      else if (1 === nType) this.setType(oStream.GetUChar());
      else if (2 === nType) this.setCxnId(oStream.GetString2());
    };
    Point.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          var sppr = new AscFormat.CSpPr();
          this.setSpPr(sppr);
          pReader.ReadSpPr(this.spPr);
          break;
        }
        case 1: {
          this.setT(pReader.ReadTextBody());
          break;
        }
        case 2: {
          this.setPrSet(new PrSet());
          this.prSet.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    Point.prototype.getChildren = function() {
      return [this.spPr, this.t, this.prSet];
    };

    Point.prototype.getPresStyleLbl = function () {
      return this.prSet && this.prSet.presStyleLbl;
    };

    Point.prototype.getCustAng = function () {
      var prSet = this.getPrSet();
      if (prSet) {
        return prSet.getCustAng();
      }
    };



    changesFactory[AscDFH.historyitem_PrSetCoherent3DOff] = CChangeBool;
    changesFactory[AscDFH.historyitem_PrSetCsCatId] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetCsTypeId] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetCustAng] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustFlipHor] = CChangeBool;
    changesFactory[AscDFH.historyitem_PrSetCustFlipVert] = CChangeBool;
    changesFactory[AscDFH.historyitem_PrSetCustLinFactNeighborX] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustLinFactNeighborY] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustLinFactX] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustLinFactY] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustRadScaleInc] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustRadScaleRad] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustScaleX] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustScaleY] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_PrSetCustSzX] = CChangeLong;
    changesFactory[AscDFH.historyitem_PrSetCustSzY] = CChangeLong;
    changesFactory[AscDFH.historyitem_PrSetCustT] = CChangeBool;
    changesFactory[AscDFH.historyitem_PrSetLoCatId] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetLoTypeId] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetPhldr] = CChangeBool;
    changesFactory[AscDFH.historyitem_PrSetPhldrT] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetPresAssocID] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetPresName] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetPresStyleCnt] = CChangeLong;
    changesFactory[AscDFH.historyitem_PrSetPresStyleIdx] = CChangeLong;
    changesFactory[AscDFH.historyitem_PrSetPresStyleLbl] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetQsCatId] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetQsTypeId] = CChangeString;
    changesFactory[AscDFH.historyitem_PrSetStyle] = CChangeObject;
    changesFactory[AscDFH.historyitem_PrSetPresLayoutVars] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_PrSetCoherent3DOff] = function (oClass, value) {
      oClass.coherent3DOff = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCsCatId] = function (oClass, value) {
      oClass.csCatId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCsTypeId] = function (oClass, value) {
      oClass.csTypeId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustAng] = function (oClass, value) {
      oClass.custAng = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustFlipHor] = function (oClass, value) {
      oClass.custFlipHor = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustFlipVert] = function (oClass, value) {
      oClass.custFlipVert = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustLinFactNeighborX] = function (oClass, value) {
      oClass.custLinFactNeighborX = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustLinFactNeighborY] = function (oClass, value) {
      oClass.custLinFactNeighborY = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustLinFactX] = function (oClass, value) {
      oClass.custLinFactX = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustLinFactY] = function (oClass, value) {
      oClass.custLinFactY = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustRadScaleInc] = function (oClass, value) {
      oClass.custRadScaleInc = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustRadScaleRad] = function (oClass, value) {
      oClass.custRadScaleRad = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustScaleX] = function (oClass, value) {
      oClass.custScaleX = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustScaleY] = function (oClass, value) {
      oClass.custScaleY = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustSzX] = function (oClass, value) {
      oClass.custSzX = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustSzY] = function (oClass, value) {
      oClass.custSzY = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetCustT] = function (oClass, value) {
      oClass.custT = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetLoCatId] = function (oClass, value) {
      oClass.loCatId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetLoTypeId] = function (oClass, value) {
      oClass.loTypeId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetPhldr] = function (oClass, value) {
      oClass.phldr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetPhldrT] = function (oClass, value) {
      oClass.phldrT = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetPresAssocID] = function (oClass, value) {
      oClass.presAssocID = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetPresName] = function (oClass, value) {
      oClass.presName = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetPresStyleCnt] = function (oClass, value) {
      oClass.presStyleCnt = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetPresStyleIdx] = function (oClass, value) {
      oClass.presStyleIdx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetPresStyleLbl] = function (oClass, value) {
      oClass.presStyleLbl = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetQsCatId] = function (oClass, value) {
      oClass.qsCatId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetQsTypeId] = function (oClass, value) {
      oClass.qsTypeId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetStyle] = function (oClass, value) {
      oClass.style = value;
    };
    drawingsChangesMap[AscDFH.historyitem_PrSetPresLayoutVars] = function (oClass, value) {
      oClass.presLayoutVars = value;
    };

    function PrSet() {
      CBaseFormatObject.call(this);
      this.coherent3DOff = null;
      this.csCatId = null;
      this.csTypeId = null;
      this.custAng = null;
      this.custFlipHor = null;
      this.custFlipVert = null;
      this.custLinFactNeighborX = null;
      this.custLinFactNeighborY = null;
      this.custLinFactX = null;
      this.custLinFactY = null;
      this.custRadScaleInc = null;
      this.custRadScaleRad = null;
      this.custScaleX = null;
      this.custScaleY = null;
      this.custSzX = null;
      this.custSzY = null;
      this.custT = null;
      this.loCatId = null;
      this.loTypeId = null;
      this.phldr = null;
      this.phldrT = null;
      this.presAssocID = null;
      this.presName = null;
      this.presStyleCnt = null;
      this.presStyleIdx = null;
      this.presStyleLbl = null;
      this.qsCatId = null;
      this.qsTypeId = null;

      this.presLayoutVars = null;
      this.style = null;
    }

    InitClass(PrSet, CBaseFormatObject, AscDFH.historyitem_type_PrSet);


    PrSet.prototype.setCoherent3DOff = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetCoherent3DOff, this.getCoherent3DOff(), pr));
      this.coherent3DOff = pr;
    }

    PrSet.prototype.setCsCatId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetCsCatId, this.getCsCatId(), pr));
      this.csCatId = pr;
    }

    PrSet.prototype.setCsTypeId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetCsTypeId, this.getCsTypeId(), pr))
      this.csTypeId = pr;
    }

    PrSet.prototype.setCustAng = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustAng, this.getCustAng(), pr));
      this.custAng = pr;
    }

    PrSet.prototype.setCustFlipHor = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetCustFlipHor, this.getCustFlipHor(), pr));
      this.custFlipHor = pr;
    }

    PrSet.prototype.setCustFlipVert = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetCustFlipVert, this.getCustFlipVert(), pr));
      this.custFlipVert = pr;
    }

    PrSet.prototype.setCustLinFactNeighborX = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustLinFactNeighborX, this.getCustLinFactNeighborX(), pr));
      this.custLinFactNeighborX = pr;
    }

    PrSet.prototype.setCustLinFactNeighborY = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustLinFactNeighborY, this.getCustLinFactNeighborY(), pr));
      this.custLinFactNeighborY = pr;
    }

    PrSet.prototype.setCustLinFactX = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustLinFactX, this.getCustLinFactX(), pr));
      this.custLinFactX = pr;
    }

    PrSet.prototype.setCustLinFactY = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustLinFactY, this.getCustLinFactY(), pr));
      this.custLinFactY = pr;
    }

    PrSet.prototype.setCustRadScaleInc = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustRadScaleInc, this.getCustRadScaleInc(), pr));
      this.custRadScaleInc = pr;
    }

    PrSet.prototype.setCustRadScaleRad = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustRadScaleRad, this.getCustRadScaleRad(), pr));
      this.custRadScaleRad = pr;
    }

    PrSet.prototype.setCustScaleX = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustScaleX, this.getCustScaleX(), pr));
      this.custScaleX = pr;
    }

    PrSet.prototype.setCustScaleY = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustScaleY, this.getCustScaleY(), pr));
      this.custScaleY = pr;
    }

    PrSet.prototype.setCustSzX = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PrSetCustSzX, this.getCustSzX(), pr));
      this.custSzX = pr;
    }

    PrSet.prototype.setCustSzY = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PrSetCustSzY, this.getCustSzY(), pr));
      this.custSzY = pr;
    }

    PrSet.prototype.setCustT = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetCustT, this.getCustT(), pr));
      this.custT = pr;
    }

    PrSet.prototype.setLoCatId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetLoCatId, this.getLoCatId(), pr));
      this.loCatId = pr;
    }

    PrSet.prototype.setLoTypeId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetLoTypeId, this.getLoTypeId(), pr));
      this.loTypeId = pr;
    }

    PrSet.prototype.setPhldr = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetPhldr, this.getPhldr(), pr));
      this.phldr = pr;
    }

    PrSet.prototype.setPhldrT = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetPhldrT, this.getPhldrT(), pr));
      this.phldrT = pr;
    }

    PrSet.prototype.setPresAssocID = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetPresAssocID, this.getPresAssocID(), pr));
      this.presAssocID = pr;
    }

    PrSet.prototype.setPresName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetPresName, this.getPresName(), pr));
      this.presName = pr;
    }
    PrSet.prototype.setPresStyleCnt = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PrSetPresStyleCnt, this.getPresStyleCnt(), pr));
      this.presStyleCnt = pr;
    }

    PrSet.prototype.setPresStyleIdx = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PrSetPresStyleIdx, this.getPresStyleIdx(), pr));
      this.presStyleIdx = pr;
    }

    PrSet.prototype.setPresStyleLbl = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetPresStyleLbl, this.getPresStyleLbl(), pr));
      this.presStyleLbl = pr;
    }

    PrSet.prototype.setQsCatId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetQsCatId, this.getQsCatId(), pr));
      this.qsCatId = pr;
    }

    PrSet.prototype.setQsTypeId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetQsTypeId, this.getQsTypeId(), pr));
      this.qsTypeId = pr;
    }

    PrSet.prototype.setStyle = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PrSetStyle, this.getStyle(), oPr));
      this.style = oPr;
      this.setParentToChild(oPr);
    }

    PrSet.prototype.setPresLayoutVars = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PrSetPresLayoutVars, this.getPresLayoutVars(), oPr));
      this.presLayoutVars = oPr;
      this.setParentToChild(oPr);
    }

    PrSet.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setCoherent3DOff(this.getCoherent3DOff());
      oCopy.setCsCatId(this.getCsCatId());
      oCopy.setCsTypeId(this.getCsTypeId());
      oCopy.setCustAng(this.getCustAng());
      oCopy.setCustFlipHor(this.getCustFlipHor());
      oCopy.setCustFlipVert(this.getCustFlipVert());
      oCopy.setCustLinFactNeighborX(this.getCustLinFactNeighborX());
      oCopy.setCustLinFactNeighborY(this.getCustLinFactNeighborY());
      oCopy.setCustLinFactX(this.getCustLinFactX());
      oCopy.setCustLinFactY(this.getCustLinFactY());
      oCopy.setCustRadScaleInc(this.getCustRadScaleInc());
      oCopy.setCustRadScaleRad(this.getCustRadScaleRad());
      oCopy.setCustScaleX(this.getCustScaleX());
      oCopy.setCustScaleY(this.getCustScaleY());
      oCopy.setCustSzX(this.getCustSzX());
      oCopy.setCustSzY(this.getCustSzY());
      oCopy.setCustT(this.getCustT());
      oCopy.setLoCatId(this.getLoCatId());
      oCopy.setLoTypeId(this.getLoTypeId());
      oCopy.setPhldr(this.getPhldr());
      oCopy.setPhldrT(this.getPhldrT());
      oCopy.setPresAssocID(this.getPresAssocID());
      oCopy.setPresName(this.getPresName());
      oCopy.setPresStyleCnt(this.getPresStyleCnt());
      oCopy.setPresStyleIdx(this.getPresStyleIdx());
      oCopy.setPresStyleLbl(this.getPresStyleLbl());
      oCopy.setQsCatId(this.getQsCatId());
      oCopy.setQsTypeId(this.getQsTypeId());
      if (this.getStyle()) {
        oCopy.setStyle(this.style.createDuplicate(oIdMap));
      }
      if (this.getPresLayoutVars()) {
        oCopy.setPresLayoutVars(this.presLayoutVars.createDuplicate(oIdMap));
      }
    }

    PrSet.prototype.getCoherent3DOff = function () {
      return this.coherent3DOff;
    }

    PrSet.prototype.getCsCatId = function () {
      return this.csCatId;
    }

    PrSet.prototype.getCsTypeId = function () {
      return this.csTypeId;
    }

    PrSet.prototype.getCustAng = function () {
      return this.custAng;
    }

    PrSet.prototype.getCustFlipHor = function () {
      return this.custFlipHor;
    }

    PrSet.prototype.getCustFlipVert = function () {
      return this.custFlipVert;
    }

    PrSet.prototype.getCustLinFactNeighborX = function () {
      return this.custLinFactNeighborX;
    }

    PrSet.prototype.getCustLinFactNeighborY = function () {
      return this.custLinFactNeighborY;
    }

    PrSet.prototype.getCustLinFactX = function () {
      return this.custLinFactX;
    }

    PrSet.prototype.getCustLinFactY = function () {
      return this.custLinFactY;
    }

    PrSet.prototype.getCustRadScaleInc = function () {
      return this.custRadScaleInc;
    }

    PrSet.prototype.getCustRadScaleRad = function () {
      return this.custRadScaleRad;
    }

    PrSet.prototype.getCustScaleX = function () {
      return this.custScaleX;
    }

    PrSet.prototype.getCustScaleY = function () {
      return this.custScaleY;
    }

    PrSet.prototype.getCustSzX = function () {
      return this.custSzX;
    }

    PrSet.prototype.getCustSzY = function () {
      return this.custSzY;
    }

    PrSet.prototype.getCustT = function () {
      return this.custT;
    }

    PrSet.prototype.getLoCatId = function () {
      return this.loCatId;
    }

    PrSet.prototype.getLoTypeId = function () {
      return this.loTypeId;
    }

    PrSet.prototype.getPhldr = function () {
      return this.phldr;
    }

    PrSet.prototype.getPhldrT = function () {
      return this.phldrT;
    }

    PrSet.prototype.getPresAssocID = function () {
      return this.presAssocID;
    }

    PrSet.prototype.getPresName = function () {
      return this.presName;
    }

    PrSet.prototype.getPresStyleCnt = function () {
      return this.presStyleCnt;
    }

    PrSet.prototype.getPresStyleIdx = function () {
      return this.presStyleIdx;
    }

    PrSet.prototype.getPresStyleLbl = function () {
      return this.presStyleLbl;
    }

    PrSet.prototype.getQsCatId = function () {
      return this.qsCatId;
    }

    PrSet.prototype.getQsTypeId = function () {
      return this.qsTypeId;
    }

    PrSet.prototype.getStyle = function () {
      return this.style;
    }

    PrSet.prototype.getPresLayoutVars = function () {
      return this.presLayoutVars;
    }

    PrSet.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteBool2(1, this.coherent3DOff);
      pWriter._WriteString2(2, this.csCatId);
      pWriter._WriteString2(3, this.csTypeId);
      pWriter._WriteInt2(4, this.custAng ? (this.custAng / AscFormat.cToRad + 0.5) >> 0 : null);
      pWriter._WriteBool2(5, this.custFlipHor);
      pWriter._WriteBool2(6, this.custFlipVert);
      pWriter._WriteInt2(7, this.custLinFactNeighborX ? Math.floor(this.custLinFactNeighborX * 100000) : null);
      pWriter._WriteInt2(8, this.custLinFactNeighborY ? Math.floor(this.custLinFactNeighborY * 100000) : null);
      pWriter._WriteInt2(9, this.custLinFactX ? Math.floor(this.custLinFactX * 100000) : null);
      pWriter._WriteInt2(10, this.custLinFactY ? Math.floor(this.custLinFactY * 100000) : null);
      pWriter._WriteInt2(11, this.custRadScaleInc);
      pWriter._WriteInt2(12, this.custRadScaleRad);
      pWriter._WriteInt2(13, this.custScaleX ? Math.floor(this.custScaleX * 100000) : null);
      pWriter._WriteInt2(14, this.custScaleY ? Math.floor(this.custScaleY * 100000) : null);
      pWriter._WriteInt2(15, this.custSzX);
      pWriter._WriteInt2(16, this.custSzY);
      pWriter._WriteBool2(17, this.custT);
      pWriter._WriteString2(18, this.loCatId);
      pWriter._WriteString2(19, this.loTypeId);
      pWriter._WriteBool2(20, this.phldr);
      pWriter._WriteString2(21, this.phldrT);
      pWriter._WriteString2(22, this.presAssocID);
      pWriter._WriteString2(23, this.presName);
      pWriter._WriteInt2(24, this.presStyleCnt);
      pWriter._WriteInt2(25, this.presStyleIdx);
      pWriter._WriteString2(26, this.presStyleLbl);
      pWriter._WriteString2(27, this.qsCatId);
      pWriter._WriteString2(28, this.qsTypeId);
    };
    PrSet.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.presLayoutVars);
      pWriter.WriteRecord2(1, this.style, pWriter.WriteShapeStyle);
    };
    PrSet.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (1 === nType) this.setCoherent3DOff(oStream.GetBool());
      else if (2 === nType) this.setCsCatId(oStream.GetString2());
      else if (3 === nType) this.setCsTypeId(oStream.GetString2());
      else if (4 === nType) this.setCustAng(oStream.GetLong() * AscFormat.cToRad);
      else if (5 === nType) this.setCustFlipHor(oStream.GetBool());
      else if (6 === nType) this.setCustFlipVert(oStream.GetBool());
      else if (7 === nType) this.setCustLinFactNeighborX(oStream.GetLong() / 100000);
      else if (8 === nType) this.setCustLinFactNeighborY(oStream.GetLong() / 100000);
      else if (9 === nType) this.setCustLinFactX(oStream.GetLong() / 100000);
      else if (10 === nType) this.setCustLinFactY(oStream.GetLong() / 100000);
      else if (11 === nType) this.setCustRadScaleInc(oStream.GetLong());
      else if (12 === nType) this.setCustRadScaleRad(oStream.GetLong());
      else if (13 === nType) this.setCustScaleX(oStream.GetLong() / 100000);
      else if (14 === nType) this.setCustScaleY(oStream.GetLong() / 100000);
      else if (15 === nType) this.setCustSzX(oStream.GetLong());
      else if (16 === nType) this.setCustSzY(oStream.GetLong());
      else if (17 === nType) this.setCustT(oStream.GetBool());
      else if (18 === nType) this.setLoCatId(oStream.GetString2());
      else if (19 === nType) this.setLoTypeId(oStream.GetString2());
      else if (20 === nType) this.setPhldr(oStream.GetBool());
      else if (21 === nType) this.setPhldrT(oStream.GetString2());
      else if (22 === nType) this.setPresAssocID(oStream.GetString2());
      else if (23 === nType) this.setPresName(oStream.GetString2());
      else if (24 === nType) this.setPresStyleCnt(oStream.GetLong());
      else if (25 === nType) this.setPresStyleIdx(oStream.GetLong());
      else if (26 === nType) this.setPresStyleLbl(oStream.GetString2());
      else if (27 === nType) this.setQsCatId(oStream.GetString2());
      else if (28 === nType) this.setQsTypeId(oStream.GetString2());
    };
    PrSet.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setPresLayoutVars(new VarLst());
          this.presLayoutVars.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setStyle(pReader.ReadShapeStyle());
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    PrSet.prototype.getChildren = function() {
      return [this.presLayoutVars, this.style];
    }



    changesFactory[AscDFH.historyitem_LayoutDefDefStyle] = CChangeString;
    changesFactory[AscDFH.historyitem_LayoutDefMinVer] = CChangeString;
    changesFactory[AscDFH.historyitem_LayoutDefUniqueId] = CChangeString;
    changesFactory[AscDFH.historyitem_LayoutDefCatLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefClrData] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefTitle] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefLayoutNode] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefSampData] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefStyleData] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefDesc] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_LayoutDefDefStyle] = function (oClass, value) {
      oClass.defStyle = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefMinVer] = function (oClass, value) {
      oClass.minVer = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefUniqueId] = function (oClass, value) {
      oClass.uniqueId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefCatLst] = function (oClass, value) {
      oClass.catLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefClrData] = function (oClass, value) {
      oClass.clrData = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefLayoutNode] = function (oClass, value) {
      oClass.layoutNode = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefSampData] = function (oClass, value) {
      oClass.sampData = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefStyleData] = function (oClass, value) {
      oClass.styleData = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefTitle] = function (oClass, value) {
      oClass.title = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefDesc] = function (oClass, value) {
      oClass.desc = value;
    };

    function LayoutDef() {
      CBaseFormatObject.call(this);
      this.defStyle = null;
      this.minVer = null;
      this.uniqueId = null;
      this.catLst = null;
      this.clrData = null;
      this.desc = null;
      this.extLst = null;
      this.layoutNode = null;
      this.sampData = null;
      this.styleData = null;
      this.title = null;
    }

    InitClass(LayoutDef, CBaseFormatObject, AscDFH.historyitem_type_LayoutDef);

    LayoutDef.prototype.setDefStyle = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutDefDefStyle, this.getDefStyle(), pr));
      this.defStyle = pr;
    }

    LayoutDef.prototype.setMinVer = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutDefMinVer, this.getMinVer(), pr));
      this.minVer = pr;
    }

    LayoutDef.prototype.setUniqueId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutDefUniqueId, this.getUniqueId(), pr));
      this.uniqueId = pr;
    }

    LayoutDef.prototype.setCatLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefCatLst, this.getCatLst(), oPr));
      this.catLst = oPr;
      this.setParentToChild(oPr);
    }

    LayoutDef.prototype.setClrData = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefClrData, this.getClrData(), oPr));
      this.clrData = oPr;
      this.setParentToChild(oPr);
    }

    LayoutDef.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    LayoutDef.prototype.setLayoutNode = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefLayoutNode, this.getLayoutNode(), oPr));
      this.layoutNode = oPr;
      this.setParentToChild(oPr);
    }

    LayoutDef.prototype.setSampData = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefSampData, this.getSampData(), oPr));
      this.sampData = oPr;
      this.setParentToChild(oPr);
    }

    LayoutDef.prototype.setStyleData = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefStyleData, this.getStyleData(), oPr));
      this.styleData = oPr;
      this.setParentToChild(oPr);
    }

    LayoutDef.prototype.setTitle = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefTitle, this.getTitle(), oPr));
      this.title = oPr;
      this.setParentToChild(oPr);
    };

    LayoutDef.prototype.setDesc = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefDesc, this.getDesc(), oPr));
      this.desc = oPr;
      this.setParentToChild(oPr);
    };

    LayoutDef.prototype.getDefStyle = function () {
      return this.defStyle;
    }

    LayoutDef.prototype.getMinVer = function () {
      return this.minVer;
    }

    LayoutDef.prototype.getUniqueId = function () {
      return this.uniqueId;
    }

    LayoutDef.prototype.getCatLst = function () {
      return this.catLst;
    }

    LayoutDef.prototype.getClrData = function () {
      return this.clrData;
    }

    LayoutDef.prototype.getDesc = function () {
      return this.desc;
    }

    LayoutDef.prototype.getExtLst = function () {
      return this.extLst;
    }

    LayoutDef.prototype.getLayoutNode = function () {
      return this.layoutNode;
    }

    LayoutDef.prototype.getSampData = function () {
      return this.sampData;
    }

    LayoutDef.prototype.getStyleData = function () {
      return this.styleData;
    }

    LayoutDef.prototype.getTitle = function () {
      return this.title;
    }

    LayoutDef.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setDefStyle(this.getDefStyle());
      oCopy.setMinVer(this.getMinVer());
      oCopy.setUniqueId(this.getUniqueId());
      if (this.getCatLst()) {
        oCopy.setCatLst(this.getCatLst().createDuplicate(oIdMap));
      }
      if (this.getClrData()) {
        oCopy.setClrData(this.getClrData().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getLayoutNode()) {
        oCopy.setLayoutNode(this.getLayoutNode().createDuplicate(oIdMap));
      }
      if (this.getSampData()) {
        oCopy.setSampData(this.getSampData().createDuplicate(oIdMap));
      }
      if (this.getStyleData()) {
        oCopy.setStyleData(this.getStyleData().createDuplicate(oIdMap));
      }
      if (this.getTitle()) {
        oCopy.setTitle(this.getTitle().createDuplicate(oIdMap));
      }
      if (this.getDesc()) {
        oCopy.setDesc(this.getDesc().createDuplicate(oIdMap));
      }
    }

    LayoutDef.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.uniqueId);
      pWriter._WriteString2(1, this.minVer);
      pWriter._WriteString2(2, this.defStyle);
    };
    LayoutDef.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.title);
      this.writeRecord2(pWriter, 1, this.desc);
      this.writeRecord2(pWriter, 2, this.catLst);
      this.writeRecord2(pWriter, 3, this.sampData);
      this.writeRecord2(pWriter, 4, this.styleData);
      this.writeRecord2(pWriter, 5, this.clrData);
      this.writeRecord2(pWriter, 0xb5, this.layoutNode);
    };
    LayoutDef.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setUniqueId(oStream.GetString2());
      else if (1 === nType) this.setMinVer(oStream.GetString2());
      else if (2 === nType) this.setDefStyle(oStream.GetString2());
    };
    LayoutDef.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setTitle(new DiagramTitle());
          this.title.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setDesc(new Desc());
          this.desc.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setCatLst(new CatLst());
          this.catLst.fromPPTY(pReader);
          break;
        }
        case 3: {
          this.setSampData(new SampData());
          this.sampData.fromPPTY(pReader);
          break;
        }
        case 4: {
          this.setStyleData(new StyleData());
          this.styleData.fromPPTY(pReader);
          break;
        }
        case 5: {
          this.setClrData(new ClrData());
          this.clrData.fromPPTY(pReader);
          break;
        }
        case 0xb5: {
          this.setLayoutNode(new LayoutNode());
          this.layoutNode.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    LayoutDef.prototype.getChildren = function() {
      return [this.title, this.desc, this.catLst, this.sampData, this.styleData, this.clrData, this.layoutNode];
    };

    LayoutDef.prototype.startAlgorithm = function (pointTree) {
      var entry = this.getLayoutNode();
      if (entry) {
        entry.startAlgorithm(pointTree);
      }
    }


    function CatLst() {
      CCommonDataList.call(this);
    }

    InitClass(CatLst, CCommonDataList, AscDFH.historyitem_type_CatLst);

    CatLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          var oChild = new SCat();
          oChild.fromPPTY(pReader);
          this.addToLst(this.list.length, oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    changesFactory[AscDFH.historyitem_SCatPri] = CChangeLong;
    changesFactory[AscDFH.historyitem_SCatType] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_SCatPri] = function (oClass, value) {
      oClass.pri = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SCatType] = function (oClass, value) {
      oClass.type = value;
    };

    function SCat() {
      CBaseFormatObject.call(this);
      this.pri = null;
      this.type = null;
    }

    InitClass(SCat, CBaseFormatObject, AscDFH.historyitem_type_SCat);

    SCat.prototype.setPri = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_SCatPri, this.getPri(), pr));
      this.pri = pr;
    }

    SCat.prototype.setType = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_SCatType, this.getType(), pr));
      this.type = pr;
    }

    SCat.prototype.getPri = function () {
      return this.pri;
    }

    SCat.prototype.getType = function () {
      return this.type;
    }

    SCat.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setPri(this.getPri());
      oCopy.setType(this.getType());
    }

    SCat.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.type);
      pWriter._WriteUInt2(1, this.pri);
    };
    SCat.prototype.writeChildren = function(pWriter) {
    };
    SCat.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setType(oStream.GetString2());
      else if (1 === nType) this.setPri(oStream.GetULong());
    };
    SCat.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_ClrDataUseDef] = CChangeBool;
    changesFactory[AscDFH.historyitem_ClrDataDataModel] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_ClrDataUseDef] = function (oClass, value) {
      oClass.useDef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ClrDataDataModel] = function (oClass, value) {
      oClass.dataModel = value;
    };

    function ClrData() {
      CBaseFormatObject.call(this);
      this.useDef = null;
      this.dataModel = null;
    }

    InitClass(ClrData, CBaseFormatObject, AscDFH.historyitem_type_ClrData);

    ClrData.prototype.setUseDef = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_ClrDataUseDef, this.getUseDef(), pr));
      this.useDef = pr;
    }

    ClrData.prototype.setDataModel = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ClrDataDataModel, this.getDataModel(), oPr));
      this.dataModel = oPr;
      this.setParentToChild(oPr);
    }

    ClrData.prototype.getUseDef = function () {
      return this.useDef;
    }

    ClrData.prototype.getDataModel = function () {
      return this.dataModel;
    }

    ClrData.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setUseDef(this.getUseDef());
      if (this.getDataModel()) {
        oCopy.setDataModel(this.getDataModel().createDuplicate(oIdMap));
      }
    }

    ClrData.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteBool2(0, this.useDef);
    };
    ClrData.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.dataModel); // TODO: add record number
    };
    ClrData.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setUseDef(oStream.GetBool());
    };
    ClrData.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setDataModel(new DataModel());
          this.dataModel.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    ClrData.prototype.getChildren = function() {
      return [this.dataModel];
    };


    changesFactory[AscDFH.historyitem_DescLang] = CChangeString;
    changesFactory[AscDFH.historyitem_DescVal] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_DescLang] = function (oClass, value) {
      oClass.lang = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DescVal] = function (oClass, value) {
      oClass.val = value;
    };

    function Desc() {
      CBaseFormatObject.call(this);
      this.lang = null;
      this.val = null;
    }

    InitClass(Desc, CBaseFormatObject, AscDFH.historyitem_type_Desc);

    Desc.prototype.setLang = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_DescLang, this.getLang(), pr));
      this.lang = pr;
    }

    Desc.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_DescVal, this.getVal(), pr));
      this.val = pr;
    }

    Desc.prototype.getLang = function () {
      return this.lang;
    }

    Desc.prototype.getVal = function () {
      return this.val;
    }

    Desc.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setLang(this.getLang());
      oCopy.setVal(this.getVal());
    }

    Desc.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.lang);
      pWriter._WriteString2(1, this.val);
    };
    Desc.prototype.writeChildren = function(pWriter) {
    };
    Desc.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setLang(oStream.GetString2());
      else if (1 === nType) this.setVal(oStream.GetString2());
    };
    Desc.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_LayoutNodeChOrder] = CChangeLong;
    changesFactory[AscDFH.historyitem_LayoutNodeMoveWith] = CChangeString;
    changesFactory[AscDFH.historyitem_LayoutNodeName] = CChangeString;
    changesFactory[AscDFH.historyitem_LayoutNodeStyleLbl] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_LayoutNodeChOrder] = function (oClass, value) {
      oClass.chOrder = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutNodeMoveWith] = function (oClass, value) {
      oClass.moveWith = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutNodeName] = function (oClass, value) {
      oClass.name = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutNodeStyleLbl] = function (oClass, value) {
      oClass.styleLbl = value;
    };

    function LayoutNode() {
      CCommonDataList.call(this);
      this.chOrder = null;
      this.moveWith = null;
      this.name = null;
      this.styleLbl = null;
    }

    InitClass(LayoutNode, CCommonDataList, AscDFH.historyitem_type_LayoutNode);

    LayoutNode.prototype.setChOrder = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_LayoutNodeChOrder, this.getChOrder(), pr));
      this.chOrder = pr;
    }

    LayoutNode.prototype.setMoveWith = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutNodeMoveWith, this.getMoveWith(), pr));
      this.moveWith = pr;
    }

    LayoutNode.prototype.setName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutNodeName, this.getName(), pr));
      this.name = pr;
    }

    LayoutNode.prototype.setStyleLbl = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutNodeStyleLbl, this.getStyleLbl(), pr));
      this.styleLbl = pr;
    }

    LayoutNode.prototype.getChOrder = function () {
      return this.chOrder;
    }

    LayoutNode.prototype.getMoveWith = function () {
      return this.moveWith;
    }

    LayoutNode.prototype.getName = function () {
      return this.name;
    }

    LayoutNode.prototype.getStyleLbl = function () {
      return this.styleLbl;
    }

    LayoutNode.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setChOrder(this.getChOrder());
      oCopy.setMoveWith(this.getMoveWith());
      oCopy.setName(this.getName());
      oCopy.setStyleLbl(this.getStyleLbl());
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        oCopy.addToLst(nIdx, this.list[nIdx].createDuplicate(oIdMap));
      }
    }
    LayoutNode.prototype.readElement = function(pReader, nType) {
      var oElement = null;
      switch(nType) {
        case 0xb1: oElement = new Alg(); break;
        case 0xb2: oElement = new Choose(); break;
        case 0xb3: oElement = new ConstrLst(); break;
        case 0xb4: oElement = new ForEach(); break;
        case 0xb5: oElement = new LayoutNode(); break;
        case 0xb6: oElement = new PresOf(); break;
        case 0xb7: oElement = new RuleLst(); break;
        case 0xb8: oElement = new SShape(); break;
        case 0xb9: oElement = new VarLst(); break;
        default:break;
      }
      if(oElement) {
        oElement.fromPPTY(pReader);
        this.addToLst(this.list.length, oElement);
      }
    };
    LayoutNode.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
      pWriter._WriteString2(1, this.styleLbl);
      pWriter._WriteString2(2, this.moveWith);
    };
    LayoutNode.prototype.writeChildren = function(pWriter) {
      for(var nIndex = 0; nIndex < this.list.length; ++nIndex) {
        var oElement = this.list[nIndex];
        switch (oElement.getObjectType()) {
          case AscDFH.historyitem_type_Alg: this.writeRecord2(pWriter, 0xb1, oElement); break;
          case AscDFH.historyitem_type_Choose: this.writeRecord2(pWriter, 0xb2, oElement); break;
          case AscDFH.historyitem_type_ConstrLst: this.writeRecord2(pWriter, 0xb3, oElement); break;
          case AscDFH.historyitem_type_ForEach: this.writeRecord2(pWriter, 0xb4, oElement); break;
          case AscDFH.historyitem_type_LayoutNode: this.writeRecord2(pWriter, 0xb5, oElement); break;
          case AscDFH.historyitem_type_PresOf: this.writeRecord2(pWriter, 0xb6, oElement); break;
          case AscDFH.historyitem_type_RuleLst: this.writeRecord2(pWriter, 0xb7, oElement); break;
          case AscDFH.historyitem_type_SShape: this.writeRecord2(pWriter, 0xb8, oElement); break;
          case AscDFH.historyitem_type_VarLst: this.writeRecord2(pWriter, 0xb9, oElement); break;
          default: break;
        }
      }
    };
    LayoutNode.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
      else if (1 === nType) this.setStyleLbl(oStream.GetString2());
      else if (2 === nType) this.setMoveWith(oStream.GetString2());

    };
    LayoutNode.prototype.readChild = function(nType, pReader) {
      this.readElement(pReader, nType);
    };
    LayoutNode.prototype.getChildren = function() {
      return [].concat(this.list);
    };


    LayoutNode.prototype.getConstrLst = function () {
      return this.list.reduce(function (save, next) {
        return next instanceof ConstrLst ? next : save;
      }, undefined);
    }

    LayoutNode.prototype.startAlgorithm = function (pointTree) {
      if (pointTree) {
        var nodes = pointTree.findNodeByNameAndStyleLbl.call(pointTree, this.name, this.styleLbl);
        if (nodes) {
          var that = this;
          nodes.forEach(function (node) {
            var transfer = {node: node, name: that.name, styleLbl: that.styleLbl};
            var constrLst = that.list.reduce(function (acc, current) {
              return current instanceof ConstrLst ? current : acc;
            }, null);
            if (constrLst) {
              constrLst.startSetConstr(pointTree, transfer);
            }
            if (that.list) {
              that.list.forEach(function (element) {
                if (element instanceof AscFormat.ForEach || element instanceof AscFormat.LayoutNode || element instanceof AscFormat.Choose) {
                  element.startAlgorithm(pointTree, transfer);
                }
              });
            }
          });
        }
      }
    }

    LayoutNode.prototype.findPoint = function (pointInfo) {

    }


    changesFactory[AscDFH.historyitem_AlgRev] = CChangeLong;
    changesFactory[AscDFH.historyitem_AlgType] = CChangeLong;
    changesFactory[AscDFH.historyitem_AlgExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_AlgAddParam] = CChangeContent;
    changesFactory[AscDFH.historyitem_AlgRemoveParam] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_AlgRev] = function (oClass, value) {
      oClass.rev = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AlgType] = function (oClass, value) {
      oClass.type = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AlgExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingContentChanges[AscDFH.historyitem_AlgAddParam] = function (oClass) {
      return oClass.param;
    };
    drawingContentChanges[AscDFH.historyitem_AlgRemoveParam] = function (oClass) {
      return oClass.param;
    };

    function Alg() {
      CBaseFormatObject.call(this);
      this.rev = null;
      this.type = null;
      this.extLst = null;
      this.param = [];
    }

    InitClass(Alg, CBaseFormatObject, AscDFH.historyitem_type_Alg);

    Alg.prototype.setRev = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AlgRev, this.getRev(), pr));
      this.rev = pr;
    }

    Alg.prototype.setType = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AlgType, this.getType(), pr));
      this.type = pr;
    }

    Alg.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_AlgExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    Alg.prototype.addToLstParam = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.param.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_AlgAddParam, nInsertIdx, [oPr], true));
      this.param.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    Alg.prototype.removeFromLstParam = function (nIdx) {
      if (nIdx > -1 && nIdx < this.param.length) {
        this.param[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_AlgRemoveParam, nIdx, [this.param[nIdx]], false));
        this.param.splice(nIdx, 1);
      }
    };

    Alg.prototype.getRev = function () {
      return this.rev;
    }

    Alg.prototype.getType = function () {
      return this.type;
    }

    Alg.prototype.getExtLst = function () {
      return this.extLst;
    }

    Alg.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setRev(this.getRev());
      oCopy.setType(this.getType());
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.param.length; ++nIdx) {
        oCopy.addToLstParam(nIdx, this.param[nIdx].createDuplicate(oIdMap));
      }
    }

    Alg.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUInt2(0, this.rev);
      pWriter._WriteUChar2(1, this.type);
    };
    Alg.prototype.writeChildren = function(pWriter) {
      for (var i = 0;i < this.param.length; i += 1) {
        this.writeRecord2(pWriter,0, this.param[i]);
      }
    };
    Alg.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setRev(oStream.GetULong());
      else if (1 === nType) this.setType(oStream.GetUChar());
    };
    Alg.prototype.readChild = function(nType, pReader) {
      switch (nType) {
        case 0: {
          var oChild = new Param();
          oChild.fromPPTY(pReader);
          this.addToLstParam(this.param.length, oChild);
          break;
        }
        default:
          pReader.SkipRecord();
          break;
      }
    };


    changesFactory[AscDFH.historyitem_ParameterValArrowheadStyle] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValAutoTextRotation] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValBendPoint] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValBreakpoint] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValCenterShapeMapping] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValChildAlignment] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValChildDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValConnectorDimension] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValConnectorPoint] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValConnectorRouting] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValContinueDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValDiagramHorizontalAlignment] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValDiagramTextAlignment] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValFallbackDimension] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValFlowDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValGrowDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValHierarchyAlignment] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValLinearDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValNodeHorizontalAlignment] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValNodeVerticalAlignment] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValOffset] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValPyramidAccentPosition] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValPyramidAccentTextMargin] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValRotationPath] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValSecondaryChildAlignment] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValSecondaryLinearDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValStartingElement] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValTextAnchorHorizontal] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValTextAnchorVertical] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValTextBlockDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValTextDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValVerticalAlignment] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValBool] = CChangeBool;
    changesFactory[AscDFH.historyitem_ParameterValDouble] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_ParameterValInt] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParameterValStr] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_ParameterValArrowheadStyle] = function (oClass, value) {
      oClass.arrowheadStyle = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValAutoTextRotation] = function (oClass, value) {
      oClass.autoTextRotation = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValBendPoint] = function (oClass, value) {
      oClass.bendPoint = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValBreakpoint] = function (oClass, value) {
      oClass.breakpoint = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValCenterShapeMapping] = function (oClass, value) {
      oClass.centerShapeMapping = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValChildAlignment] = function (oClass, value) {
      oClass.childAlignment = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValChildDirection] = function (oClass, value) {
      oClass.childDirection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValConnectorDimension] = function (oClass, value) {
      oClass.connectorDimension = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValConnectorPoint] = function (oClass, value) {
      oClass.connectorPoint = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValConnectorRouting] = function (oClass, value) {
      oClass.connectorRouting = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValContinueDirection] = function (oClass, value) {
      oClass.continueDirection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValDiagramHorizontalAlignment] = function (oClass, value) {
      oClass.diagramHorizontalAlignment = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValDiagramTextAlignment] = function (oClass, value) {
      oClass.diagramTextAlignment = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValFallbackDimension] = function (oClass, value) {
      oClass.fallbackDimension = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValFlowDirection] = function (oClass, value) {
      oClass.flowDirection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValGrowDirection] = function (oClass, value) {
      oClass.growDirection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValHierarchyAlignment] = function (oClass, value) {
      oClass.hierarchyAlignment = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValLinearDirection] = function (oClass, value) {
      oClass.linearDirection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValNodeHorizontalAlignment] = function (oClass, value) {
      oClass.nodeHorizontalAlignment = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValNodeVerticalAlignment] = function (oClass, value) {
      oClass.nodeVerticalAlignment = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValOffset] = function (oClass, value) {
      oClass.offset = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValPyramidAccentPosition] = function (oClass, value) {
      oClass.pyramidAccentPosition = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValPyramidAccentTextMargin] = function (oClass, value) {
      oClass.pyramidAccentTextMargin = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValRotationPath] = function (oClass, value) {
      oClass.rotationPath = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValSecondaryChildAlignment] = function (oClass, value) {
      oClass.secondaryChildAlignment = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValSecondaryLinearDirection] = function (oClass, value) {
      oClass.secondaryLinearDirection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValStartingElement] = function (oClass, value) {
      oClass.startingElement = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValTextAnchorHorizontal] = function (oClass, value) {
      oClass.textAnchorHorizontal = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValTextAnchorVertical] = function (oClass, value) {
      oClass.textAnchorVertical = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValTextBlockDirection] = function (oClass, value) {
      oClass.textBlockDirection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValTextDirection] = function (oClass, value) {
      oClass.textDirection = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValVerticalAlignment] = function (oClass, value) {
      oClass.verticalAlignment = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValBool] = function (oClass, value) {
      oClass.bool = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValDouble] = function (oClass, value) {
      oClass.double = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValInt] = function (oClass, value) {
      oClass.int = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParameterValStr] = function (oClass, value) {
      oClass.str = value;
    };

    function ParameterVal() {
      CBaseFormatObject.call(this);
      this.arrowheadStyle = null;
      this.autoTextRotation = null;
      this.bendPoint = null;
      this.breakpoint = null;
      this.centerShapeMapping = null;
      this.childAlignment = null;
      this.childDirection = null;
      this.connectorDimension = null;
      this.connectorPoint = null;
      this.connectorRouting = null;
      this.continueDirection = null;
      this.diagramHorizontalAlignment = null;
      this.diagramTextAlignment = null;
      this.fallbackDimension = null;
      this.flowDirection = null;
      this.growDirection = null;
      this.hierarchyAlignment = null;
      this.linearDirection = null;
      this.nodeHorizontalAlignment = null;
      this.nodeVerticalAlignment = null;
      this.offset = null;
      this.pyramidAccentPosition = null;
      this.pyramidAccentTextMargin = null;
      this.rotationPath = null;
      this.secondaryChildAlignment = null;
      this.secondaryLinearDirection = null;
      this.startingElement = null;
      this.textAnchorHorizontal = null;
      this.textAnchorVertical = null;
      this.textBlockDirection = null;
      this.textDirection = null;
      this.verticalAlignment = null;
      this.bool = null;
      this.double = null;
      this.int = null;
      this.str = null;
    }

    InitClass(ParameterVal, CBaseFormatObject, AscDFH.historyitem_type_ParameterVal);

    ParameterVal.prototype.setArrowheadStyle = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValArrowheadStyle, this.getArrowheadStyle(), pr));
      this.arrowheadStyle = pr;
    }

    ParameterVal.prototype.setAutoTextRotation = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValAutoTextRotation, this.getAutoTextRotation(), pr));
      this.autoTextRotation = pr;
    }

    ParameterVal.prototype.setBendPoint = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValBendPoint, this.getBendPoint(), pr));
      this.bendPoint = pr;
    }

    ParameterVal.prototype.setBreakpoint = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValBreakpoint, this.getBreakpoint(), pr));
      this.breakpoint = pr;
    }

    ParameterVal.prototype.setCenterShapeMapping = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValCenterShapeMapping, this.getCenterShapeMapping(), pr));
      this.centerShapeMapping = pr;
    }

    ParameterVal.prototype.setChildAlignment = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValChildAlignment, this.getChildAlignment(), pr));
      this.childAlignment = pr;
    }

    ParameterVal.prototype.setChildDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValChildDirection, this.getChildDirection(), pr));
      this.childDirection = pr;
    }

    ParameterVal.prototype.setConnectorDimension = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValConnectorDimension, this.getConnectorDimension(), pr));
      this.connectorDimension = pr;
    }

    ParameterVal.prototype.setConnectorPoint = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValConnectorPoint, this.getConnectorPoint(), pr));
      this.connectorPoint = pr;
    }

    ParameterVal.prototype.setConnectorRouting = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValConnectorRouting, this.getConnectorRouting(), pr));
      this.connectorRouting = pr;
    }

    ParameterVal.prototype.setContinueDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValContinueDirection, this.getContinueDirection(), pr));
      this.continueDirection = pr;
    }

    ParameterVal.prototype.setDiagramHorizontalAlignment = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValDiagramHorizontalAlignment, this.getDiagramHorizontalAlignment(), pr));
      this.diagramHorizontalAlignment = pr;
    }

    ParameterVal.prototype.setDiagramTextAlignment = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValDiagramTextAlignment, this.getDiagramTextAlignment(), pr));
      this.diagramTextAlignment = pr;
    }

    ParameterVal.prototype.setFallbackDimension = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValFallbackDimension, this.getFallbackDimension(), pr));
      this.fallbackDimension = pr;
    }

    ParameterVal.prototype.setFlowDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValFlowDirection, this.getFlowDirection(), pr));
      this.flowDirection = pr;
    }

    ParameterVal.prototype.setGrowDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValGrowDirection, this.getGrowDirection(), pr));
      this.growDirection = pr;
    }

    ParameterVal.prototype.setHierarchyAlignment = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValHierarchyAlignment, this.getHierarchyAlignment(), pr));
      this.hierarchyAlignment = pr;
    }

    ParameterVal.prototype.setLinearDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValLinearDirection, this.getLinearDirection(), pr));
      this.linearDirection = pr;
    }

    ParameterVal.prototype.setNodeHorizontalAlignment = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValNodeHorizontalAlignment, this.getNodeHorizontalAlignment(), pr));
      this.nodeHorizontalAlignment = pr;
    }

    ParameterVal.prototype.setNodeVerticalAlignment = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValNodeVerticalAlignment, this.getNodeVerticalAlignment(), pr));
      this.nodeVerticalAlignment = pr;
    }

    ParameterVal.prototype.setOffset = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValOffset, this.getOffset(), pr));
      this.offset = pr;
    }

    ParameterVal.prototype.setPyramidAccentPosition = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValPyramidAccentPosition, this.getPyramidAccentPosition(), pr));
      this.pyramidAccentPosition = pr;
    }

    ParameterVal.prototype.setPyramidAccentTextMargin = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValPyramidAccentTextMargin, this.getPyramidAccentTextMargin(), pr));
      this.pyramidAccentTextMargin = pr;
    }

    ParameterVal.prototype.setRotationPath = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValRotationPath, this.getRotationPath(), pr));
      this.rotationPath = pr;
    }

    ParameterVal.prototype.setSecondaryChildAlignment = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValSecondaryChildAlignment, this.getSecondaryChildAlignment(), pr));
      this.secondaryChildAlignment = pr;
    }

    ParameterVal.prototype.setSecondaryLinearDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValSecondaryLinearDirection, this.getSecondaryLinearDirection(), pr));
      this.secondaryLinearDirection = pr;
    }

    ParameterVal.prototype.setStartingElement = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValStartingElement, this.getStartingElement(), pr));
      this.startingElement = pr;
    }

    ParameterVal.prototype.setTextAnchorHorizontal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValTextAnchorHorizontal, this.getTextAnchorHorizontal(), pr));
      this.textAnchorHorizontal = pr;
    }

    ParameterVal.prototype.setTextAnchorVertical = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValTextAnchorVertical, this.getTextAnchorVertical(), pr));
      this.textAnchorVertical = pr;
    }

    ParameterVal.prototype.setTextBlockDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValTextBlockDirection, this.getTextBlockDirection(), pr));
      this.textBlockDirection = pr;
    }

    ParameterVal.prototype.setTextDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValTextDirection, this.getTextDirection(), pr));
      this.textDirection = pr;
    }

    ParameterVal.prototype.setVerticalAlignment = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValVerticalAlignment, this.getVerticalAlignment(), pr));
      this.verticalAlignment = pr;
    }

    ParameterVal.prototype.setBool = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_ParameterValBool, this.getBool(), pr));
      this.bool = pr;
    }

    ParameterVal.prototype.setDouble = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_ParameterValDouble, this.getDouble(), pr));
      this.double = pr;
    }

    ParameterVal.prototype.setInt = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParameterValInt, this.getInt(), pr));
      this.int = pr;
    }

    ParameterVal.prototype.setStr = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ParameterValStr, this.getStr(), pr));
      this.str = pr;
    }

    ParameterVal.prototype.getArrowheadStyle = function () {
      return this.arrowheadStyle;
    }

    ParameterVal.prototype.getAutoTextRotation = function () {
      return this.autoTextRotation;
    }

    ParameterVal.prototype.getBendPoint = function () {
      return this.bendPoint;
    }

    ParameterVal.prototype.getBreakpoint = function () {
      return this.breakpoint;
    }

    ParameterVal.prototype.getCenterShapeMapping = function () {
      return this.centerShapeMapping;
    }

    ParameterVal.prototype.getChildAlignment = function () {
      return this.childAlignment;
    }

    ParameterVal.prototype.getChildDirection = function () {
      return this.childDirection;
    }

    ParameterVal.prototype.getConnectorDimension = function () {
      return this.connectorDimension;
    }

    ParameterVal.prototype.getConnectorPoint = function () {
      return this.connectorPoint;
    }

    ParameterVal.prototype.getConnectorRouting = function () {
      return this.connectorRouting;
    }

    ParameterVal.prototype.getContinueDirection = function () {
      return this.continueDirection;
    }

    ParameterVal.prototype.getDiagramHorizontalAlignment = function () {
      return this.diagramHorizontalAlignment;
    }

    ParameterVal.prototype.getDiagramTextAlignment = function () {
      return this.diagramTextAlignment;
    }

    ParameterVal.prototype.getFallbackDimension = function () {
      return this.fallbackDimension;
    }

    ParameterVal.prototype.getFlowDirection = function () {
      return this.flowDirection;
    }

    ParameterVal.prototype.getGrowDirection = function () {
      return this.growDirection;
    }

    ParameterVal.prototype.getHierarchyAlignment = function () {
      return this.hierarchyAlignment;
    }

    ParameterVal.prototype.getLinearDirection = function () {
      return this.linearDirection;
    }

    ParameterVal.prototype.getNodeHorizontalAlignment = function () {
      return this.nodeHorizontalAlignment;
    }

    ParameterVal.prototype.getNodeVerticalAlignment = function () {
      return this.nodeVerticalAlignment;
    }

    ParameterVal.prototype.getOffset = function () {
      return this.offset;
    }

    ParameterVal.prototype.getPyramidAccentPosition = function () {
      return this.pyramidAccentPosition;
    }

    ParameterVal.prototype.getPyramidAccentTextMargin = function () {
      return this.pyramidAccentTextMargin;
    }

    ParameterVal.prototype.getRotationPath = function () {
      return this.rotationPath;
    }

    ParameterVal.prototype.getSecondaryChildAlignment = function () {
      return this.secondaryChildAlignment;
    }

    ParameterVal.prototype.getSecondaryLinearDirection = function () {
      return this.secondaryLinearDirection;
    }

    ParameterVal.prototype.getStartingElement = function () {
      return this.startingElement;
    }

    ParameterVal.prototype.getTextAnchorHorizontal = function () {
      return this.textAnchorHorizontal;
    }

    ParameterVal.prototype.getTextAnchorVertical = function () {
      return this.textAnchorVertical;
    }

    ParameterVal.prototype.getTextBlockDirection = function () {
      return this.textBlockDirection;
    }

    ParameterVal.prototype.getTextDirection = function () {
      return this.textDirection;
    }

    ParameterVal.prototype.getVerticalAlignment = function () {
      return this.verticalAlignment;
    }

    ParameterVal.prototype.getBool = function () {
      return this.bool;
    }

    ParameterVal.prototype.getDouble = function () {
      return this.double;
    }

    ParameterVal.prototype.getInt = function () {
      return this.int;
    }

    ParameterVal.prototype.getStr = function () {
      return this.str;
    }

    ParameterVal.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setArrowheadStyle(this.getArrowheadStyle());
      oCopy.setAutoTextRotation(this.getAutoTextRotation());
      oCopy.setBendPoint(this.getBendPoint());
      oCopy.setBreakpoint(this.getBreakpoint());
      oCopy.setCenterShapeMapping(this.getCenterShapeMapping());
      oCopy.setChildAlignment(this.getChildAlignment());
      oCopy.setChildDirection(this.getChildDirection());
      oCopy.setConnectorDimension(this.getConnectorDimension());
      oCopy.setConnectorPoint(this.getConnectorPoint());
      oCopy.setConnectorRouting(this.getConnectorRouting());
      oCopy.setContinueDirection(this.getContinueDirection());
      oCopy.setDiagramHorizontalAlignment(this.getDiagramHorizontalAlignment());
      oCopy.setDiagramTextAlignment(this.getDiagramTextAlignment());
      oCopy.setFallbackDimension(this.getFallbackDimension());
      oCopy.setFlowDirection(this.getFlowDirection());
      oCopy.setGrowDirection(this.getGrowDirection());
      oCopy.setHierarchyAlignment(this.getHierarchyAlignment());
      oCopy.setLinearDirection(this.getLinearDirection());
      oCopy.setNodeHorizontalAlignment(this.getNodeHorizontalAlignment());
      oCopy.setNodeVerticalAlignment(this.getNodeVerticalAlignment());
      oCopy.setOffset(this.getOffset());
      oCopy.setPyramidAccentPosition(this.getPyramidAccentPosition());
      oCopy.setPyramidAccentTextMargin(this.getPyramidAccentTextMargin());
      oCopy.setRotationPath(this.getRotationPath());
      oCopy.setSecondaryChildAlignment(this.getSecondaryChildAlignment());
      oCopy.setSecondaryLinearDirection(this.getSecondaryLinearDirection());
      oCopy.setStartingElement(this.getStartingElement());
      oCopy.setTextAnchorHorizontal(this.getTextAnchorHorizontal());
      oCopy.setTextAnchorVertical(this.getTextAnchorVertical());
      oCopy.setTextBlockDirection(this.getTextBlockDirection());
      oCopy.setTextDirection(this.getTextDirection());
      oCopy.setVerticalAlignment(this.getVerticalAlignment());
      oCopy.setBool(this.getBool());
      oCopy.setDouble(this.getDouble());
      oCopy.setInt(this.getInt());
      oCopy.setStr(this.getStr());
    }


    changesFactory[AscDFH.historyitem_ParamType] = CChangeLong;
    changesFactory[AscDFH.historyitem_ParamVal] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_ParamType] = function (oClass, value) {
      oClass.type = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ParamVal] = function (oClass, value) {
      oClass.val = value;
    };

    function Param() {
      CBaseFormatObject.call(this);
      this.type = null;
      this.val = null;
    }

    InitClass(Param, CBaseFormatObject, AscDFH.historyitem_type_Param);

    Param.prototype.setType = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ParamType, this.getType(), pr));
      this.type = pr;
    }

    Param.prototype.setVal = function (oPr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ParamVal, this.getVal(), oPr));
      this.val = oPr;
    }

    Param.prototype.getType = function () {
      return this.type;
    }

    Param.prototype.getVal = function () {
      return this.val;
    }

    Param.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setType(this.getType());
      oCopy.setVal(this.getVal());
    }

    Param.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.val);
      pWriter._WriteUChar2(1, this.type);
    };
    Param.prototype.writeChildren = function(pWriter) {
    };
    Param.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setVal(oStream.GetString2());
      else if (1 === nType) this.setType(oStream.GetUChar());
    };
    Param.prototype.readChild = function(nType, pReader) {
    };



    changesFactory[AscDFH.historyitem_ChooseName] = CChangeString;
    changesFactory[AscDFH.historyitem_ChooseElse] = CChangeObject;
    changesFactory[AscDFH.historyitem_ChooseAddToLstIf] = CChangeContent;
    changesFactory[AscDFH.historyitem_ChooseRemoveFromLstIf] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_ChooseName] = function (oClass, value) {
      oClass.name = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ChooseElse] = function (oClass, value) {
      oClass.else = value;
    };
    drawingContentChanges[AscDFH.historyitem_ChooseAddToLstIf] = function (oClass) {
      return oClass.if;
    };
    drawingContentChanges[AscDFH.historyitem_ChooseRemoveFromLstIf] = function (oClass) {
      return oClass.if;
    };

    function Choose() {
      CBaseFormatObject.call(this);
      this.name = null;
      this.else = null;
      this.if = [];
    }

    InitClass(Choose, CBaseFormatObject, AscDFH.historyitem_type_Choose);

    Choose.prototype.setName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ChooseName, this.getName(), pr));
      this.name = pr;
    }

    Choose.prototype.setElse = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ChooseElse, this.getElse(), oPr));
      this.else = oPr;
      this.setParentToChild(oPr);
    }

    Choose.prototype.addToLstIf = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.if.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ChooseAddToLstIf, nInsertIdx, [oPr], true));
      this.if.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    }

    Choose.prototype.removeFromLstIf = function (nIdx) {
      if (nIdx > -1 && nIdx < this.if.length) {
        this.if[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ChooseRemoveFromLstIf, nIdx, [this.if[nIdx]], false));
        this.if.splice(nIdx, 1);
      }
    }

    Choose.prototype.getName = function () {
      return this.name;
    }

    Choose.prototype.getElse = function () {
      return this.else;
    }

    Choose.prototype.getIf = function () {
      return this.if;
    }

    Choose.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setName(this.getName());
      if (this.getElse()) {
        oCopy.setElse(this.getElse().createDuplicate(oIdMap));
      }
      for (var i = 0; i < this.if.length; i += 1) {
        oCopy.addToLstIf(i, this.if[i].createDuplicate(oIdMap));
      }
    }

    Choose.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
    };
    Choose.prototype.writeChildren = function(pWriter) {
      for (var i = 0; i < this.if.length; i += 1) {
        this.writeRecord2(pWriter, 0, this.if[i]);
      }
      this.writeRecord2(pWriter, 1, this.else);
    };
    Choose.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
    };
    Choose.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.addToLstIf(0, new If());
          this.if[0].fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setElse(new Else());
          this.else.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    Choose.prototype.getChildren = function() {
      return [this.if, this.else];
    };

    Choose.prototype.startAlgorithm = function (pointTree, node) {
      var check;
      for (var i = 0; i < this.if.length; i += 1) {
        if (!check) {
          check = this.if[i].startAlgorithm(pointTree, node);
        } else {
          this.if[i].startAlgorithm(pointTree, node);
        }
      }
      if (!check) {
        if (this.else) {
          this.else.startAlgorithm(pointTree, node);
        }
      }
    }



    changesFactory[AscDFH.historyitem_ElseName] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_ElseName] = function (oClass, value) {
      oClass.name = value;
    };

    function Else() {
      CCommonDataList.call(this);
      this.name = null;
    }

    InitClass(Else, CCommonDataList, AscDFH.historyitem_type_Else);

    Else.prototype.setName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ElseName, this.getName(), pr));
      this.name = pr;
    }

    Else.prototype.getName = function () {
      return this.name;
    }

    Else.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setName(this.getName());
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        oCopy.addToLst(nIdx, this.list[nIdx].createDuplicate(oIdMap));
      }
    }

    Else.prototype.readElement = function(pReader, nType) {
      var oElement = null;
      switch(nType) {
        case 0xb1: oElement = new Alg(); break;
        case 0xb2: oElement = new Choose(); break;
        case 0xb3: oElement = new ConstrLst(); break;
        case 0xb4: oElement = new ForEach(); break;
        case 0xb5: oElement = new LayoutNode(); break;
        case 0xb6: oElement = new PresOf(); break;
        case 0xb7: oElement = new RuleLst(); break;
        case 0xb8: oElement = new SShape(); break;
        case 0xb9: oElement = new VarLst(); break;
        default:break;
      }
      if(oElement) {
        oElement.fromPPTY(pReader);
        this.addToLst(this.list.length, oElement);
      }
    };

    Else.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
    };

    Else.prototype.writeChildren = function(pWriter) {
      for(var nIndex = 0; nIndex < this.list.length; ++nIndex) {
        var oElement = this.list[nIndex];
        switch (oElement.getObjectType()) {
          case AscDFH.historyitem_type_Alg: this.writeRecord2(pWriter, 0xb1, oElement); break;
          case AscDFH.historyitem_type_Choose: this.writeRecord2(pWriter, 0xb2, oElement); break;
          case AscDFH.historyitem_type_ConstrLst: this.writeRecord2(pWriter, 0xb3, oElement); break;
          case AscDFH.historyitem_type_ForEach: this.writeRecord2(pWriter, 0xb4, oElement); break;
          case AscDFH.historyitem_type_LayoutNode: this.writeRecord2(pWriter, 0xb5, oElement); break;
          case AscDFH.historyitem_type_PresOf: this.writeRecord2(pWriter, 0xb6, oElement); break;
          case AscDFH.historyitem_type_RuleLst: this.writeRecord2(pWriter, 0xb7, oElement); break;
          case AscDFH.historyitem_type_SShape: this.writeRecord2(pWriter, 0xb8, oElement); break;
          case AscDFH.historyitem_type_VarLst: this.writeRecord2(pWriter, 0xb9, oElement); break;
          default: break;
        }
      }
    };
    Else.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
    };

    Else.prototype.readChild = function(nType, pReader) {
      this.readElement(pReader, nType);
    };
    Else.prototype.getChildren = function() {
      return [].concat(this.list);
    };

    Else.prototype.startAlgorithm = function (pointTree, node) {
      var constrLst = this.list.reduce(function (acc, current) {
        return current instanceof ConstrLst ? current : acc;
      }, null);
      if (constrLst) {
        constrLst.startSetConstr(pointTree, node);
      }
      this.list.forEach(function (element) {
        if (element instanceof AscFormat.ForEach || element instanceof AscFormat.LayoutNode || element instanceof AscFormat.Choose) {
          element.startAlgorithm(pointTree, node);
        }
      });
    }

    changesFactory[AscDFH.historyitem_IteratorAttributesAddAxis] = CChangeContent;
    changesFactory[AscDFH.historyitem_IteratorAttributesRemoveAxis] = CChangeContent;
    changesFactory[AscDFH.historyitem_IteratorAttributesAddCnt] = AscDFH.CChangesDrawingsContentLong;
    changesFactory[AscDFH.historyitem_IteratorAttributesRemoveCnt] = AscDFH.CChangesDrawingsContentLong;
    changesFactory[AscDFH.historyitem_IteratorAttributesAddHideLastTrans] = AscDFH.CChangesDrawingsContentBool;
    changesFactory[AscDFH.historyitem_IteratorAttributesRemoveHideLastTrans] = AscDFH.CChangesDrawingsContentBool;
    changesFactory[AscDFH.historyitem_IteratorAttributesAddPtType] = CChangeContent;
    changesFactory[AscDFH.historyitem_IteratorAttributesRemovePtType] = CChangeContent;
    changesFactory[AscDFH.historyitem_IteratorAttributesAddSt] = AscDFH.CChangesDrawingsContentLong;
    changesFactory[AscDFH.historyitem_IteratorAttributesRemoveSt] = AscDFH.CChangesDrawingsContentLong;
    changesFactory[AscDFH.historyitem_IteratorAttributesAddStep] = AscDFH.CChangesDrawingsContentLong;
    changesFactory[AscDFH.historyitem_IteratorAttributesRemoveStep] = AscDFH.CChangesDrawingsContentLong;
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesAddAxis] = function (oClass) {
      return oClass.axis;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesRemoveAxis] = function (oClass) {
      return oClass.axis;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesAddCnt] = function (oClass) {
      return oClass.cnt;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesRemoveCnt] = function (oClass) {
      return oClass.cnt;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesAddHideLastTrans] = function (oClass) {
      return oClass.hideLastTrans;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesRemoveHideLastTrans] = function (oClass) {
      return oClass.hideLastTrans;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesAddPtType] = function (oClass) {
      return oClass.ptType;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesRemovePtType] = function (oClass) {
      return oClass.ptType;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesAddSt] = function (oClass) {
      return oClass.st;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesRemoveSt] = function (oClass) {
      return oClass.st;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesAddStep] = function (oClass) {
      return oClass.step;
    };
    drawingContentChanges[AscDFH.historyitem_IteratorAttributesRemoveStep] = function (oClass) {
      return oClass.step;
    };

    function IteratorAttributes() {
      CBaseFormatObject.call(this);
      this.axis = [];
      this.cnt = [];
      this.hideLastTrans = [];
      this.ptType = [];
      this.st = [];
      this.step = [];
    }

    InitClass(IteratorAttributes, CBaseFormatObject, AscDFH.historyitem_type_IteratorAttributes);

    IteratorAttributes.prototype.addToLstAxis = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.axis.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_IteratorAttributesAddAxis, nInsertIdx, [oPr], true));
      this.axis.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    IteratorAttributes.prototype.removeFromLstAxis = function (nIdx) {
      if (nIdx > -1 && nIdx < this.axis.length) {
        this.axis[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_IteratorAttributesRemoveAxis, nIdx, [this.axis[nIdx]], false));
        this.axis.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstCnt = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.cnt.length, Math.max(0, nIdx));
      oHistory.Add(new AscDFH.CChangesDrawingsContentLong(this, AscDFH.historyitem_IteratorAttributesAddCnt, nInsertIdx, [oPr], true));
      this.cnt.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    IteratorAttributes.prototype.removeFromLstCnt = function (nIdx) {
      if (nIdx > -1 && nIdx < this.cnt.length) {
        this.cnt[nIdx].setParent(null);
        oHistory.Add(new AscDFH.CChangesDrawingsContentLong(this, AscDFH.historyitem_IteratorAttributesRemoveCnt, nIdx, [this.cnt[nIdx]], false));
        this.cnt.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstHideLastTrans = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.hideLastTrans.length, Math.max(0, nIdx));
      oHistory.Add(new AscDFH.CChangesDrawingsContentBool(this, AscDFH.historyitem_IteratorAttributesAddHideLastTrans, nInsertIdx, [oPr], true));
      this.hideLastTrans.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    IteratorAttributes.prototype.removeFromLstHideLastTrans = function (nIdx) {
      if (nIdx > -1 && nIdx < this.hideLastTrans.length) {
        this.hideLastTrans[nIdx].setParent(null);
        oHistory.Add(new AscDFH.CChangesDrawingsContentBool(this, AscDFH.historyitem_IteratorAttributesRemoveHideLastTrans, nIdx, [this.hideLastTrans[nIdx]], false));
        this.hideLastTrans.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstPtType = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.ptType.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_IteratorAttributesAddPtType, nInsertIdx, [oPr], true));
      this.ptType.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    IteratorAttributes.prototype.removeFromLstPtType = function (nIdx) {
      if (nIdx > -1 && nIdx < this.ptType.length) {
        this.ptType[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_IteratorAttributesRemovePtType, nIdx, [this.ptType[nIdx]], false));
        this.ptType.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstSt = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.st.length, Math.max(0, nIdx));
      oHistory.Add(new AscDFH.CChangesDrawingsContentLong(this, AscDFH.historyitem_IteratorAttributesAddSt, nInsertIdx, [oPr], true));
      this.st.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    IteratorAttributes.prototype.removeFromLstSt = function (nIdx) {
      if (nIdx > -1 && nIdx < this.st.length) {
        this.st[nIdx].setParent(null);
        oHistory.Add(new AscDFH.CChangesDrawingsContentLong(this, AscDFH.historyitem_IteratorAttributesRemoveSt, nIdx, [this.st[nIdx]], false));
        this.st.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstStep = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.step.length, Math.max(0, nIdx));
      oHistory.Add(new AscDFH.CChangesDrawingsContentLong(this, AscDFH.historyitem_IteratorAttributesAddStep, nInsertIdx, [oPr], true));
      this.step.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    IteratorAttributes.prototype.removeFromLstStep = function (nIdx) {
      if (nIdx > -1 && nIdx < this.step.length) {
        this.step[nIdx].setParent(null);
        oHistory.Add(new AscDFH.CChangesDrawingsContentLong(this, AscDFH.historyitem_IteratorAttributesRemoveStep, nIdx, [this.step[nIdx]], false));
        this.step.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.fillObject = function (oCopy, oIdMap) {
      for (var nIdx = 0; nIdx < this.axis.length; ++nIdx) {
        oCopy.addToLstAxis(nIdx, this.axis[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.cnt.length; ++nIdx) {
        oCopy.addToLstCnt(nIdx, this.cnt[nIdx]);
      }
      for (nIdx = 0; nIdx < this.hideLastTrans.length; ++nIdx) {
        oCopy.addToLstHideLastTrans(nIdx, this.hideLastTrans[nIdx]);
      }
      for (nIdx = 0; nIdx < this.ptType.length; ++nIdx) {
        oCopy.addToLstPtType(nIdx, this.ptType[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.st.length; ++nIdx) {
        oCopy.addToLstSt(nIdx, this.st[nIdx]);
      }
      for (nIdx = 0; nIdx < this.step.length; ++nIdx) {
        oCopy.addToLstStep(nIdx, this.step[nIdx]);
      }
    };

    IteratorAttributes.prototype.Clear_ContentChanges = function() {
    };
    IteratorAttributes.prototype.Add_ContentChanges = function(Changes) {
    };
    IteratorAttributes.prototype.Refresh_ContentChanges = function() {
    };


    changesFactory[AscDFH.historyitem_AxisTypeVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_AxisTypeVal] = function (oClass, value) {
      oClass.val = value;
    };

    function AxisType() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(AxisType, CBaseFormatObject, AscDFH.historyitem_type_AxisType);

    AxisType.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AxisTypeVal, this.getVal(), pr));
      this.val = pr;
    }

    AxisType.prototype.getVal = function () {
      return this.val;
    }

    AxisType.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }

    changesFactory[AscDFH.historyitem_ElementTypeVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_ElementTypeVal] = function (oClass, value) {
      oClass.val = value;
    };

    function ElementType() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(ElementType, CBaseFormatObject, AscDFH.historyitem_type_ElementType);

    ElementType.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ElementTypeVal, this.getVal(), pr));
      this.val = pr;
    }

    ElementType.prototype.getVal = function () {
      return this.val;
    }

    ElementType.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }

    changesFactory[AscDFH.historyitem_FunctionValueAnimLvlStr] = CChangeLong;
    changesFactory[AscDFH.historyitem_FunctionValueAnimOneStr] = CChangeLong;
    changesFactory[AscDFH.historyitem_FunctionValueDirection] = CChangeLong;
    changesFactory[AscDFH.historyitem_FunctionValueHierBranchStyle] = CChangeLong;
    changesFactory[AscDFH.historyitem_FunctionValueResizeHandlesStr] = CChangeLong;
    changesFactory[AscDFH.historyitem_FunctionValueBool] = CChangeBool;
    changesFactory[AscDFH.historyitem_FunctionValueInt] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_FunctionValueAnimLvlStr] = function (oClass, value) {
      oClass.animLvlStr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_FunctionValueAnimOneStr] = function (oClass, value) {
      oClass.animOneStr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_FunctionValueDirection] = function (oClass, value) {
      oClass.direction = value;
    };
    drawingsChangesMap[AscDFH.historyitem_FunctionValueHierBranchStyle] = function (oClass, value) {
      oClass.hierBranchStyle = value;
    };
    drawingsChangesMap[AscDFH.historyitem_FunctionValueResizeHandlesStr] = function (oClass, value) {
      oClass.resizeHandlesStr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_FunctionValueBool] = function (oClass, value) {
      oClass.bool = value;
    };
    drawingsChangesMap[AscDFH.historyitem_FunctionValueInt] = function (oClass, value) {
      oClass.int = value;
    };

    function FunctionValue() {
      CBaseFormatObject.call(this);
      this.animLvlStr = null;
      this.animOneStr = null;
      this.direction = null;
      this.hierBranchStyle = null;
      this.resizeHandlesStr = null;
      this.bool = null;
      this.int = null;
    }

    InitClass(FunctionValue, CBaseFormatObject, AscDFH.historyitem_type_FunctionValue);

    FunctionValue.prototype.setAnimLvlStr = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_FunctionValueAnimLvlStr, this.getAnimLvlStr(), pr));
      this.animLvlStr = pr;
    }

    FunctionValue.prototype.setAnimOneStr = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_FunctionValueAnimOneStr, this.getAnimOneStr(), pr));
      this.animOneStr = pr;
    }

    FunctionValue.prototype.setDirection = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_FunctionValueDirection, this.getDirection(), pr));
      this.direction = pr;
    }

    FunctionValue.prototype.setHierBranchStyle = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_FunctionValueHierBranchStyle, this.getHierBranchStyle(), pr));
      this.hierBranchStyle = pr;
    }

    FunctionValue.prototype.setResizeHandlesStr = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_FunctionValueResizeHandlesStr, this.getResizeHandlesStr(), pr));
      this.resizeHandlesStr = pr;
    }

    FunctionValue.prototype.setBool = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_FunctionValueBool, this.getBool(), pr));
      this.bool = pr;
    }

    FunctionValue.prototype.setInt = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_FunctionValueInt, this.getInt(), pr));
      this.int = pr;
    }

    FunctionValue.prototype.setVal = function (pr) {
      this.setInt(pr);
    }

    FunctionValue.prototype.getVal = function () {
      return this.int;
    }

    FunctionValue.prototype.getAnimLvlStr = function () {
      return this.animLvlStr;
    }

    FunctionValue.prototype.getAnimOneStr = function () {
      return this.animOneStr;
    }

    FunctionValue.prototype.getDirection = function () {
      return this.direction;
    }

    FunctionValue.prototype.getHierBranchStyle = function () {
      return this.hierBranchStyle;
    }

    FunctionValue.prototype.getResizeHandlesStr = function () {
      return this.resizeHandlesStr;
    }

    FunctionValue.prototype.getBool = function () {
      return this.bool;
    }

    FunctionValue.prototype.getInt = function () {
      return this.int;
    }

    FunctionValue.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setAnimLvlStr(this.getAnimLvlStr());
      oCopy.setAnimOneStr(this.getAnimOneStr());
      oCopy.setDirection(this.getDirection());
      oCopy.setHierBranchStyle(this.getHierBranchStyle());
      oCopy.setResizeHandlesStr(this.getResizeHandlesStr());
      oCopy.setBool(this.getBool());
      oCopy.setInt(this.getInt());
    }


    changesFactory[AscDFH.historyitem_IfArg] = CChangeString;
    changesFactory[AscDFH.historyitem_IfFunc] = CChangeLong;
    changesFactory[AscDFH.historyitem_IfName] = CChangeString;
    changesFactory[AscDFH.historyitem_IfOp] = CChangeLong;
    changesFactory[AscDFH.historyitem_IfVal] = CChangeString;
    changesFactory[AscDFH.historyitem_IfAddList] = CChangeContent;
    changesFactory[AscDFH.historyitem_IfRemoveList] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_IfArg] = function (oClass, value) {
      oClass.arg = value;
    };
    drawingsChangesMap[AscDFH.historyitem_IfFunc] = function (oClass, value) {
      oClass.func = value;
    };
    drawingsChangesMap[AscDFH.historyitem_IfName] = function (oClass, value) {
      oClass.name = value;
    };
    drawingsChangesMap[AscDFH.historyitem_IfOp] = function (oClass, value) {
      oClass.op = value;
    };
    drawingsChangesMap[AscDFH.historyitem_IfVal] = function (oClass, value) {
      oClass.val = value;
    };
    drawingContentChanges[AscDFH.historyitem_IfAddList] = function (oClass) {
      return oClass.list;
    };
    drawingContentChanges[AscDFH.historyitem_IfRemoveList] = function (oClass) {
      return oClass.list;
    };

    function If() {
      IteratorAttributes.call(this);
      this.arg = null;
      this.func = null;
      this.name = null;
      this.op = null;
      this.val = null;
      this.list = [];
    }

    InitClass(If, IteratorAttributes, AscDFH.historyitem_type_If);

    If.prototype.setArg = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_IfArg, this.getArg(), pr));
      this.arg = pr;
    }

    If.prototype.setFunc = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_IfFunc, this.getFunc(), pr));
      this.func = pr;
    }

    If.prototype.setName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_IfName, this.getName(), pr));
      this.name = pr;
    }

    If.prototype.setOp = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_IfOp, this.getOp(), pr));
      this.op = pr;
    }

    If.prototype.setVal = function (oPr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_IfVal, this.getVal(), oPr));
      this.val = oPr;
    }

    If.prototype.addToLstList = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_IfAddList, nInsertIdx, [oPr], true));
      this.list.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    If.prototype.removeFromLstList = function (nIdx) {
      if (nIdx > -1 && nIdx < this.list.length) {
        this.list[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_IfRemoveList, nIdx, [this.list[nIdx]], false));
        this.list.splice(nIdx, 1);
      }
    };

    If.prototype.getArg = function () {
      return this.arg;
    }

    If.prototype.getFunc = function () {
      return this.func;
    }

    If.prototype.getName = function () {
      return this.name;
    }

    If.prototype.getOp = function () {
      return this.op;
    }

    If.prototype.getVal = function () {
      return this.val;
    }

    If.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setArg(this.getArg());
      oCopy.setFunc(this.getFunc());
      oCopy.setName(this.getName());
      oCopy.setOp(this.getOp());
      oCopy.setVal(this.getVal());
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        oCopy.addToLstList(nIdx, this.list[nIdx].createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.axis.length; ++nIdx) {
        oCopy.addToLstAxis(nIdx, this.axis[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.cnt.length; ++nIdx) {
        oCopy.addToLstCnt(nIdx, this.cnt[nIdx]);
      }
      for (nIdx = 0; nIdx < this.hideLastTrans.length; ++nIdx) {
        oCopy.addToLstHideLastTrans(nIdx, this.hideLastTrans[nIdx]);
      }
      for (nIdx = 0; nIdx < this.ptType.length; ++nIdx) {
        oCopy.addToLstPtType(nIdx, this.ptType[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.st.length; ++nIdx) {
        oCopy.addToLstSt(nIdx, this.st[nIdx]);
      }
      for (nIdx = 0; nIdx < this.step.length; ++nIdx) {
        oCopy.addToLstStep(nIdx, this.step[nIdx]);
      }
    }

    If.prototype.readElement = function(pReader, nType) {
      var oElement = null;
      switch(nType) {
        case 0xb1: oElement = new Alg(); break;
        case 0xb2: oElement = new Choose(); break;
        case 0xb3: oElement = new ConstrLst(); break;
        case 0xb4: oElement = new ForEach(); break;
        case 0xb5: oElement = new LayoutNode(); break;
        case 0xb6: oElement = new PresOf(); break;
        case 0xb7: oElement = new RuleLst(); break;
        case 0xb8: oElement = new SShape(); break;
        case 0xb9: oElement = new VarLst(); break;
        default:break;
      }
      if(oElement) {
        oElement.fromPPTY(pReader);
        this.addToLstList(this.list.length, oElement);
      }
    };

    If.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
      for (var i = 0; i < this.st.length; i += 1) {
        pWriter._WriteInt1(1, this.st[i]);
      }
      for (i = 0; i < this.step.length; i += 1) {
        pWriter._WriteInt1(2, this.step[i]);
      }
      for (i = 0; i < this.hideLastTrans.length; i += 1) {
        pWriter._WriteBool1(3, this.hideLastTrans[i]);
      }
      for (i = 0; i < this.cnt.length; i += 1) {
        pWriter._WriteInt1(4, this.cnt[i]);
      }
      for (i = 0; i < this.axis.length; i += 1) {
        pWriter._WriteUChar1(5, this.axis[i].getVal());
      }
      for (i = 0; i < this.ptType.length; i += 1) {
        pWriter._WriteUChar1(6, this.ptType[i].getVal());
      }
      pWriter._WriteString2(7, this.ref);
      pWriter._WriteUChar2(8, this.op);
      pWriter._WriteUChar2(9, this.func);
      pWriter._WriteString2(10, this.val);
      pWriter._WriteString2(11, this.arg);
    };
    If.prototype.writeChildren = function(pWriter) {
      for(var nIndex = 0; nIndex < this.list.length; ++nIndex) {
        var oElement = this.list[nIndex];
        switch (oElement.getObjectType()) {
          case AscDFH.historyitem_type_Alg: this.writeRecord2(pWriter, 0xb1, oElement); break;
          case AscDFH.historyitem_type_Choose: this.writeRecord2(pWriter, 0xb2, oElement); break;
          case AscDFH.historyitem_type_ConstrLst: this.writeRecord2(pWriter, 0xb3, oElement); break;
          case AscDFH.historyitem_type_ForEach: this.writeRecord2(pWriter, 0xb4, oElement); break;
          case AscDFH.historyitem_type_LayoutNode: this.writeRecord2(pWriter, 0xb5, oElement); break;
          case AscDFH.historyitem_type_PresOf: this.writeRecord2(pWriter, 0xb6, oElement); break;
          case AscDFH.historyitem_type_RuleLst: this.writeRecord2(pWriter, 0xb7, oElement); break;
          case AscDFH.historyitem_type_SShape: this.writeRecord2(pWriter, 0xb8, oElement); break;
          case AscDFH.historyitem_type_VarLst: this.writeRecord2(pWriter, 0xb9, oElement); break;
          default: break;
        }
      }
    };
    If.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
      else if (1 === nType) this.addToLstSt(this.st.length, oStream.GetLong());
      else if (2 === nType) this.addToLstStep(this.step.length, oStream.GetLong());
      else if (3 === nType) this.addToLstHideLastTrans(this.hideLastTrans.length, oStream.GetBool());
      else if (4 === nType) this.addToLstCnt(this.cnt.length, oStream.GetLong());
      else if (5 === nType) {
        var axis = new AxisType();
        axis.setVal(oStream.GetUChar());
        this.addToLstAxis(this.axis.length, axis);
      }
      else if (6 === nType) {
        var ptType = new ElementType();
        ptType.setVal(oStream.GetUChar());
        this.addToLstPtType(this.ptType.length, ptType);
      }
      else if (7 === nType) this.setRef(oStream.GetString2());
      else if (8 === nType) this.setOp(oStream.GetUChar());
      else if (9 === nType) this.setFunc(oStream.GetUChar());
      else if (10 === nType) this.setVal(oStream.GetString2());
      else if (11 === nType) this.setArg(oStream.GetString2());
    };

    If.prototype.readChild = function(nType, pReader) {
      this.readElement(pReader, nType);
    };
    If.prototype.getChildren = function() {
      return [].concat(this.list);
    };

    If.prototype.startAlgorithm = function (pointTree, node) {
      var check = this.checkCondition(pointTree, node);
      if (check) {
        var constrLst = this.list.reduce(function (acc, current) {
          return current instanceof ConstrLst ? current : acc;
        }, null);
        if (constrLst) {
          constrLst.startSetConstr(pointTree, node);
        }
        this.list.forEach(function (element) {
          if (element instanceof AscFormat.ForEach || element instanceof AscFormat.LayoutNode || element instanceof AscFormat.Choose) {
            element.startAlgorithm(pointTree, node);
          }
        });
      }
      return check;
    }


    If.prototype.checkCondition = function (pointTree, nodeData) {
      if (nodeData.node) {
        var nodes = [];
        var smartArtNode = nodeData.node.parent;
        this.axis.forEach(function (type) {
          nodes = nodes.concat(smartArtNode.getAxis(type.getVal()));
        });
      }

      switch (this.func) {
        case If_func_cnt:
          return this.funcCnt(nodes);
        case If_func_depth:
          return this.funcDepth(nodes);
        case If_func_maxDepth:
          return this.funcMaxDepth(nodes);
        case If_func_pos:
          return this.funcPos(nodes);
        case If_func_posEven:
          return this.funcPosEven(nodes);
        case If_func_posOdd:
          return this.funcPosOdd(nodes);
        case If_func_revPos:
          return this.funcRevPos(nodes);
        case If_func_var:
          return this.funcVar(nodeData);
        default:
          return false;
      }
    }

    If.prototype.funcVar = function (nodeData) {
      if (nodeData && nodeData.node) {
        var nodeElement = nodeData.node.parent;
        var rootOfTree = nodeElement.getRoot()[0];
        var rootPres = rootOfTree.data.getPresWithVarLst();
        var currentPres = nodeData.node.getPresByNameAndStyleLbl(nodeData.name, nodeData.styleLbl);
        switch (this.arg) {
          case 'animLvl':
          case 'hierBranch':
          case 'bulletEnabled':
            if (currentPres && currentPres.prSet && currentPres.prSet.presLayoutVars) {
              return this.compare(currentPres.prSet.presLayoutVars.getVal(this.arg));
            }
            break;

          case 'animOne':
          case 'dir':
          case 'chPref':
          case 'chMax':
          case 'orgChart':
          case 'resizeHandles':
            if (rootPres && rootPres.prSet && rootPres.prSet.presLayoutVars && rootPres.prSet.presLayoutVars[this.arg]) {
              return this.compare(rootPres.prSet.presLayoutVars.getVal(this.arg));
            }
            break;
          default:
            return false;
        }
      }
      return false;
    }

    If.prototype.funcCnt = function (axis) {
      var count = 0;
      this.ptType.forEach(function (Type) {
        var typeOfPoint = Type.getVal();
        axis.forEach(function (node) {
          count += node.getPoints(typeOfPoint).length;
        });
      });
      return this.compare(count);
    }

    If.prototype.funcDepth = function (axis) {
      return this.compare(axis[0].data.depth);
    }

    If.prototype.funcMaxDepth = function (axis) {
      var maxDepth = axis.reduce(function (acc, b) {
        if (b.data.depth > acc) {
          return b.data.depth;
        }
        return acc;
      }, 0);
      return this.compare(maxDepth);
    }

    If.prototype.funcPos = function () {

    }

    If.prototype.funcPosEven = function (axis) {
      return this.funcPos(axis) % 2 === 0;
    }

    If.prototype.funcPosOdd = function (axis) {
      return this.funcPos(axis) % 2 === 1;
    }

    If.prototype.funcRevPos = function () {

    }

    If.prototype.compare = function (comparingArg) {
      var val = this.valAdapter(this.val);
      var adaptComparingArg = this.valAdapter(comparingArg);
      switch (this.op) {
        case If_op_equ:
          return adaptComparingArg === val;
        case If_op_gt:
          return adaptComparingArg > val;
        case If_op_gte:
          return adaptComparingArg >= val;
        case If_op_lt:
          return adaptComparingArg < val;
        case If_op_lte:
          return adaptComparingArg <= val;
        case If_op_neq:
          return adaptComparingArg !== val;
        default:
          return false;
      }
    }

    If.prototype.valAdapter = function (value) {
      var adaptVal;
      if (!(parseFloat(value) !== parseFloat(value))) {
        adaptVal = parseFloat(value);
      } else {
        switch (value) {
          case 'norm':
          case 'none':
          case 'hang':
          case 'exact':
          case false:
          case 'false':
            adaptVal = 0;
            break;
          case 'rev':
          case 'ctr':
          case 'branch':
          case 'init':
          case 'rel':
          case true:
          case 'true':
            adaptVal = 1;
            break;
          case 'lvl':
          case 'one':
          case 'l':
            adaptVal = 2;
            break;
          case 'r':
            adaptVal = 3;
            break;
          case 'std':
            adaptVal = 4;
            break;
          default:
            return;
        }
      }
      return adaptVal;
    }

    function ConstrLst() {
      CCommonDataList.call(this);
    }

    InitClass(ConstrLst, CCommonDataList, AscDFH.historyitem_type_ConstrLst);

    ConstrLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          var oChild = new Constr();
          oChild.fromPPTY(pReader);
          this.addToLst(this.list.length, oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    ConstrLst.prototype.startSetConstr = function (pointTree, node) {
      var constrWithPrimFont = [];
      var constrWithSecFont = [];
      this.list.forEach(function (constr) {
        if (constr.type === Constr_type_primFontSz) {
          constrWithPrimFont.push({
            constr: constr,
            node: node,
          });
        } else if (constr.type === Constr_type_secFontSz) {
          constrWithSecFont.push({
            constr: constr,
            node: node,
          });
        } else {
          constr.setConstr(pointTree, [{node: node, constr: constr}]);
        }
      });
      if (constrWithPrimFont.length !== 0) {
        constrWithPrimFont[0].constr.setConstr(pointTree, constrWithPrimFont);
      }
      if (constrWithSecFont.length !== 0) {
        constrWithSecFont[0].constr.setConstr(pointTree, constrWithSecFont);
      }
    }



    changesFactory[AscDFH.historyitem_ConstrFact] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_ConstrFor] = CChangeLong;
    changesFactory[AscDFH.historyitem_ConstrForName] = CChangeString;
    changesFactory[AscDFH.historyitem_ConstrOp] = CChangeLong;
    changesFactory[AscDFH.historyitem_ConstrPtType] = CChangeObject;
    changesFactory[AscDFH.historyitem_ConstrRefFor] = CChangeLong;
    changesFactory[AscDFH.historyitem_ConstrRefForName] = CChangeString;
    changesFactory[AscDFH.historyitem_ConstrRefPtType] = CChangeObject;
    changesFactory[AscDFH.historyitem_ConstrRefType] = CChangeLong;
    changesFactory[AscDFH.historyitem_ConstrType] = CChangeLong;
    changesFactory[AscDFH.historyitem_ConstrVal] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_ConstrExtLst] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_ConstrFact] = function (oClass, value) {
      oClass.fact = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrFor] = function (oClass, value) {
      oClass.for = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrForName] = function (oClass, value) {
      oClass.forName = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrOp] = function (oClass, value) {
      oClass.op = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrPtType] = function (oClass, value) {
      oClass.ptType = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrRefFor] = function (oClass, value) {
      oClass.refFor = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrRefForName] = function (oClass, value) {
      oClass.refForName = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrRefPtType] = function (oClass, value) {
      oClass.refPtType = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrRefType] = function (oClass, value) {
      oClass.refType = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrType] = function (oClass, value) {
      oClass.type = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrVal] = function (oClass, value) {
      oClass.val = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ConstrExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };

    function Constr() {
      CBaseFormatObject.call(this);
      this.fact = null;
      this.for = null;
      this.forName = null;
      this.op = null;
      this.setPtType(new ElementType());
      this.refFor = null;
      this.refForName = null;
      this.setRefPtType(new ElementType());
      this.refType = null;
      this.type = null;
      this.val = null;
      this.extLst = null;
    }

    InitClass(Constr, CBaseFormatObject, AscDFH.historyitem_type_Constr);

    Constr.prototype.setFact = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_ConstrFact, this.getFact(), pr));
      this.fact = pr;
    }

    Constr.prototype.setFor = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ConstrFor, this.getFor(), pr));
      this.for = pr;
    }

    Constr.prototype.setForName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ConstrForName, this.getForName(), pr));
      this.forName = pr;
    }

    Constr.prototype.setOp = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ConstrOp, this.getOp(), pr));
      this.op = pr;
    }

    Constr.prototype.setPtType = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ConstrPtType, this.getPtType(), oPr));
      this.ptType = oPr;
      this.setParentToChild(oPr);
    }

    Constr.prototype.setRefFor = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ConstrRefFor, this.getRefFor(), pr));
      this.refFor = pr;
    }

    Constr.prototype.setRefForName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ConstrRefForName, this.getRefForName(), pr));
      this.refForName = pr;
    }

    Constr.prototype.setRefPtType = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ConstrRefPtType, this.getRefPtType(), oPr));
      this.refPtType = oPr;
      this.setParentToChild(oPr);
    }

    Constr.prototype.setRefType = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ConstrRefType, this.getRefType(), pr));
      this.refType = pr;
    }

    Constr.prototype.setType = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ConstrType, this.getType(), pr));
      this.type = pr;
    }

    Constr.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_ConstrVal, this.getVal(), pr));
      this.val = pr;
    }

    Constr.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ConstrExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    Constr.prototype.getFact = function () {
      return this.fact;
    }

    Constr.prototype.getFor = function () {
      return this.for;
    }

    Constr.prototype.getForName = function () {
      return this.forName;
    }

    Constr.prototype.getOp = function () {
      return this.op;
    }

    Constr.prototype.getPtType = function () {
      return this.ptType;
    }

    Constr.prototype.getRefFor = function () {
      return this.refFor;
    }

    Constr.prototype.getRefForName = function () {
      return this.refForName;
    }

    Constr.prototype.getRefPtType = function () {
      return this.refPtType;
    }

    Constr.prototype.getRefType = function () {
      return this.refType;
    }

    Constr.prototype.getType = function () {
      return this.type;
    }

    Constr.prototype.getVal = function () {
      return this.val;
    }

    Constr.prototype.getExtLst = function () {
      return this.extLst;
    }

    Constr.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setFact(this.getFact());
      oCopy.setFor(this.getFor());
      oCopy.setForName(this.getForName());
      oCopy.setOp(this.getOp());
      oCopy.setRefFor(this.getRefFor());
      oCopy.setRefForName(this.getRefForName());
      oCopy.setRefType(this.getRefType());
      oCopy.setType(this.getType());
      oCopy.setVal(this.getVal());
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getPtType()) {
        oCopy.setPtType(this.getPtType().createDuplicate(oIdMap));
      }
      if (this.getRefPtType()) {
        oCopy.setRefPtType(this.getRefPtType().createDuplicate(oIdMap));
      }
    }

    Constr.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteDoubleReal2(0, this.fact);
      pWriter._WriteUChar2(1, this.for);
      pWriter._WriteString2(2, this.forName);
      pWriter._WriteUChar2(3, this.op);
      pWriter._WriteUChar2(4, this.ptType.getVal());
      pWriter._WriteUChar2(5, this.refFor);
      pWriter._WriteString2(6, this.refForName);
      pWriter._WriteUChar2(7, this.refPtType.getVal());
      pWriter._WriteUChar2(8, this.refType);
      pWriter._WriteUChar2(9, this.type);
      pWriter._WriteDoubleReal2(10, this.val);
    };
    Constr.prototype.writeChildren = function(pWriter) {
    };
    Constr.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setFact(oStream.GetDouble());
      else if (1 === nType) this.setFor(oStream.GetUChar());
      else if (2 === nType) this.setForName(oStream.GetString2());
      else if (3 === nType) this.setOp(oStream.GetUChar());
      else if (4 === nType) {
        var pt = new ElementType();
        pt.setVal(oStream.GetUChar());
        this.setPtType(pt);
      }
      else if (5 === nType) this.setRefFor(oStream.GetUChar());
      else if (6 === nType) this.setRefForName(oStream.GetString2());
      else if (7 === nType) {
        var pt = new ElementType();
        pt.setVal(oStream.GetUChar());
        this.setRefPtType(pt);
      }
      else if (8 === nType) this.setRefType(oStream.GetUChar());
      else if (9 === nType) this.setType(oStream.GetUChar());
      else if (10 === nType) this.setVal(oStream.GetDouble());
    };
    Constr.prototype.readChild = function(nType, pReader) {
    };

    Constr.prototype.getFieldScale = function (type) {
      if (type) {
        switch (type) {
          case Constr_type_alignOff:
            return 1;
          case Constr_type_b:
            return 1;
          case Constr_type_begMarg:
            return 1;
          case Constr_type_begPad:
            return 1;
          case Constr_type_bendDist:
            return 1;
          case Constr_type_bMarg:
            return 1;
          case Constr_type_bOff:
            return 1;
          case Constr_type_connDist:
            return 1;
          case Constr_type_ctrX:
            return 1;
          case Constr_type_ctrXOff:
            return 1;
          case Constr_type_ctrY:
            return 1;
          case Constr_type_ctrYOff:
            return 1;
          case Constr_type_diam:
            return 1;
          case Constr_type_endMarg:
            return 1;
          case Constr_type_endPad:
            return 1;
          case Constr_type_h:
          case Constr_type_w:
            return 36000;
          case Constr_type_hArH:
            return 1;
          case Constr_type_hOff: // TODO: add to constr type in x2t
            return 1;
          case Constr_type_l:
            return 1;
          case Constr_type_lMarg:
            return 1;
          case Constr_type_lOff:
            return 1;
          case Constr_type_none:
            return 1;
          case Constr_type_primFontSz:
          case Constr_type_secFontSz:
            return 100 * Constr_font_scale;
          case Constr_type_pyraAcctRatio:
            return 1;
          case Constr_type_r:
            return 1;
          case Constr_type_rMarg:
            return 1;
          case Constr_type_rOff:
            return 1;
          case Constr_type_secSibSp:
            return 1;
          case Constr_type_sibSp:
            return 1;
          case Constr_type_sp:
            return 1;
          case Constr_type_stemThick:
            return 1;
          case Constr_type_t:
            return 1;
          case Constr_type_tMarg:
            return 1;
          case Constr_type_tOff:
            return 1;
          case Constr_type_userA:
            return 1;
          case Constr_type_userB:
            return 1;
          case Constr_type_userC:
            return 1;
          case Constr_type_userD:
            return 1;
          case Constr_type_userE:
            return 1;
          case Constr_type_userF:
            return 1;
          case Constr_type_userG:
            return 1;
          case Constr_type_userH:
            return 1;
          case Constr_type_userI:
            return 1;
          case Constr_type_userJ:
            return 1;
          case Constr_type_userK:
            return 1;
          case Constr_type_userL:
            return 1;
          case Constr_type_userM:
            return 1;
          case Constr_type_userN:
            return 1;
          case Constr_type_userO:
            return 1;
          case Constr_type_userP:
            return 1;
          case Constr_type_userQ:
            return 1;
          case Constr_type_userR:
            return 1;
          case Constr_type_userS:
            return 1;
          case Constr_type_userT:
            return 1;
          case Constr_type_userU:
            return 1;
          case Constr_type_userV:
            return 1;
          case Constr_type_userW:
            return 1;
          case Constr_type_userX:
            return 1;
          case Constr_type_userY:
            return 1;
          case Constr_type_userZ:
            return 1;
          case Constr_type_wArH:
            return 1;
          case Constr_type_wOff:
            return 1;
          default:
            return 1;
        }
      }
    return 1;
    }

    Constr.prototype.getConstrVal = function (shape) {
      var result;
      if (this.val) {
        result = this.val;
        return;
      } else {
        if (shape) {
          var constrType = this.refType;

          switch (constrType) {
            case Constr_type_alignOff:
              break;
            case Constr_type_b:
              break;
            case Constr_type_begMarg:
              break;
            case Constr_type_begPad:
              break;
            case Constr_type_bendDist:
              break;
            case Constr_type_bMarg:
              break;
            case Constr_type_bOff:
              break;
            case Constr_type_connDist:
              break;
            case Constr_type_ctrX:
              break;
            case Constr_type_ctrXOff:
              break;
            case Constr_type_ctrY:
              break;
            case Constr_type_ctrYOff:
              break;
            case Constr_type_diam:
              break;
            case Constr_type_endMarg:
              break;
            case Constr_type_endPad:
              break;
            case Constr_type_h:
              //result = shape.extX;
              break;
            case Constr_type_hArH:
              break;
            case Constr_type_hOff: // TODO: add to constr type in x2t
              break;
            case Constr_type_l:
              break;
            case Constr_type_lMarg:
              break;
            case Constr_type_lOff:
              break;
            case Constr_type_none:
              break;
            case Constr_type_primFontSz:
            case Constr_type_secFontSz:
              if (shape.txBody && shape.txBody.content) {
                result = shape.txBody.content.getFontSizeForConstr();
              }
              break;
            case Constr_type_pyraAcctRatio:
              break;
            case Constr_type_r:
              break;
            case Constr_type_rMarg:
              break;
            case Constr_type_rOff:
              break;
            case Constr_type_secSibSp:
              break;
            case Constr_type_sibSp:
              break;
            case Constr_type_sp:
              break;
            case Constr_type_stemThick:
              break;
            case Constr_type_t:
              break;
            case Constr_type_tMarg:
              break;
            case Constr_type_tOff:
              break;
            case Constr_type_userA:
              break;
            case Constr_type_userB:
              break;
            case Constr_type_userC:
              break;
            case Constr_type_userD:
              break;
            case Constr_type_userE:
              break;
            case Constr_type_userF:
              break;
            case Constr_type_userG:
              break;
            case Constr_type_userH:
              break;
            case Constr_type_userI:
              break;
            case Constr_type_userJ:
              break;
            case Constr_type_userK:
              break;
            case Constr_type_userL:
              break;
            case Constr_type_userM:
              break;
            case Constr_type_userN:
              break;
            case Constr_type_userO:
              break;
            case Constr_type_userP:
              break;
            case Constr_type_userQ:
              break;
            case Constr_type_userR:
              break;
            case Constr_type_userS:
              break;
            case Constr_type_userT:
              break;
            case Constr_type_userU:
              break;
            case Constr_type_userV:
              break;
            case Constr_type_userW:
              break;
            case Constr_type_userX:
              break;
            case Constr_type_userY:
              break;
            case Constr_type_userZ:
              break;
            case Constr_type_w:
              //result = shape.extY;
              break;
            case Constr_type_wArH:
              break;
            case Constr_type_wOff:
              break;
            default:
              return;
          }
          if (typeof result === 'number' && !(result !== result)) {
            result *= this.getFieldScale(constrType);
          }
        }
      }
      if (typeof result === 'number'&& !(result !== result)) {
        var fact = Math.abs(this.fact ? this.fact : 1);
        result *= fact;
        return result;
      }
    }

    Constr.prototype.getShapesFromAxis = function (transfer, isRef) {
      var node = transfer && transfer.node && transfer.node.parent;
      if (node) {
        var axisNodes;
        var shapes = [];
        var fGetAxis = this.getAxisFromParent.bind(this, node);
        var that = this;
        axisNodes = isRef ? fGetAxis(this.refFor) : fGetAxis(this.for);
        if (axisNodes && axisNodes.length !== 0) {
          var ptType;
          if (isRef) {
            ptType = this.refPtType && this.refPtType.getVal();
            shapes = axisNodes.reduce(function (acc, axisNode) {
              if (that.refForName && ptType) {
                return acc.concat(axisNode.getShape(that.refForName, ptType));
              }
              else if (that.refForName) {
                return acc.concat(axisNode.getShape(that.refForName));
              }
              else if (ptType) {
                return acc.concat(axisNode.getShape(transfer.name, ptType));
              } else {
                return acc.concat(axisNode.getShape(transfer.name));
              }
            }, []);
          } else {
            ptType = this.ptType && this.ptType.getVal();
            shapes = axisNodes.reduce(function (acc, axisNode) {
              if (that.forName && ptType) {
                return acc.concat(axisNode.getShape(that.forName, ptType));
              }
              else if (that.forName) {
                return acc.concat(axisNode.getShape(that.forName));
              }
              else if (ptType) {
                return acc.concat(axisNode.getShape(undefined, ptType));
              } else {
                return acc.concat(axisNode.getShape());
              }
            }, []);
          }
        }
        return shapes;
      }
    }

    Constr.prototype.getAxisFromParent = function (node, constrAxisType) {
      switch (constrAxisType) {
        case Constr_for_ch:
          return node.getAxis(AxisType_value_ch);
        case Constr_for_des:
          return node.getAxis(AxisType_value_des);
        case Constr_for_self:
          return node.getAxis(AxisType_value_self);
        default:
          return node.getAxis(AxisType_value_self);
      }
    }

    Constr.prototype.setConstr = function (pointTree, transfer) {
      // var node = transfer;
      // var smartArt = pointTree && pointTree.parent;
      // if (node) {
      //   var shapes = this.getShapesFromAxis(node, false);
      //   var refShape = this.getShapesFromAxis(node, true);
      //   if (this.type === Constr_type_primFontSz || this.type === Constr_type_secFontSz) {
      //     if (smartArt) {
      //       smartArt.setTruthFontSizeInSmartArt(shapes);
      //     }
      //     return;
      //   }
      //
      //   var that = this;
      //   var value = refShape.reduce(function (acc, shape) {
      //     var compute = that.getConstrVal(shape);
      //     if (!acc && typeof compute === 'number') {
      //       return compute;
      //     } else if (typeof compute === 'number' && compute > acc) {
      //       return compute;
      //     }
      //     return acc;
      //   }, undefined);
      //
      //   if (typeof value === 'number' && !(value !== value)) {
      //     value /= this.getFieldScale(this.type);
      //     shapes.forEach(function (shape) {
      //       var setter = that.getConstrSetter(shape);
      //       if (setter) {
      //         setter.call(shape, value);
      //       }
      //     });
      //   }
      // }
    };

    Constr.prototype.getConstrSetter = function (shape) {
      if (shape) {
        var constrType = this.type;
        switch (constrType) {
          case Constr_type_alignOff:
            break;
          case Constr_type_b:
            break;
          case Constr_type_begMarg:
            break;
          case Constr_type_begPad:
            break;
          case Constr_type_bendDist:
            break;
          case Constr_type_bMarg:
            break;
          case Constr_type_bOff:
            break;
          case Constr_type_connDist:
            break;
          case Constr_type_ctrX:
            break;
          case Constr_type_ctrXOff:
            break;
          case Constr_type_ctrY:
            break;
          case Constr_type_ctrYOff:
            break;
          case Constr_type_diam:
            break;
          case Constr_type_endMarg:
            break;
          case Constr_type_endPad:
            break;
          case Constr_type_h:
            //return shape.setResizeHeightConstr;
            break;
          case Constr_type_hArH:
            break;
          case Constr_type_hOff: // TODO: add to constr type in x2t
            break;
          case Constr_type_l:
            break;
          case Constr_type_lMarg:
            break;
          case Constr_type_lOff:
            break;
          case Constr_type_none:
            break;
          case Constr_type_primFontSz:
          case Constr_type_secFontSz:
            //return shape.setFontSizeInSmartArt;
            break;
          case Constr_type_pyraAcctRatio:
            break;
          case Constr_type_r:
            break;
          case Constr_type_rMarg:
            break;
          case Constr_type_rOff:
            break;
          case Constr_type_secSibSp:
            break;
          case Constr_type_sibSp:
            break;
          case Constr_type_sp:
            break;
          case Constr_type_stemThick:
            break;
          case Constr_type_t:
            break;
          case Constr_type_tMarg:
            break;
          case Constr_type_tOff:
            break;
          case Constr_type_userA:
            break;
          case Constr_type_userB:
            break;
          case Constr_type_userC:
            break;
          case Constr_type_userD:
            break;
          case Constr_type_userE:
            break;
          case Constr_type_userF:
            break;
          case Constr_type_userG:
            break;
          case Constr_type_userH:
            break;
          case Constr_type_userI:
            break;
          case Constr_type_userJ:
            break;
          case Constr_type_userK:
            break;
          case Constr_type_userL:
            break;
          case Constr_type_userM:
            break;
          case Constr_type_userN:
            break;
          case Constr_type_userO:
            break;
          case Constr_type_userP:
            break;
          case Constr_type_userQ:
            break;
          case Constr_type_userR:
            break;
          case Constr_type_userS:
            break;
          case Constr_type_userT:
            break;
          case Constr_type_userU:
            break;
          case Constr_type_userV:
            break;
          case Constr_type_userW:
            break;
          case Constr_type_userX:
            break;
          case Constr_type_userY:
            break;
          case Constr_type_userZ:
            break;
          case Constr_type_w:
            //return shape.setResizeWidthConstr;
            break;
          case Constr_type_wArH:
            break;
          case Constr_type_wOff:
            break;
          default:
            return;
        }
      }
    }

    changesFactory[AscDFH.historyitem_PresOfExtLst] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_PresOfExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };

    function PresOf() {
      IteratorAttributes.call(this);
      this.extLst = null;
    }

    InitClass(PresOf, IteratorAttributes, AscDFH.historyitem_type_PresOf);

    PresOf.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PresOfExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    PresOf.prototype.getExtLst = function () {
      return this.extLst;
    }

    PresOf.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.axis.length; ++nIdx) {
        oCopy.addToLstAxis(nIdx, this.axis[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.cnt.length; ++nIdx) {
        oCopy.addToLstCnt(nIdx, this.cnt[nIdx]);
      }
      for (nIdx = 0; nIdx < this.hideLastTrans.length; ++nIdx) {
        oCopy.addToLstHideLastTrans(nIdx, this.hideLastTrans[nIdx]);
      }
      for (nIdx = 0; nIdx < this.ptType.length; ++nIdx) {
        oCopy.addToLstPtType(nIdx, this.ptType[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.st.length; ++nIdx) {
        oCopy.addToLstSt(nIdx, this.st[nIdx]);
      }
      for (nIdx = 0; nIdx < this.step.length; ++nIdx) {
        oCopy.addToLstStep(nIdx, this.step[nIdx]);
      }
    }

    PresOf.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
      for (var i = 0; i < this.st.length; i += 1) {
        pWriter._WriteInt1(1, this.st[i]);
      }
      for (i = 0; i < this.step.length; i += 1) {
        pWriter._WriteInt1(2, this.step[i]);
      }
      for (i = 0; i < this.hideLastTrans.length; i += 1) {
        pWriter._WriteBool1(3, this.hideLastTrans[i]);
      }
      for (i = 0; i < this.cnt.length; i += 1) {
        pWriter._WriteInt1(4, this.cnt[i]);
      }
      for (i = 0; i < this.axis.length; i += 1) {
        pWriter._WriteUChar1(5, this.axis[i].getVal());
      }
      for (i = 0; i < this.ptType.length; i += 1) {
        pWriter._WriteUChar1(6, this.ptType[i].getVal());
      }
    };
    PresOf.prototype.writeChildren = function(pWriter) {
    };
    PresOf.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
      else if (1 === nType) this.addToLstSt(this.st.length, oStream.GetLong());
      else if (2 === nType) this.addToLstStep(this.step.length, oStream.GetLong());
      else if (3 === nType) this.addToLstHideLastTrans(this.hideLastTrans.length, oStream.GetBool());
      else if (4 === nType) this.addToLstCnt(this.cnt.length, oStream.GetLong());
      else if (5 === nType) {
        var axis = new AxisType();
        this.addToLstAxis(this.axis.length, axis);
        axis.setVal(oStream.GetUChar());
      }
      else if (6 === nType) {
        var ptType = new ElementType();
        this.addToLstPtType(this.ptType.length, ptType);
        ptType.setVal(oStream.GetUChar());
      }
    };

    PresOf.prototype.readChild = function(nType, pReader) {
    };

    function RuleLst() {
      CCommonDataList.call(this);
    }

    InitClass(RuleLst, CCommonDataList, AscDFH.historyitem_type_RuleLst);

    RuleLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          var oChild = new Rule();
          oChild.fromPPTY(pReader);
          this.addToLst(this.list.length, oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    changesFactory[AscDFH.historyitem_RuleFact] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_RuleFor] = CChangeLong;
    changesFactory[AscDFH.historyitem_RuleForName] = CChangeString;
    changesFactory[AscDFH.historyitem_RuleMax] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_RuleType] = CChangeLong;
    changesFactory[AscDFH.historyitem_RuleVal] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_RuleExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_RulePtType] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_RuleFact] = function (oClass, value) {
      oClass.fact = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RuleFor] = function (oClass, value) {
      oClass.for = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RuleForName] = function (oClass, value) {
      oClass.forName = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RuleMax] = function (oClass, value) {
      oClass.max = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RuleType] = function (oClass, value) {
      oClass.type = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RuleVal] = function (oClass, value) {
      oClass.val = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RuleExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RulePtType] = function (oClass, value) {
      oClass.ptType = value;
    };

    function Rule() {
      CBaseFormatObject.call(this);
      this.fact = null;
      this.for = null;
      this.forName = null;
      this.max = null;
      this.type = null;
      this.val = null;
      this.extLst = null;
      this.setPtType(new ElementType());
    }

    InitClass(Rule, CBaseFormatObject, AscDFH.historyitem_type_Rule);

    Rule.prototype.setFact = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_RuleFact, this.getFact(), pr));
      this.fact = pr;
    }

    Rule.prototype.setFor = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_RuleFor, this.getFor(), pr));
      this.for = pr;
    }

    Rule.prototype.setForName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_RuleForName, this.getForName(), pr));
      this.forName = pr;
    }

    Rule.prototype.setMax = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_RuleMax, this.getMax(), pr));
      this.max = pr;
    }

    Rule.prototype.setType = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_RuleType, this.getType(), pr));
      this.type = pr;
    }

    Rule.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_RuleVal, this.getVal(), pr));
      this.val = pr;
    }

    Rule.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_RuleExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    Rule.prototype.setPtType = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_RulePtType, this.getPtType(), oPr));
      this.ptType = oPr;
      this.setParentToChild(oPr);
    }

    Rule.prototype.getFact = function () {
      return this.fact;
    }

    Rule.prototype.getFor = function () {
      return this.for;
    }

    Rule.prototype.getForName = function () {
      return this.forName;
    }

    Rule.prototype.getMax = function () {
      return this.max;
    }

    Rule.prototype.getType = function () {
      return this.type;
    }

    Rule.prototype.getVal = function () {
      return this.val;
    }

    Rule.prototype.getExtLst = function () {
      return this.extLst;
    }

    Rule.prototype.getPtType = function () {
      return this.ptType;
    }

    Rule.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setFact(this.getFact());
      oCopy.setFor(this.getFor());
      oCopy.setForName(this.getForName());
      oCopy.setMax(this.getMax());
      oCopy.setType(this.getType());
      oCopy.setVal(this.getVal());
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getPtType()) {
        oCopy.setPtType(this.getPtType().createDuplicate(oIdMap));
      }
    }

    Rule.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteDoubleReal2(0, this.fact);
      pWriter._WriteUChar2(1, this.for);
      pWriter._WriteString2(2, this.forName);
      pWriter._WriteUChar2(3, this.ptType.getVal());
      pWriter._WriteUChar2(4, this.type);
      pWriter._WriteDoubleReal2(5, this.val);
      pWriter._WriteDoubleReal2(6, this.max);
    };
    Rule.prototype.writeChildren = function(pWriter) {
    };
    Rule.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setFact(oStream.GetDouble());
      else if (1 === nType) this.setFor(oStream.GetUChar());
      else if (2 === nType) this.setForName(oStream.GetString2());
      else if (3 === nType) {
        var pt = new ElementType();
        pt.setVal(oStream.GetUChar());
        this.setPtType(pt);
      }
      else if (4 === nType) this.setType(oStream.GetUChar());
      else if (5 === nType) this.setVal(oStream.GetDouble());
      else if (6 === nType) this.setMax(oStream.GetDouble());
    };
    Rule.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_SShapeBlip] = CChangeString;
    changesFactory[AscDFH.historyitem_SShapeBlipPhldr] = CChangeBool;
    changesFactory[AscDFH.historyitem_SShapeHideGeom] = CChangeBool;
    changesFactory[AscDFH.historyitem_SShapeLkTxEntry] = CChangeBool;
    changesFactory[AscDFH.historyitem_SShapeRot] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_SShapeType] = CChangeString;
    changesFactory[AscDFH.historyitem_SShapeZOrderOff] = CChangeLong;
    changesFactory[AscDFH.historyitem_SShapeAdjLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_SShapeExtLst] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_SShapeBlip] = function (oClass, value) {
      oClass.blip = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SShapeBlipPhldr] = function (oClass, value) {
      oClass.blipPhldr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SShapeHideGeom] = function (oClass, value) {
      oClass.hideGeom = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SShapeLkTxEntry] = function (oClass, value) {
      oClass.lkTxEntry = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SShapeRot] = function (oClass, value) {
      oClass.rot = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SShapeType] = function (oClass, value) {
      oClass.type = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SShapeZOrderOff] = function (oClass, value) {
      oClass.zOrderOff = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SShapeAdjLst] = function (oClass, value) {
      oClass.adjLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SShapeExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };

    function SShape() {
      CBaseFormatObject.call(this);
      this.blip = null;
      this.blipPhldr = null;
      this.hideGeom = null;
      this.lkTxEntry = null;
      this.rot = null;
      this.type = null;
      this.zOrderOff = null;
      this.adjLst = null;
      this.extLst = null;
    }

    InitClass(SShape, CBaseFormatObject, AscDFH.historyitem_type_SShape);

    SShape.prototype.setBlip = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_SShapeBlip, this.getBlip(), pr));
      this.blip = pr;
    }

    SShape.prototype.setBlipPhldr = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_SShapeBlipPhldr, this.getBlipPhldr(), pr));
      this.blipPhldr = pr;
    }

    SShape.prototype.setHideGeom = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_SShapeHideGeom, this.getHideGeom(), pr));
      this.hideGeom = pr;
    }

    SShape.prototype.setLkTxEntry = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_SShapeLkTxEntry, this.getLkTxEntry(), pr));
      this.lkTxEntry = pr;
    }

    SShape.prototype.setRot = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_SShapeRot, this.getRot(), pr));
      this.rot = pr;
    }

    SShape.prototype.setType = function (oPr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_SShapeType, this.getType(), oPr));
      this.type = oPr;
    }

    SShape.prototype.setZOrderOff = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_SShapeZOrderOff, this.getZOrderOff(), pr));
      this.zOrderOff = pr;
    }

    SShape.prototype.setAdjLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SShapeAdjLst, this.getAdjLst(), oPr));
      this.adjLst = oPr;
      this.setParentToChild(oPr);
    }

    SShape.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SShapeExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    SShape.prototype.getBlip = function () {
      return this.blip;
    }

    SShape.prototype.getBlipPhldr = function () {
      return this.blipPhldr;
    }

    SShape.prototype.getHideGeom = function () {
      return this.hideGeom;
    }

    SShape.prototype.getLkTxEntry = function () {
      return this.lkTxEntry;
    }

    SShape.prototype.getRot = function () {
      return this.rot;
    }

    SShape.prototype.getType = function () {
      return this.type;
    }

    SShape.prototype.getZOrderOff = function () {
      return this.zOrderOff;
    }

    SShape.prototype.getAdjLst = function () {
      return this.adjLst;
    }

    SShape.prototype.getExtLst = function () {
      return this.extLst;
    }

    SShape.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setBlip(this.getBlip());
      oCopy.setBlipPhldr(this.getBlipPhldr());
      oCopy.setHideGeom(this.getHideGeom());
      oCopy.setLkTxEntry(this.getLkTxEntry());
      oCopy.setRot(this.getRot());
      oCopy.setType(this.getType());
      oCopy.setZOrderOff(this.getZOrderOff());
      if (this.getAdjLst()) {
        oCopy.setAdjLst(this.getAdjLst().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
    }
    SShape.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.blip);
      pWriter._WriteBool2(1, this.blipPhldr);
      pWriter._WriteBool2(2, this.hideGeom);
      pWriter._WriteBool2(3, this.lkTxEntry);
      pWriter._WriteDoubleReal2(4, this.rot);
      pWriter._WriteInt2(5, this.zOrderOff);
      pWriter._WriteString2(6, this.type);

    };
    SShape.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.adjLst);
    };
    SShape.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setBlip(oStream.GetString2());
      else if (1 === nType) this.setBlipPhldr(oStream.GetBool());
      else if (2 === nType) this.setHideGeom(oStream.GetBool());
      else if (3 === nType) this.setLkTxEntry(oStream.GetBool());
      else if (4 === nType) this.setRot(oStream.GetDouble());
      else if (5 === nType) this.setZOrderOff(oStream.GetULong());
      else if (6 === nType) this.setType(oStream.GetString2());
    };
    SShape.prototype.readChild = function(nType, pReader) {
      switch (nType) {
        case 0: {
          var oLst = new AdjLst();
          oLst.fromPPTY(pReader);
          this.setAdjLst(oLst);
          break;
        }
      }
    };

    function AdjLst() {
      CCommonDataList.call(this);
    }

    InitClass(AdjLst, CCommonDataList, AscDFH.historyitem_type_AdjLst);

    AdjLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          var oChild = new Adj();
          oChild.fromPPTY(pReader);
          this.addToLst(this.list.length, oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    changesFactory[AscDFH.historyitem_AdjIdx] = CChangeLong;
    changesFactory[AscDFH.historyitem_AdjVal] = CChangeDouble2;
    drawingsChangesMap[AscDFH.historyitem_AdjIdx] = function (oClass, value) {
      oClass.idx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_AdjVal] = function (oClass, value) {
      oClass.val = value;
    };

    function Adj() {
      CBaseFormatObject.call(this);
      this.idx = null;
      this.val = null;
    }

    InitClass(Adj, CBaseFormatObject, AscDFH.historyitem_type_Adj);

    Adj.prototype.setIdx = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AdjIdx, this.getIdx(), pr));
      this.idx = pr;
    }

    Adj.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_AdjVal, this.getVal(), pr));
      this.val = pr;
    }

    Adj.prototype.getIdx = function () {
      return this.idx;
    }

    Adj.prototype.getVal = function () {
      return this.val;
    }

    Adj.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setIdx(this.getIdx());
      oCopy.setVal(this.getVal());
    }

    Adj.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUInt2(0, this.idx);
      pWriter._WriteDoubleReal2(1, this.val);
    };
    Adj.prototype.writeChildren = function(pWriter) {
    };
    Adj.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setIdx(oStream.GetULong());
      else if (1 === nType) this.setVal(oStream.GetDouble());
    };
    Adj.prototype.readChild = function(nType, pReader) {
    };


    changesFactory[AscDFH.historyitem_VarLstAnimLvl] = CChangeObject;
    changesFactory[AscDFH.historyitem_VarLstAnimOne] = CChangeObject;
    changesFactory[AscDFH.historyitem_VarLstBulletEnabled] = CChangeObject;
    changesFactory[AscDFH.historyitem_VarLstChMax] = CChangeObject;
    changesFactory[AscDFH.historyitem_VarLstChPref] = CChangeObject;
    changesFactory[AscDFH.historyitem_VarLstDir] = CChangeObject;
    changesFactory[AscDFH.historyitem_VarLstHierBranch] = CChangeObject;
    changesFactory[AscDFH.historyitem_VarLstOrgChart] = CChangeObject;
    changesFactory[AscDFH.historyitem_VarLstResizeHandles] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_VarLstAnimLvl] = function (oClass, value) {
      oClass.animLvl = value;
    };
    drawingsChangesMap[AscDFH.historyitem_VarLstAnimOne] = function (oClass, value) {
      oClass.animOne = value;
    };
    drawingsChangesMap[AscDFH.historyitem_VarLstBulletEnabled] = function (oClass, value) {
      oClass.bulletEnabled = value;
    };
    drawingsChangesMap[AscDFH.historyitem_VarLstChMax] = function (oClass, value) {
      oClass.chMax = value;
    };
    drawingsChangesMap[AscDFH.historyitem_VarLstChPref] = function (oClass, value) {
      oClass.chPref = value;
    };
    drawingsChangesMap[AscDFH.historyitem_VarLstDir] = function (oClass, value) {
      oClass.dir = value;
    };
    drawingsChangesMap[AscDFH.historyitem_VarLstHierBranch] = function (oClass, value) {
      oClass.hierBranch = value;
    };
    drawingsChangesMap[AscDFH.historyitem_VarLstOrgChart] = function (oClass, value) {
      oClass.orgChart = value;
    };
    drawingsChangesMap[AscDFH.historyitem_VarLstResizeHandles] = function (oClass, value) {
      oClass.resizeHandles = value;
    };

    function VarLst() {
      CBaseFormatObject.call(this);
      this.animLvl = null;
      this.animOne = null;
      this.bulletEnabled = null;
      this.chMax = null;
      this.chPref = null;
      this.dir = null;
      this.hierBranch = null;
      this.orgChart = null;
      this.resizeHandles = null;
    }

    InitClass(VarLst, CBaseFormatObject, AscDFH.historyitem_type_VarLst);

    VarLst.prototype.setAnimLvl = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstAnimLvl, this.getAnimLvl(), oPr));
      this.animLvl = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.setAnimOne = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstAnimOne, this.getAnimOne(), oPr));
      this.animOne = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.setBulletEnabled = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstBulletEnabled, this.getBulletEnabled(), oPr));
      this.bulletEnabled = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.setChMax = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstChMax, this.getChMax(), oPr));
      this.chMax = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.setChPref = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstChPref, this.getChPref(), oPr));
      this.chPref = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.setDir = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstDir, this.getDir(), oPr));
      this.dir = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.setHierBranch = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstHierBranch, this.getHierBranch(), oPr));
      this.hierBranch = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.setOrgChart = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstOrgChart, this.getOrgChart(), oPr));
      this.orgChart = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.setResizeHandles = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_VarLstResizeHandles, this.getResizeHandles(), oPr));
      this.resizeHandles = oPr;
      this.setParentToChild(oPr);
    }

    VarLst.prototype.getAnimLvl = function () {
      return this.animLvl;
    }

    VarLst.prototype.getAnimOne = function () {
      return this.animOne;
    }

    VarLst.prototype.getBulletEnabled = function () {
      return this.bulletEnabled;
    }

    VarLst.prototype.getChMax = function () {
      return this.chMax;
    }

    VarLst.prototype.getChPref = function () {
      return this.chPref;
    }

    VarLst.prototype.getDir = function () {
      return this.dir;
    }

    VarLst.prototype.getHierBranch = function () {
      return this.hierBranch;
    }

    VarLst.prototype.getOrgChart = function () {
      return this.orgChart;
    }

    VarLst.prototype.getResizeHandles = function () {
      return this.resizeHandles;
    }

    VarLst.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getAnimLvl()) {
        oCopy.setAnimLvl(this.getAnimLvl().createDuplicate(oIdMap));
      }
      if (this.getAnimOne()) {
        oCopy.setAnimOne(this.getAnimOne().createDuplicate(oIdMap));
      }
      if (this.getBulletEnabled()) {
        oCopy.setBulletEnabled(this.getBulletEnabled().createDuplicate(oIdMap));
      }
      if (this.getChMax()) {
        oCopy.setChMax(this.getChMax().createDuplicate(oIdMap));
      }
      if (this.getChPref()) {
        oCopy.setChPref(this.getChPref().createDuplicate(oIdMap));
      }
      if (this.getDir()) {
        oCopy.setDir(this.getDir().createDuplicate(oIdMap));
      }
      if (this.getHierBranch()) {
        oCopy.setHierBranch(this.getHierBranch().createDuplicate(oIdMap));
      }
      if (this.getOrgChart()) {
        oCopy.setOrgChart(this.getOrgChart().createDuplicate(oIdMap));
      }
      if (this.getResizeHandles()) {
        oCopy.setResizeHandles(this.getResizeHandles().createDuplicate(oIdMap));
      }
    }

    VarLst.prototype.privateWriteAttributes = null;
    VarLst.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.animLvl);
      this.writeRecord2(pWriter, 1, this.animOne);
      this.writeRecord2(pWriter, 2, this.bulletEnabled);
      this.writeRecord2(pWriter, 3, this.chMax);
      this.writeRecord2(pWriter, 4, this.chPref);
      this.writeRecord2(pWriter, 5, this.dir);
      this.writeRecord2(pWriter, 6, this.hierBranch);
      this.writeRecord2(pWriter, 7, this.orgChart);
      this.writeRecord2(pWriter, 8, this.resizeHandles);
    };
    VarLst.prototype.readAttribute = null;
    VarLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setAnimLvl(new AnimLvl());
          this.animLvl.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setAnimOne(new AnimOne());
          this.animOne.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setBulletEnabled(new BulletEnabled());
          this.bulletEnabled.fromPPTY(pReader);
          break;
        }
        case 3: {
          this.setChMax(new ChMax());
          this.chMax.fromPPTY(pReader);
          break;
        }
        case 4: {
          this.setChPref(new ChPref());
          this.chPref.fromPPTY(pReader);
          break;
        }
        case 5: {
          this.setDir(new DiagramDirection());
          this.dir.fromPPTY(pReader);
          break;
        }
        case 6: {
          this.setHierBranch(new HierBranch());
          this.hierBranch.fromPPTY(pReader);
          break;
        }
        case 7: {
          this.setOrgChart(new OrgChart());
          this.orgChart.fromPPTY(pReader);
          break;
        }
        case 8: {
          this.setResizeHandles(new ResizeHandles());
          this.resizeHandles.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    VarLst.prototype.getVal = function (fieldType) {
      var defaultValues = {
        'animLvl': 'none',
        'animOne': 'one',
        'bulletEnabled': 'false',
        'chMax': '-1',
        'chPref': '-1',
        'dir': 'norm',
        'hierBranch': 'std',
        'orgChart': 'false',
        'resizeHandles': 'rel',
      };
      var nonDefaultValue = this[fieldType] && this[fieldType].getVal();
      return nonDefaultValue ? nonDefaultValue : defaultValues[fieldType];
    }

    VarLst.prototype.getChildren = function() {
      return [this.animLvl, this.animOne, this.bulletEnabled, this.chMax, this.chPref, this.dir, this.hierBranch, this.orgChart, this.resizeHandles];
    };


    changesFactory[AscDFH.historyitem_AnimLvlVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_AnimLvlVal] = function (oClass, value) {
      oClass.val = value;
    };

    function AnimLvl() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(AnimLvl, CBaseFormatObject, AscDFH.historyitem_type_AnimLvl);

    AnimLvl.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimLvlVal, this.getVal(), pr));
      this.val = pr;
    }

    AnimLvl.prototype.getVal = function () {
      return this.val;
    }

    AnimLvl.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }
    AnimLvl.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.val !== null ? this.val : 0);
    };
    AnimLvl.prototype.writeChildren = function(pWriter) {
    };
    AnimLvl.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) {
        var nVal = oStream.GetUChar();
        if(nVal !== 0) {
          this.setVal(nVal);
        }
      }
    };
    AnimLvl.prototype.readChild = function(nType, pReader) {

    };

    changesFactory[AscDFH.historyitem_AnimOneVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_AnimOneVal] = function (oClass, value) {
      oClass.val = value;
    };

    function AnimOne() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(AnimOne, CBaseFormatObject, AscDFH.historyitem_type_AnimOne);

    AnimOne.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_AnimOneVal, this.getVal(), pr));
      this.val = pr;
    }

    AnimOne.prototype.getVal = function () {
      return this.val;
    }

    AnimOne.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }
    AnimOne.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.val !== null ? this.val : 0);
    };
    AnimOne.prototype.writeChildren = function(pWriter) {
    };
    AnimOne.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) {
        var nVal = oStream.GetUChar();
        if(nVal !== 0) {
          this.setVal(nVal);
        }
      }
    };
    AnimOne.prototype.readChild = function(nType, pReader) {

    };

    changesFactory[AscDFH.historyitem_BulletEnabledVal] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_BulletEnabledVal] = function (oClass, value) {
      oClass.val = value;
    };

    function BulletEnabled() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(BulletEnabled, CBaseFormatObject, AscDFH.historyitem_type_BulletEnabled);

    BulletEnabled.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_BulletEnabledVal, this.getVal(), pr));
      this.val = pr;
    }

    BulletEnabled.prototype.getVal = function () {
      return this.val;
    }

    BulletEnabled.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }
    BulletEnabled.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.val !== null ? (this.val ? 1 : 0) : 0);
    };
    BulletEnabled.prototype.writeChildren = function(pWriter) {
    };
    BulletEnabled.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) {
        var nVal = oStream.GetUChar();
        this.setVal(nVal ? true : false);
      }
    };
    BulletEnabled.prototype.readChild = function(nType, pReader) {

    };

    changesFactory[AscDFH.historyitem_ChMaxVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_ChMaxVal] = function (oClass, value) {
      oClass.val = value;
    };

    function ChMax() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(ChMax, CBaseFormatObject, AscDFH.historyitem_type_ChMax);

    ChMax.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ChMaxVal, this.getVal(), pr));
      this.val = pr;
    }

    ChMax.prototype.getVal = function () {
      return this.val;
    }

    ChMax.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }

    changesFactory[AscDFH.historyitem_ChPrefVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_ChPrefVal] = function (oClass, value) {
      oClass.val = value;
    };

    function ChPref() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(ChPref, CBaseFormatObject, AscDFH.historyitem_type_ChPref);

    ChPref.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ChPrefVal, this.getVal(), pr));
      this.val = pr;
    }

    ChPref.prototype.getVal = function () {
      return this.val;
    }

    ChPref.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }

    changesFactory[AscDFH.historyitem_DiagramDirectionVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_DiagramDirectionVal] = function (oClass, value) {
      oClass.val = value;
    };

    function DiagramDirection() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(DiagramDirection, CBaseFormatObject, AscDFH.historyitem_type_DiagramDirection);

    DiagramDirection.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_DiagramDirectionVal, this.getVal(), pr));
      this.val = pr;
    }

    DiagramDirection.prototype.getVal = function () {
      return this.val;
    }

    DiagramDirection.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }
    DiagramDirection.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.val !== null ? this.val : 0);
    };
    DiagramDirection.prototype.writeChildren = function(pWriter) {
    };
    DiagramDirection.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) {
        var nVal = oStream.GetUChar();
        if(nVal !== 0) {
          this.setVal(nVal);
        }
      }
    };
    DiagramDirection.prototype.readChild = function(nType, pReader) {

    };

    changesFactory[AscDFH.historyitem_HierBranchVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_HierBranchVal] = function (oClass, value) {
      oClass.val = value;
    };

    function HierBranch() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(HierBranch, CBaseFormatObject, AscDFH.historyitem_type_HierBranch);

    HierBranch.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_HierBranchVal, this.getVal(), pr));
      this.val = pr;
    }

    HierBranch.prototype.getVal = function () {
      return this.val;
    }

    HierBranch.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    };
    HierBranch.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.val !== null ? this.val : 0);
    };
    HierBranch.prototype.writeChildren = function(pWriter) {
    };
    HierBranch.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) {
        var nVal = oStream.GetUChar();
        if(nVal !== 0) {
          this.setVal(nVal);
        }
      }
    };
    HierBranch.prototype.readChild = function(nType, pReader) {

    };

    changesFactory[AscDFH.historyitem_OrgChartVal] = CChangeBool;
    drawingsChangesMap[AscDFH.historyitem_OrgChartVal] = function (oClass, value) {
      oClass.val = value;
    };

    function OrgChart() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(OrgChart, CBaseFormatObject, AscDFH.historyitem_type_OrgChart);

    OrgChart.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_OrgChartVal, this.getVal(), pr));
      this.val = pr;
    }

    OrgChart.prototype.getVal = function () {
      return this.val;
    }

    OrgChart.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }
    OrgChart.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.val !== null ? (this.val ? 1 : 0) : 0);
    };
    OrgChart.prototype.writeChildren = function(pWriter) {
    };
    OrgChart.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) {
        var nVal = oStream.GetUChar();
        this.setVal(nVal ? true : false);
      }
    };
    OrgChart.prototype.readChild = function(nType, pReader) {

    };
    changesFactory[AscDFH.historyitem_ResizeHandlesVal] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_ResizeHandlesVal] = function (oClass, value) {
      oClass.val = value;
    };

    function ResizeHandles() {
      CBaseFormatObject.call(this);
      this.val = null;
    }

    InitClass(ResizeHandles, CBaseFormatObject, AscDFH.historyitem_type_ResizeHandles);

    ResizeHandles.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ResizeHandlesVal, this.getVal(), pr));
      this.val = pr;
    }

    ResizeHandles.prototype.getVal = function () {
      return this.val;
    }

    ResizeHandles.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setVal(this.getVal());
    }
    ResizeHandles.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.val !== null ? this.val : 0);
    };
    ResizeHandles.prototype.writeChildren = function(pWriter) {
    };
    ResizeHandles.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) {
        var nVal = oStream.GetUChar();
        if(nVal !== 0) {
          this.setVal(nVal);
        }
      }
    };
    ResizeHandles.prototype.readChild = function(nType, pReader) {

    };


    changesFactory[AscDFH.historyitem_ForEachName] = CChangeString;
    changesFactory[AscDFH.historyitem_ForEachRef] = CChangeString;
    changesFactory[AscDFH.historyitem_ForEachAddList] = CChangeContent;
    changesFactory[AscDFH.historyitem_ForEachRemoveList] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_ForEachName] = function (oClass, value) {
      oClass.name = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ForEachRef] = function (oClass, value) {
      oClass.ref = value;
    };
    drawingContentChanges[AscDFH.historyitem_ForEachAddList] = function (oClass) {
      return oClass.list;
    };
    drawingContentChanges[AscDFH.historyitem_ForEachRemoveList] = function (oClass) {
      return oClass.list;
    };

    function ForEach() {
      IteratorAttributes.call(this);
      this.name = null;
      this.ref = null;
      this.list = [];
    }

    InitClass(ForEach, IteratorAttributes, AscDFH.historyitem_type_ForEach);

    ForEach.prototype.setName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ForEachName, this.getName(), pr));
      this.name = pr;
    }

    ForEach.prototype.setRef = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ForEachRef, this.getRef(), pr));
      this.ref = pr;
    }

    ForEach.prototype.addToLstList = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ForEachAddList, nInsertIdx, [oPr], true));
      this.list.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    ForEach.prototype.removeFromLstList = function (nIdx) {
      if (nIdx > -1 && nIdx < this.list.length) {
        this.list[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ForEachRemoveList, nIdx, [this.list[nIdx]], false));
        this.list.splice(nIdx, 1);
      }
    };

    ForEach.prototype.getName = function () {
      return this.name;
    }

    ForEach.prototype.getRef = function () {
      return this.ref;
    }

    ForEach.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setName(this.getName());
      oCopy.setRef(this.getRef());
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        oCopy.addToLstList(nIdx, this.list[nIdx].createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.axis.length; ++nIdx) {
        oCopy.addToLstAxis(nIdx, this.axis[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.cnt.length; ++nIdx) {
        oCopy.addToLstCnt(nIdx, this.cnt[nIdx]);
      }
      for (nIdx = 0; nIdx < this.hideLastTrans.length; ++nIdx) {
        oCopy.addToLstHideLastTrans(nIdx, this.hideLastTrans[nIdx]);
      }
      for (nIdx = 0; nIdx < this.ptType.length; ++nIdx) {
        oCopy.addToLstPtType(nIdx, this.ptType[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.st.length; ++nIdx) {
        oCopy.addToLstSt(nIdx, this.st[nIdx]);
      }
      for (nIdx = 0; nIdx < this.step.length; ++nIdx) {
        oCopy.addToLstStep(nIdx, this.step[nIdx]);
      }
    }

    ForEach.prototype.readElement = function(pReader, nType) {
      var oElement = null;
      switch(nType) {
        case 0xb1: oElement = new Alg(); break;
        case 0xb2: oElement = new Choose(); break;
        case 0xb3: oElement = new ConstrLst(); break;
        case 0xb4: oElement = new ForEach(); break;
        case 0xb5: oElement = new LayoutNode(); break;
        case 0xb6: oElement = new PresOf(); break;
        case 0xb7: oElement = new RuleLst(); break;
        case 0xb8: oElement = new SShape(); break;
        case 0xb9: oElement = new VarLst(); break;
        default:break;
      }
      if(oElement) {
        oElement.fromPPTY(pReader);
        this.addToLstList(this.list.length, oElement);
      }
    };

    ForEach.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
      for (var i = 0; i < this.st.length; i += 1) {
        pWriter._WriteInt1(1, this.st[i]);
      }
      for (i = 0; i < this.step.length; i += 1) {
        pWriter._WriteInt1(2, this.step[i]);
      }
      for (i = 0; i < this.hideLastTrans.length; i += 1) {
        pWriter._WriteBool1(3, this.hideLastTrans[i]);
      }
      for (i = 0; i < this.cnt.length; i += 1) {
        pWriter._WriteInt1(4, this.cnt[i]);
      }
      for (i = 0; i < this.axis.length; i += 1) {
        pWriter._WriteUChar1(5, this.axis[i].getVal());
      }
      for (i = 0; i < this.ptType.length; i += 1) {
        pWriter._WriteUChar1(6, this.ptType[i].getVal());
      }
      pWriter._WriteString2(7, this.ref);
    };
    ForEach.prototype.writeChildren = function(pWriter) {
      for(var nIndex = 0; nIndex < this.list.length; ++nIndex) {
        var oElement = this.list[nIndex];
        switch (oElement.getObjectType()) {
          case AscDFH.historyitem_type_Alg: this.writeRecord2(pWriter, 0xb1, oElement); break;
          case AscDFH.historyitem_type_Choose: this.writeRecord2(pWriter, 0xb2, oElement); break;
          case AscDFH.historyitem_type_ConstrLst: this.writeRecord2(pWriter, 0xb3, oElement); break;
          case AscDFH.historyitem_type_ForEach: this.writeRecord2(pWriter, 0xb4, oElement); break;
          case AscDFH.historyitem_type_LayoutNode: this.writeRecord2(pWriter, 0xb5, oElement); break;
          case AscDFH.historyitem_type_PresOf: this.writeRecord2(pWriter, 0xb6, oElement); break;
          case AscDFH.historyitem_type_RuleLst: this.writeRecord2(pWriter, 0xb7, oElement); break;
          case AscDFH.historyitem_type_SShape: this.writeRecord2(pWriter, 0xb8, oElement); break;
          case AscDFH.historyitem_type_VarLst: this.writeRecord2(pWriter, 0xb9, oElement); break;
          default: break;
        }
      }
    };
    ForEach.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
      else if (1 === nType) this.addToLstSt(this.st.length, oStream.GetLong());
      else if (2 === nType) this.addToLstStep(this.step.length, oStream.GetLong());
      else if (3 === nType) this.addToLstHideLastTrans(this.hideLastTrans.length, oStream.GetBool());
      else if (4 === nType) this.addToLstCnt(this.cnt.length, oStream.GetLong());
      else if (5 === nType) {
        var axis = new AxisType();
        this.addToLstAxis(this.axis.length, axis);
        axis.setVal(oStream.GetUChar());
      }
      else if (6 === nType) {
        var ptType = new ElementType();
        this.addToLstPtType(this.ptType.length, ptType);
        ptType.setVal(oStream.GetUChar());
      }
      else if (7 === nType) this.setRef(oStream.GetString2());
    };

    ForEach.prototype.readChild = function(nType, pReader) {
      this.readElement(pReader, nType);
    };
    ForEach.prototype.getChildren = function() {
      return [].concat(this.list);
    };

    ForEach.prototype.startAlgorithm = function (pointTree, node) {
      this.list.forEach(function (element) {
        if (element instanceof AscFormat.ForEach || element instanceof AscFormat.LayoutNode || element instanceof AscFormat.Choose) {
          element.startAlgorithm(pointTree, node);
        }
      });
    }



    changesFactory[AscDFH.historyitem_SampDataUseDef] = CChangeBool;
    changesFactory[AscDFH.historyitem_SampDataDataModel] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_SampDataUseDef] = function (oClass, value) {
      oClass.useDef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SampDataDataModel] = function (oClass, value) {
      oClass.dataModel = value;
    };

    function SampData() {
      CBaseFormatObject.call(this);
      this.useDef = null;
      this.dataModel = null;
    }

    InitClass(SampData, CBaseFormatObject, AscDFH.historyitem_type_SampData);

    SampData.prototype.setUseDef = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_SampDataUseDef, this.getUseDef(), pr));
      this.useDef = pr;
    }

    SampData.prototype.setDataModel = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SampDataDataModel, this.getDataModel(), oPr));
      this.dataModel = oPr;
      this.setParentToChild(oPr);
    }

    SampData.prototype.getUseDef = function () {
      return this.useDef;
    }

    SampData.prototype.getDataModel = function () {
      return this.dataModel;
    }

    SampData.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setUseDef(this.getUseDef());
      if (this.getDataModel()) {
        oCopy.setDataModel(this.getDataModel().createDuplicate(oIdMap));
      }
    }

    SampData.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteBool2(0, this.useDef);
    };
    SampData.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.dataModel);
    };
    SampData.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setUseDef(oStream.GetBool());
    };
    SampData.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setDataModel(new DataModel());
          this.dataModel.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    SampData.prototype.getChildren = function() {
      return [this.dataModel];
    };


    changesFactory[AscDFH.historyitem_StyleDataUseDef] = CChangeBool;
    changesFactory[AscDFH.historyitem_StyleDataDataModel] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_StyleDataUseDef] = function (oClass, value) {
      oClass.useDef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDataDataModel] = function (oClass, value) {
      oClass.dataModel = value;
    };

    function StyleData() {
      CBaseFormatObject.call(this);
      this.useDef = null;
      this.dataModel = null;
    }

    InitClass(StyleData, CBaseFormatObject, AscDFH.historyitem_type_StyleData);

    StyleData.prototype.setUseDef = function (pr) {
      oHistory.Add(new CChangeBool(this, AscDFH.historyitem_StyleDataUseDef, this.getUseDef(), pr));
      this.useDef = pr;
    }

    StyleData.prototype.setDataModel = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDataDataModel, this.getDataModel(), oPr));
      this.dataModel = oPr;
      this.setParentToChild(oPr);
    }

    StyleData.prototype.getUseDef = function () {
      return this.useDef;
    }

    StyleData.prototype.getDataModel = function () {
      return this.dataModel;
    }

    StyleData.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setUseDef(this.getUseDef());
      if (this.getDataModel()) {
        oCopy.setDataModel(this.getDataModel().createDuplicate(oIdMap));
      }
    }

    StyleData.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteBool2(0, this.useDef);
    };
    StyleData.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.dataModel);
    };
    StyleData.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setUseDef(oStream.GetBool());
    };
    StyleData.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setDataModel(new DataModel());
          this.dataModel.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    StyleData.prototype.getChildren = function() {
      return [this.dataModel];
    };



    changesFactory[AscDFH.historyitem_DiagramTitleLang] = CChangeString;
    changesFactory[AscDFH.historyitem_DiagramTitleVal] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_DiagramTitleLang] = function (oClass, value) {
      oClass.lang = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DiagramTitleVal] = function (oClass, value) {
      oClass.val = value;
    };

    function DiagramTitle() {
      CBaseFormatObject.call(this);
      this.lang = null;
      this.val = null;
    }

    InitClass(DiagramTitle, CBaseFormatObject, AscDFH.historyitem_type_DiagramTitle);

    DiagramTitle.prototype.setLang = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_DiagramTitleLang, this.getLang(), pr));
      this.lang = pr;
    }

    DiagramTitle.prototype.setVal = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_DiagramTitleVal, this.getVal(), pr));
      this.val = pr;
    }

    DiagramTitle.prototype.getLang = function () {
      return this.lang;
    }

    DiagramTitle.prototype.getVal = function () {
      return this.val;
    }

    DiagramTitle.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setLang(this.getLang());
      oCopy.setVal(this.getVal());
    }

    DiagramTitle.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.lang);
      pWriter._WriteString2(1, this.val);
    };
    DiagramTitle.prototype.writeChildren = function(pWriter) {
    };
    DiagramTitle.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setLang(oStream.GetString2());
      else if (1 === nType) this.setVal(oStream.GetString2());
    };
    DiagramTitle.prototype.readChild = function(nType, pReader) {
    };



    function LayoutDefHdrLst() {
      CCommonDataList.call(this);
    }

    InitClass(LayoutDefHdrLst, CCommonDataList, AscDFH.historyitem_type_LayoutDefHdrLst);


    changesFactory[AscDFH.historyitem_LayoutDefHdrDefStyle] = CChangeString;
    changesFactory[AscDFH.historyitem_LayoutDefHdrMinVer] = CChangeString;
    changesFactory[AscDFH.historyitem_LayoutDefHdrResId] = CChangeLong;
    changesFactory[AscDFH.historyitem_LayoutDefHdrUniqueId] = CChangeString;
    changesFactory[AscDFH.historyitem_LayoutDefHdrCatLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefHdrExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_LayoutDefHdrAddTitle] = CChangeContent;
    changesFactory[AscDFH.historyitem_LayoutDefHdrRemoveTitle] = CChangeContent;
    changesFactory[AscDFH.historyitem_LayoutDefHdrAddDesc] = CChangeContent;
    changesFactory[AscDFH.historyitem_LayoutDefHdrRemoveDesc] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_LayoutDefHdrDefStyle] = function (oClass, value) {
      oClass.defStyle = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefHdrMinVer] = function (oClass, value) {
      oClass.minVer = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefHdrResId] = function (oClass, value) {
      oClass.resId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefHdrUniqueId] = function (oClass, value) {
      oClass.uniqueId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefHdrCatLst] = function (oClass, value) {
      oClass.catLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LayoutDefHdrExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingContentChanges[AscDFH.historyitem_LayoutDefHdrAddTitle] = function (oClass) {
      return oClass.title;
    };
    drawingContentChanges[AscDFH.historyitem_LayoutDefHdrRemoveTitle] = function (oClass) {
      return oClass.title;
    };
    drawingContentChanges[AscDFH.historyitem_LayoutDefHdrAddDesc] = function (oClass) {
      return oClass.desc;
    };
    drawingContentChanges[AscDFH.historyitem_LayoutDefHdrRemoveDesc] = function (oClass) {
      return oClass.desc;
    };

    function LayoutDefHdr() {
      CBaseFormatObject.call(this);
      this.defStyle = null;
      this.minVer = null;
      this.resId = null;
      this.uniqueId = null;
      this.catLst = null;
      this.extLst = null;
      this.title = [];
      this.desc = [];
    }

    InitClass(LayoutDefHdr, CBaseFormatObject, AscDFH.historyitem_type_LayoutDefHdr);

    LayoutDefHdr.prototype.setDefStyle = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutDefHdrDefStyle, this.getDefStyle(), pr));
      this.defStyle = pr;
    }

    LayoutDefHdr.prototype.setMinVer = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutDefHdrMinVer, this.getMinVer(), pr));
      this.minVer = pr;
    }

    LayoutDefHdr.prototype.setResId = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_LayoutDefHdrResId, this.getResId(), pr));
      this.resId = pr;
    }

    LayoutDefHdr.prototype.setUniqueId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_LayoutDefHdrUniqueId, this.getUniqueId(), pr));
      this.uniqueId = pr;
    }

    LayoutDefHdr.prototype.setCatLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefHdrCatLst, this.getCatLst(), oPr));
      this.catLst = oPr;
      this.setParentToChild(oPr);
    }

    LayoutDefHdr.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LayoutDefHdrExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    LayoutDefHdr.prototype.addToLstTitle = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.title.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_LayoutDefHdrAddTitle, nInsertIdx, [oPr], true));
      this.title.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    LayoutDefHdr.prototype.removeFromLstTitle = function (nIdx) {
      if (nIdx > -1 && nIdx < this.title.length) {
        this.title[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_LayoutDefHdrRemoveTitle, nIdx, [this.title[nIdx]], false));
        this.title.splice(nIdx, 1);
      }
    };

    LayoutDefHdr.prototype.addToLstDesc = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.desc.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_LayoutDefHdrAddDesc, nInsertIdx, [oPr], true));
      this.desc.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    LayoutDefHdr.prototype.removeFromLstDesc = function (nIdx) {
      if (nIdx > -1 && nIdx < this.desc.length) {
        this.desc[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_LayoutDefHdrRemoveDesc, nIdx, [this.desc[nIdx]], false));
        this.desc.splice(nIdx, 1);
      }
    };

    LayoutDefHdr.prototype.getDefStyle = function () {
      return this.defStyle;
    }

    LayoutDefHdr.prototype.getMinVer = function () {
      return this.minVer;
    }

    LayoutDefHdr.prototype.getResId = function () {
      return this.resId;
    }

    LayoutDefHdr.prototype.getUniqueId = function () {
      return this.uniqueId;
    }

    LayoutDefHdr.prototype.getCatLst = function () {
      return this.catLst;
    }

    LayoutDefHdr.prototype.getExtLst = function () {
      return this.extLst;
    }

    LayoutDefHdr.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setDefStyle(this.getDefStyle());
      oCopy.setMinVer(this.getMinVer());
      oCopy.setResId(this.getResId());
      oCopy.setUniqueId(this.getUniqueId());
      if (this.getCatLst()) {
        oCopy.setCatLst(this.getCatLst().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.title.length; ++nIdx) {
        oCopy.addToLstTitle(nIdx, this.title[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.desc.length; ++nIdx) {
        oCopy.addToLstDesc(nIdx, this.desc[nIdx].createDuplicate(oIdMap));
      }
    }

    changesFactory[AscDFH.historyitem_RelIdsCs] = CChangeString;
    changesFactory[AscDFH.historyitem_RelIdsDm] = CChangeString;
    changesFactory[AscDFH.historyitem_RelIdsLo] = CChangeString;
    changesFactory[AscDFH.historyitem_RelIdsQs] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_RelIdsCs] = function (oClass, value) {
      oClass.cs = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RelIdsDm] = function (oClass, value) {
      oClass.dm = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RelIdsLo] = function (oClass, value) {
      oClass.lo = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RelIdsQs] = function (oClass, value) {
      oClass.qs = value;
    };

    function RelIds() {
      CBaseFormatObject.call(this);
      this.cs = null;
      this.dm = null;
      this.lo = null;
      this.qs = null;
    }

    InitClass(RelIds, CBaseFormatObject, AscDFH.historyitem_type_RelIds);

    RelIds.prototype.setCs = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_RelIdsCs, this.getCs(), pr));
      this.cs = pr;
    }

    RelIds.prototype.setDm = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_RelIdsDm, this.getDm(), pr));
      this.dm = pr;
    }

    RelIds.prototype.setLo = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_RelIdsLo, this.getLo(), pr));
      this.lo = pr;
    }

    RelIds.prototype.setQs = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_RelIdsQs, this.getQs(), pr));
      this.qs = pr;
    }

    RelIds.prototype.getCs = function () {
      return this.cs;
    }

    RelIds.prototype.getDm = function () {
      return this.dm;
    }

    RelIds.prototype.getLo = function () {
      return this.lo;
    }

    RelIds.prototype.getQs = function () {
      return this.qs;
    }

    RelIds.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setCs(this.getCs());
      oCopy.setDm(this.getDm());
      oCopy.setLo(this.getLo());
      oCopy.setQs(this.getQs());
    }


    function CChangesDrawingsContentStyleLbl(Class, Type, Pos, Items, isAdd) {
      AscDFH.CChangesDrawingsContent.call(this, Class, Type, Pos, Items, isAdd);
    }

    CChangesDrawingsContentStyleLbl.prototype = Object.create(AscDFH.CChangesDrawingsContent.prototype);
    CChangesDrawingsContentStyleLbl.prototype.constructor = CChangesDrawingsContentStyleLbl;

    CChangesDrawingsContentStyleLbl.prototype.Undo = function () {
      AscDFH.CChangesDrawingsContent.prototype.Undo.call(this);
      this.updateStyleLbl();
    };

    CChangesDrawingsContentStyleLbl.prototype.Redo = function () {
      AscDFH.CChangesDrawingsContent.prototype.Redo.call(this);
      this.updateStyleLbl();
    };

    CChangesDrawingsContentStyleLbl.prototype.updateStyleLbl = function () {
      //console.log(this.Items);
      if (this.IsAdd()) {
        for (var i = 0; i < this.Items.length; i += 1) {
          this.Class.styleLblByName[this.Items[i].name] = this.Items[i];
        }
      } else {
        for (var i = 0; i < this.Items.length; i += 1) {
          delete this.Class.styleLblByName[this.Items[i].name];
        }
      }
    };

    changesFactory[AscDFH.historyitem_ColorsDefMinVer] = CChangeString;
    changesFactory[AscDFH.historyitem_ColorsDefUniqueId] = CChangeString;
    changesFactory[AscDFH.historyitem_ColorsDefCatLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorsDefExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorsDefDesc] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorsDefTitle] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorsDefAddStyleLbl] = CChangesDrawingsContentStyleLbl;
    changesFactory[AscDFH.historyitem_ColorsDefRemoveStyleLbl] = CChangesDrawingsContentStyleLbl;
    drawingsChangesMap[AscDFH.historyitem_ColorsDefMinVer] = function (oClass, value) {
      oClass.minVer = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefUniqueId] = function (oClass, value) {
      oClass.uniqueId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefCatLst] = function (oClass, value) {
      oClass.catLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefDesc] = function (oClass, value) {
      oClass.desc = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefTitle] = function (oClass, value) {
      oClass.title = value;
    };
    drawingContentChanges[AscDFH.historyitem_ColorsDefAddStyleLbl] = function (oClass) {
      return oClass.styleLbl;
    };
    drawingContentChanges[AscDFH.historyitem_ColorsDefRemoveStyleLbl] = function (oClass) {
      return oClass.styleLbl;
    };

    function ColorsDef() {
      CBaseFormatObject.call(this);
      this.minVer = null;
      this.uniqueId = null;
      this.catLst = null;
      this.extLst = null;
      this.desc = null;
      this.title = null;
      this.styleLbl = [];
      this.styleLblByName = {};
    }

    InitClass(ColorsDef, CBaseFormatObject, AscDFH.historyitem_type_ColorsDef);

    ColorsDef.prototype.setMinVer = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ColorsDefMinVer, this.getMinVer(), pr));
      this.minVer = pr;
    }

    ColorsDef.prototype.setUniqueId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ColorsDefUniqueId, this.getUniqueId(), pr));
      this.uniqueId = pr;
    }

    ColorsDef.prototype.setCatLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorsDefCatLst, this.getCatLst(), oPr));
      this.catLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorsDef.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorsDefExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorsDef.prototype.setDesc = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorsDefDesc, this.getDesc(), oPr));
      this.desc = oPr;
      this.setParentToChild(oPr);
    };

    ColorsDef.prototype.setTitle = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorsDefTitle, this.getTitle(), oPr));
      this.title = oPr;
      this.setParentToChild(oPr);
    };

    ColorsDef.prototype.addToLstStyleLbl = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.styleLbl.length, Math.max(0, nIdx));
      oHistory.Add(new CChangesDrawingsContentStyleLbl(this, AscDFH.historyitem_ColorsDefAddStyleLbl, nInsertIdx, [oPr], true));
      this.styleLbl.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
      this.styleLblByName[oPr.name] = oPr;
    };

    ColorsDef.prototype.removeFromLstStyleLbl = function (nIdx) {
      if (nIdx > -1 && nIdx < this.styleLbl.length) {
        this.styleLbl[nIdx].setParent(null);
        oHistory.Add(new CChangesDrawingsContentStyleLbl(this, AscDFH.historyitem_ColorsDefRemoveStyleLbl, nIdx, [this.styleLbl[nIdx]], false));
        var deleteObj = this.styleLbl.splice(nIdx, 1);
        delete this.styleLblByName[deleteObj[0].name];
      }
    };

    ColorsDef.prototype.getMinVer = function () {
      return this.minVer;
    }

    ColorsDef.prototype.getUniqueId = function () {
      return this.uniqueId;
    }

    ColorsDef.prototype.getCatLst = function () {
      return this.catLst;
    }

    ColorsDef.prototype.getExtLst = function () {
      return this.extLst;
    }

    ColorsDef.prototype.getDesc = function () {
      return this.desc;
    }

    ColorsDef.prototype.getTitle = function () {
      return this.title;
    }

    ColorsDef.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setMinVer(this.getMinVer());
      oCopy.setUniqueId(this.getUniqueId());
      if (this.getCatLst()) {
        oCopy.setCatLst(this.getCatLst().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getDesc()) {
        oCopy.setDesc(this.getDesc().createDuplicate(oIdMap));
      }
      if (this.getTitle()) {
        oCopy.setTitle(this.getTitle().createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.styleLbl.length; ++nIdx) {
        oCopy.addToLstStyleLbl(nIdx, this.styleLbl[nIdx].createDuplicate(oIdMap));
      }
    }

    ColorsDef.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.uniqueId);
      pWriter._WriteString2(1, this.minVer);
    };
    ColorsDef.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.title);
      this.writeRecord2(pWriter, 1, this.desc);
      this.writeRecord2(pWriter, 2, this.catLst);
      for(var nStyleLbl = 0; nStyleLbl < this.styleLbl.length; ++nStyleLbl) {
        this.writeRecord2(pWriter, 3, this.styleLbl[nStyleLbl]);
      }
    };
    ColorsDef.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setUniqueId(oStream.GetString2());
      else if (1 === nType) this.setMinVer(oStream.GetString2());
    };
    ColorsDef.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setTitle(new DiagramTitle());
          this.title.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setDesc(new Desc());
          this.desc.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setCatLst(new CatLst());
          this.catLst.fromPPTY(pReader);
          break;
        }
        case 3: {
          var oDefStyle = new ColorDefStyleLbl();
          oDefStyle.fromPPTY(pReader);
          this.addToLstStyleLbl(this.styleLbl.length, oDefStyle);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    ColorsDef.prototype.getChildren = function() {
      return [this.title, this.desc, this.catLst].concat(this.styleLbl);
    };


    changesFactory[AscDFH.historyitem_ColorDefStyleLblName] = CChangeString;
    changesFactory[AscDFH.historyitem_ColorDefStyleLblEffectClrLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorDefStyleLblExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorDefStyleLblFillClrLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorDefStyleLblLinClrLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorDefStyleLblTxEffectClrLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorDefStyleLblTxFillClrLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorDefStyleLblTxLinClrLst] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_ColorDefStyleLblName] = function (oClass, value) {
      oClass.name = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorDefStyleLblEffectClrLst] = function (oClass, value) {
      oClass.effectClrLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorDefStyleLblExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorDefStyleLblFillClrLst] = function (oClass, value) {
      oClass.fillClrLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorDefStyleLblLinClrLst] = function (oClass, value) {
      oClass.linClrLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorDefStyleLblTxEffectClrLst] = function (oClass, value) {
      oClass.txEffectClrLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorDefStyleLblTxFillClrLst] = function (oClass, value) {
      oClass.txFillClrLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorDefStyleLblTxLinClrLst] = function (oClass, value) {
      oClass.txLinClrLst = value;
    };

    function ColorDefStyleLbl() {
      CBaseFormatObject.call(this);
      this.name = null;
      this.effectClrLst = null;
      this.extLst = null;
      this.fillClrLst = null;
      this.linClrLst = null;
      this.txEffectClrLst = null;
      this.txFillClrLst = null;
      this.txLinClrLst = null;
    }

    InitClass(ColorDefStyleLbl, CBaseFormatObject, AscDFH.historyitem_type_ColorDefStyleLbl);

    ColorDefStyleLbl.prototype.setName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ColorDefStyleLblName, this.getName(), pr));
      this.name = pr;
    }

    ColorDefStyleLbl.prototype.setEffectClrLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorDefStyleLblEffectClrLst, this.getEffectClrLst(), oPr));
      this.effectClrLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorDefStyleLbl.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorDefStyleLblExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorDefStyleLbl.prototype.setFillClrLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorDefStyleLblFillClrLst, this.getFillClrLst(), oPr));
      this.fillClrLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorDefStyleLbl.prototype.setLinClrLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorDefStyleLblLinClrLst, this.getLinClrLst(), oPr));
      this.linClrLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorDefStyleLbl.prototype.setTxEffectClrLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorDefStyleLblTxEffectClrLst, this.getTxEffectClrLst(), oPr));
      this.txEffectClrLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorDefStyleLbl.prototype.setTxFillClrLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorDefStyleLblTxFillClrLst, this.getTxFillClrLst(), oPr));
      this.txFillClrLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorDefStyleLbl.prototype.setTxLinClrLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorDefStyleLblTxLinClrLst, this.getTxLinClrLst(), oPr));
      this.txLinClrLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorDefStyleLbl.prototype.getName = function () {
      return this.name;
    }

    ColorDefStyleLbl.prototype.getEffectClrLst = function () {
      return this.effectClrLst;
    }

    ColorDefStyleLbl.prototype.getExtLst = function () {
      return this.extLst;
    }

    ColorDefStyleLbl.prototype.getFillClrLst = function () {
      return this.fillClrLst;
    }

    ColorDefStyleLbl.prototype.getLinClrLst = function () {
      return this.linClrLst;
    }

    ColorDefStyleLbl.prototype.getTxEffectClrLst = function () {
      return this.txEffectClrLst;
    }

    ColorDefStyleLbl.prototype.getTxFillClrLst = function () {
      return this.txFillClrLst;
    }

    ColorDefStyleLbl.prototype.getTxLinClrLst = function () {
      return this.txLinClrLst;
    }

    ColorDefStyleLbl.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setName(this.getName());
      if (this.getEffectClrLst()) {
        oCopy.setEffectClrLst(this.getEffectClrLst().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getFillClrLst()) {
        oCopy.setFillClrLst(this.getFillClrLst().createDuplicate(oIdMap));
      }
      if (this.getLinClrLst()) {
        oCopy.setLinClrLst(this.getLinClrLst().createDuplicate(oIdMap));
      }
      if (this.getTxEffectClrLst()) {
        oCopy.setTxEffectClrLst(this.getTxEffectClrLst().createDuplicate(oIdMap));
      }
      if (this.getTxFillClrLst()) {
        oCopy.setTxFillClrLst(this.getTxFillClrLst().createDuplicate(oIdMap));
      }
      if (this.getTxLinClrLst()) {
        oCopy.setTxLinClrLst(this.getTxLinClrLst().createDuplicate(oIdMap));
      }
    }

    ColorDefStyleLbl.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
    };
    ColorDefStyleLbl.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.effectClrLst);
      this.writeRecord2(pWriter, 1, this.fillClrLst);
      this.writeRecord2(pWriter, 2, this.linClrLst);
      this.writeRecord2(pWriter, 3, this.txEffectClrLst);
      this.writeRecord2(pWriter, 4, this.txFillClrLst);
      this.writeRecord2(pWriter, 5, this.txLinClrLst);
    };
    ColorDefStyleLbl.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
    };
    ColorDefStyleLbl.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setEffectClrLst(new EffectClrLst());
          this.effectClrLst.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setFillClrLst(new FillClrLst());
          this.fillClrLst.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setLinClrLst(new LinClrLst());
          this.linClrLst.fromPPTY(pReader);
          break;
        }
        case 3: {
          this.setTxEffectClrLst(new TxEffectClrLst());
          this.txEffectClrLst.fromPPTY(pReader);
          break;
        }
        case 4: {
          this.setTxFillClrLst(new TxFillClrLst());
          this.txFillClrLst.fromPPTY(pReader);
          break;
        }
        case 5: {
          this.setTxLinClrLst(new TxLinClrLst());
          this.txLinClrLst.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    ColorDefStyleLbl.prototype.getChildren = function() {
      return [this.effectClrLst, this.fillClrLst, this.linClrLst, this.txEffectClrLst, this.txFillClrLst, this.txLinClrLst];
    };


    changesFactory[AscDFH.historyitem_CCommonDataClrListAdd] = CChangesContentNoId;
    changesFactory[AscDFH.historyitem_CCommonDataClrListRemove] = CChangesContentNoId;
    drawingConstructorsMap[AscDFH.historyitem_CCommonDataClrListAdd] = AscFormat.CUniColor;
    drawingConstructorsMap[AscDFH.historyitem_CCommonDataClrListRemove] = AscFormat.CUniColor;
    drawingContentChanges[AscDFH.historyitem_CCommonDataClrListAdd] = function (oClass) {
      return oClass.list;
    };
    drawingContentChanges[AscDFH.historyitem_CCommonDataClrListRemove] = function (oClass) {
      return oClass.list;
    };

    function CCommonDataClrList(type, ind, item, isAdd) {
      CBaseFormatObject.call(this, type, ind, item, isAdd);
      this.list = [];
    }

    InitClass(CCommonDataClrList, CBaseFormatObject, AscDFH.historyitem_type_CCommonDataClrList);

    CCommonDataClrList.prototype.addToLst = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
      oHistory.Add(new CChangesContentNoId(this, AscDFH.historyitem_CCommonDataClrListAdd, nInsertIdx, [oPr], true));
      this.list.splice(nInsertIdx, 0, oPr);
    };

    CCommonDataClrList.prototype.removeFromLst = function (nIdx) {
      if (nIdx > -1 && nIdx < this.list.length) {
        this.list[nIdx].setParent(null);
        oHistory.Add(new CChangesContentNoId(this, AscDFH.historyitem_CCommonDataClrListRemove, nIdx, [this.list[nIdx]], false));
        this.list.splice(nIdx, 1);
      }
    };

    CCommonDataClrList.prototype.fillObject = function (oCopy, oIdMap) {
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        oCopy.addToLst(nIdx, this.list[nIdx].createDuplicate(oIdMap));
      }
    };

    CCommonDataClrList.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar2(0, this.hueDir);
      pWriter._WriteUChar2(1, this.meth);
    };
    CCommonDataClrList.prototype.writeChildren = function(pWriter) {
      for (var i = 0; i < this.list.length; i += 1) {
        pWriter.WriteRecord2(0, this.list[i], pWriter.WriteUniColor);
      }
    };
    CCommonDataClrList.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setHueDir(oStream.GetUChar());
      else if (1 === nType) this.setMeth(oStream.GetUChar());
    };
    CCommonDataClrList.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0:
          this.addToLst(this.list.length, pReader.ReadUniColor());
          break;
        default:
          s.SkipRecord();
          break;
      }

    };
    CCommonDataClrList.prototype.getChildren = function() {
      return [].concat(this.list);
    };



    changesFactory[AscDFH.historyitem_ClrLstHueDir] = CChangeLong;
    changesFactory[AscDFH.historyitem_ClrLstMeth] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_ClrLstHueDir] = function (oClass, value) {
      oClass.hueDir = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ClrLstMeth] = function (oClass, value) {
      oClass.meth = value;
    };

    function ClrLst() {
      CCommonDataClrList.call(this);
      this.hueDir = null;
      this.meth = null;
    }

    InitClass(ClrLst, CCommonDataClrList, AscDFH.historyitem_type_ClrLst);

    ClrLst.prototype.setHueDir = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ClrLstHueDir, this.getHueDir(), pr));
      this.hueDir = pr;
    }

    ClrLst.prototype.setMeth = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ClrLstMeth, this.getMeth(), pr));
      this.meth = pr;
    }

    ClrLst.prototype.getHueDir = function () {
      return this.hueDir;
    }

    ClrLst.prototype.getMeth = function () {
      return this.meth;
    }

    ClrLst.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setHueDir(this.getHueDir());
      oCopy.setMeth(this.getMeth());
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        var oColor = this.list[nIdx].createDuplicate(oIdMap);
        oCopy.addToLst(nIdx, oColor);
      }
    }

    function EffectClrLst() {
      ClrLst.call(this);
    }

    InitClass(EffectClrLst, ClrLst, AscDFH.historyitem_type_EffectClrLst);

    function FillClrLst() {
      ClrLst.call(this);
    }

    InitClass(FillClrLst, ClrLst, AscDFH.historyitem_type_FillClrLst);

    function LinClrLst() {
      ClrLst.call(this);
    }

    InitClass(LinClrLst, ClrLst, AscDFH.historyitem_type_LinClrLst);

    function TxEffectClrLst() {
      ClrLst.call(this);
    }

    InitClass(TxEffectClrLst, ClrLst, AscDFH.historyitem_type_TxEffectClrLst);

    function TxFillClrLst() {
      ClrLst.call(this);
    }

    InitClass(TxFillClrLst, ClrLst, AscDFH.historyitem_type_TxFillClrLst);

    function TxLinClrLst() {
      ClrLst.call(this);
    }

    InitClass(TxLinClrLst, ClrLst, AscDFH.historyitem_type_TxLinClrLst);

    function ColorsDefHdrLst() {
      CCommonDataList.call(this);
    }

    InitClass(ColorsDefHdrLst, CCommonDataList, AscDFH.historyitem_type_ColorsDefHdrLst);


    changesFactory[AscDFH.historyitem_ColorsDefHdrMinVer] = CChangeString;
    changesFactory[AscDFH.historyitem_ColorsDefHdrResId] = CChangeLong;
    changesFactory[AscDFH.historyitem_ColorsDefHdrUniqueId] = CChangeString;
    changesFactory[AscDFH.historyitem_ColorsDefHdrCatLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorsDefHdrExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_ColorsDefHdrAddTitle] = CChangeContent;
    changesFactory[AscDFH.historyitem_ColorsDefHdrRemoveTitle] = CChangeContent;
    changesFactory[AscDFH.historyitem_ColorsDefHdrAddDesc] = CChangeContent;
    changesFactory[AscDFH.historyitem_ColorsDefHdrRemoveDesc] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_ColorsDefHdrMinVer] = function (oClass, value) {
      oClass.minVer = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefHdrResId] = function (oClass, value) {
      oClass.resId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefHdrUniqueId] = function (oClass, value) {
      oClass.uniqueId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefHdrCatLst] = function (oClass, value) {
      oClass.catLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ColorsDefHdrExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingContentChanges[AscDFH.historyitem_ColorsDefHdrAddTitle] = function (oClass) {
      return oClass.title;
    };
    drawingContentChanges[AscDFH.historyitem_ColorsDefHdrRemoveTitle] = function (oClass) {
      return oClass.title;
    };
    drawingContentChanges[AscDFH.historyitem_ColorsDefHdrAddDesc] = function (oClass) {
      return oClass.desc;
    };
    drawingContentChanges[AscDFH.historyitem_ColorsDefHdrRemoveDesc] = function (oClass) {
      return oClass.desc;
    };

    function ColorsDefHdr() {
      CBaseFormatObject.call(this);
      this.minVer = null;
      this.resId = null;
      this.uniqueId = null;
      this.catLst = null;
      this.extLst = null;
      this.title = [];
      this.desc = [];
    }

    InitClass(ColorsDefHdr, CBaseFormatObject, AscDFH.historyitem_type_ColorsDefHdr);

    ColorsDefHdr.prototype.setMinVer = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ColorsDefHdrMinVer, this.getMinVer(), pr));
      this.minVer = pr;
    }

    ColorsDefHdr.prototype.setResId = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_ColorsDefHdrResId, this.getResId(), pr));
      this.resId = pr;
    }

    ColorsDefHdr.prototype.setUniqueId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_ColorsDefHdrUniqueId, this.getUniqueId(), pr));
      this.uniqueId = pr;
    }

    ColorsDefHdr.prototype.setCatLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorsDefHdrCatLst, this.getCatLst(), oPr));
      this.catLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorsDefHdr.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ColorsDefHdrExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    ColorsDefHdr.prototype.addToLstTitle = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.title.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ColorsDefHdrAddTitle, nInsertIdx, [oPr], true));
      this.title.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    ColorsDefHdr.prototype.removeFromLstTitle = function (nIdx) {
      if (nIdx > -1 && nIdx < this.title.length) {
        this.title[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ColorsDefHdrRemoveTitle, nIdx, [this.title[nIdx]], false));
        this.title.splice(nIdx, 1);
      }
    };

    ColorsDefHdr.prototype.addToLstDesc = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.desc.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ColorsDefHdrAddDesc, nInsertIdx, [oPr], true));
      this.desc.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    ColorsDefHdr.prototype.removeFromLstDesc = function (nIdx) {
      if (nIdx > -1 && nIdx < this.desc.length) {
        this.desc[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ColorsDefHdrRemoveDesc, nIdx, [this.desc[nIdx]], false));
        this.desc.splice(nIdx, 1);
      }
    };

    ColorsDefHdr.prototype.getMinVer = function () {
      return this.minVer;
    }

    ColorsDefHdr.prototype.getResId = function () {
      return this.resId;
    }

    ColorsDefHdr.prototype.getUniqueId = function () {
      return this.uniqueId;
    }

    ColorsDefHdr.prototype.getCatLst = function () {
      return this.catLst;
    }

    ColorsDefHdr.prototype.getExtLst = function () {
      return this.extLst;
    }

    ColorsDefHdr.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setMinVer(this.getMinVer());
      oCopy.setResId(this.getResId());
      oCopy.setUniqueId(this.getUniqueId());
      if (this.getCatLst()) {
        oCopy.setCatLst(this.getCatLst().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.title.length; ++nIdx) {
        oCopy.addToLstTitle(nIdx, this.title[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.desc.length; ++nIdx) {
        oCopy.addToLstDesc(nIdx, this.desc[nIdx].createDuplicate(oIdMap));
      }
    }


    changesFactory[AscDFH.historyitem_StyleDefMinVer] = CChangeString;
    changesFactory[AscDFH.historyitem_StyleDefUniqueId] = CChangeString;
    changesFactory[AscDFH.historyitem_StyleDefCatLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefScene3d] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefTitle] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefDesc] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefAddStyleLbl] = CChangeContent;
    changesFactory[AscDFH.historyitem_StyleDefRemoveStyleLbl] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_StyleDefMinVer] = function (oClass, value) {
      oClass.minVer = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefUniqueId] = function (oClass, value) {
      oClass.uniqueId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefCatLst] = function (oClass, value) {
      oClass.catLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefScene3d] = function (oClass, value) {
      oClass.scene3d = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefTitle] = function (oClass, value) {
      oClass.title = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefDesc] = function (oClass, value) {
      oClass.desc = value;
    };
    drawingContentChanges[AscDFH.historyitem_StyleDefAddStyleLbl] = function (oClass) {
      return oClass.styleLbl;
    };
    drawingContentChanges[AscDFH.historyitem_StyleDefRemoveStyleLbl] = function (oClass) {
      return oClass.styleLbl;
    };

    function StyleDef() {
      CBaseFormatObject.call(this);
      this.minVer = null;
      this.uniqueId = null;
      this.catLst = null;
      this.extLst = null;
      this.scene3d = null;
      this.title = null;
      this.desc = null;
      this.styleLbl = [];
    }

    InitClass(StyleDef, CBaseFormatObject, AscDFH.historyitem_type_StyleDef);

    StyleDef.prototype.setMinVer = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_StyleDefMinVer, this.getMinVer(), pr));
      this.minVer = pr;
    }

    StyleDef.prototype.setUniqueId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_StyleDefUniqueId, this.getUniqueId(), pr));
      this.uniqueId = pr;
    }

    StyleDef.prototype.setCatLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefCatLst, this.getCatLst(), oPr));
      this.catLst = oPr;
      this.setParentToChild(oPr);
    }

    StyleDef.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    StyleDef.prototype.setScene3d = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefScene3d, this.getScene3d(), oPr));
      this.scene3d = oPr;
      this.setParentToChild(oPr);
    }

    StyleDef.prototype.setTitle = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefTitle, this.getTitle(), oPr));
      this.title = oPr;
      this.setParentToChild(oPr);
    }

    StyleDef.prototype.setDesc = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefDesc, this.getDesc(), oPr));
      this.desc = oPr;
      this.setParentToChild(oPr);
    }

    StyleDef.prototype.addToLstStyleLbl = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.styleLbl.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_StyleDefAddStyleLbl, nInsertIdx, [oPr], true));
      this.styleLbl.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    StyleDef.prototype.removeFromLstStyleLbl = function (nIdx) {
      if (nIdx > -1 && nIdx < this.styleLbl.length) {
        this.styleLbl[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_StyleDefRemoveStyleLbl, nIdx, [this.styleLbl[nIdx]], false));
        this.styleLbl.splice(nIdx, 1);
      }
    };

    StyleDef.prototype.getMinVer = function () {
      return this.minVer;
    }

    StyleDef.prototype.getUniqueId = function () {
      return this.uniqueId;
    }

    StyleDef.prototype.getCatLst = function () {
      return this.catLst;
    }

    StyleDef.prototype.getExtLst = function () {
      return this.extLst;
    }

    StyleDef.prototype.getScene3d = function () {
      return this.scene3d;
    }

    StyleDef.prototype.getTitle = function () {
      return this.title;
    }

    StyleDef.prototype.getDesc = function () {
      return this.desc;
    }

    StyleDef.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setMinVer(this.getMinVer());
      oCopy.setUniqueId(this.getUniqueId());
      if (this.getCatLst()) {
        oCopy.setCatLst(this.getCatLst().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getScene3d()) {
        oCopy.setScene3d(this.getScene3d().createDuplicate(oIdMap));
      }
      if (this.getTitle()) {
        oCopy.setTitle(this.getTitle().createDuplicate(oIdMap));
      }
      if (this.getDesc()) {
        oCopy.setDesc(this.getDesc().createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.styleLbl.length; ++nIdx) {
        oCopy.addToLstStyleLbl(nIdx, this.styleLbl[nIdx].createDuplicate(oIdMap));
      }
    }

    StyleDef.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.uniqueId);
      pWriter._WriteString2(1, this.minVer);
    };
    StyleDef.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.title);
      this.writeRecord2(pWriter, 1, this.desc);
      this.writeRecord2(pWriter, 2, this.catLst);
      this.writeRecord2(pWriter, 3, this.scene3d);
      for (var i = 0; i < this.styleLbl.length; i += 1) {
        this.writeRecord2(pWriter, 4, this.styleLbl[i]);
      }
    };
    StyleDef.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setUniqueId(oStream.GetString2());
      else if (1 === nType) this.setMinVer(oStream.GetString2());
    };
    StyleDef.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setTitle(new DiagramTitle());
          this.title.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setDesc(new Desc());
          this.desc.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setCatLst(new CatLst());
          this.catLst.fromPPTY(pReader);
          break;
        }
        case 3: {
          this.setScene3d(new Scene3d());
          this.scene3d.fromPPTY(pReader);
          break;
        }
        case 4: {
          var oChild = new StyleDefStyleLbl();
          oChild.fromPPTY(pReader);
          this.addToLstStyleLbl(this.styleLbl.length, oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    StyleDef.prototype.getChildren = function() {
      return [this.title, this.desc, this.catLst, this.scene3d].concat(this.styleLbl);
    };



    function BuNone() {
      CBaseFormatObject.call(this);
    }

    InitClass(BuNone, CBaseFormatObject, AscDFH.historyitem_type_BuNone);


    changesFactory[AscDFH.historyitem_Scene3dBackdrop] = CChangeObject;
    changesFactory[AscDFH.historyitem_Scene3dCamera] = CChangeObject;
    changesFactory[AscDFH.historyitem_Scene3dExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_Scene3dLightRig] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_Scene3dBackdrop] = function (oClass, value) {
      oClass.backdrop = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Scene3dCamera] = function (oClass, value) {
      oClass.camera = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Scene3dExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Scene3dLightRig] = function (oClass, value) {
      oClass.lightRig = value;
    };

    function Scene3d() {
      CBaseFormatObject.call(this);
      this.backdrop = null;
      this.camera = null;
      this.extLst = null;
      this.lightRig = null;
    }

    InitClass(Scene3d, CBaseFormatObject, AscDFH.historyitem_type_Scene3d);

    Scene3d.prototype.setBackdrop = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Scene3dBackdrop, this.getBackdrop(), oPr));
      this.backdrop = oPr;
      this.setParentToChild(oPr);
    }

    Scene3d.prototype.setCamera = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Scene3dCamera, this.getCamera(), oPr));
      this.camera = oPr;
      this.setParentToChild(oPr);
    }

    Scene3d.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Scene3dExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    Scene3d.prototype.setLightRig = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Scene3dLightRig, this.getLightRig(), oPr));
      this.lightRig = oPr;
      this.setParentToChild(oPr);
    }

    Scene3d.prototype.getBackdrop = function () {
      return this.backdrop;
    }

    Scene3d.prototype.getCamera = function () {
      return this.camera;
    }

    Scene3d.prototype.getExtLst = function () {
      return this.extLst;
    }

    Scene3d.prototype.getLightRig = function () {
      return this.lightRig;
    }

    Scene3d.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getBackdrop()) {
        oCopy.setBackdrop(this.getBackdrop().createDuplicate(oIdMap));
      }
      if (this.getCamera()) {
        oCopy.setCamera(this.getCamera().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getLightRig()) {
        oCopy.setLightRig(this.getLightRig().createDuplicate(oIdMap));
      }
    }

    Scene3d.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.camera);
      this.writeRecord2(pWriter, 1, this.lightRig);
      this.writeRecord2(pWriter, 2, this.backdrop);
    };

    Scene3d.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
         case 0:
           this.setCamera(new Camera());
           this.camera.fromPPTY(pReader);
           break;
        case 1:
          this.setLightRig(new LightRig());
          this.lightRig.fromPPTY(pReader);
          break;
        case 2:
          this.setBackdrop(new Backdrop());
          this.backdrop.fromPPTY(pReader);
          break;
        default: {
          break;
        }
      }
    };

    Scene3d.prototype.getChildren = function () {
      return [this.camera, this.lightRig, this.backdrop];
    }


    changesFactory[AscDFH.historyitem_StyleDefStyleLblName] = CChangeString;
    changesFactory[AscDFH.historyitem_StyleDefStyleLblExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefStyleLblScene3d] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefStyleLblSp3d] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefStyleLblStyle] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefStyleLblTxPr] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_StyleDefStyleLblName] = function (oClass, value) {
      oClass.name = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefStyleLblExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefStyleLblScene3d] = function (oClass, value) {
      oClass.scene3d = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefStyleLblSp3d] = function (oClass, value) {
      oClass.sp3d = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefStyleLblStyle] = function (oClass, value) {
      oClass.style = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefStyleLblTxPr] = function (oClass, value) {
      oClass.txPr = value;
    };

    function StyleDefStyleLbl() {
      CBaseFormatObject.call(this);
      this.name = null;
      this.extLst = null;
      this.scene3d = null;
      this.sp3d = null;
      this.style = null;
      this.txPr = null;
    }

    InitClass(StyleDefStyleLbl, CBaseFormatObject, AscDFH.historyitem_type_StyleDefStyleLbl);

    StyleDefStyleLbl.prototype.setName = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_StyleDefStyleLblName, this.getName(), pr));
      this.name = pr;
    }

    StyleDefStyleLbl.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefStyleLblExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    StyleDefStyleLbl.prototype.setScene3d = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefStyleLblScene3d, this.getScene3d(), oPr));
      this.scene3d = oPr;
      this.setParentToChild(oPr);
    }

    StyleDefStyleLbl.prototype.setSp3d = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefStyleLblSp3d, this.getSp3d(), oPr));
      this.sp3d = oPr;
      this.setParentToChild(oPr);
    }

    StyleDefStyleLbl.prototype.setStyle = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefStyleLblStyle, this.getStyle(), oPr));
      this.style = oPr;
      // this.setParentToChild(oPr); TODO: fix set Parent
    }

    StyleDefStyleLbl.prototype.setTxPr = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefStyleLblTxPr, this.getTxPr(), oPr));
      this.txPr = oPr;
      this.setParentToChild(oPr);
    }

    StyleDefStyleLbl.prototype.getName = function () {
      return this.name;
    }

    StyleDefStyleLbl.prototype.getExtLst = function () {
      return this.extLst;
    }

    StyleDefStyleLbl.prototype.getScene3d = function () {
      return this.scene3d;
    }

    StyleDefStyleLbl.prototype.getSp3d = function () {
      return this.sp3d;
    }

    StyleDefStyleLbl.prototype.getStyle = function () {
      return this.style;
    }

    StyleDefStyleLbl.prototype.getTxPr = function () {
      return this.txPr;
    }

    StyleDefStyleLbl.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setName(this.getName());
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getScene3d()) {
        oCopy.setScene3d(this.getScene3d().createDuplicate(oIdMap));
      }
      if (this.getSp3d()) {
        oCopy.setSp3d(this.getSp3d().createDuplicate(oIdMap));
      }
      if (this.getStyle()) {
        oCopy.setStyle(this.getStyle().createDuplicate(oIdMap));
      }
      if (this.getTxPr()) {
        oCopy.setTxPr(this.getTxPr().createDuplicate(oIdMap));
      }
    }

    StyleDefStyleLbl.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
    };
    StyleDefStyleLbl.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.scene3d);
      this.writeRecord2(pWriter, 1, this.sp3d);
      pWriter.WriteRecord2(2, this.style, pWriter.WriteShapeStyle);
      pWriter.WriteRecord2(3, this.txPr, pWriter.WriteTxBody);
    };
    StyleDefStyleLbl.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
    };
    StyleDefStyleLbl.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setScene3d(new Scene3d());
          this.scene3d.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setSp3d(new Sp3d());
          this.sp3d.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setStyle(pReader.ReadShapeStyle());
          break;
        }
        case 3: {
          this.setTxPr(pReader.ReadTextBody());
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    StyleDefStyleLbl.prototype.getChildren = function() {
      return [this.scene3d, this.sp3d, this.style, this.txPr];
    };

    changesFactory[AscDFH.historyitem_BackdropAnchor] = CChangeObject;
    changesFactory[AscDFH.historyitem_BackdropExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_BackdropNorm] = CChangeObject;
    changesFactory[AscDFH.historyitem_BackdropUp] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_BackdropAnchor] = function (oClass, value) {
      oClass.anchor = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropNorm] = function (oClass, value) {
      oClass.norm = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropUp] = function (oClass, value) {
      oClass.up = value;
    };

    function Backdrop() {
      CBaseFormatObject.call(this);
      this.anchor = new BackdropAnchor();
      this.extLst = null;
      this.norm = new BackdropNorm();
      this.up = new BackdropUp();
    }

    InitClass(Backdrop, CBaseFormatObject, AscDFH.historyitem_type_Backdrop);

    Backdrop.prototype.setAnchor = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_BackdropAnchor, this.getAnchor(), oPr));
      this.anchor = oPr;
      this.setParentToChild(oPr);
    }

    Backdrop.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_BackdropExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    Backdrop.prototype.setNorm = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_BackdropNorm, this.getNorm(), oPr));
      this.norm = oPr;
      this.setParentToChild(oPr);
    }

    Backdrop.prototype.setUp = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_BackdropUp, this.getUp(), oPr));
      this.up = oPr;
      this.setParentToChild(oPr);
    }

    Backdrop.prototype.getAnchor = function () {
      return this.anchor;
    }

    Backdrop.prototype.getExtLst = function () {
      return this.extLst;
    }

    Backdrop.prototype.getNorm = function () {
      return this.norm;
    }

    Backdrop.prototype.getUp = function () {
      return this.up;
    }

    Backdrop.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getAnchor()) {
        oCopy.setAnchor(this.getAnchor().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getNorm()) {
        oCopy.setNorm(this.getNorm().createDuplicate(oIdMap));
      }
      if (this.getUp()) {
        oCopy.setUp(this.getUp().createDuplicate(oIdMap));
      }
    }

    Backdrop.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteInt1(0, this.anchor.x);
      pWriter._WriteInt1(1, this.anchor.y);
      pWriter._WriteInt1(2, this.anchor.z);

      pWriter._WriteInt1(3, this.norm.dx);
      pWriter._WriteInt1(4, this.norm.dy);
      pWriter._WriteInt1(5, this.norm.dz);

      pWriter._WriteInt1(6, this.up.dx);
      pWriter._WriteInt1(7, this.up.dy);
      pWriter._WriteInt1(8, this.up.dz);
    };
    Backdrop.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.anchor.setX(oStream.GetLong());
      else if (1 === nType) this.anchor.setY(oStream.GetLong());
      else if (2 === nType) this.anchor.setZ(oStream.GetLong());
      else if (3 === nType) this.norm.setDx(oStream.GetLong());
      else if (4 === nType) this.norm.setDy(oStream.GetLong());
      else if (5 === nType) this.norm.setDz(oStream.GetLong());
      else if (6 === nType) this.up.setDx(oStream.GetLong());
      else if (7 === nType) this.up.setDy(oStream.GetLong());
      else if (8 === nType) this.up.setDz(oStream.GetLong());
    };

    changesFactory[AscDFH.historyitem_CoordinateCoordinateUnqualified] = CChangeLong;
    changesFactory[AscDFH.historyitem_CoordinateUniversalMeasure] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_CoordinateCoordinateUnqualified] = function (oClass, value) {
      oClass.coordinateUnqualified = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CoordinateUniversalMeasure] = function (oClass, value) {
      oClass.universalMeasure = value;
    };

    function Coordinate() {
      CBaseFormatObject.call(this);
      this.coordinateUnqualified = null;
      this.universalMeasure = null;
    }

    InitClass(Coordinate, CBaseFormatObject, AscDFH.historyitem_type_Coordinate);

    Coordinate.prototype.setCoordinateUnqualified = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CoordinateCoordinateUnqualified, this.getCoordinateUnqualified(), pr));
      this.coordinateUnqualified = pr;
    }

    Coordinate.prototype.setUniversalMeasure = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CoordinateUniversalMeasure, this.getUniversalMeasure(), pr));
      this.universalMeasure = pr;
    }

    Coordinate.prototype.getCoordinateUnqualified = function () {
      return this.coordinateUnqualified;
    }

    Coordinate.prototype.getUniversalMeasure = function () {
      return this.universalMeasure;
    }

    Coordinate.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setCoordinateUnqualified(this.getCoordinateUnqualified());
      oCopy.setUniversalMeasure(this.getUniversalMeasure());
    }

    changesFactory[AscDFH.historyitem_BackdropAnchorX] = CChangeLong;
    changesFactory[AscDFH.historyitem_BackdropAnchorY] = CChangeLong;
    changesFactory[AscDFH.historyitem_BackdropAnchorZ] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_BackdropAnchorX] = function (oClass, value) {
      oClass.x = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropAnchorY] = function (oClass, value) {
      oClass.y = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropAnchorZ] = function (oClass, value) {
      oClass.z = value;
    };

    function BackdropAnchor() {
      CBaseFormatObject.call(this);
      this.x = null;
      this.y = null;
      this.z = null;
    }

    InitClass(BackdropAnchor, CBaseFormatObject, AscDFH.historyitem_type_BackdropAnchor);

    BackdropAnchor.prototype.setX = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropAnchorX, this.getX(), oPr));
      this.x = oPr;
    }

    BackdropAnchor.prototype.setY = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropAnchorY, this.getY(), oPr));
      this.y = oPr;
    }

    BackdropAnchor.prototype.setZ = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropAnchorZ, this.getZ(), oPr));
      this.z = oPr;
    }

    BackdropAnchor.prototype.getX = function () {
      return this.x;
    }

    BackdropAnchor.prototype.getY = function () {
      return this.y;
    }

    BackdropAnchor.prototype.getZ = function () {
      return this.z;
    }

    BackdropAnchor.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getX() !== null) {
        oCopy.setX(this.getX());
      }
      if (this.getY() !== null) {
        oCopy.setY(this.getY());
      }
      if (this.getZ() !== null) {
        oCopy.setZ(this.getZ());
      }
    }

    function Drawing() {
      CGroupShape.call(this);
    }

    InitClass(Drawing, CGroupShape, AscDFH.historyitem_type_SmartArtDrawing);

    Drawing.prototype.getObjectType = function () {
      return AscDFH.historyitem_type_SmartArtDrawing;
    }
    Drawing.prototype.getName = function () {
      return 'Drawing';
    }
    Drawing.prototype.updateCoordinatesAfterInternalResize = function () {

    }

    Drawing.prototype.writeChildren = function(pWriter) {
      pWriter.WriteGroupShape(this, 0);
    };
    Drawing.prototype.toPPTY = function(pWriter) {
      this.writeChildren(pWriter);
    };

    Drawing.prototype.copy = function(oPr)
    {
      var copy = new Drawing();
      this.copy2(copy, oPr);
      return copy;
    };

    Drawing.prototype.copy2 = function(copy, oPr)
    {
      if(this.nvGrpSpPr)
      {
        copy.setNvGrpSpPr(this.nvGrpSpPr.createDuplicate());
      }
      if(this.spPr)
      {
        copy.setSpPr(this.spPr.createDuplicate());
        copy.spPr.setParent(copy);
      }
      for(var i = 0; i < this.spTree.length; ++i)
      {
        var _copy;
        if(this.spTree[i].getObjectType() === AscDFH.historyitem_type_GroupShape) {
          _copy = this.spTree[i].copy(oPr);
        }
        else{
          if(oPr && oPr.bSaveSourceFormatting){
            _copy = this.spTree[i].getCopyWithSourceFormatting();
          }
          else{
            _copy = this.spTree[i].copy(oPr);
          }

        }
        if(oPr && AscCommon.isRealObject(oPr.idMap)){
          oPr.idMap[this.spTree[i].Id] = _copy.Id;
        }
        copy.addToSpTree(copy.spTree.length, _copy);
        copy.spTree[copy.spTree.length-1].setGroup(copy);
      }
      copy.setBDeleted(this.bDeleted);
      if(this.macro !== null) {
        copy.setMacro(this.macro);
      }
      if(this.textLink !== null) {
        copy.setTextLink(this.textLink);
      }
      copy.cachedImage = this.getBase64Img();
      copy.cachedPixH = this.cachedPixH;
      copy.cachedPixW = this.cachedPixW;
      copy.setLocks(this.locks);
      if (this.group) {
        copy.setGroup(this.group);
      }

      return copy;
    };
    Drawing.prototype.createPlaceholderControl = function(aControls) {
      for(var nSp = 0; nSp < this.spTree.length; ++nSp) {
        var oShape = this.spTree[nSp];
        if (oShape.isActiveBlipFillPlaceholder()) {
          oShape.createPlaceholderControl(aControls);
        }
      }
    };

    Drawing.prototype.setXfrmByParent = function () {
      var oXfrm = this.spPr.xfrm;
      if (oXfrm.isZero()) {
        var parent = this.group;
        if (parent && parent.spPr.xfrm) {
          oXfrm.setExtX(parent.spPr.xfrm.extX);
          oXfrm.setExtY(parent.spPr.xfrm.extY);
        }
      }
    };

    Drawing.prototype.handleUpdateExtents = function(bExt)
    {
      this.recalcTransform();
      this.recalcBounds();
      this.addToRecalculate();
      this.recalcWrapPolygon();
      if(this.spTree)
      {
        for(var i = 0; i < this.spTree.length; ++i)
        {
          this.spTree[i].handleUpdateExtents(bExt);
        }
      }
    };

    changesFactory[AscDFH.historyitem_BackdropNormDx] = CChangeLong;
    changesFactory[AscDFH.historyitem_BackdropNormDy] = CChangeLong;
    changesFactory[AscDFH.historyitem_BackdropNormDz] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_BackdropNormDx] = function (oClass, value) {
      oClass.dx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropNormDy] = function (oClass, value) {
      oClass.dy = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropNormDz] = function (oClass, value) {
      oClass.dz = value;
    };

    function BackdropNorm() {
      CBaseFormatObject.call(this);
      this.dx = null;
      this.dy = null;
      this.dz = null;
    }

    InitClass(BackdropNorm, CBaseFormatObject, AscDFH.historyitem_type_BackdropNorm);

    BackdropNorm.prototype.setDx = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropNormDx, this.getDx(), oPr));
      this.dx = oPr;
    }

    BackdropNorm.prototype.setDy = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropNormDy, this.getDy(), oPr));
      this.dy = oPr;
    }

    BackdropNorm.prototype.setDz = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropNormDz, this.getDz(), oPr));
      this.dz = oPr;
    }

    BackdropNorm.prototype.getDx = function () {
      return this.dx;
    }

    BackdropNorm.prototype.getDy = function () {
      return this.dy;
    }

    BackdropNorm.prototype.getDz = function () {
      return this.dz;
    }

    BackdropNorm.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getDx()) {
        oCopy.setDx(this.getDx());
      }
      if (this.getDy()) {
        oCopy.setDy(this.getDy());
      }
      if (this.getDz()) {
        oCopy.setDz(this.getDz());
      }
    }


    changesFactory[AscDFH.historyitem_BackdropUpDx] = CChangeLong;
    changesFactory[AscDFH.historyitem_BackdropUpDy] = CChangeLong;
    changesFactory[AscDFH.historyitem_BackdropUpDz] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_BackdropUpDx] = function (oClass, value) {
      oClass.dx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropUpDy] = function (oClass, value) {
      oClass.dy = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BackdropUpDz] = function (oClass, value) {
      oClass.dz = value;
    };

    function BackdropUp() {
      CBaseFormatObject.call(this);
      this.dx = null;
      this.dy = null;
      this.dz = null;
    }

    InitClass(BackdropUp, CBaseFormatObject, AscDFH.historyitem_type_BackdropUp);

    BackdropUp.prototype.setDx = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropUpDx, this.getDx(), oPr));
      this.dx = oPr;
    }

    BackdropUp.prototype.setDy = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropUpDy, this.getDy(), oPr));
      this.dy = oPr;
    }

    BackdropUp.prototype.setDz = function (oPr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BackdropUpDz, this.getDz(), oPr));
      this.dz = oPr;
    }

    BackdropUp.prototype.getDx = function () {
      return this.dx;
    }

    BackdropUp.prototype.getDy = function () {
      return this.dy;
    }

    BackdropUp.prototype.getDz = function () {
      return this.dz;
    }

    BackdropUp.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getDx()) {
        oCopy.setDx(this.getDx());
      }
      if (this.getDy()) {
        oCopy.setDy(this.getDy());
      }
      if (this.getDz()) {
        oCopy.setDz(this.getDz());
      }
    }

    changesFactory[AscDFH.historyitem_CameraFov] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_CameraPrst] = CChangeLong;
    changesFactory[AscDFH.historyitem_CameraZoom] = CChangeDouble2;
    changesFactory[AscDFH.historyitem_CameraRot] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_CameraFov] = function (oClass, value) {
      oClass.fov = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CameraPrst] = function (oClass, value) {
      oClass.prst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CameraZoom] = function (oClass, value) {
      oClass.zoom = value;
    };
    drawingsChangesMap[AscDFH.historyitem_CameraRot] = function (oClass, value) {
      oClass.rot = value;
    };

    function Camera() {
      CBaseFormatObject.call(this);
      this.fov = null;
      this.prst = null;
      this.zoom = null;
      this.rot = null;
    }

    InitClass(Camera, CBaseFormatObject, AscDFH.historyitem_type_Camera);

    Camera.prototype.setFov = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_CameraFov, this.getFov(), pr));
      this.fov = pr;
    }

    Camera.prototype.setPrst = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CameraPrst, this.getPrst(), pr));
      this.prst = pr;
    }

    Camera.prototype.setZoom = function (pr) {
      oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_CameraZoom, this.getZoom(), pr));
      this.zoom = pr;
    }

    Camera.prototype.setRot = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_CameraRot, this.getRot(), oPr));
      this.rot = oPr;
      this.setParentToChild(oPr);
    }

    Camera.prototype.getFov = function () {
      return this.fov;
    }

    Camera.prototype.getPrst = function () {
      return this.prst;
    }

    Camera.prototype.getZoom = function () {
      return this.zoom;
    }

    Camera.prototype.getRot = function () {
      return this.rot;
    }

    Camera.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setFov(this.getFov());
      oCopy.setPrst(this.getPrst());
      oCopy.setZoom(this.getZoom());
      if (this.getRot()) {
        oCopy.setRot(this.getRot().createDuplicate(oIdMap));
      }
    }

    Camera.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.prst);
      pWriter._WriteInt2(1, this.fov);
      pWriter._WriteInt2(2, this.zoom);
    };
    Camera.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.rot);
    };
    Camera.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setPrst(oStream.GetUChar());
      else if (1 === nType) this.setFov(oStream.GetLong());
      else if (2 === nType) this.setZoom(oStream.GetLong());
    };
    Camera.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setRot(new Rot());
          this.rot.fromPPTY(pReader);
          break;
        }
        default: {
          break;
        }
      }
    };

    Camera.prototype.getChildren = function () {
      return [this.rot];
    }

    changesFactory[AscDFH.historyitem_RotLat] = CChangeLong;
    changesFactory[AscDFH.historyitem_RotLon] = CChangeLong;
    changesFactory[AscDFH.historyitem_RotRev] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_RotLat] = function (oClass, value) {
      oClass.lat = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RotLon] = function (oClass, value) {
      oClass.lon = value;
    };
    drawingsChangesMap[AscDFH.historyitem_RotRev] = function (oClass, value) {
      oClass.rev = value;
    };

    function Rot() {
      CBaseFormatObject.call(this);
      this.lat = null;
      this.lon = null;
      this.rev = null;
    }

    InitClass(Rot, CBaseFormatObject, AscDFH.historyitem_type_Rot);

    Rot.prototype.setLat = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_RotLat, this.getLat(), pr))
      this.lat = pr;
    }

    Rot.prototype.setLon = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_RotLon, this.getLon(), pr));
      this.lon = pr;
    }

    Rot.prototype.setRev = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_RotRev, this.getRev(), pr));
      this.rev = pr;
    }

    Rot.prototype.getLat = function () {
      return this.lat;
    }

    Rot.prototype.getLon = function () {
      return this.lon;
    }

    Rot.prototype.getRev = function () {
      return this.rev;
    }

    Rot.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setLat(this.getLat());
      oCopy.setLon(this.getLon());
      oCopy.setRev(this.getRev());
    }

    Rot.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteInt2(0, this.lat);
      pWriter._WriteInt2(1, this.lon);
      pWriter._WriteInt2(2, this.rev);
    };

    Rot.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setLat(oStream.GetLong());
      else if (1 === nType) this.setLon(oStream.GetLong());
      else if (2 === nType) this.setRev(oStream.GetLong());
    };


    changesFactory[AscDFH.historyitem_LightRigDir] = CChangeLong;
    changesFactory[AscDFH.historyitem_LightRigRig] = CChangeLong;
    changesFactory[AscDFH.historyitem_LightRigRot] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_LightRigDir] = function (oClass, value) {
      oClass.dir = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LightRigRig] = function (oClass, value) {
      oClass.rig = value;
    };
    drawingsChangesMap[AscDFH.historyitem_LightRigRot] = function (oClass, value) {
      oClass.rot = value;
    };

    function LightRig() {
      CBaseFormatObject.call(this);
      this.dir = null;
      this.rig = null;
      this.rot = null;
    }

    InitClass(LightRig, CBaseFormatObject, AscDFH.historyitem_type_LightRig);

    LightRig.prototype.setDir = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_LightRigDir, this.getDir(), pr));
      this.dir = pr;
    }

    LightRig.prototype.setRig = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_LightRigRig, this.getRig(), pr));
      this.rig = pr;
    }

    LightRig.prototype.setRot = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_LightRigRot, this.getRot(), oPr));
      this.rot = oPr;
      this.setParentToChild(oPr);
    }

    LightRig.prototype.getDir = function () {
      return this.dir;
    }

    LightRig.prototype.getRig = function () {
      return this.rig;
    }

    LightRig.prototype.getRot = function () {
      return this.rot;
    }

    LightRig.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setDir(this.getDir());
      oCopy.setRig(this.getRig());
      if (this.getRot()) {
        oCopy.setRot(this.getRot().createDuplicate(oIdMap));
      }
    }

    LightRig.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteUChar1(0, this.dir);
      pWriter._WriteUChar1(1, this.rig);
    };
    LightRig.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.rot);
      };
    LightRig.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setDir(oStream.GetUChar());
      else if (1 === nType) this.setRig(oStream.GetUChar());
    };
    LightRig.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setRot(new Rot());
          this.rot.fromPPTY(pReader);
          break;
        }
        default: {
          break;
        }
      }
    };
    LightRig.prototype.getChildren = function() {
      return [this.rot];
    };


    changesFactory[AscDFH.historyitem_Sp3dContourW] = CChangeLong;
    changesFactory[AscDFH.historyitem_Sp3dExtrusionH] = CChangeLong;
    changesFactory[AscDFH.historyitem_Sp3dPrstMaterial] = CChangeLong;
    changesFactory[AscDFH.historyitem_Sp3dZ] = CChangeObject;
    changesFactory[AscDFH.historyitem_Sp3dBevelB] = CChangeObject;
    changesFactory[AscDFH.historyitem_Sp3dBevelT] = CChangeObject;
    changesFactory[AscDFH.historyitem_Sp3dContourClr] = CChangeObject;
    changesFactory[AscDFH.historyitem_Sp3dExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_Sp3dExtrusionClr] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_Sp3dContourW] = function (oClass, value) {
      oClass.contourW = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Sp3dExtrusionH] = function (oClass, value) {
      oClass.extrusionH = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Sp3dPrstMaterial] = function (oClass, value) {
      oClass.prstMaterial = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Sp3dZ] = function (oClass, value) {
      oClass.z = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Sp3dBevelB] = function (oClass, value) {
      oClass.bevelB = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Sp3dBevelT] = function (oClass, value) {
      oClass.bevelT = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Sp3dContourClr] = function (oClass, value) {
      oClass.contourClr = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Sp3dExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_Sp3dExtrusionClr] = function (oClass, value) {
      oClass.extrusionClr = value;
    };

    function Sp3d() {
      CBaseFormatObject.call(this);
      this.contourW = null;
      this.extrusionH = null;
      this.prstMaterial = null;
      this.z = null;
      this.bevelB = null;
      this.bevelT = null;
      this.contourClr = null;
      this.extLst = null;
      this.extrusionClr = null;
    }

    InitClass(Sp3d, CBaseFormatObject, AscDFH.historyitem_type_Sp3d);

    Sp3d.prototype.setContourW = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_Sp3dContourW, this.getContourW(), pr));
      this.contourW = pr;
    }

    Sp3d.prototype.setExtrusionH = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_Sp3dExtrusionH, this.getExtrusionH(), pr));
      this.extrusionH = pr;
    }

    Sp3d.prototype.setPrstMaterial = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_Sp3dPrstMaterial, this.getPrstMaterial(), pr));
      this.prstMaterial = pr;
    }

    Sp3d.prototype.setZ = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Sp3dZ, this.getZ(), oPr));
      this.z = oPr;
    }

    Sp3d.prototype.setBevelB = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Sp3dBevelB, this.getBevelB(), oPr));
      this.bevelB = oPr;
      this.setParentToChild(oPr);
    }

    Sp3d.prototype.setBevelT = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Sp3dBevelT, this.getBevelT(), oPr));
      this.bevelT = oPr;
      this.setParentToChild(oPr);
    }

    Sp3d.prototype.setContourClr = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Sp3dContourClr, this.getContourClr(), oPr));
      this.contourClr = oPr;
      this.setParentToChild(oPr);
    }

    Sp3d.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Sp3dExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    Sp3d.prototype.setExtrusionClr = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_Sp3dExtrusionClr, this.getExtrusionClr(), oPr));
      this.extrusionClr = oPr;
      this.setParentToChild(oPr);
    }

    Sp3d.prototype.getContourW = function () {
      return this.contourW;
    }

    Sp3d.prototype.getExtrusionH = function () {
      return this.extrusionH;
    }

    Sp3d.prototype.getPrstMaterial = function () {
      return this.prstMaterial;
    }

    Sp3d.prototype.getZ = function () {
      return this.z;
    }

    Sp3d.prototype.getBevelB = function () {
      return this.bevelB;
    }

    Sp3d.prototype.getBevelT = function () {
      return this.bevelT;
    }

    Sp3d.prototype.getContourClr = function () {
      return this.contourClr;
    }

    Sp3d.prototype.getExtLst = function () {
      return this.extLst;
    }

    Sp3d.prototype.getExtrusionClr = function () {
      return this.extrusionClr;
    }

    Sp3d.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setContourW(this.getContourW());
      oCopy.setExtrusionH(this.getExtrusionH());
      oCopy.setPrstMaterial(this.getPrstMaterial());
      if (this.getZ()) {
        oCopy.setZ(this.getZ().createDuplicate(oIdMap));
      }
      if (this.getBevelB()) {
        oCopy.setBevelB(this.getBevelB().createDuplicate(oIdMap));
      }
      if (this.getBevelT()) {
        oCopy.setBevelT(this.getBevelT().createDuplicate(oIdMap));
      }
      if (this.getContourClr()) {
        oCopy.setContourClr(this.getContourClr().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      if (this.getExtrusionClr()) {
        oCopy.setExtrusionClr(this.getExtrusionClr().createDuplicate(oIdMap));
      }
    }

    changesFactory[AscDFH.historyitem_ContourClrColor] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_ContourClrColor] = function (oClass, value) {
      oClass.color = value;
    };

    function ContourClr() {
      CBaseFormatObject.call(this);
      this.color = null;
    }

    InitClass(ContourClr, CBaseFormatObject, AscDFH.historyitem_type_ContourClr);

    ContourClr.prototype.setColor = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ContourClrColor, this.getColor(), oPr));
      this.color = oPr;
      this.setParentToChild(oPr);
    }

    ContourClr.prototype.getColor = function () {
      return this.color;
    }

    ContourClr.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getColor()) {
        oCopy.setColor(this.getColor().createDuplicate(oIdMap));
      }
    }

    changesFactory[AscDFH.historyitem_ExtrusionClrColor] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_ExtrusionClrColor] = function (oClass, value) {
      oClass.color = value;
    };

    function ExtrusionClr() {
      CBaseFormatObject.call(this);
      this.color = null;
    }

    InitClass(ExtrusionClr, CBaseFormatObject, AscDFH.historyitem_type_ExtrusionClr);

    ExtrusionClr.prototype.setColor = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ExtrusionClrColor, this.getColor(), oPr));
      this.color = oPr;
      this.setParentToChild(oPr);
    }

    ExtrusionClr.prototype.getColor = function () {
      return this.color;
    }

    ExtrusionClr.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getColor()) {
        oCopy.setColor(this.getColor().createDuplicate(oIdMap));
      }
    }

    changesFactory[AscDFH.historyitem_BevelH] = CChangeLong;
    changesFactory[AscDFH.historyitem_BevelPrst] = CChangeLong;
    changesFactory[AscDFH.historyitem_BevelW] = CChangeLong;
    drawingsChangesMap[AscDFH.historyitem_BevelH] = function (oClass, value) {
      oClass.h = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BevelPrst] = function (oClass, value) {
      oClass.prst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_BevelW] = function (oClass, value) {
      oClass.w = value;
    };

    function Bevel() {
      CBaseFormatObject.call(this);
      this.h = null;
      this.prst = null;
      this.w = null;
    }

    InitClass(Bevel, CBaseFormatObject, AscDFH.historyitem_type_Bevel);

    Bevel.prototype.setH = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BevelH, this.getH(), pr));
      this.h = pr;
    }

    Bevel.prototype.setPrst = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BevelPrst, this.getPrst(), pr));
      this.prst = pr;
    }

    Bevel.prototype.setW = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_BevelW, this.getW(), pr));
      this.w = pr;
    }

    Bevel.prototype.getH = function () {
      return this.h;
    }

    Bevel.prototype.getPrst = function () {
      return this.prst;
    }

    Bevel.prototype.getW = function () {
      return this.w;
    }

    Bevel.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setH(this.getH());
      oCopy.setPrst(this.getPrst());
      oCopy.setW(this.getW());
    }

    function BevelB() {
      Bevel.call(this);
    }

    InitClass(BevelB, Bevel, AscDFH.historyitem_type_BevelB);

    function BevelT() {
      Bevel.call(this);
    }

    InitClass(BevelT, Bevel, AscDFH.historyitem_type_BevelT);


    changesFactory[AscDFH.historyitem_TxPrFlatTx] = CChangeObject;
    changesFactory[AscDFH.historyitem_TxPrSp3d] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_TxPrFlatTx] = function (oClass, value) {
      oClass.flatTx = value;
    };
    drawingsChangesMap[AscDFH.historyitem_TxPrSp3d] = function (oClass, value) {
      oClass.sp3d = value;
    };

    function TxPr() {
      CBaseFormatObject.call(this);
      this.flatTx = null;
      this.sp3d = null;
    }

    InitClass(TxPr, CBaseFormatObject, AscDFH.historyitem_type_TxPr);

    TxPr.prototype.setFlatTx = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TxPrFlatTx, this.getFlatTx(), oPr));
      this.flatTx = oPr;
      this.setParentToChild(oPr);
    }

    TxPr.prototype.setSp3d = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_TxPrSp3d, this.getSp3d(), oPr));
      this.sp3d = oPr;
      this.setParentToChild(oPr);
    }

    TxPr.prototype.getFlatTx = function () {
      return this.flatTx;
    }

    TxPr.prototype.getSp3d = function () {
      return this.sp3d;
    }

    TxPr.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getFlatTx()) {
        oCopy.setFlatTx(this.getFlatTx().createDuplicate(oIdMap));
      }
      if (this.getSp3d()) {
        oCopy.setSp3d(this.getSp3d().createDuplicate(oIdMap));
      }
    }

    changesFactory[AscDFH.historyitem_FlatTxZ] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_FlatTxZ] = function (oClass, value) {
      oClass.z = value;
    };

    function FlatTx() {
      CBaseFormatObject.call(this);
      this.z = null;
    }

    InitClass(FlatTx, CBaseFormatObject, AscDFH.historyitem_type_FlatTx);

    FlatTx.prototype.setZ = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_FlatTxZ, this.getZ(), oPr));
      this.z = oPr;
    }

    FlatTx.prototype.getZ = function () {
      return this.z;
    }

    FlatTx.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getZ()) {
        oCopy.setZ(this.getZ().createDuplicate(oIdMap));
      }
    }

    function StyleDefHdrLst() {
      CCommonDataList.call(this);
    }

    InitClass(StyleDefHdrLst, CCommonDataList, AscDFH.historyitem_type_StyleDefHdrLst);

    changesFactory[AscDFH.historyitem_StyleDefHdrMinVer] = CChangeString;
    changesFactory[AscDFH.historyitem_StyleDefHdrResId] = CChangeLong;
    changesFactory[AscDFH.historyitem_StyleDefHdrUniqueId] = CChangeString;
    changesFactory[AscDFH.historyitem_StyleDefHdrCatLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefHdrExtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_StyleDefHdrAddDesc] = CChangeContent;
    changesFactory[AscDFH.historyitem_StyleDefHdrRemoveDesc] = CChangeContent;
    changesFactory[AscDFH.historyitem_StyleDefHdrAddList] = CChangeContent;
    changesFactory[AscDFH.historyitem_StyleDefHdrRemoveList] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_StyleDefHdrMinVer] = function (oClass, value) {
      oClass.minVer = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefHdrResId] = function (oClass, value) {
      oClass.resId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefHdrUniqueId] = function (oClass, value) {
      oClass.uniqueId = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefHdrCatLst] = function (oClass, value) {
      oClass.catLst = value;
    };
    drawingsChangesMap[AscDFH.historyitem_StyleDefHdrExtLst] = function (oClass, value) {
      oClass.extLst = value;
    };
    drawingContentChanges[AscDFH.historyitem_StyleDefHdrAddDesc] = function (oClass) {
      return oClass.desc;
    };
    drawingContentChanges[AscDFH.historyitem_StyleDefHdrRemoveDesc] = function (oClass) {
      return oClass.desc;
    };
    drawingContentChanges[AscDFH.historyitem_StyleDefHdrAddList] = function (oClass) {
      return oClass.list;
    };
    drawingContentChanges[AscDFH.historyitem_StyleDefHdrRemoveList] = function (oClass) {
      return oClass.list;
    };

    function StyleDefHdr() {
      CBaseFormatObject.call(this);
      this.minVer = null;
      this.resId = null;
      this.uniqueId = null;
      this.catLst = null;
      this.extLst = null;
      this.desc = [];
      this.list = [];
    }

    InitClass(StyleDefHdr, CBaseFormatObject, AscDFH.historyitem_type_StyleDefHdr);

    StyleDefHdr.prototype.setMinVer = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_StyleDefHdrMinVer, this.getMinVer(), pr));
      this.minVer = pr;
    }

    StyleDefHdr.prototype.setResId = function (pr) {
      oHistory.Add(new CChangeLong(this, AscDFH.historyitem_StyleDefHdrResId, this.getResId(), pr));
      this.resId = pr;
    }

    StyleDefHdr.prototype.setUniqueId = function (pr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_StyleDefHdrUniqueId, this.getUniqueId(), pr));
      this.uniqueId = pr;
    }

    StyleDefHdr.prototype.setCatLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefHdrCatLst, this.getCatLst(), oPr));
      this.catLst = oPr;
      this.setParentToChild(oPr);
    }

    StyleDefHdr.prototype.setExtLst = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_StyleDefHdrExtLst, this.getExtLst(), oPr));
      this.extLst = oPr;
      this.setParentToChild(oPr);
    }

    StyleDefHdr.prototype.addToLstDesc = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.desc.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_StyleDefHdrAddDesc, nInsertIdx, [oPr], true));
      this.desc.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    StyleDefHdr.prototype.removeFromLstDesc = function (nIdx) {
      if (nIdx > -1 && nIdx < this.desc.length) {
        this.desc[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_StyleDefHdrRemoveDesc, nIdx, [this.desc[nIdx]], false));
        this.desc.splice(nIdx, 1);
      }
    };

    StyleDefHdr.prototype.addToLstList = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_StyleDefHdrAddList, nInsertIdx, [oPr], true));
      this.list.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    StyleDefHdr.prototype.removeFromLstList = function (nIdx) {
      if (nIdx > -1 && nIdx < this.list.length) {
        this.list[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_StyleDefHdrRemoveList, nIdx, [this.list[nIdx]], false));
        this.list.splice(nIdx, 1);
      }
    };

    StyleDefHdr.prototype.getMinVer = function () {
      return this.minVer;
    }

    StyleDefHdr.prototype.getResId = function () {
      return this.resId;
    }

    StyleDefHdr.prototype.getUniqueId = function () {
      return this.uniqueId;
    }

    StyleDefHdr.prototype.getCatLst = function () {
      return this.catLst;
    }

    StyleDefHdr.prototype.getExtLst = function () {
      return this.extLst;
    }

    StyleDefHdr.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setMinVer(this.getMinVer());
      oCopy.setResId(this.getResId());
      oCopy.setUniqueId(this.getUniqueId());
      if (this.getCatLst()) {
        oCopy.setCatLst(this.getCatLst().createDuplicate(oIdMap));
      }
      if (this.getExtLst()) {
        oCopy.setExtLst(this.getExtLst().createDuplicate(oIdMap));
      }
      for (var nIdx = 0; nIdx < this.desc.length; ++nIdx) {
        oCopy.addToLstDesc(nIdx, this.desc[nIdx].createDuplicate(oIdMap));
      }
      for (nIdx = 0; nIdx < this.list.length; ++nIdx) {
        oCopy.addToLstList(nIdx, this.list[nIdx].createDuplicate(oIdMap));
      }
    }


    changesFactory[AscDFH.historyitem_ShapeSmartArtInfoSpPrPoint] = CChangeObject;
    changesFactory[AscDFH.historyitem_ShapeSmartArtInfoShapePoint] = CChangeObject;
    changesFactory[AscDFH.historyitem_ShapeSmartArtInfoAddLstContentPoint] = CChangeContent;
    changesFactory[AscDFH.historyitem_ShapeSmartArtInfoRemoveLstContentPoint] = CChangeContent;
    drawingsChangesMap[AscDFH.historyitem_ShapeSmartArtInfoSpPrPoint] = function (oClass, value) {
      oClass.spPrPoint = value;
    };
    drawingsChangesMap[AscDFH.historyitem_ShapeSmartArtInfoShapePoint] = function (oClass, value) {
      oClass.shapePoint = value;
    };
    drawingContentChanges[AscDFH.historyitem_ShapeSmartArtInfoAddLstContentPoint] = function (oClass) {
      return oClass.contentPoint;
    };
    drawingContentChanges[AscDFH.historyitem_ShapeSmartArtInfoRemoveLstContentPoint] = function (oClass) {
      return oClass.contentPoint;
    };

    function ShapeSmartArtInfo() {
      CBaseFormatObject.call(this);
      this.shapePoint = null;
      this.contentPoint = [];
    }
    InitClass(ShapeSmartArtInfo, CBaseFormatObject, AscDFH.historyitem_type_ShapeSmartArtInfo);

    ShapeSmartArtInfo.prototype.setShapePoint = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_ShapeSmartArtInfoShapePoint, this.shapePoint, oPr));
      this.shapePoint = oPr;
      this.setParentToChild(oPr);
    }

    ShapeSmartArtInfo.prototype.addToLstContentPoint = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.contentPoint.length, Math.max(0, nIdx));
      oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ShapeSmartArtInfoAddLstContentPoint, nInsertIdx, [oPr], true));
      this.contentPoint.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    }

    ShapeSmartArtInfo.prototype.removeFromLstContentPoint = function (nIdx) {
      if (nIdx > -1 && nIdx < this.contentPoint.length) {
        this.contentPoint[nIdx].setParent(null);
        oHistory.Add(new CChangeContent(this, AscDFH.historyitem_ShapeSmartArtInfoRemoveLstContentPoint, nIdx, [this.contentPoint[nIdx]], false));
        this.contentPoint.splice(nIdx, 1);
      }
    }

    changesFactory[AscDFH.historyitem_SmartArtColorsDef] = CChangeObject;
    changesFactory[AscDFH.historyitem_SmartArtDrawing] = CChangeObject;
    changesFactory[AscDFH.historyitem_SmartArtLayoutDef] = CChangeObject;
    changesFactory[AscDFH.historyitem_SmartArtDataModel] = CChangeObject;
    changesFactory[AscDFH.historyitem_SmartArtStyleDef] = CChangeObject;
    changesFactory[AscDFH.historyitem_SmartArtParent] = CChangeObject;
    changesFactory[AscDFH.historyitem_SmartArtType] = CChangeString;
    drawingsChangesMap[AscDFH.historyitem_SmartArtColorsDef] = function (oClass, value) {
      oClass.colorsDef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SmartArtType] = function (oClass, value) {
      oClass.type = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SmartArtDrawing] = function (oClass, value) {
      oClass.drawing = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SmartArtLayoutDef] = function (oClass, value) {
      oClass.layoutDef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SmartArtDataModel] = function (oClass, value) {
      oClass.dataModel = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SmartArtStyleDef] = function (oClass, value) {
      oClass.styleDef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SmartArtParent] = function (oClass, value) {
      oClass.parent = value;
    };

    function SmartArt() {
      CGroupShape.call(this);
      this.colorsDef = null;
      this.drawing = null;
      this.layoutDef = null;
      this.dataModel = null;
      this.styleDef = null;
      this.parent = null;
      this.type = null;

      this.calcGeometry = null;
    }

    InitClass(SmartArt, CGroupShape, AscDFH.historyitem_type_SmartArt);

    SmartArt.prototype.getObjectType = function() {
      return AscDFH.historyitem_type_SmartArt;
    };
    SmartArt.prototype.getName = function () {
      return 'SmartArt';
    };

    SmartArt.prototype.recalculate = function () {
    var oldParaMarks = editor && editor.ShowParaMarks;
      if (oldParaMarks) {
        editor.ShowParaMarks = false;
      }
      CGroupShape.prototype.recalculate.call(this);
      if (oldParaMarks) {
        editor.ShowParaMarks = oldParaMarks;
      }
    }

    SmartArt.prototype.handleUpdateExtents = function(bExt)
    {
      this.recalcTransform();
      this.recalcBounds();
      this.addToRecalculate();
      this.recalcWrapPolygon();
      if(this.spTree)
      {
        for(var i = 0; i < this.spTree.length; ++i)
        {
          this.spTree[i].handleUpdateExtents(bExt);
        }
      }
    };

    SmartArt.prototype.startAlgorithm = function (pointTree) {
      var layoutDef = this.getLayoutDef();
      if (layoutDef) {
        layoutDef.startAlgorithm(pointTree);
      }
    }

    SmartArt.prototype.getShapeMap = function () {
      var shapes = this.getDrawing() && this.getDrawing().spTree;
      var shapeMap = {};
      if (shapes) {
        shapes.forEach(function (shape) {
          shapeMap[shape.modelId] = shape;
        });
        return shapeMap;
      }
    }

    SmartArt.prototype.recalculateSmartArt = function () {
      var tree = this.createHierarchy();
      this.startAlgorithm(tree);
    }

    SmartArt.prototype.getPtMap = function () {
      var ptLst = this.getPtLst();
      var ptMap = {};
      if (ptLst) {
        ptLst.forEach(function (point) {
          ptMap[point.modelId] = point;
        });
        return ptMap;
      }
    }

    SmartArt.prototype.getPtLst = function () {
      var dataModel = this.getDataModel() && this.getDataModel().getDataModel();
      return dataModel && dataModel.getPtLst() && dataModel.getPtLst().list;
    }

    SmartArt.prototype.getCxnLst = function () {
      var dataModel = this.getDataModel() && this.getDataModel().getDataModel();
      return dataModel && dataModel.getCxnLst() && dataModel.getCxnLst().list;
    }

    SmartArt.prototype.createDataForHierarchy = function () {
      var shapeMap = this.getShapeMap();
      var ptMap = this.getPtMap();
      var ptLst = this.getPtLst();
      var cxnLst = this.getCxnLst();
      var elements = [];
      var nodePoints = [];
      var ptLstWithTypePres = [];
      var docPoint;
      if (cxnLst && ptLst) {
        var connectionsParOf = cxnLst.filter(function (cxn) {return !cxn.type;});

        ptLst.forEach(function (point) {
          if (point.type === Point_type_pres) {
            ptLstWithTypePres.push(point);
          } else if (!point.type || point.type === Point_type_node || point.type === Point_type_asst) {
            nodePoints.push(point);
          } else if (point.type === Point_type_doc) {
            docPoint = point;
          }
        });

        for (var i = 0; i <= nodePoints.length; i += 1) {
          var elem = new SmartArtNodeData();

          if (i === nodePoints.length) {
            var mPoint = docPoint;
            elem.setDocPoint(mPoint);

          } else {
            mPoint = nodePoints[i];
            if (!mPoint.type || mPoint.type === Point_type_node) {
              elem.setNodePoint(mPoint);
            } else if (mPoint.type === Point_type_asst) {
              elem.setAsstPoint(mPoint);
            }
          }

          connectionsParOf.forEach(function (cxn) {
            if (cxn.destId === mPoint.modelId) {
              elem.setCxn(cxn);
              if (ptMap) {
                elem.addToLstSibPoint(0, ptMap[cxn.sibTransId]);
                elem.addToLstParPoint(0, ptMap[cxn.parTransId]);
              }
            }
          });

          ptLstWithTypePres.forEach(function (pointWithTypePres) {
            var prSet = pointWithTypePres.prSet;
            if (prSet && (prSet.presAssocID === mPoint.modelId || (elem.sibPoint && prSet.presAssocID === elem.sibPoint.modelId) || (elem.parPoint && prSet.presAssocID === elem.parPoint.modelId))) {
              elem.addToLstPresPoint(elem.presPoint.length, pointWithTypePres);
              if (shapeMap && shapeMap[pointWithTypePres.modelId]) {
                elem.addToLstShapes(elem.shapes.length, shapeMap[pointWithTypePres.modelId]);
              }
            }
          });
          elements.push(elem);
        }
        return elements;
      }
    }

    SmartArt.prototype.createHierarchy = function () {
      var cxnLst = this.getCxnLst();

      if (cxnLst) {
        var cxnWithNoPres = cxnLst.filter(function (cxn) {return !cxn.type;});

        var elements = this.createDataForHierarchy();

        var root = elements.reduce(function (acc, next) {
          if (next.docPoint) {
            return next;
          }
          return acc;
        }, undefined);

        if (root) {
          var rootInfo = root.docPoint.modelId;
          var tree = new SmartArtTree(rootInfo, root, this);

          cxnWithNoPres = cxnWithNoPres.sort(function (a, b) {
            return parseInt(a.srcOrd) - parseInt(b.srcOrd);
          });
          for (var i = 0; i < cxnWithNoPres.length; i += 1) {
            for (var j = 0; j < cxnWithNoPres.length; j += 1) {
              var _cxn = cxnWithNoPres[j];
              var childData = elements.reduce(function (acc, next) {
                var nodePoint = next.nodePoint || next.asstPoint;
                if (nodePoint && nodePoint.modelId === _cxn.destId) {
                  return next;
                }
                return acc;
              }, undefined);
              tree.add(_cxn.destId, _cxn.srcId, childData);
            }
          }
        }
      }
      return tree;
    }

    SmartArt.prototype.getDefColorsByName = function () {
      var colorsDef = this.getColorsDef();
      return colorsDef && colorsDef.styleLblByName;
    }

    SmartArt.prototype.getDefaultColorsForPoint = function (point) {
      var styleLbl = point.getPresStyleLbl();
      var defaultColors = this.getDefColorsByName();
      if (defaultColors && styleLbl) {
        return defaultColors[styleLbl];
      }
    }

    SmartArt.prototype.getDefaultTxColorFromPoint = function (point) {
      var currentDefaultColors = this.getDefaultColorsForPoint(point);
      if (currentDefaultColors) {
        var txFillLst = currentDefaultColors.txFillClrLst;
        if (txFillLst) {
          return txFillLst.list[0];
        }
      }
    }

    SmartArt.prototype.getSmartArtDefaultTxFill = function (shape) {
      var shapePoint = shape && shape.getSmartArtShapePoint();
      var defaultTxColorFromShape = shapePoint && this.getDefaultTxColorFromPoint(shapePoint);
      var defaultTxFill;

      if (defaultTxColorFromShape) {
        defaultTxFill = AscFormat.CreateUniFillByUniColorCopy(defaultTxColorFromShape);
      }
      return defaultTxFill;
    }

    SmartArt.prototype.getTypeOfSmartArt = function () {
      // Russian name -> type
      //
      //
      // Акцентируемый рисунок -> AccentedPicture
      // Баланс -> balance1
      // Блоки рисунков с названиями -> TitledPictureBlocks
      // Блоки со смещенными рисунками -> PictureAccentBlocks
      // Блочный цикл -> cycle5
      // Венна в столбик -> venn2
      // Вертикальное уравнение -> equation2
      // Вертикальный блочный список -> vList5
      // Вертикальный ломаный процесс -> bProcess4
      // Вертикальный маркированный список -> vList2
      // Вертикальный нелинейный список -> VerticalCurvedList
      // Вертикальный процесс -> process2
      // Вертикальный список -> list1
      // Вертикальный список рисунков -> vList4
      // Вертикальный список с кругами -> VerticalCircleList
      // Вертикальный список со смещенными рисунками -> vList3
      // Вертикальный список со стрелкой -> vList6
      // Вертикальный уголковый список -> chevron2
      // Вертикальный уголковый список2 -> VerticalAccentList
      // Вложенная целевая -> target2
      // Воронка -> funnel1
      // Восходящая стрелка -> arrow2
      // Восходящая стрелка процесса -> IncreasingArrowsProcess
      // Восходящий процесс -> StepUpProcess
      // Выноска с круглыми рисунками -> CircularPictureCallout
      // Горизонтальная иерархия -> hierarchy2
      // Горизонтальная иерархия с подписями -> hierarchy5
      // Горизонтальная многоуровневая иерархия -> HorizontalMultiLevelHierarchy
      // Горизонтальная организационная диаграмма -> HorizontalOrganizationChart
      // Горизонтальный маркированный список -> hList1
      // Горизонтальный список рисунков -> pList2
      // Закрытый уголковый процесс -> hChevron3
      // Иерархический список -> hierarchy3
      // Иерархия -> hierarchy1
      // Иерархия с круглыми рисунками -> CirclePictureHierarchy
      // Иерархия с подписями -> hierarchy6
      // Инвертированная пирамида -> pyramid3
      // Кластер шестиугольников -> HexagonCluster
      // Круг связей -> CircleRelationship
      // Круглая временная шкала -> CircleAccentTimeline
      // Круглый ломаный процесс -> bProcess2
      // Лента со стрелками -> arrow6
      // Линейная Венна -> venn3
      // Линия рисунков -> PictureLineup
      // Линия рисунков с названиями -> TitlePictureLineup
      // Ломаный список рисунков с подписями -> BendingPictureCaptionList
      // Ломаный список со смещенными рисунками -> bList2
      // Матрица с заголовками -> matrix1
      // Нарастающий процесс с кругами -> IncreasingCircleProcess
      // Нелинейные рисунки с блоками -> BendingPictureBlocks
      // Нелинейные рисунки с подписями -> BendingPictureCaption
      // Нелинейные рисунки с полупрозрачным текстом -> BendingPictureSemiTransparentText
      // Ненаправленный цикл -> cycle6
      // Непрерывный блочный процесс -> hProcess9
      // Непрерывный список с рисунками -> hList7
      // Непрерывный цикл -> cycle3
      // Нисходящий блочный список -> BlockDescendingList
      // Нисходящий процесс -> StepDownProcess
      // Обратный список -> ReverseList
      // Организационная диаграмма -> orgChart1
      // Организационная диаграмма с именами и должностями -> NameandTitleOrganizationalChart
      // Переменный поток -> hProcess4
      // Пирамидальный список -> pyramid2
      // Плюс и минус -> PlusandMinus
      // Повторяющийся ломаный процесс -> bProcess3
      // Подписанные рисунки -> CaptionedPictures
      // Подробный процесс -> hProcess7
      // Полосы рисунков -> PictureStrips
      // Полукруглая организационная диаграмма -> HalfCircleOrganizationChart
      // Поэтапный процесс -> PhasedProcess
      // Простая Венна -> venn1
      // Простая временная шкала -> hProcess11
      // Простая круговая -> chart3
      // Простая матрица -> matrix3
      // Простая пирамида -> pyramid1
      // Простая радиальная -> radial1
      // Простая целевая -> target1
      // Простой блочный список -> default
      // Простой ломаный процесс -> process5
      // Простой процесс -> process1
      // Простой уголковый процесс -> chevron1
      // Простой цикл -> cycle2
      // Противоположные идеи -> OpposingIdeas
      // Противостоящие стрелки -> arrow4
      // Процесс от случайности к результату -> RandomtoResultProcess
      // Процесс с вложенными шагами -> SubStepProcess
      // Процесс с круговой диаграммой -> PieProcess
      // Процесс со смещением -> process3
      // Процесс со смещенными по возрастанию рисунками -> AscendingPictureAccentProcess
      // Процесс со смещенными рисунками -> hProcess10
      // Радиальная Венна -> radial3
      // Радиальная циклическая -> radial6
      // Радиальный кластер -> RadialCluster
      // Радиальный список -> radial2
      // Разнонаправленный цикл -> cycle7
      // Расходящаяся радиальная -> radial5
      // Расходящиеся стрелки -> arrow1
      // Рисунок с текстом в рамке -> FramedTextPicture
      // Сгруппированный список -> lProcess2
      // Сегментированная пирамида -> pyramid4
      // Сегментированный процесс -> process4
      // Сегментированный цикл -> cycle8
      // Сетка рисунков -> PictureGrid
      // Сетчатая матрица -> matrix2
      // Спираль рисунков -> SpiralPicture
      // Список в столбик -> hList9
      // Список названий рисунков -> pList1
      // Список процессов -> lProcess1
      // Список рисунков с выносками -> BubblePictureList
      // Список с квадратиками -> SquareAccentList
      // Список с линиями -> LinedList
      // Список со смещенными рисунками -> hList2
      // Список со смещенными рисунками и заголовком -> PictureAccentList
      // Список со снимками -> SnapshotPictureList
      // Стрелка непрерывного процесса -> hProcess3
      // Стрелка процесса с кругами -> CircleArrowProcess
      // Стрелки процесса -> hProcess6
      // Ступенчатый процесс -> vProcess5
      // Сходящаяся радиальная -> radial4
      // Сходящиеся стрелки -> arrow5
      // Табличная иерархия -> hierarchy4
      // Табличный список -> hList3
      // Текстовый цикл -> cycle1
      // Трапецевидный список -> hList6
      // Убывающий процесс -> DescendingProcess
      // Уголковый список -> lProcess3
      // Уравнение -> equation1
      // Уравновешивающие стрелки -> arrow3
      // Целевой список -> target3
      // Циклическая матрица -> cycle4
      // Чередующиеся блоки рисунков -> AlternatingPictureBlocks
      // Чередующиеся круги рисунков -> AlternatingPictureCircles
      // Чередующиеся шестиугольники -> AlternatingHexagons
      // Шестеренки -> gear1
      var dataModel = this.getDataModel() && this.getDataModel().getDataModel();
      var ptLst = dataModel.ptLst.list;
      var type;
      ptLst.forEach(function (point) {
        if (point.type === Point_type_doc) {
          if (point.prSet && point.prSet.loTypeId) {
            var typeSplit = point.prSet.loTypeId.split('/');
            type = typeSplit[typeSplit.length - 1];
            type = type.split('#')[0];
          }
        }
      });
      return type;
    };

    SmartArt.prototype.getShapesForFitText = function (callShape) {

      var smartArtType = this.getTypeOfSmartArt();
      var shapeGeometry = callShape.getPresetGeom();
      var shapes = this.arrGraphicObjects.slice();
      var callShapePoint = callShape.getSmartArtShapePoint();
      var prSet = callShapePoint && callShapePoint.prSet;
      var getShapesFromPresStyleLbl = function(arrOfStyleLbl, returnThis) {
        if (prSet) {
          for (var i = 0; i < arrOfStyleLbl.length; i += 1) {
            if (prSet.presStyleLbl === arrOfStyleLbl[i]) {
              return shapes.filter(function (shape) {
                var smartArtShapePoint = shape.getSmartArtShapePoint();
                var shapePrSet = smartArtShapePoint && smartArtShapePoint.prSet;
                return shapePrSet.presStyleLbl === arrOfStyleLbl[i];
              });
            }
          }
        }
        return returnThis ? [callShape] : shapes;
      }
      var getShapesFromPresName = function(arrOfPresName) {
        if (prSet) {
          for (var i = 0; i < arrOfPresName.length; i += 1) {
            if (typeof prSet.presName === 'string' && prSet.presName.includes(arrOfPresName[i])) {
              return shapes.filter(function (shape) {
                var smartArtShapePoint = shape.getSmartArtShapePoint();
                var shapePrSet = smartArtShapePoint && smartArtShapePoint.prSet;
                return typeof shapePrSet.presName === 'string' && shapePrSet.presName.includes(arrOfPresName[i]);
              });
            }
          }
        }
        return shapes;
      }
      var getShapesFromPresetGeom = function() {
        var result = shapes.filter(function (shape) {
          return shape.getPresetGeom() === shapeGeometry;
        });
        return result.length === 0 ? shapes : result;
      }

      switch (smartArtType) {
        case "PictureAccentBlocks":
        case "cycle5":
        case "venn2":
        case "bProcess4":
        case "vList2":
        case "VerticalCurvedList":
        case "process2":
        case "list1":
        case "vList4":
        case "VerticalCircleList":
        case "arrow2":
        case "StepUpProcess":
        case "hierarchy2":
        case "HorizontalMultiLevelHierarchy":
        case "HorizontalOrganizationChart":
        case "hList1":
        case "pList2":
        case "hChevron3":
        case "hierarchy1":
        case "CirclePictureHierarchy":
        case "HexagonCluster":
        case "CircleRelationship":
        case "CircleAccentTimeline":
        case "bProcess2":
        case "arrow6":
        case "venn3":
        case "PictureLineup":
        case "BendingPictureCaptionList":
        case "matrix1":
        case "BendingPictureBlocks":
        case "BendingPictureCaption":
        case "BendingPictureSemiTransparentText":
        case "cycle6":
        case "hProcess9":
        case "hList7":
        case "cycle3":
        case "StepDownProcess":
        case "ReverseList":
        case "orgChart1":
        case "pyramid2":
        case "PlusandMinus":
        case "bProcess3":
        case "CaptionedPictures":// TODO: think about it
        case "PictureStrips":
        case "HalfCircleOrganizationChart":
        case "venn1":
        case "hProcess11":
        case "chart3":
        case "matrix3":
        case "target1":
        case "default":
        case "process5":
        case "process1":
        case "chevron1":
        case "cycle2":
        case "arrow4":
        case "RandomtoResultProcess": // TODO: recalculate in next
        case "process3":
        case "hProcess10":
        case "radial6": // TODO: think about it
        case "cycle7":
        case "arrow1":
        case "FramedTextPicture":
        case "pyramid4":
        case "cycle8":
        case "PictureGrid":
        case "matrix2":
        case "SpiralPicture":
        case "pList1":
        case "BubblePictureList":
        case "SnapshotPictureList": // TODO: think about it
        case "hProcess3":
        case "CircleArrowProcess": // TODO: think about it
        case "vProcess5":
        case "radial4":
        case "arrow5":
        case "hierarchy4":
        case "cycle1":
        case "hList6":
        case "DescendingProcess":
        case "equation1":
        case "arrow3":
        case "AlternatingPictureBlocks":
        case "AlternatingPictureCircles":
        case "gear1":
          return shapes;
          break;
        case "AlternatingHexagons":
          return getShapesFromPresetGeom(['rect']);
        case "LinedList":
          return getShapesFromPresName(['tx1', 'tx2', 'tx3', 'tx4']);
        case "SquareAccentList":
        case "IncreasingCircleProcess":
        case "PieProcess":
          return getShapesFromPresName(['Child', 'Parent']);
        case "hList2":
          return getShapesFromPresStyleLbl(['node1', 'revTx']);
        case "lProcess2": // TODO: check transform
          return getShapesFromPresStyleLbl(['bgShp', 'node1']);
        case "PictureAccentList":
          return getShapesFromPresStyleLbl(['lnNode1']); // TODO: think about it
        case "vList5":
        case "chevron2":
        case "bList2":
        case "hList9":
        case "hProcess7":
        case "vList6":
        case "hProcess6":
        case "SubStepProcess":
        case "funnel1":
        case "PhasedProcess":
        case "cycle4":
        case "pyramid1":
        case "pyramid3":
        case "vList3":
        case "radial2":
        case "TitledPictureBlocks":
        case "OpposingIdeas":
        case "hierarchy6":
        case "hierarchy5":
        case "IncreasingArrowsProcess":
          return getShapesFromPresetGeom();
        case "VerticalAccentList":
          return getShapesFromPresStyleLbl(['revTx', 'solidFgAcc1']);
        case "BlockDescendingList":
          return getShapesFromPresName(['childText']);
        case "hList3":
          return getShapesFromPresStyleLbl(['node1', 'dkBgShp']);
        case "process4":
          return getShapesFromPresStyleLbl(['node1', 'fgAccFollowNode1']);
        case "target3":
          return getShapesFromPresName(['СhTx', 'rect']); // TODO: think about it
        case "hierarchy3":
          return getShapesFromPresStyleLbl(['node1', 'bgAcc1']);
        case "hProcess4":
          return getShapesFromPresStyleLbl(['node1', 'bgAcc1']);
        case "lProcess3":
          return getShapesFromPresStyleLbl(['node1', 'alignAccFollowNode1']);
        case "lProcess1":
          return getShapesFromPresStyleLbl(['alignAccFollowNode1', 'node1']);
        case "AscendingPictureAccentProcess":
          return getShapesFromPresStyleLbl(['node1', 'revTx']);
        case "equation2":
          return getShapesFromPresName(['lastNode', 'node']);
        case "radial1":
        case "radial5": // TODO: think
          return getShapesFromPresStyleLbl(['node1', 'node0']);
        case "radial3":
          return getShapesFromPresName(['centerShape', 'node']);
        case "RadialCluster":
          return [callShape];
        case "NameandTitleOrganizationalChart":
          return getShapesFromPresStyleLbl(['node0'], true);
        case "balance1":
          return getShapesFromPresStyleLbl(['alignAccFollowNode1', 'node1']);
        case "target2":
          return getShapesFromPresStyleLbl(['node1', 'fgAcc1']);
        case "AccentedPicture":
        case "CircularPictureCallout":
          return getShapesFromPresStyleLbl(['revTx', 'node1']);
        case "TitlePictureLineup":
          return getShapesFromPresStyleLbl(['revTx', 'alignNode1']);
        default:
          return [];
      }
    }


    SmartArt.prototype.setParent = function (parent) {
      History.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SmartArtParent, this.parent, parent));
      this.parent = parent;
    };
    SmartArt.prototype.setColorsDef = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SmartArtColorsDef, this.getColorsDef(), oPr));
      this.colorsDef = oPr;
      oPr.setParent(this);
    };
    SmartArt.prototype.setType = function (oPr) {
      oHistory.Add(new CChangeString(this, AscDFH.historyitem_SmartArtType, this.type, oPr));
      this.type = oPr;
    };

    SmartArt.prototype.setDrawing = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SmartArtDrawing, this.getDrawing(), oPr));
      this.drawing = oPr;
      oPr.setParent(this);
    };
    SmartArt.prototype.setLayoutDef = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SmartArtLayoutDef, this.getLayoutDef(), oPr));
      this.layoutDef = oPr;
      oPr.setParent(this);
    };
    SmartArt.prototype.setDataModel = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SmartArtDataModel, this.getDataModel(), oPr));
      this.dataModel = oPr;
      oPr.setParent(this);
    };
    SmartArt.prototype.setStyleDef = function (oPr) {
      oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SmartArtStyleDef, this.getStyleDef(), oPr));
      this.styleDef = oPr;
      oPr.setParent(this);
    };
    SmartArt.prototype.getColorsDef = function () {
      return this.colorsDef;
    };
    SmartArt.prototype.getDrawing = function () {
      return this.drawing;
    };
    SmartArt.prototype.getLayoutDef = function () {
      return this.layoutDef;
    };
    SmartArt.prototype.getDataModel = function () {
      return this.dataModel;
    };
    SmartArt.prototype.getStyleDef = function () {
      return this.styleDef;
    };
    SmartArt.prototype.fillObject = function (oCopy, oIdMap) {
      if (this.getColorsDef()) {
        oCopy.setColorsDef(this.getColorsDef().createDuplicate(oIdMap));
      }
      if (this.getDrawing()) {
        oCopy.setDrawing(this.getDrawing().createDuplicate(oIdMap));
      }
      if (this.getLayoutDef()) {
        oCopy.setLayoutDef(this.getLayoutDef().createDuplicate(oIdMap));
      }
      if (this.getDataModel()) {
        oCopy.setDataModel(this.getDataModel().createDuplicate(oIdMap));
      }
      if (this.getStyleDef()) {
        oCopy.setStyleDef(this.getStyleDef().createDuplicate(oIdMap));
      }
    };
    SmartArt.prototype.createPlaceholderControl = function (aControls) {
      if(this.drawing) {
        this.drawing.createPlaceholderControl(aControls);
      }
    };

    SmartArt.prototype.getRelationOfContent = function () {
      var dataModel = this.getDataModel() && this.getDataModel().getDataModel();
      if (dataModel) {
        var connections = {};
        var ptMap = this.getPtMap();
        var shapeMap = this.getShapeMap();
        var cxnLst = dataModel.cxnLst.list;
        var presCxnLst = cxnLst.filter(function (cxn) {
          return cxn.type === 'presOf';
        });

        presCxnLst.forEach(function (cxn) {
          var shape = shapeMap[cxn.destId];
          if (shape) {
            if (!connections[cxn.destId]) {
              connections[cxn.destId] = [];
            }
            connections[cxn.destId].push({
              point: ptMap[cxn.srcId],
              srcOrd: parseInt(cxn.srcOrd, 10),
              destOrd: parseInt(cxn.destOrd, 10)
            });
          }
        });

        for (var key in connections) {
          connections[key].sort(function (firstConnection, secondConnection) {
            return firstConnection.destOrd - secondConnection.destOrd;
          });
        }

        return connections;
      }
    };

    SmartArt.prototype.getRelationOfShapes = function () {
      var dataModel = this.getDataModel() && this.getDataModel().getDataModel();
      if (dataModel) {
        var connections = {};
        var ptMap = this.getPtMap();
        var shapeMap = this.getShapeMap();
        var cxnLst = dataModel.cxnLst.list;
        var presCxnLst = cxnLst.filter(function (cxn) {
          return cxn.type === 'presOf' || cxn.type === 'presParOf';
        });

        presCxnLst.forEach(function (cxn) {
          var shape = shapeMap[cxn.destId];
          if (shape) {
            if (!connections[cxn.destId]) {
              connections[cxn.destId] = ptMap[cxn.destId];
            }
          }
        });
        return connections;
      }
    };

    SmartArt.prototype.setConnections2 = function () {
      var dataModel = this.getDataModel() && this.getDataModel().getDataModel();
      if (dataModel) {

        var shapeMap = this.getShapeMap();
        var contentConnections = this.getRelationOfContent();
        var shapeConnections = this.getRelationOfShapes();

        for (var modelId in shapeMap) {
          var shape = shapeMap[modelId];
          var smartArtInfo = new ShapeSmartArtInfo();
          shape.setShapeSmartArtInfo(smartArtInfo);
          if (contentConnections[modelId]) {
            contentConnections[modelId].forEach(function (el) {
              smartArtInfo.addToLstContentPoint(smartArtInfo.contentPoint.length, el.point);
            });
          }
          if (shapeConnections[modelId]) {
            smartArtInfo.setShapePoint(shapeConnections[modelId]);
          }
        }
      }
    };

    SmartArt.prototype.getShapeMap = function () {
      var result = {};
      var shapeTree = this.getDrawing().spTree;
      shapeTree.forEach(function (shape) {
        result[shape.modelId] = shape;
      });
      return result;
    };

    SmartArt.prototype.privateWriteAttributes = null;
    SmartArt.prototype.writeChildren = function(pWriter) {
      pWriter.StartRecord(0);
      pWriter.StartRecord(0);
      pWriter.WriteGroupShape(this.drawing);
      pWriter.EndRecord();
      pWriter.EndRecord();
      this.writeRecord2(pWriter, 1, this.dataModel);
      this.writeRecord2(pWriter, 2, this.colorsDef);
      this.writeRecord2(pWriter, 3, this.layoutDef);
      this.writeRecord2(pWriter, 4, this.styleDef);
    };
    SmartArt.prototype.readAttribute = null;
    SmartArt.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setDrawing(new Drawing());
          pReader.ReadSmartArtGroup(this.drawing);
          this.drawing.setGroup(this);
          this.addToSpTree(0, this.drawing);
          break;
        }
        case 1: {
          this.setDataModel(new DiagramData());
          this.dataModel.fromPPTY(pReader);
          this.setConnections2();
          break;
        }
        case 2: {
          this.setColorsDef(new ColorsDef());
          this.colorsDef.fromPPTY(pReader);
          break;
        }
        case 3: {
          this.setLayoutDef(new LayoutDef());
          this.layoutDef.fromPPTY(pReader);
          break;
        }
        case 4: {
          this.setStyleDef(new StyleDef());
          this.styleDef.fromPPTY(pReader);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    SmartArt.prototype.getChildren = function() {
      return [this.drawing, this.dataModel, this.colorsDef, this.layoutDef, this.styleDef];
    };

    SmartArt.prototype.switchRightToLeft = function () {
      var isRightToLeft;
      var point = this.getNullNamePoint();
      var pointDir = point && point.prSet && point.prSet.presLayoutVars && point.prSet.presLayoutVars.dir && (point.prSet.presLayoutVars.dir.val === 'norm' || point.prSet.presLayoutVars.dir.val === null);
      if (pointDir) {
        isRightToLeft = true;
        point.prSet.presLayoutVars.dir.setVal('rev');
      } else {
        isRightToLeft = false;
        if (!point.prSet) {
          point.setPrSet(new PrSet());
        }
        if (!point.prSet.presLayoutVars) {
          point.prSet.setPresLayoutVars(new VarLst());
        }
        if (!point.prSet.presLayoutVars.dir) {
          point.prSet.presLayoutVars.setDir(new DiagramDirection());
        }
        point.prSet.presLayoutVars.dir.setVal('norm');
      }
      return isRightToLeft;
    }

    SmartArt.prototype.getNullNamePoint = function () {
      var dataModel = this.getDataModel() && this.getDataModel().getDataModel();
      var ptLst = dataModel.ptLst.list;
      for (var i = 0; i < ptLst.length; i += 1) {
        if (ptLst[i].type === Point_type_pres && ptLst[i].prSet && ptLst[i].prSet.presName === 'Name0') { // TODO: ptLst type = 'pres' is 4
          return ptLst[i];
        }
      }
    }

    SmartArt.prototype.isPlaceholder = function () {
      return false;
    };
    SmartArt.prototype.canRotate = function () {
      return false;
    };
    SmartArt.prototype.canFill = function () {
      return true;
    };
    SmartArt.prototype.updateCoordinatesAfterInternalResize = function () {
      if(!this.spPr) {
        this.setSpPr(new AscFormat.CSpPr());
        this.spPr.setParent(this);
      }
      if(!this.spPr.xfrm) {
        this.spPr.setXfrm(new AscFormat.CXfrm());
        this.spPr.xfrm.setParent(this.spPr);
      }
      var oXfrm = this.spPr.xfrm;
      if(AscCommonWord.ParaDrawing && (this.parent instanceof AscCommonWord.ParaDrawing)) {
        oXfrm.setOffX(0);
        oXfrm.setOffY(0);
      }
      else {
        oXfrm.setOffX(this.x);
        oXfrm.setOffY(this.y);
      }
      oXfrm.setChOffX(0);
      oXfrm.setChOffY(0);
      oXfrm.setChExtX(this.extX);
      oXfrm.setChExtY(this.extY);
      oXfrm.setExtX(this.extX);
      oXfrm.setExtY(this.extY);
      return {posX: oXfrm.offX, posY: oXfrm.offY};
    };

    SmartArt.prototype.setXfrmByParent = function () {
      var oXfrm = this.spPr.xfrm;
      if (oXfrm.isZero()) {
        var parent = this.parent;
        if (parent instanceof AscCommonWord.ParaDrawing) {
          oXfrm.setExtX(parent.Extent.W);
          oXfrm.setExtY(parent.Extent.H);
        }
      }
      for (var i = 0; i < this.spTree.length; i += 1) {
        this.spTree[i].setXfrmByParent();
      }
    };

    SmartArt.prototype.recalculateTransform = function() {
      var oThis = this;
      AscFormat.ExecuteNoHistory(function(){
        AscFormat.CGroupShape.prototype.recalculateTransform.call(this);
        this.calcGeometry = AscFormat.CreateGeometry("rect");
        this.calcGeometry.Recalculate(this.extX, this.extY);
      }, this, []);
    };
    SmartArt.prototype.draw = function(graphics) {
      if(this.checkNeedRecalculate()){
        return;
      }

      var oldParaMarks = editor && editor.ShowParaMarks;
      if (oldParaMarks) {
        editor.ShowParaMarks = false;
      }

      if(this.calcGeometry) {
        graphics.SaveGrState();
        graphics.SetIntegerGrid(false);
        graphics.transform3(this.transform, false);
        var oDrawer = new AscCommon.CShapeDrawer();
        oDrawer.fromShape2(this, graphics, this.calcGeometry);
        oDrawer.draw(this.calcGeometry);
        graphics.RestoreGrState();
      }
      AscFormat.CGroupShape.prototype.draw.call(this, graphics);
      if (oldParaMarks) {
        editor.ShowParaMarks = oldParaMarks;
      }
    };
    SmartArt.prototype.getBg = function() {
      var oDataModel = this.getDataModel() && this.getDataModel().getDataModel();
      if(!oDataModel) {
        return;
      }
      return oDataModel.bg;
    };
    SmartArt.prototype.getWhole = function() {
      var oDataModel = this.getDataModel() && this.getDataModel().getDataModel();
      if(!oDataModel) {
        return;
      }
      return oDataModel.whole;
    };
    SmartArt.prototype.recalculateBrush = function () {
      this.brush = null;
      var oBg = this.getBg();
      if(!oBg) {
        return;
      }
      if(!oBg.fill) {
        return;
      }
      this.brush = oBg.fill.createDuplicate();
      var oParents = this.getParentObjects();
      var RGBA = {R: 0, G: 0, B: 0, A: 255};
      this.brush.calculate(oParents.theme, oParents.slide, oParents.layout, oParents.master, RGBA);
    };

    SmartArt.prototype.recalculatePen = function () {
      this.pen = null;
      var oWhole = this.getWhole();
      if(!oWhole) {
        return;
      }
      if(!oWhole.ln) {
        return;
      }
      this.pen = oWhole.ln.createDuplicate();
      if(this.pen.Fill) {
        var oParents = this.getParentObjects();
        var RGBA = {R: 0, G: 0, B: 0, A: 255};
        this.pen.calculate(oParents.theme, oParents.slide, oParents.layout, oParents.master, RGBA);
      }
    };
    SmartArt.prototype.changeFill = function (fill) {
      var oBg = this.getBg();
      if(!oBg) {
        return;
      }
      if(this.recalcInfo.recalculateBrush)
      {
        this.recalculateBrush();
      }
      var oUniFill = AscFormat.CorrectUniFill(fill, this.brush, this.getEditorType());
      oUniFill.convertToPPTXMods();
      oBg.setFill(oUniFill);
    };
    SmartArt.prototype.changeLine = function (line) {
      var oWhole = this.getWhole();
      if(!oWhole) {
        return;
      }
      if(this.recalcInfo.recalculatePen) {
        this.recalculatePen();
      }
      var stroke = AscFormat.CorrectUniStroke(line, this.pen);
      if(stroke.Fill) {
        stroke.Fill.convertToPPTXMods();
      }
      oWhole.setLn(stroke);
    };
    SmartArt.prototype.changeShadow = function (oShadow) {
      var oBg = this.getBg();
      if(!oBg) {
        return;
      }
      if(oShadow) {
        var oEffectProps = oBg.effect ? oBg.effect.createDuplicate() : new AscFormat.CEffectProperties();
        if(!oEffectProps.EffectLst) {
          oEffectProps.EffectLst = new AscFormat.CEffectLst();
        }
        oEffectProps.EffectLst.outerShdw = oShadow.createDuplicate();
        oBg.setEffect(oEffectProps);
      }
      else {
        if(oBg.effect) {
          if(oBg.effect.EffectLst) {
            if(oBg.effect.EffectLst.outerShdw) {
              var oEffectProps = oBg.effect.createDuplicate();
              oEffectProps.EffectLst.outerShdw = null;
              oBg.setEffect(oEffectProps);
            }
          }
        }
      }
    };

    SmartArt.prototype.getOuterShdw = function(){
      var oBg = this.getBg();
      if(oBg && oBg.effect && oBg.effect.EffectLst && oBg.effect.EffectLst.outerShdw) {
        return oBg.effect.EffectLst.outerShdw;
      }
      return null;
    };
    SmartArt.prototype.fromPPTY = function(pReader) {
      var oStream = pReader.stream;
      var nStart = oStream.cur;
      var nEnd = nStart + oStream.GetULong() + 4;
      this.readAttributes(pReader);
      this.readChildren(nEnd, pReader);
      oStream.Seek2(nEnd);
    };
    SmartArt.prototype.readAttributes = function(pReader) {
    };

    SmartArt.prototype.readChildren = function(nEnd, pReader) {
      var oStream = pReader.stream;
      while (oStream.cur < nEnd) {
        var nType = oStream.GetUChar();
        this.readChild(nType, pReader);
      }
    };

    SmartArt.prototype.toPPTY = function(pWriter) {
      this.writeAttributes(pWriter);
      this.writeChildren(pWriter);
    };
    SmartArt.prototype.writeAttributes = function(pWriter) {
    };

    SmartArt.prototype.copy = function(oPr)
    {
      var copy = new SmartArt();
      this.copy2(copy, oPr);
      var drawing = copy.getDrawing();
      if (drawing) {
        for (var i = 0; i < drawing.spTree.length; i += 1) {
          var obj = drawing.spTree[i];
          if (obj.getObjectType() === AscDFH.historyitem_type_Shape) {
            obj.copyTextInfoFromShapeToPoint();
          }
        }
      }
      return copy;
    };
    SmartArt.prototype.copy2 = function(copy, oPr)
    {
      if(this.nvGrpSpPr)
      {
        copy.setNvGrpSpPr(this.nvGrpSpPr.createDuplicate());
      }
      if(this.spPr)
      {
        copy.setSpPr(this.spPr.createDuplicate());
        copy.spPr.setParent(copy);
      }

      copy.setBDeleted(this.bDeleted);
      if(this.macro !== null) {
        copy.setMacro(this.macro);
      }
      if(this.textLink !== null) {
        copy.setTextLink(this.textLink);
      }
      if (this.drawing) {
        copy.setDrawing(this.drawing.copy(oPr));
        copy.addToSpTree(0, copy.drawing);
        copy.drawing.setGroup(copy);
      }
      if (this.layoutDef) {
        copy.setLayoutDef(this.layoutDef.createDuplicate());
      }
      if (this.styleDef) {
        copy.setStyleDef(this.styleDef.createDuplicate());
      }
      if (this.dataModel) {
        copy.setDataModel(this.dataModel.createDuplicate());
      }
      if (this.colorsDef) {
        copy.setColorsDef(this.colorsDef.createDuplicate());
      }
      copy.cachedImage = this.getBase64Img();
      copy.cachedPixH = this.cachedPixH;
      copy.cachedPixW = this.cachedPixW;
      copy.setLocks(this.locks);
      copy.setConnections2();
      return copy;
    };
    SmartArt.prototype.handleUpdateFill = function() {
      this.recalcInfo.recalculateBrush = true;
      CGroupShape.prototype.handleUpdateFill.call(this);
      this.addToRecalculate();
    };
    SmartArt.prototype.handleUpdateLn = function() {
      this.recalcInfo.recalculatePen = true;
      CGroupShape.prototype.handleUpdateLn.call(this);
      this.addToRecalculate();
    };

    SmartArt.prototype.convertToWord = function(document) {
      var oCopy = this.copy();
      oCopy.setBDeleted2(false);
      return oCopy;
    };
    SmartArt.prototype.convertToPPTX = function(drawingDocument, worksheet) {
      var oCopy = this.copy();
      oCopy.setBDeleted2(false);
      return oCopy;
    };
    SmartArt.prototype.getTypeName = function() {
      return AscCommon.translateManager.getValue("Diagram");
    };

    function SmartArtTree(rootInfo, rootData, parent) {
      CBaseFormatObject.call(this);
      var child = new SmartArtNode();
      if (rootInfo) {
        child.setInfo(rootInfo);
      }
      if (rootData) {
        child.setData(rootData);
      }
      child.setParent(this);
      this.parent = parent;
      child.depth = 0;
      this.root = child;
    }
    InitClass(SmartArtTree, CBaseFormatObject, AscDFH.historyitem_type_SmartArtTree);

    SmartArtTree.prototype.traverseDF = function (callback) {
      (function recurse(currentNode) {
        for (var i = 0; i < currentNode.children.length; i += 1) {
          recurse(currentNode.children[i]);
        }
        callback(currentNode);
      })(this.root);
    }
    SmartArtTree.prototype.setRoot = function (oPr) {
      this.root = oPr;
      oPr.setParent(this);
    }

    SmartArtTree.prototype.traverseBF = function (callback) {
      var queue = [];
      var currentTree = this.root;

      while (currentTree) {
        for (var i = 0; i < currentTree.children.length; i += 1) {
          queue.push(currentTree.children[i]);
        }
        callback(currentTree);
        currentTree = queue.shift();
      }
    }

    SmartArtTree.prototype.findNodeByNameAndStyleLbl = function (name, styleLbl) {
      var layoutNode = [];
      function callback(node) {
        if (node.data) {
          if (name && styleLbl) {
            var check = node.data.presPoint.some(function (point) {
              if (point.prSet) {
                return point.prSet.presName === name && point.prSet.styleLbl === styleLbl;
              }
            });
            if (check) {
              layoutNode.push(node.data);
            }
          } else if (name) {
            var check = node.data.presPoint.some(function (point) {
              if (point.prSet) {
                return point.prSet.presName === name;
              }
            });
            if (check) {
              layoutNode.push(node.data);
            }
          } else if (styleLbl) {
            var check = node.data.presPoint.some(function (point) {
              if (point.prSet) {
                return point.prSet.styleLbl === styleLbl;
              }
            });
            if (check) {
              layoutNode.push(node.data);
            }
          }
        }
      }
      this.contains(callback, true);
      return layoutNode;
    }

    SmartArtTree.prototype.contains = function (callback, isTraverseDF) {
      var traverse = isTraverseDF ? this.traverseDF : this.traverseBF;
      traverse.call(this, callback);
    }

    SmartArtTree.prototype.add = function (childInfo, parentInfo, childData, isTraverseDF) {
      var traverse = isTraverseDF ? this.traverseDF : this.traverseBF;
      var parent;

      var callback = function (node) {
        if (node.info === parentInfo) {
          parent = node;
        }
      };
      this.contains(callback, traverse === this.traverseDF);

      if (parent) {
        var parentHaveChild = parent.children.some(function (ch) {
          return ch.info === childInfo;
        });
        if (!parentHaveChild) {
          var child = new SmartArtNode();
          if (childInfo) {
            child.setInfo(childInfo);
          }
          if (childData) {
            child.setData(childData);
          }
          parent.addToLstChildren(parent.children.length, child);
        }
      }
    };

    SmartArtTree.prototype.remove = function (node, parent) {

    }

    function SmartArtNode() {
      CBaseFormatObject.call(this);
      this.info = null;
      this.data = null;
      this.children = [];
      this.depth = null;
    }
    InitClass(SmartArtNode, CBaseFormatObject, AscDFH.historyitem_type_SmartArtNode);

    SmartArtNode.prototype.setInfo = function (oPr) {
      this.info = oPr;
    }

    SmartArtNode.prototype.getInfo = function () {
      return this.info;
    }

    SmartArtNode.prototype.setData = function (oPr) {
      this.data = oPr;
      oPr.setParent(this);
    }

    SmartArtNode.prototype.getData = function () {
      return this.data;
    }

    SmartArtNode.prototype.addToLstChildren = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.children.length, Math.max(0, nIdx));
      this.children.splice(nInsertIdx, 0, oPr);
      oPr.depth = this.depth + 1;
      oPr.setParent(this);
    }

    SmartArtNode.prototype.removeFromLstChildren = function (nIdx) {
      if (nIdx > -1 && nIdx < this.children.length) {
        this.children[nIdx].setParent(null);
        this.children.splice(nIdx, 1);
      }
    }

    SmartArtNode.prototype.getAxis = function (axisType) {
      switch (axisType) {
        case AxisType_value_ancst:
          return this.getAncst();
        case AxisType_value_ancstOrSelf:
          return this.getAncstOrSelf();
        case AxisType_value_ch:
          return this.getCh();
        case AxisType_value_des:
          return this.getDes();
        case AxisType_value_desOrSelf:
          return this.getDesOrSelf();
        case AxisType_value_follow:
          return this.getFollow();
        case AxisType_value_followSib:
          return this.getFollowSib();
        case AxisType_value_none:
          return this.getNone();
        case AxisType_value_par:
          return this.getPar();
        case AxisType_value_preced:
          return this.getPreced();
        case AxisType_value_precedSib:
          return this.getPrecedSib();
        case AxisType_value_root:
          return this.getRoot();
        case AxisType_value_self:
          return this.getSelf();
        default:
          return;

      }
    }

    SmartArtNode.prototype.getAncst = function () {
      var ancestors = [];
      var root = this;
      while (!(root instanceof SmartArtTree)) {
        if (root !== this) {
          ancestors.unshift(root);
        }
        root = root.parent;
      }
      return ancestors;
    }

    SmartArtNode.prototype.getAncstOrSelf = function () {
      return this.getAncst().concat(this.getSelf());
    }

    SmartArtNode.prototype.getCh = function () {
      return this.children;
    }

    SmartArtNode.prototype.getDes = function () {
      var descendant = [];
      (function recurse(context) {
        context.children.forEach(function (children) {
          descendant.push(children);
          recurse(children);
        });
      })(this);
      return descendant;
    }

    SmartArtNode.prototype.getDesOrSelf = function () {
      return this.getSelf().concat(this.getDes());
    }

    SmartArtNode.prototype.getFollow = function () {
      if (this.parent instanceof SmartArtTree) {
        return;
      }
      var follow = [];
      var followSib = this.getFollowSib();
      followSib.forEach(function (follower) {
        follow = follow.concat(follower.getDesOrSelf());
      });
      return follow;
    }

    SmartArtNode.prototype.getFollowSib = function () {
      if (this.parent instanceof SmartArtTree) {
        return;
      }
      var followSib = [];
      var isFollowSib;
      var parent = this.parent;
      var childs = parent.children;
      for (var i = 0; i < childs.length; i += 1) {
        var child = childs[i];
        if (isFollowSib) {
          followSib.push(child);
        }
        if (this === child) {
          isFollowSib = true;
        }
      }
      return followSib;
    }

    SmartArtNode.prototype.getNone = function () {
      return [];
    }

    SmartArtNode.prototype.getPar = function () {
      return [this.parent];
    }

    SmartArtNode.prototype.getPreced = function () {
      if (this.parent instanceof SmartArtTree) {
        return;
      }
      var preced = [];
      var precedSib = this.getPrecedSib();
      precedSib.forEach(function (preceder) {
        preced = preced.concat(preceder.getDesOrSelf());
      });
      return preced;
    }

    SmartArtNode.prototype.getPrecedSib = function () {
      if (this.parent instanceof SmartArtTree) {
        return;
      }
      var precedSib = [];
      var isPrecedSib;
      var parent = this.parent;
      var childs = parent.children;
      for (var i = childs.length - 1; i >= 0; i -= 1) {
        var child = childs[i];
        if (isPrecedSib) {
          precedSib.push(child);
        }
        if (this === child) {
          isPrecedSib = true;
        }
      }
      return precedSib;
    }

    SmartArtNode.prototype.getRoot = function () {
      var root = this;
      while (!(root instanceof SmartArtTree)) {
        root = root.parent;
      }
      return [root.root];
    }

    SmartArtNode.prototype.getSelf = function () {
      return [this];
    }

    SmartArtNode.prototype.getPoints = function (ptType, name) {
      var points = [];
      if (this.data) {
        switch (ptType) {
          case ElementType_value_all:
            points = this.data.getAllPoints();
            break;
          case ElementType_value_asst:
            points = this.data.getAsstPoint();
            break;
          case ElementType_value_node:
            points = this.data.getNodePoint();
            break;
          case ElementType_value_doc:
            points = this.data.getDocPoint();
            break;
          case ElementType_value_norm:
            points = this.data.getNormPoint();
            break;
          case ElementType_value_nonAsst:
            points = this.data.getNonAsstPoint();
            break;
          case ElementType_value_nonNorm:
            points = this.data.getNonNormPoint();
            break;
          case ElementType_value_parTrans:
            points = this.data.getParPoint();
            break;
          case ElementType_value_pres:
            points = this.data.getPresPoint();
            break;
          case ElementType_value_sibTrans:
            points = this.data.getSibPoint();
            break;
          default:
            points = this.data.getAllPoints();
            break;
        }
        if (name) {
          return points.filter(function (point) {
            return point && point.prSet && point.prSet.presName && point.prSet.presName === name;
          });
        }
      }
      return points;
    }

    SmartArtNode.prototype.getShape = function (name, ptType) {
      var points;
      if (ptType) {
        points = this.getPoints(ptType, name);
      } else {
        points = this.getPoints(ElementType_value_all, name);
      }
      points = points.map(function (point) {
        return point.getShape(name);
      }).filter(function (shape) {
        return !!shape;
      });
      return points;
    }

    function SmartArtNodeData() {
      CBaseFormatObject.call(this);
      this.cxn = null;
      this.sibPoint = [];
      this.parPoint = [];
      this.docPoint = null;
      this.nodePoint = null;
      this.normPoint = [];
      this.asstPoint = null;
      this.presPoint = [];
      this.shapes = [];
    }
    InitClass(SmartArtNodeData, CBaseFormatObject, AscDFH.historyitem_type_SmartArtNodeData);

    SmartArtNodeData.prototype.setSibPoint = function (oPr) {
      this.sibPoint = oPr;
    }

    SmartArtNodeData.prototype.addToLstSibPoint = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.sibPoint.length, Math.max(0, nIdx));
      this.sibPoint.splice(nInsertIdx, 0, oPr);
    }

    SmartArtNodeData.prototype.addToLstNormPoint = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.normPoint.length, Math.max(0, nIdx));
      this.normPoint.splice(nInsertIdx, 0, oPr);
    }

    SmartArtNodeData.prototype.addToLstParPoint = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.parPoint.length, Math.max(0, nIdx));
      this.parPoint.splice(nInsertIdx, 0, oPr);
    }

    SmartArtNodeData.prototype.setDocPoint = function (oPr) {
      this.docPoint = oPr;
    }

    SmartArtNodeData.prototype.addToLstAsstPoint = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.asstPoint.length, Math.max(0, nIdx));
      this.asstPoint.splice(nInsertIdx, 0, oPr);
    }

    SmartArtNodeData.prototype.setAsstPoint = function (oPr) {
      this.asstPoint = oPr;
    }

    SmartArtNodeData.prototype.setNodePoint = function (oPr) {
      this.nodePoint = oPr;
    }

    SmartArtNodeData.prototype.getPresPoint = function () {
      return this.presPoint;
    }

    SmartArtNodeData.prototype.getAllPoints = function () {
      var typeOfPoints = ['sibPoint', 'parPoint', 'docPoint', 'normPoint', 'nodePoint', 'asstPoint', 'presPoint'];
      var result = [];
      var that = this;
      typeOfPoints.forEach(function (pointType) {
        result = result.concat(!!that[pointType] ? that[pointType] : []);
      });
      return result;
    }

    SmartArtNodeData.prototype.getDocPoint = function () {
      return this.docPoint ? [this.docPoint] : [];
    }

    SmartArtNodeData.prototype.getNodePoint = function () {
      return this.nodePoint ? [this.nodePoint] : [];
    }

    SmartArtNodeData.prototype.getAsstPoint = function () {
      return this.asstPoint;
    }

    SmartArtNodeData.prototype.getNormPoint = function () {
      return this.normPoint;
    }

    SmartArtNodeData.prototype.getParPoint = function () {
      return this.parPoint;
    }

    SmartArtNodeData.prototype.getNonAsstPoint = function () {
      return this.getAllPoints().filter(function (point) {
        return point.type !== Point_type_asst;
      });
    }

    SmartArtNodeData.prototype.getNonNormPoint = function () {
      var typeOfPoints = ['sibPoint', 'parPoint', 'docPoint', 'nodePoint', 'asstPoint'];
      var result = [];
      var that = this;
      typeOfPoints.forEach(function (pointType) {
        result = result.concat(that[pointType]);
      });
      return result;
    }

    SmartArtNodeData.prototype.getSibPoint = function () {
      return this.sibPoint;
    }

    SmartArtNodeData.prototype.setCxn = function (oPr) {
      this.cxn = oPr;
    }
    SmartArtNodeData.prototype.getCxn = function () {
      return this.cxn;
    }

    SmartArtNodeData.prototype.setParPoint = function (oPr) {
      this.parPoint = oPr;
    }

    SmartArtNodeData.prototype.getParPoint = function () {
      return this.parPoint;
    }


    SmartArtNodeData.prototype.addToLstShapes = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.shapes.length, Math.max(0, nIdx));
      this.shapes.splice(nInsertIdx, 0, oPr);
    }

    SmartArtNodeData.prototype.removeFromLstShapes = function (nIdx) {
      if (nIdx > -1 && nIdx < this.shapes.length) {
        this.shapes[nIdx].setParent(null);
        this.shapes.splice(nIdx, 1);
      }
    }

    SmartArtNodeData.prototype.addToLstPresPoint = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.presPoint.length, Math.max(0, nIdx));
      this.presPoint.splice(nInsertIdx, 0, oPr);
    }

    SmartArtNodeData.prototype.removeFromLstPresPoint = function (nIdx) {
      if (nIdx > -1 && nIdx < this.presPoint.length) {
        this.presPoint[nIdx].setParent(null);
        this.presPoint.splice(nIdx, 1);
      }
    }

    SmartArtNodeData.prototype.getPresByNameAndStyleLbl = function (name, styleLbl) {
      var presValue;
      this.presPoint.forEach(function (pres) {
        if (pres.prSet) {
          if (name && styleLbl) {
            if (pres.prSet.presName === name && pres.prSet.styleLbl === styleLbl) {
              presValue = pres;
            }
          } else if (name) {
            if (pres.prSet.presName === name) {
              presValue = pres;
            }
          } else if (styleLbl) {
            if (pres.prSet.styleLbl === styleLbl) {
              presValue = pres;
            }
          }
        }
      });
      return presValue;
    }

    SmartArtNodeData.prototype.getPresWithVarLst = function () {
      for (var i = 0; i < this.presPoint.length; i += 1) {
        var pres = this.presPoint[i];
        if (pres && pres.prSet && pres.prSet.presLayoutVars) {
          return pres;
        }
      }
    }

    // Акцентируемый рисунок
    function SmartArtAccentedPicture() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtAccentedPicture, SmartArt);

    // Баланс
    function SmartArtBalance1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBalance1, SmartArt);

    // Блоки рисунков с названиями
    function SmartArtTitledPictureBlocks() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtTitledPictureBlocks, SmartArt);

    // Блоки со смещенными рисунками
    function SmartArtPictureAccentBlocks() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPictureAccentBlocks, SmartArt);
    // Блочный цикл
    function SmartArtCycle5() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCycle5, SmartArt);

    // Венна в столбик
    function SmartArtVenn2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVenn2, SmartArt);

    // Вертикальное уравнение
    function SmartArtEquation2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtEquation2, SmartArt);

    // Вертикальный блочный список
    function SmartArtVList5() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVList5, SmartArt);

    // Вертикальный ломаный процесс
    function SmartArtBProcess4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBProcess4, SmartArt);

    // Вертикальный маркированный список
    function SmartArtVList2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVList2, SmartArt);

    // Вертикальный нелинейный список
    function SmartArtVerticalCurvedList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVerticalCurvedList, SmartArt);

    // Вертикальный процесс
    function SmartArtProcess2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtProcess2, SmartArt);

    // Вертикальный список
    function SmartArtList1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtList1, SmartArt);

    // Вертикальный список рисунков
    function SmartArtVList4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVList4, SmartArt);

    // Вертикальный список с кругами
    function SmartArtVerticalCircleList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVerticalCircleList, SmartArt);

    // Вертикальный список со смещенными рисунками
    function SmartArtVList3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVList3, SmartArt);

    // Вертикальный список со стрелкой
    function SmartArtVList6() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVList6, SmartArt);

    // Вертикальный уголковый список
    function SmartArtChevron2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtChevron2, SmartArt);

    // Вертикальный уголковый список2
    function SmartArtVerticalAccentList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVerticalAccentList, SmartArt);

    // Вложенная целевая
    function SmartArtTarget2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtTarget2, SmartArt);

    // Воронка
    function SmartArtFunnel1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtFunnel1, SmartArt);

    // Восходящая стрелка
    function SmartArtArrow2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtArrow2, SmartArt);

    // Восходящая стрелка процесса
    function SmartArtIncreasingArrowsProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtIncreasingArrowsProcess, SmartArt);

    // Восходящий процесс
    function SmartArtStepUpProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtStepUpProcess, SmartArt);

    // Выноска с круглыми рисунками
    function SmartArtCircularPictureCallout() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCircularPictureCallout, SmartArt);

    // Горизонтальная иерархия
    function SmartArtHierarchy2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHierarchy2, SmartArt);

    // Горизонтальная иерархия с подписями
    function SmartArtHierarchy5() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHierarchy5, SmartArt);

    // Горизонтальная многоуровневая иерархия
    function SmartArtHorizontalMultiLevelHierarchy() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHorizontalMultiLevelHierarchy, SmartArt);

    // Горизонтальная организационная диаграмма
    function SmartArtHorizontalOrganizationChart() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHorizontalOrganizationChart, SmartArt);

    // Горизонтальный маркированный список
    function SmartArtHList1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHList1, SmartArt);

    // Горизонтальный список рисунков
    function SmartArtPList2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPList2, SmartArt);

    // Закрытый уголковый процесс
    function SmartArtHChevron3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHChevron3, SmartArt);

    // Иерархический список
    function SmartArtHierarchy3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHierarchy3, SmartArt);

    // Иерархия
    function SmartArtHierarchy1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHierarchy1, SmartArt);

    // Иерархия с круглыми рисунками
    function SmartArtCirclePictureHierarchy() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCirclePictureHierarchy, SmartArt);

    // Иерархия с подписями
    function SmartArtHierarchy6() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHierarchy6, SmartArt);

    // Инвертированная пирамида
    function SmartArtPyramid3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPyramid3, SmartArt);

    // Кластер шестиугольников
    function SmartArtHexagonCluster() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHexagonCluster, SmartArt);

    // Круг связей
    function SmartArtCircleRelationship() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCircleRelationship, SmartArt);

    // Круглая временная шкала
    function SmartArtCircleAccentTimeline() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCircleAccentTimeline, SmartArt);

    // Круглый ломаный процесс
    function SmartArtBProcess2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBProcess2, SmartArt);

    // Лента со стрелками
    function SmartArtArrow6() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtArrow6, SmartArt);

    // Линейная Венна
    function SmartArtVenn3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVenn3, SmartArt);

    // Линия рисунков
    function SmartArtPictureLineup() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPictureLineup, SmartArt);

    // Линия рисунков с названиями
    function SmartArtTitlePictureLineup() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtTitlePictureLineup, SmartArt);

    // Ломаный список рисунков с подписями
    function SmartArtBendingPictureCaptionList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBendingPictureCaptionList, SmartArt);

    // Ломаный список со смещенными рисунками
    function SmartArtBList2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBList2, SmartArt);

    // Матрица с заголовками
    function SmartArtMatrix1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtMatrix1, SmartArt);

    // Нарастающий процесс с кругами
    function SmartArtIncreasingCircleProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtIncreasingCircleProcess, SmartArt);

    // Нелинейные рисунки с блоками
    function SmartArtBendingPictureBlocks() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBendingPictureBlocks, SmartArt);

    // Нелинейные рисунки с подписями
    function SmartArtBendingPictureCaption() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBendingPictureCaption, SmartArt);

    // Нелинейные рисунки с полупрозрачным текстом
    function SmartArtBendingPictureSemiTransparentText() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBendingPictureSemiTransparentText, SmartArt);

    // Ненаправленный цикл
    function SmartArtCycle6() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCycle6, SmartArt);

    // Непрерывный блочный процесс
    function SmartArtHProcess9() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHProcess9, SmartArt);

    // Непрерывный список с рисунками
    function SmartArtHList7() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHList7, SmartArt);

    // Непрерывный цикл
    function SmartArtCycle3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCycle3, SmartArt);

    // Нисходящий блочный список
    function SmartArtBlockDescendingList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBlockDescendingList, SmartArt);

    // Нисходящий процесс
    function SmartArtStepDownProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtStepDownProcess, SmartArt);

    // Обратный список
    function SmartArtReverseList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtReverseList, SmartArt);

    // Организационная диаграмма
    function SmartArtOrgChart1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtOrgChart1, SmartArt);

    // Организационная диаграмма с именами и должностями
    function SmartArtNameandTitleOrganizationalChart() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtNameandTitleOrganizationalChart, SmartArt);

    // Переменный поток
    function SmartArtHProcess4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHProcess4, SmartArt);

    // Пирамидальный список
    function SmartArtPyramid2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPyramid2, SmartArt);

    // Плюс и минус
    function SmartArtPlusandMinus() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPlusandMinus, SmartArt);

    // Повторяющийся ломаный процесс
    function SmartArtBProcess3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBProcess3, SmartArt);

    // Подписанные рисунки
    function SmartArtCaptionedPictures() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCaptionedPictures, SmartArt);

    // Подробный процесс
    function SmartArtHProcess7() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHProcess7, SmartArt);

    // Полосы рисунков
    function SmartArtPictureStrips() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPictureStrips, SmartArt);

    // Полукруглая организационная диаграмма
    function SmartArtHalfCircleOrganizationChart() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHalfCircleOrganizationChart, SmartArt);

    // Поэтапный процесс
    function SmartArtPhasedProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPhasedProcess, SmartArt);

    // Простая Венна
    function SmartArtVenn1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVenn1, SmartArt);

    // Простая временная шкала
    function SmartArtHProcess11() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHProcess11, SmartArt);

    // Простая круговая
    function SmartArtChart3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtChart3, SmartArt);

    // Простая матрица
    function SmartArtMatrix3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtMatrix3, SmartArt);

    // Простая пирамида
    function SmartArtPyramid1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPyramid1, SmartArt);

    // Простая радиальная
    function SmartArtRadial1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtRadial1, SmartArt);

    // Простая целевая
    function SmartArtTarget1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtTarget1, SmartArt);

    // Простой блочный список
    function SmartArtDefault() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtDefault, SmartArt);

    // Простой ломаный процесс
    function SmartArtProcess5() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtProcess5, SmartArt);

    // Простой процесс
    function SmartArtProcess1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtProcess1, SmartArt);

    // Простой уголковый процесс
    function SmartArtChevron1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtChevron1, SmartArt);

    // Простой цикл
    function SmartArtCycle2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCycle2, SmartArt);

    // Противоположные идеи
    function SmartArtOpposingIdeas() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtOpposingIdeas, SmartArt);

    // Противостоящие стрелки
    function SmartArtArrow4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtArrow4, SmartArt);

    // Процесс от случайности к результату
    function SmartArtRandomtoResultProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtRandomtoResultProcess, SmartArt);

    // Процесс с вложенными шагами
    function SmartArtSubStepProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtSubStepProcess, SmartArt);

    // Процесс с круговой диаграммой
    function SmartArtPieProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPieProcess, SmartArt);

    // Процесс со смещением
    function SmartArtProcess3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtProcess3, SmartArt);

    // Процесс со смещенными по возрастанию рисунками
    function SmartArtAscendingPictureAccentProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtAscendingPictureAccentProcess, SmartArt);

    // Процесс со смещенными рисунками
    function SmartArtHProcess10() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHProcess10, SmartArt);

    // Радиальная Венна
    function SmartArtRadial3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtRadial3, SmartArt);

    // Радиальная циклическая
    function SmartArtRadial6() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtRadial6, SmartArt);

    // Радиальный кластер
    function SmartArtRadialCluster() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtRadialCluster, SmartArt);

    // Радиальный список
    function SmartArtRadial2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtRadial2, SmartArt);

    // Разнонаправленный цикл
    function SmartArtCycle7() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCycle7, SmartArt);

    // Расходящаяся радиальная
    function SmartArtRadial5() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtRadial5, SmartArt);

    // Расходящиеся стрелки
    function SmartArtArrow1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtArrow1, SmartArt);

    // Рисунок с текстом в рамке
    function SmartArtFramedTextPicture() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtFramedTextPicture, SmartArt);

    // Сгруппированный список
    function SmartArtLProcess2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtLProcess2, SmartArt);

    // Сегментированная пирамида
    function SmartArtPyramid4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPyramid4, SmartArt);

    // Сегментированный процесс
    function SmartArtProcess4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtProcess4, SmartArt);

    // Сегментированный цикл
    function SmartArtCycle8() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCycle8, SmartArt);

    // Сетка рисунков
    function SmartArtPictureGrid() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPictureGrid, SmartArt);

    // Сетчатая матрица
    function SmartArtMatrix2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtMatrix2, SmartArt);

    // Спираль рисунков
    function SmartArtSpiralPicture() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtSpiralPicture, SmartArt);

    // Список в столбик
    function SmartArtHList9() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHList9, SmartArt);

    // Список названий рисунков
    function SmartArtPList1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPList1, SmartArt);

    // Список процессов
    function SmartArtLProcess1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtLProcess1, SmartArt);

    // Список рисунков с выносками
    function SmartArtBubblePictureList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtBubblePictureList, SmartArt);

    // Список с квадратиками
    function SmartArtSquareAccentList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtSquareAccentList, SmartArt);

    // Список с линиями
    function SmartArtLinedList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtLinedList, SmartArt);

    // Список со смещенными рисунками
    function SmartArtHList2() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHList2, SmartArt);

    // Список со смещенными рисунками и заголовком
    function SmartArtPictureAccentList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtPictureAccentList, SmartArt);

    // Список со снимками
    function SmartArtSnapshotPictureList() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtSnapshotPictureList, SmartArt);

    // Стрелка непрерывного процесса
    function SmartArtHProcess3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHProcess3, SmartArt);

    // Стрелка процесса с кругами
    function SmartArtCircleArrowProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCircleArrowProcess, SmartArt);

    // Стрелки процесса
    function SmartArtHProcess6() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHProcess6, SmartArt);

    // Ступенчатый процесс
    function SmartArtVProcess5() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtVProcess5, SmartArt);

    // Сходящаяся радиальная
    function SmartArtRadial4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtRadial4, SmartArt);

    // Сходящиеся стрелки
    function SmartArtArrow5() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtArrow5, SmartArt);

    // Табличная иерархия
    function SmartArtHierarchy4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHierarchy4, SmartArt);

    // Табличный список
    function SmartArtHList3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHList3, SmartArt);

    // Текстовый цикл
    function SmartArtCycle1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCycle1, SmartArt);

    // Трапецевидный список
    function SmartArtHList6() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtHList6, SmartArt);

    // Убывающий процесс
    function SmartArtDescendingProcess() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtDescendingProcess, SmartArt);

    // Уголковый список
    function SmartArtLProcess3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtLProcess3, SmartArt);

    // Уравнение
    function SmartArtEquation1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtEquation1, SmartArt);

    // Уравновешивающие стрелки
    function SmartArtArrow3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtArrow3, SmartArt);

    // Целевой список
    function SmartArtTarget3() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtTarget3, SmartArt);

    // Циклическая матрица
    function SmartArtCycle4() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtCycle4, SmartArt);

    // Чередующиеся блоки рисунков
    function SmartArtAlternatingPictureBlocks() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtAlternatingPictureBlocks, SmartArt);

    // Чередующиеся круги рисунков
    function SmartArtAlternatingPictureCircles() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtAlternatingPictureCircles, SmartArt);

    // Чередующиеся шестиугольники
    function SmartArtAlternatingHexagons() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtAlternatingHexagons, SmartArt);

    // Шестеренки
    function SmartArtGear1() {
      SmartArt.call(this);
    }
    InitClassWithoutType(SmartArtGear1, SmartArt);



    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].kForInsFitFontSize     = kForInsFitFontSize;
    window['AscFormat'].PrSet                  = PrSet;
    window['AscFormat'].CCommonDataList        = CCommonDataList;
    window['AscFormat'].Point                  = Point;
    window['AscFormat'].PtLst                  = PtLst;
    window['AscFormat'].DataModel              = DataModel;
    window['AscFormat'].CxnLst                 = CxnLst;
    window['AscFormat'].ExtLst                 = ExtLst;
    window['AscFormat'].BgFormat               = BgFormat;
    window['AscFormat'].Whole                  = Whole;
    window['AscFormat'].Cxn                    = Cxn;
    window['AscFormat'].Ext                    = Ext;
    window['AscFormat'].LayoutDef              = LayoutDef;
    window['AscFormat'].CatLst                 = CatLst;
    window['AscFormat'].SCat                   = SCat;
    window['AscFormat'].ClrData                = ClrData;
    window['AscFormat'].Desc                   = Desc;
    window['AscFormat'].LayoutNode             = LayoutNode;
    window['AscFormat'].Alg                    = Alg;
    window['AscFormat'].Param                  = Param;
    window['AscFormat'].Choose                 = Choose;
    window['AscFormat'].IteratorAttributes     = IteratorAttributes;
    window['AscFormat'].Else                   = Else;
    window['AscFormat'].AxisType               = AxisType;
    window['AscFormat'].If                     = If;
    window['AscFormat'].ElementType            = ElementType;
    window['AscFormat'].ConstrLst              = ConstrLst;
    window['AscFormat'].Constr                 = Constr;
    window['AscFormat'].PresOf                 = PresOf;
    window['AscFormat'].RuleLst                = RuleLst;
    window['AscFormat'].Rule                   = Rule;
    window['AscFormat'].SShape                 = SShape;
    window['AscFormat'].AdjLst                 = AdjLst;
    window['AscFormat'].Adj                    = Adj;
    window['AscFormat'].AnimLvl                = AnimLvl;
    window['AscFormat'].AnimOne                = AnimOne;
    window['AscFormat'].BulletEnabled          = BulletEnabled;
    window['AscFormat'].ChMax                  = ChMax;
    window['AscFormat'].ChPref                 = ChPref;
    window['AscFormat'].DiagramDirection       = DiagramDirection;
    window['AscFormat'].DiagramTitle           = DiagramTitle;
    window['AscFormat'].LayoutDefHdrLst        = LayoutDefHdrLst;
    window['AscFormat'].LayoutDefHdr           = LayoutDefHdr;
    window['AscFormat'].RelIds                 = RelIds;
    window['AscFormat'].VarLst                 = VarLst;
    window['AscFormat'].ColorsDef              = ColorsDef;
    window['AscFormat'].ColorDefStyleLbl       = ColorDefStyleLbl;
    window['AscFormat'].ClrLst                 = ClrLst;
    window['AscFormat'].EffectClrLst           = EffectClrLst;
    window['AscFormat'].FillClrLst             = FillClrLst;
    window['AscFormat'].LinClrLst              = LinClrLst;
    window['AscFormat'].TxEffectClrLst         = TxEffectClrLst;
    window['AscFormat'].TxFillClrLst           = TxFillClrLst;
    window['AscFormat'].TxLinClrLst            = TxLinClrLst;
    window['AscFormat'].ColorsDefHdr           = ColorsDefHdr;
    window['AscFormat'].ColorsDefHdrLst        = ColorsDefHdrLst;
    window['AscFormat'].StyleDef               = StyleDef;
    window['AscFormat'].Scene3d                = Scene3d;
    window['AscFormat'].StyleDefStyleLbl       = StyleDefStyleLbl;
    window['AscFormat'].Scene3d                = Scene3d;
    window['AscFormat'].Backdrop               = Backdrop;
    window['AscFormat'].BackdropNorm           = BackdropNorm;
    window['AscFormat'].BackdropUp             = BackdropUp;
    window['AscFormat'].Camera                 = Camera;
    window['AscFormat'].Rot                    = Rot;
    window['AscFormat'].LightRig               = LightRig;
    window['AscFormat'].Sp3d                   = Sp3d;
    window['AscFormat'].Bevel                  = Bevel;
    window['AscFormat'].BevelB                 = BevelB;
    window['AscFormat'].BevelT                 = BevelT;
    window['AscFormat'].TxPr                   = TxPr;
    window['AscFormat'].FlatTx                 = FlatTx;
    window['AscFormat'].StyleDefHdrLst         = StyleDefHdrLst;
    window['AscFormat'].StyleDefHdr            = StyleDefHdr;
    window['AscFormat'].BackdropAnchor         = BackdropAnchor;
    window['AscFormat'].StyleData              = StyleData;
    window['AscFormat'].SampData               = SampData;
    window['AscFormat'].ForEach                = ForEach;
    window['AscFormat'].ResizeHandles          = ResizeHandles;
    window['AscFormat'].OrgChart               = OrgChart;
    window['AscFormat'].HierBranch             = HierBranch;
    window['AscFormat'].ParameterVal           = ParameterVal;
    window['AscFormat'].Coordinate             = Coordinate;
    window['AscFormat'].ExtrusionClr           = ExtrusionClr;
    window['AscFormat'].ContourClr             = ContourClr;
    window['AscFormat'].SmartArt               = SmartArt;
    window['AscFormat'].CCommonDataClrList     = CCommonDataClrList;
    window['AscFormat'].BuNone                 = BuNone;
    window['AscFormat'].Drawing                = Drawing;
    window['AscFormat'].DiagramData            = DiagramData;
    window['AscFormat'].FunctionValue          = FunctionValue;
    window['AscFormat'].ShapeSmartArtInfo      = ShapeSmartArtInfo;
    window['AscFormat'].SmartArtTree           = SmartArtTree;
    window['AscFormat'].SmartArtNode           = SmartArtNode;
    window['AscFormat'].SmartArtNodeData       = SmartArtNodeData;

  })(window)
