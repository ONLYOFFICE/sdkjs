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
    var CBaseFormatObject = AscFormat.CBaseFormatObject;
    var CBaseFormatNoIdObject = AscFormat.CBaseFormatNoIdObject;
    var oHistory = AscCommon.History;
    var CChangeBool = AscDFH.CChangesDrawingsBool;
    var CChangeLong = AscDFH.CChangesDrawingsLong;
    var CChangeString = AscDFH.CChangesDrawingsString;
    var CChangeObjectNoId = AscDFH.CChangesDrawingsObjectNoId;
    var CChangeObject = AscDFH.CChangesDrawingsObject;
    var CChangeContent = AscDFH.CChangesDrawingsContent;
    var CChangeDouble2 = AscDFH.CChangesDrawingsDouble2;

    var drawingsChangesMap = AscDFH.drawingsChangesMap;
    var drawingContentChanges = AscDFH.drawingContentChanges;
    var changesFactory = AscDFH.changesFactory;
    var drawingConstructorsMap = window['AscDFH'].drawingsConstructorsMap;
    var CUniColor = AscFormat.CUniColor;
    var CGraphicObjectBase = AscFormat.CGraphicObjectBase;
    var CGroupShape = AscFormat.CGroupShape;

    // consts
    const GRAYSCALE_TRESHHOLD = 150;

	  const smartArtContentFillingType_parentWithChildren = 0x0011;
	  const smartArtContentFillingType_onlyChildren = 0x0001;
	  const smartArtContentFillingType_onlyParent = 0x0010;

	  const Point_type_asst = 1;
    const Point_type_doc = 2;
    const Point_type_node = 0;
    const Point_type_parTrans = 4;
    const Point_type_pres = 3;
    const Point_type_sibTrans = 5;

    const Cxn_type_parOf = 0;
    const Cxn_type_presOf = 1;
    const Cxn_type_presParOf = 2;
    const Cxn_type_unknownRelationShip = 3;

    const LayoutNode_type_b = 0;
    const LayoutNode_type_t = 1;

    const Alg_type_composite = 0;
    const Alg_type_conn = 1;
    const Alg_type_cycle = 2;
    const Alg_type_hierChild = 3;
    const Alg_type_hierRoot = 4;
    const Alg_type_lin = 6;
    const Alg_type_pyra = 5;
    const Alg_type_snake = 9;
    const Alg_type_sp = 7;
    const Alg_type_tx = 8;


    const Param_type_horzAlign = 0;
    const Param_type_vertAlign = 1;
    const Param_type_chDir = 2;
    const Param_type_chAlign = 3;
    const Param_type_secChAlign = 4;
    const Param_type_linDir = 5;
    const Param_type_secLinDir = 6;
    const Param_type_stElem = 7;
    const Param_type_bendPt = 8;
    const Param_type_connRout = 9;
    const Param_type_begSty = 10;
    const Param_type_endSty = 11;
    const Param_type_dim = 12;
    const Param_type_rotPath = 13;
    const Param_type_ctrShpMap = 14;
    const Param_type_nodeHorzAlign = 15;
    const Param_type_nodeVertAlign = 16;
    const Param_type_fallback = 17;
    const Param_type_txDir = 18;
    const Param_type_pyraAcctPos = 19;
    const Param_type_pyraAcctTxMar = 20;
    const Param_type_txBlDir = 21;
    const Param_type_txAnchorHorz = 22;
    const Param_type_txAnchorVert = 23;
    const Param_type_txAnchorHorzCh = 24;
    const Param_type_txAnchorVertCh = 25;
    const Param_type_parTxLTRAlign = 26;
    const Param_type_parTxRTLAlign = 27;
    const Param_type_shpTxLTRAlignCh = 28;
    const Param_type_shpTxRTLAlignCh = 29;
    const Param_type_autoTxRot = 30;
    const Param_type_grDir = 31;
    const Param_type_flowDir = 32;
    const Param_type_contDir = 33;
    const Param_type_bkpt = 34;
    const Param_type_off = 35;
    const Param_type_hierAlign = 36;
    const Param_type_bkPtFixedVal = 37;
    const Param_type_stBulletLvl = 38;
    const Param_type_stAng = 39;
    const Param_type_spanAng = 40;
    const Param_type_ar = 41;
    const Param_type_lnSpPar = 42;
    const Param_type_lnSpAfParP = 43;
    const Param_type_lnSpCh = 44;
    const Param_type_lnSpAfChP = 45;
    const Param_type_rtShortDist = 46;
    const Param_type_alignTx = 47;
    const Param_type_pyraLvlNode = 48;
    const Param_type_pyraAcctBkgdNode = 49;
    const Param_type_pyraAcctTxNode = 50;
    const Param_type_srcNode = 51;
    const Param_type_dstNode = 52;
    const Param_type_begPts = 53;
    const Param_type_endPts = 54;

    const AxisType_value_ancst = 6;
    const AxisType_value_ancstOrSelf = 7;
    const AxisType_value_ch = 2;
    const AxisType_value_des = 3;
    const AxisType_value_desOrSelf = 4;
    const AxisType_value_follow = 10;
    const AxisType_value_followSib = 8;
    const AxisType_value_none = 0;
    const AxisType_value_par = 5;
    const AxisType_value_preced = 11;
    const AxisType_value_precedSib = 9;
    const AxisType_value_root = 12;
    const AxisType_value_self = 1;

    const ElementType_value_all = 0;
    const ElementType_value_asst = 5;
    const ElementType_value_doc = 1;
    const ElementType_value_node = 2;
    const ElementType_value_nonAsst = 6;
    const ElementType_value_nonNorm = 4;
    const ElementType_value_norm = 3;
    const ElementType_value_parTrans = 7;
    const ElementType_value_pres = 8;
    const ElementType_value_sibTrans = 9;

    const If_op_equ = 0;
    const If_op_neq = 1;
    const If_op_gt = 2;
    const If_op_lt = 3;
    const If_op_gte = 4;
    const If_op_lte = 5;

    const If_func_cnt = 0;
    const If_func_depth = 6;
    const If_func_maxDepth = 7;
    const If_func_pos = 1;
    const If_func_posEven = 3;
    const If_func_posOdd = 4;
    const If_func_revPos = 2;
    const If_func_var = 5;

    const If_arg_animLvl = 0;
    const If_arg_animOne = 1;
    const If_arg_bulEnabled = 2;
    const If_arg_chMax = 3;
    const If_arg_chPref = 4;
    const If_arg_dir = 5;
    const If_arg_hierBranch = 6;
    const If_arg_none = 7;
    const If_arg_orgChart = 8;
    const If_arg_resizeHandles = 9;

    const Constr_for_ch = 1;
    const Constr_for_des = 2;
    const Constr_for_self = 0;

    const Constr_op_equ = 1;
    const Constr_op_gte = 2;
    const Constr_op_lte = 3;
    const Constr_op_none = 0;

    const Constr_type_alignOff = 1;
    const Constr_type_b = 5;
    const Constr_type_begMarg = 2;
    const Constr_type_begPad = 4;
    const Constr_type_bendDist = 3;
    const Constr_type_bMarg = 6;
    const Constr_type_bOff = 7;
    const Constr_type_connDist = 12;
    const Constr_type_ctrX = 8;
    const Constr_type_ctrXOff = 9;
    const Constr_type_ctrY = 10;
    const Constr_type_ctrYOff = 11;
    const Constr_type_diam = 13;
    const Constr_type_endMarg = 14;
    const Constr_type_endPad = 15;
    const Constr_type_h = 16;
    const Constr_type_hArH = 17;
    const Constr_type_hOff = 63;
    const Constr_type_l = 18;
    const Constr_type_lMarg = 19;
    const Constr_type_lOff = 20;
    const Constr_type_none = 0;
    const Constr_type_primFontSz = 24;
    const Constr_type_pyraAcctRatio = 25;
    const Constr_type_r = 21;
    const Constr_type_rMarg = 22;
    const Constr_type_rOff = 23;
    const Constr_type_secFontSz = 26;
    const Constr_type_secSibSp = 28;
    const Constr_type_sibSp = 27;
    const Constr_type_sp = 29;
    const Constr_type_stemThick = 30;
    const Constr_type_t = 31;
    const Constr_type_tMarg = 32;
    const Constr_type_tOff = 33;
    const Constr_type_userA = 34;
    const Constr_type_userB = 35;
    const Constr_type_userC = 36;
    const Constr_type_userD = 37;
    const Constr_type_userE = 38;
    const Constr_type_userF = 39;
    const Constr_type_userG = 40;
    const Constr_type_userH = 41;
    const Constr_type_userI = 42;
    const Constr_type_userJ = 43;
    const Constr_type_userK = 44;
    const Constr_type_userL = 45;
    const Constr_type_userM = 46;
    const Constr_type_userN = 47;
    const Constr_type_userO = 48;
    const Constr_type_userP = 49;
    const Constr_type_userQ = 50;
    const Constr_type_userR = 51;
    const Constr_type_userS = 52;
    const Constr_type_userT = 53;
    const Constr_type_userU = 54;
    const Constr_type_userV = 55;
    const Constr_type_userW = 56;
    const Constr_type_userX = 57;
    const Constr_type_userY = 58;
    const Constr_type_userZ = 59;
    const Constr_type_w = 60;
    const Constr_type_wArH = 61;
    const Constr_type_wOff = 62;

	  const EChOrder_chOrderB = 0;
	  const EChOrder_chOrderT = 0;

    const kForInsFitFontSize = 71.12 / 360;

    const LayoutShapeType_outputShapeType_conn = 0;
    const LayoutShapeType_outputShapeType_none = 1;
    const LayoutShapeType_shapeType_accentBorderCallout1 = 2;
    const LayoutShapeType_shapeType_accentBorderCallout2 = 3;
    const LayoutShapeType_shapeType_accentBorderCallout3 = 4;
    const LayoutShapeType_shapeType_accentCallout1 = 5;
    const LayoutShapeType_shapeType_accentCallout2 = 6;
    const LayoutShapeType_shapeType_accentCallout3 = 7;
    const LayoutShapeType_shapeType_actionButtonBackPrevious = 8;
    const LayoutShapeType_shapeType_actionButtonBeginning = 9;
    const LayoutShapeType_shapeType_actionButtonBlank = 10;
    const LayoutShapeType_shapeType_actionButtonDocument = 11;
    const LayoutShapeType_shapeType_actionButtonEnd = 12;
    const LayoutShapeType_shapeType_actionButtonForwardNext = 13;
    const LayoutShapeType_shapeType_actionButtonHelp = 14;
    const LayoutShapeType_shapeType_actionButtonHome = 15;
    const LayoutShapeType_shapeType_actionButtonInformation = 16;
    const LayoutShapeType_shapeType_actionButtonMovie = 17;
    const LayoutShapeType_shapeType_actionButtonReturn = 18;
    const LayoutShapeType_shapeType_actionButtonSound = 19;
    const LayoutShapeType_shapeType_arc = 20;
    const LayoutShapeType_shapeType_bentArrow = 21;
    const LayoutShapeType_shapeType_bentConnector2 = 22;
    const LayoutShapeType_shapeType_bentConnector3 = 23;
    const LayoutShapeType_shapeType_bentConnector4 = 24;
    const LayoutShapeType_shapeType_bentConnector5 = 25;
    const LayoutShapeType_shapeType_bentUpArrow = 26;
    const LayoutShapeType_shapeType_bevel = 27;
    const LayoutShapeType_shapeType_blockArc = 28;
    const LayoutShapeType_shapeType_borderCallout1 = 29;
    const LayoutShapeType_shapeType_borderCallout2 = 30;
    const LayoutShapeType_shapeType_borderCallout3 = 31;
    const LayoutShapeType_shapeType_bracePair = 32;
    const LayoutShapeType_shapeType_bracketPair = 33;
    const LayoutShapeType_shapeType_callout1 = 34;
    const LayoutShapeType_shapeType_callout2 = 35;
    const LayoutShapeType_shapeType_callout3 = 36;
    const LayoutShapeType_shapeType_can = 37;
    const LayoutShapeType_shapeType_chartPlus = 38;
    const LayoutShapeType_shapeType_chartStar = 39;
    const LayoutShapeType_shapeType_chartX = 40;
    const LayoutShapeType_shapeType_chevron = 41;
    const LayoutShapeType_shapeType_chord = 42;
    const LayoutShapeType_shapeType_circularArrow = 43;
    const LayoutShapeType_shapeType_cloud = 44;
    const LayoutShapeType_shapeType_cloudCallout = 45;
    const LayoutShapeType_shapeType_corner = 46;
    const LayoutShapeType_shapeType_cornerTabs = 47;
    const LayoutShapeType_shapeType_cube = 48;
    const LayoutShapeType_shapeType_curvedConnector2 = 49;
    const LayoutShapeType_shapeType_curvedConnector3 = 50;
    const LayoutShapeType_shapeType_curvedConnector4 = 51;
    const LayoutShapeType_shapeType_curvedConnector5 = 52;
    const LayoutShapeType_shapeType_curvedDownArrow = 53;
    const LayoutShapeType_shapeType_curvedLeftArrow = 54;
    const LayoutShapeType_shapeType_curvedRightArrow = 55;
    const LayoutShapeType_shapeType_curvedUpArrow = 56;
    const LayoutShapeType_shapeType_decagon = 57;
    const LayoutShapeType_shapeType_diagStripe = 58;
    const LayoutShapeType_shapeType_diamond = 59;
    const LayoutShapeType_shapeType_dodecagon = 60;
    const LayoutShapeType_shapeType_donut = 61;
    const LayoutShapeType_shapeType_doubleWave = 62;
    const LayoutShapeType_shapeType_downArrow = 63;
    const LayoutShapeType_shapeType_downArrowCallout = 64;
    const LayoutShapeType_shapeType_ellipse = 65;
    const LayoutShapeType_shapeType_ellipseRibbon = 66;
    const LayoutShapeType_shapeType_ellipseRibbon2 = 67;
    const LayoutShapeType_shapeType_flowChartAlternateProcess = 68;
    const LayoutShapeType_shapeType_flowChartCollate = 69;
    const LayoutShapeType_shapeType_flowChartConnector = 70;
    const LayoutShapeType_shapeType_flowChartDecision = 71;
    const LayoutShapeType_shapeType_flowChartDelay = 72;
    const LayoutShapeType_shapeType_flowChartDisplay = 73;
    const LayoutShapeType_shapeType_flowChartDocument = 74;
    const LayoutShapeType_shapeType_flowChartExtract = 75;
    const LayoutShapeType_shapeType_flowChartInputOutput = 76;
    const LayoutShapeType_shapeType_flowChartInternalStorage = 77;
    const LayoutShapeType_shapeType_flowChartMagneticDisk = 78;
    const LayoutShapeType_shapeType_flowChartMagneticDrum = 79;
    const LayoutShapeType_shapeType_flowChartMagneticTape = 80;
    const LayoutShapeType_shapeType_flowChartManualInput = 81;
    const LayoutShapeType_shapeType_flowChartManualOperation = 82;
    const LayoutShapeType_shapeType_flowChartMerge = 83;
    const LayoutShapeType_shapeType_flowChartMultidocument = 84;
    const LayoutShapeType_shapeType_flowChartOfflineStorage = 85;
    const LayoutShapeType_shapeType_flowChartOffpageConnector = 86;
    const LayoutShapeType_shapeType_flowChartOnlineStorage = 87;
    const LayoutShapeType_shapeType_flowChartOr = 88;
    const LayoutShapeType_shapeType_flowChartPredefinedProcess = 89;
    const LayoutShapeType_shapeType_flowChartPreparation = 90;
    const LayoutShapeType_shapeType_flowChartProcess = 91;
    const LayoutShapeType_shapeType_flowChartPunchedCard = 92;
    const LayoutShapeType_shapeType_flowChartPunchedTape = 93;
    const LayoutShapeType_shapeType_flowChartSort = 94;
    const LayoutShapeType_shapeType_flowChartSummingJunction = 95;
    const LayoutShapeType_shapeType_flowChartTerminator = 96;
    const LayoutShapeType_shapeType_foldedCorner = 97;
    const LayoutShapeType_shapeType_frame = 98;
    const LayoutShapeType_shapeType_funnel = 99;
    const LayoutShapeType_shapeType_gear6 = 100;
    const LayoutShapeType_shapeType_gear9 = 101;
    const LayoutShapeType_shapeType_halfFrame = 102;
    const LayoutShapeType_shapeType_heart = 103;
    const LayoutShapeType_shapeType_heptagon = 104;
    const LayoutShapeType_shapeType_hexagon = 105;
    const LayoutShapeType_shapeType_homePlate = 106;
    const LayoutShapeType_shapeType_horizontalScroll = 107;
    const LayoutShapeType_shapeType_irregularSeal1 = 108;
    const LayoutShapeType_shapeType_irregularSeal2 = 109;
    const LayoutShapeType_shapeType_leftArrow = 110;
    const LayoutShapeType_shapeType_leftArrowCallout = 111;
    const LayoutShapeType_shapeType_leftBrace = 112;
    const LayoutShapeType_shapeType_leftBracket = 113;
    const LayoutShapeType_shapeType_leftCircularArrow = 114;
    const LayoutShapeType_shapeType_leftRightArrow = 115;
    const LayoutShapeType_shapeType_leftRightArrowCallout = 116;
    const LayoutShapeType_shapeType_leftRightCircularArrow = 117;
    const LayoutShapeType_shapeType_leftRightRibbon = 118;
    const LayoutShapeType_shapeType_leftRightUpArrow = 119;
    const LayoutShapeType_shapeType_leftUpArrow = 120;
    const LayoutShapeType_shapeType_lightningBolt = 121;
    const LayoutShapeType_shapeType_line = 122;
    const LayoutShapeType_shapeType_lineInv = 123;
    const LayoutShapeType_shapeType_mathDivide = 124;
    const LayoutShapeType_shapeType_mathEqual = 125;
    const LayoutShapeType_shapeType_mathMinus = 126;
    const LayoutShapeType_shapeType_mathMultiply = 127;
    const LayoutShapeType_shapeType_mathNotEqual = 128;
    const LayoutShapeType_shapeType_mathPlus = 129;
    const LayoutShapeType_shapeType_moon = 130;
    const LayoutShapeType_shapeType_nonIsoscelesTrapezoid = 131;
    const LayoutShapeType_shapeType_noSmoking = 132;
    const LayoutShapeType_shapeType_notchedRightArrow = 133;
    const LayoutShapeType_shapeType_octagon = 134;
    const LayoutShapeType_shapeType_parallelogram = 135;
    const LayoutShapeType_shapeType_pentagon = 136;
    const LayoutShapeType_shapeType_pie = 137;
    const LayoutShapeType_shapeType_pieWedge = 138;
    const LayoutShapeType_shapeType_plaque = 139;
    const LayoutShapeType_shapeType_plaqueTabs = 140;
    const LayoutShapeType_shapeType_plus = 141;
    const LayoutShapeType_shapeType_quadArrow = 142;
    const LayoutShapeType_shapeType_quadArrowCallout = 143;
    const LayoutShapeType_shapeType_rect = 144;
    const LayoutShapeType_shapeType_ribbon = 145;
    const LayoutShapeType_shapeType_ribbon2 = 146;
    const LayoutShapeType_shapeType_rightArrow = 147;
    const LayoutShapeType_shapeType_rightArrowCallout = 148;
    const LayoutShapeType_shapeType_rightBrace = 149;
    const LayoutShapeType_shapeType_rightBracket = 150;
    const LayoutShapeType_shapeType_round1Rect = 151;
    const LayoutShapeType_shapeType_round2DiagRect = 152;
    const LayoutShapeType_shapeType_round2SameRect = 153;
    const LayoutShapeType_shapeType_roundRect = 154;
    const LayoutShapeType_shapeType_rtTriangle = 155;
    const LayoutShapeType_shapeType_smileyFace = 156;
    const LayoutShapeType_shapeType_snip1Rect = 157;
    const LayoutShapeType_shapeType_snip2DiagRect = 158;
    const LayoutShapeType_shapeType_snip2SameRect = 159;
    const LayoutShapeType_shapeType_snipRoundRect = 160;
    const LayoutShapeType_shapeType_squareTabs = 161;
    const LayoutShapeType_shapeType_star10 = 162;
    const LayoutShapeType_shapeType_star12 = 163;
    const LayoutShapeType_shapeType_star16 = 164;
    const LayoutShapeType_shapeType_star24 = 165;
    const LayoutShapeType_shapeType_star32 = 166;
    const LayoutShapeType_shapeType_star4 = 167;
    const LayoutShapeType_shapeType_star5 = 168;
    const LayoutShapeType_shapeType_star6 = 169;
    const LayoutShapeType_shapeType_star7 = 170;
    const LayoutShapeType_shapeType_star8 = 171;
    const LayoutShapeType_shapeType_straightConnector1 = 172;
    const LayoutShapeType_shapeType_stripedRightArrow = 173;
    const LayoutShapeType_shapeType_sun = 174;
    const LayoutShapeType_shapeType_swooshArrow = 175;
    const LayoutShapeType_shapeType_teardrop = 176;
    const LayoutShapeType_shapeType_trapezoid = 177;
    const LayoutShapeType_shapeType_triangle = 178;
    const LayoutShapeType_shapeType_upArrow = 179;
    const LayoutShapeType_shapeType_upArrowCallout = 180;
    const LayoutShapeType_shapeType_upDownArrow = 181;
    const LayoutShapeType_shapeType_upDownArrowCallout = 182;
    const LayoutShapeType_shapeType_uturnArrow = 183;
    const LayoutShapeType_shapeType_verticalScroll = 184;
    const LayoutShapeType_shapeType_wave = 185;
    const LayoutShapeType_shapeType_wedgeEllipseCallout = 186;
    const LayoutShapeType_shapeType_wedgeRectCallout = 187;
    const LayoutShapeType_shapeType_wedgeRoundRectCallout = 188;


    const AnimLvl_val_ctr = 1;
    const AnimLvl_val_lvl = 2;
    const AnimLvl_val_none = 0;

    const AnimOne_val_branch = 1;
    const AnimOne_val_none = 0;
    const AnimOne_val_one = 2;

    const DiagramDirection_val_norm = 0;
    const DiagramDirection_val_rev = 1;

    const HierBranch_val_hang = 0;
    const HierBranch_val_init = 1;
    const HierBranch_val_l = 2;
    const HierBranch_val_r = 3;
    const HierBranch_val_std = 4;

    const ResizeHandles_val_exact = 0;
    const ResizeHandles_val_rel = 1;

    const ClrLst_hueDir_ccw = 0;
    const ClrLst_hueDir_cw = 1;
    const ClrLst_meth_cycle = 0;
    const ClrLst_meth_repeat = 1;
    const ClrLst_meth_span = 2;

    const Camera_prst_isometricBottomDown = 0;
    const Camera_prst_isometricBottomUp = 1;
    const Camera_prst_isometricLeftDown = 2;
    const Camera_prst_isometricLeftUp = 3;
    const Camera_prst_isometricOffAxis1Left = 4;
    const Camera_prst_isometricOffAxis1Right = 5;
    const Camera_prst_isometricOffAxis1Top = 6;
    const Camera_prst_isometricOffAxis2Left = 7;
    const Camera_prst_isometricOffAxis2Right = 8;
    const Camera_prst_isometricOffAxis2Top = 9;
    const Camera_prst_isometricOffAxis3Bottom = 10;
    const Camera_prst_isometricOffAxis3Left = 11;
    const Camera_prst_isometricOffAxis3Right = 12;
    const Camera_prst_isometricOffAxis4Bottom = 13;
    const Camera_prst_isometricOffAxis4Left = 14;
    const Camera_prst_isometricOffAxis4Right = 15;
    const Camera_prst_isometricRightDown = 16;
    const Camera_prst_isometricRightUp = 17;
    const Camera_prst_isometricTopDown = 18;
    const Camera_prst_isometricTopUp = 19;
    const Camera_prst_legacyObliqueBottom = 20;
    const Camera_prst_legacyObliqueBottomLeft = 21;
    const Camera_prst_legacyObliqueBottomRight = 22;
    const Camera_prst_legacyObliqueFront = 23;
    const Camera_prst_legacyObliqueLeft = 24;
    const Camera_prst_legacyObliqueRight = 25;
    const Camera_prst_legacyObliqueTop = 26;
    const Camera_prst_legacyObliqueTopLeft = 27;
    const Camera_prst_legacyObliqueTopRight = 28;
    const Camera_prst_legacyPerspectiveBottom = 29;
    const Camera_prst_legacyPerspectiveBottomLeft = 30;
    const Camera_prst_legacyPerspectiveBottomRight = 31;
    const Camera_prst_legacyPerspectiveFront = 32;
    const Camera_prst_legacyPerspectiveLeft = 33;
    const Camera_prst_legacyPerspectiveRight = 34;
    const Camera_prst_legacyPerspectiveTop = 35;
    const Camera_prst_legacyPerspectiveTopLeft = 36;
    const Camera_prst_legacyPerspectiveTopRight = 37;
    const Camera_prst_obliqueBottom = 38;
    const Camera_prst_obliqueBottomLeft = 39;
    const Camera_prst_obliqueBottomRight = 40;
    const Camera_prst_obliqueLeft = 41;
    const Camera_prst_obliqueRight = 42;
    const Camera_prst_obliqueTop = 43;
    const Camera_prst_obliqueTopLeft = 44;
    const Camera_prst_obliqueTopRight = 45;
    const Camera_prst_orthographicFront = 46;
    const Camera_prst_perspectiveAbove = 47;
    const Camera_prst_perspectiveAboveLeftFacing = 48;
    const Camera_prst_perspectiveAboveRightFacing = 49;
    const Camera_prst_perspectiveBelow = 50;
    const Camera_prst_perspectiveContrastingLeftFacing = 51;
    const Camera_prst_perspectiveContrastingRightFacing = 52;
    const Camera_prst_perspectiveFront = 53;
    const Camera_prst_perspectiveHeroicExtremeLeftFacing = 54;
    const Camera_prst_perspectiveHeroicExtremeRightFacing = 55;
    const Camera_prst_perspectiveHeroicLeftFacing = 56;
    const Camera_prst_perspectiveHeroicRightFacing = 57;
    const Camera_prst_perspectiveLeft = 58;
    const Camera_prst_perspectiveRelaxed = 59;
    const Camera_prst_perspectiveRelaxedModerately = 60;
    const Camera_prst_perspectiveRight = 61;

    const Sp3d_prstMaterial_clear = 0;
    const Sp3d_prstMaterial_dkEdge = 1;
    const Sp3d_prstMaterial_flat = 2;
    const Sp3d_prstMaterial_legacyMatte = 3;
    const Sp3d_prstMaterial_legacyMetal = 4;
    const Sp3d_prstMaterial_legacyPlastic = 5;
    const Sp3d_prstMaterial_legacyWireframe = 6;
    const Sp3d_prstMaterial_matte = 7;
    const Sp3d_prstMaterial_metal = 8;
    const Sp3d_prstMaterial_plastic = 9;
    const Sp3d_prstMaterial_powder = 10;
    const Sp3d_prstMaterial_softEdge = 11;
    const Sp3d_prstMaterial_softmetal = 12;
    const Sp3d_prstMaterial_translucentPowder = 13;
    const Sp3d_prstMaterial_warmMatte = 14;

    const LightRig_dir_b = 0;
    const LightRig_dir_bl = 1;
    const LightRig_dir_br = 2;
    const LightRig_dir_l = 4;
    const LightRig_dir_r = 5;
    const LightRig_dir_t = 6;
    const LightRig_dir_tl = 7;
    const LightRig_dir_tr = 8;

    const LightRig_rig_balanced = 0;
    const LightRig_rig_brightRoom = 1;
    const LightRig_rig_chilly = 2;
    const LightRig_rig_contrasting = 3;
    const LightRig_rig_flat = 4;
    const LightRig_rig_flood = 5;
    const LightRig_rig_freezing = 6;
    const LightRig_rig_glow = 7;
    const LightRig_rig_harsh = 8;
    const LightRig_rig_legacyFlat1 = 9;
    const LightRig_rig_legacyFlat2 = 10;
    const LightRig_rig_legacyFlat3 = 11;
    const LightRig_rig_legacyFlat4 = 12;
    const LightRig_rig_legacyHarsh1 = 13;
    const LightRig_rig_legacyHarsh2 = 14;
    const LightRig_rig_legacyHarsh3 = 15;
    const LightRig_rig_legacyHarsh4 = 16;
    const LightRig_rig_legacyNormal1 = 17;
    const LightRig_rig_legacyNormal2 = 18;
    const LightRig_rig_legacyNormal3 = 19;
    const LightRig_rig_legacyNormal4 = 20;
    const LightRig_rig_morning = 21;
    const LightRig_rig_soft = 22;
    const LightRig_rig_sunrise = 23;
    const LightRig_rig_sunset = 24;
    const LightRig_rig_threePt = 25;
    const LightRig_rig_twoPt = 26;

    const Bevel_prst_angle = 0;
    const Bevel_prst_artDeco = 1;
    const Bevel_prst_circle = 2;
    const Bevel_prst_convex = 3;
    const Bevel_prst_coolSlant = 4;
    const Bevel_prst_cross = 5;
    const Bevel_prst_divot = 6;
    const Bevel_prst_hardEdge = 7;
    const Bevel_prst_relaxedInset = 8;
    const Bevel_prst_riblet = 9;
    const Bevel_prst_slope = 10;
    const Bevel_prst_softRound = 11;

    const ParameterVal_arrowheadStyle_arr = 0;
    const ParameterVal_arrowheadStyle_auto = 1;
    const ParameterVal_arrowheadStyle_noArr = 2;
    const ParameterVal_autoTextRotation_grav = 0;
    const ParameterVal_autoTextRotation_none = 1;
    const ParameterVal_autoTextRotation_upr = 2;
    const ParameterVal_bendPoint_beg = 0;
    const ParameterVal_bendPoint_def = 1;
    const ParameterVal_bendPoint_end = 2;
    const ParameterVal_breakpoint_bal = 0;
    const ParameterVal_breakpoint_endCnv = 1;
    const ParameterVal_breakpoint_fixed = 2;
    const ParameterVal_centerShapeMapping_fNode = 0;
    const ParameterVal_centerShapeMapping_none = 1;
    const ParameterVal_childAlignment_b = 0;
    const ParameterVal_childAlignment_l = 1;
    const ParameterVal_childAlignment_r = 2;
    const ParameterVal_childAlignment_t = 3;
    const ParameterVal_childDirection_horz = 0;
    const ParameterVal_childDirection_vert = 1;
    const ParameterVal_connectorDimension_1D = 0;
    const ParameterVal_connectorDimension_2D = 1;
    const ParameterVal_connectorDimension_cust = 2;
    const ParameterVal_connectorPoint_auto = 0;
    const ParameterVal_connectorPoint_bCtr = 1;
    const ParameterVal_connectorPoint_bL = 2;
    const ParameterVal_connectorPoint_bR = 3;
    const ParameterVal_connectorPoint_ctr = 4;
    const ParameterVal_connectorPoint_midL = 5;
    const ParameterVal_connectorPoint_midR = 6;
    const ParameterVal_connectorPoint_radial = 7;
    const ParameterVal_connectorPoint_tCtr = 8;
    const ParameterVal_connectorPoint_tL = 9;
    const ParameterVal_connectorPoint_tR = 10;
    const ParameterVal_connectorRouting_bend = 0;
    const ParameterVal_connectorRouting_curve = 1;
    const ParameterVal_connectorRouting_longCurve = 2;
    const ParameterVal_connectorRouting_stra = 3;
    const ParameterVal_continueDirection_revDir = 0;
    const ParameterVal_continueDirection_sameDir = 1;
    const ParameterVal_horizontalAlignment_ctr = 0;
    const ParameterVal_horizontalAlignment_l = 1;
    const ParameterVal_horizontalAlignment_none = 2;
    const ParameterVal_horizontalAlignment_r = 3;
    const ParameterVal_diagramTextAlignment_ctr = 0;
    const ParameterVal_diagramTextAlignment_l = 1;
    const ParameterVal_diagramTextAlignment_r = 2;
    const ParameterVal_fallbackDimension_1D = 0;
    const ParameterVal_fallbackDimension_2D = 1;
    const ParameterVal_flowDirection_col = 0;
    const ParameterVal_flowDirection_row = 1;
    const ParameterVal_growDirection_bL = 0;
    const ParameterVal_growDirection_bR = 1;
    const ParameterVal_growDirection_tL = 2;
    const ParameterVal_growDirection_tR = 3;
    const ParameterVal_hierarchyAlignment_bCtrCh = 0;
    const ParameterVal_hierarchyAlignment_bCtrDes = 1;
    const ParameterVal_hierarchyAlignment_bL = 2;
    const ParameterVal_hierarchyAlignment_bR = 3;
    const ParameterVal_hierarchyAlignment_lB = 4;
    const ParameterVal_hierarchyAlignment_lCtrCh = 5;
    const ParameterVal_hierarchyAlignment_lCtrDes = 6;
    const ParameterVal_hierarchyAlignment_lT = 7;
    const ParameterVal_hierarchyAlignment_rB = 8;
    const ParameterVal_hierarchyAlignment_rCtrCh = 9;
    const ParameterVal_hierarchyAlignment_rCtrDes = 10;
    const ParameterVal_hierarchyAlignment_rT = 11;
    const ParameterVal_hierarchyAlignment_tCtrCh = 12;
    const ParameterVal_hierarchyAlignment_tCtrDes = 13;
    const ParameterVal_hierarchyAlignment_tL = 14;
    const ParameterVal_hierarchyAlignment_tR = 15;
    const ParameterVal_linearDirection_fromB = 0;
    const ParameterVal_linearDirection_fromL = 1;
    const ParameterVal_linearDirection_fromR = 2;
    const ParameterVal_linearDirection_fromT = 3;
    const ParameterVal_nodeHorizontalAlignment_ctr = 0;
    const ParameterVal_nodeHorizontalAlignment_l = 1;
    const ParameterVal_nodeHorizontalAlignment_r = 2;
    const ParameterVal_nodeVerticalAlignment_b = 0;
    const ParameterVal_nodeVerticalAlignment_mid = 1;
    const ParameterVal_nodeVerticalAlignment_t = 2;
    const ParameterVal_offset_ctr = 0;
    const ParameterVal_offset_off = 1;
    const ParameterVal_pyramidAccentPosition_aft = 0;
    const ParameterVal_pyramidAccentPosition_bef = 1;
    const ParameterVal_pyramidAccentTextMargin_stack = 0;
    const ParameterVal_pyramidAccentTextMargin_step = 1;
    const ParameterVal_rotationPath_alongPath = 0;
    const ParameterVal_rotationPath_none = 1;
    const ParameterVal_secondaryChildAlignment_b = 0;
    const ParameterVal_secondaryChildAlignment_l = 1;
    const ParameterVal_secondaryChildAlignment_none = 2;
    const ParameterVal_secondaryChildAlignment_r = 3;
    const ParameterVal_secondaryChildAlignment_t = 4;
    const ParameterVal_secondaryLinearDirection_fromB = 0;
    const ParameterVal_secondaryLinearDirection_fromL = 1;
    const ParameterVal_secondaryLinearDirection_fromR = 2;
    const ParameterVal_secondaryLinearDirection_fromT = 3;
    const ParameterVal_secondaryLinearDirection_none = 4;
    const ParameterVal_startingElement_node = 0;
    const ParameterVal_startingElement_trans = 1;
    const ParameterVal_textAnchorHorizontal_ctr = 0;
    const ParameterVal_textAnchorHorizontal_none = 1;
    const ParameterVal_textAnchorVertical_b = 0;
    const ParameterVal_textAnchorVertical_mid = 1;
    const ParameterVal_textAnchorVertical_t = 2;
    const ParameterVal_textBlockDirection_horz = 0;
    const ParameterVal_textBlockDirection_vert = 1;
    const ParameterVal_textDirection_fromB = 0;
    const ParameterVal_textDirection_fromT = 1;
    const ParameterVal_verticalAlignment_b = 0;
    const ParameterVal_verticalAlignment_mid = 1;
    const ParameterVal_verticalAlignment_none = 2;
    const ParameterVal_verticalAlignment_t = 3;

    const Coordinate_universalMeasure_cm = 0;
    const Coordinate_universalMeasure_mm = 1;
    const Coordinate_universalMeasure_in = 2;
    const Coordinate_universalMeasure_pt = 3;
    const Coordinate_universalMeasure_pc = 4;
    const Coordinate_universalMeasure_pi = 5;

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
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DiagramDataDataModel, this.getDataModel(), oPr));
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
    changesFactory[AscDFH.historyitem_DataModelPtLst] = CChangeObject;
    changesFactory[AscDFH.historyitem_DataModelWhole] = CChangeObject;
    drawingsChangesMap[AscDFH.historyitem_DataModelBg] = function (oClass, value) {
      oClass.bg = value;
    };
    drawingsChangesMap[AscDFH.historyitem_DataModelCxnLst] = function (oClass, value) {
      oClass.cxnLst = value;
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
      this.ptLst = null;
      this.whole = null;
    }

    InitClass(DataModel, CBaseFormatObject, AscDFH.historyitem_type_DataModel);
    DataModel.prototype.getCxnMap = function() {
      if (this.cxnLst) {
        return this.cxnLst.getCxnMap();
      }
    }
    DataModel.prototype.Write_ToBinary = function (w) {
      AscFormat.writeObjectNoId(w, this.bg);
      AscFormat.writeObjectNoId(w, this.cxnLst);
      AscFormat.writeObjectNoId(w, this.ptLst);
      AscFormat.writeObjectNoId(w, this.whole);
    };
    DataModel.prototype.Read_FromBinary = function (r) {
      this.setBg(AscFormat.readObjectNoId(r));
      this.setCxnLst(AscFormat.readObjectNoId(r));
      this.setPtLst(AscFormat.readObjectNoId(r));
      this.setWhole(AscFormat.readObjectNoId(r));
    };
    DataModel.prototype.setBg = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelBg, this.getBg(), oPr));
      this.bg = oPr;
      this.setParentToChild(oPr);
    }

    DataModel.prototype.setCxnLst = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelCxnLst, this.getCxnLst(), oPr));
      this.cxnLst = oPr;
      this.setParentToChild(oPr);
    }

    DataModel.prototype.setPtLst = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelPtLst, this.getPtLst(), oPr));
      this.ptLst = oPr;
      this.setParentToChild(oPr);
    }

    DataModel.prototype.setWhole = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_DataModelWhole, this.getWhole(), oPr));
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

	  DataModel.prototype.getMainPoint = function () {
		  const ptLst = this.getPtLst();

			if (ptLst) {
				const arrPoints = ptLst.list;
				for (let i  = 0; i < arrPoints.length; i += 1) {
					const oPoint = arrPoints[i];
					if (oPoint.getType() === Point_type_doc) {
						return oPoint;
					}
				}
			}
	  }


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
    CCommonDataList.prototype.Write_ToBinary = function (w) {
      w.WriteLong(this.list.length);
      for (let i = 0; i < this.list.length; i += 1) {
        AscFormat.writeObjectNoId(w, this.list[i]);
      }
    };
    CCommonDataList.prototype.Read_FromBinary = function (r) {
      for (let i = r.GetLong(); i > 0; i -= 1) {
        this.list.push(AscFormat.readObjectNoId(r));
      }
    };
    CCommonDataList.prototype.addToLst = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
      oHistory.CanAddChanges() && oHistory.Add(new CChangeContent(this, AscDFH.historyitem_CCommonDataListAdd, nInsertIdx, [oPr], true));
      nInsertIdx === this.list.length ? this.list.push(oPr) : this.list.splice(nInsertIdx, 0, oPr);
      this.setParentToChild(oPr);
    };

    CCommonDataList.prototype.removeFromLst = function (nIdx) {
      if (nIdx > -1 && nIdx < this.list.length) {
        this.list[nIdx].setParent(null);
        oHistory.CanAddChanges() && oHistory.Add(new CChangeContent(this, AscDFH.historyitem_CCommonDataListRemove, nIdx, [this.list[nIdx]], false));
        nIdx === this.list.length - 1 ? this.list.pop() : this.list.splice(nIdx, 1);
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

    function CCommonDataListNoId() {
      AscFormat.CBaseFormatNoIdObject.call(this);
      this.list = [];
    }

    InitClass(CCommonDataListNoId, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_CCommonDataListNoId);
    CCommonDataListNoId.prototype.Write_ToBinary = function (w) {
      w.WriteLong(this.list.length);
      for (let i = 0; i < this.list.length; i += 1) {
        AscFormat.writeObjectNoId(w, this.list[i]);
      }
    };
    CCommonDataListNoId.prototype.Read_FromBinary = function (r) {
      for (let i = r.GetLong(); i > 0; i -= 1) {
        this.list.push(AscFormat.readObjectNoId(r));
      }
    };
    CCommonDataListNoId.prototype.addToLst = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
      nInsertIdx === this.list.length ? this.list.push(oPr) : this.list.splice(nInsertIdx, 0, oPr);
    };

    CCommonDataListNoId.prototype.removeFromLst = function (nIdx) {
      if (nIdx > -1 && nIdx < this.list.length) {
        nIdx === this.list.length - 1 ? this.list.pop() : this.list.splice(nIdx, 1);
      }
    };

    CCommonDataListNoId.prototype.fillObject = function (oCopy, oIdMap) {
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        oCopy.addToLst(nIdx, this.list[nIdx].createDuplicate(oIdMap));
      }
    };

    CCommonDataListNoId.prototype.getChildren = function() {
      return [].concat(this.list);
    };
    CCommonDataListNoId.prototype.privateWriteAttributes = null;
    CCommonDataListNoId.prototype.writeChildren = function(pWriter) {
      for (var i = 0; i < this.list.length; i += 1) {
        this.writeRecord2(pWriter, 0, this.list[i]);
      }
    };
    CCommonDataListNoId.prototype.readAttribute = null;


    function PtLst() {
      CCommonDataList.call(this);
    }

    InitClass(PtLst, CCommonDataList, AscDFH.historyitem_type_PtLst);
    PtLst.prototype.getModelIdStr = function() {
      return this.list.map(function(pt) {
        return pt.getModelId();
      }).join("");
    };
    PtLst.prototype.getPtMap = function() {
      var ptMap = {};
        this.list.forEach(function (point) {
          ptMap[point.modelId] = point;
        });
        return ptMap;
    };
    PtLst.prototype.removeChildrenFromMap = function(pointMap) {
      for (let i = this.list.length - 1; i >= 0; i -= 1) {
        const point = this.list[i];
        if (pointMap[point.getModelId()]) {
          this.removeFromLst(i);
        }
      }
    }
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
    CxnLst.prototype.getCxnMap = function() {
      const cxnMap = {};
      for (let i = 0; i < this.list.length; i++) {
        const cxn = this.list[i];
        if (!cxnMap[cxn.srcId]) {
          cxnMap[cxn.srcId] = {};
        }
        cxnMap[cxn.srcId][cxn.destId] = cxn;
      }
      return cxnMap;
    };
    CxnLst.prototype.removeChildrenFromMap = function(srcDestCxnMap) {
      const cxnMap = {};
      for (let srcId in srcDestCxnMap) {
        const srcConnections = srcDestCxnMap[srcId];
        for (let destId in srcConnections) {
          const cxn = srcConnections[destId];
          cxnMap[cxn.getModelId()] = cxn;
        }
      }

      for (let i = this.list.length - 1; i >= 0; i -= 1) {
        const cxn = this.list[i];
        if (cxnMap[cxn.getModelId()]) {
          this.removeFromLst(i);
        }
      }
    };
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
    changesFactory[AscDFH.historyitem_CxnDestOrd] = CChangeLong;
    changesFactory[AscDFH.historyitem_CxnModelId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnParTransId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnPresId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnSibTransId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnSrcId] = CChangeString;
    changesFactory[AscDFH.historyitem_CxnSrcOrd] = CChangeLong;
    changesFactory[AscDFH.historyitem_CxnType] = CChangeLong;
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
      this.type = Cxn_type_parOf;
    }

    InitClass(Cxn, CBaseFormatObject, AscDFH.historyitem_type_Cxn);
    Cxn.prototype.Write_ToBinary = function (w) {
      AscFormat.writeString(w, this.destId);
      AscFormat.writeLong(w, this.destOrd);
      AscFormat.writeString(w, this.modelId);
      AscFormat.writeString(w, this.parTransId);
      AscFormat.writeString(w, this.presId);
      AscFormat.writeString(w, this.sibTransId);
      AscFormat.writeString(w, this.srcId);
      AscFormat.writeLong(w, this.srcOrd);
      AscFormat.writeLong(w, this.type);
    };
    Cxn.prototype.Read_FromBinary = function (r) {
      this.setDestId(AscFormat.readString(r));
      this.setDestOrd(AscFormat.readLong(r));
      this.setModelId(AscFormat.readString(r));
      this.setParTransId(AscFormat.readString(r));
      this.setPresId(AscFormat.readString(r));
      this.setSibTransId(AscFormat.readString(r));
      this.setSrcId(AscFormat.readString(r));
      this.setSrcOrd(AscFormat.readLong(r));
      this.setType(AscFormat.readLong(r));
    };
		Cxn.getTypeEnum = function (sType) {
			switch (sType) {
				case 'presOf':
					return Cxn_type_presOf;
				case 'parOf':
					return Cxn_type_parOf;
				case 'presParOf':
					return Cxn_type_presParOf;
				default:
					return;
			}
		};
	  Cxn.getTypeString = function (sType) {
		  switch (sType) {
			  case Cxn_type_presOf:
				  return 'presOf';
			  case Cxn_type_parOf:
				  return 'parOf';
			  case Cxn_type_presParOf:
				  return 'presParOf';
			  default:
				  return '';
		  }
	  };
    Cxn.prototype.setDestId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnDestId, this.getDestId(), pr));
      this.destId = pr;
    }

    Cxn.prototype.setDestOrd = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CxnDestOrd, this.getDestOrd(), pr));
      this.destOrd = pr;
    }

    Cxn.prototype.setModelId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnModelId, this.getModelId(), pr));
      this.modelId = pr;
    }

    Cxn.prototype.setParTransId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnParTransId, this.getParTransId(), pr));
      this.parTransId = pr;
    }

    Cxn.prototype.setPresId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnPresId, this.getPresId(), pr));
      this.presId = pr;
    }

    Cxn.prototype.setSibTransId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnSibTransId, this.getSibTransId(), pr));
      this.sibTransId = pr;
    }

    Cxn.prototype.setSrcId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_CxnSrcId, this.getSrcId(), pr));
      this.srcId = pr;
    }

    Cxn.prototype.setSrcOrd = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CxnSrcOrd, this.getSrcOrd(), pr));
      this.srcOrd = pr;
    }

    Cxn.prototype.setType = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeLong(this, AscDFH.historyitem_CxnType, this.getType(), pr));
      this.type = pr;
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
    }

    Cxn.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.modelId);
      pWriter._WriteString2(1, Cxn.getTypeString(this.type));
      pWriter._WriteString2(2, this.destId);
      pWriter._WriteString2(3, typeof this.destOrd === "number" ? this.destOrd.toString() : null);
      pWriter._WriteString2(4, this.srcId);
      pWriter._WriteString2(5, typeof this.srcOrd === "number" ? this.srcOrd.toString() : null);
      pWriter._WriteString2(6, this.parTransId);
      pWriter._WriteString2(7, this.sibTransId);
      pWriter._WriteString2(8, this.presId);
    };
    Cxn.prototype.writeChildren = function(pWriter) {
    };
    Cxn.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setModelId(oStream.GetString2());
      else if (1 === nType) this.setType(Cxn.getTypeEnum(oStream.GetString2()));
      else if (2 === nType) this.setDestId(oStream.GetString2());
      else if (3 === nType) this.setDestOrd(parseInt(oStream.GetString2(), 10));
      else if (4 === nType) this.setSrcId(oStream.GetString2());
      else if (5 === nType) this.setSrcOrd(parseInt(oStream.GetString2(), 10));
      else if (6 === nType) this.setParTransId(oStream.GetString2());
      else if (7 === nType) this.setSibTransId(oStream.GetString2());
      else if (8 === nType) this.setPresId(oStream.GetString2());
    };
    Cxn.prototype.readChild = function(nType, pReader) {

    };


    changesFactory[AscDFH.historyitem_BgFormatFill] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_BgFormatEffect] = CChangeObjectNoId;
    drawingsChangesMap[AscDFH.historyitem_BgFormatFill] = function (oClass, value, bFromLoad) {
      oClass.fill = value;
      oClass.handleUpdateFill();
	    if (bFromLoad) {
		    if (typeof AscCommon.CollaborativeEditing !== "undefined") {
			    if (oClass.fill && oClass.fill.fill && oClass.fill.fill.type ===  Asc.c_oAscFill.FILL_TYPE_BLIP && typeof oClass.fill.fill.RasterImageId === "string" && oClass.fill.fill.RasterImageId.length > 0) {
				    AscCommon.CollaborativeEditing.Add_NewImage(oClass.fill.fill.RasterImageId);
			    }
		    }
	    }
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

    BgFormat.prototype.Write_ToBinary = function (w) {
      AscFormat.writeObjectNoIdNoType(w, this.fill);
      AscFormat.writeObjectNoId(w, this.effect);
    };
    BgFormat.prototype.Read_FromBinary = function (r) {
      this.setFill(AscFormat.readObjectNoIdNoType(r, AscFormat.CUniFill));
      this.setEffect(AscFormat.readObjectNoId(r));
    };
    BgFormat.prototype.setFill = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_BgFormatFill, this.getFill(), oPr));
      this.fill = oPr;
      this.handleUpdateFill();
    }

    BgFormat.prototype.setEffect = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_BgFormatEffect, this.getEffect(), oPr));
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
    Whole.prototype.Write_ToBinary = function (w) {
      AscFormat.writeObjectNoIdNoType(w, this.ln);
      AscFormat.writeObjectNoId(w, this.effect);
    };
    Whole.prototype.Read_FromBinary = function (r) {
      this.setLn(AscFormat.readObjectNoIdNoType(r, AscFormat.CLn));
      this.setEffect(AscFormat.readObjectNoId(r));
    };
    Whole.prototype.setEffect = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_WholeEffect, this.getEffect(), oPr));
      this.effect = oPr;
      this.setParentToChild(oPr);
    }

    Whole.prototype.setLn = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_WholeLn, this.getLn(), oPr));
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

    changesFactory[AscDFH.historyitem_PointCxnId] = CChangeString;
    changesFactory[AscDFH.historyitem_PointModelId] = CChangeString;
    changesFactory[AscDFH.historyitem_PointType] = CChangeLong;
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
      this.type = Point_type_node;

      this.prSet = null;
      this.spPr = null;
      this.t = null;
    }

    InitClass(Point, CBaseFormatObject, AscDFH.historyitem_type_Point);
    // These methods are needed to write sampData
    Point.prototype.Write_ToBinary = function (w) {
      AscFormat.writeString(w, this.cxnId);
      AscFormat.writeString(w, this.modelId);
      AscFormat.writeLong(w, this.type);
      AscFormat.writeObjectNoId(w, this.prSet);
      // AscFormat.writeObject(w, this.spPr);
      // AscFormat.writeObject(w, this.t);
    };
    Point.prototype.Read_FromBinary = function (r) {
      this.cxnId = AscFormat.readString(r);
      this.modelId = AscFormat.readString(r);
      this.type = AscFormat.readLong(r);
      this.prSet = AscFormat.readObjectNoId(r);
      // this.spPr = AscFormat.readObject(r);
      // this.t = AscFormat.readObject(r);
    }
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
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PointCxnId, this.getCxnId(), pr));
      this.cxnId = pr;
    }

    Point.prototype.setModelId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PointModelId, this.getModelId(), pr));
      this.modelId = pr;

    }

    Point.prototype.getShape = function () {
      if (this.parent && this.parent.parent instanceof AscFormat.CShape) {
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
      oHistory.CanAddChanges() && oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PointType, this.getType(), pr));
      this.type = pr;
    }

    Point.prototype.setPrSet = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PointPrSet, this.getPrSet(), oPr));
      this.prSet = oPr;
      this.setParentToChild(oPr);
    }

    Point.prototype.setSpPr = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PointSpPr, this.getSpPr(), oPr));
      this.spPr = oPr;
      this.setParentToChild(oPr);
    }

    Point.prototype.setT = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PointT, this.getT(), oPr));
      this.t = oPr;
      this.setParentToChild(oPr);
    }

    Point.prototype.setPhldrT = function(pr) {
      var prSet = this.getPrSet();
      prSet && prSet.setPhldrT(pr);
    }

    Point.prototype.getPhldrT = function() {
      var prSet = this.getPrSet();
      return prSet && prSet.getPhldrT();
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
      //TODO: The method is a crutch. in the future, you need to determine the picture placeholder from the layout.xml file
      const pointAssociationPrSet = this.prSet;
      if (pointAssociationPrSet) {
        const sStyleLbl = pointAssociationPrSet.presStyleLbl;
        const sName = pointAssociationPrSet.presName;
        const oExcludes = {
          'node1': ['imageRepeatNode'],
          'alignImgPlace1': ['ChildAccent', 'bentUpArrow1', 'ParentShape1', 'ParentShape2', 'Text1', 'Text2', 'Text3', 'Text4', 'Text5', 'Text6'],
          'bgImgPlace1': ['LeftNode', 'RightNode', 'Background'],
	        'alignNode1': ['imageAccentRepeatNode']
        };
        if (oExcludes[sStyleLbl]) {
          if (oExcludes[sStyleLbl].indexOf(sName) !== -1) {
            return false;
          }
        }
        const imagePlaceholderArrStylelbl = ['alignImgPlace1', 'bgImgPlace1', 'fgImgPlace1'];
        const imagePlaceholderArrName = ['Image', 'imageRepeatNode', 'pictRect'];
        let bCheckImagePlaceholderStyleLbl = imagePlaceholderArrStylelbl.indexOf(sStyleLbl) !== -1;
        let bCheckImagePlaceholderName = false;
        if (sName) {
          bCheckImagePlaceholderName = imagePlaceholderArrName.indexOf(sName) !== -1 || sName.indexOf('image') !== -1;
        }

        return (bCheckImagePlaceholderStyleLbl ||
          bCheckImagePlaceholderName ||
          (sName === 'rect1' && sStyleLbl === 'bgShp') ||
          (sName === 'rect1' && sStyleLbl === 'lnNode1') ||
          (sName === 'adorn' && sStyleLbl === 'fgAccFollowNode1'));
      }

      return false;
    }

    Point.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setCxnId(this.getCxnId());
      oCopy.setModelId(this.getModelId());
      oCopy.setType(this.getType());
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

	  Point.prototype.getPresName = function () {
		  return this.prSet && this.prSet.presName;
	  };

	  Point.prototype.getPresAssocID = function () {
		  return this.prSet && this.prSet.presAssocID;
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
    changesFactory[AscDFH.historyitem_PrSetPresLayoutVars] = CChangeObjectNoId;
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
    drawingConstructorsMap[AscDFH.historyitem_PrSetPresLayoutVars] = VarLst;

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
      PrSet.prototype.Write_ToBinary = function (w) {
        AscFormat.writeBool(w, this.coherent3DOff);
        AscFormat.writeString(w, this.csCatId);
        AscFormat.writeString(w, this.csTypeId);
        AscFormat.writeLong(w, this.custAng);
        AscFormat.writeBool(w, this.custFlipHor);
        AscFormat.writeBool(w, this.custFlipVert);
        AscFormat.writeLong(w, this.custLinFactNeighborX);
        AscFormat.writeLong(w, this.custLinFactNeighborY);
        AscFormat.writeLong(w, this.custLinFactX);
        AscFormat.writeLong(w, this.custLinFactY);
        AscFormat.writeLong(w, this.custRadScaleInc);
        AscFormat.writeLong(w, this.custRadScaleRad);
        AscFormat.writeLong(w, this.custScaleX);
        AscFormat.writeLong(w, this.custScaleY);
        AscFormat.writeLong(w, this.custSzX);
        AscFormat.writeLong(w, this.custSzY);
        AscFormat.writeBool(w, this.custT);
        AscFormat.writeString(w, this.loCatId);
        AscFormat.writeString(w, this.loTypeId);
        AscFormat.writeBool(w, this.phldr);
        AscFormat.writeString(w, this.phldrT);
        AscFormat.writeString(w, this.presAssocID);
        AscFormat.writeString(w, this.presName);
        AscFormat.writeLong(w, this.presStyleCnt);
        AscFormat.writeLong(w, this.presStyleIdx);
        AscFormat.writeString(w, this.presStyleLbl);
        AscFormat.writeString(w, this.qsCatId);
        AscFormat.writeString(w, this.qsTypeId);
        AscFormat.writeObjectNoId(w, this.presLayoutVars);
        AscFormat.writeObject(w, this.style);
      }
    PrSet.prototype.Read_FromBinary = function (r) {
      this.coherent3DOff = AscFormat.readBool(r);
      this.csCatId = AscFormat.readString(r);
      this.csTypeId = AscFormat.readString(r);
      this.custAng = AscFormat.readLong(r);
      this.custFlipHor = AscFormat.readBool(r);
      this.custFlipVert = AscFormat.readBool(r);
      this.custLinFactNeighborX = AscFormat.readLong(r);
      this.custLinFactNeighborY = AscFormat.readLong(r);
      this.custLinFactX = AscFormat.readLong(r);
      this.custLinFactY = AscFormat.readLong(r);
      this.custRadScaleInc = AscFormat.readLong(r);
      this.custRadScaleRad = AscFormat.readLong(r);
      this.custScaleX = AscFormat.readLong(r);
      this.custScaleY = AscFormat.readLong(r);
      this.custSzX = AscFormat.readLong(r);
      this.custSzY = AscFormat.readLong(r);
      this.custT = AscFormat.readBool(r);
      this.loCatId = AscFormat.readString(r);
      this.loTypeId = AscFormat.readString(r);
      this.phldr = AscFormat.readBool(r);
      this.phldrT = AscFormat.readString(r);
      this.presAssocID = AscFormat.readString(r);
      this.presName = AscFormat.readString(r);
      this.presStyleCnt = AscFormat.readLong(r);
      this.presStyleIdx = AscFormat.readLong(r);
      this.presStyleLbl = AscFormat.readString(r);
      this.qsCatId = AscFormat.readString(r);
      this.qsTypeId = AscFormat.readString(r);
      this.presLayoutVars = AscFormat.readObjectNoId(r);
      this.style = AscFormat.readObject(r);
    }
    PrSet.prototype.setCoherent3DOff = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetCoherent3DOff, this.getCoherent3DOff(), pr));
      this.coherent3DOff = pr;
    };

    PrSet.prototype.setCsCatId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetCsCatId, this.getCsCatId(), pr));
      this.csCatId = pr;
    };

    PrSet.prototype.setCsTypeId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetCsTypeId, this.getCsTypeId(), pr))
      this.csTypeId = pr;
    };

    PrSet.prototype.setCustAng = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustAng, this.getCustAng(), pr));
      this.custAng = pr;
    };

    PrSet.prototype.setCustFlipHor = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetCustFlipHor, this.getCustFlipHor(), pr));
      this.custFlipHor = pr;
    };

    PrSet.prototype.setCustFlipVert = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetCustFlipVert, this.getCustFlipVert(), pr));
      this.custFlipVert = pr;
    };

    PrSet.prototype.setCustLinFactNeighborX = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustLinFactNeighborX, this.getCustLinFactNeighborX(), pr));
      this.custLinFactNeighborX = pr;
    };

    PrSet.prototype.setCustLinFactNeighborY = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustLinFactNeighborY, this.getCustLinFactNeighborY(), pr));
      this.custLinFactNeighborY = pr;
    };

    PrSet.prototype.setCustLinFactX = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustLinFactX, this.getCustLinFactX(), pr));
      this.custLinFactX = pr;
    };

    PrSet.prototype.setCustLinFactY = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustLinFactY, this.getCustLinFactY(), pr));
      this.custLinFactY = pr;
    };

    PrSet.prototype.setCustRadScaleInc = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustRadScaleInc, this.getCustRadScaleInc(), pr));
      this.custRadScaleInc = pr;
    };

    PrSet.prototype.setCustRadScaleRad = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustRadScaleRad, this.getCustRadScaleRad(), pr));
      this.custRadScaleRad = pr;
    };

    PrSet.prototype.setCustScaleX = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustScaleX, this.getCustScaleX(), pr));
      this.custScaleX = pr;
    };

    PrSet.prototype.setCustScaleY = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeDouble2(this, AscDFH.historyitem_PrSetCustScaleY, this.getCustScaleY(), pr));
      this.custScaleY = pr;
    };

    PrSet.prototype.setCustSzX = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PrSetCustSzX, this.getCustSzX(), pr));
      this.custSzX = pr;
    };

    PrSet.prototype.setCustSzY = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PrSetCustSzY, this.getCustSzY(), pr));
      this.custSzY = pr;
    };

    PrSet.prototype.setCustT = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetCustT, this.getCustT(), pr));
      this.custT = pr;
    };

    PrSet.prototype.setLoCatId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetLoCatId, this.getLoCatId(), pr));
      this.loCatId = pr;
    };

    PrSet.prototype.setLoTypeId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetLoTypeId, this.getLoTypeId(), pr));
      this.loTypeId = pr;
    };

    PrSet.prototype.setPhldr = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeBool(this, AscDFH.historyitem_PrSetPhldr, this.getPhldr(), pr));
      this.phldr = pr;
    };

    PrSet.prototype.setPhldrT = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetPhldrT, this.getPhldrT(), pr));
      this.phldrT = pr;
    };

    PrSet.prototype.setPresAssocID = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetPresAssocID, this.getPresAssocID(), pr));
      this.presAssocID = pr;
    };

    PrSet.prototype.setPresName = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetPresName, this.getPresName(), pr));
      this.presName = pr;
    };
    PrSet.prototype.setPresStyleCnt = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PrSetPresStyleCnt, this.getPresStyleCnt(), pr));
      this.presStyleCnt = pr;
    };

    PrSet.prototype.setPresStyleIdx = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeLong(this, AscDFH.historyitem_PrSetPresStyleIdx, this.getPresStyleIdx(), pr));
      this.presStyleIdx = pr;
    };

    PrSet.prototype.setPresStyleLbl = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetPresStyleLbl, this.getPresStyleLbl(), pr));
      this.presStyleLbl = pr;
    };

    PrSet.prototype.setQsCatId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetQsCatId, this.getQsCatId(), pr));
      this.qsCatId = pr;
    };

    PrSet.prototype.setQsTypeId = function (pr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeString(this, AscDFH.historyitem_PrSetQsTypeId, this.getQsTypeId(), pr));
      this.qsTypeId = pr;
    };

    PrSet.prototype.setStyle = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_PrSetStyle, this.getStyle(), oPr));
      this.style = oPr;
    };

    PrSet.prototype.setPresLayoutVars = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_PrSetPresLayoutVars, this.getPresLayoutVars(), oPr));
      this.presLayoutVars = oPr;
    };

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
    };

    PrSet.prototype.getCoherent3DOff = function () {
      return this.coherent3DOff;
    };

    PrSet.prototype.getCsCatId = function () {
      return this.csCatId;
    };

    PrSet.prototype.getCsTypeId = function () {
      return this.csTypeId;
    };

    PrSet.prototype.getCustAng = function () {
      return this.custAng;
    };

    PrSet.prototype.getCustFlipHor = function () {
      return this.custFlipHor;
    };

    PrSet.prototype.getCustFlipVert = function () {
      return this.custFlipVert;
    };

    PrSet.prototype.getCustLinFactNeighborX = function () {
      return this.custLinFactNeighborX;
    };

    PrSet.prototype.getCustLinFactNeighborY = function () {
      return this.custLinFactNeighborY;
    };

    PrSet.prototype.getCustLinFactX = function () {
      return this.custLinFactX;
    };

    PrSet.prototype.getCustLinFactY = function () {
      return this.custLinFactY;
    };

    PrSet.prototype.getCustRadScaleInc = function () {
      return this.custRadScaleInc;
    };

    PrSet.prototype.getCustRadScaleRad = function () {
      return this.custRadScaleRad;
    };

    PrSet.prototype.getCustScaleX = function () {
      return this.custScaleX;
    };

    PrSet.prototype.getCustScaleY = function () {
      return this.custScaleY;
    };

    PrSet.prototype.getCustSzX = function () {
      return this.custSzX;
    };

    PrSet.prototype.getCustSzY = function () {
      return this.custSzY;
    };

    PrSet.prototype.getCustT = function () {
      return this.custT;
    };

    PrSet.prototype.getLoCatId = function () {
      return this.loCatId;
    };

    PrSet.prototype.getLoTypeId = function () {
      return this.loTypeId;
    };

    PrSet.prototype.getPhldr = function () {
      return this.phldr;
    };

    PrSet.prototype.getPhldrT = function () {
      return this.phldrT;
    };

    PrSet.prototype.getPresAssocID = function () {
      return this.presAssocID;
    };

    PrSet.prototype.getPresName = function () {
      return this.presName;
    };

    PrSet.prototype.getPresStyleCnt = function () {
      return this.presStyleCnt;
    };

    PrSet.prototype.getPresStyleIdx = function () {
      return this.presStyleIdx;
    };

    PrSet.prototype.getPresStyleLbl = function () {
      return this.presStyleLbl;
    };

    PrSet.prototype.getQsCatId = function () {
      return this.qsCatId;
    };

    PrSet.prototype.getQsTypeId = function () {
      return this.qsTypeId;
    };

    PrSet.prototype.getStyle = function () {
      return this.style;
    };

    PrSet.prototype.getPresLayoutVars = function () {
      return this.presLayoutVars;
    };

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
      pWriter._WriteInt2(11, this.custRadScaleInc ? Math.floor(this.custRadScaleInc * 100000) : null);
      pWriter._WriteInt2(12, this.custRadScaleRad ? Math.floor(this.custRadScaleRad * 100000) : null);
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
      else if (11 === nType) this.setCustRadScaleInc(oStream.GetLong() / 100000);
      else if (12 === nType) this.setCustRadScaleRad(oStream.GetLong() / 100000);
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
          const varLst = new VarLst();
          varLst.fromPPTY(pReader);
          this.setPresLayoutVars(varLst);
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
    };

    function LayoutDef() {
      CBaseFormatNoIdObject.call(this);
      this.defStyle = null;
      this.minVer = null;
      this.uniqueId = null;
      this.catLst = null;
      this.clrData = null;
      this.desc = null;
      this.layoutNode = null;
      this.sampData = null;
      this.styleData = null;
      this.title = null;
    }

    InitClass(LayoutDef, CBaseFormatNoIdObject, AscDFH.historyitem_type_LayoutDef);

    LayoutDef.prototype.Write_ToBinary = function (w) {
      AscFormat.writeString(w, this.defStyle);
      AscFormat.writeString(w, this.minVer);
      AscFormat.writeString(w, this.uniqueId);
      AscFormat.writeObjectNoId(w, this.catLst);
      AscFormat.writeObjectNoId(w, this.clrData);
      AscFormat.writeObjectNoId(w, this.title);
      AscFormat.writeObjectNoId(w, this.layoutNode);
      AscFormat.writeObjectNoId(w, this.sampData);
      AscFormat.writeObjectNoId(w, this.styleData);
      AscFormat.writeObjectNoId(w, this.desc);
    };
    LayoutDef.prototype.Read_FromBinary = function (r) {
      this.defStyle = AscFormat.readString(r);
      this.minVer = AscFormat.readString(r);
      this.uniqueId = AscFormat.readString(r);
      this.catLst = AscFormat.readObjectNoId(r);
      this.clrData = AscFormat.readObjectNoId(r);
      this.title = AscFormat.readObjectNoId(r);
      this.layoutNode = AscFormat.readObjectNoId(r);
      this.sampData = AscFormat.readObjectNoId(r);
      this.styleData = AscFormat.readObjectNoId(r);
      this.desc = AscFormat.readObjectNoId(r);
    };
    LayoutDef.prototype.setDefStyle = function (pr) {
      this.defStyle = pr;
    };

    LayoutDef.prototype.setMinVer = function (pr) {
      this.minVer = pr;
    };

    LayoutDef.prototype.setUniqueId = function (pr) {
      this.uniqueId = pr;
    };

    LayoutDef.prototype.setCatLst = function (oPr) {
      this.catLst = oPr;
    };

    LayoutDef.prototype.setClrData = function (oPr) {
      this.clrData = oPr;
    };

    LayoutDef.prototype.setLayoutNode = function (oPr) {
      this.layoutNode = oPr;
    };

    LayoutDef.prototype.setSampData = function (oPr) {
      this.sampData = oPr;
    };

    LayoutDef.prototype.setStyleData = function (oPr) {
      this.styleData = oPr;
    };

    LayoutDef.prototype.setTitle = function (oPr) {
      this.title = oPr;
    };

    LayoutDef.prototype.setDesc = function (oPr) {
      this.desc = oPr;
    };

    LayoutDef.prototype.getDefStyle = function () {
      return this.defStyle;
    };

    LayoutDef.prototype.getMinVer = function () {
      return this.minVer;
    };

    LayoutDef.prototype.getUniqueId = function () {
      return this.uniqueId;
    };

    LayoutDef.prototype.getCatLst = function () {
      return this.catLst;
    };

    LayoutDef.prototype.getClrData = function () {
      return this.clrData;
    };

    LayoutDef.prototype.getDesc = function () {
      return this.desc;
    };

    LayoutDef.prototype.getLayoutNode = function () {
      return this.layoutNode;
    };

    LayoutDef.prototype.getSampData = function () {
      return this.sampData;
    };

    LayoutDef.prototype.getStyleData = function () {
      return this.styleData;
    };

    LayoutDef.prototype.getTitle = function () {
      return this.title;
    };

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
    };

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
          this.setDesc(new DiagramTitle());
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
          this.setStyleData(new SampData());
          this.styleData.fromPPTY(pReader);
          break;
        }
        case 5: {
          this.setClrData(new SampData());
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


    function CatLst() {
      CCommonDataListNoId.call(this);
    }

    InitClass(CatLst, CCommonDataListNoId, AscDFH.historyitem_type_CatLst);

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

    function SCat() {
      CBaseFormatNoIdObject.call(this);
      this.pri = null;
      this.type = null;
    }

    InitClass(SCat, CBaseFormatNoIdObject, AscDFH.historyitem_type_SCat);
    SCat.prototype.Write_ToBinary = function (w) {
      AscFormat.writeLong(w, this.pri);
      AscFormat.writeString(w, this.type);
    };
    SCat.prototype.Read_FromBinary = function (r) {
      this.pri = AscFormat.readLong(r);
      this.type = AscFormat.readString(r);
    };
    SCat.prototype.setPri = function (pr) {
      this.pri = pr;
    }

    SCat.prototype.setType = function (pr) {
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

		function LayoutBaseClass() {
			AscFormat.CBaseFormatNoIdObject.call(this);
			this.list = [];
			this.alg = null;
			this.shape = null;
			this.presOf = null;
			this.constrLst = null;
			this.ruleLst = null;
			this.varLst = null;
		}
		AscFormat.InitClassWithoutType(LayoutBaseClass, AscFormat.CBaseFormatNoIdObject);
	  LayoutBaseClass.prototype.addToLst = function (idx , pr) {
			idx = Math.min(this.list.length, Math.max(idx, 0));
		  this.list.splice(idx, 0, pr);
	  }
	  LayoutBaseClass.prototype.removeFromLst = function (idx) {
		  idx = Math.min(this.list.length, Math.max(idx, 0));
		  this.list.splice(idx, 1);
	  }
	  LayoutBaseClass.prototype.Write_ToBinary = function (w) {
		  AscFormat.writeObjectNoId(w, this.alg);
		  AscFormat.writeObjectNoId(w, this.shape);
		  AscFormat.writeObjectNoId(w, this.presOf);
		  AscFormat.writeObjectNoId(w, this.constrLst);
		  AscFormat.writeObjectNoId(w, this.ruleLst);
		  AscFormat.writeObjectNoId(w, this.varLst);

		  w.WriteLong(this.list.length);
		  for (let i = 0; i < this.list.length; i += 1) {
			  AscFormat.writeObjectNoId(w, this.list[i]);
		  }
	  };
	  LayoutBaseClass.prototype.Read_FromBinary = function (r) {
		  this.setAlg(AscFormat.readObjectNoId(r));
		  this.setShape(AscFormat.readObjectNoId(r));
		  this.setPresOf(AscFormat.readObjectNoId(r));
		  this.setConstrLst(AscFormat.readObjectNoId(r));
		  this.setRuleLst(AscFormat.readObjectNoId(r));
		  this.setVarLst(AscFormat.readObjectNoId(r));

		  for (let i = r.GetLong(); i > 0; i -= 1) {
			  this.list.push(AscFormat.readObjectNoId(r));
		  }
	  };
	  LayoutBaseClass.prototype.fillObject = function (oCopy, oIdMap) {
		  if (this.alg) {
			  oCopy.setAlg(this.alg.createDuplicate());
		  }
		  if (this.shape) {
			  oCopy.setShape(this.shape.createDuplicate());
		  }
		  if (this.presOf) {
			  oCopy.setPresOf(this.presOf.createDuplicate());
		  }
		  if (this.constrLst) {
			  oCopy.setConstrLst(this.constrLst.createDuplicate());
		  }
		  if (this.ruleLst) {
			  oCopy.setRuleLst(this.ruleLst.createDuplicate());
		  }
		  if (this.varLst) {
			  oCopy.setVarLst(this.varLst.createDuplicate());
		  }
		  for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
			  oCopy.addToLst(nIdx, this.list[nIdx].createDuplicate(oIdMap));
		  }
	  }
	  LayoutBaseClass.prototype.setAlg = function (pr) {
		  this.alg = pr;
	  }
	  LayoutBaseClass.prototype.setShape = function (pr) {
		  this.shape = pr;
	  }
	  LayoutBaseClass.prototype.setPresOf = function (pr) {
		  this.presOf = pr;
	  }
	  LayoutBaseClass.prototype.setConstrLst = function (pr) {
		  this.constrLst = pr;
	  }
	  LayoutBaseClass.prototype.setRuleLst = function (pr) {
		  this.ruleLst = pr;
	  }
	  LayoutBaseClass.prototype.setVarLst = function (pr) {
		  this.varLst = pr;
	  }
	  LayoutBaseClass.prototype.readElement = function(pReader, nType) {
		  let oListElement = null;
		  switch(nType) {
			  case 0xb1: {
				  const oElement = new Alg();
				  oElement.fromPPTY(pReader);
				  this.setAlg(oElement);
				  break;
			  }
			  case 0xb2:
				  oListElement = new Choose();
				  break;
			  case 0xb3: {
				  const oElement = new ConstrLst();
				  oElement.fromPPTY(pReader);
				  this.setConstrLst(oElement);
				  break;
			  }
			  case 0xb4:
				  oListElement = new ForEach();
				  break;
			  case 0xb5:
				  oListElement = new LayoutNode();
				  break;
			  case 0xb6: {
				  const oElement = new PresOf();
				  oElement.fromPPTY(pReader);
				  this.setPresOf(oElement);
				  break;
			  }
			  case 0xb7: {
				  const oElement = new RuleLst();
				  oElement.fromPPTY(pReader);
				  this.setRuleLst(oElement);
				  break;
			  }
			  case 0xb8: {
				  const oElement = new SShape();
				  oElement.fromPPTY(pReader);
				  this.setShape(oElement);
				  break;
			  }
			  case 0xb9: {
				  const oElement = new VarLst();
				  oElement.fromPPTY(pReader);
				  this.setVarLst(oElement);
				  break;
			  }
			  default: {
				  pReader.stream.SkipRecord();
				  break;
			  }
		  }
		  if(oListElement) {
			  oListElement.fromPPTY(pReader);
			  this.addToLst(this.list.length, oListElement);
		  }
	  };
	  LayoutBaseClass.prototype.writeChildren = function(pWriter) {
		  if (this.alg) {
			  this.writeRecord2(pWriter, 0xb1, this.alg);
		  }
		  if (this.shape) {
			  this.writeRecord2(pWriter, 0xb8, this.shape);
		  }
		  if (this.presOf) {
			  this.writeRecord2(pWriter, 0xb6, this.presOf);
		  }
		  if (this.constrLst) {
			  this.writeRecord2(pWriter, 0xb3, this.constrLst);
		  }
		  if (this.ruleLst) {
			  this.writeRecord2(pWriter, 0xb7, this.ruleLst);
		  }
		  if (this.varLst) {
			  this.writeRecord2(pWriter, 0xb9, this.varLst);
		  }
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
	  LayoutBaseClass.prototype.getChildren = function() {
		  return [this.alg, this.shape, this.presOf, this.constrLst, this.ruleLst, this.varLst].concat(this.list);
	  };

    function LayoutNode() {
	    LayoutBaseClass.call(this);
      this.chOrder = null;
      this.moveWith = null;
      this.name = null;
      this.styleLbl = null;
    }

    InitClass(LayoutNode, LayoutBaseClass, AscDFH.historyitem_type_LayoutNode);
	  LayoutNode.prototype.Write_ToBinary = function (w) {
			AscFormat.writeLong(w, this.chOrder);
			AscFormat.writeString(w, this.moveWith);
			AscFormat.writeString(w, this.name);
			AscFormat.writeString(w, this.styleLbl);

		  LayoutBaseClass.prototype.Write_ToBinary.call(this, w);
	  };
	  LayoutNode.prototype.Read_FromBinary = function (r) {
		  this.chOrder = AscFormat.readLong(r);
		  this.moveWith = AscFormat.readString(r);
		  this.name = AscFormat.readString(r);
		  this.styleLbl = AscFormat.readString(r);

		  LayoutBaseClass.prototype.Read_FromBinary.call(this, r);
	  };

    LayoutNode.prototype.setChOrder = function (pr) {
      this.chOrder = pr;
    };

    LayoutNode.prototype.setMoveWith = function (pr) {
      this.moveWith = pr;
    };

    LayoutNode.prototype.setName = function (pr) {
      this.name = pr;
    };

    LayoutNode.prototype.setStyleLbl = function (pr) {
      this.styleLbl = pr;
    };

    LayoutNode.prototype.getChOrder = function () {
      return this.chOrder;
    };

    LayoutNode.prototype.getMoveWith = function () {
      return this.moveWith;
    };

    LayoutNode.prototype.getName = function () {
      return this.name;
    };

    LayoutNode.prototype.getStyleLbl = function () {
      return this.styleLbl;
    };

    LayoutNode.prototype.fillObject = function (oCopy, oIdMap) {
	    LayoutBaseClass.prototype.fillObject.call(this, oCopy, oIdMap);
      oCopy.setChOrder(this.getChOrder());
      oCopy.setMoveWith(this.getMoveWith());
      oCopy.setName(this.getName());
      oCopy.setStyleLbl(this.getStyleLbl());
    }
    LayoutNode.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
      pWriter._WriteString2(1, this.styleLbl);
      pWriter._WriteString2(2, this.moveWith);
      pWriter._WriteUChar2(3, this.chOrder);
    };
    LayoutNode.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
      else if (1 === nType) this.setStyleLbl(oStream.GetString2());
      else if (2 === nType) this.setMoveWith(oStream.GetString2());
      else if (3 === nType) this.setChOrder(oStream.GetUChar());

    };
    LayoutNode.prototype.readChild = function(nType, pReader) {
      this.readElement(pReader, nType);
    };

    function Alg() {
      CBaseFormatNoIdObject.call(this);
      this.rev = null;
      this.type = null;
      this.param = [];
    }

    InitClass(Alg, CBaseFormatNoIdObject, AscDFH.historyitem_type_Alg);
    Alg.prototype.Write_ToBinary = function (w) {
      AscFormat.writeLong(w, this.rev);
      AscFormat.writeLong(w, this.type);
      w.WriteLong(this.param.length);
      for (let i = 0; i < this.param.length; i++) {
        AscFormat.writeObjectNoId(w, this.param[i]);
      }
    };
    Alg.prototype.Read_FromBinary = function (r) {
      this.rev = AscFormat.readLong(r);
      this.type = AscFormat.readLong(r);
      for (let i = r.GetLong(); i > 0; i -= 1) {
        this.param.push(AscFormat.readObjectNoId(r));
      }
    };
    Alg.prototype.setRev = function (pr) {
      this.rev = pr;
    }

    Alg.prototype.setType = function (pr) {
      this.type = pr;
    }

    Alg.prototype.addToLstParam = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.param.length, Math.max(0, nIdx));
      nInsertIdx === this.param.length ? this.param.push(oPr) : this.param.splice(nInsertIdx, 0, oPr);
    };

    Alg.prototype.removeFromLstParam = function (nIdx) {
      if (nIdx > -1 && nIdx < this.param.length) {
        nIdx === this.param.length - 1 ? this.param.pop() : this.param.splice(nIdx, 1);
      }
    };

    Alg.prototype.getRev = function () {
      return this.rev;
    }

    Alg.prototype.getType = function () {
      return this.type;
    }

    Alg.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setRev(this.getRev());
      oCopy.setType(this.getType());
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

    function Param() {
      CBaseFormatNoIdObject.call(this);
      this.type = null;
      this.val = null;
    }

    InitClass(Param, CBaseFormatNoIdObject, AscDFH.historyitem_type_Param);
    Param.prototype.Write_ToBinary = function (w) {
      AscFormat.writeLong(w, this.type);
      AscFormat.writeString(w, this.val);
    };
    Param.prototype.Read_FromBinary = function (r) {
      this.type = AscFormat.readLong(r);
      this.val = AscFormat.readString(r);
    };
		Param.prototype.getValEnum = function () {
			switch (this.type) {
				case Param_type_ctrShpMap: {
					switch (this.val) {
						case 'fNode':
							return ParameterVal_centerShapeMapping_fNode;
						case 'none':
							return ParameterVal_centerShapeMapping_none;
						default:
							return this.val;
					}
				}
				case Param_type_autoTxRot: {
					switch (this.val) {
						case 'upr':
							return ParameterVal_autoTextRotation_upr;
						case 'grav':
							return ParameterVal_autoTextRotation_grav;
						case 'none':
							return ParameterVal_autoTextRotation_none;
						default:
							return this.val;
					}
				}
				case Param_type_rotPath: {
					switch (this.val) {
						case 'alongPath':
							return ParameterVal_rotationPath_alongPath;
						case 'none':
							return ParameterVal_rotationPath_none;
						default:
							return this.val;
					}
				}
				case Param_type_fallback: {
					switch (this.val) {
						case '1D':
							return ParameterVal_fallbackDimension_1D;
						case '2D':
							return ParameterVal_fallbackDimension_2D;
						default:
							return this.val;
					}
				}
				case Param_type_bendPt: {
					switch (this.val) {
						case 'end':
							return ParameterVal_bendPoint_end;
						case 'beg':
							return ParameterVal_bendPoint_beg;
						case 'def':
							return ParameterVal_bendPoint_def;
						default:
							return this.val;
					}
				}
				case Param_type_pyraAcctPos: {
					switch (this.val) {
						case 'bef':
							return ParameterVal_pyramidAccentPosition_bef;
						case 'aft':
							return ParameterVal_pyramidAccentPosition_aft;
						default:
							return this.val;
					}
				}
				case Param_type_nodeHorzAlign: {
					switch (this.val) {
						case 'l':
							return ParameterVal_nodeHorizontalAlignment_l;
						case 'ctr':
							return ParameterVal_nodeHorizontalAlignment_ctr;
						case 'r':
							return ParameterVal_nodeHorizontalAlignment_r;
						default:
							return this.val;
					}
				}
				case Param_type_nodeVertAlign: {
					switch (this.val) {
						case 'b':
							return ParameterVal_nodeVerticalAlignment_b;
						case 'mid':
							return ParameterVal_nodeVerticalAlignment_mid;
						case 't':
							return ParameterVal_nodeVerticalAlignment_t;
						default:
							return this.val;
					}
				}
				case Param_type_connRout: {
					switch (this.val) {
						case 'longCurve':
							return ParameterVal_connectorRouting_longCurve;
						case 'curve':
							return ParameterVal_connectorRouting_curve;
						case 'stra':
							return ParameterVal_connectorRouting_stra;
						case 'bend':
							return ParameterVal_connectorRouting_bend;
						default:
							return this.val;
					}
				}
				case Param_type_dim: {
					switch (this.val) {
						case '2D':
							return ParameterVal_connectorDimension_2D;
						case '1D':
							return ParameterVal_connectorDimension_1D;
						case 'cust':
							return ParameterVal_connectorDimension_cust;
						default:
							return this.val;
					}
				}
				case Param_type_begSty:
				case Param_type_endSty: {
					switch (this.val) {
						case 'arr':
							return ParameterVal_arrowheadStyle_arr;
						case 'noArr':
							return ParameterVal_arrowheadStyle_noArr;
						case 'auto':
							return ParameterVal_arrowheadStyle_auto;
						default:
							return this.val;
					}
				}
				case Param_type_grDir: {
					switch (this.val) {
						case 'bL':
							return ParameterVal_growDirection_bL;
						case 'bR':
							return ParameterVal_growDirection_bR;
						case 'tL':
							return ParameterVal_growDirection_tL;
						case 'tR':
							return ParameterVal_growDirection_tR;
						default:
							return this.val;
					}
				}
				case Param_type_contDir: {
					switch (this.val) {
						case 'sameDir':
							return ParameterVal_continueDirection_sameDir;
						case 'revDir':
							return ParameterVal_continueDirection_revDir;
						default:
							return this.val;
					}
				}
				case Param_type_secLinDir:
				case Param_type_linDir: {
					switch (this.val) {
						case 'fromL':
							return ParameterVal_linearDirection_fromL;
						case 'fromT':
							return ParameterVal_linearDirection_fromT;
						case 'fromR':
							return ParameterVal_linearDirection_fromR;
						case 'fromB':
							return ParameterVal_linearDirection_fromB;
						default:
							return this.val;
					}
				}
				case Param_type_hierAlign:
				{
					switch (this.val) {
						case 'bCtrCh':
							return ParameterVal_hierarchyAlignment_bCtrCh;
						case 'bCtrDes':
							return ParameterVal_hierarchyAlignment_bCtrDes;
						case 'bL':
							return ParameterVal_hierarchyAlignment_bL;
						case 'bR':
							return ParameterVal_hierarchyAlignment_bR;
						case 'lB':
							return ParameterVal_hierarchyAlignment_lB;
						case 'lCtrCh':
							return ParameterVal_hierarchyAlignment_lCtrCh;
						case 'lCtrDes':
							return ParameterVal_hierarchyAlignment_lCtrDes;
						case 'lT':
							return ParameterVal_hierarchyAlignment_lT;
						case 'rB':
							return ParameterVal_hierarchyAlignment_rB;
						case 'rCtrCh':
							return ParameterVal_hierarchyAlignment_rCtrCh;
						case 'rCtrDes':
							return ParameterVal_hierarchyAlignment_rCtrDes;
						case 'rT':
							return ParameterVal_hierarchyAlignment_rT;
						case 'tCtrCh':
							return ParameterVal_hierarchyAlignment_tCtrCh;
						case 'tCtrDes':
							return ParameterVal_hierarchyAlignment_tCtrDes;
						case 'tL':
							return ParameterVal_hierarchyAlignment_tL;
						case 'tR':
							return ParameterVal_hierarchyAlignment_tR;
						default:
							return this.val;
					}
				}
				case Param_type_secChAlign:
				case Param_type_chAlign: {
					switch (this.val) {
						case 'l':
							return ParameterVal_childAlignment_l;
						case 't':
							return ParameterVal_childAlignment_t;
						case 'r':
							return ParameterVal_childAlignment_r;
						case 'b':
							return ParameterVal_childAlignment_b;
						default:
							return this.val;
					}
				}
				case Param_type_flowDir: {
					switch (this.val) {
						case 'col':
							return ParameterVal_flowDirection_col;
						case 'row':
							return ParameterVal_flowDirection_row;
						default:
							return this.val;
					}
				}
				case Param_type_off:
					switch (this.val) {
						case 'ctr':
							return ParameterVal_offset_ctr;
						case 'off':
							return ParameterVal_offset_off;
						default:
							return this.val;
					}
				case Param_type_bkpt:
					switch (this.val) {
						case 'fixed':
							return ParameterVal_breakpoint_fixed;
						case 'endCnv':
							return ParameterVal_breakpoint_endCnv;
						case 'bal':
							return ParameterVal_breakpoint_bal;
						default:
							return this.val;
					}
				case Param_type_vertAlign:
					switch (this.val) {
						case 'mid':
							return ParameterVal_verticalAlignment_mid;
						case 'b':
							return ParameterVal_verticalAlignment_b;
						case 't':
							return ParameterVal_verticalAlignment_t;
						case 'none':
							return ParameterVal_verticalAlignment_none;
						default:
							return this.val;
					}
				case Param_type_horzAlign:
					switch (this.val) {
						case 'r':
							return ParameterVal_horizontalAlignment_r;
						case 'l':
							return ParameterVal_horizontalAlignment_l;
						case 'ctr':
							return ParameterVal_horizontalAlignment_ctr;
						case 'none':
							return ParameterVal_horizontalAlignment_none;
						default:
							return this.val;
					}
				case Param_type_begPts:
				case Param_type_endPts:
					const arrVal = this.val.split(' ');
					const arrResult = [];
					for (let i = 0; i < arrVal.length; i++) {
						const val = arrVal[i];
						switch (val) {
							case "auto":
								arrResult.push(ParameterVal_connectorPoint_auto);
								break;
							case "bCtr":
								arrResult.push(ParameterVal_connectorPoint_bCtr);
								break;
							case "bL":
								arrResult.push(ParameterVal_connectorPoint_bL);
								break;
							case "bR":
								arrResult.push(ParameterVal_connectorPoint_bR);
								break;
							case "ctr":
								arrResult.push(ParameterVal_connectorPoint_ctr);
								break;
							case "midL":
								arrResult.push(ParameterVal_connectorPoint_midL);
								break;
							case "midR":
								arrResult.push(ParameterVal_connectorPoint_midR);
								break;
							case "radial":
								arrResult.push(ParameterVal_connectorPoint_radial);
								break;
							case "tCtr":
								arrResult.push(ParameterVal_connectorPoint_tCtr);
								break;
							case "tL":
								arrResult.push(ParameterVal_connectorPoint_tL);
								break;
							case "tR":
								arrResult.push(ParameterVal_connectorPoint_tR);
								break;
							default:
								break;
						}
					}
					return arrResult;
				case Param_type_parTxLTRAlign:
				case Param_type_parTxRTLAlign:
				case Param_type_shpTxLTRAlignCh:
				case Param_type_shpTxRTLAlignCh:
					switch (this.val) {
						case "r":
							return ParameterVal_horizontalAlignment_r;
						case "l":
							return ParameterVal_horizontalAlignment_l;
						case "ctr":
							return ParameterVal_horizontalAlignment_ctr;
						default:
							return this.val;
					}
				case Param_type_txAnchorVert:
				case Param_type_txAnchorVertCh:
					switch (this.val) {
						case "b":
							return ParameterVal_textAnchorVertical_b;
						case "t":
							return ParameterVal_textAnchorVertical_t;
						case "mid":
							return ParameterVal_textAnchorVertical_mid;
						default:
							return this.val;
					}
				case Param_type_txAnchorHorz:
				case Param_type_txAnchorHorzCh:
					switch (this.val) {
						case "ctr":
							return true;
						default:
							return false;
					}
				case Param_type_lnSpAfParP:
				case Param_type_lnSpAfChP:
					return parseFloat(this.val) / 100 * 1.2 * g_dKoef_pt_to_mm;
				case Param_type_ar:
				case Param_type_spanAng:
				case Param_type_stAng:
				case Param_type_bkPtFixedVal:
				case Param_type_stBulletLvl:
					return parseFloat(this.val);
				default:
					return this.val;
			}
		};

    Param.prototype.setType = function (pr) {
      this.type = pr;
    }

    Param.prototype.setVal = function (oPr) {
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

    function Choose() {
      CBaseFormatNoIdObject.call(this);
      this.name = null;
      this.else = null;
      this.if = [];
    }

    InitClass(Choose, CBaseFormatNoIdObject, AscDFH.historyitem_type_Choose);
    Choose.prototype.Write_ToBinary = function (w) {
      AscFormat.writeString(w, this.name);
      AscFormat.writeObjectNoId(w, this.else);
      w.WriteLong(this.if.length);
      for (let i = 0; i < this.if.length; i++) {
        AscFormat.writeObjectNoId(w, this.if[i]);
      }
    };
    Choose.prototype.Read_FromBinary = function (r) {
      this.name = AscFormat.readString(r);
      this.else = AscFormat.readObjectNoId(r);
      for (let i = r.GetLong(); i >  0;  i -= 1) {
        this.if.push(AscFormat.readObjectNoId(r));
      }
    };
    Choose.prototype.setName = function (pr) {
      this.name = pr;
    };

    Choose.prototype.setElse = function (oPr) {
      this.else = oPr;
    };

    Choose.prototype.addToLstIf = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.if.length, Math.max(0, nIdx));
      nInsertIdx === this.if.length ? this.if.push(oPr) : this.if.splice(nInsertIdx, 0, oPr);
    };

    Choose.prototype.removeFromLstIf = function (nIdx) {
      if (nIdx > -1 && nIdx < this.if.length) {
        nIdx === this.if.length - 1 ? this.if.pop() : this.if.splice(nIdx, 1);
      }
    };

    Choose.prototype.getName = function () {
      return this.name;
    };

    Choose.prototype.getElse = function () {
      return this.else;
    };

    Choose.prototype.getIf = function () {
      return this.if;
    };

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
          var ifObj = new If();
          this.addToLstIf(this.if.length, ifObj);
          ifObj.fromPPTY(pReader);
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
      return [this.else].concat(this.if);
    };

    function Else() {
      LayoutBaseClass.call(this);
      this.name = null;
    }

    InitClass(Else, LayoutBaseClass, AscDFH.historyitem_type_Else);
    Else.prototype.Write_ToBinary = function (w) {
			LayoutBaseClass.prototype.Write_ToBinary.call(this, w);
      AscFormat.writeString(w, this.name);
    };
    Else.prototype.Read_FromBinary = function (r) {
			LayoutBaseClass.prototype.Read_FromBinary.call(this, r);
      this.name = AscFormat.readString(r);
    };
    Else.prototype.setName = function (pr) {
      this.name = pr;
    }
    Else.prototype.getName = function () {
      return this.name;
    }
    Else.prototype.fillObject = function (oCopy, oIdMap) {
      LayoutBaseClass.prototype.fillObject.call(this, oCopy, oIdMap);
      oCopy.setName(this.getName());
    };

    Else.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
    };

    Else.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
    };

    Else.prototype.readChild = function(nType, pReader) {
      this.readElement(pReader, nType);
    };
    Else.prototype.getChildren = function() {
      return [this.alg, this.shape, this.presOf, this.constrLst, this.ruleLst, this.varLst].concat(this.list);
    };

    function IteratorAttributes() {
      AscFormat.CBaseFormatNoIdObject.call(this);
      this.axis = [];
      this.cnt = [];
      this.hideLastTrans = [];
      this.ptType = [];
      this.st = [];
      this.step = [];
    }

    InitClass(IteratorAttributes, AscFormat.CBaseFormatNoIdObject, AscDFH.historyitem_type_IteratorAttributes);
		IteratorAttributes.prototype.Write_ToBinary = function (w) {
			w.WriteLong(this.axis.length);
			for (let i = 0; i < this.axis.length; i += 1) {
				w.WriteLong(this.axis[i]);
			}
			w.WriteLong(this.cnt.length);
			for (let i = 0; i < this.cnt.length; i += 1) {
				w.WriteLong(this.cnt[i]);
			}
			w.WriteLong(this.hideLastTrans.length);
			for (let i = 0; i < this.hideLastTrans.length; i += 1) {
				w.WriteBool(this.hideLastTrans[i]);
			}
			w.WriteLong(this.ptType.length);
			for (let i = 0; i < this.ptType.length; i += 1) {
				w.WriteLong(this.ptType[i]);
			}
			w.WriteLong(this.st.length);
			for (let i = 0; i < this.st.length; i += 1) {
				w.WriteLong(this.st[i]);
			}
			w.WriteLong(this.step.length);
			for (let i = 0; i < this.step.length; i += 1) {
				w.WriteLong(this.step[i]);
			}
		};
	  IteratorAttributes.prototype.Read_FromBinary = function (r) {
		  for (let i = r.GetLong(); i > 0; i -= 1) {
			  this.axis.push(r.GetLong());
		  }
		  for (let i = r.GetLong(); i > 0; i -= 1) {
			  this.cnt.push(r.GetLong());
		  }
		  for (let i = r.GetLong(); i > 0; i -= 1) {
			  this.hideLastTrans.push(r.GetBool());
		  }
		  for (let i = r.GetLong(); i > 0; i -= 1) {
			  this.ptType.push(r.GetLong());
		  }
		  for (let i = r.GetLong(); i > 0; i -= 1) {
			  this.st.push(r.GetLong());
		  }
		  for (let i = r.GetLong(); i > 0; i -= 1) {
			  this.step.push(r.GetLong());
		  }
	  };
	  IteratorAttributes.prototype.getNodesArray = function (smartartAlgorithm) {
		  if (!this.axis.length) {
			  return [];
		  }
		  const currentNode = smartartAlgorithm.getCurrentNode();
		  let currentNodes = [currentNode];
		  for (let i = 0; i < this.axis.length; i += 1) {
			  const newCurrentNodes = [];
			  for (let j = 0; j < currentNodes.length; j += 1) {
				  const node = currentNodes[j];
				  const tempNodes = [];
				  node.getNodesByAxis(tempNodes, this.getAxis(i), this.getPtType(i));
				  const step = this.getStep(i);
					let count = this.getCount(i) || tempNodes.length;
				  for (let k = this.getStart(i, tempNodes.length); k < tempNodes.length; k += step) {
					  if (!count) {
						  break;
					  }
					  newCurrentNodes.push(tempNodes[k]);
						count -= 1;
				  }
			  }
			  currentNodes = newCurrentNodes;
		  }
			const lastNode = currentNodes[currentNodes.length - 1];
			if (lastNode && lastNode.isHideLastTrans) {
				currentNodes.pop();
			}
		  return currentNodes;
	  };
	  IteratorAttributes.prototype.getPtType = function (index) {
		  if (AscFormat.isRealNumber(this.ptType[index])) {
				return this.ptType[index];
		  }
			return ElementType_value_all;
	  };
	  IteratorAttributes.prototype.getCount = function (index) {
			return this.cnt[index] || 0;
	  };
	  IteratorAttributes.prototype.getAxis = function (index) {
			if (AscFormat.isRealNumber(this.axis[index])) {
			return this.axis[index];
			}
			return AxisType_value_none;
	  };
	  IteratorAttributes.prototype.getStart = function (index, nodesLength) {
		  if (AscFormat.isRealNumber(this.st[index])) {
				const start = this.st[index];
				if (start <= 0) {
					return Math.max(nodesLength + start, 0);
				}
			  return start - 1;
		  }
		  return 0;
	  };
	  IteratorAttributes.prototype.getStep = function (index) {
		  if (AscFormat.isRealNumber(this.step[index])) {
			  return this.st[index];
		  }
		  return 1;
	  };
	  IteratorAttributes.prototype.getHideLastTrans = function (index) {
		  if (this.hideLastTrans[index] !== undefined) {
			  return this.hideLastTrans[index];
		  }
		  return true;
	  };
    IteratorAttributes.prototype.addToLstAxis = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.axis.length, Math.max(0, nIdx));
      nInsertIdx === this.axis.length ? this.axis.push(oPr) : this.axis.splice(nInsertIdx, 0, oPr);
    };

    IteratorAttributes.prototype.removeFromLstAxis = function (nIdx) {
      if (nIdx > -1 && nIdx < this.axis.length) {
        nIdx === this.axis.length - 1 ? this.axis.pop() : this.axis.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstCnt = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.cnt.length, Math.max(0, nIdx));
      nInsertIdx === this.cnt.length ? this.cnt.push(oPr) : this.cnt.splice(nInsertIdx, 0, oPr);
    };

    IteratorAttributes.prototype.removeFromLstCnt = function (nIdx) {
      if (nIdx > -1 && nIdx < this.cnt.length) {
        nIdx === this.cnt.length - 1 ? this.cnt.pop() : this.cnt.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstHideLastTrans = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.hideLastTrans.length, Math.max(0, nIdx));
      nInsertIdx === this.hideLastTrans.length ? this.hideLastTrans.push(oPr) : this.hideLastTrans.splice(nInsertIdx, 0, oPr);
    };

    IteratorAttributes.prototype.removeFromLstHideLastTrans = function (nIdx) {
      if (nIdx > -1 && nIdx < this.hideLastTrans.length) {
        nIdx === this.hideLastTrans.length - 1 ? this.hideLastTrans.pop() : this.hideLastTrans.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstPtType = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.ptType.length, Math.max(0, nIdx));
      nInsertIdx === this.ptType.length ? this.ptType.push(oPr) : this.ptType.splice(nInsertIdx, 0, oPr);
    };

    IteratorAttributes.prototype.removeFromLstPtType = function (nIdx) {
      if (nIdx > -1 && nIdx < this.ptType.length) {
        nIdx === this.ptType.length - 1 ? this.ptType.pop() : this.ptType.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstSt = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.st.length, Math.max(0, nIdx));
      nInsertIdx === this.st.length ? this.st.push(oPr) : this.st.splice(nInsertIdx, 0, oPr);
    };

    IteratorAttributes.prototype.removeFromLstSt = function (nIdx) {
      if (nIdx > -1 && nIdx < this.st.length) {
        nIdx === this.st.length - 1 ? this.st.pop() : this.st.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.addToLstStep = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.step.length, Math.max(0, nIdx));
      nInsertIdx === this.step.length ? this.step.push(oPr) : this.step.splice(nInsertIdx, 0, oPr);
    };

    IteratorAttributes.prototype.removeFromLstStep = function (nIdx) {
      if (nIdx > -1 && nIdx < this.step.length) {
        nIdx === this.step.length - 1 ? this.step.pop() : this.step.splice(nIdx, 1);
      }
    };

    IteratorAttributes.prototype.fillObject = function (oCopy, oIdMap) {
      for (var nIdx = 0; nIdx < this.axis.length; ++nIdx) {
        oCopy.addToLstAxis(nIdx, this.axis[nIdx]);
      }
      for (nIdx = 0; nIdx < this.cnt.length; ++nIdx) {
        oCopy.addToLstCnt(nIdx, this.cnt[nIdx]);
      }
      for (nIdx = 0; nIdx < this.hideLastTrans.length; ++nIdx) {
        oCopy.addToLstHideLastTrans(nIdx, this.hideLastTrans[nIdx]);
      }
      for (nIdx = 0; nIdx < this.ptType.length; ++nIdx) {
        oCopy.addToLstPtType(nIdx, this.ptType[nIdx]);
      }
      for (nIdx = 0; nIdx < this.st.length; ++nIdx) {
        oCopy.addToLstSt(nIdx, this.st[nIdx]);
      }
      for (nIdx = 0; nIdx < this.step.length; ++nIdx) {
        oCopy.addToLstStep(nIdx, this.step[nIdx]);
      }
    };

		function IteratorLayoutBase() {
			IteratorAttributes.call(this);
			this.list = [];
			this.alg = null;
			this.shape = null;
			this.presOf = null;
			this.constrLst = null;
			this.ruleLst = null;
			this.varLst = null;
		}
		AscFormat.InitClass(IteratorLayoutBase, IteratorAttributes, 0);
	  IteratorLayoutBase.prototype.fillObject = function (oCopy, oIdMap) {
		  IteratorAttributes.prototype.fillObject.call(this, oCopy, oIdMap);
		  LayoutBaseClass.prototype.fillObject.call(this, oCopy, oIdMap);
	  };
	  IteratorLayoutBase.prototype.Write_ToBinary = function (w) {
		  IteratorAttributes.prototype.Write_ToBinary.call(this, w);
		  LayoutBaseClass.prototype.Write_ToBinary.call(this, w);
	  };
	  IteratorLayoutBase.prototype.Read_FromBinary = function (r) {
		  IteratorAttributes.prototype.Read_FromBinary.call(this, r);
		  LayoutBaseClass.prototype.Read_FromBinary.call(this, r);
	  };
	  IteratorLayoutBase.prototype.setAlg = LayoutBaseClass.prototype.setAlg;
	  IteratorLayoutBase.prototype.setShape = LayoutBaseClass.prototype.setShape;
	  IteratorLayoutBase.prototype.setPresOf = LayoutBaseClass.prototype.setPresOf;
	  IteratorLayoutBase.prototype.setConstrLst = LayoutBaseClass.prototype.setConstrLst;
	  IteratorLayoutBase.prototype.setRuleLst = LayoutBaseClass.prototype.setRuleLst;
	  IteratorLayoutBase.prototype.setVarLst = LayoutBaseClass.prototype.setVarLst;
	  IteratorLayoutBase.prototype.addToLst = LayoutBaseClass.prototype.addToLst;
	  IteratorLayoutBase.prototype.removeFromLst = LayoutBaseClass.prototype.removeFromLst;
    IteratorLayoutBase.prototype.readElement = LayoutBaseClass.prototype.readElement;
    IteratorLayoutBase.prototype.writeChildren = LayoutBaseClass.prototype.writeChildren;

    function If() {
      IteratorLayoutBase.call(this);

      this.arg = If_arg_none;
      this.func = null;
      this.name = null;
      this.op = null;
      this.val = null;
      this.ref = null;
    }

    InitClass(If, IteratorLayoutBase, AscDFH.historyitem_type_If);
    If.prototype.Write_ToBinary = function (w) {
      IteratorLayoutBase.prototype.Write_ToBinary.call(this, w);
      AscFormat.writeLong(w, this.arg);
      AscFormat.writeString(w, this.ref);
      AscFormat.writeLong(w, this.func);
      AscFormat.writeString(w, this.name);
      AscFormat.writeLong(w, this.op);
      AscFormat.writeString(w, this.val);
    }
    If.prototype.Read_FromBinary = function (r) {
      IteratorLayoutBase.prototype.Read_FromBinary.call(this, r);
      this.arg = AscFormat.readLong(r);
      this.ref = AscFormat.readString(r);
      this.func = AscFormat.readLong(r);
      this.name = AscFormat.readString(r);
      this.op = AscFormat.readLong(r);
      this.val = AscFormat.readString(r);
    };

		If.getArgEnum = function (arg) {
			switch (arg) {
				case 'animLvl':
					return If_arg_animLvl;
				case 'animOne':
					return If_arg_animOne;
				case 'bulEnabled':
					return If_arg_bulEnabled;
				case 'chMax':
					return If_arg_chMax;
				case 'chPref':
					return If_arg_chPref;
				case 'dir':
					return If_arg_dir;
				case 'hierBranch':
					return If_arg_hierBranch;
				case 'orgChart':
					return If_arg_orgChart;
				case 'resizeHandles':
					return If_arg_resizeHandles;
				case 'none':
				default:
					return If_arg_none;
			}
		};
		If.getArgString = function (arg) {
			switch (arg) {
				case If_arg_animLvl:
					return 'animLvl';
				case If_arg_animOne:
					return 'animOne';
				case If_arg_bulEnabled:
					return 'bulEnabled';
				case If_arg_chMax:
					return 'chMax';
				case If_arg_chPref:
					return 'chPref';
				case If_arg_dir:
					return 'dir';
				case If_arg_hierBranch:
					return 'hierBranch';
				case If_arg_none:
					return 'none';
				case If_arg_orgChart:
					return 'orgChart';
				case If_arg_resizeHandles:
					return 'resizeHandles';

			}
		}
    If.prototype.setArg = function (pr) {
      this.arg = pr;
    };

    If.prototype.setFunc = function (pr) {
      this.func = pr;
    };

    If.prototype.setRef = function (pr) {
      this.ref = pr;
    };

    If.prototype.setName = function (pr) {
      this.name = pr;
    };

    If.prototype.setOp = function (pr) {
      this.op = pr;
    };

    If.prototype.setVal = function (oPr) {
      this.val = oPr;
    };

    If.prototype.getArg = function () {
      return this.arg;
    };

    If.prototype.getRef = function () {
      return this.ref;
    };

    If.prototype.getFunc = function () {
      return this.func;
    };

    If.prototype.getName = function () {
      return this.name;
    };

    If.prototype.getOp = function () {
      return this.op;
    };

    If.prototype.getVal = function () {
      return this.val;
    };

    If.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setArg(this.getArg());
      oCopy.setFunc(this.getFunc());
      oCopy.setName(this.getName());
      oCopy.setOp(this.getOp());
      oCopy.setVal(this.getVal());
      IteratorLayoutBase.prototype.fillObject.call(this, oCopy, oIdMap);
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
        pWriter._WriteUChar1(5, this.axis[i]);
      }
      for (i = 0; i < this.ptType.length; i += 1) {
        pWriter._WriteUChar1(6, this.ptType[i]);
      }
      pWriter._WriteString2(7, this.ref);
      pWriter._WriteUChar2(8, this.op);
      pWriter._WriteUChar2(9, this.func);
      pWriter._WriteString2(10, this.val);
      pWriter._WriteString2(11, If.getArgString(this.arg));
    };
    If.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
      else if (1 === nType) this.addToLstSt(this.st.length, oStream.GetLong());
      else if (2 === nType) this.addToLstStep(this.step.length, oStream.GetLong());
      else if (3 === nType) this.addToLstHideLastTrans(this.hideLastTrans.length, oStream.GetBool());
      else if (4 === nType) this.addToLstCnt(this.cnt.length, oStream.GetLong());
      else if (5 === nType) this.addToLstAxis(this.axis.length, oStream.GetUChar());
      else if (6 === nType) this.addToLstPtType(this.ptType.length, oStream.GetUChar());
      else if (7 === nType) this.setRef(oStream.GetString2());
      else if (8 === nType) this.setOp(oStream.GetUChar());
      else if (9 === nType) this.setFunc(oStream.GetUChar());
      else if (10 === nType) this.setVal(oStream.GetString2());
      else if (11 === nType) this.setArg(If.getArgEnum(oStream.GetString2()));
    };

    If.prototype.readChild = function(nType, pReader) {
      this.readElement(pReader, nType);
    };
    If.prototype.getChildren = function() {
      return [this.alg, this.shape, this.presOf, this.constrLst, this.ruleLst, this.varLst].concat(this.list);
    };

    function ConstrLst() {
      CCommonDataListNoId.call(this);
    }

    InitClass(ConstrLst, CCommonDataListNoId, AscDFH.historyitem_type_ConstrLst);

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

    function Constr() {
      CBaseFormatNoIdObject.call(this);
      this.fact = 1;
      this.for = Constr_for_self;
      this.forName = null;
      this.op = Constr_op_none;
			this.ptType = ElementType_value_all;
      this.refFor = Constr_for_self;
      this.refForName = null;
	    this.refPtType = ElementType_value_all;
      this.refType = Constr_type_none;
      this.type = null;
      this.val = 0;
    }

    InitClass(Constr, CBaseFormatNoIdObject, AscDFH.historyitem_type_Constr);

		Constr.prototype.Write_ToBinary = function (w) {
			AscFormat.writeDouble2(w, this.fact);
			AscFormat.writeLong(w, this.for);
			AscFormat.writeString(w, this.forName);
			AscFormat.writeLong(w, this.op);
			AscFormat.writeLong(w, this.ptType);
			AscFormat.writeLong(w, this.refFor);
			AscFormat.writeString(w, this.refForName);
			AscFormat.writeLong(w, this.refPtType);
			AscFormat.writeLong(w, this.refType);
			AscFormat.writeLong(w, this.type);
			AscFormat.writeDouble2(w, this.val);
		}
	  Constr.prototype.Read_FromBinary = function (r) {
		  this.fact = AscFormat.readDouble2(r);
		  this.for = AscFormat.readLong(r);
		  this.forName = AscFormat.readString(r);
		  this.op = AscFormat.readLong(r);
		  this.ptType = AscFormat.readLong(r);
		  this.refFor = AscFormat.readLong(r);
		  this.refForName = AscFormat.readString(r);
		  this.refPtType = AscFormat.readLong(r);
		  this.refType = AscFormat.readLong(r);
		  this.type = AscFormat.readLong(r);
		  this.val = AscFormat.readDouble2(r);
	  }
    Constr.prototype.setFact = function (pr) {
      this.fact = pr;
    };

    Constr.prototype.setFor = function (pr) {
      this.for = pr;
    };

    Constr.prototype.setForName = function (pr) {
      this.forName = pr;
    };

    Constr.prototype.setOp = function (pr) {
      this.op = pr;
    };

    Constr.prototype.setPtType = function (oPr) {
      this.ptType = oPr;
    };

    Constr.prototype.setRefFor = function (pr) {
      this.refFor = pr;
    };

    Constr.prototype.setRefForName = function (pr) {
      this.refForName = pr;
    };

    Constr.prototype.setRefPtType = function (oPr) {
      this.refPtType = oPr;
    };

    Constr.prototype.setRefType = function (pr) {
      this.refType = pr;
    };

    Constr.prototype.setType = function (pr) {
      this.type = pr;
    };

    Constr.prototype.setVal = function (pr) {
      this.val = pr;
    };

    Constr.prototype.getFact = function () {
      return this.fact;
    };

    Constr.prototype.getFor = function () {
      return this.for;
    };

    Constr.prototype.getForName = function () {
      return this.forName;
    };

    Constr.prototype.getOp = function () {
      return this.op;
    };

    Constr.prototype.getPtType = function () {
      return this.ptType;
    };

    Constr.prototype.getRefFor = function () {
      return this.refFor;
    };

    Constr.prototype.getRefForName = function () {
      return this.refForName;
    };

    Constr.prototype.getRefPtType = function () {
      return this.refPtType;
    };

    Constr.prototype.getRefType = function () {
      return this.refType;
    };

    Constr.prototype.getType = function () {
      return this.type;
    };

    Constr.prototype.getVal = function () {
      return this.val;
    };

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
      oCopy.setPtType(this.getPtType());
      oCopy.setRefPtType(this.getRefPtType());
    };

    Constr.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteDoubleReal2(0, this.fact);
      pWriter._WriteUChar2(1, this.for);
      pWriter._WriteString2(2, this.forName);
      pWriter._WriteUChar2(3, this.op);
      pWriter._WriteUChar2(4, this.ptType);
      pWriter._WriteUChar2(5, this.refFor);
      pWriter._WriteString2(6, this.refForName);
      pWriter._WriteUChar2(7, this.refPtType);
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
      else if (4 === nType) this.setPtType(oStream.GetUChar());
      else if (5 === nType) this.setRefFor(oStream.GetUChar());
      else if (6 === nType) this.setRefForName(oStream.GetString2());
      else if (7 === nType) this.setRefPtType(oStream.GetUChar());
      else if (8 === nType) this.setRefType(oStream.GetUChar());
      else if (9 === nType) this.setType(oStream.GetUChar());
      else if (10 === nType) this.setVal(oStream.GetDouble());
    };
    Constr.prototype.readChild = function(nType, pReader) {
    };
    function PresOf() {
      IteratorAttributes.call(this);
    }

    InitClass(PresOf, IteratorAttributes, AscDFH.historyitem_type_PresOf);

    PresOf.prototype.fillObject = function (oCopy, oIdMap) {
      for (var nIdx = 0; nIdx < this.axis.length; ++nIdx) {
        oCopy.addToLstAxis(nIdx, this.axis[nIdx]);
      }
      for (nIdx = 0; nIdx < this.cnt.length; ++nIdx) {
        oCopy.addToLstCnt(nIdx, this.cnt[nIdx]);
      }
      for (nIdx = 0; nIdx < this.hideLastTrans.length; ++nIdx) {
        oCopy.addToLstHideLastTrans(nIdx, this.hideLastTrans[nIdx]);
      }
      for (nIdx = 0; nIdx < this.ptType.length; ++nIdx) {
        oCopy.addToLstPtType(nIdx, this.ptType[nIdx]);
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
        pWriter._WriteUChar1(5, this.axis[i]);
      }
      for (i = 0; i < this.ptType.length; i += 1) {
        pWriter._WriteUChar1(6, this.ptType[i]);
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
      else if (5 === nType) this.addToLstAxis(this.axis.length, oStream.GetUChar());
      else if (6 === nType) this.addToLstPtType(this.ptType.length, oStream.GetUChar());
    };

    PresOf.prototype.readChild = function(nType, pReader) {
    };

    function RuleLst() {
      CCommonDataListNoId.call(this);
    }

    InitClass(RuleLst, CCommonDataListNoId, AscDFH.historyitem_type_RuleLst);

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

    function Rule() {
      CBaseFormatNoIdObject.call(this);
      this.fact = null;
      this.for = AscFormat.Constr_for_self;
      this.forName = null;
      this.max = null;
      this.type = null;
      this.val = null;
	    this.ptType = ElementType_value_all;
    }

    InitClass(Rule, CBaseFormatNoIdObject, AscDFH.historyitem_type_Rule);
	  Rule.prototype.Write_ToBinary = function (w) {
		  AscFormat.writeDouble2(w, this.fact);
		  AscFormat.writeLong(w, this.for);
			AscFormat.writeString(w, this.forName);
			AscFormat.writeDouble2(w, this.max);
			AscFormat.writeLong(w, this.type);
			AscFormat.writeDouble2(w, this.val);
			AscFormat.writeLong(w, this.ptType);
	  }
	  Rule.prototype.Read_FromBinary = function (r) {
		  this.setFact(AscFormat.readDouble2(r));
		  this.setFor(AscFormat.readLong(r));
		  this.setForName(AscFormat.readString(r));
		  this.setMax(AscFormat.readDouble2(r));
		  this.setType(AscFormat.readLong(r));
		  this.setVal(AscFormat.readDouble2(r));
		  this.setPtType(AscFormat.readLong(r));
	  }
    Rule.prototype.setFact = function (pr) {
      this.fact = pr;
    }

    Rule.prototype.setFor = function (pr) {
      this.for = pr;
    }

    Rule.prototype.setForName = function (pr) {
      this.forName = pr;
    }

    Rule.prototype.setMax = function (pr) {
      this.max = pr;
    }

    Rule.prototype.setType = function (pr) {
      this.type = pr;
    }

    Rule.prototype.setVal = function (pr) {
      this.val = pr;
    }

    Rule.prototype.setPtType = function (oPr) {
      this.ptType = oPr;
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
      oCopy.setPtType(this.getPtType());
    }

    Rule.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteDoubleReal2(0, this.fact);
      pWriter._WriteUChar2(1, this.for);
      pWriter._WriteString2(2, this.forName);
      pWriter._WriteUChar2(3, this.ptType);
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
      else if (3 === nType) this.setPtType(oStream.GetUChar());
      else if (4 === nType) this.setType(oStream.GetUChar());
      else if (5 === nType) this.setVal(oStream.GetDouble());
      else if (6 === nType) this.setMax(oStream.GetDouble());
    };
    Rule.prototype.readChild = function(nType, pReader) {
    };


		function SShape() {
      CBaseFormatNoIdObject.call(this);
      this.blip = null;
      this.blipPhldr = false;
      this.hideGeom = false;
      this.lkTxEntry = false;
      this.rot = 0;
      this.type = LayoutShapeType_outputShapeType_none;
      this.zOrderOff = 0;
      this.adjLst = null;
    }

    InitClass(SShape, CBaseFormatNoIdObject, AscDFH.historyitem_type_SShape);

		SShape.prototype.Write_ToBinary = function (w) {
			AscFormat.writeString(w, this.blip);
			AscFormat.writeBool(w, this.blipPhldr);
			AscFormat.writeBool(w, this.hideGeom);
			AscFormat.writeBool(w, this.lkTxEntry);
			AscFormat.writeDouble2(w, this.rot);
			AscFormat.writeLong(w, this.type);
			AscFormat.writeLong(w, this.zOrderOff);
			AscFormat.writeObjectNoId(w, this.adjLst);
		};
	  SShape.prototype.Read_FromBinary = function (r) {
		  this.blip = AscFormat.readString(r);
		  this.blipPhldr = AscFormat.readBool(r);
		  this.hideGeom = AscFormat.readBool(r);
		  this.lkTxEntry = AscFormat.readBool(r);
		  this.rot = AscFormat.readDouble2(r);
		  this.type = AscFormat.readLong(r);
		  this.zOrderOff = AscFormat.readLong(r);
		  this.adjLst = AscFormat.readObjectNoId(r);
	  };
    SShape.prototype.setBlip = function (pr) {
      this.blip = pr;
    }

    SShape.prototype.setBlipPhldr = function (pr) {
      this.blipPhldr = pr;
    }

    SShape.prototype.setHideGeom = function (pr) {
      this.hideGeom = pr;
    }

    SShape.prototype.setLkTxEntry = function (pr) {
      this.lkTxEntry = pr;
    }

    SShape.prototype.setRot = function (pr) {
      this.rot = pr;
    }

    SShape.prototype.setType = function (oPr) {
      this.type = oPr;
    }

    SShape.prototype.setZOrderOff = function (pr) {
      this.zOrderOff = pr;
    }

    SShape.prototype.setAdjLst = function (oPr) {
      this.adjLst = oPr;
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
    }
    SShape.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.blip);
      pWriter._WriteBool2(1, this.blipPhldr);
      pWriter._WriteBool2(2, this.hideGeom);
      pWriter._WriteBool2(3, this.lkTxEntry);
      pWriter._WriteDoubleReal2(4, this.rot);
      pWriter._WriteInt2(5, this.zOrderOff);
      pWriter._WriteString2(6, AscCommon.To_XML_ST_LayoutShapeType(this.type));

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
      else if (5 === nType) this.setZOrderOff(oStream.GetLong());
      else if (6 === nType) this.setType(AscCommon.From_XML_ST_LayoutShapeType(oStream.GetString2()));
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
      CCommonDataListNoId.call(this);
    }

    InitClass(AdjLst, CCommonDataListNoId, AscDFH.historyitem_type_AdjLst);

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

    function Adj() {
      CBaseFormatNoIdObject.call(this);
      this.idx = null;
      this.val = null;
    }

    InitClass(Adj, CBaseFormatNoIdObject, AscDFH.historyitem_type_Adj);

	  Adj.prototype.Write_ToBinary = function (w) {
		  AscFormat.writeLong(w, this.idx);
		  AscFormat.writeDouble2(w, this.val);
	  };
	  Adj.prototype.Read_FromBinary = function (r) {
		  this.idx = AscFormat.readLong(r);
		  this.val = AscFormat.readDouble2(r);
	  };
    Adj.prototype.setIdx = function (pr) {
      this.idx = pr;
    }

    Adj.prototype.setVal = function (pr) {
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

    function VarLst() {
      CBaseFormatNoIdObject.call(this);
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

    InitClass(VarLst, CBaseFormatNoIdObject, AscDFH.historyitem_type_VarLst);

	  VarLst.prototype.Write_ToBinary = function (w) {
		  AscFormat.writeLong(w, this.animLvl);
		  AscFormat.writeLong(w, this.animOne);
		  AscFormat.writeBool(w, this.bulletEnabled);
		  AscFormat.writeLong(w, this.chMax);
		  AscFormat.writeLong(w, this.chPref);
		  AscFormat.writeLong(w, this.dir);
		  AscFormat.writeLong(w, this.hierBranch);
		  AscFormat.writeBool(w, this.orgChart);
		  AscFormat.writeLong(w, this.resizeHandles);
	  };
	  VarLst.prototype.Read_FromBinary = function (r) {
		  this.animLvl = AscFormat.readLong(r);
		  this.animOne = AscFormat.readLong(r);
		  this.bulletEnabled = AscFormat.readBool(r);
		  this.chMax = AscFormat.readLong(r);
		  this.chPref = AscFormat.readLong(r);
		  this.dir = AscFormat.readLong(r);
		  this.hierBranch = AscFormat.readLong(r);
		  this.orgChart = AscFormat.readBool(r);
		  this.resizeHandles = AscFormat.readLong(r);
	  };
    VarLst.prototype.setAnimLvl = function (oPr) {
      this.animLvl = oPr;
    }

    VarLst.prototype.setAnimOne = function (oPr) {
      this.animOne = oPr;
    }

    VarLst.prototype.setBulletEnabled = function (oPr) {
      this.bulletEnabled = oPr;
    }

    VarLst.prototype.setChMax = function (oPr) {
      this.chMax = oPr;
    }

    VarLst.prototype.setChPref = function (oPr) {
      this.chPref = oPr;
    }

    VarLst.prototype.setDir = function (oPr) {
      this.dir = oPr;
    }

    VarLst.prototype.setHierBranch = function (oPr) {
      this.hierBranch = oPr;
    }

    VarLst.prototype.setOrgChart = function (oPr) {
      this.orgChart = oPr;
    }

    VarLst.prototype.setResizeHandles = function (oPr) {
      this.resizeHandles = oPr;
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
      oCopy.setAnimLvl(this.getAnimLvl());
      oCopy.setAnimOne(this.getAnimOne());
      oCopy.setBulletEnabled(this.getBulletEnabled());
      oCopy.setChMax(this.getChMax());
      oCopy.setChPref(this.getChPref());
      oCopy.setDir(this.getDir());
      oCopy.setHierBranch(this.getHierBranch());
      oCopy.setOrgChart(this.getOrgChart());
      oCopy.setResizeHandles(this.getResizeHandles());
    }

    VarLst.prototype.privateWriteAttributes = null;
    VarLst.prototype.writeChildren = function(pWriter) {
      if (this.animLvl !== null) pWriter.WriteRecord1(0, this.animLvl, pWriter.WriteByteToPPTY.bind(pWriter));
      if (this.animOne !== null) pWriter.WriteRecord1(1, this.animOne, pWriter.WriteByteToPPTY.bind(pWriter));
      if (this.bulletEnabled !== null) pWriter.WriteRecord1(2, Number(this.bulletEnabled), pWriter.WriteByteToPPTY.bind(pWriter));
      if (this.chMax !== null) pWriter.WriteRecord1(3, this.chMax, pWriter.WriteIntToPPTY.bind(pWriter));
      if (this.chPref !== null) pWriter.WriteRecord1(4, this.chPref, pWriter.WriteIntToPPTY.bind(pWriter));
      if (this.dir !== null) pWriter.WriteRecord1(5, this.dir, pWriter.WriteByteToPPTY.bind(pWriter));
      if (this.hierBranch !== null) pWriter.WriteRecord1(6, this.hierBranch, pWriter.WriteByteToPPTY.bind(pWriter));
      if (this.orgChart !== null) pWriter.WriteRecord1(7, Number(this.orgChart), pWriter.WriteByteToPPTY.bind(pWriter));
      if (this.resizeHandles !== null) pWriter.WriteRecord1(8, this.resizeHandles, pWriter.WriteByteToPPTY.bind(pWriter));
    };
    VarLst.prototype.readAttribute = null;
    VarLst.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setAnimLvl(s.ReadByteFromPPTY());
          break;
        }
        case 1: {
          this.setAnimOne(s.ReadByteFromPPTY());
          break;
        }
        case 2: {
          this.setBulletEnabled(!!s.ReadByteFromPPTY());
          break;
        }
        case 3: {
          this.setChMax(s.ReadIntFromPPTY());
          break;
        }
        case 4: {
          this.setChPref(s.ReadIntFromPPTY());
          break;
        }
        case 5: {
          this.setDir(s.ReadByteFromPPTY());
          break;
        }
        case 6: {
          this.setHierBranch(s.ReadByteFromPPTY());
          break;
        }
        case 7: {
          this.setOrgChart(!!s.ReadByteFromPPTY());
          break;
        }
        case 8: {
          this.setResizeHandles(s.ReadByteFromPPTY());
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };

    VarLst.prototype.getChildren = function() {
      return [this.animLvl, this.animOne, this.bulletEnabled, this.chMax, this.chPref, this.dir, this.hierBranch, this.orgChart, this.resizeHandles];
    };

    function ForEach() {
      IteratorLayoutBase.call(this);
      this.name = null;
      this.ref = null;
    }

    InitClass(ForEach, IteratorLayoutBase, AscDFH.historyitem_type_ForEach);

    ForEach.prototype.Write_ToBinary = function (w) {
      IteratorLayoutBase.prototype.Write_ToBinary.call(this, w);
      AscFormat.writeString(w, this.ref);
      AscFormat.writeString(w, this.name);
    };
    ForEach.prototype.Read_FromBinary = function (r) {
      IteratorLayoutBase.prototype.Read_FromBinary.call(this, r);
      this.ref = AscFormat.readString(r);
      this.name = AscFormat.readString(r);
    };
    ForEach.prototype.setName = function (pr) {
      this.name = pr;
    }

    ForEach.prototype.setRef = function (pr) {
      this.ref = pr;
    }

    ForEach.prototype.getName = function () {
      return this.name;
    }

    ForEach.prototype.getRef = function () {
      return this.ref;
    }

    ForEach.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setName(this.getName());
      oCopy.setRef(this.getRef());
      IteratorLayoutBase.prototype.fillObject.call(this, oCopy, oIdMap);
    }

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
        pWriter._WriteUChar1(5, this.axis[i]);
      }
      for (i = 0; i < this.ptType.length; i += 1) {
        pWriter._WriteUChar1(6, this.ptType[i]);
      }
      pWriter._WriteString2(7, this.ref);
    };
    ForEach.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      if (0 === nType) this.setName(oStream.GetString2());
      else if (1 === nType) this.addToLstSt(this.st.length, oStream.GetLong());
      else if (2 === nType) this.addToLstStep(this.step.length, oStream.GetLong());
      else if (3 === nType) this.addToLstHideLastTrans(this.hideLastTrans.length, oStream.GetBool());
      else if (4 === nType) this.addToLstCnt(this.cnt.length, oStream.GetLong());
      else if (5 === nType) {
        this.addToLstAxis(this.axis.length, oStream.GetUChar());
      }
      else if (6 === nType) {
        this.addToLstPtType(this.ptType.length, oStream.GetUChar());
      }
      else if (7 === nType) this.setRef(oStream.GetString2());
    };

    ForEach.prototype.readChild = function(nType, pReader) {
      this.readElement(pReader, nType);
    };
    ForEach.prototype.getChildren = function() {
      return [this.alg, this.shape, this.presOf, this.constrLst, this.ruleLst, this.varLst].concat(this.list);
    };

    function SampData() {
      CBaseFormatNoIdObject.call(this);
      this.useDef = null;
      this.dataModel = null;
    }

    InitClass(SampData, CBaseFormatNoIdObject, AscDFH.historyitem_type_SampData);

    SampData.prototype.Write_ToBinary = function (w) {
      AscFormat.writeBool(w, this.useDef);
      AscFormat.writeObjectNoId(w, this.dataModel);
    };
    SampData.prototype.Read_FromBinary = function (r) {
      this.useDef = AscFormat.readBool(r);
      this.dataModel = AscFormat.ExecuteNoHistory(AscFormat.readObjectNoId, this, [r]);
    };
    SampData.prototype.setUseDef = function (pr) {
      this.useDef = pr;
    }

    SampData.prototype.setDataModel = function (oPr) {
      this.dataModel = oPr;
    }

    SampData.prototype.getUseDef = function () {
      return this.useDef;
    }

    SampData.prototype.getDataModel = function () {
      return this.dataModel;
    }

    SampData.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setUseDef(this.getUseDef());
      const dataModel = this.getDataModel();
      if (dataModel) {
        oCopy.setDataModel(AscFormat.ExecuteNoHistory(dataModel.createDuplicate, dataModel, [oIdMap]));
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
          AscFormat.ExecuteNoHistory(function() {
            this.setDataModel(new DataModel());
            this.dataModel.fromPPTY(pReader);
          }, this, []);
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

    function DiagramTitle() {
      CBaseFormatNoIdObject.call(this);
      this.lang = null;
      this.val = "";
    }

    InitClass(DiagramTitle, CBaseFormatNoIdObject, AscDFH.historyitem_type_DiagramTitle);

    DiagramTitle.prototype.Write_ToBinary = function (w) {
      AscFormat.writeString(w, this.lang);
      AscFormat.writeString(w, this.val);
    };
    DiagramTitle.prototype.Read_FromBinary = function (r) {
      this.lang = AscFormat.readString(r);
      this.val = AscFormat.readString(r);
    };
    DiagramTitle.prototype.setLang = function (pr) {
      this.lang = pr;
    }

    DiagramTitle.prototype.setVal = function (pr) {
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

    function ColorsDef() {
      CBaseFormatNoIdObject.call(this);
      this.minVer = null;
      this.uniqueId = null;
      this.catLst = null;
      this.desc = null;
      this.title = null;
      this.styleLbl = {};
    }

    InitClass(ColorsDef, CBaseFormatNoIdObject, AscDFH.historyitem_type_ColorsDef);
	  ColorsDef.prototype.Write_ToBinary = function (w) {
		  AscFormat.writeString(w, this.minVer);
		  AscFormat.writeString(w, this.uniqueId);
		  AscFormat.writeObjectNoId(w, this.catLst);
		  AscFormat.writeObjectNoId(w, this.desc);
		  AscFormat.writeObjectNoId(w, this.title);

		  const keys = Object.keys(this.styleLbl);
			w.WriteLong(keys.length);
		  for (let name in this.styleLbl) {
				AscFormat.writeObjectNoId(w, this.styleLbl[name]);
		  }
	  }
	  ColorsDef.prototype.Read_FromBinary = function (r) {
		  this.minVer = AscFormat.readString(r);
		  this.uniqueId = AscFormat.readString(r);
		  this.catLst = AscFormat.readObjectNoId(r);
		  this.desc = AscFormat.readObjectNoId(r);
		  this.title = AscFormat.readObjectNoId(r);

		  for (let i = r.GetLong(); i > 0; i -= 1) {
			  this.addToLstStyleLbl(AscFormat.readObjectNoId(r));
		  }
	  }
    ColorsDef.prototype.setMinVer = function (pr) {
      this.minVer = pr;
    }

    ColorsDef.prototype.setUniqueId = function (pr) {
      this.uniqueId = pr;
    }

    ColorsDef.prototype.setCatLst = function (oPr) {
      this.catLst = oPr;
    }

    ColorsDef.prototype.setDesc = function (oPr) {
      this.desc = oPr;
    };

    ColorsDef.prototype.setTitle = function (oPr) {
      this.title = oPr;
    };

    ColorsDef.prototype.addToLstStyleLbl = function (oPr) {
      this.styleLbl[oPr.name] = oPr;
    };

    ColorsDef.prototype.removeFromLstStyleLbl = function (name) {
      delete this.styleLbl[name];
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

    ColorsDef.prototype.getDesc = function () {
      return this.desc;
    }

    ColorsDef.prototype.getTitle = function () {
      return this.title;
    }
	  ColorsDef.prototype.getStyleLblList = function () {
		  const keys = Object.keys(this.styleLbl).sort();
			return keys.map(function (e) {
				return this.styleLbl[e];
			}, this);
	  };

    ColorsDef.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setMinVer(this.getMinVer());
      oCopy.setUniqueId(this.getUniqueId());
      if (this.getCatLst()) {
        oCopy.setCatLst(this.getCatLst().createDuplicate(oIdMap));
      }
      if (this.getDesc()) {
        oCopy.setDesc(this.getDesc().createDuplicate(oIdMap));
      }
      if (this.getTitle()) {
        oCopy.setTitle(this.getTitle().createDuplicate(oIdMap));
      }
			for (let name in this.styleLbl) {
				oCopy.addToLstStyleLbl(this.styleLbl[name].createDuplicate(oIdMap));
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
	    for (let name in this.styleLbl) {
		    this.writeRecord2(pWriter, 3, this.styleLbl[name]);
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
          this.setDesc(new DiagramTitle());
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
          this.addToLstStyleLbl(oDefStyle);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    ColorsDef.prototype.getChildren = function() {
      return [this.title, this.desc, this.catLst].concat(this.getStyleLblList());
    };

    function ColorDefStyleLbl() {
      CBaseFormatNoIdObject.call(this);
      this.name = null;
      this.effectClrLst = null;
      this.fillClrLst = null;
      this.linClrLst = null;
      this.txEffectClrLst = null;
      this.txFillClrLst = null;
      this.txLinClrLst = null;
    }

    InitClass(ColorDefStyleLbl, CBaseFormatNoIdObject, AscDFH.historyitem_type_ColorDefStyleLbl);
    ColorDefStyleLbl.prototype.Write_ToBinary = function (w) {
      AscFormat.writeString(w, this.name);
      AscFormat.writeObjectNoId(w, this.effectClrLst);
      AscFormat.writeObjectNoId(w, this.fillClrLst);
      AscFormat.writeObjectNoId(w, this.linClrLst);
      AscFormat.writeObjectNoId(w, this.txEffectClrLst);
      AscFormat.writeObjectNoId(w, this.txFillClrLst);
      AscFormat.writeObjectNoId(w, this.txLinClrLst);
    }
    ColorDefStyleLbl.prototype.Read_FromBinary = function (r) {
      this.name = AscFormat.readString(r);
      this.effectClrLst = AscFormat.readObjectNoId(r);
      this.fillClrLst = AscFormat.readObjectNoId(r);
      this.linClrLst = AscFormat.readObjectNoId(r);
      this.txEffectClrLst = AscFormat.readObjectNoId(r);
      this.txFillClrLst = AscFormat.readObjectNoId(r);
      this.txLinClrLst = AscFormat.readObjectNoId(r);
    }
    ColorDefStyleLbl.prototype.setName = function (pr) {
      this.name = pr;
    }
    ColorDefStyleLbl.prototype.setEffectClrLst = function (oPr) {
      this.effectClrLst = oPr;
    }

    ColorDefStyleLbl.prototype.setFillClrLst = function (oPr) {
      this.fillClrLst = oPr;
    }

    ColorDefStyleLbl.prototype.setLinClrLst = function (oPr) {
      this.linClrLst = oPr;
    }

    ColorDefStyleLbl.prototype.setTxEffectClrLst = function (oPr) {
      this.txEffectClrLst = oPr;
    }

    ColorDefStyleLbl.prototype.setTxFillClrLst = function (oPr) {
      this.txFillClrLst = oPr;
    }

    ColorDefStyleLbl.prototype.setTxLinClrLst = function (oPr) {
      this.txLinClrLst = oPr;
    }

    ColorDefStyleLbl.prototype.getName = function () {
      return this.name;
    }

    ColorDefStyleLbl.prototype.getEffectClrLst = function () {
      return this.effectClrLst;
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
          this.setEffectClrLst(new ClrLst());
          this.effectClrLst.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setFillClrLst(new ClrLst());
          this.fillClrLst.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setLinClrLst(new ClrLst());
          this.linClrLst.fromPPTY(pReader);
          break;
        }
        case 3: {
          this.setTxEffectClrLst(new ClrLst());
          this.txEffectClrLst.fromPPTY(pReader);
          break;
        }
        case 4: {
          this.setTxFillClrLst(new ClrLst());
          this.txFillClrLst.fromPPTY(pReader);
          break;
        }
        case 5: {
          this.setTxLinClrLst(new ClrLst());
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
	  ColorDefStyleLbl.prototype.checkNoFill = function () {
			switch (this.name) {
				case "parChTrans1D1":
					return true;
				default:
					return false;
			}
	  };

	  ColorDefStyleLbl.prototype.checkNoLn = function () {
		  switch (this.name) {
			  case "sibTrans2D1":
			  case "trBgShp":
			  case "bgShp":
			  case "dkBgShp":
			  case "fgSibTrans2D1":
			  case "revTx":
					return true;
			  default:
				  return false;
		  }
	  };
	  ColorDefStyleLbl.prototype.setShapeFill = function (shapes, parentObjects) {
			if (shapes.length && this.fillClrLst && !this.checkNoFill()) {
				const fillShapes = [];
				for (let i = 0; i < shapes.length; i++) {
					if (!(shapes[i].shape.hideGeom || shapes[i].type === AscFormat.LayoutShapeType_outputShapeType_conn)) {
						fillShapes.push(shapes[i]);
					}
				}
				const fills = this.fillClrLst.getCurColor(fillShapes.length, parentObjects);
				if (fills) {
					for (let i = 0; i < fills.length; i++) {
						fillShapes[i].setFill(fills[i]);
					}
				}
			}
	  };

	  ColorDefStyleLbl.prototype.getLineWidth = function (shadowShape) {
		  switch (this.name) {
			  case "trAlignAcc1":
			  case "sibTrans1D1":
				  return 6350;
			  default:
					return shadowShape.tailLnArrow || shadowShape.headLnArrow ? 6350 : 12700;
		  }
	  }
	  ColorDefStyleLbl.prototype.setShapeLn = function (shapes, parentObjects) {
		  if (shapes.length && this.linClrLst && !this.checkNoLn()) {
				const fillShapes = [];
				for (let i = 0; i < shapes.length; i += 1) {
					const shadowShape = shapes[i];
					if (!(shadowShape.shape.hideGeom ||
						shadowShape.node.isParNode() && shadowShape.type !== AscFormat.LayoutShapeType_outputShapeType_conn)) {
						fillShapes.push(shadowShape);
					}
				}
			  const fills = this.linClrLst.getCurColor(fillShapes.length, parentObjects);
			  if (fills) {
				  for (let i = 0; i < fills.length; i++) {
					  const shadowShape = fillShapes[i];
					  const ln = new AscFormat.CLn();
					  ln.setW(this.getLineWidth(shadowShape));
					  ln.setFill(fills[i]);
					  shadowShape.setLn(ln);
				  }
			  }
		  }
	  };

    function ClrLst() {
	    CBaseFormatNoIdObject.call(this);
	    this.list = [];
      this.hueDir = ClrLst_hueDir_cw;
      this.meth = ClrLst_meth_span;
    }
    InitClass(ClrLst, CBaseFormatNoIdObject, AscDFH.historyitem_type_ClrLst);
      ClrLst.prototype.Write_ToBinary = function (w) {
        w.WriteLong(this.list.length);
        for (let i = 0; i < this.list.length; i++) {
          AscFormat.writeObjectNoIdNoType(w, this.list[i]);
        }

        AscFormat.writeLong(w, this.hueDir);
        AscFormat.writeLong(w, this.meth);
      };
    ClrLst.prototype.Read_FromBinary = function (r) {
      for (let i = r.GetLong(); i > 0; i -= 1) {
        this.list.push(AscFormat.readObjectNoIdNoType(r, CUniColor));
      }

      this.hueDir = AscFormat.readLong(r);
      this.meth = AscFormat.readLong(r);
    };
	  ClrLst.prototype.getCurColor = function (length, parentObjects) {
			if (!length) {
				return;
			}
			if (this.list.length === 0) {
				return this.getNoFillColor(length);
			}

			if (this.meth === ClrLst_meth_repeat || this.list.length === 1) {
				return this.getRepeatColor(length);
			}

		  const startColor = this.list[0];
		  const endColor = this.list[1];
		  startColor.Calculate(parentObjects.theme, parentObjects.slide, parentObjects.layout, parentObjects.master);
		  endColor.Calculate(parentObjects.theme, parentObjects.slide, parentObjects.layout, parentObjects.master);
		  const startHSL = {};
		  AscFormat.CColorModifiers.prototype.RGB2HSL(startColor.RGBA.R, startColor.RGBA.G, startColor.RGBA.B, startHSL);

		  const endHSL = {};
		  AscFormat.CColorModifiers.prototype.RGB2HSL(endColor.RGBA.R, endColor.RGBA.G, endColor.RGBA.B, endHSL);

		  const diffHSL = {};
		  if (this.hueDir === ClrLst_hueDir_ccw) {
			  diffHSL.H = startHSL.H - endHSL.H;
		  } else {
			  diffHSL.H = endHSL.H - startHSL.H;
		  }
		  diffHSL.S = endHSL.S - startHSL.S;
		  diffHSL.L = endHSL.L - startHSL.L;
			diffHSL.A = endColor.RGBA.A - startColor.RGBA.A;
		  if (this.meth === ClrLst_meth_cycle) {
			  return this.getCycleColor(diffHSL, length);
		  } else {
				return this.getSpanColor(diffHSL, length);
		  }
	  };

	  ClrLst.prototype.getNoFillColor = function (length) {
		  const colors = [];
		  for (let i = 0; i < length; i++) {
			  colors.push(AscFormat.CreateNoFillUniFill());
		  }
		  return colors;
	  };

	  ClrLst.prototype.getRepeatColor = function (length) {
			const colors = [];
		  for (let i = 0; i < length; i++) {
			  const truthIndex = i % this.list.length;
			  const uniColor = this.list[truthIndex];
				colors.push(AscFormat.CreateUniFillByUniColorCopy(uniColor));
		  }

			return colors;
	  };
	  ClrLst.prototype.getInterpolateStartColor = function (diffHSL, scale) {
		  const startColor = this.list[0];

		  const hueOff = diffHSL.H * scale;
		  const satOff = diffHSL.S * scale;
		  const lumOff = diffHSL.L * scale;
		  const alphaOff = diffHSL.A * scale;
		  const copyColor = startColor.createDuplicate();
		  copyColor.addColorMod(new AscFormat.CColorMod("hueOff", ((hueOff / 255) * (360 * 60000)) >> 0));
		  copyColor.addColorMod(new AscFormat.CColorMod("satOff", (satOff / 255 * 100000) >> 0));
		  copyColor.addColorMod(new AscFormat.CColorMod("lumOff", (lumOff / 255 * 100000) >> 0));
		  copyColor.addColorMod(new AscFormat.CColorMod("alphaOff", (alphaOff / 255 * 100000) >> 0));
			return AscFormat.CreateUniFillByUniColor(copyColor);
	  };
		ClrLst.prototype.getCycleColor = function (diffHSL, length) {
			if (!length) {
				return;
			}
			const resultColors = [];

			for (let i = 0; i < length; i++) {
				const scale = (i * 2) / length;
				resultColors.push(this.getInterpolateStartColor(diffHSL, scale > 1 ? 2 - scale : scale));
			}
			return resultColors;
		};
		ClrLst.prototype.getSpanColor = function (diffHSL, length) {
			if (!length) {
				return;
			}

			const resultColors = [];
			for (let i = 0; i < length; i += 1) {
				const scale = i / (length - 1);
				resultColors.push(this.getInterpolateStartColor(diffHSL, scale));
			}
			return resultColors;
		}
    ClrLst.prototype.setHueDir = function (pr) {
      this.hueDir = pr;
    }

    ClrLst.prototype.setMeth = function (pr) {
      this.meth = pr;
    }

	  ClrLst.prototype.addToLst = function (nIdx, oPr) {
		  var nInsertIdx = Math.min(this.list.length, Math.max(0, nIdx));
		  nInsertIdx === this.list.length ? this.list.push(oPr) : this.list.splice(nInsertIdx, 0, oPr);
	  };

	  ClrLst.prototype.removeFromLst = function (nIdx) {
		  if (nIdx > -1 && nIdx < this.list.length) {
			  nIdx === this.list.length - 1 ? this.list.pop() : this.list.splice(nIdx, 1);
		  }
	  };

    ClrLst.prototype.getHueDir = function () {
      return this.hueDir;
    }

    ClrLst.prototype.getMeth = function () {
      return this.meth;
    }

	  ClrLst.prototype.privateWriteAttributes = function(pWriter) {
		  pWriter._WriteUChar2(0, this.hueDir);
		  pWriter._WriteUChar2(1, this.meth);
	  };
	  ClrLst.prototype.writeChildren = function(pWriter) {
		  for (var i = 0; i < this.list.length; i += 1) {
			  pWriter.WriteRecord2(0, this.list[i], pWriter.WriteUniColor);
		  }
	  };
	  ClrLst.prototype.readAttribute = function(nType, pReader) {
		  var oStream = pReader.stream;
		  if (0 === nType) this.setHueDir(oStream.GetUChar());
		  else if (1 === nType) this.setMeth(oStream.GetUChar());
	  };
	  ClrLst.prototype.readChild = function(nType, pReader) {
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

    ClrLst.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setHueDir(this.getHueDir());
      oCopy.setMeth(this.getMeth());
      for (var nIdx = 0; nIdx < this.list.length; ++nIdx) {
        var oColor = this.list[nIdx].createDuplicate(oIdMap);
        oCopy.addToLst(nIdx, oColor);
      }
    }
	  ClrLst.prototype.getChildren = function() {
		  return [].concat(this.list);
	  };

    function StyleDef() {
      CBaseFormatNoIdObject.call(this);
      this.minVer = null;
      this.uniqueId = null;
      this.catLst = null;
      this.scene3d = null;
      this.title = null;
      this.desc = null;
      this.styleLbl = {};
    }

    InitClass(StyleDef, CBaseFormatNoIdObject, AscDFH.historyitem_type_StyleDef);
    StyleDef.prototype.isEmptyStyleLbls = function() {
      for (let name in this.styleLbl) {
        return false;
      }
      return true;
    };
    StyleDef.prototype.Write_ToBinary = function (w) {
      AscFormat.writeString(w, this.minVer);
      AscFormat.writeString(w, this.uniqueId);
      AscFormat.writeObjectNoId(w, this.catLst);
      AscFormat.writeObjectNoId(w, this.scene3d);
      AscFormat.writeObjectNoId(w, this.title);
      AscFormat.writeObjectNoId(w, this.desc);
      const styleLbls = this.getStyleLblList();
      w.WriteLong(styleLbls.length);
      for (let i = 0; i < styleLbls.length; i += 1) {
        AscFormat.writeObjectNoId(w, styleLbls[i]);
      }
    };
    StyleDef.prototype.Read_FromBinary = function (r) {
      this.minVer = AscFormat.readString(r);
      this.uniqueId = AscFormat.readString(r);
      this.catLst = AscFormat.readObjectNoId(r);
      this.scene3d = AscFormat.readObjectNoId(r);
      this.title = AscFormat.readObjectNoId(r);
      this.desc = AscFormat.readObjectNoId(r);
      for (let i = r.GetLong(); i > 0; i -= 1) {
        this.addToLstStyleLbl(AscFormat.readObjectNoId(r));
      }
    };
    StyleDef.prototype.getStyleLblList = function () {
      const keys = Object.keys(this.styleLbl).sort();
      return keys.map(function (e) {
        return this.styleLbl[e];
      }, this);
    };
    StyleDef.prototype.setMinVer = function (pr) {
      this.minVer = pr;
    }

    StyleDef.prototype.setUniqueId = function (pr) {
      this.uniqueId = pr;
    }

    StyleDef.prototype.setCatLst = function (oPr) {
      this.catLst = oPr;
    }

    StyleDef.prototype.setScene3d = function (oPr) {
      this.scene3d = oPr;
    }

    StyleDef.prototype.setTitle = function (oPr) {
      this.title = oPr;
    }

    StyleDef.prototype.setDesc = function (oPr) {
      this.desc = oPr;
    }

    StyleDef.prototype.addToLstStyleLbl = function (oPr) {
      this.styleLbl[oPr.name] = oPr;
    };

    StyleDef.prototype.removeFromLstStyleLbl = function (name) {
        delete this.styleLbl[name];
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
      if (this.getScene3d()) {
        oCopy.setScene3d(this.getScene3d().createDuplicate(oIdMap));
      }
      if (this.getTitle()) {
        oCopy.setTitle(this.getTitle().createDuplicate(oIdMap));
      }
      if (this.getDesc()) {
        oCopy.setDesc(this.getDesc().createDuplicate(oIdMap));
      }
      const styleLbls = this.getStyleLblList();
      for (var nIdx = 0; nIdx < styleLbls.length; ++nIdx) {
        oCopy.addToLstStyleLbl(styleLbls[nIdx].createDuplicate(oIdMap));
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
      const styleLbls = this.getStyleLblList();
      for (var i = 0; i < styleLbls.length; ++i) {
        this.writeRecord2(pWriter, 4, styleLbls[i]);
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
          this.setDesc(new DiagramTitle());
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
          this.addToLstStyleLbl(oChild);
          break;
        }
        default: {
          s.SkipRecord();
          break;
        }
      }
    };
    StyleDef.prototype.getChildren = function() {
      return [this.title, this.desc, this.catLst, this.scene3d].concat(this.getStyleLblList());
    };

    function Scene3d() {
      CBaseFormatNoIdObject.call(this);
      this.backdrop = null;
      this.camera = null;
      this.lightRig = null;
    }

    InitClass(Scene3d, CBaseFormatNoIdObject, AscDFH.historyitem_type_Scene3d);
    Scene3d.prototype.Write_ToBinary = function (w) {
      AscFormat.writeObjectNoId(w, this.backdrop);
      AscFormat.writeObjectNoId(w, this.camera);
      AscFormat.writeObjectNoId(w, this.lightRig);
    };
    Scene3d.prototype.Read_FromBinary = function (r) {
      this.backdrop = AscFormat.readObjectNoId(r);
      this.camera = AscFormat.readObjectNoId(r);
      this.lightRig = AscFormat.readObjectNoId(r);
    };
    Scene3d.prototype.setBackdrop = function (oPr) {
      this.backdrop = oPr;
    }

    Scene3d.prototype.setCamera = function (oPr) {
      this.camera = oPr;
    }

    Scene3d.prototype.setLightRig = function (oPr) {
      this.lightRig = oPr;
    }

    Scene3d.prototype.getBackdrop = function () {
      return this.backdrop;
    }

    Scene3d.prototype.getCamera = function () {
      return this.camera;
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
    };
    function StyleDefStyleLbl() {
      CBaseFormatNoIdObject.call(this);
      this.name = null;
      this.scene3d = null;
      this.sp3d = null;
      this.style = null;
      this.txPr = null;
    }

    InitClass(StyleDefStyleLbl, CBaseFormatNoIdObject, AscDFH.historyitem_type_StyleDefStyleLbl);
    StyleDefStyleLbl.prototype.Write_ToBinary = function (w) {
      AscFormat.writeString(w, this.name);
      AscFormat.writeObjectNoId(w, this.scene3d);
      AscFormat.writeObjectNoId(w, this.sp3d);
      AscFormat.writeObjectNoId(w, this.style);
      AscFormat.writeObjectNoIdNoType(w, this.txPr);
    };
    StyleDefStyleLbl.prototype.Read_FromBinary = function (r) {
      this.name = AscFormat.readString(r);
      this.scene3d = AscFormat.readObjectNoId(r);
      this.sp3d = AscFormat.readObjectNoId(r);
      this.style = AscFormat.ExecuteNoHistory(AscFormat.readObjectNoId, this, [r]);
      this.txPr = AscFormat.ExecuteNoHistory(AscFormat.readObjectNoIdNoType, this, [r, AscFormat.CBodyPr]);
    };
	  StyleDefStyleLbl.prototype.setShapeStyle = function (shapes) {
		  if (this.style) {
			  for (let i = 0; i < shapes.length; i += 1) {
				  const shape = shapes[i];
					shape.setStyle(this.style.createDuplicate());
			  }
		  }
	  }
    StyleDefStyleLbl.prototype.setName = function (pr) {
      this.name = pr;
    }

    StyleDefStyleLbl.prototype.setScene3d = function (oPr) {
      this.scene3d = oPr;
    }

    StyleDefStyleLbl.prototype.setSp3d = function (oPr) {
      this.sp3d = oPr;
    }

    StyleDefStyleLbl.prototype.setStyle = function (oPr) {
      this.style = oPr;
    }

    StyleDefStyleLbl.prototype.setTxPr = function (oPr) {
      this.txPr = oPr;
    }

    StyleDefStyleLbl.prototype.getName = function () {
      return this.name;
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
      if (this.getScene3d()) {
        oCopy.setScene3d(this.getScene3d().createDuplicate(oIdMap));
      }
      if (this.getSp3d()) {
        oCopy.setSp3d(this.getSp3d().createDuplicate(oIdMap));
      }
      const style = this.getStyle();
      if (style) {
        oCopy.setStyle(AscFormat.ExecuteNoHistory(style.createDuplicate, style, [oIdMap]));
      }
      const txPr = this.getTxPr();
      if (txPr) {
        oCopy.setTxPr(AscFormat.ExecuteNoHistory(txPr.createDuplicate, txPr, [oIdMap]));
      }
    }

    StyleDefStyleLbl.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteString2(0, this.name);
    };
    StyleDefStyleLbl.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.scene3d);
      this.writeRecord2(pWriter, 1, this.sp3d);
      pWriter.WriteRecord2(2, this.style, pWriter.WriteShapeStyle);
      if (this.txPr) {
        AscFormat.ExecuteNoHistory(function() {
          const textBody = new AscFormat.CTextBody();
          textBody.setBodyPr(this.txPr);
          pWriter.WriteRecord2(3, textBody, pWriter.WriteTxBody);
        }, this, []);
      }

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
          AscFormat.ExecuteNoHistory(function() {
          this.setStyle(pReader.ReadShapeStyle());
          }, this, []);
          break;
        }
        case 3: {
          AscFormat.ExecuteNoHistory(function() {
            const textBody = pReader.ReadTextBody();
            this.setTxPr(textBody.bodyPr);
          }, this, []);
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
    function Point3D() {
      AscFormat.CBaseNoIdObject.call(this);
      this.x = null;
      this.y = null;
      this.z = null;
    }
    InitClass(Point3D, AscFormat.CBaseNoIdObject, 0);
    Point3D.prototype.Write_ToBinary = function(w) {
      AscFormat.writeLong(w, this.x);
      AscFormat.writeLong(w, this.y);
      AscFormat.writeLong(w, this.z);
    };
    Point3D.prototype.Read_FromBinary = function(r) {
      this.x = AscFormat.readLong(r);
      this.y = AscFormat.readLong(r);
      this.z = AscFormat.readLong(r);
    };
    Point3D.prototype.setX = function(pr) {
      this.x = pr;
    }
    Point3D.prototype.setY = function(pr) {
      this.y = pr;
    }
    Point3D.prototype.setZ = function(pr) {
      this.z = pr;
    }
    Point3D.prototype.fillObject = function(oCopy, oIdMap) {
      if (this.x !== null) {
        oCopy.x = this.x;
      }
      if (this.y !== null) {
        oCopy.y = this.y;
      }
      if (this.z !== null) {
        oCopy.z = this.z;
      }
    }
    Point3D.prototype.getChildren = function() {
      return [this.x, this.y, this.z];
    };
    function Vector3D() {
      AscFormat.CBaseNoIdObject.call(this);
      this.dx = null;
      this.dy = null;
      this.dz = null;
    }
    InitClass(Vector3D, AscFormat.CBaseNoIdObject, 0);
    Vector3D.prototype.Write_ToBinary = function(w) {
      AscFormat.writeLong(w, this.dx);
      AscFormat.writeLong(w, this.dy);
      AscFormat.writeLong(w, this.dz);
    };
    Vector3D.prototype.Read_FromBinary = function(r) {
      this.dx = AscFormat.readLong(r);
      this.dy = AscFormat.readLong(r);
      this.dz = AscFormat.readLong(r);
    };
    Vector3D.prototype.setDx = function(pr) {
      this.dx = pr;
    }
    Vector3D.prototype.setDy = function(pr) {
      this.dy = pr;
    }
    Vector3D.prototype.setDz = function(pr) {
      this.dz = pr;
    }
    Vector3D.prototype.fillObject = function(oCopy, oIdMap) {
      if (this.dx !== null) {
        oCopy.dx = this.dx;
      }
      if (this.dy !== null) {
        oCopy.dy = this.dy;
      }
      if (this.dz !== null) {
        oCopy.dz = this.dz;
      }
    }
    Vector3D.prototype.getChildren = function() {
      return [this.dx, this.dy, this.dz];
    }
    function Backdrop() {
      CBaseFormatNoIdObject.call(this);
      this.anchor = new Point3D();
      this.norm = new Vector3D();
      this.up = new Vector3D();
    }

    InitClass(Backdrop, CBaseFormatNoIdObject, AscDFH.historyitem_type_Backdrop);
    Backdrop.prototype.Write_ToBinary = function(w) {
      AscFormat.writeObjectNoIdNoType(w, this.anchor);
      AscFormat.writeObjectNoIdNoType(w, this.norm);
      AscFormat.writeObjectNoIdNoType(w, this.up);
    };
    Backdrop.prototype.Read_FromBinary = function(r) {
      this.anchor = AscFormat.readObjectNoIdNoType(r, Point3D);
      this.norm = AscFormat.readObjectNoIdNoType(r, Vector3D);
      this.up = AscFormat.readObjectNoIdNoType(r, Vector3D);
    };
    Backdrop.prototype.setAnchor = function (oPr) {
      this.anchor = oPr;
    }

    Backdrop.prototype.setNorm = function (oPr) {
      this.norm = oPr;
    }

    Backdrop.prototype.setUp = function (oPr) {
      this.up = oPr;
    }

    Backdrop.prototype.getAnchor = function () {
      return this.anchor;
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

    function Drawing() {
      CGroupShape.call(this);
    }

    InitClass(Drawing, CGroupShape, AscDFH.historyitem_type_SmartArtDrawing);

		Drawing.prototype.recalcSmartArtConnections = function () {
			if (this.group) {
				this.group.recalcSmartArtConnections();
			}
		};
    Drawing.prototype.getObjectType = function () {
      return AscDFH.historyitem_type_SmartArtDrawing;
    }
    Drawing.prototype.getName = function () {
      return 'Drawing';
    }
	  Drawing.prototype.Get_ParentParagraph = function ()
	  {
			if (this.group)
			{
				return this.group.parent && this.group.parent.Get_ParentParagraph && this.group.parent.Get_ParentParagraph();
			}
	  };
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
      if(!oPr || false !== oPr.cacheImage) {
        copy.cachedImage = this.getBase64Img();
        copy.cachedPixH = this.cachedPixH;
        copy.cachedPixW = this.cachedPixW;
      }
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

    Drawing.prototype.getResultScaleCoefficients = function() {
      let oParaDrawing = AscFormat.getParaDrawing(this);
      if(oParaDrawing) {
        let dScaleCoefficient = oParaDrawing.GetScaleCoefficient();
        return {cx: dScaleCoefficient, cy: dScaleCoefficient};
      }
      return {cx: 1, cy: 1};
    };

    Drawing.prototype.setXfrmByParent = function () {
      if (!this.spPr) {
        this.setSpPr(new AscFormat.CSpPr());
      }
      if (!this.spPr.xfrm) {
        this.spPr.setXfrm(new AscFormat.CXfrm());
      }
      var oXfrm = this.spPr.xfrm;
      if (oXfrm.isNull() || oXfrm.isZero()) {
        var parent = this.group;
        if (parent && parent.spPr.xfrm) {
          oXfrm.setOffX(0);
          oXfrm.setOffY(0);
          oXfrm.setExtX(parent.spPr.xfrm.extX);
          oXfrm.setExtY(parent.spPr.xfrm.extY);
          oXfrm.setChOffX(0);
          oXfrm.setChOffY(0);
          oXfrm.setChExtX(parent.spPr.xfrm.extX);
          oXfrm.setChExtY(parent.spPr.xfrm.extY);
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

    function Camera() {
      CBaseFormatNoIdObject.call(this);
      this.fov = null;
      this.prst = null;
      this.zoom = null;
      this.rot = null;
    }

    InitClass(Camera, CBaseFormatNoIdObject, AscDFH.historyitem_type_Camera);
    Camera.prototype.Write_ToBinary = function(w) {
      AscFormat.writeDouble2(w, this.fov);
      AscFormat.writeLong(w, this.prst);
      AscFormat.writeDouble2(w, this.zoom);
      AscFormat.writeObjectNoId(w, this.rot);
    };
    Camera.prototype.Read_FromBinary = function(r) {
      this.fov = AscFormat.readDouble2(r);
      this.prst = AscFormat.readLong(r);
      this.zoom = AscFormat.readDouble2(r);
      this.rot = AscFormat.readObjectNoId(r);
    };
    Camera.prototype.setFov = function (pr) {
      this.fov = pr;
    }

    Camera.prototype.setPrst = function (pr) {
      this.prst = pr;
    }

    Camera.prototype.setZoom = function (pr) {
      this.zoom = pr;
    }

    Camera.prototype.setRot = function (oPr) {
      this.rot = oPr;
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
    };

    function Rot() {
      CBaseFormatNoIdObject.call(this);
      this.lat = null;
      this.lon = null;
      this.rev = null;
    }

    InitClass(Rot, CBaseFormatNoIdObject, AscDFH.historyitem_type_Rot);
    Rot.prototype.Write_ToBinary = function(w) {
      AscFormat.writeLong(w, this.lat);
      AscFormat.writeLong(w, this.lon);
      AscFormat.writeLong(w, this.rev);
    };
    Rot.prototype.Read_FromBinary = function(r) {
      this.lat = AscFormat.readLong(r);
      this.lon = AscFormat.readLong(r);
      this.rev = AscFormat.readLong(r);
    };
    Rot.prototype.setLat = function (pr) {
      this.lat = pr;
    }

    Rot.prototype.setLon = function (pr) {
      this.lon = pr;
    }

    Rot.prototype.setRev = function (pr) {
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

    function LightRig() {
      CBaseFormatNoIdObject.call(this);
      this.dir = null;
      this.rig = null;
      this.rot = null;
    }

    InitClass(LightRig, CBaseFormatNoIdObject, AscDFH.historyitem_type_LightRig);
    LightRig.prototype.Write_ToBinary = function(w) {
      AscFormat.writeLong(w, this.dir);
      AscFormat.writeLong(w, this.rig);
      AscFormat.writeObjectNoId(w, this.rot);
    }
    LightRig.prototype.Read_FromBinary = function(r) {
      this.dir = AscFormat.readLong(r);
      this.rig = AscFormat.readLong(r);
      this.rot = AscFormat.readObjectNoId(r);
    }
    LightRig.prototype.setDir = function (pr) {
      this.dir = pr;
    }

    LightRig.prototype.setRig = function (pr) {
      this.rig = pr;
    }

    LightRig.prototype.setRot = function (oPr) {
      this.rot = oPr;
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

    function Sp3d() {
      CBaseFormatNoIdObject.call(this);
      this.contourW = null;
      this.extrusionH = null;
      this.prstMaterial = null;
      this.z = null;
      this.bevelB = null;
      this.bevelT = null;
      this.contourClr = null;
      this.extrusionClr = null;
    }

    InitClass(Sp3d, CBaseFormatNoIdObject, AscDFH.historyitem_type_Sp3d);
    Sp3d.prototype.Write_ToBinary = function(w) {
      AscFormat.writeLong(w, this.contourW);
      AscFormat.writeLong(w, this.extrusionH);
      AscFormat.writeLong(w, this.prstMaterial);
      AscFormat.writeLong(w, this.z);
      AscFormat.writeObjectNoId(w, this.bevelB);
      AscFormat.writeObjectNoId(w, this.bevelT);
      AscFormat.writeObjectNoIdNoType(w, this.contourClr);
      AscFormat.writeObjectNoIdNoType(w, this.extrusionClr);
    }
    Sp3d.prototype.Read_FromBinary = function(r) {
      this.contourW = AscFormat.readLong(r);
      this.extrusionH = AscFormat.readLong(r);
      this.prstMaterial = AscFormat.readLong(r);
      this.z = AscFormat.readLong(r);
      this.bevelB = AscFormat.readObjectNoId(r);
      this.bevelT = AscFormat.readObjectNoId(r);
      this.contourClr = AscFormat.readObjectNoIdNoType(r, CUniColor);
      this.extrusionClr = AscFormat.readObjectNoIdNoType(r, CUniColor);
    }
    Sp3d.prototype.setContourW = function (pr) {
      this.contourW = pr;
    }

    Sp3d.prototype.setExtrusionH = function (pr) {
      this.extrusionH = pr;
    }

    Sp3d.prototype.setPrstMaterial = function (pr) {
      this.prstMaterial = pr;
    }

    Sp3d.prototype.setZ = function (oPr) {
      this.z = oPr;
    }

    Sp3d.prototype.setBevelB = function (oPr) {
      this.bevelB = oPr;
    }

    Sp3d.prototype.setBevelT = function (oPr) {
      this.bevelT = oPr;
    }

    Sp3d.prototype.setContourClr = function (oPr) {
      this.contourClr = oPr;
    }

    Sp3d.prototype.setExtrusionClr = function (oPr) {
      this.extrusionClr = oPr;
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

    Sp3d.prototype.getExtrusionClr = function () {
      return this.extrusionClr;
    }

    Sp3d.prototype.fillObject = function (oCopy, oIdMap) {
      oCopy.setContourW(this.getContourW());
      oCopy.setExtrusionH(this.getExtrusionH());
      oCopy.setPrstMaterial(this.getPrstMaterial());
      oCopy.setZ(this.getZ());
      if (this.getBevelB()) {
        oCopy.setBevelB(this.getBevelB().createDuplicate(oIdMap));
      }
      if (this.getBevelT()) {
        oCopy.setBevelT(this.getBevelT().createDuplicate(oIdMap));
      }
      if (this.getContourClr()) {
        oCopy.setContourClr(this.getContourClr().createDuplicate(oIdMap));
      }
      if (this.getExtrusionClr()) {
        oCopy.setExtrusionClr(this.getExtrusionClr().createDuplicate(oIdMap));
      }
    }

    Sp3d.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteInt2(0, this.contourW);
      pWriter._WriteInt2(1, this.extrusionH);
      pWriter._WriteLimit2(2, this.prstMaterial);
      pWriter._WriteInt2(3, this.z);
    };
    Sp3d.prototype.writeChildren = function(pWriter) {
      this.writeRecord2(pWriter, 0, this.bevelT);
      this.writeRecord2(pWriter, 1, this.bevelB);
      pWriter.WriteRecord1(2, this.extrusionClr, pWriter.WriteUniColor);
      pWriter.WriteRecord1(3, this.contourClr, pWriter.WriteUniColor);

    };
    Sp3d.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      switch (nType) {
        case 0:
          this.setContourW(oStream.GetLong());
          break;
        case 1:
          this.setExtrusionH(oStream.GetLong());
          break;
        case 2:
          this.setPrstMaterial(oStream.GetUChar());
          break;
        case 3:
          this.setZ(oStream.GetLong());
          break;
        default:
          break;
      }
    };
    Sp3d.prototype.readChild = function(nType, pReader) {
      var s = pReader.stream;
      switch (nType) {
        case 0: {
          this.setBevelT(new Bevel());
          this.bevelT.fromPPTY(pReader);
          break;
        }
        case 1: {
          this.setBevelB(new Bevel());
          this.bevelB.fromPPTY(pReader);
          break;
        }
        case 2: {
          this.setExtrusionClr(pReader.ReadUniColor());
          break;
        }
        case 3: {
          this.setContourClr(pReader.ReadUniColor());
          break;
        }
        default: {
          break;
        }
      }
    };

    Sp3d.prototype.getChildren = function() {
      return [this.contourW, this.extrusionH, this.prstMaterial, this.z, this.bevelB, this.bevelT, this.contourClr, this.extrusionClr];
    };

    function Bevel() {
      CBaseFormatNoIdObject.call(this);
      this.h = null;
      this.prst = null;
      this.w = null;
    }

    InitClass(Bevel, CBaseFormatNoIdObject, AscDFH.historyitem_type_Bevel);
    Bevel.prototype.Write_ToBinary = function(w) {
      AscFormat.writeLong(w, this.h);
      AscFormat.writeLong(w, this.prst);
      AscFormat.writeLong(w, this.w);
    }
    Bevel.prototype.Read_FromBinary = function(r) {
      this.h = AscFormat.readLong(r);
      this.prst = AscFormat.readLong(r);
      this.w = AscFormat.readLong(r);
    }
    Bevel.prototype.setH = function (pr) {
      this.h = pr;
    }

    Bevel.prototype.setPrst = function (pr) {
      this.prst = pr;
    }

    Bevel.prototype.setW = function (pr) {
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

    Bevel.prototype.privateWriteAttributes = function(pWriter) {
      pWriter._WriteInt2(0, this.w);
      pWriter._WriteInt2(1, this.h);
      pWriter._WriteLimit2(2, this.prst);
    };
    Bevel.prototype.readAttribute = function(nType, pReader) {
      var oStream = pReader.stream;
      switch (nType) {
        case 0:
          this.setW(oStream.GetLong());
          break;
        case 1:
          this.setH(oStream.GetLong());
          break;
        case 2:
          this.setPrst(oStream.GetUChar());
          break;
        default:
          break;
      }
    };

    Bevel.prototype.getChildren = function() {
      return [this.h, this.prst, this.w];
    };

    function ShapeSmartArtInfo() {
			this.shape = null;
      this.shapePoint = null;
      this.contentPoint = [];
      this.maxFontSize = null;
			this.textConstraints = {};
			this.params = {};
			this.textConstraintRelations = [];
			this.adaptFontSizeShapesInfo = null;
	    this.secondaryFontSizeScale = null;
    }
	  ShapeSmartArtInfo.prototype.getContentFillingType = function (shapes) {

			let res = 0x00;
		  for (let i = 0; i < shapes.length; i += 1) {
			  const shape = shapes[i];
			  if (!shape.isCanFitFontSize()) {
				  continue;
			  }
				const docContent = shape.getDocContent();
				if (docContent) {
					for (let j = 0; j < docContent.Content.length; j++) {
						const item = docContent.Content[j];
						if (item.PresentationPr && item.PresentationPr.Bullet && !item.PresentationPr.Bullet.IsNone()) {
							res |= smartArtContentFillingType_onlyChildren;
						} else {
							res |= smartArtContentFillingType_onlyParent;
						}
						if (res === smartArtContentFillingType_parentWithChildren) {
							return res;
						}
					}
				}
		  }
			return res;
	  };
		ShapeSmartArtInfo.prototype.getSecondaryFontSizeCoefficient = function () {
			if (this.secondaryFontSizeScale === null) {
				const primaryConstr = this.textConstraints[AscFormat.Constr_type_primFontSz];
				if (primaryConstr) {
					const coefficient = primaryConstr.getSecondaryFontSizeCoefficient();
					if (coefficient !== null) {
						this.secondaryFontSizeScale = coefficient;
						return coefficient;
					}
				}

				const secondaryConstr = this.textConstraints[AscFormat.Constr_type_secFontSz];
				if (secondaryConstr) {
					const coefficient = secondaryConstr.getSecondaryFontSizeCoefficient();
					if (coefficient !== null) {
						this.secondaryFontSizeScale = coefficient;
						return coefficient;
					}
				}

				this.secondaryFontSizeScale = 0.78;
			}
			return this.secondaryFontSizeScale;
		};
		ShapeSmartArtInfo.prototype.getAdaptFontSizeInfo = function () {
			if (this.adaptFontSizeShapesInfo === null) {
				const checkShapes = {};

				this.adaptFontSizeShapesInfo = {};
				this.adaptFontSizeShapesInfo.shapes = [];
				const res = this.adaptFontSizeShapesInfo.shapes;
				const checkShapeSmartArtInfos = [this];
				while (checkShapeSmartArtInfos.length) {
					const smartArtInfo = checkShapeSmartArtInfos.pop();
					const shape = smartArtInfo.shape;
					if (!checkShapes[shape.GetId()]) {
						res.push(shape);
						checkShapes[shape.GetId()] = true;
					}


					const primFontConstraint = smartArtInfo.textConstraints[AscFormat.Constr_type_primFontSz];
					if (primFontConstraint) {
						primFontConstraint.collectShapeSmartArtInfo(AscFormat.Constr_op_equ, checkShapeSmartArtInfos, checkShapes);
						primFontConstraint.collectShapeSmartArtInfo(AscFormat.Constr_op_equ, checkShapeSmartArtInfos, checkShapes, true);
						primFontConstraint.collectShapeSmartArtInfo(AscFormat.Constr_op_none, checkShapeSmartArtInfos, checkShapes);
						primFontConstraint.collectShapeSmartArtInfo(AscFormat.Constr_op_lte, checkShapeSmartArtInfos, checkShapes);
					}
					const secFontConstraint = smartArtInfo.textConstraints[AscFormat.Constr_type_secFontSz];
					if (secFontConstraint) {
						secFontConstraint.collectShapeSmartArtInfo(AscFormat.Constr_op_equ, checkShapeSmartArtInfos, checkShapes);
						secFontConstraint.collectShapeSmartArtInfo(AscFormat.Constr_op_equ, checkShapeSmartArtInfos, checkShapes, true);
						secFontConstraint.collectShapeSmartArtInfo(AscFormat.Constr_op_none, checkShapeSmartArtInfos, checkShapes);
						secFontConstraint.collectShapeSmartArtInfo(AscFormat.Constr_op_lte, checkShapeSmartArtInfos, checkShapes);
					}
				}
				this.adaptFontSizeShapesInfo.contentFillingType = this.getContentFillingType(res);
			}

			return this.adaptFontSizeShapesInfo;
		};
	  ShapeSmartArtInfo.prototype.getChildrenSpacingScale = function () {
		  if (this.params[AscFormat.Param_type_lnSpAfChP] !== undefined) {
			  return this.params[AscFormat.Param_type_lnSpAfChP];
		  }
		  return g_dKoef_pt_to_mm * 0.18;
	  };
	  ShapeSmartArtInfo.prototype.getParentSpacingScale = function () {
		  if (this.params[AscFormat.Param_type_lnSpAfParP] !== undefined) {
			  return this.params[AscFormat.Param_type_lnSpAfParP];
		  }
			return g_dKoef_pt_to_mm * 0.42;
	  };
	  ShapeSmartArtInfo.prototype.setShape = function (shape) {
		  this.shape = shape;
	  }
		ShapeSmartArtInfo.prototype.getMarginFactors = function () {
			const res = {};
			res.bMarg = this.textConstraints[AscFormat.Constr_type_bMarg];
			res.tMarg = this.textConstraints[AscFormat.Constr_type_tMarg];
			res.rMarg = this.textConstraints[AscFormat.Constr_type_rMarg];
			res.lMarg = this.textConstraints[AscFormat.Constr_type_lMarg];
			return res;
		};
		ShapeSmartArtInfo.prototype.getRelFitFontSize = function (isUseChildrenCoefficient) {
			const adaptInfo = this.getAdaptFontSizeInfo();
			const shapes = adaptInfo.shapes;
			let minFontSize = null;
			for (let i = 0; i < shapes.length; i++) {
				const shape = shapes[i];
				const content = shape.getDocContent();
				const shapeInfo = shape.getSmartArtInfo();
				const isNotPlaceholder = content && !content.Is_Empty({SkipEnd: true, SkipPlcHldr: false}) && shapeInfo.contentPoint.some(function (node) {
					const point = node.point;
					return point && point.prSet && !point.prSet.custT && !point.prSet.phldr;
				});
				if (isNotPlaceholder) {

					if (shapeInfo.maxFontSize === null) {
						this.setMaxFontSize(shape.findFitFontSizeForSmartArt());
					}
					if (minFontSize === null || minFontSize > shapeInfo.maxFontSize) {
						minFontSize = shapeInfo.maxFontSize;
					}
				}
			if (minFontSize !== null && isUseChildrenCoefficient) {
				const contentFillingType = this.getContentFillingType(shapes);
				if (contentFillingType === smartArtContentFillingType_onlyChildren || contentFillingType === smartArtContentFillingType_parentWithChildren) {
					return minFontSize;
				} else if (contentFillingType === smartArtContentFillingType_onlyParent) {
					return Math.round(minFontSize * shapeInfo.getSecondaryFontSizeCoefficient());
				}
			}
			}
			return minFontSize;
		};
		ShapeSmartArtInfo.prototype.getMaxConstrFontSize = function (isUseChildrenCoefficient) {
			const fontSizes = [65];
			const primTextConstr = this.textConstraints[AscFormat.Constr_type_primFontSz];
			if (primTextConstr) {
				fontSizes.push(primTextConstr.getMaxFontSize(isUseChildrenCoefficient));
			}
			const secTextConstr = this.textConstraints[AscFormat.Constr_type_secFontSz];
			if (secTextConstr) {
				fontSizes.push(secTextConstr.getMaxFontSize(isUseChildrenCoefficient));
			}
			return Math.min.apply(Math, fontSizes);
		};
	  ShapeSmartArtInfo.prototype.getMinConstrFontSize = function () {
		  const fontSizes = [5];
		  const primTextConstr = this.textConstraints[AscFormat.Constr_type_primFontSz];
		  if (primTextConstr) {
			  fontSizes.push(primTextConstr.getMinFontSize());
		  }
		  const secTextConstr = this.textConstraints[AscFormat.Constr_type_secFontSz];
		  if (secTextConstr) {
			  fontSizes.push(secTextConstr.getMinFontSize());
		  }
		  return Math.min.apply(Math, fontSizes);
	  };
		ShapeSmartArtInfo.prototype.getShape = function () {
			return this.shape;
		};
    ShapeSmartArtInfo.prototype.setShapePoint = function (oPr) {
      this.shapePoint = oPr;
    }

    ShapeSmartArtInfo.prototype.addToLstContentPoint = function (nIdx, oPr) {
      var nInsertIdx = Math.min(this.contentPoint.length, Math.max(0, nIdx));
      nInsertIdx === this.contentPoint.length ? this.contentPoint.push(oPr) : this.contentPoint.splice(nInsertIdx, 0, oPr);
    }

    ShapeSmartArtInfo.prototype.removeFromLstContentPoint = function (nIdx) {
      if (nIdx > -1 && nIdx < this.contentPoint.length) {
        nIdx === this.contentPoint.length - 1 ? this.contentPoint.pop() : this.contentPoint.splice(nIdx, 1);
      }
    }
    ShapeSmartArtInfo.prototype.setMaxFontSize = function (nPr) {
      this.maxFontSize = nPr;
    };
	  ShapeSmartArtInfo.prototype.collectTextConstraintRelations = function (array) {
			array = array || [];
		  for (let i = 0; i < this.textConstraintRelations.length; i += 1) {
			  const presNodeArray = this.textConstraintRelations[i];
			  if (!array.includes(presNodeArray)) {
				  array.push(presNodeArray);
			  }
		  }
			return array;
	  };

    changesFactory[AscDFH.historyitem_SmartArtColorsDef] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_SmartArtDrawing] = CChangeObject;
    changesFactory[AscDFH.historyitem_SmartArtLayoutDef] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_SmartArtDataModel] = CChangeObject;
    changesFactory[AscDFH.historyitem_SmartArtStyleDef] = CChangeObjectNoId;
    changesFactory[AscDFH.historyitem_SmartArtParent] = CChangeObject;
    drawingConstructorsMap[AscDFH.historyitem_SmartArtColorsDef] = ColorsDef;
    drawingConstructorsMap[AscDFH.historyitem_SmartArtLayoutDef] = LayoutDef;
    drawingConstructorsMap[AscDFH.historyitem_SmartArtStyleDef] = StyleDef;
    drawingsChangesMap[AscDFH.historyitem_SmartArtColorsDef] = function (oClass, value) {
      oClass.colorsDef = value;
    };
    drawingsChangesMap[AscDFH.historyitem_SmartArtDrawing] = function (oClass, value) {
      oClass.drawing = value;
			oClass.isLocalDrawingPart = oClass.localDrawingId !== null ? oClass.localDrawingId === (value && value.Get_Id()) : false;
			oClass.recalcSmartArtConnections();
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
		oClass.oldParent = oClass.parent;
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
      this.bNeedUpdatePosition = true;

      this.calcGeometry = null;
			this.isLocalDrawingPart = true;
			this.localDrawingId = null;
    }

    InitClass(SmartArt, CGroupShape, AscDFH.historyitem_type_SmartArt);
		SmartArt.prototype.correctUngeneratedSmartArtContent = function() {
			if (this.isCanGenerateSmartArt()) {
				return;
			}
			const drawing = this.getDrawing();
			if (drawing) {
				this.reconnectSmartArtShapes();
				for (let i = 0; i < drawing.spTree.length; i += 1) {
					const shape = drawing.spTree[i];
					if (!shape.isCorrectSmartArtContentPoints()) {
						shape.correctUngeneratedSmartArtContent();
					}
				}
			}
		};
    SmartArt.prototype.getDataModelFromData = function() {
      if (this.dataModel) {
        return this.dataModel.getDataModel();
      }
    };
		SmartArt.prototype.isEmptyColors = function () {
			return !this.colorsDef;
		};
	  SmartArt.prototype.isEmptyStyles = function () {
		  return !(this.styleDef && this.styleDef.uniqueId && !this.styleDef.isEmptyStyleLbls());
	  };
	  SmartArt.prototype.isEmptyLayout = function () {
			return !(this.layoutDef && this.layoutDef.layoutNode && this.layoutDef.uniqueId);
	  };
		SmartArt.prototype.generateDefaultStructures = function () {
			if (this.isEmptyColors()) {
				this.setColorsDef(AscFormat.generateDefaultSmartArtColors());
			}
			if (this.isEmptyStyles()){
				this.setStyleDef(AscFormat.generateDefaultSmartArtQuickStyle());
			}
      if (this.isEmptyLayout()) {
        this.setLayoutDef(AscFormat.generateDefaultSmartArtLayout());
      }
		};
    SmartArt.prototype.initSmartArtAlgorithm = function() {
      if (!this.smartArtTree) {
        this.smartArtTree = new AscFormat.SmartArtAlgorithm(this);
      }
    };
    SmartArt.prototype.checkDataModel = function() {
      if (this.isCanGenerateSmartArt()) {
        this.initSmartArtAlgorithm();
        this.smartArtTree.checkDataModel();
      }
    }
	  SmartArt.prototype.recalcFitFontSize = function () {
		  this.recalcInfo.fitFontSize = true;
	  };
	  SmartArt.prototype.recalcSmartArtConnections = function () {
		  this.recalcInfo.reconnectSmartArtShapes = true;
	  };
		SmartArt.prototype.setRecalculateInfo = function () {
			CGroupShape.prototype.setRecalculateInfo.call(this);
			this.recalcInfo.fitFontSize = true;
			this.recalcInfo.reconnectSmartArtShapes = true;
		};
    SmartArt.prototype.getObjectType = function() {
      return AscDFH.historyitem_type_SmartArt;
    };
    SmartArt.prototype.getName = function () {
      return 'SmartArt';
    };

		SmartArt.prototype.isCanGenerateSmartArt = function () {
			const smartartType = this.getTypeOfSmartArt();
			switch (smartartType) {
				case Asc.c_oAscSmartArtTypes.AlternatingHexagonList:
				case Asc.c_oAscSmartArtTypes.AlternatingPictureBlocks:
				case Asc.c_oAscSmartArtTypes.ArrowRibbon:
				case Asc.c_oAscSmartArtTypes.AscendingPictureAccentProcess:
				case Asc.c_oAscSmartArtTypes.Balance:
				case Asc.c_oAscSmartArtTypes.BasicBlockList:
				case Asc.c_oAscSmartArtTypes.BasicCycle:
				case Asc.c_oAscSmartArtTypes.BasicMatrix:
				case Asc.c_oAscSmartArtTypes.BasicBendingProcess:
				case Asc.c_oAscSmartArtTypes.BasicPie:
				case Asc.c_oAscSmartArtTypes.BasicProcess:
				case Asc.c_oAscSmartArtTypes.BasicPyramid:
				case Asc.c_oAscSmartArtTypes.BasicRadial:
				case Asc.c_oAscSmartArtTypes.BasicTarget:
				case Asc.c_oAscSmartArtTypes.BasicVenn:
				case Asc.c_oAscSmartArtTypes.BendingPictureAccentList:
				case Asc.c_oAscSmartArtTypes.BendingPictureBlocks:
				case Asc.c_oAscSmartArtTypes.BendingPictureCaptionList:
				case Asc.c_oAscSmartArtTypes.BendingPictureCaption:
				case Asc.c_oAscSmartArtTypes.BendingPictureSemiTransparentText:
				case Asc.c_oAscSmartArtTypes.BlockCycle:
				case Asc.c_oAscSmartArtTypes.CaptionedPictures:
				case Asc.c_oAscSmartArtTypes.ChevronAccentProcess:
				case Asc.c_oAscSmartArtTypes.CircleArrowProcess:
				case Asc.c_oAscSmartArtTypes.CircleProcess:
				case Asc.c_oAscSmartArtTypes.CircleRelationship:
				case Asc.c_oAscSmartArtTypes.ClosedChevronProcess:
				case Asc.c_oAscSmartArtTypes.ContinuousCycle:
				case Asc.c_oAscSmartArtTypes.ConvergingArrows:
				case Asc.c_oAscSmartArtTypes.ConvergingRadial:
				case Asc.c_oAscSmartArtTypes.ConvergingText:
				case Asc.c_oAscSmartArtTypes.CounterbalanceArrows:
				case Asc.c_oAscSmartArtTypes.CycleMatrix:
				case Asc.c_oAscSmartArtTypes.DescendingProcess:
				case Asc.c_oAscSmartArtTypes.DivergingArrows:
				case Asc.c_oAscSmartArtTypes.DivergingRadial:
				case Asc.c_oAscSmartArtTypes.Equation:
				case Asc.c_oAscSmartArtTypes.Funnel:
				case Asc.c_oAscSmartArtTypes.Gear:
				case Asc.c_oAscSmartArtTypes.GridMatrix:
				case Asc.c_oAscSmartArtTypes.GroupedList:
				case Asc.c_oAscSmartArtTypes.HexagonCluster:
				case Asc.c_oAscSmartArtTypes.IncreasingCircleProcess:
				case Asc.c_oAscSmartArtTypes.InterconnectedBlockProcess:
				case Asc.c_oAscSmartArtTypes.InterconnectedRings:
				case Asc.c_oAscSmartArtTypes.InvertedPyramid:
				case Asc.c_oAscSmartArtTypes.LinearVenn:
				case Asc.c_oAscSmartArtTypes.MultiDirectionalCycle:
				case Asc.c_oAscSmartArtTypes.NestedTarget:
				case Asc.c_oAscSmartArtTypes.NonDirectionalCycle:
				case Asc.c_oAscSmartArtTypes.OpposingArrows:
				case Asc.c_oAscSmartArtTypes.OpposingIdeas:
				case Asc.c_oAscSmartArtTypes.PhasedProcess:
				case Asc.c_oAscSmartArtTypes.PictureAccentProcess:
				case Asc.c_oAscSmartArtTypes.PictureCaptionList:
				case Asc.c_oAscSmartArtTypes.PictureFrame:
				case Asc.c_oAscSmartArtTypes.PictureLineup:
				case Asc.c_oAscSmartArtTypes.PictureOrganizationChart:
				case Asc.c_oAscSmartArtTypes.PictureStrips:
				case Asc.c_oAscSmartArtTypes.PlusAndMinus:
				case Asc.c_oAscSmartArtTypes.ProcessList:
				case Asc.c_oAscSmartArtTypes.RadialCluster:
				case Asc.c_oAscSmartArtTypes.RadialCycle:
				case Asc.c_oAscSmartArtTypes.RadialList:
				case Asc.c_oAscSmartArtTypes.RadialPictureList:
				case Asc.c_oAscSmartArtTypes.RadialVenn:
				case Asc.c_oAscSmartArtTypes.RepeatingBendingProcess:
				case Asc.c_oAscSmartArtTypes.ReverseList:
				case Asc.c_oAscSmartArtTypes.SegmentedCycle:
				case Asc.c_oAscSmartArtTypes.SegmentedProcess:
				case Asc.c_oAscSmartArtTypes.SegmentedPyramid:
				case Asc.c_oAscSmartArtTypes.SpiralPicture:
				case Asc.c_oAscSmartArtTypes.StackedVenn:
				case Asc.c_oAscSmartArtTypes.StaggeredProcess:
				case Asc.c_oAscSmartArtTypes.TabList:
				case Asc.c_oAscSmartArtTypes.TabbedArc:
				case Asc.c_oAscSmartArtTypes.TableList:
				case Asc.c_oAscSmartArtTypes.TextCycle:
				case Asc.c_oAscSmartArtTypes.ThemePictureAccent:
				case Asc.c_oAscSmartArtTypes.ThemePictureAlternatingAccent:
				case Asc.c_oAscSmartArtTypes.ThemePictureGrid:
				case Asc.c_oAscSmartArtTypes.TitledMatrix:
				case Asc.c_oAscSmartArtTypes.TitledPictureAccentList:
				case Asc.c_oAscSmartArtTypes.TitledPictureBlocks:
				case Asc.c_oAscSmartArtTypes.TitlePictureLineup:
				case Asc.c_oAscSmartArtTypes.TrapezoidList:
				case Asc.c_oAscSmartArtTypes.UpwardArrow:
				case Asc.c_oAscSmartArtTypes.VerticalPictureAccentList:
				case Asc.c_oAscSmartArtTypes.VerticalPictureList:
				case Asc.c_oAscSmartArtTypes.HorizontalOrganizationChart:
				case Asc.c_oAscSmartArtTypes.HalfCircleOrganizationChart:
				case Asc.c_oAscSmartArtTypes.StackedList:
				case Asc.c_oAscSmartArtTypes.HorizontalHierarchy:
				case Asc.c_oAscSmartArtTypes.HorizontalMultiLevelHierarchy:
				case Asc.c_oAscSmartArtTypes.OrganizationChart:
				case Asc.c_oAscSmartArtTypes.NameAndTitleOrganizationChart:
				case Asc.c_oAscSmartArtTypes.CirclePictureHierarchy:
				case Asc.c_oAscSmartArtTypes.HierarchyList:
				case Asc.c_oAscSmartArtTypes.Hierarchy:
					return true;
				default:
					return false;
			}
		};
	  SmartArt.prototype.getSizes = function () {
			const sizes = {width: 0, height: 0};
			if (this.drawingBase) {
				const drawingMetrics = this.drawingBase.getGraphicObjectMetrics();
				sizes.width = drawingMetrics.extX;
				sizes.height = drawingMetrics.extY;
			} else {
				sizes.width = this.spPr.xfrm.extX;
				sizes.height = this.spPr.xfrm.extY;
			}
			return sizes;
	  };
		SmartArt.prototype.generateLocalDrawingPart = function () {
			if (this.isLocalDrawingPart) {
				if (!this.isCanGenerateSmartArt()) {
					this.isLocalDrawingPart = false;
					return;
				}
				AscFormat.ExecuteNoHistory(function () {
					this.generateDrawingPart();
					this.isLocalDrawingPart = true;
					this.localDrawingId = this.drawing.Get_Id();
				}, this, []);
			}
		};
	  SmartArt.prototype.reconnectSmartArtShapes = function () {
			if (!this.dataModel) {
				return;
			}
      this.initSmartArtAlgorithm();
			if (this.isCanGenerateSmartArt()) {
				this.smartArtTree.connectShapeSmartArtInfo();
			} else {
				this.setConnections2();
			}
	  };
		SmartArt.prototype.checkDrawingPartWithHistory = function (handleShape) {
			if (this.worksheet && this.isLocalDrawingPart && AscCommon.History.CanAddChanges()) {
				this.isLocalDrawingPart = false;
				const oldDrawing = this.drawing;
				const mainGroup = this.getMainGroup();
				const selectedObjects = mainGroup.selectedObjects;
				const mapShapes = {};
				for (let i = 0; i < oldDrawing.spTree.length; i += 1) {
					const shape = oldDrawing.spTree[i];
					mapShapes[shape.GetId()] = {index: i, state: null};
					const docContent = shape.getDocContent();
					if (docContent) {
						mapShapes[shape.GetId()].state = docContent.GetSelectionState();
					}
				}

				const copyDrawing = this.drawing.copy();
				copyDrawing.setWorksheet(this.drawing.worksheet);
				copyDrawing.setDrawingObjects(this.drawing.drawingObjects);
				this.removeFromSpTreeByPos(0);
				this.addToSpTree(0, copyDrawing);
				this.setDrawing(copyDrawing);
				this.reconnectSmartArtShapes();
				copyDrawing.recalculate();
				for (let i = 0; i < selectedObjects.length; i += 1) {
					const shapeInfo = mapShapes[selectedObjects[i].Id];
					if (shapeInfo) {
						copyDrawing.spTree[shapeInfo.index].selectStartPage = selectedObjects[i].selectStartPage;
						selectedObjects[i] = copyDrawing.spTree[shapeInfo.index];
						selectedObjects[i].selected = true;
						if (shapeInfo.state) {
							const docContent = selectedObjects[i].getDocContent();
							docContent && docContent.SetSelectionState(shapeInfo.state);
						}
					}
				}
				if (mainGroup.selection.textSelection) {
					const info = mapShapes[mainGroup.selection.textSelection.GetId()];
					if (info) {
						mainGroup.selection.textSelection = copyDrawing.spTree[info.index];
					}
				}
				if (handleShape) {
					const info = mapShapes[handleShape.GetId()];
					if (info) {
						return copyDrawing.spTree[info.index];
					}
				}
			}
		};
	  SmartArt.prototype.initDrawing = function () {
			if (!this.drawing) {
				this.drawing = new Drawing();
				this.drawing.setBDeleted(false);
				this.addToSpTree(0, this.drawing);
				const spPr = new AscFormat.CSpPr();
				this.drawing.setSpPr(spPr);
				spPr.setParent(this.drawing);
				const nvSpPr = new AscFormat.UniNvPr();
				nvSpPr.cNvPr.setId(0);
				this.drawing.setNvSpPr(nvSpPr);
			}
	  }
    SmartArt.prototype.generateDrawingPart = function () {
	    this.isLocalDrawingPart = false;
			if (!this.isCanGenerateSmartArt()) {
				return;
			}
      this.initSmartArtAlgorithm();
	    this.smartArtTree.startFromBegin();
      const drawing = this.getDrawing();
      const shapeLength = drawing.spTree.length;
      for (let i = 0; i < shapeLength; i++) {
        drawing.removeFromSpTreeByPos(0);
      }
      const shapes = this.smartArtTree.getShapes();
      for (let i = shapes.length - 1; i >= 0; i -= 1) {
        drawing.addToSpTree(0, shapes[i]);
      }
			this.recalcFitFontSize();
    };
	  SmartArt.prototype.generateSmartArtDrawingPart = SmartArt.prototype.generateDrawingPart;
	  SmartArt.prototype.findConnector = function (x, y) {
		  return null;
	  };
	  SmartArt.prototype.getAllRasterImages = function (arrImages)
	  {
			const oBgFormat = this.getBg();
			if (oBgFormat)
			{
				if (oBgFormat.fill && oBgFormat.fill.fill && typeof (oBgFormat.fill.fill.RasterImageId) === "string" && oBgFormat.fill.fill.RasterImageId.length > 0)
					arrImages.push(oBgFormat.fill.fill.RasterImageId);
			}
		  CGroupShape.prototype.getAllRasterImages.call(this, arrImages);
	  };

    SmartArt.prototype.hasSmartArt = function (bRetSmartArt) {
      return bRetSmartArt ? this : true;
    }

    SmartArt.prototype.getContrastDrawing = function () {
      const arrShapes = this.spTree[0] && this.spTree[0].spTree;
      if (arrShapes) {
        arrShapes.forEach(function (oShape) {
          if (oShape.spPr) {
            if (oShape.spPr.Fill && oShape.spPr.Fill.fill && !(oShape.spPr.Fill.fill instanceof AscFormat.CNoFill)) {
              const id = oShape.spPr.Fill.fill.color.color.id;
              let standardColor;
              if (id === 0) {
                standardColor = {R: 0x5B, G: 0x9B, B: 0xD5, A: 255};
              } else if (id === 12) {
                standardColor = {R: 255, G: 255, B: 255, A: 255};
              } else {
                standardColor = {R: 0, G: 0, B: 0, A: 255};
              }

              if (oShape.spPr.Fill.fill.color.Mods) {
                oShape.spPr.Fill.fill.color.Mods.Apply(standardColor);
              }
              const grayscaleValue = AscFormat.getGrayscaleValue(standardColor);

              if (grayscaleValue < GRAYSCALE_TRESHHOLD) {
                const oHeavyBrush = new AscFormat.CreateSolidFillRGB(211, 211, 211);
                oShape.spPr.setFill(oHeavyBrush);
              } else {
                const oLightBrush = new AscFormat.CreateSolidFillRGB(255, 255, 255);
                oShape.spPr.setFill(oLightBrush);
              }
              if (oShape.spPr.ln) {
                const oPen = new AscFormat.CreateSolidFillRGB(0, 0, 0);
                oShape.spPr.ln.setFill(oPen);
                oShape.spPr.ln.setW(12700 * 3);
              }
            } else if (oShape.spPr.ln && oShape.spPr.ln.Fill && oShape.spPr.ln.Fill.fill && !(oShape.spPr.ln.Fill.fill instanceof AscFormat.CNoFill)) {
              const oPen = new AscFormat.CreateSolidFillRGB(0, 0, 0);
              oShape.spPr.ln.setFill(oPen);
              oShape.spPr.ln.setW(12700 * 3);
            }
          }
        });
      }
    }

    SmartArt.prototype.recalculate = function () {
      if(this.bDeleted)
        return;
      AscFormat.ExecuteNoHistory(function () {
        var oldParaMarks = editor && editor.ShowParaMarks;
        if (oldParaMarks) {
          editor.ShowParaMarks = false;
        }
				if (this.recalcInfo.reconnectSmartArtShapes) {
					this.recalcInfo.reconnectSmartArtShapes = false;
					this.reconnectSmartArtShapes();
				}
        CGroupShape.prototype.recalculate.call(this);
				if (this.recalcInfo.fitFontSize) {
					this.recalcInfo.fitFontSize = false;
					this.fitFontSize();
				}
        if (oldParaMarks) {
          editor.ShowParaMarks = oldParaMarks;
        }
      }, this, []);
    }

    SmartArt.prototype.decorateParaDrawing = function (drawingObjects) {
      var drawing = new ParaDrawing(this.spPr.xfrm.extX, this.spPr.xfrm.extY, this, drawingObjects.drawingDocument, drawingObjects.document, null);
      drawing.setExtent(this.spPr.xfrm.extX, this.spPr.xfrm.extY);
      drawing.Set_GraphicObject(this);
      this.setParent(drawing);
      drawing.CheckWH();
      return drawing;
    };

    SmartArt.prototype.fillByPreset = function (nSmartArtType, bLoadOnlyDrawing) {
      const oApi = Asc.editor || editor;
			const drawingInfo = AscCommon.g_oBinarySmartArts.getDrawingInfo(nSmartArtType);
			const dataBin = AscCommon.g_oBinarySmartArts.getDataBinary(nSmartArtType);
      if (oApi && drawingInfo && (bLoadOnlyDrawing || dataBin)) {
        const oDrawingDocument = oApi.getDrawingDocument();
        const oLogicDocument = oApi.getLogicDocument();

        const pReader = new AscCommon.BinaryPPTYLoader();
	      pReader.IsFillingSmartArt = true;
	      pReader.presentation = oLogicDocument;
	      pReader.DrawingDocument = oDrawingDocument;

        pReader.stream = new AscCommon.FileStream(drawingInfo.bin, drawingInfo.bin.length);
        pReader.stream.cur = drawingInfo.pos;
        this.readChild(pReader.stream.GetUChar(), pReader);

	      if (!bLoadOnlyDrawing)
	      {
		      pReader.stream = new AscCommon.FileStream(dataBin, dataBin.length);
		      this.readChild(pReader.stream.GetUChar(), pReader);
		      this.readChild(pReader.stream.GetUChar(), pReader);
		      this.readChild(pReader.stream.GetUChar(), pReader);
		      this.readChild(pReader.stream.GetUChar(), pReader);
		      this.checkNodePointsAfterRead(true);
	      }
        this.setSpPr(new AscFormat.CSpPr());
        this.spPr.setParent(this);
        const smXfrm = new AscFormat.CXfrm();
        smXfrm.fillStandardSmartArtXfrm();
        this.spPr.setXfrm(smXfrm);
        this.setBDeleted2(false);
        this.x = smXfrm.offX;
        this.y = smXfrm.offY;
        this.extX = smXfrm.extX;
        this.extY = smXfrm.extY;
        this.drawing.setXfrmByParent();

	      pReader.IsFillingSmartArt = false;
      }
      return this;
    }

	  SmartArt.prototype.fitToPageSize = function ()
	  {
		  const oApi = Asc.editor || editor;
		  if (oApi)
		  {
			  const bFromWord = oApi.isDocumentEditor;
			  if (bFromWord)
			  {
				  let W;
				  let H;
				  const logicDocument = oApi.getLogicDocument();
				  const oColumnSize = logicDocument.GetColumnSize();
				  if (oColumnSize)
				  {
					  W = oColumnSize.W;
					  H = oColumnSize.H;
				  }
				  else
				  {
					  W = AscCommon.Page_Width - (AscCommon.X_Left_Margin + AscCommon.X_Right_Margin);
					  H = AscCommon.Page_Height - (AscCommon.Y_Top_Margin + AscCommon.Y_Bottom_Margin);
				  }
				  this.fitForSizes(H, W);
			  }
		  }
	  };
	  SmartArt.prototype.fitForSizes = function (nFitHeight, nFitWidth) {
		  const cH = nFitWidth / this.extX;
		  const cW = nFitHeight / this.extY;
			  const minCoefficient = Math.min(cH, cW);
			  this.changeSize(minCoefficient, minCoefficient);
	  };
    SmartArt.prototype.fitFontSize = function () {
      this.spTree[0] && this.spTree[0].spTree.forEach(function (oShape) {
        oShape.recalculateContentWitCompiledPr();
        oShape.setTruthFontSizeInSmartArt();
        oShape.recalculateContentWitCompiledPr();
      });
    };

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

    SmartArt.prototype.getPtMap = function () {
      const data = this.getDataModelFromData();
      var ptLst = data && data.getPtLst();
      if (ptLst) {
        return ptLst.getPtMap();
      }
    };

    SmartArt.prototype.getDefColorsByName = function () {
      var colorsDef = this.getColorsDef();
      return colorsDef && colorsDef.styleLbl;
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
      const smartArtInfo = shape.getSmartArtInfo();
			if (!smartArtInfo) {
				return;
			}
			let textPoint;
			if (this.isCanGenerateSmartArt()) {
				const contentNodes = smartArtInfo.contentPoint;
				if (contentNodes.length) {
					const textNode = contentNodes[0].getTextNode();
					textPoint = textNode && textNode.presPoint;
				}
			} else {
				textPoint = smartArtInfo.shapePoint;
			}
      var defaultTxColorFromShape = textPoint && this.getDefaultTxColorFromPoint(textPoint);
      var defaultTxFill;

      if (defaultTxColorFromShape) {
        defaultTxFill = AscFormat.CreateUniFillByUniColorCopy(defaultTxColorFromShape);
      }
      return defaultTxFill;
    }

    SmartArt.prototype.getTypeOfSmartArt = function () {
	    const layoutDef = this.getLayoutDef();
      var type;
			if (layoutDef) {
				var typeSplit = layoutDef.uniqueId.split('/');
				type = typeSplit[typeSplit.length - 1].split('#')[0];
			}

      switch (type) {
        case "AccentedPicture": {
          return Asc.c_oAscSmartArtTypes.AccentedPicture;
        }
        case "balance1": {
          return Asc.c_oAscSmartArtTypes.Balance;
        }
        case "TitledPictureBlocks": {
          return Asc.c_oAscSmartArtTypes.TitledPictureBlocks;
        }
        case "PictureAccentBlocks": {
          return Asc.c_oAscSmartArtTypes.PictureAccentBlocks;
        }
        case "cycle5": {
          return Asc.c_oAscSmartArtTypes.BlockCycle;
        }
        case "venn2": {
          return Asc.c_oAscSmartArtTypes.StackedVenn;
        }
        case "equation2": {
          return Asc.c_oAscSmartArtTypes.VerticalEquation;
        }
        case "vList5": {
          return Asc.c_oAscSmartArtTypes.VerticalBlockList;
        }
        case "bProcess4": {
          return Asc.c_oAscSmartArtTypes.VerticalBendingProcess;
        }
        case "vList2": {
          return Asc.c_oAscSmartArtTypes.VerticalBulletList;
        }
        case "VerticalCurvedList": {
          return Asc.c_oAscSmartArtTypes.VerticalCurvedList;
        }
        case "process2": {
          return Asc.c_oAscSmartArtTypes.VerticalProcess;
        }
        case "list1": {
          return Asc.c_oAscSmartArtTypes.VerticalBoxList;
        }
        case "vList4": {
          return Asc.c_oAscSmartArtTypes.VerticalPictureList;
        }
        case "VerticalCircleList": {
          return Asc.c_oAscSmartArtTypes.VerticalCircleList;
        }
        case "vList3": {
          return Asc.c_oAscSmartArtTypes.VerticalPictureAccentList;
        }
        case "vList6": {
          return Asc.c_oAscSmartArtTypes.VerticalArrowList;
        }
        case "chevron2": {
          return Asc.c_oAscSmartArtTypes.VerticalChevronList;
        }
        case "VerticalAccentList": {
          return Asc.c_oAscSmartArtTypes.VerticalAccentList;
        }
        case "target2": {
          return Asc.c_oAscSmartArtTypes.NestedTarget;
        }
        case "funnel1": {
          return Asc.c_oAscSmartArtTypes.Funnel;
        }
        case "arrow2": {
          return Asc.c_oAscSmartArtTypes.UpwardArrow;
        }
        case "IncreasingArrowsProcess": {
          return Asc.c_oAscSmartArtTypes.IncreasingArrowsProcess;
        }
        case "StepUpProcess": {
          return Asc.c_oAscSmartArtTypes.StepUpProcess;
        }
        case "CircularPictureCallout": {
          return Asc.c_oAscSmartArtTypes.CircularPictureCallout;
        }
        case "hierarchy2": {
          return Asc.c_oAscSmartArtTypes.HorizontalHierarchy;
        }
        case "hierarchy5": {
          return Asc.c_oAscSmartArtTypes.HorizontalLabeledHierarchy;
        }
        case "HorizontalMultiLevelHierarchy": {
          return Asc.c_oAscSmartArtTypes.HorizontalMultiLevelHierarchy;
        }
        case "HorizontalOrganizationChart": {
          return Asc.c_oAscSmartArtTypes.HorizontalOrganizationChart;
        }
        case "hList1": {
          return Asc.c_oAscSmartArtTypes.HorizontalBulletList;
        }
        case "pList2": {
          return Asc.c_oAscSmartArtTypes.HorizontalPictureList;
        }
        case "hChevron3": {
          return Asc.c_oAscSmartArtTypes.ClosedChevronProcess;
        }
        case "hierarchy3": {
          return Asc.c_oAscSmartArtTypes.HierarchyList;
        }
        case "hierarchy1": {
          return Asc.c_oAscSmartArtTypes.Hierarchy;
        }
        case "CirclePictureHierarchy": {
          return Asc.c_oAscSmartArtTypes.CirclePictureHierarchy;
        }
        case "hierarchy6": {
          return Asc.c_oAscSmartArtTypes.LabeledHierarchy;
        }
        case "pyramid3": {
          return Asc.c_oAscSmartArtTypes.InvertedPyramid;
        }
        case "HexagonCluster": {
          return Asc.c_oAscSmartArtTypes.HexagonCluster;
        }
        case "CircleRelationship": {
          return Asc.c_oAscSmartArtTypes.CircleRelationship;
        }
        case "CircleAccentTimeline": {
          return Asc.c_oAscSmartArtTypes.CircleAccentTimeline;
        }
        case "bProcess2": {
          return Asc.c_oAscSmartArtTypes.CircularBendingProcess;
        }
        case "arrow6": {
          return Asc.c_oAscSmartArtTypes.ArrowRibbon;
        }
        case "venn3": {
          return Asc.c_oAscSmartArtTypes.LinearVenn;
        }
        case "PictureLineup": {
          return Asc.c_oAscSmartArtTypes.PictureLineup;
        }
        case "TitlePictureLineup": {
          return Asc.c_oAscSmartArtTypes.TitlePictureLineup;
        }
        case "BendingPictureCaptionList": {
          return Asc.c_oAscSmartArtTypes.BendingPictureCaptionList;
        }
        case "bList2": {
          return Asc.c_oAscSmartArtTypes.BendingPictureAccentList;
        }
        case "matrix1": {
          return Asc.c_oAscSmartArtTypes.TitledMatrix;
        }
        case "IncreasingCircleProcess": {
          return Asc.c_oAscSmartArtTypes.IncreasingCircleProcess;
        }
        case "BendingPictureBlocks": {
          return Asc.c_oAscSmartArtTypes.BendingPictureBlocks;
        }
        case "BendingPictureCaption": {
          return Asc.c_oAscSmartArtTypes.BendingPictureCaption;
        }
        case "BendingPictureSemiTransparentText": {
          return Asc.c_oAscSmartArtTypes.BendingPictureSemiTransparentText;
        }
        case "cycle6": {
          return Asc.c_oAscSmartArtTypes.NonDirectionalCycle;
        }
        case "hProcess9": {
          return Asc.c_oAscSmartArtTypes.ContinuousBlockProcess;
        }
        case "hList7": {
          return Asc.c_oAscSmartArtTypes.ContinuousPictureList;
        }
        case "cycle3": {
          return Asc.c_oAscSmartArtTypes.ContinuousCycle;
        }
        case "BlockDescendingList": {
          return Asc.c_oAscSmartArtTypes.DescendingBlockList;
        }
        case "StepDownProcess": {
          return Asc.c_oAscSmartArtTypes.StepDownProcess;
        }
        case "ReverseList": {
          return Asc.c_oAscSmartArtTypes.ReverseList;
        }
        case "orgChart1": {
          return Asc.c_oAscSmartArtTypes.OrganizationChart;
        }
        case "NameandTitleOrganizationalChart": {
          return Asc.c_oAscSmartArtTypes.NameAndTitleOrganizationChart;
        }
        case "hProcess4": {
          return Asc.c_oAscSmartArtTypes.AlternatingFlow;
        }
        case "pyramid2": {
          return Asc.c_oAscSmartArtTypes.PyramidList;
        }
        case "PlusandMinus": {
          return Asc.c_oAscSmartArtTypes.PlusAndMinus;
        }
        case "bProcess3": {
          return Asc.c_oAscSmartArtTypes.RepeatingBendingProcess;
        }
        case "CaptionedPictures": {
          return Asc.c_oAscSmartArtTypes.CaptionedPictures;
        }
        case "hProcess7": {
          return Asc.c_oAscSmartArtTypes.DetailedProcess;
        }
        case "PictureStrips": {
          return Asc.c_oAscSmartArtTypes.PictureStrips;
        }
        case "HalfCircleOrganizationChart": {
          return Asc.c_oAscSmartArtTypes.HalfCircleOrganizationChart;
        }
        case "PhasedProcess": {
          return Asc.c_oAscSmartArtTypes.PhasedProcess;
        }
        case "venn1": {
          return Asc.c_oAscSmartArtTypes.BasicVenn;
        }
        case "hProcess11": {
          return Asc.c_oAscSmartArtTypes.BasicTimeline;
        }
        case "chart3": {
          return Asc.c_oAscSmartArtTypes.BasicPie;
        }
        case "matrix3": {
          return Asc.c_oAscSmartArtTypes.BasicMatrix;
        }
        case "pyramid1": {
          return Asc.c_oAscSmartArtTypes.BasicPyramid;
        }
        case "radial1": {
          return Asc.c_oAscSmartArtTypes.BasicRadial;
        }
        case "target1": {
          return Asc.c_oAscSmartArtTypes.BasicTarget;
        }
        case "default": {
          return Asc.c_oAscSmartArtTypes.BasicBlockList;
        }
        case "process5": {
          return Asc.c_oAscSmartArtTypes.BasicBendingProcess;
        }
        case "process1": {
          return Asc.c_oAscSmartArtTypes.BasicProcess;
        }
        case "chevron1": {
          return Asc.c_oAscSmartArtTypes.BasicChevronProcess;
        }
        case "cycle2": {
          return Asc.c_oAscSmartArtTypes.BasicCycle;
        }
        case "OpposingIdeas": {
          return Asc.c_oAscSmartArtTypes.OpposingIdeas;
        }
        case "arrow4": {
          return Asc.c_oAscSmartArtTypes.OpposingArrows;
        }
        case "RandomtoResultProcess": {
          return Asc.c_oAscSmartArtTypes.RandomToResultProcess;
        }
        case "SubStepProcess": {
          return Asc.c_oAscSmartArtTypes.SubStepProcess;
        }
        case "PieProcess": {
          return Asc.c_oAscSmartArtTypes.PieProcess;
        }
        case "process3": {
          return Asc.c_oAscSmartArtTypes.AccentProcess;
        }
        case "AscendingPictureAccentProcess": {
          return Asc.c_oAscSmartArtTypes.AscendingPictureAccentProcess;
        }
        case "hProcess10": {
          return Asc.c_oAscSmartArtTypes.PictureAccentProcess;
        }
        case "radial3": {
          return Asc.c_oAscSmartArtTypes.RadialVenn;
        }
        case "radial6": {
          return Asc.c_oAscSmartArtTypes.RadialCycle;
        }
        case "RadialCluster": {
          return Asc.c_oAscSmartArtTypes.RadialCluster;
        }
        case "radial2": {
          return Asc.c_oAscSmartArtTypes.RadialList;
        }
        case "cycle7": {
          return Asc.c_oAscSmartArtTypes.MultiDirectionalCycle;
        }
        case "radial5": {
          return Asc.c_oAscSmartArtTypes.DivergingRadial;
        }
        case "arrow1": {
          return Asc.c_oAscSmartArtTypes.DivergingArrows;
        }
        case "FramedTextPicture": {
          return Asc.c_oAscSmartArtTypes.FramedTextPicture;
        }
        case "lProcess2": {
          return Asc.c_oAscSmartArtTypes.GroupedList;
        }
        case "pyramid4": {
          return Asc.c_oAscSmartArtTypes.SegmentedPyramid;
        }
        case "process4": {
          return Asc.c_oAscSmartArtTypes.SegmentedProcess;
        }
        case "cycle8": {
          return Asc.c_oAscSmartArtTypes.SegmentedCycle;
        }
        case "PictureGrid": {
          return Asc.c_oAscSmartArtTypes.PictureGrid;
        }
        case "matrix2": {
          return Asc.c_oAscSmartArtTypes.GridMatrix;
        }
        case "SpiralPicture": {
          return Asc.c_oAscSmartArtTypes.SpiralPicture;
        }
        case "hList9": {
          return Asc.c_oAscSmartArtTypes.StackedList;
        }
        case "pList1": {
          return Asc.c_oAscSmartArtTypes.PictureCaptionList;
        }
        case "lProcess1": {
          return Asc.c_oAscSmartArtTypes.ProcessList;
        }
        case "BubblePictureList": {
          return Asc.c_oAscSmartArtTypes.BubblePictureList;
        }
        case "SquareAccentList": {
          return Asc.c_oAscSmartArtTypes.SquareAccentList;
        }
        case "LinedList": {
          return Asc.c_oAscSmartArtTypes.LinedList;
        }
        case "hList2": {
          return Asc.c_oAscSmartArtTypes.PictureAccentList;
        }
        case "PictureAccentList": {
          return Asc.c_oAscSmartArtTypes.TitledPictureAccentList;
        }
        case "SnapshotPictureList": {
          return Asc.c_oAscSmartArtTypes.SnapshotPictureList;
        }
        case "hProcess3": {
          return Asc.c_oAscSmartArtTypes.ContinuousArrowProcess;
        }
        case "CircleArrowProcess": {
          return Asc.c_oAscSmartArtTypes.CircleArrowProcess;
        }
        case "hProcess6": {
          return Asc.c_oAscSmartArtTypes.ProcessArrows;
        }
        case "vProcess5": {
          return Asc.c_oAscSmartArtTypes.StaggeredProcess;
        }
        case "radial4": {
          return Asc.c_oAscSmartArtTypes.ConvergingRadial;
        }
        case "arrow5": {
          return Asc.c_oAscSmartArtTypes.ConvergingArrows;
        }
        case "hierarchy4": {
          return Asc.c_oAscSmartArtTypes.TableHierarchy;
        }
        case "hList3": {
          return Asc.c_oAscSmartArtTypes.TableList;
        }
        case "cycle1": {
          return Asc.c_oAscSmartArtTypes.TextCycle;
        }
        case "hList6": {
          return Asc.c_oAscSmartArtTypes.TrapezoidList;
        }
        case "DescendingProcess": {
          return Asc.c_oAscSmartArtTypes.DescendingProcess;
        }
        case "lProcess3": {
          return Asc.c_oAscSmartArtTypes.ChevronList;
        }
        case "equation1": {
          return Asc.c_oAscSmartArtTypes.Equation;
        }
        case "arrow3": {
          return Asc.c_oAscSmartArtTypes.CounterbalanceArrows;
        }
        case "target3": {
          return Asc.c_oAscSmartArtTypes.TargetList;
        }
        case "cycle4": {
          return Asc.c_oAscSmartArtTypes.CycleMatrix;
        }
        case "AlternatingPictureBlocks": {
          return Asc.c_oAscSmartArtTypes.AlternatingPictureBlocks;
        }
        case "AlternatingPictureCircles": {
          return Asc.c_oAscSmartArtTypes.AlternatingPictureCircles;
        }
        case "AlternatingHexagons": {
          return Asc.c_oAscSmartArtTypes.AlternatingHexagonList;
        }
        case "gear1": {
          return Asc.c_oAscSmartArtTypes.Gear;
        }
        case "architecture": {
          return Asc.c_oAscSmartArtTypes.ArchitectureLayout;
        }
        case "chevronAccent+Icon": {
          return Asc.c_oAscSmartArtTypes.ChevronAccentProcess;
        }
        case "CircleProcess": {
          return Asc.c_oAscSmartArtTypes.CircleProcess;
        }
        case "ConvergingText": {
          return Asc.c_oAscSmartArtTypes.ConvergingText;
        }
        case "HexagonRadial": {
          return Asc.c_oAscSmartArtTypes.HexagonRadial;
        }
        case "InterconnectedBlockProcess": {
          return Asc.c_oAscSmartArtTypes.InterconnectedBlockProcess;
        }
        case "rings+Icon": {
          return Asc.c_oAscSmartArtTypes.InterconnectedRings;
        }
        case "Picture Frame": {
          return Asc.c_oAscSmartArtTypes.PictureFrame;
        }
        case "pictureOrgChart+Icon": {
          return Asc.c_oAscSmartArtTypes.PictureOrganizationChart;
        }
        case "RadialPictureList": {
          return Asc.c_oAscSmartArtTypes.RadialPictureList;
        }
        case "TabList": {
          return Asc.c_oAscSmartArtTypes.TabList;
        }
        case "TabbedArc+Icon": {
          return Asc.c_oAscSmartArtTypes.TabbedArc;
        }
        case "ThemePictureAccent": {
          return Asc.c_oAscSmartArtTypes.ThemePictureAccent;
        }
        case "ThemePictureAlternatingAccent": {
          return Asc.c_oAscSmartArtTypes.ThemePictureAlternatingAccent;
        }
        case "ThemePictureGrid": {
          return Asc.c_oAscSmartArtTypes.ThemePictureGrid;
        }
        case "VaryingWidthList": {
          return Asc.c_oAscSmartArtTypes.VaryingWidthList;
        }
        case "BracketList": {
          return Asc.c_oAscSmartArtTypes.VerticalBracketList;
        }
        default: {
          return type;
        }
      }
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
        case Asc.c_oAscSmartArtTypes.PictureAccentBlocks:
        case Asc.c_oAscSmartArtTypes.BlockCycle:
        case Asc.c_oAscSmartArtTypes.StackedVenn:
        case Asc.c_oAscSmartArtTypes.VerticalBendingProcess:
        case Asc.c_oAscSmartArtTypes.VerticalBulletList:
        case Asc.c_oAscSmartArtTypes.VerticalCurvedList:
        case Asc.c_oAscSmartArtTypes.VerticalProcess:
        case Asc.c_oAscSmartArtTypes.VerticalBoxList:
        case Asc.c_oAscSmartArtTypes.VerticalPictureList:
        case Asc.c_oAscSmartArtTypes.VerticalCircleList:
        case Asc.c_oAscSmartArtTypes.UpwardArrow:
        case Asc.c_oAscSmartArtTypes.StepUpProcess:
        case Asc.c_oAscSmartArtTypes.HorizontalHierarchy:
        case Asc.c_oAscSmartArtTypes.HorizontalMultiLevelHierarchy:
        case Asc.c_oAscSmartArtTypes.HorizontalOrganizationChart:
        case Asc.c_oAscSmartArtTypes.HorizontalBulletList:
        case Asc.c_oAscSmartArtTypes.HorizontalPictureList:
        case Asc.c_oAscSmartArtTypes.ClosedChevronProcess:
        case Asc.c_oAscSmartArtTypes.Hierarchy:
        case Asc.c_oAscSmartArtTypes.CirclePictureHierarchy:
        case Asc.c_oAscSmartArtTypes.HexagonCluster:
        case Asc.c_oAscSmartArtTypes.CircleRelationship:
        case Asc.c_oAscSmartArtTypes.CircleAccentTimeline:
        case Asc.c_oAscSmartArtTypes.CircularBendingProcess:
        case Asc.c_oAscSmartArtTypes.ArrowRibbon:
        case Asc.c_oAscSmartArtTypes.LinearVenn:
        case Asc.c_oAscSmartArtTypes.PictureLineup:
        case Asc.c_oAscSmartArtTypes.BendingPictureCaptionList:
        case Asc.c_oAscSmartArtTypes.TitledMatrix:
        case Asc.c_oAscSmartArtTypes.BendingPictureBlocks:
        case Asc.c_oAscSmartArtTypes.BendingPictureCaption:
        case Asc.c_oAscSmartArtTypes.BendingPictureSemiTransparentText:
        case Asc.c_oAscSmartArtTypes.NonDirectionalCycle:
        case Asc.c_oAscSmartArtTypes.ContinuousBlockProcess:
        case Asc.c_oAscSmartArtTypes.ContinuousPictureList:
        case Asc.c_oAscSmartArtTypes.ContinuousCycle:
        case Asc.c_oAscSmartArtTypes.StepDownProcess:
        case Asc.c_oAscSmartArtTypes.ReverseList:
        case Asc.c_oAscSmartArtTypes.OrganizationChart:
        case Asc.c_oAscSmartArtTypes.PictureOrganizationChart:
        case Asc.c_oAscSmartArtTypes.PyramidList:
        case Asc.c_oAscSmartArtTypes.PlusAndMinus:
        case Asc.c_oAscSmartArtTypes.RepeatingBendingProcess:
        case Asc.c_oAscSmartArtTypes.CaptionedPictures:
        case Asc.c_oAscSmartArtTypes.PictureStrips:
        case Asc.c_oAscSmartArtTypes.HalfCircleOrganizationChart:
        case Asc.c_oAscSmartArtTypes.BasicVenn:
        case Asc.c_oAscSmartArtTypes.BasicTimeline:
        case Asc.c_oAscSmartArtTypes.BasicPie:
        case Asc.c_oAscSmartArtTypes.BasicMatrix:
        case Asc.c_oAscSmartArtTypes.BasicTarget:
        case Asc.c_oAscSmartArtTypes.BasicBlockList:
        case Asc.c_oAscSmartArtTypes.BasicBendingProcess:
        case Asc.c_oAscSmartArtTypes.BasicProcess:
        case Asc.c_oAscSmartArtTypes.BasicChevronProcess:
        case Asc.c_oAscSmartArtTypes.BasicCycle:
        case Asc.c_oAscSmartArtTypes.OpposingArrows:
        case Asc.c_oAscSmartArtTypes.RandomToResultProcess:
        case Asc.c_oAscSmartArtTypes.AccentProcess:
        case Asc.c_oAscSmartArtTypes.PictureAccentProcess:
        case Asc.c_oAscSmartArtTypes.RadialCycle:
        case Asc.c_oAscSmartArtTypes.MultiDirectionalCycle:
        case Asc.c_oAscSmartArtTypes.DivergingArrows:
        case Asc.c_oAscSmartArtTypes.FramedTextPicture:
        case Asc.c_oAscSmartArtTypes.SegmentedPyramid:
        case Asc.c_oAscSmartArtTypes.SegmentedCycle:
        case Asc.c_oAscSmartArtTypes.PictureGrid:
        case Asc.c_oAscSmartArtTypes.GridMatrix:
        case Asc.c_oAscSmartArtTypes.SpiralPicture:
        case Asc.c_oAscSmartArtTypes.PictureCaptionList:
        case Asc.c_oAscSmartArtTypes.BubblePictureList:
        case Asc.c_oAscSmartArtTypes.SnapshotPictureList:
        case Asc.c_oAscSmartArtTypes.ContinuousArrowProcess:
        case Asc.c_oAscSmartArtTypes.CircleArrowProcess:
        case Asc.c_oAscSmartArtTypes.StaggeredProcess:
        case Asc.c_oAscSmartArtTypes.ConvergingRadial:
        case Asc.c_oAscSmartArtTypes.ConvergingArrows:
        case Asc.c_oAscSmartArtTypes.TableHierarchy: //TODO: think about it
        case Asc.c_oAscSmartArtTypes.ArchitectureLayout: //TODO: think about it
        case Asc.c_oAscSmartArtTypes.TextCycle:
        case Asc.c_oAscSmartArtTypes.TrapezoidList:
        case Asc.c_oAscSmartArtTypes.DescendingProcess:
        case Asc.c_oAscSmartArtTypes.Equation:
        case Asc.c_oAscSmartArtTypes.CounterbalanceArrows:
        case Asc.c_oAscSmartArtTypes.AlternatingPictureBlocks:
        case Asc.c_oAscSmartArtTypes.AlternatingPictureCircles:
        case Asc.c_oAscSmartArtTypes.ChevronAccentProcess:
        case Asc.c_oAscSmartArtTypes.TabbedArc:
        case Asc.c_oAscSmartArtTypes.ThemePictureAccent:
        case Asc.c_oAscSmartArtTypes.VaryingWidthList:
        case Asc.c_oAscSmartArtTypes.InterconnectedRings:
        case Asc.c_oAscSmartArtTypes.ThemePictureAlternatingAccent:
        case Asc.c_oAscSmartArtTypes.HexagonRadial:
        case Asc.c_oAscSmartArtTypes.PictureFrame:
        case Asc.c_oAscSmartArtTypes.TabList:
        case Asc.c_oAscSmartArtTypes.VerticalBracketList:
        case Asc.c_oAscSmartArtTypes.Gear: {
          return shapes;
        }
        case Asc.c_oAscSmartArtTypes.AlternatingHexagonList:
          return getShapesFromPresetGeom(['rect']);
        case Asc.c_oAscSmartArtTypes.LinedList:
          return getShapesFromPresName(['tx1', 'tx2', 'tx3', 'tx4']);
        case Asc.c_oAscSmartArtTypes.SquareAccentList:
        case Asc.c_oAscSmartArtTypes.IncreasingCircleProcess:
        case Asc.c_oAscSmartArtTypes.PieProcess:
          return getShapesFromPresName(['Child', 'Parent']);
        case Asc.c_oAscSmartArtTypes.PictureAccentList:
          return getShapesFromPresStyleLbl(['node1', 'revTx']);
        case Asc.c_oAscSmartArtTypes.GroupedList:// TODO: check transform
          return getShapesFromPresStyleLbl(['bgShp', 'node1']);
        case Asc.c_oAscSmartArtTypes.InterconnectedBlockProcess:
          return getShapesFromPresStyleLbl(['alignImgPlace1', 'node1']);
        case Asc.c_oAscSmartArtTypes.TitledPictureAccentList:
          return getShapesFromPresStyleLbl(['lnNode1']); // TODO: think about it
        case Asc.c_oAscSmartArtTypes.VerticalBlockList:
        case Asc.c_oAscSmartArtTypes.VerticalChevronList:
        case Asc.c_oAscSmartArtTypes.BendingPictureAccentList:
        case Asc.c_oAscSmartArtTypes.StackedList:
        case Asc.c_oAscSmartArtTypes.DetailedProcess:
        case Asc.c_oAscSmartArtTypes.VerticalArrowList:
        case Asc.c_oAscSmartArtTypes.ProcessArrows:
        case Asc.c_oAscSmartArtTypes.SubStepProcess:
        case Asc.c_oAscSmartArtTypes.Funnel:
        case Asc.c_oAscSmartArtTypes.PhasedProcess:
        case Asc.c_oAscSmartArtTypes.CycleMatrix:
        case Asc.c_oAscSmartArtTypes.BasicPyramid:
        case Asc.c_oAscSmartArtTypes.InvertedPyramid:
        case Asc.c_oAscSmartArtTypes.VerticalPictureAccentList:
        case Asc.c_oAscSmartArtTypes.RadialList:
        case Asc.c_oAscSmartArtTypes.TitledPictureBlocks:
        case Asc.c_oAscSmartArtTypes.OpposingIdeas:
        case Asc.c_oAscSmartArtTypes.LabeledHierarchy:
        case Asc.c_oAscSmartArtTypes.RadialPictureList:
        case Asc.c_oAscSmartArtTypes.ConvergingText:
        case Asc.c_oAscSmartArtTypes.CircleProcess:
        case Asc.c_oAscSmartArtTypes.HorizontalLabeledHierarchy:
        case Asc.c_oAscSmartArtTypes.IncreasingArrowsProcess: {
          return getShapesFromPresetGeom();
        }
        case Asc.c_oAscSmartArtTypes.VerticalAccentList:
          return getShapesFromPresStyleLbl(['revTx', 'solidFgAcc1']);
        case Asc.c_oAscSmartArtTypes.DescendingBlockList:
          return getShapesFromPresName(['childText']);
        case Asc.c_oAscSmartArtTypes.TableList:
          return getShapesFromPresStyleLbl(['node1', 'dkBgShp']);
        case Asc.c_oAscSmartArtTypes.SegmentedProcess:
          return getShapesFromPresStyleLbl(['node1', 'fgAccFollowNode1']);
        case Asc.c_oAscSmartArtTypes.TargetList:
          return getShapesFromPresName(['СhTx', 'rect']); // TODO: think about it
        case Asc.c_oAscSmartArtTypes.HierarchyList:
          return getShapesFromPresStyleLbl(['node1', 'bgAcc1']);
        case Asc.c_oAscSmartArtTypes.AlternatingFlow:
          return getShapesFromPresStyleLbl(['node1', 'bgAcc1']);
        case Asc.c_oAscSmartArtTypes.ChevronList:
          return getShapesFromPresStyleLbl(['node1', 'alignAccFollowNode1']);
        case Asc.c_oAscSmartArtTypes.ProcessList:
          return getShapesFromPresStyleLbl(['alignAccFollowNode1', 'node1']);
        case Asc.c_oAscSmartArtTypes.AscendingPictureAccentProcess:
          return getShapesFromPresStyleLbl(['node1', 'revTx']);
        case Asc.c_oAscSmartArtTypes.VerticalEquation:
          return getShapesFromPresName(['lastNode', 'node']);
        case Asc.c_oAscSmartArtTypes.BasicRadial:
        case Asc.c_oAscSmartArtTypes.DivergingRadial: // TODO: think
          return getShapesFromPresStyleLbl(['node1', 'node0']);
        case Asc.c_oAscSmartArtTypes.RadialVenn:
          return getShapesFromPresName(['centerShape', 'node']);
        case Asc.c_oAscSmartArtTypes.RadialCluster:
          return [callShape];
        case Asc.c_oAscSmartArtTypes.NameAndTitleOrganizationChart:
          return getShapesFromPresStyleLbl(['node0'], true);
        case Asc.c_oAscSmartArtTypes.Balance:
          return getShapesFromPresStyleLbl(['alignAccFollowNode1', 'node1']);
        case Asc.c_oAscSmartArtTypes.NestedTarget:
          return getShapesFromPresStyleLbl(['node1', 'fgAcc1']);
        case Asc.c_oAscSmartArtTypes.AccentedPicture:
        case Asc.c_oAscSmartArtTypes.CircularPictureCallout:
          return getShapesFromPresStyleLbl(['revTx', 'node1']);
        case Asc.c_oAscSmartArtTypes.ThemePictureGrid:
          return getShapesFromPresStyleLbl(['revTx', 'trBgShp']);
        case Asc.c_oAscSmartArtTypes.TitlePictureLineup:
          return getShapesFromPresStyleLbl(['revTx', 'alignNode1']);
        default:
          return [];
      }
    }
    SmartArt.prototype.setParent = function (parent) {
      oHistory.CanAddChanges() && oHistory.Add(new AscDFH.CChangesDrawingsObject(this, AscDFH.historyitem_SmartArtParent, this.parent, parent));
      this.parent = parent;
    };
    SmartArt.prototype.setColorsDef = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_SmartArtColorsDef, this.getColorsDef(), oPr));
      this.colorsDef = oPr;
    };
    SmartArt.prototype.setType = function (oPr) {
      this.type = oPr;
    };
    SmartArt.prototype.setDrawing = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SmartArtDrawing, this.getDrawing(), oPr));
			this.drawing = oPr;
	    if (oPr) {
		    oPr.setParent(this);
	    }
    };
    SmartArt.prototype.setLayoutDef = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_SmartArtLayoutDef, this.getLayoutDef(), oPr));
      this.layoutDef = oPr;
    };
    SmartArt.prototype.setDataModel = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObject(this, AscDFH.historyitem_SmartArtDataModel, this.getDataModel(), oPr));
      this.dataModel = oPr;
	    if (oPr) {
		    oPr.setParent(this);
	    }
    };
    SmartArt.prototype.setStyleDef = function (oPr) {
      oHistory.CanAddChanges() && oHistory.Add(new CChangeObjectNoId(this, AscDFH.historyitem_SmartArtStyleDef, this.getStyleDef(), oPr));
      this.styleDef = oPr;
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

	  SmartArt.prototype.getRelationOfContent2 = function () {
		  const dataModel = this.getDataModel() && this.getDataModel().getDataModel();
		  if (dataModel) {
			  const connections = {};
				connections[Cxn_type_presOf] = {};
				connections[Cxn_type_parOf] = {};
				connections[Cxn_type_presParOf] = {};
			  const ptMap = this.getPtMap();
			  const cxnLst = dataModel.cxnLst.list;

			  cxnLst.forEach(function (cxn) {
					switch (cxn.type) {
						case Cxn_type_presOf: {
							const point = ptMap[cxn.srcId];
							if (!connections[Cxn_type_presOf][point.getModelId()]) {
								connections[Cxn_type_presOf][point.getModelId()] = {};
							}
							const destPresPoint = ptMap[cxn.destId]
							connections[Cxn_type_presOf][point.getModelId()][destPresPoint.getPresName()] = destPresPoint;
							break;
						}
						case Cxn_type_parOf: {
							if (!connections[Cxn_type_parOf][cxn.srcId]) {
								connections[Cxn_type_parOf][cxn.srcId] = [];
							}
							connections[Cxn_type_parOf][cxn.srcId].push({
								point   : ptMap[cxn.destId],
								sibPoint: ptMap[cxn.sibTransId],
								parPoint: ptMap[cxn.parTransId],
								index: cxn.srcOrd
							});
							break;
						}
						case Cxn_type_presParOf: {
							if (!connections[Cxn_type_presParOf][cxn.srcId]) {
								connections[Cxn_type_presParOf][cxn.srcId] = {};
							}
              const point = ptMap[cxn.destId];
              if (point) {
               if (!connections[Cxn_type_presParOf][cxn.srcId][point.getPresName()]) {
                 connections[Cxn_type_presParOf][cxn.srcId][point.getPresName()] = [];
                }
                connections[Cxn_type_presParOf][cxn.srcId][point.getPresName()].push({point: point, index: cxn.srcOrd});
              }
							break;
						}
						default:
							break;
					}
			  });
			  for (let id in connections[Cxn_type_parOf]) {
				  const points = connections[Cxn_type_parOf][id];
				  points.sort(function (a, b) {
					  return a.index - b.index;
				  });
			  }
        for (let id in connections[Cxn_type_presParOf]) {
          const presNameMap = connections[Cxn_type_presParOf][id];
          for (let presName in presNameMap) {
            presNameMap[presName].sort(function(a, b) {
              return a.index - b.index;
            });
          }
        }
			  return connections;
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
          return cxn.type === Cxn_type_presOf;
        });

        presCxnLst.forEach(function (cxn) {
          var shape = shapeMap[cxn.destId];
          if (shape) {
            if (!connections[cxn.destId]) {
              connections[cxn.destId] = [];
            }
            if (ptMap[cxn.srcId]) {
              connections[cxn.destId].push({
                point: ptMap[cxn.srcId],
                srcOrd: cxn.srcOrd,
                destOrd: cxn.destOrd
              });
            }

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
          return cxn.type === Cxn_type_presOf || cxn.type === Cxn_type_presParOf;
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
				const contentNodeRelations = this.smartArtTree.getDataPointRelations();
        for (var modelId in shapeMap) {
          var shape = shapeMap[modelId];
          var smartArtInfo = new ShapeSmartArtInfo();
          shape.setShapeSmartArtInfo(smartArtInfo);
          if (contentConnections[modelId]) {
            contentConnections[modelId].forEach(function (el) {
							const contentNode = contentNodeRelations[el.point.getModelId()];
              smartArtInfo.addToLstContentPoint(smartArtInfo.contentPoint.length, contentNode);
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
					this.drawing.setBDeleted(false);
          pReader.ReadSmartArtGroup(this.drawing);
          this.drawing.setGroup(this);
          this.addToSpTree(0, this.drawing);
          break;
        }
        case 1: {
          this.setDataModel(new DiagramData());
          this.dataModel.fromPPTY(pReader);
          break;
        }
        case 2: {
          const colorsDef = new ColorsDef();
          colorsDef.fromPPTY(pReader);
          this.setColorsDef(colorsDef);
          break;
        }
        case 3: {
          const layoutDef = new LayoutDef();
          layoutDef.fromPPTY(pReader);
          this.setLayoutDef(layoutDef);
          break;
        }
        case 4: {
          const styleDef = new StyleDef();
          styleDef.fromPPTY(pReader);
          this.setStyleDef(styleDef);
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

    SmartArt.prototype.getNullNamePoint = function () {
      var dataModel = this.getDataModel() && this.getDataModel().getDataModel();
      var ptLst = dataModel.ptLst.list;
      for (var i = 0; i < ptLst.length; i += 1) {
        if (ptLst[i].type === Point_type_pres && ptLst[i].prSet && ptLst[i].prSet.presName === 'Name0') { // TODO: ptLst type = 'pres' is 4
          return ptLst[i];
        }
      }
    }

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
      if (oXfrm.isZero && oXfrm.isZero()) {
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

      if (graphics.animationDrawer) {
        graphics.animationDrawer.drawObject(this, graphics);
        return;
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
		SmartArt.prototype.check_bounds = function (oChecker)
		{
			oChecker._s();
			oChecker._m(0, 0);
			oChecker._l(this.extX, 0);
			oChecker._l(this.extX, this.extY);
			oChecker._l(0, this.extY);
			oChecker._z();
			oChecker._e();
		}
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
	  SmartArt.prototype.getCompiledFill = function () {
			return null;
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
	    if (oPr && oPr.contentCopyPr && oPr.contentCopyPr.Comparison) {
		    copy.generateDrawingPart();
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
      if(!oPr || false !== oPr.cacheImage) {
        copy.cachedImage = this.getBase64Img();
        copy.cachedPixH = this.cachedPixH;
        copy.cachedPixW = this.cachedPixW;
      }
      copy.setLocks(this.locks);
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
      oCopy.removePlaceholder();
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
    SmartArt.prototype.checkEmptySpPrAndXfrm = function(_xfrm) {
      CGraphicObjectBase.prototype.checkEmptySpPrAndXfrm.call(this, _xfrm);
      if(this.drawing) {
        var oDrawing = this.drawing;
        oDrawing.setSpPr(new AscFormat.CSpPr());
        oDrawing.spPr.setParent(this);
        oDrawing.spPr.setXfrm(this.spPr.xfrm.createDuplicate());
        oDrawing.spPr.xfrm.setParent(oDrawing.spPr);
        oDrawing.spPr.xfrm.setOffX(0);
        oDrawing.spPr.xfrm.setOffY(0);
      }
    };
    SmartArt.prototype.checkNodePointsAfterRead = function(bReplaceAll) {
			//todo
      // let tree = this.createHierarchy();
      // tree.traverseBF(function (node) {
      //   let nodePoint = node.data && (node.data.nodePoint || node.data.asstPoint);
      //   if (nodePoint) {
      //     if (bReplaceAll) {
      //       nodePoint.setPhldrT('[' + AscCommon.translateManager.getValue('Text') + ']');
      //     } else {
      //       const oPlaceholderText = nodePoint.getPhldrT();
      //       if (typeof oPlaceholderText !== 'string') {
      //         nodePoint.setPhldrT('');
      //       }
      //     }
      //   }
      // });
    };


    window['AscFormat'] = window['AscFormat'] || {};
    window['AscFormat'].kForInsFitFontSize     = kForInsFitFontSize;
    window['AscFormat'].PrSet                  = PrSet;
    window['AscFormat'].CCommonDataList        = CCommonDataList;
    window['AscFormat'].CCommonDataListNoId    = CCommonDataListNoId;
    window['AscFormat'].Point                  = Point;
    window['AscFormat'].PtLst                  = PtLst;
    window['AscFormat'].DataModel              = DataModel;
    window['AscFormat'].CxnLst                 = CxnLst;
    window['AscFormat'].BgFormat               = BgFormat;
    window['AscFormat'].Whole                  = Whole;
    window['AscFormat'].Cxn                    = Cxn;
    window['AscFormat'].LayoutDef              = LayoutDef;
    window['AscFormat'].CatLst                 = CatLst;
    window['AscFormat'].SCat                   = SCat;
    window['AscFormat'].LayoutNode             = LayoutNode;
    window['AscFormat'].Alg                    = Alg;
    window['AscFormat'].Param                  = Param;
    window['AscFormat'].Choose                 = Choose;
    window['AscFormat'].IteratorAttributes     = IteratorAttributes;
    window['AscFormat'].Else                   = Else;
    window['AscFormat'].If                     = If;
    window['AscFormat'].ConstrLst              = ConstrLst;
    window['AscFormat'].Constr                 = Constr;
    window['AscFormat'].PresOf                 = PresOf;
    window['AscFormat'].RuleLst                = RuleLst;
    window['AscFormat'].Rule                   = Rule;
    window['AscFormat'].SShape                 = SShape;
    window['AscFormat'].AdjLst                 = AdjLst;
    window['AscFormat'].Adj                    = Adj;
    window['AscFormat'].DiagramTitle           = DiagramTitle;
    window['AscFormat'].VarLst                 = VarLst;
    window['AscFormat'].ColorsDef              = ColorsDef;
    window['AscFormat'].ColorDefStyleLbl       = ColorDefStyleLbl;
    window['AscFormat'].ClrLst                 = ClrLst;
    window['AscFormat'].StyleDef               = StyleDef;
    window['AscFormat'].Scene3d                = Scene3d;
    window['AscFormat'].StyleDefStyleLbl       = StyleDefStyleLbl;
    window['AscFormat'].Backdrop               = Backdrop;
    window['AscFormat'].Point3D                = Point3D;
    window['AscFormat'].Vector3D               = Vector3D;
    window['AscFormat'].Camera                 = Camera;
    window['AscFormat'].Rot                    = Rot;
    window['AscFormat'].LightRig               = LightRig;
    window['AscFormat'].Sp3d                   = Sp3d;
    window['AscFormat'].Bevel                  = Bevel;
    window['AscFormat'].SampData               = SampData;
    window['AscFormat'].ForEach                = ForEach;
    window['AscFormat'].SmartArt               = SmartArt;
    window['AscFormat'].Drawing                = Drawing;
    window['AscFormat'].DiagramData            = DiagramData;
    window['AscFormat'].ShapeSmartArtInfo      = ShapeSmartArtInfo;
    window['AscFormat'].LayoutBaseClass        = LayoutBaseClass;
    window['AscFormat'].IteratorLayoutBase     = IteratorLayoutBase;

    window['AscFormat'].Point_type_asst = Point_type_asst;
    window['AscFormat'].Point_type_doc = Point_type_doc;
    window['AscFormat'].Point_type_node = Point_type_node;
    window['AscFormat'].Point_type_parTrans = Point_type_parTrans;
    window['AscFormat'].Point_type_pres = Point_type_pres;
    window['AscFormat'].Point_type_sibTrans = Point_type_sibTrans

    window['AscFormat'].Cxn_type_parOf = Cxn_type_parOf;
    window['AscFormat'].Cxn_type_presOf = Cxn_type_presOf;
    window['AscFormat'].Cxn_type_presParOf = Cxn_type_presParOf;
    window['AscFormat'].Cxn_type_unknownRelationShip = Cxn_type_unknownRelationShip;

    window['AscFormat'].LayoutNode_type_b = LayoutNode_type_b;
    window['AscFormat'].LayoutNode_type_t = LayoutNode_type_t;

    window['AscFormat'].Alg_type_composite = Alg_type_composite;
    window['AscFormat'].Alg_type_conn = Alg_type_conn;
    window['AscFormat'].Alg_type_cycle = Alg_type_cycle;
    window['AscFormat'].Alg_type_hierChild = Alg_type_hierChild;
    window['AscFormat'].Alg_type_hierRoot = Alg_type_hierRoot;
    window['AscFormat'].Alg_type_lin = Alg_type_lin;
    window['AscFormat'].Alg_type_pyra = Alg_type_pyra;
    window['AscFormat'].Alg_type_snake = Alg_type_snake;
    window['AscFormat'].Alg_type_sp = Alg_type_sp;
    window['AscFormat'].Alg_type_tx = Alg_type_tx;


    window['AscFormat'].Param_type_horzAlign = Param_type_horzAlign;
    window['AscFormat'].Param_type_vertAlign = Param_type_vertAlign;
    window['AscFormat'].Param_type_chDir = Param_type_chDir;
    window['AscFormat'].Param_type_chAlign = Param_type_chAlign;
    window['AscFormat'].Param_type_secChAlign = Param_type_secChAlign;
    window['AscFormat'].Param_type_linDir = Param_type_linDir;
    window['AscFormat'].Param_type_secLinDir = Param_type_secLinDir;
    window['AscFormat'].Param_type_stElem = Param_type_stElem;
    window['AscFormat'].Param_type_bendPt = Param_type_bendPt;
    window['AscFormat'].Param_type_connRout = Param_type_connRout;
    window['AscFormat'].Param_type_begSty = Param_type_begSty;
    window['AscFormat'].Param_type_endSty = Param_type_endSty;
    window['AscFormat'].Param_type_dim = Param_type_dim;
    window['AscFormat'].Param_type_rotPath = Param_type_rotPath;
    window['AscFormat'].Param_type_ctrShpMap = Param_type_ctrShpMap;
    window['AscFormat'].Param_type_nodeHorzAlign = Param_type_nodeHorzAlign;
    window['AscFormat'].Param_type_nodeVertAlign = Param_type_nodeVertAlign;
    window['AscFormat'].Param_type_fallback = Param_type_fallback;
    window['AscFormat'].Param_type_txDir = Param_type_txDir;
    window['AscFormat'].Param_type_pyraAcctPos = Param_type_pyraAcctPos;
    window['AscFormat'].Param_type_pyraAcctTxMar = Param_type_pyraAcctTxMar;
    window['AscFormat'].Param_type_txBlDir = Param_type_txBlDir;
    window['AscFormat'].Param_type_txAnchorHorz = Param_type_txAnchorHorz;
    window['AscFormat'].Param_type_txAnchorVert = Param_type_txAnchorVert;
    window['AscFormat'].Param_type_txAnchorHorzCh = Param_type_txAnchorHorzCh;
    window['AscFormat'].Param_type_txAnchorVertCh = Param_type_txAnchorVertCh;
    window['AscFormat'].Param_type_parTxLTRAlign = Param_type_parTxLTRAlign;
    window['AscFormat'].Param_type_parTxRTLAlign = Param_type_parTxRTLAlign;
    window['AscFormat'].Param_type_shpTxLTRAlignCh = Param_type_shpTxLTRAlignCh;
    window['AscFormat'].Param_type_shpTxRTLAlignCh = Param_type_shpTxRTLAlignCh;
    window['AscFormat'].Param_type_autoTxRot = Param_type_autoTxRot;
    window['AscFormat'].Param_type_grDir = Param_type_grDir;
    window['AscFormat'].Param_type_flowDir = Param_type_flowDir;
    window['AscFormat'].Param_type_contDir = Param_type_contDir;
    window['AscFormat'].Param_type_bkpt = Param_type_bkpt;
    window['AscFormat'].Param_type_off = Param_type_off;
    window['AscFormat'].Param_type_hierAlign = Param_type_hierAlign;
    window['AscFormat'].Param_type_bkPtFixedVal = Param_type_bkPtFixedVal;
    window['AscFormat'].Param_type_stBulletLvl = Param_type_stBulletLvl;
    window['AscFormat'].Param_type_stAng = Param_type_stAng;
    window['AscFormat'].Param_type_spanAng = Param_type_spanAng;
    window['AscFormat'].Param_type_ar = Param_type_ar;
    window['AscFormat'].Param_type_lnSpPar = Param_type_lnSpPar;
    window['AscFormat'].Param_type_lnSpAfParP = Param_type_lnSpAfParP;
    window['AscFormat'].Param_type_lnSpCh = Param_type_lnSpCh;
    window['AscFormat'].Param_type_lnSpAfChP = Param_type_lnSpAfChP;
    window['AscFormat'].Param_type_rtShortDist = Param_type_rtShortDist;
    window['AscFormat'].Param_type_alignTx = Param_type_alignTx;
    window['AscFormat'].Param_type_pyraLvlNode = Param_type_pyraLvlNode;
    window['AscFormat'].Param_type_pyraAcctBkgdNode = Param_type_pyraAcctBkgdNode;
    window['AscFormat'].Param_type_pyraAcctTxNode = Param_type_pyraAcctTxNode;
    window['AscFormat'].Param_type_srcNode = Param_type_srcNode;
    window['AscFormat'].Param_type_dstNode = Param_type_dstNode;
    window['AscFormat'].Param_type_begPts = Param_type_begPts;
    window['AscFormat'].Param_type_endPts = Param_type_endPts;

    window['AscFormat'].AxisType_value_ancst = AxisType_value_ancst;
    window['AscFormat'].AxisType_value_ancstOrSelf = AxisType_value_ancstOrSelf;
    window['AscFormat'].AxisType_value_ch = AxisType_value_ch;
    window['AscFormat'].AxisType_value_des = AxisType_value_des;
    window['AscFormat'].AxisType_value_desOrSelf = AxisType_value_desOrSelf;
    window['AscFormat'].AxisType_value_follow = AxisType_value_follow;
    window['AscFormat'].AxisType_value_followSib = AxisType_value_followSib;
    window['AscFormat'].AxisType_value_none = AxisType_value_none;
    window['AscFormat'].AxisType_value_par = AxisType_value_par;
    window['AscFormat'].AxisType_value_preced = AxisType_value_preced;
    window['AscFormat'].AxisType_value_precedSib = AxisType_value_precedSib;
    window['AscFormat'].AxisType_value_root = AxisType_value_root;
    window['AscFormat'].AxisType_value_self = AxisType_value_self;

    window['AscFormat'].ElementType_value_all = ElementType_value_all;
    window['AscFormat'].ElementType_value_asst = ElementType_value_asst;
    window['AscFormat'].ElementType_value_doc = ElementType_value_doc;
    window['AscFormat'].ElementType_value_node = ElementType_value_node;
    window['AscFormat'].ElementType_value_nonAsst = ElementType_value_nonAsst;
    window['AscFormat'].ElementType_value_nonNorm = ElementType_value_nonNorm;
    window['AscFormat'].ElementType_value_norm = ElementType_value_norm;
    window['AscFormat'].ElementType_value_parTrans = ElementType_value_parTrans;
    window['AscFormat'].ElementType_value_pres = ElementType_value_pres;
    window['AscFormat'].ElementType_value_sibTrans = ElementType_value_sibTrans;

    window['AscFormat'].If_op_equ = If_op_equ;
    window['AscFormat'].If_op_neq = If_op_neq;
    window['AscFormat'].If_op_gt = If_op_gt;
    window['AscFormat'].If_op_lt = If_op_lt;
    window['AscFormat'].If_op_gte = If_op_gte;
    window['AscFormat'].If_op_lte = If_op_lte;

    window['AscFormat'].If_func_cnt = If_func_cnt;
    window['AscFormat'].If_func_depth = If_func_depth;
    window['AscFormat'].If_func_maxDepth = If_func_maxDepth;
    window['AscFormat'].If_func_pos = If_func_pos;
    window['AscFormat'].If_func_posEven = If_func_posEven;
    window['AscFormat'].If_func_posOdd = If_func_posOdd;
    window['AscFormat'].If_func_revPos = If_func_revPos;
    window['AscFormat'].If_func_var = If_func_var;

    window['AscFormat'].If_arg_animLvl = If_arg_animLvl;
    window['AscFormat'].If_arg_animOne = If_arg_animOne;
    window['AscFormat'].If_arg_bulEnabled = If_arg_bulEnabled;
    window['AscFormat'].If_arg_chMax = If_arg_chMax;
    window['AscFormat'].If_arg_chPref = If_arg_chPref;
    window['AscFormat'].If_arg_dir = If_arg_dir;
    window['AscFormat'].If_arg_hierBranch = If_arg_hierBranch;
    window['AscFormat'].If_arg_none = If_arg_none;
    window['AscFormat'].If_arg_orgChart = If_arg_orgChart;
    window['AscFormat'].If_arg_resizeHandles = If_arg_resizeHandles;

    window['AscFormat'].Constr_for_ch = Constr_for_ch;
    window['AscFormat'].Constr_for_des = Constr_for_des;
    window['AscFormat'].Constr_for_self = Constr_for_self;

    window['AscFormat'].Constr_op_equ = Constr_op_equ;
    window['AscFormat'].Constr_op_gte = Constr_op_gte;
    window['AscFormat'].Constr_op_lte = Constr_op_lte;
    window['AscFormat'].Constr_op_none = Constr_op_none;

    window['AscFormat'].Constr_type_alignOff =Constr_type_alignOff;
    window['AscFormat'].Constr_type_b =Constr_type_b;
    window['AscFormat'].Constr_type_begMarg =Constr_type_begMarg;
    window['AscFormat'].Constr_type_begPad =Constr_type_begPad;
    window['AscFormat'].Constr_type_bendDist =Constr_type_bendDist;
    window['AscFormat'].Constr_type_bMarg =Constr_type_bMarg;
    window['AscFormat'].Constr_type_bOff =Constr_type_bOff;
    window['AscFormat'].Constr_type_connDist = Constr_type_connDist;
    window['AscFormat'].Constr_type_ctrX =Constr_type_ctrX;
    window['AscFormat'].Constr_type_ctrXOff =Constr_type_ctrXOff;
    window['AscFormat'].Constr_type_ctrY = Constr_type_ctrY;
    window['AscFormat'].Constr_type_ctrYOff = Constr_type_ctrYOff;
    window['AscFormat'].Constr_type_diam = Constr_type_diam;
    window['AscFormat'].Constr_type_endMarg = Constr_type_endMarg;
    window['AscFormat'].Constr_type_endPad = Constr_type_endPad;
    window['AscFormat'].Constr_type_h = Constr_type_h;
    window['AscFormat'].Constr_type_hArH = Constr_type_hArH;
    window['AscFormat'].Constr_type_hOff = Constr_type_hOff;
    window['AscFormat'].Constr_type_l = Constr_type_l;
    window['AscFormat'].Constr_type_lMarg = Constr_type_lMarg;
    window['AscFormat'].Constr_type_lOff = Constr_type_lOff;
    window['AscFormat'].Constr_type_none =Constr_type_none;
    window['AscFormat'].Constr_type_primFontSz = Constr_type_primFontSz;
    window['AscFormat'].Constr_type_pyraAcctRatio = Constr_type_pyraAcctRatio;
    window['AscFormat'].Constr_type_r = Constr_type_r;
    window['AscFormat'].Constr_type_rMarg = Constr_type_rMarg;
    window['AscFormat'].Constr_type_rOff = Constr_type_rOff;
    window['AscFormat'].Constr_type_secFontSz = Constr_type_secFontSz;
    window['AscFormat'].Constr_type_secSibSp = Constr_type_secSibSp;
    window['AscFormat'].Constr_type_sibSp = Constr_type_sibSp;
    window['AscFormat'].Constr_type_sp = Constr_type_sp;
    window['AscFormat'].Constr_type_stemThick = Constr_type_stemThick;
    window['AscFormat'].Constr_type_t = Constr_type_t;
    window['AscFormat'].Constr_type_tMarg = Constr_type_tMarg;
    window['AscFormat'].Constr_type_tOff = Constr_type_tOff;
    window['AscFormat'].Constr_type_userA = Constr_type_userA;
    window['AscFormat'].Constr_type_userB = Constr_type_userB;
    window['AscFormat'].Constr_type_userC = Constr_type_userC;
    window['AscFormat'].Constr_type_userD = Constr_type_userD;
    window['AscFormat'].Constr_type_userE = Constr_type_userE;
    window['AscFormat'].Constr_type_userF = Constr_type_userF;
    window['AscFormat'].Constr_type_userG = Constr_type_userG;
    window['AscFormat'].Constr_type_userH = Constr_type_userH;
    window['AscFormat'].Constr_type_userI = Constr_type_userI;
    window['AscFormat'].Constr_type_userJ = Constr_type_userJ;
    window['AscFormat'].Constr_type_userK = Constr_type_userK;
    window['AscFormat'].Constr_type_userL = Constr_type_userL;
    window['AscFormat'].Constr_type_userM = Constr_type_userM;
    window['AscFormat'].Constr_type_userN = Constr_type_userN;
    window['AscFormat'].Constr_type_userO = Constr_type_userO;
    window['AscFormat'].Constr_type_userP = Constr_type_userP;
    window['AscFormat'].Constr_type_userQ = Constr_type_userQ;
    window['AscFormat'].Constr_type_userR = Constr_type_userR;
    window['AscFormat'].Constr_type_userS = Constr_type_userS;
    window['AscFormat'].Constr_type_userT = Constr_type_userT;
    window['AscFormat'].Constr_type_userU = Constr_type_userU;
    window['AscFormat'].Constr_type_userV = Constr_type_userV;
    window['AscFormat'].Constr_type_userW = Constr_type_userW;
    window['AscFormat'].Constr_type_userX = Constr_type_userX;
    window['AscFormat'].Constr_type_userY = Constr_type_userY;
    window['AscFormat'].Constr_type_userZ = Constr_type_userZ;
    window['AscFormat'].Constr_type_w = Constr_type_w;
    window['AscFormat'].Constr_type_wArH = Constr_type_wArH;
    window['AscFormat'].Constr_type_wOff = Constr_type_wOff;

    window['AscFormat'].kForInsFitFontSize = kForInsFitFontSize;

    window['AscFormat'].LayoutShapeType_outputShapeType_conn = LayoutShapeType_outputShapeType_conn;
    window['AscFormat'].LayoutShapeType_outputShapeType_none = LayoutShapeType_outputShapeType_none;
    window['AscFormat'].LayoutShapeType_shapeType_accentBorderCallout1 = LayoutShapeType_shapeType_accentBorderCallout1;
    window['AscFormat'].LayoutShapeType_shapeType_accentBorderCallout2 = LayoutShapeType_shapeType_accentBorderCallout2;
    window['AscFormat'].LayoutShapeType_shapeType_accentBorderCallout3 = LayoutShapeType_shapeType_accentBorderCallout3;
    window['AscFormat'].LayoutShapeType_shapeType_accentCallout1 = LayoutShapeType_shapeType_accentCallout1;
    window['AscFormat'].LayoutShapeType_shapeType_accentCallout2 = LayoutShapeType_shapeType_accentCallout2;
    window['AscFormat'].LayoutShapeType_shapeType_accentCallout3 = LayoutShapeType_shapeType_accentCallout3;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonBackPrevious = LayoutShapeType_shapeType_actionButtonBackPrevious;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonBeginning = LayoutShapeType_shapeType_actionButtonBeginning;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonBlank = LayoutShapeType_shapeType_actionButtonBlank;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonDocument = LayoutShapeType_shapeType_actionButtonDocument;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonEnd = LayoutShapeType_shapeType_actionButtonEnd;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonForwardNext = LayoutShapeType_shapeType_actionButtonForwardNext;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonHelp = LayoutShapeType_shapeType_actionButtonHelp;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonHome = LayoutShapeType_shapeType_actionButtonHome;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonInformation = LayoutShapeType_shapeType_actionButtonInformation;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonMovie = LayoutShapeType_shapeType_actionButtonMovie;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonReturn = LayoutShapeType_shapeType_actionButtonReturn;
    window['AscFormat'].LayoutShapeType_shapeType_actionButtonSound = LayoutShapeType_shapeType_actionButtonSound;
    window['AscFormat'].LayoutShapeType_shapeType_arc = LayoutShapeType_shapeType_arc;
    window['AscFormat'].LayoutShapeType_shapeType_bentArrow = LayoutShapeType_shapeType_bentArrow;
    window['AscFormat'].LayoutShapeType_shapeType_bentConnector2 = LayoutShapeType_shapeType_bentConnector2;
    window['AscFormat'].LayoutShapeType_shapeType_bentConnector3 = LayoutShapeType_shapeType_bentConnector3;
    window['AscFormat'].LayoutShapeType_shapeType_bentConnector4 = LayoutShapeType_shapeType_bentConnector4;
    window['AscFormat'].LayoutShapeType_shapeType_bentConnector5 = LayoutShapeType_shapeType_bentConnector5;
    window['AscFormat'].LayoutShapeType_shapeType_bentUpArrow = LayoutShapeType_shapeType_bentUpArrow;
    window['AscFormat'].LayoutShapeType_shapeType_bevel = LayoutShapeType_shapeType_bevel;
    window['AscFormat'].LayoutShapeType_shapeType_blockArc = LayoutShapeType_shapeType_blockArc;
    window['AscFormat'].LayoutShapeType_shapeType_borderCallout1 = LayoutShapeType_shapeType_borderCallout1;
    window['AscFormat'].LayoutShapeType_shapeType_borderCallout2 = LayoutShapeType_shapeType_borderCallout2;
    window['AscFormat'].LayoutShapeType_shapeType_borderCallout3 = LayoutShapeType_shapeType_borderCallout3;
    window['AscFormat'].LayoutShapeType_shapeType_bracePair = LayoutShapeType_shapeType_bracePair;
    window['AscFormat'].LayoutShapeType_shapeType_bracketPair = LayoutShapeType_shapeType_bracketPair;
    window['AscFormat'].LayoutShapeType_shapeType_callout1 = LayoutShapeType_shapeType_callout1;
    window['AscFormat'].LayoutShapeType_shapeType_callout2 = LayoutShapeType_shapeType_callout2;
    window['AscFormat'].LayoutShapeType_shapeType_callout3 = LayoutShapeType_shapeType_callout3;
    window['AscFormat'].LayoutShapeType_shapeType_can = LayoutShapeType_shapeType_can;
    window['AscFormat'].LayoutShapeType_shapeType_chartPlus = LayoutShapeType_shapeType_chartPlus;
    window['AscFormat'].LayoutShapeType_shapeType_chartStar = LayoutShapeType_shapeType_chartStar;
    window['AscFormat'].LayoutShapeType_shapeType_chartX = LayoutShapeType_shapeType_chartX;
    window['AscFormat'].LayoutShapeType_shapeType_chevron = LayoutShapeType_shapeType_chevron;
    window['AscFormat'].LayoutShapeType_shapeType_chord = LayoutShapeType_shapeType_chord;
    window['AscFormat'].LayoutShapeType_shapeType_circularArrow = LayoutShapeType_shapeType_circularArrow;
    window['AscFormat'].LayoutShapeType_shapeType_cloud = LayoutShapeType_shapeType_cloud;
    window['AscFormat'].LayoutShapeType_shapeType_cloudCallout = LayoutShapeType_shapeType_cloudCallout;
    window['AscFormat'].LayoutShapeType_shapeType_corner = LayoutShapeType_shapeType_corner;
    window['AscFormat'].LayoutShapeType_shapeType_cornerTabs = LayoutShapeType_shapeType_cornerTabs;
    window['AscFormat'].LayoutShapeType_shapeType_cube = LayoutShapeType_shapeType_cube;
    window['AscFormat'].LayoutShapeType_shapeType_curvedConnector2 = LayoutShapeType_shapeType_curvedConnector2;
    window['AscFormat'].LayoutShapeType_shapeType_curvedConnector3 = LayoutShapeType_shapeType_curvedConnector3;
    window['AscFormat'].LayoutShapeType_shapeType_curvedConnector4 = LayoutShapeType_shapeType_curvedConnector4;
    window['AscFormat'].LayoutShapeType_shapeType_curvedConnector5 = LayoutShapeType_shapeType_curvedConnector5;
    window['AscFormat'].LayoutShapeType_shapeType_curvedDownArrow = LayoutShapeType_shapeType_curvedDownArrow;
    window['AscFormat'].LayoutShapeType_shapeType_curvedLeftArrow = LayoutShapeType_shapeType_curvedLeftArrow;
    window['AscFormat'].LayoutShapeType_shapeType_curvedRightArrow = LayoutShapeType_shapeType_curvedRightArrow;
    window['AscFormat'].LayoutShapeType_shapeType_curvedUpArrow = LayoutShapeType_shapeType_curvedUpArrow;
    window['AscFormat'].LayoutShapeType_shapeType_decagon = LayoutShapeType_shapeType_decagon;
    window['AscFormat'].LayoutShapeType_shapeType_diagStripe = LayoutShapeType_shapeType_diagStripe;
    window['AscFormat'].LayoutShapeType_shapeType_diamond = LayoutShapeType_shapeType_diamond;
    window['AscFormat'].LayoutShapeType_shapeType_dodecagon = LayoutShapeType_shapeType_dodecagon;
    window['AscFormat'].LayoutShapeType_shapeType_donut = LayoutShapeType_shapeType_donut;
    window['AscFormat'].LayoutShapeType_shapeType_doubleWave = LayoutShapeType_shapeType_doubleWave;
    window['AscFormat'].LayoutShapeType_shapeType_downArrow = LayoutShapeType_shapeType_downArrow;
    window['AscFormat'].LayoutShapeType_shapeType_downArrowCallout = LayoutShapeType_shapeType_downArrowCallout;
    window['AscFormat'].LayoutShapeType_shapeType_ellipse = LayoutShapeType_shapeType_ellipse;
    window['AscFormat'].LayoutShapeType_shapeType_ellipseRibbon = LayoutShapeType_shapeType_ellipseRibbon;
    window['AscFormat'].LayoutShapeType_shapeType_ellipseRibbon2 = LayoutShapeType_shapeType_ellipseRibbon2;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartAlternateProcess = LayoutShapeType_shapeType_flowChartAlternateProcess;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartCollate = LayoutShapeType_shapeType_flowChartCollate;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartConnector = LayoutShapeType_shapeType_flowChartConnector;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartDecision = LayoutShapeType_shapeType_flowChartDecision;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartDelay = LayoutShapeType_shapeType_flowChartDelay;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartDisplay = LayoutShapeType_shapeType_flowChartDisplay;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartDocument = LayoutShapeType_shapeType_flowChartDocument;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartExtract = LayoutShapeType_shapeType_flowChartExtract;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartInputOutput = LayoutShapeType_shapeType_flowChartInputOutput;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartInternalStorage = LayoutShapeType_shapeType_flowChartInternalStorage;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartMagneticDisk = LayoutShapeType_shapeType_flowChartMagneticDisk;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartMagneticDrum = LayoutShapeType_shapeType_flowChartMagneticDrum;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartMagneticTape = LayoutShapeType_shapeType_flowChartMagneticTape;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartManualInput = LayoutShapeType_shapeType_flowChartManualInput;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartManualOperation = LayoutShapeType_shapeType_flowChartManualOperation;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartMerge = LayoutShapeType_shapeType_flowChartMerge;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartMultidocument = LayoutShapeType_shapeType_flowChartMultidocument;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartOfflineStorage = LayoutShapeType_shapeType_flowChartOfflineStorage;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartOffpageConnector = LayoutShapeType_shapeType_flowChartOffpageConnector;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartOnlineStorage = LayoutShapeType_shapeType_flowChartOnlineStorage;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartOr = LayoutShapeType_shapeType_flowChartOr;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartPredefinedProcess = LayoutShapeType_shapeType_flowChartPredefinedProcess;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartPreparation = LayoutShapeType_shapeType_flowChartPreparation;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartProcess = LayoutShapeType_shapeType_flowChartProcess;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartPunchedCard = LayoutShapeType_shapeType_flowChartPunchedCard;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartPunchedTape = LayoutShapeType_shapeType_flowChartPunchedTape;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartSort = LayoutShapeType_shapeType_flowChartSort;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartSummingJunction = LayoutShapeType_shapeType_flowChartSummingJunction;
    window['AscFormat'].LayoutShapeType_shapeType_flowChartTerminator = LayoutShapeType_shapeType_flowChartTerminator;
    window['AscFormat'].LayoutShapeType_shapeType_foldedCorner = LayoutShapeType_shapeType_foldedCorner;
    window['AscFormat'].LayoutShapeType_shapeType_frame = LayoutShapeType_shapeType_frame;
    window['AscFormat'].LayoutShapeType_shapeType_funnel = LayoutShapeType_shapeType_funnel;
    window['AscFormat'].LayoutShapeType_shapeType_gear6 = LayoutShapeType_shapeType_gear6;
    window['AscFormat'].LayoutShapeType_shapeType_gear9 = LayoutShapeType_shapeType_gear9;
    window['AscFormat'].LayoutShapeType_shapeType_halfFrame = LayoutShapeType_shapeType_halfFrame;
    window['AscFormat'].LayoutShapeType_shapeType_heart = LayoutShapeType_shapeType_heart;
    window['AscFormat'].LayoutShapeType_shapeType_heptagon = LayoutShapeType_shapeType_heptagon;
    window['AscFormat'].LayoutShapeType_shapeType_hexagon = LayoutShapeType_shapeType_hexagon;
    window['AscFormat'].LayoutShapeType_shapeType_homePlate = LayoutShapeType_shapeType_homePlate;
    window['AscFormat'].LayoutShapeType_shapeType_horizontalScroll = LayoutShapeType_shapeType_horizontalScroll;
    window['AscFormat'].LayoutShapeType_shapeType_irregularSeal1 = LayoutShapeType_shapeType_irregularSeal1;
    window['AscFormat'].LayoutShapeType_shapeType_irregularSeal2 = LayoutShapeType_shapeType_irregularSeal2;
    window['AscFormat'].LayoutShapeType_shapeType_leftArrow = LayoutShapeType_shapeType_leftArrow;
    window['AscFormat'].LayoutShapeType_shapeType_leftArrowCallout = LayoutShapeType_shapeType_leftArrowCallout;
    window['AscFormat'].LayoutShapeType_shapeType_leftBrace = LayoutShapeType_shapeType_leftBrace;
    window['AscFormat'].LayoutShapeType_shapeType_leftBracket = LayoutShapeType_shapeType_leftBracket;
    window['AscFormat'].LayoutShapeType_shapeType_leftCircularArrow = LayoutShapeType_shapeType_leftCircularArrow;
    window['AscFormat'].LayoutShapeType_shapeType_leftRightArrow = LayoutShapeType_shapeType_leftRightArrow;
    window['AscFormat'].LayoutShapeType_shapeType_leftRightArrowCallout = LayoutShapeType_shapeType_leftRightArrowCallout;
    window['AscFormat'].LayoutShapeType_shapeType_leftRightCircularArrow = LayoutShapeType_shapeType_leftRightCircularArrow;
    window['AscFormat'].LayoutShapeType_shapeType_leftRightRibbon = LayoutShapeType_shapeType_leftRightRibbon;
    window['AscFormat'].LayoutShapeType_shapeType_leftRightUpArrow = LayoutShapeType_shapeType_leftRightUpArrow;
    window['AscFormat'].LayoutShapeType_shapeType_leftUpArrow = LayoutShapeType_shapeType_leftUpArrow;
    window['AscFormat'].LayoutShapeType_shapeType_lightningBolt = LayoutShapeType_shapeType_lightningBolt;
    window['AscFormat'].LayoutShapeType_shapeType_line = LayoutShapeType_shapeType_line;
    window['AscFormat'].LayoutShapeType_shapeType_lineInv = LayoutShapeType_shapeType_lineInv;
    window['AscFormat'].LayoutShapeType_shapeType_mathDivide = LayoutShapeType_shapeType_mathDivide;
    window['AscFormat'].LayoutShapeType_shapeType_mathEqual = LayoutShapeType_shapeType_mathEqual;
    window['AscFormat'].LayoutShapeType_shapeType_mathMinus = LayoutShapeType_shapeType_mathMinus;
    window['AscFormat'].LayoutShapeType_shapeType_mathMultiply = LayoutShapeType_shapeType_mathMultiply;
    window['AscFormat'].LayoutShapeType_shapeType_mathNotEqual = LayoutShapeType_shapeType_mathNotEqual;
    window['AscFormat'].LayoutShapeType_shapeType_mathPlus = LayoutShapeType_shapeType_mathPlus;
    window['AscFormat'].LayoutShapeType_shapeType_moon = LayoutShapeType_shapeType_moon;
    window['AscFormat'].LayoutShapeType_shapeType_nonIsoscelesTrapezoid = LayoutShapeType_shapeType_nonIsoscelesTrapezoid;
    window['AscFormat'].LayoutShapeType_shapeType_noSmoking = LayoutShapeType_shapeType_noSmoking;
    window['AscFormat'].LayoutShapeType_shapeType_notchedRightArrow = LayoutShapeType_shapeType_notchedRightArrow;
    window['AscFormat'].LayoutShapeType_shapeType_octagon = LayoutShapeType_shapeType_octagon;
    window['AscFormat'].LayoutShapeType_shapeType_parallelogram = LayoutShapeType_shapeType_parallelogram;
    window['AscFormat'].LayoutShapeType_shapeType_pentagon = LayoutShapeType_shapeType_pentagon;
    window['AscFormat'].LayoutShapeType_shapeType_pie = LayoutShapeType_shapeType_pie;
    window['AscFormat'].LayoutShapeType_shapeType_pieWedge = LayoutShapeType_shapeType_pieWedge;
    window['AscFormat'].LayoutShapeType_shapeType_plaque = LayoutShapeType_shapeType_plaque;
    window['AscFormat'].LayoutShapeType_shapeType_plaqueTabs = LayoutShapeType_shapeType_plaqueTabs;
    window['AscFormat'].LayoutShapeType_shapeType_plus = LayoutShapeType_shapeType_plus;
    window['AscFormat'].LayoutShapeType_shapeType_quadArrow = LayoutShapeType_shapeType_quadArrow;
    window['AscFormat'].LayoutShapeType_shapeType_quadArrowCallout = LayoutShapeType_shapeType_quadArrowCallout;
    window['AscFormat'].LayoutShapeType_shapeType_rect = LayoutShapeType_shapeType_rect;
    window['AscFormat'].LayoutShapeType_shapeType_ribbon = LayoutShapeType_shapeType_ribbon;
    window['AscFormat'].LayoutShapeType_shapeType_ribbon2 = LayoutShapeType_shapeType_ribbon2;
    window['AscFormat'].LayoutShapeType_shapeType_rightArrow = LayoutShapeType_shapeType_rightArrow;
    window['AscFormat'].LayoutShapeType_shapeType_rightArrowCallout = LayoutShapeType_shapeType_rightArrowCallout;
    window['AscFormat'].LayoutShapeType_shapeType_rightBrace = LayoutShapeType_shapeType_rightBrace;
    window['AscFormat'].LayoutShapeType_shapeType_rightBracket = LayoutShapeType_shapeType_rightBracket;
    window['AscFormat'].LayoutShapeType_shapeType_round1Rect = LayoutShapeType_shapeType_round1Rect;
    window['AscFormat'].LayoutShapeType_shapeType_round2DiagRect = LayoutShapeType_shapeType_round2DiagRect;
    window['AscFormat'].LayoutShapeType_shapeType_round2SameRect = LayoutShapeType_shapeType_round2SameRect;
    window['AscFormat'].LayoutShapeType_shapeType_roundRect = LayoutShapeType_shapeType_roundRect;
    window['AscFormat'].LayoutShapeType_shapeType_rtTriangle = LayoutShapeType_shapeType_rtTriangle;
    window['AscFormat'].LayoutShapeType_shapeType_smileyFace = LayoutShapeType_shapeType_smileyFace;
    window['AscFormat'].LayoutShapeType_shapeType_snip1Rect = LayoutShapeType_shapeType_snip1Rect;
    window['AscFormat'].LayoutShapeType_shapeType_snip2DiagRect = LayoutShapeType_shapeType_snip2DiagRect;
    window['AscFormat'].LayoutShapeType_shapeType_snip2SameRect = LayoutShapeType_shapeType_snip2SameRect;
    window['AscFormat'].LayoutShapeType_shapeType_snipRoundRect = LayoutShapeType_shapeType_snipRoundRect;
    window['AscFormat'].LayoutShapeType_shapeType_squareTabs = LayoutShapeType_shapeType_squareTabs;
    window['AscFormat'].LayoutShapeType_shapeType_star10 = LayoutShapeType_shapeType_star10;
    window['AscFormat'].LayoutShapeType_shapeType_star12 = LayoutShapeType_shapeType_star12;
    window['AscFormat'].LayoutShapeType_shapeType_star16 = LayoutShapeType_shapeType_star16;
    window['AscFormat'].LayoutShapeType_shapeType_star24 = LayoutShapeType_shapeType_star24;
    window['AscFormat'].LayoutShapeType_shapeType_star32 = LayoutShapeType_shapeType_star32;
    window['AscFormat'].LayoutShapeType_shapeType_star4 = LayoutShapeType_shapeType_star4;
    window['AscFormat'].LayoutShapeType_shapeType_star5 = LayoutShapeType_shapeType_star5;
    window['AscFormat'].LayoutShapeType_shapeType_star6 = LayoutShapeType_shapeType_star6;
    window['AscFormat'].LayoutShapeType_shapeType_star7 = LayoutShapeType_shapeType_star7;
    window['AscFormat'].LayoutShapeType_shapeType_star8 = LayoutShapeType_shapeType_star8;
    window['AscFormat'].LayoutShapeType_shapeType_straightConnector1 = LayoutShapeType_shapeType_straightConnector1;
    window['AscFormat'].LayoutShapeType_shapeType_stripedRightArrow = LayoutShapeType_shapeType_stripedRightArrow;
    window['AscFormat'].LayoutShapeType_shapeType_sun = LayoutShapeType_shapeType_sun;
    window['AscFormat'].LayoutShapeType_shapeType_swooshArrow = LayoutShapeType_shapeType_swooshArrow;
    window['AscFormat'].LayoutShapeType_shapeType_teardrop = LayoutShapeType_shapeType_teardrop;
    window['AscFormat'].LayoutShapeType_shapeType_trapezoid = LayoutShapeType_shapeType_trapezoid;
    window['AscFormat'].LayoutShapeType_shapeType_triangle = LayoutShapeType_shapeType_triangle;
    window['AscFormat'].LayoutShapeType_shapeType_upArrow = LayoutShapeType_shapeType_upArrow;
    window['AscFormat'].LayoutShapeType_shapeType_upArrowCallout = LayoutShapeType_shapeType_upArrowCallout;
    window['AscFormat'].LayoutShapeType_shapeType_upDownArrow = LayoutShapeType_shapeType_upDownArrow;
    window['AscFormat'].LayoutShapeType_shapeType_upDownArrowCallout = LayoutShapeType_shapeType_upDownArrowCallout;
    window['AscFormat'].LayoutShapeType_shapeType_uturnArrow = LayoutShapeType_shapeType_uturnArrow;
    window['AscFormat'].LayoutShapeType_shapeType_verticalScroll = LayoutShapeType_shapeType_verticalScroll;
    window['AscFormat'].LayoutShapeType_shapeType_wave = LayoutShapeType_shapeType_wave;
    window['AscFormat'].LayoutShapeType_shapeType_wedgeEllipseCallout = LayoutShapeType_shapeType_wedgeEllipseCallout;
    window['AscFormat'].LayoutShapeType_shapeType_wedgeRectCallout = LayoutShapeType_shapeType_wedgeRectCallout;
    window['AscFormat'].LayoutShapeType_shapeType_wedgeRoundRectCallout = LayoutShapeType_shapeType_wedgeRoundRectCallout;


    window['AscFormat'].AnimLvl_val_ctr = AnimLvl_val_ctr;
    window['AscFormat'].AnimLvl_val_lvl = AnimLvl_val_lvl;
    window['AscFormat'].AnimLvl_val_none = AnimLvl_val_none;

    window['AscFormat'].AnimOne_val_branch = AnimOne_val_branch;
    window['AscFormat'].AnimOne_val_none = AnimOne_val_none;
    window['AscFormat'].AnimOne_val_one = AnimOne_val_one;

    window['AscFormat'].DiagramDirection_val_norm = DiagramDirection_val_norm;
    window['AscFormat'].DiagramDirection_val_rev = DiagramDirection_val_rev;

    window['AscFormat'].HierBranch_val_hang = HierBranch_val_hang;
    window['AscFormat'].HierBranch_val_init = HierBranch_val_init;
    window['AscFormat'].HierBranch_val_l = HierBranch_val_l;
    window['AscFormat'].HierBranch_val_r = HierBranch_val_r;
    window['AscFormat'].HierBranch_val_std = HierBranch_val_std;

    window['AscFormat'].ResizeHandles_val_exact = ResizeHandles_val_exact;
    window['AscFormat'].ResizeHandles_val_rel = ResizeHandles_val_rel;

    window['AscFormat'].ClrLst_hueDir_ccw = ClrLst_hueDir_ccw;
    window['AscFormat'].ClrLst_hueDir_cw = ClrLst_hueDir_cw;
    window['AscFormat'].ClrLst_meth_cycle = ClrLst_meth_cycle;
    window['AscFormat'].ClrLst_meth_repeat = ClrLst_meth_repeat;
    window['AscFormat'].ClrLst_meth_span = ClrLst_meth_span;

    window['AscFormat'].Camera_prst_isometricBottomDown = Camera_prst_isometricBottomDown;
    window['AscFormat'].Camera_prst_isometricBottomUp = Camera_prst_isometricBottomUp;
    window['AscFormat'].Camera_prst_isometricLeftDown = Camera_prst_isometricLeftDown;
    window['AscFormat'].Camera_prst_isometricLeftUp = Camera_prst_isometricLeftUp;
    window['AscFormat'].Camera_prst_isometricOffAxis1Left = Camera_prst_isometricOffAxis1Left;
    window['AscFormat'].Camera_prst_isometricOffAxis1Right = Camera_prst_isometricOffAxis1Right;
    window['AscFormat'].Camera_prst_isometricOffAxis1Top = Camera_prst_isometricOffAxis1Top;
    window['AscFormat'].Camera_prst_isometricOffAxis2Left = Camera_prst_isometricOffAxis2Left;
    window['AscFormat'].Camera_prst_isometricOffAxis2Right = Camera_prst_isometricOffAxis2Right;
    window['AscFormat'].Camera_prst_isometricOffAxis2Top = Camera_prst_isometricOffAxis2Top;
    window['AscFormat'].Camera_prst_isometricOffAxis3Bottom = Camera_prst_isometricOffAxis3Bottom;
    window['AscFormat'].Camera_prst_isometricOffAxis3Left = Camera_prst_isometricOffAxis3Left;
    window['AscFormat'].Camera_prst_isometricOffAxis3Right = Camera_prst_isometricOffAxis3Right;
    window['AscFormat'].Camera_prst_isometricOffAxis4Bottom = Camera_prst_isometricOffAxis4Bottom;
    window['AscFormat'].Camera_prst_isometricOffAxis4Left = Camera_prst_isometricOffAxis4Left;
    window['AscFormat'].Camera_prst_isometricOffAxis4Right = Camera_prst_isometricOffAxis4Right;
    window['AscFormat'].Camera_prst_isometricRightDown = Camera_prst_isometricRightDown;
    window['AscFormat'].Camera_prst_isometricRightUp = Camera_prst_isometricRightUp;
    window['AscFormat'].Camera_prst_isometricTopDown = Camera_prst_isometricTopDown;
    window['AscFormat'].Camera_prst_isometricTopUp = Camera_prst_isometricTopUp;
    window['AscFormat'].Camera_prst_legacyObliqueBottom = Camera_prst_legacyObliqueBottom;
    window['AscFormat'].Camera_prst_legacyObliqueBottomLeft = Camera_prst_legacyObliqueBottomLeft;
    window['AscFormat'].Camera_prst_legacyObliqueBottomRight = Camera_prst_legacyObliqueBottomRight;
    window['AscFormat'].Camera_prst_legacyObliqueFront = Camera_prst_legacyObliqueFront;
    window['AscFormat'].Camera_prst_legacyObliqueLeft = Camera_prst_legacyObliqueLeft;
    window['AscFormat'].Camera_prst_legacyObliqueRight = Camera_prst_legacyObliqueRight;
    window['AscFormat'].Camera_prst_legacyObliqueTop = Camera_prst_legacyObliqueTop;
    window['AscFormat'].Camera_prst_legacyObliqueTopLeft = Camera_prst_legacyObliqueTopLeft;
    window['AscFormat'].Camera_prst_legacyObliqueTopRight = Camera_prst_legacyObliqueTopRight;
    window['AscFormat'].Camera_prst_legacyPerspectiveBottom = Camera_prst_legacyPerspectiveBottom;
    window['AscFormat'].Camera_prst_legacyPerspectiveBottomLeft = Camera_prst_legacyPerspectiveBottomLeft;
    window['AscFormat'].Camera_prst_legacyPerspectiveBottomRight = Camera_prst_legacyPerspectiveBottomRight;
    window['AscFormat'].Camera_prst_legacyPerspectiveFront = Camera_prst_legacyPerspectiveFront;
    window['AscFormat'].Camera_prst_legacyPerspectiveLeft = Camera_prst_legacyPerspectiveLeft;
    window['AscFormat'].Camera_prst_legacyPerspectiveRight = Camera_prst_legacyPerspectiveRight;
    window['AscFormat'].Camera_prst_legacyPerspectiveTop = Camera_prst_legacyPerspectiveTop;
    window['AscFormat'].Camera_prst_legacyPerspectiveTopLeft = Camera_prst_legacyPerspectiveTopLeft;
    window['AscFormat'].Camera_prst_legacyPerspectiveTopRight = Camera_prst_legacyPerspectiveTopRight;
    window['AscFormat'].Camera_prst_obliqueBottom = Camera_prst_obliqueBottom;
    window['AscFormat'].Camera_prst_obliqueBottomLeft = Camera_prst_obliqueBottomLeft;
    window['AscFormat'].Camera_prst_obliqueBottomRight = Camera_prst_obliqueBottomRight;
    window['AscFormat'].Camera_prst_obliqueLeft = Camera_prst_obliqueLeft;
    window['AscFormat'].Camera_prst_obliqueRight = Camera_prst_obliqueRight;
    window['AscFormat'].Camera_prst_obliqueTop = Camera_prst_obliqueTop;
    window['AscFormat'].Camera_prst_obliqueTopLeft = Camera_prst_obliqueTopLeft;
    window['AscFormat'].Camera_prst_obliqueTopRight = Camera_prst_obliqueTopRight;
    window['AscFormat'].Camera_prst_orthographicFront = Camera_prst_orthographicFront;
    window['AscFormat'].Camera_prst_perspectiveAbove = Camera_prst_perspectiveAbove;
    window['AscFormat'].Camera_prst_perspectiveAboveLeftFacing = Camera_prst_perspectiveAboveLeftFacing;
    window['AscFormat'].Camera_prst_perspectiveAboveRightFacing = Camera_prst_perspectiveAboveRightFacing;
    window['AscFormat'].Camera_prst_perspectiveBelow = Camera_prst_perspectiveBelow;
    window['AscFormat'].Camera_prst_perspectiveContrastingLeftFacing = Camera_prst_perspectiveContrastingLeftFacing;
    window['AscFormat'].Camera_prst_perspectiveContrastingRightFacing = Camera_prst_perspectiveContrastingRightFacing;
    window['AscFormat'].Camera_prst_perspectiveFront = Camera_prst_perspectiveFront;
    window['AscFormat'].Camera_prst_perspectiveHeroicExtremeLeftFacing = Camera_prst_perspectiveHeroicExtremeLeftFacing;
    window['AscFormat'].Camera_prst_perspectiveHeroicExtremeRightFacing = Camera_prst_perspectiveHeroicExtremeRightFacing;
    window['AscFormat'].Camera_prst_perspectiveHeroicLeftFacing = Camera_prst_perspectiveHeroicLeftFacing;
    window['AscFormat'].Camera_prst_perspectiveHeroicRightFacing = Camera_prst_perspectiveHeroicRightFacing;
    window['AscFormat'].Camera_prst_perspectiveLeft = Camera_prst_perspectiveLeft;
    window['AscFormat'].Camera_prst_perspectiveRelaxed = Camera_prst_perspectiveRelaxed;
    window['AscFormat'].Camera_prst_perspectiveRelaxedModerately = Camera_prst_perspectiveRelaxedModerately;
    window['AscFormat'].Camera_prst_perspectiveRight = Camera_prst_perspectiveRight;

    window['AscFormat'].Sp3d_prstMaterial_clear = Sp3d_prstMaterial_clear;
    window['AscFormat'].Sp3d_prstMaterial_dkEdge = Sp3d_prstMaterial_dkEdge;
    window['AscFormat'].Sp3d_prstMaterial_flat = Sp3d_prstMaterial_flat;
    window['AscFormat'].Sp3d_prstMaterial_legacyMatte = Sp3d_prstMaterial_legacyMatte;
    window['AscFormat'].Sp3d_prstMaterial_legacyMetal = Sp3d_prstMaterial_legacyMetal;
    window['AscFormat'].Sp3d_prstMaterial_legacyPlastic = Sp3d_prstMaterial_legacyPlastic;
    window['AscFormat'].Sp3d_prstMaterial_legacyWireframe = Sp3d_prstMaterial_legacyWireframe;
    window['AscFormat'].Sp3d_prstMaterial_matte = Sp3d_prstMaterial_matte;
    window['AscFormat'].Sp3d_prstMaterial_metal = Sp3d_prstMaterial_metal;
    window['AscFormat'].Sp3d_prstMaterial_plastic = Sp3d_prstMaterial_plastic;
    window['AscFormat'].Sp3d_prstMaterial_powder = Sp3d_prstMaterial_powder;
    window['AscFormat'].Sp3d_prstMaterial_softEdge = Sp3d_prstMaterial_softEdge;
    window['AscFormat'].Sp3d_prstMaterial_softmetal = Sp3d_prstMaterial_softmetal;
    window['AscFormat'].Sp3d_prstMaterial_translucentPowder = Sp3d_prstMaterial_translucentPowder;
    window['AscFormat'].Sp3d_prstMaterial_warmMatte = Sp3d_prstMaterial_warmMatte;

    window['AscFormat'].LightRig_dir_b = LightRig_dir_b;
    window['AscFormat'].LightRig_dir_bl = LightRig_dir_bl;
    window['AscFormat'].LightRig_dir_br = LightRig_dir_br;
    window['AscFormat'].LightRig_dir_l = LightRig_dir_l;
    window['AscFormat'].LightRig_dir_r = LightRig_dir_r;
    window['AscFormat'].LightRig_dir_t = LightRig_dir_t;
    window['AscFormat'].LightRig_dir_tl = LightRig_dir_tl;
    window['AscFormat'].LightRig_dir_tr = LightRig_dir_tr;

    window['AscFormat'].LightRig_rig_balanced = LightRig_rig_balanced;
    window['AscFormat'].LightRig_rig_brightRoom = LightRig_rig_brightRoom;
    window['AscFormat'].LightRig_rig_chilly = LightRig_rig_chilly;
    window['AscFormat'].LightRig_rig_contrasting = LightRig_rig_contrasting;
    window['AscFormat'].LightRig_rig_flat = LightRig_rig_flat;
    window['AscFormat'].LightRig_rig_flood = LightRig_rig_flood;
    window['AscFormat'].LightRig_rig_freezing = LightRig_rig_freezing;
    window['AscFormat'].LightRig_rig_glow = LightRig_rig_glow;
    window['AscFormat'].LightRig_rig_harsh = LightRig_rig_harsh;
    window['AscFormat'].LightRig_rig_legacyFlat1 = LightRig_rig_legacyFlat1;
    window['AscFormat'].LightRig_rig_legacyFlat2 = LightRig_rig_legacyFlat2;
    window['AscFormat'].LightRig_rig_legacyFlat3 = LightRig_rig_legacyFlat3;
    window['AscFormat'].LightRig_rig_legacyFlat4 = LightRig_rig_legacyFlat4;
    window['AscFormat'].LightRig_rig_legacyHarsh1 = LightRig_rig_legacyHarsh1;
    window['AscFormat'].LightRig_rig_legacyHarsh2 = LightRig_rig_legacyHarsh2;
    window['AscFormat'].LightRig_rig_legacyHarsh3 = LightRig_rig_legacyHarsh3;
    window['AscFormat'].LightRig_rig_legacyHarsh4 = LightRig_rig_legacyHarsh4;
    window['AscFormat'].LightRig_rig_legacyNormal1 = LightRig_rig_legacyNormal1;
    window['AscFormat'].LightRig_rig_legacyNormal2 = LightRig_rig_legacyNormal2;
    window['AscFormat'].LightRig_rig_legacyNormal3 = LightRig_rig_legacyNormal3;
    window['AscFormat'].LightRig_rig_legacyNormal4 = LightRig_rig_legacyNormal4;
    window['AscFormat'].LightRig_rig_morning = LightRig_rig_morning;
    window['AscFormat'].LightRig_rig_soft = LightRig_rig_soft;
    window['AscFormat'].LightRig_rig_sunrise = LightRig_rig_sunrise;
    window['AscFormat'].LightRig_rig_sunset = LightRig_rig_sunset;
    window['AscFormat'].LightRig_rig_threePt = LightRig_rig_threePt;
    window['AscFormat'].LightRig_rig_twoPt = LightRig_rig_twoPt;

    window['AscFormat'].Bevel_prst_angle = Bevel_prst_angle;
    window['AscFormat'].Bevel_prst_artDeco = Bevel_prst_artDeco;
    window['AscFormat'].Bevel_prst_circle = Bevel_prst_circle;
    window['AscFormat'].Bevel_prst_convex = Bevel_prst_convex;
    window['AscFormat'].Bevel_prst_coolSlant = Bevel_prst_coolSlant;
    window['AscFormat'].Bevel_prst_cross = Bevel_prst_cross;
    window['AscFormat'].Bevel_prst_divot = Bevel_prst_divot;
    window['AscFormat'].Bevel_prst_hardEdge = Bevel_prst_hardEdge;
    window['AscFormat'].Bevel_prst_relaxedInset = Bevel_prst_relaxedInset;
    window['AscFormat'].Bevel_prst_riblet = Bevel_prst_riblet;
    window['AscFormat'].Bevel_prst_slope = Bevel_prst_slope;
    window['AscFormat'].Bevel_prst_softRound = Bevel_prst_softRound;

    window['AscFormat'].ParameterVal_arrowheadStyle_arr = ParameterVal_arrowheadStyle_arr;
    window['AscFormat'].ParameterVal_arrowheadStyle_auto = ParameterVal_arrowheadStyle_auto;
    window['AscFormat'].ParameterVal_arrowheadStyle_noArr = ParameterVal_arrowheadStyle_noArr;
    window['AscFormat'].ParameterVal_autoTextRotation_grav = ParameterVal_autoTextRotation_grav;
    window['AscFormat'].ParameterVal_autoTextRotation_none = ParameterVal_autoTextRotation_none;
    window['AscFormat'].ParameterVal_autoTextRotation_upr = ParameterVal_autoTextRotation_upr;
    window['AscFormat'].ParameterVal_bendPoint_beg = ParameterVal_bendPoint_beg;
    window['AscFormat'].ParameterVal_bendPoint_def = ParameterVal_bendPoint_def;
    window['AscFormat'].ParameterVal_bendPoint_end = ParameterVal_bendPoint_end;
    window['AscFormat'].ParameterVal_breakpoint_bal = ParameterVal_breakpoint_bal;
    window['AscFormat'].ParameterVal_breakpoint_endCnv = ParameterVal_breakpoint_endCnv;
    window['AscFormat'].ParameterVal_breakpoint_fixed = ParameterVal_breakpoint_fixed;
    window['AscFormat'].ParameterVal_centerShapeMapping_fNode = ParameterVal_centerShapeMapping_fNode;
    window['AscFormat'].ParameterVal_centerShapeMapping_none = ParameterVal_centerShapeMapping_none;
    window['AscFormat'].ParameterVal_childAlignment_b = ParameterVal_childAlignment_b;
    window['AscFormat'].ParameterVal_childAlignment_l = ParameterVal_childAlignment_l;
    window['AscFormat'].ParameterVal_childAlignment_r = ParameterVal_childAlignment_r;
    window['AscFormat'].ParameterVal_childAlignment_t = ParameterVal_childAlignment_t;
    window['AscFormat'].ParameterVal_childDirection_horz = ParameterVal_childDirection_horz;
    window['AscFormat'].ParameterVal_childDirection_vert = ParameterVal_childDirection_vert;
    window['AscFormat'].ParameterVal_connectorDimension_1D = ParameterVal_connectorDimension_1D;
    window['AscFormat'].ParameterVal_connectorDimension_2D = ParameterVal_connectorDimension_2D;
    window['AscFormat'].ParameterVal_connectorDimension_cust = ParameterVal_connectorDimension_cust;
    window['AscFormat'].ParameterVal_connectorPoint_auto = ParameterVal_connectorPoint_auto;
    window['AscFormat'].ParameterVal_connectorPoint_bCtr = ParameterVal_connectorPoint_bCtr;
    window['AscFormat'].ParameterVal_connectorPoint_bL = ParameterVal_connectorPoint_bL;
    window['AscFormat'].ParameterVal_connectorPoint_bR = ParameterVal_connectorPoint_bR;
    window['AscFormat'].ParameterVal_connectorPoint_ctr = ParameterVal_connectorPoint_ctr;
    window['AscFormat'].ParameterVal_connectorPoint_midL = ParameterVal_connectorPoint_midL;
    window['AscFormat'].ParameterVal_connectorPoint_midR = ParameterVal_connectorPoint_midR;
    window['AscFormat'].ParameterVal_connectorPoint_radial = ParameterVal_connectorPoint_radial;
    window['AscFormat'].ParameterVal_connectorPoint_tCtr = ParameterVal_connectorPoint_tCtr;
    window['AscFormat'].ParameterVal_connectorPoint_tL = ParameterVal_connectorPoint_tL;
    window['AscFormat'].ParameterVal_connectorPoint_tR = ParameterVal_connectorPoint_tR;
    window['AscFormat'].ParameterVal_connectorRouting_bend = ParameterVal_connectorRouting_bend;
    window['AscFormat'].ParameterVal_connectorRouting_curve = ParameterVal_connectorRouting_curve;
    window['AscFormat'].ParameterVal_connectorRouting_longCurve = ParameterVal_connectorRouting_longCurve;
    window['AscFormat'].ParameterVal_connectorRouting_stra = ParameterVal_connectorRouting_stra;
    window['AscFormat'].ParameterVal_continueDirection_revDir = ParameterVal_continueDirection_revDir;
    window['AscFormat'].ParameterVal_continueDirection_sameDir = ParameterVal_continueDirection_sameDir;
    window['AscFormat'].ParameterVal_horizontalAlignment_ctr = ParameterVal_horizontalAlignment_ctr;
    window['AscFormat'].ParameterVal_horizontalAlignment_l = ParameterVal_horizontalAlignment_l;
    window['AscFormat'].ParameterVal_horizontalAlignment_none = ParameterVal_horizontalAlignment_none;
    window['AscFormat'].ParameterVal_horizontalAlignment_r = ParameterVal_horizontalAlignment_r;
    window['AscFormat'].ParameterVal_diagramTextAlignment_ctr = ParameterVal_diagramTextAlignment_ctr;
    window['AscFormat'].ParameterVal_diagramTextAlignment_l = ParameterVal_diagramTextAlignment_l;
    window['AscFormat'].ParameterVal_diagramTextAlignment_r = ParameterVal_diagramTextAlignment_r;
    window['AscFormat'].ParameterVal_fallbackDimension_1D = ParameterVal_fallbackDimension_1D;
    window['AscFormat'].ParameterVal_fallbackDimension_2D = ParameterVal_fallbackDimension_2D;
    window['AscFormat'].ParameterVal_flowDirection_col = ParameterVal_flowDirection_col;
    window['AscFormat'].ParameterVal_flowDirection_row = ParameterVal_flowDirection_row;
    window['AscFormat'].ParameterVal_growDirection_bL = ParameterVal_growDirection_bL;
    window['AscFormat'].ParameterVal_growDirection_bR = ParameterVal_growDirection_bR;
    window['AscFormat'].ParameterVal_growDirection_tL = ParameterVal_growDirection_tL;
    window['AscFormat'].ParameterVal_growDirection_tR = ParameterVal_growDirection_tR;
    window['AscFormat'].ParameterVal_hierarchyAlignment_bCtrCh = ParameterVal_hierarchyAlignment_bCtrCh;
    window['AscFormat'].ParameterVal_hierarchyAlignment_bCtrDes = ParameterVal_hierarchyAlignment_bCtrDes;
    window['AscFormat'].ParameterVal_hierarchyAlignment_bL = ParameterVal_hierarchyAlignment_bL;
    window['AscFormat'].ParameterVal_hierarchyAlignment_bR = ParameterVal_hierarchyAlignment_bR;
    window['AscFormat'].ParameterVal_hierarchyAlignment_lB = ParameterVal_hierarchyAlignment_lB;
    window['AscFormat'].ParameterVal_hierarchyAlignment_lCtrCh = ParameterVal_hierarchyAlignment_lCtrCh;
    window['AscFormat'].ParameterVal_hierarchyAlignment_lCtrDes = ParameterVal_hierarchyAlignment_lCtrDes;
    window['AscFormat'].ParameterVal_hierarchyAlignment_lT = ParameterVal_hierarchyAlignment_lT;
    window['AscFormat'].ParameterVal_hierarchyAlignment_rB = ParameterVal_hierarchyAlignment_rB;
    window['AscFormat'].ParameterVal_hierarchyAlignment_rCtrCh = ParameterVal_hierarchyAlignment_rCtrCh;
    window['AscFormat'].ParameterVal_hierarchyAlignment_rCtrDes = ParameterVal_hierarchyAlignment_rCtrDes;
    window['AscFormat'].ParameterVal_hierarchyAlignment_rT = ParameterVal_hierarchyAlignment_rT;
    window['AscFormat'].ParameterVal_hierarchyAlignment_tCtrCh = ParameterVal_hierarchyAlignment_tCtrCh;
    window['AscFormat'].ParameterVal_hierarchyAlignment_tCtrDes = ParameterVal_hierarchyAlignment_tCtrDes;
    window['AscFormat'].ParameterVal_hierarchyAlignment_tL = ParameterVal_hierarchyAlignment_tL;
    window['AscFormat'].ParameterVal_hierarchyAlignment_tR = ParameterVal_hierarchyAlignment_tR;
    window['AscFormat'].ParameterVal_linearDirection_fromB = ParameterVal_linearDirection_fromB;
    window['AscFormat'].ParameterVal_linearDirection_fromL = ParameterVal_linearDirection_fromL;
    window['AscFormat'].ParameterVal_linearDirection_fromR = ParameterVal_linearDirection_fromR;
    window['AscFormat'].ParameterVal_linearDirection_fromT = ParameterVal_linearDirection_fromT;
    window['AscFormat'].ParameterVal_nodeHorizontalAlignment_ctr = ParameterVal_nodeHorizontalAlignment_ctr;
    window['AscFormat'].ParameterVal_nodeHorizontalAlignment_l = ParameterVal_nodeHorizontalAlignment_l;
    window['AscFormat'].ParameterVal_nodeHorizontalAlignment_r = ParameterVal_nodeHorizontalAlignment_r;
    window['AscFormat'].ParameterVal_nodeVerticalAlignment_b = ParameterVal_nodeVerticalAlignment_b;
    window['AscFormat'].ParameterVal_nodeVerticalAlignment_mid = ParameterVal_nodeVerticalAlignment_mid;
    window['AscFormat'].ParameterVal_nodeVerticalAlignment_t = ParameterVal_nodeVerticalAlignment_t;
    window['AscFormat'].ParameterVal_offset_ctr = ParameterVal_offset_ctr;
    window['AscFormat'].ParameterVal_offset_off = ParameterVal_offset_off;
    window['AscFormat'].ParameterVal_pyramidAccentPosition_aft = ParameterVal_pyramidAccentPosition_aft;
    window['AscFormat'].ParameterVal_pyramidAccentPosition_bef = ParameterVal_pyramidAccentPosition_bef;
    window['AscFormat'].ParameterVal_pyramidAccentTextMargin_stack = ParameterVal_pyramidAccentTextMargin_stack;
    window['AscFormat'].ParameterVal_pyramidAccentTextMargin_step = ParameterVal_pyramidAccentTextMargin_step;
    window['AscFormat'].ParameterVal_rotationPath_alongPath = ParameterVal_rotationPath_alongPath;
    window['AscFormat'].ParameterVal_rotationPath_none = ParameterVal_rotationPath_none;
    window['AscFormat'].ParameterVal_secondaryChildAlignment_b = ParameterVal_secondaryChildAlignment_b;
    window['AscFormat'].ParameterVal_secondaryChildAlignment_l = ParameterVal_secondaryChildAlignment_l;
    window['AscFormat'].ParameterVal_secondaryChildAlignment_none = ParameterVal_secondaryChildAlignment_none;
    window['AscFormat'].ParameterVal_secondaryChildAlignment_r = ParameterVal_secondaryChildAlignment_r;
    window['AscFormat'].ParameterVal_secondaryChildAlignment_t = ParameterVal_secondaryChildAlignment_t;
    window['AscFormat'].ParameterVal_secondaryLinearDirection_fromB = ParameterVal_secondaryLinearDirection_fromB;
    window['AscFormat'].ParameterVal_secondaryLinearDirection_fromL = ParameterVal_secondaryLinearDirection_fromL;
    window['AscFormat'].ParameterVal_secondaryLinearDirection_fromR = ParameterVal_secondaryLinearDirection_fromR;
    window['AscFormat'].ParameterVal_secondaryLinearDirection_fromT = ParameterVal_secondaryLinearDirection_fromT;
    window['AscFormat'].ParameterVal_secondaryLinearDirection_none = ParameterVal_secondaryLinearDirection_none;
    window['AscFormat'].ParameterVal_startingElement_node = ParameterVal_startingElement_node;
    window['AscFormat'].ParameterVal_startingElement_trans = ParameterVal_startingElement_trans;
    window['AscFormat'].ParameterVal_textAnchorHorizontal_ctr = ParameterVal_textAnchorHorizontal_ctr;
    window['AscFormat'].ParameterVal_textAnchorHorizontal_none = ParameterVal_textAnchorHorizontal_none;
    window['AscFormat'].ParameterVal_textAnchorVertical_b = ParameterVal_textAnchorVertical_b;
    window['AscFormat'].ParameterVal_textAnchorVertical_mid = ParameterVal_textAnchorVertical_mid;
    window['AscFormat'].ParameterVal_textAnchorVertical_t = ParameterVal_textAnchorVertical_t;
    window['AscFormat'].ParameterVal_textBlockDirection_horz = ParameterVal_textBlockDirection_horz;
    window['AscFormat'].ParameterVal_textBlockDirection_vert = ParameterVal_textBlockDirection_vert;
    window['AscFormat'].ParameterVal_textDirection_fromB = ParameterVal_textDirection_fromB;
    window['AscFormat'].ParameterVal_textDirection_fromT = ParameterVal_textDirection_fromT;
    window['AscFormat'].ParameterVal_verticalAlignment_b = ParameterVal_verticalAlignment_b;
    window['AscFormat'].ParameterVal_verticalAlignment_mid = ParameterVal_verticalAlignment_mid;
    window['AscFormat'].ParameterVal_verticalAlignment_none = ParameterVal_verticalAlignment_none;
    window['AscFormat'].ParameterVal_verticalAlignment_t = ParameterVal_verticalAlignment_t;

    window['AscFormat'].Coordinate_universalMeasure_cm = Coordinate_universalMeasure_cm;
    window['AscFormat'].Coordinate_universalMeasure_mm = Coordinate_universalMeasure_mm;
    window['AscFormat'].Coordinate_universalMeasure_in = Coordinate_universalMeasure_in;
    window['AscFormat'].Coordinate_universalMeasure_pt = Coordinate_universalMeasure_pt;
    window['AscFormat'].Coordinate_universalMeasure_pc = Coordinate_universalMeasure_pc;
    window['AscFormat'].Coordinate_universalMeasure_pi = Coordinate_universalMeasure_pi;

    window['AscFormat'].EChOrder_chOrderB = EChOrder_chOrderB;
    window['AscFormat'].EChOrder_chOrderT = EChOrder_chOrderT;

	  window['AscCommon'].smartArtContentFillingType_parentWithChildren = smartArtContentFillingType_parentWithChildren;
	  window['AscCommon'].smartArtContentFillingType_onlyChildren = smartArtContentFillingType_onlyChildren;
	  window['AscCommon'].smartArtContentFillingType_onlyParent = smartArtContentFillingType_onlyParent;

  })(window)
