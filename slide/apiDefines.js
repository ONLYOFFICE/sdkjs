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

/** @enum {number} */
var c_oAscZoomType = {
	Current  : 0,
	FitWidth : 1,
	FitPage  : 2
};

/** @enum {number} */
var c_oAscCollaborativeMarksShowType = {
	All         : 0,
	LastChanges : 1
};

/** @enum {number} */
var c_oAscVertAlignJc = {
	Top    : 0x00, // var vertalignjc_Top    = 0x00;
	Center : 0x01, // var vertalignjc_Center = 0x01;
	Bottom : 0x02  // var vertalignjc_Bottom = 0x02
};

/** @enum {number} */
var c_oAscAlignType = {
	LEFT    : 0,
	CENTER  : 1,
	RIGHT   : 2,
	JUSTIFY : 3,
	TOP     : 4,
	MIDDLE  : 5,
	BOTTOM  : 6
};

/** @enum {number} */
var c_oAscContextMenuTypes = {
	Main       : 0,
	Thumbnails : 1
};

var THEME_THUMBNAIL_WIDTH   = 180;
var THEME_THUMBNAIL_HEIGHT  = 135;
var LAYOUT_THUMBNAIL_WIDTH  = 180;
var LAYOUT_THUMBNAIL_HEIGHT = 135;

/** @enum {number} */
var c_oAscTableSelectionType = {
	Cell   : 0,
	Row    : 1,
	Column : 2,
	Table  : 3
};

/** @enum {number} */
var c_oAscAlignShapeType = {
	ALIGN_LEFT   : 0,
	ALIGN_RIGHT  : 1,
	ALIGN_TOP    : 2,
	ALIGN_BOTTOM : 3,
	ALIGN_CENTER : 4,
	ALIGN_MIDDLE : 5
};

/** @enum {number} */
var c_oAscTableLayout = {
	AutoFit : 0x00,
	Fixed   : 0x01
};

/** @enum {number} */
var c_oAscSlideTransitionTypes = {
	None    : 0,
	Fade    : 1,
	Push    : 2,
	Wipe    : 3,
	Split   : 4,
	UnCover : 5,
	Cover   : 6,
	Clock   : 7,
	Zoom    : 8
};

/** @enum {number} */
var c_oAscSlideTransitionParams = {
	Fade_Smoothly      : 0,
	Fade_Through_Black : 1,

	Param_Left        : 0,
	Param_Top         : 1,
	Param_Right       : 2,
	Param_Bottom      : 3,
	Param_TopLeft     : 4,
	Param_TopRight    : 5,
	Param_BottomLeft  : 6,
	Param_BottomRight : 7,

	Split_VerticalIn    : 8,
	Split_VerticalOut   : 9,
	Split_HorizontalIn  : 10,
	Split_HorizontalOut : 11,

	Clock_Clockwise        : 0,
	Clock_Counterclockwise : 1,
	Clock_Wedge            : 2,

	Zoom_In        : 0,
	Zoom_Out       : 1,
	Zoom_AndRotate : 2
};

/** @enum {number} */
var c_oAscLockTypeElemPresentation = {
	Object       : 1,
	Slide        : 2,
	Presentation : 3
};

var c_oSerFormat = {
	Version   : 1,
	Signature : "PPTY"
};

var c_oAscPresentationShortcutType = {
	EditSelectAll   : 1,
	EditUndo        : 2,
	EditRedo        : 3,
	Cut             : 4,
	Copy            : 5,
	Paste           : 6,
	Duplicate       : 7,
	Print           : 8,
	Save            : 9,
	ShowContextMenu : 10
};

var TABLE_STYLE_WIDTH_PIX  = 72;
var TABLE_STYLE_HEIGHT_PIX = 52;

//------------------------------------------------------------export---------------------------------------------------
var prot;
window['Asc'] = window['Asc'] || {};

prot = window['Asc']['c_oAscCollaborativeMarksShowType'] = c_oAscCollaborativeMarksShowType;
prot['All']         = c_oAscCollaborativeMarksShowType.All;
prot['LastChanges'] = c_oAscCollaborativeMarksShowType.LastChanges;

prot = window['Asc']['c_oAscVertAlignJc'] = c_oAscVertAlignJc;
prot['Top']    = c_oAscVertAlignJc.Top;
prot['Center'] = c_oAscVertAlignJc.Center;
prot['Bottom'] = c_oAscVertAlignJc.Bottom;

prot = window['Asc']['c_oAscContextMenuTypes'] = window['Asc'].c_oAscContextMenuTypes = c_oAscContextMenuTypes;
prot['Main']       = c_oAscContextMenuTypes.Main;
prot['Thumbnails'] = c_oAscContextMenuTypes.Thumbnails;

prot = window['Asc']['c_oAscAlignShapeType'] = c_oAscAlignShapeType;
prot['ALIGN_LEFT']   = c_oAscAlignShapeType.ALIGN_LEFT;
prot['ALIGN_RIGHT']  = c_oAscAlignShapeType.ALIGN_RIGHT;
prot['ALIGN_TOP']    = c_oAscAlignShapeType.ALIGN_TOP;
prot['ALIGN_BOTTOM'] = c_oAscAlignShapeType.ALIGN_BOTTOM;
prot['ALIGN_CENTER'] = c_oAscAlignShapeType.ALIGN_CENTER;
prot['ALIGN_MIDDLE'] = c_oAscAlignShapeType.ALIGN_MIDDLE;

prot = window['Asc']['c_oAscTableLayout'] = c_oAscTableLayout;
prot['AutoFit'] = c_oAscTableLayout.AutoFit;
prot['Fixed']   = c_oAscTableLayout.Fixed;

prot = window['Asc']['c_oAscSlideTransitionTypes'] = c_oAscSlideTransitionTypes;
prot['None']    = c_oAscSlideTransitionTypes.None;
prot['Fade']    = c_oAscSlideTransitionTypes.Fade;
prot['Push']    = c_oAscSlideTransitionTypes.Push;
prot['Wipe']    = c_oAscSlideTransitionTypes.Wipe;
prot['Split']   = c_oAscSlideTransitionTypes.Split;
prot['UnCover'] = c_oAscSlideTransitionTypes.UnCover;
prot['Cover']   = c_oAscSlideTransitionTypes.Cover;
prot['Clock']   = c_oAscSlideTransitionTypes.Clock;
prot['Zoom']    = c_oAscSlideTransitionTypes.Zoom;

prot = window['Asc']['c_oAscSlideTransitionParams'] = c_oAscSlideTransitionParams;
prot['Fade_Smoothly']          = c_oAscSlideTransitionParams.Fade_Smoothly;
prot['Fade_Through_Black']     = c_oAscSlideTransitionParams.Fade_Through_Black;
prot['Param_Left']             = c_oAscSlideTransitionParams.Param_Left;
prot['Param_Top']              = c_oAscSlideTransitionParams.Param_Top;
prot['Param_Right']            = c_oAscSlideTransitionParams.Param_Right;
prot['Param_Bottom']           = c_oAscSlideTransitionParams.Param_Bottom;
prot['Param_TopLeft']          = c_oAscSlideTransitionParams.Param_TopLeft;
prot['Param_TopRight']         = c_oAscSlideTransitionParams.Param_TopRight;
prot['Param_BottomLeft']       = c_oAscSlideTransitionParams.Param_BottomLeft;
prot['Param_BottomRight']      = c_oAscSlideTransitionParams.Param_BottomRight;
prot['Split_VerticalIn']       = c_oAscSlideTransitionParams.Split_VerticalIn;
prot['Split_VerticalOut']      = c_oAscSlideTransitionParams.Split_VerticalOut;
prot['Split_HorizontalIn']     = c_oAscSlideTransitionParams.Split_HorizontalIn;
prot['Split_HorizontalOut']    = c_oAscSlideTransitionParams.Split_HorizontalOut;
prot['Clock_Clockwise']        = c_oAscSlideTransitionParams.Clock_Clockwise;
prot['Clock_Counterclockwise'] = c_oAscSlideTransitionParams.Clock_Counterclockwise;
prot['Clock_Wedge']            = c_oAscSlideTransitionParams.Clock_Wedge;
prot['Zoom_In']                = c_oAscSlideTransitionParams.Zoom_In;
prot['Zoom_Out']               = c_oAscSlideTransitionParams.Zoom_Out;
prot['Zoom_AndRotate']         = c_oAscSlideTransitionParams.Zoom_AndRotate;

prot = window['Asc']['c_oAscPresentationShortcutType'] = window['Asc'].c_oAscPresentationShortcutType = c_oAscPresentationShortcutType;
prot['EditSelectAll']   = c_oAscPresentationShortcutType.EditSelectAll;
prot['EditUndo']        = c_oAscPresentationShortcutType.EditUndo;
prot['EditRedo']        = c_oAscPresentationShortcutType.EditRedo;
prot['Cut']             = c_oAscPresentationShortcutType.Cut;
prot['Copy']            = c_oAscPresentationShortcutType.Copy;
prot['Paste']           = c_oAscPresentationShortcutType.Paste;
prot['Duplicate']       = c_oAscPresentationShortcutType.Duplicate;
prot['Print']           = c_oAscPresentationShortcutType.Print;
prot['Save']            = c_oAscPresentationShortcutType.Save;
prot['ShowContextMenu'] = c_oAscPresentationShortcutType.ShowContextMenu;

window['AscCommon']                = window['AscCommon'] || {};
window['AscCommon'].c_oSerFormat   = c_oSerFormat;
window['AscCommon'].CurFileVersion = c_oSerFormat.Version;

window['AscFormat'] = window['AscFormat'] || {};

//classes of animations
window['AscFormat'].PRESET_CLASS_EMPH = window['AscFormat']["PRESET_CLASS_EMPH"] = 0;
window['AscFormat'].PRESET_CLASS_ENTR = window['AscFormat']["PRESET_CLASS_ENTR"] = 1;
window['AscFormat'].PRESET_CLASS_EXIT = window['AscFormat']["PRESET_CLASS_EXIT"] = 2;
window['AscFormat'].PRESET_CLASS_MEDIACALL = window['AscFormat']["PRESET_CLASS_MEDIACALL"] = 3;
window['AscFormat'].PRESET_CLASS_PATH = window['AscFormat']["PRESET_CLASS_PATH"] = 4;
window['AscFormat'].PRESET_CLASS_VERB = window['AscFormat']["PRESET_CLASS_VERB"] = 5;

//presets of animations
window['AscFormat'].ANIM_PRESET_NONE = window['AscFormat']["ANIM_PRESET_NONE"] = -2;
window['AscFormat'].ANIM_PRESET_MULTIPLE = window['AscFormat']["ANIM_PRESET_MULTIPLE"] = -1;
window['AscFormat'].MOTION_CUSTOM_PATH = window['AscFormat']["MOTION_CUSTOM_PATH"] = 0;
window['AscFormat'].MOTION_CIRCLE = window['AscFormat']["MOTION_CIRCLE"] = 1;
window['AscFormat'].MOTION_RIGHT_TRIANGLE = window['AscFormat']["MOTION_RIGHT_TRIANGLE"] = 2;
window['AscFormat'].MOTION_DIAMOND = window['AscFormat']["MOTION_DIAMOND"] = 3;
window['AscFormat'].MOTION_HEXAGON = window['AscFormat']["MOTION_HEXAGON"] = 4;
window['AscFormat'].MOTION_PATH_5_POINT_STAR = window['AscFormat']["MOTION_PATH_5_POINT_STAR"] = 5;
window['AscFormat'].MOTION_CRESCENT_MOON = window['AscFormat']["MOTION_CRESCENT_MOON"] = 6;
window['AscFormat'].MOTION_SQUARE = window['AscFormat']["MOTION_SQUARE"] = 7;
window['AscFormat'].MOTION_TRAPEZOID = window['AscFormat']["MOTION_TRAPEZOID"] = 8;
window['AscFormat'].MOTION_HEART = window['AscFormat']["MOTION_HEART"] = 9;
window['AscFormat'].MOTION_OCTAGON = window['AscFormat']["MOTION_OCTAGON"] = 10;
window['AscFormat'].MOTION_PATH_6_POINT_STAR = window['AscFormat']["MOTION_PATH_6_POINT_STAR"] = 11;
window['AscFormat'].MOTION_FOOTBALL = window['AscFormat']["MOTION_FOOTBALL"] = 12;
window['AscFormat'].MOTION_EQUAL_TRIANGLE = window['AscFormat']["MOTION_EQUAL_TRIANGLE"] = 13;
window['AscFormat'].MOTION_PARALLELOGRAM = window['AscFormat']["MOTION_PARALLELOGRAM"] = 14;
window['AscFormat'].MOTION_PENTAGON = window['AscFormat']["MOTION_PENTAGON"] = 15;
window['AscFormat'].MOTION_PATH_4_POINT_STAR = window['AscFormat']["MOTION_PATH_4_POINT_STAR"] = 16;
window['AscFormat'].MOTION_PATH_8_POINT_STAR = window['AscFormat']["MOTION_PATH_8_POINT_STAR"] = 17;
window['AscFormat'].MOTION_TEARDROP = window['AscFormat']["MOTION_TEARDROP"] = 18;
window['AscFormat'].MOTION_POINTY_STAR = window['AscFormat']["MOTION_POINTY_STAR"] = 19;
window['AscFormat'].MOTION_CURVED_SQUARE = window['AscFormat']["MOTION_CURVED_SQUARE"] = 20;
window['AscFormat'].MOTION_CURVED_X = window['AscFormat']["MOTION_CURVED_X"] = 21;
window['AscFormat'].MOTION_VERTICAL_FIGURE_8 = window['AscFormat']["MOTION_VERTICAL_FIGURE_8"] = 22;
window['AscFormat'].MOTION_CURVY_STAR = window['AscFormat']["MOTION_CURVY_STAR"] = 23;
window['AscFormat'].MOTION_LOOP_DE_LOOP = window['AscFormat']["MOTION_LOOP_DE_LOOP"] = 24;
window['AscFormat'].MOTION_HORIZONTAL_FIGURE_8_FOUR = window['AscFormat']["MOTION_HORIZONTAL_FIGURE_8_FOUR"] = 26;
window['AscFormat'].MOTION_PEANUT = window['AscFormat']["MOTION_PEANUT"] = 27;
window['AscFormat'].MOTION_FIGURE_8_FOUR = window['AscFormat']["MOTION_FIGURE_8_FOUR"] = 28;
window['AscFormat'].MOTION_NEUTRON = window['AscFormat']["MOTION_NEUTRON"] = 29;
window['AscFormat'].MOTION_SWOOSH = window['AscFormat']["MOTION_SWOOSH"] = 30;
window['AscFormat'].MOTION_BEAN = window['AscFormat']["MOTION_BEAN"] = 31;
window['AscFormat'].MOTION_PLUS = window['AscFormat']["MOTION_PLUS"] = 32;
window['AscFormat'].MOTION_INVERTED_TRIANGLE = window['AscFormat']["MOTION_INVERTED_TRIANGLE"] = 33;
window['AscFormat'].MOTION_INVERTED_SQUARE = window['AscFormat']["MOTION_INVERTED_SQUARE"] = 34;
window['AscFormat'].MOTION_LEFT = window['AscFormat']["MOTION_LEFT"] = 35;
window['AscFormat'].MOTION_TURN_DOWN_RIGHT = window['AscFormat']["MOTION_TURN_DOWN_RIGHT"] = 36;
window['AscFormat'].MOTION_ARC_DOWN = window['AscFormat']["MOTION_ARC_DOWN"] = 37;
window['AscFormat'].MOTION_ZIGZAG = window['AscFormat']["MOTION_ZIGZAG"] = 38;
window['AscFormat'].MOTION_S_CURVE_2 = window['AscFormat']["MOTION_S_CURVE_2"] = 39;
window['AscFormat'].MOTION_SINE_WAVE = window['AscFormat']["MOTION_SINE_WAVE"] = 40;
window['AscFormat'].MOTION_BOUNCE_LEFT = window['AscFormat']["MOTION_BOUNCE_LEFT"] = 41;
window['AscFormat'].MOTION_DOWN = window['AscFormat']["MOTION_DOWN"] = 42;
window['AscFormat'].MOTION_TURN_UP = window['AscFormat']["MOTION_TURN_UP"] = 43;
window['AscFormat'].MOTION_ARC_UP = window['AscFormat']["MOTION_ARC_UP"] = 44;
window['AscFormat'].MOTION_HEARTBEAT = window['AscFormat']["MOTION_HEARTBEAT"] = 45;
window['AscFormat'].MOTION_SINE_SPIRAL_RIGHT = window['AscFormat']["MOTION_SINE_SPIRAL_RIGHT"] = 46;
window['AscFormat'].MOTION_WAVE = window['AscFormat']["MOTION_WAVE"] = 47;
window['AscFormat'].MOTION_CURVY_LEFT = window['AscFormat']["MOTION_CURVY_LEFT"] = 48;
window['AscFormat'].MOTION_DIAGONAL_DOWN_RIGHT = window['AscFormat']["MOTION_DIAGONAL_DOWN_RIGHT"] = 49;
window['AscFormat'].MOTION_TURN_DOWN = window['AscFormat']["MOTION_TURN_DOWN"] = 50;
window['AscFormat'].MOTION_ARC_LEFT = window['AscFormat']["MOTION_ARC_LEFT"] = 51;
window['AscFormat'].MOTION_FUNNEL = window['AscFormat']["MOTION_FUNNEL"] = 52;
window['AscFormat'].MOTION_SPRING = window['AscFormat']["MOTION_SPRING"] = 53;
window['AscFormat'].MOTION_BOUNCE_RIGHT = window['AscFormat']["MOTION_BOUNCE_RIGHT"] = 54;
window['AscFormat'].MOTION_SINE_SPIRAL_LEFT = window['AscFormat']["MOTION_SINE_SPIRAL_LEFT"] = 55;
window['AscFormat'].MOTION_DIAGONAL_UP_RIGHT = window['AscFormat']["MOTION_DIAGONAL_UP_RIGHT"] = 56;
window['AscFormat'].MOTION_TURN_UP_RIGHT = window['AscFormat']["MOTION_TURN_UP_RIGHT"] = 57;
window['AscFormat'].MOTION_ARC_RIGHT = window['AscFormat']["MOTION_ARC_RIGHT"] = 58;
window['AscFormat'].MOTION_S_CURVE_1 = window['AscFormat']["MOTION_S_CURVE_1"] = 59;
window['AscFormat'].MOTION_DECAYING_WAVE = window['AscFormat']["MOTION_DECAYING_WAVE"] = 60;
window['AscFormat'].MOTION_CURVY_RIGHT = window['AscFormat']["MOTION_CURVY_RIGHT"] = 61;
window['AscFormat'].MOTION_STAIRS_DOWN = window['AscFormat']["MOTION_STAIRS_DOWN"] = 62;
window['AscFormat'].MOTION_RIGHT = window['AscFormat']["MOTION_RIGHT"] = 63;
window['AscFormat'].MOTION_UP = window['AscFormat']["MOTION_UP"] = 64;

window['AscFormat'].EXIT_DISAPPEAR = window['AscFormat']["EXIT_DISAPPEAR"] = 1;
window['AscFormat'].EXIT_FLY_OUT_TO = window['AscFormat']["EXIT_FLY_OUT_TO"] = 2;
window['AscFormat'].EXIT_BLINDS = window['AscFormat']["EXIT_BLINDS"] = 3;
window['AscFormat'].EXIT_BOX = window['AscFormat']["EXIT_BOX"] = 4;
window['AscFormat'].EXIT_CHECKERBOARD = window['AscFormat']["EXIT_CHECKERBOARD"] = 5;
window['AscFormat'].EXIT_CIRCLE = window['AscFormat']["EXIT_CIRCLE"] = 6;
window['AscFormat'].EXIT_DIAMOND = window['AscFormat']["EXIT_DIAMOND"] = 8;
window['AscFormat'].EXIT_DISSOLVE_OUT = window['AscFormat']["EXIT_DISSOLVE_OUT"] = 9;
window['AscFormat'].EXIT_FADE = window['AscFormat']["EXIT_FADE"] = 10;
window['AscFormat'].EXIT_PEEK_OUT_TO = window['AscFormat']["EXIT_PEEK_OUT_TO"] = 12;
window['AscFormat'].EXIT_PLUS = window['AscFormat']["EXIT_PLUS"] = 13;
window['AscFormat'].EXIT_RANDOM_BARS = window['AscFormat']["EXIT_RANDOM_BARS"] = 14;
window['AscFormat'].EXIT_SPIRAL_OUT = window['AscFormat']["EXIT_SPIRAL_OUT"] = 15;
window['AscFormat'].EXIT_SPLIT = window['AscFormat']["EXIT_SPLIT"] = 16;
window['AscFormat'].EXIT_COLLAPSE = window['AscFormat']["EXIT_COLLAPSE"] = 17;
window['AscFormat'].EXIT_STRIPS = window['AscFormat']["EXIT_STRIPS"] = 18;
window['AscFormat'].EXIT_BASIC_SWIVEL = window['AscFormat']["EXIT_BASIC_SWIVEL"] = 19;
window['AscFormat'].EXIT_WEDGE = window['AscFormat']["EXIT_WEDGE"] = 20;
window['AscFormat'].EXIT_WHEEL = window['AscFormat']["EXIT_WHEEL"] = 21;
window['AscFormat'].EXIT_WIPE_FROM = window['AscFormat']["EXIT_WIPE_FROM"] = 22;
window['AscFormat'].EXIT_BASIC_ZOOM = window['AscFormat']["EXIT_BASIC_ZOOM"] = 23;
window['AscFormat'].EXIT_BOOMERANG = window['AscFormat']["EXIT_BOOMERANG"] = 25;
window['AscFormat'].EXIT_BOUNCE = window['AscFormat']["EXIT_BOUNCE"] = 26;
window['AscFormat'].EXIT_CREDITS = window['AscFormat']["EXIT_CREDITS"] = 28;
window['AscFormat'].EXIT_FLOAT = window['AscFormat']["EXIT_FLOAT"] = 30;
window['AscFormat'].EXIT_SHRINK_AND_TURN = window['AscFormat']["EXIT_SHRINK_AND_TURN"] = 31;
window['AscFormat'].EXIT_PINWHEEL = window['AscFormat']["EXIT_PINWHEEL"] = 35;
window['AscFormat'].EXIT_SINK_DOWN = window['AscFormat']["EXIT_SINK_DOWN"] = 37;
window['AscFormat'].EXIT_DROP = window['AscFormat']["EXIT_DROP"] = 38;
window['AscFormat'].EXIT_WHIP = window['AscFormat']["EXIT_WHIP"] = 41;
window['AscFormat'].EXIT_FLOAT_DOWN = window['AscFormat']["EXIT_FLOAT_DOWN"] = 42;
window['AscFormat'].EXIT_CENTER_REVOLVE = window['AscFormat']["EXIT_CENTER_REVOLVE"] = 43;
window['AscFormat'].EXIT_SWIVEL = window['AscFormat']["EXIT_SWIVEL"] = 45;
window['AscFormat'].EXIT_FLOAT_UP = window['AscFormat']["EXIT_FLOAT_UP"] = 47;
window['AscFormat'].EXIT_SPINNER = window['AscFormat']["EXIT_SPINNER"] = 49;
window['AscFormat'].EXIT_STRETCHY = window['AscFormat']["EXIT_STRETCHY"] = 50;
window['AscFormat'].EXIT_CURVE_DOWN = window['AscFormat']["EXIT_CURVE_DOWN"] = 52;
window['AscFormat'].EXIT_ZOOM = window['AscFormat']["EXIT_ZOOM"] = 53;
window['AscFormat'].EXIT_CONTRACT = window['AscFormat']["EXIT_CONTRACT"] = 55;
window['AscFormat'].EXIT_FLIP = window['AscFormat']["EXIT_FLIP"] = 56;

window['AscFormat'].ENTRANCE_APPEAR = window['AscFormat']["ENTRANCE_APPEAR"] = 1;
window['AscFormat'].ENTRANCE_FLY_IN_FROM = window['AscFormat']["ENTRANCE_FLY_IN_FROM"] = 2;
window['AscFormat'].ENTRANCE_BLINDS = window['AscFormat']["ENTRANCE_BLINDS"] = 3;
window['AscFormat'].ENTRANCE_BOX = window['AscFormat']["ENTRANCE_BOX"] = 4;
window['AscFormat'].ENTRANCE_CHECKERBOARD = window['AscFormat']["ENTRANCE_CHECKERBOARD"] = 5;
window['AscFormat'].ENTRANCE_CIRCLE = window['AscFormat']["ENTRANCE_CIRCLE"] = 6;
window['AscFormat'].ENTRANCE_DIAMOND = window['AscFormat']["ENTRANCE_DIAMOND"] = 8;
window['AscFormat'].ENTRANCE_DISSOLVE_IN = window['AscFormat']["ENTRANCE_DISSOLVE_IN"] = 9;
window['AscFormat'].ENTRANCE_FADE = window['AscFormat']["ENTRANCE_FADE"] = 10;
window['AscFormat'].ENTRANCE_PEEK_IN_FROM = window['AscFormat']["ENTRANCE_PEEK_IN_FROM"] = 12;
window['AscFormat'].ENTRANCE_PLUS = window['AscFormat']["ENTRANCE_PLUS"] = 13;
window['AscFormat'].ENTRANCE_RANDOM_BARS = window['AscFormat']["ENTRANCE_RANDOM_BARS"] = 14;
window['AscFormat'].ENTRANCE_SPIRAL_IN = window['AscFormat']["ENTRANCE_SPIRAL_IN"] = 15;
window['AscFormat'].ENTRANCE_SPLIT = window['AscFormat']["ENTRANCE_SPLIT"] = 16;
window['AscFormat'].ENTRANCE_STRETCH = window['AscFormat']["ENTRANCE_STRETCH"] = 17;
window['AscFormat'].ENTRANCE_STRIPS = window['AscFormat']["ENTRANCE_STRIPS"] = 18;
window['AscFormat'].ENTRANCE_BASIC_SWIVEL = window['AscFormat']["ENTRANCE_BASIC_SWIVEL"] = 19;
window['AscFormat'].ENTRANCE_WEDGE = window['AscFormat']["ENTRANCE_WEDGE"] = 20;
window['AscFormat'].ENTRANCE_WHEEL = window['AscFormat']["ENTRANCE_WHEEL"] = 21;
window['AscFormat'].ENTRANCE_WIPE_FROM = window['AscFormat']["ENTRANCE_WIPE_FROM"] = 22;
window['AscFormat'].ENTRANCE_BASIC_ZOOM = window['AscFormat']["ENTRANCE_BASIC_ZOOM"] = 23;
window['AscFormat'].ENTRANCE_BOOMERANG = window['AscFormat']["ENTRANCE_BOOMERANG"] = 25;
window['AscFormat'].ENTRANCE_BOUNCE = window['AscFormat']["ENTRANCE_BOUNCE"] = 26;
window['AscFormat'].ENTRANCE_CREDITS = window['AscFormat']["ENTRANCE_CREDITS"] = 28;
window['AscFormat'].ENTRANCE_FLOAT = window['AscFormat']["ENTRANCE_FLOAT"] = 30;
window['AscFormat'].ENTRANCE_GROW_AND_TURN = window['AscFormat']["ENTRANCE_GROW_AND_TURN"] = 31;
window['AscFormat'].ENTRANCE_PINWHEEL = window['AscFormat']["ENTRANCE_PINWHEEL"] = 35;
window['AscFormat'].ENTRANCE_RISE_UP = window['AscFormat']["ENTRANCE_RISE_UP"] = 37;
window['AscFormat'].ENTRANCE_DROP = window['AscFormat']["ENTRANCE_DROP"] = 38;
window['AscFormat'].ENTRANCE_WHIP = window['AscFormat']["ENTRANCE_WHIP"] = 41;
window['AscFormat'].ENTRANCE_FLOAT_UP = window['AscFormat']["ENTRANCE_FLOAT_UP"] = 42;
window['AscFormat'].ENTRANCE_CENTER_REVOLVE = window['AscFormat']["ENTRANCE_CENTER_REVOLVE"] = 43;
window['AscFormat'].ENTRANCE_SWIVEL = window['AscFormat']["ENTRANCE_SWIVEL"] = 45;
window['AscFormat'].ENTRANCE_FLOAT_DOWN = window['AscFormat']["ENTRANCE_FLOAT_DOWN"] = 47;
window['AscFormat'].ENTRANCE_SPINNER = window['AscFormat']["ENTRANCE_SPINNER"] = 49;
window['AscFormat'].ENTRANCE_CENTER_COMPRESS = window['AscFormat']["ENTRANCE_CENTER_COMPRESS"] = 50;
window['AscFormat'].ENTRANCE_CURVE_UP = window['AscFormat']["ENTRANCE_CURVE_UP"] = 52;
window['AscFormat'].ENTRANCE_ZOOM = window['AscFormat']["ENTRANCE_ZOOM"] = 53;
window['AscFormat'].ENTRANCE_EXPAND = window['AscFormat']["ENTRANCE_EXPAND"] = 55;
window['AscFormat'].ENTRANCE_FLIP = window['AscFormat']["ENTRANCE_FLIP"] = 56;

window['AscFormat'].EMPHASIS_FILL_COLOR = window['AscFormat']["EMPHASIS_FILL_COLOR"] = 1;
window['AscFormat'].EMPHASIS_FONT_COLOR = window['AscFormat']["EMPHASIS_FONT_COLOR"] = 3;
window['AscFormat'].EMPHASIS_GROW_SHRINK = window['AscFormat']["EMPHASIS_GROW_SHRINK"] = 6;
window['AscFormat'].EMPHASIS_LINE_COLOR = window['AscFormat']["EMPHASIS_LINE_COLOR"] = 7;
window['AscFormat'].EMPHASIS_SPIN = window['AscFormat']["EMPHASIS_SPIN"] = 8;
window['AscFormat'].EMPHASIS_TRANSPARENCY = window['AscFormat']["EMPHASIS_TRANSPARENCY"] = 9;
window['AscFormat'].EMPHASIS_BOLD_FLASH = window['AscFormat']["EMPHASIS_BOLD_FLASH"] = 10;
window['AscFormat'].EMPHASIS_BOLD_REVEAL = window['AscFormat']["EMPHASIS_BOLD_REVEAL"] = 15;
window['AscFormat'].EMPHASIS_BRUSH_COLOR = window['AscFormat']["EMPHASIS_BRUSH_COLOR"] = 16;
window['AscFormat'].EMPHASIS_UNDERLINE = window['AscFormat']["EMPHASIS_UNDERLINE"] = 18;
window['AscFormat'].EMPHASIS_OBJECT_COLOR = window['AscFormat']["EMPHASIS_OBJECT_COLOR"] = 19;
window['AscFormat'].EMPHASIS_COMPLEMENTARY_COLOR = window['AscFormat']["EMPHASIS_COMPLEMENTARY_COLOR"] = 21;
window['AscFormat'].EMPHASIS_COMPLEMENTARY_COLOR_2 = window['AscFormat']["EMPHASIS_COMPLEMENTARY_COLOR_2"] = 22;
window['AscFormat'].EMPHASIS_CONTRASTING_COLOR = window['AscFormat']["EMPHASIS_CONTRASTING_COLOR"] = 23;
window['AscFormat'].EMPHASIS_CONTRASTING_DARKEN = window['AscFormat']["EMPHASIS_CONTRASTING_DARKEN"] = 24;
window['AscFormat'].EMPHASIS_DESATURATE = window['AscFormat']["EMPHASIS_DESATURATE"] = 25;
window['AscFormat'].EMPHASIS_PULSE = window['AscFormat']["EMPHASIS_PULSE"] = 26;
window['AscFormat'].EMPHASIS_COLOR_PULSE = window['AscFormat']["EMPHASIS_COLOR_PULSE"] = 27;
window['AscFormat'].EMPHASIS_GROW_WITH_COLOR = window['AscFormat']["EMPHASIS_GROW_WITH_COLOR"] = 28;
window['AscFormat'].EMPHASIS_LIGHTEN = window['AscFormat']["EMPHASIS_LIGHTEN"] = 30;
window['AscFormat'].EMPHASIS_TEETER = window['AscFormat']["EMPHASIS_TEETER"] = 32;
window['AscFormat'].EMPHASIS_WAVE = window['AscFormat']["EMPHASIS_WAVE"] = 34;
window['AscFormat'].EMPHASIS_BLINK = window['AscFormat']["EMPHASIS_BLINK"] = 35;
window['AscFormat'].EMPHASIS_SHIMMER = window['AscFormat']["EMPHASIS_SHIMMER"] = 36;


//preset subtypes
window['AscFormat'].EXIT_ZOOM_OBJECT_CENTER = window['AscFormat']["EXIT_ZOOM_OBJECT_CENTER"] = 32;
window['AscFormat'].EXIT_ZOOM_SLIDE_CENTER = window['AscFormat']["EXIT_ZOOM_SLIDE_CENTER"] = 544;

window['AscFormat'].EXIT_WIPE_FROM_TOP = window['AscFormat']["EXIT_WIPE_FROM_TOP"] = 1;
window['AscFormat'].EXIT_WIPE_FROM_RIGHT = window['AscFormat']["EXIT_WIPE_FROM_RIGHT"] = 2;
window['AscFormat'].EXIT_WIPE_FROM_BOTTOM = window['AscFormat']["EXIT_WIPE_FROM_BOTTOM"] = 4;
window['AscFormat'].EXIT_WIPE_FROM_LEFT = window['AscFormat']["EXIT_WIPE_FROM_LEFT"] = 8;

window['AscFormat'].EXIT_WHEEL_1_SPOKE = window['AscFormat']["EXIT_WHEEL_1_SPOKE"] = 1;
window['AscFormat'].EXIT_WHEEL_2_SPOKES = window['AscFormat']["EXIT_WHEEL_2_SPOKES"] = 2;
window['AscFormat'].EXIT_WHEEL_3_SPOKES = window['AscFormat']["EXIT_WHEEL_3_SPOKES"] = 3;
window['AscFormat'].EXIT_WHEEL_4_SPOKES = window['AscFormat']["EXIT_WHEEL_4_SPOKES"] = 4;
window['AscFormat'].EXIT_WHEEL_8_SPOKES = window['AscFormat']["EXIT_WHEEL_8_SPOKES"] = 8;

window['AscFormat'].EXIT_STRIPS_RIGHT_UP = window['AscFormat']["EXIT_STRIPS_RIGHT_UP"] = 3;
window['AscFormat'].EXIT_STRIPS_RIGHT_DOWN = window['AscFormat']["EXIT_STRIPS_RIGHT_DOWN"] = 6;
window['AscFormat'].EXIT_STRIPS_LEFT_UP = window['AscFormat']["EXIT_STRIPS_LEFT_UP"] = 9;
window['AscFormat'].EXIT_STRIPS_LEFT_DOWN = window['AscFormat']["EXIT_STRIPS_LEFT_DOWN"] = 12;

window['AscFormat'].EXIT_SPLIT_VERTICAL_IN = window['AscFormat']["EXIT_SPLIT_VERTICAL_IN"] = 21;
window['AscFormat'].EXIT_SPLIT_HORIZONTAL_IN = window['AscFormat']["EXIT_SPLIT_HORIZONTAL_IN"] = 26;
window['AscFormat'].EXIT_SPLIT_VERTICAL_OUT = window['AscFormat']["EXIT_SPLIT_VERTICAL_OUT"] = 37;
window['AscFormat'].EXIT_SPLIT_HORIZONTAL_OUT = window['AscFormat']["EXIT_SPLIT_HORIZONTAL_OUT"] = 42;

window['AscFormat'].EXIT_RANDOM_BARS_VERTICAL = window['AscFormat']["EXIT_RANDOM_BARS_VERTICAL"] = 5;
window['AscFormat'].EXIT_RANDOM_BARS_HORIZONTAL = window['AscFormat']["EXIT_RANDOM_BARS_HORIZONTAL"] = 10;

window['AscFormat'].EXIT_PLUS_IN = window['AscFormat']["EXIT_PLUS_IN"] = 16;
window['AscFormat'].EXIT_PLUS_OUT = window['AscFormat']["EXIT_PLUS_OUT"] = 32;

window['AscFormat'].EXIT_PEEK_OUT_TO_TOP = window['AscFormat']["EXIT_PEEK_OUT_TO_TOP"] = 1;
window['AscFormat'].EXIT_PEEK_OUT_TO_RIGHT = window['AscFormat']["EXIT_PEEK_OUT_TO_RIGHT"] = 2;
window['AscFormat'].EXIT_PEEK_OUT_TO_BOTTOM = window['AscFormat']["EXIT_PEEK_OUT_TO_BOTTOM"] = 4;
window['AscFormat'].EXIT_PEEK_OUT_TO_LEFT = window['AscFormat']["EXIT_PEEK_OUT_TO_LEFT"] = 8;

window['AscFormat'].EXIT_FLY_OUT_TO_TOP = window['AscFormat']["EXIT_FLY_OUT_TO_TOP"] = 1;
window['AscFormat'].EXIT_FLY_OUT_TO_RIGHT = window['AscFormat']["EXIT_FLY_OUT_TO_RIGHT"] = 2;
window['AscFormat'].EXIT_FLY_OUT_TO_TOP_RIGHT = window['AscFormat']["EXIT_FLY_OUT_TO_TOP_RIGHT"] = 3;
window['AscFormat'].EXIT_FLY_OUT_TO_BOTTOM = window['AscFormat']["EXIT_FLY_OUT_TO_BOTTOM"] = 4;
window['AscFormat'].EXIT_FLY_OUT_TO_BOTTOM_RIGHT = window['AscFormat']["EXIT_FLY_OUT_TO_BOTTOM_RIGHT"] = 6;
window['AscFormat'].EXIT_FLY_OUT_TO_LEFT = window['AscFormat']["EXIT_FLY_OUT_TO_LEFT"] = 8;
window['AscFormat'].EXIT_FLY_OUT_TO_TOP_LEFT = window['AscFormat']["EXIT_FLY_OUT_TO_TOP_LEFT"] = 9;
window['AscFormat'].EXIT_FLY_OUT_TO_BOTTOM_LEFT = window['AscFormat']["EXIT_FLY_OUT_TO_BOTTOM_LEFT"] = 12;

window['AscFormat'].EXIT_DIAMOND_IN = window['AscFormat']["EXIT_DIAMOND_IN"] = 16;
window['AscFormat'].EXIT_DIAMOND_OUT = window['AscFormat']["EXIT_DIAMOND_OUT"] = 32;

window['AscFormat'].EXIT_COLLAPSE_TO_TOP = window['AscFormat']["EXIT_COLLAPSE_TO_TOP"] = 1;
window['AscFormat'].EXIT_COLLAPSE_TO_RIGHT = window['AscFormat']["EXIT_COLLAPSE_TO_RIGHT"] = 2;
window['AscFormat'].EXIT_COLLAPSE_TO_BOTTOM = window['AscFormat']["EXIT_COLLAPSE_TO_BOTTOM"] = 4;
window['AscFormat'].EXIT_COLLAPSE_TO_LEFT = window['AscFormat']["EXIT_COLLAPSE_TO_LEFT"] = 8;
window['AscFormat'].EXIT_COLLAPSE_ACROSS = window['AscFormat']["EXIT_COLLAPSE_ACROSS"] = 10;

window['AscFormat'].EXIT_CIRCLE_IN = window['AscFormat']["EXIT_CIRCLE_IN"] = 16;
window['AscFormat'].EXIT_CIRCLE_OUT = window['AscFormat']["EXIT_CIRCLE_OUT"] = 32;

window['AscFormat'].EXIT_CHECKERBOARD_UP = window['AscFormat']["EXIT_CHECKERBOARD_UP"] = 5;
window['AscFormat'].EXIT_CHECKERBOARD_ACROSS = window['AscFormat']["EXIT_CHECKERBOARD_ACROSS"] = 10;

window['AscFormat'].EXIT_BOX_IN = window['AscFormat']["EXIT_BOX_IN"] = 16;
window['AscFormat'].EXIT_BOX_OUT = window['AscFormat']["EXIT_BOX_OUT"] = 32;

window['AscFormat'].EXIT_BLINDS_VERTICAL = window['AscFormat']["EXIT_BLINDS_VERTICAL"] = 5;
window['AscFormat'].EXIT_BLINDS_HORIZONTAL = window['AscFormat']["EXIT_BLINDS_HORIZONTAL"] = 10;

window['AscFormat'].EXIT_BASIC_ZOOM_IN = window['AscFormat']["EXIT_BASIC_ZOOM_IN"] = 16;
window['AscFormat'].EXIT_BASIC_ZOOM_IN_TO_SCREEN_BOTTOM = window['AscFormat']["EXIT_BASIC_ZOOM_IN_TO_SCREEN_BOTTOM"] = 20;
window['AscFormat'].EXIT_BASIC_ZOOM_OUT = window['AscFormat']["EXIT_BASIC_ZOOM_OUT"] = 32;
window['AscFormat'].EXIT_BASIC_ZOOM_IN_SLIGHTLY = window['AscFormat']["EXIT_BASIC_ZOOM_IN_SLIGHTLY"] = 272;
window['AscFormat'].EXIT_BASIC_ZOOM_OUT_SLIGHTLY = window['AscFormat']["EXIT_BASIC_ZOOM_OUT_SLIGHTLY"] = 288;
window['AscFormat'].EXIT_BASIC_ZOOM_OUT_TO_SCREEN_CENTER = window['AscFormat']["EXIT_BASIC_ZOOM_OUT_TO_SCREEN_CENTER"] = 544;

window['AscFormat'].EXIT_BASIC_SWIVEL_VERTICAL = window['AscFormat']["EXIT_BASIC_SWIVEL_VERTICAL"] = 5;
window['AscFormat'].EXIT_BASIC_SWIVEL_HORIZONTAL = window['AscFormat']["EXIT_BASIC_SWIVEL_HORIZONTAL"] = 10;

window['AscFormat'].ENTRANCE_ZOOM_OBJECT_CENTER = window['AscFormat']["ENTRANCE_ZOOM_OBJECT_CENTER"] = 16;
window['AscFormat'].ENTRANCE_ZOOM_SLIDE_CENTER = window['AscFormat']["ENTRANCE_ZOOM_SLIDE_CENTER"] = 528;

window['AscFormat'].ENTRANCE_WIPE_FROM_TOP = window['AscFormat']["ENTRANCE_WIPE_FROM_TOP"] = 1;
window['AscFormat'].ENTRANCE_WIPE_FROM_RIGHT = window['AscFormat']["ENTRANCE_WIPE_FROM_RIGHT"] = 2;
window['AscFormat'].ENTRANCE_WIPE_FROM_BOTTOM = window['AscFormat']["ENTRANCE_WIPE_FROM_BOTTOM"] = 4;
window['AscFormat'].ENTRANCE_WIPE_FROM_LEFT = window['AscFormat']["ENTRANCE_WIPE_FROM_LEFT"] = 8;

window['AscFormat'].ENTRANCE_WHEEL_1_SPOKE = window['AscFormat']["ENTRANCE_WHEEL_1_SPOKE"] = 1;
window['AscFormat'].ENTRANCE_WHEEL_2_SPOKES = window['AscFormat']["ENTRANCE_WHEEL_2_SPOKES"] = 2;
window['AscFormat'].ENTRANCE_WHEEL_3_SPOKES = window['AscFormat']["ENTRANCE_WHEEL_3_SPOKES"] = 3;
window['AscFormat'].ENTRANCE_WHEEL_4_SPOKES = window['AscFormat']["ENTRANCE_WHEEL_4_SPOKES"] = 4;
window['AscFormat'].ENTRANCE_WHEEL_8_SPOKES = window['AscFormat']["ENTRANCE_WHEEL_8_SPOKES"] = 8;

window['AscFormat'].ENTRANCE_STRIPS_RIGHT_UP = window['AscFormat']["ENTRANCE_STRIPS_RIGHT_UP"] = 3;
window['AscFormat'].ENTRANCE_STRIPS_RIGHT_DOWN = window['AscFormat']["ENTRANCE_STRIPS_RIGHT_DOWN"] = 6;
window['AscFormat'].ENTRANCE_STRIPS_LEFT_UP = window['AscFormat']["ENTRANCE_STRIPS_LEFT_UP"] = 9;
window['AscFormat'].ENTRANCE_STRIPS_LEFT_DOWN = window['AscFormat']["ENTRANCE_STRIPS_LEFT_DOWN"] = 12;

window['AscFormat'].ENTRANCE_STRETCH_FROM_TOP = window['AscFormat']["ENTRANCE_STRETCH_FROM_TOP"] = 1;
window['AscFormat'].ENTRANCE_STRETCH_FROM_RIGHT = window['AscFormat']["ENTRANCE_STRETCH_FROM_RIGHT"] = 2;
window['AscFormat'].ENTRANCE_STRETCH_FROM_BOTTOM = window['AscFormat']["ENTRANCE_STRETCH_FROM_BOTTOM"] = 4;
window['AscFormat'].ENTRANCE_STRETCH_FROM_LEFT = window['AscFormat']["ENTRANCE_STRETCH_FROM_LEFT"] = 8;
window['AscFormat'].ENTRANCE_STRETCH_ACROSS = window['AscFormat']["ENTRANCE_STRETCH_ACROSS"] = 10;

window['AscFormat'].ENTRANCE_SPLIT_VERTICAL_IN = window['AscFormat']["ENTRANCE_SPLIT_VERTICAL_IN"] = 21;
window['AscFormat'].ENTRANCE_SPLIT_HORIZONTAL_IN = window['AscFormat']["ENTRANCE_SPLIT_HORIZONTAL_IN"] = 26;
window['AscFormat'].ENTRANCE_SPLIT_VERTICAL_OUT = window['AscFormat']["ENTRANCE_SPLIT_VERTICAL_OUT"] = 37;
window['AscFormat'].ENTRANCE_SPLIT_HORIZONTAL_OUT = window['AscFormat']["ENTRANCE_SPLIT_HORIZONTAL_OUT"] = 42;

window['AscFormat'].ENTRANCE_RANDOM_BARS_VERTICAL = window['AscFormat']["ENTRANCE_RANDOM_BARS_VERTICAL"] = 5;
window['AscFormat'].ENTRANCE_RANDOM_BARS_HORIZONTAL = window['AscFormat']["ENTRANCE_RANDOM_BARS_HORIZONTAL"] = 10;

window['AscFormat'].ENTRANCE_PLUS_IN = window['AscFormat']["ENTRANCE_PLUS_IN"] = 16;
window['AscFormat'].ENTRANCE_PLUS_OUT = window['AscFormat']["ENTRANCE_PLUS_OUT"] = 32;

window['AscFormat'].ENTRANCE_PEEK_IN_FROM_TOP = window['AscFormat']["ENTRANCE_PEEK_IN_FROM_TOP"] = 1;
window['AscFormat'].ENTRANCE_PEEK_IN_FROM_RIGHT = window['AscFormat']["ENTRANCE_PEEK_IN_FROM_RIGHT"] = 2;
window['AscFormat'].ENTRANCE_PEEK_IN_FROM_BOTTOM = window['AscFormat']["ENTRANCE_PEEK_IN_FROM_BOTTOM"] = 4;
window['AscFormat'].ENTRANCE_PEEK_IN_FROM_LEFT = window['AscFormat']["ENTRANCE_PEEK_IN_FROM_LEFT"] = 8;

window['AscFormat'].ENTRANCE_FLY_IN_FROM_TOP = window['AscFormat']["ENTRANCE_FLY_IN_FROM_TOP"] = 1;
window['AscFormat'].ENTRANCE_FLY_IN_FROM_RIGHT = window['AscFormat']["ENTRANCE_FLY_IN_FROM_RIGHT"] = 2;
window['AscFormat'].ENTRANCE_FLY_IN_FROM_TOP_RIGHT = window['AscFormat']["ENTRANCE_FLY_IN_FROM_TOP_RIGHT"] = 3;
window['AscFormat'].ENTRANCE_FLY_IN_FROM_BOTTOM = window['AscFormat']["ENTRANCE_FLY_IN_FROM_BOTTOM"] = 4;
window['AscFormat'].ENTRANCE_FLY_IN_FROM_BOTTOM_RIGHT = window['AscFormat']["ENTRANCE_FLY_IN_FROM_BOTTOM_RIGHT"] = 6;
window['AscFormat'].ENTRANCE_FLY_IN_FROM_LEFT = window['AscFormat']["ENTRANCE_FLY_IN_FROM_LEFT"] = 8;
window['AscFormat'].ENTRANCE_FLY_IN_FROM_TOP_LEFT = window['AscFormat']["ENTRANCE_FLY_IN_FROM_TOP_LEFT"] = 9;
window['AscFormat'].ENTRANCE_FLY_IN_FROM_BOTTOM_LEFT = window['AscFormat']["ENTRANCE_FLY_IN_FROM_BOTTOM_LEFT"] = 12;

window['AscFormat'].ENTRANCE_DIAMOND_IN = window['AscFormat']["ENTRANCE_DIAMOND_IN"] = 16;
window['AscFormat'].ENTRANCE_DIAMOND_OUT = window['AscFormat']["ENTRANCE_DIAMOND_OUT"] = 32;

window['AscFormat'].ENTRANCE_CIRCLE_IN = window['AscFormat']["ENTRANCE_CIRCLE_IN"] = 16;
window['AscFormat'].ENTRANCE_CIRCLE_OUT = window['AscFormat']["ENTRANCE_CIRCLE_OUT"] = 32;

window['AscFormat'].ENTRANCE_CHECKERBOARD_DOWN = window['AscFormat']["ENTRANCE_CHECKERBOARD_DOWN"] = 5;
window['AscFormat'].ENTRANCE_CHECKERBOARD_ACROSS = window['AscFormat']["ENTRANCE_CHECKERBOARD_ACROSS"] = 10;

window['AscFormat'].ENTRANCE_BOX_IN = window['AscFormat']["ENTRANCE_BOX_IN"] = 16;
window['AscFormat'].ENTRANCE_BOX_OUT = window['AscFormat']["ENTRANCE_BOX_OUT"] = 32;

window['AscFormat'].ENTRANCE_BLINDS_VERTICAL = window['AscFormat']["ENTRANCE_BLINDS_VERTICAL"] = 5;
window['AscFormat'].ENTRANCE_BLINDS_HORIZONTAL = window['AscFormat']["ENTRANCE_BLINDS_HORIZONTAL"] = 10;

window['AscFormat'].ENTRANCE_BASIC_ZOOM_IN = window['AscFormat']["ENTRANCE_BASIC_ZOOM_IN"] = 16;
window['AscFormat'].ENTRANCE_BASIC_ZOOM_OUT = window['AscFormat']["ENTRANCE_BASIC_ZOOM_OUT"] = 32;
window['AscFormat'].ENTRANCE_BASIC_ZOOM_OUT_FROM_SCREEN_BOTTOM = window['AscFormat']["ENTRANCE_BASIC_ZOOM_OUT_FROM_SCREEN_BOTTOM"] = 36;
window['AscFormat'].ENTRANCE_BASIC_ZOOM_IN_SLIGHTLY = window['AscFormat']["ENTRANCE_BASIC_ZOOM_IN_SLIGHTLY"] = 272;
window['AscFormat'].ENTRANCE_BASIC_ZOOM_OUT_SLIGHTLY = window['AscFormat']["ENTRANCE_BASIC_ZOOM_OUT_SLIGHTLY"] = 288;
window['AscFormat'].ENTRANCE_BASIC_ZOOM_IN_FROM_SCREEN_CENTER = window['AscFormat']["ENTRANCE_BASIC_ZOOM_IN_FROM_SCREEN_CENTER"] = 528;

window['AscFormat'].ENTRANCE_BASIC_SWIVEL_VERTICAL = window['AscFormat']["ENTRANCE_BASIC_SWIVEL_VERTICAL"] = 5;
window['AscFormat'].ENTRANCE_BASIC_SWIVEL_HORIZONTAL = window['AscFormat']["ENTRANCE_BASIC_SWIVEL_HORIZONTAL"] = 10;

//animation node types
window['AscFormat'].NODE_TYPE_AFTEREFFECT = window['AscFormat']["NODE_TYPE_AFTEREFFECT"] = 0;
window['AscFormat'].NODE_TYPE_AFTERGROUP = window['AscFormat']["NODE_TYPE_AFTERGROUP"] = 1;
window['AscFormat'].NODE_TYPE_CLICKEFFECT = window['AscFormat']["NODE_TYPE_CLICKEFFECT"] = 2;
window['AscFormat'].NODE_TYPE_CLICKPAR = window['AscFormat']["NODE_TYPE_CLICKPAR"] = 3;
window['AscFormat'].NODE_TYPE_INTERACTIVESEQ = window['AscFormat']["NODE_TYPE_INTERACTIVESEQ"] = 4;
window['AscFormat'].NODE_TYPE_MAINSEQ = window['AscFormat']["NODE_TYPE_MAINSEQ"] = 5;
window['AscFormat'].NODE_TYPE_TMROOT = window['AscFormat']["NODE_TYPE_TMROOT"] = 6;
window['AscFormat'].NODE_TYPE_WITHEFFECT = window['AscFormat']["NODE_TYPE_WITHEFFECT"] = 7;
window['AscFormat'].NODE_TYPE_WITHGROUP = window['AscFormat']["NODE_TYPE_WITHGROUP"] = 8;

//special values for repeat count
window['AscFormat'].untilNextClick = window['AscFormat']["untilNextClick"] = -1;
window['AscFormat'].untilNextSlide = window['AscFormat']["untilNextSlide"] = -2;
