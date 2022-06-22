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
var c_oAscWrapStyle2 = {
	Inline       : 0,
	Square       : 1,
	Tight        : 2,
	Through      : 3,
	TopAndBottom : 4,
	Behind       : 5,
	InFront      : 6
};

/** @enum {number} */
var c_oAscTableSelectionType = {
	Cell   : 0,
	Row    : 1,
	Column : 2,
	Table  : 3
};

/** @enum {number} */
var c_oAscContextMenuTypes = {
	Common       : 0, // Обычное контекстное меню
	ChangeHdrFtr : 1  // Специальное контестное меню для попадания в колонтитул
};

/** @enum {number} */
var c_oAscMouseMoveLockedObjectType = {
	Common : 0,
	Header : 1,
	Footer : 2
};

/** @enum {number} */
var c_oAscCollaborativeMarksShowType = {
	None        : -1,
	All         : 0,
	LastChanges : 1
};

/** @enum {number} */
var c_oAscChangeLevel = {
	BringToFront  : 0x00,
	BringForward  : 0x01,
	SendToBack    : 0x02,
	BringBackward : 0x03
};



/** @enum {number} */
var c_oAscVertAlignJc = {
	Top    : 0x00, // var vertalignjc_Top    = 0x00;
	Center : 0x01, // var vertalignjc_Center = 0x01;
	Bottom : 0x02  // var vertalignjc_Bottom = 0x02
};

/** @enum {number} */
var c_oAscTableLayout = {
	AutoFit : 0x00,
	Fixed   : 0x01
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

var TABLE_STYLE_WIDTH_PIX  = 72;
var TABLE_STYLE_HEIGHT_PIX = 52;




/** @enum {number} */
var c_oAscRevisionsObjectType = {
	Image        : 0,
	Shape        : 1,
	Chart        : 2,
	MathEquation : 3
};

var c_oSerFormat = {
	Version   : 5, //1.0.0.5
	Signature : "DOCY"
};

var c_oAscFootnotePos = {
	BeneathText : 0x00,
	DocEnd      : 0x01,
	PageBottom  : 0x02,
	SectEnd     : 0x03
};

var c_oAscEndnotePos = {
	DocEnd      : 0x00,
	SectEnd     : 0x01
};

var c_oAscFootnoteRestart = {
	Continuous : 0x00, //section_footnote_RestartContinuous,
	EachSect   : 0x01, //section_footnote_RestartEachSect,
	EachPage   : 0x02  //section_footnote_RestartEachPage
};


var c_oAscSdtLevelType = {
	Block  : 0x01,
	Inline : 0x02,
	Row    : 0x03,
	Cell   : 0x04
};

var c_oAscTOCStylesType = {
	Current  : 0,
	Simple   : 1,
	Standard : 2,
	Modern   : 3,
	Classic  : 4,
	Web      : 5
};

var c_oAscTOFStylesType = {
	Current     : 0,
	Classic     : 1,
	Distinctive : 2,
	Centered    : 3,
	Formal      : 4,
	Simple      : 5,
	Web         : 6
};

var c_oAscStyleType = {
	Paragraph : 1,
	Numbering : 2,
	Table     : 3,
	Character : 4
};

var c_oAscHyperlinkAnchor = {
	Heading       : 1,
	Bookmark      : 2
};


var c_oAscSdtCheckBoxDefaults = {
	CheckedSymbol   : 0x2612,
	UncheckedSymbol : 0x2610,
	CheckedFont     : "MS Gothic",
	UncheckedFont   : "MS Gothic"
};

window["flat_desine"] = false;

var c_oAscDocumentShortcutType = {
	InsertPageBreak      : 1,
	InsertLineBreak      : 2,
	InsertColumnBreak    : 3,
	ResetChar            : 4,
	NonBreakingSpace     : 5,
	ApplyHeading1        : 6,
	ApplyHeading2        : 7,
	ApplyHeading3        : 8,
	Strikeout            : 9,
	ShowAll              : 10, // Показать/скрыть непечатаемые символы
	EditSelectAll        : 11,
	Bold                 : 12,
	CopyFormat           : 13,
	CopyrightSign        : 14,
	InsertEndnoteNow     : 15,
	CenterPara           : 16,
	EuroSign             : 17,
	InsertFootnoteNow    : 18,
	Italic               : 19,
	JustifyPara          : 20,
	InsertHyperlink      : 21,
	ApplyListBullet      : 22,
	LeftPara             : 23,
	Indent               : 24,
	UnIndent             : 25,
	PrintPreviewAndPrint : 26,
	InsertPageNumber     : 27,
	RegisteredSign       : 28,
	RightPara            : 29,
	Save                 : 30,
	TrademarkSign        : 31,
	Underline            : 32,
	PasteFormat          : 33,
	EditUndo             : 34,
	EditRedo             : 35,
	EmDash               : 36,
	EnDash               : 37,
	UpdateFields         : 38,
	InsertEquation       : 39,
	Superscript          : 40,
	NonBreakingHyphen    : 41,
	SoftHyphen           : 42,
	HorizontalEllipsis   : 43,
	Subscript            : 44,
	IncreaseFontSize     : 45,
	DecreaseFontSize     : 46
};

var c_oAscDocumentRefenceToType = {
	Text               : 0,
	PageNum            : 1,
	ParaNum            : 2,
	ParaNumNoContext   : 3,
	ParaNumFullContex  : 4,
	AboveBelow         : 5,
	OnlyLabelAndNumber : 6,
	OnlyCaptionText    : 7,
	NoteNumber         : 8,
	NoteNumberFormatted : 9
};

var c_oAscTextToTableSeparator = {
	Paragraph : 1,
	Tab       : 2,
	Symbol    : 3
};

var c_oAscTextToTableAutoFitType = {
	Fixed   : 1,
	Content : 2,
	Window  : 3
};

var c_oAscDateTimeFormat = {};
c_oAscDateTimeFormat[lcid_azLatnAZ] = [
	"dd.mm.yyyy",
	"d Month yyyy, dddd",
	"d Mmmm yyyy",
	"dd.mm.yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"dd/mm/yyyy",
	"d mmm. yy",
	"dd/mm/yy",
	"Mmmm yy",
	"mmm-yy",
	"dd.mm.yyyy hh:mm",
	"dd.mm.yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_bgBG]     = [
	"d.m.yyyy",
	"dddd, dd mmmm yyyy г.",
	"dd mmmm yyyy г.",
	"d.m.yy г.",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"d/m/yyyy",
	"dd mmm. yy г.",
	"d/m/yy",
	"mmmm yy г",
	"mmm-yy",
	"d.m.yyyy г. hh:mm",
	"d.m.yyyy г. hh:mm:ss",
	"hh:mm",
	"hh:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_csCZ]     = [
	"dd.mm.yyyy",
	"dddd d. mmmm yyyy",
	"d. mmmm yyyy",
	"dd.mm.yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"dd/mm/yyyy",
	"d. mmm. yy",
	"dd/mm/yy",
	"mmmm ’yy",
	"mmm-yy",
	"dd.mm.yyyy hh:mm",
	"dd.mm.yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_deAT]     = c_oAscDateTimeFormat[lcid_deDE] = [
	"dd.mm.yyyyy",
	"Dddd, d. Mmmm yyyy",
	"d. Mmmm yyyy",
	"dd.mm.yy",
	"yyyy-mm-dd",
	"yy-mm-dd",
	"dd/mm/yyyy",
	"dd. Mmm. yyyy",
	"dd/mm/yy",
	"Mmmm yy",
	"Mmm-yy",
	"dd.mm.yyyy hh:mm",
	"dd.mm.yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_deCH]     = [
	"dd.mm.yyyyy",
	"Dddd, d. Mmmm yyyy",
	"d. Mmmm yyyy",
	"dd.mm.yy",
	"yyyy-mm-dd",
	"yy-mm-dd",
	"dd/mm/yyyy",
	"dd. Mmm. yyyy",
	"dd/mm/yy",
	"Mmmm yy",
	"Mmm-yy",
	"dd.mm.yyyy hh:mm",
	"dd.mm.yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_elGR]     = [
	"d/m/yyyy",
	"Dddd, d Mmmm yyyy",
	"d Mmmm yyyy",
	"d/m/yy",
	"yyyy-mm-dd",
	"d-Mmm-yy",
	"d.m.yyyy",
	"d Mmm. yy",
	"d.m.yy",
	"Mmmm yy",
	"Mmm-yy",
	"d/m/yyyy H:mm am/pm",
	"d/m/yyyy H:mm:ss am/pm",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_enAU]     = [
	"d/mm/yyyy",
	"Dddd, d Mmmm yyyy",
	"d Mmmm yyyy",
	"d/mm/yy",
	"yyyy-mm-dd",
	"d-Mmm-yyyy",
	"d.mm.yyyy",
	"d Mmm. yy",
	"Mmmm yy",
	"Mmm-yy",
	"d/mm/yyyy H:mm am/pm",
	"d/mm/yyyy H:mm:ss am/pm",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_enGB]     = [
	"dd/mm/yyyy",
	"Dddd, dd Mmmm yyyy",
	"dd Mmmm yyyy",
	"dd/mm/yy",
	"yyyy-mm-dd",
	"d-Mmm-yy",
	"dd.mm.yyyy",
	"dd Mmm. yy",
	"d Mmmm yyyy",
	"Mmmm yy",
	"Mmm-yy",
	"dd/mm/yyyy hh:mm",
	"dd/mm/yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_esES]     = [
	"dd/mm/yyyy",
	"dddd, d de mmmm de yyyy",
	"d de mmmm de yyyy",
	"dd/mm/yy",
	"yyyy-mm-dd",
	"d-mmm.-yy",
	"dd.mm.yyyy",
	"d mmm. yy",
	"dd.mm.yy",
	"mmmm de yy",
	"mmm.-yy",
	"dd/mm/yyyy hh:mm",
	"dd/mm/yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_esMX]     = [
	"dd/mm/yyyy",
	"dddd, d de mmmm de yyyy",
	"d de mmmm de yyyy",
	"dd/mm/yy",
	"yyyy-mm-dd",
	"d-mmm.-yy",
	"dd.mm.yyyy",
	"d mmm. yy",
	"dd.mm.yy",
	"mmmm de yyyy",
	"mmm.-yy",
	"dd/mm/yyyy HH:mm am/pm",
	"dd/mm/yyyy HH:mm:ss am/pm",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_fiFI]     = [
	"d.m.yyyy",
	"dddd d. mmmm yyyy",
	"d. mmmm yyyy",
	"d.m.yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"d/m/yyyy",
	"d. mmm. yy",
	"d/m/yy",
	"mmmm yy",
	"mmm-yy",
	"d.m.yyyy hh.mm",
	"d.m.yyyy hh.mm.ss",
	"H.mm am/pm",
	"H.mm.ss am/pm",
	"hh.mm",
	"hh.mm.ss"
];
c_oAscDateTimeFormat[lcid_frFR]     = [
	"dd/mm/yyyy",
	"dddd d Mmmm yyyy",
	"d mmmm yyyy",
	"dd/mm/yy",
	"yyyy-mm-dd",
	"d-mmm.-yy",
	"dd.mm.yyyy",
	"d mmm. yy",
	"dd.mm.yy",
	"mmmm yy",
	"mmm.-yy",
	"dd/mm/yyyy hh:mm",
	"dd/mm/yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss",
];
c_oAscDateTimeFormat[lcid_itIT]     = [
	"dd/mm/yyyy",
	"dddd d mmmm yyyy",
	"d mmmm yyyy",
	"dd/mm/yy",
	"yyyy-mm-dd",
	"d-mm.-yy",
	"dd.mm.yyyy",
	"d mmm. yy",
	"mmm. ’yy",
	"mmmm ’yy",
	"mmm-yy",
	"dd/mm/yyyy hh:mm",
	"dd/mm/yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_jaJP]     = [
	"yyyy/mm/dd",
	"yyyy年m月d日",
	"yyyy年m月d日(JDDD)",
	"yyyy年m月",
	"JYYYY年JM月JD日(JDDD)",
	"yyyy/m/d",
	"yy/m/d h時m分",
	"yy/m/d h時m分s秒",
	"am/pmH時M分",
	"am/pmH時M分s秒",
	"h時M分",
	"h時M分s秒",
	"yyyy-mm-dd"
];
c_oAscDateTimeFormat[lcid_koKR]     = [
	"yyyy-mm-dd",
	"yyyy년 m월 d일 dddd",
	"yyyy년 m월 d일",
	"yyyy/m/d",
	"yymmdd",
	"yyyy년 m월",
	"yyyy년 m월 d일 am/pm h시 M분",
	"y년 m월 d일 h시 M분 s초",
	"am/pm h시 M분",
	"am/pm h시 M분 s초",
	"h시 M분",
	"h시 M분 s초"
];
c_oAscDateTimeFormat[lcid_lvLV]     = [
	"dd.mm.yyyy",
	"dddd, yyyy. gada d. mmmm",
	"yyyy. gada d. mmmm",
	"dd.mm.yy",
	"yyyy-mm-dd",
	"d-mmm.-yy",
	"dd/mm/yyyy",
	"yy. gada d. mmm.",
	"dd/mm/yy",
	"yy mmmm",
	"mmm.-yy",
	"dd.mm.yyyy hh:mm",
	"dd.mm.yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_huHU]     = [
	"yyyy. mm. dd.",
	"yyyy. mmmm d., dddd",
	"yyyy. mmmm d.",
	"yy. mm. dd.",
	"yyyy-mm-dd",
	"yy-mmm.-d",
	"yyyy/mm/dd",
	"yy. mmm. d.",
	"’yy mmm.",
	"’yy mmmm",
	"mmm.-yy",
	"yyyy. mm. dd. hh:mm",
	"yyyy. mm. dd. hh:mm:ss",
	"am/pm H:mm",
	"am/pm H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_nlNL]     = [
	"d-m-yyyy",
	"dddd d mmmm yyyy",
	"d mmmm yyyy",
	"d-m-yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"d/m/yyyy",
	"d mmm. yy",
	"d/m/yy",
	"mmmm ’yy",
	"mmm-yy",
	"d-m-yyyy hh:mm",
	"d-m-yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_plPL]     = [
	"dd.mm.yyyy",
	"dddd, d mmmm yyyy",
	"d mmmm yyyy",
	"dd.mm.yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"dd/mm/yyyy",
	"dd mmm yy",
	"dd/mm/yy",
	"mmmm yy",
	"mmm-yy",
	"dd.mm.yyyy hh:mm",
	"dd.mm.yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_ptBR]     = [
	"dd/mm/yyyy",
	"dddd, d de mmmm de yyyy",
	"d de mmmm de yyyy",
	"dd/mm/yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"dd.mm.yyyy",
	"d mmm. yy",
	"dd.mm.yy",
	"d mmmm de yy",
	"mmm-yy",
	"dd/mm/yyyy hh:mm",
	"dd/mm/yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_ptPT]     = [
	"dd/mm/yyyy",
	"dddd, d de mmmm de yyyy",
	"d de mmmm de yyyy",
	"dd/mm/yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"dd.mm.yyyy",
	"d mmm. yy",
	"dd.mm.yy",
	"mmmm de yy",
	"mmm-yy",
	"dd/mm/yyyy hh:mm",
	"dd/mm/yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_ruRU]     = [
	"dd.mm.yyyy",
	"dddd, d mmmm yyyy г.",
	"d mmmm yyyy г.",
	"dd.mm.yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"dd/mm/yyyy",
	"d mmm. yy г.",
	"dd/mm/yy",
	"mmmm yy",
	"mmm-yy",
	"dd.mm.yyyy hh:mm",
	"dd.mm.yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_skSK]     = [
	"d. m. yyyy",
	"dddd d. mmmm yyyy",
	"d. mmmm yyyy",
	"d. m. yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"d/m/yyyy",
	"d. mmm. yy",
	"d/m/yy",
	"mmmm yy",
	"mmm-yy",
	"d. m. yyyy hh:mm",
	"d. m. yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_slSI]     = [
	"d. mm. yyyy",
	"dddd, dd. mmmm yyyy",
	"dd. mmmm yyyy",
	"d. mm. yy",
	"yyyy-mm-dd",
	"d-mmm.-yy",
	"d/mm/yyyy",
	"dd. mmm. yy",
	"d/mm/yy",
	"mmmm yy",
	"mmm.-yy",
	"d. mm. yyyy hh:mm",
	"d. mm. yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_svFI]     = [
	"d. mm. yyyy",
	"dddd, dd. mmmm yyyy",
	"dd. mmmm yyyy",
	"d. mm. yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"d/mm/yyyy",
	"dd. mmm. yy",
	"d/mm/yy",
	"mmmm yy",
	"mmm.-yy",
	"d. mm. yyyy hh:mm",
	"d. mm. yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_svSE]     = [
	"yyyy-mm-dd",
	"dddd 'den' d mmmm yyyy",
	"d mmmm yyyy",
	"dd-mm-yy",
	"yymmdd",
	"d mmm yy",
	"d/m yyyy",
	"d mmm -yy",
	"d/m/yy",
	"mmmm yyyy",
	"'den' d mmmm yyyy",
	"yyyy-mm-dd hh:mm",
	"yy-mm-dd HH:mm",
	"H.mm",
	"H.mm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_trTR]     = [
	"d.mm.yyyy",
	"d Mmmm yyyy Dddd",
	"d Mmmm yyyy",
	"d.mm.yy",
	"yyyy-mm-dd",
	"d-Mmm-yy",
	"d/mm/yyyy",
	"d Mmm. yy",
	"d/mm/yy",
	"Mmmm yy",
	"Mmm-yy",
	"d.mm.yyyy hh:mm",
	"d.mm.yyyy hh:mm:ss",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_ukUA]     = [
	"dd.mm.yyyy",
	"dddd, d mmmm yyyy р.",
	"d mmmm yyyy р.",
	"dd.mm.yy",
	"yyyy-mm-dd",
	"d-mmm-yy",
	"dd/mm/yyyy",
	"d mmm. yy р.",
	"dd/mm/yy",
	"mmmm yy р",
	"mmm-yy",
	"dd.mm.yyyy hh:mm",
	"dd.mm.yyyy hh:mm:ss",
	"H:mm",
	"H:mm:ss",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_viVN]     = [
	"dd/mm/yyyy",
	"Dddd, dd Mmmm yyyy",
	"dd Mmmm yyyy",
	"dd/mm/yy",
	"yyyy-mm-dd",
	"d-Mmm-yy",
	"d-Mmmm-yy",
	"dd Mmm. yy",
	"dd.mm.yy",
	"Mmmm yy",
	"Mmm-yy",
	"dd/mm/yyyy H:mm am/pm",
	"dd/mm/yyyy H:mm:ss am/pm",
	"H:mm am/pm",
	"H:mm:ss am/pm",
	"hh:mm",
	"hh:mm:ss"
];
c_oAscDateTimeFormat[lcid_zhCN]     = [
	"yyyy/m/d",
	"yyyy年m月d日",
	"yyyy年m月d日dddd",
	"yy.m.d",
	"yyyy年m月",
	"h时m分ss秒",
	"h时m分",
	"am/pmh时m分",
	"JYYYY年mmmmd日",
	"JYYYY年JM月JD日dddd",
	"JYYYY年JM月JD"
];
c_oAscDateTimeFormat[lcid_enUS]     = [
	"M/d/yyyy",
	"dddd, MMMM d, yyyy",
	"MMMM d, yyyy",
	"M/d/yy",
	"yyyy-MM-dd",
	"d-MMM-yy",
	"M.d.yyyy",
	"MMM. d, yy",
	"d MMMM yyyy",
	"MMMM yy",
	"MMM-yy",
	"M/d/yyyy h:mm am/pm",
	"M/d/yyyy h:mm:ss am/pm",
	"h:mm am/pm",
	"h:mm:ss am/pm",
	"HH:mm",
	"HH:mm:ss"
];

//------------------------------------------------------------export---------------------------------------------------
var prot;
window['Asc'] = window['Asc'] || {};
window['AscCommonWord'] = window['AscCommonWord'] || {};

prot          = window['Asc']['c_oAscWrapStyle2'] = c_oAscWrapStyle2;
prot['Inline']       = c_oAscWrapStyle2.Inline;
prot['Square']       = c_oAscWrapStyle2.Square;
prot['Tight']        = c_oAscWrapStyle2.Tight;
prot['Through']      = c_oAscWrapStyle2.Through;
prot['TopAndBottom'] = c_oAscWrapStyle2.TopAndBottom;
prot['Behind']       = c_oAscWrapStyle2.Behind;
prot['InFront']      = c_oAscWrapStyle2.InFront;

prot = window['Asc']['c_oAscContextMenuTypes'] = window['Asc'].c_oAscContextMenuTypes = c_oAscContextMenuTypes;
prot['Common']       = c_oAscContextMenuTypes.Common;
prot['ChangeHdrFtr'] = c_oAscContextMenuTypes.ChangeHdrFtr;

prot = window['Asc']['c_oAscCollaborativeMarksShowType'] = c_oAscCollaborativeMarksShowType;
prot['None']        = c_oAscCollaborativeMarksShowType.None;
prot['All']         = c_oAscCollaborativeMarksShowType.All;
prot['LastChanges'] = c_oAscCollaborativeMarksShowType.LastChanges;


prot = window['Asc']['c_oAscChangeLevel'] = c_oAscChangeLevel;
prot['BringToFront']  = c_oAscChangeLevel.BringToFront;
prot['BringForward']  = c_oAscChangeLevel.BringForward;
prot['SendToBack']    = c_oAscChangeLevel.SendToBack;
prot['BringBackward'] = c_oAscChangeLevel.BringBackward;

prot = window['Asc']['c_oAscVertAlignJc'] = c_oAscVertAlignJc;
prot['Top']    = c_oAscVertAlignJc.Top;
prot['Center'] = c_oAscVertAlignJc.Center;
prot['Bottom'] = c_oAscVertAlignJc.Bottom;

prot = window['Asc']['c_oAscTableLayout'] = c_oAscTableLayout;
prot['AutoFit'] = c_oAscTableLayout.AutoFit;
prot['Fixed']   = c_oAscTableLayout.Fixed;

prot = window['Asc']['c_oAscAlignShapeType'] = c_oAscAlignShapeType;
prot['ALIGN_LEFT']   = c_oAscAlignShapeType.ALIGN_LEFT;
prot['ALIGN_RIGHT']  = c_oAscAlignShapeType.ALIGN_RIGHT;
prot['ALIGN_TOP']    = c_oAscAlignShapeType.ALIGN_TOP;
prot['ALIGN_BOTTOM'] = c_oAscAlignShapeType.ALIGN_BOTTOM;
prot['ALIGN_CENTER'] = c_oAscAlignShapeType.ALIGN_CENTER;
prot['ALIGN_MIDDLE'] = c_oAscAlignShapeType.ALIGN_MIDDLE;



prot = window['Asc']['c_oAscFootnotePos'] = window['Asc'].c_oAscFootnotePos = c_oAscFootnotePos;
prot['BeneathText'] = c_oAscFootnotePos.BeneathText;
prot['DocEnd']      = c_oAscFootnotePos.DocEnd;
prot['PageBottom']  = c_oAscFootnotePos.PageBottom;
prot['SectEnd']     = c_oAscFootnotePos.SectEnd;

prot = window['Asc']['c_oAscEndnotePos'] = window['Asc'].c_oAscEndnotePos = c_oAscEndnotePos;
prot['DocEnd']      = c_oAscEndnotePos.DocEnd;
prot['SectEnd']     = c_oAscEndnotePos.SectEnd;

prot = window['Asc']['c_oAscFootnoteRestart'] = window['Asc'].c_oAscFootnoteRestart = c_oAscFootnoteRestart;
prot['Continuous'] = c_oAscFootnoteRestart.Continuous;
prot['EachSect']   = c_oAscFootnoteRestart.EachSect;
prot['EachPage']   = c_oAscFootnoteRestart.EachPage;


prot = window['Asc']['c_oAscSdtLevelType'] = window['Asc'].c_oAscSdtLevelType = c_oAscSdtLevelType;
prot['Block']  = c_oAscSdtLevelType.Block;
prot['Inline'] = c_oAscSdtLevelType.Inline;
prot['Row']    = c_oAscSdtLevelType.Row;
prot['Cell']   = c_oAscSdtLevelType.Cell;

prot = window['Asc']['c_oAscTOCStylesType'] = window['Asc'].c_oAscTOCStylesType = c_oAscTOCStylesType;
prot['Current']  = c_oAscTOCStylesType.Current;
prot['Simple']   = c_oAscTOCStylesType.Simple;
prot['Standard'] = c_oAscTOCStylesType.Standard;
prot['Modern']   = c_oAscTOCStylesType.Modern;
prot['Classic']  = c_oAscTOCStylesType.Classic;
prot['Web']      = c_oAscTOCStylesType.Web;

prot = window['Asc']['c_oAscTOFStylesType'] = window['Asc'].c_oAscTOFStylesType = c_oAscTOFStylesType;
prot['Current']     = prot.Current;
prot['Classic']     = prot.Classic;
prot['Distinctive'] = prot.Distinctive;
prot['Centered']    = prot.Centered;
prot['Formal']      = prot.Formal;
prot['Simple']      = prot.Simple;
prot['Web']         = prot.Web;

prot = window['Asc']['c_oAscStyleType'] = window['Asc'].c_oAscStyleType = c_oAscStyleType;
prot['Paragraph'] = c_oAscStyleType.Paragraph;
prot['Numbering'] = c_oAscStyleType.Numbering;
prot['Table']     = c_oAscStyleType.Table;
prot['Character'] = c_oAscStyleType.Character;

prot = window['Asc']['c_oAscHyperlinkAnchor'] = window['Asc'].c_oAscHyperlinkAnchor = c_oAscHyperlinkAnchor;
prot['Heading']       = c_oAscHyperlinkAnchor.Heading;
prot['Bookmark']      = c_oAscHyperlinkAnchor.Bookmark;

window['AscCommon']                = window['AscCommon'] || {};
window['AscCommon'].c_oSerFormat   = c_oSerFormat;
window['AscCommon'].CurFileVersion = c_oSerFormat.Version;

prot = window['Asc']['c_oAscSdtCheckBoxDefaults'] = window['Asc'].c_oAscSdtCheckBoxDefaults = c_oAscSdtCheckBoxDefaults;
prot['CheckedSymbol']   = prot.CheckedSymbol;
prot['UncheckedSymbol'] = prot.UncheckedSymbol;
prot['CheckedFont']     = prot.CheckedFont;
prot['UncheckedFont']   = prot.UncheckedFont;

window['Asc']['NoYLimit'] = window['Asc'].NoYLimit = 0x7FFFFFFF;

prot = window['Asc']['c_oAscDocumentShortcutType'] = window['Asc'].c_oAscDocumentShortcutType = c_oAscDocumentShortcutType;
prot['InsertPageBreak']      = prot.InsertPageBreak;
prot['InsertLineBreak']      = prot.InsertLineBreak;
prot['InsertColumnBreak']    = prot.InsertColumnBreak;
prot['ResetChar']            = prot.ResetChar;
prot['NonBreakingSpace']     = prot.NonBreakingSpace;
prot['ApplyHeading1']        = prot.ApplyHeading1;
prot['ApplyHeading2']        = prot.ApplyHeading2;
prot['ApplyHeading3']        = prot.ApplyHeading3;
prot['Strikeout']            = prot.Strikeout;
prot['ShowAll']              = prot.ShowAll;
prot['EditSelectAll']        = prot.EditSelectAll;
prot['Bold']                 = prot.Bold;
prot['CopyFormat']           = prot.CopyFormat;
prot['CopyrightSign']        = prot.CopyrightSign;
prot['InsertEndnoteNow']     = prot.InsertEndnoteNow;
prot['CenterPara']           = prot.CenterPara;
prot['EuroSign']             = prot.EuroSign;
prot['InsertFootnoteNow']    = prot.InsertFootnoteNow;
prot['Italic']               = prot.Italic;
prot['JustifyPara']          = prot.JustifyPara;
prot['InsertHyperlink']      = prot.InsertHyperlink;
prot['ApplyListBullet']      = prot.ApplyListBullet;
prot['LeftPara']             = prot.LeftPara;
prot['Indent']               = prot.Indent;
prot['UnIndent']             = prot.UnIndent;
prot['PrintPreviewAndPrint'] = prot.PrintPreviewAndPrint;
prot['InsertPageNumber']     = prot.InsertPageNumber;
prot['RegisteredSign']       = prot.RegisteredSign;
prot['RightPara']            = prot.RightPara;
prot['Save']                 = prot.Save;
prot['TrademarkSign']        = prot.TrademarkSign;
prot['Underline']            = prot.Underline;
prot['PasteFormat']          = prot.PasteFormat;
prot['EditUndo']             = prot.EditUndo;
prot['EditRedo']             = prot.EditRedo;
prot['EmDash']               = prot.EmDash;
prot['EnDash']               = prot.EnDash;
prot['UpdateFields']         = prot.UpdateFields;
prot['InsertEquation']       = prot.InsertEquation;
prot['Superscript']          = prot.Superscript;
prot['NonBreakingHyphen']    = prot.NonBreakingHyphen;
prot['SoftHyphen']           = prot.SoftHyphen;
prot['HorizontalEllipsis']   = prot.HorizontalEllipsis;
prot['Subscript']            = prot.Subscript;
prot['IncreaseFontSize']     = prot.IncreaseFontSize;
prot['DecreaseFontSize']     = prot.DecreaseFontSize;

prot = window['Asc']['c_oAscDocumentRefenceToType'] = window['Asc'].c_oAscDocumentRefenceToType = c_oAscDocumentRefenceToType;
prot['Text']               = prot.Text;
prot['PageNum']            = prot.PageNum;
prot['ParaNum']            = prot.ParaNum;
prot['ParaNumNoContext']   = prot.ParaNumNoContext;
prot['ParaNumFullContex']  = prot.ParaNumFullContex;
prot['AboveBelow']         = prot.AboveBelow;
prot['OnlyLabelAndNumber'] = prot.OnlyLabelAndNumber;
prot['OnlyCaptionText']    = prot.OnlyCaptionText;
prot['NoteNumber']         = prot.NoteNumber;
prot['NoteNumberFormatted'] = prot.NoteNumberFormatted;

prot = window['Asc']['c_oAscTextToTableSeparator'] = window['Asc'].c_oAscTextToTableSeparator = c_oAscTextToTableSeparator;
prot['Paragraph'] = prot.Paragraph;
prot['Tab']       = prot.Tab;
prot['Symbol']    = prot.Symbol;

prot = window['Asc']['c_oAscTextToTableAutoFitType'] = window['Asc'].c_oAscTextToTableAutoFitType = c_oAscTextToTableAutoFitType;
prot['Fixed']   = prot.Fixed;
prot['Content'] = prot.Content;
prot['Window']  = prot.Window;

window['Asc']['c_oAscDateTimeFormat'] = window['Asc'].c_oAscDateTimeFormat = c_oAscDateTimeFormat;
