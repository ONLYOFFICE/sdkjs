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

var TABLE_STYLE_WIDTH_PIX  = 70;
var TABLE_STYLE_HEIGHT_PIX = 50;




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
	BeneathText : 0x00, //section_footnote_PosBeneathText
	DocEnd      : 0x01, //section_footnote_PosDocEnd
	PageBottom  : 0x02, //section_footnote_PosPageBottom,
	SectEnd     : 0x03  //section_footnote_PosSectEnd
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
	AddPageBreak                    : 1,
	AddLineBreak                    : 2,
	AddColumnBreak                  : 3,
	ClearFormatting                 : 4,
	NonBreakingSpace                : 5,
	StyleHeading1                   : 6,
	StyleHeading2                   : 7,
	StyleHeading3                   : 8,
	Strikeout                       : 9,
	ToggleNonPrintingCharacters     : 10,
	SelectAll                       : 11,
	Bold                            : 12,
	CopyFormatting                  : 13,
	CopyrightSign                   : 14,
	AddEndnote                      : 15,
	ToggleParagraphAlignCenterLeft  : 16,
	EuroSign                        : 17,
	AddFootnote                     : 18,
	Italic                          : 19,
	ToggleParagraphAlignJustifyLeft : 20,
	AddHyperlink                    : 21,
	AddBulletedList                 : 22,
	ToggleParagraphAlignLeftJustify : 23,
	IncreaseIndent                  : 24,
	DecreaseIndent                  : 25,
	Print                           : 26,
	AddPageNumber                   : 27,
	RegisteredSign                  : 28,
	ToggleParagraphAlignRightLeft   : 29,
	Save                            : 30,
	TrademarkSign                   : 31,
	Underline                       : 32,
	PasteFormatting                 : 33,
	Undo                            : 34,
	Redo                            : 35,
	EmDash                          : 36,
	EnDash                          : 37,
	UpdateFields                    : 38,
	AddEquation                     : 39,
	Superscript                     : 40,
	NonBreakingHyphen               : 41,
	SoftHyphen                      : 42,
	HorizontalEllipsis              : 43,
	Subscript                       : 44,
	IncreaseFontSize                : 45,
	DecreaseFontSize                : 46
};

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




prot = window['Asc']['c_oAscFootnotePos'] = c_oAscFootnotePos;
prot['BeneathText'] = c_oAscFootnotePos.BeneathText;
prot['DocEnd']      = c_oAscFootnotePos.DocEnd;
prot['PageBottom']  = c_oAscFootnotePos.PageBottom;
prot['SectEnd']     = c_oAscFootnotePos.SectEnd;

prot = window['Asc']['c_oAscFootnoteRestart'] = c_oAscFootnoteRestart;
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
prot['AddPageBreak']                    = prot.AddPageBreak;
prot['AddLineBreak']                    = prot.AddLineBreak;
prot['AddColumnBreak']                  = prot.AddColumnBreak;
prot['ClearFormatting']                 = prot.ClearFormatting;
prot['NonBreakingSpace']                = prot.NonBreakingSpace;
prot['StyleHeading1']                   = prot.StyleHeading1;
prot['StyleHeading2']                   = prot.StyleHeading2;
prot['StyleHeading3']                   = prot.StyleHeading3;
prot['Strikeout']                       = prot.Strikeout;
prot['ToggleNonPrintingCharacters']     = prot.ToggleNonPrintingCharacters;
prot['SelectAll']                       = prot.SelectAll;
prot['Bold']                            = prot.Bold;
prot['CopyFormatting']                  = prot.CopyFormatting;
prot['CopyrightSign']                   = prot.CopyrightSign;
prot['AddEndnote']                      = prot.AddEndnote;
prot['ToggleParagraphAlignCenterLeft']  = prot.ToggleParagraphAlignCenterLeft;
prot['EuroSign']                        = prot.EuroSign;
prot['AddFootnote']                     = prot.AddFootnote;
prot['Italic']                          = prot.Italic;
prot['ToggleParagraphAlignJustifyLeft'] = prot.ToggleParagraphAlignJustifyLeft;
prot['AddHyperlink']                    = prot.AddHyperlink;
prot['AddBulletedList']                 = prot.AddBulletedList;
prot['ToggleParagraphAlignLeftJustify'] = prot.ToggleParagraphAlignLeftJustify;
prot['IncreaseIndent']                  = prot.IncreaseIndent;
prot['DecreaseIndent']                  = prot.DecreaseIndent;
prot['Print']                           = prot.Print;
prot['AddPageNumber']                   = prot.AddPageNumber;
prot['RegisteredSign']                  = prot.RegisteredSign;
prot['ToggleParagraphAlignRightLeft']   = prot.ToggleParagraphAlignRightLeft;
prot['Save']                            = prot.Save;
prot['TrademarkSign']                   = prot.TrademarkSign;
prot['Underline']                       = prot.Underline;
prot['PasteFormatting']                 = prot.PasteFormatting;
prot['Undo']                            = prot.Undo;
prot['Redo']                            = prot.Redo;
prot['EmDash']                          = prot.EmDash;
prot['EnDash']                          = prot.EnDash;
prot['UpdateFields']                    = prot.UpdateFields;
prot['AddEquation']                     = prot.AddEquation;
prot['Superscript']                     = prot.Superscript;
prot['NonBreakingHyphen']               = prot.NonBreakingHyphen;
prot['SoftHyphen']                      = prot.SoftHyphen;
prot['HorizontalEllipsis']              = prot.HorizontalEllipsis;
prot['Subscript']                       = prot.Subscript;
prot['IncreaseFontSize']                = prot.IncreaseFontSize;
prot['DecreaseFontSize']                = prot.DecreaseFontSize;
