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

(function()
{
	const DEFAULT_HEADING_FONT = "Arial";
	const DEFAULT_FONT         = "Arial";
	
	function tw2mm(tw)
	{
		return tw * g_dKoef_twips_to_mm
	}
	function pt2mm(pt)
	{
		return pt * g_dKoef_pt_to_mm;
	}
	
	const DEFAULT_STYLE_LIST = [
		{
			StyleId : "Normal",
			Default : true,
			Type    : AscWord.styletype_Paragraph,
			Name    : "Normal",
			QFormat : true
		},
		{
			StyleId    : "Heading1",
			Type       : AscWord.styletype_Paragraph,
			Name       : "Heading 1",
			UiPriority : 9,
			BasedOn    : "Normal",
			Next       : "Normal",
			Link       : "Heading1Char",
			QFormat    : true,
			ParaPr     : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					Before : tw2mm(360),
					After  : tw2mm(80)
				},
				OutlineLvl : 0
			},
			TextPr     : {
				FontSize   : 20,
				FontSizeCS : 20,
				Color      : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "Heading2",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Heading 2",
			BasedOn        : "Normal",
			Next           : "Normal",
			Link           : "Heading2Char",
			UiPriority     : 9,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					Before : tw2mm(160),
					After  : tw2mm(80)
				},
				OutlineLvl : 1
			},
			TextPr         : {
				FontSize   : 16,
				FontSizeCS : 16,
				Color      : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "Heading3",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Heading 3",
			BasedOn        : "Normal",
			Next           : "Normal",
			Link           : "Heading3Char",
			UiPriority     : 9,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					Before : tw2mm(160),
					After  : tw2mm(80)
				},
				OutlineLvl : 2
			},
			TextPr         : {
				FontSize   : 14,
				FontSizeCS : 14,
				Color      : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "Heading4",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Heading 4",
			BasedOn        : "Normal",
			Next           : "Normal",
			Link           : "Heading4Char",
			UiPriority     : 9,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					Before : tw2mm(80),
					After  : tw2mm(40)
				},
				OutlineLvl : 3
			},
			TextPr         : {
				Italic     : true,
				ItalicCS   : true,
				Color      : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "Heading5",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Heading 5",
			BasedOn        : "Normal",
			Next           : "Normal",
			Link           : "Heading5Char",
			UiPriority     : 9,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					Before : tw2mm(80),
					After  : tw2mm(40)
				},
				OutlineLvl : 4
			},
			TextPr         : {
				Color      : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "Heading6",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Heading 6",
			BasedOn        : "Normal",
			Next           : "Normal",
			Link           : "Heading6Char",
			UiPriority     : 9,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					Before : tw2mm(40),
					After  : 0
				},
				OutlineLvl : 5
			},
			TextPr         : {
				Italic     : true,
				ItalicCS   : true,
				Color      : {r : 0x59, g : 0x59, b : 0x59},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xA6, null),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "Heading7",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Heading 7",
			BasedOn        : "Normal",
			Next           : "Normal",
			Link           : "Heading7Char",
			UiPriority     : 9,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					Before : tw2mm(40),
					After  : 0
				},
				OutlineLvl : 6
			},
			TextPr         : {
				Color      : {r : 0x59, g : 0x59, b : 0x59},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xA6, null),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "Heading8",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Heading 8",
			BasedOn        : "Normal",
			Next           : "Normal",
			Link           : "Heading8Char",
			UiPriority     : 9,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					After : 0
				},
				OutlineLvl : 7
			},
			TextPr         : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x27, g : 0x27, b : 0x27},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xD8, null),
				RFonts   : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "Heading9",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Heading 9",
			BasedOn        : "Normal",
			Next           : "Normal",
			Link           : "Heading9Char",
			UiPriority     : 9,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				KeepNext   : true,
				KeepLines  : true,
				Spacing    : {
					After : 0
				},
				OutlineLvl : 8
			},
			TextPr         : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x27, g : 0x27, b : 0x27},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xD8, null),
				RFonts   : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId        : "DefaultParagraphFont",
			Type           : AscWord.styletype_Character,
			Name           : "Default Paragraph Font",
			Default        : true,
			UiPriority     : 1,
			SemiHidden     : true,
			UnhideWhenUsed : true
		},
		{
			StyleId        : "NoList",
			Default        : true,
			Type           : AscWord.styletype_Numbering,
			Name           : "No List",
			UiPriority     : 99,
			SemiHidden     : true,
			UnhideWhenUsed : true,
		},
		{
			StyleId    : "Heading1Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 1 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading1",
			UiPriority : 9,
			TextPr     : {
				FontSize   : 20,
				FontSizeCS : 20,
				Color      : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Heading2Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 2 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading2",
			UiPriority : 9,
			TextPr     : {
				FontSize   : 16,
				FontSizeCS : 16,
				Color      : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Heading3Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 3 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading3",
			UiPriority : 9,
			TextPr     : {
				FontSize   : 14,
				FontSizeCS : 14,
				Color      : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts     : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Heading4Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 4 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading4",
			UiPriority : 9,
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts   : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Heading5Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 5 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading5",
			UiPriority : 9,
			TextPr     : {
				Color   : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF),
				RFonts  : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Heading6Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 6 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading6",
			UiPriority : 9,
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x59, g : 0x59, b : 0x59},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xA6, null),
				RFonts   : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Heading7Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 7 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading7",
			UiPriority : 9,
			TextPr     : {
				Color   : {r : 0x59, g : 0x59, b : 0x59},
				Unifill : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xA6, null),
				RFonts  : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Heading8Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 8 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading8",
			UiPriority : 9,
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x27, g : 0x27, b : 0x27},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xD8, null),
				RFonts   : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Heading9Char",
			Type       : AscWord.styletype_Character,
			Name       : "Heading 9 Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Heading9",
			UiPriority : 9,
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x27, g : 0x27, b : 0x27},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xD8, null),
				RFonts   : {
					Ascii    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_HEADING_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_HEADING_FONT, Index : -1},
					CS       : {Name : DEFAULT_HEADING_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Title",
			Type       : AscWord.styletype_Paragraph,
			Name       : "Title",
			BasedOn    : "Normal",
			Next       : "Normal",
			Link       : "TitleChar",
			UiPriority : 10,
			QFormat    : true,
			ParaPr     : {
				Spacing           : {
					After    : tw2mm(80),
					Line     : 1,
					LineRule : Asc.linerule_Auto
				},
				ContextualSpacing : true
			},
			TextPr     : {
				Spacing    : tw2mm(-10),
				FontSize   : 28,
				FontSizeCS : 28,
				RFonts     : {
					Ascii    : {Name : DEFAULT_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_FONT, Index : -1},
					CS       : {Name : DEFAULT_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "TitleChar",
			Type       : AscWord.styletype_Character,
			Name       : "Title Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Title",
			UiPriority : 10,
			TextPr     : {
				Spacing    : tw2mm(-10),
				FontSize   : 28,
				FontSizeCS : 28,
				RFonts     : {
					Ascii    : {Name : DEFAULT_FONT, Index : -1},
					EastAsia : {Name : DEFAULT_FONT, Index : -1},
					HAnsi    : {Name : DEFAULT_FONT, Index : -1},
					CS       : {Name : DEFAULT_FONT, Index : -1}
				}
			}
		},
		{
			StyleId    : "Subtitle",
			Type       : AscWord.styletype_Paragraph,
			Name       : "Subtitle",
			BasedOn    : "Normal",
			Next       : "Normal",
			Link       : "SubtitleChar",
			UiPriority : 11,
			QFormat    : true,
			ParaPr     : {
				NumPr : {
					Lvl : 1
				}
			},
			TextPr     : {
				Color      : {r : 0x59, g : 0x59, b : 0x59},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xA6, null),
				Spacing    : tw2mm(15),
				FontSize   : 14,
				FontSizeCS : 14
			}
		},
		{
			StyleId    : "SubtitleChar",
			Type       : AscWord.styletype_Character,
			Name       : "Subtitle Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Subtitle",
			UiPriority : 11,
			TextPr     : {
				Color      : {r : 0x59, g : 0x59, b : 0x59},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xA6, null),
				Spacing    : tw2mm(15),
				FontSize   : 14,
				FontSizeCS : 14
			}
		},
		{
			StyleId    : "Quote",
			Type       : AscWord.styletype_Paragraph,
			Name       : "Quote",
			BasedOn    : "Normal",
			Next       : "Normal",
			Link       : "QuoteChar",
			UiPriority : 29,
			QFormat    : true,
			ParaPr     : {
				Spacing : {
					Before : tw2mm(160)
				},
				Jc      : AscCommon.align_Center
			},
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x40, g : 0x40, b : 0x40},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xBF, null)
			}
		},
		{
			StyleId    : "QuoteChar",
			Type       : AscWord.styletype_Character,
			Name       : "Quote Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Quote",
			UiPriority : 29,
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x40, g : 0x40, b : 0x40},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xBF, null)
			}
		},
		{
			StyleId    : "ListParagraph",
			Type       : AscWord.styletype_Paragraph,
			Name       : "List Paragraph",
			BasedOn    : "Normal",
			UiPriority : 34,
			QFormat    : true,
			ParaPr     : {
				Ind : {
					Left : tw2mm(720)
				},
				
				ContextualSpacing : true
			}
		},
		{
			StyleId        : "IntenseEmphasis",
			Type           : AscWord.styletype_Character,
			Name           : "Intense Emphasis",
			BasedOn        : "DefaultParagraphFont",
			UiPriority     : 21,
			QFormat        : true,
			TextPr         : {
				Italic : true,
				ItalicCS : true,
				Color    : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF)
			}
		},
		{
			StyleId    : "IntenseQuote",
			Type       : AscWord.styletype_Paragraph,
			Name       : "Intense Quote",
			BasedOn    : "Normal",
			Next       : "Normal",
			Link       : "IntenseQuoteChar",
			UiPriority : 30,
			QFormat    : true,
			ParaPr     : {
				Brd : {
					Top    : {
						Value   : border_Single,
						Size    : pt2mm(0.5),
						Space   : pt2mm(10),
						Color   : {r : 0x0F, g : 0x47, b : 0x61},
						Unifill : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF)
					},
					Bottom : {
						Value   : border_Single,
						Size    : pt2mm(0.5),
						Space   : pt2mm(10),
						Color   : {r : 0x0F, g : 0x47, b : 0x61},
						Unifill : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF)
					}
				},
				
				Spacing : {
					Before : tw2mm(360),
					After  : tw2mm(360)
				},
				Ind     : {
					Left  : tw2mm(864),
					Right : tw2mm(864)
				},
				Jc      : AscCommon.align_Center
			},
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF)
			}
		},
		{
			StyleId    : "IntenseQuoteChar",
			Type       : AscWord.styletype_Character,
			Name       : "Intense Quote Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "IntenseQuote",
			UiPriority : 30,
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF)
			}
		},
		{
			StyleId    : "IntenseReference",
			Type       : AscWord.styletype_Character,
			Name       : "Intense Reference",
			BasedOn    : "DefaultParagraphFont",
			UiPriority : 32,
			QFormat    : true,
			TextPr     : {
				Bold      : true,
				BoldCS    : true,
				SmallCaps : true,
				Spacing   : tw2mm(5),
				Color     : {r : 0x0F, g : 0x47, b : 0x61},
				Unifill   : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorAccent1, null, 0xBF)
			}
		},
		{
			StyleId    : "NoSpacing",
			Type       : AscWord.styletype_Paragraph,
			Name       : "No Spacing",
			BasedOn    : "Normal",
			UiPriority : 1,
			QFormat    : true,
			ParaPr     : {
				Spacing : {
					After    : 0,
					Line     : 1,
					LineRule : Asc.linerule_Auto
				}
			}
		},
		{
			StyleId    : "SubtleEmphasis",
			Type       : AscWord.styletype_Character,
			Name       : "Subtle Emphasis",
			BasedOn    : "DefaultParagraphFont",
			UiPriority : 19,
			QFormat    : true,
			TextPr     : {
				Italic   : true,
				ItalicCS : true,
				Color    : {r : 0x40, g : 0x40, b : 0x40},
				Unifill  : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xBF, null)
			}
		},
		{
			StyleId    : "Emphasis",
			Type       : AscWord.styletype_Character,
			Name       : "Emphasis",
			BasedOn    : "DefaultParagraphFont",
			UiPriority : 20,
			QFormat    : true,
			TextPr     : {
				Italic   : true,
				ItalicCS : true
			}
		},
		{
			StyleId    : "Strong",
			Type       : AscWord.styletype_Character,
			Name       : "Strong",
			BasedOn    : "DefaultParagraphFont",
			UiPriority : 22,
			QFormat    : true,
			TextPr     : {
				Bold   : true,
				BoldCS : true
			}
		},
		{
			StyleId    : "SubtleReference",
			Type       : AscWord.styletype_Character,
			Name       : "Subtle Reference",
			BasedOn    : "DefaultParagraphFont",
			UiPriority : 31,
			QFormat    : true,
			TextPr     : {
				SmallCaps : true,
				Color     : {r : 0x5A, g : 0x5A, b : 0x5A},
				Unifill   : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText1, 0xA5, null)
			}
		},
		{
			StyleId    : "BookTitle",
			Type       : AscWord.styletype_Character,
			Name       : "Book Title",
			BasedOn    : "DefaultParagraphFont",
			UiPriority : 33,
			QFormat    : true,
			TextPr     : {
				Bold     : true,
				BoldCS   : true,
				Italic   : true,
				ItalicCS : true,
				Spacing  : tw2mm(5)
			}
		},
		{
			StyleId        : "Header",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Header",
			BasedOn        : "Normal",
			Link           : "HeaderChar",
			UiPriority     : 99,
			UnhideWhenUsed : true,
			ParaPr         : {
				Tabs    : {
					Tabs : [
						{
							Value : Asc.c_oAscTabType.Center,
							Pos   : tw2mm(4844)
						},
						{
							Value : Asc.c_oAscTabType.Right,
							Pos   : tw2mm(9689)
						}
					]
				},
				Spacing : {
					After    : 0,
					Line     : 1,
					LineRule : Asc.linerule_Auto
				}
			}
		},
		{
			StyleId    : "HeaderChar",
			Type       : AscWord.styletype_Character,
			Name       : "Header Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Header",
			UiPriority : 99
		},
		{
			StyleId        : "Footer",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Footer",
			BasedOn        : "Normal",
			Link           : "FooterChar",
			UiPriority     : 99,
			UnhideWhenUsed : true,
			ParaPr         : {
				Tabs : {
					Tabs : [
						{
							Value : Asc.c_oAscTabType.Center,
							Pos   : tw2mm(4844)
						},
						{
							Value : Asc.c_oAscTabType.Right,
							Pos   : tw2mm(9689)
						}
					]
				},
				Spacing : {
					After : 0,
					Line  : 1,
					LineRule : Asc.linerule_Auto
				}
			}
		},
		{
			StyleId    : "FooterChar",
			Type       : AscWord.styletype_Character,
			Name       : "Footer Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "Footer",
			UiPriority : 99
		},
		{
			StyleId        : "Caption",
			Type           : AscWord.styletype_Paragraph,
			Name           : "Caption",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 35,
			UnhideWhenUsed : true,
			QFormat        : true,
			ParaPr         : {
				Spacing : {
					After    : tw2mm(200),
					Line     : 1,
					LineRule : Asc.linerule_Auto
				}
			},
			TextPr         : {
				Italic     : true,
				ItalicCS   : true,
				FontSize   : 9,
				FontSizeCS : 9,
				Color      : {r : 0x0E, g : 0x28, b : 0x41},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorText2, null, null)
			}
		},
		{
			StyleId        : "FootnoteText",
			Type           : AscWord.styletype_Paragraph,
			Name           : "footnote text",
			BasedOn        : "Normal",
			Link           : "FootnoteTextChar",
			UiPriority     : 99,
			SemiHidden     : true,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After    : 0,
					Line     : 1,
					LineRule : Asc.linerule_Auto
				}
			},
			TextPr         : {
				FontSize   : 10,
				FontSizeCS : 10
			}
		},
		{
			StyleId    : "FootnoteTextChar",
			Type       : AscWord.styletype_Character,
			Name       : "Footnote Text Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "FootnoteText",
			UiPriority : 99,
			SemiHidden : true,
			TextPr     : {
				FontSize   : 10,
				FontSizeCS : 10
			}
		},
		{
			StyleId        : "FootnoteReference",
			Type           : AscWord.styletype_Character,
			Name           : "footnote reference",
			BasedOn        : "DefaultParagraphFont",
			UiPriority     : 99,
			SemiHidden     : true,
			UnhideWhenUsed : true,
			TextPr         : {
				VertAlign : AscCommon.vertalign_SuperScript
			}
		},
		{
			StyleId        : "EndnoteText",
			Type           : AscWord.styletype_Paragraph,
			Name           : "endnote text",
			BasedOn        : "Normal",
			Link           : "EndnoteTextChar",
			UiPriority     : 99,
			SemiHidden     : true,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After    : 0,
					Line     : 1,
					LineRule : Asc.linerule_Auto
				}
			},
			TextPr         : {
				FontSize   : 10,
				FontSizeCS : 10
			}
		},
		{
			StyleId    : "EndnoteTextChar",
			Type       : AscWord.styletype_Character,
			Name       : "Endnote Text Char",
			BasedOn    : "DefaultParagraphFont",
			Link       : "EndnoteText",
			UiPriority : 99,
			SemiHidden : true,
			TextPr     : {
				FontSize   : 10,
				FontSizeCS : 10
			}
		},
		{
			StyleId        : "EndnoteReference",
			Type           : AscWord.styletype_Character,
			Name           : "endnote reference",
			BasedOn        : "DefaultParagraphFont",
			UiPriority     : 99,
			SemiHidden     : true,
			UnhideWhenUsed : true,
			TextPr         : {
				VertAlign : AscCommon.vertalign_SuperScript
			}
		},
		{
			StyleId        : "Hyperlink",
			Type           : AscWord.styletype_Character,
			Name           : "Hyperlink",
			BasedOn        : "DefaultParagraphFont",
			UiPriority     : 99,
			UnhideWhenUsed : true,
			TextPr         : {
				Color      : {r : 0x05, g : 0x63, b : 0xC1},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorHyperlink, null, null),
				Underline  : true
			}
		},
		{
			StyleId        : "FollowedHyperlink",
			Type           : AscWord.styletype_Character,
			Name           : "FollowedHyperlink",
			BasedOn        : "DefaultParagraphFont",
			UiPriority     : 99,
			SemiHidden     : true,
			UnhideWhenUsed : true,
			TextPr         : {
				Color      : {r : 0x95, g : 0x4F, b : 0x72},
				Unifill    : AscCommonWord.CreateThemeUnifill(EThemeColor.themecolorFollowedHyperlink, null, null),
				Underline  : true
			}
		},
		{
			StyleId        : "Toc1",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 1",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				}
			}
		},
		{
			StyleId        : "Toc2",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 2",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				},
				Ind     : {
					Left : tw2mm(220)
				}
			}
		},
		{
			StyleId        : "Toc3",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 3",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				},
				Ind     : {
					Left : tw2mm(440)
				}
			}
		},
		{
			StyleId        : "Toc4",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 4",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				},
				Ind     : {
					Left : tw2mm(660)
				}
			}
		},
		{
			StyleId        : "Toc5",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 5",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				},
				Ind     : {
					Left : tw2mm(880)
				}
			}
		},
		{
			StyleId        : "Toc6",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 6",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				},
				Ind     : {
					Left : tw2mm(1100)
				}
			}
		},
		{
			StyleId        : "Toc7",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 7",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				},
				Ind     : {
					Left : tw2mm(1320)
				}
			}
		},
		{
			StyleId        : "Toc8",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 8",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				},
				Ind     : {
					Left : tw2mm(1540)
				}
			}
		},
		{
			StyleId        : "Toc9",
			Type           : AscWord.styletype_Paragraph,
			Name           : "toc 9",
			BasedOn        : "Normal",
			Next           : "Normal",
			UiPriority     : 39,
			UnhideWhenUsed : true,
			ParaPr         : {
				Spacing : {
					After : tw2mm(100)
				},
				Ind     : {
					Left : tw2mm(1760)
				}
			}
		}
	];
	
	function getDefaultStyleObject(styleName)
	{
		for (let i = 0; i < DEFAULT_STYLE_LIST.length; ++i)
		{
			if (DEFAULT_STYLE_LIST[i].Name === styleName)
				return DEFAULT_STYLE_LIST[i];
		}
		
		return null;
	}
	/**
	 *  StyleId здесь - это местный идентификтор, а не тот что используется в менеджере стилей
	 * @param defaultStyleId
	 * @returns {string}
	 */
	function getDefaultStyleName(defaultStyleId)
	{
		for (let i = 0; i < DEFAULT_STYLE_LIST.length; ++i)
		{
			if (DEFAULT_STYLE_LIST[i].StyleId === defaultStyleId)
				return DEFAULT_STYLE_LIST[i].Name;
		}
		
		return "";
	}
	
	function setDefaultHeadingColor(r, g, b)
	{
		for (let i = 0; i < DEFAULT_STYLE_LIST.length; ++i)
		{
			if (DEFAULT_STYLE_LIST[i].StyleId
				&& DEFAULT_STYLE_LIST[i].TextPr
				&& 0 === DEFAULT_STYLE_LIST[i].StyleId.indexOf("Heading"))
			{
				DEFAULT_STYLE_LIST[i].TextPr.Color   = {r : r, g : g, b : b};
				DEFAULT_STYLE_LIST[i].TextPr.Unifill = undefined;
			}
		}
	}
	function setDefaultHeadingColorStr(rgbStr)
	{
		let rgba = AscCommon.RgbaTextToRGBA(rgbStr);
		setDefaultHeadingColor(rgba.R, rgba.G, rgba.B);
	}
	
	//--------------------------------------------------------export----------------------------------------------------
	AscWord.DEFAULT_STYLE_LIST     = DEFAULT_STYLE_LIST;
	AscWord.DEFAULT_HEADING_FONT   = DEFAULT_HEADING_FONT;
	AscWord.DEFAULT_FONT           = DEFAULT_FONT;
	AscWord.getDefaultStyleObject  = getDefaultStyleObject;
	AscWord.getDefaultStyleName    = getDefaultStyleName;
	AscWord["setDefaultHeadingColor"] = AscWord.setDefaultHeadingColor = setDefaultHeadingColor;
	AscWord["setDefaultHeadingColorStr"] = AscWord.setDefaultHeadingColorStr = setDefaultHeadingColorStr;
})(window);
