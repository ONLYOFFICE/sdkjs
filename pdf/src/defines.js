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

(function(){
    if (!window["AscPDF"])
	    window["AscPDF"] = {};

    let asc = window["AscPDF"];

    let ANNOTATIONS_TYPES = {
        Text:           0,
        Link:           1,
        FreeText:       2,
        Line:           3,
        Square:         4,
        Circle:         5,
        Polygon:        6,
        PolyLine:       7,
        Highlight:      8,
        Underline:      9,
        Squiggly:       10,
        Strikeout:      11,
        Stamp:          12,
        Caret:          13,
        Ink:            14,
        Popup:          15,
        FileAttachment: 16,
        Sound:          17,
        Movie:          18,
        Widget:         19,
        Screen:         20,
        PrinterMark:    21,
        TrapNet:        22,
        Watermark:      23,
        Type3D:         24,
        Redact:         25
    };

    let BORDER_EFFECT_STYLES = {
        None:   0,
        Cloud:  1
    }

    let REF_TO_REASON = {
        Reply: 0,
        Group: 1
    }

    ANNOTATIONS_TYPES["Underline"]   = ANNOTATIONS_TYPES.Underline;
    ANNOTATIONS_TYPES["Strikeout"]   = ANNOTATIONS_TYPES.Strikeout;
    ANNOTATIONS_TYPES["Highlight"]   = ANNOTATIONS_TYPES.Highlight;
    asc["ANNOTATIONS_TYPES"] = asc.ANNOTATIONS_TYPES = ANNOTATIONS_TYPES;

    let FIELD_TYPES = {
        unknown:        26,
        button:         27,
        radiobutton:    28,
        checkbox:       29,
        text:           30,
        combobox:       31,
        listbox:        32,
        signature:      33
    };

    FIELD_TYPES["unknown"]      = FIELD_TYPES.unknown;
    FIELD_TYPES["button"]       = FIELD_TYPES.button;
    FIELD_TYPES["radiobutton"]  = FIELD_TYPES.radiobutton;
    FIELD_TYPES["checkbox"]     = FIELD_TYPES.checkbox;
    FIELD_TYPES["text"]         = FIELD_TYPES.text;
    FIELD_TYPES["combobox"]     = FIELD_TYPES.combobox;
    FIELD_TYPES["listbox"]      = FIELD_TYPES.listbox;
    FIELD_TYPES["signature"]    = FIELD_TYPES.signature;

    let CommandType = {
        editPage:   0,
        addPage:    1,
        removePage: 2,
        widgetInfo: 3,
        movePage:   4
    }

    const FormatType = {
        NONE:       -1,
        NUMBER:     0,
        PERCENTAGE: 1,
        DATE:       2,
        TIME:       3,
        SPECIAL:    4,
        REGULAR:    5,
        CUSTOM:     6
    }

    FormatType["NONE"]          = FormatType.NONE;
    FormatType["NUMBER"]        = FormatType.NUMBER;
    FormatType["PERCENTAGE"]    = FormatType.PERCENTAGE;
    FormatType["DATE"]          = FormatType.DATE;
    FormatType["TIME"]          = FormatType.TIME;
    FormatType["SPECIAL"]       = FormatType.SPECIAL;
    FormatType["REGULAR"]       = FormatType.REGULAR;
    FormatType["CUSTOM"]        = FormatType.CUSTOM;

    const SpectialFormatType = {
        ZIP_CODE: 0,       // Почтовый индекс (ZIP Code)
        ZIP_PLUS_4: 1,     // Почтовый индекс + 4 (ZIP + 4)
        PHONE: 2,          // Телефонный номер
        SSN: 3             // Социальное страхование (SSN)
    }

    SpectialFormatType["ZIP_CODE"] = SpectialFormatType.ZIP_CODE;
    SpectialFormatType["ZIP_PLUS_4"] = SpectialFormatType.ZIP_PLUS_4;
    SpectialFormatType["PHONE"] = SpectialFormatType.PHONE;
    SpectialFormatType["SSN"] = SpectialFormatType.SSN;

    const SeparatorStyle = {
        COMMA_DOT: 0,       // 1,234.56
        NO_SEPARATOR: 1,    // 1234.56
        DOT_COMMA: 2,       // 1.234,56
        NO_SEPARATOR_COMMA: 3, // 1234,56
        APOSTROPHE_DOT: 4   // 1'234.56
    };

    SeparatorStyle["COMMA_DOT"] = SeparatorStyle.COMMA_DOT;
    SeparatorStyle["NO_SEPARATOR"] = SeparatorStyle.NO_SEPARATOR;
    SeparatorStyle["DOT_COMMA"] = SeparatorStyle.DOT_COMMA;
    SeparatorStyle["NO_SEPARATOR_COMMA"] = SeparatorStyle.NO_SEPARATOR_COMMA;
    SeparatorStyle["APOSTROPHE_DOT"] = SeparatorStyle.APOSTROPHE_DOT;

    const NegativeStyle = {
        BLACK_MINUS: 0,   // black minus
        RED_MINUS: 1,     // red minus
        PARENS_BLACK: 2,  // black parens
        PARENS_RED: 3     // red parens
    }

    NegativeStyle["BLACK_MINUS"] = NegativeStyle.BLACK_MINUS;
    NegativeStyle["RED_MINUS"] = NegativeStyle.RED_MINUS;
    NegativeStyle["PARENS_BLACK"] = NegativeStyle.PARENS_BLACK;
    NegativeStyle["PARENS_RED"] = NegativeStyle.PARENS_RED;

    const CalculateType = {
        SUM:        0,
        PRODUCT:    1,
        AVERAGE:    2,
        MIN:        3,
        MAX:        4
    }

    CalculateType["SUM"]        = CalculateType.SUM;
    CalculateType["PRODUCT"]    = CalculateType.PRODUCT;
    CalculateType["AVERAGE"]    = CalculateType.AVERAGE;
    CalculateType["MIN"]        = CalculateType.MIN;
    CalculateType["MAX"]        = CalculateType.MAX;

    const ValidateType = {
        NONE:   -1,
        NUMBER: 0,
        CUSTOM: 1
    }

    ValidateType["NONE"]      = ValidateType.NONE;
    ValidateType["NUMBER"]    = ValidateType.NUMBER;
    ValidateType["CUSTOM"]    = ValidateType.CUSTOM;

    Object.freeze(FIELD_TYPES);
    asc["FIELD_TYPES"]              = asc.FIELD_TYPES               = FIELD_TYPES;
    asc["BORDER_EFFECT_STYLES"]     = asc.BORDER_EFFECT_STYLES      = BORDER_EFFECT_STYLES;
    asc["REF_TO_REASON"]            = asc.REF_TO_REASON             = REF_TO_REASON;
    asc["FormatType"]               = asc.FormatType                = FormatType;
    asc["SpectialFormatType"]   = asc.SpectialFormatType    = SpectialFormatType;
    asc["SeparatorStyle"]           = asc.SeparatorStyle            = SeparatorStyle;
    asc["NegativeStyle"]            = asc.NegativeStyle             = NegativeStyle;
    asc["CalculateType"]            = asc.CalculateType             = CalculateType;
    asc["ValidateType"]             = asc.ValidateType              = ValidateType;
    asc.CommandType                 = CommandType;
})();
