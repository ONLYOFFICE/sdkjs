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

    ANNOTATIONS_TYPES["Underline"]   = ANNOTATIONS_TYPES.Underline;
    ANNOTATIONS_TYPES["Strikeout"]   = ANNOTATIONS_TYPES.Strikeout;
    ANNOTATIONS_TYPES["Highlight"]   = ANNOTATIONS_TYPES.Highlight;
    asc["ANNOTATIONS_TYPES"] = asc.ANNOTATIONS_TYPES = ANNOTATIONS_TYPES;

    let FIELD_TYPES = {
        unknown:        0,
        button:         1,
        radiobutton:    2,
        checkbox:       3,
        text:           4,
        combobox:       5,
        listbox:        6,
        signature:      7
    };

    FIELD_TYPES["unknown"]      = FIELD_TYPES.unknown;
    FIELD_TYPES["button"]       = FIELD_TYPES.button;
    FIELD_TYPES["radiobutton"]  = FIELD_TYPES.radiobutton;
    FIELD_TYPES["checkbox"]     = FIELD_TYPES.checkbox;
    FIELD_TYPES["text"]         = FIELD_TYPES.text;
    FIELD_TYPES["combobox"]     = FIELD_TYPES.combobox;
    FIELD_TYPES["listbox"]      = FIELD_TYPES.listbox;
    FIELD_TYPES["signature"]    = FIELD_TYPES.signature;

    Object.freeze(FIELD_TYPES);
    asc["FIELD_TYPES"] = asc.FIELD_TYPES = FIELD_TYPES;
})();
