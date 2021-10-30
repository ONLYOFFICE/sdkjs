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

(function(window, undefined){

var CAscColorScheme = AscCommon.CAscColorScheme;
var CColor = AscCommon.CColor;

var g_oAutoShapesGroups = [
    "Basic shapes",
    "Figured arrows",
    "Math",
    "Charts",
    "Stars & Ribbons",
    "Callouts",
    "Buttons",
    "Rectangles",
    "Lines"
];

var autoShapes = [
	[
		"textRect", "rect", "ellipse", "triangle", "rtTriangle", "parallelogram", "trapezoid", "diamond", "pentagon", "hexagon",  "heptagon",
		"octagon", "decagon", "dodecagon", "pie", "chord", "teardrop", "frame", "halfFrame", "corner", "diagStripe", "plus", "plaque",
		"can", "cube", "bevel", "donut", "noSmoking", "blockArc", "foldedCorner", "smileyFace", "heart", "lightningBolt", "sun",
		"moon", "cloud", "arc", "bracePair", "leftBracket", "rightBracket", "leftBrace", "rightBrace"
	],
	[
		"rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow", "quadArrow", "leftRightUpArrow", "bentArrow",
		"uturnArrow", "leftUpArrow", "bentUpArrow", "curvedRightArrow", "curvedLeftArrow", "curvedUpArrow", "curvedDownArrow", "stripedRightArrow",
		"notchedRightArrow", "homePlate", "chevron", "rightArrowCallout", "downArrowCallout", "leftArrowCallout", "upArrowCallout",
		"leftRightArrowCallout", "quadArrowCallout", "circularArrow"
	],
	[
        "mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual"
	],
	[
		"flowChartProcess", "flowChartAlternateProcess", "flowChartDecision", "flowChartInputOutput", "flowChartPredefinedProcess",
		"flowChartInternalStorage", "flowChartDocument", "flowChartMultidocument", "flowChartTerminator", "flowChartPreparation",
		"flowChartManualInput", "flowChartManualOperation", "flowChartConnector", "flowChartOffpageConnector", "flowChartPunchedCard",
		"flowChartPunchedTape", "flowChartSummingJunction", "flowChartOr", "flowChartCollate", "flowChartSort", "flowChartExtract",
		"flowChartMerge", "flowChartOnlineStorage", "flowChartDelay", "flowChartMagneticTape", "flowChartMagneticDisk", "flowChartMagneticDrum", "flowChartDisplay"
	],
	[
		"irregularSeal1", "irregularSeal2", "star4", "star5", "star6", "star7", "star8", "star10", "star12", "star16", "star24", "star32",
		"ribbon2", "ribbon", "ellipseRibbon2", "ellipseRibbon", "verticalScroll", "horizontalScroll", "wave", "doubleWave"
	],
	[
		"wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout", "borderCallout1", "borderCallout2", "borderCallout3",
		"accentCallout1", "accentCallout2", "accentCallout3", "callout1", "callout2", "callout3", "accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3"
	],
	[
		"actionButtonBackPrevious", "actionButtonForwardNext", "actionButtonBeginning", "actionButtonEnd", "actionButtonHome", "actionButtonInformation",
		"actionButtonReturn", "actionButtonMovie", "actionButtonDocument", "actionButtonSound", "actionButtonHelp", "actionButtonBlank"
	],
	[
		"rect", "roundRect", "snip1Rect", "snip2SameRect", "snip2DiagRect", "snipRoundRect", "round1Rect", "round2SameRect", "round2DiagRect"
	],
	[
		"line", "lineWithArrow", "lineWithTwoArrows", "bentConnector5", "bentConnector5WithArrow", "bentConnector5WithTwoArrows", "curvedConnector3",
		"curvedConnector3WithArrow", "curvedConnector3WithTwoArrows", "spline", "polyline1", "polyline2"
	]
];

var g_oAutoShapesTypes = [];

for (var i = 0, length = autoShapes.length; i < length; i++)
{
	g_oAutoShapesTypes[i] = [];
	for (var j = 0, length_group = autoShapes[i].length; j < length_group; j++)
	{
        g_oAutoShapesTypes[i].push({ "Type" : autoShapes[i][j] });
	}
}

var g_oStandartColors = [
    {R: 0xC0, G: 0x00, B: 0x00},
    {R: 0xFF, G: 0x00, B: 0x00},
    {R: 0xFF, G: 0xC0, B: 0x00},
    {R: 0xFF, G: 0xFF, B: 0x00},
    {R: 0x92, G: 0xD0, B: 0x50},
    {R: 0x00, G: 0xB0, B: 0x50},
    {R: 0x00, G: 0xB0, B: 0xF0},
    {R: 0x00, G: 0x70, B: 0xC0},
    {R: 0x00, G: 0x20, B: 0x60},
    {R: 0x70, G: 0x30, B: 0xA0}
];

var g_oThemeColorsDefaultModsWord = [
    [
        { name : "wordShade", val : 0xF2 },
        { name : "wordShade", val : 0xD9 },
        { name : "wordShade", val : 0xBF },
        { name : "wordShade", val : 0xA6 },
        { name : "wordShade", val : 0x80 }
    ],
    [
        { name : "wordShade", val : 0xE6 },
        { name : "wordShade", val : 0xBF },
        { name : "wordShade", val : 0x80 },
        { name : "wordShade", val : 0x40 },
        { name : "wordShade", val : 0x1A }
    ],
    [
        { name : "wordTint", val : 0x33 },
        { name : "wordTint", val : 0x66 },
        { name : "wordTint", val : 0x99 },
        { name : "wordShade", val : 0xBF },
        { name : "wordShade", val : 0x80 }
    ],
    [
        { name : "wordTint", val : 0x1A },
        { name : "wordTint", val : 0x40 },
        { name : "wordTint", val : 0x80 },
        { name : "wordTint", val : 0xBF },
        { name : "wordTint", val : 0xE6 }
    ],
    [
        { name : "wordTint", val : 0x80 },
        { name : "wordTint", val : 0xA6 },
        { name : "wordTint", val : 0xBF },
        { name : "wordTint", val : 0xD9 },
        { name : "wordTint", val : 0xF2 }
    ]
];

var g_oThemeColorsDefaultModsPowerPoint = [
    [
        { lumMod : 95000, lumOff : -1 },
        { lumMod : 85000, lumOff : -1 },
        { lumMod : 75000, lumOff : -1 },
        { lumMod : 65000, lumOff : -1 },
        { lumMod : 50000, lumOff : -1 }
    ],
    [
        { lumMod : 90000, lumOff : -1 },
        { lumMod : 75000, lumOff : -1 },
        { lumMod : 50000, lumOff : -1 },
        { lumMod : 25000, lumOff : -1 },
        { lumMod : 10000, lumOff : -1 }
    ],
    [
        { lumMod : 20000, lumOff : 80000 },
        { lumMod : 40000, lumOff : 60000 },
        { lumMod : 60000, lumOff : 40000 },
        { lumMod : 75000, lumOff : -1 },
        { lumMod : 50000, lumOff : -1 }
    ],
    [
        { lumMod : 10000, lumOff : 90000 },
        { lumMod : 25000, lumOff : 75000 },
        { lumMod : 50000, lumOff : 50000 },
        { lumMod : 75000, lumOff : 25000 },
        { lumMod : 90000, lumOff : 10000 }
    ],
    [
        { lumMod : 50000, lumOff : 50000 },
        { lumMod : 65000, lumOff : 35000 },
        { lumMod : 75000, lumOff : 25000 },
        { lumMod : 85000, lumOff : 15000 },
        { lumMod : 90000, lumOff : 5000 }
    ]
];

/* 0..4 */
function GetDefaultColorModsIndex(r, g, b)
{
    var L = (Math.max(r, Math.max(g, b)) + Math.min(r, Math.min(g, b))) / 2;
    L /= 255;
    if (L == 1)
        return 0;
    if (L >= 0.8)
        return 1;
    if (L >= 0.2)
        return 2;
    if (L > 0)
        return 3;
    return 4;
}

/* 0 - ppt, 1 - word, 2 - excel */
function GetDefaultMods(r, g, b, pos, editor_id)
{
    if (pos < 1 || pos > 5)
        return [];

    var index = GetDefaultColorModsIndex(r, g, b);
    var _obj, _mods = [], _mod;

    if (editor_id == 0)
    {
        _obj = g_oThemeColorsDefaultModsPowerPoint[index][pos - 1];

        if (_obj.lumMod !== -1)
        {
            _mod = new AscFormat.CColorMod();
            _mod["name"] = "lumMod";
            _mod["val"] = _obj.lumMod;
            _mods.push(_mod);
        }
        if (_obj.lumOff !== -1)
        {
            _mod = new AscFormat.CColorMod();
            _mod.name = "lumOff";
            _mod.val = _obj.lumOff;
            _mods.push(_mod);
        }

        return _mods;
    }
    if (editor_id == 1)
    {
        _obj = g_oThemeColorsDefaultModsWord[index][pos - 1];

        _mod = new AscFormat.CColorMod();
        _mod.name = _obj.name;
        _mod.val = _obj.val /** 100000 / 255) >> 0*/;
        _mods.push(_mod);

        return _mods;
    }
    // TODO: excel
    return [];
}

	var g_oUserColorScheme = [];
	var elem;
	elem = new CAscColorScheme();
	elem.name = 'New Office';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(0xFF, 0xFF, 0xFF));
	elem.putColor(new CColor(0x44, 0x54, 0x6A));
	elem.putColor(new CColor(0xE7, 0xE6, 0xE6));
	elem.putColor(new CColor(0x5B, 0x9B, 0xD5));
	elem.putColor(new CColor(0xED, 0x7D, 0x31));
	elem.putColor(new CColor(0xA5, 0xA5, 0xA5));
	elem.putColor(new CColor(0xFF, 0xC0, 0x00));
	elem.putColor(new CColor(0x44, 0x72, 0xC4));
	elem.putColor(new CColor(0x70, 0xAD, 0x47));
	elem.putColor(new CColor(0x05, 0x63, 0xC1));
	elem.putColor(new CColor(0x95, 0x4F, 0x72));
	g_oUserColorScheme.push(elem);

	elem = new CAscColorScheme();
	elem.name = 'Office';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(31, 73, 125));
	elem.putColor(new CColor(238, 236, 225));
	elem.putColor(new CColor(79, 129, 189));
	elem.putColor(new CColor(192, 80, 77));
	elem.putColor(new CColor(155, 187, 89));
	elem.putColor(new CColor(128, 100, 162));
	elem.putColor(new CColor(75, 172, 198));
	elem.putColor(new CColor(247, 150, 70));
	elem.putColor(new CColor(0, 0, 255));
	elem.putColor(new CColor(128, 0, 128));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Grayscale';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(248, 248, 248));
	elem.putColor(new CColor(221, 221, 221));
	elem.putColor(new CColor(178, 178, 178));
	elem.putColor(new CColor(150, 150, 150));
	elem.putColor(new CColor(128, 128, 128));
	elem.putColor(new CColor(95, 95, 95));
	elem.putColor(new CColor(77, 77, 77));
	elem.putColor(new CColor(95, 95, 95));
	elem.putColor(new CColor(145, 145, 145));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Apex';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(105, 103, 109));
	elem.putColor(new CColor(201, 194, 209));
	elem.putColor(new CColor(206, 185, 102));
	elem.putColor(new CColor(156, 176, 132));
	elem.putColor(new CColor(107, 177, 201));
	elem.putColor(new CColor(101, 133, 207));
	elem.putColor(new CColor(126, 107, 201));
	elem.putColor(new CColor(163, 121, 187));
	elem.putColor(new CColor(65, 0, 130));
	elem.putColor(new CColor(147, 41, 104));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Aspect';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(50, 50, 50));
	elem.putColor(new CColor(227, 222, 209));
	elem.putColor(new CColor(240, 127, 9));
	elem.putColor(new CColor(159, 41, 54));
	elem.putColor(new CColor(27, 88, 124));
	elem.putColor(new CColor(78, 133, 66));
	elem.putColor(new CColor(96, 72, 120));
	elem.putColor(new CColor(193, 152, 89));
	elem.putColor(new CColor(107, 159, 37));
	elem.putColor(new CColor(178, 107, 2));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Civic';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(100, 107, 134));
	elem.putColor(new CColor(197, 209, 215));
	elem.putColor(new CColor(209, 99, 73));
	elem.putColor(new CColor(204, 180, 0));
	elem.putColor(new CColor(140, 173, 174));
	elem.putColor(new CColor(140, 123, 112));
	elem.putColor(new CColor(143, 176, 140));
	elem.putColor(new CColor(209, 144, 73));
	elem.putColor(new CColor(0, 163, 214));
	elem.putColor(new CColor(105, 79, 7));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Concourse';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(70, 70, 70));
	elem.putColor(new CColor(222, 245, 250));
	elem.putColor(new CColor(45, 162, 191));
	elem.putColor(new CColor(218, 31, 40));
	elem.putColor(new CColor(235, 100, 27));
	elem.putColor(new CColor(57, 99, 157));
	elem.putColor(new CColor(71, 75, 120));
	elem.putColor(new CColor(125, 60, 74));
	elem.putColor(new CColor(255, 129, 25));
	elem.putColor(new CColor(68, 185, 232));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Equity';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(105, 100, 100));
	elem.putColor(new CColor(233, 229, 220));
	elem.putColor(new CColor(211, 72, 23));
	elem.putColor(new CColor(155, 45, 31));
	elem.putColor(new CColor(162, 142, 106));
	elem.putColor(new CColor(149, 98, 81));
	elem.putColor(new CColor(145, 132, 133));
	elem.putColor(new CColor(133, 93, 93));
	elem.putColor(new CColor(204, 153, 0));
	elem.putColor(new CColor(150, 169, 169));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Flow';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(4, 97, 123));
	elem.putColor(new CColor(219, 245, 249));
	elem.putColor(new CColor(15, 111, 198));
	elem.putColor(new CColor(0, 157, 217));
	elem.putColor(new CColor(11, 208, 217));
	elem.putColor(new CColor(16, 207, 155));
	elem.putColor(new CColor(124, 202, 98));
	elem.putColor(new CColor(165, 194, 73));
	elem.putColor(new CColor(244, 145, 0));
	elem.putColor(new CColor(133, 223, 208));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Foundry';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(103, 106, 85));
	elem.putColor(new CColor(234, 235, 222));
	elem.putColor(new CColor(114, 163, 118));
	elem.putColor(new CColor(176, 204, 176));
	elem.putColor(new CColor(168, 205, 215));
	elem.putColor(new CColor(192, 190, 175));
	elem.putColor(new CColor(206, 197, 151));
	elem.putColor(new CColor(232, 183, 183));
	elem.putColor(new CColor(219, 83, 83));
	elem.putColor(new CColor(144, 54, 56));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Median';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(119, 95, 85));
	elem.putColor(new CColor(235, 221, 195));
	elem.putColor(new CColor(148, 182, 210));
	elem.putColor(new CColor(221, 128, 71));
	elem.putColor(new CColor(165, 171, 129));
	elem.putColor(new CColor(216, 178, 92));
	elem.putColor(new CColor(123, 167, 157));
	elem.putColor(new CColor(150, 140, 140));
	elem.putColor(new CColor(247, 182, 21));
	elem.putColor(new CColor(112, 68, 4));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Metro';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(78, 91, 111));
	elem.putColor(new CColor(214, 236, 255));
	elem.putColor(new CColor(127, 209, 59));
	elem.putColor(new CColor(234, 21, 122));
	elem.putColor(new CColor(254, 184, 10));
	elem.putColor(new CColor(0, 173, 220));
	elem.putColor(new CColor(115, 138, 200));
	elem.putColor(new CColor(26, 179, 159));
	elem.putColor(new CColor(235, 136, 3));
	elem.putColor(new CColor(95, 119, 145));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Module';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(90, 99, 120));
	elem.putColor(new CColor(212, 212, 214));
	elem.putColor(new CColor(240, 173, 0));
	elem.putColor(new CColor(96, 181, 204));
	elem.putColor(new CColor(230, 108, 125));
	elem.putColor(new CColor(107, 183, 109));
	elem.putColor(new CColor(232, 134, 81));
	elem.putColor(new CColor(198, 72, 71));
	elem.putColor(new CColor(22, 139, 186));
	elem.putColor(new CColor(104, 0, 0));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Opulent';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(177, 63, 154));
	elem.putColor(new CColor(244, 231, 237));
	elem.putColor(new CColor(184, 61, 104));
	elem.putColor(new CColor(172, 102, 187));
	elem.putColor(new CColor(222, 108, 54));
	elem.putColor(new CColor(249, 182, 57));
	elem.putColor(new CColor(207, 109, 164));
	elem.putColor(new CColor(250, 141, 61));
	elem.putColor(new CColor(255, 222, 102));
	elem.putColor(new CColor(212, 144, 197));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Oriel';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(87, 95, 109));
	elem.putColor(new CColor(255, 243, 157));
	elem.putColor(new CColor(254, 134, 55));
	elem.putColor(new CColor(117, 152, 217));
	elem.putColor(new CColor(179, 44, 22));
	elem.putColor(new CColor(245, 205, 45));
	elem.putColor(new CColor(174, 186, 213));
	elem.putColor(new CColor(119, 124, 132));
	elem.putColor(new CColor(210, 97, 28));
	elem.putColor(new CColor(59, 67, 91));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Origin';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(70, 70, 83));
	elem.putColor(new CColor(221, 233, 236));
	elem.putColor(new CColor(114, 124, 163));
	elem.putColor(new CColor(159, 184, 205));
	elem.putColor(new CColor(210, 218, 122));
	elem.putColor(new CColor(250, 218, 122));
	elem.putColor(new CColor(184, 132, 114));
	elem.putColor(new CColor(142, 115, 106));
	elem.putColor(new CColor(178, 146, 202));
	elem.putColor(new CColor(107, 86, 128));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Paper';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(68, 77, 38));
	elem.putColor(new CColor(254, 250, 201));
	elem.putColor(new CColor(165, 181, 146));
	elem.putColor(new CColor(243, 164, 71));
	elem.putColor(new CColor(231, 188, 41));
	elem.putColor(new CColor(208, 146, 167));
	elem.putColor(new CColor(156, 133, 192));
	elem.putColor(new CColor(128, 158, 194));
	elem.putColor(new CColor(142, 88, 182));
	elem.putColor(new CColor(127, 111, 111));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Solstice';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(79, 39, 28));
	elem.putColor(new CColor(231, 222, 201));
	elem.putColor(new CColor(56, 145, 167));
	elem.putColor(new CColor(254, 184, 10));
	elem.putColor(new CColor(195, 45, 46));
	elem.putColor(new CColor(132, 170, 51));
	elem.putColor(new CColor(150, 67, 5));
	elem.putColor(new CColor(71, 90, 141));
	elem.putColor(new CColor(141, 199, 101));
	elem.putColor(new CColor(170, 138, 20));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Technic';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(59, 59, 59));
	elem.putColor(new CColor(212, 210, 208));
	elem.putColor(new CColor(110, 160, 176));
	elem.putColor(new CColor(204, 175, 10));
	elem.putColor(new CColor(141, 137, 164));
	elem.putColor(new CColor(116, 133, 96));
	elem.putColor(new CColor(158, 146, 115));
	elem.putColor(new CColor(126, 132, 141));
	elem.putColor(new CColor(0, 200, 195));
	elem.putColor(new CColor(161, 22, 224));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Trek';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(78, 59, 48));
	elem.putColor(new CColor(251, 238, 201));
	elem.putColor(new CColor(240, 162, 46));
	elem.putColor(new CColor(165, 100, 78));
	elem.putColor(new CColor(181, 139, 128));
	elem.putColor(new CColor(195, 152, 109));
	elem.putColor(new CColor(161, 149, 116));
	elem.putColor(new CColor(193, 117, 41));
	elem.putColor(new CColor(173, 31, 31));
	elem.putColor(new CColor(255, 196, 47));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Urban';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(66, 68, 86));
	elem.putColor(new CColor(222, 222, 222));
	elem.putColor(new CColor(83, 84, 138));
	elem.putColor(new CColor(67, 128, 134));
	elem.putColor(new CColor(160, 77, 163));
	elem.putColor(new CColor(196, 101, 45));
	elem.putColor(new CColor(139, 93, 61));
	elem.putColor(new CColor(92, 146, 181));
	elem.putColor(new CColor(103, 175, 189));
	elem.putColor(new CColor(194, 168, 116));
	g_oUserColorScheme.push(elem);
	elem = new CAscColorScheme();
	elem.name = 'Verve';
	elem.putColor(new CColor(0, 0, 0));
	elem.putColor(new CColor(255, 255, 255));
	elem.putColor(new CColor(102, 102, 102));
	elem.putColor(new CColor(210, 210, 210));
	elem.putColor(new CColor(255, 56, 140));
	elem.putColor(new CColor(228, 0, 89));
	elem.putColor(new CColor(156, 0, 127));
	elem.putColor(new CColor(104, 0, 127));
	elem.putColor(new CColor(0, 91, 211));
	elem.putColor(new CColor(0, 52, 158));
	elem.putColor(new CColor(23, 187, 253));
	elem.putColor(new CColor(255, 121, 194));
	g_oUserColorScheme.push(elem);

var g_oUserTexturePresets = [
    "data:image/png;base64,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",
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAIAAACRXR/mAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAClNJREFUeNpkmdmO3DYQRbVQUm9jGEbgR7/5//8myS84BuzJTHerN0m5xcO+Fhw9CGyKLNZya2PXp39/LMvSNM3tdtN7nufz+bzZbF5eXk6nU9/3mhy222qev3371nXddrdJKV3z07atvtZ1rWXj6fz29rbf7w+Hw/F4FB0NqqoathsRvFwuWqmNWqyBNm764efPn4/HQ5M6Syu1XXQ00DvWaZuoaOl2u9UicanBIz8aa4GW8oaPOT/6pLfmNRjHUVINwyC+oatJDaZpen19Fc3dbmc6DCRVnx+t0UGa0VtjPiWN0BYU9Q3pJbEGoqiz7/e7xlogLVbVop+aRE9aACGxlb9WUoze+soxWqk1IqgBZ+ktCvdrrMdKWimB+SQd6x1speeDAkRUu6rnY/n0jsVdC4tajJSoUOYTBY052+OuWuCVee3VEdooRqVdrKRJbOWNCaERRWP0Lyr6LDk0YA/S6Oeh28OHSHMAcBFgoKBHUjEOW9cV+kCFKFh6Sk1wo5VgDmBgd2aSFa79OhKFawVmZV5rREszgFcUYYuxPgnysIhSUbw2pr5DJQimRwPp6XIecS8dpxnAyqNDE7CFelj8etXZ6B89wXT4jkQchuvtAk8ADhNItnDSbAU8FJ8CwkiOc+gnawwPe6h4xXtEOSwFjwIHDgiK9Y1POm/IT5h1HOXMIqHYgWrxEgzU73b3MR74QDFztZi/Lj8aS+X77Q7sogvwClJDI49LHCBl6oPY0gd8Z7zEjDaIuliRqowYEQWC+on+NRi6Xgve39/F5YcPHySGfmrc9RE1AAn2wg5NeFHADp+D12A9BVSKETVCpfAemg9LBgjwZ/abdRTMRhBGHNFewItrR4BYKvwAvBoY06M4O1+JSgWyS51YDUQwP4Dd7beY3OAFJbg6NrIXh8MHujqFTauQaF4vvw5GsHV8Xr8BjwSJMci1TxGQkMkclwMyPIsJsvR6g63wuCZ8AtRP+SkL6gakgzAmw7ip5SDEthtyXCJE8ZtQBvTMK2ETDemrPANeCTmcp41aQxCRQNPTXeK8qvj8OvlohlTBdiKLgSGMpHVAB2cEpOvtV9LgjPDGzWbJJiORsZiQ3dahA7mh9Q3a2tTZFMQCUD+eT+t4ZvuUJEGUIg6xCKxcruUA/cTnwzqrnIo1WSwlDf0AESRkYySffsB2NhbyMzY3K7fQfJVAA0KDVmTaNluWkqdwk8jtmT/SlGakyDUQ7QSgAplhAkngiZiHw6ItC8PPpEiD28uJqKI4r2Crbje77dCHGsbzRawPmx4wQUuGxShTNT3mKUMhG6sNH9dMfQtJ0tDbOe63q87vNhHTpYnb5Q5BYXu/OUz3OCu0RfZwacUAPVlKw9POaIfHVwAKC4A2KXl+TCYLZeKOs9b0fOR8Yno75HDqkoZoy/GEEBvFtLKTLmaRHCCFEYfX9Qxchmz1DE94iWFOnq3/9+DRic/IR60I9bWHogCWKVUDEYxC/UiV4byGjoFUvRT4e9KmQNmuclHQ+XgKvZCbHVSwkam4+nP0b9rGD+JCF73CE9U2T1OVMsG6L3vrEsnsB+wik5ZUb+k5A17xOBfvJdl1vTsF8M5hpB1kWKtkmWZnw3WLcb9ebDgfQdAObZGMkcyVIfByODGKy0nPGG2wr5OxKaCDNu91lDL8jQQmiQtBKhUoJ5BIUBa8qCZQJhsgUaqr6Y4+IlQ+keRC2VEbggLcYbe3L685U9WKEzgPPsWeoyGjbKUDwXyKZMHBstCEaEFE1xx1KV00ryx0GUd1dppRFAi+o2qvlII0UOyJtkx+3rTd0N9Opz///ktW/vr16y+oPCaFAZlYcSx2ycGXiK4iVFK16y0CRmkNsnBIGbko1zNxfF2pNEVV4g9Nh8R1gwwomBIvep63CNGCitgCc5HO6/r93ze0q5pRJ2oNwCiBsHs+KN85UdqKiJzDhCNIWK2J7oqC/VkhxaMeygGWIEkKr5dZmv7y5YvYWjJNsogYjXIo20EDSniaADESkIe0o9k9P0OuHajl6RcYL/UC8AET7Lp04UgDLhyiWmBCk9+/f5dI+Mrjdicaiya1PD09hVb4EVWyPhO0YBEPdZP4C9FdS2Xh8poe3wGaCG5f0VuWRXJZv30+l8foLsjdbHYdEWkLZ8SIdXQhmdjPnbnmakbbBAXwG4xOJR6WwJPtK9PsDntr+vPnz1SwoqBy0nUUkBJxsVi20wA6Yfkm4ng8ur9b5wAleKpC9qMqQqFI3fLDeYjET6TFxDKT2Ora5DIEkFDnRD0rtiQQJSicshMusQiyuuolnNK/iyGuHmQd5T7o4NpgK9qy6QGwqGocHW0BUYOtctUz5wpvnh5N3Y/Xy+16+fjxY98Fo+qWPv3xiVbsei+9Zd9FeOurTn26+pwh/xRy9TO8NYsekfAZ9/HoaODmxdWvLKPGNa4zbtfb445XPS4T0aBum9N4jDiFW2lz9PIp4agAUFC1ZVGn3rtNqc/c/GALTtUu6Y8aiUoEwIE2IEuK+y3uu5EpBRKAIAhF5T+OkNNApqGBxppuQKyVKCOfPFF2g3SKJcDnNgv/xY6ESZcVvqFYcmoJOs59bqz18/X1Vfr0LY/rSa05vR85lX7fly109OQD4qRDhjHq+jHQ3XeGLPy5YoiOBoVbRFmBtLh/OWzy456M2IaS3aO6QiRqG7ncCTjy4ZIuy367NnJfVKAmbSknQFGs4KL4F1ksivHsO1zLWOdA2FHevR4A4IYHlyQDYhCIr4Mi3v3bZUSi9Hb96fAY6TO1AJZLEcj5ZoELIF8j4hOuUqSt6JRyK0vQR69OJNR8hDfXcL6nCNkUNqmWuNgQLUkpE+y3B8xM78p1Lc6IE1hJTGoLnHGSvFjz0rrTMKWsFpRioSrZdt0nskD1UrmkRFxCAy7GpTxqN0rocwjleJkkCT9QqbPMipzRT9TV+TIqJg3bTfSJS0E6MBBNbaEfodJXcfasKBe18lWTjUi94QLSdbf7xHX+CQFS5x6aPlsDLuUQd90AalLVqZscl8vwgec2qycskJIgVBov5wEz4TsSly7rex8yDzFPY0KoK3q7Vb42GSBLeHOiBLiIug5ypUUAub749wMVPAtoO8/jBECbIMkFJKEVYRwXuHr0bcz6On19n+vL6a6NXjWw7DbfISQGVdENOnCd2DatLYWjiKHw/+cB7gRLhJyK8wIS3+q654GO7/e4dE3rSzpbMHxzevge0D6MFvESQgD1E38cuRdyKwtW3Ak6IhDDHWbJp8QgeU+katyV7OHL0lzuLb/dpj4dqpztSx9RAJrGg0NlWD/fha7bfy0jBvni05FC88fz6eXlJVnhvhsu0OmSge+MS5fsv9f40yYKoWkS6tf9PkaJlH8622msNlKntoNIJz1qvmD3/cc//v8NIPs6nijqf7DIiRpIz/rK/1t2OmnXRem6Ymmb5P+2pCF6r8iPuTx010osDKma7AdrhbsqWqd6JsG1b37t5+CJDgc1ODuVWDgtLodQTKlZqpLvwRYbo0nOlP8TYAD8wdiTL7IMZAAAAABJRU5ErkJggg==",
    "data:image/png;base64,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",
    "data:image/png;base64,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",
    "data:image/png;base64,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",
    "data:image/png;base64,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",
    "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADIAAAAyCAIAAACRXR/mAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAADYFJREFUeNo82VernNcVxnGdGXXJ6r33XgzG+jKxY4eQkIvUixACIeTDBXIhJIzKUT3qVpesZhWr5ffOH85cDO/sd5VnPWvttctMzVyaXrFixd27d5cuXfrp06enT58uWbJk48aNFy5cOHTo0OvXr1+8eGHk5cuXHz9+3LRp0/3793/55Zc9e/Zcv36d/BdffPHgwYPly5evX79+ZmZm586dT548GY1GP//8s+/xeLxo0aJ58+axs2rVqnfv3t27d4+wt58/f2bkypUr69atmz9/vnEA3rx5A8C2bdumzv1w6u3bt3RIG4XAi0ePHq1evZrXhQsXkmaOlTlz5sDkYfv27aBMTU0xB66fIMI6d+5c8q9evdq6dSsjnlnjD3RvgWNnw4YNdLds2fL48WNYhcr11atXQXz//j0LYGBk5AdlIOhAIz7PK1euZMiIn9zz5Cd8xhnCh8E1a9YsXrxYJCzs2LFj2bJl5EVIHf0fPnwwTgaL3PspKph8J+aZup8//fSTwNauXcvm8+fPgUH5SEyU/RYf99euXUOAZJEGnJrgjD979oyMyHBe0DwZQR4VCDCPXZigRyoogrx58yZJwUglL8gwQljuGNy8eTO+cSkqpqgbPH36tMEptcUNx8LFhDjib9a32uJbCjBn0CvWwWXCWwVkxLMYBMYTU0rHIDvEUAKTUCH+8ccfybPGFDGphEYd37lzhzBJg2zyNdSWROCQAgSE+CDHSpluHiyffBDTDOCPCn12xcorAshzKYPoER5FgZkHCgCOBQsWwMe3apM1AphjxyuS/IJhcN++fTdu3BhhSPoZglfivPAsjqZSxHjFMWXjkl5+sUvAKynzU96rG+OyoKJv377NmuRyHEPy+GzyURi80KIusL1790ql2iLWnB0pBfGhx29uzpw5c+vWLUz4SdO34PhTKCYITWmihktiJiAVzcUrNcexoAl7lYNLly4BKmzBU0EDhoQnEo6Ywt+RI0eYlQekYvTixYtSOXXv9g1DwALx8OFDJFenIkYMfeZ4IiCVeCIACn187N69Gx8YrWhE6LnwqIu2LgBH1EoThkDkS0iEK/kAEfs8+QzWODbEATUupemrr76iD4HXoteEwBLBwYMH620IUGHHjh1jl5jUaLyevZIRMpoQ62QUSs1CSbCgCuUBHwwePny4Bk4MW4B6hg8LxEZhCgRlBcHKrl27SItDvlD4ZvJRvBQA5dKDoGn55sZbzQYyrHBgXLRqCPEsFJiAZRlJEqe2qDCFbJhgFaryFZWmyMKoaam2ZrtDFVMV167EYdIpFOYIYJ5paOADXRigs8sULfKqralHXhI5RoySMM43lAJQjvxCgB5hNCuJyfvly5enbly9hKehsU6KVzoUozJXpzXxAwcOCL1Sm56eJgYucDU2ugS491DDFJsHikCjgU2gTQj4GFTmlhpTDxTTyAjLgqdCni4Kz58/PzIEssgYIiF3dKA0LhopaDWUX6alGw3M4c/Mb1Vha9Pk44Ewu1JZRvAEOiMZRwwBzygxrnuR4Ygpr1IhoIiH2pL1J5OPikMgNJDhk4M2F/x5JWu1tLofTKw0pcmgQaIBKnecodMcxx9h+CpKyZJ3/KlUcHmhIlQ9hcGsKaepKxfOgU+t3QsdGaRgBMOagm9yigbu5jD3BrlUJXEJVg1WomUzgdojm7wa57J9hCChgTsKpY8KAYRhMRgjkQl62eTDkExx9nzykSyARNmCiPMFkw9btMxw6PkwK1UkSeFyxvr+/fuJUVEMHtrAcOaZfc9qTuKkRZy+5UTixCA2PQWy8XfffsMosE1AtcJZXZ5QRabk1RyjIHowIpVIEiWelF31BCKyRe9BUto/Ae1BhEyRQT8XJj4XVYJJDRYVRPiJUWU6/s+//+WFlCsF2L3GJASU+fCT75pvFe2D1DalIIpE9GpF2QEqUHaNs4ZUbmgZZ82rQsK9kbaHPOZLhbRlJUBy2NiQRjj4+CSq6jnDGU+ibB6phoJu2wMKJtqxaCW2D/iuhnxjpf1J65IREAFtTlDRrs+dG2qafD2Fl1Q4HXYQYHIjtfVS+pYO3+1IlQJKKVQQZrX4qmJTr3kg9C+//FJJvZh82gJ5kJRy6kFOwWrXwAIWrV1BrGw4apMHmZ/jP/z+dxyjoSpBo0kEgVQyIZr2kwx5y7cRAagtbMOkVgTNFs48qBLWjaseUQEkkiYTlXZBzdzaWFOEIzllquwPy9Q///F3OqSF4gGlXnso6wXtVacX4xUKtpj2s3OE2jLiFRU0N/MrTdY8NzGJMTJv8gHIeC2jCNsr+LYATJ3833+9M1/w1DnCKCH4RpOPtQUCPpiWC548zB6WdEUbpsYluqUQQ63BrFlqGMQig75poV/NoMe2oFoUhu4jKgA6CI2///U3UoZJvskJXdFAY6Ihv57uufMW2lDVesUQf8aZo9IqVGdCg7Il1hJiXGpg0hHUBv5aT/GktqjrHQIwbypH1sZ/+8ufjALY5H86+eh13NAUh37WccqcQg8OmCNw9uzZduKo6hDQRpQAPiLepOZGqFRYa9dg/jafvGK/hZ+RtpBGuJu6fuViJyHREMVTi493Mj27HUVAm0E4+CNZU63ezRjpAIhLsSGphLYqC6aZRKXeYXrWqEDxrOZKRZXAy3AqZ1RGhtPZZEHV9JFJuWVBjwGI733798t1/ZYDKwxnxKjYWbT4cOABMdx4ZSWhwsXLyYcK3+FuL14vFJIVDzsCsEvwavzb33wPU8T4kDNaTISUnbAUivE3r1+3KIFrQWRdLdbi8ScXHmY3tM3N1py2AqitSGCSZSoqofrrdKmClQHhYaNgG0hCQJ0HPXOMABBr4uip/5avjhJYZNo3QxATMJdl30i+a0umUYd9sASp1KQMjm4o6pe1CY46ZopzOL7evXW9W5p2iQApfxL1MEGIqfMxAfvGEi3L7ILCccdi6FtJOvK3OknCqVOn5k4+3hqUGSjNVhOI8XY7KGhPxpEH69LQCLCCt1eTD9QKojnCQcs+r8Iipt+UAvLaFXAtzMZ5qgl7hTOhatHQHD16VI460lmjpEXwnanarBohUK49g/7111+P//rnP+JZKXQwamVVByBSE66pS5Rdamq5iW0eaKR815RrVx68AoW1tr8mh4yTEW03Ds4pMi71HWgVrhg8dFKiC9kw+W7OXMZ/5/faP25Yl2AWYZJfHHQn1haqJZZXz90W6W3UyQgMrzxJWXcK3ShhulONeoLMzAVR/F6RBKV2GCzuhh0EYsTanRaLphvmWPSNHjkFlw7a25oKC6w1k08NSRa6rOsarb0A4c4ELcBQ8iLgdgB0jXesaD9icHYzMhQjQ7yaRyQgSIKOV5jnrwM7N2oc28pRErvhAQu7ug4KlSr31LvarGvXn5pVHHUZiQu04cyBoCnZnRstQQ5Lxf07N+mTJtq1iRf0y7oQNad6YCsMT575VoUdWWvNNQgcsNCxB278UUESWIyw+XHyaU8LFhXxMOihmgNmgHH14vnhADS5COUYcBJSPnsKhV2+u8oiIPRuIiqgEg06KCBCMFTGaES96uwkjfXmDRDqtdvX9gvGGel2iEGYhuNrazhPcsQ0Iaa7P2nft37DBoAANYeRlFjNCY7is6GFnnVzkEA3cmRUevcLXLJApv7nlTLwMHsi6uaRJGSWkKGW4eDYCuNZL/VarSycfGhOnz9PmoCqaoNPwCbMRNG1OaPCK7tIEiuGWnwMIobLEydOIIYi3G0elRSDFl+WqSBPqEzJskk9NWxCZi53Pun2mxUZAYuCsui6poKbvYnodAocZ12SUwFU7Yuw69bx5MOyXNMyCBYLkqhNGEQbqppJ3VMqnu5phnP2d9/+ymtlhO22Ax5EOdt1FGZHD2XHH2GI5VdGkNfNpbf1G1ShQfRkFEaHxNZyoHllhIAMCsmEmDP5dJongxf1MJzzuBcHu0B0UgXfOogDSYFSKP2FQUZ5tSsyrqTo40PxwiEwCIxgupLnPuOi96wuJbojnVlctxMMC7CC5RmRFIftMdpJGKLJEwTe1cm6DrUCluVuIrlhHaPk0W4cGRq3QJkGvV1h3di4B69anq0lhDujstbWnkFe0FHw1jRMD7XVRqz7BRC9RoCRfCusmntktEDNjgNhSp48eZL14Xw3HsMKd9cv/YQDYV0C4hUC2fDtbbuMTkE8ctHhbLg7bccnx50RBN01fzUhX/0noJj8FJDZ1CUnrf4OchZt5y1Ckl2YkVF2VLr9BrTLFZmFRj0IzGDLuSrszwfFwPKofWNnLEJcemcdZJRcK1prQAceueCGUcEZRMD09LQYqLQt6zZFVejsrUstR3SJsS8AAcsJUlHItWzW88Ss3Qx3EHancIAsIBnpD5LOM+ZL86gzY/f9/HW6gqk1tCsQqzL0IuQAZ+LpzhyULqG96najs11cMiKhXXp3T+EzbJrP/XCKIVC86w8FK27/yHUPdvz48dygh12htPH3s1bELs7qRpjAcYwyIhJd2qBXVVt/TejjbLZPRBgoAqbYFRdGhn/I+hOr23bWNd8Ox1VD/+A1TYAmT0ZqvOKvhaVrEuptHoXbVTaXcioJoFBEz+z/MZqCYu1g18HEuGdFJc7/CzAAb2tiTmxUvU0AAAAASUVORK5CYII=",
    "data:image/png;base64,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",
    "data:image/png;base64,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",
    "data:image/png;base64,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",
    "data:image/png;base64,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"
];

var g_sWordPlaceholderImage = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAIAAAAiOjnJAAAAAXNSR0IArs4c6QAAAAlwSFlzAAAOxAAADsQBlSsOGwAAA5dJREFUeF7t0rEVgkAAREG4BgjpvzvJsAA9DSiBHzFbwAb/zTrnPN7fxRS4r8C+jfV1fu479KTAVWAooUBRAKyiqs8FLAiSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSAKwkq1OwGEgKgJVkdQoWA0kBsJKsTsFiICkAVpLVKVgMJAXASrI6BYuBpABYSVanYDGQFAAryeoULAaSAmAlWZ2CxUBSYOwbW0nZJ5/+Uf0Ahm8Ksdfm760AAAAASUVORK5CYII=";




	var g_aTextArtPresets = [
		"textNoShape",
		"textPlain",
		"textStop",
		"textTriangle",
		"textTriangleInverted",
		"textChevron",
		"textChevronInverted",
		"textRingInside",
		"textRingOutside",
		"textArchUpPour",
		"textArchDownPour",
		"textCirclePour",
		"textButtonPour",
		"textCurveUp",
		"textCurveDown",
		"textCanUp",
		"textCanDown",
		"textWave1",
		"textWave2",
		"textDoubleWave1",
		"textWave4",
		"textInflate",
		"textDeflate",
		"textInflateBottom",
		"textDeflateBottom",
		"textInflateTop",
		"textDeflateTop",
		"textDeflateInflate",
		"textDeflateInflateDeflate",
		"textFadeRight",
		"textFadeLeft",
		"textFadeUp",
		"textFadeDown",
		"textSlantUp",
		"textCascadeUp",
		"textCascadeDown",
		"textArchUp",
		"textArchDown",
		"textCircle",
		"textButton"
	];



	function fSaveStream(oStream, nLength) {
		/*var aData = oStream.data.slice(oStream.cur, oStream.cur + nLength);
		var sData = "XLSY;;";
		sData += (nLength + ";");
		sData += AscCommon.Base64.encode(aData);
		var nCRC32 = AscCommon.g_oCRC32.Calculate_ByString(sData, sData.length);
		return {data: sData, crc32: nCRC32};*/
		return undefined;
	}

    //----------------------------------------------------------export----------------------------------------------------


  window['AscCommon'] = window['AscCommon'] || {};
  window['AscCommon'].g_oAutoShapesGroups = g_oAutoShapesGroups;
  window['AscCommon'].g_oAutoShapesTypes = g_oAutoShapesTypes;
  window['AscCommon'].g_oStandartColors = g_oStandartColors;
  window['AscCommon'].GetDefaultColorModsIndex = GetDefaultColorModsIndex;
  window['AscCommon'].GetDefaultMods = GetDefaultMods;
  window['AscCommon'].g_oUserColorScheme = g_oUserColorScheme;
  window['AscCommon'].g_oUserTexturePresets = g_oUserTexturePresets;
  window['AscCommon'].g_sWordPlaceholderImage = g_sWordPlaceholderImage;
  window['AscCommon'].g_aTextArtPresets = g_aTextArtPresets;
  window['AscCommon'].sChartStyles = "";




	window['AscCommon'].g_oChartStyles = window['AscCommon']['g_oChartStyles'] || {};
	window['AscCommon']['g_oChartStyles'] = window['AscCommon'].g_oChartStyles;

	window['AscCommon'].g_oStylesBinaries = window['AscCommon']['g_oStylesBinaries'] || {};
	window['AscCommon']['g_oStylesBinaries'] = window['AscCommon'].g_oStylesBinaries;

	window['AscCommon'].g_oColorsBinaries = window['AscCommon']['g_oColorsBinaries'] || {};
	window['AscCommon']['g_oColorsBinaries'] = window['AscCommon'].g_oColorsBinaries;

	window['AscCommon'].g_oDataLabelsBinaries = window['AscCommon']['g_oDataLabelsBinaries'] || {};
	window['AscCommon']['g_oDataLabelsBinaries'] = window['AscCommon'].g_oDataLabelsBinaries;

	window['AscCommon'].g_oLegendBinaries = window['AscCommon']['g_oLegendBinaries'] || {};
	window['AscCommon']['g_oLegendBinaries'] = window['AscCommon'].g_oLegendBinaries;

	window['AscCommon'].g_oCatBinaries = window['AscCommon']['g_oCatBinaries'] || {};
	window['AscCommon']['g_oCatBinaries'] = window['AscCommon'].g_oCatBinaries;

	window['AscCommon'].g_oValBinaries = window['AscCommon']['g_oValBinaries'] || {};
	window['AscCommon']['g_oValBinaries'] = window['AscCommon'].g_oValBinaries;

	window['AscCommon'].g_oBarParams = window['AscCommon']['g_oBarParams'] || {};
	window['AscCommon']['g_oBarParams'] = window['AscCommon'].g_oBarParams;

	window['AscCommon'].g_oView3dBinaries = window['AscCommon']['g_oView3dBinaries'] || {};
	window['AscCommon']['g_oView3dBinaries'] = window['AscCommon'].g_oView3dBinaries;

	window['AscCommon'].g_oChartStylesIdMap = window['AscCommon']['g_oChartStylesIdMap'] || {};
	window['AscCommon']['g_oChartStylesIdMap'] = window['AscCommon'].g_oChartStylesIdMap;

	function fGetAttributeString(attribute) {
		if(Array.isArray(attribute)) {
			var sResult = "[";
			for(var nIdx = 0; nIdx < attribute.length; ++nIdx) {
				sResult += fGetAttributeString(attribute[nIdx]);
				if(nIdx < attribute.length - 1) {
					sResult += ", ";
				}
			}
			sResult += "]";
			return sResult;
		}
		var nParsedNum = parseInt(attribute);
		if(!isNaN(nParsedNum) && (!(typeof attribute === "string") || attribute === ("" + nParsedNum))) {
			return "" + attribute;
		}
		if(typeof attribute === "string") {
			return "\"" + attribute + "\"";
		}
		return "" + attribute;
	}
	function fGenerateScriptMap(sObjectName, oMap, bKeysAsIs) {
		var sMapScript = sObjectName + " = {};\n";
		var oKeysMap = {};
		for(var key in oMap) {
			if(oMap.hasOwnProperty(key)) {
				var sValue = fGetAttributeString(oMap[key]);
				if(!Array.isArray(oKeysMap[sValue])) { 
					oKeysMap[sValue] = [];
				}
				oKeysMap[sValue].push(key);
			}
		}
		for(var sValues in oKeysMap) {
			if(oKeysMap.hasOwnProperty(sValues)) {
				var aKeys = oKeysMap[sValues];
				for(var nKey = 0; nKey < aKeys.length; ++nKey) {
					var sKey = aKeys[nKey];
					if(!bKeysAsIs) {
						sMapScript += (sObjectName + "[" + fGetAttributeString(sKey) + "] = ");
					}
					else {
						sMapScript += (sObjectName + "[" + sKey + "] = ");
					}
				}
				sMapScript += (sValues + ";\n");
			}
		}
		sMapScript += "\n";
		return sMapScript;
	}
	function fGenerateStyles() {
		var sResultScript = "";
		sResultScript += fGenerateScriptMap("g_oDataLabelsBinaries", AscCommon.g_oDataLabelsBinaries);
		sResultScript += fGenerateScriptMap("g_oCatBinaries", AscCommon.g_oCatBinaries);
		sResultScript += fGenerateScriptMap("g_oValBinaries", AscCommon.g_oValBinaries);
		sResultScript += fGenerateScriptMap("g_oBarParams", AscCommon.g_oBarParams);
		sResultScript += fGenerateScriptMap("g_oView3dBinaries", AscCommon.g_oView3dBinaries);
		sResultScript += fGenerateScriptMap("g_oLegendBinaries", AscCommon.g_oLegendBinaries);
		sResultScript += fGenerateScriptMap("g_oStylesBinaries", AscCommon.g_oStylesBinaries);
		sResultScript += fGenerateScriptMap("g_oColorsBinaries", AscCommon.g_oColorsBinaries);
		sResultScript += fGenerateScriptMap("g_oChartStyles", AscCommon.g_oChartStyles, true);
		sResultScript += fGenerateScriptMap("g_oChartStylesIdMap", AscCommon.g_oChartStylesIdMap, true);
		return sResultScript;
	}
  window['AscCommon'].fSaveStream = fSaveStream;
  window['AscCommon'].fGenerateStyles = fGenerateStyles;

})(window);
