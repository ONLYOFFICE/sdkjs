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

// не скрываем переменные, скин используется напрямую в sdk-all.js
// а экспорт в AscCommon - только для sdk-all-min.js
// если хочется скрыть - то везде GlobalSkin => AscCommon.GlobalSkin

var EditorSkins = {
	"theme-light" : {
		Name                      : "theme-light",
		Type                      : "light",

		RulersButton              : false,
		NavigationButtons         : false,

		BackgroundColor           : "#EEEEEE",
		PageOutline               : "#BBBEC2",

		RulerDark                 : "#D9D9D9",
		RulerLight                : "#FFFFFF",
		RulerOutline              : "#CBCBCB",
		RulerMarkersOutlineColor  : "#555555",
		RulerMarkersOutlineColorOld : "#AAAAAA",
		RulerMarkersFillColor     : "#FFFFFF",
		RulerMarkersFillColorOld  : "#FFFFFF",
		RulerTextColor            : "#555555",
		RulerTabsColor            : "#000000",
		RulerTabsColorOld         : "#828282",
		RulerTableColor1          : "#FFFFFF",
		RulerTableColor2          : "#555555",

		ScrollBackgroundColor     : "#EEEEEE",
		ScrollOutlineColor        : "#CBCBCB",
		ScrollOutlineHoverColor   : "#CBCBCB",
		ScrollOutlineActiveColor  : "#ADADAD",
		ScrollerColor             : "#F7F7F7",
		ScrollerHoverColor        : "#C0C0C0",
		ScrollerActiveColor       : "#ADADAD",
		ScrollArrowColor          : "#ADADAD",
		ScrollArrowHoverColor     : "#F7F7F7",
		ScrollArrowActiveColor    : "#F7F7F7",
		ScrollerTargetColor       : "#CFCFCF",
		ScrollerTargetHoverColor  : "#F1F1F1",
		ScrollerTargetActiveColor : "#F1F1F1",

		/* word */
		STYLE_THUMBNAIL_WIDTH     : 104,
		STYLE_THUMBNAIL_HEIGHT    : 40,

		isNeedInvertOnActive      : false,
		ContentControlsBack    : "#F1F1F1",
		ContentControlsHover   : "#D8DADC",
		ContentControlsActive  : "#7C838A",
		ContentControlsText    : "#444444",
		ContentControlsTextActive   : "#FFFFFF",
		ContentControlsAnchorActive : "#CFCFCF",
		FormsContentControlsOutlineHover : "rgba(0, 0, 0, 0.3)",
		FormsContentControlsOutlineActive : "rgba(0, 0, 0, 0.3)",
		FormsContentControlsOutlineBorderRadiusHover : 0,
		FormsContentControlsOutlineBorderRadiusActive : 2,
		FormsContentControlsMarkersBackground : "#FFFFFF",
		FormsContentControlsMarkersBackgroundHover : "#E1E1E1",
		FormsContentControlsMarkersBackgroundActive : "#CCCCCC",
		FormsContentControlsOutlineMoverHover : "#444444",
		FormsContentControlsOutlineMoverActive : "#444444",

		/* presentations */
		BackgroundColorThumbnails       : "#F4F4F4",
		BackgroundColorThumbnailsActive : "#F4F4F4",
		BackgroundColorThumbnailsHover  : "#F4F4F4",
		ThumbnailsPageOutlineActive     : "#848484",
		ThumbnailsPageOutlineHover      : "#CFCFCF",
		ThumbnailsPageNumberText        : "#000000",
		ThumbnailsPageNumberTextActive  : "#000000",
		ThumbnailsPageNumberTextHover   : "#000000",
		ThumbnailsLockColor             : "#D34F4F",
		BackgroundColorNotes            : "#F0F0F0",

		THEMES_THUMBNAIL_WIDTH          : 88,
		THEMES_THUMBNAIL_HEIGHT         : 40,
		THEMES_LAYOUT_THUMBNAIL_HEIGHT  : 68,

		BorderSplitterColor             : "#CBCBCB",
		SupportNotes                    : true,
		SplitterWidthMM                 : 1,
		ThumbnailScrollWidthNullIfNoScrolling : false,

		// demonstration
		DemBackgroundColor              : "#F0F0F0",
		DemButtonBackgroundColor        : "#FFFFFF",
		DemButtonBackgroundColorHover   : "#D8DADC",
		DemButtonBackgroundColorActive  : "#7D858C",
		DemButtonBorderColor            : "#CFCFCF",
		DemButtonTextColor              : "#444444",
		DemButtonTextColorActive        : "#FFFFFF",
		DemSplitterColor                : "#CBCBCB",
		DemTextColor                    : "#666666",

		/* spreadsheets */
		//TODO названия не менял. использую такие же как и были ранее. пересмотреть!
		Background               : "#F0F0F0",
		BackgroundActive         : "#c1c1c1",
		BackgroundHighlighted    : "#dfdfdf",

		Border                   : "#d5d5d5",
		BorderActive             : "#929292",
		BorderHighlighted        : "#afafaf",

		Color                    : "#363636",
		ColorActive              : "#363636",
		ColorHighlighted         : "#6a6a70",

		BackgroundDark           : "#444444",
		BackgroundDarkActive     : "#111111",
		BackgroundDarkHighlighted: "#666666",

		ColorDark                : "#ffffff",
		ColorDarkActive          : "#ffffff",
		ColorDarkHighlighted     : "#c1c1c1",

		GroupDataBorder          : "#000000",
		EditorBorder             : "#cbcbcb"
	},
	"theme-dark" : {
		Name                      : "theme-dark",
		Type                      : "dark",

		RulersButton              : false,
		NavigationButtons         : false,

		BackgroundColor           : "#666666",
		PageOutline               : "#BBBEC2",

		RulerDark                 : "#373737",
		RulerLight                : "#555555",
		RulerOutline              : "#2A2A2A",
		RulerMarkersOutlineColor  : "#B6B6B6",
		RulerMarkersOutlineColorOld : "#808080",
		RulerMarkersFillColor     : "#555555",
		RulerMarkersFillColorOld  : "#555555",
		RulerTextColor            : "#B6B6B6",
		RulerTabsColor            : "#FFFFFF",
		RulerTabsColorOld         : "#999999",
		RulerTableColor1          : "#FFFFFF",
		RulerTableColor2          : "#B2B2B2",

		ScrollBackgroundColor       : "#666666",
		ScrollOutlineColor          : "#2A2A2A",
		ScrollOutlineHoverColor     : "#999999",
		ScrollOutlineActiveColor    : "#ADADAD",
		ScrollerColor               : "#404040",
		ScrollerHoverColor          : "#999999",
		ScrollerActiveColor         : "#ADADAD",
		ScrollArrowColor            : "#999999",
		ScrollArrowHoverColor       : "#404040",
		ScrollArrowActiveColor      : "#404040",
		ScrollerTargetColor         : "#999999",
		ScrollerTargetHoverColor    : "#404040",
		ScrollerTargetActiveColor   : "#404040",

		/* word */
		STYLE_THUMBNAIL_WIDTH       : 104,
		STYLE_THUMBNAIL_HEIGHT      : 40,
		THEMES_LAYOUT_THUMBNAIL_HEIGHT : 68,

		isNeedInvertOnActive        : false,
		ContentControlsBack         : "#F1F1F1",
		ContentControlsHover        : "#D8DADC",
		ContentControlsActive       : "#7C838A",
		ContentControlsText         : "#444444",
		ContentControlsTextActive   : "#FFFFFF",
		ContentControlsAnchorActive : "#CFCFCF",
		FormsContentControlsOutlineHover : "rgba(0, 0, 0, 0.3)",
		FormsContentControlsOutlineActive : "rgba(0, 0, 0, 0.3)",
		FormsContentControlsOutlineBorderRadiusHover : 0,
		FormsContentControlsOutlineBorderRadiusActive : 2,
		FormsContentControlsMarkersBackground : "#FFFFFF",
		FormsContentControlsMarkersBackgroundHover : "#E1E1E1",
		FormsContentControlsMarkersBackgroundActive : "#CCCCCC",
		FormsContentControlsOutlineMoverHover : "#444444",
		FormsContentControlsOutlineMoverActive : "#444444",

		/* presentations */
		BackgroundColorThumbnails       : "#666666",
		BackgroundColorThumbnailsActive : "#666666",
		BackgroundColorThumbnailsHover  : "#666666",
		ThumbnailsPageOutlineActive     : "#848484",
		ThumbnailsPageOutlineHover      : "#CFCFCF",
		ThumbnailsPageNumberText        : "#FFFFFF",
		ThumbnailsPageNumberTextActive  : "#FFFFFF",
		ThumbnailsPageNumberTextHover   : "#FFFFFF",
		ThumbnailsLockColor             : "#D34F4F",
		BackgroundColorNotes            : "#666666",

		THEMES_THUMBNAIL_WIDTH  : 88,
		THEMES_THUMBNAIL_HEIGHT : 40,

		BorderSplitterColor                   : "#CBCBCB",
		SupportNotes                          : true,
		SplitterWidthMM                       : 1,
		ThumbnailScrollWidthNullIfNoScrolling : false,

		// demonstration
		DemBackgroundColor              : "#666666",
		DemButtonBackgroundColor        : "#333333",
		DemButtonBackgroundColorHover   : "#555555",
		DemButtonBackgroundColorActive  : "#DDDDDD",
		DemButtonBorderColor            : "#CFCFCF",
		DemButtonTextColor              : "#FFFFFF",
		DemButtonTextColorActive        : "#333333",
		DemSplitterColor                : "#CBCBCB",
		DemTextColor                    : "#FFFFFF",

		/* spreadsheets */
		Background               : "#666666",
		BackgroundActive         : "#939393",
		BackgroundHighlighted    : "#787878",

		Border                   : "#757575",
		BorderActive             : "#9e9e9e",
		BorderHighlighted        : "#858585",

		Color                    : "#d9d9d9",
		ColorActive              : "#d9d9d9",
		ColorHighlighted         : "#d9d9d9",

		BackgroundDark           : "#111111",
		BackgroundDarkActive     : "#333333",
		BackgroundDarkHighlighted: "#000000",

		ColorDark                : "#ffffff",
		ColorDarkActive          : "#ffffff",
		ColorDarkHighlighted     : "#ffffff",

		GroupDataBorder          : "#ffffff",
		EditorBorder             : "#2a2a2a"
	}
};

/*
функция для генерации "else" updateGlobalSkin
function setter_from_interface(obj)
{
	var code = "";
	for (var i in obj) {
		code += ("if (obj[\"" + i + "\"]) GlobalSkin." + i + " = obj[\"" + i + "\"];\n");
	}
	copy(code);
}
*/

var GlobalSkin = EditorSkins["theme-light"];

function updateGlobalSkinColors(theme)
{
	var skin = GlobalSkin;

	var correctColor = function (c) {
		return AscCommon.RgbaTextToRgbaHex(c);
	};

	var colorMap = {
		BackgroundColor                : "canvas-background",
		PageOutline                    : "canvas-page-border",

		RulerDark                      : "canvas-ruler-margins-background",
		RulerLight                     : "canvas-ruler-background",
		RulerOutline                   : "border-toolbar",
		RulerMarkersOutlineColor       : "canvas-ruler-handle-border",
		RulerMarkersOutlineColorOld    : "canvas-ruler-handle-border-disabled",
		RulerMarkersFillColor          : "background-normal",
		RulerMarkersFillColorOld       : "background-normal",
		RulerTextColor                 : "canvas-ruler-mark",
		RulerTabsColor                 : "canvas-high-contrast",
		RulerTabsColorOld              : "canvas-high-contrast-disabled",
		RulerTableColor1               : "background-normal",
		RulerTableColor2               : "canvas-ruler-handle-border",

		ScrollBackgroundColor          : "canvas-background",
		ScrollOutlineColor             : "canvas-scroll-thumb-border",
		ScrollOutlineHoverColor        : "canvas-scroll-thumb-border-hover",
		ScrollOutlineActiveColor       : "canvas-scroll-thumb-border-pressed",
		ScrollerColor                  : "canvas-scroll-thumb",
		ScrollerHoverColor             : "canvas-scroll-thumb-hover",
		ScrollerActiveColor            : "canvas-scroll-thumb-pressed",
		ScrollArrowColor               : "canvas-scroll-arrow",
		ScrollArrowHoverColor          : "canvas-scroll-arrow-hover",
		ScrollArrowActiveColor         : "canvas-scroll-arrow-pressed",
		ScrollerTargetColor            : "canvas-scroll-thumb-target",
		ScrollerTargetHoverColor       : "canvas-scroll-thumb-target-hover",
		ScrollerTargetActiveColor      : "canvas-scroll-thumb-target-pressed",

		/* presentations */
		BackgroundColorThumbnails       : "background-toolbar",
		BackgroundColorThumbnailsActive : "background-toolbar",
		BackgroundColorThumbnailsHover  : "background-toolbar",
		ThumbnailsPageOutlineActive     : "border-preview-select",
		ThumbnailsPageOutlineHover      : "border-preview-hover",
		ThumbnailsPageNumberText        : "text-normal",
		ThumbnailsPageNumberTextActive  : "text-normal",
		ThumbnailsPageNumberTextHover   : "text-normal",
		BackgroundColorNotes            : "canvas-background",

		BorderSplitterColor             : "border-toolbar",

		// demonstration
		DemBackgroundColor              : "background-toolbar",
		DemButtonBackgroundColor        : "background-normal",
		DemButtonBackgroundColorHover   : "highlight-buttin-hover",
		DemButtonBackgroundColorActive  : "highlight-button-pressed",
		DemButtonBorderColor            : "border-regular-control",
		DemButtonTextColor              : "text-normal",
		DemButtonTextColorActive        : "text-normal-pressed",
		DemSplitterColor                : "border-divider",
		DemTextColor                    : "text-normal",

		/* spreadsheets */
		Background                      : "canvas-background",
		BackgroundActive                : "canvas-cell-title-selected",
		BackgroundHighlighted           : "canvas-cell-title-hover",

		Border                          : "canvas-cell-title-border",
		BorderActive                    : "canvas-cell-title-border-selected",
		BorderHighlighted               : "canvas-cell-title-border-hover",

		Color                           : "text-normal",
		ColorActive                     : "text-normal",
		ColorHighlighted                : "text-normal",

		BackgroundDark                  : "canvas-dark-cell-title",
		BackgroundDarkActive            : "canvas-dark-cell-title-selected",
		BackgroundDarkHighlighted       : "canvas-dark-cell-title-hover",

		ColorDark                       : "text-contrast-background",
		ColorDarkActive                 : "text-contrast-background",
		ColorDarkHighlighted            : "text-contrast-background",

		GroupDataBorder                 : "canvas-high-contrast",
		EditorBorder                    : "border-toolbar"
	};

	// корректируем цвета для старого хрома:
	// в старых хромах (desktop windows XP)
	// если начинается цвет с цифры (#0-9) - то помечается символом \3 (конец текста)
	for (var item in theme)
	{
		var testValue = theme[item];
		if (typeof testValue !== "string")
			continue;

		if (0 === testValue.indexOf("#\\3"))
		{
			testValue = testValue.replace("\\3", "");
			testValue = testValue.replace(" ", "");
			theme[item] = testValue;
		}
	}

	for (var color in colorMap)
	{
		if (undefined === GlobalSkin[color])
			continue;
		if ("" === colorMap[color])
			continue;
		if (undefined === theme[colorMap[color]])
			continue;

		if (0 === GlobalSkin[color].indexOf("rgb"))
			GlobalSkin[color] = theme[colorMap[color]];
		else
			GlobalSkin[color] = correctColor(theme[colorMap[color]]);
	}
}

function updateGlobalSkin(obj)
{
	if (!obj) return;

	if (typeof obj === "string")
	{
		var name = obj;
		obj = {
			"name" : name,
			"type" : ((-1 !== name.indexOf("dark")) ? "dark" : "light")
		};
	}

	if (obj["name"] && undefined !== EditorSkins[obj["name"]])
		GlobalSkin = EditorSkins[obj["name"]];
	else if (obj["type"])
	{
		for (var item in EditorSkins)
		{
			if (obj["type"] === EditorSkins[item].Type)
			{
				GlobalSkin = EditorSkins[item];
				break;
			}
		}
	}

	updateGlobalSkinColors(obj);
	for (var item in obj)
		GlobalSkin[item] = obj[item];

	if (window.g_asc_plugins)
		window.g_asc_plugins.onThemeChanged(GlobalSkin);

	window['AscCommon'].GlobalSkin = GlobalSkin;
}

window['AscCommon'] = window['AscCommon'] || {};
window['AscCommon'].GlobalSkin = GlobalSkin;
window['AscCommon'].updateGlobalSkin = updateGlobalSkin;

window['AscCommon'].RgbaHexToRGBA = function(color)
{
	var index = 0;
	if ("#".charCodeAt(0) === color.charCodeAt(0))
		index++;

	var ret = {
		R : 0,
		G : 0,
		B : 0,
		A : 255
	};

	if (6 <= color.length)
	{
		ret.R = parseInt( color.substring( index, index + 2 ), 16 );
		ret.G = parseInt( color.substring( index + 2, index + 4 ), 16 );
		ret.B = parseInt( color.substring( index + 4, index + 6 ), 16 );
	}
	else
	{
		ret.R = parseInt( color.substring( index, index + 1 ), 16 );
		ret.G = parseInt( color.substring( index + 1, index + 2 ), 16 );
		ret.B = parseInt( color.substring( index + 2, index + 3 ), 16 );

		ret.R = (ret.R << 4) | ret.R;
		ret.G = (ret.G << 4) | ret.G;
		ret.B = (ret.B << 4) | ret.B;
	}

	return ret;
};
window['AscCommon'].RgbaTextToRgbaHex = function(color)
{
	var toHex = function (c) {
		var res = Number(c).toString(16);
		return res.length === 1 ? "0" + res : res;
	};

	if (0 !== color.indexOf("rgb"))
	{
		if (color.length < 6)
		{
			var rgba = AscCommon.RgbaHexToRGBA(color);
			return "#" + toHex(rgba.R) + toHex(rgba.G) + toHex(rgba.B);
		}
		return color;
	}

	var start = color.indexOf("(");
	var end = color.indexOf(")");
	var tmp = color.substring(start + 1, end);
	var colors = tmp.split(",");

	for (var i in colors)
		colors[i] = colors[i].trim();

	var r = colors[0] || 0;
	var g = colors[1] || 0;
	var b = colors[2] || 0;
	var a = (colors[3] === undefined) ? 255 : colors[3];

	return "#" + toHex(r) + toHex(g) + toHex(b);
};

if (AscCommon.TEMP_STYLE_THUMBNAIL_WIDTH !== undefined && AscCommon.TEMP_STYLE_THUMBNAIL_HEIGHT !== undefined)
{
	// TODO: переделать.
	GlobalSkin.STYLE_THUMBNAIL_WIDTH = AscCommon.TEMP_STYLE_THUMBNAIL_WIDTH;
	GlobalSkin.STYLE_THUMBNAIL_HEIGHT = AscCommon.TEMP_STYLE_THUMBNAIL_HEIGHT;
}
