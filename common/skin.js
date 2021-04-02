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
	"flat" : {
		Name                      : "flat",

		RulersButton              : false,
		NavigationButtons         : false,

		BackgroundColor           : "#F0F0F0",
		PageOutline               : "#BBBEC2",

		RulerDark                 : "#D9D9D9",
		RulerLight                : "#FFFFFF",
		RulerOutline              : "#BBBEC2",
		RulerMarkersOutlineColor  : "#81878F",
		RulerMarkersFillColor     : "#FFFFFF",
		RulerMarkersFillColorOld  : "#CDD1D6",
		RulerTextColor            : "#585B5E",
		RulerTabsColor            : "#000000",
		RulerTabsColorOld         : "#81878F",
		RulerTableColor1          : "#FFFFFF",
		RulerTableColor2          : "#646464",

		ScrollBackgroundColor     : "#F0F0F0",
		ScrollOutlineColor        : "#CFCFCF",
		ScrollOutlineHoverColor   : "#CFCFCF",
		ScrollOutlineActiveColor  : "#ADADAD",
		ScrollerColor             : "#F1F1F1",
		ScrollerHoverColor        : "#CFCFCF",
		ScrollerActiveColor       : "#ADADAD",
		ScrollArrowColor          : "#ADADAD",
		ScrollArrowHoverColor     : "#F1F1F1",
		ScrollArrowActiveColor    : "#F1F1F1",
		ScrollerTargetColor       : "#CFCFCF",
		ScrollerTargetHoverColor  : "#F1F1F1",
		ScrollerTargetActiveColor : "#F1F1F1",

		/* word */
		STYLE_THUMBNAIL_WIDTH     : 104,
		STYLE_THUMBNAIL_HEIGHT    : 38,

		isNeedInvertOnActive      : false,
		ContentControlsBack	   : "#F1F1F1",
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
		BackgroundColorNotes       	    : "#F0F0F0",

		THEMES_THUMBNAIL_WIDTH          : 85,
		THEMES_THUMBNAIL_HEIGHT         : 38,

		BorderSplitterColor             : "#CBCBCB",
		SupportNotes                    : true,
		SplitterWidthMM                 : 1,
		ThumbnailScrollWidthNullIfNoScrolling : false,

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
	"flatDark" : {
		Name                      : "flatDark",

		RulersButton              : false,
		NavigationButtons         : false,

		BackgroundColor           : "#666666",
		PageOutline               : "#BBBEC2",

		RulerDark                 : "#444444",
		RulerLight                : "#555555",
		RulerOutline              : "#B2B2B2",
		RulerMarkersOutlineColor  : "#81878F",
		RulerMarkersFillColor     : "#666666",
		RulerMarkersFillColorOld  : "#CDD1D6",
		RulerTextColor            : "#B2B2B2",
		RulerTabsColor            : "#FFFFFF",
		RulerTabsColorOld         : "#888888",
		RulerTableColor1          : "#FFFFFF",
		RulerTableColor2          : "#646464",

		ScrollBackgroundColor       : "#666666",
		ScrollOutlineColor          : "#404040",
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
		STYLE_THUMBNAIL_HEIGHT      : 38,

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
		BackgroundColorNotes       	    : "#666666",

		THEMES_THUMBNAIL_WIDTH  : 85,
		THEMES_THUMBNAIL_HEIGHT : 38,

		BorderSplitterColor                   : "#CBCBCB",
		SupportNotes                          : true,
		SplitterWidthMM                       : 1,
		ThumbnailScrollWidthNullIfNoScrolling : false,

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

var GlobalSkin = EditorSkins["flat"];

function updateGlobalSkin(obj)
{
	console.log(obj);
	if (!obj) return;

	if (typeof obj === "string" && undefined !== EditorSkins[obj])
	{
		GlobalSkin = EditorSkins[obj];
	}
	else
	{
		if (obj["Name"]) GlobalSkin.Name = obj["Name"];
		if (obj["RulersButton"]) GlobalSkin.RulersButton = obj["RulersButton"];
		if (obj["NavigationButtons"]) GlobalSkin.NavigationButtons = obj["NavigationButtons"];
		if (obj["BackgroundColor"]) GlobalSkin.BackgroundColor = obj["BackgroundColor"];
		if (obj["PageOutline"]) GlobalSkin.PageOutline = obj["PageOutline"];
		if (obj["RulerDark"]) GlobalSkin.RulerDark = obj["RulerDark"];
		if (obj["RulerLight"]) GlobalSkin.RulerLight = obj["RulerLight"];
		if (obj["RulerOutline"]) GlobalSkin.RulerOutline = obj["RulerOutline"];
		if (obj["RulerMarkersOutlineColor"]) GlobalSkin.RulerMarkersOutlineColor = obj["RulerMarkersOutlineColor"];
		if (obj["RulerMarkersFillColor"]) GlobalSkin.RulerMarkersFillColor = obj["RulerMarkersFillColor"];
		if (obj["RulerMarkersFillColorOld"]) GlobalSkin.RulerMarkersFillColorOld = obj["RulerMarkersFillColorOld"];
		if (obj["RulerTextColor"]) GlobalSkin.RulerTextColor = obj["RulerTextColor"];
		if (obj["RulerTabsColor"]) GlobalSkin.RulerTabsColor = obj["RulerTabsColor"];
		if (obj["RulerTabsColorOld"]) GlobalSkin.RulerTabsColorOld = obj["RulerTabsColorOld"];
		if (obj["RulerTableColor1"]) GlobalSkin.RulerTableColor1 = obj["RulerTableColor1"];
		if (obj["RulerTableColor2"]) GlobalSkin.RulerTableColor2 = obj["RulerTableColor2"];
		if (obj["ScrollBackgroundColor"]) GlobalSkin.ScrollBackgroundColor = obj["ScrollBackgroundColor"];
		if (obj["ScrollOutlineColor"]) GlobalSkin.ScrollOutlineColor = obj["ScrollOutlineColor"];
		if (obj["ScrollOutlineHoverColor"]) GlobalSkin.ScrollOutlineHoverColor = obj["ScrollOutlineHoverColor"];
		if (obj["ScrollOutlineActiveColor"]) GlobalSkin.ScrollOutlineActiveColor = obj["ScrollOutlineActiveColor"];
		if (obj["ScrollerColor"]) GlobalSkin.ScrollerColor = obj["ScrollerColor"];
		if (obj["ScrollerHoverColor"]) GlobalSkin.ScrollerHoverColor = obj["ScrollerHoverColor"];
		if (obj["ScrollerActiveColor"]) GlobalSkin.ScrollerActiveColor = obj["ScrollerActiveColor"];
		if (obj["ScrollArrowColor"]) GlobalSkin.ScrollArrowColor = obj["ScrollArrowColor"];
		if (obj["ScrollArrowHoverColor"]) GlobalSkin.ScrollArrowHoverColor = obj["ScrollArrowHoverColor"];
		if (obj["ScrollArrowActiveColor"]) GlobalSkin.ScrollArrowActiveColor = obj["ScrollArrowActiveColor"];
		if (obj["ScrollerTargetColor"]) GlobalSkin.ScrollerTargetColor = obj["ScrollerTargetColor"];
		if (obj["ScrollerTargetHoverColor"]) GlobalSkin.ScrollerTargetHoverColor = obj["ScrollerTargetHoverColor"];
		if (obj["ScrollerTargetActiveColor"]) GlobalSkin.ScrollerTargetActiveColor = obj["ScrollerTargetActiveColor"];
		if (obj["STYLE_THUMBNAIL_WIDTH"]) GlobalSkin.STYLE_THUMBNAIL_WIDTH = obj["STYLE_THUMBNAIL_WIDTH"];
		if (obj["STYLE_THUMBNAIL_HEIGHT"]) GlobalSkin.STYLE_THUMBNAIL_HEIGHT = obj["STYLE_THUMBNAIL_HEIGHT"];
		if (obj["isNeedInvertOnActive"]) GlobalSkin.isNeedInvertOnActive = obj["isNeedInvertOnActive"];
		if (obj["ContentControlsBack"]) GlobalSkin.ContentControlsBack = obj["ContentControlsBack"];
		if (obj["ContentControlsHover"]) GlobalSkin.ContentControlsHover = obj["ContentControlsHover"];
		if (obj["ContentControlsActive"]) GlobalSkin.ContentControlsActive = obj["ContentControlsActive"];
		if (obj["ContentControlsText"]) GlobalSkin.ContentControlsText = obj["ContentControlsText"];
		if (obj["ContentControlsTextActive"]) GlobalSkin.ContentControlsTextActive = obj["ContentControlsTextActive"];
		if (obj["ContentControlsAnchorActive"]) GlobalSkin.ContentControlsAnchorActive = obj["ContentControlsAnchorActive"];
		if (obj["BackgroundColorThumbnails"]) GlobalSkin.BackgroundColorThumbnails = obj["BackgroundColorThumbnails"];
		if (obj["BackgroundColorThumbnailsActive"]) GlobalSkin.BackgroundColorThumbnailsActive = obj["BackgroundColorThumbnailsActive"];
		if (obj["BackgroundColorThumbnailsHover"]) GlobalSkin.BackgroundColorThumbnailsHover = obj["BackgroundColorThumbnailsHover"];
		if (obj["ThumbnailsPageOutlineActive"]) GlobalSkin.ThumbnailsPageOutlineActive = obj["ThumbnailsPageOutlineActive"];
		if (obj["ThumbnailsPageOutlineHover"]) GlobalSkin.ThumbnailsPageOutlineHover = obj["ThumbnailsPageOutlineHover"];
		if (obj["ThumbnailsPageNumberText"]) GlobalSkin.ThumbnailsPageNumberText = obj["ThumbnailsPageNumberText"];
		if (obj["ThumbnailsPageNumberTextActive"]) GlobalSkin.ThumbnailsPageNumberTextActive = obj["ThumbnailsPageNumberTextActive"];
		if (obj["ThumbnailsPageNumberTextHover"]) GlobalSkin.ThumbnailsPageNumberTextHover = obj["ThumbnailsPageNumberTextHover"];
		if (obj["THEMES_THUMBNAIL_WIDTH"]) GlobalSkin.THEMES_THUMBNAIL_WIDTH = obj["THEMES_THUMBNAIL_WIDTH"];
		if (obj["THEMES_THUMBNAIL_HEIGHT"]) GlobalSkin.THEMES_THUMBNAIL_HEIGHT = obj["THEMES_THUMBNAIL_HEIGHT"];
		if (obj["BorderSplitterColor"]) GlobalSkin.BorderSplitterColor = obj["BorderSplitterColor"];
		if (obj["SupportNotes"]) GlobalSkin.SupportNotes = obj["SupportNotes"];
		if (obj["SplitterWidthMM"]) GlobalSkin.SplitterWidthMM = obj["SplitterWidthMM"];
		if (obj["ThumbnailScrollWidthNullIfNoScrolling"]) GlobalSkin.ThumbnailScrollWidthNullIfNoScrolling = obj["ThumbnailScrollWidthNullIfNoScrolling"];

		if (obj["Background"]) GlobalSkin.Background = obj["Background"];
		if (obj["Border"]) GlobalSkin.Border = obj["Border"];
		if (obj["Color"]) GlobalSkin.Color = obj["Color"];
		if (obj["BackgroundDark"]) GlobalSkin.BackgroundDark = obj["BackgroundDark"];
		if (obj["ColorDark"]) GlobalSkin.ColorDark = obj["ColorDark"];
		if (obj["BackgroundActive"]) GlobalSkin.BackgroundActive = obj["BackgroundActive"];
		if (obj["BorderActive"]) GlobalSkin.BorderActive = obj["BorderActive"];
		if (obj["ColorActive"]) GlobalSkin.ColorActive = obj["ColorActive"];
		if (obj["BackgroundDarkActive"]) GlobalSkin.BackgroundDarkActive = obj["BackgroundDarkActive"];
		if (obj["ColorDarkActive"]) GlobalSkin.ColorDarkActive = obj["ColorDarkActive"];
		if (obj["BackgroundHighlighted"]) GlobalSkin.BackgroundHighlighted = obj["BackgroundHighlighted"];
		if (obj["BorderHighlighted"]) GlobalSkin.BorderHighlighted = obj["BorderHighlighted"];
		if (obj["ColorHighlighted"]) GlobalSkin.ColorHighlighted = obj["ColorHighlighted"];
		if (obj["BackgroundDarkHighlighted"]) GlobalSkin.BackgroundDarkHighlighted = obj["BackgroundDarkHighlighted"];
		if (obj["ColorDarkHighlighted"]) GlobalSkin.ColorDarkHighlighted = obj["ColorDarkHighlighted"];
		if (obj["GroupDataBorder"]) GlobalSkin.GroupDataBorder = obj["GroupDataBorder"];
		if (obj["EditorBorder"]) GlobalSkin.EditorBorder = obj["EditorBorder"];
	}

	window['AscCommon'].GlobalSkin = GlobalSkin;
}

window['AscCommon'] = window['AscCommon'] || {};
window['AscCommon'].GlobalSkin = GlobalSkin;
window['AscCommon'].updateGlobalSkin = updateGlobalSkin;

if (AscCommon.TEMP_STYLE_THUMBNAIL_WIDTH !== undefined && AscCommon.TEMP_STYLE_THUMBNAIL_HEIGHT !== undefined)
{
	// TODO: переделать.
	GlobalSkin.STYLE_THUMBNAIL_WIDTH = AscCommon.TEMP_STYLE_THUMBNAIL_WIDTH;
	GlobalSkin.STYLE_THUMBNAIL_HEIGHT = AscCommon.TEMP_STYLE_THUMBNAIL_HEIGHT;
}
