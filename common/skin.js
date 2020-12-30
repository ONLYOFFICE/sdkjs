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
		ContentControlsBack       : "#F1F1F1",
		ContentControlsHover      : "#D8DADC",
		ContentControlsActive     : "#7C838A",
		ContentControlsText       : "#444444",
		ContentControlsTextActive : "#FFFFFF",
		ContentControlsAnchorActive : "#CFCFCF",

		/* presentations */
		BackgroundColorThumbnails       : "#F4F4F4",
		BackgroundColorThumbnailsActive : "#F4F4F4",
		BackgroundColorThumbnailsHover  : "#F4F4F4",
		ThumbnailsPageOutlineActive     : "#848484",
		ThumbnailsPageOutlineHover      : "#CFCFCF",
		ThumbnailsPageNumberText        : "#000000",
		ThumbnailsPageNumberTextActive  : "#000000",
		ThumbnailsPageNumberTextHover   : "#000000",

		THEMES_THUMBNAIL_WIDTH          : 85,
		THEMES_THUMBNAIL_HEIGHT         : 38,

		BorderSplitterColor             : "#CBCBCB",
		SupportNotes                    : true,
		SplitterWidthMM                 : 1,
		ThumbnailScrollWidthNullIfNoScrolling : false,

		/* spreadsheets */
		//TODO названия не менял. использую такие же как и были ранее. пересмотреть!
		Background               : "#f1f1f1",
		Border                   : "#d5d5d5",
		Color                    : "#363636",
		BackgroundDark           : "#444444",
		ColorDark                : "#ffffff",
		BackgroundActive         : "#c1c1c1",
		BorderActive             : "#929292",
		ColorActive              : "#363636",
		BackgroundDarkActive     : "#000000",
		ColorDarkActive          : "#ffffff",
		BackgroundHighlighted    : "#dfdfdf",
		BorderHighlighted        : "#afafaf",
		ColorHighlighted         : "#6a6a70",
		BackgroundDarkHighlighted: "#ffffff",
		ColorDarkHighlighted     : "#c1c1c1"

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
		RulerTabsColorOld         : "#81878F",
		RulerTableColor1          : "#FFFFFF",
		RulerTableColor2          : "#646464",

		ScrollBackgroundColor       : "#666666",
		ScrollOutlineColor          : "#CFCFCF",
		ScrollOutlineHoverColor     : "#CFCFCF",
		ScrollOutlineActiveColor    : "#ADADAD",
		ScrollerColor               : "#F1F1F1",
		ScrollerHoverColor          : "#CFCFCF",
		ScrollerActiveColor         : "#ADADAD",
		ScrollArrowColor            : "#ADADAD",
		ScrollArrowHoverColor       : "#F1F1F1",
		ScrollArrowActiveColor      : "#F1F1F1",
		ScrollerTargetColor         : "#CFCFCF",
		ScrollerTargetHoverColor    : "#F1F1F1",
		ScrollerTargetActiveColor   : "#F1F1F1",

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

		/* presentations */
		BackgroundColorThumbnails       : "#666666",
		BackgroundColorThumbnailsActive : "#666666",
		BackgroundColorThumbnailsHover  : "#666666",
		ThumbnailsPageOutlineActive     : "#848484",
		ThumbnailsPageOutlineHover      : "#CFCFCF",
		ThumbnailsPageNumberText        : "#FFFFFF",
		ThumbnailsPageNumberTextActive  : "#FFFFFF",
		ThumbnailsPageNumberTextHover   : "#FFFFFF",

		THEMES_THUMBNAIL_WIDTH  : 85,
		THEMES_THUMBNAIL_HEIGHT : 38,

		BorderSplitterColor                   : "#CBCBCB",
		SupportNotes                          : true,
		SplitterWidthMM                       : 1,
		ThumbnailScrollWidthNullIfNoScrolling : false,

		Background               : "#f1f1f1",
		Border                   : "#d5d5d5",
		Color                    : "#363636",
		BackgroundDark           : "#444444",
		ColorDark                : "#ffffff",
		BackgroundActive         : "#c1c1c1",
		BorderActive             : "#929292",
		ColorActive              : "#363636",
		BackgroundDarkActive     : "#000000",
		ColorDarkActive          : "#ffffff",
		BackgroundHighlighted    : "#dfdfdf",
		BorderHighlighted        : "#afafaf",
		ColorHighlighted         : "#6a6a70",
		BackgroundDarkHighlighted: "#ffffff",
		ColorDarkHighlighted     : "#c1c1c1"
	}
};

/*
функция для генерации "else" updateGlobalSkin
function setter_from_interface(obj)
{
	var code = "";
	for (var i in obj) {
		code += ("if (undefined !== obj[\"" + i + "\"]) GlobalSkin." + i + " = obj[\"" + i + "\"];\n");
	}
	copy(code);
}
*/

var GlobalSkin = EditorSkins["flat"];

function updateGlobalSkin(obj)
{
	if (!obj) return;

	if (typeof obj === "string" && undefined !== EditorSkins[obj])
	{
		GlobalSkin = EditorSkins[obj];
	}
	else
	{
		if (undefined !== obj["Name"]) GlobalSkin.Name = obj["Name"];
		if (undefined !== obj["RulersButton"]) GlobalSkin.RulersButton = obj["RulersButton"];
		if (undefined !== obj["NavigationButtons"]) GlobalSkin.NavigationButtons = obj["NavigationButtons"];
		if (undefined !== obj["BackgroundColor"]) GlobalSkin.BackgroundColor = obj["BackgroundColor"];
		if (undefined !== obj["PageOutline"]) GlobalSkin.PageOutline = obj["PageOutline"];
		if (undefined !== obj["RulerDark"]) GlobalSkin.RulerDark = obj["RulerDark"];
		if (undefined !== obj["RulerLight"]) GlobalSkin.RulerLight = obj["RulerLight"];
		if (undefined !== obj["RulerOutline"]) GlobalSkin.RulerOutline = obj["RulerOutline"];
		if (undefined !== obj["RulerMarkersOutlineColor"]) GlobalSkin.RulerMarkersOutlineColor = obj["RulerMarkersOutlineColor"];
		if (undefined !== obj["RulerMarkersFillColor"]) GlobalSkin.RulerMarkersFillColor = obj["RulerMarkersFillColor"];
		if (undefined !== obj["RulerMarkersFillColorOld"]) GlobalSkin.RulerMarkersFillColorOld = obj["RulerMarkersFillColorOld"];
		if (undefined !== obj["RulerTextColor"]) GlobalSkin.RulerTextColor = obj["RulerTextColor"];
		if (undefined !== obj["RulerTabsColor"]) GlobalSkin.RulerTabsColor = obj["RulerTabsColor"];
		if (undefined !== obj["RulerTabsColorOld"]) GlobalSkin.RulerTabsColorOld = obj["RulerTabsColorOld"];
		if (undefined !== obj["RulerTableColor1"]) GlobalSkin.RulerTableColor1 = obj["RulerTableColor1"];
		if (undefined !== obj["RulerTableColor2"]) GlobalSkin.RulerTableColor2 = obj["RulerTableColor2"];
		if (undefined !== obj["ScrollBackgroundColor"]) GlobalSkin.ScrollBackgroundColor = obj["ScrollBackgroundColor"];
		if (undefined !== obj["ScrollOutlineColor"]) GlobalSkin.ScrollOutlineColor = obj["ScrollOutlineColor"];
		if (undefined !== obj["ScrollOutlineHoverColor"]) GlobalSkin.ScrollOutlineHoverColor = obj["ScrollOutlineHoverColor"];
		if (undefined !== obj["ScrollOutlineActiveColor"]) GlobalSkin.ScrollOutlineActiveColor = obj["ScrollOutlineActiveColor"];
		if (undefined !== obj["ScrollerColor"]) GlobalSkin.ScrollerColor = obj["ScrollerColor"];
		if (undefined !== obj["ScrollerHoverColor"]) GlobalSkin.ScrollerHoverColor = obj["ScrollerHoverColor"];
		if (undefined !== obj["ScrollerActiveColor"]) GlobalSkin.ScrollerActiveColor = obj["ScrollerActiveColor"];
		if (undefined !== obj["ScrollArrowColor"]) GlobalSkin.ScrollArrowColor = obj["ScrollArrowColor"];
		if (undefined !== obj["ScrollArrowHoverColor"]) GlobalSkin.ScrollArrowHoverColor = obj["ScrollArrowHoverColor"];
		if (undefined !== obj["ScrollArrowActiveColor"]) GlobalSkin.ScrollArrowActiveColor = obj["ScrollArrowActiveColor"];
		if (undefined !== obj["ScrollerTargetColor"]) GlobalSkin.ScrollerTargetColor = obj["ScrollerTargetColor"];
		if (undefined !== obj["ScrollerTargetHoverColor"]) GlobalSkin.ScrollerTargetHoverColor = obj["ScrollerTargetHoverColor"];
		if (undefined !== obj["ScrollerTargetActiveColor"]) GlobalSkin.ScrollerTargetActiveColor = obj["ScrollerTargetActiveColor"];
		if (undefined !== obj["STYLE_THUMBNAIL_WIDTH"]) GlobalSkin.STYLE_THUMBNAIL_WIDTH = obj["STYLE_THUMBNAIL_WIDTH"];
		if (undefined !== obj["STYLE_THUMBNAIL_HEIGHT"]) GlobalSkin.STYLE_THUMBNAIL_HEIGHT = obj["STYLE_THUMBNAIL_HEIGHT"];
		if (undefined !== obj["isNeedInvertOnActive"]) GlobalSkin.isNeedInvertOnActive = obj["isNeedInvertOnActive"];
		if (undefined !== obj["ContentControlsBack"]) GlobalSkin.ContentControlsBack = obj["ContentControlsBack"];
		if (undefined !== obj["ContentControlsHover"]) GlobalSkin.ContentControlsHover = obj["ContentControlsHover"];
		if (undefined !== obj["ContentControlsActive"]) GlobalSkin.ContentControlsActive = obj["ContentControlsActive"];
		if (undefined !== obj["ContentControlsText"]) GlobalSkin.ContentControlsText = obj["ContentControlsText"];
		if (undefined !== obj["ContentControlsTextActive"]) GlobalSkin.ContentControlsTextActive = obj["ContentControlsTextActive"];
		if (undefined !== obj["ContentControlsAnchorActive"]) GlobalSkin.ContentControlsAnchorActive = obj["ContentControlsAnchorActive"];
		if (undefined !== obj["BackgroundColorThumbnails"]) GlobalSkin.BackgroundColorThumbnails = obj["BackgroundColorThumbnails"];
		if (undefined !== obj["BackgroundColorThumbnailsActive"]) GlobalSkin.BackgroundColorThumbnailsActive = obj["BackgroundColorThumbnailsActive"];
		if (undefined !== obj["BackgroundColorThumbnailsHover"]) GlobalSkin.BackgroundColorThumbnailsHover = obj["BackgroundColorThumbnailsHover"];
		if (undefined !== obj["ThumbnailsPageOutlineActive"]) GlobalSkin.ThumbnailsPageOutlineActive = obj["ThumbnailsPageOutlineActive"];
		if (undefined !== obj["ThumbnailsPageOutlineHover"]) GlobalSkin.ThumbnailsPageOutlineHover = obj["ThumbnailsPageOutlineHover"];
		if (undefined !== obj["ThumbnailsPageNumberText"]) GlobalSkin.ThumbnailsPageNumberText = obj["ThumbnailsPageNumberText"];
		if (undefined !== obj["ThumbnailsPageNumberTextActive"]) GlobalSkin.ThumbnailsPageNumberTextActive = obj["ThumbnailsPageNumberTextActive"];
		if (undefined !== obj["ThumbnailsPageNumberTextHover"]) GlobalSkin.ThumbnailsPageNumberTextHover = obj["ThumbnailsPageNumberTextHover"];
		if (undefined !== obj["THEMES_THUMBNAIL_WIDTH"]) GlobalSkin.THEMES_THUMBNAIL_WIDTH = obj["THEMES_THUMBNAIL_WIDTH"];
		if (undefined !== obj["THEMES_THUMBNAIL_HEIGHT"]) GlobalSkin.THEMES_THUMBNAIL_HEIGHT = obj["THEMES_THUMBNAIL_HEIGHT"];
		if (undefined !== obj["BorderSplitterColor"]) GlobalSkin.BorderSplitterColor = obj["BorderSplitterColor"];
		if (undefined !== obj["SupportNotes"]) GlobalSkin.SupportNotes = obj["SupportNotes"];
		if (undefined !== obj["SplitterWidthMM"]) GlobalSkin.SplitterWidthMM = obj["SplitterWidthMM"];
		if (undefined !== obj["ThumbnailScrollWidthNullIfNoScrolling"]) GlobalSkin.ThumbnailScrollWidthNullIfNoScrolling = obj["ThumbnailScrollWidthNullIfNoScrolling"];

		if (undefined !== obj["Background"]) GlobalSkin.Background = obj["Background"];
		if (undefined !== obj["Border"]) GlobalSkin.Border = obj["Border"];
		if (undefined !== obj["Color"]) GlobalSkin.Color = obj["Color"];
		if (undefined !== obj["BackgroundDark"]) GlobalSkin.BackgroundDark = obj["BackgroundDark"];
		if (undefined !== obj["ColorDark"]) GlobalSkin.ColorDark = obj["ColorDark"];
		if (undefined !== obj["BackgroundActive"]) GlobalSkin.BackgroundActive = obj["BackgroundActive"];
		if (undefined !== obj["BorderActive"]) GlobalSkin.BorderActive = obj["BorderActive"];
		if (undefined !== obj["ColorActive"]) GlobalSkin.ColorActive = obj["ColorActive"];
		if (undefined !== obj["BackgroundDarkActive"]) GlobalSkin.BackgroundDarkActive = obj["BackgroundDarkActive"];
		if (undefined !== obj["ColorDarkActive"]) GlobalSkin.ColorDarkActive = obj["ColorDarkActive"];
		if (undefined !== obj["BackgroundHighlighted"]) GlobalSkin.BackgroundHighlighted = obj["BackgroundHighlighted"];
		if (undefined !== obj["BorderHighlighted"]) GlobalSkin.BorderHighlighted = obj["BorderHighlighted"];
		if (undefined !== obj["ColorHighlighted"]) GlobalSkin.ColorHighlighted = obj["ColorHighlighted"];
		if (undefined !== obj["BackgroundDarkHighlighted"]) GlobalSkin.BackgroundDarkHighlighted = obj["BackgroundDarkHighlighted"];
		if (undefined !== obj["ColorDarkHighlighted"]) GlobalSkin.ColorDarkHighlighted = obj["ColorDarkHighlighted"];
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
