/*
 * (c) Copyright Ascensio System SIA 2010-2025
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

$(function()
{
    AscTest.Editor.GetPresentation = AscCommon.SlideEditorApi.prototype.GetPresentation.bind(AscTest.Editor);

    AscTest.Editor.private_checkPlaceholders = function(){};
    AscTest.Editor.private_CreateApiDocContent = AscCommon.SlideEditorApi.prototype.private_CreateApiDocContent.bind(AscTest.Editor);;
    AscTest.Editor.private_CreateApiParagraph = AscCommon.SlideEditorApi.prototype.private_CreateApiParagraph.bind(AscTest.Editor);;

	AscTest.JsApi = {};

	AscTest.JsApi.GetPresentation = AscCommon.SlideEditorApi.prototype.GetPresentation.bind(AscTest.Editor);
	AscTest.JsApi.CreateSlide = AscCommon.SlideEditorApi.prototype.CreateSlide.bind(AscTest.Editor);
	AscTest.JsApi.CreateMaster = AscCommon.SlideEditorApi.prototype.CreateMaster.bind(AscTest.Editor);
	AscTest.JsApi.CreateLayout = AscCommon.SlideEditorApi.prototype.CreateLayout.bind(AscTest.Editor);
	AscTest.JsApi.CreatePlaceholder = AscCommon.SlideEditorApi.prototype.CreatePlaceholder.bind(AscTest.Editor);
	AscTest.JsApi.CreateTheme = AscCommon.SlideEditorApi.prototype.CreateTheme.bind(AscTest.Editor);

	AscTest.JsApi.CreateImage = AscCommon.SlideEditorApi.prototype.CreateImage.bind(AscTest.Editor);
	AscTest.JsApi.CreateShape = AscCommon.SlideEditorApi.prototype.CreateShape.bind(AscTest.Editor);
	AscTest.JsApi.CreateChart = AscCommon.SlideEditorApi.prototype.CreateChart.bind(AscTest.Editor);
	AscTest.JsApi.CreateGroup = AscCommon.SlideEditorApi.prototype.CreateGroup.bind(AscTest.Editor);
	AscTest.JsApi.CreateTable = AscCommon.SlideEditorApi.prototype.CreateTable.bind(AscTest.Editor);
	AscTest.JsApi.CreateParagraph = AscCommon.SlideEditorApi.prototype.CreateParagraph.bind(AscTest.Editor);
	AscTest.JsApi.CreateWordArt = AscCommon.SlideEditorApi.prototype.CreateWordArt.bind(AscTest.Editor);

	AscTest.JsApi.CreateSolidFill = AscCommon.SlideEditorApi.prototype.CreateSolidFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateLinearGradientFill = AscCommon.SlideEditorApi.prototype.CreateLinearGradientFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateRadialGradientFill = AscCommon.SlideEditorApi.prototype.CreateRadialGradientFill.bind(AscTest.Editor);
	AscTest.JsApi.CreatePatternFill = AscCommon.SlideEditorApi.prototype.CreatePatternFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateBlipFill = AscCommon.SlideEditorApi.prototype.CreateBlipFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateNoFill = AscCommon.SlideEditorApi.prototype.CreateNoFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateStroke = AscCommon.SlideEditorApi.prototype.CreateStroke.bind(AscTest.Editor);
	AscTest.JsApi.CreateGradientStop = AscCommon.SlideEditorApi.prototype.CreateGradientStop.bind(AscTest.Editor);

	AscTest.JsApi.CreateRGBColor = AscCommon.SlideEditorApi.prototype.CreateRGBColor.bind(AscTest.Editor);
	AscTest.JsApi.CreateSchemeColor = AscCommon.SlideEditorApi.prototype.CreateSchemeColor.bind(AscTest.Editor);
	AscTest.JsApi.CreatePresetColor = AscCommon.SlideEditorApi.prototype.CreatePresetColor.bind(AscTest.Editor);

	// QUnit.testStart(function()
	// {
	// 	AscTest.CreateLogicDocument();
	// 	AscCommon.History.Clear();
	// 	AscTest.ClearDocument();
	// });
});
