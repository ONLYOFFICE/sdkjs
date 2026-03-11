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
	AscBuilder.Slide.init();

	Asc.editor.getLogicDocument = function() { return Asc.editor.WordControl.m_oLogicDocument; };

	AscTest.JsApi = {};

	AscTest.JsApi.GetPresentation = AscBuilder.Slide.Api.GetPresentation.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateSlide = AscBuilder.Slide.Api.CreateSlide.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateMaster = AscBuilder.Slide.Api.CreateMaster.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateLayout = AscBuilder.Slide.Api.CreateLayout.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreatePlaceholder = AscBuilder.Slide.Api.CreatePlaceholder.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateTheme = AscBuilder.Slide.Api.CreateTheme.bind(AscBuilder.Slide.Api);

	AscTest.JsApi.CreateImage = AscBuilder.Slide.Api.CreateImage.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateShape = AscBuilder.Slide.Api.CreateShape.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateChart = AscBuilder.Slide.Api.CreateChart.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateGroup = AscBuilder.Slide.Api.CreateGroup.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateTable = AscBuilder.Slide.Api.CreateTable.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateParagraph = AscBuilder.Slide.Api.CreateParagraph.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateWordArt = AscBuilder.Slide.Api.CreateWordArt.bind(AscBuilder.Slide.Api);

	AscTest.JsApi.CreateSolidFill = AscBuilder.Slide.Api.CreateSolidFill.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateLinearGradientFill = AscBuilder.Slide.Api.CreateLinearGradientFill.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateRadialGradientFill = AscBuilder.Slide.Api.CreateRadialGradientFill.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreatePatternFill = AscBuilder.Slide.Api.CreatePatternFill.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateBlipFill = AscBuilder.Slide.Api.CreateBlipFill.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateNoFill = AscBuilder.Slide.Api.CreateNoFill.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateStroke = AscBuilder.Slide.Api.CreateStroke.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateGradientStop = AscBuilder.Slide.Api.CreateGradientStop.bind(AscBuilder.Slide.Api);

	AscTest.JsApi.CreateRGBColor = AscBuilder.Slide.Api.CreateRGBColor.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreateSchemeColor = AscBuilder.Slide.Api.CreateSchemeColor.bind(AscBuilder.Slide.Api);
	AscTest.JsApi.CreatePresetColor = AscBuilder.Slide.Api.CreatePresetColor.bind(AscBuilder.Slide.Api);
});
