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
	AscTest.Editor.GetDocument = AscBuilder.Word.Api.GetDocument.bind(AscTest.Editor);
	AscTest.JsApi = {};
	
	AscTest.JsApi.GetDocument = AscBuilder.Word.Api.GetDocument.bind(AscTest.Editor);
	AscTest.JsApi.ReplaceTextSmart = AscBuilder.Word.Api.ReplaceTextSmart.bind(AscTest.Editor);
	AscTest.JsApi.CreateRun = AscBuilder.Word.Api.CreateRun.bind(AscTest.Editor);
	AscTest.JsApi.CreateParagraph = AscBuilder.Word.Api.CreateParagraph.bind(AscTest.Editor);
	AscTest.JsApi.CreateInlineLvlSdt = AscBuilder.Word.Api.CreateInlineLvlSdt.bind(AscTest.Editor);
	AscTest.JsApi.CreateTable = AscBuilder.Word.Api.CreateTable.bind(AscTest.Editor);
	AscTest.JsApi.CreateShape = AscBuilder.Word.Api.CreateShape.bind(AscTest.Editor);
	AscTest.JsApi.CreateSolidFill = AscBuilder.Word.Api.CreateSolidFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateStroke = AscBuilder.Word.Api.CreateStroke.bind(AscTest.Editor);
	AscTest.JsApi.CreateNoFill = AscBuilder.Word.Api.CreateNoFill.bind(AscTest.Editor);
	AscTest.JsApi.HexColor = AscBuilder.Word.Api.HexColor.bind(AscTest.Editor);
	AscTest.JsApi.ThemeColor = AscBuilder.Word.Api.ThemeColor.bind(AscTest.Editor);
	AscTest.JsApi.AutoColor = AscBuilder.Word.Api.AutoColor.bind(AscTest.Editor);
	AscTest.JsApi.RGBA = AscBuilder.Word.Api.RGBA.bind(AscTest.Editor);
	AscTest.JsApi.RGB = AscBuilder.Word.Api.RGB.bind(AscTest.Editor);
	AscTest.JsApi.FromJSON = AscBuilder.Word.Api.FromJSON.bind(AscTest.Editor);
	
	AscTest.JsApi.CreateDocContent = function()
	{
		let docContent = new AscWord.CDocumentContent();
		return new AscBuilder.ApiDocumentContent(docContent);
	};

	QUnit.testStart(function()
	{
		AscTest.CreateLogicDocument();
		AscCommon.History.Clear();
		AscTest.ClearDocument();
	});
});
