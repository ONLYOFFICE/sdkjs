/*
 * (c) Copyright Ascensio System SIA 2010-2026
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

var AscTest = AscTest || {};

$(function()
{
	//------------------------------------------------------------------------------------------------------------------
	// Move to Common
	const drawingDocument = {
		CanvasHit : null,
		CanvasHitContext : null,
		
		OnStartRecalculate : function(){},
		OnRecalculatePage : function(){},
		OnEndRecalculate : function(){},
		UpdateTargetTransform : function(){},
		SelectEnabled : function(){},
		SelectShow : function(){},
		TargetStart : function(){},
		TargetShow : function(){},
		TargetEnd : function(){},
		showTarget : function(){},
		Set_RulerState_Start : function(){},
		Set_RulerState_Paragraph : function(){},
		Set_RulerState_End : function(){},
		Update_MathTrack : function(){},
		startCollectContentControlTracks : function(){},
		endCollectContentControlTracks : function(){},
		addContentControlTrack : function(obj, state, geom){},
		removeContentControlTrackHover : function(){},
		Update_FieldTrack : function(){},
		SetTargetColor : function(){},
		SetTargetSize : function(){},
		UpdateTarget : function(){},
		ClearCachePages : function(){},
		OnRepaintPage : function(){},
		FirePaint : function(){},
		GetMMPerDot : function(value){return value / this.GetDotsPerMM(1);},
		GetDotsPerMM : function(value) {return 72;},
		EndTrackTable : function() {},
		SetCurrentPage : function(pageNum) {},
		SelectClear : function() {},
		Start_CollaborationEditing : function() {},
		End_CollaborationEditing : function() {},
		ConvertCoordsToCursorWR : function() {return {X : 0, Y : 0};},
		Set_RulerState_Table : function() {},
		scrollToTarget : function() {},
		GetVisibleRegion : function() {return [{Page : 0, Y : 0}, {Page : 0, Y : 0}]},
	};
	
	drawingDocument.CanvasHit = document.createElement('canvas');
	drawingDocument.CanvasHitContext = drawingDocument.CanvasHit.getContext('2d');
	//------------------------------------------------------------------------------------------------------------------
	
	 // ====== REQUIRED ENVIRONMENT SETUP (preserve these stubs/settings) ======
	Asc.spreadsheet_api.prototype._init = function() {};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	AscCommonExcel.Workbook.prototype._getSnapshot = function() {return null;};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function() {};
	AscCommonExcel.WorkbookView.prototype._init = function() {};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function() {};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function() {};
	AscCommonExcel.WorksheetView.prototype._init = function () {};
	AscCommonExcel.WorksheetView.prototype._onUpdateFormatTable = function() {};
	AscCommonExcel.WorksheetView.prototype.setSelection = function () {};
	AscCommonExcel.WorksheetView.prototype.draw = function() {};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function() {};
	AscCommonExcel.WorksheetView.prototype._reinitializeScroll = function() {};
	AscCommonExcel.WorksheetView.prototype.getZoom = function() {};
	AscCommonExcel.WorksheetView.prototype._getPPIY = function() {};
	AscCommonExcel.WorksheetView.prototype._getPPIX = function() {};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function() {};
	Asc.ReadDefTableStyles = function() {};
	
	var api = new Asc.spreadsheet_api({ "id-view": "editor_sdk" });
	
	AscTest.Editor = api;
	api.FontLoader = { LoadDocumentFonts: function() {} };
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	api.initCollaborativeEditing({});
	api.wbModel.DrawingDocument = drawingDocument;
	api.wbModel.mathTrackHandler = new AscWord.CMathTrackHandler(drawingDocument, api);
	api.wb = new AscCommonExcel.WorkbookView(
		api.wbModel,
		api.controller,
		api.handlers,
		api.HtmlElement,
		api.topLineEditorElement,
		api,
		api.collaborativeEditing,
		api.fontRenderingMode
	);
	
	var wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();
	wsView.objectRender.OnUpdateOverlay = function() {};
	wsView.objectRender.drawingDocument = drawingDocument;
	wsView.objectRender.controller = new AscFormat.DrawingObjectsController(wsView.objectRender);
	
	var ws = api.GetActiveSheet();

	AscTest.JsApi = {};
	
	AscTest.JsApi.GetActiveSheet = Asc.spreadsheet_api.prototype.GetActiveSheet.bind(AscTest.Editor);
	AscTest.JsApi.GetRange = Asc.spreadsheet_api.prototype.GetRange.bind(AscTest.Editor);
	AscTest.JsApi.CreateRGBColor = Asc.spreadsheet_api.prototype.CreateRGBColor.bind(AscTest.Editor);
	AscTest.JsApi.CreateSolidFill = Asc.spreadsheet_api.prototype.CreateSolidFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateStroke = Asc.spreadsheet_api.prototype.CreateStroke.bind(AscTest.Editor);
	AscTest.JsApi.CreateNoFill = Asc.spreadsheet_api.prototype.CreateNoFill.bind(AscTest.Editor);
	AscTest.JsApi.CreateColorFromRGB = Asc.spreadsheet_api.prototype.CreateColorFromRGB.bind(AscTest.Editor);
	AscTest.JsApi.CreateGradientStop = Asc.spreadsheet_api.prototype.CreateGradientStop.bind(AscTest.Editor);
	AscTest.JsApi.CreateRadialGradientFill = Asc.spreadsheet_api.prototype.CreateRadialGradientFill.bind(AscTest.Editor);
	AscTest.JsApi.AddDefName = Asc.spreadsheet_api.prototype.AddDefName.bind(AscTest.Editor);
	
	AscTest.JsApi.wbModel = AscTest.Editor.wbModel;
	
	AscTest.Workbook = AscTest.Editor.wbModel;
	AscTest.WorkbookView = AscTest.Editor.wb;
	
	QUnit.testStart(function()
	{
		let range = AscTest.JsApi.GetRange('A1:Z100');
		range.Clear();
		
		ws.worksheet.AutoFilter = null;

		if (ws && ws.dataValidations)
			ws.dataValidations.clear(ws, true);
		
		AscTest.Editor.asc_cleanWorksheet();
		let wsView = AscTest.WorkbookView.getWorksheet();
		wsView && wsView.objectRender && wsView.objectRender.controller && wsView.objectRender.controller.remove(1);
	});
	
	AscTest.GetActiveWorksheet = function()
	{
		return ws;
	};
});
