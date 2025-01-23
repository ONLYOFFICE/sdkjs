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
QUnit.config.autostart = false;
$(function() {


	Asc.spreadsheet_api.prototype._init = function() {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype.asyncImagesDocumentEndLoaded = function() {
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function(fonts, callback) {
		openDocument();
	};
	Asc.spreadsheet_api.prototype._onEndLoadSdk = function () {
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk.call(this);
	};
	Asc.spreadsheet_api.prototype.initGlobalObjects = function(wbModel) {
		AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wbModel);
		AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wbModel);
		AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wbModel);
		AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wbModel, false);
		AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wbModel, true);
		AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wbModel);
		AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wbModel);
		AscCommonExcel.g_oUndoRedoSparklines = new AscCommonExcel.UndoRedoSparklines(wbModel);
		AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wbModel, -1);
		AscCommonExcel.g_oUndoRedoSharedFormula = new AscCommonExcel.UndoRedoSharedFormula(wbModel);
		AscCommonExcel.g_oUndoRedoLayout = new AscCommonExcel.UndoRedoRedoLayout(wbModel);
		AscCommonExcel.g_oUndoRedoHeaderFooter = new AscCommonExcel.UndoRedoHeaderFooter(wbModel);
		AscCommonExcel.g_oUndoRedoArrayFormula = new AscCommonExcel.UndoRedoArrayFormula(wbModel);
		AscCommonExcel.g_oUndoRedoSortState = new AscCommonExcel.UndoRedoSortState(wbModel);
		AscCommonExcel.g_oUndoRedoSlicer = new AscCommonExcel.UndoRedoSlicer(wbModel);
		AscCommonExcel.g_oUndoRedoPivotTables = new AscCommonExcel.UndoRedoPivotTables(wbModel);
		AscCommonExcel.g_oUndoRedoPivotFields = new AscCommonExcel.UndoRedoPivotFields(wbModel);
		AscCommonExcel.g_oUndoRedoPivotFieldItems = new AscCommonExcel.UndoRedoPivotFieldItems(wbModel);
		AscCommonExcel.g_oUndoRedoPivotCache = new AscCommonExcel.UndoRedoPivotCache(wbModel);
		AscCommonExcel.g_oUndoRedoCacheFields = new AscCommonExcel.UndoRedoCacheFields(wbModel);
		AscCommonExcel.g_oUndoRedoCF = new AscCommonExcel.UndoRedoCF(wbModel);
		AscCommonExcel.g_oUndoRedoProtectedRange = new AscCommonExcel.UndoRedoProtectedRange(wbModel);
		AscCommonExcel.g_oUndoRedoProtectedSheet = new AscCommonExcel.UndoRedoProtectedSheet(wbModel);
		AscCommonExcel.g_oUndoRedoProtectedWorkbook = new AscCommonExcel.UndoRedoProtectedWorkbook(wbModel);
		AscCommonExcel.g_oUndoRedoNamedSheetViews = new AscCommonExcel.UndoRedoNamedSheetViews(wbModel);
		AscCommonExcel.g_oUndoRedoUserProtectedRange = new AscCommonExcel.UndoRedoUserProtectedRange(wbModel);
	};
	Asc.spreadsheet_api.prototype._onUpdateDocumentCanSave = function() {
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function() {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function() {
	};
	AscCommonExcel.WorkbookView.prototype.getZoom = function() {
		return 1;
	};
	AscCommonExcel.WorkbookView.prototype.changeZoom = function() {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function() {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function() {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function() {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function() {
		this.objectRender = new AscFormat.DrawingObjects();
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function() {
		return 1;
	};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function() {
		AscFonts.g_fontApplication.Init();

		this.FontLoader  = AscCommon.g_font_loader;
		this.ImageLoader = AscCommon.g_image_loader;
		this.FontLoader.put_Api(this);
		this.ImageLoader.put_Api(this);
	};

	Asc.spreadsheet_api.prototype.fAfterLoad = function(fonts, callback) {
		api.initCollaborativeEditing({});
		api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
			api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
		wb = api.wbModel;
		wb.handlers.add("getSelectionState", function() {
			return null;
		});
		wsView = api.wb.getWorksheet();

		startTests();
	};

	AscCommon.CHistory.prototype.UndoRedoEnd = function () {
		this.TurnOn();
	};

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});

	function openDocument(){
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api._openDocument(binaryData);
		api._openOnClient();
	}

	api.HtmlElement = document.createElement("div");
	var curElem = document.getElementById("editor_sdk");
	curElem.appendChild(api.HtmlElement);
	window["Asc"]["editor"] = api;

	function comparePrintPageSettings (assert, obj1, obj2, desc) {
		for (let i in obj1) {
			if (obj1.hasOwnProperty(i)) {
				if (typeof(obj1[i]) === "object" && (i === "pageRange" || i === "titleRowRange" || i === "titleColRange")) {
					for (let j in obj1[i]) {
						if (obj1[i].hasOwnProperty(j)) {
							assert.strictEqual(obj1[i][j], obj2[i][j], desc + j);
						}
					}
				} else {
					assert.strictEqual(obj1[i], obj2[i], desc + i);
				}
			}
		}
	}

	function updateView () {
		wsView._cleanCache(new Asc.Range(0, 0, wsView.cols.length - 1, wsView.rows.length - 1));
		wsView.changeWorksheet("update", {reinitRanges: true});
	}

	function undoAll() {
		while(AscCommon.History.Index !== -1) {
			AscCommon.History.Undo();
		}
	}

	let wb, ws, wsView;
	function testPrintFileSettings() {
		QUnit.test("Test: open print settings ", function (assert) {
			let printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 1, "Compare pages length 1");
			let page = printPagesData.arrPages[0];
			let referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 700.7800000000003,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 796.2399999999999,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 0.74,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 18,
					r1: 0,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings 1:");

			ws = api.wbModel.aWorksheets[0];
			ws.setColWidth(80, 0, 0);
			ws.setColWidth(50, 1, 1);

			updateView();

			referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 549.2600000000001,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 1029.4999999999998,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 0.58,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 18,
					r1: 0,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 1, "Compare pages length 2");
			page = printPagesData.arrPages[0];

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes width cols: ");

			wsView._setPrintScale(100);
			wsView._changeFitToPage(0, 0);
			updateView();

			printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 4, "Compare pages length 3");
			page = printPagesData.arrPages[0];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:707,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:1000,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:1,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 6,
					r1: 0,
					r2: 45,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page1: ");

			page = printPagesData.arrPages[1];

			referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 707,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 775,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 1,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 7,
					c2: 18,
					r1: 0,
					r2: 45,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page2: ");

			page = printPagesData.arrPages[2];

			referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 240,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 1000,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 1,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 6,
					r1: 46,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page3: ");


			page = printPagesData.arrPages[3];

			referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 240,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 775,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 1,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 7,
					c2: 18,
					r1: 46,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page4: ");

			undoAll();
			updateView();

			printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 1, "Compare pages length 1");
			page = printPagesData.arrPages[0];
			referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 700.7800000000003,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 796.2399999999999,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 0.74,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 18,
					r1: 0,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings 1:");
			AscCommon.History.Clear();
		});
	}

	function testPageBreaksSimple() {
		QUnit.test("Test: page break settings ", function (assert) {
			//change active cell and add page break
			//C5
			wsView.setSelection(new Asc.Range(2, 4, 2, 4));
			api.asc_InsertPageBreak();

			let printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 1, "Compare pages length 1");
			let page = printPagesData.arrPages[0];
			let referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 700.7800000000003,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 796.2399999999999,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 0.74,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 18,
					r1: 0,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings 1:");
			wsView._changeFitToPage(0, 0);
			updateView();


			printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 4, "Compare pages length 2");

			page = printPagesData.arrPages[0];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:43.66,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:67.34,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:0.74,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 1,
					r1: 0,
					r2: 3,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes with scale page1: ");

			page = printPagesData.arrPages[1];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:43.66,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:728.9,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:0.74,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 2,
					c2: 18,
					r1: 0,
					r2: 3,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes with scale page2: ");


			page = printPagesData.arrPages[2];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:657.1200000000002,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:67.34,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:0.74,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 1,
					r1: 4,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes with scale page3: ");

			page = printPagesData.arrPages[3];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:657.1200000000002,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:728.9,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:0.74,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 2,
					c2: 18,
					r1: 4,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes with scale page4: ");

			api.asc_SetPrintScale(null, null, 100);

			updateView();

			printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 6, "Compare pages length 2");

			page = printPagesData.arrPages[0];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:59,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:91,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:1,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 1,
					r1: 0,
					r2: 3,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page1: ");

			page = printPagesData.arrPages[1];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:59,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:985,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:1,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 2,
					c2: 18,
					r1: 0,
					r2: 3,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page2: ");

			page = printPagesData.arrPages[2];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:705,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:91,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:1,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 1,
					r1: 4,
					r2: 48,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page3: ");

			page = printPagesData.arrPages[3];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:705,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:985,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:1,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 2,
					c2: 18,
					r1: 4,
					r2: 48,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page4: ");

			page = printPagesData.arrPages[4];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:183,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:91,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:1,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 1,
					r1: 49,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page5: ");

			page = printPagesData.arrPages[5];

			referenceObj = {
				indexWorksheet:0,
				leftFieldInPx:38.79527559055118,
				pageClipRectHeight:183,
				pageClipRectLeft:37.79527559055118,
				pageClipRectTop:37.79527559055118,
				pageClipRectWidth:985,
				pageGridLines:false,
				pageHeadings:false,
				pageHeight:210,
				pageWidth:297,
				scale:1,
				startOffset:0,
				startOffsetPx:0,
				titleColRange:null,
				titleHeight:0,
				titleRowRange:null,
				titleWidth:0,
				topFieldInPx:38.79527559055118,
				pageRange: {
					c1: 2,
					c2: 18,
					r1: 49,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings changes without scale page6: ");

			undoAll();
			updateView();

			printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 1, "Compare pages length 6 after undo");
			page = printPagesData.arrPages[0];
			referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 700.7800000000003,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 796.2399999999999,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 0.74,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 18,
					r1: 0,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings after undo:");
			AscCommon.History.Clear();
		});
	}

	function testPageBreaksAndTitles() {
		QUnit.test("Test: page break and titles settings ", function (assert) {
			api.asc_SetPrintScale(null, null, 100);
			api.asc_changePrintTitles("$A:$D", "$1:$5", 0);
			wsView.setSelection(new Asc.Range(1, 3, 1, 3));
			api.asc_InsertPageBreak();

			let printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 9, "Compare pages length with print titles");

			let page = printPagesData.arrPages[0];

			let referenceObj = {
				"pageWidth": 297,
				"pageHeight": 210,
				"pageClipRectLeft": 37.79527559055118,
				"pageClipRectTop": 37.79527559055118,
				"pageClipRectWidth": 21,
				"pageClipRectHeight": 45,
				"pageRange": {"c1": 0, "r1": 0, "c2": 0, "r2": 2, "refType1": 3, "refType2": 3},
				"leftFieldInPx": 38.79527559055118,
				"topFieldInPx": 38.79527559055118,
				"pageGridLines": false,
				"pageHeadings": false,
				"indexWorksheet": 0,
				"startOffset": 0,
				"startOffsetPx": 0,
				"scale": 1,
				"titleRowRange": null,
				"titleColRange": null,
				"titleWidth": 0,
				"titleHeight": 0
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings with titles 1:");

			page = printPagesData.arrPages[1];

			referenceObj = {
				"pageWidth": 297,
				"pageHeight": 210,
				"pageClipRectLeft": 37.79527559055118,
				"pageClipRectTop": 37.79527559055118,
				"pageClipRectWidth": 966,
				"pageClipRectHeight": 45,
				"pageRange": {"c1": 1, "r1": 0, "c2": 16, "r2": 2, "refType1": 3, "refType2": 3},
				"leftFieldInPx": 38.79527559055118,
				"topFieldInPx": 38.79527559055118,
				"pageGridLines": false,
				"pageHeadings": false,
				"indexWorksheet": 0,
				"startOffset": 0,
				"startOffsetPx": 0,
				"scale": 1,
				"titleRowRange": null,
				"titleColRange": {"c1": 0, "r1": 0, "c2": 0, "r2": 2, "refType1": 3, "refType2": 3},
				"titleWidth": 21,
				"titleHeight": 0
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings with titles 2:");

			page = printPagesData.arrPages[2];

			referenceObj = {
				"pageWidth": 297,
				"pageHeight": 210,
				"pageClipRectLeft": 37.79527559055118,
				"pageClipRectTop": 37.79527559055118,
				"pageClipRectWidth": 89,
				"pageClipRectHeight": 45,
				"pageRange": {"c1": 17, "r1": 0, "c2": 18, "r2": 2, "refType1": 3, "refType2": 3},
				"leftFieldInPx": 38.79527559055118,
				"topFieldInPx": 38.79527559055118,
				"pageGridLines": false,
				"pageHeadings": false,
				"indexWorksheet": 0,
				"startOffset": 0,
				"startOffsetPx": 0,
				"scale": 1,
				"titleRowRange": null,
				"titleColRange": {"c1": 0, "r1": 0, "c2": 3, "r2": 2, "refType1": 3, "refType2": 3},
				"titleWidth": 177,
				"titleHeight": 0
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings with titles 3:");

			page = printPagesData.arrPages[3];

			referenceObj = {
				"pageWidth": 297,
				"pageHeight": 210,
				"pageClipRectLeft": 37.79527559055118,
				"pageClipRectTop": 37.79527559055118,
				"pageClipRectWidth": 21,
				"pageClipRectHeight": 662,
				"pageRange": {"c1": 0, "r1": 3, "c2": 0, "r2": 45, "refType1": 3, "refType2": 3},
				"leftFieldInPx": 38.79527559055118,
				"topFieldInPx": 38.79527559055118,
				"pageGridLines": false,
				"pageHeadings": false,
				"indexWorksheet": 0,
				"startOffset": 0,
				"startOffsetPx": 0,
				"scale": 1,
				"titleRowRange": {"c1": 0, "r1": 0, "c2": 0, "r2": 2, "refType1": 3, "refType2": 3},
				"titleColRange": null,
				"titleWidth": 0,
				"titleHeight": 45
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings with titles 4:");

			page = printPagesData.arrPages[4];

			referenceObj = {
				"pageWidth": 297,
				"pageHeight": 210,
				"pageClipRectLeft": 37.79527559055118,
				"pageClipRectTop": 37.79527559055118,
				"pageClipRectWidth": 966,
				"pageClipRectHeight": 662,
				"pageRange": {"c1": 1, "r1": 3, "c2": 16, "r2": 45, "refType1": 3, "refType2": 3},
				"leftFieldInPx": 38.79527559055118,
				"topFieldInPx": 38.79527559055118,
				"pageGridLines": false,
				"pageHeadings": false,
				"indexWorksheet": 0,
				"startOffset": 0,
				"startOffsetPx": 0,
				"scale": 1,
				"titleRowRange": {"c1": 1, "r1": 0, "c2": 16, "r2": 2, "refType1": 3, "refType2": 3},
				"titleColRange": {"c1": 0, "r1": 3, "c2": 0, "r2": 45, "refType1": 3, "refType2": 3},
				"titleWidth": 21,
				"titleHeight": 45
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings with titles 5:");

			page = printPagesData.arrPages[5];

			referenceObj = {
				"pageWidth": 297,
				"pageHeight": 210,
				"pageClipRectLeft": 37.79527559055118,
				"pageClipRectTop": 37.79527559055118,
				"pageClipRectWidth": 89,
				"pageClipRectHeight": 662,
				"pageRange": {"c1": 17, "r1": 3, "c2": 18, "r2": 45, "refType1": 3, "refType2": 3},
				"leftFieldInPx": 38.79527559055118,
				"topFieldInPx": 38.79527559055118,
				"pageGridLines": false,
				"pageHeadings": false,
				"indexWorksheet": 0,
				"startOffset": 0,
				"startOffsetPx": 0,
				"scale": 1,
				"titleRowRange": {"c1": 17, "r1": 0, "c2": 18, "r2": 2, "refType1": 3, "refType2": 3},
				"titleColRange": {"c1": 0, "r1": 3, "c2": 3, "r2": 45, "refType1": 3, "refType2": 3},
				"titleWidth": 177,
				"titleHeight": 45
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings with titles 6:");

			page = printPagesData.arrPages[7];

			referenceObj = {
				"pageWidth": 297,
				"pageHeight": 210,
				"pageClipRectLeft": 37.79527559055118,
				"pageClipRectTop": 37.79527559055118,
				"pageClipRectWidth": 966,
				"pageClipRectHeight": 240,
				"pageRange": {"c1": 1, "r1": 46, "c2": 16, "r2": 57, "refType1": 3, "refType2": 3},
				"leftFieldInPx": 38.79527559055118,
				"topFieldInPx": 38.79527559055118,
				"pageGridLines": false,
				"pageHeadings": false,
				"indexWorksheet": 0,
				"startOffset": 0,
				"startOffsetPx": 0,
				"scale": 1,
				"titleRowRange": {"c1": 1, "r1": 0, "c2": 16, "r2": 4, "refType1": 3, "refType2": 3},
				"titleColRange": {"c1": 0, "r1": 46, "c2": 0, "r2": 57, "refType1": 3, "refType2": 3},
				"titleWidth": 21,
				"titleHeight": 73
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings with titles 8:");

			undoAll();
			updateView();

			printPagesData = api.wb.calcPagesPrint(new Asc.asc_CAdjustPrint());
			assert.strictEqual(printPagesData.arrPages.length, 1, "Compare pages length 6 after undo");
			page = printPagesData.arrPages[0];
			referenceObj = {
				indexWorksheet: 0,
				leftFieldInPx: 38.79527559055118,
				pageClipRectHeight: 700.7800000000003,
				pageClipRectLeft: 37.79527559055118,
				pageClipRectTop: 37.79527559055118,
				pageClipRectWidth: 796.2399999999999,
				pageGridLines: false,
				pageHeadings: false,
				pageHeight: 210,
				pageWidth: 297,
				scale: 0.74,
				startOffset: 0,
				startOffsetPx: 0,
				titleColRange: null,
				titleHeight: 0,
				titleRowRange: null,
				titleWidth: 0,
				topFieldInPx: 38.79527559055118,
				pageRange: {
					c1: 0,
					c2: 18,
					r1: 0,
					r2: 57,
					refType1: 3,
					refType2: 3
				}
			};

			comparePrintPageSettings(assert, page, referenceObj, "Compare pages settings after undo:");
			AscCommon.History.Clear();
		});
	}

	function checkUndoRedo(fBefore, fAfter, desc, skipLastUndo) {
		fAfter("after_" + desc);
		AscCommon.History.Undo();
		fBefore("undo_" + desc);
		AscCommon.History.Redo();
		fAfter("redo_" + desc);
		if (!skipLastUndo) {
			AscCommon.History.Undo();
		}
	}

	function testPageBreaksManipulation() {
		QUnit.test("Test: page break manipulation ", function (assert) {
			//add breaks
			ws = api.wbModel.aWorksheets[0];

			let beforeFunc = function(desc) {
				assert.strictEqual((ws.colBreaks == null || ws.colBreaks.getCount() === 0) ? null : 1, null, desc);
				assert.strictEqual((ws.rowBreaks == null || ws.rowBreaks.getCount() === 0) ? null : 1, null, desc);
			};

			let insertColBreakId = 1;
			let insertRowBreakId = 3;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 1, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(insertColBreakId), true, desc + " check col contains");
				assert.strictEqual(ws.rowBreaks.containsBreak(insertRowBreakId), true, desc + " check row contains");
			}, "insert page break_col1row3");


			insertColBreakId = 5;
			insertRowBreakId = 5;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 1, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(insertColBreakId), true, desc + " check col contains");
				assert.strictEqual(ws.rowBreaks.containsBreak(insertRowBreakId), true, desc + " check row contains");
			}, "insert page break_col5row5");



			insertColBreakId = 1;
			insertRowBreakId = 1;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 1, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(insertColBreakId), true, desc + " check col contains");
				assert.strictEqual(ws.rowBreaks.containsBreak(insertRowBreakId), true, desc + " check row contains");
			}, "insert page break_col1row1");

			insertColBreakId = 0;
			insertRowBreakId = 0;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 0, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 0, desc + " check row count");
			}, "insert page break_col0row0");

			insertColBreakId = 0;
			insertRowBreakId = 1;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 0, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc + " check row count");

				assert.strictEqual(ws.rowBreaks.containsBreak(insertRowBreakId), true, desc + " check row contains");
			}, "insert page break_col0row1");

			insertColBreakId = 1;
			insertRowBreakId = 0;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 1, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 0, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(insertColBreakId), true, desc + " check col contains");
			}, "insert page break_col1row0");

			insertColBreakId = 3;
			insertRowBreakId = 3;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(beforeFunc, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 1, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(insertColBreakId), true, desc + " check col contains");
				assert.strictEqual(ws.rowBreaks.containsBreak(insertRowBreakId), true, desc + " check row contains");
			}, "insert page break_col3row3", true);

			//remove
			insertColBreakId = 4;
			insertRowBreakId = 4;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_RemovePageBreak();

			assert.strictEqual(ws.colBreaks.getCount(), 1, " check col count + remove1");
			assert.strictEqual(ws.rowBreaks.getCount(), 1, " check row count + remove1");

			assert.strictEqual(ws.colBreaks.containsBreak(3), true, " check col contains + remove1");
			assert.strictEqual(ws.rowBreaks.containsBreak(3), true, " check row contains + remove1");


			insertColBreakId = 3;
			insertRowBreakId = 3;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_RemovePageBreak();

			checkUndoRedo(function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 1, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(insertColBreakId), true, desc + " check col contains");
				assert.strictEqual(ws.rowBreaks.containsBreak(insertRowBreakId), true, desc + " check row contains");
			}, beforeFunc, "remove page break_col3row3");

			assert.strictEqual(ws.colBreaks.getCount(), 1, " check col count + remove2");
			assert.strictEqual(ws.rowBreaks.getCount(), 1, " check row count + remove2");

			assert.strictEqual(ws.colBreaks.containsBreak(3), true, " check col contains + remove2");
			assert.strictEqual(ws.rowBreaks.containsBreak(3), true, " check row contains + remove2");

			insertColBreakId = 5;
			insertRowBreakId = 5;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(function (desc) {
				assert.strictEqual(ws.colBreaks.getCount(), 1, desc);
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc);
			}, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 2, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 2, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(insertColBreakId), true, desc + " check col contains");
				assert.strictEqual(ws.rowBreaks.containsBreak(insertRowBreakId), true, desc + " check row contains");
			}, "insert page break_col5row5", true);

			//reset all
			api.asc_ResetAllPageBreaks();
			checkUndoRedo(function (desc) {
				assert.strictEqual(ws.colBreaks.getCount(), 2, desc);
				assert.strictEqual(ws.rowBreaks.getCount(), 2, desc);

				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains");
				assert.strictEqual(ws.colBreaks.containsBreak(3), true, desc + " check col contains");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains");

			}, beforeFunc, "remove all page breaks");

			//move
			insertColBreakId = 8;
			insertRowBreakId = 8;
			wsView.setSelection(new Asc.Range(insertColBreakId, insertRowBreakId, insertColBreakId, insertRowBreakId));
			api.asc_InsertPageBreak();

			checkUndoRedo(function (desc) {
				assert.strictEqual(ws.colBreaks.getCount(), 2, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 2, desc + " check row count");
			}, function (desc){
				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");
			}, "insert page break_col8row8", true);

			wsView.changeRowColBreaks(3, 4, new Asc.Range(0, 0, 16, 57), true, true);

			checkUndoRedo(function (desc) {
				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(8), true, desc + " check col contains 8");
				assert.strictEqual(ws.rowBreaks.containsBreak(8), true, desc + " check row contains 8");
				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains 5");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains 5");
				assert.strictEqual(ws.colBreaks.containsBreak(3), true, desc + " check col contains 3");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains 3");
			}, function (desc){

				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(8), true, desc + " check col contains 8");
				assert.strictEqual(ws.rowBreaks.containsBreak(8), true, desc + " check row contains 8");
				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains 5");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains 5");
				assert.strictEqual(ws.colBreaks.containsBreak(4), true, desc + " check col contains 4");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains 4");

			}, "change page col break_from3to4");

			wsView.changeRowColBreaks(3, 12, new Asc.Range(0, 0, 16, 57), true, true);

			checkUndoRedo(function (desc) {
				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(8), true, desc + " check col contains 8");
				assert.strictEqual(ws.rowBreaks.containsBreak(8), true, desc + " check row contains 8");
				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains 5");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains 5");
				assert.strictEqual(ws.colBreaks.containsBreak(3), true, desc + " check col contains 3");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains 3");
			}, function (desc){

				assert.strictEqual(ws.colBreaks.getCount(), 1, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(12), true, desc + " check col contains 12");
				assert.strictEqual(ws.rowBreaks.containsBreak(8), true, desc + " check row contains 8");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains 5");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains 4");

			}, "change page col break_from3to12");

			wsView.changeRowColBreaks(3, 12, new Asc.Range(0, 0, 16, 57), null, true);

			checkUndoRedo(function (desc) {
				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(8), true, desc + " check col contains 8");
				assert.strictEqual(ws.rowBreaks.containsBreak(8), true, desc + " check row contains 8");
				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains 5");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains 5");
				assert.strictEqual(ws.colBreaks.containsBreak(3), true, desc + " check col contains 3");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains 3");
			}, function (desc){

				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc + " check row count");

				assert.strictEqual(ws.rowBreaks.containsBreak(12), true, desc + " check row contains 12");
				assert.strictEqual(ws.colBreaks.containsBreak(8), true, desc + " check col contains 8");
				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains 5");
				assert.strictEqual(ws.colBreaks.containsBreak(3), true, desc + " check col contains 4");

			}, "change page row break_from3to12");


			wsView.changeRowColBreaks(8, 2, new Asc.Range(0, 0, 16, 57), true, true);

			checkUndoRedo(function (desc) {
				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(8), true, desc + " check col contains 8");
				assert.strictEqual(ws.rowBreaks.containsBreak(8), true, desc + " check row contains 8");
				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains 5");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains 5");
				assert.strictEqual(ws.colBreaks.containsBreak(3), true, desc + " check col contains 3");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains 3");
			}, function (desc){

				assert.strictEqual(ws.colBreaks.getCount(), 1, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(2), true, desc + " check col contains 12");
				assert.strictEqual(ws.rowBreaks.containsBreak(8), true, desc + " check row contains 8");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains 5");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains 4");

			}, "change page col break_from8to2");

			wsView.changeRowColBreaks(8, 2, new Asc.Range(0, 0, 16, 57), null, true);

			checkUndoRedo(function (desc) {
				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 3, desc + " check row count");

				assert.strictEqual(ws.colBreaks.containsBreak(8), true, desc + " check col contains 8");
				assert.strictEqual(ws.rowBreaks.containsBreak(8), true, desc + " check row contains 8");
				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains 5");
				assert.strictEqual(ws.rowBreaks.containsBreak(5), true, desc + " check row contains 5");
				assert.strictEqual(ws.colBreaks.containsBreak(3), true, desc + " check col contains 3");
				assert.strictEqual(ws.rowBreaks.containsBreak(3), true, desc + " check row contains 3");
			}, function (desc){

				assert.strictEqual(ws.colBreaks.getCount(), 3, desc + " check col count");
				assert.strictEqual(ws.rowBreaks.getCount(), 1, desc + " check row count");

				assert.strictEqual(ws.rowBreaks.containsBreak(2), true, desc + " check row contains 12");
				assert.strictEqual(ws.colBreaks.containsBreak(8), true, desc + " check col contains 8");
				assert.strictEqual(ws.colBreaks.containsBreak(5), true, desc + " check col contains 5");
				assert.strictEqual(ws.colBreaks.containsBreak(3), true, desc + " check col contains 4");

			}, "change page row break_from8to2");
		});
	}

	QUnit.module("Print");

	function startTests() {
		QUnit.start();
		testPrintFileSettings();
		testPageBreaksSimple();
		testPageBreaksAndTitles();
		testPageBreaksManipulation();
	}



	//test file
	//sdkjs-tests-printtests.xlsx
	let binaryData = new Uint8Array([88,76,83,89,59,118,49,48,59,48,59,8,6,138,30,0,0,7,219,30,0,0,1,186,31,0,0,2,146,41,0,0,3,139,62,0,0,4,187,62,0,0,0,209,142,0,0,10,87,162,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,145,1,0,0,17,0,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,17,0,0,0,2,0,0,0,0,0,0,0,0,0,0,17,1,0,0,0,0,0,0,0,3,1,0,32,0,0,0,0,0,7,14,0,0,0,0,24,0,0,0,1,0,0,0,0,0,1,10,1,0,0,0,24,0,0,0,0,0,1,10,2,0,0,0,24,0,0,0,0,0,1,10,3,0,0,0,24,0,0,0,0,0,1,10,4,0,0,0,24,0,0,0,0,0,1,10,5,0,0,0,24,0,0,0,0,0,1,10,6,0,0,0,24,0,0,0,0,0,1,10,7,0,0,0,24,0,0,0,0,0,1,10,8,0,0,0,24,0,0,0,0,0,1,10,9,0,0,0,24,0,0,0,0,0,1,10,10,0,0,0,24,0,0,0,0,0,1,10,11,0,0,0,24,0,0,0,0,0,1,10,12,0,0,0,24,0,0,0,0,0,1,10,13,0,0,0,24,0,0,0,0,0,1,10,14,0,0,0,3,0,0,0,0,0,1,10,15,0,0,0,3,0,0,0,0,0,7,14,16,0,0,0,3,0,0,0,2,0,0,0,0,0,7,14,17,0,0,0,4,0,0,0,3,0,0,0,0,0,0,17,2,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,0,0,0,0,25,0,0,0,4,0,0,0,0,0,1,10,1,0,0,0,25,0,0,0,0,0,1,10,2,0,0,0,25,0,0,0,0,0,1,10,3,0,0,0,25,0,0,0,0,0,1,10,4,0,0,0,25,0,0,0,0,0,1,10,5,0,0,0,25,0,0,0,0,0,1,10,6,0,0,0,25,0,0,0,0,0,1,10,7,0,0,0,25,0,0,0,0,0,1,10,8,0,0,0,25,0,0,0,0,0,1,10,9,0,0,0,25,0,0,0,0,0,1,10,10,0,0,0,25,0,0,0,0,0,1,10,11,0,0,0,25,0,0,0,0,0,1,10,12,0,0,0,25,0,0,0,0,0,1,10,13,0,0,0,25,0,0,0,0,0,1,10,14,0,0,0,3,0,0,0,0,0,1,10,15,0,0,0,3,0,0,0,0,0,7,14,16,0,0,0,3,0,0,0,5,0,0,0,0,0,7,14,17,0,0,0,5,0,0,0,6,0,0,0,0,0,0,17,3,0,0,0,1,0,0,0,222,0,0,96,0,0,0,0,0,7,14,5,0,0,0,26,0,0,0,7,0,0,0,0,0,1,10,6,0,0,0,26,0,0,0,0,0,1,10,7,0,0,0,26,0,0,0,0,0,1,10,8,0,0,0,26,0,0,0,0,0,1,10,9,0,0,0,26,0,0,0,0,0,1,10,10,0,0,0,26,0,0,0,0,0,1,10,11,0,0,0,26,0,0,0,0,0,1,10,12,0,0,0,26,0,0,0,0,0,1,10,13,0,0,0,26,0,0,0,0,0,1,10,14,0,0,0,3,0,0,0,0,0,1,10,15,0,0,0,3,0,0,0,0,0,7,14,16,0,0,0,3,0,0,0,8,0,0,0,0,0,7,14,17,0,0,0,5,0,0,0,9,0,0,0,0,0,0,17,4,0,0,0,1,0,0,0,222,0,0,96,0,0,0,0,0,7,14,0,0,0,0,1,0,0,0,10,0,0,0,0,0,1,10,5,0,0,0,27,0,0,0,0,0,1,10,6,0,0,0,27,0,0,0,0,0,1,10,7,0,0,0,27,0,0,0,0,0,1,10,8,0,0,0,27,0,0,0,0,0,1,10,9,0,0,0,27,0,0,0,0,0,1,10,10,0,0,0,27,0,0,0,0,0,1,10,11,0,0,0,27,0,0,0,0,0,1,10,12,0,0,0,27,0,0,0,0,0,1,10,13,0,0,0,27,0,0,0,0,0,1,10,14,0,0,0,3,0,0,0,0,0,1,10,15,0,0,0,3,0,0,0,0,0,7,14,16,0,0,0,3,0,0,0,11,0,0,0,0,0,7,14,17,0,0,0,5,0,0,0,12,0,0,0,0,0,0,17,5,0,0,0,1,0,0,0,102,0,0,96,0,0,0,0,0,1,10,14,0,0,0,6,0,0,0,0,0,1,10,15,0,0,0,6,0,0,0,0,0,1,10,16,0,0,0,6,0,0,0,0,0,7,14,17,0,0,0,28,0,0,0,13,0,0,0,0,0,0,17,6,0,0,0,1,0,0,0,222,0,0,96,0,0,0,0,0,7,14,10,0,0,0,30,0,0,0,14,0,0,0,0,0,1,10,11,0,0,0,30,0,0,0,0,0,1,10,12,0,0,0,30,0,0,0,0,0,1,10,13,0,0,0,30,0,0,0,0,0,1,10,14,0,0,0,3,0,0,0,0,0,1,10,15,0,0,0,3,0,0,0,0,0,7,14,16,0,0,0,3,0,0,0,15,0,0,0,0,0,1,10,17,0,0,0,29,0,0,0,0,0,0,17,7,0,0,0,1,0,0,0,222,0,0,96,0,0,0,0,0,1,10,14,0,0,0,3,0,0,0,0,0,1,10,15,0,0,0,3,0,0,0,0,0,7,14,16,0,0,0,3,0,0,0,8,0,0,0,0,0,1,10,17,0,0,0,7,0,0,0,0,0,0,17,8,0,0,0,1,0,0,0,222,0,0,96,0,0,0,0,0,7,14,0,0,0,0,1,0,0,0,16,0,0,0,0,0,7,14,5,0,0,0,27,0,0,0,17,0,0,0,0,0,1,10,6,0,0,0,27,0,0,0,0,0,1,10,7,0,0,0,27,0,0,0,0,0,1,10,8,0,0,0,27,0,0,0,0,0,1,10,9,0,0,0,27,0,0,0,0,0,1,10,10,0,0,0,27,0,0,0,0,0,1,10,11,0,0,0,27,0,0,0,0,0,1,10,12,0,0,0,27,0,0,0,0,0,1,10,13,0,0,0,27,0,0,0,0,0,1,10,14,0,0,0,3,0,0,0,0,0,1,10,15,0,0,0,3,0,0,0,0,0,7,14,16,0,0,0,3,0,0,0,11,0,0,0,0,0,7,14,17,0,0,0,7,0,0,0,12,0,0,0,0,0,0,17,9,0,0,0,1,0,0,0,102,0,0,96,0,0,0,0,0,7,14,17,0,0,0,31,0,0,0,13,0,0,0,0,0,0,17,10,0,0,0,1,0,0,0,222,0,0,96,0,0,0,0,0,7,14,10,0,0,0,30,0,0,0,14,0,0,0,0,0,1,10,11,0,0,0,30,0,0,0,0,0,1,10,12,0,0,0,30,0,0,0,0,0,1,10,13,0,0,0,30,0,0,0,0,0,1,10,14,0,0,0,3,0,0,0,0,0,1,10,15,0,0,0,3,0,0,0,0,0,7,14,16,0,0,0,3,0,0,0,15,0,0,0,0,0,1,10,17,0,0,0,29,0,0,0,0,0,0,17,11,0,0,0,1,0,0,0,102,0,0,96,0,0,0,0,0,0,17,12,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,0,0,0,0,32,0,0,0,18,0,0,0,0,0,1,10,1,0,0,0,32,0,0,0,0,0,1,10,2,0,0,0,32,0,0,0,0,0,1,10,3,0,0,0,32,0,0,0,0,0,1,10,4,0,0,0,32,0,0,0,0,0,1,10,5,0,0,0,32,0,0,0,0,0,1,10,6,0,0,0,32,0,0,0,0,0,1,10,7,0,0,0,32,0,0,0,0,0,1,10,8,0,0,0,32,0,0,0,0,0,1,10,9,0,0,0,32,0,0,0,0,0,1,10,10,0,0,0,32,0,0,0,0,0,1,10,11,0,0,0,32,0,0,0,0,0,1,10,12,0,0,0,32,0,0,0,0,0,1,10,13,0,0,0,32,0,0,0,0,0,1,10,14,0,0,0,32,0,0,0,0,0,1,10,15,0,0,0,32,0,0,0,0,0,1,10,16,0,0,0,32,0,0,0,0,0,1,10,17,0,0,0,32,0,0,0,0,0,0,17,13,0,0,0,1,0,0,0,120,0,0,96,0,0,0,0,0,0,17,14,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,0,0,0,0,33,0,0,0,19,0,0,0,0,0,1,10,1,0,0,0,33,0,0,0,0,0,1,10,2,0,0,0,33,0,0,0,0,0,1,10,3,0,0,0,33,0,0,0,0,0,1,10,4,0,0,0,33,0,0,0,0,0,7,14,5,0,0,0,40,0,0,0,10,0,0,0,0,0,1,10,6,0,0,0,40,0,0,0,0,0,1,10,7,0,0,0,40,0,0,0,0,0,1,10,8,0,0,0,40,0,0,0,0,0,1,10,9,0,0,0,40,0,0,0,0,0,1,10,10,0,0,0,40,0,0,0,0,0,1,10,11,0,0,0,40,0,0,0,0,0,7,14,12,0,0,0,40,0,0,0,16,0,0,0,0,0,1,10,13,0,0,0,40,0,0,0,0,0,1,10,14,0,0,0,40,0,0,0,0,0,1,10,15,0,0,0,40,0,0,0,0,0,1,10,16,0,0,0,40,0,0,0,0,0,1,10,17,0,0,0,40,0,0,0,0,0,0,17,15,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,34,0,0,0,0,0,1,10,1,0,0,0,35,0,0,0,0,0,1,10,2,0,0,0,35,0,0,0,0,0,1,10,3,0,0,0,35,0,0,0,0,0,1,10,4,0,0,0,36,0,0,0,0,0,7,14,5,0,0,0,40,0,0,0,20,0,0,0,0,0,1,10,6,0,0,0,40,0,0,0,0,0,1,10,7,0,0,0,40,0,0,0,0,0,1,10,8,0,0,0,40,0,0,0,0,0,1,10,9,0,0,0,40,0,0,0,0,0,7,14,10,0,0,0,41,0,0,0,21,0,0,0,0,0,1,10,11,0,0,0,41,0,0,0,0,0,7,14,12,0,0,0,40,0,0,0,20,0,0,0,0,0,1,10,13,0,0,0,40,0,0,0,0,0,1,10,14,0,0,0,40,0,0,0,0,0,1,10,15,0,0,0,40,0,0,0,0,0,1,10,16,0,0,0,40,0,0,0,0,0,7,14,17,0,0,0,41,0,0,0,21,0,0,0,0,0,0,17,16,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,37,0,0,0,0,0,1,10,1,0,0,0,38,0,0,0,0,0,1,10,2,0,0,0,38,0,0,0,0,0,1,10,3,0,0,0,38,0,0,0,0,0,1,10,4,0,0,0,39,0,0,0,0,0,7,14,5,0,0,0,40,0,0,0,22,0,0,0,0,0,1,10,6,0,0,0,40,0,0,0,0,0,1,10,7,0,0,0,40,0,0,0,0,0,7,14,8,0,0,0,40,0,0,0,23,0,0,0,0,0,1,10,9,0,0,0,40,0,0,0,0,0,1,10,10,0,0,0,42,0,0,0,0,0,1,10,11,0,0,0,43,0,0,0,0,0,7,14,12,0,0,0,40,0,0,0,22,0,0,0,0,0,1,10,13,0,0,0,40,0,0,0,0,0,7,14,14,0,0,0,40,0,0,0,23,0,0,0,0,0,1,10,15,0,0,0,40,0,0,0,0,0,1,10,16,0,0,0,40,0,0,0,0,0,1,10,17,0,0,0,44,0,0,0,0,0,0,17,17,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,0,0,0,0,40,0,0,0,24,0,0,0,0,0,1,10,1,0,0,0,40,0,0,0,0,0,1,10,2,0,0,0,40,0,0,0,0,0,1,10,3,0,0,0,40,0,0,0,0,0,1,10,4,0,0,0,40,0,0,0,0,0,7,14,5,0,0,0,40,0,0,0,25,0,0,0,0,0,1,10,6,0,0,0,40,0,0,0,0,0,1,10,7,0,0,0,40,0,0,0,0,0,7,14,8,0,0,0,40,0,0,0,26,0,0,0,0,0,1,10,9,0,0,0,40,0,0,0,0,0,7,14,10,0,0,0,40,0,0,0,27,0,0,0,0,0,1,10,11,0,0,0,40,0,0,0,0,0,7,14,12,0,0,0,40,0,0,0,28,0,0,0,0,0,1,10,13,0,0,0,40,0,0,0,0,0,7,14,14,0,0,0,40,0,0,0,29,0,0,0,0,0,1,10,15,0,0,0,40,0,0,0,0,0,1,10,16,0,0,0,40,0,0,0,0,0,7,14,17,0,0,0,8,0,0,0,30,0,0,0,0,0,0,17,18,0,0,0,1,0,0,0,98,4,0,96,0,0,0,0,0,7,14,0,0,0,0,45,0,0,0,55,0,0,0,0,0,1,10,1,0,0,0,45,0,0,0,0,0,1,10,2,0,0,0,45,0,0,0,0,0,1,10,3,0,0,0,45,0,0,0,0,0,1,10,4,0,0,0,45,0,0,0,0,0,7,14,5,0,0,0,46,0,0,0,31,0,0,0,0,0,1,10,6,0,0,0,46,0,0,0,0,0,1,10,7,0,0,0,46,0,0,0,0,0,7,14,8,0,0,0,47,0,0,0,32,0,0,0,0,0,1,10,9,0,0,0,47,0,0,0,0,0,1,10,10,0,0,0,48,0,0,0,0,0,1,10,11,0,0,0,48,0,0,0,0,0,1,10,12,0,0,0,9,0,0,0,0,0,1,10,13,0,0,0,10,0,0,0,0,0,1,10,14,0,0,0,9,0,0,0,0,0,1,10,15,0,0,0,10,0,0,0,0,0,1,10,16,0,0,0,10,0,0,0,0,0,1,10,17,0,0,0,11,0,0,0,0,0,0,17,19,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,12,0,0,0,0,0,1,10,1,0,0,0,12,0,0,0,0,0,1,10,2,0,0,0,12,0,0,0,0,0,1,10,3,0,0,0,13,0,0,0,0,0,7,14,4,0,0,0,13,0,0,0,33,0,0,0,0,0,7,14,5,0,0,0,49,0,0,0,34,0,0,0,0,0,1,10,6,0,0,0,49,0,0,0,0,0,1,10,7,0,0,0,49,0,0,0,0,0,7,14,8,0,0,0,49,0,0,0,34,0,0,0,0,0,1,10,9,0,0,0,49,0,0,0,0,0,1,10,10,0,0,0,50,0,0,0,0,0,1,10,11,0,0,0,50,0,0,0,0,0,7,14,12,0,0,0,49,0,0,0,34,0,0,0,0,0,1,10,13,0,0,0,49,0,0,0,0,0,7,14,14,0,0,0,49,0,0,0,34,0,0,0,0,0,1,10,15,0,0,0,49,0,0,0,0,0,1,10,16,0,0,0,49,0,0,0,0,0,1,10,17,0,0,0,14,0,0,0,0,0,0,17,20,0,0,0,1,0,0,0,139,0,0,96,0,0,0,0,0,0,17,21,0,0,0,1,0,0,0,26,1,0,96,0,0,0,0,0,7,14,5,0,0,0,51,0,0,0,35,0,0,0,0,0,1,10,6,0,0,0,51,0,0,0,0,0,1,10,7,0,0,0,51,0,0,0,0,0,1,10,8,0,0,0,51,0,0,0,0,0,1,10,9,0,0,0,51,0,0,0,0,0,1,10,10,0,0,0,51,0,0,0,0,0,1,10,11,0,0,0,51,0,0,0,0,0,1,10,12,0,0,0,51,0,0,0,0,0,1,10,13,0,0,0,51,0,0,0,0,0,1,10,14,0,0,0,51,0,0,0,0,0,1,10,15,0,0,0,51,0,0,0,0,0,1,10,16,0,0,0,51,0,0,0,0,0,1,10,17,0,0,0,51,0,0,0,0,0,0,17,22,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,5,0,0,0,52,0,0,0,31,0,0,0,0,0,1,10,6,0,0,0,52,0,0,0,0,0,1,10,7,0,0,0,52,0,0,0,0,0,7,14,8,0,0,0,47,0,0,0,32,0,0,0,0,0,1,10,9,0,0,0,47,0,0,0,0,0,1,10,10,0,0,0,48,0,0,0,0,0,1,10,11,0,0,0,48,0,0,0,0,0,1,10,12,0,0,0,15,0,0,0,0,0,1,10,13,0,0,0,10,0,0,0,0,0,1,10,14,0,0,0,9,0,0,0,0,0,1,10,15,0,0,0,10,0,0,0,0,0,1,10,16,0,0,0,10,0,0,0,0,0,1,10,17,0,0,0,11,0,0,0,0,0,0,17,23,0,0,0,1,0,0,0,139,0,0,96,0,0,0,0,0,1,10,5,0,0,0,16,0,0,0,0,0,1,10,6,0,0,0,16,0,0,0,0,0,1,10,7,0,0,0,16,0,0,0,0,0,1,10,8,0,0,0,16,0,0,0,0,0,1,10,9,0,0,0,16,0,0,0,0,0,1,10,10,0,0,0,16,0,0,0,0,0,1,10,11,0,0,0,16,0,0,0,0,0,1,10,12,0,0,0,16,0,0,0,0,0,1,10,13,0,0,0,16,0,0,0,0,0,1,10,14,0,0,0,16,0,0,0,0,0,1,10,15,0,0,0,16,0,0,0,0,0,1,10,16,0,0,0,16,0,0,0,0,0,1,10,17,0,0,0,16,0,0,0,0,0,0,17,24,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,0,17,25,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,0,0,0,0,1,0,0,0,36,0,0,0,0,0,1,10,4,0,0,0,17,0,0,0,0,0,1,10,5,0,0,0,17,0,0,0,0,0,7,14,6,0,0,0,30,0,0,0,37,0,0,0,0,0,1,10,7,0,0,0,30,0,0,0,0,0,1,10,8,0,0,0,30,0,0,0,0,0,1,10,9,0,0,0,30,0,0,0,0,0,1,10,10,0,0,0,30,0,0,0,0,0,1,10,11,0,0,0,30,0,0,0,0,0,1,10,12,0,0,0,30,0,0,0,0,0,1,10,13,0,0,0,30,0,0,0,0,0,1,10,14,0,0,0,30,0,0,0,0,0,1,10,15,0,0,0,30,0,0,0,0,0,1,10,16,0,0,0,30,0,0,0,0,0,1,10,17,0,0,0,30,0,0,0,0,0,1,10,18,0,0,0,30,0,0,0,0,0,0,17,26,0,0,0,1,0,0,0,162,0,0,96,0,0,0,0,0,0,17,27,0,0,0,18,0,0,0,63,1,0,96,0,0,0,0,0,7,14,0,0,0,0,19,0,0,0,38,0,0,0,0,0,1,10,1,0,0,0,19,0,0,0,0,0,1,10,2,0,0,0,19,0,0,0,0,0,1,10,3,0,0,0,19,0,0,0,0,0,7,14,11,0,0,0,53,0,0,0,39,0,0,0,0,0,1,10,12,0,0,0,53,0,0,0,0,0,1,10,13,0,0,0,53,0,0,0,0,0,1,10,14,0,0,0,53,0,0,0,0,0,1,10,15,0,0,0,53,0,0,0,0,0,1,10,16,0,0,0,53,0,0,0,0,0,1,10,17,0,0,0,53,0,0,0,0,0,1,10,18,0,0,0,53,0,0,0,0,0,0,17,28,0,0,0,1,0,0,0,222,0,0,96,0,0,0,0,0,7,14,0,0,0,0,26,0,0,0,40,0,0,0,0,0,1,10,1,0,0,0,26,0,0,0,0,0,1,10,2,0,0,0,26,0,0,0,0,0,1,10,3,0,0,0,26,0,0,0,0,0,1,10,4,0,0,0,26,0,0,0,0,0,1,10,5,0,0,0,26,0,0,0,0,0,1,10,9,0,0,0,26,0,0,0,0,0,1,10,10,0,0,0,26,0,0,0,0,0,7,14,11,0,0,0,26,0,0,0,40,0,0,0,0,0,1,10,12,0,0,0,26,0,0,0,0,0,1,10,13,0,0,0,26,0,0,0,0,0,1,10,14,0,0,0,26,0,0,0,0,0,0,17,29,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,30,0,0,0,0,0,1,10,1,0,0,0,30,0,0,0,0,0,1,10,2,0,0,0,30,0,0,0,0,0,1,10,3,0,0,0,30,0,0,0,0,0,1,10,4,0,0,0,30,0,0,0,0,0,1,10,5,0,0,0,30,0,0,0,0,0,1,10,9,0,0,0,26,0,0,0,0,0,1,10,10,0,0,0,26,0,0,0,0,0,1,10,11,0,0,0,30,0,0,0,0,0,1,10,12,0,0,0,30,0,0,0,0,0,1,10,13,0,0,0,30,0,0,0,0,0,1,10,14,0,0,0,30,0,0,0,0,0,0,17,30,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,0,17,31,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,0,0,0,0,26,0,0,0,41,0,0,0,0,0,1,10,1,0,0,0,26,0,0,0,0,0,1,10,2,0,0,0,26,0,0,0,0,0,1,10,3,0,0,0,26,0,0,0,0,0,1,10,4,0,0,0,26,0,0,0,0,0,1,10,5,0,0,0,26,0,0,0,0,0,1,10,9,0,0,0,26,0,0,0,0,0,1,10,10,0,0,0,26,0,0,0,0,0,7,14,11,0,0,0,26,0,0,0,41,0,0,0,0,0,1,10,12,0,0,0,26,0,0,0,0,0,1,10,13,0,0,0,26,0,0,0,0,0,1,10,14,0,0,0,26,0,0,0,0,0,0,17,32,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,30,0,0,0,0,0,1,10,1,0,0,0,30,0,0,0,0,0,1,10,2,0,0,0,30,0,0,0,0,0,1,10,3,0,0,0,30,0,0,0,0,0,1,10,4,0,0,0,30,0,0,0,0,0,1,10,5,0,0,0,30,0,0,0,0,0,1,10,9,0,0,0,26,0,0,0,0,0,1,10,10,0,0,0,26,0,0,0,0,0,1,10,11,0,0,0,30,0,0,0,0,0,1,10,12,0,0,0,30,0,0,0,0,0,1,10,13,0,0,0,30,0,0,0,0,0,1,10,14,0,0,0,30,0,0,0,0,0,0,17,33,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,0,17,34,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,0,0,0,0,30,0,0,0,42,0,0,0,0,0,1,10,1,0,0,0,30,0,0,0,0,0,1,10,2,0,0,0,30,0,0,0,0,0,1,10,3,0,0,0,30,0,0,0,0,0,1,10,4,0,0,0,30,0,0,0,0,0,1,10,5,0,0,0,30,0,0,0,0,0,7,14,11,0,0,0,30,0,0,0,42,0,0,0,0,0,1,10,12,0,0,0,30,0,0,0,0,0,1,10,13,0,0,0,30,0,0,0,0,0,1,10,14,0,0,0,30,0,0,0,0,0,0,17,35,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,7,14,0,0,0,0,26,0,0,0,43,0,0,0,0,0,1,10,1,0,0,0,26,0,0,0,0,0,1,10,2,0,0,0,26,0,0,0,0,0,1,10,3,0,0,0,26,0,0,0,0,0,1,10,4,0,0,0,26,0,0,0,0,0,1,10,5,0,0,0,26,0,0,0,0,0,1,10,9,0,0,0,26,0,0,0,0,0,1,10,10,0,0,0,26,0,0,0,0,0,0,17,36,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,0,17,37,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,26,0,0,0,0,0,1,10,1,0,0,0,26,0,0,0,0,0,1,10,2,0,0,0,26,0,0,0,0,0,1,10,3,0,0,0,26,0,0,0,0,0,1,10,4,0,0,0,26,0,0,0,0,0,1,10,5,0,0,0,26,0,0,0,0,0,0,17,38,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,0,17,39,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,0,17,40,0,0,0,1,0,0,0,222,0,0,96,0,0,0,0,0,7,14,0,0,0,0,1,0,0,0,44,0,0,0,0,0,7,14,11,0,0,0,30,0,0,0,44,0,0,0,0,0,1,10,12,0,0,0,30,0,0,0,0,0,1,10,13,0,0,0,30,0,0,0,0,0,1,10,14,0,0,0,30,0,0,0,0,0,0,17,41,0,0,0,0,0,0,0,183,1,0,32,0,0,0,0,0,1,10,0,0,0,0,20,0,0,0,0,0,1,10,1,0,0,0,54,0,0,0,0,0,7,14,2,0,0,0,56,0,0,0,45,0,0,0,0,0,1,10,3,0,0,0,56,0,0,0,0,0,1,10,4,0,0,0,56,0,0,0,0,0,1,10,5,0,0,0,56,0,0,0,0,0,1,10,6,0,0,0,56,0,0,0,0,0,0,17,42,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,1,0,0,0,55,0,0,0,0,0,7,14,2,0,0,0,57,0,0,0,46,0,0,0,0,0,1,10,3,0,0,0,57,0,0,0,0,0,1,10,4,0,0,0,57,0,0,0,0,0,1,10,5,0,0,0,57,0,0,0,0,0,1,10,6,0,0,0,57,0,0,0,0,0,0,17,43,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,22,0,0,0,47,0,0,0,0,0,1,10,2,0,0,0,58,0,0,0,0,0,1,10,3,0,0,0,58,0,0,0,0,0,1,10,4,0,0,0,58,0,0,0,0,0,1,10,5,0,0,0,58,0,0,0,0,0,1,10,6,0,0,0,58,0,0,0,0,0,0,17,44,0,0,0,0,0,0,0,224,1,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,22,0,0,0,48,0,0,0,0,0,1,10,2,0,0,0,58,0,0,0,0,0,1,10,3,0,0,0,58,0,0,0,0,0,1,10,4,0,0,0,58,0,0,0,0,0,1,10,5,0,0,0,58,0,0,0,0,0,1,10,6,0,0,0,58,0,0,0,0,0,0,17,45,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,23,0,0,0,49,0,0,0,0,0,7,14,2,0,0,0,59,0,0,0,50,0,0,0,0,0,1,10,3,0,0,0,59,0,0,0,0,0,1,10,4,0,0,0,59,0,0,0,0,0,1,10,5,0,0,0,59,0,0,0,0,0,1,10,6,0,0,0,59,0,0,0,0,0,0,17,46,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,0,17,47,0,0,0,0,0,0,0,183,1,0,32,0,0,0,0,0,1,10,0,0,0,0,20,0,0,0,0,0,1,10,1,0,0,0,54,0,0,0,0,0,7,14,2,0,0,0,56,0,0,0,45,0,0,0,0,0,1,10,3,0,0,0,56,0,0,0,0,0,1,10,4,0,0,0,56,0,0,0,0,0,1,10,5,0,0,0,56,0,0,0,0,0,1,10,6,0,0,0,56,0,0,0,0,0,0,17,48,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,1,0,0,0,55,0,0,0,0,0,7,14,2,0,0,0,57,0,0,0,51,0,0,0,0,0,1,10,3,0,0,0,57,0,0,0,0,0,1,10,4,0,0,0,57,0,0,0,0,0,1,10,5,0,0,0,57,0,0,0,0,0,1,10,6,0,0,0,57,0,0,0,0,0,0,17,49,0,0,0,0,0,0,0,63,1,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,22,0,0,0,47,0,0,0,0,0,1,10,2,0,0,0,58,0,0,0,0,0,1,10,3,0,0,0,58,0,0,0,0,0,1,10,4,0,0,0,58,0,0,0,0,0,1,10,5,0,0,0,58,0,0,0,0,0,1,10,6,0,0,0,58,0,0,0,0,0,0,17,50,0,0,0,0,0,0,0,63,1,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,22,0,0,0,48,0,0,0,0,0,1,10,2,0,0,0,58,0,0,0,0,0,1,10,3,0,0,0,58,0,0,0,0,0,1,10,4,0,0,0,58,0,0,0,0,0,1,10,5,0,0,0,58,0,0,0,0,0,1,10,6,0,0,0,58,0,0,0,0,0,0,17,51,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,23,0,0,0,49,0,0,0,0,0,7,14,2,0,0,0,59,0,0,0,52,0,0,0,0,0,1,10,3,0,0,0,59,0,0,0,0,0,1,10,4,0,0,0,59,0,0,0,0,0,1,10,5,0,0,0,59,0,0,0,0,0,1,10,6,0,0,0,59,0,0,0,0,0,0,17,52,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,0,17,53,0,0,0,0,0,0,0,183,1,0,32,0,0,0,0,0,1,10,0,0,0,0,20,0,0,0,0,0,1,10,1,0,0,0,54,0,0,0,0,0,7,14,2,0,0,0,56,0,0,0,45,0,0,0,0,0,1,10,3,0,0,0,56,0,0,0,0,0,1,10,4,0,0,0,56,0,0,0,0,0,1,10,5,0,0,0,56,0,0,0,0,0,1,10,6,0,0,0,56,0,0,0,0,0,0,17,54,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,1,0,0,0,55,0,0,0,0,0,7,14,2,0,0,0,57,0,0,0,53,0,0,0,0,0,1,10,3,0,0,0,57,0,0,0,0,0,1,10,4,0,0,0,57,0,0,0,0,0,1,10,5,0,0,0,57,0,0,0,0,0,1,10,6,0,0,0,57,0,0,0,0,0,0,17,55,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,22,0,0,0,47,0,0,0,0,0,1,10,2,0,0,0,58,0,0,0,0,0,1,10,3,0,0,0,58,0,0,0,0,0,1,10,4,0,0,0,58,0,0,0,0,0,1,10,5,0,0,0,58,0,0,0,0,0,1,10,6,0,0,0,58,0,0,0,0,0,0,17,56,0,0,0,0,0,0,0,130,2,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,22,0,0,0,48,0,0,0,0,0,1,10,2,0,0,0,58,0,0,0,0,0,1,10,3,0,0,0,58,0,0,0,0,0,1,10,4,0,0,0,58,0,0,0,0,0,1,10,5,0,0,0,58,0,0,0,0,0,1,10,6,0,0,0,58,0,0,0,0,0,0,17,57,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,1,10,0,0,0,0,21,0,0,0,0,0,7,14,1,0,0,0,23,0,0,0,49,0,0,0,0,0,7,14,2,0,0,0,59,0,0,0,54,0,0,0,0,0,1,10,3,0,0,0,59,0,0,0,0,0,1,10,4,0,0,0,59,0,0,0,0,0,1,10,5,0,0,0,59,0,0,0,0,0,1,10,6,0,0,0,59,0,0,0,0,0,0,17,58,0,0,0,0,0,0,0,222,0,0,32,0,0,0,0,0,146,1,0,1,76,0,0,0,250,1,15,0,0,0,77,0,105,0,99,0,114,0,111,0,115,0,111,0,102,0,116,0,32,0,69,0,120,0,99,0,101,0,108,0,4,7,0,0,0,49,0,53,0,46,0,48,0,51,0,48,0,48,0,12,0,13,0,14,0,15,0,251,0,7,0,0,0,250,18,0,0,0,0,251,2,218,0,0,0,250,1,16,0,0,0,31,4,48,4,61,4,66,4,78,4,69,4,62,4,50,4,48,4,32,0,28,4,48,4,64,4,56,4,61,4,48,4,2,16,0,0,0,31,4,48,4,61,4,66,4,78,4,58,4,62,4,50,4,48,4,32,0,28,4,48,4,64,4,56,4,61,4,48,4,4,20,0,0,0,50,0,48,0,50,0,50,0,45,0,49,0,50,0,45,0,50,0,49,0,84,0,48,0,54,0,58,0,49,0,57,0,58,0,51,0,55,0,90,0,5,20,0,0,0,50,0,48,0,50,0,50,0,45,0,49,0,50,0,45,0,50,0,49,0,84,0,48,0,54,0,58,0,49,0,57,0,58,0,51,0,55,0,90,0,251,0,47,0,0,0,250,12,20,0,0,0,50,0,48,0,50,0,50,0,45,0,49,0,50,0,45,0,50,0,49,0,84,0,48,0,54,0,58,0,49,0,57,0,58,0,50,0,51,0,90,0,251,212,9,0,0,0,13,0,0,0,3,8,0,0,0,26,4,30,4,20,4,43,4,0,51,0,0,0,3,46,0,0,0,24,4,23,4,18,4,21,4,41,4,21,4,29,4,24,4,21,4,32,0,22,33,32,0,48,0,48,0,24,4,24,4,45,0,48,0,48,0,48,0,48,0,54,0,51,0,0,31,0,0,0,3,26,0,0,0,36,4,62,4,64,4,60,4,48,4,32,0,63,4,62,4,32,0,30,4,26,4,35,4,20,4,0,19,0,0,0,3,14,0,0,0,48,0,53,0,48,0,52,0,56,0,48,0,53,0,0,39,0,0,0,3,34,0,0,0,62,4,66,4,32,0,50,0,53,0,32,0,61,4,62,4,79,4,49,4,64,4,79,4,32,0,50,0,48,0,50,0,50,0,0,13,0,0,0,3,8,0,0,0,20,4,48,4,66,4,48,4,0,25,0,0,0,3,20,0,0,0,50,0,53,0,46,0,49,0,49,0,46,0,50,0,48,0,50,0,50,0,0,61,0,0,0,3,56,0,0,0,36,4,21,4,20,4,21,4,32,4,16,4,27,4,44,4,29,4,16,4,47,4,32,0,31,4,32,4,30,4,17,4,24,4,32,4,29,4,16,4,47,4,32,0,31,4,16,4,27,4,16,4,34,4,16,4,0,19,0,0,0,3,14,0,0,0,63,4,62,4,32,0,30,4,26,4,31,4,30,4,0,21,0,0,0,3,16,0,0,0,55,0,53,0,55,0,51,0,53,0,53,0,53,0,56,0,0,53,0,0,0,3,48,0,0,0,35,4,71,4,64,4,53,4,54,4,52,4,53,4,61,4,56,4,53,4,32,0,40,0,62,4,66,4,63,4,64,4,48,4,50,4,56,4,66,4,53,4,59,4,76,4,41,0,0,27,0,0,0,3,22,0,0,0,51,4,59,4,48,4,50,4,75,4,32,0,63,4,62,4,32,0,17,4,26,4,0,11,0,0,0,3,6,0,0,0,49,0,52,0,53,0,0,23,0,0,0,3,18,0,0,0,55,0,55,0,48,0,51,0,48,0,49,0,48,0,48,0,49,0,0,53,0,0,0,3,48,0,0,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,32,0,24,4,29,4,29,4,0,11,0,0,0,3,6,0,0,0,26,4,31,4,31,4,0,51,0,0,0,3,46,0,0,0,35,4,71,4,64,4,53,4,54,4,52,4,53,4,61,4,56,4,53,4,32,0,40,0,63,4,62,4,59,4,67,4,71,4,48,4,66,4,53,4,59,4,76,4,41,0,0,187,0,0,0,3,182,0,0,0,28,4,53,4,54,4,64,4,53,4,51,4,56,4,62,4,61,4,48,4,59,4,76,4,61,4,62,4,53,4,32,0,67,4,63,4,64,4,48,4,50,4,59,4,53,4,61,4,56,4,53,4,32,0,36,4,53,4,52,4,53,4,64,4,48,4,59,4,76,4,61,4,62,4,57,4,32,0,63,4,64,4,62,4,49,4,56,4,64,4,61,4,62,4,57,4,32,0,63,4,48,4,59,4,48,4,66,4,75,4,32,0,63,4,62,4,32,0,38,4,53,4,61,4,66,4,64,4,48,4,59,4,76,4,61,4,62,4,60,4,67,4,32,0,68,4,53,4,52,4,53,4,64,4,48,4,59,4,76,4,61,4,62,4,60,4,67,4,32,0,62,4,58,4,64,4,67,4,51,4,67,4,0,205,0,0,0,3,200,0,0,0,32,0,32,0,29,4,48,4,65,4,66,4,62,4,79,4,73,4,56,4,60,4,32,0,63,4,62,4,52,4,66,4,50,4,53,4,64,4,54,4,52,4,48,4,53,4,66,4,65,4,79,4,32,0,62,4,65,4,67,4,73,4,53,4,65,4,66,4,50,4,59,4,53,4,61,4,56,4,53,4,32,0,64,4,48,4,65,4,71,4,53,4,66,4,62,4,50,4,32,0,60,4,53,4,54,4,52,4,67,4,32,0,67,4,71,4,64,4,53,4,54,4,52,4,53,4,61,4,56,4,79,4,60,4,56,4,32,0,65,4,32,0,62,4,66,4,64,4,48,4,54,4,53,4,61,4,56,4,53,4,60,4,32,0,65,4,59,4,53,4,52,4,67,4,78,4,73,4,56,4,69,4,32,0,55,4,48,4,63,4,56,4,65,4,53,4,57,4,58,0,0,39,0,0,0,3,34,0,0,0,33,4,62,4,52,4,53,4,64,4,54,4,48,4,61,4,56,4,53,4,32,0,55,4,48,4,63,4,56,4,65,4,56,4,0,27,0,0,0,3,22,0,0,0,61,4,62,4,60,4,53,4,64,4,32,0,65,4,71,4,53,4,66,4,48,4,0,27,0,0,0,3,22,0,0,0,33,4,67,4,60,4,60,4,48,4,44,0,32,0,64,4,67,4,49,4,46,0,0,15,0,0,0,3,10,0,0,0,52,4,53,4,49,4,53,4,66,4,0,17,0,0,0,3,12,0,0,0,58,4,64,4,53,4,52,4,56,4,66,4,0,7,0,0,0,3,2,0,0,0,49,0,0,7,0,0,0,3,2,0,0,0,50,0,0,7,0,0,0,3,2,0,0,0,51,0,0,7,0,0,0,3,2,0,0,0,52,0,0,7,0,0,0,3,2,0,0,0,53,0,0,7,0,0,0,3,2,0,0,0,54,0,0,7,0,0,0,3,2,0,0,0,55,0,0,65,0,0,0,3,60,0,0,0,48,0,48,0,48,0,48,0,48,0,48,0,48,0,48,0,48,0,48,0,48,0,48,0,48,0,48,0,56,0,48,0,50,0,46,0,49,0,46,0,51,0,48,0,52,0,46,0,48,0,52,0,46,0,51,0,52,0,54,0,0,65,0,0,0,3,60,0,0,0,48,0,49,0,49,0,51,0,51,0,57,0,52,0,49,0,49,0,57,0,48,0,48,0,49,0,57,0,50,0,52,0,52,0,46,0,49,0,46,0,51,0,48,0,50,0,46,0,51,0,52,0,46,0,55,0,51,0,52,0,0,15,0,0,0,3,10,0,0,0,24,4,66,4,62,4,51,4,62,4,0,7,0,0,0,3,2,0,0,0,37,4,0,55,0,0,0,3,50,0,0,0,30,4,49,4,62,4,64,4,62,4,66,4,75,4,32,0,50,4,32,0,54,4,67,4,64,4,61,4,48,4,59,4,32,0,62,4,63,4,53,4,64,4,48,4,70,4,56,4,57,4,0,27,0,0,0,3,22,0,0,0,31,4,64,4,56,4,59,4,62,4,54,4,53,4,61,4,56,4,53,4,58,0,0,33,0,0,0,3,28,0,0,0,32,0,32,0,32,0,32,0,52,4,62,4,58,4,67,4,60,4,53,4,61,4,66,4,62,4,50,4,0,27,0,0,0,3,22,0,0,0,30,4,66,4,63,4,64,4,48,4,50,4,56,4,66,4,53,4,59,4,76,4,0,25,0,0,0,3,20,0,0,0,31,4,62,4,59,4,67,4,71,4,48,4,66,4,53,4,59,4,76,4,0,95,0,0,0,3,90,0,0,0,32,4,67,4,58,4,62,4,50,4,62,4,52,4,56,4,66,4,53,4,59,4,76,4,32,0,67,4,71,4,64,4,53,4,54,4,52,4,53,4,61,4,56,4,79,4,10,0,40,0,67,4,63,4,62,4,59,4,61,4,62,4,60,4,62,4,71,4,53,4,61,4,61,4,62,4,53,4,32,0,59,4,56,4,70,4,62,4,41,0,0,105,0,0,0,3,100,0,0,0,19,4,59,4,48,4,50,4,61,4,75,4,57,4,32,0,49,4,67,4,69,4,51,4,48,4,59,4,66,4,53,4,64,4,32,0,67,4,71,4,64,4,53,4,54,4,52,4,53,4,61,4,56,4,79,4,10,0,40,0,67,4,63,4,62,4,59,4,61,4,62,4,60,4,62,4,71,4,53,4,61,4,61,4,62,4,53,4,32,0,59,4,56,4,70,4,62,4,41,0,0,27,0,0,0,3,22,0,0,0,24,4,65,4,63,4,62,4,59,4,61,4,56,4,66,4,53,4,59,4,76,4,0,37,0,0,0,3,32,0,0,0,19,4,59,4,48,4,50,4,61,4,75,4,57,4,32,0,58,4,48,4,55,4,61,4,48,4,71,4,53,4,57,4,0,69,0,0,0,3,64,0,0,0,34,0,95,0,95,0,95,0,95,0,95,0,34,0,32,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,95,0,32,0,50,0,48,0,95,0,95,0,95,0,32,0,51,4,46,0,0,81,0,0,0,3,76,0,0,0,20,4,30,4,26,4,35,4,28,4,21,4,29,4,34,4,32,0,31,4,30,4,20,4,31,4,24,4,33,4,16,4,29,4,10,0,45,4,27,4,21,4,26,4,34,4,32,4,30,4,29,4,29,4,30,4,25,4,32,0,31,4,30,4,20,4,31,4,24,4,33,4,44,4,46,4,0,43,0,0,0,3,38,0,0,0,48,0,50,0,46,0,49,0,50,0,46,0,50,0,48,0,50,0,50,0,32,0,49,0,50,0,58,0,48,0,53,0,58,0,49,0,56,0,0,29,0,0,0,3,24,0,0,0,32,0,33,4,53,4,64,4,66,4,56,4,68,4,56,4,58,4,48,4,66,4,58,0,0,25,0,0,0,3,20,0,0,0,32,0,18,4,59,4,48,4,52,4,53,4,59,4,53,4,70,4,58,0,0,33,0,0,0,3,28,0,0,0,32,0,20,4,53,4,57,4,65,4,66,4,50,4,56,4,66,4,53,4,59,4,53,4,61,4,58,0,0,57,0,0,0,3,52,0,0,0,65,4,32,0,49,0,56,0,46,0,48,0,53,0,46,0,50,0,48,0,50,0,50,0,32,0,63,4,62,4,32,0,49,0,49,0,46,0,48,0,56,0,46,0,50,0,48,0,50,0,51,0,0,43,0,0,0,3,38,0,0,0,48,0,50,0,46,0,49,0,50,0,46,0,50,0,48,0,50,0,50,0,32,0,49,0,51,0,58,0,51,0,49,0,58,0,48,0,54,0,0,57,0,0,0,3,52,0,0,0,65,4,32,0,48,0,56,0,46,0,48,0,57,0,46,0,50,0,48,0,50,0,49,0,32,0,63,4,62,4,32,0,48,0,56,0,46,0,49,0,50,0,46,0,50,0,48,0,50,0,50,0,0,43,0,0,0,3,38,0,0,0,48,0,54,0,46,0,49,0,50,0,46,0,50,0,48,0,50,0,50,0,32,0,49,0,48,0,58,0,51,0,49,0,58,0,51,0,57,0,0,57,0,0,0,3,52,0,0,0,65,4,32,0,48,0,53,0,46,0,48,0,57,0,46,0,50,0,48,0,50,0,50,0,32,0,63,4,62,4,32,0,50,0,57,0,46,0,49,0,49,0,46,0,50,0,48,0,50,0,51,0,0,29,0,0,0,3,24,0,0,0,49,0,46,0,32,0,31,4,62,4,65,4,66,4,48,4,50,4,58,4,48,4,32,0,245,20,0,0,0,219,6,0,0,1,25,0,0,0,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,4,0,0,0,0,5,0,0,0,0,1,85,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,85,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,1,85,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,40,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,0,0,0,0,4,0,0,0,0,5,0,0,0,0,1,70,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,5,0,0,0,0,1,70,0,0,0,0,0,0,0,0,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,85,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,55,0,0,0,0,0,0,0,0,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,5,0,0,0,0,1,40,0,0,0,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,5,0,0,0,0,1,40,0,0,0,0,0,0,0,0,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,4,0,0,0,0,5,0,0,0,0,1,55,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,0,0,0,0,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,5,0,0,0,0,1,55,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,4,0,0,0,0,5,0,0,0,0,1,70,0,0,0,0,0,0,0,0,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,70,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,5,0,0,0,0,1,70,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,0,0,0,0,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,70,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,0,0,0,0,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,55,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,0,0,0,0,4,0,0,0,0,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,40,0,0,0,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,4,0,0,0,0,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,40,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,1,0,0,0,0,2,0,0,0,0,4,0,0,0,0,5,0,0,0,0,1,70,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,0,0,0,0,2,0,0,0,0,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,13,1,40,0,0,0,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,4,0,0,0,0,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,6,1,40,0,0,0,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,5,0,0,0,0,1,55,0,0,0,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,1,55,0,0,0,0,0,0,0,0,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,4,0,0,0,0,5,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,1,40,0,0,0,0,0,0,0,0,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,4,0,0,0,0,5,0,0,0,0,1,55,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,1,0,0,0,0,2,0,0,0,0,4,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,5,0,0,0,0,1,55,0,0,0,0,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,1,0,0,0,0,2,15,0,0,0,0,6,6,0,0,0,0,4,0,0,0,255,1,1,12,4,0,0,0,0,5,0,0,0,0,4,32,0,0,0,5,11,0,0,0,0,6,0,0,0,2,1,0,0,0,17,5,11,0,0,0,0,6,0,0,0,2,1,0,0,0,8,6,167,0,0,0,7,26,0,0,0,4,6,10,0,0,0,65,0,114,0,105,0,97,0,108,0,6,5,0,0,0,0,0,0,32,64,7,29,0,0,0,0,1,1,4,6,10,0,0,0,65,0,114,0,105,0,97,0,108,0,6,5,0,0,0,0,0,0,34,64,7,29,0,0,0,0,1,1,4,6,10,0,0,0,65,0,114,0,105,0,97,0,108,0,6,5,0,0,0,0,0,0,32,64,7,32,0,0,0,0,1,1,3,1,1,4,6,10,0,0,0,65,0,114,0,105,0,97,0,108,0,6,5,0,0,0,0,0,0,34,64,7,26,0,0,0,4,6,10,0,0,0,65,0,114,0,105,0,97,0,108,0,6,5,0,0,0,0,0,0,24,64,14,29,0,0,0,3,24,0,0,0,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,2,90,12,0,0,3,30,0,0,0,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,3,42,0,0,0,0,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,6,3,48,0,0,0,0,1,1,1,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,48,0,0,0,0,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,7,1,4,1,0,0,0,3,45,0,0,0,0,1,1,1,1,1,6,4,2,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,45,0,0,0,0,1,1,1,1,1,6,4,3,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,42,0,0,0,0,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,7,3,45,0,0,0,0,1,1,1,1,1,6,4,7,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,48,0,0,0,0,1,1,1,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,4,3,51,0,0,0,0,1,1,1,1,1,6,4,15,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,7,7,1,4,8,1,1,3,51,0,0,0,0,1,1,1,1,1,6,4,17,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,7,7,1,4,8,1,1,3,48,0,0,0,0,1,1,1,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,7,7,1,4,3,45,0,0,0,0,1,1,1,1,1,6,4,18,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,6,3,45,0,0,0,0,1,1,1,1,1,6,4,18,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,7,3,45,0,0,0,0,1,1,1,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,7,3,51,0,0,0,0,1,1,1,1,1,6,4,20,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,7,7,1,4,8,1,1,3,45,0,0,0,0,1,1,1,1,1,6,4,21,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,6,3,45,0,0,0,0,1,1,1,1,1,6,4,4,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,45,0,0,0,0,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,7,1,4,3,48,0,0,0,0,1,1,3,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,3,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,7,1,4,3,48,0,0,0,0,1,1,3,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,7,1,4,3,54,0,0,0,0,1,1,3,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,12,0,0,0,0,1,6,1,4,1,0,0,0,7,1,4,3,57,0,0,0,0,1,1,1,1,1,3,1,1,6,4,22,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,12,0,0,0,0,1,6,1,4,1,0,0,0,7,1,4,3,57,0,0,0,0,1,1,1,1,1,3,1,1,6,4,26,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,12,0,0,0,0,1,6,1,4,1,0,0,0,7,1,4,3,45,0,0,0,0,1,1,3,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,1,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,45,0,0,0,0,1,1,3,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,2,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,45,0,0,0,0,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,8,1,1,3,48,0,0,0,0,1,1,1,1,1,6,4,4,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,8,1,1,3,45,0,0,0,0,1,1,1,1,1,6,4,6,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,45,0,0,0,0,1,1,1,1,1,6,4,5,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,42,0,0,0,0,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,6,3,45,0,0,0,0,1,1,1,1,1,6,4,8,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,45,0,0,0,0,1,1,3,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,2,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,6,3,48,0,0,0,0,1,1,1,1,1,6,4,13,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,48,0,0,0,0,1,1,1,1,1,6,4,9,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,45,0,0,0,0,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,48,0,0,0,0,1,1,1,1,1,6,4,10,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,48,0,0,0,0,1,1,1,1,1,6,4,11,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,48,0,0,0,0,1,1,1,1,1,6,4,4,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,48,0,0,0,0,1,1,1,1,1,6,4,12,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,48,0,0,0,0,1,1,1,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,4,3,48,0,0,0,0,1,1,1,1,1,6,4,13,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,4,3,48,0,0,0,0,1,1,1,1,1,6,4,11,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,4,3,48,0,0,0,0,1,1,1,1,1,6,4,12,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,4,3,48,0,0,0,0,1,1,1,1,1,6,4,14,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,4,3,51,0,0,0,0,1,1,1,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,6,7,1,4,8,1,1,3,51,0,0,0,0,1,1,1,1,1,6,4,15,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,7,7,1,4,8,1,1,3,51,0,0,0,0,1,1,1,1,1,6,4,16,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,7,7,1,4,8,1,1,3,51,0,0,0,0,1,1,1,1,1,4,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,4,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,7,7,1,4,3,48,0,0,0,0,1,1,1,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,7,1,1,3,48,0,0,0,0,1,1,1,1,1,4,1,1,6,4,1,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,4,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,7,3,48,0,0,0,0,1,1,1,1,1,6,4,19,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,7,1,1,3,51,0,0,0,0,1,1,1,1,1,6,4,20,0,0,0,7,4,0,0,0,0,8,4,0,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,7,7,1,4,8,1,1,3,48,0,0,0,0,1,1,3,1,1,6,4,0,0,0,0,7,4,0,0,0,0,8,4,3,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,7,1,4,3,51,0,0,0,0,1,1,1,1,1,3,1,1,6,4,23,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,7,1,4,3,51,0,0,0,0,1,1,1,1,1,3,1,1,6,4,22,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,6,7,1,4,3,51,0,0,0,0,1,1,1,1,1,3,1,1,6,4,24,0,0,0,7,4,0,0,0,0,8,4,2,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,6,0,0,0,0,1,0,8,1,1,3,48,0,0,0,0,1,1,1,1,1,3,1,1,6,4,25,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,3,0,0,0,0,1,0,3,54,0,0,0,0,1,1,1,1,1,3,1,1,6,4,25,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,6,7,1,4,8,1,1,3,54,0,0,0,0,1,1,1,1,1,3,1,1,6,4,27,0,0,0,7,4,0,0,0,0,8,4,4,0,0,0,9,4,0,0,0,0,12,4,0,0,0,0,13,6,9,0,0,0,0,1,6,7,1,4,8,1,1,15,42,0,0,0,16,37,0,0,0,0,4,0,0,0,0,0,0,0,4,14,0,0,0,30,4,49,4,75,4,71,4,61,4,75,4,57,4,5,4,0,0,0,0,0,0,0,10,0,0,0,0,12,78,0,0,0,0,34,0,0,0,84,0,97,0,98,0,108,0,101,0,83,0,116,0,121,0,108,0,101,0,77,0,101,0,100,0,105,0,117,0,109,0,57,0,1,34,0,0,0,80,0,105,0,118,0,111,0,116,0,83,0,116,0,121,0,108,0,101,0,76,0,105,0,103,0,104,0,116,0,49,0,54,0,17,55,0,0,0,0,50,0,0,0,250,0,17,0,0,0,83,0,108,0,105,0,99,0,101,0,114,0,83,0,116,0,121,0,108,0,101,0,76,0,105,0,103,0,104,0,116,0,49,0,251,0,4,0,0,0,0,0,0,0,44,0,0,0,0,0,0,0,0,1,5,0,0,0,2,0,0,0,0,15,15,0,0,0,0,4,0,0,0,103,230,1,0,3,1,0,0,0,1,17,4,0,0,0,120,0,108,0,18,80,0,0,0,13,80,0,0,1,26,0,0,0,0,6,14,0,0,0,84,0,68,0,83,0,104,0,101,0,101,0,116,0,1,4,1,0,0,0,2,172,2,0,0,3,31,0,0,0,2,4,1,0,0,0,3,4,1,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,0,12,64,6,1,1,3,31,0,0,0,2,4,2,0,0,0,3,4,2,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,84,39,64,6,1,1,3,31,0,0,0,2,4,3,0,0,0,3,4,3,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,0,18,64,6,1,1,3,31,0,0,0,2,4,4,0,0,0,3,4,4,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,170,35,64,6,1,1,3,31,0,0,0,2,4,5,0,0,0,3,4,5,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,80,9,64,6,1,1,3,31,0,0,0,2,4,6,0,0,0,3,4,6,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,168,24,64,6,1,1,3,31,0,0,0,2,4,7,0,0,0,3,4,7,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,170,38,64,6,1,1,3,31,0,0,0,2,4,8,0,0,0,3,4,8,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,84,42,64,6,1,1,3,31,0,0,0,2,4,9,0,0,0,3,4,9,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,168,2,64,6,1,1,3,31,0,0,0,2,4,10,0,0,0,3,4,10,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,170,59,64,6,1,1,3,31,0,0,0,2,4,11,0,0,0,3,4,11,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,168,14,64,6,1,1,3,31,0,0,0,2,4,12,0,0,0,3,4,12,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,0,37,64,6,1,1,3,31,0,0,0,2,4,13,0,0,0,3,4,13,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,84,44,64,6,1,1,3,31,0,0,0,2,4,14,0,0,0,3,4,14,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,170,44,64,6,1,1,3,31,0,0,0,2,4,15,0,0,0,3,4,15,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,80,245,63,6,1,1,3,31,0,0,0,2,4,16,0,0,0,3,4,16,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,0,39,64,6,1,1,3,31,0,0,0,2,4,17,0,0,0,3,4,17,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,0,47,64,6,1,1,3,31,0,0,0,2,4,18,0,0,0,3,4,18,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,84,45,64,6,1,1,3,31,0,0,0,2,4,19,0,0,0,3,4,19,0,0,0,4,4,1,0,0,0,5,5,0,0,0,0,0,0,197,63,6,1,1,4,12,0,0,0,65,0,49,0,58,0,83,0,53,0,57,0,22,47,0,0,0,23,42,0,0,0,10,1,0,0,0,1,14,4,0,0,0,0,0,0,0,20,22,0,0,0,0,6,0,0,0,78,0,53,0,49,0,2,6,0,0,0,78,0,53,0,49,0,11,23,0,0,0,0,5,0,0,0,0,0,0,37,64,1,5,102,102,102,102,102,230,38,64,3,1,1,14,60,0,0,0,0,5,0,0,0,0,0,0,36,64,1,5,0,0,0,0,0,0,36,64,2,5,0,0,0,0,0,0,36,64,3,5,0,0,0,0,0,0,36,64,4,5,0,0,0,0,0,0,0,0,5,5,0,0,0,0,0,0,0,0,15,12,0,0,0,0,1,0,11,1,1,15,4,83,0,0,0,7,45,5,0,0,8,14,0,0,0,67,0,53,0,54,0,58,0,71,0,53,0,54,0,8,14,0,0,0,67,0,53,0,55,0,58,0,71,0,53,0,55,0,8,14,0,0,0,67,0,53,0,56,0,58,0,71,0,53,0,56,0,8,14,0,0,0,67,0,53,0,48,0,58,0,71,0,53,0,48,0,8,14,0,0,0,67,0,53,0,49,0,58,0,71,0,53,0,49,0,8,14,0,0,0,67,0,53,0,50,0,58,0,71,0,53,0,50,0,8,14,0,0,0,66,0,53,0,52,0,58,0,66,0,53,0,53,0,8,14,0,0,0,67,0,53,0,52,0,58,0,71,0,53,0,52,0,8,14,0,0,0,67,0,53,0,53,0,58,0,71,0,53,0,53,0,8,14,0,0,0,67,0,52,0,53,0,58,0,71,0,52,0,53,0,8,14,0,0,0,67,0,52,0,54,0,58,0,71,0,52,0,54,0,8,14,0,0,0,66,0,52,0,56,0,58,0,66,0,52,0,57,0,8,14,0,0,0,67,0,52,0,56,0,58,0,71,0,52,0,56,0,8,14,0,0,0,67,0,52,0,57,0,58,0,71,0,52,0,57,0,8,14,0,0,0,76,0,52,0,49,0,58,0,79,0,52,0,49,0,8,14,0,0,0,66,0,52,0,50,0,58,0,66,0,52,0,51,0,8,14,0,0,0,67,0,52,0,50,0,58,0,71,0,52,0,50,0,8,14,0,0,0,67,0,52,0,51,0,58,0,71,0,52,0,51,0,8,14,0,0,0,67,0,52,0,52,0,58,0,71,0,52,0,52,0,8,14,0,0,0,65,0,51,0,53,0,58,0,70,0,51,0,53,0,8,14,0,0,0,76,0,51,0,53,0,58,0,79,0,51,0,53,0,8,14,0,0,0,65,0,51,0,54,0,58,0,70,0,51,0,54,0,8,14,0,0,0,74,0,51,0,54,0,58,0,75,0,51,0,54,0,8,14,0,0,0,65,0,51,0,56,0,58,0,70,0,51,0,56,0,8,14,0,0,0,76,0,50,0,56,0,58,0,83,0,50,0,56,0,8,14,0,0,0,65,0,50,0,57,0,58,0,70,0,51,0,48,0,8,14,0,0,0,74,0,50,0,57,0,58,0,75,0,51,0,48,0,8,14,0,0,0,76,0,50,0,57,0,58,0,79,0,51,0,48,0,8,14,0,0,0,65,0,51,0,50,0,58,0,70,0,51,0,51,0,8,14,0,0,0,74,0,51,0,50,0,58,0,75,0,51,0,51,0,8,14,0,0,0,76,0,51,0,50,0,58,0,79,0,51,0,51,0,8,14,0,0,0,70,0,50,0,50,0,58,0,82,0,50,0,50,0,8,14,0,0,0,70,0,50,0,51,0,58,0,72,0,50,0,51,0,8,14,0,0,0,73,0,50,0,51,0,58,0,74,0,50,0,51,0,8,14,0,0,0,75,0,50,0,51,0,58,0,76,0,50,0,51,0,8,14,0,0,0,71,0,50,0,54,0,58,0,83,0,50,0,54,0,8,14,0,0,0,70,0,50,0,48,0,58,0,72,0,50,0,48,0,8,14,0,0,0,73,0,50,0,48,0,58,0,74,0,50,0,48,0,8,14,0,0,0,75,0,50,0,48,0,58,0,76,0,50,0,48,0,8,14,0,0,0,77,0,50,0,48,0,58,0,78,0,50,0,48,0,8,14,0,0,0,79,0,50,0,48,0,58,0,81,0,50,0,48,0,8,14,0,0,0,79,0,49,0,56,0,58,0,81,0,49,0,56,0,8,14,0,0,0,65,0,49,0,57,0,58,0,69,0,49,0,57,0,8,14,0,0,0,70,0,49,0,57,0,58,0,72,0,49,0,57,0,8,14,0,0,0,73,0,49,0,57,0,58,0,74,0,49,0,57,0,8,14,0,0,0,75,0,49,0,57,0,58,0,76,0,49,0,57,0,8,14,0,0,0,65,0,49,0,56,0,58,0,69,0,49,0,56,0,8,14,0,0,0,70,0,49,0,56,0,58,0,72,0,49,0,56,0,8,14,0,0,0,73,0,49,0,56,0,58,0,74,0,49,0,56,0,8,14,0,0,0,75,0,49,0,56,0,58,0,76,0,49,0,56,0,8,14,0,0,0,77,0,49,0,56,0,58,0,78,0,49,0,56,0,8,10,0,0,0,70,0,57,0,58,0,78,0,57,0,8,14,0,0,0,82,0,49,0,48,0,58,0,82,0,49,0,49,0,8,14,0,0,0,75,0,49,0,49,0,58,0,78,0,49,0,49,0,8,14,0,0,0,65,0,49,0,51,0,58,0,82,0,49,0,51,0,8,14,0,0,0,65,0,49,0,53,0,58,0,69,0,49,0,55,0,8,14,0,0,0,70,0,49,0,53,0,58,0,76,0,49,0,53,0,8,14,0,0,0,77,0,49,0,53,0,58,0,82,0,49,0,53,0,8,14,0,0,0,70,0,49,0,54,0,58,0,74,0,49,0,54,0,8,14,0,0,0,75,0,49,0,54,0,58,0,76,0,49,0,55,0,8,14,0,0,0,77,0,49,0,54,0,58,0,81,0,49,0,54,0,8,14,0,0,0,82,0,49,0,54,0,58,0,82,0,49,0,55,0,8,14,0,0,0,70,0,49,0,55,0,58,0,72,0,49,0,55,0,8,14,0,0,0,73,0,49,0,55,0,58,0,74,0,49,0,55,0,8,14,0,0,0,77,0,49,0,55,0,58,0,78,0,49,0,55,0,8,14,0,0,0,79,0,49,0,55,0,58,0,81,0,49,0,55,0,8,10,0,0,0,65,0,50,0,58,0,78,0,50,0,8,10,0,0,0,65,0,51,0,58,0,78,0,51,0,8,10,0,0,0,70,0,52,0,58,0,78,0,53,0,8,10,0,0,0,82,0,54,0,58,0,82,0,55,0,8,10,0,0,0,75,0,55,0,58,0,78,0,55,0,9,9,0,0,0,35,4,0,0,0,139,2,0,0,12,30,71,0,0,13,39,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,1,0,0,0,1,5,221,221,221,221,221,253,33,64,2,4,41,0,0,0,3,5,154,153,153,153,153,57,22,64,2,32,0,0,0,0,4,1,0,0,0,1,5,221,221,221,221,221,253,33,64,2,4,41,0,0,0,3,5,154,153,153,153,153,57,22,64,14,0,0,0,0,9,205,0,0,0,0,200,0,0,0,1,195,0,0,0,1,190,0,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,2,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,120,59,80,119,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,39,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,1,0,0,0,1,5,221,221,221,221,221,253,33,64,2,4,47,0,0,0,3,5,154,153,153,153,153,57,22,64,2,32,0,0,0,0,4,1,0,0,0,1,5,221,221,221,221,221,253,33,64,2,4,47,0,0,0,3,5,154,153,153,153,153,57,22,64,14,0,0,0,0,9,205,0,0,0,0,200,0,0,0,1,195,0,0,0,1,190,0,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,3,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,39,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,1,0,0,0,1,5,221,221,221,221,221,253,33,64,2,4,53,0,0,0,3,5,154,153,153,153,153,57,22,64,2,32,0,0,0,0,4,1,0,0,0,1,5,221,221,221,221,221,253,33,64,2,4,53,0,0,0,3,5,154,153,153,153,153,57,22,64,14,0,0,0,0,9,205,0,0,0,0,200,0,0,0,1,195,0,0,0,1,190,0,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,4,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,82,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,12,0,0,0,1,5,204,204,204,204,204,12,3,64,2,4,6,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,13,0,0,0,1,5,119,119,119,119,119,95,52,64,2,4,6,0,0,0,3,5,36,34,34,34,34,162,13,64,14,0,0,0,0,9,248,1,0,0,0,243,1,0,0,1,238,1,0,0,1,233,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,5,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,130,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,56,0,0,0,250,3,106,74,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,105,0,223,15,251,4,0,0,0,0,3,6,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,205,0,0,0,1,0,0,0,0,196,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,153,0,0,0,1,0,0,0,0,144,0,0,0,1,139,0,0,0,250,0,10,0,0,0,57,0,55,0,48,0,51,0,48,0,49,0,49,0,54,0,49,0,50,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,32,3,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,82,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,12,0,0,0,1,5,204,204,204,204,204,12,3,64,2,4,9,0,0,0,3,5,102,102,102,102,102,102,249,63,2,32,0,0,0,0,4,13,0,0,0,1,5,119,119,119,119,119,95,52,64,2,4,10,0,0,0,3,5,68,68,68,68,68,132,11,64,14,0,0,0,0,9,248,1,0,0,0,243,1,0,0,1,238,1,0,0,1,233,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,6,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,130,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,56,0,0,0,250,3,106,74,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,73,100,61,34,251,4,0,0,0,0,3,6,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,205,0,0,0,1,0,0,0,0,196,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,153,0,0,0,1,0,0,0,0,144,0,0,0,1,139,0,0,0,250,0,10,0,0,0,57,0,55,0,48,0,51,0,48,0,49,0,50,0,48,0,50,0,48,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,32,3,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,68,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,9,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,2,32,0,0,0,0,4,9,0,0,0,1,5,103,102,102,102,102,114,67,64,2,4,39,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,234,1,0,0,0,229,1,0,0,1,224,1,0,0,1,219,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,7,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,24,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,223,0,0,0,1,0,0,0,0,214,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,171,0,0,0,1,0,0,0,0,162,0,0,0,1,157,0,0,0,250,0,19,0,0,0,40,0,77,4,59,4,53,4,58,4,66,4,64,4,62,4,61,4,61,4,75,4,57,4,32,0,48,4,52,4,64,4,53,4,65,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,9,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,2,32,0,0,0,0,4,9,0,0,0,1,5,103,102,102,102,102,114,67,64,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,8,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,99,101,110,116,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,86,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,0,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,2,32,0,0,0,0,4,3,0,0,0,1,5,36,34,34,34,34,162,45,64,2,4,38,0,0,0,3,5,36,34,34,34,34,162,13,64,14,0,0,0,0,9,252,1,0,0,0,247,1,0,0,1,242,1,0,0,1,237,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,9,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,42,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,241,0,0,0,1,0,0,0,0,232,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,189,0,0,0,1,0,0,0,0,180,0,0,0,1,175,0,0,0,250,0,28,0,0,0,40,0,61,4,62,4,60,4,53,4,64,4,32,0,58,4,62,4,61,4,66,4,48,4,58,4,66,4,61,4,62,4,51,4,62,4,32,0,66,4,53,4,59,4,53,4,68,4,62,4,61,4,48,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,0,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,2,32,0,0,0,0,4,3,0,0,0,1,5,36,34,34,34,34,162,45,64,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,10,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,68,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,16,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,2,32,0,0,0,0,4,17,0,0,0,1,5,205,204,204,204,204,4,55,64,2,4,39,0,0,0,3,5,239,238,238,238,238,238,208,63,14,0,0,0,0,9,234,1,0,0,0,229,1,0,0,1,224,1,0,0,1,219,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,11,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,24,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,223,0,0,0,1,0,0,0,0,214,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,171,0,0,0,1,0,0,0,0,162,0,0,0,1,157,0,0,0,250,0,19,0,0,0,40,0,77,4,59,4,53,4,58,4,66,4,64,4,62,4,61,4,61,4,75,4,57,4,32,0,48,4,52,4,64,4,53,4,65,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,16,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,2,32,0,0,0,0,4,17,0,0,0,1,5,205,204,204,204,204,4,55,64,2,4,38,0,0,0,3,5,239,238,238,238,238,238,208,63,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,12,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,9,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,72,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,9,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,9,0,0,0,1,5,35,34,34,34,34,182,67,64,2,4,31,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,238,1,0,0,0,233,1,0,0,1,228,1,0,0,1,223,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,13,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,28,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,227,0,0,0,1,0,0,0,0,218,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,175,0,0,0,1,0,0,0,0,166,0,0,0,1,161,0,0,0,250,0,21,0,0,0,40,0,64,4,48,4,65,4,72,4,56,4,68,4,64,4,62,4,50,4,58,4,48,4,32,0,63,4,62,4,52,4,63,4,56,4,65,4,56,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,9,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,9,0,0,0,1,5,35,34,34,34,34,182,67,64,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,14,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,48,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,5,0,0,0,1,5,36,34,34,34,34,162,253,63,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,7,0,0,0,1,5,85,85,85,85,85,133,50,64,2,4,31,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,214,1,0,0,0,209,1,0,0,1,204,1,0,0,1,199,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,15,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,7,0,0,0,251,4,0,0,0,0,3,4,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,203,0,0,0,1,0,0,0,0,194,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,151,0,0,0,1,0,0,0,0,142,0,0,0,1,137,0,0,0,250,0,9,0,0,0,40,0,63,4,62,4,52,4,63,4,56,4,65,4,76,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,5,0,0,0,1,5,36,34,34,34,34,162,253,63,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,7,0,0,0,1,5,85,85,85,85,85,133,50,64,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,16,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,48,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,5,0,0,0,1,5,239,238,238,238,238,238,0,64,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,7,0,0,0,1,5,85,85,85,85,85,133,50,64,2,4,34,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,214,1,0,0,0,209,1,0,0,1,204,1,0,0,1,199,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,17,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,4,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,203,0,0,0,1,0,0,0,0,194,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,151,0,0,0,1,0,0,0,0,142,0,0,0,1,137,0,0,0,250,0,9,0,0,0,40,0,63,4,62,4,52,4,63,4,56,4,65,4,76,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,5,0,0,0,1,5,239,238,238,238,238,238,0,64,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,7,0,0,0,1,5,85,85,85,85,85,133,50,64,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,18,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,72,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,9,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,9,0,0,0,1,5,103,102,102,102,102,114,67,64,2,4,33,0,0,0,3,5,36,34,34,34,34,162,13,64,14,0,0,0,0,9,238,1,0,0,0,233,1,0,0,1,228,1,0,0,1,223,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,19,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,28,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,227,0,0,0,1,0,0,0,0,218,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,175,0,0,0,1,0,0,0,0,166,0,0,0,1,161,0,0,0,250,0,21,0,0,0,40,0,64,4,48,4,65,4,72,4,56,4,68,4,64,4,62,4,50,4,58,4,48,4,32,0,63,4,62,4,52,4,63,4,56,4,65,4,56,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,9,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,9,0,0,0,1,5,103,102,102,102,102,114,67,64,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,20,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,72,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,16,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,18,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,31,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,238,1,0,0,0,233,1,0,0,1,228,1,0,0,1,223,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,21,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,28,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,227,0,0,0,1,0,0,0,0,218,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,175,0,0,0,1,0,0,0,0,166,0,0,0,1,161,0,0,0,250,0,21,0,0,0,40,0,64,4,48,4,65,4,72,4,56,4,68,4,64,4,62,4,50,4,58,4,48,4,32,0,63,4,62,4,52,4,63,4,56,4,65,4,56,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,16,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,18,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,22,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,72,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,16,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,18,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,33,0,0,0,3,5,36,34,34,34,34,162,13,64,14,0,0,0,0,9,238,1,0,0,0,233,1,0,0,1,228,1,0,0,1,223,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,23,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,28,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,227,0,0,0,1,0,0,0,0,218,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,175,0,0,0,1,0,0,0,0,166,0,0,0,1,161,0,0,0,250,0,21,0,0,0,40,0,64,4,48,4,65,4,72,4,56,4,68,4,64,4,62,4,50,4,58,4,48,4,32,0,63,4,62,4,52,4,63,4,56,4,65,4,56,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,16,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,18,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,24,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,72,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,9,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,9,0,0,0,1,5,103,102,102,102,102,114,67,64,2,4,36,0,0,0,3,5,36,34,34,34,34,162,13,64,14,0,0,0,0,9,238,1,0,0,0,233,1,0,0,1,228,1,0,0,1,223,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,25,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,28,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,227,0,0,0,1,0,0,0,0,218,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,175,0,0,0,1,0,0,0,0,166,0,0,0,1,161,0,0,0,250,0,21,0,0,0,40,0,64,4,48,4,65,4,72,4,56,4,68,4,64,4,62,4,50,4,58,4,48,4,32,0,63,4,62,4,52,4,63,4,56,4,65,4,56,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,9,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,9,0,0,0,1,5,103,102,102,102,102,114,67,64,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,26,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,48,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,5,0,0,0,1,5,36,34,34,34,34,162,253,63,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,7,0,0,0,1,5,85,85,85,85,85,133,50,64,2,4,36,0,0,0,3,5,102,102,102,102,102,102,9,64,14,0,0,0,0,9,214,1,0,0,0,209,1,0,0,1,204,1,0,0,1,199,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,27,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,4,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,203,0,0,0,1,0,0,0,0,194,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,151,0,0,0,1,0,0,0,0,142,0,0,0,1,137,0,0,0,250,0,9,0,0,0,40,0,63,4,62,4,52,4,63,4,56,4,65,4,76,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,5,0,0,0,1,5,36,34,34,34,34,162,253,63,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,7,0,0,0,1,5,85,85,85,85,85,133,50,64,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,28,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,52,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,0,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,4,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,37,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,218,1,0,0,0,213,1,0,0,1,208,1,0,0,1,203,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,29,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,8,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,207,0,0,0,1,0,0,0,0,198,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,155,0,0,0,1,0,0,0,0,146,0,0,0,1,141,0,0,0,250,0,11,0,0,0,40,0,52,4,62,4,59,4,54,4,61,4,62,4,65,4,66,4,76,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,0,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,4,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,30,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,72,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,16,0,0,0,1,5,239,238,238,238,238,238,208,63,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,18,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,37,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,238,1,0,0,0,233,1,0,0,1,228,1,0,0,1,223,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,31,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,28,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,227,0,0,0,1,0,0,0,0,218,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,175,0,0,0,1,0,0,0,0,166,0,0,0,1,161,0,0,0,250,0,21,0,0,0,40,0,64,4,48,4,65,4,72,4,56,4,68,4,64,4,62,4,50,4,58,4,48,4,32,0,63,4,62,4,52,4,63,4,56,4,65,4,56,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,16,0,0,0,1,5,239,238,238,238,238,238,208,63,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,18,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,32,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,48,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,13,0,0,0,1,5,137,136,136,136,136,56,47,64,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,15,0,0,0,1,5,137,136,136,136,136,56,47,64,2,4,36,0,0,0,3,5,102,102,102,102,102,102,9,64,14,0,0,0,0,9,214,1,0,0,0,209,1,0,0,1,204,1,0,0,1,199,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,33,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,4,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,203,0,0,0,1,0,0,0,0,194,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,151,0,0,0,1,0,0,0,0,142,0,0,0,1,137,0,0,0,250,0,9,0,0,0,40,0,63,4,62,4,52,4,63,4,56,4,65,4,76,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,13,0,0,0,1,5,137,136,136,136,136,56,47,64,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,15,0,0,0,1,5,137,136,136,136,136,56,47,64,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,34,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,48,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,13,0,0,0,1,5,36,34,34,34,34,162,45,64,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,15,0,0,0,1,5,137,136,136,136,136,56,47,64,2,4,33,0,0,0,3,5,102,102,102,102,102,102,9,64,14,0,0,0,0,9,214,1,0,0,0,209,1,0,0,1,204,1,0,0,1,199,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,35,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,4,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,203,0,0,0,1,0,0,0,0,194,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,151,0,0,0,1,0,0,0,0,142,0,0,0,1,137,0,0,0,250,0,9,0,0,0,40,0,63,4,62,4,52,4,63,4,56,4,65,4,76,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,13,0,0,0,1,5,36,34,34,34,34,162,45,64,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,15,0,0,0,1,5,137,136,136,136,136,56,47,64,2,4,33,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,36,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,48,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,13,0,0,0,1,5,52,51,51,51,51,147,44,64,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,15,0,0,0,1,5,137,136,136,136,136,56,47,64,2,4,31,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,214,1,0,0,0,209,1,0,0,1,204,1,0,0,1,199,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,37,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,4,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,203,0,0,0,1,0,0,0,0,194,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,151,0,0,0,1,0,0,0,0,142,0,0,0,1,137,0,0,0,250,0,9,0,0,0,40,0,63,4,62,4,52,4,63,4,56,4,65,4,76,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,13,0,0,0,1,5,52,51,51,51,51,147,44,64,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,15,0,0,0,1,5,137,136,136,136,136,56,47,64,2,4,30,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,38,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,86,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,11,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,38,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,13,0,0,0,1,5,102,102,102,102,102,102,41,64,2,4,38,0,0,0,3,5,68,68,68,68,68,132,11,64,14,0,0,0,0,9,252,1,0,0,0,247,1,0,0,1,242,1,0,0,1,237,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,39,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,42,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,241,0,0,0,1,0,0,0,0,232,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,189,0,0,0,1,0,0,0,0,180,0,0,0,1,175,0,0,0,250,0,28,0,0,0,40,0,61,4,62,4,60,4,53,4,64,4,32,0,58,4,62,4,61,4,66,4,48,4,58,4,66,4,61,4,62,4,51,4,62,4,32,0,66,4,53,4,59,4,53,4,68,4,62,4,61,4,48,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,11,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,38,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,13,0,0,0,1,5,102,102,102,102,102,102,41,64,2,4,38,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,40,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,13,52,2,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,11,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,13,0,0,0,1,5,222,221,221,221,221,237,41,64,2,4,37,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,218,1,0,0,0,213,1,0,0,1,208,1,0,0,1,203,1,0,0,250,251,0,70,0,0,0,0,37,0,0,0,250,0,41,0,0,0,1,4,0,0,0,24,4,60,4,79,4,32,0,4,6,0,0,0,68,0,101,0,115,0,99,0,114,0,32,0,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,98,0,0,0,250,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,114,0,101,0,99,0,116,0,251,0,4,0,0,0,0,0,0,0,2,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,255,1,255,2,255,251,3,24,0,0,0,250,251,0,5,0,0,0,2,0,0,0,0,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,3,8,1,0,0,0,42,0,0,0,250,1,4,3,80,70,0,0,7,0,8,160,140,0,0,10,160,140,0,0,15,80,70,0,0,18,0,19,1,251,1,7,0,0,0,250,0,0,0,0,0,251,1,0,0,0,0,2,207,0,0,0,1,0,0,0,0,198,0,0,0,0,33,0,0,0,250,0,0,251,3,0,0,0,0,4,0,0,0,0,5,0,0,0,0,6,0,0,0,0,7,4,0,0,0,0,0,0,0,2,155,0,0,0,1,0,0,0,0,146,0,0,0,1,141,0,0,0,250,0,11,0,0,0,40,0,52,4,62,4,59,4,54,4,61,4,62,4,65,4,66,4,76,4,41,0,251,0,107,0,0,0,250,1,0,7,0,10,5,0,0,0,101,0,110,0,45,0,85,0,83,0,16,1,17,88,2,0,0,18,12,251,1,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,2,0,0,0,0,3,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,5,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,161,5,0,0,0,0,0,0,0,0,13,54,1,0,0,0,1,0,0,0,2,1,32,0,0,0,0,4,11,0,0,0,1,5,0,0,0,0,0,0,0,0,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,2,32,0,0,0,0,4,13,0,0,0,1,5,222,221,221,221,221,237,41,64,2,4,36,0,0,0,3,5,0,0,0,0,0,0,0,0,14,0,0,0,0,9,220,0,0,0,0,215,0,0,0,1,210,0,0,0,3,205,0,0,0,0,51,0,0,0,0,18,0,0,0,250,0,42,0,0,0,1,3,0,0,0,24,4,60,4,79,4,251,1,2,0,0,0,250,251,2,16,0,0,0,250,251,1,0,0,0,0,2,4,0,0,0,0,0,0,0,1,134,0,0,0,250,251,0,22,0,0,0,250,0,0,0,0,0,1,0,0,0,0,2,0,0,0,0,3,0,0,0,0,251,1,29,0,0,0,1,24,0,0,0,250,0,4,0,0,0,108,0,105,0,110,0,101,0,251,0,4,0,0,0,0,0,0,0,2,0,0,0,0,3,56,0,0,0,250,3,53,37,0,0,251,0,23,0,0,0,3,18,0,0,0,0,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,4,0,0,0,0,161,5,0,0,0,0,0,0,0,0,24,34,0,0,0,10,12,0,0,0,11,1,0,0,0,0,12,1,0,0,0,1,13,12,0,0,0,16,1,0,0,0,0,17,1,0,0,0,0,130,19,0,0,5,125,19,0,0,20,120,19,0,0,250,0,11,0,0,0,34,4,53,4,60,4,48,4,32,0,79,0,102,0,102,0,105,0,99,0,101,0,251,0,77,19,0,0,0,31,1,0,0,250,0,11,0,0,0,33,4,66,4,48,4,61,4,52,4,48,4,64,4,66,4,61,4,48,4,79,4,251,0,13,0,0,0,1,8,0,0,0,250,0,79,1,129,2,189,251,1,13,0,0,0,1,8,0,0,0,250,0,192,1,80,2,77,251,2,13,0,0,0,1,8,0,0,0,250,0,155,1,187,2,89,251,3,13,0,0,0,1,8,0,0,0,250,0,128,1,100,2,162,251,4,13,0,0,0,1,8,0,0,0,250,0,75,1,172,2,198,251,5,13,0,0,0,1,8,0,0,0,250,0,247,1,150,2,70,251,8,38,0,0,0,4,33,0,0,0,250,0,10,0,0,0,119,0,105,0,110,0,100,0,111,0,119,0,84,0,101,0,120,0,116,0,1,0,2,0,3,0,251,9,13,0,0,0,1,8,0,0,0,250,0,31,1,73,2,125,251,10,13,0,0,0,1,8,0,0,0,250,0,128,1,0,2,128,251,11,13,0,0,0,1,8,0,0,0,250,0,0,1,0,2,255,251,12,30,0,0,0,4,25,0,0,0,250,0,6,0,0,0,119,0,105,0,110,0,100,0,111,0,119,0,1,255,2,255,3,255,251,13,13,0,0,0,1,8,0,0,0,250,0,238,1,236,2,225,251,1,49,11,0,0,250,0,11,0,0,0,33,4,66,4,48,4,61,4,52,4,48,4,64,4,66,4,61,4,48,4,79,4,251,0,164,5,0,0,0,66,0,0,0,250,1,20,0,0,0,48,0,50,0,48,0,70,0,48,0,51,0,48,0,50,0,48,0,50,0,48,0,50,0,48,0,52,0,48,0,51,0,48,0,50,0,48,0,52,0,3,7,0,0,0,67,0,97,0,109,0,98,0,114,0,105,0,97,0,251,1,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,2,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,3,44,5,0,0,30,0,0,0,0,36,0,0,0,250,0,4,0,0,0,74,0,112,0,97,0,110,0,1,8,0,0,0,45,255,51,255,32,0,48,255,180,48,183,48,195,48,175,48,251,0,30,0,0,0,250,0,4,0,0,0,72,0,97,0,110,0,103,0,1,5,0,0,0,209,185,64,199,32,0,224,172,21,181,251,0,24,0,0,0,250,0,4,0,0,0,72,0,97,0,110,0,115,0,1,2,0,0,0,139,91,83,79,251,0,28,0,0,0,250,0,4,0,0,0,72,0,97,0,110,0,116,0,1,4,0,0,0,176,101,48,125,14,102,212,154,251,0,50,0,0,0,250,0,4,0,0,0,65,0,114,0,97,0,98,0,1,15,0,0,0,84,0,105,0,109,0,101,0,115,0,32,0,78,0,101,0,119,0,32,0,82,0,111,0,109,0,97,0,110,0,251,0,50,0,0,0,250,0,4,0,0,0,72,0,101,0,98,0,114,0,1,15,0,0,0,84,0,105,0,109,0,101,0,115,0,32,0,78,0,101,0,119,0,32,0,82,0,111,0,109,0,97,0,110,0,251,0,32,0,0,0,250,0,4,0,0,0,84,0,104,0,97,0,105,0,1,6,0,0,0,84,0,97,0,104,0,111,0,109,0,97,0,251,0,30,0,0,0,250,0,4,0,0,0,69,0,116,0,104,0,105,0,1,5,0,0,0,78,0,121,0,97,0,108,0,97,0,251,0,32,0,0,0,250,0,4,0,0,0,66,0,101,0,110,0,103,0,1,6,0,0,0,86,0,114,0,105,0,110,0,100,0,97,0,251,0,32,0,0,0,250,0,4,0,0,0,71,0,117,0,106,0,114,0,1,6,0,0,0,83,0,104,0,114,0,117,0,116,0,105,0,251,0,38,0,0,0,250,0,4,0,0,0,75,0,104,0,109,0,114,0,1,9,0,0,0,77,0,111,0,111,0,108,0,66,0,111,0,114,0,97,0,110,0,251,0,30,0,0,0,250,0,4,0,0,0,75,0,110,0,100,0,97,0,1,5,0,0,0,84,0,117,0,110,0,103,0,97,0,251,0,30,0,0,0,250,0,4,0,0,0,71,0,117,0,114,0,117,0,1,5,0,0,0,82,0,97,0,97,0,118,0,105,0,251,0,36,0,0,0,250,0,4,0,0,0,67,0,97,0,110,0,115,0,1,8,0,0,0,69,0,117,0,112,0,104,0,101,0,109,0,105,0,97,0,251,0,60,0,0,0,250,0,4,0,0,0,67,0,104,0,101,0,114,0,1,20,0,0,0,80,0,108,0,97,0,110,0,116,0,97,0,103,0,101,0,110,0,101,0,116,0,32,0,67,0,104,0,101,0,114,0,111,0,107,0,101,0,101,0,251,0,56,0,0,0,250,0,4,0,0,0,89,0,105,0,105,0,105,0,1,18,0,0,0,77,0,105,0,99,0,114,0,111,0,115,0,111,0,102,0,116,0,32,0,89,0,105,0,32,0,66,0,97,0,105,0,116,0,105,0,251,0,56,0,0,0,250,0,4,0,0,0,84,0,105,0,98,0,116,0,1,18,0,0,0,77,0,105,0,99,0,114,0,111,0,115,0,111,0,102,0,116,0,32,0,72,0,105,0,109,0,97,0,108,0,97,0,121,0,97,0,251,0,34,0,0,0,250,0,4,0,0,0,84,0,104,0,97,0,97,0,1,7,0,0,0,77,0,86,0,32,0,66,0,111,0,108,0,105,0,251,0,32,0,0,0,250,0,4,0,0,0,68,0,101,0,118,0,97,0,1,6,0,0,0,77,0,97,0,110,0,103,0,97,0,108,0,251,0,34,0,0,0,250,0,4,0,0,0,84,0,101,0,108,0,117,0,1,7,0,0,0,71,0,97,0,117,0,116,0,97,0,109,0,105,0,251,0,30,0,0,0,250,0,4,0,0,0,84,0,97,0,109,0,108,0,1,5,0,0,0,76,0,97,0,116,0,104,0,97,0,251,0,54,0,0,0,250,0,4,0,0,0,83,0,121,0,114,0,99,0,1,17,0,0,0,69,0,115,0,116,0,114,0,97,0,110,0,103,0,101,0,108,0,111,0,32,0,69,0,100,0,101,0,115,0,115,0,97,0,251,0,34,0,0,0,250,0,4,0,0,0,79,0,114,0,121,0,97,0,1,7,0,0,0,75,0,97,0,108,0,105,0,110,0,103,0,97,0,251,0,34,0,0,0,250,0,4,0,0,0,77,0,108,0,121,0,109,0,1,7,0,0,0,75,0,97,0,114,0,116,0,105,0,107,0,97,0,251,0,38,0,0,0,250,0,4,0,0,0,76,0,97,0,111,0,111,0,1,9,0,0,0,68,0,111,0,107,0,67,0,104,0,97,0,109,0,112,0,97,0,251,0,44,0,0,0,250,0,4,0,0,0,83,0,105,0,110,0,104,0,1,12,0,0,0,73,0,115,0,107,0,111,0,111,0,108,0,97,0,32,0,80,0,111,0,116,0,97,0,251,0,50,0,0,0,250,0,4,0,0,0,77,0,111,0,110,0,103,0,1,15,0,0,0,77,0,111,0,110,0,103,0,111,0,108,0,105,0,97,0,110,0,32,0,66,0,97,0,105,0,116,0,105,0,251,0,50,0,0,0,250,0,4,0,0,0,86,0,105,0,101,0,116,0,1,15,0,0,0,84,0,105,0,109,0,101,0,115,0,32,0,78,0,101,0,119,0,32,0,82,0,111,0,109,0,97,0,110,0,251,0,52,0,0,0,250,0,4,0,0,0,85,0,105,0,103,0,104,0,1,16,0,0,0,77,0,105,0,99,0,114,0,111,0,115,0,111,0,102,0,116,0,32,0,85,0,105,0,103,0,104,0,117,0,114,0,251,0,34,0,0,0,250,0,4,0,0,0,71,0,101,0,111,0,114,0,1,7,0,0,0,83,0,121,0,108,0,102,0,97,0,101,0,110,0,251,1,102,5,0,0,0,66,0,0,0,250,1,20,0,0,0,48,0,50,0,48,0,70,0,48,0,53,0,48,0,50,0,48,0,50,0,48,0,50,0,48,0,52,0,48,0,51,0,48,0,50,0,48,0,52,0,3,7,0,0,0,67,0,97,0,108,0,105,0,98,0,114,0,105,0,251,1,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,2,17,0,0,0,250,3,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,3,238,4,0,0,30,0,0,0,0,36,0,0,0,250,0,4,0,0,0,74,0,112,0,97,0,110,0,1,8,0,0,0,45,255,51,255,32,0,48,255,180,48,183,48,195,48,175,48,251,0,30,0,0,0,250,0,4,0,0,0,72,0,97,0,110,0,103,0,1,5,0,0,0,209,185,64,199,32,0,224,172,21,181,251,0,24,0,0,0,250,0,4,0,0,0,72,0,97,0,110,0,115,0,1,2,0,0,0,139,91,83,79,251,0,28,0,0,0,250,0,4,0,0,0,72,0,97,0,110,0,116,0,1,4,0,0,0,176,101,48,125,14,102,212,154,251,0,30,0,0,0,250,0,4,0,0,0,65,0,114,0,97,0,98,0,1,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,0,30,0,0,0,250,0,4,0,0,0,72,0,101,0,98,0,114,0,1,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,0,32,0,0,0,250,0,4,0,0,0,84,0,104,0,97,0,105,0,1,6,0,0,0,84,0,97,0,104,0,111,0,109,0,97,0,251,0,30,0,0,0,250,0,4,0,0,0,69,0,116,0,104,0,105,0,1,5,0,0,0,78,0,121,0,97,0,108,0,97,0,251,0,32,0,0,0,250,0,4,0,0,0,66,0,101,0,110,0,103,0,1,6,0,0,0,86,0,114,0,105,0,110,0,100,0,97,0,251,0,32,0,0,0,250,0,4,0,0,0,71,0,117,0,106,0,114,0,1,6,0,0,0,83,0,104,0,114,0,117,0,116,0,105,0,251,0,36,0,0,0,250,0,4,0,0,0,75,0,104,0,109,0,114,0,1,8,0,0,0,68,0,97,0,117,0,110,0,80,0,101,0,110,0,104,0,251,0,30,0,0,0,250,0,4,0,0,0,75,0,110,0,100,0,97,0,1,5,0,0,0,84,0,117,0,110,0,103,0,97,0,251,0,30,0,0,0,250,0,4,0,0,0,71,0,117,0,114,0,117,0,1,5,0,0,0,82,0,97,0,97,0,118,0,105,0,251,0,36,0,0,0,250,0,4,0,0,0,67,0,97,0,110,0,115,0,1,8,0,0,0,69,0,117,0,112,0,104,0,101,0,109,0,105,0,97,0,251,0,60,0,0,0,250,0,4,0,0,0,67,0,104,0,101,0,114,0,1,20,0,0,0,80,0,108,0,97,0,110,0,116,0,97,0,103,0,101,0,110,0,101,0,116,0,32,0,67,0,104,0,101,0,114,0,111,0,107,0,101,0,101,0,251,0,56,0,0,0,250,0,4,0,0,0,89,0,105,0,105,0,105,0,1,18,0,0,0,77,0,105,0,99,0,114,0,111,0,115,0,111,0,102,0,116,0,32,0,89,0,105,0,32,0,66,0,97,0,105,0,116,0,105,0,251,0,56,0,0,0,250,0,4,0,0,0,84,0,105,0,98,0,116,0,1,18,0,0,0,77,0,105,0,99,0,114,0,111,0,115,0,111,0,102,0,116,0,32,0,72,0,105,0,109,0,97,0,108,0,97,0,121,0,97,0,251,0,34,0,0,0,250,0,4,0,0,0,84,0,104,0,97,0,97,0,1,7,0,0,0,77,0,86,0,32,0,66,0,111,0,108,0,105,0,251,0,32,0,0,0,250,0,4,0,0,0,68,0,101,0,118,0,97,0,1,6,0,0,0,77,0,97,0,110,0,103,0,97,0,108,0,251,0,34,0,0,0,250,0,4,0,0,0,84,0,101,0,108,0,117,0,1,7,0,0,0,71,0,97,0,117,0,116,0,97,0,109,0,105,0,251,0,30,0,0,0,250,0,4,0,0,0,84,0,97,0,109,0,108,0,1,5,0,0,0,76,0,97,0,116,0,104,0,97,0,251,0,54,0,0,0,250,0,4,0,0,0,83,0,121,0,114,0,99,0,1,17,0,0,0,69,0,115,0,116,0,114,0,97,0,110,0,103,0,101,0,108,0,111,0,32,0,69,0,100,0,101,0,115,0,115,0,97,0,251,0,34,0,0,0,250,0,4,0,0,0,79,0,114,0,121,0,97,0,1,7,0,0,0,75,0,97,0,108,0,105,0,110,0,103,0,97,0,251,0,34,0,0,0,250,0,4,0,0,0,77,0,108,0,121,0,109,0,1,7,0,0,0,75,0,97,0,114,0,116,0,105,0,107,0,97,0,251,0,38,0,0,0,250,0,4,0,0,0,76,0,97,0,111,0,111,0,1,9,0,0,0,68,0,111,0,107,0,67,0,104,0,97,0,109,0,112,0,97,0,251,0,44,0,0,0,250,0,4,0,0,0,83,0,105,0,110,0,104,0,1,12,0,0,0,73,0,115,0,107,0,111,0,111,0,108,0,97,0,32,0,80,0,111,0,116,0,97,0,251,0,50,0,0,0,250,0,4,0,0,0,77,0,111,0,110,0,103,0,1,15,0,0,0,77,0,111,0,110,0,103,0,111,0,108,0,105,0,97,0,110,0,32,0,66,0,97,0,105,0,116,0,105,0,251,0,30,0,0,0,250,0,4,0,0,0,86,0,105,0,101,0,116,0,1,5,0,0,0,65,0,114,0,105,0,97,0,108,0,251,0,52,0,0,0,250,0,4,0,0,0,85,0,105,0,103,0,104,0,1,16,0,0,0,77,0,105,0,99,0,114,0,111,0,115,0,111,0,102,0,116,0,32,0,85,0,105,0,103,0,104,0,117,0,114,0,251,0,34,0,0,0,250,0,4,0,0,0,71,0,101,0,111,0,114,0,1,7,0,0,0,83,0,121,0,108,0,102,0,97,0,101,0,110,0,251,2,238,6,0,0,250,0,11,0,0,0,33,4,66,4,48,4,61,4,52,4,48,4,64,4,66,4,61,4,48,4,79,4,251,0,178,2,0,0,3,0,0,0,0,19,0,0,0,3,14,0,0,0,0,9,0,0,0,3,4,0,0,0,250,0,14,251,0,67,1,0,0,4,62,1,0,0,250,1,1,251,0,39,1,0,0,3,0,0,0,0,92,0,0,0,250,0,0,0,0,0,251,0,80,0,0,0,3,75,0,0,0,250,0,14,251,0,66,0,0,0,2,0,0,0,1,24,0,0,0,250,0,6,0,0,0,97,0,58,0,116,0,105,0,110,0,116,0,1,80,195,0,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,224,147,4,0,251,0,92,0,0,0,250,0,184,136,0,0,251,0,80,0,0,0,3,75,0,0,0,250,0,14,251,0,66,0,0,0,2,0,0,0,1,24,0,0,0,250,0,6,0,0,0,97,0,58,0,116,0,105,0,110,0,116,0,1,136,144,0,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,224,147,4,0,251,0,92,0,0,0,250,0,160,134,1,0,251,0,80,0,0,0,3,75,0,0,0,250,0,14,251,0,66,0,0,0,2,0,0,0,1,24,0,0,0,250,0,6,0,0,0,97,0,58,0,116,0,105,0,110,0,116,0,1,152,58,0,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,48,87,5,0,251,1,9,0,0,0,250,0,64,49,247,0,1,1,251,0,73,1,0,0,4,68,1,0,0,250,1,1,251,0,45,1,0,0,3,0,0,0,0,94,0,0,0,250,0,0,0,0,0,251,0,82,0,0,0,3,77,0,0,0,250,0,14,251,0,68,0,0,0,2,0,0,0,1,26,0,0,0,250,0,7,0,0,0,97,0,58,0,115,0,104,0,97,0,100,0,101,0,1,56,199,0,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,208,251,1,0,251,0,94,0,0,0,250,0,128,56,1,0,251,0,82,0,0,0,3,77,0,0,0,250,0,14,251,0,68,0,0,0,2,0,0,0,1,26,0,0,0,250,0,7,0,0,0,97,0,58,0,115,0,104,0,97,0,100,0,101,0,1,72,107,1,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,208,251,1,0,251,0,94,0,0,0,250,0,160,134,1,0,251,0,82,0,0,0,3,77,0,0,0,250,0,14,251,0,68,0,0,0,2,0,0,0,1,26,0,0,0,250,0,7,0,0,0,97,0,58,0,115,0,104,0,97,0,100,0,101,0,1,48,111,1,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,88,15,2,0,251,1,9,0,0,0,250,0,64,49,247,0,1,0,251,1,10,1,0,0,3,0,0,0,0,131,0,0,0,250,0,0,1,0,2,1,3,53,37,0,0,251,0,92,0,0,0,3,87,0,0,0,0,82,0,0,0,3,77,0,0,0,250,0,14,251,0,68,0,0,0,2,0,0,0,1,26,0,0,0,250,0,7,0,0,0,97,0,58,0,115,0,104,0,97,0,100,0,101,0,1,24,115,1,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,40,154,1,0,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,0,58,0,0,0,250,0,0,1,0,2,1,3,56,99,0,0,251,0,19,0,0,0,3,14,0,0,0,0,9,0,0,0,3,4,0,0,0,250,0,14,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,0,58,0,0,0,250,0,0,1,0,2,1,3,212,148,0,0,251,0,19,0,0,0,3,14,0,0,0,0,9,0,0,0,3,4,0,0,0,250,0,14,251,1,4,0,0,0,250,0,6,251,2,7,0,0,0,250,0,0,0,0,0,251,2,19,0,0,0,3,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,3,238,2,0,0,3,0,0,0,0,19,0,0,0,3,14,0,0,0,0,9,0,0,0,3,4,0,0,0,250,0,14,251,0,166,1,0,0,4,161,1,0,0,250,1,1,251,0,72,1,0,0,3,0,0,0,0,92,0,0,0,250,0,0,0,0,0,251,0,80,0,0,0,3,75,0,0,0,250,0,14,251,0,66,0,0,0,2,0,0,0,1,24,0,0,0,250,0,6,0,0,0,97,0,58,0,116,0,105,0,110,0,116,0,1,64,156,0,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,48,87,5,0,251,0,123,0,0,0,250,0,64,156,0,0,251,0,111,0,0,0,3,106,0,0,0,250,0,14,251,0,97,0,0,0,3,0,0,0,1,24,0,0,0,250,0,6,0,0,0,97,0,58,0,116,0,105,0,110,0,116,0,1,200,175,0,0,251,1,26,0,0,0,250,0,7,0,0,0,97,0,58,0,115,0,104,0,97,0,100,0,101,0,1,184,130,1,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,48,87,5,0,251,0,94,0,0,0,250,0,160,134,1,0,251,0,82,0,0,0,3,77,0,0,0,250,0,14,251,0,68,0,0,0,2,0,0,0,1,26,0,0,0,250,0,7,0,0,0,97,0,58,0,115,0,104,0,97,0,100,0,101,0,1,32,78,0,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,24,228,3,0,251,2,75,0,0,0,250,0,0,251,0,66,0,0,0,250,0,5,0,0,0,53,0,48,0,48,0,48,0,48,0,1,6,0,0,0,45,0,56,0,48,0,48,0,48,0,48,0,2,5,0,0,0,53,0,48,0,48,0,48,0,48,0,3,6,0,0,0,49,0,56,0,48,0,48,0,48,0,48,0,251,0,34,1,0,0,4,29,1,0,0,250,1,1,251,0,200,0,0,0,2,0,0,0,0,92,0,0,0,250,0,0,0,0,0,251,0,80,0,0,0,3,75,0,0,0,250,0,14,251,0,66,0,0,0,2,0,0,0,1,24,0,0,0,250,0,6,0,0,0,97,0,58,0,116,0,105,0,110,0,116,0,1,128,56,1,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,224,147,4,0,251,0,94,0,0,0,250,0,160,134,1,0,251,0,82,0,0,0,3,77,0,0,0,250,0,14,251,0,68,0,0,0,2,0,0,0,1,26,0,0,0,250,0,7,0,0,0,97,0,58,0,115,0,104,0,97,0,100,0,101,0,1,48,117,0,0,251,1,28,0,0,0,250,0,8,0,0,0,97,0,58,0,115,0,97,0,116,0,77,0,111,0,100,0,1,64,13,3,0,251,2,71,0,0,0,250,0,0,251,0,62,0,0,0,250,0,5,0,0,0,53,0,48,0,48,0,48,0,48,0,1,5,0,0,0,53,0,48,0,48,0,48,0,48,0,2,5,0,0,0,53,0,48,0,48,0,48,0,48,0,3,5,0,0,0,53,0,48,0,48,0,48,0,48,0,251,4,4,0,0,0,0,0,0,0,0,0,0,0]);
});
