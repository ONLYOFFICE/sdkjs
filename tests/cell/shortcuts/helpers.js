/*
 * (c) Copyright Ascensio System SIA 2010-2023
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

var AscTestShortcut = AscTestShortcut || {};
(function (window)
{
	window.AscFonts = AscFonts || {};
	window.setTimeout = function (callback)
	{
		callback();
	}
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback)
	{
		callback();
	};
	Asc.spreadsheet_api.prototype._loadModules = function () {};
	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function() {
		this.ImageLoader = AscCommon.g_image_loader;
		this.chartPreviewManager   = new AscCommon.ChartPreviewManager();
		this.textArtPreviewManager = new AscCommon.TextArtPreviewManager();

		AscFormat.initStyleManager();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function(fonts, callback) {
		callback();
	};
	AscCommonExcel.WorksheetView.prototype._calcVisibleRows = function ()
	{

	};
	AscCommonExcel.WorksheetView.prototype._calcVisibleColumns = function ()
	{

	};
	AscCommonExcel.WorksheetView.prototype.scrollVertical = function ()
	{

	};
	AscCommonExcel.WorksheetView.prototype.scrollHorizontal = function ()
	{

	};
	AscCommonExcel.WorksheetView.prototype._normalizeViewRange = function ()
	{

	};
	AscCommonExcel.WorksheetView.prototype._fixVisibleRange = function ()
	{

	};

	AscCommonExcel.WorkbookView.prototype.sendCursor = function ()
	{

	}

	const fOldCellEditor = AscCommonExcel.CellEditor.prototype.open;
	AscCommonExcel.CellEditor.prototype.open = function (options)
	{
		options.getSides = function ()
		{
			return {l: [0], r: [100], b: [10], cellX: 0, cellY: 0, ri: 0, bi: 0};
		}
		fOldCellEditor.call(this, options);
	};
	Asc.DrawingContext.prototype.measureChar = function ()
	{
		return Asc.TextMetrics(5, 9, 10, 1, 1, 10, 5);
	}
	AscFonts.CFontManager.prototype.MeasureChar = function ()
	{
		return {fAdvanceX: 5, oBBox: {fMaxX: 0, fMinX: 0}};
	};
	AscCommon.ZLib = function ()
	{
		this.open = function ()
		{
			return false;
		}
	};

	Asc.DrawingContext.prototype.setFont = function ()
	{
	};
	Asc.DrawingContext.prototype.fillText = function ()
	{
	};
	Asc.DrawingContext.prototype.getFontMetrics = function ()
	{
		return {ascender: 15, descender: 4, lineGap: 1, nat_scale: 1000, nat_y1: 1000, nat_y2: -1000};
	};

	function initEditor()
	{
		const editor = new Asc.spreadsheet_api({'id-view': 'editor_sdk', 'id-input': 'ce-cell-content'});
		editor.FontLoader = {
			LoadDocumentFonts: function ()
		{
			editor.ServerIdWaitComplete = true;
			editor._coAuthoringInitEnd();
			editor.asyncFontsDocumentEndLoaded();
		}
		}


		const sStream = AscCommon.getEmpty();
		const oFile = new AscCommon.OpenFileResult();
		oFile.bSerFormat = AscCommon.checkStreamSignature(sStream, AscCommon.c_oSerFormat.Signature);
		oFile.data = sStream;

		editor.openDocument(oFile);


		editor.asc_setZoom(1);
		return editor;
	}

	const editor = initEditor();

	function createEvent(nKeyCode, bIsCtrl, bIsShift, bIsAlt, bIsAltGr, bIsMacCmdKey)
	{
		const oKeyBoardEvent = {
			preventDefault : function ()
			{
				this.isDefaultPrevented = true;
			},
			stopPropagation: function ()
			{
				this.isPropagationStopped = true;
			}
		};
		oKeyBoardEvent.isDefaultPrevented = false;
		oKeyBoardEvent.isPropagationStopped = false;
		oKeyBoardEvent.which = nKeyCode;
		oKeyBoardEvent.keyCode = nKeyCode;
		oKeyBoardEvent.shiftKey = bIsShift;
		oKeyBoardEvent.altKey = bIsAlt;
		oKeyBoardEvent.ctrlKey = bIsCtrl;
		oKeyBoardEvent.metaKey = bIsMacCmdKey;
		oKeyBoardEvent.altGr = bIsAltGr;
		return oKeyBoardEvent;
	}

	function wbModel()
	{
		return editor.wbModel;
	}

	function wbView()
	{
		return editor.wb;
	}

	function executeTestWithCatchEvent(sSendEvent, fCustomCheck, customExpectedValue, oEvent, oAssert, fBeforeCallback)
	{
		fBeforeCallback && fBeforeCallback();

		let bCheck = false;

		const fCheck = function (...args)
		{
			if (fCustomCheck)
			{
				bCheck = fCustomCheck(...args);
			} else
			{
				bCheck = true;
			}
		}
		editor.asc_registerCallback(sSendEvent, fCheck);

		onKeyDown(oEvent);
		oAssert.strictEqual(bCheck, customExpectedValue === undefined ? true : customExpectedValue, 'Check catch ' + sSendEvent + ' event');
		editor.asc_unregisterCallback(sSendEvent, fCheck);
	}

	function getFragments(start, length)
	{
		return cellEditor()._getFragments(start, length);
	}

	function getSelectionCellEditor()
	{
		return cellEditor().copySelection().map((e) => e.getFragmentText()).join('');
	}

	function moveToStartCellEditor()
	{
		cellEditor()._moveCursor(-2);
	}

	function moveToEndCellEditor()
	{
		cellEditor()._moveCursor(-4);
	}

	function moveRight()
	{
		wbView()._onChangeSelection(true, 1, 0, false, false);
	}

	function moveToCell(nRow, nCol)
	{
		const nCurrentCell = activeCell();
		wbView()._onChangeSelection(true, nCol - nCurrentCell.c1, nRow - nCurrentCell.r1, false, false);
	}

	function selectToCell(nRow, nCol)
	{
		const nCurrentCell = activeCell();
		wbView()._onChangeSelection(false, nCol - nCurrentCell.c1, nRow - nCurrentCell.r1, false, false);
	}

	function onKeyDown(oEvent)
	{
		if (oEvent instanceof Object)
		{
			editor.onKeyDown(oEvent);
		} else
		{
			const oRetEvent = createEvent.apply(null, arguments);
			editor.onKeyDown(oRetEvent);
			return oRetEvent;
		}
	}


	function checkTextAfterKeyDownHelperEmpty(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, '');
	}

	function remove()
	{
		editor.asc_Remove();
	}

	function closeCellEditor(bSkip)
	{
		wbView().closeCellEditor();
	}

	function enterTextWithoutClose(sString)
	{
		wbView().EnterText(sString.split('').map((e) => e.charCodeAt(0)));
		setCheckOpenCellEditor(true);
	}

	function enterText(sString)
	{
		enterTextWithoutClose(sString);
		closeCellEditor();
	}

	function cellEditor()
	{
		return wbView().cellEditor;
	}

	function getCellText()
	{
		closeCellEditor(true);
		return activeCellRange().getValueWithFormat();
	}

	function getCellTextWithoutFormat()
	{
		closeCellEditor();
		return activeCellRange().getValueWithoutFormat();
	}

	function moveDown()
	{
		wbView()._onChangeSelection(true, 0, 1, false, false);
	}

	function wsView()
	{
		return wbView().getWorksheet();
	}

	function ws()
	{
		return wsView().model;
	}

	function moveAndEnterText(sText, nRow, nCol)
	{
		moveToCell(nRow, nCol);
		enterText(sText);
	}

	function createTest(oAssert)
	{
		const deep = (result, expected, sPrompt) => oAssert.deepEqual(result, expected, sPrompt);
		const equal = (result, expected, sPrompt) => oAssert.strictEqual(result, expected, sPrompt);
		return {deep, equal};
	}

	function moveAndGetCellText(nRow, nCol)
	{

		moveToCell(nRow, nCol);
		return getCellText();
	}

	function goToSheet(i)
	{
		wbView().showWorksheet(i);
	}

	let id = 0;

	function createWorksheet()
	{
		const sName = 'name' + id;
		editor.asc_addWorksheet(sName);
		id += 1;
		return sName;
	}

	function removeCurrentWorksheet()
	{
		editor.asc_deleteWorksheet();
	}

	function cleanCell(oRange)
	{
		return {r: oRange.r1, c: oRange.c1};
	}

	function cleanRange(oRange)
	{
		return {r1: oRange.r1, r2: oRange.r2, c1: oRange.c1, c2: oRange.c2};
	}

	function cleanSelection()
	{
		return cleanRange(selectionRange());
	}

	function cleanActiveCell()
	{
		return cleanCell(activeCell());
	}

	function checkRange(nRow1, nRow2, nCol1, nCol2)
	{
		return cleanRange(Asc.Range(nCol1, nRow1, nCol2, nRow2, true));
	}

	function openCellEditor()
	{
		var enterOptions = new AscCommonExcel.CEditorEnterOptions();
		enterOptions.newText = '';
		enterOptions.quickInput = true;
		enterOptions.focus = true;
		handlers().trigger('editCell', enterOptions);
		setCheckOpenCellEditor(true);
	}

	function checkActiveCell(nRow, nCol)
	{
		return cleanCell(new Asc.Range(nCol, nRow, nCol, nRow, true));
	}

	function cleanCache()
	{
		wsView()._cleanCache();
	}

	function selectAll()
	{
		wbView().selectAll();
	}

	function cleanAll()
	{
		selectAll();
		handlers().trigger("empty");
		cleanCache();
		wsView().changeZoomResize();
		wsView().visibleRange = new Asc.Range(0, 0, 23, 37);
		moveToCell(0, 0);
	}

	function setCellFormat(nFormat)
	{
		handlers().trigger('setCellFormat', nFormat);


	}

	function selectionInfo()
	{
		return wbView().getSelectionInfo();
	}

	function xfs()
	{
		return selectionInfo().asc_getXfs();
	}

	function undo()
	{
		handlers().trigger('undo');
	}

	function selectAllCell()
	{
		cellEditor()._moveCursor(-2);
		cellEditor()._selectChars(-4);
	}

	function cellPosition()
	{
		return cellEditor().cursorPos;
	}

	function getCellEditMode()
	{
		return wsView().getCellEditMode();
	}

	function testPreventDefaultAndStopPropagation(oEvent, oAssert, bInvert)
	{
		onKeyDown(oEvent);
		oAssert.true(!!bInvert ? !oEvent.isDefaultPrevented : oEvent.isDefaultPrevented);
		oAssert.true(!!bInvert ? !oEvent.isPropagationStopped : oEvent.isPropagationStopped);
	}

	function controller()
	{
		return editor.wb.controller;
	}

	function handlers()
	{
		return controller().handlers;
	}

	function activeCell()
	{
		return wsView().getActiveCell();
	}

	function selectionRange()
	{
		return wsView().getSelectedRange().bbox;
	}

	function activeCellRange()
	{
		return ws().getRange3(activeCell().r1, activeCell().c1, activeCell().r2, activeCell().c2);
	}

	function addHyperlink(sLink, sText)
	{
		const oHyperlink = new AscCommonExcel.Hyperlink();
		oHyperlink.Hyperlink = sLink;

		const oAscHyperlink = new Asc.asc_CHyperlink(oHyperlink);
		oAscHyperlink.text = sText;
		editor.asc_insertHyperlink(oAscHyperlink);
	}

	function createMathInShape()
	{
		const {oShape, oParagraph} = createShapeWithContent('', true);
		selectOnlyObjects([oShape]);
		moveToParagraph(oParagraph, true);
		editor.asc_AddMath2(c_oAscMathType.FractionVertical);
		return {oShape, oParagraph};
	}

	function moveCursorRight(bAddToSelect, bWord)
	{
		graphicController().cursorMoveRight(bAddToSelect, bWord);
	}

	function moveCursorLeft(bAddToSelect, bWord)
	{
		graphicController().cursorMoveLeft(bAddToSelect, bWord);
	}

	function moveCursorUp(bAddToSelect, bWord)
	{
		graphicController().cursorMoveUp(bAddToSelect, bWord);
	}

	function moveCursorDown(bAddToSelect, bWord)
	{
		graphicController().cursorMoveDown(bAddToSelect, bWord);

	}

	function checkMoveContentShapeHelper(arrExpected, oEvent, oAssert)
	{
		const {deep, equal} = createTest(oAssert);
		const {oShape} = createShapeWithContent('Hello World Hello World Hello World Hello World Hello World');
		moveToShapeParagraph(oShape, 0);
		onKeyDown(oEvent);
		equal(contentPosition(), arrExpected[0]);

		moveToShapeParagraph(oShape, 0, true);
		onKeyDown(oEvent);
		equal(contentPosition(), arrExpected[1]);
	}

	function contentPosition()
	{
		const arrContentPosition = graphicController().getTargetDocContent().GetContentPosition();
		return arrContentPosition[arrContentPosition.length - 1].Position;
	}

	function selectedContent()
	{
		return graphicController().GetSelectedText(false, {TabSymbol: '\t'});
	}

	function checkSelectContentShapeHelper(arrExpected, oEvent, oAssert)
	{
		const {deep, equal} = createTest(oAssert);
		const {oShape} = createShapeWithContent('Hello World Hello World Hello World Hello World Hello World');
		moveToShapeParagraph(oShape, 0, true);
		onKeyDown(oEvent);
		equal(selectedContent(), arrExpected[0]);

		moveToShapeParagraph(oShape, 0);
		onKeyDown(oEvent);
		equal(selectedContent(), arrExpected[1]);
	}

	function round(nNumber, nAmount)
	{
		const nPower = Math.pow(10, nAmount);
		return Math.round(nNumber * nPower) / nPower;
	}

	function moveShapeHelper(oExpected, oEvent, oAssert)
	{
		const oShape = createShape();
		const {deep, equal} = createTest(oAssert);
		selectOnlyObjects([oShape]);

		onKeyDown(oEvent);
		oShape.recalculateTransform();
		deep({x: round(oShape.x, 12), y: round(oShape.y, 12)}, {x: round(oExpected.x, 12), y: round(oExpected.y, 12)});
	}

	function createTable()
	{
		moveToCell(0, 0);
		enterText('Hello');

		moveToCell(1, 0);
		enterText('1');
		moveToCell(2, 0);
		enterText('2');
		moveToCell(3, 0);
		enterText('3');
		moveToCell(4, 0);
		enterText('4');
		moveToCell(0, 0);
		selectToCell(4, 0);

		editor.asc_addAutoFilter('TableStyleMedium2', editor.asc_getAddFormatTableOptions());
	}

	function createSlicer()
	{
		createTable();
		editor.asc_insertSlicer(['Hello']);
		return graphicController().getSelectedArray()[0];
	}

	function moveToShapeParagraph(oShape, nParagraph, bIsMoveToStart)
	{
		const oParagraph = oShape.getDocContent().Content[nParagraph];
		moveToParagraph(oParagraph, bIsMoveToStart);

		return oParagraph;
	}


	function addProperty(oPr)
	{
		graphicController().paragraphAdd(new AscCommonWord.ParaTextPr(oPr), true);
	}

	function targetContent()
	{
		return graphicController().getTargetDocContent();
	}

	function textPr()
	{
		return graphicController().getParagraphTextPr();
	}

	function paraPr()
	{
		return graphicController().getParagraphParaPr();
	}

	function moveToParagraph(oParagraph, bToStart)
	{
		oParagraph.SetThisElementCurrent();
		if (bToStart)
		{
			oParagraph.MoveCursorToStartPos();
		} else
		{
			oParagraph.MoveCursorToEndPos();
		}

		graphicController().recalculateCurPos(true, true);
	}

	function selectedObjects()
	{
		return graphicController().getSelectedArray();
	}

	function cleanGraphic()
	{
		graphicController().resetSelection();
		graphicController().selectAll();
		graphicController().remove();
	}

	function graphicController()
	{
		return editor.getGraphicController();
	}

	function createShape()
	{
		const oGraphicController = graphicController();
		const oTrack = new AscFormat.NewShapeTrack('rect', 0, 0, wbModel().theme, null, null, null, 0, oGraphicController);
		const oShape = oTrack.getShape(false, drawingDocument(), oGraphicController.drawingObjects);
		oShape.setWorksheet(wbView().getWorksheet().model);
		oShape.spPr.xfrm.setOffX(100);
		oShape.spPr.xfrm.setOffY(100);
		oShape.spPr.xfrm.setExtX(100);
		oShape.spPr.xfrm.setExtY(100);
		oShape.setPaddings({Left: 0, Right: 0, Top: 0, Bottom: 0});
		oShape.setParent(oGraphicController.drawingObjects);
		oShape.setRecalculateInfo();

		oShape.addToDrawingObjects(undefined, AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
		oShape.checkDrawingBaseCoords();
		oGraphicController.checkChartTextSelection();
		oGraphicController.resetSelection();
		oShape.select(oGraphicController, 0);
		oGraphicController.startRecalculate();
		oGraphicController.drawingObjects.sendGraphicObjectProps();
		oGraphicController.recalculate();
		oShape.recalculateTransform();

		return oShape;
	}

	function drawingDocument()
	{
		return wbModel().DrawingDocument;
	}

	function selectOnlyObjects(arrShapes)
	{
		const oController = graphicController();
		oController.resetSelection();
		for (let i = 0; i < arrShapes.length; i += 1)
		{
			const oObject = arrShapes[i];
			if (oObject.group)
			{
				const oMainGroup = oObject.group.getMainGroup();
				oMainGroup.select(oController, 0);
				oController.selection.groupSelection = oMainGroup;
			}
			oObject.select(oController, 0);
		}
	}

	function createShapeWithContent(sText)
	{
		const oShape = createShape();
		selectOnlyObjects([oShape])
		oShape.setTxBody(AscFormat.CreateTextBodyFromString(sText, drawingDocument(), oShape));
		const oParagraph = oShape.txBody.content.Content[0];
		oShape.recalculateContentWitCompiledPr();
		return {oShape, oParagraph};
	}

	function checkTextAfterKeyDownHelperHelloWorld(sCheckText, oEvent, oAssert, sPrompt)
	{
		checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, 'Hello World');
	}

	function getParagraphText(oParagraph)
	{
		return oParagraph.GetText({ParaEndToSpace: false});
	}

	function checkTextAfterKeyDownHelper(sCheckText, oEvent, oAssert, sPrompt, sInitText)
	{
		const {oParagraph} = createShapeWithContent(sInitText);
		moveToParagraph(oParagraph);
		onKeyDown(oEvent);
		const sTextAfterKeyDown = getParagraphText(oParagraph);
		oAssert.strictEqual(sTextAfterKeyDown, sCheckText, sPrompt);
	}

	function checkDirectTextPrAfterKeyDown(fCallback, oEvent, oAssert, nExpectedValue, sPrompt)
	{
		const {oParagraph} = createShapeWithContent('Hello World');
		let oTextPr = getDirectTextPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oTextPr), nExpectedValue, sPrompt);
		return function continueCheck(fCallback2, oEvent2, oAssert2, nExpectedValue2, sPrompt2)
		{
			oTextPr = getDirectTextPrHelper(oParagraph, oEvent2);
			oAssert2.strictEqual(fCallback2(oTextPr), nExpectedValue2, sPrompt2);
			return continueCheck;
		}
	}

	function checkDirectParaPrAfterKeyDown(fCallback, oEvent, oAssert, nExpectedValue, sPrompt)
	{
		const {oParagraph} = createShapeWithContent('Hello World');
		let oParaPr = getDirectParaPrHelper(oParagraph, oEvent);
		oAssert.strictEqual(fCallback(oParaPr), nExpectedValue, sPrompt);

		return function continueCheck(fCallback2, oEvent2, oAssert2, nExpectedValue2, sPrompt2)
		{
			oParaPr = getDirectParaPrHelper(oParagraph, oEvent2);
			oAssert2.strictEqual(fCallback2(oParaPr), nExpectedValue2, sPrompt2);
			return continueCheck;
		}
	}

	function getDirectTextPrHelper(oParagraph, oEvent)
	{
		oParagraph.SetThisElementCurrent();
		graphicController().selectAll();
		onKeyDown(oEvent);
		return textPr();
	}

	function getDirectParaPrHelper(oParagraph, oEvent)
	{
		oParagraph.SetThisElementCurrent();
		graphicController().selectAll();
		onKeyDown(oEvent);
		return paraPr();
	}

	function checkRemoveObject(oShape, arrShapes)
	{
		return !!(arrShapes.indexOf(oShape) === -1 && oShape.bDeleted);
	}

	function createGroup(arrShapes)
	{
		graphicController().resetSelection();
		selectOnlyObjects(arrShapes);
		return graphicController().createGroup();
	}

	function createChart()
	{
		moveAndEnterText('1', 0, 0);
		moveAndEnterText('2', 1, 0);
		moveAndEnterText('3', 2, 0);

		moveAndEnterText('1', 0, 1);
		moveAndEnterText('2', 1, 1);
		moveAndEnterText('3', 2, 1);

		selectToCell(0, 0);

		const oProps = editor.asc_getChartObject(true);
		oProps.changeType(0);
		editor.asc_addChartDrawingObject(oProps);
		return selectedObjects()[0];
	}

	let bIsCellEditorOpened = false;

	function setCheckOpenCellEditor(oPr)
	{
		bIsCellEditorOpened = oPr;
	}

	function checkOpenCellEditor()
	{
		return bIsCellEditorOpened;
	}

	AscTestShortcut.wbModel = wbModel;
	AscTestShortcut.wbView = wbView;
	AscTestShortcut.executeTestWithCatchEvent = executeTestWithCatchEvent;
	AscTestShortcut.getFragments = getFragments;
	AscTestShortcut.getSelectionCellEditor = getSelectionCellEditor;
	AscTestShortcut.moveToStartCellEditor = moveToStartCellEditor;
	AscTestShortcut.moveToEndCellEditor = moveToEndCellEditor;
	AscTestShortcut.moveRight = moveRight;
	AscTestShortcut.moveToCell = moveToCell;
	AscTestShortcut.selectToCell = selectToCell;
	AscTestShortcut.onKeyDown = onKeyDown;
	AscTestShortcut.remove = remove;
	AscTestShortcut.closeCellEditor = closeCellEditor;
	AscTestShortcut.enterTextWithoutClose = enterTextWithoutClose;
	AscTestShortcut.enterText = enterText;
	AscTestShortcut.cellEditor = cellEditor;
	AscTestShortcut.getCellText = getCellText;
	AscTestShortcut.getCellTextWithoutFormat = getCellTextWithoutFormat;
	AscTestShortcut.moveDown = moveDown;
	AscTestShortcut.wsView = wsView;
	AscTestShortcut.ws = ws;
	AscTestShortcut.moveAndEnterText = moveAndEnterText;
	AscTestShortcut.createTest = createTest;
	AscTestShortcut.moveAndGetCellText = moveAndGetCellText;
	AscTestShortcut.goToSheet = goToSheet;
	AscTestShortcut.createWorksheet = createWorksheet;
	AscTestShortcut.removeCurrentWorksheet = removeCurrentWorksheet;
	AscTestShortcut.cleanCell = cleanCell;
	AscTestShortcut.cleanRange = cleanRange;
	AscTestShortcut.cleanSelection = cleanSelection;
	AscTestShortcut.cleanActiveCell = cleanActiveCell;
	AscTestShortcut.checkRange = checkRange;
	AscTestShortcut.openCellEditor = openCellEditor;
	AscTestShortcut.checkActiveCell = checkActiveCell;
	AscTestShortcut.cleanCache = cleanCache;
	AscTestShortcut.selectAll = selectAll;
	AscTestShortcut.cleanAll = cleanAll;
	AscTestShortcut.setCellFormat = setCellFormat;
	AscTestShortcut.selectionInfo = selectionInfo;
	AscTestShortcut.xfs = xfs;
	AscTestShortcut.undo = undo;
	AscTestShortcut.createEvent = createEvent;
	AscTestShortcut.selectAllCell = selectAllCell;
	AscTestShortcut.cellPosition = cellPosition;
	AscTestShortcut.getCellEditMode = getCellEditMode;
	AscTestShortcut.testPreventDefaultAndStopPropagation = testPreventDefaultAndStopPropagation;
	AscTestShortcut.controller = controller;
	AscTestShortcut.handlers = handlers;
	AscTestShortcut.activeCell = activeCell;
	AscTestShortcut.selectionRange = selectionRange;
	AscTestShortcut.checkOpenCellEditor = checkOpenCellEditor;
	AscTestShortcut.setCheckOpenCellEditor = setCheckOpenCellEditor;
	AscTestShortcut.activeCellRange = activeCellRange;
	AscTestShortcut.addHyperlink = addHyperlink;
	AscTestShortcut.createMathInShape = createMathInShape;
	AscTestShortcut.moveCursorRight = moveCursorRight;
	AscTestShortcut.moveCursorLeft = moveCursorLeft;
	AscTestShortcut.moveCursorUp = moveCursorUp;
	AscTestShortcut.moveCursorDown = moveCursorDown;
	AscTestShortcut.checkMoveContentShapeHelper = checkMoveContentShapeHelper;
	AscTestShortcut.contentPosition = contentPosition;
	AscTestShortcut.selectedContent = selectedContent;
	AscTestShortcut.checkSelectContentShapeHelper = checkSelectContentShapeHelper;
	AscTestShortcut.round = round;
	AscTestShortcut.moveShapeHelper = moveShapeHelper;
	AscTestShortcut.createTable = createTable;
	AscTestShortcut.createSlicer = createSlicer;
	AscTestShortcut.moveToShapeParagraph = moveToShapeParagraph;
	AscTestShortcut.addProperty = addProperty;
	AscTestShortcut.targetContent = targetContent;
	AscTestShortcut.textPr = textPr;
	AscTestShortcut.paraPr = paraPr;
	AscTestShortcut.moveToParagraph = moveToParagraph;
	AscTestShortcut.selectedObjects = selectedObjects;
	AscTestShortcut.cleanGraphic = cleanGraphic;
	AscTestShortcut.graphicController = graphicController;
	AscTestShortcut.createShape = createShape;
	AscTestShortcut.drawingDocument = drawingDocument;
	AscTestShortcut.selectOnlyObjects = selectOnlyObjects;
	AscTestShortcut.createShapeWithContent = createShapeWithContent;
	AscTestShortcut.checkTextAfterKeyDownHelperHelloWorld = checkTextAfterKeyDownHelperHelloWorld;
	AscTestShortcut.checkTextAfterKeyDownHelperEmpty = checkTextAfterKeyDownHelperEmpty;
	AscTestShortcut.getParagraphText = getParagraphText;
	AscTestShortcut.checkTextAfterKeyDownHelper = checkTextAfterKeyDownHelper;
	AscTestShortcut.checkDirectTextPrAfterKeyDown = checkDirectTextPrAfterKeyDown;
	AscTestShortcut.checkDirectParaPrAfterKeyDown = checkDirectParaPrAfterKeyDown;
	AscTestShortcut.getDirectTextPrHelper = getDirectTextPrHelper;
	AscTestShortcut.getDirectParaPrHelper = getDirectParaPrHelper;
	AscTestShortcut.checkRemoveObject = checkRemoveObject;
	AscTestShortcut.createGroup = createGroup;
	AscTestShortcut.createChart = createChart;
})(window);
