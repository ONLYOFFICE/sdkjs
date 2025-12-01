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

$(function () {

	Asc.spreadsheet_api.prototype._init = function () {
		this._loadModules();
	};
	Asc.spreadsheet_api.prototype._loadFonts = function (fonts, callback) {
		callback();
	};
	Asc.spreadsheet_api.prototype.onEndLoadFile = function (fonts, callback) {
		openDocument();
	};
	AscCommonExcel.WorkbookView.prototype._calcMaxDigitWidth = function () {
	};
	AscCommonExcel.WorkbookView.prototype._init = function () {
	};
	AscCommonExcel.WorkbookView.prototype._isLockedUserProtectedRange = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorkbookView.prototype._onWSSelectionChanged = function () {
	};
	AscCommonExcel.WorkbookView.prototype.showWorksheet = function () {
	};
	AscCommonExcel.WorkbookView.prototype.recalculateDrawingObjects = function () {
	};
	AscCommonExcel.WorkbookView.prototype.restoreFocus = function () {
	};
	AscCommonExcel.WorkbookView.prototype._onChangeSelection = function (isStartPoint, dc, dr, isCoord, isCtrl, callback) {
        if (!this._checkStopCellEditorInFormulas()) {
            return;
        }

        var ws = this.getWorksheet();
		if (ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.selectUnlockedCells)) {
			return;
		}
		if (ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.selectLockedCells)) {
			//TODO _getRangeByXY ?
			var newRange = isCoord ? ws._getRangeByXY(dc, dr) :
				ws._calcSelectionEndPointByOffset(dc, dr);
			var lockedCell = ws.model.getLockedCell(newRange.c2, newRange.r2);
			if (lockedCell || lockedCell === null) {
				return;
			}
		}

        if (this.selectionDialogMode && !ws.model.selectionRange) {
            if (isCoord) {
                ws.model.selectionRange = new AscCommonExcel.SelectionRange(ws.model);

				// remove first range if we paste argument with ctrl key
				if (isCtrl && ws.model.selectionRange.ranges && Array.isArray(ws.model.selectionRange.ranges)) {
					ws.model.selectionRange.ranges.shift();
				}

                isStartPoint = true;
            } else {
                ws.model.selectionRange = ws.model.copySelection.clone();
            }
        }

        var t = this;
        var d = isStartPoint ? ws.changeSelectionStartPoint(dc, dr, isCoord, isCtrl) :
            ws.changeSelectionEndPoint(dc, dr, isCoord, isCoord && this.keepType);
        if (!isCoord && !isStartPoint) {
            // Выделение с зажатым shift
            this.canUpdateAfterShiftUp = true;
        }
        this.keepType = isCoord;
        // if (isCoord && !this.timerEnd && this.timerId === null) {
        //     this.timerId = setTimeout(function () {
        //         var arrClose = [];
        //         arrClose.push(new asc_CMM({type: c_oAscMouseMoveType.None}));
        //         t.handlers.trigger("asc_onMouseMove", arrClose);
        //         t._onUpdateCursor(AscCommon.Cursors.CellCur);
        //         t.timerId = null;
        //         t.timerEnd = true;
        //     }, 1000);
        // }

        if (this.isFormulaEditMode && this.isCellEditMode && this.cellEditor && this.cellEditor.openFromTopLine) {
            /* set focus to the top formula entry line */
            this.cellEditor.restoreFocus();
        }

        AscCommonExcel.applyFunction(callback, d);
    };
	AscCommonExcel.WorksheetView.prototype._init = function () {
	};
	AscCommonExcel.WorksheetView.prototype.updateRanges = function () {
	};
	AscCommonExcel.WorksheetView.prototype._autoFitColumnsWidth = function () {
	};
	AscCommonExcel.WorksheetView.prototype.cleanSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._drawSelection = function () {
	};
	AscCommonExcel.WorksheetView.prototype._scrollToRange = function () {
	};
	AscCommonExcel.WorksheetView.prototype.draw = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareDrawingObjects = function () {
	};
	AscCommonExcel.WorksheetView.prototype._initCellsArea = function () {
	};
	AscCommonExcel.WorksheetView.prototype.getZoom = function () {
	};
	AscCommonExcel.WorksheetView.prototype._prepareCellTextMetricsCache = function () {
	};
	AscCommonExcel.WorksheetView.prototype._getCellCache = function (col, row) {
		let _cell = null;
		this.model.getRange3(row, col, row, col)._foreachNoEmpty(function(cell, row, col) {
			if (cell && !cell.isEmptyTextString()) {
				_cell = {cellType: cell.getType()}
			}
		}, null, true);
		return _cell;
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	AscCommonExcel.WorksheetView.prototype._isLockedCells = function (range, subType, callback) {
		callback(true);
		return true;
	};
	AscCommonExcel.WorksheetView.prototype._isLockedAll = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._isLockedFrozenPane = function (callback) {
		callback(true);
	};
	AscCommonExcel.WorksheetView.prototype._updateVisibleColsCount = function () {
	};
	AscCommonExcel.WorksheetView.prototype._calcActiveCellOffset = function () {
	};

	AscCommon.baseEditorsApi.prototype._onEndLoadSdk = function () {
	};
	Asc.ReadDefTableStyles = function(){};

	function openDocument(){
		AscCommon.g_oTableId.init();
		api._onEndLoadSdk();
		api.isOpenOOXInBrowser = false;
		api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	}

	var api = new Asc.spreadsheet_api({
		'id-view': 'editor_sdk'
	});
	api.FontLoader = {
		LoadDocumentFonts: function() {}
	};
	window["Asc"]["editor"] = api;
	AscCommon.g_oTableId.init();
	api._onEndLoadSdk();
	api.isOpenOOXInBrowser = false;
	api.OpenDocumentFromBin(null, AscCommon.getEmpty());
	api.initCollaborativeEditing({});
	api.wb = new AscCommonExcel.WorkbookView(api.wbModel, api.controller, api.handlers, api.HtmlElement,
		api.topLineEditorElement, api, api.collaborativeEditing, api.fontRenderingMode);
	var wb = api.wbModel;
	wb.handlers.add("getSelectionState", function () {
		return null;
	});
	wb.handlers.add("getLockDefNameManagerStatus", function () {
		return true;
	});
	wb.handlers.add("asc_onConfirmAction", function (test1, callback) {
		callback(true);
	});
	api.wb.cellCommentator = new AscCommonExcel.CCellCommentator({
		model: api.wbModel.aWorksheets[0],
		collaborativeEditing: null,
		draw: function() {
		},
		handlers: {
			trigger: function() {
				return false;
			}
		}
	});

	AscCommonExcel.CCellCommentator.prototype.isLockedComment = function (oComment, callbackFunc) {
		callbackFunc(true);
	};
	AscCommonExcel.CCellCommentator.prototype.drawCommentCells = function () {
	};
	AscCommonExcel.CCellCommentator.prototype.ascCvtRatio = function () {
	};

	var wsView = api.wb.getWorksheet(0);
	wsView.handlers = api.handlers;
	wsView.objectRender = new AscFormat.DrawingObjects();
	var ws = api.wbModel.aWorksheets[0];

	var getRange = function (c1, r1, c2, r2) {
		return new window["Asc"].Range(c1, r1, c2, r2);
	};
	const clearData = function (c1, r1, c2, r2) {
		ws.autoFilters.deleteAutoFilter(getRange(0,0,0,0));
		ws.TableParts = [];
		ws.removeRows(r1, r2, false);
		ws.removeCols(c1, c2);
	};

	function checkUndoRedo(fBefore, fAfter, desc) {
		fAfter("after_" + desc);
		AscCommon.History.Undo();
		fBefore("undo_" + desc);
		AscCommon.History.Redo();
		fAfter("redo_" + desc);
		AscCommon.History.Undo();
	}

	function compareData (assert, range, data, desc) {
		for (let i = range.r1; i <= range.r2; i++) {
			for (let j = range.c1; j <= range.c2; j++) {
				let rangeVal = ws.getCell3(i, j);
				let dataVal = data[i - range.r1][j - range.c1];
				assert.strictEqual(rangeVal.getValue(), dataVal, desc + " compare " + rangeVal.getName());
			}
		}
	}
	function autofillData (assert, rangeTo, expectedData, description) {
		for (let i = rangeTo.r1; i <= rangeTo.r2; i++) {
			for (let j = rangeTo.c1; j <= rangeTo.c2; j++) {
				let rangeToVal = ws.getCell3(i, j);
				let dataVal = expectedData[i - rangeTo.r1][j - rangeTo.c1];
				assert.strictEqual(rangeToVal.getValue(), dataVal, `${description} Cell: ${rangeToVal.getName()}, Value: ${dataVal}`);
			}
		}
	}
	function reverseAutofillData (assert, rangeTo, expectedData, description) {
		for (let i = rangeTo.r1; i >= rangeTo.r2; i--) {
			for (let j = rangeTo.c1; j >= rangeTo.c2; j--) {
				let rangeToVal = ws.getCell3(i, j);
				let dataVal = expectedData[Math.abs(i - rangeTo.r1)][Math.abs(j - rangeTo.c1)];
				assert.strictEqual(rangeToVal.getValue(), dataVal, `${description} Cell: ${rangeToVal.getName()}, Value: ${dataVal}`);
			}
		}
	}
	function getAutoFillRange(wsView, c1To, r1To, c2To, r2To, nHandleDirection, nFillHandleArea) {
		wsView.fillHandleArea = nFillHandleArea;
		wsView.fillHandleDirection = nHandleDirection;
		wsView.activeFillHandle = getRange(c1To, r1To, c2To, r2To);
		wsView.applyFillHandle(0,0,false);

		return wsView;
	}
	function updateDataToUpCase (aExpectedData) {
		return aExpectedData.map (function (expectedData) {
			if (Array.isArray(expectedData)) {
				return [expectedData[0].toUpperCase()]
			}
			return expectedData.toUpperCase();
		});
	}
	function updateDataToLowCase (aExpectedData) {
		return aExpectedData.map (function (expectedData) {
			if (Array.isArray(expectedData)) {
				return [expectedData[0].toLowerCase()]
			}
			return expectedData.toLowerCase();
		});
	}
	function getHorizontalAutofillCases(c1From, c2From, c1To, c2To, assert, expectedData, nFillHandleArea) {
		const [
			expectedDataCapitalized,
			expectedDataUpper,
			expectedDataLower,
			expectedDataShortCapitalized,
			expectedDataShortUpper,
			expectedDataShortLower
		] = expectedData;

		const nHandleDirection = 0; // 0 - Horizontal, 1 - Vertical
		let autofillC1 =  nFillHandleArea === 3 ? c2From + 1 : c1From - 1;
		const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;
		const descSequenceType = nFillHandleArea === 3 ? 'Asc sequence.' : 'Reverse sequence.';
		// With capitalized
		ws.selectionRange.ranges = [getRange(c1From, 0, c2From, 0)];
		wsView = getAutoFillRange(wsView, c1To, 0, c2To, 0, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(autofillC1, 0, c2To, 0);
		autoFillAssert(assert, autoFillRange, [expectedDataCapitalized], `Case: ${descSequenceType} With capitalized`);

		//Upper-registry
		ws.selectionRange.ranges = [getRange(c1From, 1, c2From, 1)];
		wsView = getAutoFillRange(wsView, c1To, 1, c2To, 1, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 1, c2To, 1);
		autoFillAssert(assert, autoFillRange, [expectedDataUpper], `Case: ${descSequenceType} Upper-registry`);

		// Lower-registry
		ws.selectionRange.ranges = [getRange(c1From, 2, c2From, 2)];
		wsView = getAutoFillRange(wsView, c1To, 2, c2To, 2, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 2, c2To, 2);
		autoFillAssert(assert, autoFillRange, [expectedDataLower], `Case: ${descSequenceType} Lower-registry`);

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(c1From, 3, c2From, 3)];
		wsView = getAutoFillRange(wsView, c1To, 3, c2To, 3, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 3, c2To, 3);
		autoFillAssert(assert, autoFillRange, [expectedDataCapitalized], `Case: ${descSequenceType} Camel-registry - Su.`);

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(c1From, 4, c2From, 4)];
		wsView = getAutoFillRange(wsView, c1To, 4, c2To, 4, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 4, c2To, 4);
		autoFillAssert(assert, autoFillRange, [expectedDataUpper], `Case: ${descSequenceType} Camel-registry - SU.`);

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(c1From, 5, c2From, 5)];
		wsView = getAutoFillRange(wsView, c1To, 5, c2To, 5, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 5, c2To, 5);
		autoFillAssert(assert, autoFillRange, [expectedDataLower], `Case: ${descSequenceType} Camel-registry - sU.`);

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(c1From, 6, c2From, 6)];
		wsView = getAutoFillRange(wsView, c1To, 6, c2To, 6, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 6, c2To, 6);
		autoFillAssert(assert, autoFillRange, [expectedDataLower], `Case: ${descSequenceType} Camel-registry - su.`);

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(c1From, 7, c2From, 7)];
		wsView = getAutoFillRange(wsView, c1To, 7, c2To, 7, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 7, c2To, 7);
		autoFillAssert(assert, autoFillRange, [expectedDataShortCapitalized], `Case: ${descSequenceType} Short name with capitalized`);

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(c1From, 8, c2From,8)];
		wsView = getAutoFillRange(wsView, c1To, 8, c2To, 8, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 8, c2To, 8);
		autoFillAssert(assert, autoFillRange, [expectedDataShortUpper], `Case: ${descSequenceType} Short name Upper-registry start from Sun`);

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(c1From,9,c2From,9)];
		wsView = getAutoFillRange(wsView, c1To, 9, c2To, 9, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 9, c2To, 9);
		autoFillAssert(assert, autoFillRange, [expectedDataShortLower], `Case: ${descSequenceType} Short name Lower-registry`);

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(c1From, 10, c2From, 10)];
		wsView = getAutoFillRange(wsView, c1To, 10, c2To, 10, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 10, c2To, 10);
		autoFillAssert(assert, autoFillRange, [expectedDataShortCapitalized], `Case: ${descSequenceType} Short name Camel-registry - Su.`);

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(c1From, 11, c2From, 11)];
		wsView = getAutoFillRange(wsView, c1To, 11, c2To, 11, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 11, c2To, 11);
		autoFillAssert(assert, autoFillRange, [expectedDataShortUpper], `Case: ${descSequenceType} Short name Camel-registry - SU.`);

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(c1From, 12, c2From, 12)];
		wsView = getAutoFillRange(wsView, c1To, 12, c2To, 12, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 12, c2To, 12);
		autoFillAssert(assert, autoFillRange, [expectedDataShortLower], `Case: ${descSequenceType} Short name Camel-registry - sU.`);

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(c1From, 13, c2From, 13)];
		wsView = getAutoFillRange(wsView, c1To, 13, c2To, 13, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(autofillC1, 13, c2To, 13);
		autoFillAssert(assert, autoFillRange, [expectedDataShortLower], `Case: ${descSequenceType} Short name Camel-registry - su.`);
	}

	function getVerticalAutofillCases (r1From, r2From, r1To, r2To, assert, expectedData, nFillHandleArea) {
		const [
			expectedDataCapitalized,
			expectedDataUpper,
			expectedDataLower,
			expectedDataShortCapitalized,
			expectedDataShortUpper,
			expectedDataShortLower
		] = expectedData;

		const nHandleDirection = 1; // 0 - Horizontal, 1 - Vertical,
		let autofillR1 =  nFillHandleArea === 3 ? r2From + 1 : r1From - 1;
		const autoFillAssert = nFillHandleArea === 3 ? autofillData : reverseAutofillData;
		const descSequenceType = nFillHandleArea === 3 ? 'Asc sequence.' : 'Reverse sequence.';
		// With capitalized
		ws.selectionRange.ranges = [getRange(0, r1From, 0, r2From)];
		wsView = getAutoFillRange(wsView, 0, r1To, 0, r2To, nHandleDirection, nFillHandleArea);
		let autoFillRange = getRange(0, autofillR1, 0, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataCapitalized, `Case: ${descSequenceType} With capitalized`);

		//Upper-registry
		ws.selectionRange.ranges = [getRange(1, r1From, 1, r2From)];
		wsView = getAutoFillRange(wsView, 1, r1To, 1, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(1, autofillR1, 1, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataUpper, `Case: ${descSequenceType} Upper-registry`);

		// Lower-registry
		ws.selectionRange.ranges = [getRange(2, r1From, 2, r2From)];
		wsView = getAutoFillRange(wsView, 2, r1To, 2, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(2, autofillR1, 2, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataLower, `Case: ${descSequenceType} Lower-registry`);

		// Camel-registry - SuNdAy
		ws.selectionRange.ranges = [getRange(3, r1From, 3, r2From)];
		wsView = getAutoFillRange(wsView, 3, r1To, 3, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(3, autofillR1, 3, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataCapitalized, `Case: ${descSequenceType} Camel-registry - Su.`);

		// Camel-registry - SUnDaY
		ws.selectionRange.ranges = [getRange(4, r1From, 4, r2From)];
		wsView = getAutoFillRange(wsView, 4, r1To, 4, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(4, autofillR1, 4, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataUpper, `Case: ${descSequenceType} Camel-registry - SU.`);

		// Camel-registry - sUnDaY
		ws.selectionRange.ranges = [getRange(5, r1From, 5, r2From)];
		wsView = getAutoFillRange(wsView, 5, r1To, 5, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(5, autofillR1, 5, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataLower, `Case: ${descSequenceType} Camel-registry - sU.`);

		// Camel-registry - suNDay
		ws.selectionRange.ranges = [getRange(6, r1From, 6, r2From)];
		wsView = getAutoFillRange(wsView, 6, r1To, 6, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(6, autofillR1, 6, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataLower, `Case: ${descSequenceType} Camel-registry - su.`);

		// Short name day of the week with capitalized
		ws.selectionRange.ranges = [getRange(7, r1From, 7, r2From)];
		wsView = getAutoFillRange(wsView, 7, r1To, 7, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(7, autofillR1, 7, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortCapitalized, `Case: ${descSequenceType} Short name with capitalized`);

		// Short name day of the week Upper-registry
		ws.selectionRange.ranges = [getRange(8, r1From, 8, r2From)];
		wsView = getAutoFillRange(wsView, 8, r1To, 8, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(8, autofillR1, 8, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortUpper, `Case: ${descSequenceType} Short name Upper-registry`);

		// Short name day of the week Lower-registry
		ws.selectionRange.ranges = [getRange(9, r1From, 9, r2From)];
		wsView = getAutoFillRange(wsView, 9, r1To, 9, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(9, autofillR1, 9, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortLower, `Case: ${descSequenceType} Short name Lower-registry`);

		// Short name  day of the week Camel-registry - SuN
		ws.selectionRange.ranges = [getRange(10, r1From, 10, r2From)];
		wsView = getAutoFillRange(wsView, 10, r1To, 10, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(10, autofillR1, 10, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortCapitalized, `Case: ${descSequenceType} Short name Camel-registry - Su.`);

		// Short name day of the week Camel-registry - SUn
		ws.selectionRange.ranges = [getRange(11, r1From, 11, r2From)];
		wsView = getAutoFillRange(wsView, 11, r1To, 11, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(11, autofillR1, 11, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortUpper, `Case: ${descSequenceType} Short name Camel-registry - SU.`);

		// Short name day of the week Camel-registry - sUn
		ws.selectionRange.ranges = [getRange(12, r1From, 12, r2From)];
		wsView = getAutoFillRange(wsView, 12, r1To, 12, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(12, autofillR1, 12, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortLower, `Case: ${descSequenceType} Short name Camel-registry - sU.`);

		// Short name day of the week Camel-registry - suN
		ws.selectionRange.ranges = [getRange(13, r1From, 13, r2From)];
		wsView = getAutoFillRange(wsView, 13, r1To, 13, r2To, nHandleDirection, nFillHandleArea);
		autoFillRange = getRange(13, autofillR1, 13, r2To);
		autoFillAssert(assert, autoFillRange, expectedDataShortLower, `Case: ${descSequenceType} Short name Camel-registry - su.`);

	}

	function CacheColumn() {
	    this.left = 0;
		this.width = 0;

		this._widthForPrint = null;
	}

	const getCell = function (oRange) {
		let oCell = null;

		oRange._foreach2(function (cell) {
			oCell = cell;
		})

		return oCell;
	};

	QUnit.test('Test @ -> single function with different types', function (assert) {
		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SIN(@B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIN(_xlfn.SINGLE(B1))", "SIN(@B1) -> SIN(SINGLE(B1))");

		fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText("=SUM(@B1:B3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A2"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(B1:B3))", "SUM(@B1:B3) -> SUM(SINGLE(B1:B3))");

		fillRange = ws.getRange2("A4");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A4").getValueForEdit2();
		fragment[0].setFragmentText("=SUM(@3:3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A4"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(3:3))", "SUM(@3:3) -> SUM(SINGLE(3:3))");

		fillRange = ws.getRange2("A5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A5").getValueForEdit2();
		fragment[0].setFragmentText("=SUM(@B:B)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A5"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(B:B))", "SUM(@B:B) -> SUM(SINGLE(B:B))");

		fillRange = ws.getRange2("A6");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A6").getValueForEdit2();
		fragment[0].setFragmentText("=IF(@TRUE,1,0)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A6"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(TRUE),1,0)", "IF(@TRUE,1,0) -> IF(SINGLE(TRUE),1,0)");

		fillRange = ws.getRange2("A7");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A7").getValueForEdit2();
		fragment[0].setFragmentText('=LEN(@"test")');
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A7"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), 'LEN(_xlfn.SINGLE("test"))', 'LEN(@"test") -> LEN(SINGLE("test"))');

		fillRange = ws.getRange2("A7");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A7").getValueForEdit2();
		fragment[0].setFragmentText('=LEN(@{1,2,3})');
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A7"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), 'LEN(_xlfn.SINGLE({1,2,3}))', 'LEN(@{1,2,3}) -> LEN(SINGLE({1,2,3}))');

		fillRange = ws.getRange2("A9");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A9").getValueForEdit2();
		fragment[0].setFragmentText("=SUM(IF(@B1:B3>0,@B1:B3,0))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A9"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(IF(_xlfn.SINGLE(B1:B3)>0,_xlfn.SINGLE(B1:B3),0))", "SUM(IF(@B1:B3>0,@B1:B3,0)) -> SUM(IF(SINGLE(B1:B3)>0,SINGLE(B1:B3),0))");

		fillRange = ws.getRange2("A10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A10").getValueForEdit2();
		fragment[0].setFragmentText("=AVERAGE(IF(@B1:B3<>0,@B1:B3))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A10"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "AVERAGE(IF(_xlfn.SINGLE(B1:B3)<>0,_xlfn.SINGLE(B1:B3)))", "AVERAGE(IF(@B1:B3<>0,@B1:B3)) -> AVERAGE(IF(SINGLE(B1:B3)<>0,SINGLE(B1:B3)))");

		fillRange = ws.getRange2("A11");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A11").getValueForEdit2();
		fragment[0].setFragmentText("=IF(AND(@B1>0,@C1>0),SUM(@B1:C1),0)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A11"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(AND(_xlfn.SINGLE(B1)>0,_xlfn.SINGLE(C1)>0),SUM(_xlfn.SINGLE(B1:C1)),0)", "IF(AND(@B1>0,@C1>0),SUM(@B1:C1),0) -> IF(AND(SINGLE(B1)>0,SINGLE(C1)>0),SUM(SINGLE(B1:C1)),0)");

		fillRange = ws.getRange2("A12");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A12").getValueForEdit2();
		fragment[0].setFragmentText("=SUM(IF(OR(@B1:B3>10,@B1:B3<0),@B1:B3,0))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A12"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(IF(OR(_xlfn.SINGLE(B1:B3)>10,_xlfn.SINGLE(B1:B3)<0),_xlfn.SINGLE(B1:B3),0))", "SUM(IF(OR(@B1:B3>10,@B1:B3<0),@B1:B3,0)) -> SUM(IF(OR(SINGLE(B1:B3)>10,SINGLE(B1:B3)<0),SINGLE(B1:B3),0))");

		fillRange = ws.getRange2("A13");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A13").getValueForEdit2();
		fragment[0].setFragmentText("=ROUND(AVERAGE(@B1:B3),2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A13"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUND(AVERAGE(_xlfn.SINGLE(B1:B3)),2)", "ROUND(AVERAGE(@B1:B3),2) -> ROUND(AVERAGE(SINGLE(B1:B3)),2)");

		fillRange = ws.getRange2("A14");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A14").getValueForEdit2();
		fragment[0].setFragmentText("=VLOOKUP(@B1,@D1:E10,2,FALSE)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A14"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:E10),2,FALSE)", "VLOOKUP(@B1,@D1:E10,2,FALSE) -> VLOOKUP(SINGLE(B1),SINGLE(D1:E10),2,FALSE)");

		fillRange = ws.getRange2("A15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A15").getValueForEdit2();
		fragment[0].setFragmentText("=SUMPRODUCT((@B1:B3>5)*(@C1:C3<10)*@B1:B3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A15"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUMPRODUCT((_xlfn.SINGLE(B1:B3)>5)*(_xlfn.SINGLE(C1:C3)<10)*_xlfn.SINGLE(B1:B3))", "SUMPRODUCT((@B1:B3>5)*(@C1:C3<10)*@B1:B3) -> SUMPRODUCT((SINGLE(B1:B3)>5)*(SINGLE(C1:C3)<10)*SINGLE(B1:B3))");

		fillRange = ws.getRange2("A16");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A16").getValueForEdit2();
		fragment[0].setFragmentText("=IF(@B1>0,IF(@C1>0,SUM(@B1:C1),@B1),0)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A16"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(B1)>0,IF(_xlfn.SINGLE(C1)>0,SUM(_xlfn.SINGLE(B1:C1)),_xlfn.SINGLE(B1)),0)", "IF(@B1>0,IF(@C1>0,SUM(@B1:C1),@B1),0) -> IF(SINGLE(B1)>0,IF(SINGLE(C1)>0,SUM(SINGLE(B1:C1)),SINGLE(B1)),0)");

		fillRange = ws.getRange2("A17");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A17").getValueForEdit2();
		fragment[0].setFragmentText("=IFERROR(VLOOKUP(@B1,@D1:E10,2,FALSE),@B1*2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A17"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IFERROR(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:E10),2,FALSE),_xlfn.SINGLE(B1)*2)", "IFERROR(VLOOKUP(@B1,@D1:E10,2,FALSE),@B1*2) -> IFERROR(VLOOKUP(SINGLE(B1),SINGLE(D1:E10),2,FALSE),SINGLE(B1)*2)");

		//TODO =SUMIF(SINGLE(B1:B10),">"&SINGLE(D1),SINGLE(C1:C10)) - don't check parser
		// fillRange = ws.getRange2("A18");
		// wsView.setSelection(fillRange.bbox);
		// fragment = ws.getRange2("A18").getValueForEdit2();
		// fragment[0].setFragmentText("=SUMIF(@B1:B10,\">\"&@D1,@C1:C10)");
		// wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		// resCell = getCell(ws.getRange2("A18"));
		// assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUMIF(_xlfn.SINGLE(B1:B10),\">\"&_xlfn.SINGLE(D1),_xlfn.SINGLE(C1:C10))", "SUMIF(@B1:B10,\">\"&@D1,@C1:C10) -> SUMIF(SINGLE(B1:B10),\">\"&SINGLE(D1),SINGLE(C1:C10))");

		fillRange = ws.getRange2("A19");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A19").getValueForEdit2();
		fragment[0].setFragmentText("=INDEX(@B1:B10,MATCH(MAX(@B1:B10),@B1:B10,0))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A19"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "INDEX(_xlfn.SINGLE(B1:B10),MATCH(MAX(_xlfn.SINGLE(B1:B10)),_xlfn.SINGLE(B1:B10),0))", "INDEX(@B1:B10,MATCH(MAX(@B1:B10),@B1:B10,0)) -> INDEX(SINGLE(B1:B10),MATCH(MAX(SINGLE(B1:B10)),SINGLE(B1:B10),0))");

		fillRange = ws.getRange2("A20");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A20").getValueForEdit2();
		fragment[0].setFragmentText("=CONCATENATE(@B1,\" \",@C1,\" \",@D1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A20"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CONCATENATE(_xlfn.SINGLE(B1),\" \",_xlfn.SINGLE(C1),\" \",_xlfn.SINGLE(D1))", "CONCATENATE(@B1,\" \",@C1,\" \",@D1) -> CONCATENATE(SINGLE(B1),\" \",SINGLE(C1),\" \",SINGLE(D1))");

		fillRange = ws.getRange2("A9");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A9").getValueForEdit2();
		fragment[0].setFragmentText("=@SIN(@B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A9"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SIN(_xlfn.SINGLE(B1)))", "@SIN(@B1) -> SINGLE(SIN(SINGLE(B1)))");

		fillRange = ws.getRange2("A10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A10").getValueForEdit2();
		fragment[0].setFragmentText("=@SUM(@B1:B3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A10"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:B3)))", "@SUM(@B1:B3) -> SINGLE(SUM(SINGLE(B1:B3)))");

		fillRange = ws.getRange2("A11");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A11").getValueForEdit2();
		fragment[0].setFragmentText("=@IF(@B1>0,@C1,0)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A11"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)>0,_xlfn.SINGLE(C1),0))", "@IF(@B1>0,@C1,0) -> SINGLE(IF(SINGLE(B1)>0,SINGLE(C1),0))");

		fillRange = ws.getRange2("A12");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A12").getValueForEdit2();
		fragment[0].setFragmentText("=@ROUND(@AVERAGE(@B1:B3),2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A12"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(ROUND(_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:B3))),2))", "@ROUND(@AVERAGE(@B1:B3),2) -> SINGLE(ROUND(SINGLE(AVERAGE(SINGLE(B1:B3))),2))");

		fillRange = ws.getRange2("A13");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A13").getValueForEdit2();
		fragment[0].setFragmentText("=@ABS(@MIN(@B1:B3))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A13"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(ABS(_xlfn.SINGLE(MIN(_xlfn.SINGLE(B1:B3)))))", "@ABS(@MIN(@B1:B3)) -> SINGLE(ABS(SINGLE(MIN(SINGLE(B1:B3)))))");

		fillRange = ws.getRange2("A14");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A14").getValueForEdit2();
		fragment[0].setFragmentText("=@SQRT(@ABS(@B1))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A14"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SQRT(_xlfn.SINGLE(ABS(_xlfn.SINGLE(B1)))))", "@SQRT(@ABS(@B1)) -> SINGLE(SQRT(SINGLE(ABS(SINGLE(B1)))))");

		fillRange = ws.getRange2("A15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A15").getValueForEdit2();
		fragment[0].setFragmentText("=@IF(@AND(@B1>0,@C1>0),@SUM(@B1:C1),0)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A15"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(AND(_xlfn.SINGLE(B1)>0,_xlfn.SINGLE(C1)>0)),_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:C1))),0))", "@IF(@AND(@B1>0,@C1>0),@SUM(@B1:C1),0) -> SINGLE(IF(SINGLE(AND(SINGLE(B1)>0,SINGLE(C1)>0)),SINGLE(SUM(SINGLE(B1:C1))),0))");

		fillRange = ws.getRange2("A16");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A16").getValueForEdit2();
		fragment[0].setFragmentText("=@VLOOKUP(@B1,@D1:E10,2,FALSE)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A16"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:E10),2,FALSE))", "@VLOOKUP(@B1,@D1:E10,2,FALSE) -> SINGLE(VLOOKUP(SINGLE(B1),SINGLE(D1:E10),2,FALSE))");

		fillRange = ws.getRange2("A17");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A17").getValueForEdit2();
		fragment[0].setFragmentText("=@IFERROR(@VLOOKUP(@B1,@D1:E10,2,FALSE),@B1*2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A17"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IFERROR(_xlfn.SINGLE(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:E10),2,FALSE)),_xlfn.SINGLE(B1)*2))", "@IFERROR(@VLOOKUP(@B1,@D1:E10,2,FALSE),@B1*2) -> SINGLE(IFERROR(SINGLE(VLOOKUP(SINGLE(B1),SINGLE(D1:E10),2,FALSE)),SINGLE(B1)*2))");

		fillRange = ws.getRange2("A18");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A18").getValueForEdit2();
		fragment[0].setFragmentText("=@TEXT(@ROUND(@AVERAGE(@B1:B3),2),\"0.00\")");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A18"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(TEXT(_xlfn.SINGLE(ROUND(_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:B3))),2)),\"0.00\"))", "@TEXT(@ROUND(@AVERAGE(@B1:B3),2),\"0.00\") -> SINGLE(TEXT(SINGLE(ROUND(SINGLE(AVERAGE(SINGLE(B1:B3))),2)),\"0.00\"))");


		fillRange = ws.getRange2("A18");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A18").getValueForEdit2();
		fragment[0].setFragmentText("=@TEXT(@ROUND(@AVERAGE(@B1:B3),2),\"0.00\")");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A18"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(TEXT(_xlfn.SINGLE(ROUND(_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:B3))),2)),\"0.00\"))", "@TEXT(@ROUND(@AVERAGE(@B1:B3),2),\"0.00\") -> SINGLE(TEXT(SINGLE(ROUND(SINGLE(AVERAGE(SINGLE(B1:B3))),2)),\"0.00\"))");

		// Additional test cases for @ operator

		// Test @ with numbers
		fillRange = ws.getRange2("A21");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A21").getValueForEdit2();
		fragment[0].setFragmentText("=@123");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A21"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(123)", "@123 -> SINGLE(123)");

		// Test @ with negative numbers
		fillRange = ws.getRange2("A22");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A22").getValueForEdit2();
		fragment[0].setFragmentText("=SUM(@-5,@B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A22"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(-5),_xlfn.SINGLE(B1))", "SUM(@-5,@B1) -> SUM(SINGLE(-5),SINGLE(B1))");

		// Test @ with error values
		fillRange = ws.getRange2("A23");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A23").getValueForEdit2();
		fragment[0].setFragmentText("=IFERROR(@#N/A,0)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A23"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IFERROR(_xlfn.SINGLE(#N/A),0)", "IFERROR(@#N/A,0) -> IFERROR(SINGLE(#N/A),0)");

		// Test @ with 3D references
		fillRange = ws.getRange2("A24");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A24").getValueForEdit2();
		fragment[0].setFragmentText("=SUM(@Sheet1!A1:A10)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A24"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(Sheet1!A1:A10))", "SUM(@Sheet1!A1:A10) -> SUM(SINGLE(Sheet1!A1:A10))");

		// Test @ with CHOOSE function
		fillRange = ws.getRange2("A25");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A25").getValueForEdit2();
		fragment[0].setFragmentText("=CHOOSE(@B1,@C1,@D1,@E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A25"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CHOOSE(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1),_xlfn.SINGLE(E1))", "CHOOSE(@B1,@C1,@D1,@E1) -> CHOOSE(SINGLE(B1),SINGLE(C1),SINGLE(D1),SINGLE(E1))");

		// Test @ with COUNT functions
		fillRange = ws.getRange2("A26");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A26").getValueForEdit2();
		fragment[0].setFragmentText("=COUNTIF(@B1:B10,@C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A26"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COUNTIF(_xlfn.SINGLE(B1:B10),_xlfn.SINGLE(C1))", "COUNTIF(@B1:B10,@C1) -> COUNTIF(SINGLE(B1:B10),SINGLE(C1))");

		// Test @ with nested MAX/MIN
		fillRange = ws.getRange2("A27");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A27").getValueForEdit2();
		fragment[0].setFragmentText("=@MAX(@MIN(@B1:B10),@MIN(@C1:C10))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A27"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(MAX(_xlfn.SINGLE(MIN(_xlfn.SINGLE(B1:B10))),_xlfn.SINGLE(MIN(_xlfn.SINGLE(C1:C10)))))", "@MAX(@MIN(@B1:B10),@MIN(@C1:C10)) -> SINGLE(MAX(SINGLE(MIN(SINGLE(B1:B10))),SINGLE(MIN(SINGLE(C1:C10)))))");

		// Test @ with LEFT/RIGHT/MID string functions
		fillRange = ws.getRange2("A28");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A28").getValueForEdit2();
		fragment[0].setFragmentText("=LEFT(@B1,@C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A28"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LEFT(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "LEFT(@B1,@C1) -> LEFT(SINGLE(B1),SINGLE(C1))");

		// Test @ with MID function
		fillRange = ws.getRange2("A29");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A29").getValueForEdit2();
		fragment[0].setFragmentText("=MID(@B1,@C1,@D1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A29"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MID(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))", "MID(@B1,@C1,@D1) -> MID(SINGLE(B1),SINGLE(C1),SINGLE(D1))");

		// Test @ with DATE/TIME functions
		fillRange = ws.getRange2("A30");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A30").getValueForEdit2();
		fragment[0].setFragmentText("=DATE(@B1,@C1,@D1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A30"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DATE(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))", "DATE(@B1,@C1,@D1) -> DATE(SINGLE(B1),SINGLE(C1),SINGLE(D1))");

		// Test @ with arithmetic operations between @ operands
		fillRange = ws.getRange2("A31");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A31").getValueForEdit2();
		fragment[0].setFragmentText("=@B1+@C1*@D1");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A31"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(B1)+_xlfn.SINGLE(C1)*_xlfn.SINGLE(D1)", "@B1+@C1*@D1 -> SINGLE(B1)+SINGLE(C1)*SINGLE(D1)");

		// Test @ with comparison operators
		fillRange = ws.getRange2("A32");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A32").getValueForEdit2();
		fragment[0].setFragmentText("=IF(@B1>=@C1,@B1,@C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A32"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(B1)>=_xlfn.SINGLE(C1),_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "IF(@B1>=@C1,@B1,@C1) -> IF(SINGLE(B1)>=SINGLE(C1),SINGLE(B1),SINGLE(C1))");

		// Test @ with POWER function
		fillRange = ws.getRange2("A33");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A33").getValueForEdit2();
		fragment[0].setFragmentText("=POWER(@B1,@C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A33"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "POWER(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "POWER(@B1,@C1) -> POWER(SINGLE(B1),SINGLE(C1))");

		// Test @ with MOD function
		fillRange = ws.getRange2("A34");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A34").getValueForEdit2();
		fragment[0].setFragmentText("=MOD(@B1,@C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A34"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MOD(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "MOD(@B1,@C1) -> MOD(SINGLE(B1),SINGLE(C1))");

		// Test @ with SUBSTITUTE function
		fillRange = ws.getRange2("A35");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A35").getValueForEdit2();
		fragment[0].setFragmentText("=SUBSTITUTE(@B1,@C1,@D1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A35"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUBSTITUTE(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))", "SUBSTITUTE(@B1,@C1,@D1) -> SUBSTITUTE(SINGLE(B1),SINGLE(C1),SINGLE(D1))");

		// Test @ with INDIRECT (dynamic reference)
		fillRange = ws.getRange2("A36");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A36").getValueForEdit2();
		fragment[0].setFragmentText("=@INDIRECT(@B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A36"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(INDIRECT(_xlfn.SINGLE(B1)))", "@INDIRECT(@B1) -> SINGLE(INDIRECT(SINGLE(B1)))");

		// Test @ with OFFSET function
		fillRange = ws.getRange2("A37");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A37").getValueForEdit2();
		fragment[0].setFragmentText("=SUM(@OFFSET(@B1,@C1,@D1))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A37"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(OFFSET(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))))", "SUM(@OFFSET(@B1,@C1,@D1)) -> SUM(SINGLE(OFFSET(SINGLE(B1),SINGLE(C1),SINGLE(D1))))");

		// Test @ with LARGE/SMALL functions
		fillRange = ws.getRange2("A38");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A38").getValueForEdit2();
		fragment[0].setFragmentText("=LARGE(@B1:B10,@C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A38"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LARGE(_xlfn.SINGLE(B1:B10),_xlfn.SINGLE(C1))", "LARGE(@B1:B10,@C1) -> LARGE(SINGLE(B1:B10),SINGLE(C1))");

		// Test @ with SMALL function
		fillRange = ws.getRange2("A39");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A39").getValueForEdit2();
		fragment[0].setFragmentText("=@SMALL(@B1:B10,@C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A39"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SMALL(_xlfn.SINGLE(B1:B10),_xlfn.SINGLE(C1)))", "@SMALL(@B1:B10,@C1) -> SINGLE(SMALL(SINGLE(B1:B10),SINGLE(C1)))");

		// Test @ with LOOKUP function
		fillRange = ws.getRange2("A40");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A40").getValueForEdit2();
		fragment[0].setFragmentText("=LOOKUP(@B1,@C1:C10,@D1:D10)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A40"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1:C10),_xlfn.SINGLE(D1:D10))", "LOOKUP(@B1,@C1:C10,@D1:D10) -> LOOKUP(SINGLE(B1),SINGLE(C1:C10),SINGLE(D1:D10))");

		// Test @ with HLOOKUP function
		fillRange = ws.getRange2("A41");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A41").getValueForEdit2();
		fragment[0].setFragmentText("=HLOOKUP(@B1,@C1:G3,@D1,FALSE)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A41"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "HLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1:G3),_xlfn.SINGLE(D1),FALSE)", "HLOOKUP(@B1,@C1:G3,@D1,FALSE) -> HLOOKUP(SINGLE(B1),SINGLE(C1:G3),SINGLE(D1),FALSE)");

		// Test @ with NOT function
		fillRange = ws.getRange2("A42");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A42").getValueForEdit2();
		fragment[0].setFragmentText("=IF(@NOT(@B1),@C1,@D1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A42"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(NOT(_xlfn.SINGLE(B1))),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))", "IF(@NOT(@B1),@C1,@D1) -> IF(SINGLE(NOT(SINGLE(B1))),SINGLE(C1),SINGLE(D1))");

		// Test @ with ISBLANK function
		fillRange = ws.getRange2("A43");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A43").getValueForEdit2();
		fragment[0].setFragmentText("=IF(@ISBLANK(@B1),@C1,@B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A43"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(ISBLANK(_xlfn.SINGLE(B1))),_xlfn.SINGLE(C1),_xlfn.SINGLE(B1))", "IF(@ISBLANK(@B1),@C1,@B1) -> IF(SINGLE(ISBLANK(SINGLE(B1))),SINGLE(C1),SINGLE(B1))");

		// Test @ with ISNUMBER function
		fillRange = ws.getRange2("A44");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A44").getValueForEdit2();
		fragment[0].setFragmentText("=@ISNUMBER(@B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A44"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(ISNUMBER(_xlfn.SINGLE(B1)))", "@ISNUMBER(@B1) -> SINGLE(ISNUMBER(SINGLE(B1)))");

		// Test @ with TRIM and CLEAN functions
		fillRange = ws.getRange2("A45");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A45").getValueForEdit2();
		fragment[0].setFragmentText("=@TRIM(@CLEAN(@B1))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A45"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(TRIM(_xlfn.SINGLE(CLEAN(_xlfn.SINGLE(B1)))))", "@TRIM(@CLEAN(@B1)) -> SINGLE(TRIM(SINGLE(CLEAN(SINGLE(B1)))))");

		// Test @ with VALUE function
		fillRange = ws.getRange2("A46");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A46").getValueForEdit2();
		fragment[0].setFragmentText("=@VALUE(@B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A46"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(VALUE(_xlfn.SINGLE(B1)))", "@VALUE(@B1) -> SINGLE(VALUE(SINGLE(B1)))");

		// Test @ with UPPER/LOWER/PROPER functions
		fillRange = ws.getRange2("A47");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A47").getValueForEdit2();
		fragment[0].setFragmentText("=@UPPER(@LOWER(@PROPER(@B1)))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A47"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(UPPER(_xlfn.SINGLE(LOWER(_xlfn.SINGLE(PROPER(_xlfn.SINGLE(B1)))))))", "@UPPER(@LOWER(@PROPER(@B1))) -> SINGLE(UPPER(SINGLE(LOWER(SINGLE(PROPER(SINGLE(B1)))))))");

		// Test @ with REPT function
		fillRange = ws.getRange2("A48");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A48").getValueForEdit2();
		fragment[0].setFragmentText("=REPT(@B1,@C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A48"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "REPT(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "REPT(@B1,@C1) -> REPT(SINGLE(B1),SINGLE(C1))");

		// Test @ with complex nested expression
		fillRange = ws.getRange2("A49");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A49").getValueForEdit2();
		fragment[0].setFragmentText("=@IF(@AND(@B1>0,@OR(@C1<10,@D1=5)),@SUM(@B1:D1),@AVERAGE(@B1:D1))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A49"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(AND(_xlfn.SINGLE(B1)>0,_xlfn.SINGLE(OR(_xlfn.SINGLE(C1)<10,_xlfn.SINGLE(D1)=5)))),_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:D1))),_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:D1)))))", "@IF(@AND(@B1>0,@OR(@C1<10,@D1=5)),@SUM(@B1:D1),@AVERAGE(@B1:D1)) -> SINGLE(IF(SINGLE(AND(SINGLE(B1)>0,SINGLE(OR(SINGLE(C1)<10,SINGLE(D1)=5)))),SINGLE(SUM(SINGLE(B1:D1))),SINGLE(AVERAGE(SINGLE(B1:D1)))))");

		// Test @ with ROW and COLUMN functions
		fillRange = ws.getRange2("A50");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A50").getValueForEdit2();
		fragment[0].setFragmentText("=@ROW(@B1)+@COLUMN(@B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A50"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(ROW(_xlfn.SINGLE(B1)))+_xlfn.SINGLE(COLUMN(_xlfn.SINGLE(B1)))", "@ROW(@B1)+@COLUMN(@B1) -> SINGLE(ROW(SINGLE(B1)))+SINGLE(COLUMN(SINGLE(B1)))");


		// Test @ with deeply nested IFS and multiple conditions
		fillRange = ws.getRange2("A51");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A51").getValueForEdit2();
		fragment[0].setFragmentText("=@IF(@B1>100,@IF(@C1>50,@SUM(@B1:C1)*@D1,@AVERAGE(@B1:D1)),@IF(@B1<0,@ABS(@B1),@MIN(@B1:D1)))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A51"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)>100,_xlfn.SINGLE(IF(_xlfn.SINGLE(C1)>50,_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:C1)))*_xlfn.SINGLE(D1),_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:D1))))),_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)<0,_xlfn.SINGLE(ABS(_xlfn.SINGLE(B1))),_xlfn.SINGLE(MIN(_xlfn.SINGLE(B1:D1)))))))", "@IF(@B1>100,@IF(@C1>50,@SUM(@B1:C1)*@D1,@AVERAGE(@B1:D1)),@IF(@B1<0,@ABS(@B1),@MIN(@B1:D1))) -> complex nested IF");

		// Test @ with SUMPRODUCT and multiple array operations
		fillRange = ws.getRange2("A52");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A52").getValueForEdit2();
		fragment[0].setFragmentText("=@SUMPRODUCT((@B1:B10>@C1)*(@C1:C10<@D1)*(@D1:D10=@E1)*@B1:B10/@F1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A52"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SUMPRODUCT((_xlfn.SINGLE(B1:B10)>_xlfn.SINGLE(C1))*(_xlfn.SINGLE(C1:C10)<_xlfn.SINGLE(D1))*(_xlfn.SINGLE(D1:D10)=_xlfn.SINGLE(E1))*_xlfn.SINGLE(B1:B10)/_xlfn.SINGLE(F1)))", "@SUMPRODUCT with multiple @ array operations");

		// Test @ with INDEX/MATCH combination and nested functions
		fillRange = ws.getRange2("A53");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A53").getValueForEdit2();
		fragment[0].setFragmentText("=@IFERROR(@INDEX(@B1:E10,@MATCH(@MAX(@A1:A10),@A1:A10,0),@MATCH(@MIN(@F1:F10),@F1:F10,0)),@AVERAGE(@B1:E10))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A53"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IFERROR(_xlfn.SINGLE(INDEX(_xlfn.SINGLE(B1:E10),_xlfn.SINGLE(MATCH(_xlfn.SINGLE(MAX(_xlfn.SINGLE(A1:A10))),_xlfn.SINGLE(A1:A10),0)),_xlfn.SINGLE(MATCH(_xlfn.SINGLE(MIN(_xlfn.SINGLE(F1:F10))),_xlfn.SINGLE(F1:F10),0)))),_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:E10)))))", "@IFERROR(@INDEX with nested @MATCH and @MAX/@MIN");

		// Test @ with TEXT formatting and nested calculations
		fillRange = ws.getRange2("A54");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A54").getValueForEdit2();
		fragment[0].setFragmentText("=@CONCATENATE(@TEXT(@ROUND(@SUM(@B1:B10)/@COUNT(@B1:B10),2),\"#,##0.00\"),\" \",@TEXT(@MAX(@B1:B10)-@MIN(@B1:B10),\"0.00%\"))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A54"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(CONCATENATE(_xlfn.SINGLE(TEXT(_xlfn.SINGLE(ROUND(_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:B10)))/_xlfn.SINGLE(COUNT(_xlfn.SINGLE(B1:B10))),2)),\"#,##0.00\")),\" \",_xlfn.SINGLE(TEXT(_xlfn.SINGLE(MAX(_xlfn.SINGLE(B1:B10)))-_xlfn.SINGLE(MIN(_xlfn.SINGLE(B1:B10))),\"0.00%\"))))", "@CONCATENATE with @TEXT and nested arithmetic");

		// Test @ with multiple logical functions combined
		fillRange = ws.getRange2("A55");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A55").getValueForEdit2();
		fragment[0].setFragmentText("=@IF(@AND(@NOT(@ISBLANK(@B1)),@OR(@ISNUMBER(@B1),@ISTEXT(@B1)),@ISERROR(@C1)=FALSE),@VLOOKUP(@B1,@D1:F10,@IF(@B1>0,2,3),@FALSE),@NA())");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A55"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(AND(_xlfn.SINGLE(NOT(_xlfn.SINGLE(ISBLANK(_xlfn.SINGLE(B1))))),_xlfn.SINGLE(OR(_xlfn.SINGLE(ISNUMBER(_xlfn.SINGLE(B1))),_xlfn.SINGLE(ISTEXT(_xlfn.SINGLE(B1))))),_xlfn.SINGLE(ISERROR(_xlfn.SINGLE(C1)))=FALSE)),_xlfn.SINGLE(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:F10),_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)>0,2,3)),_xlfn.SINGLE(FALSE))),_xlfn.SINGLE(NA())))", "@IF with @AND, @NOT, @OR, @ISNUMBER, @ISTEXT, @ISERROR, @VLOOKUP");

		ws.getRange2("A1:Z100").cleanAll();
	});


	QUnit.module("Sheet structure");
});


