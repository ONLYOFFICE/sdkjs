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

	const parserFormula = AscCommonExcel.parserFormula;

	QUnit.test('Test @ -> single() + single() -> @', function (assert) {
		let fillRange, resCell, fragment, assembledVal;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		let formula = "=SIN(@B1)";
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIN(_xlfn.SINGLE(B1))", "SIN(@B1) -> SIN(SINGLE(B1))");
		assembledVal = ws.getRange2("A1").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SUM(@B1:B3)";
		fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A2"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(B1:B3))", "SUM(@B1:B3) -> SUM(SINGLE(B1:B3))");
		assembledVal = ws.getRange2("A2").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SUM(@3:3)";
		fillRange = ws.getRange2("A4");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A4").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A4"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(3:3))", "SUM(@3:3) -> SUM(SINGLE(3:3))");
		assembledVal = ws.getRange2("A4").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SUM(@B:B)";
		fillRange = ws.getRange2("A5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A5"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(B:B))", "SUM(@B:B) -> SUM(SINGLE(B:B))");
		assembledVal = ws.getRange2("A5").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=IF(@TRUE,1,0)";
		fillRange = ws.getRange2("A6");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A6").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A6"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(TRUE),1,0)", "IF(@TRUE,1,0) -> IF(SINGLE(TRUE),1,0)");
		assembledVal = ws.getRange2("A6").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = '=LEN(@"test")';
		fillRange = ws.getRange2("A7");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A7").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A7"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), 'LEN(_xlfn.SINGLE("test"))', 'LEN(@"test") -> LEN(SINGLE("test"))');
		assembledVal = ws.getRange2("A7").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = '=LEN(@{1,2,3})';
		fillRange = ws.getRange2("A7");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A7").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A7"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), 'LEN(_xlfn.SINGLE({1,2,3}))', 'LEN(@{1,2,3}) -> LEN(SINGLE({1,2,3}))');
		assembledVal = ws.getRange2("A7").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SUM(IF(@B1:B3>0,@B1:B3,0))";
		fillRange = ws.getRange2("A9");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A9").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A9"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(IF(_xlfn.SINGLE(B1:B3)>0,_xlfn.SINGLE(B1:B3),0))", "SUM(IF(@B1:B3>0,@B1:B3,0)) -> SUM(IF(SINGLE(B1:B3)>0,SINGLE(B1:B3),0))");
		assembledVal = ws.getRange2("A9").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=AVERAGE(IF(@B1:B3<>0,@B1:B3))";
		fillRange = ws.getRange2("A10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A10"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "AVERAGE(IF(_xlfn.SINGLE(B1:B3)<>0,_xlfn.SINGLE(B1:B3)))", "AVERAGE(IF(@B1:B3<>0,@B1:B3)) -> AVERAGE(IF(SINGLE(B1:B3)<>0,SINGLE(B1:B3)))");
		assembledVal = ws.getRange2("A10").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=IF(AND(@B1>0,@C1>0),SUM(@B1:C1),0)";
		fillRange = ws.getRange2("A11");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A11").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A11"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(AND(_xlfn.SINGLE(B1)>0,_xlfn.SINGLE(C1)>0),SUM(_xlfn.SINGLE(B1:C1)),0)", "IF(AND(@B1>0,@C1>0),SUM(@B1:C1),0) -> IF(AND(SINGLE(B1)>0,SINGLE(C1)>0),SUM(SINGLE(B1:C1)),0)");
		assembledVal = ws.getRange2("A11").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SUM(IF(OR(@B1:B3>10,@B1:B3<0),@B1:B3,0))";
		fillRange = ws.getRange2("A12");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A12").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A12"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(IF(OR(_xlfn.SINGLE(B1:B3)>10,_xlfn.SINGLE(B1:B3)<0),_xlfn.SINGLE(B1:B3),0))", "SUM(IF(OR(@B1:B3>10,@B1:B3<0),@B1:B3,0)) -> SUM(IF(OR(SINGLE(B1:B3)>10,SINGLE(B1:B3)<0),SINGLE(B1:B3),0))");
		assembledVal = ws.getRange2("A12").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ROUND(AVERAGE(@B1:B3),2)";
		fillRange = ws.getRange2("A13");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A13").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A13"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUND(AVERAGE(_xlfn.SINGLE(B1:B3)),2)", "ROUND(AVERAGE(@B1:B3),2) -> ROUND(AVERAGE(SINGLE(B1:B3)),2)");
		assembledVal = ws.getRange2("A13").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=VLOOKUP(@B1,@D1:E10,2,FALSE)";
		fillRange = ws.getRange2("A14");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A14").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A14"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:E10),2,FALSE)", "VLOOKUP(@B1,@D1:E10,2,FALSE) -> VLOOKUP(SINGLE(B1),SINGLE(D1:E10),2,FALSE)");
		assembledVal = ws.getRange2("A14").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SUMPRODUCT((@B1:B3>5)*(@C1:C3<10)*@B1:B3)";
		fillRange = ws.getRange2("A15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A15"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUMPRODUCT((_xlfn.SINGLE(B1:B3)>5)*(_xlfn.SINGLE(C1:C3)<10)*_xlfn.SINGLE(B1:B3))", "SUMPRODUCT((@B1:B3>5)*(@C1:C3<10)*@B1:B3) -> SUMPRODUCT((SINGLE(B1:B3)>5)*(SINGLE(C1:C3)<10)*SINGLE(B1:B3))");
		assembledVal = ws.getRange2("A15").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=IF(@B1>0,IF(@C1>0,SUM(@B1:C1),@B1),0)";
		fillRange = ws.getRange2("A16");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A16").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A16"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(B1)>0,IF(_xlfn.SINGLE(C1)>0,SUM(_xlfn.SINGLE(B1:C1)),_xlfn.SINGLE(B1)),0)", "IF(@B1>0,IF(@C1>0,SUM(@B1:C1),@B1),0) -> IF(SINGLE(B1)>0,IF(SINGLE(C1)>0,SUM(SINGLE(B1:C1)),SINGLE(B1)),0)");
		assembledVal = ws.getRange2("A16").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=IFERROR(VLOOKUP(@B1,@D1:E10,2,FALSE),@B1*2)";
		fillRange = ws.getRange2("A17");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A17").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A17"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IFERROR(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:E10),2,FALSE),_xlfn.SINGLE(B1)*2)", "IFERROR(VLOOKUP(@B1,@D1:E10,2,FALSE),@B1*2) -> IFERROR(VLOOKUP(SINGLE(B1),SINGLE(D1:E10),2,FALSE),SINGLE(B1)*2)");
		assembledVal = ws.getRange2("A17").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=INDEX(@B1:B10,MATCH(MAX(@B1:B10),@B1:B10,0))";
		fillRange = ws.getRange2("A19");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A19").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A19"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "INDEX(_xlfn.SINGLE(B1:B10),MATCH(MAX(_xlfn.SINGLE(B1:B10)),_xlfn.SINGLE(B1:B10),0))", "INDEX(@B1:B10,MATCH(MAX(@B1:B10),@B1:B10,0)) -> INDEX(SINGLE(B1:B10),MATCH(MAX(SINGLE(B1:B10)),SINGLE(B1:B10),0))");
		assembledVal = ws.getRange2("A19").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=CONCATENATE(@B1,\" \",@C1,\" \",@D1)";
		fillRange = ws.getRange2("A20");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A20").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A20"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CONCATENATE(_xlfn.SINGLE(B1),\" \",_xlfn.SINGLE(C1),\" \",_xlfn.SINGLE(D1))", "CONCATENATE(@B1,\" \",@C1,\" \",@D1) -> CONCATENATE(SINGLE(B1),\" \",SINGLE(C1),\" \",SINGLE(D1))");
		assembledVal = ws.getRange2("A20").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@SIN(@B1)";
		fillRange = ws.getRange2("A9");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A9").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A9"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SIN(_xlfn.SINGLE(B1)))", "@SIN(@B1) -> SINGLE(SIN(SINGLE(B1)))");
		assembledVal = ws.getRange2("A9").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@SUM(@B1:B3)";
		fillRange = ws.getRange2("A10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A10"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:B3)))", "@SUM(@B1:B3) -> SINGLE(SUM(SINGLE(B1:B3)))");
		assembledVal = ws.getRange2("A10").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@IF(@B1>0,@C1,0)";
		fillRange = ws.getRange2("A11");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A11").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A11"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)>0,_xlfn.SINGLE(C1),0))", "@IF(@B1>0,@C1,0) -> SINGLE(IF(SINGLE(B1)>0,SINGLE(C1),0))");
		assembledVal = ws.getRange2("A11").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@ROUND(@AVERAGE(@B1:B3),2)";
		fillRange = ws.getRange2("A12");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A12").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A12"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(ROUND(_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:B3))),2))", "@ROUND(@AVERAGE(@B1:B3),2) -> SINGLE(ROUND(SINGLE(AVERAGE(SINGLE(B1:B3))),2))");
		assembledVal = ws.getRange2("A12").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@ABS(@MIN(@B1:B3))";
		fillRange = ws.getRange2("A13");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A13").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A13"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(ABS(_xlfn.SINGLE(MIN(_xlfn.SINGLE(B1:B3)))))", "@ABS(@MIN(@B1:B3)) -> SINGLE(ABS(SINGLE(MIN(SINGLE(B1:B3)))))");
		assembledVal = ws.getRange2("A13").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@SQRT(@ABS(@B1))";
		fillRange = ws.getRange2("A14");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A14").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A14"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SQRT(_xlfn.SINGLE(ABS(_xlfn.SINGLE(B1)))))", "@SQRT(@ABS(@B1)) -> SINGLE(SQRT(SINGLE(ABS(SINGLE(B1)))))");
		assembledVal = ws.getRange2("A14").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@IF(@AND(@B1>0,@C1>0),@SUM(@B1:C1),0)";
		fillRange = ws.getRange2("A15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A15"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(AND(_xlfn.SINGLE(B1)>0,_xlfn.SINGLE(C1)>0)),_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:C1))),0))", "@IF(@AND(@B1>0,@C1>0),@SUM(@B1:C1),0) -> SINGLE(IF(SINGLE(AND(SINGLE(B1)>0,SINGLE(C1)>0)),SINGLE(SUM(SINGLE(B1:C1))),0))");
		assembledVal = ws.getRange2("A15").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@VLOOKUP(@B1,@D1:E10,2,FALSE)";
		fillRange = ws.getRange2("A16");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A16").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A16"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:E10),2,FALSE))", "@VLOOKUP(@B1,@D1:E10,2,FALSE) -> SINGLE(VLOOKUP(SINGLE(B1),SINGLE(D1:E10),2,FALSE))");
		assembledVal = ws.getRange2("A16").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@IFERROR(@VLOOKUP(@B1,@D1:E10,2,FALSE),@B1*2)";
		fillRange = ws.getRange2("A17");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A17").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A17"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IFERROR(_xlfn.SINGLE(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:E10),2,FALSE)),_xlfn.SINGLE(B1)*2))", "@IFERROR(@VLOOKUP(@B1,@D1:E10,2,FALSE),@B1*2) -> SINGLE(IFERROR(SINGLE(VLOOKUP(SINGLE(B1),SINGLE(D1:E10),2,FALSE)),SINGLE(B1)*2))");
		assembledVal = ws.getRange2("A17").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@TEXT(@ROUND(@AVERAGE(@B1:B3),2),\"0.00\")";
		fillRange = ws.getRange2("A18");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A18").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A18"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(TEXT(_xlfn.SINGLE(ROUND(_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:B3))),2)),\"0.00\"))", "@TEXT(@ROUND(@AVERAGE(@B1:B3),2),\"0.00\") -> SINGLE(TEXT(SINGLE(ROUND(SINGLE(AVERAGE(SINGLE(B1:B3))),2)),\"0.00\"))");
		assembledVal = ws.getRange2("A18").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);


		fillRange = ws.getRange2("A18");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A18").getValueForEdit2();
		fragment[0].setFragmentText("=@TEXT(@ROUND(@AVERAGE(@B1:B3),2),\"0.00\")");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A18"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(TEXT(_xlfn.SINGLE(ROUND(_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:B3))),2)),\"0.00\"))", "@TEXT(@ROUND(@AVERAGE(@B1:B3),2),\"0.00\") -> SINGLE(TEXT(SINGLE(ROUND(SINGLE(AVERAGE(SINGLE(B1:B3))),2)),\"0.00\"))");

		// Additional test cases for @ operator

		// Test @ with numbers
		formula = "=@123";
		fillRange = ws.getRange2("A21");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A21").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A21"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(123)", "@123 -> SINGLE(123)");
		assembledVal = ws.getRange2("A21").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with negative numbers
		formula = "=SUM(@-5,@B1)";
		fillRange = ws.getRange2("A22");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A22").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A22"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(-5),_xlfn.SINGLE(B1))", "SUM(@-5,@B1) -> SUM(SINGLE(-5),SINGLE(B1))");
		assembledVal = ws.getRange2("A22").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with error values
		formula = "=IFERROR(@#N/A,0)";
		fillRange = ws.getRange2("A23");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A23").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A23"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IFERROR(_xlfn.SINGLE(#N/A),0)", "IFERROR(@#N/A,0) -> IFERROR(SINGLE(#N/A),0)");
		assembledVal = ws.getRange2("A23").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with 3D references
		formula = "=SUM(@Sheet1!A1:A10)";
		fillRange = ws.getRange2("A24");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A24").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A24"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(Sheet1!A1:A10))", "SUM(@Sheet1!A1:A10) -> SUM(SINGLE(Sheet1!A1:A10))");
		assembledVal = ws.getRange2("A24").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with CHOOSE function
		formula = "=CHOOSE(@B1,@C1,@D1,@E1)";
		fillRange = ws.getRange2("A25");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A25").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A25"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CHOOSE(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1),_xlfn.SINGLE(E1))", "CHOOSE(@B1,@C1,@D1,@E1) -> CHOOSE(SINGLE(B1),SINGLE(C1),SINGLE(D1),SINGLE(E1))");
		assembledVal = ws.getRange2("A25").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with COUNT functions
		formula = "=COUNTIF(@B1:B10,@C1)";
		fillRange = ws.getRange2("A26");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A26").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A26"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COUNTIF(_xlfn.SINGLE(B1:B10),_xlfn.SINGLE(C1))", "COUNTIF(@B1:B10,@C1) -> COUNTIF(SINGLE(B1:B10),SINGLE(C1))");
		assembledVal = ws.getRange2("A26").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with nested MAX/MIN
		formula = "=@MAX(@MIN(@B1:B10),@MIN(@C1:C10))";
		fillRange = ws.getRange2("A27");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A27").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A27"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(MAX(_xlfn.SINGLE(MIN(_xlfn.SINGLE(B1:B10))),_xlfn.SINGLE(MIN(_xlfn.SINGLE(C1:C10)))))", "@MAX(@MIN(@B1:B10),@MIN(@C1:C10)) -> SINGLE(MAX(SINGLE(MIN(SINGLE(B1:B10))),SINGLE(MIN(SINGLE(C1:C10)))))");
		assembledVal = ws.getRange2("A27").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with LEFT/RIGHT/MID string functions
		formula = "=LEFT(@B1,@C1)";
		fillRange = ws.getRange2("A28");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A28").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A28"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LEFT(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "LEFT(@B1,@C1) -> LEFT(SINGLE(B1),SINGLE(C1))");
		assembledVal = ws.getRange2("A28").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with MID function
		formula = "=MID(@B1,@C1,@D1)";
		fillRange = ws.getRange2("A29");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A29").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A29"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MID(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))", "MID(@B1,@C1,@D1) -> MID(SINGLE(B1),SINGLE(C1),SINGLE(D1))");
		assembledVal = ws.getRange2("A29").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with DATE/TIME functions
		formula = "=DATE(@B1,@C1,@D1)";
		fillRange = ws.getRange2("A30");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A30").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A30"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DATE(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))", "DATE(@B1,@C1,@D1) -> DATE(SINGLE(B1),SINGLE(C1),SINGLE(D1))");
		assembledVal = ws.getRange2("A30").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with arithmetic operations between @ operands
		formula = "=@B1+@C1*@D1";
		fillRange = ws.getRange2("A31");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A31").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A31"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(B1)+_xlfn.SINGLE(C1)*_xlfn.SINGLE(D1)", "@B1+@C1*@D1 -> SINGLE(B1)+SINGLE(C1)*SINGLE(D1)");
		assembledVal = ws.getRange2("A31").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with comparison operators
		formula = "=IF(@B1>=@C1,@B1,@C1)";
		fillRange = ws.getRange2("A32");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A32").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A32"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(B1)>=_xlfn.SINGLE(C1),_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "IF(@B1>=@C1,@B1,@C1) -> IF(SINGLE(B1)>=SINGLE(C1),SINGLE(B1),SINGLE(C1))");
		assembledVal = ws.getRange2("A32").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with POWER function
		formula = "=POWER(@B1,@C1)";
		fillRange = ws.getRange2("A33");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A33").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A33"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "POWER(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "POWER(@B1,@C1) -> POWER(SINGLE(B1),SINGLE(C1))");
		assembledVal = ws.getRange2("A33").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with MOD function
		formula = "=MOD(@B1,@C1)";
		fillRange = ws.getRange2("A34");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A34").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A34"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MOD(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "MOD(@B1,@C1) -> MOD(SINGLE(B1),SINGLE(C1))");
		assembledVal = ws.getRange2("A34").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with SUBSTITUTE function
		formula = "=SUBSTITUTE(@B1,@C1,@D1)";
		fillRange = ws.getRange2("A35");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A35").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A35"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUBSTITUTE(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))", "SUBSTITUTE(@B1,@C1,@D1) -> SUBSTITUTE(SINGLE(B1),SINGLE(C1),SINGLE(D1))");
		assembledVal = ws.getRange2("A35").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with INDIRECT (dynamic reference)
		formula = "=@INDIRECT(@B1)";
		fillRange = ws.getRange2("A36");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A36").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A36"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(INDIRECT(_xlfn.SINGLE(B1)))", "@INDIRECT(@B1) -> SINGLE(INDIRECT(SINGLE(B1)))");
		assembledVal = ws.getRange2("A36").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with OFFSET function
		formula = "=SUM(@OFFSET(@B1,@C1,@D1))";
		fillRange = ws.getRange2("A37");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A37").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A37"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUM(_xlfn.SINGLE(OFFSET(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))))", "SUM(@OFFSET(@B1,@C1,@D1)) -> SUM(SINGLE(OFFSET(SINGLE(B1),SINGLE(C1),SINGLE(D1))))");
		assembledVal = ws.getRange2("A37").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with LARGE/SMALL functions
		formula = "=LARGE(@B1:B10,@C1)";
		fillRange = ws.getRange2("A38");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A38").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A38"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LARGE(_xlfn.SINGLE(B1:B10),_xlfn.SINGLE(C1))", "LARGE(@B1:B10,@C1) -> LARGE(SINGLE(B1:B10),SINGLE(C1))");
		assembledVal = ws.getRange2("A38").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with SMALL function
		formula = "=@SMALL(@B1:B10,@C1)";
		fillRange = ws.getRange2("A39");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A39").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A39"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SMALL(_xlfn.SINGLE(B1:B10),_xlfn.SINGLE(C1)))", "@SMALL(@B1:B10,@C1) -> SINGLE(SMALL(SINGLE(B1:B10),SINGLE(C1)))");
		assembledVal = ws.getRange2("A39").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with LOOKUP function
		formula = "=LOOKUP(@B1,@C1:C10,@D1:D10)";
		fillRange = ws.getRange2("A40");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A40").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A40"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1:C10),_xlfn.SINGLE(D1:D10))", "LOOKUP(@B1,@C1:C10,@D1:D10) -> LOOKUP(SINGLE(B1),SINGLE(C1:C10),SINGLE(D1:D10))");
		assembledVal = ws.getRange2("A40").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with HLOOKUP function
		formula = "=HLOOKUP(@B1,@C1:G3,@D1,FALSE)";
		fillRange = ws.getRange2("A41");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A41").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A41"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "HLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1:G3),_xlfn.SINGLE(D1),FALSE)", "HLOOKUP(@B1,@C1:G3,@D1,FALSE) -> HLOOKUP(SINGLE(B1),SINGLE(C1:G3),SINGLE(D1),FALSE)");
		assembledVal = ws.getRange2("A41").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with NOT function
		formula = "=IF(@NOT(@B1),@C1,@D1)";
		fillRange = ws.getRange2("A42");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A42").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A42"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(NOT(_xlfn.SINGLE(B1))),_xlfn.SINGLE(C1),_xlfn.SINGLE(D1))", "IF(@NOT(@B1),@C1,@D1) -> IF(SINGLE(NOT(SINGLE(B1))),SINGLE(C1),SINGLE(D1))");
		assembledVal = ws.getRange2("A42").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with ISBLANK function
		formula = "=IF(@ISBLANK(@B1),@C1,@B1)";
		fillRange = ws.getRange2("A43");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A43").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A43"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(_xlfn.SINGLE(ISBLANK(_xlfn.SINGLE(B1))),_xlfn.SINGLE(C1),_xlfn.SINGLE(B1))", "IF(@ISBLANK(@B1),@C1,@B1) -> IF(SINGLE(ISBLANK(SINGLE(B1))),SINGLE(C1),SINGLE(B1))");
		assembledVal = ws.getRange2("A43").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with ISNUMBER function
		formula = "=@ISNUMBER(@B1)";
		fillRange = ws.getRange2("A44");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A44").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A44"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(ISNUMBER(_xlfn.SINGLE(B1)))", "@ISNUMBER(@B1) -> SINGLE(ISNUMBER(SINGLE(B1)))");
		assembledVal = ws.getRange2("A44").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with TRIM and CLEAN functions
		formula = "=@TRIM(@CLEAN(@B1))";
		fillRange = ws.getRange2("A45");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A45").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A45"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(TRIM(_xlfn.SINGLE(CLEAN(_xlfn.SINGLE(B1)))))", "@TRIM(@CLEAN(@B1)) -> SINGLE(TRIM(SINGLE(CLEAN(SINGLE(B1)))))");
		assembledVal = ws.getRange2("A45").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with VALUE function
		formula = "=@VALUE(@B1)";
		fillRange = ws.getRange2("A46");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A46").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A46"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(VALUE(_xlfn.SINGLE(B1)))", "@VALUE(@B1) -> SINGLE(VALUE(SINGLE(B1)))");
		assembledVal = ws.getRange2("A46").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with UPPER/LOWER/PROPER functions
		formula = "=@UPPER(@LOWER(@PROPER(@B1)))";
		fillRange = ws.getRange2("A47");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A47").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A47"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(UPPER(_xlfn.SINGLE(LOWER(_xlfn.SINGLE(PROPER(_xlfn.SINGLE(B1)))))))", "@UPPER(@LOWER(@PROPER(@B1))) -> SINGLE(UPPER(SINGLE(LOWER(SINGLE(PROPER(SINGLE(B1)))))))");
		assembledVal = ws.getRange2("A47").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with REPT function
		formula = "=REPT(@B1,@C1)";
		fillRange = ws.getRange2("A48");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A48").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A48"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "REPT(_xlfn.SINGLE(B1),_xlfn.SINGLE(C1))", "REPT(@B1,@C1) -> REPT(SINGLE(B1),SINGLE(C1))");
		assembledVal = ws.getRange2("A48").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with complex nested expression
		formula = "=@IF(@AND(@B1>0,@OR(@C1<10,@D1=5)),@SUM(@B1:D1),@AVERAGE(@B1:D1))";
		fillRange = ws.getRange2("A49");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A49").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A49"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(AND(_xlfn.SINGLE(B1)>0,_xlfn.SINGLE(OR(_xlfn.SINGLE(C1)<10,_xlfn.SINGLE(D1)=5)))),_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:D1))),_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:D1)))))", "@IF(@AND(@B1>0,@OR(@C1<10,@D1=5)),@SUM(@B1:D1),@AVERAGE(@B1:D1)) -> SINGLE(IF(SINGLE(AND(SINGLE(B1)>0,SINGLE(OR(SINGLE(C1)<10,SINGLE(D1)=5)))),SINGLE(SUM(SINGLE(B1:D1))),SINGLE(AVERAGE(SINGLE(B1:D1)))))");
		assembledVal = ws.getRange2("A49").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with ROW and COLUMN functions
		formula = "=@ROW(@B1)+@COLUMN(@B1)";
		fillRange = ws.getRange2("A50");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A50").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A50"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(ROW(_xlfn.SINGLE(B1)))+_xlfn.SINGLE(COLUMN(_xlfn.SINGLE(B1)))", "@ROW(@B1)+@COLUMN(@B1) -> SINGLE(ROW(SINGLE(B1)))+SINGLE(COLUMN(SINGLE(B1)))");
		assembledVal = ws.getRange2("A50").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with deeply nested IFS and multiple conditions
		formula = "=@IF(@B1>100,@IF(@C1>50,@SUM(@B1:C1)*@D1,@AVERAGE(@B1:D1)),@IF(@B1<0,@ABS(@B1),@MIN(@B1:D1)))";
		fillRange = ws.getRange2("A51");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A51").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A51"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)>100,_xlfn.SINGLE(IF(_xlfn.SINGLE(C1)>50,_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:C1)))*_xlfn.SINGLE(D1),_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:D1))))),_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)<0,_xlfn.SINGLE(ABS(_xlfn.SINGLE(B1))),_xlfn.SINGLE(MIN(_xlfn.SINGLE(B1:D1)))))))", "@IF(@B1>100,@IF(@C1>50,@SUM(@B1:C1)*@D1,@AVERAGE(@B1:D1)),@IF(@B1<0,@ABS(@B1),@MIN(@B1:D1))) -> complex nested IF");
		assembledVal = ws.getRange2("A51").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);


		// Test @ with SUMPRODUCT and multiple array operations
		formula = "=@SUMPRODUCT((@B1:B10>@C1)*(@C1:C10<@D1)*(@D1:D10=@E1)*@B1:B10/@F1)";
		fillRange = ws.getRange2("A52");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A52").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A52"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(SUMPRODUCT((_xlfn.SINGLE(B1:B10)>_xlfn.SINGLE(C1))*(_xlfn.SINGLE(C1:C10)<_xlfn.SINGLE(D1))*(_xlfn.SINGLE(D1:D10)=_xlfn.SINGLE(E1))*_xlfn.SINGLE(B1:B10)/_xlfn.SINGLE(F1)))", "@SUMPRODUCT with multiple @ array operations");
		assembledVal = ws.getRange2("A52").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with INDEX/MATCH combination and nested functions
		formula = "=@IFERROR(@INDEX(@B1:E10,@MATCH(@MAX(@A1:A10),@A1:A10,0),@MATCH(@MIN(@F1:F10),@F1:F10,0)),@AVERAGE(@B1:E10))";
		fillRange = ws.getRange2("A53");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A53").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A53"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IFERROR(_xlfn.SINGLE(INDEX(_xlfn.SINGLE(B1:E10),_xlfn.SINGLE(MATCH(_xlfn.SINGLE(MAX(_xlfn.SINGLE(A1:A10))),_xlfn.SINGLE(A1:A10),0)),_xlfn.SINGLE(MATCH(_xlfn.SINGLE(MIN(_xlfn.SINGLE(F1:F10))),_xlfn.SINGLE(F1:F10),0)))),_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:E10)))))", "@IFERROR(@INDEX with nested @MATCH and @MAX/@MIN");
		assembledVal = ws.getRange2("A53").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with TEXT formatting and nested calculations
		formula = "=@CONCATENATE(@TEXT(@ROUND(@SUM(@B1:B10)/@COUNT(@B1:B10),2),\"#,##0.00\"),\" \",@TEXT(@MAX(@B1:B10)-@MIN(@B1:B10),\"0.00%\"))";
		fillRange = ws.getRange2("A54");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A54").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A54"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(CONCATENATE(_xlfn.SINGLE(TEXT(_xlfn.SINGLE(ROUND(_xlfn.SINGLE(SUM(_xlfn.SINGLE(B1:B10)))/_xlfn.SINGLE(COUNT(_xlfn.SINGLE(B1:B10))),2)),\"#,##0.00\")),\" \",_xlfn.SINGLE(TEXT(_xlfn.SINGLE(MAX(_xlfn.SINGLE(B1:B10)))-_xlfn.SINGLE(MIN(_xlfn.SINGLE(B1:B10))),\"0.00%\"))))", "@CONCATENATE with @TEXT and nested arithmetic");
		assembledVal = ws.getRange2("A54").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with multiple logical functions combined
		formula = "=@IF(@AND(@NOT(@ISBLANK(@B1)),@OR(@ISNUMBER(@B1),@ISTEXT(@B1)),@ISERROR(@C1)=FALSE),@VLOOKUP(@B1,@D1:F10,@IF(@B1>0,2,3),@FALSE),@NA())";
		fillRange = ws.getRange2("A55");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A55").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A55"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(AND(_xlfn.SINGLE(NOT(_xlfn.SINGLE(ISBLANK(_xlfn.SINGLE(B1))))),_xlfn.SINGLE(OR(_xlfn.SINGLE(ISNUMBER(_xlfn.SINGLE(B1))),_xlfn.SINGLE(ISTEXT(_xlfn.SINGLE(B1))))),_xlfn.SINGLE(ISERROR(_xlfn.SINGLE(C1)))=FALSE)),_xlfn.SINGLE(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:F10),_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)>0,2,3)),_xlfn.SINGLE(FALSE))),_xlfn.SINGLE(NA())))", "@IF with @AND, @NOT, @OR, @ISNUMBER, @ISTEXT, @ISERROR, @VLOOKUP");
		assembledVal = ws.getRange2("A55").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=@IF(@AND(@NOT(@ISBLANK(@B1)),@OR(@ISNUMBER(@B1),@ISTEXT(@B1)),@ISERROR(@C1)=FALSE),@VLOOKUP(@B1,@D1:F10,@IF(@B1>0,2,3),@FALSE),@NA())";
		fillRange = ws.getRange2("A55");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A55").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A55"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IF(_xlfn.SINGLE(AND(_xlfn.SINGLE(NOT(_xlfn.SINGLE(ISBLANK(_xlfn.SINGLE(B1))))),_xlfn.SINGLE(OR(_xlfn.SINGLE(ISNUMBER(_xlfn.SINGLE(B1))),_xlfn.SINGLE(ISTEXT(_xlfn.SINGLE(B1))))),_xlfn.SINGLE(ISERROR(_xlfn.SINGLE(C1)))=FALSE)),_xlfn.SINGLE(VLOOKUP(_xlfn.SINGLE(B1),_xlfn.SINGLE(D1:F10),_xlfn.SINGLE(IF(_xlfn.SINGLE(B1)>0,2,3)),_xlfn.SINGLE(FALSE))),_xlfn.SINGLE(NA())))", "@IF with @AND, @NOT, @OR, @ISNUMBER, @ISTEXT, @ISERROR, @VLOOKUP");
		assembledVal = ws.getRange2("A55").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		// Test @ with extreme deep nesting - 10+ levels of function calls
		formula = "=@IFERROR(@IF(@AND(@NOT(@ISBLANK(@INDIRECT(@TEXT(@ROW(@B1),\"0\")&\":\"&@TEXT(@COLUMN(@B1),\"0\")))),@OR(@ISNUMBER(@INDEX(@B1:F10,@MATCH(@MAX(@A1:A10),@A1:A10,0),@MATCH(@MIN(@G1:G10),@G1:G10,0))),@ISTEXT(@VLOOKUP(@SMALL(@B1:B10,@INT(@AVERAGE(@C1:C10))),@D1:F10,@MOD(@ABS(@SUM(@E1:E10)),3)+1,@FALSE)))),@SUMPRODUCT((@B1:B10>@PERCENTILE(@B1:B10,0.5))*(@C1:C10<@QUARTILE(@C1:C10,3))*@IF(@COUNTIF(@D1:D10,\">\"&@MEDIAN(@D1:D10))>0,@D1:D10/@STDEV(@D1:D10),1)),@CONCATENATE(@LEFT(@TEXT(@ROUND(@AVERAGE(@B1:B10),@INT(@SQRT(@COUNT(@B1:B10)))),\"#,##0.00\"),5),@MID(@TEXT(@VAR(@C1:C10),\"0.00E+00\"),1,@LEN(@TEXT(@VAR(@C1:C10),\"0.00E+00\"))-2))),@NA())";
		fillRange = ws.getRange2("A56");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A56").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A56"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SINGLE(IFERROR(_xlfn.SINGLE(IF(_xlfn.SINGLE(AND(_xlfn.SINGLE(NOT(_xlfn.SINGLE(ISBLANK(_xlfn.SINGLE(INDIRECT(_xlfn.SINGLE(TEXT(_xlfn.SINGLE(ROW(_xlfn.SINGLE(B1))),\"0\"))&\":\"&_xlfn.SINGLE(TEXT(_xlfn.SINGLE(COLUMN(_xlfn.SINGLE(B1))),\"0\")))))))),_xlfn.SINGLE(OR(_xlfn.SINGLE(ISNUMBER(_xlfn.SINGLE(INDEX(_xlfn.SINGLE(B1:F10),_xlfn.SINGLE(MATCH(_xlfn.SINGLE(MAX(_xlfn.SINGLE(A1:A10))),_xlfn.SINGLE(A1:A10),0)),_xlfn.SINGLE(MATCH(_xlfn.SINGLE(MIN(_xlfn.SINGLE(G1:G10))),_xlfn.SINGLE(G1:G10),0)))))),_xlfn.SINGLE(ISTEXT(_xlfn.SINGLE(VLOOKUP(_xlfn.SINGLE(SMALL(_xlfn.SINGLE(B1:B10),_xlfn.SINGLE(INT(_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(C1:C10))))))),_xlfn.SINGLE(D1:F10),_xlfn.SINGLE(MOD(_xlfn.SINGLE(ABS(_xlfn.SINGLE(SUM(_xlfn.SINGLE(E1:E10))))),3))+1,_xlfn.SINGLE(FALSE))))))))),_xlfn.SINGLE(SUMPRODUCT((_xlfn.SINGLE(B1:B10)>_xlfn.SINGLE(PERCENTILE(_xlfn.SINGLE(B1:B10),0.5)))*(_xlfn.SINGLE(C1:C10)<_xlfn.SINGLE(QUARTILE(_xlfn.SINGLE(C1:C10),3)))*_xlfn.SINGLE(IF(_xlfn.SINGLE(COUNTIF(_xlfn.SINGLE(D1:D10),\">\"&_xlfn.SINGLE(MEDIAN(_xlfn.SINGLE(D1:D10)))))>0,_xlfn.SINGLE(D1:D10)/_xlfn.SINGLE(STDEV(_xlfn.SINGLE(D1:D10))),1)))),_xlfn.SINGLE(CONCATENATE(_xlfn.SINGLE(LEFT(_xlfn.SINGLE(TEXT(_xlfn.SINGLE(ROUND(_xlfn.SINGLE(AVERAGE(_xlfn.SINGLE(B1:B10))),_xlfn.SINGLE(INT(_xlfn.SINGLE(SQRT(_xlfn.SINGLE(COUNT(_xlfn.SINGLE(B1:B10))))))))),\"#,##0.00\")),5)),_xlfn.SINGLE(MID(_xlfn.SINGLE(TEXT(_xlfn.SINGLE(VAR(_xlfn.SINGLE(C1:C10))),\"0.00E+00\")),1,_xlfn.SINGLE(LEN(_xlfn.SINGLE(TEXT(_xlfn.SINGLE(VAR(_xlfn.SINGLE(C1:C10))),\"0.00E+00\"))))-2)))))),_xlfn.SINGLE(NA())))", "Extreme deep nesting with 10+ levels of @ functions");
		assembledVal = ws.getRange2("A56").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=GCD(@P4:P13,@Q4:Q13)";
		fillRange = ws.getRange2("A84");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A84").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A84"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "GCD(_xlfn.SINGLE(P4:P13),_xlfn.SINGLE(Q4:Q13))", "GCD(@P4:P13,@Q4:Q13) -> GCD(SINGLE(P4:P13),SINGLE(Q4:Q13))");
		assembledVal = ws.getRange2("A84").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=LCM(@R4:R13,@S4:S13)";
		fillRange = ws.getRange2("A85");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A85").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A85"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LCM(_xlfn.SINGLE(R4:R13),_xlfn.SINGLE(S4:S13))", "LCM(@R4:R13,@S4:S13) -> LCM(SINGLE(R4:R13),SINGLE(S4:S13))");
		assembledVal = ws.getRange2("A85").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=MROUND(@T4:T13,@U4:U13)";
		fillRange = ws.getRange2("A86");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A86").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A86"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MROUND(_xlfn.SINGLE(T4:T13),_xlfn.SINGLE(U4:U13))", "MROUND(@T4:T13,@U4:U13) -> MROUND(SINGLE(T4:T13),SINGLE(U4:U13))");
		assembledVal = ws.getRange2("A86").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=WEEKNUM(@O5:O14)";
		fillRange = ws.getRange2("A106");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A106").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A106"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "WEEKNUM(_xlfn.SINGLE(O5:O14))", "WEEKNUM(@O5:O14) -> WEEKNUM(SINGLE(O5:O14))");
		assembledVal = ws.getRange2("A106").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=EOMONTH(@X5:X14,@Y5:Y14)";
		fillRange = ws.getRange2("A111");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A111").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A111"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EOMONTH(_xlfn.SINGLE(X5:X14),_xlfn.SINGLE(Y5:Y14))", "EOMONTH(@X5:X14,@Y5:Y14) -> EOMONTH(SINGLE(X5:X14),SINGLE(Y5:Y14))");
		assembledVal = ws.getRange2("A111").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISEVEN(@E4:E13)";
		fillRange = ws.getRange2("A74");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A74").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A74"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISEVEN(_xlfn.SINGLE(E4:E13))", "ISEVEN(@E4:E13) -> ISEVEN(SINGLE(E4:E13))");
		assembledVal = ws.getRange2("A74").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISODD(@F4:F13)";
		fillRange = ws.getRange2("A75");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A75").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A75"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISODD(_xlfn.SINGLE(F4:F13))", "ISODD(@F4:F13) -> ISODD(SINGLE(F4:F13))");
		assembledVal = ws.getRange2("A75").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=N(@G4:G13)";
		fillRange = ws.getRange2("A76");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A76").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A76"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "N(_xlfn.SINGLE(G4:G13))", "N(@G4:G13) -> N(SINGLE(G4:G13))");
		assembledVal = ws.getRange2("A76").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=T(@H4:H13)";
		fillRange = ws.getRange2("A77");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A77").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A77"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "T(_xlfn.SINGLE(H4:H13))", "T(@H4:H13) -> T(SINGLE(H4:H13))");
		assembledVal = ws.getRange2("A77").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISREF(@J4:J13)";
		fillRange = ws.getRange2("A79");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A79").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A79"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISREF(_xlfn.SINGLE(J4:J13))", "ISREF(@J4:J13) -> ISREF(SINGLE(J4:J13))");
		assembledVal = ws.getRange2("A79").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=QUOTIENT(@N4:N13,@O4:O13)";
		fillRange = ws.getRange2("A83");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A83").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A83"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "QUOTIENT(_xlfn.SINGLE(N4:N13),_xlfn.SINGLE(O4:O13))", "QUOTIENT(@N4:N13,@O4:O13) -> QUOTIENT(SINGLE(N4:N13),SINGLE(O4:O13))");
		assembledVal = ws.getRange2("A83").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=EDATE(@V5:V14,@W5:W14)";
		fillRange = ws.getRange2("A110");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A110").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A110"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EDATE(_xlfn.SINGLE(V5:V14),_xlfn.SINGLE(W5:W14))", "EDATE(@V5:V14,@W5:W14) -> EDATE(SINGLE(V5:V14),SINGLE(W5:W14))");
		assembledVal = ws.getRange2("A110").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=NETWORKDAYS(@Z5:Z14,@A6:A15)";
		fillRange = ws.getRange2("A112");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A112").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A112"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "NETWORKDAYS(_xlfn.SINGLE(Z5:Z14),_xlfn.SINGLE(A6:A15))", "NETWORKDAYS(@Z5:Z14,@A6:A15) -> NETWORKDAYS(SINGLE(Z5:Z14),SINGLE(A6:A15))");
		assembledVal = ws.getRange2("A112").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=WORKDAY(@B6:B15,@C6:C15)";
		fillRange = ws.getRange2("A113");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A113").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A113"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "WORKDAY(_xlfn.SINGLE(B6:B15),_xlfn.SINGLE(C6:C15))", "WORKDAY(@B6:B15,@C6:C15) -> WORKDAY(SINGLE(B6:B15),SINGLE(C6:C15))");
		assembledVal = ws.getRange2("A113").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=YEARFRAC(@D6:D15,@E6:E15)";
		fillRange = ws.getRange2("A114");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A114").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A114"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "YEARFRAC(_xlfn.SINGLE(D6:D15),_xlfn.SINGLE(E6:E15))", "YEARFRAC(@D6:D15,@E6:E15) -> YEARFRAC(SINGLE(D6:D15),SINGLE(E6:E15))");
		assembledVal = ws.getRange2("A114").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=FORMULATEXT(@F6:F15)";
		fillRange = ws.getRange2("A115");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A115").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A115"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.FORMULATEXT(_xlfn.SINGLE(F6:F15))", "FORMULATEXT(@F6:F15) -> _xlfn.FORMULATEXT(SINGLE(F6:F15))");
		assembledVal = ws.getRange2("A115").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISFORMULA(@G6:G15)";
		fillRange = ws.getRange2("A116");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A116").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A116"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.ISFORMULA(_xlfn.SINGLE(G6:G15))", "ISFORMULA(@G6:G15) -> _xlfn.ISFORMULA(SINGLE(G6:G15))");
		assembledVal = ws.getRange2("A116").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SHEET(@H6:H15)";
		fillRange = ws.getRange2("A117");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A117").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A117"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SHEET(_xlfn.SINGLE(H6:H15))", "SHEET(@H6:H15) -> _xlfn.SHEET(_xlfn.SINGLE(H6:H15))");
		assembledVal = ws.getRange2("A117").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		ws.getRange2("A1:Z100").cleanAll();
	});


	QUnit.test('Test @ -> not single() -> exceptions', function (assert) {
		let fillRange, resCell, fragment, assembledVal;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		let formula = "=SIN(@A1:B1)";
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIN(A1:B1)", "SIN(@A1:B1) -> SIN(A1:B1)");
		assembledVal = ws.getRange2("A1").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=COS(@B1:B10)";
		fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A2"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COS(B1:B10)", "COS(@B1:B10) -> COS(B1:B10)");
		assembledVal = ws.getRange2("A2").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=TAN(@C1:C5)";
		fillRange = ws.getRange2("A3");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A3").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A3"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TAN(C1:C5)", "TAN(@C1:C5) -> TAN(C1:C5)");
		assembledVal = ws.getRange2("A3").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=SQRT(@D1:D10)";
		fillRange = ws.getRange2("A4");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A4").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A4"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SQRT(D1:D10)", "SQRT(@D1:D10) -> SQRT(D1:D10)");
		assembledVal = ws.getRange2("A4").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ABS(@E1:E5)";
		fillRange = ws.getRange2("A5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A5"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ABS(E1:E5)", "ABS(@E1:E5) -> ABS(E1:E5)");
		assembledVal = ws.getRange2("A5").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=EXP(@F1:F10)";
		fillRange = ws.getRange2("A6");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A6").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A6"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EXP(F1:F10)", "EXP(@F1:F10) -> EXP(F1:F10)");
		assembledVal = ws.getRange2("A6").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=LN(@G1:G5)";
		fillRange = ws.getRange2("A7");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A7").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A7"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LN(G1:G5)", "LN(@G1:G5) -> LN(G1:G5)");
		assembledVal = ws.getRange2("A7").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=LOG(@H1:H10)";
		fillRange = ws.getRange2("A8");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A8").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A8"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LOG(H1:H10)", "LOG(@H1:H10) -> LOG(H1:H10)");
		assembledVal = ws.getRange2("A8").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=LOG10(@I1:I5)";
		fillRange = ws.getRange2("A9");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A9").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A9"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LOG10(I1:I5)", "LOG10(@I1:I5) -> LOG10(I1:I5)");
		assembledVal = ws.getRange2("A9").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ROUND(@J1:J10,2)";
		fillRange = ws.getRange2("A10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A10"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUND(J1:J10,2)", "ROUND(@J1:J10,2) -> ROUND(J1:J10,2)");
		assembledVal = ws.getRange2("A10").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ROUNDUP(@K1:K5,1)";
		fillRange = ws.getRange2("A11");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A11").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A11"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUNDUP(K1:K5,1)", "ROUNDUP(@K1:K5,1) -> ROUNDUP(K1:K5,1)");
		assembledVal = ws.getRange2("A11").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ROUNDDOWN(@L1:L10,0)";
		fillRange = ws.getRange2("A12");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A12").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A12"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUNDDOWN(L1:L10,0)", "ROUNDDOWN(@L1:L10,0) -> ROUNDDOWN(L1:L10,0)");
		assembledVal = ws.getRange2("A12").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=INT(@M1:M5)";
		fillRange = ws.getRange2("A13");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A13").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A13"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "INT(M1:M5)", "INT(@M1:M5) -> INT(M1:M5)");
		assembledVal = ws.getRange2("A13").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=TRUNC(@N1:N10)";
		fillRange = ws.getRange2("A14");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A14").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A14"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TRUNC(N1:N10)", "TRUNC(@N1:N10) -> TRUNC(N1:N10)");
		assembledVal = ws.getRange2("A14").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=CEILING(@O1:O5,1)";
		fillRange = ws.getRange2("A15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A15"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CEILING(O1:O5,1)", "CEILING(@O1:O5,1) -> CEILING(O1:O5,1)");
		assembledVal = ws.getRange2("A15").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=FLOOR(@P1:P10,1)";
		fillRange = ws.getRange2("A16");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A16").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A16"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FLOOR(P1:P10,1)", "FLOOR(@P1:P10,1) -> FLOOR(P1:P10,1)");
		assembledVal = ws.getRange2("A16").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=SIGN(@Q1:Q5)";
		fillRange = ws.getRange2("A17");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A17").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A17"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIGN(Q1:Q5)", "SIGN(@Q1:Q5) -> SIGN(Q1:Q5)");
		assembledVal = ws.getRange2("A17").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=FACT(@R1:R10)";
		fillRange = ws.getRange2("A18");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A18").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A18"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FACT(R1:R10)", "FACT(@R1:R10) -> FACT(R1:R10)");
		assembledVal = ws.getRange2("A18").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=POWER(@S1:S5,2)";
		fillRange = ws.getRange2("A19");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A19").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A19"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "POWER(S1:S5,2)", "POWER(@S1:S5,2) -> POWER(S1:S5,2)");
		assembledVal = ws.getRange2("A19").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=RADIANS(@T1:T10)";
		fillRange = ws.getRange2("A20");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A20").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A20"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "RADIANS(T1:T10)", "RADIANS(@T1:T10) -> RADIANS(T1:T10)");
		assembledVal = ws.getRange2("A20").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=DEGREES(@U1:U5)";
		fillRange = ws.getRange2("A21");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A21").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A21"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DEGREES(U1:U5)", "DEGREES(@U1:U5) -> DEGREES(U1:U5)");
		assembledVal = ws.getRange2("A21").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ASIN(@V1:V10)";
		fillRange = ws.getRange2("A22");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A22").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A22"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ASIN(V1:V10)", "ASIN(@V1:V10) -> ASIN(V1:V10)");
		assembledVal = ws.getRange2("A22").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ACOS(@W1:W5)";
		fillRange = ws.getRange2("A23");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A23").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A23"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ACOS(W1:W5)", "ACOS(@W1:W5) -> ACOS(W1:W5)");
		assembledVal = ws.getRange2("A23").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ATAN(@X1:X10)";
		fillRange = ws.getRange2("A24");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A24").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A24"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATAN(X1:X10)", "ATAN(@X1:X10) -> ATAN(X1:X10)");
		assembledVal = ws.getRange2("A24").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=SINH(@Y1:Y5)";
		fillRange = ws.getRange2("A25");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A25").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A25"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SINH(Y1:Y5)", "SINH(@Y1:Y5) -> SINH(Y1:Y5)");
		assembledVal = ws.getRange2("A25").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=COSH(@Z1:Z10)";
		fillRange = ws.getRange2("A26");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A26").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A26"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COSH(Z1:Z10)", "COSH(@Z1:Z10) -> COSH(Z1:Z10)");
		assembledVal = ws.getRange2("A26").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=TANH(@A2:A10)";
		fillRange = ws.getRange2("A27");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A27").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A27"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TANH(A2:A10)", "TANH(@A2:A10) -> TANH(A2:A10)");
		assembledVal = ws.getRange2("A27").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=LEFT(@B2:B10,3)";
		fillRange = ws.getRange2("A28");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A28").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A28"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LEFT(B2:B10,3)", "LEFT(@B2:B10,3) -> LEFT(B2:B10,3)");
		assembledVal = ws.getRange2("A28").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=RIGHT(@C2:C10,2)";
		fillRange = ws.getRange2("A29");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A29").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A29"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "RIGHT(C2:C10,2)", "RIGHT(@C2:C10,2) -> RIGHT(C2:C10,2)");
		assembledVal = ws.getRange2("A29").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=MID(@D2:D10,2,3)";
		fillRange = ws.getRange2("A30");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A30").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A30"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MID(D2:D10,2,3)", "MID(@D2:D10,2,3) -> MID(D2:D10,2,3)");
		assembledVal = ws.getRange2("A30").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=LEN(@E2:E10)";
		fillRange = ws.getRange2("A31");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A31").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A31"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LEN(E2:E10)", "LEN(@E2:E10) -> LEN(E2:E10)");
		assembledVal = ws.getRange2("A31").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=UPPER(@F2:F10)";
		fillRange = ws.getRange2("A32");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A32").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A32"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "UPPER(F2:F10)", "UPPER(@F2:F10) -> UPPER(F2:F10)");
		assembledVal = ws.getRange2("A32").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=LOWER(@G2:G10)";
		fillRange = ws.getRange2("A33");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A33").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A33"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LOWER(G2:G10)", "LOWER(@G2:G10) -> LOWER(G2:G10)");
		assembledVal = ws.getRange2("A33").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=PROPER(@H2:H10)";
		fillRange = ws.getRange2("A34");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A34").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A34"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "PROPER(H2:H10)", "PROPER(@H2:H10) -> PROPER(H2:H10)");
		assembledVal = ws.getRange2("A34").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=TRIM(@I2:I10)";
		fillRange = ws.getRange2("A35");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A35").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A35"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TRIM(I2:I10)", "TRIM(@I2:I10) -> TRIM(I2:I10)");
		assembledVal = ws.getRange2("A35").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=CLEAN(@J2:J10)";
		fillRange = ws.getRange2("A36");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A36").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A36"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CLEAN(J2:J10)", "CLEAN(@J2:J10) -> CLEAN(J2:J10)");
		assembledVal = ws.getRange2("A36").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=TEXT(@K2:K10,\"0.00\")";
		fillRange = ws.getRange2("A37");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A37").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A37"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TEXT(K2:K10,\"0.00\")", "TEXT(@K2:K10,\"0.00\") -> TEXT(K2:K10,\"0.00\")");
		assembledVal = ws.getRange2("A37").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=VALUE(@L2:L10)";
		fillRange = ws.getRange2("A38");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A38").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A38"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "VALUE(L2:L10)", "VALUE(@L2:L10) -> VALUE(L2:L10)");
		assembledVal = ws.getRange2("A38").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=FIND(\"a\",@M2:M10)";
		fillRange = ws.getRange2("A39");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A39").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A39"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FIND(\"a\",M2:M10)", "FIND(\"a\",@M2:M10) -> FIND(\"a\",M2:M10)");
		assembledVal = ws.getRange2("A39").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=SEARCH(\"test\",@N2:N10)";
		fillRange = ws.getRange2("A40");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A40").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A40"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SEARCH(\"test\",N2:N10)", "SEARCH(\"test\",@N2:N10) -> SEARCH(\"test\",N2:N10)");
		assembledVal = ws.getRange2("A40").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=REPLACE(@O2:O10,1,2,\"XX\")";
		fillRange = ws.getRange2("A41");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A41").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A41"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "REPLACE(O2:O10,1,2,\"XX\")", "REPLACE(@O2:O10,1,2,\"XX\") -> REPLACE(O2:O10,1,2,\"XX\")");
		assembledVal = ws.getRange2("A41").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=SUBSTITUTE(@P2:P10,\"old\",\"new\")";
		fillRange = ws.getRange2("A42");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A42").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A42"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SUBSTITUTE(P2:P10,\"old\",\"new\")", "SUBSTITUTE(@P2:P10,\"old\",\"new\") -> SUBSTITUTE(P2:P10,\"old\",\"new\")");
		assembledVal = ws.getRange2("A42").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=REPT(@Q2:Q10,3)";
		fillRange = ws.getRange2("A43");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A43").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A43"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "REPT(Q2:Q10,3)", "REPT(@Q2:Q10,3) -> REPT(Q2:Q10,3)");
		assembledVal = ws.getRange2("A43").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=CONCATENATE(@R2:R10,\"-\",@S2:S10)";
		fillRange = ws.getRange2("A44");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A44").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A44"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CONCATENATE(R2:R10,\"-\",S2:S10)", "CONCATENATE(@R2:R10,\"-\",@S2:S10) -> CONCATENATE(R2:R10,\"-\",S2:S10)");
		assembledVal = ws.getRange2("A44").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=DATE(@T2:T10,@U2:U10,@V2:V10)";
		fillRange = ws.getRange2("A45");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A45").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A45"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DATE(T2:T10,U2:U10,V2:V10)", "DATE(@T2:T10,@U2:U10,@V2:V10) -> DATE(T2:T10,U2:U10,V2:V10)");
		assembledVal = ws.getRange2("A45").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=TIME(@W2:W10,@X2:X10,@Y2:Y10)";
		fillRange = ws.getRange2("A46");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A46").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A46"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TIME(W2:W10,X2:X10,Y2:Y10)", "TIME(@W2:W10,@X2:X10,@Y2:Y10) -> TIME(W2:W10,X2:X10,Y2:Y10)");
		assembledVal = ws.getRange2("A46").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=YEAR(@Z2:Z10)";
		fillRange = ws.getRange2("A47");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A47").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A47"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "YEAR(Z2:Z10)", "YEAR(@Z2:Z10) -> YEAR(Z2:Z10)");
		assembledVal = ws.getRange2("A47").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=MONTH(@A3:A12)";
		fillRange = ws.getRange2("A48");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A48").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A48"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MONTH(A3:A12)", "MONTH(@A3:A12) -> MONTH(A3:A12)");
		assembledVal = ws.getRange2("A48").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=DAY(@B3:B12)";
		fillRange = ws.getRange2("A49");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A49").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A49"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DAY(B3:B12)", "DAY(@B3:B12) -> DAY(B3:B12)");
		assembledVal = ws.getRange2("A49").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=HOUR(@C3:C12)";
		fillRange = ws.getRange2("A50");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A50").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A50"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "HOUR(C3:C12)", "HOUR(@C3:C12) -> HOUR(C3:C12)");
		assembledVal = ws.getRange2("A50").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=MINUTE(@D3:D12)";
		fillRange = ws.getRange2("A51");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A51").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A51"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MINUTE(D3:D12)", "MINUTE(@D3:D12) -> MINUTE(D3:D12)");
		assembledVal = ws.getRange2("A51").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=SECOND(@E3:E12)";
		fillRange = ws.getRange2("A52");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A52").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A52"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SECOND(E3:E12)", "SECOND(@E3:E12) -> SECOND(E3:E12)");
		assembledVal = ws.getRange2("A52").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=WEEKDAY(@F3:F12)";
		fillRange = ws.getRange2("A53");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A53").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A53"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "WEEKDAY(F3:F12)", "WEEKDAY(@F3:F12) -> WEEKDAY(F3:F12)");
		assembledVal = ws.getRange2("A53").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=NOT(@G3:G12)";
		fillRange = ws.getRange2("A54");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A54").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A54"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "NOT(G3:G12)", "NOT(@G3:G12) -> NOT(G3:G12)");
		assembledVal = ws.getRange2("A54").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ISBLANK(@H3:H12)";
		fillRange = ws.getRange2("A55");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A55").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A55"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISBLANK(H3:H12)", "ISBLANK(@H3:H12) -> ISBLANK(H3:H12)");
		assembledVal = ws.getRange2("A55").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ISERROR(@I3:I12)";
		fillRange = ws.getRange2("A56");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A56").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A56"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISERROR(I3:I12)", "ISERROR(@I3:I12) -> ISERROR(I3:I12)");
		assembledVal = ws.getRange2("A56").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ISNA(@J3:J12)";
		fillRange = ws.getRange2("A57");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A57").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A57"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISNA(J3:J12)", "ISNA(@J3:J12) -> ISNA(J3:J12)");
		assembledVal = ws.getRange2("A57").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ISNUMBER(@K3:K12)";
		fillRange = ws.getRange2("A58");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A58").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A58"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISNUMBER(K3:K12)", "ISNUMBER(@K3:K12) -> ISNUMBER(K3:K12)");
		assembledVal = ws.getRange2("A58").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ISTEXT(@L3:L12)";
		fillRange = ws.getRange2("A59");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A59").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A59"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISTEXT(L3:L12)", "ISTEXT(@L3:L12) -> ISTEXT(L3:L12)");
		assembledVal = ws.getRange2("A59").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ISNONTEXT(@M3:M12)";
		fillRange = ws.getRange2("A60");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A60").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A60"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISNONTEXT(M3:M12)", "ISNONTEXT(@M3:M12) -> ISNONTEXT(M3:M12)");
		assembledVal = ws.getRange2("A60").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ISLOGICAL(@N3:N12)";
		fillRange = ws.getRange2("A61");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A61").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A61"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISLOGICAL(N3:N12)", "ISLOGICAL(@N3:N12) -> ISLOGICAL(N3:N12)");
		assembledVal = ws.getRange2("A61").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=MOD(@O3:O12,3)";
		fillRange = ws.getRange2("A62");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A62").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A62"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "MOD(O3:O12,3)", "MOD(@O3:O12,3) -> MOD(O3:O12,3)");
		assembledVal = ws.getRange2("A62").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=ATAN2(@P3:P12,@Q3:Q12)";
		fillRange = ws.getRange2("A63");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A63").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A63"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATAN2(P3:P12,Q3:Q12)", "ATAN2(@P3:P12,@Q3:Q12) -> ATAN2(P3:P12,Q3:Q12)");
		assembledVal = ws.getRange2("A63").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=COMBIN(@R3:R12,@S3:S12)";
		fillRange = ws.getRange2("A64");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A64").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A64"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COMBIN(R3:R12,S3:S12)", "COMBIN(@R3:R12,@S3:S12) -> COMBIN(R3:R12,S3:S12)");
		assembledVal = ws.getRange2("A64").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=PERMUT(@T3:T12,@U3:U12)";
		fillRange = ws.getRange2("A65");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A65").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A65"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "PERMUT(T3:T12,U3:U12)", "PERMUT(@T3:T12,@U3:U12) -> PERMUT(T3:T12,U3:U12)");
		assembledVal = ws.getRange2("A65").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=EXACT(@V3:V12,@W3:W12)";
		fillRange = ws.getRange2("A66");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A66").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A66"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EXACT(V3:V12,W3:W12)", "EXACT(@V3:V12,@W3:W12) -> EXACT(V3:V12,W3:W12)");
		assembledVal = ws.getRange2("A66").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		
		formula = "=CODE(@X3:X12)";
		fillRange = ws.getRange2("A67");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A67").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A67"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CODE(X3:X12)", "CODE(@X3:X12) -> CODE(X3:X12)");
		assembledVal = ws.getRange2("A67").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=CODE(@X3:X12)";
		fillRange = ws.getRange2("A67");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A67").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A67"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CODE(X3:X12)", "CODE(@X3:X12) -> CODE(X3:X12)");
		assembledVal = ws.getRange2("A67").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=CHAR(@Y3:Y12)";
		fillRange = ws.getRange2("A68");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A68").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A68"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CHAR(Y3:Y12)", "CHAR(@Y3:Y12) -> CHAR(Y3:Y12)");
		assembledVal = ws.getRange2("A68").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ASINH(@Z3:Z12)";
		fillRange = ws.getRange2("A69");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A69").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A69"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ASINH(Z3:Z12)", "ASINH(@Z3:Z12) -> ASINH(Z3:Z12)");
		assembledVal = ws.getRange2("A69").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ACOSH(@A4:A13)";
		fillRange = ws.getRange2("A70");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A70").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A70"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ACOSH(A4:A13)", "ACOSH(@A4:A13) -> ACOSH(A4:A13)");
		assembledVal = ws.getRange2("A70").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ATANH(@B4:B13)";
		fillRange = ws.getRange2("A71");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A71").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A71"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATANH(B4:B13)", "ATANH(@B4:B13) -> ATANH(B4:B13)");
		assembledVal = ws.getRange2("A71").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=EVEN(@C4:C13)";
		fillRange = ws.getRange2("A72");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A72").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A72"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EVEN(C4:C13)", "EVEN(@C4:C13) -> EVEN(C4:C13)");
		assembledVal = ws.getRange2("A72").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ODD(@D4:D13)";
		fillRange = ws.getRange2("A73");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A73").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A73"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ODD(D4:D13)", "ODD(@D4:D13) -> ODD(D4:D13)");
		assembledVal = ws.getRange2("A73").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=TYPE(@I4:I13)";
		fillRange = ws.getRange2("A78");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A78").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A78"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TYPE(I4:I13)", "TYPE(@I4:I13) -> TYPE(I4:I13)");
		assembledVal = ws.getRange2("A78").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISERR(@K4:K13)";
		fillRange = ws.getRange2("A80");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A80").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A80"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISERR(K4:K13)", "ISERR(@K4:K13) -> ISERR(K4:K13)");
		assembledVal = ws.getRange2("A80").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=DATEVALUE(@L4:L13)";
		fillRange = ws.getRange2("A81");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A81").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A81"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DATEVALUE(L4:L13)", "DATEVALUE(@L4:L13) -> DATEVALUE(L4:L13)");
		assembledVal = ws.getRange2("A81").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=TIMEVALUE(@M4:M13)";
		fillRange = ws.getRange2("A82");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A82").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A82"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TIMEVALUE(M4:M13)", "TIMEVALUE(@M4:M13) -> TIMEVALUE(M4:M13)");
		assembledVal = ws.getRange2("A82").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=DOLLAR(@V4:V13,2)";
		fillRange = ws.getRange2("A87");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A87").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A87"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DOLLAR(V4:V13,2)", "DOLLAR(@V4:V13,2) -> DOLLAR(V4:V13,2)");
		assembledVal = ws.getRange2("A87").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=FIXED(@W4:W13,2)";
		fillRange = ws.getRange2("A88");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A88").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A88"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FIXED(W4:W13,2)", "FIXED(@W4:W13,2) -> FIXED(W4:W13,2)");
		assembledVal = ws.getRange2("A88").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ROMAN(@X4:X13)";
		fillRange = ws.getRange2("A89");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A89").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A89"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROMAN(X4:X13)", "ROMAN(@X4:X13) -> ROMAN(X4:X13)");
		assembledVal = ws.getRange2("A89").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ARABIC(@Y4:Y13)";
		fillRange = ws.getRange2("A90");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A90").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A90"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.ARABIC(Y4:Y13)", "ARABIC(@Y4:Y13) -> _xlfn.ARABIC(Y4:Y13)");
		assembledVal = ws.getRange2("A90").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=BASE(@Z4:Z13,16)";
		fillRange = ws.getRange2("A91");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A91").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A91"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.BASE(Z4:Z13,16)", "BASE(@Z4:Z13,16) -> _xlfn.BASE(Z4:Z13,16)");
		assembledVal = ws.getRange2("A91").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=DECIMAL(@A5:A14,16)";
		fillRange = ws.getRange2("A92");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A92").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A92"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.DECIMAL(A5:A14,16)", "DECIMAL(@A5:A14,16) -> _xlfn.DECIMAL(A5:A14,16)");
		assembledVal = ws.getRange2("A92").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=NUMBERVALUE(@B5:B14)";
		fillRange = ws.getRange2("A93");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A93").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A93"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.NUMBERVALUE(B5:B14)", "NUMBERVALUE(@B5:B14) -> _xlfn.NUMBERVALUE(B5:B14)");
		assembledVal = ws.getRange2("A93").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=UNICHAR(@C5:C14)";
		fillRange = ws.getRange2("A94");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A94").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A94"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.UNICHAR(C5:C14)", "UNICHAR(@C5:C14) -> _xlfn.UNICHAR(C5:C14)");
		assembledVal = ws.getRange2("A94").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=UNICODE(@D5:D14)";
		fillRange = ws.getRange2("A95");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A95").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A95"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.UNICODE(D5:D14)", "UNICODE(@D5:D14) -> _xlfn.UNICODE(D5:D14)");
		assembledVal = ws.getRange2("A95").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ENCODEURL(@E5:E14)";
		fillRange = ws.getRange2("A96");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A96").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A96"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ENCODEURL(E5:E14)", "ENCODEURL(@E5:E14) -> ENCODEURL(E5:E14)");
		assembledVal = ws.getRange2("A96").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SEC(@F5:F14)";
		fillRange = ws.getRange2("A97");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A97").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A97"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SEC(F5:F14)", "SEC(@F5:F14) -> _xlfn.SEC(F5:F14)");
		assembledVal = ws.getRange2("A97").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=CSC(@G5:G14)";
		fillRange = ws.getRange2("A98");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A98").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A98"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.CSC(G5:G14)", "CSC(@G5:G14) -> _xlfn.CSC(G5:G14)");
		assembledVal = ws.getRange2("A98").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=COT(@H5:H14)";
		fillRange = ws.getRange2("A99");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A99").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A99"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.COT(H5:H14)", "COT(@H5:H14) -> _xlfn.COT(H5:H14)");
		assembledVal = ws.getRange2("A99").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=SECH(@I5:I14)";
		fillRange = ws.getRange2("A100");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A100").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A100"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.SECH(I5:I14)", "SECH(@I5:I14) -> _xlfn.SECH(I5:I14)");
		assembledVal = ws.getRange2("A100").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=CSCH(@J5:J14)";
		fillRange = ws.getRange2("A101");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A101").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A101"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.CSCH(J5:J14)", "CSCH(@J5:J14) -> _xlfn.CSCH(J5:J14)");
		assembledVal = ws.getRange2("A101").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=COTH(@K5:K14)";
		fillRange = ws.getRange2("A102");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A102").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A102"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.COTH(K5:K14)", "COTH(@K5:K14) -> _xlfn.COTH(K5:K14)");
		assembledVal = ws.getRange2("A102").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ACOT(@L5:L14)";
		fillRange = ws.getRange2("A103");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A103").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A103"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.ACOT(L5:L14)", "ACOT(@L5:L14) -> _xlfn.ACOT(L5:L14)");
		assembledVal = ws.getRange2("A103").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ACOTH(@M5:M14)";
		fillRange = ws.getRange2("A104");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A104").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A104"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.ACOTH(M5:M14)", "ACOTH(@M5:M14) -> _xlfn.ACOTH(M5:M14)");
		assembledVal = ws.getRange2("A104").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISOWEEKNUM(@N5:N14)";
		fillRange = ws.getRange2("A105");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A105").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A105"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.ISOWEEKNUM(N5:N14)", "ISOWEEKNUM(@N5:N14) -> _xlfn.ISOWEEKNUM(N5:N14)");
		assembledVal = ws.getRange2("A105").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=DAYS(@P5:P14,@Q5:Q14)";
		fillRange = ws.getRange2("A107");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A107").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A107"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.DAYS(P5:P14,Q5:Q14)", "DAYS(@P5:P14,@Q5:Q14) -> _xlfn.DAYS(P5:P14,Q5:Q14)");
		assembledVal = ws.getRange2("A107").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=DAYS360(@R5:R14,@S5:S14)";
		fillRange = ws.getRange2("A108");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A108").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A108"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DAYS360(R5:R14,S5:S14)", "DAYS360(@R5:R14,@S5:S14) -> DAYS360(R5:R14,S5:S14)");
		assembledVal = ws.getRange2("A108").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=DATEDIF(@T5:T14,@U5:U14,\"D\")";
		fillRange = ws.getRange2("A109");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A109").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A109"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DATEDIF(T5:T14,U5:U14,\"D\")", "DATEDIF(@T5:T14,@U5:U14,\"D\") -> DATEDIF(T5:T14,U5:U14,\"D\")");
		assembledVal = ws.getRange2("A109").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);


		formula = "=ERROR.TYPE(@I6:I15)";
		fillRange = ws.getRange2("A118");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A118").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A118"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ERROR.TYPE(I6:I15)", "ERROR.TYPE(@I6:I15) -> ERROR.TYPE(I6:I15)");
		assembledVal = ws.getRange2("A118").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);


		formula = "=COMBINA(@K6:K15,@L6:L15)";
		fillRange = ws.getRange2("A120");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A120").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A120"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "_xlfn.COMBINA(K6:K15,L6:L15)", "COMBINA(@K6:K15,@L6:L15) -> _xlfn.COMBINA(K6:K15,L6:L15)");
		assembledVal = ws.getRange2("A120").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		ws.getRange2("A1:Z150").cleanAll();
	});

	QUnit.test("Test: \"Dynamic array test\"", function (assert) {
		let bboxParent, cellWithFormula, formulaInfo, resultRow, resultCol, applyByArray, array, oParser;

		// wb.dependencyFormulas.unlockRecal();

		ws.getRange2("A1:Z10").cleanAll();
		ws.getRange2("A1").setValue("1");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("A3").setValue("3");
		ws.getRange2("B1").setValue("4");
		ws.getRange2("B2").setValue("str");
		ws.getRange2("B3").setValue("6");
		ws.getRange2("C1").setValue("1");
		ws.getRange2("C2").setValue();
		ws.getRange2("C3").setValue("1");

		// let parent = AscCommonExcel.g_oRangeCache.getAscRange("D1");
		bboxParent = ws.getRange2("D1").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bboxParent.r1, bboxParent.c1);

		ws.getRange2("C3").setValue("=SIN(A1:A3)", null, null, bboxParent);

		// TODO: review tests with ranges after adding dynamic arrays and add findRefByOutStack formula to use in tests
		oParser = new parserFormula('A1:A3', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A1:A3');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A1:A3 array formula');
		assert.strictEqual(resultRow, false, 'Rows in =A1:A3');
		assert.strictEqual(resultCol, false, 'Cols in =A1:A3');


		oParser = new parserFormula('{1;2;3}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{1;2;3}');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={1;2;3} array formula');
		assert.strictEqual(resultRow, 3, 'Rows in ={1;2;3}');
		assert.strictEqual(resultCol, 1, 'Cols in ={1;2;3}');

		oParser = new parserFormula('A1:C1', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A1:C1');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A1:C1 array formula');
		assert.strictEqual(resultRow, false, 'Rows in =A1:C1');
		assert.strictEqual(resultCol, false, 'Cols in =A1:C1');

		oParser = new parserFormula('{1,2,3}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{1,2,3}');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={1,2,3} array formula');
		assert.strictEqual(resultRow, 1, 'Rows in ={1,2,3}');
		assert.strictEqual(resultCol, 3, 'Cols in ={1,2,3}');

		oParser = new parserFormula('A1:C3', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A1:C3');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A1:C3 array formula');
		assert.strictEqual(resultRow, false, 'Rows in =A1:C3');
		assert.strictEqual(resultCol, false, 'Cols in =A1:C3');

		oParser = new parserFormula('{1,2;3,4}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{1,2;3,4}');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={1,2;3,4} array formula');
		assert.strictEqual(resultRow, 2, 'Rows in ={1,2;3,4}');
		assert.strictEqual(resultCol, 2, 'Cols in ={1,2;3,4}');

		oParser = new parserFormula('SIN(A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1:A3)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN(A1:A3) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SIN(A1:A3)');
		assert.strictEqual(resultCol, 1, 'Cols in =SIN(A1:A3)');

		oParser = new parserFormula('SIN({1;2;3})', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN({1;2;3})');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN({1;2;3}) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SIN({1;2;3})');
		assert.strictEqual(resultCol, 1, 'Cols in =SIN({1;2;3})');

		oParser = new parserFormula('SIN(A1:C1)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1:C1)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN(A1:C1) array formula');
		assert.strictEqual(resultRow, 1, 'Rows in =SIN(A1:C1)');
		assert.strictEqual(resultCol, 3, 'Cols in =SIN(A1:C1)');

		oParser = new parserFormula('SIN({1,2,3})', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN({1,2,3})');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN({1,2,3}) array formula');
		assert.strictEqual(resultRow, 1, 'Rows in =SIN({1,2,3})');
		assert.strictEqual(resultCol, 3, 'Cols in =SIN({1,2,3})');

		oParser = new parserFormula('SIN(A1:C3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1:C3)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN(A1:C3) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SIN(A1:C3)');
		assert.strictEqual(resultCol, 3, 'Cols in =SIN(A1:C3)');

		oParser = new parserFormula('SIN({1,2;3,4})', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN({1,2;3,4})');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN({1,2;3,4}) array formula');
		assert.strictEqual(resultRow, 2, 'Rows in =SIN({1,2;3,4})');
		assert.strictEqual(resultCol, 2, 'Cols in =SIN({1,2;3,4})');

		oParser = new parserFormula('A:A', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A:A');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A:A array formula');
		assert.strictEqual(resultRow, false /*AscCommon.gc_nMaxRow*/, 'Rows in =A:A from D1');
		assert.strictEqual(resultCol, false, 'Cols in =A:A from D1');

		oParser = new parserFormula('A1:XFD1', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'A1:XFD1');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =A1:XFD1 array formula');
		assert.strictEqual(resultRow, false, 'Rows in =A1:XFD1 from D1');
		assert.strictEqual(resultCol, false /*AscCommon.gc_nMaxCol - 3*/, 'Cols in =A1:XFD1 from D1');


		oParser = new parserFormula('SIN(A1)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SIN(A1) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SIN(A1)');
		assert.strictEqual(resultCol, false, 'Cols in =SIN(A1)');


		oParser = new parserFormula('SUM(A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(A1:A3)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SUM(A1:A3) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SUM(A1:A3)');
		assert.strictEqual(resultCol, false, 'Cols in =SUM(A1:A3)');


		oParser = new parserFormula('SUM(A1:A3+A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(A1:A3+A1:A3)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SUM(A1:A3+A1:A3) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SUM(A1:A3+A1:A3)');
		assert.strictEqual(resultCol, false, 'Cols in =SUM(A1:A3+A1:A3)');

		oParser = new parserFormula('SUM(A1:A3+A1:A3)+A1:A3', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(A1:A3+A1:A3)+A1:A3');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SUM(A1:A3+A1:A3)+A1:A3 array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SUM(A1:A3+A1:A3)+A1:A3');
		assert.strictEqual(resultCol, 1, 'Cols in =SUM(A1:A3+A1:A3)+A1:A3');


		oParser = new parserFormula('SUM(SIN(A1:A3)+A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(SIN(A1:A3)+A1:A3)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SUM(SIN(A1:A3)+A1:A3) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SUM(SIN(A1:A3)+A1:A3)');
		assert.strictEqual(resultCol, false, 'Cols in =SUM(SIN(A1:A3)+A1:A3)');


		oParser = new parserFormula('SUM(SIN(SUM(A1:A3)))', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SUM(SIN(SUM(A1:A3)))');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SUM(SIN(SUM(A1:A3))) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SUM(SIN(SUM(A1:A3)))');
		assert.strictEqual(resultCol, false, 'Cols in =SUM(SIN(SUM(A1:A3)))');


		oParser = new parserFormula('SIN(SUM(SIN(A1:A3)))', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(SUM(SIN(A1:A3)))');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, false, 'Is =SIN(SUM(SIN(A1:A3))) array formula');
		assert.strictEqual(resultRow, false, 'Rows in =SIN(SUM(SIN(A1:A3)))');
		assert.strictEqual(resultCol, false, 'Cols in =SIN(SUM(SIN(A1:A3)))');


		oParser = new parserFormula('COS(SIN(A1)*SUM(A1:A3)+A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'COS(SIN(A1)*SUM(A1:A3)+A1:A3)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =COS(SIN(A1)*SUM(A1:A3)+A1:A3) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =COS(SIN(A1)*SUM(A1:A3)+A1:A3)');
		assert.strictEqual(resultCol, 1, 'Cols in =COS(SIN(A1)*SUM(A1:A3)+A1:A3)');


		oParser = new parserFormula('SIN(A1+A1:A3)', cellWithFormula, ws);
		assert.ok(oParser.parse(), 'SIN(A1+A1:A3)');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is =SIN(A1+A1:A3) array formula');
		assert.strictEqual(resultRow, 3, 'Rows in =SIN(A1+A1:A3)');
		assert.strictEqual(resultCol, 1, 'Cols in =SIN(A1+A1:A3)');


		oParser = new parserFormula('{1,2}*{3;4}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{1,2}*{3;4}');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={1,2}*{3;4} array formula');
		assert.strictEqual(resultRow, 2, 'Rows in ={1,2}*{3;4}');
		assert.strictEqual(resultCol, 2, 'Cols in ={1,2}*{3;4}');

		oParser = new parserFormula('{2}*{2}', cellWithFormula, ws);
		assert.ok(oParser.parse(), '{2}*{2}');
		formulaInfo = ws.dynamicArrayManager.getRefDynamicInfo(oParser);
		resultRow = formulaInfo && formulaInfo.dynamicRange.getHeight();
		resultCol = formulaInfo && formulaInfo.dynamicRange.getWidth();
		applyByArray = formulaInfo && formulaInfo.applyByArray;
		assert.strictEqual(applyByArray, true, 'Is ={2}*{2} array formula');
		assert.strictEqual(resultRow, 1, 'Rows in ={1,2}*{3;4}');
		assert.strictEqual(resultCol, 1, 'Cols in ={1,2}*{3;4}');

		// #N/A check
		ws.getRange2("A100:Z110").cleanAll();

		bboxParent = ws.getRange2("D100").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bboxParent.r1, bboxParent.c1);
		oParser = new parserFormula('A100:B101', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("D100:E104").bbox);
		assert.ok(oParser.parse(), 'A100:B101');
		array = oParser.calculate();
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1).getValue(), "", "Result of =A100:B101 [0,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [0,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [0,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [0,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1).getValue(), "", "Result of =A100:B101 [1,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [1,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [1,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [1,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1).getValue(), "#N/A", "Result of =A100:B101 [2,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 1).getValue(), "#N/A", "Result of =A100:B101 [2,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [2,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [2,3]");


		ws.getRange2("A100").setValue("1");

		bboxParent = ws.getRange2("I100").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bboxParent.r1, bboxParent.c1);
		oParser = new parserFormula('A100:B101', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("I100:J104").bbox);
		assert.ok(oParser.parse(), 'A100:B101');
		array = oParser.calculate();
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1).getValue(), 1, "Result of =A100:B101 [0,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [0,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [0,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [0,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1).getValue(), "", "Result of =A100:B101 [1,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [1,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [1,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [1,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1).getValue(), "#N/A", "Result of =A100:B101 [2,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 1).getValue(), "#N/A", "Result of =A100:B101 [2,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [2,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [2,3]");


		ws.getRange2("B101").setValue("#N/A");

		bboxParent = ws.getRange2("M100").bbox;
		cellWithFormula = new window['AscCommonExcel'].CCellWithFormula(ws, bboxParent.r1, bboxParent.c1);
		oParser = new parserFormula('A100:B101', cellWithFormula, ws);
		oParser.setArrayFormulaRef(ws.getRange2("M100:O104").bbox);
		assert.ok(oParser.parse(), 'A100:B101');
		array = oParser.calculate();
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1).getValue(), 1, "Result of =A100:B101 [0,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 1).getValue(), "", "Result of =A100:B101 [0,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [0,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [0,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1).getValue(), "", "Result of =A100:B101 [1,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 1).getValue(), "#N/A", "Result of =A100:B101 [1,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [1,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 1, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [1,3]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1).getValue(), "#N/A", "Result of =A100:B101 [2,0]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 1).getValue(), "#N/A", "Result of =A100:B101 [2,1]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 2).getValue(), "#N/A", "Result of =A100:B101 [2,2]");
		assert.strictEqual(oParser.simplifyRefType(array, ws, bboxParent.r1 + 2, bboxParent.c1 + 3).getValue(), "#N/A", "Result of =A100:B101 [2,3]");

		ws.getRange2("A1:Z150").cleanAll();
	});

	QUnit.test("Test: \"Check expand dynamic array test\"", function (assert) {
		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		let formula = "=SIN(A1:B2)";
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A1"));
		let dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIN(A1:B2)", "formula result -> SIN(A1:B2)");
		assert.strictEqual(dynamicRef.getHeight(), 2, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 2, "width dynamic array: " + formula);

		// Test 2: COS with range
		ws.getRange2("E1").setValue("5");
		ws.getRange2("E2").setValue("10");
		ws.getRange2("E3").setValue("15");
		ws.getRange2("F1").setValue("20");
		ws.getRange2("F2").setValue("25");
		ws.getRange2("F3").setValue("30");

		formula = "=COS(E1:F3)";
		fillRange = ws.getRange2("H1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("H1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("H1"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COS(E1:F3)", "formula result -> COS(E1:F3)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 2, "width dynamic array: " + formula);

		// Test 3: ABS with range
		ws.getRange2("A5").setValue("-5");
		ws.getRange2("A6").setValue("-10");
		ws.getRange2("B5").setValue("-15");
		ws.getRange2("B6").setValue("-20");

		formula = "=ABS(A5:B6)";
		fillRange = ws.getRange2("D5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D5"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ABS(A5:B6)", "formula result -> ABS(A5:B6)");
		assert.strictEqual(dynamicRef.getHeight(), 2, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 2, "width dynamic array: " + formula);

		// Test 4: SQRT with range
		ws.getRange2("A8").setValue("4");
		ws.getRange2("A9").setValue("9");
		ws.getRange2("A10").setValue("16");

		formula = "=SQRT(A8:A10)";
		fillRange = ws.getRange2("C8");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C8").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("C8"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SQRT(A8:A10)", "formula result -> SQRT(A8:A10)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 5: ROUND with range
		ws.getRange2("E5").setValue("3.14159");
		ws.getRange2("E6").setValue("2.71828");
		ws.getRange2("E7").setValue("1.41421");

		formula = "=ROUND(E5:E7,2)";
		fillRange = ws.getRange2("G5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("G5"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUND(E5:E7,2)", "formula result -> ROUND(E5:E7,2)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 6: POWER with range
		ws.getRange2("A12").setValue("2");
		ws.getRange2("A13").setValue("3");
		ws.getRange2("B12").setValue("3");
		ws.getRange2("B13").setValue("2");

		formula = "=POWER(A12:A13,B12:B13)";
		fillRange = ws.getRange2("D12");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D12").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D12"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "POWER(A12:A13,B12:B13)", "formula result -> POWER(A12:A13,B12:B13)");
		assert.strictEqual(dynamicRef.getHeight(), 2, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 7: UPPER with range
		ws.getRange2("A15").setValue("hello");
		ws.getRange2("A16").setValue("world");
		ws.getRange2("A17").setValue("test");

		formula = "=UPPER(A15:A17)";
		fillRange = ws.getRange2("C15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("C15"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "UPPER(A15:A17)", "formula result -> UPPER(A15:A17)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 8: LEN with range
		ws.getRange2("E10").setValue("Hello");
		ws.getRange2("E11").setValue("World");
		ws.getRange2("F10").setValue("Test");
		ws.getRange2("F11").setValue("Array");

		formula = "=LEN(E10:F11)";
		fillRange = ws.getRange2("H10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("H10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("H10"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LEN(E10:F11)", "formula result -> LEN(E10:F11)");
		assert.strictEqual(dynamicRef.getHeight(), 2, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 2, "width dynamic array: " + formula);

		// Test 9: INT with range
		ws.getRange2("A20").setValue("5.7");
		ws.getRange2("A21").setValue("8.3");
		ws.getRange2("A22").setValue("12.9");

		formula = "=INT(A20:A22)";
		fillRange = ws.getRange2("C20");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C20").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("C20"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "INT(A20:A22)", "formula result -> INT(A20:A22)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 10: EXP with range
		ws.getRange2("E15").setValue("1");
		ws.getRange2("E16").setValue("2");
		ws.getRange2("F15").setValue("3");
		ws.getRange2("F16").setValue("4");

		formula = "=EXP(E15:F16)";
		fillRange = ws.getRange2("H15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("H15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("H15"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EXP(E15:F16)", "formula result -> EXP(E15:F16)");
		assert.strictEqual(dynamicRef.getHeight(), 2, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 2, "width dynamic array: " + formula);

		// Test 11: CHOOSE with multiple arguments
		ws.getRange2("A25").setValue("1");
		ws.getRange2("A26").setValue("2");
		ws.getRange2("A27").setValue("3");
		ws.getRange2("B25").setValue("10");
		ws.getRange2("B26").setValue("20");
		ws.getRange2("B27").setValue("30");
		ws.getRange2("C25").setValue("100");
		ws.getRange2("C26").setValue("200");
		ws.getRange2("C27").setValue("300");
		ws.getRange2("D25").setValue("1000");
		ws.getRange2("D26").setValue("2000");
		ws.getRange2("D27").setValue("3000");

		formula = "=CHOOSE(A25:A27,B25:B27,C25:C27,D25:D27)";
		fillRange = ws.getRange2("F25");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("F25").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("F25"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "CHOOSE(A25:A27,B25:B27,C25:C27,D25:D27)", "formula result -> CHOOSE(A25:A27,B25:B27,C25:C27,D25:D27)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 12: IF with nested conditions and multiple ranges
		ws.getRange2("A30").setValue("5");
		ws.getRange2("A31").setValue("15");
		ws.getRange2("A32").setValue("25");
		ws.getRange2("B30").setValue("100");
		ws.getRange2("B31").setValue("200");
		ws.getRange2("B32").setValue("300");
		ws.getRange2("C30").setValue("50");
		ws.getRange2("C31").setValue("150");
		ws.getRange2("C32").setValue("250");

		formula = "=IF(A30:A32>10,B30:B32*2,C30:C32/2)";
		fillRange = ws.getRange2("E30");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E30").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E30"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(A30:A32>10,B30:B32*2,C30:C32/2)", "formula result -> IF(A30:A32>10,B30:B32*2,C30:C32/2)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 13: SUM with multiple array operations
		ws.getRange2("A35").setValue("2");
		ws.getRange2("A36").setValue("4");
		ws.getRange2("A37").setValue("6");
		ws.getRange2("B35").setValue("3");
		ws.getRange2("B36").setValue("5");
		ws.getRange2("B37").setValue("7");
		ws.getRange2("C35").setValue("10");
		ws.getRange2("C36").setValue("20");
		ws.getRange2("C37").setValue("30");

		formula = "=(A35:A37+B35:B37)*C35:C37";
		fillRange = ws.getRange2("E35");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E35").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E35"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "(A35:A37+B35:B37)*C35:C37", "formula result -> (A35:A37+B35:B37)*C35:C37");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 14: Multiple mathematical operations
		ws.getRange2("A40").setValue("10");
		ws.getRange2("A41").setValue("20");
		ws.getRange2("A42").setValue("30");
		ws.getRange2("B40").setValue("2");
		ws.getRange2("B41").setValue("3");
		ws.getRange2("B42").setValue("4");
		ws.getRange2("C40").setValue("5");
		ws.getRange2("C41").setValue("6");
		ws.getRange2("C42").setValue("7");
		ws.getRange2("D40").setValue("1");
		ws.getRange2("D41").setValue("2");
		ws.getRange2("D42").setValue("3");

		formula = "=((A40:A42*B40:B42)+(C40:C42-D40:D42))/2";
		fillRange = ws.getRange2("F40");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("F40").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("F40"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "((A40:A42*B40:B42)+(C40:C42-D40:D42))/2", "formula result -> ((A40:A42*B40:B42)+(C40:C42-D40:D42))/2");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 15: Complex nested functions with arrays
		ws.getRange2("A45").setValue("1");
		ws.getRange2("A46").setValue("2");
		ws.getRange2("A47").setValue("3");
		ws.getRange2("B45").setValue("4");
		ws.getRange2("B46").setValue("5");
		ws.getRange2("B47").setValue("6");

		formula = "=ROUND(SQRT(A45:A47^2+B45:B47^2),2)";
		fillRange = ws.getRange2("D45");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D45").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D45"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUND(SQRT(A45:A47^2+B45:B47^2),2)", "formula result -> ROUND(SQRT(A45:A47^2+B45:B47^2),2)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 16: String concatenation with multiple ranges
		ws.getRange2("A50").setValue("First");
		ws.getRange2("A51").setValue("Second");
		ws.getRange2("A52").setValue("Third");
		ws.getRange2("B50").setValue("Name");
		ws.getRange2("B51").setValue("Title");
		ws.getRange2("B52").setValue("Label");
		ws.getRange2("C50").setValue("2024");
		ws.getRange2("C51").setValue("2025");
		ws.getRange2("C52").setValue("2026");

		formula = "=A50:A52&\" \"&B50:B52&\" \"&C50:C52";
		fillRange = ws.getRange2("E50");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E50").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E50"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "A50:A52&\" \"&B50:B52&\" \"&C50:C52", "formula result -> A50:A52&\" \"&B50:B52&\" \"&C50:C52");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 17: Multiple comparison operations
		ws.getRange2("A55").setValue("10");
		ws.getRange2("A56").setValue("20");
		ws.getRange2("A57").setValue("30");
		ws.getRange2("B55").setValue("15");
		ws.getRange2("B56").setValue("15");
		ws.getRange2("B57").setValue("15");

		formula = "=(A55:A57>B55:B57)*100+(A55:A57=B55:B57)*50+(A55:A57<B55:B57)*25";
		fillRange = ws.getRange2("D55");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D55").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D55"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "(A55:A57>B55:B57)*100+(A55:A57=B55:B57)*50+(A55:A57<B55:B57)*25", "formula result -> (A55:A57>B55:B57)*100+(A55:A57=B55:B57)*50+(A55:A57<B55:B57)*25");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 18: Nested IF with multiple conditions
		ws.getRange2("A60").setValue("A");
		ws.getRange2("A61").setValue("B");
		ws.getRange2("A62").setValue("C");
		ws.getRange2("B60").setValue("10");
		ws.getRange2("B61").setValue("20");
		ws.getRange2("B62").setValue("30");

		formula = "=IF(A60:A62=\"A\",B60:B62*1.5,IF(A60:A62=\"B\",B60:B62*2,B60:B62*2.5))";
		fillRange = ws.getRange2("D60");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D60").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D60"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(A60:A62=\"A\",B60:B62*1.5,IF(A60:A62=\"B\",B60:B62*2,B60:B62*2.5))", "formula result -> IF(A60:A62=\"A\",B60:B62*1.5,IF(A60:A62=\"B\",B60:B62*2,B60:B62*2.5))");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 19: Multi-dimensional array with 2D range
		ws.getRange2("A65").setValue("1");
		ws.getRange2("A66").setValue("2");
		ws.getRange2("B65").setValue("3");
		ws.getRange2("B66").setValue("4");
		ws.getRange2("C65").setValue("5");
		ws.getRange2("C66").setValue("6");

		formula = "=A65:C66*10";
		fillRange = ws.getRange2("E65");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E65").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E65"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "A65:C66*10", "formula result -> A65:C66*10");
		assert.strictEqual(dynamicRef.getHeight(), 2, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 3, "width dynamic array: " + formula);

		// Test 20: Complex formula with MOD and multiple conditions
		ws.getRange2("A70").setValue("5");
		ws.getRange2("A71").setValue("10");
		ws.getRange2("A72").setValue("15");
		ws.getRange2("A73").setValue("20");
		ws.getRange2("B70").setValue("2");
		ws.getRange2("B71").setValue("3");
		ws.getRange2("B72").setValue("4");
		ws.getRange2("B73").setValue("5");

		formula = "=IF(MOD(A70:A73,B70:B73)=0,A70:A73/B70:B73,A70:A73*B70:B73)";
		fillRange = ws.getRange2("D70");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D70").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D70"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(MOD(A70:A73,B70:B73)=0,A70:A73/B70:B73,A70:A73*B70:B73)", "formula result -> IF(MOD(A70:A73,B70:B73)=0,A70:A73/B70:B73,A70:A73*B70:B73)");
		assert.strictEqual(dynamicRef.getHeight(), 4, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 21: Multiple nested array operations with different dimensions
		ws.getRange2("A75").setValue("1");
		ws.getRange2("A76").setValue("2");
		ws.getRange2("A77").setValue("3");
		ws.getRange2("B75").setValue("4");
		ws.getRange2("B76").setValue("5");
		ws.getRange2("B77").setValue("6");
		ws.getRange2("C75").setValue("2");
		ws.getRange2("C76").setValue("3");
		ws.getRange2("C77").setValue("4");

		formula = "=SQRT((A75:A77^2)+(B75:B77^2))/(C75:C77)";
		fillRange = ws.getRange2("E75");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E75").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E75"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SQRT((A75:A77^2)+(B75:B77^2))/(C75:C77)", "formula result -> SQRT((A75:A77^2)+(B75:B77^2))/(C75:C77)");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 22: Array with mixed text and numeric operations
		ws.getRange2("A80").setValue("100");
		ws.getRange2("A81").setValue("200");
		ws.getRange2("A82").setValue("300");
		ws.getRange2("B80").setValue("USD");
		ws.getRange2("B81").setValue("EUR");
		ws.getRange2("B82").setValue("GBP");
		ws.getRange2("C80").setValue("1.0");
		ws.getRange2("C81").setValue("0.85");
		ws.getRange2("C82").setValue("0.73");

		formula = "=TEXT(A80:A82*C80:C82,\"#,##0.00\")&\" \"&B80:B82";
		fillRange = ws.getRange2("E80");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E80").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E80"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TEXT(A80:A82*C80:C82,\"#,##0.00\")&\" \"&B80:B82", "formula result -> TEXT(A80:A82*C80:C82,\"#,##0.00\")&\" \"&B80:B82");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 23: Complex conditional array with multiple IF levels
		ws.getRange2("A85").setValue("10");
		ws.getRange2("A86").setValue("50");
		ws.getRange2("A87").setValue("100");
		ws.getRange2("A88").setValue("150");

		formula = "=IF(A85:A88<50,A85:A88*0.9,IF(A85:A88<100,A85:A88*0.85,IF(A85:A88<150,A85:A88*0.8,A85:A88*0.75)))";
		fillRange = ws.getRange2("C85");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C85").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("C85"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF(A85:A88<50,A85:A88*0.9,IF(A85:A88<100,A85:A88*0.85,IF(A85:A88<150,A85:A88*0.8,A85:A88*0.75)))", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 4, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 24: 2D array with cross-multiplication
		ws.getRange2("A90").setValue("1");
		ws.getRange2("A91").setValue("2");
		ws.getRange2("B90").setValue("3");
		ws.getRange2("B91").setValue("4");
		ws.getRange2("D90").setValue("10");
		ws.getRange2("E90").setValue("20");

		formula = "=A90:B91*D90:E90";
		fillRange = ws.getRange2("G90");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G90").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("G90"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "A90:B91*D90:E90", "formula result -> A90:B91*D90:E90");
		assert.strictEqual(dynamicRef.getHeight(), 2, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 2, "width dynamic array: " + formula);

		// Test 25: Array with date calculations
		ws.getRange2("A95").setValue("2024-01-01");
		ws.getRange2("A96").setValue("2024-02-01");
		ws.getRange2("A97").setValue("2024-03-01");
		ws.getRange2("B95").setValue("30");
		ws.getRange2("B96").setValue("60");
		ws.getRange2("B97").setValue("90");

		formula = "=TEXT(A95:A97+B95:B97,\"YYYY-MM-DD\")";
		fillRange = ws.getRange2("D95");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D95").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D95"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TEXT(A95:A97+B95:B97,\"YYYY-MM-DD\")", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 26: Complex array with SUMPRODUCT-like calculation
		ws.getRange2("A100").setValue("10");
		ws.getRange2("A101").setValue("20");
		ws.getRange2("A102").setValue("30");
		ws.getRange2("B100").setValue("5");
		ws.getRange2("B101").setValue("4");
		ws.getRange2("B102").setValue("3");
		ws.getRange2("C100").setValue("1.1");
		ws.getRange2("C101").setValue("1.2");
		ws.getRange2("C102").setValue("1.3");

		formula = "=(A100:A102*B100:B102)*C100:C102";
		fillRange = ws.getRange2("E100");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E100").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E100"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "(A100:A102*B100:B102)*C100:C102", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 27: Array with percentage calculations and formatting
		ws.getRange2("A105").setValue("1000");
		ws.getRange2("A106").setValue("2000");
		ws.getRange2("A107").setValue("3000");
		ws.getRange2("B105").setValue("10");
		ws.getRange2("B106").setValue("15");
		ws.getRange2("B107").setValue("20");

		formula = "=A105:A107*(1+B105:B107/100)";
		fillRange = ws.getRange2("D105");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D105").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D105"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "A105:A107*(1+B105:B107/100)", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 28: Multi-condition array with AND/OR logic
		ws.getRange2("A110").setValue("10");
		ws.getRange2("A111").setValue("25");
		ws.getRange2("A112").setValue("40");
		ws.getRange2("B110").setValue("5");
		ws.getRange2("B111").setValue("30");
		ws.getRange2("B112").setValue("35");

		formula = "=IF((A110:A112>20)*(B110:B112>30),\"High\",IF((A110:A112>10)+(B110:B112>20),\"Medium\",\"Low\"))";
		fillRange = ws.getRange2("D110");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D110").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D110"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IF((A110:A112>20)*(B110:B112>30),\"High\",IF((A110:A112>10)+(B110:B112>20),\"Medium\",\"Low\"))", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 29: Array with exponential and logarithmic operations
		ws.getRange2("A115").setValue("2");
		ws.getRange2("A116").setValue("3");
		ws.getRange2("A117").setValue("4");
		ws.getRange2("B115").setValue("10");
		ws.getRange2("B116").setValue("100");
		ws.getRange2("B117").setValue("1000");

		formula = "=ROUND(LOG(B115:B117,A115:A117),4)";
		fillRange = ws.getRange2("D115");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D115").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D115"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUND(LOG(B115:B117,A115:A117),4)", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 30: 3x3 matrix-like operations
		ws.getRange2("A120").setValue("1");
		ws.getRange2("A121").setValue("2");
		ws.getRange2("A122").setValue("3");
		ws.getRange2("B120").setValue("4");
		ws.getRange2("B121").setValue("5");
		ws.getRange2("B122").setValue("6");
		ws.getRange2("C120").setValue("7");
		ws.getRange2("C121").setValue("8");
		ws.getRange2("C122").setValue("9");

		formula = "=(A120:C122*2)+(A120:C122^2)";
		fillRange = ws.getRange2("E120");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E120").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E120"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "(A120:C122*2)+(A120:C122^2)", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 3, "width dynamic array: " + formula);

		// Test 31: Array with IFERROR and complex error handling
		ws.getRange2("A125").setValue("10");
		ws.getRange2("A126").setValue("0");
		ws.getRange2("A127").setValue("5");
		ws.getRange2("B125").setValue("2");
		ws.getRange2("B126").setValue("0");
		ws.getRange2("B127").setValue("0");

		formula = "=IFERROR(A125:A127/B125:B127,IF(B125:B127=0,\"Zero Division\",\"Error\"))";
		fillRange = ws.getRange2("D125");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D125").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D125"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "IFERROR(A125:A127/B125:B127,IF(B125:B127=0,\"Zero Division\",\"Error\"))", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 32: Compound interest calculation array
		ws.getRange2("A130").setValue("1000");
		ws.getRange2("A131").setValue("2000");
		ws.getRange2("A132").setValue("3000");
		ws.getRange2("B130").setValue("5");
		ws.getRange2("B131").setValue("10");
		ws.getRange2("B132").setValue("15");
		ws.getRange2("C130").setValue("1");
		ws.getRange2("C131").setValue("2");
		ws.getRange2("C132").setValue("3");

		formula = "=ROUND(A130:A132*((1+B130:B132/100)^C130:C132),2)";
		fillRange = ws.getRange2("E130");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E130").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E130"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ROUND(A130:A132*((1+B130:B132/100)^C130:C132),2)", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 33: Array with modulo and remainder operations
		ws.getRange2("A135").setValue("23");
		ws.getRange2("A136").setValue("45");
		ws.getRange2("A137").setValue("67");
		ws.getRange2("B135").setValue("5");
		ws.getRange2("B136").setValue("7");
		ws.getRange2("B137").setValue("9");

		formula = "=A135:A137&\" mod \"&B135:B137&\" = \"&MOD(A135:A137,B135:B137)";
		fillRange = ws.getRange2("D135");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D135").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D135"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "A135:A137&\" mod \"&B135:B137&\" = \"&MOD(A135:A137,B135:B137)", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);

		// Test 34: Statistical array operations
		ws.getRange2("A140").setValue("10");
		ws.getRange2("A141").setValue("20");
		ws.getRange2("A142").setValue("30");
		ws.getRange2("B140").setValue("15");
		ws.getRange2("B141").setValue("25");
		ws.getRange2("B142").setValue("35");

		formula = "=SQRT((A140:A142-AVERAGE(A140:A142))^2+(B140:B142-AVERAGE(B140:B142))^2)";
		fillRange = ws.getRange2("D140");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D140").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("D140"));
		dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SQRT((A140:A142-AVERAGE(A140:A142))^2+(B140:B142-AVERAGE(B140:B142))^2)", "formula result");
		assert.strictEqual(dynamicRef.getHeight(), 3, "height dynamic array: " + formula);
		assert.strictEqual(dynamicRef.getWidth(), 1, "width dynamic array: " + formula);
		

		ws.getRange2("A1:Z30").cleanAll();

	});

	QUnit.test("Test: \"Dynamic array blocked expansion (#SPILL! error)\"", function (assert) {
		// Clean up the test area
		ws.getRange2("A1:Z30").cleanAll();

		let fillRange, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		ws.getRange2("A1").setValue("1");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("A3").setValue("3");
		ws.getRange2("C2").setValue("Blocking cell");

		let formula = "=A1:A3*10";
		fillRange = ws.getRange2("C1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("C1"));

		let cellValue = ws.getRange2("C1").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "C1 should contain #SPILL! error when array expansion is blocked");

		ws.getRange2("E1").setValue("5");
		ws.getRange2("E2").setValue("10");
		ws.getRange2("E3").setValue("15");
		ws.getRange2("G3").setValue("Block");

		formula = "=SIN(E1:E3)";
		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("G1"));
		cellValue = ws.getRange2("G1").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "G1 #SPILL! with SIN function");

		ws.getRange2("A5").setValue("100");
		ws.getRange2("A6").setValue("200");
		ws.getRange2("A7").setValue("300");
		ws.getRange2("C6").setValue("X");

		formula = "=A5:A7/10";
		fillRange = ws.getRange2("C5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("C5").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "C5 #SPILL! with division");

		ws.getRange2("E5").setValue("2");
		ws.getRange2("E6").setValue("4");
		ws.getRange2("E7").setValue("6");
		ws.getRange2("G6").setValue("Y");

		formula = "=SQRT(E5:E7)";
		fillRange = ws.getRange2("G5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("G5").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "G5 #SPILL! with SQRT");

		ws.getRange2("A10").setValue("text1");
		ws.getRange2("A11").setValue("text2");
		ws.getRange2("A12").setValue("text3");
		ws.getRange2("C11").setValue("Block");

		formula = "=UPPER(A10:A12)";
		fillRange = ws.getRange2("C10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("C10").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "C10 #SPILL! with UPPER");

		ws.getRange2("E10").setValue("5.5");
		ws.getRange2("E11").setValue("10.8");
		ws.getRange2("E12").setValue("15.3");
		ws.getRange2("G11").setValue("Z");

		formula = "=ROUND(E10:E12,0)";
		fillRange = ws.getRange2("G10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("G10").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "G10 #SPILL! with ROUND");

		ws.getRange2("A15").setValue("1");
		ws.getRange2("A16").setValue("2");
		ws.getRange2("B15").setValue("3");
		ws.getRange2("B16").setValue("4");
		ws.getRange2("D16").setValue("Block");

		formula = "=A15:B16*2";
		fillRange = ws.getRange2("D15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("D15").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "D15 #SPILL! with 2D array");

		ws.getRange2("F15").setValue("10");
		ws.getRange2("F16").setValue("20");
		ws.getRange2("F17").setValue("30");
		ws.getRange2("H16").setValue("X");

		formula = "=ABS(F15:F17-15)";
		fillRange = ws.getRange2("H15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("H15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("H15").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "H15 #SPILL! with ABS");

		ws.getRange2("A20").setValue("5");
		ws.getRange2("A21").setValue("10");
		ws.getRange2("A22").setValue("15");
		ws.getRange2("C21").setValue("Block");

		formula = "=COS(A20:A22)";
		fillRange = ws.getRange2("C20");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C20").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("C20").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "C20 #SPILL! with COS");

		ws.getRange2("E20").setValue("2");
		ws.getRange2("E21").setValue("3");
		ws.getRange2("E22").setValue("4");
		ws.getRange2("G21").setValue("Y");

		formula = "=POWER(E20:E22,2)";
		fillRange = ws.getRange2("G20");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G20").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("G20").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "G20 #SPILL! with POWER");

		ws.getRange2("A25").setValue("hello");
		ws.getRange2("A26").setValue("world");
		ws.getRange2("A27").setValue("test");
		ws.getRange2("C26").setValue("Block");

		formula = "=LEN(A25:A27)";
		fillRange = ws.getRange2("C25");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C25").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("C25").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "C25 #SPILL! with LEN");

		formula = "={10;20;30}+5";
		ws.getRange2("E26").setValue("Block");
		fillRange = ws.getRange2("E25");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E25").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("E25").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "E25 #SPILL! with array constant");

		formula = "=SIN({1;2;3})";
		ws.getRange2("G26").setValue("X");
		fillRange = ws.getRange2("G25");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G25").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("G25").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "G25 #SPILL! with SIN array constant");

		ws.getRange2("I1").setValue("1");
		ws.getRange2("I2").setValue("2");
		ws.getRange2("I3").setValue("3");
		ws.getRange2("K2").setValue("Block");

		formula = "=SQRT(I1:I3+10)";
		fillRange = ws.getRange2("K1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("K1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("K1").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "K1 #SPILL! with SQRT and addition");

		ws.getRange2("I5").setValue("10");
		ws.getRange2("I6").setValue("20");
		ws.getRange2("I7").setValue("30");
		ws.getRange2("K6").setValue("Y");

		formula = "=LOG(I5:I7,10)";
		fillRange = ws.getRange2("K5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("K5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("K5").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "K5 #SPILL! with LOG");

		ws.getRange2("I10").setValue("5");
		ws.getRange2("I11").setValue("10");
		ws.getRange2("I12").setValue("15");
		ws.getRange2("K11").setValue("Block");

		formula = "=MOD(I10:I12,3)";
		fillRange = ws.getRange2("K10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("K10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("K10").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "K10 #SPILL! with MOD");

		ws.getRange2("I15").setValue("text");
		ws.getRange2("I16").setValue("data");
		ws.getRange2("I17").setValue("info");
		ws.getRange2("K16").setValue("X");

		formula = "=LEFT(I15:I17,2)";
		fillRange = ws.getRange2("K15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("K15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("K15").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "K15 #SPILL! with LEFT");

		ws.getRange2("I20").setValue("3.14159");
		ws.getRange2("I21").setValue("2.71828");
		ws.getRange2("I22").setValue("1.41421");
		ws.getRange2("K21").setValue("Block");

		formula = "=ROUND(I20:I22,2)";
		fillRange = ws.getRange2("K20");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("K20").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("K20").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "K20 #SPILL! with ROUND");

		ws.getRange2("M1").setValue("1");
		ws.getRange2("M2").setValue("2");
		ws.getRange2("M3").setValue("3");
		ws.getRange2("O2").setValue("Block");

		formula = "=EXP(M1:M3)";
		fillRange = ws.getRange2("O1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("O1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("O1").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "O1 #SPILL! with EXP");

		ws.getRange2("M5").setValue("Hello");
		ws.getRange2("M6").setValue("World");
		ws.getRange2("M7").setValue("Test");
		ws.getRange2("O6").setValue("Y");

		formula = "=LOWER(M5:M7)";
		fillRange = ws.getRange2("O5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("O5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("O5").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "O5 #SPILL! with LOWER");

		ws.getRange2("M10").setValue("-5");
		ws.getRange2("M11").setValue("-10");
		ws.getRange2("M12").setValue("-15");
		ws.getRange2("O11").setValue("Block");

		formula = "=ABS(M10:M12)";
		fillRange = ws.getRange2("O10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("O10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("O10").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "O10 #SPILL! with ABS");

		ws.getRange2("M15").setValue("2");
		ws.getRange2("M16").setValue("3");
		ws.getRange2("M17").setValue("4");
		ws.getRange2("O16").setValue("X");

		formula = "=FACT(M15:M17)";
		fillRange = ws.getRange2("O15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("O15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("O15").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "O15 #SPILL! with FACT");

		ws.getRange2("Q1").setValue("1");
		ws.getRange2("Q2").setValue("2");
		ws.getRange2("R1").setValue("3");
		ws.getRange2("R2").setValue("4");
		ws.getRange2("T2").setValue("Block");

		formula = "=Q1:R2+10";
		fillRange = ws.getRange2("T1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("T1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("T1").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "T1 #SPILL! with 2D range addition");

		ws.getRange2("Q5").setValue("10");
		ws.getRange2("Q6").setValue("20");
		ws.getRange2("Q7").setValue("30");
		ws.getRange2("S6").setValue("Y");

		formula = "=IF(Q5:Q7>15,Q5:Q7*2,Q5:Q7/2)";
		fillRange = ws.getRange2("S5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("S5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("S5").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "S5 #SPILL! with IF");

		ws.getRange2("Q10").setValue("5");
		ws.getRange2("Q11").setValue("10");
		ws.getRange2("Q12").setValue("15");
		ws.getRange2("S11").setValue("Block");

		formula = "=RADIANS(Q10:Q12)";
		fillRange = ws.getRange2("S10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("S10").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("S10").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "S10 #SPILL! with RADIANS");

		ws.getRange2("Q15").setValue("data1");
		ws.getRange2("Q16").setValue("data2");
		ws.getRange2("Q17").setValue("data3");
		ws.getRange2("S16").setValue("X");

		formula = "=PROPER(Q15:Q17)";
		fillRange = ws.getRange2("S15");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("S15").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("S15").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "S15 #SPILL! with PROPER");

		ws.getRange2("Q20").setValue("100");
		ws.getRange2("Q21").setValue("200");
		ws.getRange2("Q22").setValue("300");
		ws.getRange2("S21").setValue("Block");

		formula = "=SQRT(Q20:Q22)/5";
		fillRange = ws.getRange2("S20");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("S20").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		cellValue = ws.getRange2("S20").getValue();
		assert.strictEqual(cellValue, "#SPILL!", "S20 #SPILL! with nested operations");

		ws.getRange2("A1:Z30").cleanAll();
	});

	QUnit.test("Test: \"Metadata add test\"", function (assert) {
		clearData(0, 0, 100, 200);
		var getMetadata = function () {
			return ws.workbook.metadata;
		};

		var getCellMetadata = function (r, c) {
			var _cell;
			ws.getRange3(r, c, r, c)._foreachNoEmpty(function(cell) {
				_cell = cell;
			});
			return _cell && _cell.formulaParsed && _cell.formulaParsed.getCm();
		};

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Add first array formula
		var formula1 = "=SEQUENCE(3,2)";
		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula1);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var metadata = getMetadata();
		assert.ok(metadata.cellMetadata && metadata.cellMetadata.length > 0, "cellMetadata created after first formula");
		assert.ok(metadata.metadataTypes && metadata.metadataTypes.length > 0, "metadataTypes created");
		assert.ok(metadata.aFutureMetadata && metadata.aFutureMetadata.length > 0, "aFutureMetadata created");

		var cmIndex1 = getCellMetadata(0, 0);
		assert.ok(cmIndex1 > 0, "A1 has metadata");

		var initialMetadataCount = metadata.aFutureMetadata.length;

		// Add second array formula
		var formula2 = "=SEQUENCE(2,3)";
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText(formula2);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var cmIndex2 = getCellMetadata(0, 3);
		assert.ok(cmIndex2 > 0, "D1 has metadata");

		metadata = getMetadata();
		// Check that metadata count hasn't increased (metadata is shared)
		assert.strictEqual(metadata.aFutureMetadata.length, initialMetadataCount, "Metadata is shared between formulas");

		var cellMetadataBlock = metadata.cellMetadata[cmIndex1 - 1];
		assert.ok(cellMetadataBlock, "cellMetadata block exists");
		assert.ok(cellMetadataBlock.t > 0, "cellMetadata has type");

		var typeIndex = cellMetadataBlock.t;
		var metadataType = metadata.metadataTypes[typeIndex - 1];
		assert.ok(metadataType, "metadataType exists");
		assert.strictEqual(metadataType.name, "XLDAPR", "XLDAPR check type");

		var valueIndex = cellMetadataBlock.v;
		var futureBlock = metadata.aFutureMetadata[valueIndex];
		assert.ok(futureBlock, "futureMetadataBlock exists");
		assert.strictEqual(futureBlock.name, "XLDAPR", "XLDAPR check type");

		// Delete first formula
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var cmIndexAfterFirstDelete = getCellMetadata(0, 0);
		assert.ok(!cmIndexAfterFirstDelete || cmIndexAfterFirstDelete === 0, "A1 metadata removed after deletion");
		
		metadata = getMetadata();
		// Metadata should remain as the second formula still uses it
		assert.ok(metadata != null, "Metadata still exists after first formula deletion");
		assert.ok(metadata.aFutureMetadata && metadata.aFutureMetadata.length > 0, "aFutureMetadata still exists");

		var cmIndex2AfterFirstDelete = getCellMetadata(0, 3);
		assert.ok(cmIndex2AfterFirstDelete > 0, "D1 still has metadata");

		// Delete second formula
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var cmIndexAfterSecondDelete = getCellMetadata(0, 3);
		assert.ok(!cmIndexAfterSecondDelete || cmIndexAfterSecondDelete === 0, "D1 metadata removed after deletion");
		
		metadata = getMetadata();
		// Now metadata should be removed as both formulas are deleted
		assert.ok(metadata == null, "Metadata removed after all formulas deleted");

		clearData(0, 0, 10, 20);
	});

	QUnit.module("Sheet structure");
});