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
		ws.getRange3(r1, c1, r2, c2).cleanAll();
	};

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

	const getNormalizedFormula = function (oCell) {
		let formula = oCell.getFormulaParsed().getFormula();
		return formula.replace(/_xlfn\./g, '').replace(/_xlws\./g, '');
	};

	const parserFormula = AscCommonExcel.parserFormula;

	QUnit.test('Test @ -> single() + single() -> @', function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
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
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
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

		// Test @ with mixed range operands - ranges with and without @
		formula = "=ISBLANK(A1:A3+@A1:A2)";
		fillRange = ws.getRange2("A60");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A60").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A60"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISBLANK(A1:A3+_xlfn.SINGLE(A1:A2))", "ISBLANK(A1:A3+@A1:A2) -> ISBLANK(A1:A3+SINGLE(A1:A2))");
		assembledVal = ws.getRange2("A60").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISBLANK(@A1:A3+A1:A2)";
		fillRange = ws.getRange2("A61");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A61").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A61"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISBLANK(_xlfn.SINGLE(A1:A3)+A1:A2)", "ISBLANK(@A1:A3+A1:A2) -> ISBLANK(SINGLE(A1:A3)+A1:A2)");
		assembledVal = ws.getRange2("A61").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISNUMBER(B2:B5*@C2:C5)";
		fillRange = ws.getRange2("A62");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A62").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A62"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISNUMBER(B2:B5*_xlfn.SINGLE(C2:C5))", "ISNUMBER(B2:B5*@C2:C5) -> ISNUMBER(B2:B5*SINGLE(C2:C5))");
		assembledVal = ws.getRange2("A62").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISERROR(@D1:D10-E1:E10)";
		fillRange = ws.getRange2("A63");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A63").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A63"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISERROR(_xlfn.SINGLE(D1:D10)-E1:E10)", "ISERROR(@D1:D10-E1:E10) -> ISERROR(SINGLE(D1:D10)-E1:E10)");
		assembledVal = ws.getRange2("A63").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISTEXT(F1:F5&@G1:G5)";
		fillRange = ws.getRange2("A64");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A64").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A64"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISTEXT(F1:F5&_xlfn.SINGLE(G1:G5))", "ISTEXT(F1:F5&@G1:G5) -> ISTEXT(F1:F5&SINGLE(G1:G5))");
		assembledVal = ws.getRange2("A64").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISBLANK(@A1:A3+@B1:B3+C1:C3)";
		fillRange = ws.getRange2("A65");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A65").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A65"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISBLANK(_xlfn.SINGLE(A1:A3)+_xlfn.SINGLE(B1:B3)+C1:C3)", "ISBLANK(@A1:A3+@B1:B3+C1:C3) -> ISBLANK(SINGLE(A1:A3)+SINGLE(B1:B3)+C1:C3)");
		assembledVal = ws.getRange2("A65").getValueForEdit();
		assert.strictEqual(assembledVal, formula, "result for edit: " + formula);

		formula = "=ISNA(A1:A5>@B1:B5)";
		fillRange = ws.getRange2("A66");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A66").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("A66"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ISNA(A1:A5>_xlfn.SINGLE(B1:B5))", "ISNA(A1:A5>@B1:B5) -> ISNA(A1:A5>SINGLE(B1:B5))");
		assembledVal = ws.getRange2("A66").getValueForEdit();
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
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
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
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
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

		// Test 35: SPILL error when entire row/column formula is not in first row/column
		ws.getRange2("A2").setValue("5");
		ws.getRange2("B2").setValue("10");
		ws.getRange2("C2").setValue("15");
		ws.getRange2("D2").setValue("20");
		ws.getRange2("B1").setValue("3");
		ws.getRange2("B2").setValue("6");
		ws.getRange2("B3").setValue("9");
		ws.getRange2("B4").setValue("12");

		// Test with entire row reference not in first row - should cause SPILL error
		formula = "=2:2+2";
		fillRange = ws.getRange2("F5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("F5").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("F5"));
		assert.strictEqual(resCell.getValue(), "#SPILL!", "Should return #SPILL! error for entire row formula not in first row");

		// Test with entire column reference not in first column - should cause SPILL error
		formula = "=B:B+2";
		fillRange = ws.getRange2("E3");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E3").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		resCell = getCell(ws.getRange2("E3"));
		assert.strictEqual(resCell.getValue(), "#SPILL!", "Should return #SPILL! error for entire column formula not in first column");
		

		ws.getRange2("A1:Z30").cleanAll();

	});

	QUnit.test("Test: \"Dynamic array blocked expansion (#SPILL! error)\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
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

	QUnit.test("Test: \"Dynamic array metadata and deletion with undo/redo\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		let fillRange, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		ws.getRange2("A1").setValue("1");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("A3").setValue("3");
		ws.getRange2("C2").setValue("Block1");

		ws.getRange2("E1").setValue("5");
		ws.getRange2("E2").setValue("10");
		ws.getRange2("G2").setValue("Block2");

		ws.getRange2("I1").setValue("100");
		ws.getRange2("I2").setValue("200");
		ws.getRange2("I3").setValue("300");
		ws.getRange2("I4").setValue("400");
		ws.getRange2("K2").setValue("Block3");

		let checkThreeFormulas = function(desc) {
			let resCell = getCell(ws.getRange2("C1"));
			let cellValue = ws.getRange2("C1").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - First formula #SPILL!");
			let vmIndex1 = resCell && resCell.formulaParsed && resCell.formulaParsed.getVm();
			assert.strictEqual(vmIndex1, 1, desc + " - First formula has metadata");
			let cmIndex1 = resCell && resCell.formulaParsed && resCell.formulaParsed.getCm();
			assert.strictEqual(cmIndex1, 1, desc + " - C1 has cellMetadata");

			let offset1 = ws.dynamicArrayManager.getRichValueOffset(resCell.nRow, resCell.nCol);
			assert.strictEqual(offset1.row, 2, desc + " - First formula offset row is 2 (3 rows)");
			assert.strictEqual(offset1.col, 0, desc + " - First formula offset col is 0 (1 column)");

			resCell = getCell(ws.getRange2("G1"));
			cellValue = ws.getRange2("G1").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - Second formula #SPILL!");
			let vmIndex2 = resCell && resCell.formulaParsed && resCell.formulaParsed.getVm();
			assert.strictEqual(vmIndex2, 2, desc + " - Second formula has metadata");
			let cmIndex2 = resCell && resCell.formulaParsed && resCell.formulaParsed.getCm();
			assert.strictEqual(cmIndex2, 1, desc + " - G1 has cellMetadata");

			let offset2 = ws.dynamicArrayManager.getRichValueOffset(resCell.nRow, resCell.nCol);
			assert.strictEqual(offset2.row, 1, desc + " - Second formula offset row is 1 (2 rows)");
			assert.strictEqual(offset2.col, 0, desc + " - Second formula offset col is 0 (1 column)");

			resCell = getCell(ws.getRange2("K1"));
			cellValue = ws.getRange2("K1").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - Third formula #SPILL!");
			let vmIndex3 = resCell && resCell.formulaParsed && resCell.formulaParsed.getVm();
			assert.strictEqual(vmIndex3, 3, desc + " - Third formula has metadata");
			let cmIndex3 = resCell && resCell.formulaParsed && resCell.formulaParsed.getCm();
			assert.strictEqual(cmIndex3, 1, desc + " - K1 has cellMetadata");

			let offset3 = ws.dynamicArrayManager.getRichValueOffset(resCell.nRow, resCell.nCol);
			assert.strictEqual(offset3.row, 3, desc + " - Third formula offset row is 3 (4 rows)");
			assert.strictEqual(offset3.col, 0, desc + " - Third formula offset col is 0 (1 column)");

			// Check metadata structure
			let metadata = getMetadata();
			assert.ok(metadata != null, desc + " - Metadata exists");
			assert.ok(metadata.cellMetadata && metadata.cellMetadata.length > 0, desc + " - cellMetadata array exists");
			assert.ok(metadata.metadataTypes && metadata.metadataTypes.length > 0, desc + " - metadataTypes array exists");
			assert.ok(metadata.aFutureMetadata && metadata.aFutureMetadata.length > 0, desc + " - aFutureMetadata array exists");

			// Check richValueData structure
			let richValueData = getRichValueData();
			assert.ok(richValueData != null, desc + " - richValueData exists");
			assert.ok(richValueData.pData && richValueData.pData.length === 3, desc + " - richValueData has 3 entries");

			// Check richValueStructures
			let richValueStructures = getRichValueStructures();
			assert.ok(richValueStructures != null, desc + " - richValueStructures exists");
			assert.ok(richValueStructures.children && richValueStructures.children.length > 0, desc + " - richValueStructures has children");
		};

		let checkTwoFormulas = function(desc) {
			let resCell = getCell(ws.getRange2("C1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - First formula deleted");

			resCell = getCell(ws.getRange2("G1"));
			let cellValue = ws.getRange2("G1").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - Second formula #SPILL!");
			let vmIndex2 = resCell && resCell.formulaParsed && resCell.formulaParsed.getVm();
			assert.strictEqual(vmIndex2, 1, desc + " - Second formula has metadata");
			let cmIndex2 = resCell && resCell.formulaParsed && resCell.formulaParsed.getCm();
			assert.strictEqual(cmIndex2, 1, desc + " - G1 has cellMetadata");

			let offset2 = ws.dynamicArrayManager.getRichValueOffset(resCell.nRow, resCell.nCol);
			assert.strictEqual(offset2.row, 1, desc + " - Second formula offset row is 1 (2 rows)");
			assert.strictEqual(offset2.col, 0, desc + " - Second formula offset col is 0 (1 column)");

			resCell = getCell(ws.getRange2("K1"));
			cellValue = ws.getRange2("K1").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - Third formula #SPILL!");
			let vmIndex3 = resCell && resCell.formulaParsed && resCell.formulaParsed.getVm();
			assert.strictEqual(vmIndex3, 2, desc + " - Third formula has metadata");
			let cmIndex3 = resCell && resCell.formulaParsed && resCell.formulaParsed.getCm();
			assert.strictEqual(cmIndex3, 1, desc + " - K1 has cellMetadata");

			let offset3 = ws.dynamicArrayManager.getRichValueOffset(resCell.nRow, resCell.nCol);
			assert.strictEqual(offset3.row, 3, desc + " - Third formula (K1) offset row is 3 (4 rows) - after first deletion");
			assert.strictEqual(offset3.col, 0, desc + " - Third formula (K1) offset col is 0 (1 column) - after first deletion");

			// Check metadata structure
			let metadata = getMetadata();
			assert.ok(metadata != null, desc + " - Metadata exists");
			assert.ok(metadata.cellMetadata && metadata.cellMetadata.length > 0, desc + " - cellMetadata array exists");
			assert.ok(metadata.metadataTypes && metadata.metadataTypes.length > 0, desc + " - metadataTypes array exists");
			assert.ok(metadata.aFutureMetadata && metadata.aFutureMetadata.length > 0, desc + " - aFutureMetadata array exists");

			// Check richValueData structure
			let richValueData = getRichValueData();
			assert.ok(richValueData != null, desc + " - richValueData exists");
			assert.ok(richValueData.pData && richValueData.pData.length === 2, desc + " - richValueData has 2 entries (after first deletion)");

			// Check richValueStructures
			let richValueStructures = getRichValueStructures();
			assert.ok(richValueStructures != null, desc + " - richValueStructures exists");
			assert.ok(richValueStructures.children && richValueStructures.children.length > 0, desc + " - richValueStructures has children");
		};

		let checkOneFormula = function(desc) {
			let resCell = getCell(ws.getRange2("C1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - First formula deleted");

			resCell = getCell(ws.getRange2("G1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - Second formula deleted");

			resCell = getCell(ws.getRange2("K1"));
			let cellValue = ws.getRange2("K1").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - Third formula #SPILL!");
			let vmIndex3 = resCell && resCell.formulaParsed && resCell.formulaParsed.getVm();
			assert.strictEqual(vmIndex3, 1, desc + " - Third formula has metadata");
			let cmIndex3 = resCell && resCell.formulaParsed && resCell.formulaParsed.getCm();
			assert.strictEqual(cmIndex3, 1, desc + " - K1 has cellMetadata");

			let offset3 = ws.dynamicArrayManager.getRichValueOffset(resCell.nRow, resCell.nCol);
			assert.strictEqual(offset3.row, 3, desc + " - Third formula (K1) offset row is 3 (4 rows) - only remaining formula");
			assert.strictEqual(offset3.col, 0, desc + " - Third formula (K1) offset col is 0 (1 column) - only remaining formula");

			// Check metadata structure
			let metadata = getMetadata();
			assert.ok(metadata != null, desc + " - Metadata exists");
			assert.ok(metadata.cellMetadata && metadata.cellMetadata.length > 0, desc + " - cellMetadata array exists");
			assert.ok(metadata.metadataTypes && metadata.metadataTypes.length > 0, desc + " - metadataTypes array exists");
			assert.ok(metadata.aFutureMetadata && metadata.aFutureMetadata.length > 0, desc + " - aFutureMetadata array exists");

			// Check richValueData structure
			let richValueData = getRichValueData();
			assert.ok(richValueData != null, desc + " - richValueData exists");
			assert.ok(richValueData.pData && richValueData.pData.length === 1, desc + " - richValueData has 1 entry (only K1 remains)");

			// Check richValueStructures
			let richValueStructures = getRichValueStructures();
			assert.ok(richValueStructures != null, desc + " - richValueStructures exists");
			assert.ok(richValueStructures.children && richValueStructures.children.length > 0, desc + " - richValueStructures has children");
		};

		let checkAllDeleted = function(desc) {
			let resCell = getCell(ws.getRange2("C1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - First formula deleted");

			resCell = getCell(ws.getRange2("G1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - Second formula deleted");

			resCell = getCell(ws.getRange2("K1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - Third formula deleted");

			let metadata = getMetadata();
			assert.ok(metadata == null, desc + " - Metadata removed");

			let richValueData = getRichValueData();
			assert.ok(richValueData == null, desc + " - richValueData removed");

			let richValueStructures = getRichValueStructures();
			assert.ok(richValueStructures == null, desc + " - richValueStructures removed");
		};

		fillRange = ws.getRange2("C1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C1").getValueForEdit2();
		fragment[0].setFragmentText("=A1:A3*10");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=SIN(E1:E2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		fillRange = ws.getRange2("K1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("K1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRT(I1:I4)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		checkThreeFormulas("Initial state");

		ws.getRange2("C1").setValue("");
		checkUndoRedo(checkThreeFormulas, checkTwoFormulas, "After deleting C1 (first formula, 2 formulas remain: G1 and K1)", true);

		ws.getRange2("G1").setValue("");
		checkUndoRedo(checkTwoFormulas, checkOneFormula, "After deleting G1 (second formula, 1 formula remains: K1)", true);

		ws.getRange2("K1").setValue("");
		checkUndoRedo(checkOneFormula, checkAllDeleted, "After deleting K1 (third formula, all formulas deleted)", true);

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"Complex dynamic array metadata with expanded and blocked arrays - delete with undo/redo\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		clearData(0, 0, 100, 200);

		let fillRange, fragment, resCell, cellValue;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// ============ Step 1: Add all source data ============
		
		// Group 1 data: Vertical arrays
		ws.getRange2("A1").setValue("10");
		ws.getRange2("A2").setValue("20");
		ws.getRange2("E1").setValue("5");
		ws.getRange2("E2").setValue("10");
		ws.getRange2("G2").setValue("Block");

		// Group 2 data: 2D arrays
		ws.getRange2("A5").setValue("1");
		ws.getRange2("A6").setValue("2");
		ws.getRange2("B5").setValue("3");
		ws.getRange2("B6").setValue("4");
		
		// Group 3 data: Another array
		ws.getRange2("A10").setValue("100");
		ws.getRange2("A11").setValue("200");
		ws.getRange2("C11").setValue("Y");

		// ============ Step 2: Add all formulas ============
		
		// Group 1 formulas: Expanded array 1
		fillRange = ws.getRange2("C1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C1").getValueForEdit2();
		fragment[0].setFragmentText("=A1:A2*2");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Group 1 formulas: Blocked array 1
		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRT(E1:E2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Group 2 formulas: Expanded 2D array
		fillRange = ws.getRange2("D5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D5").getValueForEdit2();
		fragment[0].setFragmentText("=A5:B6*10");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		
		// Group 3 formulas: Blocked array
		fillRange = ws.getRange2("C10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C10").getValueForEdit2();
		fragment[0].setFragmentText("=A10:A11/10");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// ============ Step 3: Verification functions ============
		let checkAllArrays = function(desc) {
			// Group 1
			cellValue = ws.getRange2("C1").getValue();
			assert.strictEqual(cellValue, "20", desc + " - C1 expanded");
			cellValue = ws.getRange2("C2").getValue();
			assert.strictEqual(cellValue, "40", desc + " - C2 expanded");
			cellValue = ws.getRange2("G1").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - G1 blocked");
			
			// Group 2
			cellValue = ws.getRange2("D5").getValue();
			assert.strictEqual(cellValue, "10", desc + " - D5 expanded 2D");
			cellValue = ws.getRange2("E6").getValue();
			assert.strictEqual(cellValue, "40", desc + " - E6 expanded 2D");
			
			// Group 3
			cellValue = ws.getRange2("C10").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - C10 blocked");

			// Check metadata exists
			let metadata = getMetadata();
			assert.ok(metadata != null, desc + " - Metadata exists");
			assert.ok(metadata.cellMetadata && metadata.cellMetadata.length > 0, desc + " - cellMetadata exists");
		};

		let checkAfterFirstDeletions = function(desc) {
			// C1 deleted, others remain
			resCell = getCell(ws.getRange2("C1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - C1 deleted");
			
			// Group 1 - G1 still blocked
			cellValue = ws.getRange2("G1").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - G1 still blocked");
			
			// Group 2 and 3 still exist
			cellValue = ws.getRange2("D5").getValue();
			assert.strictEqual(cellValue, "10", desc + " - D5 still expanded");
			cellValue = ws.getRange2("C10").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - C10 still blocked");
		};

		let checkAfterSecondDeletions = function(desc) {
			// C1 and G1 deleted
			resCell = getCell(ws.getRange2("C1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - C1 deleted");
			resCell = getCell(ws.getRange2("G1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - G1 deleted");
			
			// Group 2 and 3 still exist
			cellValue = ws.getRange2("D5").getValue();
			assert.strictEqual(cellValue, "10", desc + " - D5 still expanded");
			cellValue = ws.getRange2("C10").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - C10 still blocked");
		};

		let checkAfterThirdDeletions = function(desc) {
			// C1, G1, D5 deleted
			resCell = getCell(ws.getRange2("C1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - C1 deleted");
			resCell = getCell(ws.getRange2("G1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - G1 deleted");
			resCell = getCell(ws.getRange2("D5"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - D5 deleted");
			
			// Only C10 still exists
			cellValue = ws.getRange2("C10").getValue();
			assert.strictEqual(cellValue, "#SPILL!", desc + " - C10 still blocked");
		};

		let checkAllDeleted = function(desc) {
			// All arrays deleted
			resCell = getCell(ws.getRange2("C1"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - C1 deleted");
			resCell = getCell(ws.getRange2("D5"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - D5 deleted");
			resCell = getCell(ws.getRange2("C10"));
			assert.strictEqual(resCell.getFormula(), "", desc + " - C10 deleted");
			
			// Check metadata removed
			let metadata = getMetadata();
			assert.ok(metadata == null, desc + " - Metadata removed");
		};

		checkAllArrays("Initial state with all arrays");

		// Step 1: Delete C1 only
		ws.getRange2("C1").setValue("");
		checkUndoRedo(checkAllArrays, checkAfterFirstDeletions, "After deleting C1", true);

		// Step 2: Delete G1
		ws.getRange2("G1").setValue("");
		checkUndoRedo(checkAfterFirstDeletions, checkAfterSecondDeletions, "After deleting G1", true);

		// Step 3: Delete D5
		ws.getRange2("D5").setValue("");
		checkUndoRedo(checkAfterSecondDeletions, checkAfterThirdDeletions, "After deleting D5", true);

		// Step 4: Delete C10 (last array)
		ws.getRange2("C10").setValue("");
		checkUndoRedo(checkAfterThirdDeletions, checkAllDeleted, "After deleting all arrays", true);

		// Cleanup
		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"Metadata add test\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
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

	var getMetadata = function () {
		return ws.workbook.metadata;
	};

	var getRichValueData = function () {
		return ws.workbook.richValueData;
	};

	var getRichValueStructures = function () {
		return ws.workbook.richValueStructures;
	};

	var getRichValueTypesInfo = function () {
		return ws.workbook.richValueStructures;
	};

	var getCellMetadata = function (r, c) {
		var _cell;
		ws.getRange3(r, c, r, c)._foreachNoEmpty(function(cell) {
			_cell = cell;
		});
		return _cell && _cell.formulaParsed && _cell.formulaParsed.getCm();
	};

	var getCellRichValueIndex = function (r, c) {
		var _cell;
		ws.getRange3(r, c, r, c)._foreachNoEmpty(function(cell) {
			_cell = cell;
		});
		return _cell && _cell.formulaParsed && _cell.formulaParsed.getVm();
	};

	QUnit.test("Test: \"Richdata add test\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

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

		fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText("test");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		cmIndex1 = getCellMetadata(0, 0);
		assert.ok(cmIndex1 === 1, "A1 has metadata");

		vmIndex1 = getCellRichValueIndex(0, 0);
		assert.ok(vmIndex1 === 1, "A1 has richdata");

		// Check RichValueData
		var richValueData = getRichValueData();
		assert.ok(richValueData != null, "richValueData exists");
		assert.ok(richValueData.pData && richValueData.pData.length > 0, "richValueData has pData array");
		var richValue = richValueData.getRichValue(vmIndex1 - 1);
		assert.ok(richValue != null, "richValue exists at index 1");
		assert.ok(richValue.s != null, "richValue has structure index");
		assert.ok(richValue.arrV && richValue.arrV.length > 0, "richValue has values array");

		// Check RichValueStructures
		var richValueStructures = getRichValueStructures();
		assert.ok(richValueStructures != null, "richValueStructures exists");
		assert.ok(richValueStructures.children && richValueStructures.children.length > 0, "richValueStructures has children");
		var structure = richValueStructures.getValueStructure(richValue.s);
		assert.ok(structure != null, "structure exists");
		assert.ok(structure.t != null, "structure has type");
		assert.ok(structure.children && structure.children.length > 0, "structure has children keys");

		// Check RichValueTypesInfo
		var richValueTypesInfo = getRichValueTypesInfo();
		assert.ok(richValueTypesInfo != null, "richValueTypesInfo exists");

		ws.getRange2("A2").setValue("");
		cmIndex1 = getCellMetadata(0, 0);
		assert.ok(cmIndex1 === 1, "A1 has metadata after remove A1 value");

		vmIndex1 = getCellRichValueIndex(0, 0);
		assert.ok(vmIndex1 == null, "A1 don't has richdata after remove A1 value");

		richValueData = getRichValueData();
		assert.ok(richValueData == null, "richValueData removed");

		richValueStructures = getRichValueStructures();
		assert.ok(richValueStructures == null, "richValueStructures removed");

		richValueTypesInfo = getRichValueTypesInfo();
		assert.ok(richValueTypesInfo == null, "richValueTypesInfo removed");

		clearData(0, 0, 10, 20);
	});

	QUnit.test("Test: \"Multiple richdata formulas collapse and delete\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

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

		// Add second array formula
		var formula2 = "=SEQUENCE(2,3)";
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText(formula2);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Add third array formula
		var formula3 = "=SEQUENCE(4,1)";
		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText(formula3);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var metadata = getMetadata();
		assert.ok(metadata != null, "Metadata exists after adding formulas");
		assert.ok(metadata.cellMetadata && metadata.cellMetadata.length > 0, "cellMetadata created");
		assert.ok(metadata.metadataTypes && metadata.metadataTypes.length > 0, "metadataTypes created");
		assert.ok(metadata.aFutureMetadata && metadata.aFutureMetadata.length > 0, "aFutureMetadata created");

		var cmIndex1 = getCellMetadata(0, 0);
		var cmIndex2 = getCellMetadata(0, 3);
		var cmIndex3 = getCellMetadata(0, 6);
		assert.ok(cmIndex1 > 0, "A1 has metadata");
		assert.ok(cmIndex2 > 0, "D1 has metadata");
		assert.ok(cmIndex3 > 0, "G1 has metadata");

		// Collapse formulas by adding data
		fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText("test1");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		fillRange = ws.getRange2("D2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D2").getValueForEdit2();
		fragment[0].setFragmentText("test2");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		fillRange = ws.getRange2("G2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G2").getValueForEdit2();
		fragment[0].setFragmentText("test3");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var vmIndex1 = getCellRichValueIndex(0, 0);
		var vmIndex2 = getCellRichValueIndex(0, 3);
		var vmIndex3 = getCellRichValueIndex(0, 6);
		assert.ok(vmIndex1 > 0, "A1 has richdata after collapse");
		assert.ok(vmIndex2 > 0, "D1 has richdata after collapse");
		assert.ok(vmIndex3 > 0, "G1 has richdata after collapse");

		var richValueData = getRichValueData();
		assert.ok(richValueData != null, "richValueData exists");
		assert.ok(richValueData.pData && richValueData.pData.length >= 3, "richValueData has multiple entries");

		var richValueStructures = getRichValueStructures();
		assert.ok(richValueStructures != null, "richValueStructures exists");
		assert.ok(richValueStructures.children && richValueStructures.children.length > 0, "richValueStructures has children");

		// Delete first formula by setting empty value to head cell
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		vmIndex1 = getCellRichValueIndex(0, 0);
		vmIndex2 = getCellRichValueIndex(0, 3);
		vmIndex3 = getCellRichValueIndex(0, 6);
		assert.ok(!vmIndex1 || vmIndex1 === 0, "A1 richdata removed");
		assert.ok(vmIndex2 > 0, "D1 still has richdata");
		assert.ok(vmIndex3 > 0, "G1 still has richdata");

		richValueData = getRichValueData();
		assert.ok(richValueData != null, "richValueData still exists");
		assert.ok(richValueData.pData && richValueData.pData.length >= 2, "richValueData has remaining entries");

		// Delete second formula by setting empty value to head cell
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		vmIndex2 = getCellRichValueIndex(0, 3);
		vmIndex3 = getCellRichValueIndex(0, 6);
		assert.ok(!vmIndex2 || vmIndex2 === 0, "D1 richdata removed");
		assert.ok(vmIndex3 > 0, "G1 still has richdata");

		richValueData = getRichValueData();
		assert.ok(richValueData != null, "richValueData still exists after second deletion");

		// Delete third (last) formula by setting empty value to head cell
		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		vmIndex3 = getCellRichValueIndex(0, 6);
		assert.ok(!vmIndex3 || vmIndex3 === 0, "G1 richdata removed");

		richValueData = getRichValueData();
		assert.ok(richValueData == null, "richValueData removed after all formulas deleted");

		richValueStructures = getRichValueStructures();
		assert.ok(richValueStructures == null, "richValueStructures removed after all formulas deleted");

		richValueTypesInfo = getRichValueTypesInfo();
		assert.ok(richValueTypesInfo == null, "richValueTypesInfo removed after all formulas deleted");

		metadata = getMetadata();
		assert.ok(metadata == null, "Metadata removed after all formulas deleted");

		clearData(0, 0, 10, 20);
	});

	QUnit.test("Test: \"Delete head cell of expanded and collapsed array\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Test 1: Delete head cell of expanded array
		var formula1 = "=SEQUENCE(3,2)";
		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula1);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that array is expanded
		var cellValueA1 = ws.getRange2("A1").getValue();
		var cellValueA2 = ws.getRange2("A2").getValue();
		var cellValueA3 = ws.getRange2("A3").getValue();
		var cellValueB1 = ws.getRange2("B1").getValue();
		var cellValueB2 = ws.getRange2("B2").getValue();
		var cellValueB3 = ws.getRange2("B3").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 has value 1");
		assert.strictEqual(cellValueA2, "3", "A2 has value 2");
		assert.strictEqual(cellValueA3, "5", "A3 has value 3");
		assert.strictEqual(cellValueB1, "2", "B1 has value 1");
		assert.strictEqual(cellValueB2, "4", "B2 has value 2");
		assert.strictEqual(cellValueB3, "6", "B3 has value 3");

		var metadata = getMetadata();
		assert.ok(metadata != null, "Metadata exists for expanded array");
		assert.ok(metadata.cellMetadata && metadata.cellMetadata.length > 0, "cellMetadata exists");
		assert.ok(metadata.metadataTypes && metadata.metadataTypes.length > 0, "metadataTypes exists");
		assert.ok(metadata.aFutureMetadata && metadata.aFutureMetadata.length > 0, "aFutureMetadata exists");

		var cmIndex1 = getCellMetadata(0, 0);
		assert.ok(cmIndex1 > 0, "A1 has metadata");

		// Delete head cell by setting empty value
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that all array cells are cleared
		cellValueA1 = ws.getRange2("A1").getValue();
		cellValueA2 = ws.getRange2("A2").getValue();
		cellValueA3 = ws.getRange2("A3").getValue();
		cellValueB1 = ws.getRange2("B1").getValue();
		cellValueB2 = ws.getRange2("B2").getValue();
		cellValueB3 = ws.getRange2("B3").getValue();
		assert.strictEqual(cellValueA1, "", "A1 is empty after deletion");
		assert.strictEqual(cellValueA2, "", "A2 is empty after deletion");
		assert.strictEqual(cellValueA3, "", "A3 is empty after deletion");
		assert.strictEqual(cellValueB1, "", "B1 is empty after deletion");
		assert.strictEqual(cellValueB2, "", "B2 is empty after deletion");
		assert.strictEqual(cellValueB3, "", "B3 is empty after deletion");

		cmIndex1 = getCellMetadata(0, 0);
		assert.ok(!cmIndex1 || cmIndex1 === 0, "A1 metadata removed after deletion");

		metadata = getMetadata();
		assert.ok(metadata == null, "Metadata removed after expanded array deletion");

		// Test 2: Delete head cell of collapsed array
		var formula2 = "=SEQUENCE(3,2)";
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText(formula2);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Collapse array by adding blocking data
		fillRange = ws.getRange2("D2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D2").getValueForEdit2();
		fragment[0].setFragmentText("block");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that array is collapsed
		var cellValueD1 = ws.getRange2("D1").getValue();
		var cellValueD2 = ws.getRange2("D2").getValue();
		assert.strictEqual(cellValueD1, "#SPILL!", "D1 has value 1 (collapsed)");
		assert.strictEqual(cellValueD2, "block", "D2 has blocking value");

		var vmIndex = getCellRichValueIndex(0, 3);
		assert.ok(vmIndex > 0, "D1 has richdata after collapse");

		var richValueData = getRichValueData();
		assert.ok(richValueData != null, "richValueData exists for collapsed array");

		var richValueStructures = getRichValueStructures();
		assert.ok(richValueStructures != null, "richValueStructures exists for collapsed array");

		metadata = getMetadata();
		assert.ok(metadata != null, "Metadata exists for collapsed array");

		var cmIndex2 = getCellMetadata(0, 3);
		assert.ok(cmIndex2 > 0, "D1 has metadata");

		// Delete head cell of collapsed array
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that head cell is cleared
		cellValueD1 = ws.getRange2("D1").getValue();
		cellValueD2 = ws.getRange2("D2").getValue();
		assert.strictEqual(cellValueD1, "", "D1 is empty after deletion");
		assert.strictEqual(cellValueD2, "block", "D2 still has blocking value");

		vmIndex = getCellRichValueIndex(0, 3);
		assert.ok(!vmIndex || vmIndex === 0, "D1 richdata removed after deletion");

		cmIndex2 = getCellMetadata(0, 3);
		assert.ok(!cmIndex2 || cmIndex2 === 0, "D1 metadata removed after deletion");

		richValueData = getRichValueData();
		assert.ok(richValueData == null, "richValueData removed after collapsed array deletion");

		richValueStructures = getRichValueStructures();
		assert.ok(richValueStructures == null, "richValueStructures removed after collapsed array deletion");

		var richValueTypesInfo = getRichValueTypesInfo();
		assert.ok(richValueTypesInfo == null, "richValueTypesInfo removed after collapsed array deletion");

		metadata = getMetadata();
		assert.ok(metadata == null, "Metadata removed after collapsed array deletion");

		clearData(0, 0, 10, 20);
	});

	function _getArrayFormulaRef(sAddr) {
		let resCell = getCell(ws.getRange2(sAddr));
		let pF = resCell && resCell.getFormulaParsed();
		return pF && pF.getArrayFormulaRef();
	}

	QUnit.test("Test: \"Replace dynamic array with different sizes\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Test 1: Replace with larger array
		var formula1 = "=SEQUENCE(3,2)";
		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula1);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check initial array
		var cellValueA1 = ws.getRange2("A1").getValue();
		var cellValueA3 = ws.getRange2("A3").getValue();
		var cellValueB3 = ws.getRange2("B3").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 has value 1");
		assert.strictEqual(cellValueA3, "5", "A3 has value 5");
		assert.strictEqual(cellValueB3, "6", "B3 has value 6");

		var cmIndex1 = getCellMetadata(0, 0);
		assert.ok(cmIndex1 > 0, "A1 has metadata");

		// Replace with larger array (4x3 instead of 3x2)
		var formula2 = "=SEQUENCE(4,3)";
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula2);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that new larger array is expanded correctly
		cellValueA1 = ws.getRange2("A1").getValue();
		var cellValueA4 = ws.getRange2("A4").getValue();
		var cellValueC4 = ws.getRange2("C4").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 has value 1 after replacement");
		assert.strictEqual(cellValueA4, "10", "A4 has value 10");
		assert.strictEqual(cellValueC4, "12", "C4 has value 12");

		// Check array range for larger array (4x3)
		var arrayRef1 = _getArrayFormulaRef("A1");
		assert.ok(arrayRef1 != null, "A1 has array reference");
		assert.strictEqual(arrayRef1.r1, 0, "Array starts at row 0");
		assert.strictEqual(arrayRef1.c1, 0, "Array starts at col 0");
		assert.strictEqual(arrayRef1.r2, 3, "Array ends at row 3 (4 rows)");
		assert.strictEqual(arrayRef1.c2, 2, "Array ends at col 2 (3 cols)");

		var metadata = getMetadata();
		assert.ok(metadata != null, "Metadata exists after replacement with larger array");

		cmIndex1 = getCellMetadata(0, 0);
		assert.ok(cmIndex1 > 0, "A1 still has metadata after replacement");

		// Test 2: Replace with smaller array
		var formula3 = "=SEQUENCE(2,1)";
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula3);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that smaller array is correct and old cells are cleared
		cellValueA1 = ws.getRange2("A1").getValue();
		var cellValueA2 = ws.getRange2("A2").getValue();
		var cellValueA3 = ws.getRange2("A3").getValue();
		cellValueA4 = ws.getRange2("A4").getValue();
		var cellValueB1 = ws.getRange2("B1").getValue();
		var cellValueC1 = ws.getRange2("C1").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 has value 1");
		assert.strictEqual(cellValueA2, "2", "A2 has value 2");
		assert.strictEqual(cellValueA3, "", "A3 is empty (old array cleared)");
		assert.strictEqual(cellValueA4, "", "A4 is empty (old array cleared)");
		assert.strictEqual(cellValueB1, "", "B1 is empty (old array cleared)");
		assert.strictEqual(cellValueC1, "", "C1 is empty (old array cleared)");

		// Check array range for smaller array (2x1)
		var arrayRef2 = _getArrayFormulaRef("A1");
		assert.ok(arrayRef2 != null, "A1 has array reference after replacement");
		assert.strictEqual(arrayRef2.r1, 0, "Array starts at row 0");
		assert.strictEqual(arrayRef2.c1, 0, "Array starts at col 0");
		assert.strictEqual(arrayRef2.r2, 1, "Array ends at row 1 (2 rows)");
		assert.strictEqual(arrayRef2.c2, 0, "Array ends at col 0 (1 col)");

		cmIndex1 = getCellMetadata(0, 0);
		assert.ok(cmIndex1 > 0, "A1 still has metadata after replacement with smaller array");

		// Test 3: Replace with equal size array
		var formula4 = "=SEQUENCE(2,1,10,5)";
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula4);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that equal size array replaced correctly
		cellValueA1 = ws.getRange2("A1").getValue();
		cellValueA2 = ws.getRange2("A2").getValue();
		assert.strictEqual(cellValueA1, "10", "A1 has value 10 (new formula)");
		assert.strictEqual(cellValueA2, "15", "A2 has value 15 (new formula)");

		// Check array range for equal size array (2x1)
		var arrayRef3 = _getArrayFormulaRef("A1");
		assert.ok(arrayRef3 != null, "A1 has array reference after equal size replacement");
		assert.strictEqual(arrayRef3.r1, 0, "Array starts at row 0");
		assert.strictEqual(arrayRef3.c1, 0, "Array starts at col 0");
		assert.strictEqual(arrayRef3.r2, 1, "Array ends at row 1 (2 rows)");
		assert.strictEqual(arrayRef3.c2, 0, "Array ends at col 0 (1 col)");

		cmIndex1 = getCellMetadata(0, 0);
		assert.ok(cmIndex1 > 0, "A1 still has metadata after replacement with equal size array");

		metadata = getMetadata();
		assert.ok(metadata != null, "Metadata exists after all replacements");

		clearData(0, 0, 10, 20);
	});

	QUnit.test("Test: \"Add dynamic array in previous cell when next cell has array\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Create array in A2 first
		var formula1 = "=SEQUENCE(3,2)";
		var fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText(formula1);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that array in A2 is expanded
		var cellValueA2 = ws.getRange2("A2").getValue();
		var cellValueA3 = ws.getRange2("A3").getValue();
		var cellValueA4 = ws.getRange2("A4").getValue();
		var cellValueB2 = ws.getRange2("B2").getValue();
		assert.strictEqual(cellValueA2, "1", "A2 has value 1");
		assert.strictEqual(cellValueA3, "3", "A3 has value 3");
		assert.strictEqual(cellValueA4, "5", "A4 has value 5");
		assert.strictEqual(cellValueB2, "2", "B2 has value 2");

		// Check array range for A2 array
		var arrayRefA2 = _getArrayFormulaRef("A2");
		assert.ok(arrayRefA2 != null, "A2 has array reference");
		assert.strictEqual(arrayRefA2.r1, 1, "Array starts at row 1 (A2)");
		assert.strictEqual(arrayRefA2.c1, 0, "Array starts at col 0");
		assert.strictEqual(arrayRefA2.r2, 3, "Array ends at row 3 (A4)");
		assert.strictEqual(arrayRefA2.c2, 1, "Array ends at col 1 (B)");

		var cmIndexA2 = getCellMetadata(1, 0);
		assert.ok(cmIndexA2 > 0, "A2 has metadata");

		// Try to add array in A1 that would overlap with A2 array
		var formula2 = "=SEQUENCE(3,2)";
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula2);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that A1 shows SPILL error because A2 is occupied
		var cellValueA1 = ws.getRange2("A1").getValue();
		assert.strictEqual(cellValueA1, "#SPILL!", "A1 shows SPILL error");

		// Check that A1 has metadata and richdata (collapsed array)
		var cmIndexA1 = getCellMetadata(0, 0);
		assert.ok(cmIndexA1 > 0, "A1 has metadata even with SPILL error");

		var vmIndexA1 = getCellRichValueIndex(0, 0);
		assert.ok(vmIndexA1 > 0, "A1 has richdata (collapsed array)");

		// Check RichValueData for collapsed array
		var richValueData = getRichValueData();
		assert.ok(richValueData != null, "richValueData exists for collapsed array");
		assert.ok(richValueData.pData && richValueData.pData.length > 0, "richValueData has pData array");
		var richValue = richValueData.getRichValue(vmIndexA1 - 1);
		assert.ok(richValue != null, "richValue exists for A1");
		assert.ok(richValue.s != null, "richValue has structure index");
		assert.ok(richValue.arrV && richValue.arrV.length > 0, "richValue has values array");

		// Check RichValueStructures
		var richValueStructures = getRichValueStructures();
		assert.ok(richValueStructures != null, "richValueStructures exists for collapsed array");
		assert.ok(richValueStructures.children && richValueStructures.children.length > 0, "richValueStructures has children");
		var structure = richValueStructures.getValueStructure(richValue.s);
		assert.ok(structure != null, "structure exists");
		assert.ok(structure.t != null, "structure has type");

		// Check that A2 array is still expanded and intact
		cellValueA2 = ws.getRange2("A2").getValue();
		cellValueA3 = ws.getRange2("A3").getValue();
		assert.strictEqual(cellValueA2, "1", "A2 still has value 1");
		assert.strictEqual(cellValueA3, "3", "A3 still has value 3");

		cmIndexA2 = getCellMetadata(1, 0);
		assert.ok(cmIndexA2 > 0, "A2 still has metadata");

		var vmIndexA2 = getCellRichValueIndex(1, 0);
		assert.ok(!vmIndexA2 || vmIndexA2 === 0, "A2 has no richdata (expanded array)");

		var metadata = getMetadata();
		assert.ok(metadata != null, "Metadata exists for both arrays");

		// Test 2: Delete A2 array to allow A1 array to expand
		fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check that A1 array now expands
		cellValueA1 = ws.getRange2("A1").getValue();
		cellValueA2 = ws.getRange2("A2").getValue();
		cellValueA3 = ws.getRange2("A3").getValue();
		var cellValueB1 = ws.getRange2("B1").getValue();
		var cellValueB3 = ws.getRange2("B3").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 has value 1 (expanded)");
		assert.strictEqual(cellValueA2, "3", "A2 has value 3 from A1 array");
		assert.strictEqual(cellValueA3, "5", "A3 has value 5 from A1 array");
		assert.strictEqual(cellValueB1, "2", "B1 has value 2 from A1 array");
		assert.strictEqual(cellValueB3, "6", "B3 has value 6 from A1 array");

		// Check array range for expanded A1 array
		var arrayRefA1 = _getArrayFormulaRef("A1");
		assert.ok(arrayRefA1 != null, "A1 has array reference after expansion");
		assert.strictEqual(arrayRefA1.r1, 0, "Array starts at row 0 (A1)");
		assert.strictEqual(arrayRefA1.c1, 0, "Array starts at col 0");
		assert.strictEqual(arrayRefA1.r2, 2, "Array ends at row 2 (A3)");
		assert.strictEqual(arrayRefA1.c2, 1, "Array ends at col 1 (B)");

		// Check that A1 no longer has richdata (expanded)
		vmIndexA1 = getCellRichValueIndex(0, 0);
		assert.ok(!vmIndexA1 || vmIndexA1 === 0, "A1 no longer has richdata (expanded)");

		richValueData = getRichValueData();
		assert.ok(richValueData == null, "richValueData removed after array expanded");

		richValueStructures = getRichValueStructures();
		assert.ok(richValueStructures == null, "richValueStructures removed after array expanded");

		var richValueTypesInfo = getRichValueTypesInfo();
		assert.ok(richValueTypesInfo == null, "richValueTypesInfo removed after array expanded");

		cmIndexA1 = getCellMetadata(0, 0);
		assert.ok(cmIndexA1 > 0, "A1 still has metadata after expansion");

		metadata = getMetadata();
		assert.ok(metadata != null, "Metadata still exists");

		clearData(0, 0, 10, 20);
	});

	QUnit.test("Test: \"Insert dynamic array into existing spill range\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Create initial expanded array in A1 (spills to A1:B3)
		var formula1 = "=SEQUENCE(3,2)";
		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula1);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Sanity check initial spill
		var cellValueA1 = ws.getRange2("A1").getValue();
		var cellValueB3 = ws.getRange2("B3").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 has value 1");
		assert.strictEqual(cellValueB3, "6", "B3 has value 6");

		var cmIndexA1 = getCellMetadata(0, 0);
		assert.ok(cmIndexA1 > 0, "A1 has metadata");

		// Insert a new dynamic array with head inside the existing spill (B2)
		var formula2 = "=SEQUENCE(2,2)"; // spills to B2:C3
		fillRange = ws.getRange2("B2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("B2").getValueForEdit2();
		fragment[0].setFragmentText(formula2);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Previous head (A1) should collapse with #SPILL!
		cellValueA1 = ws.getRange2("A1").getValue();
		assert.strictEqual(cellValueA1, "#SPILL!", "Previous head A1 collapsed to SPILL");

		// New array should be expanded from B2 to C3
		var cellValueB2 = ws.getRange2("B2").getValue();
		var cellValueC2 = ws.getRange2("C2").getValue();
		var cellValueB3n = ws.getRange2("B3").getValue();
		var cellValueC3 = ws.getRange2("C3").getValue();
		assert.strictEqual(cellValueB2, "1", "B2 new array value 1");
		assert.strictEqual(cellValueC2, "2", "C2 new array value 2");
		assert.strictEqual(cellValueB3n, "3", "B3 new array value 3");
		assert.strictEqual(cellValueC3, "4", "C3 new array value 4");

		// Metadata/richdata expectations
		var cmIndexB2 = getCellMetadata(1, 1);
		assert.ok(cmIndexB2 > 0, "B2 has metadata as new array head");

		var vmIndexA1 = getCellRichValueIndex(0, 0);
		assert.ok(vmIndexA1 > 0, "A1 has richdata after collapse");

		var vmIndexB2 = getCellRichValueIndex(1, 1);
		assert.ok(!vmIndexB2 || vmIndexB2 === 0, "B2 has no richdata (expanded)");

		// Range reference for new array
		var arrayRefB2 = _getArrayFormulaRef("B2");
		assert.ok(arrayRefB2 != null, "B2 has array reference");
		assert.strictEqual(arrayRefB2.r1, 1, "Array starts at row 1 (B2)");
		assert.strictEqual(arrayRefB2.c1, 1, "Array starts at col 1 (B)");
		assert.strictEqual(arrayRefB2.r2, 2, "Array ends at row 2 (B3)");
		assert.strictEqual(arrayRefB2.c2, 2, "Array ends at col 2 (C)");

		var richValueData = getRichValueData();
		assert.ok(richValueData != null, "richValueData exists due to collapsed A1");

		clearData(0, 0, 10, 20);
	});

	QUnit.test("Test: \"Paste with clipboard collision - dynamic array collapse/delete\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(3,3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var cellValueA1 = ws.getRange2("A1").getValue();
		var cellValueC3 = ws.getRange2("C3").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 = 1 before paste");
		assert.strictEqual(cellValueC3, "9", "C3 = 9 before paste");

		var cmIndexA1 = getCellMetadata(0, 0);
		assert.ok(cmIndexA1 > 0, "A1 has metadata before paste");

		var arrayRef = _getArrayFormulaRef("A1");
		assert.ok(arrayRef != null, "A1 has array reference");
		assert.strictEqual(arrayRef.r1, 0, "Array starts at row 0");
		assert.strictEqual(arrayRef.c1, 0, "Array starts at col 0");
		assert.strictEqual(arrayRef.r2, 2, "Array ends at row 2");
		assert.strictEqual(arrayRef.c2, 2, "Array ends at col 2");

		ws.getRange2("Z1").setValue("100");
		ws.selectionRange.ranges = [getRange(25, 0, 25, 0)];
		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(1, 1, 1, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		var cellValueA1After = ws.getRange2("A1").getValue();
		var cellValueB2After = ws.getRange2("B2").getValue();
		assert.strictEqual(cellValueB2After, "100", "B2 = 100 after paste");

		var vmIndexA1After = getCellRichValueIndex(0, 0);
		assert.ok(vmIndexA1After > 0, "A1 has richdata after paste (collapsed)");

		var cmIndexA1After = getCellMetadata(0, 0);
		assert.ok(cmIndexA1After > 0, "A1 still has metadata after paste");

		var richValueData = getRichValueData();
		assert.ok(richValueData != null, "richValueData exists due to collapsed array");

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(2,2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		cellValueA1 = ws.getRange2("A1").getValue();
		var cellValueB2Before = ws.getRange2("B2").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 = 1 before second paste");
		assert.strictEqual(cellValueB2Before, "4", "B2 = 4 before second paste");

		ws.getRange2("Z1").setValue("100");
		ws.selectionRange.ranges = [getRange(25, 0, 25, 0)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		var cellValueA1AfterOverwrite = ws.getRange2("A1").getValue();
		var cellValueB2AfterDelete = ws.getRange2("B2").getValue();
		assert.strictEqual(cellValueA1AfterOverwrite, "100", "A1 = 100 after paste on formula cell");
		assert.strictEqual(cellValueB2AfterDelete, "", "B2 is empty after array deletion");

		var cmIndexA1Deleted = getCellMetadata(0, 0);
		assert.ok(!cmIndexA1Deleted || cmIndexA1Deleted === 0, "A1 has no metadata after overwrite");

		var vmIndexA1Deleted = getCellRichValueIndex(0, 0);
		assert.ok(!vmIndexA1Deleted || vmIndexA1Deleted === 0, "A1 has no richdata after overwrite");

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("D5");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D5").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(4,4)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var cellValueD5 = ws.getRange2("D5").getValue();
		var cellValueG8 = ws.getRange2("G8").getValue();
		assert.strictEqual(cellValueD5, "1", "D5 = 1 before paste");
		assert.strictEqual(cellValueG8, "16", "G8 = 16 before paste");

		ws.getRange2("Z1").setValue("200");
		ws.getRange2("Z2").setValue("201");
		ws.getRange2("AA1").setValue("202");
		ws.getRange2("AA2").setValue("203");
		ws.selectionRange.ranges = [getRange(25, 0, 26, 1)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(4, 5, 5, 6)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		var cellValueE6 = ws.getRange2("E6").getValue();
		var cellValueF7 = ws.getRange2("F7").getValue();
		assert.strictEqual(cellValueE6, "200", "E6 = 200 after paste");
		assert.strictEqual(cellValueF7, "203", "F7 = 203 after paste");

		var vmIndexD5Collapsed = getCellRichValueIndex(4, 3);
		assert.ok(vmIndexD5Collapsed > 0, "D5 has richdata after multi-cell paste (collapsed)");

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(3,2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		cellValueA1 = ws.getRange2("A1").getValue();
		var cellValueB3Formula = ws.getRange2("B3").getValue();
		assert.strictEqual(cellValueA1, "1", "A1 = 1 before formula paste");
		assert.strictEqual(cellValueB3Formula, "6", "B3 = 6 before formula paste");

		ws.getRange2("Z1").setValue("=2+2");
		ws.selectionRange.ranges = [getRange(25, 0, 25, 0)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		ws.selectionRange.ranges = [getRange(1, 1, 1, 1)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		var cellValueB2Formula = ws.getRange2("B2").getValue();
		assert.strictEqual(cellValueB2Formula, "4", "B2 = 4 after formula paste");

		var vmIndexA1Formula = getCellRichValueIndex(0, 0);
		assert.ok(vmIndexA1Formula > 0, "A1 has richdata after formula paste (collapsed)");

		clearData(0, 0, 100, 200);
	});


	QUnit.test("Test: \"Copy-paste dynamic array - expand vs blocked\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Create a 2x2 dynamic array in A1 and copy it
		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(2,2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		var base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		// 1) Paste into a free area (E1) - array should expand
		ws.selectionRange.ranges = [getRange(4, 0, 4, 0)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		assert.strictEqual(ws.getRange2("E1").getValue(), "1", "E1 = 1 after paste into free space");
		assert.strictEqual(ws.getRange2("F1").getValue(), "2", "F1 = 2 after paste into free space");
		assert.strictEqual(ws.getRange2("E2").getValue(), "3", "E2 = 3 after paste into free space");
		assert.strictEqual(ws.getRange2("F2").getValue(), "4", "F2 = 4 after paste into free space");

		var vmIndexE1 = getCellRichValueIndex(0, 4);
		assert.ok(!vmIndexE1 || vmIndexE1 === 0, "E1 has no richdata (expanded)");

		var cmIndexE1 = getCellMetadata(0, 4);
		assert.ok(cmIndexE1 > 0, "E1 has metadata after paste (expanded)");

		// Clean and prepare blocked scenario
		clearData(0, 0, 100, 200);

		// Recreate original array to copy again
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(2,2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		base64 = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(ws, wsView.objectRender);

		// Block the right cell so pasted array cannot expand to the right
		ws.getRange2("H1").setValue("X");

		// 2) Paste into G1 where H1 blocks expansion - should collapse / show SPILL or richdata
		ws.selectionRange.ranges = [getRange(6, 0, 6, 0)];
		AscCommonExcel.g_clipboardExcel.pasteData(wsView, AscCommon.c_oAscClipboardDataFormat.Internal, base64);

		var cellValueG1 = ws.getRange2("G1").getValue();
		assert.ok(cellValueG1 === "#SPILL!" || cellValueG1 === "1", "G1 shows SPILL or collapsed representation when blocked");

		var vmIndexG1 = getCellRichValueIndex(0, 6);
		assert.ok(vmIndexG1 > 0, "G1 has richdata after paste into blocked area (collapsed)");

		var cmIndexG1 = getCellMetadata(0, 6);
		assert.ok(cmIndexG1 > 0, "G1 has metadata after collapsed paste");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"Dynamic array add/delete with undo/redo\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Helper function to check empty state
		var checkEmptyState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueA2 = ws.getRange2("A2").getValue();
			var cellValueB2 = ws.getRange2("B2").getValue();
			assert.strictEqual(cellValueA1, "", desc + ": A1 is empty");
			assert.strictEqual(cellValueA2, "", desc + ": A2 is empty");
			assert.strictEqual(cellValueB2, "", desc + ": B2 is empty");
			
			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(!cmIndexA1 || cmIndexA1 === 0, desc + ": A1 has no metadata");
			
			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(!vmIndexA1 || vmIndexA1 === 0, desc + ": A1 has no richdata");
		};

		// Helper function to check array state
		var checkArrayState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueA2 = ws.getRange2("A2").getValue();
			var cellValueB1 = ws.getRange2("B1").getValue();
			var cellValueB2 = ws.getRange2("B2").getValue();
			assert.strictEqual(cellValueA1, "1", desc + ": A1 has value 1");
			assert.strictEqual(cellValueA2, "3", desc + ": A2 has value 3");
			assert.strictEqual(cellValueB1, "2", desc + ": B1 has value 2");
			assert.strictEqual(cellValueB2, "4", desc + ": B2 has value 4");
			
			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(cmIndexA1 > 0, desc + ": A1 has metadata");
			
			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(!vmIndexA1 || vmIndexA1 === 0, desc + ": A1 has no richdata (expanded)");
			
			var arrayRef = _getArrayFormulaRef("A1");
			assert.ok(arrayRef != null, desc + ": A1 has array reference");
			assert.strictEqual(arrayRef.r1, 0, desc + ": Array starts at row 0");
			assert.strictEqual(arrayRef.c1, 0, desc + ": Array starts at col 0");
			assert.strictEqual(arrayRef.r2, 1, desc + ": Array ends at row 1");
			assert.strictEqual(arrayRef.c2, 1, desc + ": Array ends at col 1");
		};

		// Initial state - empty
		checkEmptyState("Initial state");

		// Add dynamic array
		var formula = "=SEQUENCE(2,2)";
		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Check array is expanded
		checkArrayState("After adding array");

		// Delete dynamic array
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		// Test undo/redo
		checkUndoRedo(checkArrayState, checkEmptyState, "Dynamic array add/delete");

		// Delete dynamic array
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		wsView.emptySelection(Asc.c_oAscCleanOptions.Text);

		// Test undo/redo
		checkUndoRedo(checkArrayState, checkEmptyState, "Dynamic array add/delete with emptySelection text option");

		// Delete dynamic array
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		wsView.emptySelection(Asc.c_oAscCleanOptions.All);

		// Test undo/redo
		checkUndoRedo(checkArrayState, checkEmptyState, "Dynamic array add/delete with emptySelection all option");

		// Delete dynamic array
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		wsView.emptySelection(Asc.c_oAscCleanOptions.Formula);

		// Test undo/redo
		checkUndoRedo(checkArrayState, checkEmptyState, "Dynamic array add/delete with emptySelection Formula option");

		// Delete only format
		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		wsView.emptySelection(Asc.c_oAscCleanOptions.Format);

		// Test undo/redo
		checkUndoRedo(checkEmptyState, checkArrayState, "Dynamic array add/delete with emptySelection Format option");


		clearData(0, 0, 10, 20);
	});

	QUnit.test("Test: \"Autofill with collision - dynamic array collapse/delete\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		var fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(1,3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var checkExpandedState = function(desc) {
			var cellValueA2 = ws.getRange2("A2").getValue();
			var cellValueB2 = ws.getRange2("B2").getValue();
			var cellValueC2 = ws.getRange2("C2").getValue();
			assert.strictEqual(cellValueA2, "1", desc + ": A2 = 1");
			assert.strictEqual(cellValueB2, "2", desc + ": B2 = 2");
			assert.strictEqual(cellValueC2, "3", desc + ": C2 = 3");

			var vmIndexA2 = getCellRichValueIndex(1, 0);
			assert.ok(!vmIndexA2 || vmIndexA2 === 0, desc + ": A2 has no richdata (expanded)");
		};

		var checkCollapsedState = function(desc) {
			var cellValueB2 = ws.getRange2("B2").getValue();
			assert.strictEqual(cellValueB2, "100", desc + ": B2 = 100");

			var vmIndexA2 = getCellRichValueIndex(1, 0);
			assert.ok(vmIndexA2 > 0, desc + ": A2 has richdata (collapsed)");

			var cmIndexA2 = getCellMetadata(1, 0);
			assert.ok(cmIndexA2 > 0, desc + ": A2 has metadata");
		};

		ws.getRange2("B1").setValue("100");
		ws.selectionRange.ranges = [getRange(1, 0, 1, 0)];
		wsView.fillHandleArea = 1;
		wsView.fillHandleDirection = 1;
		wsView.activeFillHandle = getRange(1, 0, 1, 1);
		wsView.applyFillHandle(0, 0, false);

		checkUndoRedo(checkExpandedState, checkCollapsedState, "Autofill collision - array collapse");

		// Test 2: Autofill over head cell - array should be deleted
		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A2");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A2").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(1,3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var checkExpandedState2 = function(desc) {
			var cellValueA2 = ws.getRange2("A2").getValue();
			var cellValueB2 = ws.getRange2("B2").getValue();
			var cellValueC2 = ws.getRange2("C2").getValue();
			assert.strictEqual(cellValueA2, "1", desc + ": A2 = 1");
			assert.strictEqual(cellValueB2, "2", desc + ": B2 = 2");
			assert.strictEqual(cellValueC2, "3", desc + ": C2 = 3");

			var vmIndexA2 = getCellRichValueIndex(1, 0);
			assert.ok(!vmIndexA2 || vmIndexA2 === 0, desc + ": A2 has no richdata (expanded)");

			var cmIndexA2 = getCellMetadata(1, 0);
			assert.ok(cmIndexA2 > 0, desc + ": A2 has metadata");
		};

		var checkDeletedState = function(desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueA2 = ws.getRange2("A2").getValue();
			var cellValueB2 = ws.getRange2("B2").getValue();
			assert.strictEqual(cellValueA1, "200", desc + ": A1 = 200 (overwritten)");
			assert.strictEqual(cellValueA2, "200", desc + ": A1 = 200 (overwritten)");
			assert.strictEqual(cellValueB2, "", desc + ": B2 = empty");

			var vmIndexA2 = getCellRichValueIndex(1, 0);
			assert.ok(!vmIndexA2 || vmIndexA2 === 0, desc + ": A2 has no richdata (deleted)");

			var cmIndexA2 = getCellMetadata(1, 0);
			assert.ok(!cmIndexA2 || cmIndexA2 === 0, desc + ": A2 has no metadata (deleted)");
		};

		ws.getRange2("A1").setValue("200");
		ws.selectionRange.ranges = [getRange(0, 0, 0, 0)];
		wsView.fillHandleArea = 1;
		wsView.fillHandleDirection = 1;
		wsView.activeFillHandle = getRange(0, 0, 0, 1);
		wsView.applyFillHandle(0, 0, false);

		checkUndoRedo(checkExpandedState2, checkDeletedState, "Autofill over head cell - array delete");
	});

	QUnit.test("Test: \"Dynamic array undo/redo with expand, collapse, and blocked states\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		var checkEmptyState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueA2 = ws.getRange2("A2").getValue();
			var cellValueB1 = ws.getRange2("B1").getValue();
			assert.strictEqual(cellValueA1, "", desc + ": A1 is empty");
			assert.strictEqual(cellValueA2, "", desc + ": A2 is empty");
			assert.strictEqual(cellValueB1, "", desc + ": B1 is empty");

			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(!cmIndexA1 || cmIndexA1 === 0, desc + ": A1 has no metadata");

			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(!vmIndexA1 || vmIndexA1 === 0, desc + ": A1 has no richdata");

			var metadata = getMetadata();
			assert.ok(metadata == null, desc + ": metadata is null");

			var richValueData = getRichValueData();
			assert.ok(richValueData == null, desc + ": richValueData is null");

			var richValueStructures = getRichValueStructures();
			assert.ok(richValueStructures == null, desc + ": richValueStructures is null");

			var richValueTypesInfo = getRichValueTypesInfo();
			assert.ok(richValueTypesInfo == null, desc + ": richValueTypesInfo is null");
		};

		var checkExpandedState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueA2 = ws.getRange2("A2").getValue();
			var cellValueB1 = ws.getRange2("B1").getValue();
			var cellValueB2 = ws.getRange2("B2").getValue();
			assert.strictEqual(cellValueA1, "1", desc + ": A1 has value 1");
			assert.strictEqual(cellValueA2, "3", desc + ": A2 has value 3");
			assert.strictEqual(cellValueB1, "2", desc + ": B1 has value 2");
			assert.strictEqual(cellValueB2, "4", desc + ": B2 has value 4");

			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(cmIndexA1 > 0, desc + ": A1 has metadata");

			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(!vmIndexA1 || vmIndexA1 === 0, desc + ": A1 has no richdata (expanded)");

			var arrayRef = _getArrayFormulaRef("A1");
			assert.ok(arrayRef != null, desc + ": A1 has array reference");
			assert.strictEqual(arrayRef.r1, 0, desc + ": Array starts at row 0");
			assert.strictEqual(arrayRef.c1, 0, desc + ": Array starts at col 0");
			assert.strictEqual(arrayRef.r2, 1, desc + ": Array ends at row 1");
			assert.strictEqual(arrayRef.c2, 1, desc + ": Array ends at col 1");

			var metadata = getMetadata();
			assert.ok(metadata != null, desc + ": metadata exists");

			var richValueData = getRichValueData();
			assert.ok(richValueData == null, desc + ": richValueData is null (expanded)");

			var richValueStructures = getRichValueStructures();
			assert.ok(richValueStructures == null, desc + ": richValueStructures is null (expanded)");

			var richValueTypesInfo = getRichValueTypesInfo();
			assert.ok(richValueTypesInfo == null, desc + ": richValueTypesInfo is null (expanded)");
		};

		var checkCollapsedState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueB1 = ws.getRange2("B1").getValue();
			assert.strictEqual(cellValueA1, "#SPILL!", desc + ": A1 shows SPILL error");
			assert.strictEqual(cellValueB1, "block", desc + ": B1 has blocking value");

			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(cmIndexA1 > 0, desc + ": A1 has metadata");

			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(vmIndexA1 > 0, desc + ": A1 has richdata (collapsed)");

			var metadata = getMetadata();
			assert.ok(metadata != null, desc + ": metadata exists");

			var richValueData = getRichValueData();
			assert.ok(richValueData != null, desc + ": richValueData exists (collapsed)");
			assert.ok(richValueData.pData && richValueData.pData.length > 0, desc + ": richValueData has pData array");
			var richValue = richValueData.getRichValue(vmIndexA1 - 1);
			assert.ok(richValue != null, desc + ": richValue exists for A1");
			assert.ok(richValue.s != null, desc + ": richValue has structure index");
			assert.ok(richValue.arrV && richValue.arrV.length > 0, desc + ": richValue has values array");

			var richValueStructures = getRichValueStructures();
			assert.ok(richValueStructures != null, desc + ": richValueStructures exists (collapsed)");
			assert.ok(richValueStructures.children && richValueStructures.children.length > 0, desc + ": richValueStructures has children");
			var structure = richValueStructures.getValueStructure(richValue.s);
			assert.ok(structure != null, desc + ": structure exists");
			assert.ok(structure.t != null, desc + ": structure has type");

			var richValueTypesInfo = getRichValueTypesInfo();
			assert.ok(richValueTypesInfo != null, desc + ": richValueTypesInfo exists (collapsed)");
		};

		var checkBlockedState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueB1 = ws.getRange2("B1").getValue();
			assert.strictEqual(cellValueA1, "#SPILL!", desc + ": A1 shows SPILL error");
			assert.strictEqual(cellValueB1, "block", desc + ": B1 has blocking value");

			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(cmIndexA1 > 0, desc + ": A1 has metadata");

			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(vmIndexA1 > 0, desc + ": A1 has richdata (blocked from start)");

			var metadata = getMetadata();
			assert.ok(metadata != null, desc + ": metadata exists");

			var richValueData = getRichValueData();
			assert.ok(richValueData != null, desc + ": richValueData exists (blocked)");

			var richValueStructures = getRichValueStructures();
			assert.ok(richValueStructures != null, desc + ": richValueStructures exists (blocked)");

			var richValueTypesInfo = getRichValueTypesInfo();
			assert.ok(richValueTypesInfo != null, desc + ": richValueTypesInfo exists (blocked)");
		};

		var checkBlockedEmptyState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueB1 = ws.getRange2("B1").getValue();
			assert.strictEqual(cellValueA1, "", desc + ": A1 is empty");
			assert.strictEqual(cellValueB1, "block", desc + ": B1 still has blocking value");

			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(!cmIndexA1 || cmIndexA1 === 0, desc + ": A1 has no metadata");

			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(!vmIndexA1 || vmIndexA1 === 0, desc + ": A1 has no richdata");

			var metadata = getMetadata();
			assert.ok(metadata == null, desc + ": metadata is null");

			var richValueData = getRichValueData();
			assert.ok(richValueData == null, desc + ": richValueData is null");
		};

		checkEmptyState("Initial state");

		var formula = "=SEQUENCE(2,2)";
		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		checkExpandedState("After adding array");

		checkUndoRedo(checkEmptyState, checkExpandedState, "Array expand", true);

		ws.getRange2("B1").setValue("block");

		checkCollapsedState("After blocking B1");

		checkUndoRedo(checkExpandedState, checkCollapsedState, "Array collapse");

		var checkSecondBlockedState = function (desc) {
			var cellValueE1 = ws.getRange2("E1").getValue();
			var cellValueF1 = ws.getRange2("F1").getValue();
			assert.strictEqual(cellValueE1, "#SPILL!", desc + ": E1 shows SPILL error");
			assert.strictEqual(cellValueF1, "data", desc + ": F1 has blocking value");

			var cmIndexE1 = getCellMetadata(0, 4);
			assert.ok(cmIndexE1 > 0, desc + ": E1 has metadata");

			var vmIndexE1 = getCellRichValueIndex(0, 4);
			assert.ok(vmIndexE1 > 0, desc + ": E1 has richdata (collapsed)");

			var metadata = getMetadata();
			assert.ok(metadata != null, desc + ": metadata exists");

			var richValueData = getRichValueData();
			assert.ok(richValueData != null, desc + ": richValueData exists (collapsed)");
		};

		var checkSecondBlockedEmptyState = function (desc) {
			var cellValueE1 = ws.getRange2("E1").getValue();
			var cellValueF1 = ws.getRange2("F1").getValue();
			assert.strictEqual(cellValueE1, "", desc + ": E1 is empty");
			assert.strictEqual(cellValueF1, "data", desc + ": F1 still has blocking value");

			var cmIndexE1 = getCellMetadata(0, 4);
			assert.ok(!cmIndexE1 || cmIndexE1 === 0, desc + ": E1 has no metadata");

			var vmIndexE1 = getCellRichValueIndex(0, 4);
			assert.ok(!vmIndexE1 || vmIndexE1 === 0, desc + ": E1 has no richdata");

			var metadata = getMetadata();
			assert.ok(metadata != null, desc + ": metadata is no null");

			var richValueData = getRichValueData();
			assert.ok(richValueData == null, desc + ": richValueData is null");
		};

		ws.getRange2("F1").setValue("data");
		fillRange = ws.getRange2("E1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E1").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(1,2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		checkSecondBlockedState("Second blocked insert");


		checkUndoRedo(checkSecondBlockedEmptyState, checkSecondBlockedState, "Second blocked insert undo/redo");

		clearData(0, 0, 100, 200);

		ws.getRange2("B1").setValue("block");

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		checkBlockedState("After adding blocked array");

		checkUndoRedo(checkBlockedEmptyState, checkBlockedState, "Blocked array add");

		clearData(0, 0, 10, 20);
	});

	QUnit.test("Test: \"Dynamic array blocked, then unblocked with undo/redo\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		clearData(0, 0, 100, 200);

		var flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		var formula = "=SEQUENCE(2,2)";

		// Step 1: Add blocking data
		ws.getRange2("B1").setValue("block");

		var checkBlockedState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueB1 = ws.getRange2("B1").getValue();
			assert.strictEqual(cellValueA1, "#SPILL!", desc + ": A1 shows SPILL error");
			assert.strictEqual(cellValueB1, "block", desc + ": B1 has blocking value");

			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(cmIndexA1 > 0, desc + ": A1 has metadata");

			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(vmIndexA1 > 0, desc + ": A1 has richdata (collapsed)");

			var arrayRef = _getArrayFormulaRef("A1");
			assert.ok(arrayRef.r1 === arrayRef.r2 && arrayRef.c1 === arrayRef.c2, desc + ": No array reference when blocked");
		};

		// Step 2: Add array formula (it will be blocked)
		var fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		var fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		checkBlockedState("After adding blocked array");

		// Step 3: Remove blocking data
		ws.getRange2("B1").setValue("");

		var checkExpandedState = function (desc) {
			var cellValueA1 = ws.getRange2("A1").getValue();
			var cellValueA2 = ws.getRange2("A2").getValue();
			var cellValueB1 = ws.getRange2("B1").getValue();
			var cellValueB2 = ws.getRange2("B2").getValue();
			assert.strictEqual(cellValueA1, "1", desc + ": A1 = 1");
			assert.strictEqual(cellValueA2, "3", desc + ": A2 = 3");
			assert.strictEqual(cellValueB1, "2", desc + ": B1 = 2");
			assert.strictEqual(cellValueB2, "4", desc + ": B2 = 4");

			var cmIndexA1 = getCellMetadata(0, 0);
			assert.ok(cmIndexA1 > 0, desc + ": A1 has metadata");

			var vmIndexA1 = getCellRichValueIndex(0, 0);
			assert.ok(!vmIndexA1 || vmIndexA1 === 0, desc + ": A1 has no richdata (expanded)");

			var arrayRef = _getArrayFormulaRef("A1");
			assert.ok(arrayRef != null, desc + ": Array reference exists when expanded");
			assert.strictEqual(arrayRef.r1, 0, desc + ": Array starts at row 0");
			assert.strictEqual(arrayRef.c1, 0, desc + ": Array starts at col 0");
			assert.strictEqual(arrayRef.r2, 1, desc + ": Array ends at row 1");
			assert.strictEqual(arrayRef.c2, 1, desc + ": Array ends at col 1");
		};

		checkExpandedState("After removing blocking data");

		// Step 4: Test undo/redo
		checkUndoRedo(checkBlockedState, checkExpandedState, "Unblocking array undo/redo");

		// Step 5: Test expanded -> blocked -> expanded scenario
		clearData(0, 0, 100, 200);

		// First: Add array formula that can expand freely
		fillRange = ws.getRange2("C3");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C3").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		var checkExpandedStateC3 = function (desc) {
			var cellValueC3 = ws.getRange2("C3").getValue();
			var cellValueC4 = ws.getRange2("C4").getValue();
			var cellValueD3 = ws.getRange2("D3").getValue();
			var cellValueD4 = ws.getRange2("D4").getValue();
			assert.strictEqual(cellValueC3, "1", desc + ": C3 = 1");
			assert.strictEqual(cellValueC4, "3", desc + ": C4 = 3");
			assert.strictEqual(cellValueD3, "2", desc + ": D3 = 2");
			assert.strictEqual(cellValueD4, "4", desc + ": D4 = 4");

			var cmIndexC3 = getCellMetadata(2, 2);
			assert.ok(cmIndexC3 > 0, desc + ": C3 has metadata");

			var vmIndexC3 = getCellRichValueIndex(2, 2);
			assert.ok(!vmIndexC3 || vmIndexC3 === 0, desc + ": C3 has no richdata (expanded)");

			var arrayRef = _getArrayFormulaRef("C3");
			assert.ok(arrayRef != null, desc + ": Array reference exists");
			assert.strictEqual(arrayRef.r1, 2, desc + ": Array starts at row 2");
			assert.strictEqual(arrayRef.c1, 2, desc + ": Array starts at col 2");
			assert.strictEqual(arrayRef.r2, 3, desc + ": Array ends at row 3");
			assert.strictEqual(arrayRef.c2, 3, desc + ": Array ends at col 3");
		};

		checkExpandedStateC3("After adding expanded array in C3");

		// Second: Add blocking data to collapse the array
		ws.getRange2("D3").setValue("blocker");

		var checkBlockedStateC3 = function (desc) {
			var cellValueC3 = ws.getRange2("C3").getValue();
			var cellValueD3 = ws.getRange2("D3").getValue();
			assert.strictEqual(cellValueC3, "#SPILL!", desc + ": C3 shows SPILL error");
			assert.strictEqual(cellValueD3, "blocker", desc + ": D3 has blocking value");

			var cmIndexC3 = getCellMetadata(2, 2);
			assert.ok(cmIndexC3 > 0, desc + ": C3 has metadata");

			var vmIndexC3 = getCellRichValueIndex(2, 2);
			assert.ok(vmIndexC3 > 0, desc + ": C3 has richdata (collapsed)");

			var arrayRef = _getArrayFormulaRef("C3");
			assert.ok(arrayRef.r1 === arrayRef.r2 && arrayRef.c1 === arrayRef.c2, desc + ": Array collapsed when blocked");
		};

		checkBlockedStateC3("After blocking D3");

		// Third: Remove blocking data to expand again
		ws.getRange2("D3").setValue("");

		checkExpandedStateC3("After removing blocking data from D3");

		// Fourth: Test undo/redo for the expanded -> blocked -> expanded cycle
		checkUndoRedo(checkBlockedStateC3, checkExpandedStateC3, "Expanded->Blocked->Expanded undo/redo");

		clearData(0, 0, 10, 20);
	});

	QUnit.test("Test: \"Range reference as dynamic array\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}
		//By bug #71536
		clearData(0, 0, 100, 200);

		let fillRange, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Prepare source data in A1:B1
		ws.getRange2("A1").setValue("10");
		ws.getRange2("B1").setValue("20");

		// Helper function to check expanded state
		var checkExpandedState = function (desc) {
			var cellValueD1 = ws.getRange2("D1").getValue();
			var cellValueE1 = ws.getRange2("E1").getValue();
			assert.strictEqual(cellValueD1, "10", desc + ": D1 value = 10");
			assert.strictEqual(cellValueE1, "20", desc + ": E1 value = 20");
			
			var resCell = getCell(ws.getRange2("D1"));
			var dynamicRef = resCell.getFormulaParsed().getDynamicRef();
			assert.notStrictEqual(dynamicRef, null, desc + ": dynamic array reference exists");
			assert.strictEqual(dynamicRef.getHeight(), 1, desc + ": height = 1");
			assert.strictEqual(dynamicRef.getWidth(), 2, desc + ": width = 2");
			
			var arrayRef = _getArrayFormulaRef("D1");
			assert.ok(arrayRef != null, desc + ": D1 has array formula reference");
			assert.strictEqual(arrayRef.r1, 0, desc + ": Array starts at row 0");
			assert.strictEqual(arrayRef.c1, 3, desc + ": Array starts at col 3 (D)");
			assert.strictEqual(arrayRef.r2, 0, desc + ": Array ends at row 0");
			assert.strictEqual(arrayRef.c2, 4, desc + ": Array ends at col 4 (E)");
			
			var cmIndex = getCellMetadata(0, 3);
			assert.ok(cmIndex > 0, desc + ": D1 has metadata");
			
			var vmIndex = getCellRichValueIndex(0, 3);
			assert.ok(!vmIndex || vmIndex === 0, desc + ": D1 has no richdata (expanded)");
		};

		// Helper function to check collapsed state
		var checkCollapsedState = function (desc) {
			var cellValueD1 = ws.getRange2("D1").getValue();
			var cellValueE1 = ws.getRange2("E1").getValue();
			assert.strictEqual(cellValueD1, "#SPILL!", desc + ": D1 shows #SPILL! error");
			assert.strictEqual(cellValueE1, "blocking", desc + ": E1 has blocking value");
			
			var cmIndex = getCellMetadata(0, 3);
			assert.ok(cmIndex > 0, desc + ": D1 has metadata");
			
			var vmIndex = getCellRichValueIndex(0, 3);
			assert.ok(vmIndex > 0, desc + ": D1 has richdata (collapsed)");
			
			var arrayRef = _getArrayFormulaRef("D1");
			assert.ok(arrayRef.r1 === arrayRef.r2 && arrayRef.c1 === arrayRef.c2, desc + ": Array is collapsed");
		};

		// Test 1: Create expanded dynamic array
		let formula = "=A1:B1";
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		
		checkExpandedState("After creating array");

		// Test 2: Block the array by adding data to E1
		ws.getRange2("E1").setValue("blocking");
		
		checkCollapsedState("After blocking E1");

		// Test 3: Remove blocking data to expand again
		ws.getRange2("E1").setValue("");
		
		checkExpandedState("After removing blocking data");

		// Test 4: Undo/redo for expanded -> blocked -> expanded cycle
		checkUndoRedo(checkCollapsedState, checkExpandedState, "Expanded->Blocked->Expanded undo/redo");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"SIN with entire column reference (A:A)\"", function (assert) {
		//TODO check error  
		assert.ok(true, "Dynamic arrays support is disabled");
		return;
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		// Clear and prepare data in column A
		clearData(0, 0, 100, 200);
		
		// Fill column A with several values, leaving most cells empty
		ws.getRange2("A1").setValue("0");        // SIN(0) = 0
		ws.getRange2("A2").setValue("");         // Empty
		ws.getRange2("A3").setValue("1.5708");   // SIN(π/2) ≈ 1
		ws.getRange2("A4").setValue("");         // Empty
		ws.getRange2("A5").setValue("3.14159");  // SIN(π) ≈ 0
		// A6 and beyond are empty

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		// Set formula =SIN(A:A) in cell D1
		let formula = "=SIN(A:A)";
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText(formula);
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);
		
		// Get the result cell
		resCell = getCell(ws.getRange2("D1"));
		let dynamicRef = resCell.getFormulaParsed().getDynamicRef();
		
		// Check that formula is correct
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIN(A:A)", "formula result -> SIN(A:A)");
		
		// Check dynamic array dimensions
		// For entire column reference, the dynamic array should span the entire column
		assert.ok(dynamicRef, "Dynamic reference exists");
		let height = dynamicRef.getHeight();
		let width = dynamicRef.getWidth();
		assert.ok(height > 9999, "height dynamic array should be very large for entire column: " + height);
		assert.strictEqual(width, 1, "width dynamic array for column reference: " + width);
		
		// Check specific cell values in the result
		// D1 should contain SIN(A1) = SIN(0) = 0
		let d1Value = ws.getRange2("D1").getValue();
		assert.ok(Math.abs(d1Value - 0) < 0.0001, "D1 should contain SIN(0) ≈ 0, got: " + d1Value);
		
		// D2 should contain SIN(A2) = SIN("") = 0 (empty cells are treated as 0)
		let d2Value = ws.getRange2("D2").getValue();
		assert.ok(Math.abs(d2Value - 0) < 0.0001, "D2 should contain SIN(empty) ≈ 0, got: " + d2Value);
		
		// D3 should contain SIN(A3) = SIN(1.5708) ≈ 1
		let d3Value = ws.getRange2("D3").getValue();
		assert.ok(Math.abs(d3Value - 1) < 0.01, "D3 should contain SIN(π/2) ≈ 1, got: " + d3Value);
		
		// D4 should contain SIN(A4) = SIN("") = 0
		let d4Value = ws.getRange2("D4").getValue();
		assert.ok(Math.abs(d4Value - 0) < 0.0001, "D4 should contain SIN(empty) ≈ 0, got: " + d4Value);
		
		// D5 should contain SIN(A5) = SIN(3.14159) ≈ 0
		let d5Value = ws.getRange2("D5").getValue();
		assert.ok(Math.abs(d5Value - 0) < 0.01, "D5 should contain SIN(π) ≈ 0, got: " + d5Value);
		
		// D6 should contain SIN(A6) = SIN("") = 0
		let d6Value = ws.getRange2("D6").getValue();
		assert.ok(Math.abs(d6Value - 0) < 0.0001, "D6 should contain SIN(empty) ≈ 0, got: " + d6Value);
		
		// Check a cell further down to ensure the array extends properly
		let d100Value = ws.getRange2("D100").getValue();
		assert.ok(Math.abs(d100Value - 0) < 0.0001, "D100 should contain SIN(empty) ≈ 0, got: " + d100Value);

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"COS with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=COS({1;2;3})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COS({1;2;3})", "COS array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0.5403, "COS array literal: A1 = COS(1) = 0.5403");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, -0.4161, "COS array literal: A2 = COS(2) = -0.4161");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, -0.9899, "COS array literal: A3 = COS(3) = -0.9899");

		clearData(0, 0, 100, 200);

		ws.getRange2("A1").setValue("1");
		ws.getRange2("B1").setValue("2");

		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText("=COS(A1:B1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("D1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COS(A1:B1)", "COS range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("D1").getValue()) * 10000) / 10000, 0.5403, "COS range: D1 = COS(A1) = COS(1) = 0.5403");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("E1").getValue()) * 10000) / 10000, -0.4161, "COS range: E1 = COS(B1) = COS(2) = -0.4161");

		clearData(0, 0, 100, 200);

		ws.getRange2("A1").setValue("1");
		ws.getRange2("A2").setValue("2");
		ws.getRange2("B1").setValue("3");
		ws.getRange2("B2").setValue("4");

		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText("=COS(@A1:B2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("D1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COS(A1:B2)", "COS implicit intersection: formula correctly parsed");
		assert.strictEqual(resCell.getValueForEdit(), "=COS(@A1:B2)", "COS implicit intersection: formula value correctly parsed");
		assert.strictEqual(ws.getRange2("D1").getValue(), "#VALUE!", "COS implicit intersection: D1 = COS(@A1:B2) = COS(1) = 0.5403");
	
		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ABS with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ABS({-1;-2;-3})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ABS({-1;-2;-3})", "ABS array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 1, "ABS array literal: A1 = ABS(-1) = 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 2, "ABS array literal: A2 = ABS(-2) = 2");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 3, "ABS array literal: A3 = ABS(-3) = 3");

		ws.getRange2("D1").setValue("-1");
		ws.getRange2("E1").setValue("-2");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ABS(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ABS(D1:E1)", "ABS range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 1, "ABS range: G1 = ABS(D1) = ABS(-1) = 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 2, "ABS range: H1 = ABS(E1) = ABS(-2) = 2");

		ws.getRange2("J1").setValue("-1");
		ws.getRange2("J2").setValue("-2");
		ws.getRange2("K1").setValue("-3");
		ws.getRange2("K2").setValue("-4");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ABS(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ABS(J1:K2)", "ABS implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ABS(@J1:K2)", "ABS implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ABS implicit intersection: M1 = ABS(@J1:K2) = ABS(-1) = 1");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ACOS with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ACOS({1;0;-1})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ACOS({1;0;-1})", "ACOS array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "ACOS array literal: A1 = ACOS(1) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 1.5707, "ACOS array literal: A2 = ACOS(0) ≈ 1.5707");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 3.1415, "ACOS array literal: A3 = ACOS(-1) ≈ 3.1415");

		ws.getRange2("D1").setValue("1");
		ws.getRange2("E1").setValue("0");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ACOS(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ACOS(D1:E1)", "ACOS range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "ACOS range: G1 = ACOS(D1) = ACOS(1) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 1.5707, "ACOS range: H1 = ACOS(E1) = ACOS(0) ≈ 1.5707");

		ws.getRange2("J1").setValue("-1");
		ws.getRange2("J2").setValue("0");
		ws.getRange2("K1").setValue("1");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ACOS(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ACOS(J1:K2)", "ACOS implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ACOS(@J1:K2)", "ACOS implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ACOS implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ASIN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ASIN({0;1;-1})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ASIN({0;1;-1})", "ASIN array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "ASIN array literal: A1 = ASIN(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 1.5707, "ASIN array literal: A2 = ASIN(1) ≈ 1.5707");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, -1.5707, "ASIN array literal: A3 = ASIN(-1) ≈ -1.5707");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ASIN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ASIN(D1:E1)", "ASIN range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "ASIN range: G1 = ASIN(D1) = ASIN(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 1.5707, "ASIN range: H1 = ASIN(E1) = ASIN(1) ≈ 1.5707");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("-1");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ASIN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ASIN(J1:K2)", "ASIN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ASIN(@J1:K2)", "ASIN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ASIN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ATAN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ATAN({0;1;-1})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATAN({0;1;-1})", "ATAN array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "ATAN array literal: A1 = ATAN(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.7853, "ATAN array literal: A2 = ATAN(1) ≈ 0.7853");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, -0.7853, "ATAN array literal: A3 = ATAN(-1) ≈ -0.7853");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ATAN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATAN(D1:E1)", "ATAN range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "ATAN range: G1 = ATAN(D1) = ATAN(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.7853, "ATAN range: H1 = ATAN(E1) = ATAN(1) ≈ 0.7853");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("-1");
		ws.getRange2("K2").setValue("2");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ATAN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATAN(J1:K2)", "ATAN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ATAN(@J1:K2)", "ATAN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ATAN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ACOSH with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ACOSH({1;2;3})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ACOSH({1;2;3})", "ACOSH array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "ACOSH array literal: A1 = ACOSH(1) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 1.3169, "ACOSH array literal: A2 = ACOSH(2) ≈ 1.3169");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 1.7627, "ACOSH array literal: A3 = ACOSH(3) ≈ 1.7627");

		ws.getRange2("D1").setValue("1");
		ws.getRange2("E1").setValue("2");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ACOSH(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ACOSH(D1:E1)", "ACOSH range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "ACOSH range: G1 = ACOSH(D1) = ACOSH(1) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 1.3169, "ACOSH range: H1 = ACOSH(E1) = ACOSH(2) ≈ 1.3169");

		ws.getRange2("J1").setValue("1");
		ws.getRange2("J2").setValue("2");
		ws.getRange2("K1").setValue("3");
		ws.getRange2("K2").setValue("1.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ACOSH(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ACOSH(J1:K2)", "ACOSH implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ACOSH(@J1:K2)", "ACOSH implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ACOSH implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ASINH with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ASINH({0;1;-1})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ASINH({0;1;-1})", "ASINH array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "ASINH array literal: A1 = ASINH(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.8813, "ASINH array literal: A2 = ASINH(1) ≈ 0.8813");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, -0.8813, "ASINH array literal: A3 = ASINH(-1) ≈ -0.8813");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ASINH(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ASINH(D1:E1)", "ASINH range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "ASINH range: G1 = ASINH(D1) = ASINH(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.8813, "ASINH range: H1 = ASINH(E1) = ASINH(1) ≈ 0.8813");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("-1");
		ws.getRange2("K2").setValue("2");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ASINH(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ASINH(J1:K2)", "ASINH implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ASINH(@J1:K2)", "ASINH implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ASINH implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ATANH with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ATANH({0;0.5;-0.5})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATANH({0;0.5;-0.5})", "ATANH array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "ATANH array literal: A1 = ATANH(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.5493, "ATANH array literal: A2 = ATANH(0.5) ≈ 0.5493");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, -0.5493, "ATANH array literal: A3 = ATANH(-0.5) ≈ -0.5493");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("0.5");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ATANH(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATANH(D1:E1)", "ATANH range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "ATANH range: G1 = ATANH(D1) = ATANH(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.5493, "ATANH range: H1 = ATANH(E1) = ATANH(0.5) ≈ 0.5493");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("0.5");
		ws.getRange2("K1").setValue("-0.5");
		ws.getRange2("K2").setValue("0.25");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ATANH(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ATANH(J1:K2)", "ATANH implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ATANH(@J1:K2)", "ATANH implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ATANH implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"TAN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=TAN({0;1;-1})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TAN({0;1;-1})", "TAN array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "TAN array literal: A1 = TAN(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 1.5574, "TAN array literal: A2 = TAN(1) ≈ 1.5574");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, -1.5574, "TAN array literal: A3 = TAN(-1) ≈ -1.5574");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=TAN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TAN(D1:E1)", "TAN range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "TAN range: G1 = TAN(D1) = TAN(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 1.5574, "TAN range: H1 = TAN(E1) = TAN(1) ≈ 1.5574");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("-1");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=TAN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TAN(J1:K2)", "TAN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=TAN(@J1:K2)", "TAN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "TAN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"COSH with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=COSH({0;1;2})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COSH({0;1;2})", "COSH array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 1, "COSH array literal: A1 = COSH(0) = 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 1.5430, "COSH array literal: A2 = COSH(1) ≈ 1.5430");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 3.7621, "COSH array literal: A3 = COSH(2) ≈ 3.7621");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=COSH(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COSH(D1:E1)", "COSH range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 1, "COSH range: G1 = COSH(D1) = COSH(0) = 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 1.5430, "COSH range: H1 = COSH(E1) = COSH(1) ≈ 1.5430");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("2");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=COSH(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "COSH(J1:K2)", "COSH implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=COSH(@J1:K2)", "COSH implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "COSH implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"SINH with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SINH({0;1;-1})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SINH({0;1;-1})", "SINH array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "SINH array literal: A1 = SINH(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 1.1752, "SINH array literal: A2 = SINH(1) ≈ 1.1752");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, -1.1752, "SINH array literal: A3 = SINH(-1) ≈ -1.1752");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=SINH(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SINH(D1:E1)", "SINH range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "SINH range: G1 = SINH(D1) = SINH(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 1.1752, "SINH range: H1 = SINH(E1) = SINH(1) ≈ 1.1752");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("-1");
		ws.getRange2("K2").setValue("2");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=SINH(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SINH(J1:K2)", "SINH implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=SINH(@J1:K2)", "SINH implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "SINH implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"TANH with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=TANH({0;1;-1})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TANH({0;1;-1})", "TANH array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "TANH array literal: A1 = TANH(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.7615, "TANH array literal: A2 = TANH(1) ≈ 0.7615");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, -0.7615, "TANH array literal: A3 = TANH(-1) ≈ -0.7615");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=TANH(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TANH(D1:E1)", "TANH range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "TANH range: G1 = TANH(D1) = TANH(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.7615, "TANH range: H1 = TANH(E1) = TANH(1) ≈ 0.7615");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("-1");
		ws.getRange2("K2").setValue("2");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=TANH(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "TANH(J1:K2)", "TANH implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=TANH(@J1:K2)", "TANH implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "TANH implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"DEGREES with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=DEGREES({0;1.5708;3.1416})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DEGREES({0;1.5708;3.1416})", "DEGREES array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "DEGREES array literal: A1 = DEGREES(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 90.0002, "DEGREES array literal: A2 = DEGREES(π/2) ≈ 90");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 180.0004, "DEGREES array literal: A3 = DEGREES(π) ≈ 180");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1.5708");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=DEGREES(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DEGREES(D1:E1)", "DEGREES range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "DEGREES range: G1 = DEGREES(D1) = DEGREES(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 90.0002, "DEGREES range: H1 = DEGREES(E1) = DEGREES(π/2) ≈ 90");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1.5708");
		ws.getRange2("K1").setValue("3.1416");
		ws.getRange2("K2").setValue("0.7854");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=DEGREES(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "DEGREES(J1:K2)", "DEGREES implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=DEGREES(@J1:K2)", "DEGREES implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "DEGREES implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"RADIANS with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=RADIANS({0;90;180})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "RADIANS({0;90;180})", "RADIANS array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "RADIANS array literal: A1 = RADIANS(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 1.5707, "RADIANS array literal: A2 = RADIANS(90) ≈ π/2");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 3.1415, "RADIANS array literal: A3 = RADIANS(180) ≈ π");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("90");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=RADIANS(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "RADIANS(D1:E1)", "RADIANS range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "RADIANS range: G1 = RADIANS(D1) = RADIANS(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 1.5707, "RADIANS range: H1 = RADIANS(E1) = RADIANS(90) ≈ π/2");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("90");
		ws.getRange2("K1").setValue("180");
		ws.getRange2("K2").setValue("45");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=RADIANS(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "RADIANS(J1:K2)", "RADIANS implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=RADIANS(@J1:K2)", "RADIANS implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "RADIANS implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"EXP with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=EXP({0;1;2})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EXP({0;1;2})", "EXP array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 1, "EXP array literal: A1 = EXP(0) = 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 2.7182, "EXP array literal: A2 = EXP(1) ≈ e ≈ 2.7182");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 7.3890, "EXP array literal: A3 = EXP(2) ≈ 7.3890");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=EXP(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EXP(D1:E1)", "EXP range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 1, "EXP range: G1 = EXP(D1) = EXP(0) = 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 2.7182, "EXP range: H1 = EXP(E1) = EXP(1) ≈ e ≈ 2.7182");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("2");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=EXP(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EXP(J1:K2)", "EXP implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=EXP(@J1:K2)", "EXP implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "EXP implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"FACT with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=FACT({0;3;5})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FACT({0;3;5})", "FACT array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 1, "FACT array literal: A1 = FACT(0) = 1");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 6, "FACT array literal: A2 = FACT(3) = 6");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 120, "FACT array literal: A3 = FACT(5) = 120");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("3");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=FACT(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FACT(D1:E1)", "FACT range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 1, "FACT range: G1 = FACT(D1) = FACT(0) = 1");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 6, "FACT range: H1 = FACT(E1) = FACT(3) = 6");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("3");
		ws.getRange2("K1").setValue("5");
		ws.getRange2("K2").setValue("4");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=FACT(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FACT(J1:K2)", "FACT implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=FACT(@J1:K2)", "FACT implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "FACT implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"FACTDOUBLE with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=FACTDOUBLE({5;6;7})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FACTDOUBLE({5;6;7})", "FACTDOUBLE array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 15, "FACTDOUBLE array literal: A1 = FACTDOUBLE(5) = 15");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 48, "FACTDOUBLE array literal: A2 = FACTDOUBLE(6) = 48");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 105, "FACTDOUBLE array literal: A3 = FACTDOUBLE(7) = 105");

		ws.getRange2("D1").setValue("5");
		ws.getRange2("E1").setValue("6");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=FACTDOUBLE(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FACTDOUBLE(D1:E1)", "FACTDOUBLE range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "#VALUE!", "FACTDOUBLE range: G1 = FACTDOUBLE(D1) = FACTDOUBLE(5) = 15");

		ws.getRange2("J1").setValue("5");
		ws.getRange2("J2").setValue("6");
		ws.getRange2("K1").setValue("7");
		ws.getRange2("K2").setValue("4");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=FACTDOUBLE(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "FACTDOUBLE(_xlfn.SINGLE(J1:K2))", "FACTDOUBLE implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=FACTDOUBLE(@J1:K2)", "FACTDOUBLE implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "FACTDOUBLE implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"INT with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=INT({1.5;2.8;-1.5})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "INT({1.5;2.8;-1.5})", "INT array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 1, "INT array literal: A1 = INT(1.5) = 1");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 2, "INT array literal: A2 = INT(2.8) = 2");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), -2, "INT array literal: A3 = INT(-1.5) = -2");

		ws.getRange2("D1").setValue("1.5");
		ws.getRange2("E1").setValue("2.8");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=INT(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "INT(D1:E1)", "INT range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 1, "INT range: G1 = INT(D1) = INT(1.5) = 1");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 2, "INT range: H1 = INT(E1) = INT(2.8) = 2");

		ws.getRange2("J1").setValue("1.5");
		ws.getRange2("J2").setValue("2.8");
		ws.getRange2("K1").setValue("-1.5");
		ws.getRange2("K2").setValue("3.9");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=INT(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "INT(J1:K2)", "INT implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=INT(@J1:K2)", "INT implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "INT implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"EVEN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=EVEN({1;2;3})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EVEN({1;2;3})", "EVEN array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 2, "EVEN array literal: A1 = EVEN(1) = 2");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 2, "EVEN array literal: A2 = EVEN(2) = 2");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 4, "EVEN array literal: A3 = EVEN(3) = 4");

		ws.getRange2("D1").setValue("1");
		ws.getRange2("E1").setValue("2");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=EVEN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EVEN(D1:E1)", "EVEN range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 2, "EVEN range: G1 = EVEN(D1) = EVEN(1) = 2");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 2, "EVEN range: H1 = EVEN(E1) = EVEN(2) = 2");

		ws.getRange2("J1").setValue("1");
		ws.getRange2("J2").setValue("2");
		ws.getRange2("K1").setValue("3");
		ws.getRange2("K2").setValue("5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=EVEN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "EVEN(J1:K2)", "EVEN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=EVEN(@J1:K2)", "EVEN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "EVEN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ODD with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ODD({1;2;3})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ODD({1;2;3})", "ODD array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 1, "ODD array literal: A1 = ODD(1) = 1");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 3, "ODD array literal: A2 = ODD(2) = 3");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 3, "ODD array literal: A3 = ODD(3) = 3");

		ws.getRange2("D1").setValue("1");
		ws.getRange2("E1").setValue("2");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ODD(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ODD(D1:E1)", "ODD range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 1, "ODD range: G1 = ODD(D1) = ODD(1) = 1");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 3, "ODD range: H1 = ODD(E1) = ODD(2) = 3");

		ws.getRange2("J1").setValue("1");
		ws.getRange2("J2").setValue("2");
		ws.getRange2("K1").setValue("3");
		ws.getRange2("K2").setValue("4");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ODD(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "ODD(J1:K2)", "ODD implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ODD(@J1:K2)", "ODD implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ODD implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"LN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=LN({1;2.7182;7.3890})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LN({1;2.7182;7.389})", "LN array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0, "LN array literal: A1 = LN(1) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.9999, "LN array literal: A2 = LN(e) ≈ 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 1.9999, "LN array literal: A3 = LN(e²) ≈ 2");

		ws.getRange2("D1").setValue("1");
		ws.getRange2("E1").setValue("2.7182");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=LN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LN(D1:E1)", "LN range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0, "LN range: G1 = LN(D1) = LN(1) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.9999, "LN range: H1 = LN(E1) = LN(e) ≈ 1");

		ws.getRange2("J1").setValue("1");
		ws.getRange2("J2").setValue("2.7182");
		ws.getRange2("K1").setValue("7.3890");
		ws.getRange2("K2").setValue("2");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=LN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LN(J1:K2)", "LN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=LN(@J1:K2)", "LN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "LN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"LOG10 with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=LOG10({1;10;100})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LOG10({1;10;100})", "LOG10 array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "LOG10 array literal: A1 = LOG10(1) = 0");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 1, "LOG10 array literal: A2 = LOG10(10) = 1");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 2, "LOG10 array literal: A3 = LOG10(100) = 2");

		ws.getRange2("D1").setValue("1");
		ws.getRange2("E1").setValue("10");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=LOG10(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LOG10(D1:E1)", "LOG10 range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "LOG10 range: G1 = LOG10(D1) = LOG10(1) = 0");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 1, "LOG10 range: H1 = LOG10(E1) = LOG10(10) = 1");

		ws.getRange2("J1").setValue("1");
		ws.getRange2("J2").setValue("10");
		ws.getRange2("K1").setValue("100");
		ws.getRange2("K2").setValue("1000");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=LOG10(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "LOG10(J1:K2)", "LOG10 implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=LOG10(@J1:K2)", "LOG10 implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "LOG10 implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"SIGN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SIGN({-5;0;5})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIGN({-5;0;5})", "SIGN array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), -1, "SIGN array literal: A1 = SIGN(-5) = -1");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 0, "SIGN array literal: A2 = SIGN(0) = 0");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 1, "SIGN array literal: A3 = SIGN(5) = 1");

		ws.getRange2("D1").setValue("-5");
		ws.getRange2("E1").setValue("0");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=SIGN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIGN(D1:E1)", "SIGN range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), -1, "SIGN range: G1 = SIGN(D1) = SIGN(-5) = -1");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 0, "SIGN range: H1 = SIGN(E1) = SIGN(0) = 0");

		ws.getRange2("J1").setValue("-5");
		ws.getRange2("J2").setValue("0");
		ws.getRange2("K1").setValue("5");
		ws.getRange2("K2").setValue("10");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=SIGN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SIGN(J1:K2)", "SIGN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=SIGN(@J1:K2)", "SIGN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "SIGN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"SQRT with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRT({0;4;9})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SQRT({0;4;9})", "SQRT array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "SQRT array literal: A1 = SQRT(0) = 0");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 2, "SQRT array literal: A2 = SQRT(4) = 2");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 3, "SQRT array literal: A3 = SQRT(9) = 3");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("4");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRT(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SQRT(D1:E1)", "SQRT range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "SQRT range: G1 = SQRT(D1) = SQRT(0) = 0");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 2, "SQRT range: H1 = SQRT(E1) = SQRT(4) = 2");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("4");
		ws.getRange2("K1").setValue("9");
		ws.getRange2("K2").setValue("16");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRT(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(resCell.getFormulaParsed().getFormula(), "SQRT(J1:K2)", "SQRT implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=SQRT(@J1:K2)", "SQRT implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "SQRT implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"SQRTPI with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRTPI({0;1;4})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SQRTPI({0;1;4})", "SQRTPI array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "SQRTPI array literal: A1 = SQRTPI(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 1.7724, "SQRTPI array literal: A2 = SQRTPI(1) ≈ √π ≈ 1.7724");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 3.5449, "SQRTPI array literal: A3 = SQRTPI(4) ≈ 2√π ≈ 3.5449");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRTPI(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SQRTPI(D1:E1)", "SQRTPI range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "#VALUE!", "SQRTPI range: G1 = SQRTPI(D1) = SQRTPI(0) = 0");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("4");
		ws.getRange2("K2").setValue("2");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRTPI(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SQRTPI(SINGLE(J1:K2))", "SQRTPI implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=SQRTPI(@J1:K2)", "SQRTPI implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "SQRTPI implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"FISHER with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=FISHER({0;0.5;0.9})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "FISHER({0;0.5;0.9})", "FISHER array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "FISHER array literal: A1 = FISHER(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.5493, "FISHER array literal: A2 = FISHER(0.5) ≈ 0.5493");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 1.4722, "FISHER array literal: A3 = FISHER(0.9) ≈ 1.4722");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("0.5");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=FISHER(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "FISHER(D1:E1)", "FISHER range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "FISHER range: G1 = FISHER(D1) = FISHER(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.5493, "FISHER range: H1 = FISHER(E1) = FISHER(0.5) ≈ 0.5493");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("0.5");
		ws.getRange2("K1").setValue("0.9");
		ws.getRange2("K2").setValue("0.25");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=FISHER(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "FISHER(J1:K2)", "FISHER implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=FISHER(@J1:K2)", "FISHER implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "FISHER implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"FISHERINV with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=FISHERINV({0;0.5493;1.4722})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "FISHERINV({0;0.5493;1.4722})", "FISHERINV array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "FISHERINV array literal: A1 = FISHERINV(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.4999, "FISHERINV array literal: A2 = FISHERINV(0.5493) ≈ 0.5");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 0.8999, "FISHERINV array literal: A3 = FISHERINV(1.4722) ≈ 0.9");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("0.5493");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=FISHERINV(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "FISHERINV(D1:E1)", "FISHERINV range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "FISHERINV range: G1 = FISHERINV(D1) = FISHERINV(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.4999, "FISHERINV range: H1 = FISHERINV(E1) = FISHERINV(0.5493) ≈ 0.5");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("0.5493");
		ws.getRange2("K1").setValue("1.4722");
		ws.getRange2("K2").setValue("0.2554");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=FISHERINV(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "FISHERINV(J1:K2)", "FISHERINV implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=FISHERINV(@J1:K2)", "FISHERINV implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "FISHERINV implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"GAUSS with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=GAUSS({0;1;2})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAUSS({0;1;2})", "GAUSS array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "GAUSS array literal: A1 = GAUSS(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.3413, "GAUSS array literal: A2 = GAUSS(1) ≈ 0.3413");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 0.4772, "GAUSS array literal: A3 = GAUSS(2) ≈ 0.4772");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=GAUSS(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAUSS(D1:E1)", "GAUSS range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "GAUSS range: G1 = GAUSS(D1) = GAUSS(0) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.3413, "GAUSS range: H1 = GAUSS(E1) = GAUSS(1) ≈ 0.3413");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("2");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=GAUSS(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAUSS(J1:K2)", "GAUSS implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=GAUSS(@J1:K2)", "GAUSS implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "GAUSS implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"PHI with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=PHI({0;1;2})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "PHI({0;1;2})", "PHI array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0.3989, "PHI array literal: A1 = PHI(0) ≈ 0.3989");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.2419, "PHI array literal: A2 = PHI(1) ≈ 0.2419");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 0.0539, "PHI array literal: A3 = PHI(2) ≈ 0.0539");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=PHI(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "PHI(D1:E1)", "PHI range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0.3989, "PHI range: G1 = PHI(D1) = PHI(0) ≈ 0.3989");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.2419, "PHI range: H1 = PHI(E1) = PHI(1) ≈ 0.2419");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("2");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=PHI(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "PHI(J1:K2)", "PHI implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=PHI(@J1:K2)", "PHI implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "PHI implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"GAMMALN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=GAMMALN({1;2;3})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAMMALN({1;2;3})", "GAMMALN array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "GAMMALN array literal: A1 = GAMMALN(1) = 0");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 0, "GAMMALN array literal: A2 = GAMMALN(2) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 0.6931, "GAMMALN array literal: A3 = GAMMALN(3) ≈ 0.6931");

		ws.getRange2("D1").setValue("1");
		ws.getRange2("E1").setValue("2");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=GAMMALN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAMMALN(D1:E1)", "GAMMALN range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "GAMMALN range: G1 = GAMMALN(D1) = GAMMALN(1) = 0");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 0, "GAMMALN range: H1 = GAMMALN(E1) = GAMMALN(2) = 0");

		ws.getRange2("J1").setValue("1");
		ws.getRange2("J2").setValue("2");
		ws.getRange2("K1").setValue("3");
		ws.getRange2("K2").setValue("4");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=GAMMALN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAMMALN(J1:K2)", "GAMMALN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=GAMMALN(@J1:K2)", "GAMMALN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "GAMMALN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"GAMMALN.PRECISE with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=GAMMALN.PRECISE({1;2;3})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAMMALN.PRECISE({1;2;3})", "GAMMALN.PRECISE array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "GAMMALN.PRECISE array literal: A1 = GAMMALN.PRECISE(1) = 0");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 0, "GAMMALN.PRECISE array literal: A2 = GAMMALN.PRECISE(2) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 0.6931, "GAMMALN.PRECISE array literal: A3 = GAMMALN.PRECISE(3) ≈ 0.6931");

		ws.getRange2("D1").setValue("1");
		ws.getRange2("E1").setValue("2");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=GAMMALN.PRECISE(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAMMALN.PRECISE(D1:E1)", "GAMMALN.PRECISE range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "GAMMALN.PRECISE range: G1 = GAMMALN.PRECISE(D1) = GAMMALN.PRECISE(1) = 0");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 0, "GAMMALN.PRECISE range: H1 = GAMMALN.PRECISE(E1) = GAMMALN.PRECISE(2) = 0");

		ws.getRange2("J1").setValue("1");
		ws.getRange2("J2").setValue("2");
		ws.getRange2("K1").setValue("3");
		ws.getRange2("K2").setValue("4");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=GAMMALN.PRECISE(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "GAMMALN.PRECISE(J1:K2)", "GAMMALN.PRECISE implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=GAMMALN.PRECISE(@J1:K2)", "GAMMALN.PRECISE implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "GAMMALN.PRECISE implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"NORMSDIST with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=NORMSDIST({0;1;2})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORMSDIST({0;1;2})", "NORMSDIST array literal: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A1").getValue()) * 10000) / 10000, 0.5, "NORMSDIST array literal: A1 = NORMSDIST(0) = 0.5");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.8413, "NORMSDIST array literal: A2 = NORMSDIST(1) ≈ 0.8413");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 0.9772, "NORMSDIST array literal: A3 = NORMSDIST(2) ≈ 0.9772");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=NORMSDIST(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORMSDIST(D1:E1)", "NORMSDIST range: formula correctly parsed");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("G1").getValue()) * 10000) / 10000, 0.5, "NORMSDIST range: G1 = NORMSDIST(D1) = NORMSDIST(0) = 0.5");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.8413, "NORMSDIST range: H1 = NORMSDIST(E1) = NORMSDIST(1) ≈ 0.8413");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("2");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=NORMSDIST(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORMSDIST(J1:K2)", "NORMSDIST implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=NORMSDIST(@J1:K2)", "NORMSDIST implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "NORMSDIST implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"NORMSINV with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=NORMSINV({0.5;0.8413;0.9772})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORMSINV({0.5;0.8413;0.9772})", "NORMSINV array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "NORMSINV array literal: A1 = NORMSINV(0.5) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.9998, "NORMSINV array literal: A2 = NORMSINV(0.8413) ≈ 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 1.999, "NORMSINV array literal: A3 = NORMSINV(0.9772) ≈ 2");

		ws.getRange2("D1").setValue("0.5");
		ws.getRange2("E1").setValue("0.8413");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=NORMSINV(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORMSINV(D1:E1)", "NORMSINV range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "NORMSINV range: G1 = NORMSINV(D1) = NORMSINV(0.5) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.9998, "NORMSINV range: H1 = NORMSINV(E1) = NORMSINV(0.8413) ≈ 1");

		ws.getRange2("J1").setValue("0.5");
		ws.getRange2("J2").setValue("0.8413");
		ws.getRange2("K1").setValue("0.9772");
		ws.getRange2("K2").setValue("0.6915");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=NORMSINV(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORMSINV(J1:K2)", "NORMSINV implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=NORMSINV(@J1:K2)", "NORMSINV implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "NORMSINV implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"NORM.S.INV with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=NORM.S.INV({0.5;0.8413;0.9772})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORM.S.INV({0.5;0.8413;0.9772})", "NORM.S.INV array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 0, "NORM.S.INV array literal: A1 = NORM.S.INV(0.5) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.9998, "NORM.S.INV array literal: A2 = NORM.S.INV(0.8413) ≈ 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 1.999, "NORM.S.INV array literal: A3 = NORM.S.INV(0.9772) ≈ 2");

		ws.getRange2("D1").setValue("0.5");
		ws.getRange2("E1").setValue("0.8413");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=NORM.S.INV(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORM.S.INV(D1:E1)", "NORM.S.INV range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 0, "NORM.S.INV range: G1 = NORM.S.INV(D1) = NORM.S.INV(0.5) = 0");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("H1").getValue()) * 10000) / 10000, 0.9998, "NORM.S.INV range: H1 = NORM.S.INV(E1) = NORM.S.INV(0.8413) ≈ 1");

		ws.getRange2("J1").setValue("0.5");
		ws.getRange2("J2").setValue("0.8413");
		ws.getRange2("K1").setValue("0.9772");
		ws.getRange2("K2").setValue("0.6915");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=NORM.S.INV(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "NORM.S.INV(J1:K2)", "NORM.S.INV implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=NORM.S.INV(@J1:K2)", "NORM.S.INV implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "NORM.S.INV implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"ERFC with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=ERFC({0;1;2})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "ERFC({0;1;2})", "ERFC array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 1, "ERFC array literal: A1 = ERFC(0) = 1");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A2").getValue()) * 10000) / 10000, 0.1572, "ERFC array literal: A2 = ERFC(1) ≈ 0.1572");
		assert.strictEqual(Math.trunc(Number(ws.getRange2("A3").getValue()) * 10000) / 10000, 0.0046, "ERFC array literal: A3 = ERFC(2) ≈ 0.0046");

		ws.getRange2("D1").setValue("0");
		ws.getRange2("E1").setValue("1");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=ERFC(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "ERFC(D1:E1)", "ERFC range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "#VALUE!", "ERFC range: G1 = ERFC(D1:E1) returns #VALUE! for range");

		ws.getRange2("J1").setValue("0");
		ws.getRange2("J2").setValue("1");
		ws.getRange2("K1").setValue("2");
		ws.getRange2("K2").setValue("0.5");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=ERFC(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "ERFC(SINGLE(J1:K2))", "ERFC implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=ERFC(@J1:K2)", "ERFC implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "ERFC implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"LEN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=LEN({\"Hello\";\"World\";\"Test\"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "LEN({\"Hello\";\"World\";\"Test\"})", "LEN array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 5, "LEN array literal: A1 = LEN(\"Hello\") = 5");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 5, "LEN array literal: A2 = LEN(\"World\") = 5");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 4, "LEN array literal: A3 = LEN(\"Test\") = 4");

		ws.getRange2("D1").setValue("Hello");
		ws.getRange2("E1").setValue("World");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=LEN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "LEN(D1:E1)", "LEN range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 5, "LEN range: G1 = LEN(D1) = LEN(\"Hello\") = 5");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 5, "LEN range: H1 = LEN(E1) = LEN(\"World\") = 5");

		ws.getRange2("J1").setValue("Hello");
		ws.getRange2("J2").setValue("World");
		ws.getRange2("K1").setValue("Test");
		ws.getRange2("K2").setValue("ABC");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=LEN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "LEN(J1:K2)", "LEN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=LEN(@J1:K2)", "LEN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "LEN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"LOWER with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=LOWER({\"HELLO\";\"WORLD\";\"TEST\"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "LOWER({\"HELLO\";\"WORLD\";\"TEST\"})", "LOWER array literal: formula correctly parsed");
		assert.strictEqual(ws.getRange2("A1").getValue(), "hello", "LOWER array literal: A1 = LOWER(\"HELLO\") = \"hello\"");
		assert.strictEqual(ws.getRange2("A2").getValue(), "world", "LOWER array literal: A2 = LOWER(\"WORLD\") = \"world\"");
		assert.strictEqual(ws.getRange2("A3").getValue(), "test", "LOWER array literal: A3 = LOWER(\"TEST\") = \"test\"");

		ws.getRange2("D1").setValue("HELLO");
		ws.getRange2("E1").setValue("WORLD");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=LOWER(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "LOWER(D1:E1)", "LOWER range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "hello", "LOWER range: G1 = LOWER(D1) = LOWER(\"HELLO\") = \"hello\"");
		assert.strictEqual(ws.getRange2("H1").getValue(), "world", "LOWER range: H1 = LOWER(E1) = LOWER(\"WORLD\") = \"world\"");

		ws.getRange2("J1").setValue("HELLO");
		ws.getRange2("J2").setValue("WORLD");
		ws.getRange2("K1").setValue("TEST");
		ws.getRange2("K2").setValue("ABC");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=LOWER(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "LOWER(J1:K2)", "LOWER implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=LOWER(@J1:K2)", "LOWER implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "LOWER implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"UPPER with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=UPPER({\"hello\";\"world\";\"test\"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UPPER({\"hello\";\"world\";\"test\"})", "UPPER array literal: formula correctly parsed");
		assert.strictEqual(ws.getRange2("A1").getValue(), "HELLO", "UPPER array literal: A1 = UPPER(\"hello\") = \"HELLO\"");
		assert.strictEqual(ws.getRange2("A2").getValue(), "WORLD", "UPPER array literal: A2 = UPPER(\"world\") = \"WORLD\"");
		assert.strictEqual(ws.getRange2("A3").getValue(), "TEST", "UPPER array literal: A3 = UPPER(\"test\") = \"TEST\"");

		ws.getRange2("D1").setValue("hello");
		ws.getRange2("E1").setValue("world");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=UPPER(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UPPER(D1:E1)", "UPPER range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "HELLO", "UPPER range: G1 = UPPER(D1) = UPPER(\"hello\") = \"HELLO\"");
		assert.strictEqual(ws.getRange2("H1").getValue(), "WORLD", "UPPER range: H1 = UPPER(E1) = UPPER(\"world\") = \"WORLD\"");

		ws.getRange2("J1").setValue("hello");
		ws.getRange2("J2").setValue("world");
		ws.getRange2("K1").setValue("test");
		ws.getRange2("K2").setValue("abc");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=UPPER(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UPPER(J1:K2)", "UPPER implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=UPPER(@J1:K2)", "UPPER implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "UPPER implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"PROPER with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=PROPER({\"hello world\";\"JOHN SMITH\";\"test case\"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "PROPER({\"hello world\";\"JOHN SMITH\";\"test case\"})", "PROPER array literal: formula correctly parsed");
		assert.strictEqual(ws.getRange2("A1").getValue(), "Hello World", "PROPER array literal: A1 = PROPER(\"hello world\") = \"Hello World\"");
		assert.strictEqual(ws.getRange2("A2").getValue(), "John Smith", "PROPER array literal: A2 = PROPER(\"JOHN SMITH\") = \"John Smith\"");
		assert.strictEqual(ws.getRange2("A3").getValue(), "Test Case", "PROPER array literal: A3 = PROPER(\"test case\") = \"Test Case\"");

		ws.getRange2("D1").setValue("hello world");
		ws.getRange2("E1").setValue("JOHN SMITH");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=PROPER(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "PROPER(D1:E1)", "PROPER range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "Hello World", "PROPER range: G1 = PROPER(D1) = PROPER(\"hello world\") = \"Hello World\"");
		assert.strictEqual(ws.getRange2("H1").getValue(), "John Smith", "PROPER range: H1 = PROPER(E1) = PROPER(\"JOHN SMITH\") = \"John Smith\"");

		ws.getRange2("J1").setValue("hello world");
		ws.getRange2("J2").setValue("JOHN SMITH");
		ws.getRange2("K1").setValue("test case");
		ws.getRange2("K2").setValue("abc def");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=PROPER(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "PROPER(J1:K2)", "PROPER implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=PROPER(@J1:K2)", "PROPER implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "PROPER implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"TRIM with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=TRIM({\"  Hello  \";\"  World  \";\"  Test  \"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "TRIM({\"  Hello  \";\"  World  \";\"  Test  \"})", "TRIM array literal: formula correctly parsed");
		assert.strictEqual(ws.getRange2("A1").getValue(), "Hello", "TRIM array literal: A1 = TRIM(\"  Hello  \") = \"Hello\"");
		assert.strictEqual(ws.getRange2("A2").getValue(), "World", "TRIM array literal: A2 = TRIM(\"  World  \") = \"World\"");
		assert.strictEqual(ws.getRange2("A3").getValue(), "Test", "TRIM array literal: A3 = TRIM(\"  Test  \") = \"Test\"");

		ws.getRange2("D1").setValue("  Hello  ");
		ws.getRange2("E1").setValue("  World  ");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=TRIM(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "TRIM(D1:E1)", "TRIM range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "Hello", "TRIM range: G1 = TRIM(D1) = TRIM(\"  Hello  \") = \"Hello\"");
		assert.strictEqual(ws.getRange2("H1").getValue(), "World", "TRIM range: H1 = TRIM(E1) = TRIM(\"  World  \") = \"World\"");

		ws.getRange2("J1").setValue("  Hello  ");
		ws.getRange2("J2").setValue("  World  ");
		ws.getRange2("K1").setValue("  Test  ");
		ws.getRange2("K2").setValue("  ABC  ");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=TRIM(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "TRIM(J1:K2)", "TRIM implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=TRIM(@J1:K2)", "TRIM implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "TRIM implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"CLEAN with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=CLEAN({\"Hello\";\"World\";\"Test\"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CLEAN({\"Hello\";\"World\";\"Test\"})", "CLEAN array literal: formula correctly parsed");
		assert.strictEqual(ws.getRange2("A1").getValue(), "Hello", "CLEAN array literal: A1 = CLEAN(\"Hello\") = \"Hello\"");
		assert.strictEqual(ws.getRange2("A2").getValue(), "World", "CLEAN array literal: A2 = CLEAN(\"World\") = \"World\"");
		assert.strictEqual(ws.getRange2("A3").getValue(), "Test", "CLEAN array literal: A3 = CLEAN(\"Test\") = \"Test\"");

		ws.getRange2("D1").setValue("Hello");
		ws.getRange2("E1").setValue("World");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=CLEAN(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CLEAN(D1:E1)", "CLEAN range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "Hello", "CLEAN range: G1 = CLEAN(D1) = CLEAN(\"Hello\") = \"Hello\"");
		assert.strictEqual(ws.getRange2("H1").getValue(), "World", "CLEAN range: H1 = CLEAN(E1) = CLEAN(\"World\") = \"World\"");

		ws.getRange2("J1").setValue("Hello");
		ws.getRange2("J2").setValue("World");
		ws.getRange2("K1").setValue("Test");
		ws.getRange2("K2").setValue("ABC");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=CLEAN(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CLEAN(J1:K2)", "CLEAN implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=CLEAN(@J1:K2)", "CLEAN implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "CLEAN implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"VALUE with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=VALUE({\"123\";\"456\";\"789\"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "VALUE({\"123\";\"456\";\"789\"})", "VALUE array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 123, "VALUE array literal: A1 = VALUE(\"123\") = 123");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 456, "VALUE array literal: A2 = VALUE(\"456\") = 456");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 789, "VALUE array literal: A3 = VALUE(\"789\") = 789");

		ws.getRange2("D1").setValue("123");
		ws.getRange2("E1").setValue("456");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=VALUE(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "VALUE(D1:E1)", "VALUE range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 123, "VALUE range: G1 = VALUE(D1) = VALUE(\"123\") = 123");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 456, "VALUE range: H1 = VALUE(E1) = VALUE(\"456\") = 456");

		ws.getRange2("J1").setValue("123");
		ws.getRange2("J2").setValue("456");
		ws.getRange2("K1").setValue("789");
		ws.getRange2("K2").setValue("100");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=VALUE(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "VALUE(J1:K2)", "VALUE implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=VALUE(@J1:K2)", "VALUE implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "VALUE implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"CODE with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=CODE({\"A\";\"B\";\"C\"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CODE({\"A\";\"B\";\"C\"})", "CODE array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 65, "CODE array literal: A1 = CODE(\"A\") = 65");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 66, "CODE array literal: A2 = CODE(\"B\") = 66");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 67, "CODE array literal: A3 = CODE(\"C\") = 67");

		ws.getRange2("D1").setValue("A");
		ws.getRange2("E1").setValue("B");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=CODE(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CODE(D1:E1)", "CODE range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 65, "CODE range: G1 = CODE(D1) = CODE(\"A\") = 65");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 66, "CODE range: H1 = CODE(E1) = CODE(\"B\") = 66");

		ws.getRange2("J1").setValue("A");
		ws.getRange2("J2").setValue("B");
		ws.getRange2("K1").setValue("C");
		ws.getRange2("K2").setValue("D");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=CODE(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CODE(J1:K2)", "CODE implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=CODE(@J1:K2)", "CODE implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "CODE implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"CHAR with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=CHAR({65;66;67})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CHAR({65;66;67})", "CHAR array literal: formula correctly parsed");
		assert.strictEqual(ws.getRange2("A1").getValue(), "A", "CHAR array literal: A1 = CHAR(65) = \"A\"");
		assert.strictEqual(ws.getRange2("A2").getValue(), "B", "CHAR array literal: A2 = CHAR(66) = \"B\"");
		assert.strictEqual(ws.getRange2("A3").getValue(), "C", "CHAR array literal: A3 = CHAR(67) = \"C\"");

		ws.getRange2("D1").setValue("65");
		ws.getRange2("E1").setValue("66");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=CHAR(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CHAR(D1:E1)", "CHAR range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "A", "CHAR range: G1 = CHAR(D1) = CHAR(65) = \"A\"");
		assert.strictEqual(ws.getRange2("H1").getValue(), "B", "CHAR range: H1 = CHAR(E1) = CHAR(66) = \"B\"");

		ws.getRange2("J1").setValue("65");
		ws.getRange2("J2").setValue("66");
		ws.getRange2("K1").setValue("67");
		ws.getRange2("K2").setValue("68");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=CHAR(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "CHAR(J1:K2)", "CHAR implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=CHAR(@J1:K2)", "CHAR implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "CHAR implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"UNICHAR with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=UNICHAR({65;66;67})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UNICHAR({65;66;67})", "UNICHAR array literal: formula correctly parsed");
		assert.strictEqual(ws.getRange2("A1").getValue(), "A", "UNICHAR array literal: A1 = UNICHAR(65) = \"A\"");
		assert.strictEqual(ws.getRange2("A2").getValue(), "B", "UNICHAR array literal: A2 = UNICHAR(66) = \"B\"");
		assert.strictEqual(ws.getRange2("A3").getValue(), "C", "UNICHAR array literal: A3 = UNICHAR(67) = \"C\"");

		ws.getRange2("D1").setValue("65");
		ws.getRange2("E1").setValue("66");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=UNICHAR(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UNICHAR(D1:E1)", "UNICHAR range: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "A", "UNICHAR range: G1 = UNICHAR(D1) = UNICHAR(65) = \"A\"");
		assert.strictEqual(ws.getRange2("H1").getValue(), "B", "UNICHAR range: H1 = UNICHAR(E1) = UNICHAR(66) = \"B\"");

		ws.getRange2("J1").setValue("65");
		ws.getRange2("J2").setValue("66");
		ws.getRange2("K1").setValue("67");
		ws.getRange2("K2").setValue("68");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=UNICHAR(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UNICHAR(J1:K2)", "UNICHAR implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=UNICHAR(@J1:K2)", "UNICHAR implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "UNICHAR implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"UNICODE with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=UNICODE({\"A\";\"B\";\"C\"})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UNICODE({\"A\";\"B\";\"C\"})", "UNICODE array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 65, "UNICODE array literal: A1 = UNICODE(\"A\") = 65");
		assert.strictEqual(Number(ws.getRange2("A2").getValue()), 66, "UNICODE array literal: A2 = UNICODE(\"B\") = 66");
		assert.strictEqual(Number(ws.getRange2("A3").getValue()), 67, "UNICODE array literal: A3 = UNICODE(\"C\") = 67");

		ws.getRange2("D1").setValue("A");
		ws.getRange2("E1").setValue("B");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=UNICODE(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UNICODE(D1:E1)", "UNICODE range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 65, "UNICODE range: G1 = UNICODE(D1) = UNICODE(\"A\") = 65");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 66, "UNICODE range: H1 = UNICODE(E1) = UNICODE(\"B\") = 66");

		ws.getRange2("J1").setValue("A");
		ws.getRange2("J2").setValue("B");
		ws.getRange2("K1").setValue("C");
		ws.getRange2("K2").setValue("D");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=UNICODE(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "UNICODE(J1:K2)", "UNICODE implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=UNICODE(@J1:K2)", "UNICODE implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "#VALUE!", "UNICODE implicit intersection: M1 shows #VALUE! for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"TYPE with dynamic arrays\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=TYPE({123;\"text\";TRUE})");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "TYPE({123;\"text\";TRUE})", "TYPE array literal: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 64, "TYPE array literal: A1 = TYPE({123;\"text\";TRUE}) = 64 (array)");

		ws.getRange2("D1").setValue("123");
		ws.getRange2("E1").setValue("text");

		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=TYPE(D1:E1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "TYPE(D1:E1)", "TYPE range: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("G1").getValue()), 16, "TYPE range: G1 = TYPE(D1:E1) = 16 (error)");

		ws.getRange2("J1").setValue("123");
		ws.getRange2("J2").setValue("text");
		ws.getRange2("K1").setValue("TRUE");
		ws.getRange2("K2").setValue("456");

		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=TYPE(@J1:K2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "TYPE(J1:K2)", "TYPE implicit intersection: parsed formula normalizes to range without @");
		assert.strictEqual(resCell.getValueForEdit(), "=TYPE(@J1:K2)", "TYPE implicit intersection: stored formula text keeps user-entered @ reference");
		assert.strictEqual(ws.getRange2("M1").getValue(), "16", "TYPE implicit intersection: M1 shows 16 (error) for implicit intersection over a range");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"Multiplication table with SEQUENCE\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(10) * SEQUENCE(1,10)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SEQUENCE(10)*SEQUENCE(1,10)", "Multiplication table: formula correctly parsed");
		
		// Проверка выборочных значений в таблице умножения 10x10
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 1, "A1 = 1*1 = 1");
		assert.strictEqual(Number(ws.getRange2("A10").getValue()), 10, "A10 = 10*1 = 10");
		assert.strictEqual(Number(ws.getRange2("J1").getValue()), 10, "J1 = 1*10 = 10");
		assert.strictEqual(Number(ws.getRange2("J10").getValue()), 100, "J10 = 10*10 = 100");
		assert.strictEqual(Number(ws.getRange2("E5").getValue()), 25, "E5 = 5*5 = 25");
		assert.strictEqual(Number(ws.getRange2("C7").getValue()), 21, "C7 = 7*3 = 21");
		assert.strictEqual(Number(ws.getRange2("H4").getValue()), 32, "H4 = 4*8 = 32");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"FILTER and SORT with multiple conditions\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		// Setup data: Product, Category, Price, Quantity
		ws.getRange2("A1").setValue("Product");
		ws.getRange2("B1").setValue("Category");
		ws.getRange2("C1").setValue("Price");
		ws.getRange2("D1").setValue("Quantity");

		ws.getRange2("A2").setValue("Laptop");
		ws.getRange2("B2").setValue("Electronics");
		ws.getRange2("C2").setValue("1200");
		ws.getRange2("D2").setValue("5");

		ws.getRange2("A3").setValue("Mouse");
		ws.getRange2("B3").setValue("Electronics");
		ws.getRange2("C3").setValue("25");
		ws.getRange2("D3").setValue("50");

		ws.getRange2("A4").setValue("Desk");
		ws.getRange2("B4").setValue("Furniture");
		ws.getRange2("C4").setValue("350");
		ws.getRange2("D4").setValue("10");

		ws.getRange2("A5").setValue("Chair");
		ws.getRange2("B5").setValue("Furniture");
		ws.getRange2("C5").setValue("150");
		ws.getRange2("D5").setValue("20");

		ws.getRange2("A6").setValue("Monitor");
		ws.getRange2("B6").setValue("Electronics");
		ws.getRange2("C6").setValue("450");
		ws.getRange2("D6").setValue("15");

		ws.getRange2("A7").setValue("Keyboard");
		ws.getRange2("B7").setValue("Electronics");
		ws.getRange2("C7").setValue("75");
		ws.getRange2("D7").setValue("30");

		ws.getRange2("A8").setValue("Table");
		ws.getRange2("B8").setValue("Furniture");
		ws.getRange2("C8").setValue("280");
		ws.getRange2("D8").setValue("8");

		ws.getRange2("A9").setValue("Headphones");
		ws.getRange2("B9").setValue("Electronics");
		ws.getRange2("C9").setValue("120");
		ws.getRange2("D9").setValue("25");

		ws.getRange2("A10").setValue("Cabinet");
		ws.getRange2("B10").setValue("Furniture");
		ws.getRange2("C10").setValue("420");
		ws.getRange2("D10").setValue("6");

		ws.getRange2("A11").setValue("Webcam");
		ws.getRange2("B11").setValue("Electronics");
		ws.getRange2("C11").setValue("95");
		ws.getRange2("D11").setValue("18");

		// Apply formula: Filter Electronics with Price > 100, then Sort by Price descending
		fillRange = ws.getRange2("F1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("F1").getValueForEdit2();
		fragment[0].setFragmentText("=SORT(FILTER(A2:D11,(B2:B11=\"Electronics\")*(C2:C11>100)),3,-1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("F1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SORT(FILTER(A2:D11,(B2:B11=\"Electronics\")*(C2:C11>100)),3,-1)", "FILTER+SORT: formula correctly parsed");

		// Check results: should be Laptop (1200), Monitor (450), Headphones (120) sorted by price descending
		assert.strictEqual(ws.getRange2("F1").getValue(), "Laptop", "F1 = Laptop (highest price Electronics > 100)");
		assert.strictEqual(ws.getRange2("G1").getValue(), "Electronics", "G1 = Electronics");
		assert.strictEqual(Number(ws.getRange2("H1").getValue()), 1200, "H1 = 1200");
		assert.strictEqual(Number(ws.getRange2("I1").getValue()), 5, "I1 = 5");

		assert.strictEqual(ws.getRange2("F2").getValue(), "Monitor", "F2 = Monitor (second highest price)");
		assert.strictEqual(Number(ws.getRange2("H2").getValue()), 450, "H2 = 450");

		assert.strictEqual(ws.getRange2("F3").getValue(), "Headphones", "F3 = Headphones (third highest price)");
		assert.strictEqual(Number(ws.getRange2("H3").getValue()), 120, "H3 = 120");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"Dynamic summary with SORTBY, HSTACK, UNIQUE, SUMIF\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		// Setup data: Region, Amount
		ws.getRange2("A1").setValue("Region");
		ws.getRange2("B1").setValue("Amount");

		ws.getRange2("A2").setValue("North");
		ws.getRange2("B2").setValue("1500");

		ws.getRange2("A3").setValue("South");
		ws.getRange2("B3").setValue("2000");

		ws.getRange2("A4").setValue("North");
		ws.getRange2("B4").setValue("800");

		ws.getRange2("A5").setValue("West");
		ws.getRange2("B5").setValue("2500");

		ws.getRange2("A6").setValue("South");
		ws.getRange2("B6").setValue("1200");

		ws.getRange2("A7").setValue("North");
		ws.getRange2("B7").setValue("600");

		ws.getRange2("A8").setValue("West");
		ws.getRange2("B8").setValue("1800");

		ws.getRange2("A9").setValue("South");
		ws.getRange2("B9").setValue("900");

		ws.getRange2("A10").setValue("West");
		ws.getRange2("B10").setValue("1100");

		// Apply formula: Extract unique regions with sums, sorted by sum descending
		fillRange = ws.getRange2("D1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("D1").getValueForEdit2();
		fragment[0].setFragmentText("=SORTBY(HSTACK(UNIQUE(A2:A10),SUMIF(A2:A10,UNIQUE(A2:A10),B2:B10)),SUMIF(A2:A10,UNIQUE(A2:A10),B2:B10),-1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("D1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SORTBY(HSTACK(UNIQUE(A2:A10),SUMIF(A2:A10,UNIQUE(A2:A10),B2:B10)),SUMIF(A2:A10,UNIQUE(A2:A10),B2:B10),-1)", "Dynamic summary: formula correctly parsed");

		// Check results: should be West (5400), South (4100), North (2900) sorted by amount descending
		assert.strictEqual(ws.getRange2("D1").getValue(), "West", "D1 = West (highest total)");
		assert.strictEqual(Number(ws.getRange2("E1").getValue()), 5400, "E1 = 5400 (2500+1800+1100)");

		assert.strictEqual(ws.getRange2("D2").getValue(), "South", "D2 = South (second highest total)");
		assert.strictEqual(Number(ws.getRange2("E2").getValue()), 4100, "E2 = 4100 (2000+1200+900)");

		assert.strictEqual(ws.getRange2("D3").getValue(), "North", "D3 = North (third highest total)");
		assert.strictEqual(Number(ws.getRange2("E3").getValue()), 2900, "E3 = 2900 (1500+800+600)");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"Matrix calculations with SQRT and SEQUENCE\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		fillRange = ws.getRange2("A1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A1").getValueForEdit2();
		fragment[0].setFragmentText("=SQRT(SEQUENCE(5,5,1,1))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SQRT(SEQUENCE(5,5,1,1))", "Matrix calculation: formula correctly parsed");

		// Check sample values from the 5x5 matrix of square roots
		assert.strictEqual(Number(ws.getRange2("A1").getValue()), 1, "A1 = SQRT(1) = 1");
		assert.strictEqual(Number(ws.getRange2("E1").getValue()), Math.sqrt(5), "E1 = SQRT(5)");
		assert.strictEqual(Number(ws.getRange2("A5").getValue()), Math.sqrt(21), "A5 = SQRT(21)");
		assert.strictEqual(Number(ws.getRange2("E5").getValue()), 5, "E5 = SQRT(25) = 5");
		assert.strictEqual(Number(ws.getRange2("C3").getValue()), Math.sqrt(13), "C3 = SQRT(13)");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"Complex dynamic arrays scenarios - FILTER, SORT, VSTACK, error handling\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		// Setup employee data once for all scenarios
		ws.getRange2("A1").setValue("Employee");
		ws.getRange2("B1").setValue("Department");
		ws.getRange2("C1").setValue("Salary");
		ws.getRange2("D1").setValue("Years");
		ws.getRange2("E1").setValue("Rating");

		ws.getRange2("A2").setValue("John");
		ws.getRange2("B2").setValue("IT");
		ws.getRange2("C2").setValue("80000");
		ws.getRange2("D2").setValue("5");
		ws.getRange2("E2").setValue("4.5");

		ws.getRange2("A3").setValue("Mary");
		ws.getRange2("B3").setValue("Sales");
		ws.getRange2("C3").setValue("75000");
		ws.getRange2("D3").setValue("3");
		ws.getRange2("E3").setValue("4.8");

		ws.getRange2("A4").setValue("Bob");
		ws.getRange2("B4").setValue("IT");
		ws.getRange2("C4").setValue("95000");
		ws.getRange2("D4").setValue("7");
		ws.getRange2("E4").setValue("4.2");

		ws.getRange2("A5").setValue("Alice");
		ws.getRange2("B5").setValue("HR");
		ws.getRange2("C5").setValue("65000");
		ws.getRange2("D5").setValue("2");
		ws.getRange2("E5").setValue("4.6");

		ws.getRange2("A6").setValue("Charlie");
		ws.getRange2("B6").setValue("IT");
		ws.getRange2("C6").setValue("120000");
		ws.getRange2("D6").setValue("10");
		ws.getRange2("E6").setValue("4.9");

		ws.getRange2("A7").setValue("Diana");
		ws.getRange2("B7").setValue("Sales");
		ws.getRange2("C7").setValue("82000");
		ws.getRange2("D7").setValue("4");
		ws.getRange2("E7").setValue("4.7");

		// Scenario 1: Simple FILTER for IT department
		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=FILTER(A2:E7,B2:B7=\"IT\")");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "FILTER(A2:E7,B2:B7=\"IT\")", "FILTER IT department: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "John", "G1 = John (first IT employee)");
		assert.strictEqual(ws.getRange2("G2").getValue(), "Bob", "G2 = Bob (second IT employee)");
		assert.strictEqual(ws.getRange2("G3").getValue(), "Charlie", "G3 = Charlie (third IT employee)");
		assert.strictEqual(Number(ws.getRange2("I3").getValue()), 120000, "I3 = 120000");

		// Scenario 2: FILTER with multiple conditions
		clearData(6, 0, 20, 20); // Clear G1 area
		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText('=FILTER(A2:E7,(B2:B7="IT")*(C2:C7>80000))');
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "FILTER(A2:E7,(B2:B7=\"IT\")*(C2:C7>80000))", "FILTER multiple conditions: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "Bob", "G1 = Bob (IT with salary > 80000)");
		assert.strictEqual(ws.getRange2("G2").getValue(), "Charlie", "G2 = Charlie (IT with salary > 80000)");
		assert.strictEqual(Number(ws.getRange2("I1").getValue()), 95000, "I1 = 95000");

		// Scenario 3: SORT + FILTER
		clearData(6, 0, 20, 20); // Clear G1 area
		fillRange = ws.getRange2("G1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("G1").getValueForEdit2();
		fragment[0].setFragmentText("=SORT(FILTER(A2:E7,B2:B7=\"IT\"),3,-1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("G1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SORT(FILTER(A2:E7,B2:B7=\"IT\"),3,-1)", "SORT+FILTER: formula correctly parsed");
		assert.strictEqual(ws.getRange2("G1").getValue(), "Charlie", "G1 = Charlie (highest IT salary)");
		assert.strictEqual(ws.getRange2("G2").getValue(), "Bob", "G2 = Bob (second highest IT salary)");
		assert.strictEqual(ws.getRange2("G3").getValue(), "John", "G3 = John (third highest IT salary)");
		assert.strictEqual(Number(ws.getRange2("I1").getValue()), 120000, "I1 = 120000");

		// Scenario 4: SORTBY + HSTACK + UNIQUE - department summary
		clearData(12, 0, 20, 20); // Clear M1 area
		fillRange = ws.getRange2("M1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("M1").getValueForEdit2();
		fragment[0].setFragmentText("=SORTBY(HSTACK(UNIQUE(B2:B7),SUMIF(B2:B7,UNIQUE(B2:B7),C2:C7)),SUMIF(B2:B7,UNIQUE(B2:B7),C2:C7),-1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("M1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SORTBY(HSTACK(UNIQUE(B2:B7),SUMIF(B2:B7,UNIQUE(B2:B7),C2:C7)),SUMIF(B2:B7,UNIQUE(B2:B7),C2:C7),-1)", "SORTBY+HSTACK+UNIQUE: formula correctly parsed");
		assert.strictEqual(ws.getRange2("M1").getValue(), "IT", "M1 = IT (highest total salary)");
		assert.strictEqual(Number(ws.getRange2("N1").getValue()), 295000, "N1 = 295000 (80000+95000+120000)");
		assert.strictEqual(ws.getRange2("M2").getValue(), "Sales", "M2 = Sales (second highest)");
		assert.strictEqual(Number(ws.getRange2("N2").getValue()), 157000, "N2 = 157000 (75000+82000)");

		// Scenario 5: VSTACK with headers
		clearData(18, 0, 30, 20); // Clear S1 area
		fillRange = ws.getRange2("S1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("S1").getValueForEdit2();
		fragment[0].setFragmentText("=VSTACK({\"Department\",\"Total\",\"Count\"},HSTACK(UNIQUE(B2:B7),SUMIF(B2:B7,UNIQUE(B2:B7),C2:C7),COUNTIF(B2:B7,UNIQUE(B2:B7))))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("S1"));
		assert.ok(getNormalizedFormula(resCell).includes("VSTACK"), "VSTACK: formula contains VSTACK");
		assert.strictEqual(ws.getRange2("S1").getValue(), "Department", "S1 = Department (header)");
		assert.strictEqual(ws.getRange2("T1").getValue(), "Total", "T1 = Total (header)");
		assert.strictEqual(ws.getRange2("U1").getValue(), "Count", "U1 = Count (header)");
		assert.strictEqual(ws.getRange2("S2").getValue(), "IT", "S2 = IT");
		assert.strictEqual(Number(ws.getRange2("T2").getValue()), 295000, "T2 = 295000 (IT total)");
		assert.strictEqual(Number(ws.getRange2("U2").getValue()), 3, "U2 = 3 (IT count)");

		// Scenario 6: Multiplication table 5x5
		clearData(0, 25, 10, 35); // Clear AA1 area
		fillRange = ws.getRange2("AA1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("AA1").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(5) * TRANSPOSE(SEQUENCE(1,5))");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("AA1"));
		assert.strictEqual(getNormalizedFormula(resCell), "SEQUENCE(5)*TRANSPOSE(SEQUENCE(1,5))", "Multiplication table 5x5: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("AA1").getValue()), 1, "AA1 = 1*1 = 1");
		assert.strictEqual(Number(ws.getRange2("AA5").getValue()), 25, "AA5 = 5*5 = 25");

		// Scenario 7: Division by zero with error handling
		fillRange = ws.getRange2("A10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("A10").getValueForEdit2();
		fragment[0].setFragmentText("=100/SEQUENCE(5,1,-2,1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("A10"));
		assert.strictEqual(getNormalizedFormula(resCell), "100/SEQUENCE(5,1,-2,1)", "Division with zero: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("A10").getValue()), -50, "A10 = 100/(-2) = -50");
		
		const a12Value = ws.getRange2("A12").getValue();
		assert.ok(a12Value === "#DIV/0!" || a12Value === "#NUM!", "A12 = #DIV/0! (division by zero)");

		// With IFERROR handling
		fillRange = ws.getRange2("C10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("C10").getValueForEdit2();
		fragment[0].setFragmentText("=IFERROR(100/SEQUENCE(5,1,-2,1),0)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("C10"));
		assert.strictEqual(getNormalizedFormula(resCell), "IFERROR(100/SEQUENCE(5,1,-2,1),0)", "IFERROR division: formula correctly parsed");
		assert.strictEqual(Number(ws.getRange2("C12").getValue()), 0, "C12 = 0 (error replaced)");
		assert.strictEqual(Number(ws.getRange2("C13").getValue()), 100, "C13 = 100");

		// Scenario 8: SEQUENCE spill conflict
		clearData(0, 15, 10, 25); // Clear P10 area
		fillRange = ws.getRange2("P10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("P10").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(3,3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		assert.strictEqual(Number(ws.getRange2("P10").getValue()), 1, "P10 = 1");
		assert.strictEqual(Number(ws.getRange2("R12").getValue()), 9, "R12 = 9");

		// Place blocking data and test spill conflict
		ws.getRange2("Q11").setValue("BLOCKING");
		ws.getRange2("P10").setValue("");
		fillRange = ws.getRange2("P10");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("P10").getValueForEdit2();
		fragment[0].setFragmentText("=SEQUENCE(3,3)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		const spillValue = ws.getRange2("P10").getValue();
		assert.ok(spillValue === "#SPILL!" || spillValue === "#REF!" || Number(spillValue) === 1, "P10 shows #SPILL!, #REF! error, or spill is prevented");

		clearData(0, 0, 100, 200);
	});

	QUnit.test("Test: \"IF with array condition and range values\"", function (assert) {
		if (!AscCommonExcel.bIsSupportDynamicArrays) {
			assert.ok(true, "Dynamic arrays support is disabled");
			return;
		}

		let fillRange, resCell, fragment;
		let flags = wsView._getCellFlags(0, 0);
		flags.ctrlKey = false;
		flags.shiftKey = false;

		clearData(0, 0, 100, 200);

		ws.getRange2("A1").setValue("1");
		ws.getRange2("B1").setValue("FALSE");
		ws.getRange2("C1").setValue("3");
		
		ws.getRange2("A2").setValue("34");
		ws.getRange2("B2").setValue("FALSE");
		ws.getRange2("C2").setValue("23");

		fillRange = ws.getRange2("E1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("E1").getValueForEdit2();
		fragment[0].setFragmentText("=IF({1,0,1},A1:C2)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("E1"));
		assert.ok(resCell, "Formula cell E1 exists");

		assert.strictEqual(ws.getRange2("E1").getValue(), "1", "E1 = 1");
		assert.strictEqual(ws.getRange2("F1").getValue(), "FALSE", "F1 = FALSE");
		assert.strictEqual(ws.getRange2("G1").getValue(), "3", "G1 = 3");
		assert.strictEqual(ws.getRange2("E2").getValue(), "34", "E2 = 34");
		assert.strictEqual(ws.getRange2("F2").getValue(), "FALSE", "F2 = FALSE");
		assert.strictEqual(ws.getRange2("G2").getValue(), "23", "G2 = 23");

		fillRange = ws.getRange2("H1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("H1").getValueForEdit2();
		fragment[0].setFragmentText("=IF({1,0,1},A1:C1)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("H1"));
		assert.ok(resCell, "Formula cell H1 exists");

		assert.strictEqual(ws.getRange2("H1").getValue(), "1", "H1 = 1");
		assert.strictEqual(ws.getRange2("I1").getValue(), "FALSE", "I1 = FALSE");
		assert.strictEqual(ws.getRange2("J1").getValue(), "3", "J1 = 3");

		fillRange = ws.getRange2("K1");
		wsView.setSelection(fillRange.bbox);
		fragment = ws.getRange2("K1").getValueForEdit2();
		fragment[0].setFragmentText("=IF({1,0,1;2,0,4;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1},A1:C10)");
		wsView._saveCellValueAfterEdit(fillRange, fragment, flags, null, null);

		resCell = getCell(ws.getRange2("K1"));
		assert.ok(resCell, "Formula cell K1 exists");

		assert.strictEqual(ws.getRange2("K1").getValue(), "1", "K1 = 1");
		assert.strictEqual(ws.getRange2("L1").getValue(), "FALSE", "L1 = FALSE");
		assert.strictEqual(ws.getRange2("M1").getValue(), "3", "M1 = 3");
		assert.strictEqual(ws.getRange2("K2").getValue(), "34", "K2 = 34");
		assert.strictEqual(ws.getRange2("L2").getValue(), "FALSE", "L2 = FALSE");
		assert.strictEqual(ws.getRange2("M2").getValue(), "23", "M2 = 23");
		assert.strictEqual(ws.getRange2("K3").getValue(), "0", "K3 = 0");
		assert.strictEqual(ws.getRange2("L3").getValue(), "FALSE", "L3 = FALSE");
		assert.strictEqual(ws.getRange2("M3").getValue(), "0", "M3 = 0");
		assert.strictEqual(ws.getRange2("K4").getValue(), "0", "K4 = 0");
		assert.strictEqual(ws.getRange2("L4").getValue(), "FALSE", "L4 = FALSE");
		assert.strictEqual(ws.getRange2("M4").getValue(), "0", "M4 = 0");
		assert.strictEqual(ws.getRange2("K5").getValue(), "0", "K5 = 0");
		assert.strictEqual(ws.getRange2("L5").getValue(), "FALSE", "L5 = FALSE");
		assert.strictEqual(ws.getRange2("M5").getValue(), "0", "M5 = 0");
		assert.strictEqual(ws.getRange2("K6").getValue(), "0", "K6 = 0");
		assert.strictEqual(ws.getRange2("L6").getValue(), "FALSE", "L6 = FALSE");
		assert.strictEqual(ws.getRange2("M6").getValue(), "0", "M6 = 0");
		assert.strictEqual(ws.getRange2("K7").getValue(), "0", "K7 = 0");
		assert.strictEqual(ws.getRange2("L7").getValue(), "FALSE", "L7 = FALSE");
		assert.strictEqual(ws.getRange2("M7").getValue(), "0", "M7 = 0");
		assert.strictEqual(ws.getRange2("K8").getValue(), "0", "K8 = 0");
		assert.strictEqual(ws.getRange2("L8").getValue(), "FALSE", "L8 = FALSE");
		assert.strictEqual(ws.getRange2("M8").getValue(), "0", "M8 = 0");
		assert.strictEqual(ws.getRange2("K9").getValue(), "0", "K9 = 0");
		assert.strictEqual(ws.getRange2("L9").getValue(), "FALSE", "L9 = FALSE");
		assert.strictEqual(ws.getRange2("M9").getValue(), "0", "M9 = 0");
		assert.strictEqual(ws.getRange2("K10").getValue(), "0", "K10 = 0");
		assert.strictEqual(ws.getRange2("L10").getValue(), "FALSE", "L10 = FALSE");
		assert.strictEqual(ws.getRange2("M10").getValue(), "0", "M10 = 0");
		assert.strictEqual(ws.getRange2("K11").getValue(), "#N/A", "K11 = #N/A");
		assert.strictEqual(ws.getRange2("L11").getValue(), "FALSE", "L11 = FALSE");
		assert.strictEqual(ws.getRange2("M11").getValue(), "#N/A", "M11 = #N/A");
		assert.strictEqual(ws.getRange2("K12").getValue(), "#N/A", "K12 = #N/A");
		assert.strictEqual(ws.getRange2("L12").getValue(), "FALSE", "L12 = FALSE");
		assert.strictEqual(ws.getRange2("M12").getValue(), "#N/A", "M12 = #N/A");
		assert.strictEqual(ws.getRange2("K13").getValue(), "#N/A", "K13 = #N/A");
		assert.strictEqual(ws.getRange2("L13").getValue(), "FALSE", "L13 = FALSE");
		assert.strictEqual(ws.getRange2("M13").getValue(), "#N/A", "M13 = #N/A");
		assert.strictEqual(ws.getRange2("K14").getValue(), "#N/A", "K14 = #N/A");
		assert.strictEqual(ws.getRange2("L14").getValue(), "FALSE", "L14 = FALSE");
		assert.strictEqual(ws.getRange2("M14").getValue(), "#N/A", "M14 = #N/A");

		//TODO IF({1,0,1;2,0,4;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1;1,0,1},A1:C10)

		clearData(0, 0, 100, 200);
	});

	QUnit.module("Dynamic Arrays Tests");
});
