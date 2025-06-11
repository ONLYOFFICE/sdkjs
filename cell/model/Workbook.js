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

"use strict";

(function(window, undefined){

// Import
	var g_memory = AscFonts.g_memory;

	var CellValueType = AscCommon.CellValueType;
	var c_oAscBorderStyles = Asc.c_oAscBorderStyles;
	var fSortAscending = AscCommon.fSortAscending;
	var parserHelp = AscCommon.parserHelp;
	var oNumFormatCache = AscCommon.oNumFormatCache;
	var gc_nMaxRow0 = AscCommon.gc_nMaxRow0;
	var gc_nMaxCol0 = AscCommon.gc_nMaxCol0;
	var g_oCellAddressUtils = AscCommon.g_oCellAddressUtils;
	var CellAddress = AscCommon.CellAddress;
	var isRealObject = AscCommon.isRealObject;
	var cBoolLocal = AscCommon.cBoolLocal;
	var cErrorLocal = AscCommon.cErrorLocal;
	var cErrorOrigin = AscCommon.cErrorOrigin;
	var c_oAscNumFormatType = Asc.c_oAscNumFormatType;
	var oSeriesInType = Asc.c_oAscSeriesInType;
	var oSeriesType = Asc.c_oAscSeriesType;
	var oSeriesDateUnitType = Asc.c_oAscDateUnitType;
	var oFillType = Asc.c_oAscFillType;


	var UndoRedoItemSerializable = AscCommonExcel.UndoRedoItemSerializable;
	var UndoRedoData_CellSimpleData = AscCommonExcel.UndoRedoData_CellSimpleData;
	var UndoRedoData_CellValueData = AscCommonExcel.UndoRedoData_CellValueData;
	var UndoRedoData_FromToRowCol = AscCommonExcel.UndoRedoData_FromToRowCol;
	var UndoRedoData_FromTo = AscCommonExcel.UndoRedoData_FromTo;
	var UndoRedoData_IndexSimpleProp = AscCommonExcel.UndoRedoData_IndexSimpleProp;
	var UndoRedoData_BBox = AscCommonExcel.UndoRedoData_BBox;
	var UndoRedoData_SheetAdd = AscCommonExcel.UndoRedoData_SheetAdd;
	var UndoRedoData_DefinedNames = AscCommonExcel.UndoRedoData_DefinedNames;
	var g_oDefaultFormat = AscCommonExcel.g_oDefaultFormat;
	var g_StyleCache = AscCommonExcel.g_StyleCache;
	var Border = AscCommonExcel.Border;
	var RangeDataManager = AscCommonExcel.RangeDataManager;

	var cElementType = AscCommonExcel.cElementType;

	var parserFormula = AscCommonExcel.parserFormula;

	var c_oAscError = Asc.c_oAscError;
	var c_oAscInsertOptions = Asc.c_oAscInsertOptions;
	var c_oAscDeleteOptions = Asc.c_oAscDeleteOptions;
	var c_oAscGetDefinedNamesList = Asc.c_oAscGetDefinedNamesList;
	var c_oAscDefinedNameReason = Asc.c_oAscDefinedNameReason;
	var c_oNotifyType = AscCommon.c_oNotifyType;
	var g_cCalcRecursion = AscCommonExcel.g_cCalcRecursion;

	var cError = AscCommonExcel.cError;
	var cErrorType = AscCommonExcel.cErrorType;

	var g_nVerticalTextAngle = 255;
	//определяется в WorksheetView.js
	var oDefaultMetrics = {
		ColWidthChars: 0,
		RowHeight: 0
	};
	var g_sNewSheetNamePattern = "Sheet";
	var g_nSheetNameMaxLength = 31;
	var g_nDefNameMaxLength = 255;
	var g_nAllColIndex = -1;
	var g_nAllRowIndex = -1;
	var aStandartNumFormats = [];
	var aStandartNumFormatsId = {};
	var oFormulaLocaleInfo = {
		Parse: true,
		DigitSep: true
	};

	(function(){
		aStandartNumFormats[0] = "General";
		aStandartNumFormats[1] = "0";
		aStandartNumFormats[2] = "0.00";
		aStandartNumFormats[3] = "#,##0";
		aStandartNumFormats[4] = "#,##0.00";
		aStandartNumFormats[9] = "0%";
		aStandartNumFormats[10] = "0.00%";
		aStandartNumFormats[11] = "0.00E+00";
		aStandartNumFormats[12] = "# ?/?";
		aStandartNumFormats[13] = "# ??/??";
		aStandartNumFormats[14] = "m/d/yyyy";
		aStandartNumFormats[15] = "d-mmm-yy";
		aStandartNumFormats[16] = "d-mmm";
		aStandartNumFormats[17] = "mmm-yy";
		aStandartNumFormats[18] = "h:mm AM/PM";
		aStandartNumFormats[19] = "h:mm:ss AM/PM";
		aStandartNumFormats[20] = "h:mm";
		aStandartNumFormats[21] = "h:mm:ss";
		aStandartNumFormats[22] = "m/d/yyyy h:mm";
		aStandartNumFormats[37] = "#,##0_);(#,##0)";
		aStandartNumFormats[38] = "#,##0_);[Red](#,##0)";
		aStandartNumFormats[39] = "#,##0.00_);(#,##0.00)";
		aStandartNumFormats[40] = "#,##0.00_);[Red](#,##0.00)";
		aStandartNumFormats[45] = "mm:ss";
		aStandartNumFormats[46] = "[h]:mm:ss";
		aStandartNumFormats[47] = "mm:ss.0";
		aStandartNumFormats[48] = "##0.0E+0";
		aStandartNumFormats[49] = "@";
		for(var i in aStandartNumFormats)
		{
			aStandartNumFormatsId[aStandartNumFormats[i]] = i - 0;
		}
	})();

	var c_oRangeType =
		{
			Range:0,
			Col:1,
			Row:2,
			All:3
		};
	var c_oSharedShiftType = {
		Processed: 1,
		NeedTransform: 2,
		PreProcessed: 3
	};



	var emptyStyleComponents = {table: [], conditional: []};
	function getRangeType(oBBox){
		if(null == oBBox)
			oBBox = this.bbox;
		if(oBBox.c1 == 0 && gc_nMaxCol0 == oBBox.c2 && oBBox.r1 == 0 && gc_nMaxRow0 == oBBox.r2)
			return c_oRangeType.All;
		if(oBBox.c1 == 0 && gc_nMaxCol0 == oBBox.c2)
			return c_oRangeType.Row;
		else if(oBBox.r1 == 0 && gc_nMaxRow0 == oBBox.r2)
			return c_oRangeType.Col;
		else
			return c_oRangeType.Range;
	}

	function getCompiledStyleFromArray(xf, xfs, isTableBorders) {
		for (var i = 0; i < xfs.length; ++i) {
			if (null == xf) {
				xf = xfs[i];
			} else {
				xf = xf.merge(xfs[i], undefined, isTableBorders);
			}
		}
		return xf;
	}
	function getCompiledStyle(sheetMergedStyles, hiddenManager, nRow, nCol, opt_cell, opt_ws, opt_styleComponents, opt_AffectingText) {
		var styleComponents = opt_styleComponents ? opt_styleComponents : sheetMergedStyles.getStyle(hiddenManager, nRow, nCol, opt_ws, opt_AffectingText);
		var xf = getCompiledStyleFromArray(null, styleComponents.table, true);
		if (opt_cell) {
			if (null === xf) {
				xf = opt_cell.xfs;
			} else if (opt_cell.xfs) {
				xf = xf.merge(opt_cell.xfs, true);
			}
		} else if (opt_ws) {
			opt_ws._getRowNoEmpty(nRow, function(row){
				if(row && null != row.xfs){
					xf = null === xf ? row.xfs : xf.merge(row.xfs, true);
				} else {
					var col = opt_ws._getColNoEmptyWithAll(nCol);
					if(null != col && null != col.xfs){
						xf = null === xf ? col.xfs : xf.merge(col.xfs, true);
					}
				}
			});

		}
		xf = getCompiledStyleFromArray(xf, styleComponents.conditional);
		return xf;
	}

	function getDefNameIndex(name) {
		//uniqueness is checked without capitalization
		return name ? name.toLowerCase() : name;
	}

	function getDefNameId(sheetId, name) {
		if (sheetId) {
			return sheetId + AscCommon.g_cCharDelimiter + getDefNameIndex(name);
		} else {
			return getDefNameIndex(name);
		}
	}

	var g_FDNI = {sheetId: null, name: null};

	function getFromDefNameId(nodeId) {
		var index = nodeId ? nodeId.indexOf(AscCommon.g_cCharDelimiter) : -1;
		if (-1 != index) {
			g_FDNI.sheetId = nodeId.substring(0, index);
			g_FDNI.name = nodeId.substring(index + 1);
		} else {
			g_FDNI.sheetId = null;
			g_FDNI.name = nodeId;
		}
	}

	function changeTextCase(fragments, type, opt_start, opt_end) {
		if (!fragments) {
			return;
		}

		let isChange = false;
		let newText = "", fragmentsMap;
		let c_oType = Asc.c_oAscChangeTextCaseType;

		let getChangedTextSimpleCase = function (_text, _text_pos) {
			if (!_text) {
				return _text;
			}

			let _newText = "";
			let _textBefore = "";
			let _textAfter = "";
			if (opt_start || opt_end) {
				if (_text_pos + _text.length < opt_start || _text_pos > opt_end) {
					return _text;
				}
				let _start = null, _end = null;
				if (_text_pos < opt_start) {
					_textBefore = _text.substring(0, opt_start - _text_pos);
					_start = opt_start - _text_pos;
				}
				if (_text_pos + _text.length > opt_end) {
					_textAfter = _text.substring(opt_end - _text_pos, _text.length);
					_end = opt_end - _text_pos;
				}
				if (_start || _end) {
					_text = _text.substring(_start ? _start : 0, _end ? _end : _text.length);
				}

			}
			switch (type) {
				case c_oType.LowerCase: {
					_newText = _text.toLowerCase();
					break;
				}
				case c_oType.UpperCase: {
					_newText = _text.toUpperCase();
					break;
				}
				case c_oType.ToggleCase: {
					for (let i = 0, length = _text.length; i < length; ++i) {
						if (_text[i].toUpperCase() === _text[i]) {
							_newText += _text[i].toLowerCase();
						} else {
							_newText += _text[i].toUpperCase();
						}
					}

					break;
				}
			}
			return _textBefore + _newText + _textAfter;
		};

		if (type === c_oType.LowerCase || type === c_oType.UpperCase || type === c_oType.ToggleCase) {
			let curTextLength = 0;
			for (let m = 0; m < fragments.length; m++) {
				let newFragmentText = getChangedTextSimpleCase(fragments[m].text, curTextLength);
				if (newFragmentText !== fragments[m].text) {
					isChange = true;
					if (fragments[m].setText) {
						fragments[m].setText(newFragmentText);
					} else if (fragments[m].setFragmentText) {
						if (!fragmentsMap) {
							fragmentsMap = {};
						}
						fragmentsMap[m] = fragments[m].clone();
						fragmentsMap[m].setFragmentText(newFragmentText);
					}
				}
				newText += newFragmentText;
				curTextLength += fragments[m].text.length;
			}
		} else {

			let getParagraphs = function (_fragments) {
				let res = [];

				AscFormat.ExecuteNoHistory(function () {
					let oCurPar = null;
					let oCurRun = null;

					let _setSelection = function (_run, pos) {
						if (!_run || !_run.Content) {
							return;
						}
						let startRun = pos;
						let run_length = _run.Content.length;
						let endRun = pos + run_length;
						if (opt_start <= startRun && endRun <= opt_end) {
							//select all
							_run.State.Selection.StartPos = 0;
							_run.State.Selection.EndPos = run_length;
							_run.State.Selection.Use = true;
						} else if (endRun > opt_start && startRun < opt_end) {
							_run.State.Selection.StartPos = opt_start > startRun ? (opt_start - startRun) : 0;
							_run.State.Selection.EndPos = (endRun > opt_end) ? (opt_end - curTextLength) : run_length;
							_run.State.Selection.Use = true;
						}
					};

					let pushParagraph = function () {
						if (!oCurPar || !oCurRun) {
							return;
						}

						oCurPar.Internal_Content_Add(0, oCurRun, false);

						if (opt_start || opt_end) {
							_setSelection(oCurRun, curTextLength);
						} else {
							oCurPar.SelectAll();
						}

						curTextLength += oCurRun.Content.length;

						res.push(oCurPar);
						oCurPar = null;
						oCurRun = null;
					};

					let curTextLength = 0;
					for (let m = 0; m < _fragments.length; m++) {
						let curMultiText = _fragments[m].text;
						for (let k = 0, length = curMultiText.length; k < length; k++) {
							if (oCurPar === null) {
								oCurPar = new AscWord.Paragraph();
								oCurRun = new ParaRun(oCurPar);
							}

							let nCharCode = curMultiText.charCodeAt(k);
							if (curMultiText[k] === "\n") {
								pushParagraph();
								curTextLength++;
								continue;
							}

							let nUnicode = null;
							if (AscCommon.isLeadingSurrogateChar(nCharCode)) {
								if (k + 1 < length) {
									k++;
									let nTrailingChar = curMultiText.charCodeAt(k);
									nUnicode = AscCommon.decodeSurrogateChar(nCharCode, nTrailingChar);
								}
							} else {
								nUnicode = nCharCode;
							}

							if (null != nUnicode) {
								let Item;
								if (0x20 !== nUnicode && 0xA0 !== nUnicode && 0x2009 !== nUnicode) {
									Item = new AscWord.CRunText(nUnicode);
								} else {
									Item = new AscWord.CRunSpace();
								}
								//add text
								oCurRun.Add_ToContent(-1, Item, false);
							}
						}
					}
					if (oCurPar) {
						pushParagraph();
					}

				}, this, []);

				return res;
			};

			let aParagraphs = getParagraphs(fragments);
			if (aParagraphs && aParagraphs.length) {
				let oChangeEngine = new AscCommonWord.CChangeTextCaseEngine(type);
				oChangeEngine.ProcessParagraphs(aParagraphs);

				for (let i = 0; i < aParagraphs.length; i++) {
					newText += aParagraphs[i].GetText({ParaSeparator : ""});
					if (i !== aParagraphs.length - 1) {
						newText += "\n";
					}
				}

				if (newText) {
					let counter = 0;
					for (let m = 0; m < fragments.length; m++) {
						let curMultiText = fragments[m].text;
						let isChangeFragment = false;
						let newFragmentText = "";
						for (let k = 0, length = curMultiText.length; k < length; k++) {
							if (curMultiText[k] !== newText[counter]) {
								isChange = true;
								isChangeFragment = true;
							}
							newFragmentText += newText[counter];
							counter++;
						}
						if (isChangeFragment) {
							if (fragments[m].setText) {
								//cell value
								fragments[m].setText(newFragmentText);
							} else if (fragments[m].setFragmentText) {
								//cell editor
								if (!fragmentsMap) {
									fragmentsMap = {};
								}
								fragmentsMap[m] = fragments[m].clone();
								fragmentsMap[m].setFragmentText(newFragmentText);
							}
						}
					}
				}
			}
		}
		return isChange ? {text: newText, fragmentsMap: fragmentsMap} : null;
	}

	function DefName(wb, name, ref, sheetId, hidden, type, isXLNM) {
		this.wb = wb;
		this.name = name;
		this.ref = ref;
		this.sheetId = sheetId;
		this.hidden = hidden;
		this.type = type;

		this.isXLNM = isXLNM;

		this.isLock = null;
		this.parsedRef = null;
	}

	DefName.prototype = {
		clone: function(wb){
			return new DefName(wb, this.name, this.ref, this.sheetId, this.hidden, this.type, this.isXLNM);
		},
		removeDependencies: function() {
			if (this.parsedRef) {
				this.parsedRef.removeDependencies();
				this.parsedRef = null;
			}
		},
		setRef: function(ref, opt_noRemoveDependencies, opt_forceBuild, opt_open) {
			if(!opt_noRemoveDependencies){
				this.removeDependencies();
			}
			//для R1C1: ref - всегда строка в виде A1B1
			//флаг opt_open - на открытие, undo/redo, принятие изменений - строка приходит в виде A1B1 - преобразовывать не нужно
			//во всех остальных случаях парсим ref и заменяем на формат A1B1
			opt_open = opt_open || this.wb.bRedoChanges || this.wb.bUndoChanges;

			this.ref = ref;
			//all ref should be 3d, so worksheet can be anyone
			this.parsedRef = new parserFormula(ref, this, AscCommonExcel.g_DefNameWorksheet);
			this.parsedRef.setIsTable(this.type);
			var t = this;
			if (opt_forceBuild) {
				if(opt_open) {
					AscCommonExcel.executeInR1C1Mode(false, function () {
						t.parsedRef.parse();
					});
				} else {
					this.parsedRef.parse();
				}
				this.parsedRef.buildDependencies();
			} else {
				if(!opt_open) {
					this.parsedRef.parse();
					this.ref = this.parsedRef.assemble();
				}
				this.wb.dependencyFormulas.addToBuildDependencyDefName(this);
			}
		},
		getRef: function(bLocale) {
			//R1C1 - отдаём в зависимости от флага bLocale(для меню в виде R1C1)
			var res, t = this;
			if(!this.parsedRef.isParsed) {
				AscCommonExcel.executeInR1C1Mode(false, function () {
					t.parsedRef.parse();
				});
			}
			if(bLocale) {
				res = this.parsedRef.assembleLocale(AscCommonExcel.cFormulaFunctionToLocale, true);
			} else {
				res = this.parsedRef.assemble();
			}
			return res;
		},
		getNodeId: function() {
			return getDefNameId(this.sheetId, this.name);
		},
		getAscCDefName: function(bLocale) {
			var index = null;
			if (this.sheetId) {
				var sheet = this.wb.getWorksheetById(this.sheetId);
				index = sheet.getIndex();
			}
			//теперь тип используется ещё и при получении результата вычисления именованного диапазона
			//так же заполняю при отркытии
			//TODO - проверить, возможно необходимо убрать
			if (!this.type && this.wb && this.wb.getSlicerCacheByName(this.name)) {
				this.type = Asc.c_oAscDefNameType.slicer;
			}
			return new Asc.asc_CDefName(this.name, this.getRef(bLocale), index, this.type, this.hidden, this.isLock, this.isXLNM);
		},
		getUndoDefName: function() {
			return new UndoRedoData_DefinedNames(this.name, this.ref, this.sheetId, this.type, this.isXLNM);
		},
		setUndoDefName: function(newUndoName, doNotChangeRef) {
			this.name = newUndoName.name;
			this.sheetId = newUndoName.sheetId;
			this.hidden = newUndoName.hidden;
			this.type = newUndoName.type;
			if (!doNotChangeRef && this.ref != newUndoName.ref) {
				this.setRef(newUndoName.ref);
			}
			this.isXLNM = newUndoName.isXLNM;
		},
		onFormulaEvent: function (type, eventData) {
			if (AscCommon.c_oNotifyParentType.IsDefName === type) {
				return null;
			} else if (AscCommon.c_oNotifyParentType.Change === type) {
				this.wb.dependencyFormulas.addToChangedDefName(this);
			} else if (AscCommon.c_oNotifyParentType.ChangeFormula === type) {
				var notifyType = eventData.notifyData.type;
				if (!(this.type === Asc.c_oAscDefNameType.table && (c_oNotifyType.Shift === notifyType || c_oNotifyType.Move === notifyType || c_oNotifyType.Delete === notifyType))) {
					var oldUndoName = this.getUndoDefName();
					this.parsedRef.setFormulaString(this.ref = eventData.assemble);
					this.wb.dependencyFormulas.addToChangedDefName(this);
					var newUndoName = this.getUndoDefName();
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_DefinedNamesChangeUndo,
						null, null, new UndoRedoData_FromTo(oldUndoName, newUndoName), true);
				}
			}
		}
	};

	function getCellIndex(row, col) {
		return row * AscCommon.gc_nMaxCol + col;
	}

	var g_FCI = {row: null, col: null};

	function getFromCellIndex(cellIndex, returnDuplicate) {
		g_FCI.row = Math.floor(cellIndex / AscCommon.gc_nMaxCol);
		g_FCI.col = cellIndex % AscCommon.gc_nMaxCol;
		return returnDuplicate ? {row: g_FCI.row, col: g_FCI.col} : null;
	}

	function getVertexIndex(bbox) {
		//without $
		//значения в areaMap хранятся в виде A1B1
		//данная функция используется только для получения данных из areaMap
		var res;
		AscCommonExcel.executeInR1C1Mode(false, function () {
			res = bbox.getName(AscCommonExcel.referenceType.R);
		});
		return res;
	}

	function DependencyGraph(wb) {
		this.wb = wb;
		//listening
		this.sheetListeners = {};
		this.volatileListeners = {};
		this.defNameListeners = {};
		this.tempGetByCells = [];
		this.volatileArrays = {};	// contains all daf formulas on the sheet that could not be expanded (vm="1")
		//set dirty
		this.isInCalc = false;
		this.changedCell = null;
		this.changedCellRepeated = null;
		this.changedRange = null;
		this.changedRangeRepeated = null;
		this.changedDefName = null;
		this.changedDefNameRepeated = null;
		this.changedShared = {};
		this.buildCell = {};
		this.buildDefName = {};
		this.buildShared = {};
		this.buildArrayFormula = [];
		this.buildPivot = [];
		this.cleanCellCache = {};
		//lock
		this.lockCounter = 0;
		//defined name
		this.defNames = {wb: {}, sheet: {}};
		this.tableNamePattern = "Table";
		this.tableNameIndex = 0;
		this.pivotNamePattern = "PivotTable";
		this.pivotNameIndex = 0;
	}

	DependencyGraph.prototype = {
		maxSharedRecursion: 50,
		maxBroadcastRecursion: 1000,
		//listening
		startListeningRange: function(sheetId, bbox, listener) {
			//todo bbox clone or bbox immutable
			var listenerId = listener.getListenerId();
			var sheetContainer = this.sheetListeners[sheetId];
			if (!sheetContainer) {
				sheetContainer = {cellMap: {}, areaMap: {}, defName3d: {}, rangesTopFrom: null, rangesTopTo: null, rangesBottomFrom: null, rangesBottomTo: null, cells: null };
				this.sheetListeners[sheetId] = sheetContainer;
			}
			if (bbox.isOneCell()) {
				var cellIndex = getCellIndex(bbox.r1, bbox.c1);
				var cellMapElem = sheetContainer.cellMap[cellIndex];
				if (!cellMapElem) {
					cellMapElem = {cellIndex: cellIndex, count: 0, listeners: {}};
					sheetContainer.cellMap[cellIndex] = cellMapElem;
					sheetContainer.cells = null;
				}
				if (!cellMapElem.listeners[listenerId]) {
					cellMapElem.listeners[listenerId] = listener;
					cellMapElem.count++;
				}
			} else {
				var vertexIndex = getVertexIndex(bbox);
				var areaSheetElem = sheetContainer.areaMap[vertexIndex];
				if (!areaSheetElem) {
					//todo clone inside or outside startListeningRange?
					bbox = bbox.clone();
					areaSheetElem = {index: 0, bbox: bbox, count: 0, listeners: {}, isActive: true};
					if (true) {
						areaSheetElem.sharedBroadcast = {changedBBox: null, prevChangedBBox: null, recursion: 0};
					}
					sheetContainer.areaMap[vertexIndex] = areaSheetElem;
					sheetContainer.rangesTopFrom = null;
					sheetContainer.rangesTopTo = null;
					sheetContainer.rangesBottomFrom = null;
					sheetContainer.rangesBottomTo = null;
				}
				if (!areaSheetElem.listeners[listenerId]) {
					areaSheetElem.listeners[listenerId] = listener;
					areaSheetElem.count++;
				}
			}
		},
		endListeningRange: function(sheetId, bbox, listener) {
			var listenerId = listener.getListenerId();
			if (null != listenerId) {
				var sheetContainer = this.sheetListeners[sheetId];
				if (sheetContainer) {
					if (bbox.isOneCell()) {
						var cellIndex = getCellIndex(bbox.r1, bbox.c1);
						var cellMapElem = sheetContainer.cellMap[cellIndex];
						if (cellMapElem && cellMapElem.listeners[listenerId]) {
							delete cellMapElem.listeners[listenerId];
							cellMapElem.count--;
							if (cellMapElem.count <= 0) {
								delete sheetContainer.cellMap[cellIndex];
								sheetContainer.cells = null;
							}
						}
					} else {
						var vertexIndex = getVertexIndex(bbox);
						var areaSheetElem = sheetContainer.areaMap[vertexIndex];
						if (areaSheetElem && areaSheetElem.listeners[listenerId]) {
							delete areaSheetElem.listeners[listenerId];
							areaSheetElem.count--;
							if (areaSheetElem.count <= 0) {
								delete sheetContainer.areaMap[vertexIndex];
								sheetContainer.rangesTopFrom = null;
								sheetContainer.rangesTopTo = null;
								sheetContainer.rangesBottomFrom = null;
								sheetContainer.rangesBottomTo = null;
							}
						}
					}
				}
			}
		},
		startListeningVolatile: function(listener) {
			var listenerId = listener.getListenerId();
			this.volatileListeners[listenerId] = listener;
		},
		endListeningVolatile: function(listener) {
			var listenerId = listener.getListenerId();
			if (null != listenerId) {
				delete this.volatileListeners[listenerId];
			}
		},
		forEachSheetListeners: function(sheetId, callback) {
			var sheetContainer = this.sheetListeners[sheetId];
			if (sheetContainer) {
				var i, j;
				for (i in sheetContainer.cellMap) {
					if (sheetContainer.cellMap[i]) {
						for (j in sheetContainer.cellMap[i].listeners) {
							callback(sheetContainer.cellMap[i].listeners[j]);
						}
					}
				}
				for (i in sheetContainer.areaMap) {
					if (sheetContainer.areaMap[i]) {
						for (j in sheetContainer.areaMap[i].listeners) {
							callback(sheetContainer.areaMap[i].listeners[j]);
						}
					}
				}
			}
		},
		startListeningDefName: function(name, listener, opt_sheetId) {
			var listenerId = listener.getListenerId();
			var nameIndex = getDefNameIndex(name);
			var container = this.defNameListeners[nameIndex];
			if (!container) {
				container = {count: 0, listeners: {}};
				this.defNameListeners[nameIndex] = container;
			}
			if (!container.listeners[listenerId]) {
				container.listeners[listenerId] = listener;
				container.count++;
			}
			if(opt_sheetId){
				var sheetContainer = this.sheetListeners[opt_sheetId];
				if (!sheetContainer) {
					sheetContainer = {cellMap: {}, areaMap: {}, defName3d: {}, rangesTopFrom: null, rangesTopTo: null, rangesBottomFrom: null, rangesBottomTo: null, cells: null };
					this.sheetListeners[opt_sheetId] = sheetContainer;
				}
				sheetContainer.defName3d[listenerId] = listener;
			}
		},
		isListeningDefName: function(name) {
			return null != this.defNameListeners[getDefNameIndex(name)];
		},
		endListeningDefName: function(name, listener, opt_sheetId) {
			var listenerId = listener.getListenerId();
			if (null != listenerId) {
				var nameIndex = getDefNameIndex(name);
				var container = this.defNameListeners[nameIndex];
				if (container && container.listeners[listenerId]) {
					delete container.listeners[listenerId];
					container.count--;
					if (container.count <= 0) {
						delete this.defNameListeners[nameIndex];
					}
				}
				if(opt_sheetId){
					var sheetContainer = this.sheetListeners[opt_sheetId];
					if (sheetContainer) {
						delete sheetContainer.defName3d[listenerId];
					}
				}
			}
		},
		//shift, move
		deleteNodes: function(sheetId, bbox) {
			this.buildDependency();
			this._shiftMoveDelete(c_oNotifyType.Delete, sheetId, bbox, null);
			this.addToChangedRange(sheetId, bbox);
		},
		shift: function(sheetId, bbox, offset) {
			this.buildDependency();
			return this._shiftMoveDelete(c_oNotifyType.Shift, sheetId, bbox, offset);
		},
		move: function(sheetId, bboxFrom, offset, sheetIdTo, isTableCreated) {
			this.buildDependency();
			this._shiftMoveDelete(c_oNotifyType.Move, sheetId, bboxFrom, offset, sheetIdTo, isTableCreated);
			this.addToChangedRange(sheetId, bboxFrom);
		},
		prepareChangeSheet: function(sheetId, data, tableNamesMap) {
			this.buildDependency();
			var listeners = {};
			var sheetContainer = this.sheetListeners[sheetId];
			if (sheetContainer) {
				for (var cellIndex in sheetContainer.cellMap) {
					var cellMapElem = sheetContainer.cellMap[cellIndex];
					for (var listenerId in cellMapElem.listeners) {
						listeners[listenerId] = cellMapElem.listeners[listenerId];
					}
				}
				for (var vertexIndex in sheetContainer.areaMap) {
					var areaSheetElem = sheetContainer.areaMap[vertexIndex];
					for (var listenerId in areaSheetElem.listeners) {
						listeners[listenerId] = areaSheetElem.listeners[listenerId];
					}
				}
				for (var listenerId in sheetContainer.defName3d) {
					listeners[listenerId] = sheetContainer.defName3d[listenerId];
				}
			}
			if (this.volatileListeners) {
				for (var listenerId in this.volatileListeners) {
					listeners[listenerId] = this.volatileListeners[listenerId];
				}
			}
			if(tableNamesMap){
				for (var tableName in tableNamesMap) {
					var nameIndex = getDefNameIndex(tableName);
					var container = this.defNameListeners[nameIndex];
					if (container) {
						for (var listenerId in container.listeners) {
							listeners[listenerId] = container.listeners[listenerId];
						}
					}
				}
			}
			var notifyData = {
				type: c_oNotifyType.Prepare, actionType: c_oNotifyType.ChangeSheet, data: data, transformed: {},
				preparedData: {}
			};
			for (var listenerId in listeners) {
				listeners[listenerId].notify(notifyData);
			}
			for (var listenerId in notifyData.transformed) {
				if (notifyData.transformed.hasOwnProperty(listenerId)) {
					delete listeners[listenerId];
					var elems = notifyData.transformed[listenerId];
					for (var i = 0; i < elems.length; ++i) {
						var elem = elems[i];
						listeners[elem.getListenerId()] = elem;
						elem.notify(notifyData);
					}
				}
			}
			return {listeners: listeners, data: data, preparedData: notifyData.preparedData};
		},
		changeSheet: function(prepared) {
			var notifyData = {type: c_oNotifyType.ChangeSheet, data: prepared.data, preparedData: prepared.preparedData};
			for (var listenerId in prepared.listeners) {
				prepared.listeners[listenerId].notify(notifyData);
			}
		},
		changeExternalLink: function(prepared) {
			var notifyData = {type: c_oNotifyType.ChangeExternalLink, data: prepared.data, preparedData: prepared.preparedData, existedWs: prepared.existedWs};
			for (var listenerId in prepared.listeners) {
				prepared.listeners[listenerId].notify(notifyData);
			}
		},
		prepareRemoveSheet: function(sheetId, tableNames) {
			var t = this;
			//cells
			var ws = this.wb.getWorksheetById(sheetId);
			var formulas = [];
			ws.getAllFormulas(formulas);
			for (var i = 0; i < formulas.length; ++i) {
				formulas[i].removeDependencies();
			}
			//defnames
			this._foreachDefNameSheet(sheetId, function(defName){
				if (!defName.type !== Asc.c_oAscDefNameType.table) {
					t._removeDefName(sheetId, defName.name, AscCH.historyitem_Workbook_DefinedNamesChangeUndo);
				}
			});
			//tables
			var tableNamesMap = {};
			var i;
			for (i = 0; i < tableNames.length; ++i) {
				var tableName = tableNames[i];
				this._removeDefName(null, tableName, null);
				tableNamesMap[tableName] = 1;
				this.wb.deleteSlicersByTable(tableName, true);
			}
			//удаляю срезы, которые остались на данном листе, но привязаны к таблицам других листов
			if (ws.aSlicers) {
				for (i = 0; i < ws.aSlicers.length; i++) {
					ws.deleteSlicer(ws.aSlicers[i].name, true);
				}
			}

			//dependence
			return this.prepareChangeSheet(sheetId, {remove: sheetId, tableNamesMap: tableNamesMap}, tableNamesMap);
		},
		removeSheet: function(prepared) {
			this.changeSheet(prepared);
		},
		//lock
		lockRecal: function() {
			++this.lockCounter;
		},
		isLockRecal: function() {
			return this.lockCounter > 0;
		},
		unlockRecal: function(notCalc, callback) {
			if (0 < this.lockCounter) {
				--this.lockCounter;
			}
			if (!notCalc && 0 >= this.lockCounter) {
				this.calcTree(callback);
			} else {
				callback && callback();
			}
		},
		lockRecalExecute: function(callback) {
			this.lockRecal();
			callback();
			this.unlockRecal();
		},
		//defined name
		getDefNameByName: function(name, sheetId, opt_exact) {
			var res = null;
			var nameIndex = getDefNameIndex(name);
			if (sheetId) {
				var sheetContainer = this.defNames.sheet[sheetId];
				if (sheetContainer) {
					res = sheetContainer[nameIndex];
				}
			}
			if (!res && !(opt_exact && sheetId)) {
				res = this.defNames.wb[nameIndex];
			}
			return res;
		},
		getDefNameByNodeId: function(nodeId) {
			getFromDefNameId(nodeId);
			return this.getDefNameByName(g_FDNI.name, g_FDNI.sheetId, true);
		},
		getDefNameByRef: function(ref, sheetId, bLocale) {
			var getByRef = function(defName) {
				if (!defName.hidden && defName.ref == ref) {
					return defName.name;
				}
			};
			var res = this._foreachDefNameSheet(sheetId, getByRef);
			if(res && bLocale) {
				res = AscCommon.translateManager.getValue(res);
			}
			if (!res) {
				res = this._foreachDefNameBook(getByRef);
			}
			return res;
		},
		getDefNameByCellInside: function(col, row, sheetId, bLocale) {
			var cellSheet = this.wb.getWorksheetById(sheetId);
			var getByRef = function(defName) {
				if (!defName.hidden) {
					var defNameParseRef = defName.ref.split('!');
					if (defNameParseRef) {
						var sheetDefName = defNameParseRef[0];
						var sRefDefName = defNameParseRef[1];

						if (cellSheet && cellSheet.sName === sheetDefName) {
							var refDefName = AscCommonExcel.g_oRangeCache.getAscRange(sRefDefName);
							if (refDefName && refDefName.contains(col, row)) {
								return defName.name;
							}
						}
					}
				}
			};
			var res = this._foreachDefNameSheet(sheetId, getByRef);
			if(res && bLocale) {
				res = AscCommon.translateManager.getValue(res);
			}
			if (!res) {
				res = this._foreachDefNameBook(getByRef);
			}
			return res;
		},
		getDefinedNamesWB: function(type, bLocale, excludeErrorRefNames) {
			var names = [], activeWS;

			function getNames(defName) {
				if (defName.ref && !defName.hidden && (defName.name.indexOf("_xlnm") < 0)) {
					if (excludeErrorRefNames && defName.parsedRef && defName.parsedRef.outStack) {
						var _stack = defName.parsedRef.outStack;
						for (var i = 0; i < _stack.length; i++) {
							if (_stack[i] && cElementType.error === _stack[i].type) {
								return;
							}
						}
					}
					names.push(defName.getAscCDefName(bLocale));
				}
			}

			function sort(a, b) {
				if (a.name > b.name) {
					return 1;
				} else if (a.name < b.name) {
					return -1;
				} else {
					return 0;
				}
			}

			switch (type) {
				case c_oAscGetDefinedNamesList.Worksheet:
				case c_oAscGetDefinedNamesList.WorksheetWorkbook:
					activeWS = this.wb.getActiveWs();
					this._foreachDefNameSheet(activeWS.getId(), getNames);
					if (c_oAscGetDefinedNamesList.WorksheetWorkbook) {
						this._foreachDefNameBook(getNames);
					}
					break;
				case c_oAscGetDefinedNamesList.All:
				default:
					this._foreachDefName(getNames);
					break;
			}
			return names.sort(sort);
		},
		getDefinedNamesWS: function(sheetId) {
			var names = [];

			function getNames(defName) {
				if (defName.ref) {
					names.push(defName);
				}
			}

			this._foreachDefNameSheet(sheetId, getNames);
			return names;
		},
		addDefNameOpen: function(name, ref, sheetIndex, hidden, type) {
			var sheetId = this.wb.getSheetIdByIndex(sheetIndex);
			var isXLNM = null;
			var XLNMName = this._checkXlnmName(name);
			if(null !== XLNMName) {
				name = XLNMName;
				isXLNM = true;
			}

			var defName = new DefName(this.wb, name, ref, sheetId, hidden, type, isXLNM);
			this._addDefName(defName);
			return defName;
		},
		_checkXlnmName: function(name) {
			var supportName = {"Print_Area": 1, "Print_Titles": 1};

			var prefix = "_xlnm.";
			var parseName = name.split(prefix)[1];
			if(supportName[parseName]) {
				return parseName;
			}

			return null;
		},
		addDefName: function(name, ref, sheetId, hidden, type, isXLNM) {
			var defName = new DefName(this.wb, name, ref, sheetId, hidden, type, isXLNM);
			defName.setRef(defName.ref, true);
			this._addDefName(defName);
			return defName;
		},
		removeDefName: function(sheetId, name) {
			this._removeDefName(sheetId, name, AscCH.historyitem_Workbook_DefinedNamesChange);
			if (!this.wb.bUndoChanges && !this.wb.bRedoChanges) {
				this.wb.handlers && this.wb.handlers.trigger("updateCellWatches", sheetId);
			}
			this.wb.handlers && this.wb.handlers.trigger("onChangePageSetupProps", sheetId);
		},
		editDefinesNames: function(oldUndoName, newUndoName) {
			var res = null;
			var isSlicer = oldUndoName && this.wb.getSlicerCacheByCacheName(oldUndoName.name);

			if (!AscCommon.rx_defName.test(getDefNameIndex(newUndoName.name)) || (!newUndoName.ref && !isSlicer) ||
				(newUndoName.ref.length === 0 && !isSlicer) || newUndoName.name.length > g_nDefNameMaxLength) {
				return res;
			}
			if (oldUndoName) {
				res = this.getDefNameByName(oldUndoName.name, oldUndoName.sheetId);
			} else {
				res = this.addDefName(newUndoName.name, newUndoName.ref, newUndoName.sheetId, newUndoName.hidden, newUndoName.type, newUndoName.isXLNM);
			}
			AscCommon.History.Create_NewPoint();
			if (res && oldUndoName) {
				if (oldUndoName.name != newUndoName.name) {
					this.buildDependency();

					res = this._delDefName(res.name, res.sheetId);
					res.setUndoDefName(newUndoName, isSlicer);
					this._addDefName(res);

					var notifyData = {type: c_oNotifyType.ChangeDefName, from: oldUndoName, to: newUndoName};
					this._broadcastDefName(oldUndoName.name, notifyData);

					this.addToChangedDefName(res);
				} else {
					res.setUndoDefName(newUndoName);
				}
			}
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_DefinedNamesChange, null, null,
				new UndoRedoData_FromTo(oldUndoName, newUndoName));

			if (!this.wb.bUndoChanges && !this.wb.bRedoChanges) {
				this.wb.handlers && this.wb.handlers.trigger("updateCellWatches");
			}
			this.wb.handlers && this.wb.handlers.trigger("onChangePageSetupProps", newUndoName.sheetId);
			return res;
		},
		checkDefName: function (name, sheetIndex) {
			var res = new Asc.asc_CCheckDefName();
			var range = AscCommonExcel.g_oRangeCache.getRange3D(name) ||
				AscCommonExcel.g_oRangeCache.getAscRange(name);
			if(!range) {
				//проверяем на совпадение с именем диапазона в другом формате
				AscCommonExcel.executeInR1C1Mode(!AscCommonExcel.g_R1C1Mode, function () {
					range = AscCommonExcel.g_oRangeCache.getRange3D(name) ||
						AscCommonExcel.g_oRangeCache.getAscRange(name);
				});
			}
			if (range || !AscCommon.rx_defName.test(name.toLowerCase()) || name.length > g_nDefNameMaxLength) {
				res.status = false;
				res.reason = c_oAscDefinedNameReason.WrongName;
				return res;
			}

			var sheetId = this.wb.getSheetIdByIndex(sheetIndex);
			var defName = this.getDefNameByName(name, sheetId, true);
			if (defName) {
				res.status = false;
				if (defName.isLock) {
					res.reason = c_oAscDefinedNameReason.IsLocked;
				} else {
					res.reason = c_oAscDefinedNameReason.Existed;
				}
			} else {
				if (this.isListeningDefName(name)) {
					res.status = false;
					res.reason = c_oAscDefinedNameReason.NameReserved;
				} else {
					res.status = true;
					res.reason = c_oAscDefinedNameReason.OK;
				}
			}

			return res;
		},
		copyDefNameByWorksheet: function(wsFrom, wsTo, renameParams, opt_sheet) {
			var sheetContainerFrom;
			var opt_df = opt_sheet && opt_sheet.workbook && opt_sheet.workbook.dependencyFormulas ? opt_sheet.workbook.dependencyFormulas : null;
			if(opt_df && opt_df.defNames && opt_df.defNames.sheet && opt_df.defNames.sheet[wsFrom.getId()]) {
				//TODO пересмотреть!
				//пока делаю только для им. диапазонов листа. в случае книгой - необходимо хранить map преообразований для redo
				sheetContainerFrom = opt_df.defNames.sheet[wsFrom.getId()];
			} else {
				sheetContainerFrom = this.defNames.sheet[wsFrom.getId()];
			}
			if (sheetContainerFrom) {
				for (var name in sheetContainerFrom) {
					var defNameOld = sheetContainerFrom[name];
					if (defNameOld.type !== Asc.c_oAscDefNameType.table && defNameOld.parsedRef) {
						var parsedRefNew = defNameOld.parsedRef.clone();
						parsedRefNew.renameSheetCopy(renameParams);
						var refNew = parsedRefNew.assemble(true);
						this.addDefName(defNameOld.name, refNew, wsTo.getId(), defNameOld.hidden, defNameOld.type);
					}
				}
			}
		},
		copyDefNameByWorkbook: function(wsFrom, wsTo, renameParams, opt_sheet) {
			var t = this;
			var opt_wb = opt_sheet && opt_sheet.workbook;
			var opt_df = opt_wb && opt_wb.dependencyFormulas;

			var doCopy = function (_sheetContainerFrom) {
				if (_sheetContainerFrom) {
					for (var name in _sheetContainerFrom) {
						var defNameOld = _sheetContainerFrom[name];

						if (defNameOld.type !== Asc.c_oAscDefNameType.table && defNameOld.parsedRef) {
							var parsedRefNew = defNameOld.parsedRef.clone();
							parsedRefNew.renameSheetCopy(renameParams);
							var refNew = parsedRefNew.assemble(true);
							var _newDefName = new Asc.asc_CDefName(defNameOld.name, refNew,  null, defNameOld.type, defNameOld.hidden);
							t.wb.editDefinesNames(null, _newDefName);
						}
					}
				}
			};

			if (opt_df) {
				doCopy(opt_df.defNames.wb);
			}
		},
		saveDefName: function(isCopySheet) {
			var list = [];
			var t = this;
			this._foreachDefName(function(defName) {
				if (defName.type !== Asc.c_oAscDefNameType.table && defName.ref) {
					if (!isCopySheet || t._checkDefNamesCopySheet(defName)) {
						list.push(defName.getAscCDefName());
					}
				}
			});
			return list;
		},
		_checkDefNamesCopySheet: function (defName) {
			var res = true;

			var ws = this.wb.getActiveWs();
			if (defName.sheetId && defName.sheetId !== ws.Id) {
				return false;
			}
			if (defName.type === Asc.c_oAscDefNameType.slicer) {
				return false;
			}

			//временная правка - не пишем именованный диапазоны, которые ссылаются на другие листы
			var parsedRef = defName.parsedRef;
			if (parsedRef) {
				for (var i = 0; i < parsedRef.outStack.length; i++) {
					var elem = parsedRef.outStack[i];
					if (cElementType.cell === elem.type || cElementType.cellsRange === elem.type ||
						cElementType.cell3D === elem.type || cElementType.cellsRange3D === elem.type ) {
						if (elem.getWS().getName() !== ws.getName()) {
							return false;
						}
					}
				}
			}

			return res;
		},
		unlockDefName: function() {
			this._foreachDefName(function(defName) {
				defName.isLock = null;
			});
		},
		unlockCurrentDefName: function(name, sheetId) {
			var defName = this.getDefNameByName(name, sheetId);
			if(defName) {
				defName.isLock = null;
			}
		},
		checkDefNameLock: function() {
			return this._foreachDefName(function(defName) {
				return defName.isLock;
			});
		},
		//defined name table
		getNextTableName: function() {
			var sNewName;
			var collaborativeIndexUser = "";
			var api = window["Asc"]["editor"];
			if (api && api.collaborativeEditing && api.collaborativeEditing.getCollaborativeEditing()) {
				collaborativeIndexUser = "_" + this.wb.oApi.CoAuthoringApi.get_indexUser();
			}
			do {
				this.tableNameIndex++;
				var tableName = AscCommon.translateManager.getValue(this.tableNamePattern);
				sNewName = tableName + this.tableNameIndex + collaborativeIndexUser;
			} while (this.getDefNameByName(sNewName, null) || this.isListeningDefName(sNewName));
			return sNewName;
		},
		getNextPivotName: function() {
			var sNewName;
			do {
				this.pivotNameIndex++;
				var tableName = AscCommon.translateManager.getValue(this.pivotNamePattern);
				sNewName = tableName + this.pivotNameIndex;
			} while (this.wb.getPivotTableByName(sNewName));
			return sNewName;
		},
		addTableName: function(ws, table, opt_isOpen) {
			var ref = table.getRangeWithoutHeaderFooter();

			var defNameRef = parserHelp.get3DRef(ws.getName(), ref.getAbsName());
			var defName = this.getDefNameByName(table.DisplayName, null);
			if (!defName) {
				if(opt_isOpen){
					this.addDefNameOpen(table.DisplayName, defNameRef, null, null, Asc.c_oAscDefNameType.table);
				} else {
					this.addDefName(table.DisplayName, defNameRef, null, null, Asc.c_oAscDefNameType.table);
				}
			} else {
				defName.setRef(defNameRef);
			}
		},
		changeTableRef: function(table) {
			var defName = this.getDefNameByName(table.DisplayName, null);
			if (defName) {
				this.buildDependency();
				var oldUndoName = defName.getUndoDefName();
				var newUndoName = defName.getUndoDefName();
				var ref = table.getRangeWithoutHeaderFooter();
				newUndoName.ref = defName.ref.split('!')[0] + '!' + ref.getAbsName();
				AscCommon.History.TurnOff();
				this.editDefinesNames(oldUndoName, newUndoName);
				var notifyData = {type: c_oNotifyType.ChangeDefName, from: oldUndoName, to: newUndoName};
				this._broadcastDefName(defName.name, notifyData);
				AscCommon.History.TurnOn();
				this.addToChangedDefName(defName);
				this.calcTree();
			}
		},
		changeTableName: function(tableName, newName) {
			var defName = this.getDefNameByName(tableName, null);
			if (defName) {
				var oldUndoName = defName.getUndoDefName();
				var newUndoName = defName.getUndoDefName();
				newUndoName.name = newName;
				AscCommon.History.TurnOff();
				this.editDefinesNames(oldUndoName, newUndoName);
				AscCommon.History.TurnOn();
			}
		},
		delTableName: function(tableName, bConvertTableFormulaToRef) {
			this.buildDependency();
			var defName = this.getDefNameByName(tableName);

			this.addToChangedDefName(defName);
			var notifyData = {type: c_oNotifyType.ChangeDefName, from: defName.getUndoDefName(), to: null, bConvertTableFormulaToRef: bConvertTableFormulaToRef};
			this._broadcastDefName(tableName, notifyData);

			this._delDefName(tableName, null);
			if (defName) {
				defName.removeDependencies();
			}
		},
		delColumnTable: function(tableName, deleted) {
			this.buildDependency();
			var notifyData = {type: c_oNotifyType.DelColumnTable, tableName: tableName, deleted: deleted};
			this._broadcastDefName(tableName, notifyData);
		},
		renameTableColumn: function(tableName) {
			var defName = this.getDefNameByName(tableName, null);
			if (defName) {
				this.buildDependency();
				var notifyData = {type: c_oNotifyType.RenameTableColumn, tableName: tableName};
				this._broadcastDefName(defName.name, notifyData);
			}
			this.calcTree();
		},
		addToVolatileArrays: function(formula) {
			if (!this.volatileArrays) {
				this.volatileArrays = {};
			}
			let listenerId = formula.getListenerId();
			this.volatileArrays[listenerId] = formula;
		},
		getVolatileArrays: function() {
			if (!this.volatileArrays) {
				return
			}

			return this.volatileArrays
		},
		endListeningVolatileArray: function(listenerId) {
			if (listenerId) {
				delete this.volatileArrays[listenerId];
			}
		},
		//set dirty
		addToChangedRange2: function(sheetId, bbox) {
			if (bbox.isOneCell()) {
				this.addToChangedCell2(sheetId, bbox.r1, bbox.c1);
				return;
			}
			if (!this.changedRange) {
				this.changedRange = {};
			}
			var changedSheet = this.changedRange[sheetId];
			if (!changedSheet) {
				//{}, а не [], потому что при сборке может придти сразу много одинаковых ячеек
				changedSheet = {};
				this.changedRange[sheetId] = changedSheet;
			}
			var name = getVertexIndex(bbox);
			if (this.isInCalc && !changedSheet[name]) {
				//todo change cell
				if (!this.changedRangeRepeated) {
					this.changedRangeRepeated = {};
				}
				var changedSheetRepeated = this.changedRangeRepeated[sheetId];
				if (!changedSheetRepeated) {
					changedSheetRepeated = {};
					this.changedRangeRepeated[sheetId] = changedSheetRepeated;
				}
				changedSheetRepeated[name] = bbox;
			}
			changedSheet[name] = bbox;
		},
		addToChangedCell2: function(sheetId, row, col) {
			var t = this;
			if (!this.changedCell) {
				this.changedCell = {};
			}
			var changedSheet = this.changedCell[sheetId];
			if (!changedSheet) {
				//{}, а не [], потому что при сборке может придти сразу много одинаковых ячеек
				changedSheet = {};
				this.changedCell[sheetId] = changedSheet;
			}

			var addChangedSheet = function(row, col) {
				var cellIndex = getCellIndex(row, col);
				if (t.isInCalc && undefined === changedSheet[cellIndex]) {
					if (!t.changedCellRepeated) {
						t.changedCellRepeated = {};
					}
					var changedSheetRepeated = t.changedCellRepeated[sheetId];
					if (!changedSheetRepeated) {
						changedSheetRepeated = {};
						t.changedCellRepeated[sheetId] = changedSheetRepeated;
					}
					changedSheetRepeated[cellIndex] = cellIndex;
				}
				changedSheet[cellIndex] = cellIndex;
			};

			addChangedSheet(row, col);
		},
		addToChangedCell: function(cell) {
			this.addToChangedCell2(cell.ws.getId(), cell.nRow, cell.nCol);
		},
		addToChangedDefName: function(defName) {
			if (!this.changedDefName) {
				this.changedDefName = {};
			}
			var nodeId = defName.getNodeId();
			if (this.isInCalc && !this.changedDefName[nodeId]) {
				if (!this.changedDefNameRepeated) {
					this.changedDefNameRepeated = {};
				}
				this.changedDefNameRepeated[nodeId] = 1;
			}
			this.changedDefName[nodeId] = 1;
		},
		addToChangedRange: function(sheetId, bbox) {
			var notifyData = {type: c_oNotifyType.Dirty};
			var sheetContainer = this.sheetListeners[sheetId];
			if (sheetContainer) {
				for (var cellIndex in sheetContainer.cellMap) {
					getFromCellIndex(cellIndex);
					if (bbox.contains(g_FCI.col, g_FCI.row)) {
						var cellMapElem = sheetContainer.cellMap[cellIndex];
						for (var listenerId in cellMapElem.listeners) {
							cellMapElem.listeners[listenerId].notify(notifyData);
						}
					}
				}
				for (var areaIndex in sheetContainer.areaMap) {
					var areaMapElem = sheetContainer.areaMap[areaIndex];
					var isIntersect = bbox.isIntersect(areaMapElem.bbox);
					if (isIntersect) {
						for (var listenerId in areaMapElem.listeners) {
							areaMapElem.listeners[listenerId].notify(notifyData);
						}
					}
				}
			}
		},
		addToChangedHiddenRows: function() {
			//notify hidden rows
			var tmpRange = new Asc.Range(0, 0, gc_nMaxCol0, 0);
			for (var i = 0; i < this.wb.aWorksheets.length; ++i) {
				var ws = this.wb.aWorksheets[i];
				var hiddenRange = ws.hiddenManager.getHiddenRowsRange();
				if (hiddenRange) {
					tmpRange.r1 = hiddenRange.r1;
					tmpRange.r2 = hiddenRange.r2;
					this.addToChangedRange(ws.getId(), tmpRange);
				}
			}
		},
		addToBuildDependencyCell: function(cell) {
			var sheetId = cell.ws.getId();
			var unparsedSheet = this.buildCell[sheetId];
			if (!unparsedSheet) {
				//{}, а не [], потому что при сборке может придти сразу много одинаковых ячеек
				unparsedSheet = {};
				this.buildCell[sheetId] = unparsedSheet;
			}
			unparsedSheet[getCellIndex(cell.nRow, cell.nCol)] = 1;
		},
		addToBuildDependencyDefName: function(defName) {
			this.buildDefName[defName.getNodeId()] = 1;
		},
		addToBuildDependencyShared: function(shared) {
			this.buildShared[shared.getIndexNumber()] = shared;
		},
		addToBuildDependencyArray: function(f) {
			//TODO переммотреть! добавлять по индексу!
			//добавляю не по индексу потому на момент вызова(setValue) ещё не проставились индексы формулам
			//происходит это позже - в Cell.prototype.saveContent
			this.buildArrayFormula.push(f);
		},
		addToBuildDependencyPivot: function(f) {
			this.buildPivot.push(f);
		},
		addToCleanCellCache: function(sheetId, row, col) {
			var sheetArea = this.cleanCellCache[sheetId];
			if (sheetArea) {
				sheetArea.union3(col, row);
			} else {
				this.cleanCellCache[sheetId] = new Asc.Range(col, row, col, row);
			}
		},
		addToChangedShared: function(parsed) {
			this.changedShared[parsed.getIndexNumber()] = parsed;
		},
		notifyChanged: function(changedFormulas) {
			var notifyData = {type: c_oNotifyType.Dirty};
			for (var listenerId in changedFormulas) {
				changedFormulas[listenerId].notify(notifyData);
			}
		},
		//build, calc
		buildDependency: function() {
			for (var sheetId in this.buildCell) {
				var ws = this.wb.getWorksheetById(sheetId);
				if (ws) {
					var unparsedSheet = this.buildCell[sheetId];
					for (var cellIndex in unparsedSheet) {
						getFromCellIndex(cellIndex);
						ws._getCellNoEmpty(g_FCI.row, g_FCI.col, function(cell) {
							if (cell) {
								cell._BuildDependencies(true, true);
							}
						});
					}
				}
			}
			for (var defNameId in this.buildDefName) {
				var defName = this.getDefNameByNodeId(defNameId);
				if (defName && defName.parsedRef) {
					defName.parsedRef.parse();
					defName.parsedRef.buildDependencies();
					this.addToChangedDefName(defName);
				}
			}
			for (var index in this.buildShared) {
				if (this.buildShared.hasOwnProperty(index)) {
					var parsed = this.wb.workbookFormulas.get(index - 0);
					if (parsed) {
						parsed.parse();
						parsed.buildDependencies();
						var shared = parsed.getShared();
						this.wb.dependencyFormulas.addToChangedRange2(parsed.getWs().getId(), shared.ref);
					}
				}
			}
			for (var index = 0; index < this.buildArrayFormula.length; ++index) {
				var parsed = this.buildArrayFormula[index];
				if (parsed) {
					parsed.parse();
					parsed.buildDependencies();
					var array = parsed.getArrayFormulaRef();
					this.wb.dependencyFormulas.addToChangedRange2(parsed.getWs().getId(), array);
				}
			}
			for (var index = 0; index < this.buildPivot.length; ++index) {
				var parsed = this.buildPivot[index];
				if (parsed) {
					parsed.parse();
					parsed.buildDependencies();
				}
			}
			this.buildCell = {};
			this.buildDefName = {};
			this.buildShared = {};
			this.buildArrayFormula = [];
			this.buildPivot = [];
		},
		calcTree: function(callback) {
			if (this.lockCounter > 0) {
				callback && callback();
				return;
			}
			this.buildDependency();
			this.addToChangedHiddenRows();
			if (!(this.changedCell || this.changedRange || this.changedDefName)) {
				callback && callback();
				return;
			}
			var notifyData = {type: c_oNotifyType.Dirty, areaData: undefined};
			if (!g_cCalcRecursion.getIsSheetCreating()) {
				this._broadscastVolatile(notifyData);
			}
			this._broadcastCellsStart();
			let repeatCount = 0;
			while (this.changedCellRepeated || this.changedRangeRepeated || this.changedDefNameRepeated) {
				if (++repeatCount > this.maxBroadcastRecursion) {
					this.notifyAllFormulasInChangedWs();
				}

				this._broadcastDefNames(notifyData);
				this._broadcastCells(notifyData);
				this._broadcastRanges(notifyData);
			}
			this._broadcastCellsEnd();
			do {
				this._calculateDirty();
				if (g_cCalcRecursion.getIsEnabledRecursion()) {
					g_cCalcRecursion.incIterStep();
				}
			} while (g_cCalcRecursion.needRecursiveCall());

			if (!this.wb.asyncFormulasManager.isRecalculating() && this.wb.asyncFormulasManager.getPromises() && !AscCommonExcel.g_LockCustomFunctionRecalculate) {
				this.wb.asyncFormulasManager.calculatePromises(callback);
			} else if (!g_cCalcRecursion.needRecursiveCall() && !this.wb.asyncFormulasManager.isRecalculating()) {
				callback && callback();
			}

			if (/*!this.wb.asyncFormulasManager.isRecalculating() &&*/ this.wb.asyncFormulasManager.getReplacedFormulas() && !AscCommonExcel.g_LockCustomFunctionRecalculate) {
				this.wb.asyncFormulasManager.calculateReplacedFormulas(callback);
			} else if (!g_cCalcRecursion.needRecursiveCall() && !this.wb.asyncFormulasManager.isRecalculating()) {
				callback && callback();
			}

			g_cCalcRecursion.resetIterStep();
			g_cCalcRecursion.setStartCellIndex(null);
			g_cCalcRecursion.setGroupChangedCells(null);
			g_cCalcRecursion.clearPrevIterResult();
			g_cCalcRecursion.clearDiffBetweenIter();
			g_cCalcRecursion.clearCycleCells();
			this.changedCell = null;
			this.changedRange = null;
			this.updateSharedFormulas();
			//copy cleanCellCache to prevent recursion in trigger("cleanCellCache")
			var tmpCellCache = this.cleanCellCache;
			this.cleanCellCache = {};
			for (var i in tmpCellCache) {
				this.wb.handlers && this.wb.handlers.trigger("cleanCellCache", i, [tmpCellCache[i]], true);
			}
			AscCommonExcel.g_oLOOKUPCache.clean();
			AscCommonExcel.g_oVLOOKUPCache.clean();
			AscCommonExcel.g_oHLOOKUPCache.clean();
			AscCommonExcel.g_oMatchCache.clean();
			AscCommonExcel.g_oSUMIFSCache.clean();
			AscCommonExcel.g_oFormulaRangesCache.clean();
			AscCommonExcel.g_oCountIfCache.clean();
		},
		notifyAllFormulasInChangedWs: function() {
			let sheetIds = {};
			for (let sheetId in this.changedCell) {
				sheetIds[sheetId] = 1;
			}
			for (let sheetId in this.changedRangeRepeated) {
				sheetIds[sheetId] = 1;
			}
			let formulas = [];
			for (let sheetId in sheetIds) {
				let ws = this.wb.getWorksheetById(sheetId);
				if (ws) {
					ws.getAllFormulas(formulas);
				}
			}
			this.notifyChanged(formulas);
		},
		initOpen: function() {
			this._foreachDefName(function(defName) {
				defName.setRef(defName.ref, true, true, true);
			});
		},
		getSnapshot: function(wb) {
			var res = new DependencyGraph(wb);
			this._foreachDefName(function(defName){
				//_addDefName because we don't need dependency
				//include table defNames too.
				res._addDefName(defName.clone(wb));
			});
			res.tableNameIndex = this.tableNameIndex;
			return res;
		},
		getAllFormulas: function(formulas) {
			this._foreachDefName(function(defName) {
				if (defName.parsedRef) {
					formulas.push(defName.parsedRef);
				}
			});
		},
		forEachFormula: function(callback) {
			this._foreachDefName(function(defName) {
				if (defName.parsedRef) {
					callback(defName.parsedRef);
				}
			});
		},
		updateSharedFormulas: function() {
			var newRef;
			for (var indexNumber in this.changedShared) {
				var parsed = this.changedShared[indexNumber];
				var shared = parsed.getShared();
				if (shared) {
					var ws = parsed.getWs();
					var ref = shared.ref;
					var r1 = gc_nMaxRow0;
					var c1 = gc_nMaxCol0;
					var r2 = 0;
					var c2 = 0;
					ws.getRange3(ref.r1, ref.c1, ref.r2, ref.c2)._foreachNoEmpty(function(cell) {
						if (parsed === cell.getFormulaParsed()) {
							r1 = Math.min(r1, cell.nRow);
							c1 = Math.min(c1, cell.nCol);
							r2 = Math.max(r2, cell.nRow);
							c2 = Math.max(c2, cell.nCol);
						}
					});
					newRef = undefined;
					if (r1 <= r2 && c1 <= c2) {
						newRef = new Asc.Range(c1, r1, c2, r2);
					}
					parsed.setSharedRef(newRef);
				}
			}
			this.changedShared = {};
		},
		//internal
		_addDefName: function(defName) {
			var nameIndex = getDefNameIndex(defName.name);
			var container;
			var sheetId = defName.sheetId;
			if (sheetId) {
				container = this.defNames.sheet[sheetId];
				if (!container) {
					container = {};
					this.defNames.sheet[sheetId] = container;
				}
			} else {
				container = this.defNames.wb;
			}
			var cur = container[nameIndex];
			if (cur) {
				cur.removeDependencies();
			}
			container[nameIndex] = defName;
		},
		_removeDefName: function(sheetId, name, historyType) {
			var defName = this._delDefName(name, sheetId);
			if (defName) {
				if (null != historyType) {
					AscCommon.History.Create_NewPoint();
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, historyType, null, null,
								new UndoRedoData_FromTo(defName.getUndoDefName(), null));
				}

				defName.removeDependencies();
				this.addToChangedDefName(defName);
			}
		},
		_delDefName: function(name, sheetId) {
			var res = null;
			var nameIndex = getDefNameIndex(name);
			var sheetContainer;
			if (sheetId) {
				sheetContainer = this.defNames.sheet[sheetId];
			}
			else {
				sheetContainer = this.defNames.wb;
			}
			if (sheetContainer) {
				res = sheetContainer[nameIndex];
				delete sheetContainer[nameIndex];
			}
			return res;
		},
		_foreachDefName: function(action) {
			var containerSheet;
			var sheetId;
			var name;
			var res;
			for (sheetId in this.defNames.sheet) {
				containerSheet = this.defNames.sheet[sheetId];
				for (name in containerSheet) {
					res = action(containerSheet[name], containerSheet);
					if (res) {
						break;
					}
				}
			}
			if (!res) {
				res = this._foreachDefNameBook(action);
			}
			return res;
		},
		_foreachDefNameSheet: function(sheetId, action) {
			var name;
			var res;
			var containerSheet = this.defNames.sheet[sheetId];
			if (containerSheet) {
				for (name in containerSheet) {
					res = action(containerSheet[name], containerSheet);
					if (res) {
						break;
					}
				}
			}
			return res;
		},
		_foreachDefNameBook: function(action) {
			var containerSheet;
			var name;
			var res;
			for (name in this.defNames.wb) {
				res = action(this.defNames.wb[name], this.defNames.wb);
				if (res) {
					break;
				}
			}
			return res;
		},
		_broadscastVolatile: function(notifyData) {
			if (this.wb.asyncFormulasManager.isRecalculating()) {
				return;
			}
			for (var i in this.volatileListeners) {
				this.volatileListeners[i].notify(notifyData);
			}
		},
		_broadcastDefName: function(name, notifyData) {
			var nameIndex = getDefNameIndex(name);
			var container = this.defNameListeners[nameIndex];
			if (container) {
				for (var listenerId in container.listeners) {
					container.listeners[listenerId].notify(notifyData);
				}
			}
		},
		_broadcastDefNames: function(notifyData) {
			if (this.changedDefNameRepeated) {
				var changedDefName = this.changedDefNameRepeated;
				this.changedDefNameRepeated = null;
				for (var nodeId in changedDefName) {
					getFromDefNameId(nodeId);
					this._broadcastDefName(g_FDNI.name, notifyData);
				}
			}
		},
		_broadcastCells: function(notifyData) {
			if (this.changedCellRepeated) {
				var changedCell = this.changedCellRepeated;
				this.changedCellRepeated = null;
				for (var sheetId in changedCell) {
					var changedSheet = changedCell[sheetId];
					var sheetContainer = this.sheetListeners[sheetId];
					if (sheetContainer) {
						var cells = [];
						for (var cellIndex in changedSheet) {
							cells.push(changedSheet[cellIndex]);
						}
						cells.sort(AscCommon.fSortAscending);
						this._broadcastCellsByCells(sheetContainer, cells, notifyData);
						//todo add changedCellRepeated after _broadcastCellsByCells
						this._broadcastRangesByCells(sheetContainer, cells, notifyData);
					}
				}
			}
		},
		_broadcastRanges: function(notifyData) {
			if (this.changedRangeRepeated) {
				var changedRange = this.changedRangeRepeated;
				this.changedRangeRepeated = null;
				for (var sheetId in changedRange) {
					var changedSheet = changedRange[sheetId];
					var sheetContainer = this.sheetListeners[sheetId];
					if (sheetContainer) {
						if (sheetContainer) {
							let rangesTopFrom = [];
							let counter = 1;
							for (let name in changedSheet) {
								rangesTopFrom.push({index: 0, bbox: changedSheet[name]});
							}
							let rangesTopTo = rangesTopFrom.concat();
							let rangesBottomFrom = rangesTopFrom.concat();
							let rangesBottomTo = rangesTopFrom.concat();

							rangesTopFrom.sort(AscCommonExcel.BroadcastHelper.prototype.cmpInsertFrom);
							rangesTopTo.sort(AscCommonExcel.BroadcastHelper.prototype.cmpInsertTo);
							rangesBottomFrom.sort(AscCommonExcel.BroadcastHelper.prototype.cmpDeleteFrom);
							rangesBottomTo.sort(AscCommonExcel.BroadcastHelper.prototype.cmpDeleteTo);
							this._broadcastCellsByRanges(sheetContainer, rangesTopFrom, rangesTopTo, rangesBottomFrom, rangesBottomTo, notifyData);
							this._broadcastRangesByRanges(sheetContainer, rangesTopFrom, rangesTopTo, rangesBottomFrom, rangesBottomTo, notifyData);
						}
					}
				}
			}
		},
		_broadcastCellsStart: function() {
			this.isInCalc = true;
			this.changedCellRepeated = this.changedCell;
			this.changedRangeRepeated = this.changedRange;
			this.changedDefNameRepeated = this.changedDefName;
			for (var sheetId in this.sheetListeners) {
				var sheetContainer = this.sheetListeners[sheetId];
				if (!sheetContainer.cells) {
					sheetContainer.cells = [];
					for (var cellIndex in sheetContainer.cellMap) {
						sheetContainer.cells.push(sheetContainer.cellMap[cellIndex]);
					}
					sheetContainer.cells.sort(function(a, b) {
						return a.cellIndex - b.cellIndex
					});
				}
				if (!sheetContainer.rangesTopFrom || !sheetContainer.rangesTopTo || !sheetContainer.rangesBottomFrom || !sheetContainer.rangesBottomTo) {
					sheetContainer.rangesTopFrom = [];
					for (let name in sheetContainer.areaMap) {
						sheetContainer.rangesTopFrom.push(sheetContainer.areaMap[name]);
					}
					sheetContainer.rangesTopTo = sheetContainer.rangesTopFrom.concat();
					sheetContainer.rangesBottomFrom = sheetContainer.rangesTopFrom.concat();
					sheetContainer.rangesBottomTo = sheetContainer.rangesTopFrom.concat();

					sheetContainer.rangesTopFrom.sort(AscCommonExcel.BroadcastHelper.prototype.cmpInsertFrom);
					sheetContainer.rangesTopTo.sort(AscCommonExcel.BroadcastHelper.prototype.cmpInsertTo);
					sheetContainer.rangesBottomFrom.sort(AscCommonExcel.BroadcastHelper.prototype.cmpDeleteFrom);
					sheetContainer.rangesBottomTo.sort(AscCommonExcel.BroadcastHelper.prototype.cmpDeleteTo);
				}
			}
		},
		_broadcastCellsEnd: function() {
			this.isInCalc = false;
			this.changedDefName = null;
			for (var i = 0; i < this.tempGetByCells.length; ++i) {
				for (var j = 0; j < this.tempGetByCells[i].length; ++j) {
					var temp = this.tempGetByCells[i][j];

					temp.isActive = true;
					if (temp.sharedBroadcast) {
						temp.sharedBroadcast.changedBBox = null;
						temp.sharedBroadcast.prevChangedBBox = null;
						temp.sharedBroadcast.recursion = 0;
					}
				}
			}
			this.tempGetByCells = [];
		},

		/**
		 * Method starts preparing for calculate cells with formula. Recognizes what cells need to calculate.
		 * @memberof DependencyGraph
		 * @private
		 */
		_calculateDirty: function() {
			const t = this;
			let needUpdateCells = [];
			this._foreachChanged(function (oCell) {
				if (oCell && oCell.isFormula()) {
					const oFormulaParsed = oCell.getFormulaParsed();
					// Logic for iterative calculation
					if (g_cCalcRecursion.getIsEnabledRecursion()) {
						const nThisCellIndex = getCellIndex(oCell.nRow, oCell.nCol);
						const sCellWsName = oCell.ws.getName().toLowerCase();
						const aRecursiveCells = g_cCalcRecursion.getRecursiveCells(oCell);

						// Fills 0 value for empty cells for potentially recursive formula.
						if (oCell.getNumberValue() == null && (oCell.getValueText() == null || oCell.getValueText() === '#NUM!')) {
							if (oCell.getType() !== CellValueType.Number) {
								oCell.setTypeInternal(CellValueType.Number);
							}
							oCell.setValueNumberInternal(0);
						}
						g_cCalcRecursion.updateStartCellIndex(aRecursiveCells);
						if (!g_cCalcRecursion.getStartCellIndex()) {
							oCell.initStartCellForIterCalc();
						}
						let oStartCellIndex = g_cCalcRecursion.getStartCellIndex();
						if (oStartCellIndex) {
							/** Checks static formulas with ca flag - false. For iterative calc static formulas that
							 	participate in recursive formulas must be calculated before calculating dependent cells.**/
							if (!oFormulaParsed.ca) {
								oCell.changeLinkedCell(function (oLinkedCell) {
									const oFormulaParsed = oLinkedCell.getFormulaParsed();
									if (!oFormulaParsed.ca) {
										oLinkedCell.setIsDirty(true);
										oLinkedCell._checkDirty();
									}
								}, true);
								oCell.setIsDirty(true);
								oCell._checkDirty();
							}
							// Check result of the formula is convergent
							let nDiffBetweenIter = g_cCalcRecursion.getDiffBetweenIter(oCell);
							if (!isNaN(nDiffBetweenIter) && nDiffBetweenIter < g_cCalcRecursion.getRelativeError()) {
								g_cCalcRecursion.removeRecursionCell(oCell);
								if (g_cCalcRecursion.getPrevIterResult(oCell) !== 0 && g_cCalcRecursion.getIterStep() === 2) {
									oCell.setValueNumberInternal(g_cCalcRecursion.getPrevIterResult(oCell));
								}
								return;
							} else  {
								// Check diff of linked cell
								let bLinkCellConvergent = false;
								oCell.changeLinkedCell(function (oLinkedCell) {
									const nLinkedDiff =  g_cCalcRecursion.getDiffBetweenIter(oLinkedCell);
									if (!isNaN(nLinkedDiff) && nLinkedDiff < g_cCalcRecursion.getRelativeError()) {
										bLinkCellConvergent = true;
									}
								}, false);
								if (bLinkCellConvergent) {
									g_cCalcRecursion.removeRecursionCell(oCell);
									return;
								}
							}
							// Save current value of cell as result of previous iteration
							g_cCalcRecursion.setPrevIterResult(oCell);
							// Fill the array with linked cells from the recursive formula
							if (!aRecursiveCells.length && oStartCellIndex.cellId === nThisCellIndex && oStartCellIndex.wsName === sCellWsName) {
								oCell.fillRecursiveCells(aRecursiveCells, oStartCellIndex);
							}
							// Disable calculating formula for linked cells
							let bRecursiveCell = aRecursiveCells.some(function(oCellIndex) {
								return oCellIndex.cellId === nThisCellIndex && oCellIndex.wsName === sCellWsName;
							});
							if ((nThisCellIndex !== oStartCellIndex.cellId || oStartCellIndex.wsName !== sCellWsName) && bRecursiveCell) {
								oCell.setIsDirty(false);
								return;
							}
						} else if (g_cCalcRecursion.getIterStep() > 1 && oCell.getValueWithoutFormat() && !g_cCalcRecursion.hasInRecursiveCells(oCell)) {
							oCell.setIsDirty(false);
							return;
						}
					} else if (oFormulaParsed.ca === true) {
						oCell.initStartCellForIterCalc();
						if (g_cCalcRecursion.getStartCellIndex()) {
							g_cCalcRecursion.addCycleCell(oCell);
							const oRefRange = oCell.getFormulaParsed().ref;
							let bEmptyCell = oCell.getNumberValue() == null && (oCell.getValueText() == null || oCell.getValueText() === '#NUM!');
							let bRefEmptyCell = !!oRefRange && oRefRange.r1 === oCell.nRow && oRefRange.c1 === oCell.nCol && bEmptyCell;
							if ((!oRefRange && bEmptyCell) || bRefEmptyCell) {
								if (oCell.getType() !== CellValueType.Number) {
									oCell.setTypeInternal(CellValueType.Number);
								}
								oCell.setValueNumberInternal(0);
							}
							oCell.setIsDirty(false);
							g_cCalcRecursion.setStartCellIndex(null);
							return;
						}
					}
					oCell.setIsDirty(true);
				}
			});

			AscCommonExcel.importRangeLinksState.startBuildImportRangeLinks = false;

			this._foreachChanged(function (oCell) {
				if (oCell) {
					oCell._checkDirty();
					if (oCell.formulaParsed && AscCommonExcel.bIsSupportDynamicArrays && (oCell.formulaParsed.getDynamicRef() || oCell.formulaParsed.getArrayFormulaRef()) && oCell.formulaParsed.aca && oCell.formulaParsed.ca) {
						t.addToVolatileArrays(oCell.formulaParsed);
					}
					// Enable calculating formula for next cell in chain
					const oFormulaParsed = oCell.getFormulaParsed();
					if (g_cCalcRecursion.getIsEnabledRecursion() && oCell.isFormula() && oFormulaParsed.ca === true) {
						const nThisCellIndex = getCellIndex(oCell.nRow, oCell.nCol);
						const sCellWsName = oCell.ws.getName().toLowerCase();
						const aRecursiveCells = g_cCalcRecursion.getRecursiveCells(oCell);
						let bLinkedCell = aRecursiveCells.some(function (oCellIndex) {
							return oCellIndex.cellId === nThisCellIndex && oCellIndex.wsName === sCellWsName;
						});
						if (aRecursiveCells.length && bLinkedCell) {
							oCell.changeLinkedCell(function (oCell) {
								const oFormulaParsed = oCell.getFormulaParsed();
								const nCellIndex = getCellIndex(oCell.nRow, oCell.nCol);
								const sCellWsName = oCell.ws.getName().toLowerCase();
								let bLinkedCell = aRecursiveCells.some(function (oCellIndex) {
									return oCellIndex.cellId === nCellIndex && oCellIndex.wsName === sCellWsName;
								});
								if (bLinkedCell && oFormulaParsed.ca === true && !oCell.getIsDirty()) {
									oCell.setIsDirty(true);
								}
							}, false);
							g_cCalcRecursion.calcDiffBetweenIter(oCell);
						}
					}
				}
			});

			if (g_cCalcRecursion.getCycleCells().length && g_cCalcRecursion.getShowCycleWarn()) {
				const oApi = Asc.editor;
				oApi.sendEvent("asc_onError", c_oAscError.ID.CircularReference, c_oAscError.Level.NoCritical);
				g_cCalcRecursion.setShowCycleWarn(false);
			}
			if (!g_cCalcRecursion.getCycleCells().length && !g_cCalcRecursion.getShowCycleWarn()) {
				g_cCalcRecursion.setShowCycleWarn(true);
			}

			if (AscCommonExcel.importRangeLinksState.importRangeLinks) {
				//need update

				let needUpdateExternalReference = [];
				let newExternalReferences = [];
				for (let i in AscCommonExcel.importRangeLinksState.importRangeLinks) {
					let newExternalReference = new AscCommonExcel.ExternalReference();
					//newExternalReference.referenceData = referenceData;
					newExternalReference.Id = i;

					for (var j = 0; j < AscCommonExcel.importRangeLinksState.importRangeLinks[i].length; j++) {
						let newSheet = AscCommonExcel.importRangeLinksState.importRangeLinks[i][j].sheet;
						newExternalReference.addSheetName(newSheet, true);
						newExternalReference.initWorksheetFromSheetDataSet(newSheet);

						let ws = newExternalReference.worksheets[newSheet];

						newExternalReference.initRows(ws.getRange2(AscCommonExcel.importRangeLinksState.importRangeLinks[i][j].range));
					}

					newExternalReference.notUpdateId = true;

					newExternalReferences.push(newExternalReference);
					needUpdateExternalReference.push(newExternalReference.getAscLink())
				}

				AscCommonExcel.importRangeLinksState.importRangeLinks = null;

				this.wb.addExternalReferences(newExternalReferences);


				this.wb.dependencyFormulas.lockRecal();
				let t = this;
				this.wb.oApi.asc_updateExternalReferences(needUpdateExternalReference, function () {
					t.wb.dependencyFormulas.unlockRecal(true);

					//needUpdateCells
					for (let i = 0; i < needUpdateCells.length; i++) {
						t.wb.dependencyFormulas.addToChangedCell(needUpdateCells[i]);
					}
					t.wb.sortDependency();
					AscCommonExcel.importRangeLinksState.startBuildImportRangeLinks = null;


					t.wb.drawWorksheet();
				});
			} else {
				AscCommonExcel.importRangeLinksState.startBuildImportRangeLinks = null;
			}
		},
		_foreachChanged: function(callback) {
			var sheetId, changedSheet, ws, bbox;
			for (sheetId in this.changedCell) {
				if (this.changedCell.hasOwnProperty(sheetId)) {
					changedSheet = this.changedCell[sheetId];
					ws = this.wb.getWorksheetById(sheetId);
					if (changedSheet && ws) {
						for (var cellIndex in changedSheet) {
							if (changedSheet.hasOwnProperty(cellIndex)) {
								getFromCellIndex(cellIndex);
								ws._getCell(g_FCI.row, g_FCI.col, callback);
							}
						}
					}
				}
			}
			for (sheetId in this.changedRange) {
				if (this.changedRange.hasOwnProperty(sheetId)) {
					changedSheet = this.changedRange[sheetId];
					ws = this.wb.getWorksheetById(sheetId);
					if (changedSheet && ws) {
						for (var name in changedSheet) {
							if (changedSheet.hasOwnProperty(name)) {
								bbox = changedSheet[name];
								ws.getRange3(bbox.r1, bbox.c1, bbox.r2, bbox.c2)._foreachNoEmpty(callback);
							}
						}
					}
				}
			}
		},
		_shiftMoveDelete: function(notifyType, sheetId, bbox, offset, sheetIdTo, isTableCreated) {
			var listeners = {};
			var res = {changed: listeners, shiftedShared: {}};
			var sheetContainer = this.sheetListeners[sheetId];
			if (sheetContainer) {
				var bboxShift;
				if (c_oNotifyType.Shift === notifyType) {
					var bHor = 0 !== offset.col;
					bboxShift = AscCommonExcel.shiftGetBBox(bbox, bHor);
				}
				var isIntersect;
				for (var cellIndex in sheetContainer.cellMap) {
					getFromCellIndex(cellIndex);
					if (c_oNotifyType.Shift === notifyType) {
						isIntersect = bboxShift.contains(g_FCI.col, g_FCI.row);
					} else {
						isIntersect = bbox.contains(g_FCI.col, g_FCI.row);
					}
					if (isIntersect) {
						var cellMapElem = sheetContainer.cellMap[cellIndex];
						for (var listenerId in cellMapElem.listeners) {
							listeners[listenerId] = cellMapElem.listeners[listenerId];
						}
					}
				}
				for (var areaIndex in sheetContainer.areaMap) {
					var areaMapElem = sheetContainer.areaMap[areaIndex];
					if (c_oNotifyType.Shift === notifyType) {
						isIntersect = bboxShift.isIntersect(areaMapElem.bbox);
					} else if (c_oNotifyType.Move === notifyType || c_oNotifyType.Delete === notifyType) {
						isIntersect = bbox.isIntersect(areaMapElem.bbox);
					}
					if (isIntersect) {
						for (var listenerId in areaMapElem.listeners) {
							listeners[listenerId] = areaMapElem.listeners[listenerId];
						}
					}
				}
				var notifyData = {
					type: notifyType, sheetId: sheetId, sheetIdTo: sheetIdTo, bbox: bbox, offset: offset, shiftedShared: res.shiftedShared, isTableCreated: isTableCreated
				};
				for (var listenerId in listeners) {
					listeners[listenerId].notify(notifyData);
				}
			}
			return res;
		},
		_broadcastCellsByCells: function(sheetContainer, cellsChanged, notifyData) {
			var cells = sheetContainer.cells;
			var indexCell = 0;
			var indexCellChanged = 0;
			var row, col, rowChanged, colChanged;
			if (indexCell < cells.length) {
				getFromCellIndex(cells[indexCell].cellIndex);
				row = g_FCI.row;
				col = g_FCI.col;
			}
			if (indexCellChanged < cellsChanged.length) {
				getFromCellIndex(cellsChanged[indexCellChanged]);
				rowChanged = g_FCI.row;
				colChanged = g_FCI.col;
			}
			while (indexCell < cells.length && indexCellChanged < cellsChanged.length) {
				var cmp = Asc.Range.prototype.compareCell(col, row, colChanged, rowChanged);
				if (cmp > 0) {
					indexCellChanged++;
					if (indexCellChanged < cellsChanged.length) {
						getFromCellIndex(cellsChanged[indexCellChanged]);
						rowChanged = g_FCI.row;
						colChanged = g_FCI.col;
					}
				} else {
					if (0 === cmp) {
						this._broadcastNotifyListeners(cells[indexCell].listeners, notifyData);
					}
					indexCell++;
					if (indexCell < cells.length) {
						getFromCellIndex(cells[indexCell].cellIndex);
						row = g_FCI.row;
						col = g_FCI.col;
					}
				}
			}
		},
		_broadcastRangesByCells: function(sheetContainer, cells, notifyData) {
			if (0 === sheetContainer.rangesTopFrom.length || 0 === cells.length) {
				return;
			}
			let t = this;
			let rangesTopFrom = sheetContainer.rangesTopFrom;
			let rangesTopFromIndex = 0;
			let rangesTopTo = sheetContainer.rangesTopTo;
			let rangesTopToIndex = 0;
			let rangesBottomFrom = sheetContainer.rangesBottomFrom;
			let rangesBottomFromIndex = 0;
			let rangesBottomTo = sheetContainer.rangesBottomTo;
			let rangesBottomToIndex = 0;

			var indexCell = 0;
			let helper = new BroadcastHelper();
			var affected = [];
			var curY, curYCell = -1, elem;
			function skipRowHead(col) {
				while (curYCell === g_FCI.row && g_FCI.col < col) {
					indexCell++;
					if (indexCell < cells.length) {
						getFromCellIndex(cells[indexCell]);
					} else {
						break;
					}
				}
			}
			function skipRowTail() {
				while (curYCell === g_FCI.row) {
					indexCell++;
					if (indexCell < cells.length) {
						getFromCellIndex(cells[indexCell]);
					} else {
						break;
					}
				}
			}
			let processCellsBeforeCurY = function (_curY) {
				while (indexCell < cells.length) {
					getFromCellIndex(cells[indexCell]);
					if (g_FCI.row <= _curY) {
						if (g_FCI.row > curYCell) {
							helper.nextRow();
							curYCell = g_FCI.row;
							if (g_FCI.col < helper.from) {
								skipRowHead(helper.from);
								continue;
							} else if (helper.to < g_FCI.col) {
								skipRowTail();
								continue;
							}
						}
						if (g_FCI.col < helper.curFrom) {
							indexCell++;
						} else if (helper.curFrom <= g_FCI.col && g_FCI.col <= helper.curTo) {
							for (let i = 0; i < helper.curElemsLen; ++i) {
								let elem = helper.curElems[i];
								if (elem) {
									t._broadcastShared(elem, helper, affected, null, g_FCI.row, g_FCI.col);
								}
							}
							indexCell++;
						} else {
							//todo bin search?
							if (!helper.next()) {
								skipRowTail();
							}
						}
					} else {
						break;
					}
				}
				helper.finishDelete();
			}
			//scanline by Y
			while (rangesBottomFromIndex < rangesBottomFrom.length && indexCell < cells.length) {
				//next curY
				if (rangesTopFromIndex < rangesTopFrom.length) {
					curY = Math.min(rangesTopFrom[rangesTopFromIndex].bbox.r1, rangesBottomFrom[rangesBottomFromIndex].bbox.r2);
				} else {
					curY = rangesBottomFrom[rangesBottomFromIndex].bbox.r2;
				}
				//process cells before curY
				processCellsBeforeCurY(curY-1);
				if (indexCell < cells.length) {
					getFromCellIndex(cells[indexCell]);
				}
				//insert new ranges
				rangesTopFromIndex = helper.addInsertByRow(curY, g_FCI.row, rangesTopFrom, rangesTopFromIndex, helper.addInsertFrom);
				rangesTopToIndex = helper.addInsertByRow(curY, g_FCI.row, rangesTopTo, rangesTopToIndex, helper.addInsertTo);
				helper.finishInsert();
				//process cells curY
				processCellsBeforeCurY(curY);
				//delete ranges before move to new curY
				rangesBottomFromIndex = helper.addDeleteByRow(curY, rangesBottomFrom, rangesBottomFromIndex, helper.addDeleteFrom);
				rangesBottomToIndex = helper.addDeleteByRow(curY, rangesBottomTo, rangesBottomToIndex, helper.addDeleteTo);
				helper.finishDelete();
			}
			this._broadcastNotifyRanges(affected, notifyData);
		},
		_broadcastCellsByRanges: function(sheetContainer, rangesTopFrom, rangesTopTo, rangesBottomFrom, rangesBottomTo, notifyData) {
			if (0 === sheetContainer.cells.length || 0 === rangesTopFrom.length) {
				return;
			}
			let t = this;
			var cells = sheetContainer.cells;
			let rangesTopFromIndex = 0;
			let rangesTopToIndex = 0;
			let rangesBottomFromIndex = 0;
			let rangesBottomToIndex = 0;
			var indexCell = 0;
			let helper = new BroadcastHelper();
			var curY, curYCell = -1, elem, bbox;
			let skipRowHead = function(col) {
				while (curYCell === g_FCI.row && g_FCI.col < col) {
					indexCell++;
					if (indexCell < cells.length) {
						getFromCellIndex(cells[indexCell]);
					} else {
						break;
					}
				}
			}
			let skipRowTail = function() {
				while (curYCell === g_FCI.row) {
					indexCell++;
					if (indexCell < cells.length) {
						getFromCellIndex(cells[indexCell]);
					} else {
						break;
					}
				}
			}
			let processCellsBeforeCurY = function (_curY) {
				while (indexCell < cells.length) {
					getFromCellIndex(cells[indexCell].cellIndex);
					if (g_FCI.row <= _curY) {
						if (g_FCI.row > curYCell) {
							helper.nextRow();
							curYCell = g_FCI.row;
							if (g_FCI.col < helper.from) {
								skipRowHead(helper.from);
								continue;
							} else if (helper.to < g_FCI.col) {
								skipRowTail();
								continue;
							}
						}
						if (g_FCI.col < helper.curFrom) {
							indexCell++;
						} else if (helper.curFrom <= g_FCI.col && g_FCI.col <= helper.curTo) {
							t._broadcastNotifyListeners(cells[indexCell].listeners, notifyData);
							indexCell++;
						} else {
							//todo bin search?
							if (!helper.next()) {
								skipRowTail();
							}
						}
					} else {
						break;
					}
				}
			}
			//scanline by Y
			while (rangesBottomFromIndex < rangesBottomFrom.length && indexCell < cells.length) {
				//next curY
				if (rangesTopFromIndex < rangesTopFrom.length) {
					curY = Math.min(rangesTopFrom[rangesTopFromIndex].bbox.r1, rangesBottomFrom[rangesBottomFromIndex].bbox.r2);
				} else {
					curY = rangesBottomFrom[rangesBottomFromIndex].bbox.r2;
				}
				//process cells before curY
				processCellsBeforeCurY(curY - 1);
				if (indexCell < cells.length) {
					getFromCellIndex(cells[indexCell].cellIndex);
				}
				//insert new ranges
				rangesTopFromIndex = helper.addInsertByRow(curY, g_FCI.row, rangesTopFrom, rangesTopFromIndex, helper.addInsertFrom);
				rangesTopToIndex = helper.addInsertByRow(curY, g_FCI.row, rangesTopTo, rangesTopToIndex, helper.addInsertTo);
				helper.finishInsert();
				//process cells curY
				processCellsBeforeCurY(curY);
				//delete ranges before move to new curY
				rangesBottomFromIndex = helper.addDeleteByRow(curY, rangesBottomFrom, rangesBottomFromIndex, helper.addDeleteFrom);
				rangesBottomToIndex = helper.addDeleteByRow(curY, rangesBottomTo, rangesBottomToIndex, helper.addDeleteTo);
				helper.finishDelete();
			}
		},
		_broadcastRangesByRanges: function(sheetContainer, rangesTopFromChanged, rangesTopToChanged, rangesBottomFromChanged, rangesBottomToChanged, notifyData) {
			if (0 === sheetContainer.rangesTopFrom.length || 0 === rangesTopFromChanged.length) {
				return;
			}
			let t = this;
			let rangesTopFrom = sheetContainer.rangesTopFrom;
			let rangesTopFromIndex = 0;
			let rangesTopTo = sheetContainer.rangesTopTo;
			let rangesTopToIndex = 0;
			let rangesBottomFrom = sheetContainer.rangesBottomFrom;
			let rangesBottomFromIndex = 0;
			let rangesBottomTo = sheetContainer.rangesBottomTo;
			let rangesBottomToIndex = 0;
			let rangesTopFromChangedIndex = 0;
			let rangesTopToChangedIndex = 0;
			let rangesBottomFromChangedIndex = 0;
			let rangesBottomToChangedIndex = 0;
			let helper = new BroadcastHelper();
			let helperChanged = new BroadcastHelper();
			var affected = [];
			var curY;

			//scanline by Y
			while (rangesBottomFromIndex < rangesBottomFrom.length && rangesBottomFromChangedIndex < rangesBottomFromChanged.length) {
				//next curY
				curY = Math.min(rangesBottomFromChanged[rangesBottomFromChangedIndex].bbox.r2, rangesBottomFrom[rangesBottomFromIndex].bbox.r2);
				let curYTop = -1;
				if (rangesTopFromIndex < rangesTopFrom.length) {
					curYTop = rangesTopFrom[rangesTopFromIndex].bbox.r1;
					curY = Math.min(curY, curYTop);
				}
				let curYTopChanged = -1;
				if (rangesTopFromChangedIndex < rangesTopFromChanged.length) {
					curYTopChanged = rangesTopFromChanged[rangesTopFromChangedIndex].bbox.r1;
					curY = Math.min(curY, curYTopChanged);
				}
				let eventsCountOld = helper.getRowEventsCount();
				let eventsCountOldChanged = helperChanged.getRowEventsCount();
				//insert new ranges
				rangesTopFromIndex = helper.addInsertByRow(curY, eventsCountOldChanged ? curY : curYTopChanged, rangesTopFrom, rangesTopFromIndex, helper.addInsertFrom);
				rangesTopToIndex = helper.addInsertByRow(curY, eventsCountOldChanged ? curY : curYTopChanged, rangesTopTo, rangesTopToIndex, helper.addInsertTo);
				helper.finishInsert();
				rangesTopFromChangedIndex = helperChanged.addInsertByRow(curY, eventsCountOld ? curY : curYTop, rangesTopFromChanged, rangesTopFromChangedIndex, helper.addInsertFrom);
				rangesTopToChangedIndex = helperChanged.addInsertByRow(curY, eventsCountOld ? curY : curYTop, rangesTopToChanged, rangesTopToChangedIndex, helper.addInsertTo);
				helperChanged.finishInsert();
				//intersect
				if ((eventsCountOld !== helper.getRowEventsCount() || eventsCountOldChanged !== helperChanged.getRowEventsCount()) &&
					!(helper.to < helperChanged.from || helperChanged.to < helper.from))
				{
					helper.nextRow();
					helper.next();
					helperChanged.nextRow();
					helperChanged.next();
					while (true) {
						if (!(helperChanged.curTo < helper.curFrom || helper.curTo < helperChanged.curFrom)) {
							for (let i = 0; i < helper.curElemsLen; ++i) {
								let elem = helper.curElems[i];
								if (elem) {
									for (let j = 0; j < helperChanged.curElemsLen; ++j) {
										let elemChanged = helperChanged.curElems[j];
										if (elemChanged) {
											var intersect = elem.bbox.intersectionSimple(elemChanged.bbox);
											t._broadcastShared(elem, helper, affected, intersect, null, null);
										}
									}
								}
							}
						}
						if (helper.curTo < helperChanged.curTo) {
							if (!helper.next()) {
								break;
							}
						} else {
							if (!helperChanged.next()) {
								break;
							}
						}
					}
					helper.finishDelete();
				}
				//delete ranges before move to new curY
				rangesBottomFromIndex = helper.addDeleteByRow(curY, rangesBottomFrom, rangesBottomFromIndex, helper.addDeleteFrom);
				rangesBottomToIndex = helper.addDeleteByRow(curY, rangesBottomTo, rangesBottomToIndex, helper.addDeleteTo);
				helper.finishDelete();
				rangesBottomFromChangedIndex = helperChanged.addDeleteByRow(curY, rangesBottomFromChanged, rangesBottomFromChangedIndex, helper.addDeleteFrom);
				rangesBottomToChangedIndex = helperChanged.addDeleteByRow(curY, rangesBottomToChanged, rangesBottomToChangedIndex, helper.addDeleteTo);
				helperChanged.finishDelete();
			}
			this._broadcastNotifyRanges(affected, notifyData);
		},
		_broadcastShared: function(elem, helper, affected, intersect, row, col) {
			var sharedBroadcast = elem.sharedBroadcast;
			if (!sharedBroadcast) {
				if (elem.isActive) {
					affected.push(elem);
					elem.isActive = false;
					helper.addDeleteFrom(elem);
					helper.addDeleteToUnsorted(elem);
				}
			} else if (!sharedBroadcast.changedBBox ||
					!(intersect && sharedBroadcast.changedBBox.containsRange(intersect)) ||
					!(!intersect && sharedBroadcast.changedBBox.contains(col, row)))
			{
				if (!sharedBroadcast.changedBBox ||
					sharedBroadcast.changedBBox.isEqual(sharedBroadcast.prevChangedBBox)) {
					sharedBroadcast.recursion++;
					if (sharedBroadcast.recursion >= this.maxSharedRecursion) {
						sharedBroadcast.changedBBox = elem.bbox.clone();
					}
					affected.push(elem);
				}
				if (intersect) {
					if (sharedBroadcast.changedBBox) {
						sharedBroadcast.changedBBox.union2(intersect);
					} else {
						sharedBroadcast.changedBBox = intersect.clone();
					}
				} else {
					if (sharedBroadcast.changedBBox) {
						sharedBroadcast.changedBBox.union3(col, row);
					} else {
						sharedBroadcast.changedBBox = new Asc.Range(col, row, col, row);
					}
				}
				if (sharedBroadcast.changedBBox.isEqual(elem.bbox)) {
					if (elem.isActive) {
						elem.isActive = false;
						helper.addDeleteFrom(elem);
						helper.addDeleteToUnsorted(elem);
					}
				}
			}
		},
		_broadcastRangesByRangesIntersect: function(bbox, tree, output) {
			var intervals = tree.searchNodes(bbox.c1, bbox.c2);
			for(var i = 0; i < intervals.length; ++i){
				var interval = intervals[i];
				var elem = interval.data;
				if (elem) {
					var intersect = elem.bbox.intersectionSimple(bbox);
					var sharedBroadcast = elem.sharedBroadcast;
					if (!sharedBroadcast) {
						output.push(elem);
						elem.isActive = false;
						tree.remove(elem.bbox.c1, elem.bbox.c2, elem);
					} else if (!(sharedBroadcast.changedBBox && sharedBroadcast.changedBBox.containsRange(intersect))) {
						if (!sharedBroadcast.changedBBox ||
							sharedBroadcast.changedBBox.isEqual(sharedBroadcast.prevChangedBBox)) {
							sharedBroadcast.recursion++;
							if (sharedBroadcast.recursion >= this.maxSharedRecursion) {
								sharedBroadcast.changedBBox = elem.bbox.clone();
							}
							output.push(elem);
						}
						if (sharedBroadcast.changedBBox) {
							sharedBroadcast.changedBBox.union2(intersect);
						} else {
							sharedBroadcast.changedBBox = intersect;
						}
						if (sharedBroadcast.changedBBox.isEqual(elem.bbox)) {
							elem.isActive = false;
							tree.remove(elem.bbox.c1, elem.bbox.c2, elem);
						}
					}
				}
			}
		},
		_broadcastNotifyRanges: function(affected, notifyData) {
			var areaData = {bbox: null, changedBBox: null};
			for (var i = 0; i < affected.length; ++i) {
				var elem = affected[i];
				var shared = elem.sharedBroadcast;
				if (shared && shared.changedBBox) {
					areaData.bbox = elem.bbox;
					notifyData.areaData = areaData;
					if (!shared.prevChangedBBox) {
						areaData.changedBBox = shared.changedBBox;
						this._broadcastNotifyListeners(elem.listeners, notifyData);
					} else {
						var ranges = shared.prevChangedBBox.difference(shared.changedBBox);
						for (var j = 0; j < ranges.length; ++j) {
							areaData.changedBBox = ranges[j];
							this._broadcastNotifyListeners(elem.listeners, notifyData);
						}
					}
					notifyData.areaData = undefined;
					shared.prevChangedBBox = shared.changedBBox.clone();
				} else {
					this._broadcastNotifyListeners(elem.listeners, notifyData);
				}
			}
			this.tempGetByCells.push(affected);
		},
		_broadcastNotifyListeners: function(listeners, notifyData) {
			for (var listenerId in listeners) {
				listeners[listenerId].notify(notifyData);
			}
		}
	};

	function BroadcastHelper() {
		this.curElems = [];
		this.curElemsLen = 0;
		this.curElemsCount = 0;
		this.curFrom = -1;
		this.curTo = -1;
		this.from = -1;
		this.to = -1;

		//many arrays to avoid temp objects
		this.eventsFrom = [];
		this.eventsFromIndex = 0;
		this.eventsFromLen = 0;
		this.eventsTo = [];
		this.eventsToIndex = 0;
		this.eventsToLen = 0;

		this.insertFrom = [];
		this.insertFromLen = 0;
		this.insertTo = [];
		this.insertToLen = 0;

		this.deleteFrom = [];
		this.deleteFromLen = 0;
		this.deleteTo = [];
		this.deleteToLen = 0;
		this.deleteToUnsorted = [];
		this.deleteToUnsortedLen = 0;
	}
	BroadcastHelper.prototype.getRowEventsCount = function () {
		return this.eventsFromLen;
	};
	BroadcastHelper.prototype.nextRow = function () {
		this.curElemsLen = 0;
		this.curElemsCount = 0;
		this.curFrom = -1;
		this.curTo = -1;
		this.from = this.eventsFromLen > 0 ? this.eventsFrom[0].bbox.c1 : -1;
		this.to = this.eventsToLen > 0 ? this.eventsTo[this.eventsToLen - 1].bbox.c2 : -1;
		this.eventsFromIndex = 0;
		this.eventsToIndex = 0;
	};
	BroadcastHelper.prototype._getNextTo = function () {
		if (this.eventsToIndex < this.eventsToLen) {
			let elem = this.eventsTo[this.eventsToIndex];
			if (this.eventsFromIndex < this.eventsFromLen && this.eventsFrom[this.eventsFromIndex].bbox.c1 <= elem.bbox.c2) {
				elem = this.eventsFrom[this.eventsFromIndex];
				return elem.bbox.c1 - 1;
			} else {
				return elem.bbox.c2;
			}
		}
		return -1;
	}
	BroadcastHelper.prototype.next = function () {
		if (this.eventsToIndex < this.eventsToLen) {
			let elem = this.eventsTo[this.eventsToIndex];
			if (this.eventsFromIndex < this.eventsFromLen && this.eventsFrom[this.eventsFromIndex].bbox.c1 <= elem.bbox.c2) {
				elem = this.eventsFrom[this.eventsFromIndex];
				this.curFrom = elem.bbox.c1;
				while (this.eventsFromIndex < this.eventsFromLen && this.eventsFrom[this.eventsFromIndex].bbox.c1 === elem.bbox.c1) {
					elem = this.eventsFrom[this.eventsFromIndex];
					elem.index = this.curElemsLen;
					this.curElems[this.curElemsLen] = elem;
					this.curElemsCount++;
					this.curElemsLen++;
					this.eventsFromIndex++;
				}
				this.curTo = this._getNextTo();
				return true;
			} else {
				while (this.eventsToIndex < this.eventsToLen && this.eventsTo[this.eventsToIndex].bbox.c2 === elem.bbox.c2) {
					elem = this.eventsTo[this.eventsToIndex];
					//todo decrement this.curElemsLen too
					this.curElems[elem.index] = undefined;
					this.curElemsCount--;
					this.eventsToIndex++;
				}
				if (this.eventsToIndex < this.eventsToLen) {
					//skip empty space
					if (0 === this.curElemsCount) {
						return this.next();
					}
					this.curFrom = elem.bbox.c2 + 1;
					this.curTo = this._getNextTo();
					return true;
				} else {
					this.curFrom = this.curTo = -1;
					return false;
				}
			}
		}
		return false;
	};
	BroadcastHelper.prototype.cmpInsertFrom = function (a, b) {
		return Asc.Range.prototype.compareByLeftTop(a.bbox, b.bbox);
	};
	BroadcastHelper.prototype.cmpInsertTo = function (a, b) {
		return Asc.Range.prototype.compareByRightTop(a.bbox, b.bbox);
	};
	BroadcastHelper.prototype.cmpDeleteFrom = function (a, b) {
		return Asc.Range.prototype.compareByLeftBottom(a.bbox, b.bbox) || Asc.Range.prototype.compareByLeftTop(a.bbox, b.bbox);
	};
	BroadcastHelper.prototype.cmpDeleteTo = function (a, b) {
		return Asc.Range.prototype.compareByRightBottom(a.bbox, b.bbox) || Asc.Range.prototype.compareByRightTop(a.bbox, b.bbox);
	};
	BroadcastHelper.prototype._cmpC1 = function (a, b) {
		return a.bbox.c1 <= b.bbox.c1;
	};
	BroadcastHelper.prototype._cmpC2 = function (a, b) {
		return a.bbox.c2 <= b.bbox.c2;
	};
	BroadcastHelper.prototype.addInsertFrom = function (elem) {
		//elem.bbox & elem.id & elem.isActive
		this.insertFrom[this.insertFromLen] = elem;
		this.insertFromLen++;
	};
	BroadcastHelper.prototype.addInsertTo = function (elem) {
		//elem.bbox & elem.id
		this.insertTo[this.insertToLen] = elem;
		this.insertToLen++;
	};
	BroadcastHelper.prototype.addInsertByRow = function (row, minRow, sorted, sortedIndex, fInsert) {
		while (sortedIndex < sorted.length && row === sorted[sortedIndex].bbox.r1) {
			let elem = sorted[sortedIndex];
			sortedIndex++;
			if (false !== elem.isActive && minRow <= elem.bbox.r2) {
				fInsert.call(this, elem);
			}
		}
		return sortedIndex;
	}
	BroadcastHelper.prototype._mergeInsertSorted = function (events, eventsLen, ins, insLen, fCmp) {
		let eventsIndex = eventsLen - 1;
		let insIndex = insLen - 1;
		let eventsInsertIndex = eventsIndex + insLen;
		while (insIndex >= 0) {
			if (eventsIndex < 0 || fCmp(events[eventsIndex], ins[insIndex])) {
				events[eventsInsertIndex] = ins[insIndex];
				insIndex--;
			} else {
				events[eventsInsertIndex] = events[eventsIndex];
				eventsIndex--;
			}
			eventsInsertIndex--;
		}
		return eventsLen + insLen;
	};
	BroadcastHelper.prototype.finishInsert = function () {
		if (0 === this.insertFromLen && 0 === this.insertToLen) {
			this.nextRow();
			return;
		}
		this.eventsFromLen = this._mergeInsertSorted(this.eventsFrom, this.eventsFromLen, this.insertFrom, this.insertFromLen, this._cmpC1);
		this.insertFromLen = 0;
		this.eventsToLen = this._mergeInsertSorted(this.eventsTo, this.eventsToLen, this.insertTo, this.insertToLen, this._cmpC2);
		this.insertToLen = 0;

		this.nextRow();
	};
	BroadcastHelper.prototype.addDeleteFrom = function (elem) {
		//elem.bbox & elem.id
		this.deleteFrom[this.deleteFromLen] = elem;
		this.deleteFromLen++;
	};
	BroadcastHelper.prototype.addDeleteTo = function (elem) {
		this.deleteTo[this.deleteToLen] = elem;
		this.deleteToLen++;
	};
	BroadcastHelper.prototype.addDeleteToUnsorted = function (elem) {
		this.deleteToUnsorted[this.deleteToUnsortedLen] = elem;
		this.deleteToUnsortedLen++;
	};
	BroadcastHelper.prototype.addDeleteByRow = function (row, sorted, sortedIndex, fDelete) {
		while (sortedIndex < sorted.length && row === sorted[sortedIndex].bbox.r2) {
			let elem = sorted[sortedIndex];
			sortedIndex++;
			if (false !== elem.isActive) {
				fDelete.call(this, elem);
			}
		}
		return sortedIndex;
	}
	BroadcastHelper.prototype._mergeDeleteSorted = function (events, eventsLen, del, delLen) {
		//todo нельзя рассчитывать, что в addDelete будут приходить элементы из curElems и в порядке forEach(есть другие вызовы addDelete)
		//The forEach() method executes per each pair in insertion order.
		let eventsIndex = 0;
		let delIndex = 0;
		let eventsCurIndex = 0;
		while (eventsIndex < eventsLen && delIndex < delLen) {
			if (events[eventsIndex] !== del[delIndex]) {
				events[eventsCurIndex] = events[eventsIndex];
				eventsCurIndex++;
			} else {
				delIndex++;
			}
			eventsIndex++;
		}
		while (eventsIndex < eventsLen) {
			events[eventsCurIndex] = events[eventsIndex];
			eventsIndex++;
			eventsCurIndex++;
		}
		return eventsCurIndex;
	};
	BroadcastHelper.prototype.finishDelete = function () {
		if (0 === this.deleteFromLen && 0 === this.deleteToLen && 0 === this.deleteToUnsortedLen) {
			this.nextRow();
			return;
		}
		//deleteFrom,deleteTo - presorted; need sort deleteToUnsorted;
		this.deleteToUnsorted.sort(this.cmpDeleteTo);
		this.eventsFromLen = this._mergeDeleteSorted(this.eventsFrom, this.eventsFromLen, this.deleteFrom, this.deleteFromLen);
		this.deleteFromLen = 0;
		this.eventsToLen = this._mergeDeleteSorted(this.eventsTo, this.eventsToLen, this.deleteTo, this.deleteToLen);
		this.deleteToLen = 0;
		this.eventsToLen = this._mergeDeleteSorted(this.eventsTo, this.eventsToLen, this.deleteToUnsorted, this.deleteToUnsortedLen);
		this.deleteToUnsortedLen = 0;
		this.nextRow();
	};

	function AsyncFormulasManager(wb) {
		this.wb = wb;
		this.promises = null;
		this.promiseMap = null;
		this.calcProcess = null;
		this.aParserFormulas = null;

		this.replacedFormulas = null;

		this.endCallback = null;
	}

	AsyncFormulasManager.prototype.addPromise = function (val, reWrite) {
		if (!this.promises) {
			this.promises = [];
		}
		let isAdded = false;
		if (this.getPromiseByIndex(val.parserFormula._index, val.index)) {
			isAdded = true;
		}
		if (!isAdded) {
			this.promises.push(val);
			if (!this.promiseMap) {
				this.promiseMap = {};
			}
			if (!this.promiseMap[val.parserFormula._index]) {
				this.promiseMap[val.parserFormula._index] = {};
			}
			this.promiseMap[val.parserFormula._index][val.index] = this.promises.length - 1;
		} else if (reWrite) {
			this.promises[this.promiseMap[val.parserFormula._index][val.index]] = val;
		}
	};
	AsyncFormulasManager.prototype.getPromiseByIndex = function (indexParserFormula, stackIndex) {
		let indexArray =  this.promiseMap && this.promiseMap[indexParserFormula] && this.promiseMap[indexParserFormula][stackIndex];
		if (indexArray != null && this.promises) {
			return this.promises[indexArray];
		}
		return null;
	};
	AsyncFormulasManager.prototype.getPromise = function (index) {
		return this.promises && this.promises[index];
	};
	AsyncFormulasManager.prototype.getPromises = function () {
		return this.promises;
	};
	AsyncFormulasManager.prototype.isPromises = function () {
		return this.promises && this.promises.length;
	};
	AsyncFormulasManager.prototype.clearPromises = function () {
		this.promises = null;
		this.promiseMap = null;
	};
	AsyncFormulasManager.prototype.getReplacedFormulas = function () {
		return this.replacedFormulas;
	};
	AsyncFormulasManager.prototype.clearReplacedFormulas = function () {
		this.replacedFormulas = null;
	};
	AsyncFormulasManager.prototype.isReplacedFormulas = function () {
		return this.replacedFormulas && this.replacedFormulas.length;
	};
	AsyncFormulasManager.prototype.addReplacedFormula = function (val) {
		if (!this.replacedFormulas) {
			this.replacedFormulas = [];
		}
		this.replacedFormulas.push(val);
	};
	AsyncFormulasManager.prototype.setRecalculating = function (val) {
		this.calcProcess = val;
	};
	AsyncFormulasManager.prototype.isRecalculating = function () {
		return this.calcProcess;
	};
	AsyncFormulasManager.prototype.addPromiseParserFormula = function (elem) {
		if (!this.aParserFormulas) {
			this.aParserFormulas = [];
		}
		this.aParserFormulas.push(elem);
	};
	AsyncFormulasManager.prototype.cleanPromiseParserFormula = function () {
		if (this.aParserFormulas) {
			for (let i = 0; i < this.aParserFormulas.length; i++) {
				this.aParserFormulas[i].promiseResult = null;
			}
			this.aParserFormulas = null;
		}
	};
	AsyncFormulasManager.prototype.calculatePromises = function (callback) {
		let t = this;
		let promises = !this.isRecalculating() && this.getPromises();
		if (promises) {
			let doPromises = function (_promises) {
				let streamInfos = [];
				let fPromises = [];
				for (let i = 0; i < _promises.length; i++) {
					const fPromise = function () {
						return _promises[i].promise;
					};
					fPromises.push(fPromise);
				}
				const promiseIterator = new AscCommon.CPromiseGetterIterator(fPromises);
				promiseIterator.forEachValue(function (streamInfo) {
					streamInfos.push(streamInfo);
				}, function (rejectInfo) {
					streamInfos.push(null);
				}, function () {
					for (let i = 0; i < streamInfos.length; i++) {
						if (!_promises[i].parserFormula.promiseResult) {
							_promises[i].parserFormula.promiseResult = [];
						}
						let resStreamInfo;
						if (streamInfos[i] == null) {
							resStreamInfo = new cError(cErrorType.wrong_value_type);
						} else {
							resStreamInfo = _promises[i].callback(streamInfos[i]);
						}
						_promises[i].parserFormula.promiseResult[_promises[i].index] = resStreamInfo;
						if (_promises[i].parserFormula.parent) {
							t.wb.dependencyFormulas.addToChangedCell(_promises[i].parserFormula.parent);
						}
						t.addPromiseParserFormula(_promises[i].parserFormula);
					}
					//t.wb.buildDependency();
					//t.wb.sortDependency();
					/*t._foreachChanged(function (oCell) {
						oCell.setIsDirty(true);
						oCell && oCell._checkDirty();
					});*/

					t.clearPromises();
					t.setRecalculating(true);
					t.wb.dependencyFormulas.calcTree();
					t.wb.handlers && t.wb.handlers.trigger("drawWS");
					/*for (let i = 0; i < streamInfos.length; i++) {
						_promises[i].parserFormula.promiseResult = null;
					}*/
					t.setRecalculating(false);

					promises = !t.isRecalculating() && t.getPromises();
					if (promises) {
						doPromises(promises);
					} else {
						t.cleanPromiseParserFormula();
						t.endCallback && t.endCallback();
						callback && callback();
					}
				});
			};
			doPromises(promises);
		} else {
			//t.endCallback && t.endCallback()
		}
	};

	AsyncFormulasManager.prototype.calculateReplacedFormulas = function (callback) {
		let t = this;
		let replacedFormulas = /*!this.isRecalculating() &&*/ this.getReplacedFormulas();
		if (replacedFormulas) {
			let doReplacedFormulas = function (_replacedFormulas) {
				for (let i = 0; i < _replacedFormulas.length; i++) {
					t.wb.dependencyFormulas.addToChangedCell(_replacedFormulas[i].parent);
				}

				t.clearReplacedFormulas();
				t.setRecalculating(true);
				t.wb.dependencyFormulas.calcTree();
				t.wb.handlers && t.wb.handlers.trigger("drawWS");
				t.setRecalculating(false);
			};
			doReplacedFormulas(replacedFormulas);
		}
	};




	function ForwardTransformationFormula(elem, formula, parsed) {
		this.elem = elem;
		this.formula = formula;
		this.parsed = parsed;
	}
	ForwardTransformationFormula.prototype = {
		onFormulaEvent: function(type, eventData) {
			if (AscCommon.c_oNotifyParentType.ChangeFormula === type) {
				this.formula = eventData.assemble;
			}
		}
	};
	function angleFormatToInterface(val)
	{
		var nRes = 0;
		if(0 <= val && val <= 180)
			nRes = val <= 90 ? val : 90 - val;
		return nRes;
	}
	function angleFormatToInterface2(val)
	{
		if(g_nVerticalTextAngle == val)
			return val;
		else
			return angleFormatToInterface(val);
	}
	function angleInterfaceToFormat(val)
	{
		var nRes = val;
		if(-90 <= val && val <= 90)
		{
			if(val < 0)
				nRes = 90 - val;
		}
		else if(g_nVerticalTextAngle != val)
			nRes = 0;
		return nRes;
	}
	function getUniqueKeys(array) {
		var i, o = {};
		for (i = 0; i < array.length; ++i) {
			o[array[i].v] = o.hasOwnProperty(array[i].v);
		}
		return o;
	}
//-------------------------------------------------------------------------------------------------
		/**
	 * @constructor
	 */
	function Workbook(eventsHandlers, oApi, isMainLogicDocument){
		this.oApi = oApi;
		this.History = History;
		if (false !== isMainLogicDocument)
		{
			if (this.History)
				this.History.Set_LogicDocument(this);
		}
		this.MainDocument = false !== isMainLogicDocument;
		this.handlers = eventsHandlers;
		this.dependencyFormulas = new DependencyGraph(this);
		this.asyncFormulasManager = new AsyncFormulasManager(this);
		this.nActive = 0;
		this.showVerticalScroll = null;
		this.showHorizontalScroll = null;
		this.App = null;
		this.Core = null;

		AscFormat.ExecuteNoHistory(function () {
			this.CustomProperties = new AscCommon.CCustomProperties();
		}, this, [], true);

		this.theme = null;
		this.clrSchemeMap = null;

		this.CellStyles = new AscCommonExcel.CCellStyles();
		this.TableStyles = new Asc.CTableStyles();
		this.SlicerStyles = new Asc.CSlicerStyles();
		this.oStyleManager = new AscCommonExcel.StyleManager();
		this.sharedStrings = new AscCommonExcel.CSharedStrings();
		this.workbookFormulas = new AscCommonExcel.CWorkbookFormulas();
		this.loadCells = [];//to return one object when nested _getCell calls
		this.DrawingDocument = new AscCommon.CDrawingDocument();
		this.mathTrackHandler = new AscWord.CMathTrackHandler(this.DrawingDocument, oApi);

		this.aComments = [];	// Комментарии к документу
		this.aWorksheets = [];
		this.aWorksheetsById = {};
		this.aCollaborativeActions = [];
		this.bCollaborativeChanges = false;
		this.bUndoChanges = false;
		this.bRedoChanges = false;
		this.aCollaborativeChangeElements = [];
		this.externalReferences = [];
		this.calcPr = new AscCommonExcel.CCalcPr();
		this.workbookPr = new AscCommonExcel.CWorkbookPr();

		this.connections = null;

		this.wsHandlers = null;

		this.openErrors = [];

		this.maxDigitWidth = 0;
		this.paddingPlusBorder = 0;

		this.lastFindOptions = null;
		this.lastFindCells = {};
		this.oleSize = null;
		if (oApi && oApi.isEditOleMode) {
			this.oleSize = new AscCommonExcel.OleSizeSelectionRange(null, new Asc.Range(0, 0, 6, 9));
		}

		//при копировании листа с одного wb на другой необходимо менять в стеке
		// формул лист и книгу(на которые ссылаемся) - например у элементов cStrucTable
		//временно добавляю новый вставляемый лист, чтобы не передавать параметры через большое количество функций
		this.addingWorksheet = null;

		this.workbookProtection = null;
		this.fileSharing = null;

		this.customXmls = null;//[]
		this.oGoalSeek = null;
		this.oSolver = null;

		this.timelineCaches = [];
		this.TimelineStyles = null;

		this.metadata = null;
		//true - rightToLeft, false/null - leftToRight
		this.defaultDirection = null;

		this.externalReferenceHelper = new CExternalReferenceHelper(this);
	}
	Workbook.prototype.init=function(oReadResult, bNoBuildDep, bSnapshot){
		let tableCustomFunc = oReadResult.tableCustomFunc || {};
		let tableIds = oReadResult.tableIds || {};
		let sheetIds = oReadResult.sheetIds || {};
		let pivotCaches = oReadResult.pivotCacheDefinitions || {};
		if(this.nActive < 0)
			this.nActive = 0;
		if(this.nActive >= this.aWorksheets.length)
			this.nActive = this.aWorksheets.length - 1;

		var self = this;

		this.wsHandlers = new AscCommonExcel.asc_CHandlersList( /*handlers*/{
			"changeRefTablePart"   : function (table) {
				self.dependencyFormulas.changeTableRef(table);
			},
			"changeColumnTablePart": function ( tableName ) {
				self.dependencyFormulas.renameTableColumn( tableName );
			},
			"deleteColumnTablePart": function(tableName, deleted) {
				self.dependencyFormulas.delColumnTable(tableName, deleted);
				var wsActive = self.getActiveWs();
				if (wsActive) {
					wsActive.deleteSlicersByTableCol(tableName, deleted);
				}
			}, 'onFilterInfo' : function () {
				self.handlers && self.handlers.trigger("asc_onFilterInfo");
			}
		} );
		for(var i = 0, length = tableCustomFunc.length; i < length; ++i) {
			var elem = tableCustomFunc[i];
			elem.column.applyTotalRowFormula(elem.formula, elem.ws, !bNoBuildDep);
		}
		//ws
		this.forEach(function (ws) {
			ws.initPostOpen(self.wsHandlers, tableIds, sheetIds, pivotCaches, self.oNumFmtsOpen, self.dxfsOpen);
		});
		//timelinecache
		this.timelineCaches.forEach(function(elem){
			elem.initPostOpen(tableIds, sheetIds);
		});
		//show active if it hidden
		var wsActive = this.getActiveWs();
		if (wsActive && wsActive.getHidden()) {
			wsActive.setHidden(false);
		}

		if(!bNoBuildDep){
			this.dependencyFormulas.initOpen();
		}
		//external references
		this.externalReferences.forEach(function(elem){
			elem && elem.initPostOpen && elem.initPostOpen();
		});
		if (bSnapshot) {
			this.snapshot = this._getSnapshot();
		}

		if (this.MainDocument) {
			g_cCalcRecursion.initCalcProperties(this.calcPr);
		}

	};
	Workbook.prototype.Get_Api = function() {
		return this.oApi;
	};
	Workbook.prototype.Get_CollaborativeEditing = function() {
		return this.oApi.collaborativeEditing
	};
	Workbook.prototype.addImages = function (aImages, obj) {
		const oApi = Asc.editor;
		if (obj && undefined !== obj.id && aImages.length === 1 && aImages[0].Image) {
			const oController = oApi.getGraphicController();
			const oDrawingObjects = oApi.getDrawingObjects();
			const oPlaceholderTarget = AscCommon.g_oTableId.Get_ById(obj.id);
			if (oPlaceholderTarget) {
				if (oPlaceholderTarget.isObjectInSmartArt && oPlaceholderTarget.isObjectInSmartArt()) {
					const oSmartArtGroup = oPlaceholderTarget.group.getMainGroup();
					const oSmartArtId = oSmartArtGroup && oSmartArtGroup.Id;
					this.checkObjectsLock([oSmartArtId], function (bLock) {
						if (bLock) {
							AscCommon.History.Create_NewPoint();
							oController.resetSelection();
							oPlaceholderTarget.applyImagePlaceholderCallback && oPlaceholderTarget.applyImagePlaceholderCallback(aImages, obj);
							oController.selectObject(oSmartArtGroup, 0);
							oController.selection.groupSelection = oSmartArtGroup;
							oSmartArtGroup.selectObject(oPlaceholderTarget, 0);
							const oWS = oApi.wb.getWorksheet();
							if (oWS) {
								oWS.setSelectionShape(true);
							}
							oSmartArtGroup.addToRecalculate();
							oController.startRecalculate();
							if (oDrawingObjects) {
								oDrawingObjects.sendGraphicObjectProps();
							}
							oController.clearPreTrackObjects();
							oController.clearTrackObjects();
							oController.updateOverlay();
							oController.changeCurrentState(new AscFormat.NullState(oController));
							oController.updateSelectionState();
						}
					});
				}
			}
		} else if (obj && obj.callback && aImages.length === 1 && aImages[0].Image) {
			obj.callback(aImages[0]);
		}
	};
	Workbook.prototype.getOleSize = function () {
		return this.oleSize;
	};
	Workbook.prototype.setOleSize = function (oPr) {
		this.oleSize = oPr;
	};

	Workbook.prototype.preparePivotForSerialization=function(pivotCaches, isCopyPaste){
		var pivotCacheIndex = 0;
		this.forEach(function(ws) {
			for (var i = 0; i < ws.pivotTables.length; ++i) {
				var pivotTable = ws.pivotTables[i];
				if (isCopyPaste && !pivotTable.isInRange(isCopyPaste)) {
					continue;
				}
				if (pivotTable.cacheDefinition) {
					var pivotCache = pivotCaches[pivotTable.cacheDefinition.Get_Id()];
					if (!pivotCache) {
						pivotCache = {id: pivotCacheIndex++, cache: pivotTable.cacheDefinition};
						pivotCaches[pivotTable.cacheDefinition.Get_Id()] = pivotCache;
					}
					pivotTable.cacheId = pivotCache.id;
				}
			}
			for (var i = 0; i < ws.aSlicers.length; ++i) {
				var slicer = ws.aSlicers[i];
				if (isCopyPaste) {
					continue;
				}
				var cacheDefinition = slicer.getPivotCache();
				if (cacheDefinition) {
					var pivotCache = pivotCaches[cacheDefinition.Get_Id()];
					if (!pivotCache) {
						pivotCache = {id: pivotCacheIndex++, cache: cacheDefinition};
						pivotCaches[cacheDefinition.Get_Id()] = pivotCache;
					}
				}
			}
		}, isCopyPaste);
		return pivotCacheIndex;
	};
	Workbook.prototype.isChartOleObject = function () {
		return this.aWorksheets.length === 2;
	};
	Workbook.prototype.setCommonIndexObjectsFrom = function(wb) {
		this.oStyleManager = wb.oStyleManager;
		this.sharedStrings = wb.sharedStrings;
		this.workbookFormulas = wb.workbookFormulas;
	};
	Workbook.prototype.forEach = function (callback, isCopyPaste, writeTabSelected) {
		//if copy/paste - use only actve ws
		if ((isCopyPaste || isCopyPaste === false) && !writeTabSelected) {
			callback(this.getActiveWs(), this.getActive());
		} else {
			for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
				if (!writeTabSelected || (writeTabSelected && this.aWorksheets[i].isSelectedTab())) {
					callback(this.aWorksheets[i], i);
				}
			}
		}
	};
	Workbook.prototype.rebuildColors=function(){
		AscCommonExcel.g_oColorManager.rebuildColors();
		this.forEach(function (ws) {
			ws.rebuildColors();
		});
	};
	Workbook.prototype.getDefaultFont=function(){
		return g_oDefaultFormat.Font.getName();
	};
	Workbook.prototype.getDefaultSize=function(){
		return g_oDefaultFormat.Font.getSize();
	};
	Workbook.prototype.getActive=function(){
		return this.nActive;
	};
	Workbook.prototype.getActiveWs = function () {
		return this.getWorksheet(this.nActive);
	};
	Workbook.prototype.setActive=function(index){
		if(index >= 0 && index < this.aWorksheets.length){
			this.nActive = index;
			// Must clean find
			this.cleanFindResults();
			return true;
		}
		return false;
	};
	Workbook.prototype.setActiveById=function(sheetId){
		var ws = this.getWorksheetById(sheetId);
		if (ws) {
			this.setActive(ws.getIndex());
		}
		return false;
	};
	Workbook.prototype.getSheetIdByIndex = function(index) {
		var ws = this.getWorksheet(index);
		return ws ? ws.getId() : null;
	};
	Workbook.prototype.getWorksheet=function(index){
		//index 0-based
		if(index >= 0 && index < this.aWorksheets.length){
			return this.aWorksheets[index];
		}
		return null;
	};
	Workbook.prototype.getWorksheetById=function(id){
		return this.aWorksheetsById[id];
	};
	Workbook.prototype.getWorksheetByName=function(name, ignoreCaseSensitive){
		for(var i = 0; i < this.aWorksheets.length; i++) {
			if (ignoreCaseSensitive) {
				if(name && this.aWorksheets[i].getName().toLowerCase() == name.toLowerCase()){
					return this.aWorksheets[i];
				}
			} else if(this.aWorksheets[i].getName() == name){
				return this.aWorksheets[i];
			}
		}

		return null;
	};
	Workbook.prototype.getWorksheetIndexByName=function(name){
		for(var i = 0; i < this.aWorksheets.length; i++)
			if(this.aWorksheets[i].getName() == name){
				return i;
			}
		return null;
	};
	Workbook.prototype.getWorksheetCount=function(){
		return this.aWorksheets.length;
	};
	Workbook.prototype.createWorksheet=function(indexBefore, sName, sId){
		this.dependencyFormulas.lockRecal();
		AscCommon.History.Create_NewPoint();
		AscCommon.History.TurnOff();
		var wsActive = this.getActiveWs();
		var oNewWorksheet = new Worksheet(this, this.aWorksheets.length, sId);
		if (this.checkValidSheetName(sName))
			oNewWorksheet.sName = sName;
		oNewWorksheet.initPostOpen(this.wsHandlers, {});
		if(null != indexBefore && indexBefore >= 0 && indexBefore < this.aWorksheets.length)
			this.aWorksheets.splice(indexBefore, 0, oNewWorksheet);
		else
		{
			indexBefore = this.aWorksheets.length;
			this.aWorksheets.push(oNewWorksheet);
		}
		this.aWorksheetsById[oNewWorksheet.getId()] = oNewWorksheet;
		this._updateWorksheetIndexes(wsActive);
		AscCommon.History.TurnOn();
		this._insertWorksheetFormula(oNewWorksheet.index);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_SheetAdd, null, null, new UndoRedoData_SheetAdd(indexBefore, oNewWorksheet.getName(), null, oNewWorksheet.getId()));
		AscCommon.History.SetSheetUndo(wsActive.getId());
		AscCommon.History.SetSheetRedo(oNewWorksheet.getId());
		this.dependencyFormulas.unlockRecal();
		return oNewWorksheet;
	};
	Workbook.prototype.copyWorksheet=function(index, insertBefore, sName, sId, bFromRedo, tableNames, opt_sheet, opt_base64){
		//insertBefore - optional
		var renameParams;
		if(index >= 0 && index < this.aWorksheets.length){
			//buildRecalc вызываем чтобы пересчиталося cwf(может быть пустым если сделать сдвиг формул и скопировать лист)
			this.dependencyFormulas.buildDependency();
			AscCommon.History.TurnOff();
			var wsActive = this.getActiveWs();
			var wsFrom = opt_sheet ? opt_sheet : this.aWorksheets[index];
			var newSheet = new Worksheet(this, -1, sId);
			if(null != insertBefore && insertBefore >= 0 && insertBefore < this.aWorksheets.length){
				//помещаем новый sheet перед insertBefore
				this.aWorksheets.splice(insertBefore, 0, newSheet);
			}
			else{
				//помещаем новый sheet в конец
				this.aWorksheets.push(newSheet);
			}

			if(opt_sheet) {
				this.addingWorksheet = newSheet;
			}

			this.aWorksheetsById[newSheet.getId()] = newSheet;
			this._updateWorksheetIndexes(wsActive);
			if (this.handlers) {
				this.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.markModifiedSearch);
			}

			//copyFrom after sheet add because formula assemble dependce on sheet structure
			renameParams = newSheet.copyFrom(wsFrom, sName, tableNames);
			
			newSheet.copyFromFormulas(renameParams);

			newSheet.initPostOpen(this.wsHandlers, {}, {});
			AscCommon.History.TurnOn();

			this.dependencyFormulas.copyDefNameByWorksheet(wsFrom, newSheet, renameParams, opt_sheet);
			if (opt_sheet /*&& !bFromRedo*/) {
				this.dependencyFormulas.copyDefNameByWorkbook(wsFrom, newSheet, renameParams, opt_sheet);
			}
			//для формул. создаем копию this.cwf[this.Id] для нового листа.
			//newSheet._BuildDependencies(wsFrom.getCwf());

			//now insertBefore is index of inserted sheet
			this._insertWorksheetFormula(insertBefore);

			if (!tableNames) {
				tableNames = newSheet.getTableNames();
			}

			/*var opt_sheet_binary = null;
			if(!bFromRedo && opt_sheet) {
				opt_sheet_binary = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(wsFrom, null, null, true);
			}
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_SheetAdd, null, null, new UndoRedoData_SheetAdd(insertBefore, newSheet.getName(), wsFrom.getId(), newSheet.getId(), tableNames, opt_sheet_binary));*/

			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_SheetAdd, null, null, new UndoRedoData_SheetAdd(insertBefore, newSheet.getName(), wsFrom.getId(), newSheet.getId(), tableNames, opt_base64));
			AscCommon.History.SetSheetUndo(wsActive.getId());
			AscCommon.History.SetSheetRedo(newSheet.getId());
			newSheet.copyFromAfterInsert(wsFrom);
			if(!(bFromRedo === true))
			{
				wsFrom.copyObjects(newSheet, renameParams);
				var i;
				if (wsFrom.aNamedSheetViews) {
					for (i = 0; i < wsFrom.aNamedSheetViews.length; ++i) {
						newSheet.addNamedSheetView(wsFrom.aNamedSheetViews[i].clone(renameParams.tableNameMap));
					}
				}
				if (wsFrom.dataValidations && wsFrom.dataValidations.elems) {
					for (i = 0; i < wsFrom.dataValidations.elems.length; ++i) {
						newSheet.addDataValidation(wsFrom.dataValidations.elems[i].clone(), true);
					}
				}
			}
			this.sortDependency();

			if(opt_sheet) {
				this.addingWorksheet = null;
			}
		}
		return renameParams;
	};
	Workbook.prototype.insertWorksheet = function (index, sheet) {
		var wsActive = this.getActiveWs();
		if(null != index && index >= 0 && index < this.aWorksheets.length){
			//помещаем новый sheet перед insertBefore
			this.aWorksheets.splice(index, 0, sheet);
		}
		else{
			//помещаем новый sheet в конец
			this.aWorksheets.push(sheet);
		}
		this.aWorksheetsById[sheet.getId()] = sheet;
		this._updateWorksheetIndexes(wsActive);
		this._insertWorksheetFormula(index);
		this._insertTablePartsName(sheet);
		//восстанавливаем список ячеек с формулами для sheet
		sheet._BuildDependencies(sheet.getCwf());
		this.sortDependency();
	};
	Workbook.prototype._insertTablePartsName = function (sheet) {
		if(sheet && sheet.TableParts && sheet.TableParts.length)
		{
			for(var i = 0; i < sheet.TableParts.length; i++)
			{
				var tablePart = sheet.TableParts[i];
				this.dependencyFormulas.addTableName(sheet, tablePart);
				tablePart.buildDependencies();
			}
		}
	};
	Workbook.prototype._insertWorksheetFormula=function(index){
		if( index > 0 && index < this.aWorksheets.length ) {
			var oWsBefore = this.aWorksheets[index - 1];
			this.dependencyFormulas.changeSheet(this.dependencyFormulas.prepareChangeSheet(oWsBefore.getId(), {insert: index}));
		}
	};
	Workbook.prototype.replaceWorksheet=function(indexFrom, indexTo){
		if(indexFrom >= 0 && indexFrom < this.aWorksheets.length &&
			indexTo >= 0 && indexTo < this.aWorksheets.length){
			var wsActive = this.getActiveWs();
			var oWsFrom = this.aWorksheets[indexFrom];
			var tempW = {
				wF: oWsFrom,
				wFI: indexFrom,
				wTI: indexTo
			};
			//wTI index insert before
			if(tempW.wFI < tempW.wTI)
				tempW.wTI++;
			this.dependencyFormulas.lockRecal();
			var prepared = this.dependencyFormulas.prepareChangeSheet(oWsFrom.getId(), {replace: tempW}, null);
			//move sheets
			var movedSheet = this.aWorksheets.splice(indexFrom, 1);
			this.aWorksheets.splice(indexTo, 0, movedSheet[0]);
			this._updateWorksheetIndexes(wsActive);
			this.dependencyFormulas.changeSheet(prepared);

			this._insertWorksheetFormula(indexTo);

			if (!AscCommon.g_oTableId.Get_ById(oWsFrom.getId())) {
				//todo add in constructor with history
				AscFormat.ExecuteNoHistory(function () {
					AscCommon.g_oTableId.Add(oWsFrom, oWsFrom.getId())
				}, this, [], true);
			}
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_SheetRemove, null, null, new AscCommonExcel.UndoRedoData_SheetRemove(indexFrom, oWsFrom.getId(), oWsFrom));
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_SheetAdd, null, null, new UndoRedoData_SheetAdd(indexTo, oWsFrom.getName(), null, oWsFrom.getId(), null, null, oWsFrom.getId()));
			this.dependencyFormulas.unlockRecal();

			if (!this.bUndoChanges && !this.bRedoChanges) {
				this.handlers && this.handlers.trigger("updateCellWatches", true);
			}
		}
	};
	Workbook.prototype.findSheetNoHidden = function (nIndex) {
		var i, ws, oRes = null, bFound = false, countWorksheets = this.getWorksheetCount();
		for (i = nIndex; i < countWorksheets; ++i) {
			ws = this.getWorksheet(i);
			if (false === ws.getHidden()) {
				oRes = ws;
				bFound = true;
				break;
			}
		}
		// Не нашли справа, ищем слева от текущего
		if (!bFound) {
			for (i = nIndex - 1; i >= 0; --i) {
				ws = this.getWorksheet(i);
				if (false === ws.getHidden()) {
					oRes = ws;
					break;
				}
			}
		}
		return oRes;
	};
	Workbook.prototype.removeWorksheet=function(nIndex, outputParams){
		//проверяем останется ли хоть один нескрытый sheet
		var bEmpty = true;
		for(var i = 0, length = this.aWorksheets.length; i < length; ++i)
		{
			var worksheet = this.aWorksheets[i];
			if(false == worksheet.getHidden() && i != nIndex)
			{
				bEmpty = false;
				break;
			}
		}
		if(bEmpty)
			return -1;

		var removedSheet = this.getWorksheet(nIndex);
		if(removedSheet)
		{
			var removedSheetId = removedSheet.getId();
			this.dependencyFormulas.lockRecal();
			var prepared = this.dependencyFormulas.prepareRemoveSheet(removedSheetId, removedSheet.getTableNames());

			//remove timeline
			if (removedSheet.timelines) {
				for (let i = 0; i < removedSheet.timelines.length; i++) {
					this.onTimeSlicerDelete(removedSheet.timelines[i].name);
					this.onTimelineCacheDelete(removedSheet.timelines[i].name, true);
				}
			}

			//delete sheet
			var wsActive = this.getActiveWs();
			var oVisibleWs = null;
			this.aWorksheets.splice(nIndex, 1);
			delete this.aWorksheetsById[removedSheetId];

			if (nIndex == this.getActive()) {
				oVisibleWs = this.findSheetNoHidden(nIndex);
				if (null != oVisibleWs)
					wsActive = oVisibleWs;
			}
			if (!AscCommon.g_oTableId.Get_ById(removedSheet.getId())) {
				//todo add in constructor with history
				AscFormat.ExecuteNoHistory(function () {
					AscCommon.g_oTableId.Add(removedSheet, removedSheet.getId())
				}, this, [], true);
			}
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_SheetRemove, null, null, new AscCommonExcel.UndoRedoData_SheetRemove(nIndex, removedSheetId, removedSheet));
			if (null != oVisibleWs) {
				AscCommon.History.SetSheetUndo(removedSheetId);
				AscCommon.History.SetSheetRedo(wsActive.getId());
			}
			if(null != outputParams)
			{
				outputParams.sheet = removedSheet;
			}
			this._updateWorksheetIndexes(wsActive);
			this.dependencyFormulas.removeSheet(prepared);
			this.dependencyFormulas.unlockRecal();
			this.handlers && this.handlers.trigger("asc_onSheetDeleted", nIndex);
			this.handlers && this.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetRemove, nIndex);
			return wsActive.getIndex();
		}
		return -1;
	};
	Workbook.prototype._updateWorksheetIndexes = function (wsActive) {
		this.forEach(function (ws, index) {
			ws._setIndex(index);
		});
		if (null != wsActive) {
			this.setActive(wsActive.getIndex());
		}
	};
	Workbook.prototype.checkUniqueSheetName=function(name){
		var workbookSheetCount = this.getWorksheetCount();
		for (var i = 0; i < workbookSheetCount; i++){
			if (this.getWorksheet(i).getName() == name)
				return i;
		}
		return -1;
	};
	Workbook.prototype.checkValidSheetName=function(name){
		return (name && name.length <= g_nSheetNameMaxLength);
	};
	Workbook.prototype.getUniqueSheetNameFrom=function(name, bCopy){
		var nIndex = 1;
		var sNewName = "";
		var fGetPostfix = null;
		if(bCopy)
		{

			var result = /^(.*)\((\d)\)$/.exec(name);
			if(result)
			{
				fGetPostfix = function(nIndex){return "(" + nIndex +")";};
				name = result[1];
			}
			else
			{
				fGetPostfix = function(nIndex){return " (" + nIndex +")";};
				name = name;
			}
		}
		else
		{
			fGetPostfix = function(nIndex){return nIndex.toString();};
		}
		var workbookSheetCount = this.getWorksheetCount();
		while(nIndex < 10000)
		{
			var sPosfix = fGetPostfix(nIndex);
			sNewName = name + sPosfix;
			if(sNewName.length > g_nSheetNameMaxLength)
			{
				name = name.substring(0, g_nSheetNameMaxLength - sPosfix.length);
				sNewName = name + sPosfix;
			}
			var bUniqueName = true;
			for (var i = 0; i < workbookSheetCount; i++){
				if (this.getWorksheet(i).getName() == sNewName)
				{
					bUniqueName = false;
					break;
				}
			}
			if(bUniqueName)
				break;
			nIndex++;
		}
		return sNewName;
	};
	Workbook.prototype._generateFontMap=function(){
		var oFontMap = {
			"Arial"		: 1
		};
		var i;

		oFontMap[g_oDefaultFormat.Font.getName()] = 1;

		//theme
		if(null != this.theme)
			AscFormat.checkThemeFonts(oFontMap, this.theme.themeElements.fontScheme);
		//xfs
		for (i = 1; i <= g_StyleCache.getXfCount(); ++i) {
			var xf = g_StyleCache.getXf(i);
			if (xf.font) {
				oFontMap[xf.font.getName()] = 1;
			}
		}
		//sharedStrings
		this.sharedStrings.generateFontMap(oFontMap);

		this.forEach(function (ws) {
			ws.generateFontMap(oFontMap);
		});
		this.CellStyles.generateFontMap(oFontMap);

		return oFontMap;
	};
	Workbook.prototype.generateFontMap=function(){
		var oFontMap = this._generateFontMap();

		var aRes = [];
		for(var i in oFontMap)
			aRes.push(i);
		return aRes;
	};
	Workbook.prototype.generateFontMap2=function(){
		var oFontMap = this._generateFontMap();

		var aRes = [];
		for(var i in oFontMap)
			aRes.push(new AscFonts.CFont(i));
		AscFonts.FontPickerByCharacter.extendFonts(aRes);
		return aRes;
	};
	Workbook.prototype.getAllImageUrls = function(){
		var aImageUrls = [];
		this.forEach(function (ws) {
			ws.getAllImageUrls(aImageUrls);
		});
		return aImageUrls;
	};
	Workbook.prototype.reassignImageUrls = function(oImages){
		this.forEach(function (ws) {
			ws.reassignImageUrls(oImages);
		});
	};
	Workbook.prototype.calculate = function (type, sheetId) {
		let formulas;
		if (type === Asc.c_oAscCalculateType.All) {
			formulas = this.getAllFormulas();
			AscCommonExcel.executeInR1C1Mode(false, function () {
				for (let i = 0; i < formulas.length; ++i) {
					let formula = formulas[i];
					let caTemp = formula.ca;
					formula.removeDependencies();
					formula.setFormula(formula.getFormula());
					formula.parse();
					formula.ca = caTemp;
					formula.buildDependencies();
				}
			});
		} else if (type === Asc.c_oAscCalculateType.ActiveSheet) {
			formulas = [];
			let ws = sheetId !== undefined ? this.getWorksheetById(sheetId) : this.getActiveWs();
			ws.getAllFormulas(formulas);
			sheetId = ws.getId();
		} else {
			formulas = this.getAllFormulas();
		}
		this.dependencyFormulas.notifyChanged(formulas);
		this.dependencyFormulas.calcTree();
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_Calculate, sheetId,
			null, new AscCommonExcel.UndoRedoData_SingleProperty(type));
		AscCommon.History.EndTransaction();
	};
	Workbook.prototype.checkDefName = function (checkName, scope) {
		return this.dependencyFormulas.checkDefName(checkName, scope);
	};
	Workbook.prototype.getDefinedNamesWB = function (defNameListId, bLocale, excludeErrorRefNames) {
		return this.dependencyFormulas.getDefinedNamesWB(defNameListId, bLocale, excludeErrorRefNames);
	};
	Workbook.prototype.getDefinedNamesWS = function (sheetId) {
		return this.dependencyFormulas.getDefinedNamesWS(sheetId);
	};
	Workbook.prototype.addDefName = function (name, ref, sheetId, hidden, isTable) {
		return this.dependencyFormulas.addDefName(name, ref, sheetId, hidden, isTable);
	};
	Workbook.prototype.getDefinesNames = function ( name, sheetId ) {
		return this.dependencyFormulas.getDefNameByName( name, sheetId );
	};
	Workbook.prototype.getDefinedName = function(name) {
		var sheetId = this.getSheetIdByIndex(name.LocalSheetId);
		return this.dependencyFormulas.getDefNameByName(name.Name, sheetId);
	};
	Workbook.prototype.delDefinesNames = function ( defName ) {
		this.delDefinesNamesUndoRedo(this.getUndoDefName(defName));
	};
	Workbook.prototype.delDefinesNamesUndoRedo = function ( defName ) {
		this.dependencyFormulas.removeDefName( defName.sheetId, defName.name );
		this.dependencyFormulas.calcTree();
	};
	Workbook.prototype.editDefinesNames = function (oldName, newName) {
		return this.editDefinesNamesUndoRedo(this.getUndoDefName(oldName), this.getUndoDefName(newName));
	};
	Workbook.prototype.editDefinesNamesUndoRedo = function (oldName, newName) {
		var res = this.dependencyFormulas.editDefinesNames(oldName, newName);
		this.dependencyFormulas.calcTree();
		return res;
	};
	Workbook.prototype.findDefinesNames = function ( ref, sheetId, bLocale ) {
		return this.dependencyFormulas.getDefNameByRef( ref, sheetId, bLocale );
	};
	Workbook.prototype.definesNamesContains = function ( col, row, sheetId, bLocale ) {
		return this.dependencyFormulas.getDefNameByCellInside( col, row, sheetId, bLocale );
	};
	Workbook.prototype.unlockDefName = function(){
		this.dependencyFormulas.unlockDefName();
	};
	Workbook.prototype.unlockCurrentDefName = function(name, sheetId){
		this.dependencyFormulas.unlockCurrentDefName(name, sheetId);
	};
	Workbook.prototype.checkDefNameLock = function(){
		return this.dependencyFormulas.checkDefNameLock();
	};
	Workbook.prototype._SerializeHistoryItem = function (oMemory, item) {
		if (this.oApi.binaryChanges) {
			if (!item.LocalChange) {
				let nLen = oMemory.WriteWithLen(this, function(){
					item.Serialize(oMemory, this.oApi.collaborativeEditing);
				});
				return oMemory.GetDataUint8(0, nLen);
			}
			return;
		} else {
			if (!item.LocalChange) {
				oMemory.Seek(0);
				item.Serialize(oMemory, this.oApi.collaborativeEditing);
				var nLen = oMemory.GetCurPosition();
				if (nLen > 0)
					return nLen + ";" + oMemory.GetBase64Memory2(0, nLen);
			}
			return;
		}
	};
	Workbook.prototype._SerializeHistoryItem2 = function (oMemory, item) {
		if (!item.LocalChange) {
			item.Serialize(oMemory, this.oApi.collaborativeEditing);
		}
	};
	Workbook.prototype._SerializeHistory = function(oMemory, item, aPointChanges) {
		let data = this._SerializeHistoryItem(oMemory, item);
		if (data) {
			aPointChanges.push(data);
		}
	};
	Workbook.prototype.SerializeHistory = function(){
		var aResData = [];
		var aResSerializable = [];
		//соединяем изменения, которые были до приема данных с теми, что получились после.

		var t, j, length2;

		// Пересчитываем позиции
		AscCommon.CollaborativeEditing.Refresh_DCChanges();

		var aActions = this.aCollaborativeActions.concat(AscCommon.History.GetSerializeArray());
		if(aActions.length > 0)
		{
			var oMemory = new AscCommon.CMemory();
			for(var i = 0, length = aActions.length; i < length; ++i)
			{
				let items = aActions[i];
				for (j = 0, length2 = items.length; j < length2; ++j) {
					var item = items[j];
					// Пересчитываем позиции
					if (item.Data && item.Data.applyCollaborative) {
						//не делаем копию oData, а сдвигаем в ней, потому что все равно после сериализации изменения потруться
						if (item.Data.applyCollaborative(item.SheetId, this.oApi.collaborativeEditing)) {
							AscCommon.History.Refresh_SpreadsheetChanges(item);
						}
					}
					if (!item.LocalChange) {
						let data = AscCommon.CCollaborativeChanges.ToBase64(item.Binary.Pos, item.Binary.Len);
						aResData.push(data);
						aResSerializable.push(AscCommon.History.Item_ToSerializable(item));
					}
				}
			}
			this.aCollaborativeActions = [];
			this.snapshot = this._getSnapshot();
		}
		return [aResData, aResSerializable];
	};
	Workbook.prototype._getSnapshot = function() {
		var wb = new Workbook(new AscCommonExcel.asc_CHandlersList(), this.oApi, false);
		wb.dependencyFormulas = this.dependencyFormulas.getSnapshot(wb);
		this.forEach(function (ws) {
			ws = ws.getSnapshot(wb);
			wb.aWorksheets.push(ws);
			wb.aWorksheetsById[ws.getId()] = ws;
		});
		//init trigger
		wb.init({}, true, false);
		return wb;
	};
	Workbook.prototype.getAllFormulas = function(needReturnCellProps) {
		var res = [];
		this.dependencyFormulas.getAllFormulas(res);
		this.forEach(function (ws) {
			ws.getAllFormulas(res, needReturnCellProps);
		});
		return res;
	};
	Workbook.prototype.addCustomFunctionToChanged = function(isNotUpdate) {
		var res = null;
		let t = this;
		this.dependencyFormulas.forEachFormula(function (fP) {
			if (fP && fP.unknownOrCustomFunction && ((fP.ca && isNotUpdate) || !isNotUpdate)) {
				fP.isParsed = false;
				fP.outStack = [];
				fP.parse();
				if (/*DefName*/fP.parent.getNodeId) {
					t.dependencyFormulas.addToChangedDefName(fP.parent);
				}
				res = true;
			}
		});
		this.forEach(function (ws) {
			ws.forEachFormula(function (fP) {
				if (fP && fP.unknownOrCustomFunction && ((fP.ca && isNotUpdate) || !isNotUpdate)) {
					fP.isParsed = false;
					fP.outStack = [];
					fP.parse();
					if (/*CCellWithFormula*/fP.parent.nCol != null) {
						t.dependencyFormulas.addToChangedCell(fP.parent);
					}
					res = true;
				}
			});
		});
		return res;
	};
	Workbook.prototype._forwardTransformation = function(wbSnapshot, changesMine, changesTheir) {
		AscCommon.History.TurnOff();
		//first mine changes to resolve conflict sheet names
		var res1 = this._forwardTransformationGetTransform(wbSnapshot, changesTheir, changesMine);
		var res2 = this._forwardTransformationGetTransform(wbSnapshot, changesMine, changesTheir);
		//modify formulas at the end - to prevent negative effect during tranformation
		var i, elem, elemWrap;
		for (i = 0; i < res1.modify.length; ++i) {
			elemWrap = res1.modify[i];
			elem = elemWrap.elem;
			elem.oClass.forwardTransformationSet(elem.nActionType, elem.oData, elem.nSheetId, elemWrap);
		}
		for (i = 0; i < res2.modify.length; ++i) {
			elemWrap = res2.modify[i];
			elem = elemWrap.elem;
			elem.oClass.forwardTransformationSet(elem.nActionType, elem.oData, elem.nSheetId, elemWrap);
		}
		//rename current wb
		for (var oldName in res1.renameSheet) {
			var ws = this.getWorksheetByName(oldName);
			if (ws) {
				ws.setName(res1.renameSheet[oldName]);
			}
		}
		AscCommon.History.TurnOn();
	};
	Workbook.prototype._forwardTransformationGetTransform = function(wbSnapshot, changesMaster, changesModify) {
		var res = {modify: [], renameSheet: {}};
		var changesMasterSelected = [];
		var i, elem;
		if (changesModify.length > 0) {
			//select useful master changes
			for ( i = 0; i < changesMaster.length; ++i) {
				elem = changesMaster[i];
				if (elem.oClass && elem.oClass.forwardTransformationIsAffect &&
					elem.oClass.forwardTransformationIsAffect(elem.nActionType)) {
					changesMasterSelected.push(elem);
				}
			}
		}
		if (changesMasterSelected.length > 0 && changesModify.length > 0) {
			var wbSnapshotCur = wbSnapshot._getSnapshot();
			var formulas = [];
			for (i = 0; i < changesModify.length; ++i) {
				elem = changesModify[i];
				var renameRes = null;
				if (elem.oClass && elem.oClass.forwardTransformationGet) {
					var getRes = elem.oClass.forwardTransformationGet(elem.nActionType, elem.oData, elem.nSheetId);
					if (getRes && getRes.formula) {
						//inserted formulas
						formulas.push(new ForwardTransformationFormula(elem, getRes.formula, null));
					}
					if (getRes && getRes.name) {
						//add/rename sheet
						//get getUniqueSheetNameFrom if need
						renameRes = this._forwardTransformationRenameStart(wbSnapshotCur._getSnapshot(),
																		   changesMasterSelected, getRes);
					}
				}
				if (elem.oClass && elem.oClass.forwardTransformationIsAffect &&
					elem.oClass.forwardTransformationIsAffect(elem.nActionType)) {
					if (formulas.length > 0) {
						//modify all formulas before apply next change
						this._forwardTransformationFormula(wbSnapshotCur._getSnapshot(), formulas,
														   changesMasterSelected, res);
						formulas = [];
					}
					//apply useful mine change
					elem.oClass.Redo(elem.nActionType, elem.oData, elem.nSheetId, wbSnapshotCur);
				}
				if (renameRes) {
					this._forwardTransformationRenameEnd(renameRes, res.renameSheet, getRes, elem);
				}
			}
			this._forwardTransformationFormula(wbSnapshotCur, formulas, changesMasterSelected, res);
		}
		return res;
	};
	Workbook.prototype._forwardTransformationRenameStart = function(wbSnapshot, changes, getRes) {
		var res = {newName: null};
		for (var i = 0; i < changes.length; ++i) {
			var elem = changes[i];
			elem.oClass.Redo(elem.nActionType, elem.oData, elem.nSheetId, wbSnapshot);
		}
		if (-1 != wbSnapshot.checkUniqueSheetName(getRes.name)) {
			res.newName = wbSnapshot.getUniqueSheetNameFrom(getRes.name, true);
		}
		return res;
	};
	Workbook.prototype._forwardTransformationRenameEnd = function(renameRes, renameSheet, getRes, elemCur) {
		var isChange = false;
		if (getRes.from) {
			var renameCur = renameSheet[getRes.from];
			if (renameCur) {
				//no need rename next formulas
				delete renameSheet[getRes.from];
				getRes.from = renameCur;
				isChange = true;
			}
		}
		if (renameRes && renameRes.newName) {
			renameSheet[getRes.name] = renameRes.newName;
			getRes.name = renameRes.newName;
			isChange = true;
		}
		//apply immediately cause it is conflict
		if (isChange && elemCur.oClass.forwardTransformationSet) {
			elemCur.oClass.forwardTransformationSet(elemCur.nActionType, elemCur.oData, elemCur.nSheetId, getRes);
		}
	};
	Workbook.prototype._forwardTransformationFormula = function(wbSnapshot, formulas, changes, res) {
		if (formulas.length > 0) {
			var i, elem, ftFormula, ws;
			//parse formulas
			for (i = 0; i < formulas.length; ++i) {
				ftFormula = formulas[i];
				ws = wbSnapshot.getWorksheetById(ftFormula.elem.nSheetId);
				if(!ws) {
					ws = new AscCommonExcel.Worksheet(wbSnapshot, -1);
				}
				if (ws) {
					ftFormula.parsed = new parserFormula(ftFormula.formula, ftFormula, ws);
					ftFormula.parsed.parse();
					ftFormula.parsed.buildDependencies();
				}
			}
			//rename sheet first to prevent name conflict
			for (var oldName in res.renameSheet) {
				ws = wbSnapshot.getWorksheetByName(oldName);
				if (ws) {
					ws.setName(res.renameSheet[oldName]);
				}
			}
			//apply useful theirs changes
			for (i = 0; i < changes.length; ++i) {
				elem = changes[i];
				elem.oClass.Redo(elem.nActionType, elem.oData, elem.nSheetId, wbSnapshot);
			}
			//assemble
			for (i = 0; i < formulas.length; ++i) {
				ftFormula = formulas[i];
				if (ftFormula.parsed) {
					ftFormula.parsed.removeDependencies();
					res.modify.push(ftFormula);
				}
			}
		}
	};
	Workbook.prototype._DeserializeUndoRedoElems = function(aChanges){
		var oThis = this;
		var aUndoRedoElems = [];
		if (this.oApi.binaryChanges) {
			for (let i = 0; i < aChanges.length; ++i) {
				let sChange = aChanges[i];
				var stream = new AscCommon.FT_Stream2(sChange, sChange.length);
				stream.Seek(4);
				stream.Seek2(4);
				var item = new UndoRedoItemSerializable();
				item.Deserialize(stream);
				if (!oThis.needSkipChange(item)) {
					aUndoRedoElems.push(item);
				}
			}
		} else {
			var dstLen = 0;
			var aIndexes = [], i, length = aChanges.length, sChange;
			for(i = 0; i < length; ++i)
			{
				sChange = aChanges[i];
				var nIndex = sChange.indexOf(";");
				if (-1 != nIndex) {
					dstLen += parseInt(sChange.substring(0, nIndex));
					nIndex++;
				}
				aIndexes.push(nIndex);
			}
			var pointer = g_memory.Alloc(dstLen);
			var stream = new AscCommon.FT_Stream2(pointer.data, dstLen);
			stream.obj = pointer.obj;
			var nCurOffset = 0;
			for (i = 0; i < length; ++i) {
				sChange = aChanges[i];
				if (sChange.startsWith("0;")) {
					continue;
				}
				var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
				nCurOffset = oBinaryFileReader.getbase64DecodedData2(sChange, aIndexes[i], stream, nCurOffset);
				var item = new UndoRedoItemSerializable();
				item.Deserialize(stream);
				if (!oThis.needSkipChange(item)) {
					aUndoRedoElems.push(item);
				}
			}
		}
		return aUndoRedoElems;
	}
	Workbook.prototype.DeserializeHistory = function(aChanges, fCallback, oColor){
		var oThis = this;
		//сохраняем те изменения, которые были до приема данных, потому что дальше undo/redo будет очищено
		this.aCollaborativeActions = this.aCollaborativeActions.concat(AscCommon.History.GetSerializeArray());
		if(aChanges.length > 0)
		{
			this.bCollaborativeChanges = true;
			//собираем общую длину
			var i, length = aChanges.length;

			var aUndoRedoElems = this._DeserializeUndoRedoElems(aChanges);
			var wsViews = window["Asc"]["editor"].wb && window["Asc"]["editor"].wb.wsViews;
			if(oThis.oApi.collaborativeEditing.getFast()){
				AscCommon.CollaborativeEditing.Clear_DocumentPositions();
			}
			if(wsViews) {
				for (var i in wsViews) {
					if (isRealObject(wsViews[i]) && isRealObject(wsViews[i].objectRender) &&
						isRealObject(wsViews[i].objectRender.controller)) {
						wsViews[i].endEditChart();
						if (oThis.oApi.collaborativeEditing.getFast()) {
							var oState = wsViews[i].objectRender.saveStateBeforeLoadChanges();
							if (oState) {
								if (oState.Pos) {
									AscCommon.CollaborativeEditing.Add_DocumentPosition(oState.Pos);
								}
								if (oState.StartPos) {
									AscCommon.CollaborativeEditing.Add_DocumentPosition(oState.StartPos);
								}
								if (oState.EndPos) {
									AscCommon.CollaborativeEditing.Add_DocumentPosition(oState.EndPos);
								}
							}
						}
						wsViews[i].objectRender.controller.resetSelection();
					}
				}
			}
			oFormulaLocaleInfo.Parse = false;
			oFormulaLocaleInfo.DigitSep = false;
			AscFonts.IsCheckSymbols = true;
			AscCommon.History.Clear();
			AscCommon.History.TurnOff();
			var history = new AscCommon.CHistory();
			history.Set_LogicDocument(this);
			history.Create_NewPoint();

			history.SetSelection(null);
			history.SetSelectionRedo(null);
			var oRedoObjectParam = new AscCommonExcel.RedoObjectParam();
			history.UndoRedoPrepare(oRedoObjectParam, false);
			var changesMine = [].concat.apply([], oThis.aCollaborativeActions);
			oThis._forwardTransformation(oThis.snapshot, changesMine, aUndoRedoElems);
			oThis.cleanCollaborativeFilterObj();
			for (var i = 0, length = aUndoRedoElems.length; i < length; ++i)
			{
				var item = aUndoRedoElems[i];
				if ((null != item.oClass || (item.oData && typeof item.oData.sChangedObjectId === "string")) && null != item.nActionType) {
					if (window["NATIVE_EDITOR_ENJINE"] === true && window["native"]["CheckNextChange"]) {
						if (!window["native"]["CheckNextChange"]())
							break;
					}
					if (true === oThis.oApi.collaborativeEditing.private_AddOverallChange(item))
					{
						// TODO if(g_oUndoRedoGraphicObjects == item.oClass && item.oData.drawingData)
						//     item.oData.drawingData.bCollaborativeChanges = true;
						AscCommonExcel.executeInR1C1Mode(false, function () {
							history.RedoAdd(oRedoObjectParam, item.oClass, item.nActionType, item.nSheetId, item.oRange, item.oData, false, item);
						});
						if (oThis.oApi.VersionHistory && item.nSheetId && item.GetChangedRange) {
							let changedRange = item.nSheetId && item.GetChangedRange();
							if (changedRange) {
								oThis.oApi.wb.addToHistoryChangedRanges(item.nSheetId, changedRange, oColor);
							}
						}

						if (item.GetClass() && item.GetClass().SetIsRecalculated && item.IsNeedRecalculate())
							item.GetClass().SetIsRecalculated(false);

					}
				}
			}
			AscFonts.IsCheckSymbols = false;

			var oFontMap = this._generateFontMap();
			window["Asc"]["editor"]._loadFonts(oFontMap, function(){
				if(oThis.oApi.collaborativeEditing.getFast()){
					if(wsViews) {
						for(var i in wsViews){
							if(isRealObject(wsViews[i]) && isRealObject(wsViews[i].objectRender) && isRealObject(wsViews[i].objectRender.controller)){
								var oState = wsViews[i].objectRender.getStateBeforeLoadChanges();
								if(oState){
									if (oState.Pos)
										AscCommon.CollaborativeEditing.Update_DocumentPosition(oState.Pos);
									if (oState.StartPos)
										AscCommon.CollaborativeEditing.Update_DocumentPosition(oState.StartPos);
									if (oState.EndPos)
										AscCommon.CollaborativeEditing.Update_DocumentPosition(oState.EndPos);
								}
								wsViews[i].objectRender.loadStateAfterLoadChanges();
							}
						}
					}
				}
				oFormulaLocaleInfo.Parse = true;
				oFormulaLocaleInfo.DigitSep = true;
				history.UndoRedoEnd(null, oRedoObjectParam, false);
				AscCommon.History.TurnOn();
				oThis.bCollaborativeChanges = false;
				//make snapshot for faormulas
				oThis.snapshot = oThis._getSnapshot();
				if(null != fCallback)
					fCallback();
			});
		} else if(null != fCallback) {
			fCallback();
		}
	};
	Workbook.prototype.DeserializeHistoryNative = function(oRedoObjectParam, data, isFull){
		if(null != data)
		{
			this.bCollaborativeChanges = true;

			if(null == oRedoObjectParam)
			{
				if(window["Asc"]["editor"].wb) {
					var wsViews = window["Asc"]["editor"].wb.wsViews;
					for (var i in wsViews) {
						if (isRealObject(wsViews[i]) && isRealObject(wsViews[i].objectRender) &&
							isRealObject(wsViews[i].objectRender.controller)) {
							wsViews[i].endEditChart();
							wsViews[i].objectRender.controller.resetSelection();
						}
					}
				}

				AscCommon.History.Clear();
				AscCommon.History.Create_NewPoint();
				AscCommon.History.SetSelection(null);
				AscCommon.History.SetSelectionRedo(null);
				oRedoObjectParam = new AscCommonExcel.RedoObjectParam();
				AscCommon.History.UndoRedoPrepare(oRedoObjectParam, false);
			}

			var stream = new AscCommon.FT_Stream2(data, data.length);
			stream.obj = null;
			// Применяем изменения, пока они есть
			var _count = stream.GetLong();
			var _pos = 4;
			for (var i = 0; i < _count; i++)
			{
				if (window["NATIVE_EDITOR_ENJINE"] === true && window["native"]["CheckNextChange"])
				{
					if (!window["native"]["CheckNextChange"]())
						break;
				}

				var _len = stream.GetLong();

				_pos += 4;
				stream.size = _pos + _len;
				stream.Seek(_pos);
				stream.Seek2(_pos);

				var item = new UndoRedoItemSerializable();
				item.Deserialize(stream);
				if ((null != item.oClass || (item.oData && typeof item.oData.sChangedObjectId === "string")) && null != item.nActionType){
					AscCommonExcel.executeInR1C1Mode(false, function () {
						AscCommon.History.RedoAdd(oRedoObjectParam, item.oClass, item.nActionType, item.nSheetId, item.oRange, item.oData, false, item);
					});
				}

				_pos += _len;
				stream.Seek2(_pos);
				stream.size = data.length;
			}

			if(isFull){
				AscCommon.History.UndoRedoEnd(null, oRedoObjectParam, false);
				AscCommon.History.Clear();
				oRedoObjectParam = null;
			}
			this.bCollaborativeChanges = false;
		}
		return oRedoObjectParam;
	};
	Workbook.prototype.needSkipChange = function(change){
		var res = false;
		var aSkipChanges = [[AscCommonExcel.g_oUndoRedoWorkbook.getClassType(), AscCH.historyitem_Workbook_Date1904]];

		//при десериализации пропускаем изменение, если такое же есть в списке текущих у данного юзера
		var isNeededChange = function (_change, _actionType, _classType) {
			if (_change && _change.oClass) {
				if (_actionType === undefined || _classType === undefined) {
					for (var j = 0; j < aSkipChanges.length; j++) {
						if (aSkipChanges[j][0] === _change.oClass.getClassType() && aSkipChanges[j][1] === _change.nActionType) {
							return true;
						}
					}
				} else {
					if (_change && _change.oClass && _change.oClass.getClassType() === _classType && _change.nActionType === _actionType) {
						return true;
					}
				}
			}

			return false;
		};

		if (this.aCollaborativeActions && this.aCollaborativeActions.length) {
			if (isNeededChange(change)) {
				var actionType = change.nActionType;
				var classType = change.oClass && change.oClass.getClassType();
				for (var i = 0, length = this.aCollaborativeActions.length; i < length; ++i) {
					if (isNeededChange(change, actionType, classType)) {
						return true;
					}
				}
			}
		}

		return res;
	};
	Workbook.prototype.getTableRangeForFormula = function(name, objectParam){
		var res = null;
		for(var i = 0, length = this.aWorksheets.length; i < length; ++i)
		{
			var ws = this.aWorksheets[i];
			res = ws.getTableRangeForFormula(name, objectParam);
			if(res !== null){
				res = {wsID:ws.getId(),range:res};
				break;
			}
		}
		return res;
	};
	Workbook.prototype.getTableIndexColumnByName = function(tableName, columnName){
		var res = null;
		for(var i = 0, length = this.aWorksheets.length; i < length; ++i)
		{
			var ws = this.aWorksheets[i];
			res = ws.getTableIndexColumnByName(tableName, columnName);
			//получаем имя через getTableNameColumnByIndex поскольку tableName приходит в том виде
			//в котором набрал пользователь, те регистр может быть произвольным
			//todo для того, чтобы два раза не бежать по колонкам можно сделать функцию которая возвращает полную информацию о колонке
			//
			if(res !== null){
				res = {wsID:ws.getId(), index: res, name: ws.getTableNameColumnByIndex(tableName, res)};
				break;
			}
		}
		return res;
	};
	Workbook.prototype.getTableNameColumnByIndex = function(tableName, columnIndex){
		var res = null;
		for(var i = 0, length = this.aWorksheets.length; i < length; ++i)
		{
			var ws = this.aWorksheets[i];
			res = ws.getTableNameColumnByIndex(tableName, columnIndex);
			if(res !== null){
				res = {wsID:ws.getId(), columnName: res};
				break;
			}
		}
		return res;
	};
	Workbook.prototype.getTableByName = function(tableName, getSheetIndex){
		var res = null;
		for (var i = 0, length = this.aWorksheets.length; i < length; ++i) {
			var ws = this.aWorksheets[i];
			res = ws.getTableByName(tableName);
			if (res !== null) {
				return getSheetIndex ? {index: i, table: res} : res;
			}
		}
		return res;
	};
	Workbook.prototype.getTableByNameAndSheet = function(tableName, wsID){
		var res = null;
		var ws = this.getWorksheetById(wsID);
		return ws.getTableByName(tableName);
	};
	Workbook.prototype.updateSparklineCache = function (sheet, ranges) {
		this.forEach(function (ws) {
			ws.updateSparklineCache(sheet, ranges);
		});
	};
	Workbook.prototype.sortDependency = function () {
		this.dependencyFormulas.calcTree();
	};
	/**
	 * Вычисляет ширину столбца для заданного количества символов
	 * @param {Number} count  Количество символов
	 * @returns {Number}      Ширина столбца в символах
	 */
	Workbook.prototype.charCountToModelColWidth = function (count) {
		if (count <= 0) {
			return 0;
		}
		return Asc.floor((count * this.maxDigitWidth + this.paddingPlusBorder) / this.maxDigitWidth * 256) / 256;
	};
	/**
	 * Вычисляет ширину столбца в px
	 * @param {Number} mcw  Количество символов
	 * @returns {Number}    Ширина столбца в px
	 */
	Workbook.prototype.modelColWidthToColWidth = function (mcw) {
		return Asc.floor(((256 * mcw + Asc.floor(128 / this.maxDigitWidth)) / 256) * this.maxDigitWidth);
	};
	/**
	 * Вычисляет количество символов по ширине столбца
	 * @param {Number} w  Ширина столбца в px
	 * @returns {Number}  Количество символов
	 */
	Workbook.prototype.colWidthToCharCount = function (w) {
		var pxInOneCharacter = this.maxDigitWidth + this.paddingPlusBorder;
		// Когда меньше 1 символа, то просто считаем по пропорции относительно размера 1-го символа
		return w < pxInOneCharacter ?
			(1 - Asc.floor(100 * (pxInOneCharacter - w) / pxInOneCharacter + 0.49999) / 100) :
			Asc.floor((w - this.paddingPlusBorder) / this.maxDigitWidth * 100 + 0.5) / 100;
	};
	Workbook.prototype.getUndoDefName = function(ascName) {
		if (!ascName) {
			return ascName;
		}
		var sheetId = this.getSheetIdByIndex(ascName.LocalSheetId);
		return new UndoRedoData_DefinedNames(ascName.Name, ascName.Ref, sheetId, ascName.type, ascName.isXLNM, ascName.Hidden);
	};
	Workbook.prototype.changeColorScheme = function (sSchemeName) {
		var scheme = AscCommon.getColorSchemeByName(sSchemeName);
		if (!scheme) {
			scheme = this.theme.getExtraClrScheme(sSchemeName);
		}
		if(!scheme)
		{
			return;
		}
		AscCommon.History.Create_NewPoint();
		//не делаем Duplicate потому что предполагаем что схема не будет менять частями, а только обьектом целиком.
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_ChangeColorScheme, null,
			null, new AscCommonExcel.UndoRedoData_ClrScheme(this.theme.themeElements.clrScheme, scheme));
		this.theme.changeColorScheme(scheme);
		this.rebuildColors();
		return true;
	};
	Workbook.prototype.changeColorSchemeByIdx = function (nIdx) {

		var scheme = this.oApi.getColorSchemeByIdx(nIdx);
		if(!scheme) {
			return;
		}
		AscCommon.History.Create_NewPoint();
		//не делаем Duplicate потому что предполагаем что схема не будет менять частями, а только обьектом целиком.
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_ChangeColorScheme, null,
			null, new AscCommonExcel.UndoRedoData_ClrScheme(this.theme.themeElements.clrScheme, scheme));
		this.theme.changeColorScheme(scheme);
		this.rebuildColors();
		return true;
	};
	// ----- Search -----
	Workbook.prototype.cleanFindResults = function () {
		this.lastFindOptions = null;
		this.lastFindCells = {};
	};
	Workbook.prototype.findCellText = function (options, searchEngine) {
		var ws = this.getActiveWs();
		var result = ws.findCellText(options, searchEngine), result2 = null;
		if (Asc.c_oAscSearchBy.Workbook === options.scanOnOnlySheet) {
			// Search on workbook
			var key = result && (result.col + "-" + result.row);
			if (!key || (options.isEqual(this.lastFindOptions) && this.lastFindCells[key])) {
				// Мы уже находили данную ячейку, попробуем на другом листе
				var i, active = this.getActive(), start = 0, end = this.getWorksheetCount();
				var inc = options.scanForward ? +1 : -1;
				for (i = active + inc; i < end && i >= start; i += inc) {
					ws = this.getWorksheet(i);
					if (ws.getHidden()) {
						continue;
					}
					result2 = ws.findCellText(options, searchEngine);
					if (result2) {
						break;
					}
				}
				if (!result2 || searchEngine) {
					// Мы дошли до конца или начала (в зависимости от направления, теперь пойдем до активного)
					if (options.scanForward) {
						i = 0;
						end = active;
					} else {
						i = end - 1;
						start = active + 1;
						inc *= -1;
					}

					for (; i < end && i >= start; i += inc) {
						ws = this.getWorksheet(i);
						if (ws.getHidden()) {
							continue;
						}
						result2 = ws.findCellText(options, searchEngine);
						if (result2 && !searchEngine) {
							break;
						}
					}
				}

				if (result2 && !searchEngine) {
					this.handlers && this.handlers.trigger('undoRedoHideSheet', i);
					key = result2.col + "-" + result2.row;
				}
			}

			if (key && !searchEngine) {
				this.lastFindOptions = options.clone();
				this.lastFindCells[key] = true;
			}
		}
		if (searchEngine) {
			return;
		}

		if (!result2 && !result) {
			this.cleanFindResults();
		}
		return result2 || result;
	};
	//Comments
	Workbook.prototype.getComment = function (id) {
		if (id) {
			var sheet;
			for (var i = 0; i < this.aWorksheets.length; ++i) {
				sheet = this.aWorksheets[i];
				for (var j = 0; j < sheet.aComments.length; ++j) {
					if (id === sheet.aComments[j].asc_getGuid()) {
						return sheet.aComments[j];
					}
				}
			}
		}
		return null;
	};
	Workbook.prototype.getPivotTableByName = function(name) {
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var pivot = this.aWorksheets[i].getPivotTableByName(name);
			if (pivot) {
				return pivot;
			}
		}
		return null;
	};
	Workbook.prototype.getPivotTableById = function(id) {
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var pivot = this.aWorksheets[i].getPivotTableById(id);
			if (pivot) {
				return pivot;
			}
		}
		return null;
	};
	Workbook.prototype.getPivotTablesByCacheId = function(pivotCacheId) {
		var res = [];
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var pivotTables = this.aWorksheets[i].getPivotTablesByCacheId(pivotCacheId);
			res = res.concat(pivotTables);
		}
		return res;
	};
	Workbook.prototype.getPivotTablesByCache = function(cache) {
		var res = [];
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var caches = this.aWorksheets[i].getPivotTablesByCache(cache);
			res = res.concat(caches);
		}
		return res;
	};
	Workbook.prototype.getPivotCacheByDataLocation = function(dataLocation) {
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var cache = this.aWorksheets[i].getPivotCacheByDataLocation(dataLocation);
			if (cache) {
				return cache;
			}
		}
		return null;
	};
	Workbook.prototype.getPivotCacheById = function(pivotCacheId, pivotCachesOpen) {
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var cache = this.aWorksheets[i].getPivotCacheById(pivotCacheId, pivotCachesOpen);
			if (cache) {
				return cache;
			}
		}
		return null;
	};
	Workbook.prototype.getSlicerCacheByName = function (name) {
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var cache = this.aWorksheets[i].getSlicerCacheByName(name);
			if (cache) {
				return cache;
			}
		}
		return null;
	};
	Workbook.prototype.getSlicerCacheByCacheName = function (name) {
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var cache = this.aWorksheets[i].getSlicerCacheByCacheName(name);
			if (cache) {
				return cache;
			}
		}
		return null;
	};
	Workbook.prototype.getSlicerCacheByPivotTableFld = function (sheetId, pivotName, fld) {
		var slicerCaches = this.getSlicerCachesByPivotTable(sheetId, pivotName);
		for (var i = 0; i < slicerCaches.length; ++i) {
			if (slicerCaches[i].getPivotFieldIndex() === fld) {
				return slicerCaches[i];
			}
		}
		return null;
	};
	Workbook.prototype.getSlicerCachesByPivotTable = function (sheetId, pivotName) {
		var res = [];
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var caches = this.aWorksheets[i].getSlicerCachesByPivotTable(sheetId, pivotName);
			res = res.concat(caches);
		}
		return res;
	};
	Workbook.prototype.getSlicerCachesByPivotCacheId = function (pivotCacheId) {
		var res = [];
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var caches = this.aWorksheets[i].getSlicerCachesByPivotCacheId(pivotCacheId);
			res = res.concat(caches);
		}
		return res;
	};
	Workbook.prototype.getSlicerStyle = function (name) {
		var slicer = this.getSlicerByName(name);
		if (slicer) {
			return slicer.style;
		}
		return null;
	};
	Workbook.prototype.getSlicerByName = function (name) {
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var slicer = this.aWorksheets[i].getSlicerByName(name);
			if (slicer) {
				return slicer;
			}
		}
		return null;
	};
	Workbook.prototype.getSlicerViewByName = function (name) {
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var slicer = this.aWorksheets[i].getSlicerViewByName(name);
			if (slicer) {
				return slicer;
			}
		}
		return null;
	};
	Workbook.prototype.getSlicersByCacheName = function (name) {
		var res = [];
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			var slicers = this.aWorksheets[i].getSlicersByCacheName(name);
			if (slicers) {
				res = res.concat(slicers);
			}
		}
		return res.length ? res : null;
	};
	Workbook.prototype.getTimelineCachesByPivotTable = function (sheetId, pivotName) {
		var res = [];
		if (this.timelineCaches) {
			for (var i = 0, l = this.timelineCaches.length; i < l; ++i) {
				if (this.timelineCaches[i].isCachePivotTable(sheetId, pivotName)) {
					res.push(this.timelineCaches[i]);
				}
			}
		}
		return res;
	};
	Workbook.prototype.getDrawingDocument = function() {
		return this.DrawingDocument;
	};
	Workbook.prototype.onSlicerUpdate = function(sName) {
		if(AscCommon.isFileBuild()) {
			return false;
		}
		var bRet = false;
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			bRet = bRet || this.aWorksheets[i].onSlicerUpdate(sName);
		}
		return bRet;
	};
	Workbook.prototype.onSlicerDelete = function(sName) {
		if(AscCommon.isFileBuild()) {
			return false;
		}
		var bRet = false;
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			bRet = bRet || this.aWorksheets[i].onSlicerDelete(sName);
		}
		return bRet;
	};
	Workbook.prototype.onSlicerLock = function(sName, bLock) {
		if(AscCommon.isFileBuild()) {
			return;
		}
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			this.aWorksheets[i].onSlicerLock(sName, bLock);
		}
	};
	Workbook.prototype.onSlicerChangeName = function(sName, sNewName) {
		if(AscCommon.isFileBuild()) {
			return;
		}
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			this.aWorksheets[i].onSlicerChangeName(sName, sNewName);
		}
	};
	Workbook.prototype.deleteSlicer = function(sName) {
		var bRet = false;
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			bRet = bRet || this.aWorksheets[i].deleteSlicer(sName);
		}
		return bRet;
	};
	Workbook.prototype.getSlicersByTableName = function (val) {
		var res = null;
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			var wsSlicers = this.aWorksheets[i].getSlicersByTableName(val);
			if (wsSlicers) {
				if (!res) {
					res = [];
				}
				res = res.concat(wsSlicers);
			}
		}
		return res;
	};
	Workbook.prototype.slicersUpdateAfterChangeTable = function (name) {
		var slicers = this.getSlicersByTableName(name);
		if (slicers) {
			for (var j = 0; j < slicers.length; j++) {
				this.onSlicerUpdate(slicers[j].name);
			}
		}
	};
	Workbook.prototype.getSlicersByPivotTable = function (sheetId, pivotName) {
		var res = [];
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			var slicers = this.aWorksheets[i].getSlicersByPivotTable(sheetId, pivotName);
			res = res.concat(slicers);
		}
		return res;
	};
	Workbook.prototype.slicersUpdateAfterChangePivotTable = function (sheetId, pivotName) {
		var slicers = this.getSlicersByPivotTable(sheetId, pivotName);
		for (var j = 0; j < slicers.length; j++) {
			this.onSlicerUpdate(slicers[j].name);
		}
	};
	Workbook.prototype.deleteSlicersByPivotTable = function (sheetId, pivotName) {
		var wb = this;
		var slicerCaches = wb.getSlicerCachesByPivotTable(sheetId, pivotName);
		slicerCaches.forEach(function (slicerCache) {
			slicerCache.deletePivotTable(sheetId, pivotName);
			if (0 === slicerCache.getPivotTablesCount()) {
				var slicers = wb.getSlicersByCacheName(slicerCache.name);
				if (slicers) {
					slicers.forEach(function (slicer) {
						wb.deleteSlicer(slicer.name);
					});
				}
			}
		});
	};
	Workbook.prototype.deleteSlicersByPivotTableAndFields = function (sheetId, pivotName, cacheFieldsToDelete) {
		var wb = this;
		var slicerCaches = wb.getSlicerCachesByPivotTable(sheetId, pivotName);
		slicerCaches.forEach(function (slicerCache) {
			if (cacheFieldsToDelete[slicerCache.getSourceName()]) {
				slicerCache.deletePivotTable(sheetId, pivotName);
				var slicers = wb.getSlicersByCacheName(slicerCache.name);
				if (slicers) {
					slicers.forEach(function (slicer) {
						wb.deleteSlicer(slicer.name);
					});
				}
			}
		});
	};

	Workbook.prototype.deleteSlicersByTable = function (tableName, doDelDefName) {
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		for(var i = 0; i < this.aWorksheets.length; ++i) {
			var wsSlicers = this.aWorksheets[i].getSlicersByTableName(tableName);
			if (wsSlicers) {
				for (var j = 0; j < wsSlicers.length; j++) {
					this.aWorksheets[i].deleteSlicer(wsSlicers[j].name, doDelDefName);
				}
			}
		}

		AscCommon.History.EndTransaction();
	};

	Workbook.prototype.deleteTimelinesByPivotTable = function (sheetId, pivotName) {
		var wb = this;
		var timelineCaches = wb.getTimelineCachesByPivotTable(sheetId, pivotName);
		timelineCaches.forEach(function (timelineCache) {
			wb.onTimeSlicerDelete(timelineCache.sourceName);
			wb.onTimelineCacheDelete(timelineCache.sourceName, true);
			wb.onTimelinesDelete(timelineCache.sourceName, true);
		});
	};

	Workbook.prototype.deleteTimelinesByCacheSourceName = function (name) {
		this.forEach(function (ws) {
			ws.deleteTimelinesByCacheSourceName(name);
		});
	};

	Workbook.prototype.onTimeSlicerDelete = function(sName) {
		if(AscCommon.isFileBuild()) {
			return false;
		}
		var bRet = false;
		for(let i = 0; i < this.aWorksheets.length; ++i) {
			bRet = bRet || this.aWorksheets[i].onTimeSlicerDelete(sName);
		}

		return bRet;
	};
	Workbook.prototype.onTimelineCacheDelete = function (sName, addToHistory) {
		if (this.timelineCaches) {
			for(let i = 0; i < this.timelineCaches.length; ++i) {
				if (this.timelineCaches[i].sourceName === sName) {
					if (addToHistory) {
						AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_TimelineCacheDelete,
							null, null, new UndoRedoData_FromTo(this.timelineCaches[i], null));
					}
					this.timelineCaches.splice(i, 1);
					i--;
					break;
				}
			}
		}
	};
	Workbook.prototype.onTimelinesDelete = function(sName, addToHistory) {
		if(AscCommon.isFileBuild()) {
			return false;
		}
		var bRet = false;
		for(let i = 0; i < this.aWorksheets.length; ++i) {
			bRet = bRet || this.aWorksheets[i].onTimelinesDelete(sName, addToHistory);
		}

		return bRet;
	};
	Workbook.prototype.handleDrawings = function (fCallback) {
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			this.aWorksheets[i].handleDrawings(fCallback);
		}
	};
	Workbook.prototype.checkObjectsLock = function (aId, fCallback) {
		if(aId.length > 0) {
			if(Asc && Asc.editor) {
				Asc.editor.checkObjectsLock(aId, fCallback);
			}
			else {
				fCallback(true, true);
			}
		}
	};
	Workbook.prototype.handleChartsOnWorksheetsRemove = function (aWorksheets) {
		if(!AscCommon.History.CanAddChanges()) {
			return;
		}
		var aRefsToChange = [];
		var aId = [];
		var aRanges = [];
		var aNames = [];
		var oWorksheet;
		for(var nWS = 0; nWS < aWorksheets.length; ++nWS) {
			oWorksheet = aWorksheets[nWS];
			aRanges.push(new AscCommonExcel.Range(oWorksheet, 0, 0, gc_nMaxRow0, gc_nMaxCol0));
			aNames.push(parserHelp.getEscapeSheetName(oWorksheet.sName));
		}
		this.handleDrawings(function(oDrawing) {
			if(oDrawing.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
				var nPrevLength = aRefsToChange.length;
				oDrawing.collectIntersectionRefs(aRanges, aRefsToChange);
				if(aRefsToChange.length > nPrevLength) {
					aId.push(oDrawing.Get_Id());
				}
			}
		});
		this.checkObjectsLock(aId, function(bNoLock) {
			if(bNoLock) {
				for(var nRef = 0; nRef < aRefsToChange.length; ++nRef) {
					aRefsToChange[nRef].handleRemoveWorksheets(aNames);
				}
				if(Asc.editor && Asc.editor.wb) {
					Asc.editor.wb.recalculateDrawingObjects(null, false);
				}
			}
		});
	};
	Workbook.prototype.getChartsWithSheetData = function(oWorksheet) {
		var aRefsToChange = [];
		var aId = [];
		var aCharts = [];
		var aRanges = [new AscCommonExcel.Range(oWorksheet, 0, 0, gc_nMaxRow0, gc_nMaxCol0)];
		this.handleDrawings(function(oDrawing) {
			if(oDrawing.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
				var nPrevLength = aRefsToChange.length;
				oDrawing.clearDataRefs();
				oDrawing.collectIntersectionRefs(aRanges, aRefsToChange);
				oDrawing.clearDataRefs();
				if(aRefsToChange.length > nPrevLength) {
					aCharts.push(oDrawing);
					aId.push(oDrawing.Get_Id());
				}
			}
		});
		return {refs: aRefsToChange, ids: aId, charts: aCharts};

	};
	Workbook.prototype.getChartSheetRenameData = function (oWorksheet, sOldName) {
		var sOldSheetName = oWorksheet.sName;
		oWorksheet.sName = sOldName;
		var oData = this.getChartsWithSheetData(oWorksheet);
		oWorksheet.sName = sOldSheetName;
		return oData;
	};
	Workbook.prototype.changeSheetNameInRefs = function (aRefsToChange, sOldName, sNewName) {
		if(aRefsToChange.length > 0) {
			var sNewNameEscaped = parserHelp.getEscapeSheetName(sNewName);
			var sOldNameEscaped = parserHelp.getEscapeSheetName(sOldName);
			for(var nRef = 0; nRef < aRefsToChange.length; ++nRef) {
				aRefsToChange[nRef].handleOnChangeSheetName(sOldNameEscaped, sNewNameEscaped);
			}
		}
	};
	Workbook.prototype.handleChartsOnChangeSheetName = function (oWorksheet, sOldName, sNewName) {
		var oData = this.getChartSheetRenameData(oWorksheet, sOldName);
		this.changeSheetNameInRefs(oData.refs, sOldName, sNewName);
	};
	Workbook.prototype.handleChartsOnMoveRange = function (oRangeFrom, oRangeTo, isInsertCol) {
		if(!AscCommon.History.CanAddChanges()) {
			return;
		}
		var aRefsToResize = [];
		var aRefsToReplace = [];
		var aId = [];
		this.handleDrawings(function(oDrawing) {
			if(oDrawing.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
				var nPrevLength = aRefsToReplace.length;
				let delta;
				let range;
				if (isInsertCol) {
					delta = oRangeTo.bbox.c2 - oRangeTo.bbox.c1 - 1;
					range = oRangeFrom.clone()
					range.bbox.c2 = oRangeFrom.bbox.c1 + delta;
				} else {
					delta = oRangeTo.bbox.r2 - oRangeTo.bbox.r1 - 1;
					range = oRangeFrom.clone()
					range.bbox.r2 = oRangeFrom.bbox.r1 + delta;
				}
				oDrawing.collectRefsInsideRange(oRangeFrom, aRefsToReplace);
				if (typeof isInsertCol != "undefined") {
					oDrawing.collectRefsInsideRangeForInsertColRow(range, aRefsToResize, isInsertCol);
				}
				for(var nRef = 0; nRef < aRefsToResize.length; ++nRef) {
					if (aRefsToReplace.filter(function(val){return val.Id === aRefsToResize[nRef].Id}).length === 0)
						aRefsToReplace.push({ ref: aRefsToResize[nRef], resize: true, Id: aRefsToResize[nRef].Id })
				}

				if(aRefsToReplace.length > nPrevLength) {
					aId.push(oDrawing.Get_Id());
				}
			}
		});
		this.checkObjectsLock(aId, function(bNoLock) {
			if(bNoLock) {
				for(var nRef = 0; nRef < aRefsToReplace.length; ++nRef) {
					if (aRefsToReplace[nRef].resize)
						aRefsToReplace[nRef].ref.handleOnMoveRange(oRangeFrom, oRangeTo, true);
					else
						aRefsToReplace[nRef].handleOnMoveRange(oRangeFrom, oRangeTo, false);
				}
				if(Asc.editor && Asc.editor.wb) {
					Asc.editor.wb.recalculateDrawingObjects(null, false);
				}
			}
		});
	};
	Workbook.prototype.convertEquationToMath = function (oEquation, isAll) {
		var aId = [];
		var aEquations = [];
		if(isAll) {
			this.handleDrawings(function(oDrawing) {
				oDrawing.collectEquations3(aEquations);
			});

		}
		else {
			aEquations.push(oEquation);
		}
		if(aEquations.length > 0) {
			var nEquation;
			var oObjectForCheck;
			for(nEquation = 0; nEquation < aEquations.length; ++nEquation) {
				oObjectForCheck = (aEquations[nEquation].getMainGroup() || aEquations[nEquation]);
				aId.push(oObjectForCheck.Get_Id());
			}
			this.checkObjectsLock(aId, function(bNoLock) {
				if(bNoLock) {
					AscCommon.History.Create_NewPoint(0);
					for(nEquation = 0; nEquation < aEquations.length; ++nEquation) {
						aEquations[nEquation].replaceToMath();
					}
					if(Asc.editor && Asc.editor.wb) {
						Asc.editor.wb.recalculateDrawingObjects(null, false);
					}
				}
			});
		}
	};

	Workbook.prototype.cleanCollaborativeFilterObj = function () {
		if (!Asc.CT_NamedSheetView.prototype.asc_getName) {
			return;
		}
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			this.aWorksheets[i].autoFilters.cleanCollaborativeObj();
		}
	};

	Workbook.prototype.checkCorrectTables = function () {
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			this.aWorksheets[i].checkCorrectTables();
		}
	};

	Workbook.prototype.getRulesByType = function(type, id, needClone) {
		var range, sheet;
		var rules = null;
		switch (type) {
			case Asc.c_oAscSelectionForCFType.selection:
				sheet = this.getActiveWs();
				// ToDo multiselect
				range = sheet.selectionRange.getLast();
				break;
			case Asc.c_oAscSelectionForCFType.worksheet:
				sheet = id ? this.getWorksheet(id) : this.getActiveWs();
				break;
			case Asc.c_oAscSelectionForCFType.table:
				var oTable;
				if (id) {
					oTable = this.getTableByName(id, true);
					if (oTable) {
						sheet = this.aWorksheets[oTable.index];
						range = oTable.table.Ref;
					}
				} else {
					//this table
					sheet = this.getActiveWs();
					var thisTableIndex = sheet.autoFilters.searchRangeInTableParts(sheet.selectionRange.getLast());
					if (thisTableIndex >= 0) {
						range = sheet.TableParts[thisTableIndex].Ref;
					} else {
						sheet = null;
					}
				}
				break;
			case Asc.c_oAscSelectionForCFType.pivot:
				sheet = this.getActiveWs();
				var _activeCell = sheet.selectionRange.activeCell;
				var _pivot = sheet.getPivotTable(_activeCell.col, _activeCell.row);
				if (_pivot) {
					range = _pivot.location && _pivot.location.ref;
				}

				if (!range) {
					sheet = null;
				}

				break;
		}
		if (sheet) {
			let _aRules = [];
			 sheet.forEachConditionalFormattingRules(function (elem) {
				_aRules.push(elem);
			});
			var aRules = _aRules.sort(function(v1, v2) {
				return v1.priority - v2.priority;
			});

			rules = [];
			if (range) {
				var putRange = function (_range, _rule) {
					multiplyRange = new AscCommonExcel.MultiplyRange(_range);
					if (multiplyRange.isIntersect(range)) {
						if (needClone) {
							var _id = _rule.id;
							var isLocked = _rule.isLock;
							_rule = _rule.clone();
							_rule.id = _id;
							_rule.isLock = isLocked;
							_rule.recalcInterfaceFormula(sheet, true);
						}
						rules.push(_rule);
					}
				};
				var oRule, ranges, multiplyRange, i, mapChangedRules = [];
				for (i = 0; i < aRules.length; ++i) {
					oRule = aRules[i];
					ranges = oRule.ranges;
					putRange(ranges, oRule);
				}
			} else {
				for (i = 0; i < aRules.length; ++i) {
					var _rule = aRules[i];
					if (needClone) {
						var _id = aRules[i].id;
						var isLocked = _rule.isLock;
						_rule = _rule.clone();
						_rule.id = _id;
						_rule.isLock = isLocked;
						_rule.recalcInterfaceFormula(sheet, true);
					}
					rules.push(_rule);
				}
			}
		}
		return rules;
	};

	Workbook.prototype.setProtectedWorkbook = function (props, addToHistory) {
		if (!this.workbookProtection) {
			this.workbookProtection = new window["Asc"].CWorkbookProtection();
		}
		this.workbookProtection.set(props, addToHistory, this);
		return true;
	};

	Workbook.prototype.getWorkbookProtection = function (type) {
		var res = false;

		if (this.workbookProtection) {
			switch (type) {
				case Asc.c_oAscWorkbookProtectType.lockWindows:
					res = this.workbookProtection.asc_getLockWindows();
					break;
				case Asc.c_oAscWorkbookProtectType.lockRevisions:
					res = this.workbookProtection.asc_getLockRevision();
					break;
				case Asc.c_oAscWorkbookProtectType.lockStructure:
				default:
					res = this.workbookProtection.asc_getLockStructure();
					break;
			}
		}

		return res;
	};

	Workbook.prototype.getPrintOptionsJson = function () {
		var res = [];
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			res[i] = this.aWorksheets[i].getPrintOptionsJson();
		}
		return res;
	};

	Workbook.prototype.setDate1904 = function (val, addToHistory) {
		var oldVal = AscCommon.bDate1904;

		AscCommon.bDate1904 = val;
		AscCommonExcel.c_DateCorrectConst = AscCommon.bDate1904 ? AscCommonExcel.c_Date1904Const : AscCommonExcel.c_Date1900Const;

		this.workbookPr.setDate1904(val);

		if (addToHistory) {
			var updateSheet = this.getActiveWs();
			var updateRange = new Asc.Range(0, 0, updateSheet.getColsCount(), updateSheet.getRowsCount());

			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_Date1904,
				updateSheet.getId(), updateRange, new UndoRedoData_FromTo(oldVal, val));
		}
	};


//-------------------------------------------------------------------------------------------------
	Workbook.prototype.getRangeAndSheetFromStr = function (sRange) {
		var ws;
		var range;
		if (sRange) {
			if (-1 !== sRange.indexOf("!")) {
				var is3DRef = AscCommon.parserHelp.parse3DRef(sRange);
				if (is3DRef) {
					ws = this.getWorksheetByName(is3DRef.sheet);
					range = AscCommonExcel.g_oRangeCache.getAscRange(is3DRef.range);
				}
			} else {
				ws = this.getActiveWs();
				range = AscCommonExcel.g_oRangeCache.getAscRange(sRange);
				//может быть именованный диапазон
				if (!range) {
					var dN = this.dependencyFormulas.getDefNameByName(sRange, ws.getId());
					if (dN && dN.parsedRef) {
						range = dN.parsedRef.getFirstRange();
						if (range) {
							range = range.bbox;
						}
					}
				}
			}
		}
		return {sheet: ws, range: range}
	};

	Workbook.prototype.addCellWatches = function (ws, range) {
		if (ws && range) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();

			//TODO protection!
			var maxCellWatchesCount = Asc.c_nAscMaxAddCellWatchesCount;
			var counter = 0;
			for (var i = range.r1; i <= range.r2; i++) {
				for (var j = range.c1; j <= range.c2; j++) {
					var _ref = new Asc.Range(j, i, j, i);
					ws.addCellWatch(_ref, true);
					counter++;
					if (counter === maxCellWatchesCount) {
						break;
					}
				}
				if (counter === maxCellWatchesCount) {
					break;
				}
			}

			AscCommon.History.EndTransaction();

			/*if (!this.bUndoChanges && !this.bRedoChanges) {
				this.handlers.trigger("asc_onUpdateCellWatches");
			}*/
		}
	};

	Workbook.prototype.delCellWatches = function (aCellWatches, addToHistory, opt_remove_all) {
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		//TODO protection!
		var i;
		if (opt_remove_all) {
			for(i = 0; i < this.aWorksheets.length; ++i) {
				this.aWorksheets[i].deleteCellWatches(addToHistory);
			}
		} else {
			for (i = 0; i < aCellWatches.length; i++) {
				var ws = aCellWatches[i]._ws;
				ws.deleteCellWatch(aCellWatches[i].r, addToHistory);
			}
		}

		AscCommon.History.EndTransaction();

		/*if (!this.bUndoChanges && !this.bRedoChanges) {
			this.handlers.trigger("asc_onUpdateCellWatches");
		}*/
	};

	Workbook.prototype.recalculateCellWatches = function (fullRecalc) {
		var count = 0;
		var changedMap = null;
		for(var i = 0; i < this.aWorksheets.length; ++i) {
			for (var j = 0; j < this.aWorksheets[i].aCellWatches.length; j++) {
				if (this.aWorksheets[i].aCellWatches[j].recalculate(fullRecalc)) {
					if (!changedMap) {
						changedMap = {};
					}
					changedMap[count] = this.aWorksheets[i].aCellWatches[j];
				}
				count++;
			}
		}
		return changedMap;
	};

	//*****external links*****
	//при открытии ждём ссылок в виде [1]Sheet1!A1:A2, но если пользователь введёт такую ссылку, то текст самой ссылки будет уже "1", а индекс самой ссылки увеличится length + 1
	//далее отображаем в таком виде в зависимости от самой ссылки 'https://s3.amazonaws.com/xlsx/[ExternalLinksDestination.xlsx]Sheet1'!A1:A2
	//при вводе с клавиатуры сложнее - мс пропускает ссылки в достаточно странном виде 'abracadabra1://abracadabra2:[file.abracadabra3]Sheet1'!A1:A2
	Workbook.prototype.getExternalLinkByIndex = function (index, needSplit) {
		var res = this.externalReferences && this.externalReferences[index];
		if (needSplit && res) {
			//разбиваем на имя и путь
			//ms обрабатывает http:/https:/ftp: и любая конструкция до : -  ..test..:
			//предполагаем, что здесь уже лежит ссылка в грамотном виде
			//ищем последний слэш или двоеточие и разделяем на части
			res = res.Id;
			var lastSlash = "/";
			var checkPrefix = "file:///";
			var prefixFile = "";
			if (0 === res.indexOf(checkPrefix)) {
				res = res.substring(8);
				prefixFile = checkPrefix;
				lastSlash = "\\";
			}

			for (let i = res.length - 1; i >= 0; i--) {
				if (res[i] === lastSlash || res[i] === ":" || (res[i] === "\\" && res.indexOf(":") !== -1)) {
					return {path: prefixFile + res.substring(0, i + 1), name: res.substring(i + 1, res.length)};
				}
			}
			if (res) {
				res = {path: prefixFile + "", name: res};
			}
		}
		return res ? res : null;
	};

	Workbook.prototype.getExternalLinkByName = function (name) {
		for (var i = 0; i < this.externalReferences.length; i++) {
			if (this.externalReferences[i].Id === name) {
				return this.externalReferences[i];
			}
		}
	};

	Workbook.prototype.getExternalLinkIndexByName = function (name) {
		for (var i = 0; i < this.externalReferences.length; i++) {
			if (this.externalReferences[i].Id === name) {
				return i + 1;
			}
		}
		return null;
	};

	Workbook.prototype.getExternalLinkIndexBySheetId = function (sheetId) {
		for (var i = 0; i < this.externalReferences.length; i++) {
			for (var j in this.externalReferences[i].worksheets) {
				if (this.externalReferences[i].worksheets[j].Id === sheetId) {
					return i + 1;
				}
			}
		}
		return null;
	};

	Workbook.prototype.getExternalWorksheet = function (val, sheet, getFirstSheet) {
		let extarnalLink = this.getExternalLink(val);
		if (extarnalLink) {
			if (getFirstSheet && extarnalLink.SheetNames) {
				sheet = extarnalLink.SheetNames[0];
			}

			if (null == sheet) {
				return null;
			}
			if (extarnalLink.worksheets && extarnalLink.worksheets[sheet]) {
				return extarnalLink.worksheets[sheet];
			}
			if (extarnalLink.SheetNames) {
				for (var i = 0; i < extarnalLink.SheetNames.length; i++) {
					if (extarnalLink.SheetNames[i] === sheet) {
						var wb = this.getTemporaryExternalWb();
						extarnalLink.worksheets[sheet] = new Worksheet(wb, wb.aWorksheets.length);
						wb.aWorksheets.push(extarnalLink.worksheets[sheet]);
						extarnalLink.worksheets[sheet].sName = sheet;
						return extarnalLink.worksheets[sheet];
					}
				}
			}
		}
		return null;
	};

	Workbook.prototype.getExternalLink = function (val) {
		return window['AscCommon'].isNumber(val) ? this.getExternalLinkByIndex(val - 1) : this.getExternalLinkByName(val);
	};

	Workbook.prototype.getExternalWorksheetByIndex = function (index, sheet) {
		var extarnalLink = this.getExternalLinkByIndex(index);
		if (extarnalLink) {
			if (extarnalLink.worksheets && extarnalLink.worksheets[sheet]) {
				return extarnalLink.worksheets[sheet];
			}
			if (extarnalLink.SheetNames) {
				for (var i = 0; i < extarnalLink.SheetNames.length; i++) {
					if (extarnalLink.SheetNames[i] === sheet) {
						var wb = this.getTemporaryExternalWb();
						extarnalLink.worksheets[sheet] = new Worksheet(wb);
						return extarnalLink.worksheets[sheet];
					}
				}
			}
		}
		return null;
	};

	Workbook.prototype.getExternalWorksheetByName = function (name, sheet) {
		var extarnalLink = this.getExternalLinkByName(name);
		if (extarnalLink) {
			if (extarnalLink.worksheets && extarnalLink.worksheets[sheet]) {
				return extarnalLink.worksheets[sheet];
			}
			if (extarnalLink.SheetNames) {
				for (var i = 0; i < extarnalLink.SheetNames.length; i++) {
					if (extarnalLink.SheetNames[i] === sheet) {
						var wb = this.getTemporaryExternalWb();
						extarnalLink.worksheets[sheet] = new Worksheet(wb);
						return extarnalLink.worksheets[sheet];
					}
				}
			}
		}
		return null;
	};

	Workbook.prototype.getTemporaryExternalWb = function () {
		for (var i = 0; i < this.externalReferences.length; i++) {
			if (this.externalReferences[i].worksheets) {
				for (var j in this.externalReferences[i].worksheets) {
					if (this.externalReferences[i].worksheets[j]) {
						return this.externalReferences[i].worksheets[j].workbook;
					}
				}
			}
		}
		return new Workbook(undefined, undefined, false);
	};

	Workbook.prototype.getExternalIndexByWorksheet = function (ws) {
		let checkedWb = ws && ws.workbook;
		for (let i = 0; i < this.externalReferences.length; i++) {
			let externalWb = this.externalReferences[i] && this.externalReferences[i].getWb && this.externalReferences[i].getWb();
			if (checkedWb === externalWb) {
				if (this.externalReferences[i].worksheets) {
					for (let j in this.externalReferences[i].worksheets) {
						if (this.externalReferences[i].worksheets[j] === ws) {
							return i + 1;
						}
					}
				}
			}
		}
		return null;
	};

	Workbook.prototype.removeExternalReferences = function (arr) {
		//пока предполагаю, что здесь будет массив asc_CExternalReference
		if (arr) {
			var isChanged = false;
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();
			for (var i = 0; i < arr.length; i++) {
				var eRIndex = this.getExternalLinkIndexByName(arr[i].externalReference.Id);
				if (eRIndex != null) {

					//TODO при undo кладутся в массив в обратном порядке - нужно всегда в одном порядке добавлять
					this.removeExternalReference(eRIndex, true);
					isChanged = true;

					//TODO нужно заменить все ячейки просто значениями, где есть формулы, которые ссылаются на эту книгу
					for (var j in arr[i].externalReference.worksheets) {
						var removedSheet = arr[i].externalReference.worksheets[j];
						if (removedSheet) {

							var removedSheetId = removedSheet.getId();
							this.dependencyFormulas.forEachSheetListeners(removedSheetId, function (parsed) {
								if (parsed) {
									var cell = parsed.parent;
									if (cell && cell.nCol != null && cell.nRow != null) {
										//нужно удалить формулу из этой ячейки
										parsed.ws._getCellNoEmpty(cell.nRow, cell.nCol, function(cell) {
											var valueData = cell.getValueData();
											valueData.formula = null;
											cell.setValueData(valueData);
											//cell.setFormulaInternal(null);
										});
									}
								}
							});


							var prepared = this.dependencyFormulas.prepareChangeSheet(removedSheetId, {remove: removedSheetId});
							this.dependencyFormulas.removeSheet(prepared);
						}
					}
				}
			}
			AscCommon.History.EndTransaction();

			if (isChanged) {
				this.handlers && this.handlers.trigger("asc_onUpdateExternalReferenceList");
			}
		}
	};

	Workbook.prototype.removeExternalReference = function (index, addToHistory) {
		if (index != null) {
			var from = this.externalReferences[index - 1];
			//this.reIndexExternalReferencesLinks(index - 1);
			this._removeExternalReference(index - 1);
			if (addToHistory) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_ChangeExternalReference,
					null, null, new UndoRedoData_FromTo(from, null));
			}
		}
	};

	Workbook.prototype._removeExternalReference = function (index) {
		if (index != null) {
			//need offset last references
			for (let i = index + 1; i < this.externalReferences.length; i++) {
				for (let j in this.externalReferences[i].worksheets) {
					var prepared = this.dependencyFormulas.prepareChangeSheet(this.externalReferences[i].worksheets[j].getId(), {from: i + 1, to: i});
					this.dependencyFormulas.changeExternalLink(prepared);
				}
			}

			this.externalReferences.splice(index, 1);
		}
	};

	Workbook.prototype.changeExternalReference = function (index, to) {
		if (index != null) {
			let from = this.externalReferences[index - 1].clone(true);
			// var from = this.externalReferences[index - 1].clone();

			// change formulas with importRange 
			this.changeImportRangeLinks(index, from, to);

			this.externalReferences[index - 1] = to;

			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_ChangeExternalReference,
				null, null, new UndoRedoData_FromTo(from, to));
		}
	};

	Workbook.prototype.changeImportRangeLinks = function (index, fromER, toER) {
		if (index != null && fromER && toER) {
			let fromLinkStr = fromER.Id;
			let toLinkStr = toER.Id;

			for (let ws in toER.worksheets) {
				let changedSheet = toER.worksheets[ws];
				if (changedSheet) {
					let changedSheetId = changedSheet.getId();

					this.dependencyFormulas.forEachSheetListeners(changedSheetId, function (parsed) {
						// we go through the external sheet listeners and change each importRange
						let cell = parsed && parsed.parent;
						if (cell && cell && cell.nCol != null && cell.nRow != null && parsed.Formula) {
							let newFormulaStr = "";
							let formula = parsed.getFormula();
							if (formula) {
								newFormulaStr = formula.split(fromLinkStr).join(toLinkStr);
							}

							if (newFormulaStr) {
								let newFP = new AscCommonExcel.parserFormula(newFormulaStr, cell, parsed.ws);
								let parseResult = new AscCommonExcel.ParseResult();;
								if (newFP.parse(AscCommonExcel.oFormulaLocaleInfo.Parse, AscCommonExcel.oFormulaLocaleInfo.DigitSep, parseResult)) {
									parsed.ws._getCellNoEmpty(cell.nRow, cell.nCol, function(cell) {
										if (cell) {
											cell.setValue("=" + newFormulaStr, null, null, parsed.ref, null, null, true);
										}
									});

								}
							}
						}
					});
				}
			}
		}
	};

	Workbook.prototype.addExternalReferences = function (arr) {
		if (arr && arr.length) {
			for (var i = 0; i < arr.length; i++) {
				this.externalReferences.push(arr[i]);
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_ChangeExternalReference,
					null, null, new UndoRedoData_FromTo(null, arr[i]));
			}
			this.handlers && this.handlers.trigger("asc_onUpdateExternalReferenceList");
		}
	};

	Workbook.prototype.getExternalLinkByReferenceData = function (referenceData) {
		for (var i = 0; i < this.externalReferences.length; i++) {
			if (this.externalReferences[i].referenceData) {
				if (this.externalReferences[i].referenceData["fileKey"] === referenceData["fileKey"] && this.externalReferences[i].referenceData["instanceId"] === referenceData["instanceId"]) {
					return {index: i + 1, val: this.externalReferences};
				}
			}
		}
		return null;
	};

	Workbook.prototype.getExternalReferenceByReferenceData = function (referenceData, returnIndex) {
		if (!referenceData) {
			return null;
		}

		for (let i = 0; i < this.externalReferences.length; i++) {
			if (this.externalReferences[i].referenceData) {
				if (this.externalReferences[i].referenceData["fileKey"] === referenceData["fileKey"] && this.externalReferences[i].referenceData["instanceId"] === referenceData["instanceId"]) {
					return returnIndex ? i + 1 : this.externalReferences[i];
				}
			}
		}
		return null;
	};

	Workbook.prototype.getExternalReferenceWithoutRefData = function (id, returnIndex) {
		/* Receive an external link only by name(id) and without reference data */
		for (let i = 0; i < this.externalReferences.length; i++) {
			if (this.externalReferences[i].Id === id && !this.externalReferences[i].referenceData) {
				return returnIndex ? i + 1 : this.externalReferences[i];
			}
		}
		return null;
	};

	Workbook.prototype.getExternalReferences = function () {
		var res = null;
		for (var i = 0; i < this.externalReferences.length; i++) {
			if (this.externalReferences[i].getAscLink) {
				if (!res) {
					res = [];
				}
				res.push(this.externalReferences[i].getAscLink());
			}
		}
		return res;
	};

	Workbook.prototype.removeExternalReferenceBySheet = function (sheetId) {
		//пока предполагаю, что здесь будет массив asc_CExternalReference
		let index = this.getExternalLinkIndexBySheetId(sheetId);
		if (index !== null) {
			let eR = this.externalReferences[index - 1];
			// If there are no more listeners for the external document, delete the external link
			this.removeExternalReference(index, true);

			// if (eR.SheetNames.length === 1) {
			// 	//удаляем ссылку
			// 	this.removeExternalReference(index, true);
			// } else {
			// 	let to = eR.clone();
			// 	eR.removeSheetById(sheetId);
			// 	this.changeExternalReference(index, eR);
			// }
			this.handlers && this.handlers.trigger("asc_onUpdateExternalReferenceList");
		}
	};

	Workbook.prototype.getExternalReferenceById = function (id, returnIndex) {
		for (var i = 0; i < this.externalReferences.length; i++) {
			if (this.externalReferences[i].Id === id) {
				return returnIndex ? i : this.externalReferences[i];
			}
		}
		return null;
	};

	Workbook.prototype.addExternalReferencesAfterParseFormulas = function (externalReferenesNeedAdd) {
		let newExternalReferences = [];
		for (let i in externalReferenesNeedAdd) {
			let needAdd = false;
			let newExternalReference = this.getExternalReferenceById(i);
			if (!newExternalReference) {
				newExternalReference = new AscCommonExcel.ExternalReference();
				newExternalReference.Id = i;
				needAdd = true;
			}

			for (let j = 0; j < externalReferenesNeedAdd[i].length; j++) {
				let oNewSheet = externalReferenesNeedAdd[i][j];
				if (null === newExternalReference.getSheetByName(oNewSheet.sheet)) {
					let newSheet = oNewSheet.sheet;
					newExternalReference.addSheetName(newSheet, true);
					newExternalReference.initWorksheetFromSheetDataSet(newSheet);
				}
				if (oNewSheet.notUpdateId) {
					newExternalReference.notUpdateId = true;
				}
			}

			if (needAdd) {
				newExternalReferences.push(newExternalReference);
			} else {
				//this.changeExternalReference();
			}
		}

		this.addExternalReferences(newExternalReferences);
	};

	Workbook.prototype.setUpdateLinks = function (val, addToHistory, bFirstStart, bNotStartTimer) {
		var from = !!(this.workbookPr.getUpdateLinks());
		if (val !== from) {
			this.workbookPr.setUpdateLinks(val);
			if (addToHistory) {
				History.Create_NewPoint();
				History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_UpdateLinks,
					null, null, new UndoRedoData_FromTo(from, val));
			}
			!bNotStartTimer && this.handlers && this.handlers.trigger("changeUpdateLinks");
		} else if (bFirstStart) {
			!bNotStartTimer && this.handlers && this.handlers.trigger("changeUpdateLinks");
		}
	};

	Workbook.prototype.getUpdateLinks = function () {
		return this.workbookPr && this.workbookPr.getUpdateLinks();
	};

	Workbook.prototype.unlockUserProtectedRanges = function(){
		this.forEach(function (ws) {
			ws.unlockUserProtectedRanges();
		});
	};

	Workbook.prototype.isUserProtectedRangesIntersection = function(range, userId){
		let res = false;
		for (var i = 0, l = this.aWorksheets.length; i < l; ++i) {
			if (this.aWorksheets[i].isUserProtectedRangesIntersection(range, userId)) {
				return true;
			}
		}
		return res;
	};

	Workbook.prototype.checkUserProtectedRangeName = function (name) {
		var res = c_oAscDefinedNameReason.OK;
		//TODO пересмотреть проверку на rx_defName
		if (!AscCommon.rx_protectedRangeName.test(name.toLowerCase()) || name.length > g_nDefNameMaxLength) {
			return c_oAscDefinedNameReason.WrongName;
		}

		return res;
	};

	/******GOAL SEEK******
	/**
	 * Initializes and starts goal seek calculation.
	 * @param {string} sFormulaCell
	 * @param {string} sExpectedValue
	 * @param {string} sChangingCell
	 * @param {Worksheet} wsFormula
	 * @param {Worksheet} wsChangingCell
	 */
	Workbook.prototype.startGoalSeek = function(sFormulaCell, sExpectedValue, sChangingCell, wsFormula, wsChangingCell) {
		let oParserFormula;
		let oFormulaCell = wsFormula.getCell2(sFormulaCell);
		wsFormula._getCell(oFormulaCell.bbox.r1, oFormulaCell.bbox.c1, function (cell) {
			oParserFormula = cell.getFormulaParsed();
		});

		this.setGoalSeek(new AscCommonExcel.CGoalSeek(oParserFormula, Number(sExpectedValue), wsChangingCell.getRange2(sChangingCell)));
		let oGoalSeek = this.getGoalSeek();
		// Run goal seek
		oGoalSeek.init();
		oGoalSeek.setIntervalId(setInterval(function () {
			let bIsFinish = oGoalSeek.calculate();
			if (bIsFinish) {
				clearInterval(oGoalSeek.getIntervalId());
			}
		}, oGoalSeek.getDelay()));
	};
	/**
	 * Returns object with goal seek result
	 * @returns {CGoalSeek}
	 */
	Workbook.prototype.getGoalSeek = function() {
		return this.oGoalSeek;
	};
	/**
	 * Sets object with goal seek result
	 * @param {CGoalSeek} oGoalSeek
	 */
	Workbook.prototype.setGoalSeek = function(oGoalSeek) {
		this.oGoalSeek = oGoalSeek;
	};
	/**
	 * Discards goal seek result for "Changing cell" to original
	 */
	Workbook.prototype.closeGoalSeek = function () {
		let oGoalSeek = this.getGoalSeek();
		if (!oGoalSeek) {
			return;
		}
		let oChangedCell = oGoalSeek.getChangingCell();
		let nFirstChangingVal = oGoalSeek.getFirstChangingValue();
		oChangedCell.setValue(nFirstChangingVal == null ? "" : nFirstChangingVal + "");
		this.setGoalSeek(null);
	};
	/**
	 * Saves goal seek result for "Changing cell"
	 */
	Workbook.prototype.saveGoalSeek = function() {
		this.setGoalSeek(null);
	};

	Workbook.prototype.pauseGoalSeek = function() {
		this.oGoalSeek && this.oGoalSeek.pause();
	};

	Workbook.prototype.continueGoalSeek = function() {
		this.oGoalSeek && this.oGoalSeek.resume();
	};

	Workbook.prototype.stepGoalSeek = function() {
		this.oGoalSeek && this.oGoalSeek.step();
	};

	/**
	 * Sets Solver object
	 * @memberof Workbook
	 * @param {CSolver} oSolver
	 */
	Workbook.prototype.setSolver = function(oSolver) {
		this.oSolver = oSolver;
	};

	/**
	 * Returns Solver object
	 * @memberof Workbook
	 * @returns {CSolver}
	 */
	Workbook.prototype.getSolver = function() {
		return this.oSolver;
	};

	Workbook.prototype.setDefaultDirection = function(val) {
		this.defaultDirection = val;
	};
	Workbook.prototype.getDefaultDirection = function() {
		return this.defaultDirection;
	};
	Workbook.prototype.setShowVerticalScroll = function(val, addToHistory) {
		var from = this.showVerticalScroll;
		if (from !== val) {
			if (addToHistory) {
				AscCommon.History.Create_NewPoint();
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_ShowVerticalScroll, null, null, new UndoRedoData_FromTo(from, val));
			}
			this.showVerticalScroll = val;
			this.handlers && this.handlers.trigger("updateScrollVisibility");
		}
	};
	Workbook.prototype.getShowVerticalScroll = function() {
		return this.showVerticalScroll;
	};
	Workbook.prototype.setShowHorizontalScroll = function(val, addToHistory) {
		var from = this.showHorizontalScroll;
		if (from !== val) {
			if (addToHistory) {
				AscCommon.History.Create_NewPoint();
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorkbook, AscCH.historyitem_Workbook_ShowHorizontalScroll, null, null, new UndoRedoData_FromTo(from, val));
			}
			this.showHorizontalScroll = val;
			this.handlers && this.handlers.trigger("updateScrollVisibility");
		}
	};
	Workbook.prototype.getShowHorizontalScroll = function() {
		return this.showHorizontalScroll;
	};
	Workbook.prototype.getAllInks = function(arrInks) {
		arrInks = arrInks || [];
		for (let i = 0; i < this.aWorksheets.length; i += 1) {
			const oWorksheet = this.aWorksheets[i];
			oWorksheet.getAllInks(arrInks);
		}
		return arrInks;
	};

	/**
	 * @constructor
	 */
	function Worksheet(wb, _index, sId){
		this.workbook = wb;
		this.sName = this.workbook.getUniqueSheetNameFrom(g_sNewSheetNamePattern, false);
		this.bHidden = false;
		this.oSheetFormatPr = new AscCommonExcel.SheetFormatPr();
		this.index = _index;
		this.Id = null != sId ? sId : AscCommon.g_oIdCounter.Get_NewId();//todo AscCommon.g_oTableId.Add
		this.nRowsCount = 0;
		this.nColsCount = 0;
		this.rowsData = new AscCommonExcel.SheetMemory(AscCommonExcel.g_nRowStructSize, gc_nMaxRow0);
		this.cellsByCol = [];
		this.cellsByColRowsCount = 0;//maximum rows count in cellsByCol
		this.aCols = [];// 0 based
		this.hiddenManager = new HiddenManager(this);
		this.Drawings = [];
		this.TableParts = [];
		this.AutoFilter = null;
		this.oAllCol = null;
		this.aComments = [];
		var oThis = this;
		this.bExcludeHiddenRows = false;
		this.bIgnoreWriteFormulas = false;
		this.mergeManager = new RangeDataManager(function(data, from, to){
			if(AscCommon.History.Is_On() && (null != from || null != to))
			{
				if(null != from)
					from = from.clone();
				if(null != to)
					to = to.clone();
				var oHistoryRange = from;
				if(null == oHistoryRange)
					oHistoryRange = to;
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ChangeMerge, oThis.getId(), oHistoryRange, new UndoRedoData_FromTo(new UndoRedoData_BBox(from), new UndoRedoData_BBox(to)));
			}
			//расширяем границы
			if(null != to){
				var maxRow = gc_nMaxRow0 !== to.r2 ? to.r2 : to.r1;
				var maxCol = gc_nMaxCol0 !== to.c2 ? to.c2 : to.c1;
				if(maxRow >= oThis.nRowsCount)
					oThis.nRowsCount = maxRow + 1;
				if(maxCol >= oThis.nColsCount)
					oThis.setColsCount(maxCol + 1);
			}
		});
		this.mergeManager.worksheet = this;
		this.hyperlinkManager = new RangeDataManager(function(data, from, to, oChangeParam){
			if(AscCommon.History.Is_On() && (null != from || null != to))
			{
				if(null != from)
					from = from.clone();
				if(null != to)
					to = to.clone();
				var oHistoryRange = from;
				if(null == oHistoryRange)
					oHistoryRange = to;
				var oHistoryData = null;
				if(null == from || null == to)
					oHistoryData = data.clone();
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ChangeHyperlink, oThis.getId(), oHistoryRange, new AscCommonExcel.UndoRedoData_FromToHyperlink(from, to, oHistoryData));
			}
			if (null != to)
				data.Ref = oThis.getRange3(to.r1, to.c1, to.r2, to.c2);
			else if (oChangeParam && oChangeParam.removeStyle && null != data.Ref)
				data.Ref.cleanFormat();
			//расширяем границы
			if(null != to){
				var maxRow = gc_nMaxRow0 !== to.r2 ? to.r2 : to.r1;
				var maxCol = gc_nMaxCol0 !== to.c2 ? to.c2 : to.c1;
				if(maxRow >= oThis.nRowsCount)
					oThis.nRowsCount = maxRow + 1;
				if(maxCol >= oThis.nColsCount)
					oThis.setColsCount(maxCol + 1);
			}
		});
		this.hyperlinkManager.setDependenceManager(this.mergeManager);
		this.sheetViews = [];
		this.aConditionalFormattingRules = null;

		this.conditionalFormattingRangeIterator = null;
		this.updateConditionalFormattingRange = null;
		this.dataValidations = null;
		this.sheetPr = null;
		this.aFormulaExt = null;

		this.autoFilters = AscCommonExcel.AutoFilters !== undefined ? new AscCommonExcel.AutoFilters(this) : null;
		this.sortState = null;

		this.contentChanges = new AscCommon.CContentChanges();

		this.aSparklineGroups = [];

		this.selectionRange = new AscCommonExcel.SelectionRange(this);
		this.copySelection = null;

		this.sheetMergedStyles = new AscCommonExcel.SheetMergedStyles();
		this.pivotTables = [];
		this.headerFooter = new Asc.CHeaderFooter(this);
		this.rowBreaks = null;
		this.colBreaks = null;
		this.legacyDrawingHF = null;
		this.picture = null;

		this.PagePrintOptions = new Asc.asc_CPageOptions(this);
		//***array-formula***
		//TODO пересмотреть. нужно для того, чтобы хранить ссылку на parserFormula главной ячейки при проходе по range массива
		this.formulaArrayLink = null;

		this.lastFindOptions = null;

		//чтобы разделять ситуации, когда группы скрываются/открываются из меню(в данном случае не скрываются внутренние группы)
		//и ситуацию, когда группы скрываются/открываются при скрытии строк/столбцов(все внутренние группы скрываются)
		//этот флаг проставляются в true при скрытии групп из меню группировки(нажатие на +/- и скрытие целиком всего уровня)
		//в данном случае не нужно при скрытии строк делать setCollapsed(из-за внутренних групп) и заносить данные в историю
		//во всех остальных ситуациях это делать необходимо
		this.bExcludeCollapsed = false;

		this.oNumFmtsOpen = {};
		this.dxfsOpen = [];//todo move to InitOpenManager

		/*handlers*/
		this.handlers = null;

		this.aSlicers = [];

		this.aNamedSheetViews = [];
		this.activeNamedSheetViewId = null;
		this.defaultViewHiddenRows = null;

		this.sheetProtection = null;
		this.aProtectedRanges = [];
		this.aCellWatches = [];
		
		this.userProtectedRanges = [];

		this.bFillHandleRightClick = false;
		this.nFillMenuChosenProp = null;
		this.activeFillType = null;
		this.timelines = [];
		this.changedArrays = null;
	}

	Worksheet.prototype.getCompiledStyle = function (row, col, opt_cell, opt_styleComponents, opt_AffectingText) {
		return getCompiledStyle(this.sheetMergedStyles, this.hiddenManager, row, col, opt_cell, this, opt_styleComponents, opt_AffectingText);
	};
	Worksheet.prototype.getCompiledStyleCustom = function(row, col, needTable, needCell, needConditional, opt_cell) {
		var res;
		var styleComponents = this.sheetMergedStyles.getStyle(this.hiddenManager, row, col, this);
		var ws = this;
		if (!needTable) {
			styleComponents.table = [];
		}
		if (!needConditional) {
			styleComponents.conditional = [];
		}
		if (!needCell) {
			res = getCompiledStyle(undefined, undefined, row, col, undefined, undefined, styleComponents);
		} else if (opt_cell) {
			res = getCompiledStyle(undefined, undefined, row, col, opt_cell, ws, styleComponents);
		} else {
			this._getCellNoEmpty(row, col, function(cell) {
				res = getCompiledStyle(undefined, undefined, row, col, cell, ws, styleComponents);
			});
		}
		return res;
	};
	Worksheet.prototype.getColData = function(index) {
		var sheetMemory = this.cellsByCol[index];
		if(!sheetMemory){
			sheetMemory = new AscCommonExcel.SheetMemory(g_nCellStructSize, gc_nMaxRow0);
			this.cellsByCol[index] = sheetMemory;
		}
		return sheetMemory;
	};
	Worksheet.prototype.getColDataNoEmpty = function(index) {
		return this.cellsByCol[index];
	};
	Worksheet.prototype.getColDataLength = function() {
		return this.cellsByCol.length;
	};
	//returns minimal range containing all no empty cells
	Worksheet.prototype.getMinimalRange = function() {
		var oRange = null;
		this._forEachCell(function(oCell) {
			if(!oCell.isEmpty()) {
				if(oRange === null) {
					oRange = new Asc.Range(oCell.nCol, oCell.nRow, oCell.nCol, oCell.nRow)
				} else {
					oRange.union3(oCell.nCol, oCell.nRow)
				}
			}
		});
		return oRange;
	};
	Worksheet.prototype.getSnapshot = function(wb) {
		var ws = new Worksheet(wb, this.index, this.Id);
		ws.sName = this.sName;
		for (var i = 0; i < this.TableParts.length; ++i) {
			var table = this.TableParts[i];
			ws.addTablePart(table.clone(null), false);
		}
		for (i = 0; i < this.sheetViews.length; ++i) {
			ws.sheetViews.push(this.sheetViews[i].clone());
		}
		return ws;
	};
	Worksheet.prototype.addContentChanges = function (changes) {
		this.contentChanges.Add(changes);
	};
	Worksheet.prototype.refreshContentChanges = function () {
		this.contentChanges.Refresh();
		this.contentChanges.Clear();
	};
	Worksheet.prototype.rebuildColors=function(){
		this.rebuildTabColor();

		for (var i = 0; i < this.aSparklineGroups.length; ++i) {
			this.aSparklineGroups[i].cleanCache();
		}
	};
	Worksheet.prototype.generateFontMap=function(oFontMap){
		//пробегаемся по Drawing
		for(var i = 0, length = this.Drawings.length; i < length; ++i)
		{
			var drawing = this.Drawings[i];
			if(drawing)
				drawing.getAllFonts(oFontMap);
		}

		//пробегаемся по header/footer
		if(this.headerFooter){
			this.headerFooter.getAllFonts(oFontMap);
		}
	};
	Worksheet.prototype.getAllImageUrls = function(aImages){
		for(var i = 0; i < this.Drawings.length; ++i){
			this.Drawings[i].graphicObject.getAllRasterImages(aImages);
		}
	};
	Worksheet.prototype.reassignImageUrls = function(oImages){
		for(var i = 0; i < this.Drawings.length; ++i){
			this.Drawings[i].graphicObject.Reassign_ImageUrls(oImages);
		}
	};
	Worksheet.prototype.copyFrom=function(wsFrom, sName, tableNames){
		var i, elem, range, _newSlicer;
		var t = this;
		this.sName = this.workbook.checkValidSheetName(sName) ? sName : this.workbook.getUniqueSheetNameFrom(wsFrom.sName, true);
		this.bHidden = wsFrom.bHidden;
		this.oSheetFormatPr = wsFrom.oSheetFormatPr.clone();
		//this.index = wsFrom.index;
		this.nRowsCount = wsFrom.nRowsCount;
		this.setColsCount(wsFrom.nColsCount);
		var renameParams = {lastName: wsFrom.getName(), newName: this.getName(), tableNameMap: {}, slicerNameMap: {}, copySlicerError: false};
		for (i = 0; i < wsFrom.TableParts.length; ++i)
		{
			var tableFrom = wsFrom.TableParts[i];
			var tableTo = tableFrom.clone(null);
			if(tableNames && tableNames.length) {
				tableTo.changeDisplayName(tableNames[i]);
			} else {
				tableTo.changeDisplayName(this.workbook.dependencyFormulas.getNextTableName());
			}
			this.addTablePart(tableTo, true);
			renameParams.tableNameMap[tableFrom.DisplayName] = tableTo.DisplayName;
		}
		for (i = 0; i < this.TableParts.length; ++i) {
			this.TableParts[i].renameSheetCopy(this, renameParams);
		}
		if(wsFrom.AutoFilter)
			this.AutoFilter = wsFrom.AutoFilter.clone();
		for (i in wsFrom.aCols) {
			var col = wsFrom.aCols[i];
			if(null != col)
				this.aCols[i] = col.clone(this);
		}
		if(null != wsFrom.oAllCol)
			this.oAllCol = wsFrom.oAllCol.clone(this);

		//copy row/cell data
		this.rowsData = wsFrom.rowsData.clone();
		wsFrom._forEachColData(function(sheetMemory, index){
			t.cellsByCol[index] = sheetMemory.clone();
		});
		this.cellsByColRowsCount = wsFrom.cellsByColRowsCount;

		var aMerged = wsFrom.mergeManager.getAll();
		for(i in aMerged)
		{
			elem = aMerged[i];
			range = this.getRange3(elem.bbox.r1, elem.bbox.c1, elem.bbox.r2, elem.bbox.c2);
			range.mergeOpen();
		}
		var aHyperlinks = wsFrom.hyperlinkManager.getAll();
		for(i in aHyperlinks)
		{
			elem = aHyperlinks[i];
			range = this.getRange3(elem.bbox.r1, elem.bbox.c1, elem.bbox.r2, elem.bbox.c2);
			range.setHyperlinkOpen(elem.data);
		}
		if(null != wsFrom.aComments) {
			for (i = 0; i < wsFrom.aComments.length; i++) {
				var comment = wsFrom.aComments[i].clone(true);
				comment.wsId = this.getId();
				comment.nId = "sheet" + comment.wsId + "_" + (i + 1);
				this.aComments.push(comment);
			}
		}
		for (i = 0; i < wsFrom.sheetViews.length; ++i) {
			this.sheetViews.push(wsFrom.sheetViews[i].clone());
		}
		if (wsFrom.isConditionalFormattingRules()) {
			let t = this;
			wsFrom.forEachConditionalFormattingRules(function (val) {
				t.addConditionalFormattingRule(val);
			})
		}
		if (wsFrom.sheetPr)
			this.sheetPr = wsFrom.sheetPr.clone();

		this.selectionRange = wsFrom.selectionRange.clone(this);

		if(wsFrom.PagePrintOptions) {
			this.PagePrintOptions = wsFrom.PagePrintOptions.clone(this);
		}

		//copy headers/footers
		if(wsFrom.headerFooter) {
			this.headerFooter = wsFrom.headerFooter.clone(this);
		}

		for (i = 0; i < wsFrom.aSlicers.length; ++i) {
			//пока только для таблиц
			var _slicer = wsFrom.aSlicers[i];
			var _table = _slicer.getTableSlicerCache();
			var pivotCache = _slicer.getPivotCache();
			if (_table) {
				var tableIdNew = renameParams.tableNameMap[_table.tableId];
				if (tableIdNew) {
					_newSlicer = this.insertSlicer(_table.column, tableIdNew, window['AscCommonExcel'].insertSlicerType.table);
					_newSlicer.set(_slicer.clone(), true);
				}

				if (_newSlicer) {
					renameParams.slicerNameMap[_slicer.name] = _newSlicer.name;
				}
			} else if (pivotCache) {
				var _newCacheDefinition = _slicer.getCacheDefinition().clone(this.workbook);
				_newCacheDefinition.name = _newCacheDefinition.generateSlicerCacheName(_newCacheDefinition.name);
				_newSlicer = this.insertSlicer(_slicer.name, undefined, window['AscCommonExcel'].insertSlicerType.pivotTable, undefined, _newCacheDefinition);
				_newCacheDefinition.forCopySheet(wsFrom.getId(), this.getId());
				_newSlicer.set(_slicer.clone(), true);
				renameParams.slicerNameMap[_slicer.name] = _newSlicer.name;
			}
		}

		if(wsFrom.headerFooter) {
			this.headerFooter = wsFrom.headerFooter.clone(this);
		}

		if(wsFrom.sheetProtection) {
			this.sheetProtection = wsFrom.sheetProtection.clone(this);
		}
		if(wsFrom.aProtectedRanges) {
			for (i = 0; i < wsFrom.aProtectedRanges.length; i++) {
				if (!this.aProtectedRanges) {
					this.aProtectedRanges = [];
				}
				this.aProtectedRanges.push(wsFrom.aProtectedRanges[i].clone(this));
			}
		}

		if(wsFrom.colBreaks) {
			this.colBreaks = wsFrom.colBreaks.clone(this);
		}
		if(wsFrom.rowBreaks) {
			this.rowBreaks = wsFrom.rowBreaks.clone(this);
		}

		return renameParams;
	};

	Worksheet.prototype.copyFromAfterInsert=function(wsFrom){
		if (!this.workbook.bUndoChanges && !this.workbook.bRedoChanges) {
			for (var i = 0; i < wsFrom.pivotTables.length; ++i) {
				this.insertPivotTable(wsFrom.pivotTables[i].cloneShallow(), true);
			}
		}
	};
	Worksheet.prototype.copyFromFormulas=function(renameParams, renameSheetMap) {
		//change cell formulas
		var t = this;
		var oldNewArrayFormulaMap = [];
		this._forEachCell(function(cell) {
			if (cell.isFormula()) {
				var parsed, notMainArrayCell;
				if (cell.transformSharedFormula()) {
					parsed = cell.getFormulaParsed();
				} else {
					parsed = cell.getFormulaParsed();
					if(parsed.getArrayFormulaRef()) {//***array-formula***
						var listenerId = parsed.getListenerId();
						//formula-array: parsed object one of all array cells
						if(oldNewArrayFormulaMap[listenerId]) {
							parsed = oldNewArrayFormulaMap[listenerId];
							notMainArrayCell = true;
						} else {
							parsed = parsed.clone(null, new CCellWithFormula(t, cell.nRow, cell.nCol), t);
							oldNewArrayFormulaMap[listenerId] = parsed;
						}
					} else {
						parsed = parsed.clone(null, new CCellWithFormula(t, cell.nRow, cell.nCol), t);
					}
				}

				if(renameSheetMap && AscCommon.History.Is_On()) {
					//пишем в историю для того, чтобы для случая redo не делать отложенное действия для всех листов
					var _oldF = parsed.Formula;
					parsed.parse(null, null, null, null, renameSheetMap);
					var _newF = parsed.Formula;
					if(_oldF !== _newF) {
						var DataOld = cell.getValueData();
						var DataNew = cell.getValueData();
						DataNew.formula = _newF;
						if (false == DataOld.isEqual(DataNew)) {
							AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue,
								cell.ws.getId(), new Asc.Range(cell.nCol, cell.nRow, cell.nCol, cell.nRow),
								new UndoRedoData_CellSimpleData(cell.nRow, cell.nCol, DataOld, DataNew));
						}
					}
				}

				if(!notMainArrayCell) {
					parsed.renameSheetCopy(renameParams);
					parsed.setFormulaString(parsed.assemble(true));
				}
				cell.setFormulaInternal(parsed, true);
				t.workbook.dependencyFormulas.addToBuildDependencyCell(cell);
			}
		});
	};
	Worksheet.prototype.cloneSelection = function (start, selectRange) {
		if (start) {
			this.copySelection = this.selectionRange.clone();
			if (selectRange) {
				this.selectionRange.assign2(selectRange);
			} else {
				this.selectionRange = null;
			}

		} else {
			if (this.copySelection) {
				this.selectionRange = this.copySelection;
			}
			this.copySelection = null;
		}
	};
	Worksheet.prototype.getSelection = function () {
		return this.copySelection || this.selectionRange;
	};

	Worksheet.prototype.copyObjects = function (oNewWs, renameParams) {
		var i;
		if (null != this.Drawings && this.Drawings.length > 0) {
			var drawingObjects = new AscFormat.DrawingObjects();
			oNewWs.Drawings = [];
			for (i = 0; i < this.Drawings.length; ++i) {
				var _isSlicer = this.Drawings[i].graphicObject.getObjectType() === AscDFH.historyitem_type_SlicerView;
				if (_isSlicer && renameParams && !renameParams.slicerNameMap[this.Drawings[i].graphicObject.name]) {
					renameParams.copySlicerError = true;
					continue;
				}

				var drawingObject = drawingObjects.cloneDrawingObject(this.Drawings[i]);
				drawingObject.graphicObject = this.Drawings[i].graphicObject.copy(undefined);
				if (_isSlicer && renameParams) {
					drawingObject.graphicObject.setName(renameParams.slicerNameMap[drawingObject.graphicObject.name]);
				}

				drawingObject.graphicObject.setWorksheet(oNewWs);
				drawingObject.graphicObject.addToDrawingObjects();
				var drawingBase = this.Drawings[i];
				drawingObject.graphicObject.setDrawingBaseCoords(drawingBase.from.col, drawingBase.from.colOff,
																 drawingBase.from.row, drawingBase.from.rowOff, drawingBase.to.col, drawingBase.to.colOff,
																 drawingBase.to.row, drawingBase.to.rowOff, drawingBase.Pos.X, drawingBase.Pos.Y, drawingBase.ext.cx,
																 drawingBase.ext.cy);
				if(drawingObject.graphicObject.setDrawingBaseType){
					drawingObject.graphicObject.setDrawingBaseType(drawingBase.Type);
					if(drawingBase.Type === AscCommon.c_oAscCellAnchorType.cellanchorTwoCell) {
						drawingObject.graphicObject.setDrawingBaseEditAs(drawingObject.editAs);
					}
				}
				oNewWs.Drawings[oNewWs.Drawings.length - 1] = drawingObject;
			}
			var aRefsToChange = [];
			var sNewName = parserHelp.getEscapeSheetName(oNewWs.sName);
			var aRanges = [new AscCommonExcel.Range(this, 0, 0, gc_nMaxRow0, gc_nMaxCol0)];
			oNewWs.handleDrawings(function(oDrawing) {
				if(oDrawing.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
					oDrawing.collectIntersectionRefs(aRanges, aRefsToChange);
				}
			});
			var sOldName = parserHelp.getEscapeSheetName(this.sName);
			for(var nRef = 0; nRef < aRefsToChange.length; ++nRef) {
				aRefsToChange[nRef].handleOnChangeSheetName(sOldName, sNewName);
			}
			drawingObjects.pushToAObjects(oNewWs.Drawings);
		}

		var newSparkline;
		for (i = 0; i < this.aSparklineGroups.length; ++i) {
			newSparkline = this.aSparklineGroups[i].clone();
			newSparkline.setWorksheet(oNewWs, this);
			oNewWs.addSparklineGroups(newSparkline);
		}
	};
	Worksheet.prototype.initColumn = function (column) {
		if (column) {
			if (null !== column.width && 0 !== column.width) {
				column.widthPx = this.modelColWidthToColWidth(column.width);
				column.charCount = this.colWidthToCharCount(column.widthPx);
			} else {
				column.widthPx = column.charCount = null;
			}
		}
	};
	Worksheet.prototype.initColumns = function () {
		this.initColumn(this.oAllCol);
		this.aCols.forEach(this.initColumn, this);
	};
	Worksheet.prototype.initPostOpen = function (handlers, tableIds, sheetIds, pivotCaches, oNumFmts, dxfsOpen) {
		var t = this;
		this.PagePrintOptions.init();
		this.headerFooter.init();
		if (this.dataValidations) {
			this.dataValidations.init(this);
		}

		// Sheet Views
		if (0 === this.sheetViews.length) {
			// Даже если не было, создадим
			this.sheetViews.push(new AscCommonExcel.asc_CSheetViewSettings());
		}
		//this.setTableFormulaAfterOpen();
		this.hiddenManager.initPostOpen();
		this.oSheetFormatPr.correction();

		this.handlers = handlers;
		this._setHandlersTablePart();
		this.aSlicers.forEach(function(elem){
			elem.initPostOpen(tableIds, sheetIds);
		});
		this.aNamedSheetViews.forEach(function(elem){
			elem.initPostOpen(tableIds, t);
		});
		this.aCellWatches.forEach(function(elem){
			elem.initPostOpen(t);
		});
		this.userProtectedRanges.forEach(function(elem){
			elem.initPostOpen(t);
		});
		this.pivotTables.forEach(function(pivotTable){
			pivotTable.initPostOpenZip(oNumFmts, dxfsOpen);
		});
		this.aSlicers.forEach(function(slicer){
			slicer.initPostOpenZip(pivotCaches);
		});
	};
	Worksheet.prototype._getValuesForConditionalFormatting = function(ranges, numbers) {
		var res = [];
		for (var i = 0; i < ranges.length; ++i) {
			var elem = ranges[i];
			var range = this.getRange3(elem.r1, elem.c1, elem.r2, elem.c2);
			res = res.concat(range._getValues(numbers));
		}
		return res;
	};
	Worksheet.prototype._isConditionalFormattingIntersect = function(range, ranges) {
		for (var i = 0; i < ranges.length; ++i) {
			if (range.isIntersect(ranges[i])) {
				return true;
			}
		}
		return false;
	};
	Worksheet.prototype.setDirtyConditionalFormatting = function(range) {
		if (!range) {
			range = new AscCommonExcel.MultiplyRange([new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0)]);
		} else if (range && range.isNull && range.isNull()) {
			return;
		} else if (range.ranges && range.getUnionRange) {
			//объединяю в один
			range = new AscCommonExcel.MultiplyRange([range.getUnionRange()]);
		}
		if (this.updateConditionalFormattingRange) {
			this.updateConditionalFormattingRange.union2(range);
			this.updateConditionalFormattingRange = new AscCommonExcel.MultiplyRange([this.updateConditionalFormattingRange.getUnionRange()]);
		} else {
			this.updateConditionalFormattingRange = range.clone();
		}
	};
	Worksheet.prototype._updateConditionalFormatting = function() {
		if (!this.updateConditionalFormattingRange) {
			return;
		}
		var range = this.updateConditionalFormattingRange;
		this.updateConditionalFormattingRange = null;
		var t = this;
		let _aRules = [];
		this.forEachConditionalFormattingRules(function (val) {
			_aRules.push(val);
		});
		var aRules = _aRules.sort(function(v1, v2) {
			return v2.priority - v1.priority;
		});
		var oGradient1, oGradient2, aWeights, oRule, oRuleElement, bboxCf, formulaParent, parsed1, parsed2;
		var o, l, cell, ranges, values, value, tmp, dxf, compareFunction, nc, sum;
		this.sheetMergedStyles.clearConditionalStyle(range);
		var getCacheFunction = function(rule, setFunc) {
			var cache = {
				cache: {},
				get: function(row, col) {
					var cacheVal;
					var cacheRow = this.cache[row];
					if (!cacheRow) {
						cacheRow = {};
						this.cache[row] = cacheRow;
					} else {
						cacheVal = cacheRow[col];
					}
					if(undefined ===cacheVal){
						cacheVal = this.set(row, col);
						cacheRow[col] = cacheVal;
					}
					return cacheVal;
				},
				set: function(row, col) {
					if(rule){
						return setFunc(row, col) ? rule.dxf : null;
					} else {
						return setFunc(row, col);
					}
				}
			};
			return function(row, col) {
				return cache.get(row, col);
			};
		};
		for (var i = 0; i < aRules.length; ++i) {
			oRule = aRules[i];
			ranges = oRule.ranges;
			if (this._isConditionalFormattingIntersect(range, ranges)) {
				oRuleElement = oRule.asc_getColorScaleOrDataBarOrIconSetRule();
					if (oRuleElement) {
						if (Asc.ECfType.colorScale !== oRuleElement.type) {
							continue;
						}
						values = this._getValuesForConditionalFormatting(ranges, true);

						// ToDo CFVO Type formula (page 2681)
						l = oRuleElement.aColors.length;
						if (0 < values.length && 2 <= l) {
							aWeights = [];
							oGradient1 = new AscCommonExcel.CGradient(oRuleElement.aColors[0], oRuleElement.aColors[1]);
							aWeights.push(oRule.getMin(values, t), oRule.getMax(values, t));
							if (2 < l) {
								oGradient2 = new AscCommonExcel.CGradient(oRuleElement.aColors[1], oRuleElement.aColors[2]);
								aWeights.push(oRule.getMid(values, t));

								aWeights.sort(AscCommon.fSortAscending);
								oGradient1.init(aWeights[0], aWeights[1]);
								oGradient2.init(aWeights[1], aWeights[2]);
							} else {
								oGradient2 = null;
								aWeights.sort(AscCommon.fSortAscending);
								oGradient1.init(aWeights[0], aWeights[1]);
							}

							compareFunction = (function (oGradient1, oGradient2) {
								return function (row, col) {
									var val, color, gradient;
									t._getCellNoEmpty(row, col, function (cell) {
										val = cell && cell.getNumberValue();
									});
									dxf = null;
									if (null !== val) {
										dxf = new AscCommonExcel.CellXfs();
										gradient = oGradient2 ? oGradient2 : oGradient1;
										if (val >= gradient.max) {
											color = gradient.getMaxColor();
										} else if (val <= oGradient1.min) {
											color = oGradient1.getMinColor();
										} else {
											gradient = (oGradient2 && val > oGradient1.max) ? oGradient2 : oGradient1;
											color = gradient.calculateColor(val);
										}
										dxf.fill = new AscCommonExcel.Fill();
										dxf.fill.fromColor(color);
										dxf = g_StyleCache.addXf(dxf, true);
									}
									return dxf;
								};
							})(oGradient1, oGradient2);
						}
					} else if (Asc.ECfType.dataBar === oRule.type) {
						continue;
					} else if (Asc.ECfType.top10 === oRule.type) {
						if (oRule.rank > 0 && oRule.dxf) {
							nc = 0;
							values = this._getValuesForConditionalFormatting(ranges, false);
							o = oRule.bottom ? Number.MAX_VALUE : -Number.MAX_VALUE;
							for (cell = 0; cell < values.length; ++cell) {
								value = values[cell];
								if (CellValueType.Number === value.type && !isNaN(tmp = parseFloat(value.v))) {
									++nc;
									value.v = tmp;
								} else {
									value.v = o;
								}
							}
							values.sort((function(condition) {
								return function(v1, v2) {
									return condition * (v2.v - v1.v);
								}
							})(oRule.bottom ? -1 : 1));

							nc = Math.max(1, oRule.percent ? Math.floor(nc * oRule.rank / 100) : oRule.rank);
							var threshold = values.length >= nc ? values[nc - 1].v : o;
							compareFunction = (function(rule, threshold) {
								return function(row, col) {
									var val;
									t._getCellNoEmpty(row, col, function(cell) {
										val = cell ? cell.getNumberValue() : null;
									});
									return (null !== val && (rule.bottom ? val <= threshold : val >= threshold)) ? rule.dxf : null;
								};
							})(oRule, threshold);
						}
					} else if (Asc.ECfType.aboveAverage === oRule.type) {
						if (!oRule.dxf) {
							continue;
						}
						values = this._getValuesForConditionalFormatting(ranges, false);
						sum = 0;
						nc = 0;
						for (cell = 0; cell < values.length; ++cell) {
							value = values[cell];
							if (CellValueType.Number === value.type && !isNaN(tmp = parseFloat(value.v))) {
								++nc;
								value.v = tmp;
								sum += tmp;
							} else {
								value.v = null;
							}
						}

						tmp = sum / nc;

						var stdDev;
						if (oRule.hasStdDev()) {
							var sum2 = 0;
						 for (cell = 0; cell < values.length; ++cell) {
						 value = values[cell];
						 if (null !== value.v) {
									sum2 += (value.v - tmp) * (value.v - tmp);
						 }
						 }
							stdDev = Math.sqrt(sum2 / nc);
						}

						compareFunction = (function(rule, average, stdDev) {
							return function(row, col) {
								var val;
								t._getCellNoEmpty(row, col, function(cell) {
									val = cell ? cell.getNumberValue() : null;
								});
								return (null !== val && rule.getAverage(val, average, stdDev)) ? rule.dxf : null;
							};
						})(oRule, tmp, stdDev);
					} else {
						if (!oRule.dxf) {
							continue;
						}
						switch (oRule.type) {
							case Asc.ECfType.duplicateValues:
							case Asc.ECfType.uniqueValues:
								o = getUniqueKeys(this._getValuesForConditionalFormatting(ranges, false));
								compareFunction = (function(rule, obj, condition) {
									return function(row, col) {
										var val;
										t._getCellNoEmpty(row, col, function(cell) {
											val = cell ? cell.getValueWithoutFormat() : "";
										});
										return (val.length > 0 ? condition === obj[val] : false) ? rule.dxf : null;
									};
								})(oRule, o, oRule.type === Asc.ECfType.duplicateValues);
								break;
							case Asc.ECfType.containsText:
							case Asc.ECfType.notContainsText:
							case Asc.ECfType.beginsWith:
							case Asc.ECfType.endsWith:
								var operator;
								switch (oRule.type) {
									case Asc.ECfType.containsText:
										operator = AscCommonExcel.ECfOperator.Operator_containsText;
										break;
									case Asc.ECfType.notContainsText:
										operator = AscCommonExcel.ECfOperator.Operator_notContains;
										break;
									case Asc.ECfType.beginsWith:
										operator = AscCommonExcel.ECfOperator.Operator_beginsWith;
										break;
									case Asc.ECfType.endsWith:
										operator = AscCommonExcel.ECfOperator.Operator_endsWith;
										break;
								}
								formulaParent = new AscCommonExcel.CConditionalFormattingFormulaParent(this, oRule, true);
								oRuleElement =  oRule.getFormulaCellIs(Asc.ECfType.notContainsText === oRule.type || Asc.ECfType.containsText === oRule.type);
								parsed1 = oRuleElement && oRuleElement.getFormula && oRuleElement.getFormula(this, formulaParent);
								if (parsed1 && parsed1.hasRelativeRefs()) {
									bboxCf = oRule.getBBox();
									compareFunction = getCacheFunction(oRule, (function(rule, operator, formulaParent, rowLT, colLT) {
										return function(row, col) {
											var offset = new AscCommon.CellBase(row - rowLT, col - colLT);
											var bboxCell = new Asc.Range(col, row, col, row);
											var v1 = rule.getValueCellIs(t, formulaParent, bboxCell, offset, false);
											var res;
											t._getCellNoEmpty(row, col, function(cell) {
												res = rule.cellIs(operator, cell, v1) ? rule.dxf : null;
											});
											return res;
										};
									})(oRule, operator,
										new AscCommonExcel.CConditionalFormattingFormulaParent(this, oRule, true),
										bboxCf ? bboxCf.r1 : 0, bboxCf ? bboxCf.c1 : 0));
								} else {
									compareFunction = (function(rule, operator, v1) {
										return function(row, col) {
											var res;
											t._getCellNoEmpty(row, col, function(cell) {
												res = rule.cellIs(operator, cell, v1) ? rule.dxf : null;
											});
											return res;
										};
									})(oRule, operator, oRule.getValueCellIs(this));
								}
								break;
							case Asc.ECfType.containsErrors:
								compareFunction = (function(rule) {
									return function(row, col) {
										var val;
										t._getCellNoEmpty(row, col, function(cell) {
											val = (cell ? CellValueType.Error === cell.getType() : false);
										});
										return val ? rule.dxf : null;
									};
								})(oRule);
								break;
							case Asc.ECfType.notContainsErrors:
								compareFunction = (function(rule) {
									return function(row, col) {
										var val;
										t._getCellNoEmpty(row, col, function(cell) {
											val = (cell ? CellValueType.Error !== cell.getType() : true);
										});
										return val ? rule.dxf : null;
									};
								})(oRule);
								break;
							case Asc.ECfType.containsBlanks:
								compareFunction = (function(rule) {
									return function(row, col) {
										var val;
										t._getCellNoEmpty(row, col, function(cell) {
											if (cell) {
												//todo LEN(TRIM(A1))=0
												val = "" === cell.getValueWithoutFormat().replace(/^ +| +$/g, '');
											} else {
												val = true;
											}
										});
										return val ? rule.dxf : null;
									};
								})(oRule);
								break;
							case Asc.ECfType.notContainsBlanks:
								compareFunction = (function(rule) {
									return function(row, col) {
										var val;
										t._getCellNoEmpty(row, col, function(cell) {
											if (cell) {
												//todo LEN(TRIM(A1))=0
												val = "" !== cell.getValueWithoutFormat().replace(/^ +| +$/g, '');
											} else {
												val = false;
											}
										});
										return val ? rule.dxf : null;
									};
								})(oRule);
								break;
							case Asc.ECfType.timePeriod:
								if (oRule.timePeriod) {
									compareFunction = (function(rule, period) {
										return function(row, col) {
											var val;
											t._getCellNoEmpty(row, col, function(cell) {
												val = cell ? cell.getValueWithoutFormat() : "";
											});
											var n = parseFloat(val);
											return (period.start <= n && n < period.end) ? rule.dxf : null;
										};
									})(oRule, oRule.getTimePeriod());
								} else {
									continue;
								}
								break;
							case Asc.ECfType.cellIs:
								formulaParent = new AscCommonExcel.CConditionalFormattingFormulaParent(this, oRule, true);
								oRuleElement = oRule.aRuleElements[0];
								parsed1 = oRuleElement && oRuleElement.getFormula && oRuleElement.getFormula(this, formulaParent);
								oRuleElement = oRule.aRuleElements[1];
								parsed2 = oRuleElement && oRuleElement.getFormula && oRuleElement.getFormula(this, formulaParent);
								if ((parsed1 && parsed1.hasRelativeRefs()) || (parsed2 && parsed2.hasRelativeRefs())) {
									bboxCf = oRule.getBBox();
									compareFunction = getCacheFunction(oRule, (function(rule, ruleElem1, ruleElem2, formulaParent, rowLT, colLT) {
										return function(row, col) {
											var offset = new AscCommon.CellBase(row - rowLT, col - colLT);
											var bboxCell = new Asc.Range(col, row, col, row);
											var v1 = ruleElem1 && ruleElem1.getValue(t, formulaParent, bboxCell, offset, false);
											var v2 = ruleElem2 && ruleElem2.getValue(t, formulaParent, bboxCell, offset, false);
											var res;
											t._getCellNoEmpty(row, col, function(cell) {
												res = rule.cellIs(rule.operator, cell, v1, v2) ? rule.dxf : null;
											});
											return res;
										};
									})(oRule, oRule.aRuleElements[0], oRule.aRuleElements[1],
										new AscCommonExcel.CConditionalFormattingFormulaParent(this, oRule, true),
										bboxCf ? bboxCf.r1 : 0, bboxCf ? bboxCf.c1 : 0));
								} else {
									compareFunction = (function(rule, v1, v2) {
										return function(row, col) {
											var res;
											t._getCellNoEmpty(row, col, function(cell) {
												res = rule.cellIs(rule.operator, cell, v1, v2) ? rule.dxf : null;
											});
											return res;
										};
									})(oRule, oRule.aRuleElements[0] && oRule.aRuleElements[0].getValue(this),
										oRule.aRuleElements[1] && oRule.aRuleElements[1].getValue(this));
								}
								break;
							case Asc.ECfType.expression:
								bboxCf = oRule.getBBox();
								compareFunction = getCacheFunction(oRule, (function(rule, formulaCF, formulaParent, rowLT, colLT) {
									return function(row, col) {
										var offset = new AscCommon.CellBase(row - rowLT, col - colLT);
										var bboxCell = new Asc.Range(col, row, col, row);
										var res = formulaCF && formulaCF.getValue(t, formulaParent, bboxCell, offset, true);
										if (res && res.tocBool) {
											res = res.tocBool();
											if (res && res.toBool) {
												return res.toBool();
											}
										}
										return false;
									};
								})(oRule, oRule.aRuleElements[0],
									new AscCommonExcel.CConditionalFormattingFormulaParent(this, oRule, true),
									bboxCf ? bboxCf.r1 : 0, bboxCf ? bboxCf.c1 : 0));
								break;
							default:
								break;
						}
					}
					if (compareFunction) {
						this.sheetMergedStyles.setConditionalStyle(oRule, ranges, compareFunction);
					}
				}
			}
	};
	Worksheet.prototype._forEachRow = function(fAction) {
		this.getRange3(0, 0, gc_nMaxRow0, 0)._foreachRowNoEmpty(fAction);
	};
	Worksheet.prototype._forEachCol = function(fAction) {
		this.getRange3(0, 0, 0, gc_nMaxCol0)._foreachColNoEmpty(fAction);
	};
	Worksheet.prototype._forEachColData = function(fAction) {
		for (var i = 0; i < this.cellsByCol.length; ++i) {
			var sheetMemory = this.cellsByCol[i];
			if (sheetMemory) {
				fAction(sheetMemory, i);
			}
		}
	};
	Worksheet.prototype._forEachCell = function(fAction) {
		this.getRange3(0, 0, gc_nMaxRow0, gc_nMaxCol0)._foreachNoEmpty(fAction);
	};
	Worksheet.prototype.Get_Id = Worksheet.prototype.getId=function(){
		return this.Id;
	};
	Worksheet.prototype.getIndex=function(){
		return this.index;
	};
	Worksheet.prototype.getName=function(){
		return this.sName != null && this.sName.length > 0 ? this.sName : "";
	};
	Worksheet.prototype.setName=function(name){
		if(name.length <= g_nSheetNameMaxLength)
		{
			var lastName = this.sName;
			AscCommon.History.Create_NewPoint();
			var prepared = this.workbook.dependencyFormulas.prepareChangeSheet(this.getId(), {rename: {from: lastName, to: name}});
			this.sName = name;
			this.workbook.dependencyFormulas.changeSheet(prepared);

			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_Rename, this.getId(), null, new UndoRedoData_FromTo(lastName, name));

			this.workbook.dependencyFormulas.calcTree();
			if (!this.workbook.bUndoChanges && !this.workbook.bRedoChanges) {
				this.workbook.handlers && this.workbook.handlers.trigger("updateCellWatches");
			}
			if (this.workbook.handlers) {
				this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetRename, this.index, name, this.getId());
			}
		} else {
			console.log(new Error('The sheet name must be less than 31 characters.'));
		}
	};
	Worksheet.prototype.getTabColor=function(){
		return this.sheetPr && this.sheetPr.TabColor ? Asc.colorObjToAscColor(this.sheetPr.TabColor) : null;
	};
	Worksheet.prototype.setTabColor=function(color){
		if (!this.sheetPr)
			this.sheetPr = new AscCommonExcel.asc_CSheetPr();

		AscCommon.History.Create_NewPoint();
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetTabColor, this.getId(), null,
					new UndoRedoData_FromTo(this.sheetPr.TabColor ? this.sheetPr.TabColor.clone() : null, color ? color.clone() : null));

		this.sheetPr.TabColor = color;
		if (!this.workbook.bUndoChanges && !this.workbook.bRedoChanges)
			this.workbook.handlers && this.workbook.handlers.trigger("asc_onUpdateTabColor", this.getIndex());
	};
	Worksheet.prototype.rebuildTabColor = function() {
		if (this.sheetPr && this.sheetPr.TabColor) {
			this.workbook.handlers && this.workbook.handlers.trigger("asc_onUpdateTabColor", this.getIndex());
		}
	};
	Worksheet.prototype.getHidden=function(){
		return true === this.bHidden;
	};
	Worksheet.prototype.setHidden = function (hidden) {
		var bOldHidden = this.bHidden, wb = this.workbook, wsActive = wb.getActiveWs(), oVisibleWs = null;
		this.bHidden = hidden;
		if (bOldHidden != hidden) {
			if (true == this.bHidden && this.getIndex() == wsActive.getIndex()) {
				oVisibleWs = wb.findSheetNoHidden(this.getIndex());
			} else if (false == this.bHidden && this.getIndex() !== wsActive.getIndex()) {
				oVisibleWs = this;
			}
			if (null != oVisibleWs) {
				var nNewIndex = oVisibleWs.getIndex();
				wb.setActive(nNewIndex);
				if (!wb.bUndoChanges && !wb.bRedoChanges) {
					wb.handlers && wb.handlers.trigger("undoRedoHideSheet", nNewIndex);
				}
			}
			AscCommon.History.Create_NewPoint();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_Hide, this.getId(), null, new UndoRedoData_FromTo(bOldHidden, hidden));
			if (null != oVisibleWs) {
				AscCommon.History.SetSheetUndo(wsActive.getId());
				AscCommon.History.SetSheetRedo(oVisibleWs.getId());
			}
		}
	};
	Worksheet.prototype.getSheetView = function () {
		return this.sheetViews[0];
	};
	Worksheet.prototype.getSheetViewSettings = function (bNotClone) {
		return bNotClone ? this.sheetViews[0] : this.sheetViews[0].clone();
	};
	Worksheet.prototype.setDisplayGridlines = function (value) {
		var view = this.sheetViews[0];
		if (value !== view.showGridLines) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetDisplayGridlines,
				this.getId(), null, new UndoRedoData_FromTo(view.showGridLines, value));
			view.showGridLines = value;

			this.workbook.handlers && this.workbook.handlers.trigger("changeSheetViewSettings", this.getId(), AscCH.historyitem_Worksheet_SetDisplayGridlines);
			if (!this.workbook.bUndoChanges && !this.workbook.bRedoChanges) {
				this.workbook.handlers && this.workbook.handlers.trigger("asc_onUpdateSheetViewSettings");
			}
		}
	};
	Worksheet.prototype.setDisplayHeadings = function (value) {
		var view = this.sheetViews[0];
		if (value !== view.showRowColHeaders) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetDisplayHeadings,
				this.getId(), null, new UndoRedoData_FromTo(view.showRowColHeaders, value));
			view.showRowColHeaders = value;

			this.workbook.handlers && this.workbook.handlers.trigger("changeSheetViewSettings", this.getId(), AscCH.historyitem_Worksheet_SetDisplayHeadings);
			if (!this.workbook.bUndoChanges && !this.workbook.bRedoChanges) {
				this.workbook.handlers && this.workbook.handlers.trigger("asc_onUpdateSheetViewSettings");
			}
		}
	};
	Worksheet.prototype.setShowZeros = function (value) {
		var view = this.sheetViews[0];
		if (value !== view.showZeros) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetShowZeros,
				this.getId(), null, new UndoRedoData_FromTo(view.showZeros, value));
			view.showZeros = value;

			//TODO
			this.workbook.handlers && this.workbook.handlers.trigger("changeSheetViewSettings", this.getId(), AscCH.historyitem_Worksheet_SetDisplayHeadings);
			if (!this.workbook.bUndoChanges && !this.workbook.bRedoChanges) {
				this.workbook.handlers && this.workbook.handlers.trigger("asc_onUpdateSheetViewSettings");
			}
		}
	};
	Worksheet.prototype.setShowFormulas = function (value) {
		var view = this.sheetViews[0];
		if (value !== view.showFormulas) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetShowFormulas,
				this.getId(), new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0), new UndoRedoData_FromTo(view.showFormulas, value));
			view.showFormulas = value;

			this.workbook.handlers.trigger("changeSheetViewSettings", this.getId(), AscCH.historyitem_Worksheet_SetShowFormulas);
			if (!this.workbook.bCollaborativeChanges) {
				this.workbook.handlers.trigger("asc_onUpdateFormulasViewSettings");
			}
		}
	};
	Worksheet.prototype.setRightToLeft = function (value, addToHistory) {
		var view = this.sheetViews[0];
		if (value !== view.rightToLeft) {
			if (addToHistory) {
				AscCommon.History.Create_NewPoint();
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetRightToLeft,
					this.getId(), new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0), new UndoRedoData_FromTo(view.rightToLeft, value));
			}
			view.rightToLeft = value;

			this.workbook.handlers.trigger("changeSheetViewSettings", this.getId(), AscCH.historyitem_Worksheet_SetRightToLeft);
			if (!this.workbook.bCollaborativeChanges) {
				this.workbook.handlers.trigger("asc_onUpdateSheetViewSettings");
			}
		}
	};
	Worksheet.prototype.getShowFormulas = function () {
		var view = this.sheetViews[0];
		return view && view.showFormulas;
	};
	Worksheet.prototype.getRowsCount=function(){
		var result = this.nRowsCount;
		var pane = this.sheetViews.length && this.sheetViews[0].pane;
		if (pane && pane.topLeftFrozenCell)
			result = Math.max(result, pane.topLeftFrozenCell.getRow0());
		return result;
	};
	Worksheet.prototype.getSheetViewType=function(){
		return this.sheetViews && this.sheetViews.length && this.sheetViews[0].view;
	};
	Worksheet.prototype.removeRows=function(start, stop, bExcludeHiddenRows){
		var removeRowsArr = bExcludeHiddenRows ? this._getNoHiddenRowsArr(start, stop) : [{start: start, stop: stop}];
		for(var i = removeRowsArr.length - 1; i >= 0; i--) {
			var oRange = this.getRange(new CellAddress(removeRowsArr[i].start, 0, 0), new CellAddress(removeRowsArr[i].stop, gc_nMaxCol0, 0));
			oRange.deleteCellsShiftUp();
		}
	};
	Worksheet.prototype._getNoHiddenRowsArr=function(start, stop){
		var res = [];
		var elem = null;
		for (var i = start; i <= stop; i++) {
			if (this.getRowHidden(i)) {
				if (elem) {
					res.push(elem);
					elem = null;
				}
			} else {
				if (!elem) {
					elem = {};
					elem.start = i;
					elem.stop = i;
				} else {
					elem.stop++;
				}
				if (i === stop) {
					res.push(elem);
				}
			}
		}
		return res;
	};
	Worksheet.prototype._updateFormulasParents = function(r1, c1, r2, c2, bbox, offset, shiftedShared) {
		var t = this;
		var cellWithFormula;
		var shiftedArrayFormula = {};
		this.getRange3(r1, c1, r2, c2)._foreachNoEmpty(function(cell){
			var newNRow = cell.nRow + offset.row;
			var newNCol = cell.nCol + offset.col;
			var bHor = 0 !== offset.col;
			var toDelete = offset.col < 0 || offset.row < 0;

			if (cell.isFormula()) {
				var processed = c_oSharedShiftType.NeedTransform;
				var parsed = cell.getFormulaParsed();
				var shared = parsed.getShared();
				var arrayFormula = parsed.getArrayFormulaRef();
				var formulaRefObj = null;
				if (shared) {
					processed = shiftedShared[parsed.getListenerId()];
					var isPreProcessed = c_oSharedShiftType.PreProcessed === processed;
					if (!processed || isPreProcessed) {
						if (!processed) {
							var bboxShift = AscCommonExcel.shiftGetBBox(bbox, bHor);
							//if shared not completly in shift range - transform
							//if shared intersect delete range - transform
							if (bboxShift.containsRange(shared.ref) && (!toDelete || !bbox.isIntersect(shared.ref))) {
								processed = c_oSharedShiftType.Processed;
							} else {
								processed = c_oSharedShiftType.NeedTransform;
							}
						} else if(isPreProcessed) {
							//At PreProcessed stage all required formula was transformed. here we need to shift shared
							processed = c_oSharedShiftType.Processed;
						}
						if (c_oSharedShiftType.Processed === processed) {
							var newRef = shared.ref.clone();
							newRef.forShift(bbox, offset, t.workbook.bUndoChanges);
							parsed.setSharedRef(newRef, !isPreProcessed);
							t.workbook.dependencyFormulas.addToChangedRange2(t.getId(), newRef);
						}
						shiftedShared[parsed.getListenerId()] = processed;
					}
				} else if(arrayFormula) {
					//***array-formula***
					if(!shiftedArrayFormula[parsed.getListenerId()] && parsed.checkFirstCellArray(cell)) {
						shiftedArrayFormula[parsed.getListenerId()] = 1;
						var newArrayRef = arrayFormula.clone();
						newArrayRef.setOffset(offset);
						parsed.setArrayFormulaRef(newArrayRef);
					} else {
						processed = c_oSharedShiftType.Processed;
					}
				}

				if (c_oSharedShiftType.NeedTransform === processed) {
					var isTransform = cell.transformSharedFormula();
					parsed = cell.getFormulaParsed();
					if (isTransform) {
						parsed.buildDependencies();
					}

					cellWithFormula = parsed.getParent();
					cellWithFormula.nRow = newNRow;
					cellWithFormula.nCol = newNCol;
					t.workbook.dependencyFormulas.addToChangedCell(cellWithFormula);
				}
			}
			t.cellsByColRowsCount = Math.max(t.cellsByColRowsCount, newNRow + 1);
			t.nRowsCount = Math.max(t.nRowsCount, t.cellsByColRowsCount);
			t.setColsCount(Math.max(t.nColsCount, newNCol + 1));
		});
	};
	Worksheet.prototype._removeRows=function(start, stop){
		var t = this;
		this.workbook.dependencyFormulas.lockRecal();
		AscCommon.History.Create_NewPoint();
		//start, stop 0 based
		var nDif = -(stop - start + 1);
		var oActualRange = new Asc.Range(0, start, gc_nMaxCol0, stop);
		var offset = new AscCommon.CellBase(nDif, 0);
		//renameDependencyNodes before move cells to store current location in history
		var renameRes = this.renameDependencyNodes(offset, oActualRange);
		var redrawTablesArr = this.autoFilters.insertRows("delCell", oActualRange, c_oAscDeleteOptions.DeleteRows);
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.updatePivotOffset(oActualRange, offset);
			this.updateUserProtectedRangesOffset(oActualRange, offset);
		}
		if (false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			this.updateSortStateOffset(oActualRange, offset);
			this.updateSparklineGroupOffset(oActualRange, offset);
			this.updateConditionalFormattingOffset(oActualRange, offset);
			this.updateProtectedRangeOffset(oActualRange, offset);
			AscCommon.History.LocalChange = false;
		}

		var collapsedInfo = null, lastRowIndex;
		var oDefRowPr = new AscCommonExcel.UndoRedoData_RowProp();
		this.getRange3(start,0,stop,gc_nMaxCol0)._foreachRowNoEmpty(function(row){
			var oOldProps = row.getHeightProp();
			lastRowIndex = row.index;
			if (false === oOldProps.isEqual(oDefRowPr))
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RowProp, t.getId(), row._getUpdateRange(), new UndoRedoData_IndexSimpleProp(row.getIndex(), true, oOldProps, oDefRowPr));
			row.setStyle(null);

			if(!t.workbook.bRedoChanges) {
				if(collapsedInfo !== null && collapsedInfo < row.getOutlineLevel()) {
					collapsedInfo = null;
				}
				if(row.getCollapsed()) {
					collapsedInfo = row.getOutlineLevel();
					t.setCollapsedRow(false, null, row);
				}
			}

		}, function(cell){
			t._removeCell(null, null, cell);
		});

		//ms не удаляет collapsed с удаляемой строки, он наследует это свойство следующей
		if(collapsedInfo !== null && lastRowIndex === stop) {
			this._getRow(stop + 1, function(row) {
				//TODO проверить!!!
				//if(collapsedInfo >= row.getOutlineLevel()) {
					//row.setCollapsed(true);
					t.setCollapsedRow(true, null, row);
				//}
			});
		}

		this._updateFormulasParents(start, 0, gc_nMaxRow0, gc_nMaxCol0, oActualRange, offset, renameRes && renameRes.shiftedShared);
		this.rowsData.deleteRange(start, (-nDif));
		this._forEachColData(function(sheetMemory) {
			sheetMemory.deleteRange(start, (-nDif));
		});
		//notifyChanged after move cells to get new locations(for intersect ranges)
		this.workbook.dependencyFormulas.notifyChanged(renameRes && renameRes.changed);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RemoveRows, this.getId(), new Asc.Range(0, start, gc_nMaxCol0, gc_nMaxRow0), new UndoRedoData_FromToRowCol(true, start, stop));

		this.autoFilters.redrawStylesTables(redrawTablesArr);

		if (this.workbook.handlers) {
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this, new Asc.Range(0, start, gc_nMaxCol0, stop), this.getId());
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.removeRows, this, new Asc.Range(0, start, gc_nMaxCol0, stop), this.getId());
		}

		this.workbook.dependencyFormulas.unlockRecal();

		return true;
	};
	Worksheet.prototype.insertRowsBefore=function(index, count){
		var oRange = this.getRange(new CellAddress(index, 0, 0), new CellAddress(index + count - 1, gc_nMaxCol0, 0));
		oRange.addCellsShiftBottom();
	};
	Worksheet.prototype._getBordersForInsert = function(bbox, bRow) {
		var t = this;
		var borders = {};
		var offsetRow = (bRow && bbox.r1 > 0) ? -1 : 0;
		var offsetCol = (!bRow && bbox.c1 > 0) ? -1 : 0;
		var r2 = bRow ? bbox.r1 : bbox.r2;
		var c2 = !bRow ? bbox.c1 : bbox.c2;
		if(0 !== offsetRow || 0 !== offsetCol){
			this.getRange3(bbox.r1, bbox.c1, r2, c2)._foreachNoEmpty(function(cell) {
				if (cell.xfs && cell.xfs.border) {
					t._getCellNoEmpty(cell.nRow + offsetRow, cell.nCol + offsetCol, function(neighbor) {
						if (neighbor && neighbor.xfs && neighbor.xfs.border) {
							var newBorder = neighbor.xfs.border.clone();
							newBorder.intersect(cell.xfs.border, true);
							borders[bRow ? cell.nCol : cell.nRow] = newBorder;
						}
					});
				}
			});
		}
		return borders;
	};
	Worksheet.prototype._insertRowsBefore=function(index, count){
		var t = this;
		this.workbook.dependencyFormulas.lockRecal();
		var oActualRange = new Asc.Range(0, index, gc_nMaxCol0, index + count - 1);
		AscCommon.History.Create_NewPoint();
		var offset = new AscCommon.CellBase(count, 0);
		//renameDependencyNodes before move cells to store current location in history
		var renameRes = this.renameDependencyNodes(offset, oActualRange);
		var redrawTablesArr = this.autoFilters.insertRows("insCell", oActualRange, c_oAscInsertOptions.InsertColumns);
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.updatePivotOffset(oActualRange, offset);
			this.updateUserProtectedRangesOffset(oActualRange, offset);
		}
		if (false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			this.updateSortStateOffset(oActualRange, offset);
			this.updateSparklineGroupOffset(oActualRange, offset);
			this.updateConditionalFormattingOffset(oActualRange, offset);
			this.updateProtectedRangeOffset(oActualRange, offset);
			AscCommon.History.LocalChange = false;
		}

		this._updateFormulasParents(index, 0, gc_nMaxRow0, gc_nMaxCol0, oActualRange, offset, renameRes.shiftedShared);
		var borders;
		if (index > 0 && !this.workbook.bUndoChanges) {
			borders = this._getBordersForInsert(oActualRange, true);
		}
		//insert new row/cell
		this.rowsData.insertRange(index, count);
		this.nRowsCount = Math.max(this.nRowsCount, this.rowsData.getMaxIndex() + 1);
		this._forEachColData(function(sheetMemory) {
			sheetMemory.insertRange(index, count);
			t.cellsByColRowsCount = Math.max(t.cellsByColRowsCount, sheetMemory.getMaxIndex() + 1);
		});
		this.nRowsCount = Math.max(this.nRowsCount, this.cellsByColRowsCount);
		//copy property from row/cell above
		if (index > 0 && !this.workbook.bUndoChanges)
		{
			this.rowsData.copyRangeByChunk((index - 1), 1, index, count);
			this.nRowsCount = Math.max(this.nRowsCount, this.rowsData.getMaxIndex() + 1);
			this._forEachColData(function(sheetMemory) {
				sheetMemory.copyRangeByChunk((index - 1), 1, index, count);
				t.cellsByColRowsCount = Math.max(t.cellsByColRowsCount, sheetMemory.getMaxIndex() + 1);
			});
			this.nRowsCount = Math.max(this.nRowsCount, this.cellsByColRowsCount);
			//show rows and remain only cell xf property
			this.getRange3(index, 0, index + count - 1, gc_nMaxCol0)._foreachRowNoEmpty(function(row) {
				row.setHidden(false);
			},function(cell) {
				cell.clearDataKeepXf(borders[cell.nCol]);
			});
		}
		//notifyChanged after move cells to get new locations(for intersect ranges)
		this.workbook.dependencyFormulas.notifyChanged(renameRes.changed);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_AddRows, this.getId(), new Asc.Range(0, index, gc_nMaxCol0, gc_nMaxRow0), new UndoRedoData_FromToRowCol(true, index, index + count - 1));

		this.autoFilters.redrawStylesTables(redrawTablesArr);
		if (this.workbook.handlers) {
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this, null, this.getId());
		}

		this.workbook.dependencyFormulas.unlockRecal();
		return true;
	};
	Worksheet.prototype.insertRowsAfter=function(index, count){
		//index 0 based
		return this.insertRowsBefore(index + 1, count);
	};
	Worksheet.prototype.getColsCount=function(){
		var result = this.nColsCount;
		var pane = this.sheetViews.length && this.sheetViews[0].pane;
		if (pane && pane.topLeftFrozenCell)
			result = Math.max(result, pane.topLeftFrozenCell.getCol0());
		return result;
	};
	Worksheet.prototype.removeCols=function(start, stop){
		var oRange = this.getRange(new CellAddress(0, start, 0), new CellAddress(gc_nMaxRow0, stop, 0));
		oRange.deleteCellsShiftLeft();
	};
	Worksheet.prototype._removeCols=function(start, stop){
		var t = this;
		this.workbook.dependencyFormulas.lockRecal();
		AscCommon.History.Create_NewPoint();
		//start, stop 0 based
		var nDif = -(stop - start + 1), i, j, length;
		var oActualRange = new Asc.Range(start, 0, stop, gc_nMaxRow0);
		var offset = new AscCommon.CellBase(0, nDif);
		//renameDependencyNodes before move cells to store current location in history
		var renameRes = this.renameDependencyNodes(offset, oActualRange);
		var redrawTablesArr = this.autoFilters.insertColumn(oActualRange, nDif);
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.updatePivotOffset(oActualRange, offset);
			this.updateUserProtectedRangesOffset(oActualRange, offset);
		}
		if (false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			this.updateSortStateOffset(oActualRange, offset);
			this.updateSparklineGroupOffset(oActualRange, offset);
			this.updateConditionalFormattingOffset(oActualRange, offset);
			this.updateProtectedRangeOffset(oActualRange, offset);
			AscCommon.History.LocalChange = false;
		}

		var collapsedInfo = null, lastRowIndex;
		var oDefColPr = new AscCommonExcel.UndoRedoData_ColProp();
		this.getRange3(0, start, gc_nMaxRow0,stop)._foreachColNoEmpty(function(col){
			var nIndex = col.getIndex();
			var oOldProps = col.getWidthProp();
			if(false === oOldProps.isEqual(oDefColPr))
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ColProp, t.getId(), new Asc.Range(nIndex, 0, nIndex, gc_nMaxRow0), new UndoRedoData_IndexSimpleProp(nIndex, false, oOldProps, oDefColPr));
			col.setStyle(null);

			lastRowIndex = col.index;
			if(!t.workbook.bRedoChanges) {
				if(collapsedInfo !== null && collapsedInfo < col.getOutlineLevel()) {
					collapsedInfo = null;
				}
				if(col.getCollapsed()) {
					collapsedInfo = col.getOutlineLevel();
					t.setCollapsedCol(false, null, col);
				}
			}
		}, function(cell){
			t._removeCell(null, null, cell);
		});

		if(collapsedInfo !== null && lastRowIndex === stop) {
			var curCol = this._getCol(stop + 1);
			//TODO проверить!!!
			if(curCol /*&& collapsedInfo >= curCol.getOutlineLevel()*/) {
				t.setCollapsedCol(true, null, curCol);
			}
		}

		this._updateFormulasParents(0, start, gc_nMaxRow0, gc_nMaxCol0, oActualRange, offset, renameRes.shiftedShared);
		this.cellsByCol.splice(start, stop - start + 1);
		this.aCols.splice(start, stop - start + 1);
		for(i = start, length = this.aCols.length; i < length; ++i)
		{
			var elem = this.aCols[i];
			if(null != elem)
				elem.moveHor(nDif);
		}
		//notifyChanged after move cells to get new locations(for intersect ranges)
		this.workbook.dependencyFormulas.notifyChanged(renameRes.changed);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RemoveCols, this.getId(), new Asc.Range(start, 0, gc_nMaxCol0, gc_nMaxRow0), new UndoRedoData_FromToRowCol(false, start, stop));

		this.autoFilters.redrawStylesTables(redrawTablesArr);
		if (this.workbook.handlers) {
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this, new Asc.Range(start, 0, stop, gc_nMaxRow0), this.getId());
		}

		this.workbook.dependencyFormulas.unlockRecal();

		return true;
	};
	Worksheet.prototype.insertColsBefore=function(index, count){
		var oRange = this.getRange3(0, index, gc_nMaxRow0, index + count - 1);
		oRange.addCellsShiftRight();
	};
	Worksheet.prototype._insertColsBefore=function(index, count){
		this.workbook.dependencyFormulas.lockRecal();
		var oActualRange = new Asc.Range(index, 0, index + count - 1, gc_nMaxRow0);
		AscCommon.History.Create_NewPoint();
		var offset = new AscCommon.CellBase(0, count);
		//renameDependencyNodes before move cells to store current location in history
		var renameRes = this.renameDependencyNodes(offset, oActualRange);
		var redrawTablesArr = this.autoFilters.insertColumn(oActualRange, count);
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.updatePivotOffset(oActualRange, offset);
			this.updateUserProtectedRangesOffset(oActualRange, offset);
		}
		if (false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			this.updateSortStateOffset(oActualRange, offset);
			this.updateSparklineGroupOffset(oActualRange, offset);
			this.updateConditionalFormattingOffset(oActualRange, offset);
			this.updateProtectedRangeOffset(oActualRange, offset);
			AscCommon.History.LocalChange = false;
		}

		this._updateFormulasParents(0, index, gc_nMaxRow0, gc_nMaxCol0, oActualRange, offset, renameRes.shiftedShared);
		var borders;
		if (index > 0 && !this.workbook.bUndoChanges) {
			borders = this._getBordersForInsert(oActualRange, false);
		}
		//remove tail
		this.cellsByCol.splice(gc_nMaxCol0 - count + 1, count);
		for(var i = this.cellsByCol.length - 1; i >= index; --i) {
			this.cellsByCol[i + count] = this.cellsByCol[i];
			this.cellsByCol[i] = undefined;
		}
		this.setColsCount(Math.max(this.nColsCount, this.getColDataLength()));
		this.aCols.splice(gc_nMaxCol0 - count + 1, count);
		for(var i = this.aCols.length - 1; i >= index; --i) {
			this.aCols[i + count] = this.aCols[i];
			this.aCols[i] = undefined;
			if (this.aCols[i + count]) {
				this.aCols[i + count].moveHor(count);
			}
		}
		this.setColsCount(Math.max(this.nColsCount, this.aCols.length));
		if (!this.workbook.bUndoChanges) {
			//copy property from col/cell above
			var oPrevCol = null;
			if (index > 0)
				oPrevCol = this.aCols[index - 1];
			if (null == oPrevCol && null != this.oAllCol)
				oPrevCol = this.oAllCol;
			if (null != oPrevCol) {
				AscCommon.History.LocalChange = true;
				for (var i = index; i < index + count; ++i) {
					var oNewCol =  oPrevCol.clone();
					oNewCol.setHidden(null);
					oNewCol.BestFit = null;
					oNewCol.index = i;
					this.aCols[i] = oNewCol;
				}
				AscCommon.History.LocalChange = false;
			}
			var prevCellsByCol = index > 0 ? this.cellsByCol[index - 1] : null;
			if (prevCellsByCol) {
				for(var i = index; i < index + count; ++i) {
					this.cellsByCol[i] = prevCellsByCol.clone();
				}
				this.setColsCount(Math.max(this.nColsCount, this.getColDataLength()));
				//show rows and remain only cell xf property
				this.getRange3(0, index, gc_nMaxRow0, index + count - 1)._foreachNoEmpty(function(cell) {
					cell.clearDataKeepXf(borders[cell.nRow]);
				});
			}
		}

		//notifyChanged after move cells to get new locations(for intersect ranges)
		this.workbook.dependencyFormulas.notifyChanged(renameRes.changed);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_AddCols, this.getId(), new Asc.Range(index, 0, gc_nMaxCol0, gc_nMaxRow0), new UndoRedoData_FromToRowCol(false, index, index + count - 1));

		this.autoFilters.redrawStylesTables(redrawTablesArr);
		if (this.workbook.handlers) {
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this, null, this.getId());
		}

		this.workbook.dependencyFormulas.unlockRecal();
		return true;
	};
	Worksheet.prototype.insertColsAfter=function(index, count){
		//index 0 based
		return this.insertColsBefore(index + 1, count);
	};
	Worksheet.prototype.getDefaultWidth=function(){
		var dRes = null;
		if (this.oSheetFormatPr.dDefaultColWidth != null) {
			dRes = this.oSheetFormatPr.dDefaultColWidth;
		} else if(null != this.oAllCol && this.oAllCol.CustomWidth) {
			dRes = this.oAllCol.width;
		}
		return dRes;
	};
	Worksheet.prototype.getDefaultFontName=function(){
		return this.workbook.getDefaultFont();
	};
	Worksheet.prototype.getDefaultFontSize=function(){
		return this.workbook.getDefaultSize();
	};
	Worksheet.prototype.getBaseColWidth = function () {
		return this.oSheetFormatPr.nBaseColWidth || 8; // Число символов для дефалтовой ширины (по умолчинию 8)
	};
	Worksheet.prototype.charCountToModelColWidth = function (count) {
		return this.workbook.charCountToModelColWidth(count);
	};
	Worksheet.prototype.modelColWidthToColWidth = function (mcw) {
		return this.workbook.modelColWidthToColWidth(mcw);
	};
	Worksheet.prototype.colWidthToCharCount = function (w) {
		return this.workbook.colWidthToCharCount(w);
	};
	Worksheet.prototype.getColWidth=function(index){
		//index 0 based
		//Результат в пунктах
		var col = this._getColNoEmptyWithAll(index);
		if(null != col && null != col.width)
			return col.width;
		var dResult = this.oSheetFormatPr.dDefaultColWidth;
		if(dResult === undefined || dResult === null || dResult == 0)
		//dResult = (8) + 5;//(EMCA-376.page 1857.)defaultColWidth = baseColumnWidth + {margin padding (2 pixels on each side, totalling 4 pixels)} + {gridline (1pixel)}
			dResult = -1; // calc default width at presentation level
		return dResult;
	};
	Worksheet.prototype.setColWidth=function(width, start, stop){
		width = this.charCountToModelColWidth(width);
		if(0 == width)
			return this.setColHidden(true, start, stop);

		//start, stop 0 based
		if(null == start)
			return;
		if(null == stop)
			stop = start;
		AscCommon.History.Create_NewPoint();
		/*var oSelection = AscCommon.History.GetSelection();
		if(null != oSelection)
		{
			oSelection = oSelection.clone();
			oSelection.assign(start, 0, stop, gc_nMaxRow0);
			AscCommon.History.SetSelection(oSelection);
			AscCommon.History.SetSelectionRedo(oSelection);
		}*/

		var bNotAddCollapsed = true == this.workbook.bUndoChanges || true == this.workbook.bRedoChanges || this.bExcludeCollapsed;
		var _summaryRight = this.sheetPr ? this.sheetPr.SummaryRight : true;
		var oThis = this, prevCol;
		var fProcessCol = function(col){
			if(col.width != width)
			{
				if(_summaryRight && !bNotAddCollapsed && col.getCollapsed()) {
					oThis.setCollapsedCol(false, null, col);
				} else if(!_summaryRight && !bNotAddCollapsed && prevCol && prevCol.getCollapsed()) {
					oThis.setCollapsedCol(false, null, prevCol);
				}
				prevCol = col;

				var oOldProps = col.getWidthProp();
				col.width = width;
				col.CustomWidth = true;
				col.BestFit = null;
				col.setHidden(null);
				oThis.initColumn(col);
				var oNewProps = col.getWidthProp();
				if(false == oOldProps.isEqual(oNewProps))
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ColProp, oThis.getId(),
								col._getUpdateRange(),
								new UndoRedoData_IndexSimpleProp(col.index, false, oOldProps, oNewProps));
			}
		};
		if(0 === start && gc_nMaxCol0 === stop)
		{
			var col = this.getAllCol();
			fProcessCol(col);
			for(var i in this.aCols){
				var col = this.aCols[i];
				if (null != col)
					fProcessCol(col);
			}
		}
		else
		{
			if(!_summaryRight) {
				if(!bNotAddCollapsed && start > 0) {
					prevCol = this._getCol(start - 1);
				}
			}

			for(var i = start; i <= stop; i++){
				var col = this._getCol(i);
				fProcessCol(col);
			}

			if(_summaryRight && !bNotAddCollapsed && prevCol) {
				col = this._getCol(stop + 1);
				if(col.getCollapsed()) {
					this.setCollapsedCol(false, null, col);
				}
			}
		}
	};
	Worksheet.prototype.getColHidden=function(index){
		var col = this._getColNoEmptyWithAll(index);
		return col ? col.getHidden() : false;
	};
	Worksheet.prototype.setColHidden=function(bHidden, start, stop){
		//start, stop 0 based
		if(null == start)
			return;
		if(null == stop)
			stop = start;
		AscCommon.History.Create_NewPoint();
		var oThis = this, outlineLevel;
		var bNotAddCollapsed = true == this.workbook.bUndoChanges || true == this.workbook.bRedoChanges || this.bExcludeCollapsed;
		var _summaryRight = this.sheetPr ? this.sheetPr.SummaryRight : true;
		var fProcessCol = function(col){

			if(col && !bNotAddCollapsed && outlineLevel !== undefined && outlineLevel !== col.getOutlineLevel()) {
				if(!_summaryRight) {
					oThis.setCollapsedCol(bHidden, col.index - 1);
				} else {
					oThis.setCollapsedCol(bHidden, null, col);
				}
			}
			outlineLevel = col ? col.getOutlineLevel() : null;

			if(col.getHidden() != bHidden)
			{
				var oOldProps = col.getWidthProp();
				if(bHidden)
				{
					col.setHidden(bHidden);
					if(null == col.width || true != col.CustomWidth)
						col.width = 0;
					col.CustomWidth = true;
					col.BestFit = null;
				}
				else
				{
					col.setHidden(null);
					if(0 >= col.width)
						col.width = null;
				}
				var oNewProps = col.getWidthProp();
				if(false == oOldProps.isEqual(oNewProps))
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ColProp, oThis.getId(),
								col._getUpdateRange(),
								new UndoRedoData_IndexSimpleProp(col.index, false, oOldProps, oNewProps));
			}
		};

		if(!bNotAddCollapsed && !_summaryRight && start > 0) {
			col = this._getCol(start - 1);
			outlineLevel = col.getOutlineLevel();
		}

		if(0 === start && gc_nMaxCol0 === stop)
		{
			var col = null;
			if(false == bHidden)
				col = this.oAllCol;
			else
				col = this.getAllCol();
			if(null != col)
				fProcessCol(col);
			for(var i in this.aCols){
				var col = this.aCols[i];
				if (null != col)
					fProcessCol(col);
			}
		}
		else
		{
			for(var i = start; i <= stop; i++){
				var col = null;
				if(false == bHidden)
					col = this._getColNoEmpty(i);
				else
					col = this._getCol(i);
				if(null != col)
					fProcessCol(col);
			}
		}

		if(!bNotAddCollapsed && outlineLevel && _summaryRight) {
			col = this._getCol(stop + 1);
			if(col && outlineLevel !== col.getOutlineLevel()) {
				oThis.setCollapsedCol(bHidden, null, col);
			}
		}
	};
	//TODO если collapsed не будет выставляться и заносится, удалить
	Worksheet.prototype.setCollapsedCol = function (bCollapse, colIndex, curCol) {
		var oThis = this;
		var fProcessCol = function(col){
			var oOldProps = col.getCollapsed();
			col.setCollapsed(bCollapse);
			var oNewProps = col.getCollapsed();

			if(oOldProps !== oNewProps) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_CollapsedCol, oThis.getId(), col._getUpdateRange(), new UndoRedoData_IndexSimpleProp(col.index, true, oOldProps, oNewProps));
			}
		};

		if(curCol) {
			fProcessCol(curCol);
		} else {
			this.getRange3(0, colIndex,0, colIndex)._foreachCol(fProcessCol);
		}
	};
	Worksheet.prototype.setSummaryRight = function (val) {
		if (!this.sheetPr){
			this.sheetPr = new AscCommonExcel.asc_CSheetPr();
		}

		AscCommon.History.Create_NewPoint();
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetSummaryRight, this.getId(), null,
			new UndoRedoData_FromTo(this.sheetPr.SummaryRight, val));

		this.sheetPr.SummaryRight = val;
	};
	Worksheet.prototype.setSummaryBelow = function (val) {
		if (!this.sheetPr){
			this.sheetPr = new AscCommonExcel.asc_CSheetPr();
		}

		AscCommon.History.Create_NewPoint();
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetSummaryBelow, this.getId(), null,
			new UndoRedoData_FromTo(this.sheetPr.SummaryBelow, val));

		this.sheetPr.SummaryBelow = val;
	};
	Worksheet.prototype.setFitToPage = function (val) {
		if((this.sheetPr && val !== this.sheetPr.FitToPage) || (!this.sheetPr && val)) {
			if (!this.sheetPr){
				this.sheetPr = new AscCommonExcel.asc_CSheetPr();
			}

			AscCommon.History.Create_NewPoint();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetFitToPage, this.getId(), null,
				new UndoRedoData_FromTo(this.sheetPr.FitToPage, val));

			this.sheetPr.FitToPage = val;
		}
	};
	Worksheet.prototype.setGroupCol = function (bDel, start, stop) {
		var oThis = this;
		var fProcessCol = function(col){
			col.setOutlineLevel(null, bDel);
		};

		this.getRange3(0, start, 0, stop)._foreachCol(fProcessCol);
	};
	Worksheet.prototype.setOutlineCol = function (val, start, stop) {
		var oThis = this;

		var fProcessCol = function(col){
			col.setOutlineLevel(val);
		};

		this.getRange3(0, start, 0, stop)._foreachCol(fProcessCol);
	};
	Worksheet.prototype.getColCustomWidth = function (index) {
		var isBestFit;
		var column = this._getColNoEmptyWithAll(index);
		if (!column) {
			isBestFit = true;
		} else if (column.getHidden()) {
			isBestFit = false;
		} else {
			isBestFit = !!(column.BestFit || (null === column.BestFit && null === column.CustomWidth));
		}
		return !isBestFit;
	};
	Worksheet.prototype.setColBestFit=function(bBestFit, width, start, stop){
		//start, stop 0 based
		if(null == start)
			return;
		if(null == stop)
			stop = start;
		AscCommon.History.Create_NewPoint();
		var oThis = this;
		var fProcessCol = function(col){
			var oOldProps = col.getWidthProp();
			if(bBestFit)
			{
				col.BestFit = bBestFit;
				col.setHidden(null);
			}
			else
				col.BestFit = null;
			col.width = width;
			oThis.initColumn(col);
			var oNewProps = col.getWidthProp();
			if(false == oOldProps.isEqual(oNewProps))
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ColProp, oThis.getId(),
							col._getUpdateRange(),
							new UndoRedoData_IndexSimpleProp(col.index, false, oOldProps, oNewProps));
		};
		if(0 === start && gc_nMaxCol0 === stop)
		{
			var col = null;
			if(bBestFit && oDefaultMetrics.ColWidthChars == width)
				col = this.oAllCol;
			else
				col = this.getAllCol();
			if(null != col)
				fProcessCol(col);
			for(var i in this.aCols){
				var col = this.aCols[i];
				if (null != col)
					fProcessCol(col);
			}
		}
		else
		{
			for(var i = start; i <= stop; i++){
				var col = null;
				if(bBestFit && oDefaultMetrics.ColWidthChars == width)
					col = this._getColNoEmpty(i);
				else
					col = this._getCol(i);
				if(null != col)
					fProcessCol(col);
			}
		}
	};
	Worksheet.prototype.isDefaultHeightHidden=function(){
		return null != this.oSheetFormatPr.oAllRow && this.oSheetFormatPr.oAllRow.getHidden();
	};
	Worksheet.prototype.isDefaultWidthHidden=function(){
		return null != this.oAllCol && this.oAllCol.getHidden();
	};
	Worksheet.prototype.setDefaultHeight = function (h){
		// ToDo refactoring this
		if (this.oSheetFormatPr.oAllRow && !this.oSheetFormatPr.oAllRow.getCustomHeight()) {
			this.oSheetFormatPr.oAllRow.h = h;
		}
	};
	Worksheet.prototype.getDefaultHeight=function(){
		// ToDo http://bugzilla.onlyoffice.com/show_bug.cgi?id=19666 (флага CustomHeight нет)
		var dRes = null;
		// Нужно возвращать выставленную, только если флаг CustomHeight = true
		if(null != this.oSheetFormatPr.oAllRow && this.oSheetFormatPr.oAllRow.getCustomHeight())
			dRes = this.oSheetFormatPr.oAllRow.h;
		return dRes;
	};
	Worksheet.prototype.getRowHeight = function(index) {
		var res;
		this._getRowNoEmptyWithAll(index, function(row){
			res = row ? row.getHeight() : -1;
		});
		return res;
	};
	Worksheet.prototype.setRowHeight=function(height, start, stop, isCustom){
		if(0 == height)
			return this.setRowHidden(true, start, stop);
		//start, stop 0 based
		if(null == start)
			return;
		if(null == stop)
			stop = start;
		AscCommon.History.Create_NewPoint();
		var oThis = this, i;
		/*var oSelection = AscCommon.History.GetSelection();
		if(null != oSelection)
		{
			oSelection = oSelection.clone();
			oSelection.assign(0, start, gc_nMaxCol0, stop);
			AscCommon.History.SetSelection(oSelection);
			AscCommon.History.SetSelectionRedo(oSelection);
		}*/
		var prevRow;
		var bNotAddCollapsed = true == this.workbook.bUndoChanges || true == this.workbook.bRedoChanges || this.bExcludeCollapsed;
		var _summaryBelow = this.sheetPr ? this.sheetPr.SummaryBelow : true;
		var fProcessRow = function(row){
			if(row)
			{
				if(_summaryBelow && !bNotAddCollapsed && row.getCollapsed()) {
					oThis.setCollapsedRow(false, null, row);
				} else if(!_summaryBelow && !bNotAddCollapsed && prevRow && prevRow.getCollapsed()) {
					oThis.setCollapsedRow(false, null, prevRow);
				}
				prevRow = row;

				var oOldProps = row.getHeightProp();
				row.setHeight(height);
				if (isCustom) {
					row.setCustomHeight(true);
				}
				row.setCalcHeight(true);
				row.setHidden(false);
				var oNewProps = row.getHeightProp();
				if(false === oOldProps.isEqual(oNewProps))
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RowProp, oThis.getId(), row._getUpdateRange(), new UndoRedoData_IndexSimpleProp(row.index, true, oOldProps, oNewProps));
			}
		};
		if(0 == start && gc_nMaxRow0 == stop)
		{
			fProcessRow(this.getAllRow());
			this._forEachRow(fProcessRow);
		}
		else
		{
			if(!_summaryBelow) {
				if(!bNotAddCollapsed && start > 0) {
					this._getRow(start - 1, function(row) {
						prevRow = row;
					});
				}
			}

			this.getRange3(start,0,stop, 0)._foreachRow(fProcessRow);

			if(_summaryBelow) {
				if(!bNotAddCollapsed && prevRow) {
					this._getRow(stop + 1, function(row) {
						if(row.getCollapsed()) {
							oThis.setCollapsedRow(false, null, row);
						}
					});
				}
			}
		}
		if(this.needRecalFormulas(start, stop)) {
			this.workbook.dependencyFormulas.calcTree();
		}
	};
	Worksheet.prototype.getRowHidden=function(index){
		var res;
		this._getRowNoEmptyWithAll(index, function(row){
			res = row ? row.getHidden() : false;
		});
		return res;
	};
	Worksheet.prototype.setRowHidden=function(bHidden, start, stop){
		//start, stop 0 based
		if(null == start)
			return;
		if(null == stop)
			stop = start;

		var oThis = this;

		var doHide = function (_start, _stop, localChange) {
			var i;
			var startIndex = null, endIndex = null, updateRange, outlineLevel;
			var bNotAddCollapsed = true == oThis.workbook.bUndoChanges || true == oThis.workbook.bRedoChanges || oThis.bExcludeCollapsed;
			var _summaryBelow = oThis.sheetPr ? oThis.sheetPr.SummaryBelow : true;
			var fProcessRow = function(row){
				if(row && !bNotAddCollapsed && outlineLevel !== undefined && outlineLevel !== row.getOutlineLevel()) {
					if(!_summaryBelow) {
						oThis.setCollapsedRow(bHidden, row.index - 1);
					} else {
						oThis.setCollapsedRow(bHidden, null, row);
					}
				}
				outlineLevel = row ? row.getOutlineLevel() : null;

				if(row && bHidden != row.getHidden())
				{
					row.setHidden(bHidden, localChange);

					if(row.index === endIndex + 1 && startIndex !== null)
						endIndex++;
					else
					{
						if(startIndex !== null)
						{
							updateRange = new Asc.Range(0, startIndex, gc_nMaxCol0, endIndex);
							AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RowHide, oThis.getId(), updateRange, new UndoRedoData_FromToRowCol(bHidden, startIndex, endIndex));
						}

						startIndex = row.index;
						endIndex = row.index;
					}
				}
			};
			if(0 == _start && gc_nMaxRow0 == _stop)
			{
				// ToDo реализовать скрытие всех строк!
			}
			else
			{
				if(!_summaryBelow && _start > 0 && !bNotAddCollapsed) {
					oThis._getRow(_start - 1, function(row) {
						if(row) {
							outlineLevel = row.getOutlineLevel();
						}
					});
				}

				for (i = _start; i <= _stop; ++i) {
					false == bHidden ? oThis._getRowNoEmpty(i, fProcessRow) : oThis._getRow(i, fProcessRow);
				}

				if(_summaryBelow && outlineLevel && !bNotAddCollapsed) {
					oThis._getRow(_stop + 1, function(row) {
						if(row && outlineLevel !== row.getOutlineLevel()) {
							oThis.setCollapsedRow(bHidden, null, row);
						}
					});
				}

				if(startIndex !== null)//заносим последние строки
				{
					updateRange = new Asc.Range(0, startIndex, gc_nMaxCol0, endIndex);
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RowHide, oThis.getId(),updateRange, new UndoRedoData_FromToRowCol(bHidden, startIndex, endIndex));
				}
			}
		};

		var bCollaborativeChanges = !this.autoFilters.useViewLocalChange && this.workbook.bCollaborativeChanges
		if (!bCollaborativeChanges && null !== this.getActiveNamedSheetViewId()) {
			var rowsArr = this.autoFilters.splitRangeByFilters(start, stop);
			if (rowsArr) {
				var j;
				AscCommon.History.Create_NewPoint();
				if (rowsArr[0] && rowsArr[0].length) {
					var oldLocalChange = AscCommon.History.LocalChange;
					AscCommon.History.LocalChange = true;
					for (j = 0; j < rowsArr[0].length; j++) {
						doHide(rowsArr[0][j].start, rowsArr[0][j].stop, true)
					}
					AscCommon.History.LocalChange = oldLocalChange;
				}
				if (rowsArr[1] && rowsArr[1].length) {
					for (j = 0; j < rowsArr[1].length; j++) {
						doHide(rowsArr[1][j].start, rowsArr[1][j].stop)
					}
				}
			}
		} else {
			AscCommon.History.Create_NewPoint();
			doHide(start, stop)
		}

		if(this.needRecalFormulas(start, stop)) {
			this.workbook.dependencyFormulas.calcTree();
		}
	};
	//TODO
	Worksheet.prototype.setCollapsedRow = function (bCollapse, rowIndex, curRow) {
		var oThis = this;
		var fProcessRow = function(row, bSave){
			var oOldProps = row.getCollapsed();
			row.setCollapsed(bCollapse);
			var oNewProps = row.getCollapsed();

			if(oOldProps !== oNewProps) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_CollapsedRow, oThis.getId(), row._getUpdateRange(), new UndoRedoData_IndexSimpleProp(row.index, true, oOldProps, oNewProps));
			}
			if(bSave) {
				row.saveContent(true);
			}
		};

		if(curRow) {
			fProcessRow(curRow, true);
		} else {
			this.getRange3(rowIndex,0,rowIndex, 0)._foreachRow(fProcessRow);
		}
	};
	Worksheet.prototype.setGroupRow = function (bDel, start, stop) {
		var oThis = this;
		var fProcessRow = function(row){
			row.setOutlineLevel(null, bDel);
		};

		this.getRange3(start,0,stop, 0)._foreachRow(fProcessRow);
	};
	Worksheet.prototype.setOutlineRow = function (val, start, stop) {
		var oThis = this;
		var fProcessRow = function(row){
			row.setOutlineLevel(val);
		};

		this.getRange3(start,0,stop, 0)._foreachRow(fProcessRow);
	};
	Worksheet.prototype.getRowCustomHeight = function (index) {
		var isCustomHeight = false;
		this._getRowNoEmptyWithAll(index, function (row) {
			if (!row) {
				isCustomHeight = false;
			} else if (row.getHidden()) {
				isCustomHeight = true;
			} else {
				isCustomHeight = row.getCustomHeight();
			}
		});
		return isCustomHeight;
	};
	Worksheet.prototype.setRowBestFit=function(bBestFit, height, start, stop){
		//start, stop 0 based
		if(null == start)
			return;
		if(null == stop)
			stop = start;
		AscCommon.History.Create_NewPoint();
		var oThis = this, i;
		var isDefaultProp = (true == bBestFit && oDefaultMetrics.RowHeight == height);
		var fProcessRow = function(row){
			if(row)
			{
				var oOldProps = row.getHeightProp();
				row.setCustomHeight(!bBestFit);
				row.setCalcHeight(true);
				row.setHeight(height);
				var oNewProps = row.getHeightProp();
				if(false == oOldProps.isEqual(oNewProps))
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RowProp, oThis.getId(), row._getUpdateRange(), new UndoRedoData_IndexSimpleProp(row.index, true, oOldProps, oNewProps));
			}
		};
		if(0 == start && gc_nMaxRow0 == stop) {
			fProcessRow(isDefaultProp ? this.oSheetFormatPr.oAllRow : this.getAllRow());
			this._forEachRow(fProcessRow);
		} else {
			var range = this.getRange3(start,0,stop, 0);
			if (isDefaultProp) {
				range._foreachRowNoEmpty(fProcessRow);
			} else {
				range._foreachRow(fProcessRow);
			}
		}
		this.workbook.dependencyFormulas.calcTree();
	};
	Worksheet.prototype.getCell=function(oCellAdd){
		return this.getRange(oCellAdd, oCellAdd);
	};
	Worksheet.prototype.getCell2=function(sCellAdd){
		if( sCellAdd.indexOf("$") > -1)
			sCellAdd = sCellAdd.replace(/\$/g,"");
		return this.getRange2(sCellAdd);
	};
	Worksheet.prototype.getCell3=function(r1, c1){
		return this.getRange3(r1, c1, r1, c1);
	};
	Worksheet.prototype.getRange=function(cellAdd1, cellAdd2){
		//Если range находится за границами ячеек расширяем их
		var nRow1 = cellAdd1.getRow0();
		var nCol1 = cellAdd1.getCol0();
		var nRow2 = cellAdd2.getRow0();
		var nCol2 = cellAdd2.getCol0();
		return this.getRange3(nRow1, nCol1, nRow2, nCol2);
	};
	Worksheet.prototype.getRange2=function(sRange){
		var bbox = AscCommonExcel.g_oRangeCache.getAscRange(sRange);
		if(null != bbox)
			return Range.prototype.createFromBBox(this, bbox);
		return null;
	};
	Worksheet.prototype.getRange3=function(r1, c1, r2, c2){
		var nRowMin = r1;
		var nRowMax = r2;
		var nColMin = c1;
		var nColMax = c2;
		if(r1 > r2){
			nRowMax = r1;
			nRowMin = r2;
		}
		if(c1 > c2){
			nColMax = c1;
			nColMin = c2;
		}
		return new Range(this, nRowMin, nColMin, nRowMax, nColMax);
	};
	Worksheet.prototype.getRange4=function(r, c){
		return new Range(this, r, c, r, c);
	};
	Worksheet.prototype.getRowIterator=function(r1, c1, c2, callback){
		var it = new RowIterator();
		it.init(this, r1, c1, c2);
		callback(it);
		it.release();
	};
	Worksheet.prototype.getCellForValidation=function(row, col, array, formula, callback, isCopyPaste, byRef){
		let cell;
		this._getCell(row, col, function (_cell) {
			cell = _cell.clone();
		});
		cell.setValueForValidation(array, formula, callback, isCopyPaste, byRef);
		return cell;
	};
	Worksheet.prototype._removeCell=function(nRow, nCol, cell){
		var t = this;
		var processCell = function(cell) {
			if(null != cell)
			{
				var sheetId = t.getId();
				if (false == cell.isEmpty()) {
					var oUndoRedoData_CellData = new AscCommonExcel.UndoRedoData_CellData(cell.getValueData(), null);
					if (null != cell.xfs)
						oUndoRedoData_CellData.style = cell.xfs.clone();
					if (cell.formulaParsed) {
						cell.checkRemoveExternalReferences(null, cell.formulaParsed);
					}
					cell.setFormulaInternal(null);
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RemoveCell, sheetId, new Asc.Range(nCol, nRow, nCol, nRow), new UndoRedoData_CellSimpleData(nRow, nCol, oUndoRedoData_CellData, null));
				}
				t.workbook.dependencyFormulas.addToChangedCell(cell);

				cell.clearData();
				cell.saveContent(true);
			}
		};
		if(null != cell)
		{
			nRow = cell.nRow;
			nCol = cell.nCol;
			processCell(cell);
		} else {
			this._getCellNoEmpty(nRow, nCol, processCell);
		}
	};
	Worksheet.prototype._getCell=function(row, col, fAction){
		var wb = this.workbook;
		var targetCell = null;
		for (var k = 0; k < wb.loadCells.length; ++k) {
			var elem = wb.loadCells[k];
			if (elem.nRow == row && elem.nCol == col && this === elem.ws) {
				targetCell = elem;
				break;
			}
		}
		if(null === targetCell){
			var cell = new Cell(this);
			wb.loadCells.push(cell);
			if (!cell.loadContent(row, col)) {
				this._initCell(cell, row, col);
			}
			fAction(cell);
			cell.saveContent(true);
			wb.loadCells.pop();
		} else {
			fAction(targetCell);
		}
	};
	Worksheet.prototype._initRow=function(row, index){
		var t = this;
		row.setChanged(true);
		if (null != this.oSheetFormatPr.oAllRow) {
			row.copyFrom(this.oSheetFormatPr.oAllRow);
			row.setIndex(index);
		}
		this.nRowsCount = index >= this.nRowsCount ? index + 1 : this.nRowsCount;
	};
	Worksheet.prototype._initCell=function(cell, nRow, nCol){
		var t = this;
		cell.setChanged(true);
		this._getRowNoEmpty(nRow, function(row) {
			var oCol = t._getColNoEmptyWithAll(nCol);
			var xfs = null;
			if (row && null != row.xfs)
				xfs = row.xfs.clone();
			else if (null != oCol && null != oCol.xfs)
				xfs = oCol.xfs.clone();
			cell.setStyleInternal(xfs);
			t.cellsByColRowsCount = Math.max(t.cellsByColRowsCount, nRow + 1);
			t.nRowsCount = Math.max(t.nRowsCount, t.cellsByColRowsCount);
			if (nCol >= t.nColsCount)
				t.setColsCount(nCol + 1);
		});
		//init ColData otherwise all 'foreach' will not return this cell until saveContent(loadCells)
		var sheetMemory = this.getColData(nCol);
		sheetMemory.checkIndex(nRow);
	};
	Worksheet.prototype._getCellNoEmpty=function(row, col, fAction){
		var wb = this.workbook;
		var targetCell = null;
		for (var k = 0; k < wb.loadCells.length; ++k) {
			var elem = wb.loadCells[k];
			if (elem.nRow == row && elem.nCol == col && this === elem.ws) {
				targetCell = elem;
				break;
			}
		}
		if (null === targetCell) {
			var cell = new Cell(this);
			var res = cell.loadContent(row, col) ? cell : null;
			if (res && fAction) {
				wb.loadCells.push(cell);
			}
			fAction(res);
			cell.saveContent(true);
			if (res) {
				wb.loadCells.pop();
			}
		} else {
			fAction(targetCell);
		}
	};
	Worksheet.prototype._getRowNoEmpty=function(nRow, fAction){
		//0-based
		var row = new AscCommonExcel.Row(this);
		if(row.loadContent(nRow)){
			fAction(row);
			row.saveContent(true);
		} else {
			fAction(null);
		}
	};
	Worksheet.prototype._getRowNoEmptyWithAll=function(nRow, fAction){
		var t = this;
		this._getRowNoEmpty(nRow, function(row){
			if(!row)
				row = t.oSheetFormatPr.oAllRow;
			fAction(row);
		});

	};
	Worksheet.prototype._getColNoEmpty = function (col) {
		//0-based
		return this.aCols[col] || null;
	};
	Worksheet.prototype._getColNoEmptyWithAll = function (col) {
		return this._getColNoEmpty(col) || this.oAllCol;
	};
	Worksheet.prototype._getRow = function(index, fAction) {
		//0-based
		var row = null;
		if (g_nAllRowIndex == index)
			row = this.getAllRow();
		else {
			row = new AscCommonExcel.Row(this);
			if (!row.loadContent(index)) {
				this._initRow(row, index);
			}
		}
		fAction(row);
		row.saveContent(true);
	};
	Worksheet.prototype._getCol=function(index){
		//0-based
		var oCurCol;
		if (g_nAllColIndex == index)
			oCurCol = this.getAllCol();
		else
		{
			oCurCol = this.aCols[index];
			if(null == oCurCol)
			{
				if(null != this.oAllCol)
				{
					oCurCol = this.oAllCol.clone();
					oCurCol.index = index;
				}
				else
					oCurCol = new AscCommonExcel.Col(this, index);
				this.aCols[index] = oCurCol;
				this.setColsCount(index >= this.nColsCount ? index + 1 : this.nColsCount)
			}
		}
		return oCurCol;
	};
	Worksheet.prototype._prepareMoveRangeGetCleanRanges=function(oBBoxFrom, oBBoxTo, wsTo){
		var intersection = oBBoxFrom.intersectionSimple(oBBoxTo);
		var aRangesToCheck = [];
		if(null != intersection && this === wsTo)
		{
			var oThis = this;
			var fAddToRangesToCheck = function(aRangesToCheck, r1, c1, r2, c2)
			{
				if(r1 <= r2 && c1 <= c2)
					aRangesToCheck.push(oThis.getRange3(r1, c1, r2, c2));
			};
			if(intersection.r1 == oBBoxTo.r1 && intersection.c1 == oBBoxTo.c1)
			{
				fAddToRangesToCheck(aRangesToCheck, oBBoxTo.r1, intersection.c2 + 1, intersection.r2, oBBoxTo.c2);
				fAddToRangesToCheck(aRangesToCheck, intersection.r2 + 1, oBBoxTo.c1, oBBoxTo.r2, oBBoxTo.c2);
			}
			else if(intersection.r2 == oBBoxTo.r2 && intersection.c1 == oBBoxTo.c1)
			{
				fAddToRangesToCheck(aRangesToCheck, oBBoxTo.r1, oBBoxTo.c1, intersection.r1 - 1, oBBoxTo.c2);
				fAddToRangesToCheck(aRangesToCheck, intersection.r1, intersection.c2 + 1, oBBoxTo.r2, oBBoxTo.c2);
			}
			else if(intersection.r1 == oBBoxTo.r1 && intersection.c2 == oBBoxTo.c2)
			{
				fAddToRangesToCheck(aRangesToCheck, oBBoxTo.r1, oBBoxTo.c1, intersection.r2, intersection.c1 - 1);
				fAddToRangesToCheck(aRangesToCheck, intersection.r2 + 1, oBBoxTo.c1, oBBoxTo.r2, oBBoxTo.c2);
			}
			else if(intersection.r2 == oBBoxTo.r2 && intersection.c2 == oBBoxTo.c2)
			{
				fAddToRangesToCheck(aRangesToCheck, oBBoxTo.r1, oBBoxTo.c1, intersection.r1 - 1, oBBoxTo.c2);
				fAddToRangesToCheck(aRangesToCheck, intersection.r1, oBBoxTo.c1, oBBoxTo.r2, intersection.c1 - 1);
			}
		}
		else
			aRangesToCheck.push(wsTo.getRange3(oBBoxTo.r1, oBBoxTo.c1, oBBoxTo.r2, oBBoxTo.c2));
		return aRangesToCheck;
	};
	Worksheet.prototype._prepareMoveRange=function(oBBoxFrom, oBBoxTo, wsTo){
		var res = 0;
		if (!wsTo) {
			wsTo = this;
		}
		if (oBBoxFrom.isEqual(oBBoxTo) && this === wsTo)
			return res;
		var range = wsTo.getRange3(oBBoxTo.r1, oBBoxTo.c1, oBBoxTo.r2, oBBoxTo.c2);
		var aMerged = wsTo.mergeManager.get(range.getBBox0());
		if(aMerged.outer.length > 0)
			return -2;
		var aRangesToCheck = this._prepareMoveRangeGetCleanRanges(oBBoxFrom, oBBoxTo, wsTo);
		for(var i = 0, length = aRangesToCheck.length; i < length; i++)
		{
			range = aRangesToCheck[i];
			range._foreachNoEmpty(
				function(cell){
					if(!cell.isNullTextString())
					{
						res = -1;
						return res;
					}
				});
			if(0 != res)
				return res;
		}
		return res;
	};
	Worksheet.prototype._movePivots = function (oBBoxFrom, oBBoxTo, copyRange, offset, wsTo) {
		if (!wsTo) {
			wsTo = this;
		}
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			if (copyRange) {
				wsTo.deletePivotTables(oBBoxTo);
				this.copyPivotTable(oBBoxFrom, offset, wsTo);
			} else {
				if (this === wsTo) {
					this.deletePivotTablesOnMove(oBBoxFrom, oBBoxTo);
					this.movePivotOffset(oBBoxFrom, offset);
				} else {
					this.copyPivotTable(oBBoxFrom, offset, wsTo);
					this._movePivotSlicerWsRefs(oBBoxTo, this, wsTo);
					this.deletePivotTables(oBBoxFrom, true);
				}
			}
		}
	};
	Worksheet.prototype._movePivotSlicerWsRefs = function (range, wsFrom, wsTo) {
		var wb = this.workbook;
		var pivotTable;
		for (var i = 0; i < wsTo.pivotTables.length; ++i) {
			pivotTable = wsTo.pivotTables[i];
			if (pivotTable.isInRange(range)) {
				var slicerCaches = wb.getSlicerCachesByPivotTable(wsFrom.getId(), pivotTable.asc_getName());
				slicerCaches.forEach(function (slicerCache) {
					slicerCache.movePivotTable(wsFrom.getId(), pivotTable.asc_getName(), wsTo.getId(), pivotTable.asc_getName());
				});
			}
		}
	};
	Worksheet.prototype._moveMergedAndHyperlinksPrepare = function(oBBoxFrom, oBBoxTo, copyRange, wsTo, offset) {
		var res = {merged: [], hyperlinks: []};
		if (!(false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges))) {
			return res;
		}
		var i, elem, bbox, data, wsFrom = this;
		var intersection = oBBoxFrom.intersectionSimple(oBBoxTo);
		AscCommon.History.LocalChange = true;
		//merged
		var merged = wsFrom.mergeManager.get(oBBoxFrom).inner;
		var mergedToRemove;
		if (!copyRange) {
			mergedToRemove = merged;
		} else if (null !== intersection) {
			mergedToRemove = wsFrom.mergeManager.get(intersection).all;
		}
		if(mergedToRemove){
			for (i = 0; i < mergedToRemove.length; i++) {
				wsFrom.mergeManager.removeElement(mergedToRemove[i]);
			}
		}

		//hyperlinks
		var hyperlinks = wsFrom.hyperlinkManager.get(oBBoxFrom).inner;
		if (!copyRange) {
			for (i = 0; i < hyperlinks.length; i++) {
				wsFrom.hyperlinkManager.removeElement(hyperlinks[i]);
			}
		}
		AscCommon.History.LocalChange = false;
		res.merged = merged;
		res.hyperlinks = hyperlinks;
		return res;
	};
	Worksheet.prototype._moveMergedAndHyperlinks = function(prepared, oBBoxFrom, oBBoxTo, copyRange, wsTo, offset) {
		var i, elem, bbox, data;
		var intersection = oBBoxFrom.intersectionSimple(oBBoxTo);
		AscCommon.History.LocalChange = true;
		for (i = 0; i < prepared.merged.length; i++) {
			elem = prepared.merged[i];
			bbox = copyRange ? elem.bbox.clone() : elem.bbox;
			bbox.setOffset(offset);
			wsTo.mergeManager.add(bbox, elem.data);
		}
		//todo сделать для пересечения
		if (!copyRange || null === intersection) {
			for (i = 0; i < prepared.hyperlinks.length; i++) {
				elem = prepared.hyperlinks[i];
				if (copyRange) {
					bbox = elem.bbox.clone();
					data = elem.data.clone();
				}
				else {
					bbox = elem.bbox;
					data = elem.data;
				}
				bbox.setOffset(offset);
				wsTo.hyperlinkManager.add(bbox, data);
			}
		}
		AscCommon.History.LocalChange = false;
	};
	Worksheet.prototype._moveCleanRanges = function(oBBoxFrom, oBBoxTo, copyRange, wsTo) {
		//удаляем to через историю, для undo
		var cleanRanges = this._prepareMoveRangeGetCleanRanges(oBBoxFrom, oBBoxTo, wsTo);
		for (var i = 0; i < cleanRanges.length; i++) {
			var range = cleanRanges[i];
			range.cleanAll();
			//выставляем для slave refError
			if (!copyRange)
				this.workbook.dependencyFormulas.deleteNodes(wsTo.getId(), range.getBBox0());
		}
	};
	Worksheet.prototype._moveFormulas = function(oBBoxFrom, oBBoxTo, copyRange, wsTo, offset, isTableCreated) {
		if(!copyRange){
			this.workbook.dependencyFormulas.move(this.Id, oBBoxFrom, offset, wsTo.getId(), isTableCreated);
		}
		//todo avoid double getRange3
		this.getRange3(oBBoxFrom.r1, oBBoxFrom.c1, oBBoxFrom.r2, oBBoxFrom.c2)._foreachNoEmpty(function(cell) {
			if (cell.transformSharedFormula()) {
				var parsed = cell.getFormulaParsed();
				parsed.buildDependencies();
			}
		});
	};
	Worksheet.prototype._moveCells = function(oBBoxFrom, oBBoxTo, copyRange, wsTo, offset) {
		var oThis = this;
		var nRowsCountNew = 0;
		var nColsCountNew = 0;
		var dependencyFormulas = oThis.workbook.dependencyFormulas;
		var moveToOtherSheet = this !== wsTo;
		var isClearFromArea = !copyRange || (copyRange && oThis.workbook.bUndoChanges);
		var moveCells = function(copyRange, from, to, r1From, r1To, count){
			var fromData = oThis.getColDataNoEmpty(from);
			var toData;
			if(fromData){
				toData = wsTo.getColData(to);
				toData.copyRange(fromData, r1From, r1To, count);
				if (isClearFromArea) {
					if(from !== to || moveToOtherSheet) {
						fromData.clear(r1From, r1From + count);
					} else {
						if (r1From < r1To) {
							fromData.clear(r1From, Math.min(r1From + count, r1To));
						} else {
							fromData.clear(Math.max(r1From, r1To + count), r1From + count);
						}
					}
				}
			} else {
				toData = wsTo.getColDataNoEmpty(to);
				if(toData) {
					toData.clear(r1To, r1To + count);
				}
			}
			if (toData) {
				nRowsCountNew = Math.max(nRowsCountNew, toData.getMaxIndex() + 1);
				nColsCountNew = Math.max(nColsCountNew, to + 1);
			}
		};
		if(oBBoxFrom.c1 < oBBoxTo.c1){
			for(var i = 0 ; i < oBBoxFrom.c2 - oBBoxFrom.c1 + 1; ++i){
				moveCells(copyRange, oBBoxFrom. c2 - i, oBBoxTo.c2 - i, oBBoxFrom.r1, oBBoxTo.r1, oBBoxFrom.r2 - oBBoxFrom.r1 + 1);
			}
		} else {
			for(var i = 0 ; i < oBBoxFrom.c2 - oBBoxFrom.c1 + 1; ++i){
				moveCells(copyRange, oBBoxFrom.c1 + i, oBBoxTo.c1 + i, oBBoxFrom.r1, oBBoxTo.r1, oBBoxFrom.r2 - oBBoxFrom.r1 + 1);
			}
		}
		// ToDo возможно нужно уменьшить диапазон обновления
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_MoveRange,
			this.getId(), new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0),
			new UndoRedoData_FromTo(new UndoRedoData_BBox(oBBoxFrom), new UndoRedoData_BBox(oBBoxTo), copyRange, wsTo.getId()));
		if(moveToOtherSheet) {
			//сделано для того, чтобы происходил пересчет/обновление данных на другом листе
			//таким образом заносим диапазон обновления в UpdateRigions
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_Null, wsTo.getId(), oBBoxTo, new UndoRedoData_FromTo(null, null));
		}

		var shiftedArrayFormula = {};
		var oldNewArrayFormulaMap = {};
		//modify nRowsCount/nColsCount for correct foreach functions
		wsTo.cellsByColRowsCount = Math.max(wsTo.cellsByColRowsCount, nRowsCountNew);
		wsTo.nRowsCount = Math.max(wsTo.nRowsCount, wsTo.cellsByColRowsCount);
		wsTo.setColsCount(Math.max(wsTo.nColsCount, nColsCountNew));
		wsTo.getRange3(oBBoxTo.r1, oBBoxTo.c1, oBBoxTo.r2, oBBoxTo.c2)._foreachNoEmpty(function(cell){
			var formula = cell.getFormulaParsed();
			if (formula) {
				var cellWithFormula = formula.getParent();
				var arrayFormula = formula.getArrayFormulaRef();
				var newArrayRef, newFormula;
				var preMoveCell = {nRow: cell.nRow - offset.row, nCol: cell.nCol - offset.col};
				var isFirstCellArray = formula.checkFirstCellArray(preMoveCell) && !shiftedArrayFormula[formula.getListenerId()];
				if (copyRange) {
					AscCommon.History.TurnOff();
					//***array-formula***
					if(!arrayFormula || (arrayFormula && isFirstCellArray)) {
						newFormula = oThis._moveCellsFormula(cell, formula, cellWithFormula, copyRange, oBBoxFrom, wsTo);
						cellWithFormula = newFormula.getParent();
						cellWithFormula = new CCellWithFormula(wsTo, cell.nRow, cell.nCol);
						newFormula = newFormula.clone(null, cellWithFormula, wsTo);
						newFormula.changeOffset(offset, false, true);
						newFormula.setFormulaString(newFormula.assemble(true));
						cell.setFormulaInternal(newFormula, !isClearFromArea);

						if(isFirstCellArray) {
							newArrayRef = arrayFormula.clone();
							newArrayRef.setOffset(offset);
							newFormula.setArrayFormulaRef(newArrayRef);
							shiftedArrayFormula[newFormula.getListenerId()] = 1;
							oldNewArrayFormulaMap[formula.getListenerId()] = newFormula;
						}
					} else if(arrayFormula && oldNewArrayFormulaMap[formula.getListenerId()]) {
						cell.setFormulaInternal(oldNewArrayFormulaMap[formula.getListenerId()], !isClearFromArea);
					}
					AscCommon.History.TurnOn();
				} else {
					//***array-formula***
					//TODO возможно стоит это делать в dependencyFormulas.move
					if(arrayFormula) {
						if(isFirstCellArray) {
							newFormula = oThis._moveCellsFormula(cell, formula, cellWithFormula, copyRange, oBBoxFrom, wsTo);
							cellWithFormula = newFormula.getParent();
							shiftedArrayFormula[formula.getListenerId()] = 1;
							newArrayRef = arrayFormula.clone();
							newArrayRef.setOffset(offset);
							newFormula.setArrayFormulaRef(newArrayRef);

							cellWithFormula.ws = wsTo;
							cellWithFormula.nRow = cell.nRow;
							cellWithFormula.nCol = cell.nCol;
						}
					} else {
						newFormula = oThis._moveCellsFormula(cell, formula, cellWithFormula, copyRange, oBBoxFrom, wsTo);
						cellWithFormula = newFormula.getParent();
						cellWithFormula.ws = wsTo;
						cellWithFormula.nRow = cell.nRow;
						cellWithFormula.nCol = cell.nCol;
					}
				}
				if(arrayFormula) {
					if(isFirstCellArray) {
						dependencyFormulas.addToBuildDependencyArray(formula);
						if(newFormula) {
							dependencyFormulas.addToBuildDependencyArray(newFormula);
						}
					}
				} else {
					dependencyFormulas.addToBuildDependencyCell(cell);
				}
			}
		});
	};
	Worksheet.prototype._moveCellsFormula = function(cell, formula, cellWithFormula, copyRange, oBBoxFrom, wsTo) {
		if (this !== wsTo) {
			if (copyRange || !this.workbook.bUndoChanges) {
				cellWithFormula = new CCellWithFormula(wsTo, cell.nRow, cell.nCol);
				formula = formula.clone(null, cellWithFormula, wsTo);
				if (!copyRange) {
					formula.convertTo3DRefs(oBBoxFrom);
				}
				formula.moveToSheet(this, wsTo);
				formula.setFormulaString(formula.assemble(true));
				cell.setFormulaParsed(formula);
			} else {
				formula.moveToSheet(this, wsTo);
				formula.setFormulaString(formula.assemble(true));
			}
		}
		return formula;
	};
	Worksheet.prototype._moveRange=function(oBBoxFrom, oBBoxTo, copyRange, wsTo, isTableCreated){
		if (!wsTo) {
			wsTo = this;
		}
		if (oBBoxFrom.isEqual(oBBoxTo) && this === wsTo)
			return;

		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		this.workbook.dependencyFormulas.lockRecal();

		var offset = new AscCommon.CellBase(oBBoxTo.r1 - oBBoxFrom.r1, oBBoxTo.c1 - oBBoxFrom.c1);
		var prepared = this._moveMergedAndHyperlinksPrepare(oBBoxFrom, oBBoxTo, copyRange, wsTo, offset);
		this._movePivots(oBBoxFrom, oBBoxTo, copyRange, offset, wsTo);
		this._moveCleanRanges(oBBoxFrom, oBBoxTo, copyRange, wsTo);
		this._moveFormulas(oBBoxFrom, oBBoxTo, copyRange, wsTo, offset, isTableCreated);
		this._moveCells(oBBoxFrom, oBBoxTo, copyRange, wsTo, offset, isTableCreated);
		this._moveMergedAndHyperlinks(prepared, oBBoxFrom, oBBoxTo, copyRange, wsTo, offset);
		this._moveDataValidation(oBBoxFrom, oBBoxTo, copyRange, offset, wsTo);
		this.moveConditionalFormatting(oBBoxFrom, oBBoxTo, copyRange, offset, wsTo);
		this.moveSparklineGroup(oBBoxFrom, oBBoxTo, copyRange, offset, wsTo);
		this.moveProtectedRange(oBBoxFrom, oBBoxTo, copyRange, offset, wsTo);
		this._moveUserProtectedRange(oBBoxFrom, oBBoxTo, copyRange, offset, wsTo);


		if(true == this.workbook.bUndoChanges || true == this.workbook.bRedoChanges) {
			wsTo.autoFilters.unmergeTablesAfterMove(oBBoxTo);
		}

		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.autoFilters._moveAutoFilters(oBBoxTo, oBBoxFrom, null, copyRange, true, oBBoxFrom, wsTo);
		}

		this.workbook.dependencyFormulas.unlockRecal();
		AscCommon.History.EndTransaction();
		return true;
	};
	Worksheet.prototype._shiftCellsLeft=function(oBBox){
		//todo удаление когда есть замерженые ячейки
		var t = this;
		var nLeft = oBBox.c1;
		var nRight = oBBox.c2;
		var dif = nLeft - nRight - 1;
		var oActualRange = new Asc.Range(nLeft, oBBox.r1, gc_nMaxCol0, oBBox.r2);
		var offset = new AscCommon.CellBase(0, dif);
		//renameDependencyNodes before move cells to store current location in history
		var renameRes = this.renameDependencyNodes(offset, oBBox);
		var redrawTablesArr = this.autoFilters.insertColumn( oBBox, dif );
		if (false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			this.updateSortStateOffset(oBBox, offset);
			this.updateSparklineGroupOffset(oBBox, offset);
			this.updateConditionalFormattingOffset(oBBox, offset);
			this.updateProtectedRangeOffset(oBBox, offset);
			AscCommon.History.LocalChange = false;
		}
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.updatePivotOffset(oBBox, offset);
			this.updateUserProtectedRangesOffset(oBBox, offset);
		}

		this.getRange3(oBBox.r1, oBBox.c1, oBBox.r2, oBBox.c2)._foreachNoEmpty(function(cell){
			t._removeCell(null, null, cell);
		});

		this._updateFormulasParents(oActualRange.r1, oActualRange.c1, oActualRange.r2, oActualRange.c2, oBBox, offset, renameRes.shiftedShared);
		var cellsByColLength = this.getColDataLength();
		for (var i = nRight + 1; i < cellsByColLength; ++i) {
			var sheetMemoryFrom = this.getColDataNoEmpty(i);
			if (sheetMemoryFrom) {
				this.getColData(i + dif).copyRange(sheetMemoryFrom, oBBox.r1, oBBox.r1, oBBox.r2 - oBBox.r1 + 1);
				sheetMemoryFrom.clear(oBBox.r1, oBBox.r2 + 1);
			}
		}
		//notifyChanged after move cells to get new locations(for intersect ranges)
		this.workbook.dependencyFormulas.notifyChanged(renameRes.changed);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ShiftCellsLeft, this.getId(), oActualRange, new UndoRedoData_BBox(oBBox));

		this.autoFilters.redrawStylesTables(redrawTablesArr);
		if (this.workbook.handlers) {
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this, null, this.getId());
		}
		//todo проверить не уменьшились ли границы таблицы
	};
	Worksheet.prototype._shiftCellsUp=function(oBBox){
		var t = this;
		var nTop = oBBox.r1;
		var nBottom = oBBox.r2;
		var dif = nTop - nBottom - 1;
		var oActualRange = new Asc.Range(oBBox.c1, oBBox.r1, oBBox.c2, gc_nMaxRow0);
		var offset = new AscCommon.CellBase(dif, 0);
		//renameDependencyNodes before move cells to store current location in history
		var renameRes = this.renameDependencyNodes(offset, oBBox);
		var redrawTablesArr = this.autoFilters.insertRows("delCell", oBBox, c_oAscDeleteOptions.DeleteCellsAndShiftTop);
		if (false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			this.updateSortStateOffset(oBBox, offset);
			this.updateSparklineGroupOffset(oBBox, offset);
			this.updateConditionalFormattingOffset(oBBox, offset);
			this.updateProtectedRangeOffset(oBBox, offset);
			AscCommon.History.LocalChange = false;
		}
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.updatePivotOffset(oBBox, offset);
			this.updateUserProtectedRangesOffset(oBBox, offset);
		}

		this.getRange3(oBBox.r1, oBBox.c1, oBBox.r2, oBBox.c2)._foreachNoEmpty(function(cell){
			t._removeCell(null, null, cell);
		});
		this._updateFormulasParents(oActualRange.r1, oActualRange.c1, oActualRange.r2, oActualRange.c2, oBBox, offset, renameRes.shiftedShared);
		for (var i = oBBox.c1; i <= oBBox.c2; ++i) {
			var sheetMemory = this.getColDataNoEmpty(i);
			if (sheetMemory) {
				sheetMemory.deleteRange(nTop, -dif);
			}
		}
		//notifyChanged after move cells to get new locations(for intersect ranges)
		this.workbook.dependencyFormulas.notifyChanged(renameRes.changed);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ShiftCellsTop, this.getId(), oActualRange, new UndoRedoData_BBox(oBBox));

		this.autoFilters.redrawStylesTables(redrawTablesArr);
		if (this.workbook.handlers) {
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this, null, this.getId());
		}
		//todo проверить не уменьшились ли границы таблицы
	};
	Worksheet.prototype._shiftCellsRight=function(oBBox, displayNameFormatTable){
		var nLeft = oBBox.c1;
		var nRight = oBBox.c2;
		var dif = nRight - nLeft + 1;
		var oActualRange = new Asc.Range(oBBox.c1, oBBox.r1, gc_nMaxCol0, oBBox.r2);
		var offset = new AscCommon.CellBase(0, dif);
		//renameDependencyNodes before move cells to store current location in history
		var renameRes = this.renameDependencyNodes(offset, oBBox);
		var redrawTablesArr = this.autoFilters.insertColumn( oBBox, dif, displayNameFormatTable );
		if (false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			this.updateSortStateOffset(oBBox, offset);
			this.updateSparklineGroupOffset(oBBox, offset);
			this.updateConditionalFormattingOffset(oBBox, offset);
			this.updateProtectedRangeOffset(oBBox, offset);
			AscCommon.History.LocalChange = false;
		}
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.updatePivotOffset(oBBox, offset);
			this.updateUserProtectedRangesOffset(oBBox, offset);
		}

		this._updateFormulasParents(oActualRange.r1, oActualRange.c1, oActualRange.r2, oActualRange.c2, oBBox, offset, renameRes.shiftedShared);
		var borders;
		if (nLeft > 0 && !this.workbook.bUndoChanges) {
			borders = this._getBordersForInsert(oBBox, false);
		}
		var cellsByColLength = this.getColDataLength();
		for (var i = cellsByColLength - 1; i >= nLeft; --i) {
			var sheetMemoryFrom = this.getColDataNoEmpty(i);
			if (sheetMemoryFrom) {
				if (i + dif <= gc_nMaxCol0) {
					this.getColData(i + dif).copyRange(sheetMemoryFrom, oBBox.r1, oBBox.r1, oBBox.r2 - oBBox.r1 + 1);
				}
				sheetMemoryFrom.clear(oBBox.r1, oBBox.r2 + 1);
			}
		}
		this.setColsCount(Math.max(this.nColsCount, this.getColDataLength()));
		//copy property from row/cell above
		if (nLeft > 0 && !this.workbook.bUndoChanges)
		{
			var prevSheetMemory = this.getColDataNoEmpty(nLeft - 1);
			if (prevSheetMemory) {
				//todo hidden, keep only style
				for (var i = nLeft; i <= nRight; ++i) {
					this.getColData(i).copyRange(prevSheetMemory, oBBox.r1, oBBox.r1, oBBox.r2 - oBBox.r1 + 1);
				}
				this.setColsCount(Math.max(this.nColsCount, this.getColDataLength()));
				//show rows and remain only cell xf property
				this.getRange3(oBBox.r1, oBBox.c1, oBBox.r2, oBBox.c2)._foreachNoEmpty(function(cell) {
					cell.clearDataKeepXf(borders[cell.nRow]);
				});
			}
		}
		//notifyChanged after move cells to get new locations(for intersect ranges)
		this.workbook.dependencyFormulas.notifyChanged(renameRes.changed);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ShiftCellsRight, this.getId(), oActualRange, new UndoRedoData_BBox(oBBox));


		this.autoFilters.redrawStylesTables(redrawTablesArr);
		if (this.workbook.handlers) {
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this, null, this.getId());
		}
	};
	Worksheet.prototype._shiftCellsBottom=function(oBBox, displayNameFormatTable){
		var t = this;
		var nTop = oBBox.r1;
		var nBottom = oBBox.r2;
		var dif = nBottom - nTop + 1;
		var oActualRange = new Asc.Range(oBBox.c1, oBBox.r1, oBBox.c2, gc_nMaxRow0);
		var offset = new AscCommon.CellBase(dif, 0);
		//renameDependencyNodes before move cells to store current location in history
		var renameRes = this.renameDependencyNodes(offset, oBBox);
		var redrawTablesArr;
		if (!this.workbook.bUndoChanges && undefined === displayNameFormatTable) {
			redrawTablesArr = this.autoFilters.insertRows("insCell", oBBox, c_oAscInsertOptions.InsertCellsAndShiftDown,
				displayNameFormatTable);
		}
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			this.updatePivotOffset(oBBox, offset);
			this.updateUserProtectedRangesOffset(oBBox, offset);
		}

		this._updateFormulasParents(oActualRange.r1, oActualRange.c1, oActualRange.r2, oActualRange.c2, oBBox, offset, renameRes.shiftedShared);
		if (false == this.workbook.bUndoChanges && (false == this.workbook.bRedoChanges || this.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			this.updateSortStateOffset(oBBox, offset);
			this.updateSparklineGroupOffset(oBBox, offset);
			this.updateConditionalFormattingOffset(oBBox, offset);
			this.updateProtectedRangeOffset(oBBox, offset);
			AscCommon.History.LocalChange = false;
		}

		var borders;
		if (nTop > 0 && !this.workbook.bUndoChanges) {
			borders = this._getBordersForInsert(oBBox, true);
		}
		//rowcount
		for (var i = oBBox.c1; i <= oBBox.c2; ++i) {
			var sheetMemory = this.getColDataNoEmpty(i);
			if (sheetMemory) {
				sheetMemory.insertRange(nTop, dif);
				t.cellsByColRowsCount = Math.max(t.cellsByColRowsCount, sheetMemory.getMaxIndex() + 1);
			}
		}
		this.nRowsCount = Math.max(this.nRowsCount, this.cellsByColRowsCount);
		if (nTop > 0 && !this.workbook.bUndoChanges)
		{
			for (var i = oBBox.c1; i <= oBBox.c2; ++i) {
				var sheetMemory = this.getColDataNoEmpty(i);
				if (sheetMemory) {
					sheetMemory.copyRangeByChunk((nTop - 1), 1, nTop, dif);
					t.cellsByColRowsCount = Math.max(t.cellsByColRowsCount, sheetMemory.getMaxIndex() + 1);
				}
			}
			this.nRowsCount = Math.max(this.nRowsCount, this.cellsByColRowsCount);
			//show rows and remain only cell xf property
			this.getRange3(oBBox.r1, oBBox.c1, oBBox.r2, oBBox.c2)._foreachNoEmpty(function(cell) {
				cell.clearDataKeepXf(borders[cell.nCol]);
			});
		}
		//notifyChanged after move cells to get new locations(for intersect ranges)
		this.workbook.dependencyFormulas.notifyChanged(renameRes.changed);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ShiftCellsBottom, this.getId(), oActualRange, new UndoRedoData_BBox(oBBox));

		//пока перенес добавление только последней строки(в данном случае порядок занесения в истрию должен быть именно в таком порядке)
		//TODO возможно стоит полностью перенести сюда обработку для ф/т и а/ф
		if (!this.workbook.bUndoChanges && undefined !== displayNameFormatTable) {
			redrawTablesArr = this.autoFilters.insertRows("insCell", oBBox, c_oAscInsertOptions.InsertCellsAndShiftDown,
				displayNameFormatTable);
		}

		if(!this.workbook.bUndoChanges)
		{
			this.autoFilters.redrawStylesTables(redrawTablesArr);
			if (this.workbook.handlers) {
				this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this, null, this.getId());
			}
		}
	};
	Worksheet.prototype._setIndex=function(ind){
		if (this.workbook.handlers) {
			this.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetChangeIndex, this.index, ind, this.getId());
		}
		this.index = ind;
	};
	Worksheet.prototype._BuildDependencies=function(cellRange){
		/*
		 Построение графа зависимостей.
		 */
		var ca;
		for (var i in cellRange) {
			if (null === cellRange[i]) {
				cellRange[i] = i;
				continue;
			}

			ca = g_oCellAddressUtils.getCellAddress(i);
			this._getCellNoEmpty(ca.getRow0(), ca.getCol0(), function(c) {
				if (c) {
					c._BuildDependencies(true);
				}
			});
		}
	};
	Worksheet.prototype._setHandlersTablePart = function(){
		if(!this.TableParts)
			return;

		for(var i = 0; i < this.TableParts.length; i++)
		{
			this.TableParts[i].setHandlers(this.handlers);
		}
	};
	Worksheet.prototype.getTableRangeForFormula = function(name, objectParam){
		var res = null;
		if(!this.TableParts || !name)
			return res;

		let _name = name.toLowerCase();
		for(var i = 0; i < this.TableParts.length; i++)
		{
			if(this.TableParts[i].DisplayName.length === _name.length && this.TableParts[i].DisplayName.toLowerCase() === _name)
			{
				return this.TableParts[i].getTableRangeForFormula(objectParam);
			}
		}
		return res;
	};
	Worksheet.prototype.getTableIndexColumnByName = function(tableName, columnName){
		var res = null;
		if(!this.TableParts || !tableName)
			return res;

		let _tableName = tableName.toLowerCase();
		for(var i = 0; i < this.TableParts.length; i++)
		{
			if(this.TableParts[i].DisplayName.length === tableName.length && this.TableParts[i].DisplayName.toLowerCase() === _tableName)
			{
				/* if there is a quote before a special character (i.e. escaping was used), remove it, while the quote itself is also a special character */
				columnName = parserHelp.escapeTableCharacters(columnName, false);

				return this.TableParts[i].getTableIndexColumnByName(columnName);
			}
		}
		return res;
	};
	Worksheet.prototype.getTableNameColumnByIndex = function(tableName, columnIndex){
		var res = null;
		if(!this.TableParts || !tableName)
			return res;

		let _tableName = tableName.toLowerCase();
		for(var i = 0; i < this.TableParts.length; i++)
		{
			if(this.TableParts[i].DisplayName.length === _tableName.length && this.TableParts[i].DisplayName.toLowerCase() === _tableName)
			{
				res = this.TableParts[i].getTableNameColumnByIndex(columnIndex);
				break;
			}
		}
		return res;
	};
	Worksheet.prototype.getTableByName = function(tableName){
		var res = null;
		if(!this.TableParts || !tableName)
			return res;

		let _tableName = tableName.toLowerCase();
		for(var i = 0; i < this.TableParts.length; i++)
		{
			if(this.TableParts[i].DisplayName.length === _tableName.length && this.TableParts[i].DisplayName.toLowerCase() === _tableName)
			{
				res = this.TableParts[i];
				break;
			}
		}
		return res;
	};
	Worksheet.prototype.getTableRangeColumnByName = function(tableName, columnName){
		var res = null;
		if(!this.TableParts || !tableName)
			return res;

		let _tableName = tableName.toLowerCase();
		for(var i = 0; i < this.TableParts.length; i++)
		{
			if(this.TableParts[i].DisplayName.length === _tableName.length && this.TableParts[i].DisplayName.toLowerCase() === _tableName)
			{
				res = this.TableParts[i].getTableRangeColumnByName(columnName);
				break;
			}
		}
		return res;
	};
	Worksheet.prototype.isTableTotal = function (col, row) {
		if (this.TableParts && this.TableParts.length > 0) {
			for (var i = 0; i < this.TableParts.length; i++) {
				if (this.TableParts[i].isTotalsRow()) {
					var ref = this.TableParts[i].Ref;
					if (ref.r2 === row && col >= ref.c1 && col <= ref.c2) {
						return {index: i, colIndex: col - ref.c1};
					}
				}
			}
		}
		return null;
	};
	Worksheet.prototype.isApplyFilterBySheet = function () {
		var res = false;

		if (this.AutoFilter && this.AutoFilter.isApplyAutoFilter()) {
			res = true;
		}

		if (false === res && this.TableParts) {
			for (var i = 0; i < this.TableParts.length; i++) {
				if (true === this.TableParts[i].isApplyAutoFilter()) {
					res = true;
					break;
				}
			}
		}

		return res;
	};
	Worksheet.prototype.getTableNames = function() {
		var res = [];
		if (this.TableParts) {
			for (var i = 0; i < this.TableParts.length; i++) {
				res.push(this.TableParts[i].DisplayName);
			}
		}
		return res;
	};
	Worksheet.prototype.renameDependencyNodes = function(offset, oBBox){
		return this.workbook.dependencyFormulas.shift(this.Id, oBBox, offset);
	};
	Worksheet.prototype.getAllCol = function(){
		if(null == this.oAllCol)
			this.oAllCol = new AscCommonExcel.Col(this, g_nAllColIndex);
		return this.oAllCol;
	};
	Worksheet.prototype.getAllRow = function(){
		if (null == this.oSheetFormatPr.oAllRow) {
			this.oSheetFormatPr.oAllRow = new AscCommonExcel.Row(this);
			this.oSheetFormatPr.oAllRow.setIndex(g_nAllRowIndex);
		}
		return this.oSheetFormatPr.oAllRow;
	};
	Worksheet.prototype.getAllRowNoEmpty = function(){
		return this.oSheetFormatPr.oAllRow;
	};
	Worksheet.prototype.getHyperlinkByCell = function(row, col){
		var oHyperlink = this.hyperlinkManager.getByCell(row, col);
		return oHyperlink ? oHyperlink.data : null;
	};
	Worksheet.prototype.getMergedByCell = function(row, col){
		var oMergeInfo = this.mergeManager.getByCell(row, col);
		return oMergeInfo ? oMergeInfo.bbox : null;
	};
	Worksheet.prototype.getMergedByRange = function(bbox){
		return this.mergeManager.get(bbox);
	};
	Worksheet.prototype._expandRangeByMergedAddToOuter = function(aOuter, range, aMerged){
		for(var i = 0, length = aMerged.all.length; i < length; i++)
		{
			var elem = aMerged.all[i];
			if(!range.containsRange(elem.bbox))
				aOuter.push(elem);
		}
	};
	Worksheet.prototype._expandRangeByMergedGetOuter = function(range){
		var aOuter = [];
		//смотрим только границы
		this._expandRangeByMergedAddToOuter(aOuter, range, this.mergeManager.get(new Asc.Range(range.c1, range.r1, range.c1, range.r2)));
		if(range.c1 != range.c2)
		{
			this._expandRangeByMergedAddToOuter(aOuter, range, this.mergeManager.get(new Asc.Range(range.c2, range.r1, range.c2, range.r2)));
			if(range.c2 - range.c1 > 1)
			{
				this._expandRangeByMergedAddToOuter(aOuter, range, this.mergeManager.get(new Asc.Range(range.c1 + 1, range.r1, range.c2 - 1, range.r1)));
				if(range.r1 != range.r2)
					this._expandRangeByMergedAddToOuter(aOuter, range, this.mergeManager.get(new Asc.Range(range.c1 + 1, range.r2, range.c2 - 1, range.r2)));
			}
		}
		return aOuter;
	};
	Worksheet.prototype.expandRangeByMerged = function(range){
		if(null != range)
		{
			var aOuter = this._expandRangeByMergedGetOuter(range);
			if(aOuter.length > 0)
			{
				range = range.clone();
				while(aOuter.length > 0)
				{
					for(var i = 0, length = aOuter.length; i < length; i++)
						range.union2(aOuter[i].bbox);
					aOuter = this._expandRangeByMergedGetOuter(range);
				}
			}
		}
		return range;
	};
	Worksheet.prototype.createTablePart = function(){

		return new AscCommonExcel.TablePart(this.handlers);
	};
	Worksheet.prototype.onUpdateRanges = function(ranges) {
		this.workbook.updateSparklineCache(this.sName, ranges);
		// ToDo do not update conditional formatting on hidden sheet
		if (ranges && ranges.length) {
			this.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange(ranges));
		}
		//this.workbook.handlers.trigger("toggleAutoCorrectOptions", null,true);
	};
	Worksheet.prototype.updateSparklineCache = function (sheet, ranges) {
		for (var i = 0; i < this.aSparklineGroups.length; ++i) {
			this.aSparklineGroups[i].updateCache(sheet, ranges);
		}
	};
	Worksheet.prototype.getSparklineGroup = function (c, r) {
		for (var i = 0; i < this.aSparklineGroups.length; ++i) {
			if (-1 !== this.aSparklineGroups[i].contains(c, r)) {
				return this.aSparklineGroups[i];
			}
		}
		return null;
	};
	Worksheet.prototype.removeSparklines = function (range) {
		for (var i = this.aSparklineGroups.length - 1; i > -1; --i) {
			if (this.aSparklineGroups[i].remove(range)) {
				AscCommon.History.Add(new AscDFH.CChangesDrawingsSparklinesRemove(this.aSparklineGroups[i]));
				this.aSparklineGroups.splice(i, 1);
			}
		}
	};
	Worksheet.prototype.removeSparklineGroups = function (range) {
		for (var i = this.aSparklineGroups.length - 1; i > -1; --i) {
			if (-1 !== this.aSparklineGroups[i].intersectionSimple(range)) {
				AscCommon.History.Add(new AscDFH.CChangesDrawingsSparklinesRemove(this.aSparklineGroups[i]));
				this.aSparklineGroups.splice(i, 1);
			}
		}
	};
	Worksheet.prototype.addSparklineGroups = function (sparklineGroups) {
		if (sparklineGroups) {
			AscCommon.History.Add(new AscDFH.CChangesDrawingsSparklinesRemove(sparklineGroups, true));
			this.insertSparklineGroup(sparklineGroups);
		}
	};
	Worksheet.prototype.insertSparklineGroup = function (sparklineGroup) {
		this.aSparklineGroups.push(sparklineGroup);
	};
	Worksheet.prototype.removeSparklineGroup = function (id) {
		for (var i = 0; i < this.aSparklineGroups.length; ++i) {
			if (id === this.aSparklineGroups[i].Get_Id()) {
				this.aSparklineGroups.splice(i, 1);
				break;
			}
		}
	};
	Worksheet.prototype.updateSparklineGroupOffset = function (range, offset) {
		if ((offset.row < 0 || offset.col < 0)) {
			this.removeSparklines(range);
		}
		this.setSparklinesOffset(range, offset);
	};
	Worksheet.prototype.setSparklinesOffset = function (range, offset) {
		this.aSparklineGroups.forEach(function (val) {
			if (val) {
				var aSparklines = [];
				var isChange = false;
				for (var i = 0; i < val.arrSparklines.length; i++) {
					var _elem = val.arrSparklines[i];

					var cloneElem = _elem.clone();
					var _isChange = false;
					if (range.isIntersectForShift(cloneElem.sqRef, offset)) {
						_isChange = cloneElem.sqRef.forShift(range, offset);
					}
					if (_isChange) {
						isChange = true;
					}

					//необходимо ещё сдвинуть _f
					if ((offset.row < 0 || offset.col < 0) && _elem && _elem._f && range.containsRange(_elem._f)) {
						_elem.f = null;
						_elem._f = null;
					} else {
						_isChange = false;
						if (cloneElem._f) {
							if (range.isIntersectForShift(cloneElem._f, offset)) {
								_isChange = cloneElem._f.forShift(range, offset);
							}
							if (_isChange) {
								isChange = true;
								AscCommonExcel.executeInR1C1Mode(false, function () {
									cloneElem.f = cloneElem._f.getName();
								});
							}
						}
					}

					aSparklines.push(cloneElem);
				}
				if (isChange) {
					val.setSparklines(aSparklines, true, true);
				}
			}
		});
	};
	Worksheet.prototype.moveSparklineGroup = function (oBBoxFrom, oBBoxTo, copyRange, offset, wsTo, wsFrom) {
		var t = this;
		var isMove = !wsFrom;
		if (!wsTo) {
			wsTo = this;
		}
		if (false === this.workbook.bUndoChanges && false === this.workbook.bRedoChanges && !copyRange) {
			if (!wsFrom) {
				wsFrom = this;
			}

			//TODO проверить, необходимо ли чистить?
			//чистим ту область, куда переносим
			//wsTo.removeSparklineGroup

			wsFrom.aSparklineGroups.forEach(function (val) {
				if (val) {
					var aSparklines = [];
					var aSparkLinesToSheet = [];
					var isChange = false;
					for (var i = 0; i < val.arrSparklines.length; i++) {
						var _elem = val.arrSparklines[i];

						var cloneElem = _elem.clone();
						var _isChange = false;
						var moveOnNewSheet = false;

						if (wsTo === wsFrom) {
							if (oBBoxFrom.containsRange(cloneElem.sqRef)) {
								cloneElem.sqRef.setOffset(offset);
								_isChange = true;
							}
							if (_isChange) {
								isChange = true;
							}
						} else {
							if (oBBoxFrom.containsRange(cloneElem.sqRef)) {
								//тут необходимо добавить новый спакрлайн на новом листе
								moveOnNewSheet = true;
								cloneElem.sqRef.setOffset(offset);
								aSparkLinesToSheet.push(cloneElem);
								isChange = true;
							}
						}

						//необходимо ещё сдвинуть _f
						if (_elem && _elem._f && cloneElem._f) {
							if (wsTo === wsFrom) {
								if (cloneElem._f.intersection(oBBoxFrom)) {
									var _isChangeF = false;
									var isContainsFormula = oBBoxFrom.containsRange(cloneElem._f);
									if (isContainsFormula) {
										cloneElem._f.setOffset(offset);
										_isChangeF = true;
									} else {

										//при перемещении есть такой нюанс - если перемещаем последнюю ячейку формулы вниз(вверх аналогично для первой строки) + не затрагиваем весь диапазон формулы
										//то формулу растягиваем до перемещаемой ячейки

										var bColF = cloneElem._f.c1 === cloneElem._f.c2;
										if (bColF) {
											if (oBBoxFrom.contains(cloneElem._f.c1, cloneElem._f.r2) && !offset.col) {
												cloneElem._f.setOffsetLast(offset);
												_isChangeF = true;
											} else if (oBBoxFrom.contains(cloneElem._f.c1, cloneElem._f.r1)&& !offset.col) {
												cloneElem._f.setOffsetFirst(offset);
												_isChangeF = true;
											}
										} else {
											if (oBBoxFrom.contains(cloneElem._f.c2, cloneElem._f.r1) && !offset.row) {
												cloneElem._f.setOffsetLast(offset);
												_isChangeF = true;
											} else if (oBBoxFrom.contains(cloneElem._f.c1, cloneElem._f.r1)&& !offset.row) {
												cloneElem._f.setOffsetFirst(offset);
												_isChangeF = true;
											}
										}
									}

									if (_isChangeF) {
										if (wsTo.sName !== cloneElem._f.sheet) {
											cloneElem._f.setSheet(wsTo.sName);
										}
										AscCommonExcel.executeInR1C1Mode(false, function () {
											cloneElem.f = cloneElem._f.getName();
										});
										isChange = true;
									}
								}
							} else {
								if (isMove) {
									//если реальный перенос внутри книги, а не копирование/вставка
									if (oBBoxFrom.containsRange(cloneElem._f)) {
										cloneElem._f.setOffset(offset);
										if (wsTo.sName !== cloneElem._f.sheet) {
											cloneElem._f.setSheet(wsTo.sName);
										}
										AscCommonExcel.executeInR1C1Mode(false, function () {
											cloneElem.f = cloneElem._f.getName();
										});
										isChange = true;
									}
								} else {
									cloneElem._f.setOffset(offset);
									if (wsTo.sName !== cloneElem._f.sheet) {
										cloneElem._f.setSheet(wsTo.sName);
									}
									AscCommonExcel.executeInR1C1Mode(false, function () {
										cloneElem.f = cloneElem._f.getName();
									});
									isChange = true;
								}
							}
						}

						if (!moveOnNewSheet) {
							aSparklines.push(cloneElem);
						}
					}
					if (isChange) {
						if (wsTo !== wsFrom && !copyRange) {
							wsFrom.removeSparklines(oBBoxFrom);
						}

						//если, допустим, перенесли все спарклайны на другой лист, то необходимо удалить группы здесь и добавить новую группы
						if (aSparklines.length !== 0) {
							val.setSparklines(aSparklines, true, true);
						}
						if (aSparkLinesToSheet.length !== 0) {
							var modelSparkline = new AscCommonExcel.sparklineGroup(true);
							modelSparkline.worksheet = wsTo;
							modelSparkline.set(val);
							modelSparkline.setSparklines(aSparkLinesToSheet, true);
							wsTo.addSparklineGroups(modelSparkline);
						}
					}
				}
			});
		}
	};
	Worksheet.prototype.getSparkLinesMaxColRow = function() {
		var sparkLinesGroup = this.aSparklineGroups;
		var r = -1, c = -1;

		if (sparkLinesGroup) {
			sparkLinesGroup.forEach(function (item) {
				if (item.arrSparklines) {
					for (var j = 0; j < item.arrSparklines.length; ++j) {
						var range = item.arrSparklines[j].sqRef;
						if (range) {
							r = Math.max(r, range.r2);
							c = Math.max(c, range.c2);
						}
					}
				}
			});
		}

		return new AscCommon.CellBase(r, c);
	};
	Worksheet.prototype.getCwf = function() {
		var cwf = {};
		var range = this.getRange3(0,0, gc_nMaxRow0, gc_nMaxCol0);
		range._setPropertyNoEmpty(null, null, function(cell){
			if(cell.isFormula()){
				var name = cell.getName();
				cwf[name] = name;
			}
		});
		return cwf;
	};
	Worksheet.prototype.getAllFormulas = function(formulas, needReturnCellProps) {
		var range = this.getRange3(0, 0, gc_nMaxRow0, gc_nMaxCol0);
		range._setPropertyNoEmpty(null, null, function(cell) {
			if (cell.isFormula()) {
				if (!cell.formulaParsed.ref || (cell.formulaParsed.ref && cell.formulaParsed.checkFirstCellArray(cell))) {
					formulas.push(needReturnCellProps ? {f: cell.getFormulaParsed(), c: cell.nCol, r: cell.nRow} : cell.getFormulaParsed());
				}
			}
		});
		for (var i = 0; i < this.TableParts.length; ++i) {
			var table = this.TableParts[i];
			table.getAllFormulas(formulas);
		}
	};
	Worksheet.prototype.forEachFormula = function(callback) {
		var range = this.getRange3(0, 0, gc_nMaxRow0, gc_nMaxCol0);
		range._setPropertyNoEmpty(null, null, function(cell) {
			if (cell.isFormula()) {
				callback(cell.getFormulaParsed());
			}
		});
		for (var i = 0; i < this.TableParts.length; ++i) {
			var table = this.TableParts[i];
			table.forEachFormula(callback);
		}
	};
	Worksheet.prototype.setTableStyleAfterOpen = function () {
		if (this.TableParts && this.TableParts.length) {
			for (var i = 0; i < this.TableParts.length; i++) {
				var table = this.TableParts[i];
				this.autoFilters._setColorStyleTable(table.Ref, table);
			}
		}
	};
	Worksheet.prototype.setTableFormulaAfterOpen = function () {
		if (this.TableParts && this.TableParts.length) {
			for (var i = 0; i < this.TableParts.length; i++) {
				var table = this.TableParts[i];
				//TODO пока заменяем при открытии на TotalsRowFormula
				table.checkTotalRowFormula(this);
			}
		}
	};
	Worksheet.prototype.isTableTotalRow = function (range) {
		var res = false;
		if (this.TableParts && this.TableParts.length) {
			for (var i = 0; i < this.TableParts.length; i++) {
				var table = this.TableParts[i];
				var totalRowRange = table.getTotalsRowRange();
				if(totalRowRange && totalRowRange.containsRange(range)) {
					res = true;
				}
			}
		}
		return res;
	};
	Worksheet.prototype.addTablePart = function (tablePart, bAddToDependencies) {
		this.TableParts.push(tablePart);
		if (bAddToDependencies) {
			this.workbook.dependencyFormulas.addTableName(this, tablePart);
			tablePart.buildDependencies();
		}
	};
	Worksheet.prototype.changeTablePart = function (index, tablePart, bChangeName) {
		var oldTablePart = this.TableParts[index];
		if (oldTablePart) {
			oldTablePart.removeDependencies();
		}
		this.TableParts[index] = tablePart;
		tablePart.buildDependencies();
		if (bChangeName && oldTablePart) {
			this.workbook.dependencyFormulas.changeTableName(oldTablePart.DisplayName, tablePart.DisplayName);
		}
	};
	Worksheet.prototype.deleteTablePart = function (index, bConvertTableFormulaToRef) {
		var tablePart = this.TableParts[index];
		this.workbook.dependencyFormulas.delTableName(tablePart.DisplayName, bConvertTableFormulaToRef);
		tablePart.removeDependencies();

		//delete table
		this.TableParts.splice(index, 1);
	};
	Worksheet.prototype.checkPivotReportLocationForError = function(ranges, exceptPivot) {
		for (var i = 0; i < ranges.length; ++i) {
			var range = ranges[i];
			if (this.autoFilters.isIntersectionTable(range)) {
				return c_oAscError.ID.PivotOverlap;
			}
			if (this.inPivotTable(range, exceptPivot)) {
				return c_oAscError.ID.PivotOverlap;
			}
			var merged = this.mergeManager.get(range);
			if (merged.outer.length > 0) {
				return c_oAscError.ID.PastInMergeAreaError;
			}
		}
		return c_oAscError.ID.No;
	};
	Worksheet.prototype.checkPivotReportLocationForConfirm = function(ranges, changed) {
		var t = this;
		if (changed && changed.oldRanges && changed.data) {
			changed.oldRanges.forEach(function(range) {
				t.getRange3(range.r1, range.c1, range.r2, range.c2).cleanAll();
			});
		}
		for (var i = 0; i < ranges.length; ++i) {
			var range = ranges[i];
			var merged = this.mergeManager.get(range);
			if (merged.inner.length > 0) {
				return c_oAscError.ID.PivotOverlap;
			}
			var warning = c_oAscError.ID.No;
			this.getRange3(range.r1, range.c1, range.r2, range.c2)._foreachNoEmptyByCol(function(cell) {
				if (!cell.isNullTextString()) {
					warning = c_oAscError.ID.PivotOverlap;
				}
			});
			if (c_oAscError.ID.No !== warning) {
				return warning;
			}
		}
		return c_oAscError.ID.No;
	};
	Worksheet.prototype.getSlicerCachesByPivotTable = function(sheetId, pivotName) {
		var res = [];
		for (var i = 0; i < this.aSlicers.length; i++) {
			var cache = this.aSlicers[i].getSlicerCache();
			if (cache && cache.indexOfPivotTable(sheetId, pivotName) >= 0) {
				res.push(cache);
			}
		}
		return res;
	};
	Worksheet.prototype.getSlicerCachesByPivotCacheId = function(pivotCacheId) {
		var res = [];
		for (var i = 0; i < this.aSlicers.length; i++) {
			var cache = this.aSlicers[i].getSlicerCache();
			if (cache && pivotCacheId === cache.getPivotCacheId()) {
				res.push(cache);
			}
		}
		return res;
	};
	Worksheet.prototype.getSlicersByPivotTable = function(sheetId, pivotName) {
		var res = [];
		for (var i = 0; i < this.aSlicers.length; i++) {
			var cache = this.aSlicers[i].getSlicerCache();
			if (cache && cache.indexOfPivotTable(sheetId, pivotName) >= 0) {
				res.push(this.aSlicers[i]);
			}
		}
		return res;
	};
	Worksheet.prototype.updatePivotTable = function(pivotTable, changed, dataRow, canModifyDocument) {
		if (!changed.data && !changed.style) {
			return;
		}
		var t = this;
		var multiplyRange = new AscCommonExcel.MultiplyRange([]);
		if (changed.oldRanges) {
			multiplyRange.union2(new AscCommonExcel.MultiplyRange(changed.oldRanges));
		}
		pivotTable.init();
		if (changed.data && dataRow) {
			var newRanges = pivotTable.getReportRanges();
			newRanges.forEach(function(range) {
				t.getRange3(range.r1, range.c1, range.r2, range.c2).cleanAll();
			});
			this._updatePivotTableCells(pivotTable, dataRow);
		}
		var res = pivotTable.getReportRanges();
		multiplyRange.union2(new AscCommonExcel.MultiplyRange(res));
		this.updatePivotTablesStyle(multiplyRange.getUnionRange(), canModifyDocument);
		return res;
	};
	Worksheet.prototype.clearPivotTableCell = function (pivotTable) {
		var ranges = pivotTable.getReportRanges();
		ranges.forEach(function(range) {
			pivotTable.GetWS().getRange3(range.r1, range.c1, range.r2, range.c2).cleanAll();
		});
	};
	Worksheet.prototype.clearPivotTableStyle = function (pivotTable) {
		var ranges = pivotTable.getReportRanges();
		ranges.forEach(function(range){
			pivotTable.GetWS().getRange3(range.r1, range.c1, range.r2, range.c2).clearTableStyle();
		});
	};
	Worksheet.prototype._updatePivotTableSetCellValue = function (cell, text) {
		var oCellValue = new AscCommonExcel.CCellValue();
		oCellValue.type = AscCommon.CellValueType.String;
		oCellValue.text = text.toString();
		cell.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, oCellValue));
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable 
	 */
	Worksheet.prototype._updatePivotTableCellsPage = function (pivotTable) {
		//CT_pivotTableDefinition.prototype.getLayoutByCellPage
		if (pivotTable.pageFieldsPositions) {
			const fieldValuesCache = [];
			for (let i = 0; i < pivotTable.pageFieldsPositions.length; ++i) {
				const pos = pivotTable.pageFieldsPositions[i];
				let cells = this.getRange4(pos.row, pos.col);
				this._updatePivotTableSetCellValue(cells, pivotTable.getPageFieldName(i));
				const pageField = pivotTable.asc_getPageFields()[i];
				const index = pageField.asc_getIndex();
				cells.setStyle(pivotTable.getFormatting({
					axis: Asc.c_oAscAxis.AxisPage,
					type: Asc.c_oAscPivotAreaType.Button,
					isData: false,
					field: index
				}));
				cells = this.getRange4(pos.row, pos.col + 1);
				const num = pivotTable.getPivotFieldNum(pos.pageField.fld);
				if (num) {
					cells.setNum(num);
				}
				const oCellValue = pivotTable.getPageFieldCellValue(i);
				if (oCellValue.type !== AscCommon.CellValueType.String) {
					cells.setAlignHorizontal(AscCommon.align_Left);
				}
				const fieldValues = {
					fieldIndex: index,
					value: pageField.item,
					type: Asc.c_oAscItemType.Data
				};
				let valuesInfo;
				if (pageField.item !== null) {
					fieldValuesCache.push(fieldValues);
					valuesInfo = fieldValuesCache;
				} else {
					valuesInfo = fieldValuesCache.concat([fieldValues]);
				}
				cells.setStyle(pivotTable.getFormatting({
					axis: Asc.c_oAscAxis.AxisPage,
					type: Asc.c_oAscPivotAreaType.Normal,
					isData: false,
					field: index,
					valuesInfo: valuesInfo
				}));
				cells.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, oCellValue));
			}
		}
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable
	 */
	Worksheet.prototype._updatePivotTableCellsHeader = function (pivotTable) {
		//CT_pivotTableDefinition.prototype.getLayoutByCellHeader
		var location = pivotTable.location;
		if (0 === location.firstHeaderRow) {
			return;
		}
		var cells;
		var pivotRange = pivotTable.getRange();
		var rowFields = pivotTable.asc_getRowFields();
		var colFields = pivotTable.asc_getColumnFields();
		var dataFields = pivotTable.asc_getDataFields();
		var cacheFields = pivotTable.asc_getCacheFields();
		var pivotFields = pivotTable.asc_getPivotFields();
		if (dataFields && 1 === dataFields.length) {
			if (pivotTable.gridDropZones) {
				cells = this.getRange4(pivotRange.r1, pivotRange.c1);
			} else {
				if (rowFields && !colFields) {
					cells = this.getRange4(pivotRange.r1, pivotRange.c1 + location.firstDataCol);
				} else if (!rowFields && colFields) {
					cells = this.getRange4(pivotRange.r1 + location.firstDataRow, pivotRange.c1);
				} else {
					cells = this.getRange4(pivotRange.r1, pivotRange.c1);
				}
			}
			this._updatePivotTableSetCellValue(cells, pivotTable.getDataFieldName(0));
			const formatting = pivotTable.getFormatting({
				type: Asc.c_oAscPivotAreaType.Normal,
				axis: Asc.c_oAscAxis.AxisValues,
				isData: false,
				field: null
			});
			cells.setStyle(formatting);
		}
		if (pivotTable.showHeaders && colFields) {
			cells = this.getRange4(pivotRange.r1, pivotRange.c1 + location.firstDataCol);
			if (pivotTable.compact) {
				let caption;
				if (1 === colFields.length && AscCommonExcel.st_VALUES === colFields[0].asc_getIndex()) {
					caption = pivotTable.dataCaption || AscCommon.translateManager.getValue(AscCommonExcel.DATA_CAPTION)
				} else {
					caption = pivotTable.colHeaderCaption || AscCommon.translateManager.getValue(AscCommonExcel.COL_HEADER_CAPTION);
				}
				this._updatePivotTableSetCellValue(cells, caption);
			} else {
				var offset = new AscCommon.CellBase(0, 1);
				for (var i = 0; i < colFields.length; ++i) {
					var index = colFields[i].asc_getIndex();
					if (AscCommonExcel.st_VALUES !== index) {
						this._updatePivotTableSetCellValue(cells, pivotFields[index].asc_getName() || cacheFields[index].asc_getName());
					} else {
						this._updatePivotTableSetCellValue(cells, pivotTable.dataCaption || AscCommon.translateManager.getValue(AscCommonExcel.DATA_CAPTION));
					}
					cells.setOffset(offset);
				}
			}
			// update all col buttons area formats
			for(let i = 0; i < colFields.length; i += 1) {
				const cells = this.getRange4(pivotRange.r1, pivotRange.c1 + location.firstDataCol + i);
				const index = colFields[i].asc_getIndex();
				const formatting = pivotTable.getFormatting({
					type: Asc.c_oAscPivotAreaType.Button,
					axis: Asc.c_oAscAxis.AxisCol,
					isData: false,
					field: index
				});
				cells.setStyle(formatting);
			}
			// update topRight pivot area formats
			const pivotTableLenC = pivotRange.c2 - pivotRange.c1;
			const lenC = pivotTableLenC - location.firstDataCol - colFields.length + 1;
			this.updatePivotTableCellsLablesOriginOffsets(pivotTable, {
				isData: false,
				type: Asc.c_oAscPivotAreaType.TopRight,
				field: null
			},  pivotRange.r1, pivotRange.c1 + location.firstDataCol + colFields.length, 1, lenC);
		}
		// update origin pivot area formats
		this.updatePivotTableCellsLablesOriginOffsets(pivotTable, {
			isData: false,
			type: Asc.c_oAscPivotAreaType.Origin,
			field: null
		},  pivotRange.r1, pivotRange.c1, location.firstDataRow - 1, location.firstDataCol);
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable
	 * @param {PivotFormatsManagerQuery} query
	 * @param {number} startR
	 * @param {number} startC
	 * @param {number} endR
	 * @param {number} endC
	 */
	Worksheet.prototype.updatePivotTableCellsLablesOriginOffsets = function(pivotTable, query, startR, startC, lenR, lenC) {
		for (let i = 0; i < lenR; i += 1) {
			for (let j = 0; j < lenC; j += 1) {
				const offset = new Asc.Range(j, i, j, i);
				const cell = this.getRange4(startR + i, startC + j);
				query.offset = offset;
				const formatting = pivotTable.getFormatting(query);
				cell.setStyle(formatting);
			}
		}
		query.offset = null;
		return;
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable
	 * @param {number[]} rowFieldsOffset
	 * @param {boolean} isRowItem
	 * @param {number} itemIndex
	 * @param {PivotFormatsManagerQuery} query
	 */
	Worksheet.prototype.updatePivotTableCellsRowColLablesSubtotalsOffsets = function(pivotTable, rowFieldsOffset, isRowItem, itemIndex, itemXIndex, query) {
		const pivotRange = pivotTable.getRange();
		const location = pivotTable.location;
		const items = isRowItem ? pivotTable.getRowItems() : pivotTable.getColItems();
		const itemR = items[itemIndex].getR();
		const r1 = isRowItem ? pivotRange.r1 + location.firstDataRow : pivotRange.r1 + location.firstHeaderRow;
		const c1 = isRowItem ? pivotRange.c1 : pivotRange.c1 + location.firstDataCol;
		const startRow = isRowItem ? r1 + itemIndex : r1 + itemR;
		const startCol = isRowItem ? c1 + rowFieldsOffset[itemXIndex] : c1 + itemIndex;
		// using for colItems Labels
		const endRow = r1 + location.firstDataRow - 1;
		// using for rowItems Labels
		const endCol = c1 + location.firstDataCol;
		const startIndex = isRowItem ? startCol : startRow;
		const endIndex = isRowItem ? endCol : endRow;
		for (let i = startIndex; i < endIndex; i += 1) {
			const r = isRowItem ? startRow : i;
			const c = isRowItem ? i : startCol;
			const offsetR = isRowItem ? Asc.st_PIVOT_AREA_OFFSET_END : (i === endIndex - 1 ? Asc.st_PIVOT_AREA_OFFSET_END : i - startIndex);
			const offsetC = isRowItem ? (i === endIndex - 1 ? Asc.st_PIVOT_AREA_OFFSET_END : i - startIndex) : Asc.st_PIVOT_AREA_OFFSET_END;
			
			const offset = new Asc.Range(offsetC, offsetR, offsetC, offsetR);
			const cell = this.getRange4(r, c);

			query.offset = offset;
			const formatting = pivotTable.getFormatting(query);
			cell.setStyle(formatting);
		}
		query.offset = null;
		return;
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable
	 * @param {number[]} rowFieldsOffset
	 * @param {boolean} isRowItem
	 * @param {number} itemIndex
	 * @param {PivotFormatsManagerQuery} query
	 */
	Worksheet.prototype.updatePivotTableCellsRowColLablesDataOffsets = function(pivotTable, rowFieldsOffset, isRowItem, itemIndex, itemXIndex, query) {
		const pivotRange = pivotTable.getRange();
		const location = pivotTable.location;
		const items = isRowItem ? pivotTable.getRowItems() : pivotTable.getColItems();
		const itemR = items[itemIndex].getR();
		const itemSummaryR = itemR + itemXIndex;

		const r1 = isRowItem ? pivotRange.r1 + location.firstDataRow : pivotRange.r1 + location.firstHeaderRow;
		const c1 = isRowItem ? pivotRange.c1 : pivotRange.c1 + location.firstDataCol;
		const startRow = isRowItem ? r1 + itemIndex: r1 + itemSummaryR;
		const startCol = isRowItem ? c1 + rowFieldsOffset[itemSummaryR] : c1 + itemIndex;
		for (let i = itemIndex; i < items.length; i += 1) {
			const item = items[i];
			const nextItem = items[i + 1];
			const r = isRowItem ? r1 + i : startRow;
			const c = isRowItem ? startCol : c1 + i;
			if (isRowItem && c1 + rowFieldsOffset[item.getR()] <= startCol && i !== itemIndex) {
				break;
			}
			if (!isRowItem && r1 + item.getR() <= startRow && i !== itemIndex) {
				break;
			}
			const offsetR = isRowItem ? ((!nextItem || (c1 + rowFieldsOffset[nextItem.getR()] <= startCol)) ? Asc.st_PIVOT_AREA_OFFSET_END : r - startRow) : Asc.st_PIVOT_AREA_OFFSET_END;
			const offsetC = isRowItem ? Asc.st_PIVOT_AREA_OFFSET_END : ((!nextItem || (r1 + nextItem.getR() <= startRow)) ? Asc.st_PIVOT_AREA_OFFSET_END : c - startCol);
			
			const offset = new Asc.Range(offsetC, offsetR, offsetC, offsetR);
			const cell = this.getRange4(r, c);

			query.offset = offset;
			const formatting = pivotTable.getFormatting(query);
			cell.setStyle(formatting);
		}
		query.offset = null;
		return;
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable
	 * @param {number[]} rowFieldsOffset
	 * @param {boolean} isRowItem
	 * @param {number} itemIndex
	 * @param {PivotFormatsManagerQuery} query
	 */
	Worksheet.prototype._updatePivotTableCellsRowColLablesOffsets = function(pivotTable, rowFieldsOffset, isRowItem, itemIndex, itemXIndex, query) {
		const items = isRowItem ? pivotTable.getRowItems() : pivotTable.getColItems();
		query.axis = isRowItem ? Asc.c_oAscAxis.AxisRow : Asc.c_oAscAxis.AxisCol;
		const item = items[itemIndex];
		if (item.t === Asc.c_oAscItemType.Data) {
			this.updatePivotTableCellsRowColLablesDataOffsets(pivotTable, rowFieldsOffset, isRowItem, itemIndex, itemXIndex, query);
		} else {
			this.updatePivotTableCellsRowColLablesSubtotalsOffsets(pivotTable, rowFieldsOffset, isRowItem, itemIndex, itemXIndex, query);
		}
		return;
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable 
	 * @param {number[]} rowFieldsOffset 
	 */
	Worksheet.prototype._updatePivotTableCellsRowColLables = function(pivotTable, rowFieldsOffset) {
		//CT_pivotTableDefinition.prototype.getLayoutByCellHeaderRowColLables
		var items, fields, field, oCellValue, fieldIndex, cells, r1, c1, i, j, valuesIndex, isRowItems = false;
		var pivotRange = pivotTable.getRange();
		var location = pivotTable.location;
		var hasLeftAlignInRowLables = false;
		var outlines = [0];
		if (rowFieldsOffset) {
			items = pivotTable.getRowItems();
			isRowItems = true;
			fields = pivotTable.asc_getRowFields();
			r1 = pivotRange.r1 + location.firstDataRow;
			c1 = pivotRange.c1;
			valuesIndex = pivotTable.getRowFieldsValuesIndex();
			hasLeftAlignInRowLables = pivotTable.hasLeftAlignInRowLables();
			for (i = 1; i < rowFieldsOffset.length; ++i) {
				outlines[i] = outlines[i - 1] + 1;
				if (rowFieldsOffset[i] !== rowFieldsOffset[i - 1]) {
					outlines[i] = 0;
				}
			}
		} else {
			items = pivotTable.getColItems();
			fields = pivotTable.asc_getColumnFields();
			r1 = pivotRange.r1 + location.firstHeaderRow;
			c1 = pivotRange.c1 + location.firstDataCol;
			valuesIndex = pivotTable.getColumnFieldsValuesIndex();
		}
		if (!items || !fields) {
			if (pivotTable.gridDropZones && !fields && pivotTable.getDataFieldsCount() > 0) {
				if (rowFieldsOffset) {
					cells = this.getRange4(pivotRange.r1 + location.firstDataRow, pivotRange.c1 + location.firstDataCol - 1);
				} else {
					cells = this.getRange4(pivotRange.r1 + location.firstDataRow - 1, pivotRange.c1 + location.firstDataCol);
				}
				oCellValue = new AscCommonExcel.CCellValue();
				oCellValue.type = AscCommon.CellValueType.String;
				oCellValue.text = AscCommon.translateManager.getValue(AscCommonExcel.ToName_ST_ItemType(Asc.c_oAscItemType.Default));
				cells.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, oCellValue));
			}
			if (rowFieldsOffset && false === pivotTable.showHeaders && false === pivotTable.gridDropZones && pivotTable.getDataFieldsCount() > 0) {
				cells = this.getRange4(pivotRange.r1 + location.firstDataRow, pivotRange.c1 + location.firstDataCol - 1);
				this._updatePivotTableSetCellValue(cells, pivotTable.getDataFieldName(0));
			}
			return;
		}
		var pivotFields = pivotTable.asc_getPivotFields();

		var totalTitleRange = [];
		var valuesCache = [];
		for (i = 0; i < items.length; ++i) {
			var item = items[i];
			/** @type {PivotItemFieldsInfo[]} */
			var fieldValues = [];
			var r = item.getR();
			for (j = 0; j < r; ++j) {
				fieldIndex = fields[j].asc_getIndex();
				field = pivotFields[fieldIndex];
				fieldValues.push(valuesCache[j]);
				if (AscCommonExcel.st_VALUES !== fieldIndex && field.asc_getFillDownLabelsDefault() && totalTitleRange[j]) {
					if (rowFieldsOffset) {
						cells = this.getRange4(r1 + i, c1 + rowFieldsOffset[j]);
					} else {
						cells = this.getRange4(r1 + j, c1 + i);
					}
					cells.setStyle(totalTitleRange[j].getStyle());
					cells.setValueData(totalTitleRange[j].getValueData());
				}
			}
			for (j = 0; j < item.x.length; ++j) {
				/** @type {PivotFormatsManagerQuery} */
				const query = {
					valuesInfo: fieldValues,
					isData: false,
					isGrandRow: isRowItems && item.t === Asc.c_oAscItemType.Grand,
					isGrandCol: !isRowItems && item.t === Asc.c_oAscItemType.Grand,
					type: Asc.c_oAscPivotAreaType.Normal,
					field: null
				};
				fieldIndex = null;
				var outline = 0;
				if (rowFieldsOffset) {
					cells = this.getRange4(r1 + i, c1 + rowFieldsOffset[r + j]);
				} else {
					cells = this.getRange4(r1 + r + j, c1 + i);
				}
				if (Asc.c_oAscItemType.Data === item.t) {
					if (rowFieldsOffset) {
						outline = outlines[r + j];
					}
					fieldIndex = fields[r + j].asc_getIndex();
					if (AscCommonExcel.st_VALUES !== fieldIndex) {
						oCellValue = pivotTable.getPivotFieldCellValue(fieldIndex, item.x[j].getV());
						fieldValues.push({
							fieldIndex: fieldIndex,
							value: item.x[j].getV(),
							type: Asc.c_oAscItemType.Data
						});
						query.field = fieldIndex;
						this._updatePivotTableCellsRowColLablesOffsets(pivotTable, rowFieldsOffset, isRowItems, i, j, query);
					} else {
						fieldValues.push({
							fieldIndex: AscCommonExcel.st_DATAFIELD_REFERENCE_FIELD,
							value: item.i,
							type: Asc.c_oAscItemType.Data
						});
						query.field = AscCommonExcel.st_DATAFIELD_REFERENCE_FIELD;
						const formatting = pivotTable.getFormatting(query);
						cells.setStyle(formatting);
						oCellValue = new AscCommonExcel.CCellValue();
						oCellValue.type = AscCommon.CellValueType.String;
						oCellValue.text = pivotTable.getDataFieldName(item.i);
					}
					totalTitleRange[r + j] = cells;
				} else if (Asc.c_oAscItemType.Grand === item.t) {
					oCellValue = new AscCommonExcel.CCellValue();
					oCellValue.type = AscCommon.CellValueType.String;
					if (-1 === valuesIndex) {
						oCellValue.text = pivotTable.grandTotalCaption || AscCommon.translateManager.getValue(AscCommonExcel.GRAND_TOTAL_CAPTION);
						this._updatePivotTableCellsRowColLablesOffsets(pivotTable, rowFieldsOffset, isRowItems, i, j, query);
					} else {
						oCellValue.text = AscCommon.translateManager.getValue(AscCommonExcel.ToName_ST_ItemType(item.t));
						oCellValue.text += ' ' + pivotTable.getDataFieldName(item.i);
						fieldValues.push({
							fieldIndex: AscCommonExcel.st_DATAFIELD_REFERENCE_FIELD,
							value: item.i,
							type: Asc.c_oAscItemType.Data
						});
						this._updatePivotTableCellsRowColLablesOffsets(pivotTable, rowFieldsOffset, isRowItems, i, j, query);
					}
					if (fields && fields[0]) {
						if (pivotFields[fields[0].asc_getIndex()] && pivotFields[fields[0].asc_getIndex()].num) {
							cells.setNum(pivotFields[fields[0].asc_getIndex()].num);
						}
					}
				} else if (Asc.c_oAscItemType.Blank === item.t) {
					break;
				} else {
					if (rowFieldsOffset) {
						outline = outlines[r + j];
					}
					oCellValue = new AscCommonExcel.CCellValue();
					oCellValue.type = AscCommon.CellValueType.String;
					if (r + j > valuesIndex) {
						fieldIndex = fields[r + j].asc_getIndex();
						field = pivotFields[fieldIndex];
						fieldValues.push({
							fieldIndex: fieldIndex,
							value: item.x[j].getV(),
							type: item.t
						});
						if (AscCommonExcel.st_VALUES !== fieldIndex) {
							if (field.subtotalCaption) {
								oCellValue.text = field.subtotalCaption;
							} else {
								oCellValue.text = totalTitleRange[r + j].getValueWithFormatSkipToSpace();
								oCellValue.text += ' ' + AscCommon.translateManager.getValue(AscCommonExcel.ToName_ST_ItemType(item.t));
								query.field = fieldIndex;
								this._updatePivotTableCellsRowColLablesOffsets(pivotTable, rowFieldsOffset, isRowItems, i, j, query);
							}
						}
					} else {
						fieldIndex = fields[r + j].asc_getIndex();
						field = pivotFields[fieldIndex];
						fieldValues.push({
							fieldIndex: fieldIndex,
							value: item.x[j].getV(),
							type: item.t
						}, {
							fieldIndex: AscCommonExcel.st_DATAFIELD_REFERENCE_FIELD,
							value: item.i,
							type: Asc.c_oAscItemType.Data
						});
						query.field = fieldIndex;
						oCellValue.text = totalTitleRange[r + j].getValueWithFormatSkipToSpace();
						oCellValue.text += ' ' + pivotTable.getDataFieldName(item.i);
						this._updatePivotTableCellsRowColLablesOffsets(pivotTable, rowFieldsOffset, isRowItems, i, j, query);
					}
				}
				if (null !== fieldIndex) {
					var num = pivotTable.getPivotFieldNum(fieldIndex);
					if (num) {
						cells.setNum(num);
					}
				}
				if (hasLeftAlignInRowLables) {
					cells.setAlignHorizontal(AscCommon.align_Left);
				}
				if (outline > 0) {
					cells.setIndent(outline);
				}
				cells.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, oCellValue));
				valuesCache = fieldValues;
			}
		}
	};
	Worksheet.prototype._updatePivotTableCellsRowHeaderLabels = function(pivotTable) {
		//CT_pivotTableDefinition.prototype.getLayoutByCellHeaderRowColLables
		var rowFieldsOffset = [0];
		var rowFields = pivotTable.asc_getRowFields();
		if (!rowFields) {
			return rowFieldsOffset;
		}
		var cells, index, field;
		var pivotFields = pivotTable.asc_getPivotFields();
		var pivotRange = pivotTable.getRange();
		var location = pivotTable.location;
		var c1 = pivotRange.c1;
		var r1 = pivotRange.r1 + location.firstDataRow - 1;
		cells = this.getRange4(r1, c1);
		if (pivotTable.showHeaders) {
			if (pivotTable.compact || location.firstDataCol !== rowFields.length) {
				if(1 === rowFields.length && AscCommonExcel.st_VALUES === rowFields[0].asc_getIndex()){
					this._updatePivotTableSetCellValue(cells, pivotTable.dataCaption || AscCommon.translateManager.getValue(AscCommonExcel.DATA_CAPTION));
				} else {
					this._updatePivotTableSetCellValue(cells, pivotTable.rowHeaderCaption || AscCommon.translateManager.getValue(AscCommonExcel.ROW_HEADER_CAPTION));
				}
			} else {
				index = rowFields[0].asc_getIndex();
				if (AscCommonExcel.st_VALUES !== index) {
					this._updatePivotTableSetCellValue(cells, pivotTable.getPivotFieldName(index));
				} else {
					this._updatePivotTableSetCellValue(cells, pivotTable.dataCaption || AscCommon.translateManager.getValue(AscCommonExcel.DATA_CAPTION));
				}
			}
		}
		for (var i = 1; i < rowFields.length; ++i) {
			index = rowFields[i - 1].asc_getIndex();
			var isTabular;
			if (AscCommonExcel.st_VALUES !== index) {
				field = pivotFields[index];
				isTabular = field && !(field.compact && field.outline);
			} else {
				isTabular = !(pivotTable.compact && pivotTable.outline);
			}
			if (isTabular) {
				index = rowFields[i].asc_getIndex();
				++c1;
				if (pivotTable.showHeaders) {
					cells = this.getRange4(r1, c1);
					if (AscCommonExcel.st_VALUES !== index) {
						this._updatePivotTableSetCellValue(cells, pivotTable.getPivotFieldName(index));
					} else {
						this._updatePivotTableSetCellValue(cells, pivotTable.dataCaption || AscCommon.translateManager.getValue(AscCommonExcel.DATA_CAPTION));
					}
				}
			}
			rowFieldsOffset[i] = c1 - pivotRange.c1;
		}
		r1 = pivotRange.r1 + location.firstDataRow - 1;
		c1 =  pivotRange.c1;
		if (pivotTable.compact) {
			const cells = this.getRange4(r1, c1);
			const index = rowFields[0].asc_getIndex();
			const formatting = pivotTable.getFormatting({
				type: Asc.c_oAscPivotAreaType.Button,
				axis: Asc.c_oAscAxis.AxisRow,
				isData: false,
				field: index
			});
			cells.setStyle(formatting);
		} else {
			for (let i = 0; i < rowFields.length; i += 1) {
				const cells = this.getRange4(r1, c1 + i);
				const index = rowFields[i].asc_getIndex();
				const formatting = pivotTable.getFormatting({
					type: Asc.c_oAscPivotAreaType.Button,
					axis: Asc.c_oAscAxis.AxisRow,
					isData: false,
					field: index
				});
				cells.setStyle(formatting);
			}
		}
		return rowFieldsOffset;
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable 
	 */
	Worksheet.prototype.setFormatsCellsDataNoDataField = function(pivotTable) {
		const pivotRange = pivotTable.getRange();
		const location = pivotTable.location;
		const r1 = pivotRange.r1 + location.firstDataRow;
		const c1 = pivotRange.c1 + location.firstDataCol;
		const r2 = pivotRange.r2;
		const c2 = pivotRange.c2;
		for(let i = r1; i <= r2; i += 1) {
			for(let j = c1; j <= c2; j += 1) {
				const cell = this.getRange4(i, j);
				cell.setStyle(pivotTable.getFormatting({
					type: Asc.c_oAscPivotAreaType.All,
				}));
			}
		}
	};
	/**
	 * A function that updates the data in the cells of a pivot table.
	 * @param {CT_pivotTableDefinition} pivotTable
	 * @param {PivotDataElem} dataRow
	 */
	Worksheet.prototype._updatePivotTableCellsData = function(pivotTable, dataRow) {
		const rowItems = pivotTable.getRowItems();
		const colItems = pivotTable.getColItems();
		const dataFields = pivotTable.asc_getDataFields();
		if (!rowItems || !colItems || !dataFields) {
			this.setFormatsCellsDataNoDataField(pivotTable);
			return;
		}
		pivotTable.dataManager.update(dataRow);
	};
	/**
	 * @param {CT_pivotTableDefinition} pivotTable 
	 * @param {PivotDataElem} dataRow 
	 */
	Worksheet.prototype._updatePivotTableCells = function (pivotTable, dataRow) {
		pivotTable.formatsManager.updateCollection();
		this._updatePivotTableCellsPage(pivotTable);
		this._updatePivotTableCellsHeader(pivotTable);
		this._updatePivotTableCellsRowColLables(pivotTable);
		var rowFieldsOffset = this._updatePivotTableCellsRowHeaderLabels(pivotTable);
		this._updatePivotTableCellsRowColLables(pivotTable, rowFieldsOffset);
		this._updatePivotTableCellsData(pivotTable, dataRow);
	};
	Worksheet.prototype.updatePivotTablesStyle = function (range, canModifyDocument) {
		var t = this;
		var pivotTable, pivotRange, pivotFields, rowFields, styleInfo, style, wholeStyle, cells, j, r, x, pos,
			firstHeaderRow0, firstDataCol0, countC, countCWValues, countR, countD, stripe1, stripe2, items, l, item,
			start, end, isOutline, arrSubheading, emptyStripe = new Asc.CTableStyleElement();
		var dxf, dxfLabels, dxfValues, grandColumn, index;
		var checkRowSubheading = function (_i, _r, _v, _dxf) {
			var sub, bSet = true;
			if ((sub = arrSubheading[_i])) {
				if (sub.v === _v) {
					bSet = false;
				} else {
					cells = t.getRange3(sub.r, pivotRange.c1 + _i, _r - 1, pivotRange.c1 + _i);
					cells.setTableStyle(sub.dxf);
				}
			}
			if (bSet) {
				arrSubheading[_i] = (null === _v) ? null : {r: _r, dxf: _dxf, v: _v};
			}
		};
		var endRowSubheadings = function (_i, _r) {
			for (;_i < arrSubheading.length; ++_i) {
				checkRowSubheading(_i, _r, null, null);
			}
		};

		for (var i = 0; i < this.pivotTables.length; ++i) {
			grandColumn = 0;
			pivotTable = this.pivotTables[i];
			pivotRange = pivotTable.getRange();
			pivotFields = pivotTable.asc_getPivotFields();
			rowFields = pivotTable.asc_getRowFields();
			styleInfo = pivotTable.asc_getStyleInfo();
			if (!pivotTable.isInit || !styleInfo || (range && !pivotTable.intersection(range))) {
				continue;
			}
			style = this.workbook.TableStyles.AllStyles[styleInfo.asc_getName()];

			wholeStyle = style && style.wholeTable && style.wholeTable.dxf;

			// Page Field Labels, Page Field Values
			dxfLabels = style && style.pageFieldLabels && style.pageFieldLabels.dxf;
			dxfValues = style && style.pageFieldValues && style.pageFieldValues.dxf;
			for (j = 0; j < pivotTable.pageFieldsPositions.length; ++j) {
				pos = pivotTable.pageFieldsPositions[j];
				cells = this.getRange4(pos.row, pos.col);
				cells.clearTableStyle();
				cells.setTableStyle(wholeStyle);
				cells.setTableStyle(dxfLabels);
				cells = this.getRange4(pos.row, pos.col + 1);
				cells.clearTableStyle();
				cells.setTableStyle(wholeStyle);
				cells.setTableStyle(dxfValues);
			}

			cells = this.getRange3(pivotRange.r1, pivotRange.c1, pivotRange.r2, pivotRange.c2);
			cells.clearTableStyle();

			if (!style) {
				continue;
			}

			countC = pivotTable.getColumnFieldsCount();
			countR = pivotTable.getRowFieldsCount(true);

			if (pivotTable.isEmptyReport()) {
				if (canModifyDocument && 0 === pivotTable.pageFieldsPositions.length) {
					//todo transparent ih, iv
					var border;
					border = new AscCommonExcel.Border();
					border.initDefault();
					border.l = new AscCommonExcel.BorderProp();
					border.l.setStyle(c_oAscBorderStyles.Thin);
					border.l.c = AscCommonExcel.createRgbColor(0, 0, 0);
					border.t = new AscCommonExcel.BorderProp();
					border.t.setStyle(c_oAscBorderStyles.Thin);
					border.t.c = AscCommonExcel.createRgbColor(0, 0, 0);
					border.r = new AscCommonExcel.BorderProp();
					border.r.setStyle(c_oAscBorderStyles.Thin);
					border.r.c = AscCommonExcel.createRgbColor(0, 0, 0);
					border.b = new AscCommonExcel.BorderProp();
					border.b.setStyle(c_oAscBorderStyles.Thin);
					border.b.c = AscCommonExcel.createRgbColor(0, 0, 0);
					border.ih = new AscCommonExcel.BorderProp();
					border.ih.setStyle(c_oAscBorderStyles.Thin);
					border.ih.c = AscCommonExcel.createRgbColor(255, 255, 255);
					border.iv = new AscCommonExcel.BorderProp();
					border.iv.setStyle(c_oAscBorderStyles.Thin);
					border.iv.c = AscCommonExcel.createRgbColor(255, 255, 255);
					cells.setBorder(border);
				}
				continue;
			}

			firstHeaderRow0 = pivotTable.getFirstHeaderRow0();
			firstDataCol0 = pivotTable.getFirstDataCol();
			countD = pivotTable.getDataFieldsCount();
			countCWValues = pivotTable.getColumnFieldsCount(true);

			// Whole Table
			cells.setTableStyle(wholeStyle);

			// First Column Stripe, Second Column Stripe
			if (styleInfo.showColStripes) {
				stripe1 = style.firstColumnStripe || emptyStripe;
				stripe2 = style.secondColumnStripe || emptyStripe;
				start = pivotRange.c1 + firstDataCol0;
				if (stripe1.dxf) {
					cells = this.getRange3(pivotRange.r1 + firstHeaderRow0 + 1, start, pivotRange.r2, pivotRange.c2);
					cells.setTableStyle(stripe1.dxf, new Asc.CTableStyleStripe(stripe1.size, stripe2.size));
				}
				if (stripe2.dxf && start + stripe1.size <= pivotRange.c2) {
					cells = this.getRange3(pivotRange.r1 + firstHeaderRow0 + 1, start + stripe1.size, pivotRange.r2,
						pivotRange.c2);
					cells.setTableStyle(stripe2.dxf, new Asc.CTableStyleStripe(stripe2.size, stripe1.size));
				}
			}
			// First Row Stripe, Second Row Stripe
			if (styleInfo.showRowStripes && countR && (pivotRange.c1 + countR - 1 !== pivotRange.c2)) {
				stripe1 = style.firstRowStripe || emptyStripe;
				stripe2 = style.secondRowStripe || emptyStripe;
				start = pivotRange.r1 + firstHeaderRow0 + 1;
				if (stripe1.dxf) {
					cells = this.getRange3(start, pivotRange.c1, pivotRange.r2, pivotRange.c2);
					cells.setTableStyle(stripe1.dxf, new Asc.CTableStyleStripe(stripe1.size, stripe2.size, true));
				}
				if (stripe2.dxf && start + stripe1.size <= pivotRange.r2) {
					cells = this.getRange3(start + stripe1.size, pivotRange.c1, pivotRange.r2, pivotRange.c2);
					cells.setTableStyle(stripe1.dxf, new Asc.CTableStyleStripe(stripe2.size, stripe1.size, true));
				}
			}

			// First Column
			dxf = style.firstColumn && style.firstColumn.dxf;
			if (styleInfo.showRowHeaders && countR && dxf) {
				cells = this.getRange3(pivotRange.r1, pivotRange.c1, pivotRange.r2, pivotRange.c1 +
					Math.max(0, firstDataCol0 - 1));
				cells.setTableStyle(dxf);
			}

			// Header Row
			dxf = style.headerRow && style.headerRow.dxf;
			if (styleInfo.showColHeaders && dxf) {
				cells = this.getRange3(pivotRange.r1, pivotRange.c1, pivotRange.r1 + firstHeaderRow0, pivotRange.c2);
				cells.setTableStyle(dxf);
			}

			// First Header Cell
			dxf = style.firstHeaderCell && style.firstHeaderCell.dxf;
			if (styleInfo.showColHeaders && styleInfo.showRowHeaders && countCWValues && (countR + countD) && dxf) {
				cells = this.getRange3(pivotRange.r1, pivotRange.c1, pivotRange.r1 + firstHeaderRow0 - (countR ? 1 : 0),
					pivotRange.c1 + Math.max(0, firstDataCol0 - 1));
				cells.setTableStyle(dxf);
			}

			// Subtotal Column + Grand Total Column
			items = pivotTable.getColItems();
			if (items) {
				start = pivotRange.c1 + firstDataCol0;
				for (j = 0; j < items.length; ++j) {
					dxf = null;
					item = items[j];
					r = item.getR();
					if (Asc.c_oAscItemType.Grand === item.t || 0 === countCWValues) {
						// Grand Total Column
						dxf = style.lastColumn;
						grandColumn = 1;
					} else {
						// Subtotal Column
						if (r + 1 !== countC) {
							if (countD && Asc.c_oAscItemType.Data !== item.t) {
								if (0 === r) {
									dxf = style.firstSubtotalColumn;
								} else if (1 === r % 2) {
									dxf = style.secondSubtotalColumn;
								} else {
									dxf = style.thirdSubtotalColumn;
								}
							}
						}
					}
					dxf = dxf && dxf.dxf;
					if (dxf) {
						cells = this.getRange3(pivotRange.r1 + 1, start + j, pivotRange.r2, start + j);
						cells.setTableStyle(dxf);
					}
				}
			}

			// Subtotal Row + Row Subheading + Grand Total Row
			items = pivotTable.getRowItems();
			if (items && countR) {
				arrSubheading = [];
				countR = pivotTable.getRowFieldsCount();
				start = pivotRange.r1 + firstHeaderRow0 + 1;
				for (j = 0; j < items.length; ++j) {
					dxf = null;
					item = items[j];
					if (Asc.c_oAscItemType.Data !== item.t) {
						if (Asc.c_oAscItemType.Grand === item.t) {
							// Grand Total Row
							dxf = style.totalRow;
							pos = 0;
						} else if (Asc.c_oAscItemType.Blank === item.t) {
							// Blank Row
							dxf = style.blankRow;
							pos = 0;
						} else if (styleInfo.showRowHeaders) {
							// Subtotal Row
							r = item.getR();
							if (r + 1 !== countR) {
								if (0 === r) {
									dxf = style.firstSubtotalRow;
								} else if (1 === r % 2) {
									dxf = style.secondSubtotalRow;
								} else {
									dxf = style.thirdSubtotalRow;
								}
								pos = pivotTable.getRowFieldPos(r);
							}
						}
						dxf = dxf && dxf.dxf;
						if (dxf) {
							cells = this.getRange3(start + j, pivotRange.c1 + pos, start + j, pivotRange.c2);
							cells.setTableStyle(dxf);
						}
						endRowSubheadings(pos, start + j);
					} else if (styleInfo.showRowHeaders) {
						// Row Subheading
						r = item.getR();
						index = rowFields[r].asc_getIndex();
						isOutline = (AscCommonExcel.st_VALUES !== index && false !== pivotFields[index].outline);
						for (x = 0, l = item.x.length; x < l; ++x, ++r) {
							dxf = null;
							if (r + 1 !== countR) {
								if (0 === r) {
									dxf = style.firstRowSubheading;
								} else if (1 === r % 2) {
									dxf = style.secondRowSubheading;
								} else {
									dxf = style.thirdRowSubheading;
								}
								dxf = dxf && dxf.dxf;
								if (dxf) {
									pos = pivotTable.getRowFieldPos(r);
									if (1 === l && isOutline) {
										endRowSubheadings(pos, start + j);
										cells = this.getRange3(start + j, pivotRange.c1 + pos, start + j, pivotRange.c2);
										cells.setTableStyle(dxf);
									} else {
										checkRowSubheading(pos, start + j, item.x[x].getV(), dxf);
									}
								}
							}
						}
					}
				}
				endRowSubheadings(0, pivotRange.r2 + 1);
			}

			// Column Subheading
			items = pivotTable.getColItems();
			if (items && styleInfo.showColHeaders) {
				start = pivotRange.c1 + firstDataCol0;
				end = pivotRange.c2 - grandColumn;
				for (j = 0; j < countCWValues; ++j) {
					if (0 === j) {
						dxf = style.firstColumnSubheading;
					} else if (1 === j % 2) {
						dxf = style.secondColumnSubheading;
					} else {
						dxf = style.thirdColumnSubheading;
					}
					dxf = dxf && dxf.dxf;
					if (dxf) {
						cells = this.getRange3(pivotRange.r1 + 1 + j, start, pivotRange.r1 + 1 + j, end);
						cells.setTableStyle(dxf);
					}
				}
				pos = pivotRange.r1 + 1 + firstHeaderRow0 - (countR ? 1 : 0);
				for (j = 0; j < items.length; ++j) {
					item = items[j];
					if (Asc.c_oAscItemType.Data !== item.t && Asc.c_oAscItemType.Grand !== item.t) {
						r = item.getR();
						if (0 === r) {
							dxf = style.firstColumnSubheading;
						} else if (1 === r % 2) {
							dxf = style.secondColumnSubheading;
						} else {
							dxf = style.thirdColumnSubheading;
						}
						dxf = dxf && dxf.dxf;
						if (dxf) {
							cells =
								this.getRange3(pivotRange.r1 + 1 + r, start + j, pos, start + j);
							cells.setTableStyle(dxf);
						}
					}
				}
			}
		}
	};
	Worksheet.prototype.updatePivotOffset = function (range, offset) {
		if (offset.row < 0 || offset.col < 0) {
			this.deletePivotTables(range);
		}
		var bboxShift = AscCommonExcel.shiftGetBBox(range, 0 !== offset.col);
		this.movePivotOffset(bboxShift, offset);
	};
	Worksheet.prototype.movePivotOffset = function (range, offset) {
		var pivotTable;
		for (var i = 0; i < this.pivotTables.length; ++i) {
			pivotTable = this.pivotTables[i];
			if (pivotTable.isInRange(range)) {
				this.workbook.oApi._changePivotSimple(pivotTable, false, false, function() {
					pivotTable.setOffset(offset, true);
				});
			}
		}
	};
	Worksheet.prototype.copyPivotTable = function (range, offset, wsTo) {
		var t = this;
		var pivotTable;
		for (var i = 0; i < this.pivotTables.length; ++i) {
			pivotTable = this.pivotTables[i];
			if (pivotTable.isInRange(range)) {
				var newPivot = pivotTable.cloneShallow();
				newPivot.prepareToPaste(wsTo, offset, wsTo === pivotTable.GetWS());
				this.workbook.oApi._changePivotSimple(newPivot, true, false, function() {
					wsTo.insertPivotTable(newPivot, true, false);
				});
			}
		}
	};
	Worksheet.prototype.inTopAutoFilter = function (range) {
		var _filterRange = this.AutoFilter && this.AutoFilter.Ref && new Asc.Range(this.AutoFilter.Ref.c1, this.AutoFilter.Ref.r1, this.AutoFilter.Ref.c2, this.AutoFilter.Ref.r1);
		return _filterRange && range.intersection(_filterRange);
	};
	Worksheet.prototype.inPivotTable = function (range, exceptPivot) {
		return this.pivotTables.find(function (element) {
			return exceptPivot !== element && element.intersection(range);
		});
	};
	Worksheet.prototype.checkShiftPivotTable = function (range, offset) {
		if ((offset.row < 0 || offset.col < 0) && this._isPivotsIntersectRangeButNotInIt(range)) {
			return true;
		}
		return this._isPivotsIntersectRangeButNotInIt(AscCommonExcel.shiftGetBBox(range, 0 !== offset.col));
	};
	Worksheet.prototype.checkMovePivotTable = function(arnFrom, arnTo, ctrlKey, opt_wsTo) {
		var t = this;
		if (!opt_wsTo) {
			opt_wsTo = this;
		}
		if (this.inPivotTable(arnFrom)) {
			var intersectionTableParts = opt_wsTo.autoFilters.getTablesIntersectionRange(arnTo);
			for (var i = 0; i < intersectionTableParts.length; i++) {
				if(intersectionTableParts[i] && intersectionTableParts[i].Ref && !arnTo.containsRange(intersectionTableParts[i].Ref)) {
					return c_oAscError.ID.PivotOverlap;
				}
			}
		}
		var res = false;
		if (ctrlKey) {
			res = this._isPivotsIntersectRangeButNotInIt(arnFrom) || opt_wsTo._isPivotsIntersectRangeButNotInIt(arnTo);
		} else {
			res = this._isPivotsIntersectRangeButNotInIt(arnFrom) || opt_wsTo.pivotTables.some(function(element) {
					return element.intersection(arnTo) && !element.isInRange(arnTo) && (opt_wsTo !== t || !element.isInRange(arnFrom));
				});
		}
		return res ? c_oAscError.ID.LockedCellPivot : c_oAscError.ID.No;
	};
	Worksheet.prototype._isPivotsIntersectRangeButNotInIt = function (bbox) {
		return this.pivotTables.some(function (element) {
			return element.intersection(bbox) && !element.isInRange(bbox);
		});
	};
	Worksheet.prototype.checkDeletePivotTables = function (range) {
		for (var i = 0; i < this.pivotTables.length; ++i) {
			if (this.pivotTables[i].intersection(range) && !this.pivotTables[i].getAllRange(this).inContains(range)) {
				return false;
			}
		}
		return true;
	};
	Worksheet.prototype.deletePivotTable = function (id) {
		for (var i = 0; i < this.pivotTables.length; ++i) {
			var pivotTable = this.pivotTables[i];
			if (id === pivotTable.Get_Id()) {
				this.clearPivotTableStyle(pivotTable);
				this.pivotTables.splice(i, 1);
				break;
			}
		}
	};
	Worksheet.prototype._deletePivotTable = function (pivotTables, pivotTable, index, withoutCells) {
		this.workbook.deleteSlicersByPivotTable(this.getId(), pivotTable.asc_getName());
		this.workbook.deleteTimelinesByPivotTable(this.getId(), pivotTable.asc_getName());

		if (!withoutCells) {
			this.clearPivotTableCell(pivotTable);
		}
		this.clearPivotTableStyle(pivotTable);
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_PivotDelete, this.getId(), null,
			new AscCommonExcel.UndoRedoData_PivotTableRedo(pivotTable.Get_Id(), pivotTable, null));
		this.pivotTables.splice(index, 1);
	};
	Worksheet.prototype.deletePivotTables = function (range, withoutCells) {
		for (var i = this.pivotTables.length - 1; i >= 0; --i) {
			var pivotTable = this.pivotTables[i];
			if (pivotTable.intersection(range)) {
				this._deletePivotTable(this.pivotTables, pivotTable, i, withoutCells);
			}
		}
		return true;
	};
	Worksheet.prototype.deletePivotTablesOnMove = function (from, to) {
		for (var i = this.pivotTables.length - 1; i >= 0; --i) {
			var pivotTable = this.pivotTables[i];
			if (pivotTable.isInRange(to) && !pivotTable.isInRange(from)) {
				this._deletePivotTable(this.pivotTables, pivotTable, i);
			}
		}
		return true;
	};
	Worksheet.prototype.getPivotTable = function (col, row) {
		var res = null;
		for (var i = 0; i < this.pivotTables.length; ++i) {
			if (this.pivotTables[i].contains(col, row)) {
				res = this.pivotTables[i];
				break;
			}
		}
		return res;
	};
	Worksheet.prototype.getPivotTablesIntersectingRange = function (bbox) {
		var res = [];
		for (var i = 0; i < this.pivotTables.length; ++i) {
			if (this.pivotTables[i].intersection(bbox)) {
				res.push(this.pivotTables[i]);
			}
		}
		return res;
	};
	Worksheet.prototype.getPivotTableByName = function (name) {
		var res = null;
		for (var i = 0; i < this.pivotTables.length; ++i) {
			if (this.pivotTables[i].asc_getName() === name) {
				res = this.pivotTables[i];
				break;
			}
		}
		return res;
	};
	Worksheet.prototype.getPivotTableById = function (id) {
		var res = null;
		for (var i = 0; i < this.pivotTables.length; ++i) {
			if (this.pivotTables[i].Get_Id() === id) {
				res = this.pivotTables[i];
				break;
			}
		}
		return res;
	};
	Worksheet.prototype.getPivotTablesByCacheId = function (pivotCacheId) {
		var res = [];
		for (var i = 0; i < this.pivotTables.length; ++i) {
			if (this.pivotTables[i].getPivotCacheId() === pivotCacheId) {
				res.push(this.pivotTables[i]);
				break;
			}
		}
		return res;
	};
	Worksheet.prototype.getPivotTablesByCache = function (cache) {
		var res = [];
		for (var i = 0; i < this.pivotTables.length; ++i) {
			if (this.pivotTables[i].cacheDefinition === cache) {
				res.push(this.pivotTables[i]);
			}
		}
		return res;
	};
	Worksheet.prototype.getPivotCacheByDataLocation = function(dataLocation) {
		return this.forEachPivotCache(undefined, function(cacheDefinition){
			if (dataLocation && dataLocation.isEqual(cacheDefinition.getDataLocation())) {
				return cacheDefinition;
			}
		});
	};

	Worksheet.prototype.getPivotCacheById = function(pivotCacheId, pivotCachesOpen) {
		return this.forEachPivotCache(pivotCachesOpen, function(cacheDefinition){
			if (pivotCacheId === cacheDefinition.getPivotCacheId()) {
				return cacheDefinition;
			}
		});
	};
	Worksheet.prototype.forEachPivotCache = function(pivotCachesOpen, callback) {
		var res, i;
		for (i = 0; i < this.pivotTables.length; ++i) {
			res = callback(this.pivotTables[i].cacheDefinition);
			if(res){
				return res;
			}
		}
		for (i = 0; i < this.aSlicers.length; ++i) {
			var cacheDefinition = this.aSlicers[i].getPivotCache();
			if (cacheDefinition) {
				res = callback(cacheDefinition);
				if(res){
					return res;
				}
			}
		}
		//for slicer with zero pivot connection
		if (pivotCachesOpen) {
			for (var cacheId in pivotCachesOpen) {
				var cacheDefinition = pivotCachesOpen[cacheId];
				if (cacheDefinition) {
					res = callback(cacheDefinition);
					if(res){
						return res;
					}
				}
			}
		}
		return null;
	};
	Worksheet.prototype.insertPivotTable = function (pivotTable, addToHistory, checkCacheDefinition) {
		pivotTable.worksheet = this;
		pivotTable.setChanged(false, true);
		if (checkCacheDefinition) {
			var cacheDefinition = this.workbook.getPivotCacheByDataLocation(pivotTable.getDataLocation());
			if (cacheDefinition) {
				pivotTable.setCacheDefinition(cacheDefinition);
			}
		}
		this.pivotTables.push(pivotTable);
		if (addToHistory) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_PivotAdd, this.getId(), null,
				new AscCommonExcel.UndoRedoData_BinaryWrapper(pivotTable));
		}
	};
	Worksheet.prototype.getPivotTableButtons = function (range) {
		var res = [];
		for (var i = 0; i < this.pivotTables.length; ++i) {
			this.pivotTables[i].getPivotTableButtons(range, res);
		}
		return res;
	};
	Worksheet.prototype.getPivotTablesClearRanges = function (range) {
		// For outline and compact pivot tables layout we need clear the grid
		var pivotTable, pivotRange, intersection, res = [];
		for (var i = 0; i < this.pivotTables.length; ++i) {
			pivotTable = this.pivotTables[i];
			if (pivotTable.clearGrid) {
				pivotRange = pivotTable.getRange();
				if (intersection = pivotRange.intersectionSimple(range)) {
					res.push(intersection);
					res.push(pivotRange);
				}
			}
		}
		return res;
	};
	Worksheet.prototype.getDisablePrompts = function () {
		return this.dataValidations && this.dataValidations.disablePrompts;
	};
	Worksheet.prototype.getDataValidation = function (c, r) {
		if (!this.dataValidations) {
			return null;
		}
		var merged = this.getMergedByCell(r, c);
		if (merged) {
			r = merged.r1;
			c = merged.c1;
		}
		for (var i = 0; i < this.dataValidations.elems.length; ++i) {
			if (this.dataValidations.elems[i].contains(c, r)) {
				return this.dataValidations.elems[i];
			}
		}
		return null;
	};
	// ----- Search -----
	Worksheet.prototype.clearFindResults = function () {
		this.lastFindOptions = null;
		this.workbook.handlers && this.workbook.handlers.trigger("clearFindResults", this.index);
	};
	Worksheet.prototype._findAllCells = function (options, searchEngine) {
		//***searchEngine
		if (options.findWhat == null) {
			options.findWhat = "";
		}
		if (true !== options.isMatchCase) {
			options.findWhat = options.findWhat.toLowerCase();
		}

		var findEmptyStr = options.findWhat === "";
		if (findEmptyStr) {
			options.findWhat = new RegExp("^$", "g");
		} else if(options.isWholeWord) {
			var length = options.findWhat.length;
			options.findWhat = '\\b' + options.findWhat + '\\b';
			options.findWhat = new RegExp(options.findWhat, "g");
			options.findWhat.length = length;
		}
		var isWholeWordTrue = null;
		if (findEmptyStr) {
			isWholeWordTrue = options.isWholeWord;
			options.isWholeWord = true;
		}

		var selectionRange = options.selectionRange || this.selectionRange || this.copySelection;
		var lastRange = selectionRange.getLast();
		var activeCell = selectionRange.activeCell;
		var merge = this.getMergedByCell(activeCell.row, activeCell.col);
		if (!searchEngine) {
			options.findInSelection = Asc.c_oAscSearchBy.Sheet === options.scanOnOnlySheet &&
				!(selectionRange.isSingleRange() && (lastRange.isOneCell() || lastRange.isEqual(merge)));
		}

		if (options.specificRange) {
			lastRange = AscCommonExcel.g_oRangeCache.getAscRange(options.specificRange);
			if (lastRange) {
				options.findInSelection = true;
			}
		}

		var findRange;
		var maxRowsCount = this.getRowsCount();
		var maxColsCount = this.getColsCount();
		var func;
		if (findEmptyStr) {
			if (maxRowsCount === 0 || maxColsCount === 0) {
				findRange = this.getRange3(0, 0, maxRowsCount, maxColsCount);
				func = findRange._foreachNoEmpty;
			} else if (options.findInSelection) {
				if (lastRange.r1 <= maxRowsCount - 1 && lastRange.c1 <= maxColsCount - 1) {
					findRange = this.getRange3(lastRange.r1, lastRange.c1, Math.min(lastRange.r2, maxRowsCount - 1), Math.min(lastRange.c2, maxColsCount - 1));
					func = findRange._foreach2;
				} else {
					findRange = this.getRange3(lastRange.r1, lastRange.c1, lastRange.r2, lastRange.c2);
					func = findRange._foreachNoEmpty;
				}
			} else {
				findRange = this.getRange3(0, 0, maxRowsCount - 1, maxColsCount - 1);
				func = findRange._foreach2;
			}
		} else {
			findRange = options.findInSelection ? this.getRange3(lastRange.r1, lastRange.c1, lastRange.r2, lastRange.c2) :
				this.getRange3(0, 0, maxRowsCount, maxColsCount);
			func = findRange._foreachNoEmpty;
		}

		if (!searchEngine && this.lastFindOptions && this.lastFindOptions.findResults && options.isEqual2(this.lastFindOptions) &&
			findRange.getBBox0().isEqual(this.lastFindOptions.findRange)) {
			return;
		}

		var oldResults = this.lastFindOptions && this.lastFindOptions.findResults.isNotEmpty();
		var result = new AscCommonExcel.findResults(), tmp;
		var emptyCell, t = this;
		func.apply(findRange, [function (cell, r, c) {
			if (cell === null) {
				if (!emptyCell) {
					emptyCell = new Cell(t);
				}
				cell = emptyCell;
			}
			if (cell && cell.isEqual(options)) {
				if (!options.scanByRows) {
					tmp = r;
					r = c;
					c = tmp;
				}
				searchEngine ? searchEngine.Add(r, c, cell, !options.scanByRows ? result : null, options) : result.add(r, c, cell);
			}
		}]);
		!options.scanByRows && searchEngine && searchEngine.endAdd(result);

		if (isWholeWordTrue !== null) {
			options.isWholeWord = isWholeWordTrue;
		}
		if (findEmptyStr) {
			options.findWhat = "";
		}
		this.lastFindOptions = options.clone();
		// ToDo support multiselect
		this.lastFindOptions.findRange = findRange.getBBox0().clone();
		this.lastFindOptions.findResults = result;

		if (!options.isReplaceAll && this.workbook.oApi.selectSearchingResults && (oldResults || result.isNotEmpty() || (searchEngine && (searchEngine._lastNotEmpty || searchEngine.isNotEmpty()))) &&
			this === this.workbook.getActiveWs()) {
			this.workbook.handlers && this.workbook.handlers.trigger("drawWS");
			if (searchEngine) {
				searchEngine._lastNotEmpty = null;
			}
		}
	};
	Worksheet.prototype.findCellText = function (options, searchEngine) {
		this._findAllCells(options, searchEngine);
		if (searchEngine) {
			return;
		}
		//CDocumentSearchExcel.prototype.SetCurrent

		var selectionRange = options.selectionRange || this.selectionRange;
		var activeCell = selectionRange.activeCell;

		var tmp, key1 = activeCell.row, key2 = activeCell.col;
		if (!options.scanByRows) {
			tmp = key1;
			key1 = key2;
			key2 = tmp;
		}

		var result = null;
		var findResults = this.lastFindOptions.findResults;
		if (findResults.find(key1, key2, options.scanForward)) {
			key1 = findResults.currentKey1;
			key2 = findResults.currentKey2;
			if (!options.scanByRows) {
				tmp = key1;
				key1 = key2;
				key2 = tmp;
			}
			result = new AscCommon.CellBase(key1, key2);
		}
		return result;
	};
	Worksheet.prototype.inFindResults = function (row, col) {
		var tmp, res = false;
		var findResults = this.lastFindOptions && this.lastFindOptions.findResults;
		if (findResults) {
			if (!this.lastFindOptions.scanByRows) {
				tmp = col;
				col = row;
				row = tmp;
			}
			res = findResults.contains(row, col);
		}
		return res;
	};
	Worksheet.prototype.excludeHiddenRows = function (bExclude) {
		this.bExcludeHiddenRows = bExclude;
	};
	Worksheet.prototype.ignoreWriteFormulas = function (val) {
		this.bIgnoreWriteFormulas = val;
	};
	Worksheet.prototype.checkShiftArrayFormulas = function (range, offset) {
		//проверка на частичный сдвиг формулы массива
		//проверка на внутренний сдвиг
		var res = true;

		var isHor = offset && offset.col;
		var isDelete = offset && (offset.col < 0 || offset.row < 0);

		var checkRange = function(formulaRange) {
			//частичное выделение при удалении столбца/строки
			if(isDelete && formulaRange.intersection(range) && !range.containsRange(formulaRange)) {
				return false;
			}

			if (isHor) {
				if(range.r1 > formulaRange.r1 && range.r1 <= formulaRange.r2) {
					return false;
				}
				if(range.r2 >= formulaRange.r1 && range.r2 < formulaRange.r2) {
					return false;
				}
				if(range.c1 > formulaRange.c1 && range.c1 <= formulaRange.c2) {
					return false;
				}
				/*if(range.c1 < formulaRange.c1 && range.c2 >= formulaRange.c1 && range.c2 < formulaRange.c2) {
					return false;
				}*/
			} else {
				if(range.c1 > formulaRange.c1 && range.c1 <= formulaRange.c2) {
					return false;
				}
				if(range.c2 >= formulaRange.c1 && range.c2 < formulaRange.c2) {
					return false;
				}
				if(range.r1 > formulaRange.r1 && range.r1 <= formulaRange.r2) {
					return false;
				}
				/*if(range.r1 < formulaRange.r1 && range.r2 >= formulaRange.r1 && range.r2 < formulaRange.r2) {
					return false;
				}*/
			}

			return true;
		};

		//if intersection with range
		var alreadyCheckFormulas = [];
		var r2 = range.r2, c2 = range.c2;
		if(isHor) {
			c2 = gc_nMaxCol0;
		} else {
			r2 = gc_nMaxRow0;
		}

		this.getRange3(range.r1, range.c1, r2, c2)._foreachNoEmpty(function(cell) {
			if(res && cell.isFormula()) {
				var formulaParsed = cell.getFormulaParsed();
				var arrayFormulaRef = formulaParsed.getArrayFormulaRef();
				if(arrayFormulaRef && !alreadyCheckFormulas[formulaParsed._index]) {
					if(!checkRange(arrayFormulaRef)) {
						res = false;
					}
					alreadyCheckFormulas[formulaParsed._index] = 1;
				}
			}
		});

		return res;
	};
	Worksheet.prototype.setRowsCount = function (val) {
		if(val > gc_nMaxRow0 || val < 0) {
			return;
		}
		this.nRowsCount = val;
	};
	Worksheet.prototype.reinitRowsCount = function () {
		let maxTableRow = 0;
		if (this.TableParts) {
			for (let i = 0; i < this.TableParts.length; i++) {
				if (this.TableParts[i] && this.TableParts[i].Ref && this.TableParts[i].Ref.r2 > maxTableRow) {
					maxTableRow = this.TableParts[i].Ref.r2;
				}
			}
		}
		//TODO не учитываются настройки для всей строки
		//по this.rowsData.indexB ориентироваться не могу, поскольку при undo в большинстве случаев он остаётся неизменным
		this.nRowsCount = Math.max(maxTableRow, this.cellsByColRowsCount/*, this.rowsData && this.rowsData.indexB ? this.rowsData.indexB : 0*/);
	};
	Worksheet.prototype.fromXLSB = function(stream, type, tmp, aCellXfs, aSharedStrings, fInitCellAfterRead) {
		stream.XlsbSkipRecord();//XLSB::rt_BEGIN_SHEET_DATA

		var oldPos = -1;
		while (oldPos !== stream.GetCurPos()) {
			oldPos = stream.GetCurPos();
			type = stream.XlsbReadRecordType();
			if (AscCommonExcel.XLSB.rt_CELL_BLANK <= type && type <= AscCommonExcel.XLSB.rt_FMLA_ERROR) {
				tmp.cell.clear();
				tmp.formula.clean();
				tmp.cell.fromXLSB(stream, type, tmp.row.index, aCellXfs, aSharedStrings, tmp);
				fInitCellAfterRead(tmp);
			}
			else if (AscCommonExcel.XLSB.rt_ROW_HDR === type) {
				tmp.row.clear();
				tmp.row.fromXLSB(stream, aCellXfs);
				tmp.row.saveContent();
				tmp.ws.cellsByColRowsCount = Math.max(tmp.ws.cellsByColRowsCount, tmp.row.index + 1);
				tmp.ws.nRowsCount = Math.max(tmp.ws.nRowsCount, tmp.ws.cellsByColRowsCount);
			}
			else if (AscCommonExcel.XLSB.rt_END_SHEET_DATA === type) {
				stream.XlsbSkipRecord();
				break;
			}
			else {
				stream.XlsbSkipRecord();
			}
		}
	};

	//need recalculate formulas after change rows
	Worksheet.prototype.needRecalFormulas = function(start, stop) {
		//TODO в данном случае необходим пересчёт только тез формул, которые зависят от данных строк + те, которые
		// меняют своё значение в зависимости от скрытия/раскрыватия строк
		var res = false;

		if(this.AutoFilter && this.AutoFilter.isApplyAutoFilter()) {
			return true;
		}

		var tableParts = this.TableParts;
		var tablePart;
		for (var i = 0; i < tableParts.length; i++) {
			tablePart = tableParts[i];
			if (tablePart && tablePart.Ref && start >= tablePart.Ref.r1 && stop <= tablePart.Ref.r2) {
				res = true;
				break;
			}
		}

		return res;
	};

	Worksheet.prototype.getRowColColors = function (columnRange, byRow, notCheckOneColor) {
		var ws = this;
		var res = {text: true, colors: [], fontColors: [], date: false};
		var alreadyAddColors = {}, alreadyAddFontColors = {};

		var getAscColor = function (color) {
			var ascColor = new Asc.asc_CColor();
			ascColor.asc_putR(color.getR());
			ascColor.asc_putG(color.getG());
			ascColor.asc_putB(color.getB());
			ascColor.asc_putA(color.getA());

			return ascColor;
		};

		var addFontColorsToArray = function (fontColor) {
			var rgb = fontColor && null !== fontColor  ? fontColor.getRgb() : null;
			if(rgb === 0) {
				rgb = null;
			}
			var isDefaultFontColor = null === rgb;

			if (true !== alreadyAddFontColors[rgb]) {
				if (isDefaultFontColor) {
					res.fontColors.push(null);
					alreadyAddFontColors[null] = true;
				} else {
					var ascFontColor = getAscColor(fontColor);
					res.fontColors.push(ascFontColor);
					alreadyAddFontColors[rgb] = true;
				}
			}
		};

		var addCellColorsToArray = function (color) {
			var rgb = null !== color && color.fill && color.fill.bg() ? color.fill.bg().getRgb() : null;
			var isDefaultCellColor = null === rgb;

			if (true !== alreadyAddColors[rgb]) {
				if (isDefaultCellColor) {
					res.colors.push(null);
					alreadyAddColors[null] = true;
				} else {
					var ascColor = getAscColor(color.fill.bg());
					res.colors.push(ascColor);
					alreadyAddColors[rgb] = true;
				}
			}
		};

		//TODO automaticRange ?
		var tempText = 0, tempDigit = 0, tempDate = 0;
		var r1 = byRow ? columnRange.r1 : columnRange.r1;
		var r2 = byRow ? columnRange.r1 : columnRange.r2;
		var c1 = byRow ? columnRange.c1 : columnRange.c1;
		var c2 = byRow ? columnRange.c2 : columnRange.c1;
		ws.getRange3(r1, c1, r2, c2)._foreachNoEmpty(function(cell) {
			//добавляем без цвета ячейку
			if (!cell) {
				if (true !== alreadyAddColors[null]) {
					alreadyAddColors[null] = true;
					res.colors.push(null);
				}
				return;
			}

			if (false === cell.isNullText()) {
				var type = cell.getType();

				if (type === window["AscCommon"].CellValueType.Number) {
					if (cell.getNumFormat().isDateTimeFormat()) {
						tempDate++;
					} else {
						tempDigit++;
					}
				} else {
					tempText++;
				}
			}

			//font colors
			var multiText = cell.getValueMultiText();
			var fontColor = null;
			var xfs = cell.getCompiledStyleCustom(false, true, true);
			if (null !== multiText) {
				for (var j = 0; j < multiText.length; j++) {
					fontColor = multiText[j].format ? multiText[j].format.getColor() : null;
					if(null !== fontColor) {
						addFontColorsToArray(fontColor);
					} else {
						fontColor = xfs && xfs.font ? xfs.font.getColor() : null;
						addFontColorsToArray(fontColor);
					}
				}
			} else {
				fontColor = xfs && xfs.font ? xfs.font.getColor() : null;
				addFontColorsToArray(fontColor);
			}

			//cell colors
			addCellColorsToArray(xfs);
		});

		//если один элемент в массиве, не отправляем его в меню
		if (res.colors.length === 1 && !notCheckOneColor) {
			res.colors = [];
		}
		if (res.fontColors.length === 1 && !notCheckOneColor) {
			res.fontColors = [];
		}

		if (tempDate > tempDigit && tempDate > tempText) {
			res.date = true;
			res.text = false;
		} else {
			res.text = tempDigit <= tempText;
		}

		return res;
	};

	Worksheet.prototype.updateSortStateOffset = function (range, offset) {
		if(!this.sortState) {
			return;
		}
		var bAlreadyDel = false;
		if (offset.row < 0 || offset.col < 0) {
			//смотрим, не попал ли в выделение целиком
			bAlreadyDel = this.deleteSortState(range);
		}
		if(!bAlreadyDel) {
			var bboxShift = AscCommonExcel.shiftGetBBox(range, 0 !== offset.col);
			this.sortState.shift(range, offset, this, true);
			this.moveSortState(bboxShift, offset);
		}
	};

	Worksheet.prototype.deleteSortState = function (range) {
		if(this.sortState && range.containsRange(this.sortState.Ref)) {
			this._deleteSortState();
			return true;
		}
		return false;
	};

	Worksheet.prototype._deleteSortState = function () {
		var oldSortState = this.sortState.clone();
		this.sortState = null;
		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoSortState, AscCH.historyitem_SortState_Add, this.getId(), null,
			new AscCommonExcel.UndoRedoData_SortState(oldSortState, null));
		return true;
	};

	Worksheet.prototype.moveSortState = function (range, offset) {
		if(this.sortState && range.containsRange(this.sortState.Ref)) {
			this.sortState.setOffset(offset, this, true);
			return true;
		}
		return false;
	};

	Worksheet.prototype.getDrawingDocument = function() {
		if(this.workbook) {
			return this.workbook.getDrawingDocument();
		}
		return null;
	};
	Worksheet.prototype.getActiveFunctionInfo = function (parser, parserResult, argNum, type, doNotCalcArgs) {
		var t = this;
		var calculateFormula = function (str) {
			var _res = null;
			if (str !== "") {
				var _formulaParsedArg = new AscCommonExcel.parserFormula(str, /*formulaParsed.parent*/null, t);
				var _parseResultArg = new AscCommonExcel.ParseResult([], []);
				_formulaParsedArg.parse(true, true, _parseResultArg, true);
				if (!_parseResultArg.error) {
					_res = _formulaParsedArg.calculate();
				}
			}
			return _res;
		};

		var convertFormulaResultByType = function (_res) {
			if (type === undefined || type === null) {
				return _res.toLocaleString();
			}

			//TODO если полная проверка, то выводим ошибки - если нет, то вовзращаем пустую строку
			var result = "";
			if (type === Asc.c_oAscFormulaArgumentType.number) {
				_res = _res.tocNumber();
				if (_res) {
					result = _res.toLocaleString();
				}
			}
			return result;
		};

		var _formulaParsed, _parseResult, valueForEdit;
		if (!parser) {
			var activeCell = this.selectionRange.activeCell;
			var formulaParsed;
			this.getCell3(activeCell.row, activeCell.col)._foreachNoEmpty(function (cell) {
				if (cell.isFormula()) {
					formulaParsed = cell.getFormulaParsed();
					if (formulaParsed) {
						valueForEdit = cell.getValueForEdit();
					}
				}
			});
			_formulaParsed = new AscCommonExcel.parserFormula(valueForEdit.substr(1), /*formulaParsed.parent*/null, this);
			_parseResult = new AscCommonExcel.ParseResult([], []);
			_formulaParsed.parse(true, true, _parseResult, true);
		} else {
			_formulaParsed = parser;
			_parseResult = parserResult;
			valueForEdit = "=" + parser.Formula;
		}

		var res, str, calcRes;
		if (_formulaParsed && _parseResult.activeFunction && _parseResult.activeFunction.func) {
			res = new AscCommonExcel.CFunctionInfo(_parseResult.activeFunction.func.name);
			if (!_parseResult.error) {
				var _parent = _formulaParsed.parent;
				_formulaParsed.parent = null;
				res.formulaResult = _formulaParsed.calculate().toLocaleString();
				_formulaParsed.parent = _parent;
			}

			//asc_getFunctionResult
			str = valueForEdit.substring(_parseResult.activeFunction.start + 1, _parseResult.activeFunction.end + 1);
			calcRes = calculateFormula(str);
			if (calcRes) {
				res.functionResult = calcRes.toLocaleString();
			}

			res._cursorPos = _parseResult.cursorPos + 1;
			var argPosArr = _parseResult.argPosArr;
			if (argPosArr && argPosArr.length && true !== doNotCalcArgs){
				for (var i = 0; i < argPosArr.length; i++) {
					if (!res.argumentsValue) {
						res.argumentsValue = [];
					}
					str = valueForEdit.substring(argPosArr[i].start, argPosArr[i].end);
					res.argumentsValue.push(str);
					if (str !== "") {
						if (!res.argumentsResult) {
							res.argumentsResult = [];
						}
						calcRes = calculateFormula(str);
						if (calcRes) {
							res.argumentsResult[i] = i === argNum ? convertFormulaResultByType(calcRes) : calcRes.toLocaleString();
						}
					}
				}
			}
		}

		return res;
	};

	Worksheet.prototype.isActiveCellFormula = function () {
		var activeCell = this.selectionRange.activeCell;
		var res;
		this.getCell3(activeCell.row, activeCell.col)._foreachNoEmpty(function (cell) {
			if (cell.isFormula()) {
				res = true;
			}
		});
		return res;
	};

	Worksheet.prototype.calculateWizardFormula = function (formula, type) {
		let res = null, resultStr = null;
		if (formula) {
			let parser = new AscCommonExcel.parserFormula(formula, /*formulaParsed.parent*/null, this);
			let parseResultArg = new AscCommonExcel.ParseResult([], []);
			parser.parse(true, true, parseResultArg, true);
			if (!parseResultArg.error) {
				res = parser.calculate();
			}

			resultStr = "";
			if (res) {
				const maxArrayRowCount = 20;
				const maxArrayColCount = 20;
				//TODO рассчеты аргументов зависят от конкретных функций
				//допустим, sum и acos - типа аргумента number, но результат для cellsRange3D разный

				if (res.type === AscCommonExcel.cElementType.cell || res.type === AscCommonExcel.cElementType.cell3D) {
					res = res.getValue();
				}

				//TODO если полная проверка, то выводим ошибки - если нет, то вовзращаем пустую строку
				if (type === undefined || type === null || res.type === AscCommonExcel.cElementType.error) {
					return {str: res.toLocaleString(), obj: res};
				}

				if (type === Asc.c_oAscFormulaArgumentType.number) {
					//in most cases in ms array/range -> {1,2,..} if NUMBER type
					if (res.type === AscCommonExcel.cElementType.array) {
						resultStr = res.toLocaleString();
					} else if (res.type === AscCommonExcel.cElementType.cellsRange || res.type === AscCommonExcel.cElementType.cellsRange3D) {
						res = res.getFullArray(new AscCommonExcel.cNumber(0), maxArrayRowCount, maxArrayColCount);
						if (res) {
							resultStr = res.toLocaleString();
						}
					} else {
						res = res.tocNumber();
						if (res) {
							resultStr = res.toLocaleString();
						}
					}
				} else if (type === Asc.c_oAscFormulaArgumentType.text) {
					if (res.type === AscCommonExcel.cElementType.array) {
						res = res.getElementRowCol(0, 0);
						res = res.tocString();
						if (res) {
							resultStr = '"' + res.toLocaleString() + '"';
						}
					} else if (res.type === AscCommonExcel.cElementType.cellsRange) {
						resultStr = res.toLocaleString();
					} else if (res.type !== AscCommonExcel.cElementType.cellsRange3D) {
						res = res.tocString();
						if (res) {
							resultStr = '"' + res.toLocaleString() + '"';
						}
					}
				} else if (type === Asc.c_oAscFormulaArgumentType.logical) {
					if (res.type === AscCommonExcel.cElementType.cellsRange || res.type === AscCommonExcel.cElementType.cellsRange3D) {
						res = res.getFullArray(new AscCommonExcel.cNumber(0), maxArrayRowCount, maxArrayColCount);
					} else if (res.type !== AscCommonExcel.cElementType.array) {
						res = res.tocBool();
					}
					if (res) {
						resultStr = res.toLocaleString();
					}
				} else if (type === Asc.c_oAscFormulaArgumentType.any) {
					if (res.type === AscCommonExcel.cElementType.array) {
						res = res.getElementRowCol(0, 0);
						res = res.tocString();
						if (res) {
							resultStr = res.toLocaleString();
						}
					} else if (res.type === AscCommonExcel.cElementType.cellsRange) {
						resultStr = res.toLocaleString();
					} else if (res.type === AscCommonExcel.cElementType.cellsRange3D) {
						resultStr = res.toLocaleString();
					} else {
						res = res.tocString();
						if (res) {
							resultStr = '"' + res.toLocaleString() + '"';
						}
					}
				} else if (type === Asc.c_oAscFormulaArgumentType.reference) {
					if (res.type === AscCommonExcel.cElementType.array) {
						resultStr = res.toLocaleString();
					} else if (res.type === AscCommonExcel.cElementType.cellsRange || res.type === AscCommonExcel.cElementType.cellsRange3D) {
						res = res.getFullArray(new AscCommonExcel.cNumber(0), maxArrayRowCount, maxArrayColCount);
						if (res) {
							resultStr = res.toLocaleString();
						}
					} else {
						resultStr = res.toLocaleString();
					}
				}
			}
		}
		return {str: resultStr, obj: res};
	};


	Worksheet.prototype.insertSlicer = function (name, obj_name, type, pivotTable, slicerCacheDefinition) {
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		//TODO недостаточно ли вместо всей данной длинной структуры использовать только tableId(name) и columnName?
		var slicer = new window['Asc'].CT_slicer(this);
		slicer.cacheDefinition = slicerCacheDefinition || null;
		var isNewCache = slicer.init(name, obj_name, type, undefined, pivotTable);
		this.aSlicers.push(slicer);
		var oCache = slicer.getCacheDefinition();

		AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SlicerAdd, this.getId(), null,
			new AscCommonExcel.UndoRedoData_FromTo(null, slicer));

		if (isNewCache) {
			slicer.updateItemsWithNoData();
		}

		if (oCache) {
			var _name = oCache.name;
			var newDefName = new Asc.asc_CDefName(_name, "#N/A", null, Asc.c_oAscDefNameType.slicer);
			this.workbook.editDefinesNames(null, newDefName);
		}

		AscCommon.History.EndTransaction();
		return slicer;
	};

	Worksheet.prototype.deleteSlicer = function (name, doDelDefName) {
		var res = false;

		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		var slicerObj = this.getSlicerIndexByName(name);
		if (null !== slicerObj) {
			this.aSlicers.splice(slicerObj.index, 1);
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SlicerDelete, this.getId(), null,
				new AscCommonExcel.UndoRedoData_FromTo(slicerObj.obj, null));
			this.workbook.onSlicerDelete(name);
			res = true;
			if ((!this.workbook.bUndoChanges && !this.workbook.bRedoChanges) || doDelDefName)
			{
				var cache = slicerObj.obj.getCacheDefinition();
				//удаляем именованный диапазон только если на данный кэш уже никто не ссылается
				if (cache && null === this.workbook.getSlicersByCacheName(cache.name)) {
					var defName = this.workbook.getDefinesNames(cache.name);
					if (defName) {
						this.workbook.delDefinesNames(defName.getAscCDefName());
					}
				}
			}
		}

		AscCommon.History.EndTransaction();
		return res;
	};

	Worksheet.prototype.getSlicerCachesBySourceName = function (name) {
		var res = null
		for (var i = 0; i < this.aSlicers.length; i++) {
			var cache = this.aSlicers[i] && this.aSlicers[i].getSlicerCache();
			if (cache && name === cache.sourceName) {
				if (!res) {
					res = [];
				}
				res.push(cache);
			}
		}

		return res;
	};

	Worksheet.prototype.getSlicersByCaption = function (name) {
		var res = [];

		for (var i = 0; i < this.aSlicers.length; i++) {
			if (this.aSlicers[i] && name === this.aSlicers[i].caption) {
				res.push({obj: this.aSlicers[i], index: i});
			}
		}

		return res.length ? res : null;
	};

	Worksheet.prototype.getSlicersByCacheName = function (name) {
		var res = [];

		for (var i = 0; i < this.aSlicers.length; i++) {
			var cache = this.aSlicers[i] && this.aSlicers[i].getSlicerCache();
			if (cache && name === cache.name) {
				res.push(this.aSlicers[i]);
			}
		}

		return res.length ? res : null;
	};

	Worksheet.prototype.getSlicerCacheByCacheName = function (name) {
		for (var i = 0; i < this.aSlicers.length; i++) {
			var cache = this.aSlicers[i] && this.aSlicers[i].getSlicerCache();
			if (cache && name === cache.name) {
				return cache;
			}
		}

		return null;
	};

	Worksheet.prototype.getSlicerByName = function (name) {
		var res = null;

		for (var i = 0; i < this.aSlicers.length; i++) {
			if (this.aSlicers[i] && name === this.aSlicers[i].name) {
				res = this.aSlicers[i];
				break;
			}
		}

		return res;
	};

	Worksheet.prototype.getSlicerViewByName = function (name) {
		var res = null;
		for (var i = 0; i < this.Drawings.length; i++) {
			res = this.Drawings[i].getSlicerViewByName(name);
			if(res) {
				break;
			}
		}
		return res;
	};

	Worksheet.prototype.getSlicerIndexByName = function (name) {
		var res = null;

		for (var i = 0; i < this.aSlicers.length; i++) {
			if (this.aSlicers[i] && name === this.aSlicers[i].name) {
				res = {obj: this.aSlicers[i], index: i};
				break;
			}
		}

		return res;
	};

	Worksheet.prototype.getSlicerCacheByName = function (name) {
		for (var i = 0; i < this.aSlicers.length; i++) {
			var cache = this.aSlicers[i] && this.aSlicers[i].getSlicerCache();
			if (cache && name === cache.name) {
				return cache;
			}
		}

		return null;
	};

	Worksheet.prototype.getSlicersByTableName = function (val) {
		var res = [];
		for (var i = 0; i < this.aSlicers.length; i++) {
			//пока сделал только для форматированных таблиц
			var tableSlicerCache = this.aSlicers[i] && this.aSlicers[i].getTableSlicerCache();
			if (tableSlicerCache && tableSlicerCache.tableId === val) {
				res.push(this.aSlicers[i]);
			}
		}
		return res.length ? res : null;
	};

	Worksheet.prototype.getSlicersByTableColName = function (tableName, colName) {
		var res = [];
		for (var i = 0; i < this.aSlicers.length; i++) {
			//пока сделал только для форматированных таблиц
			var tableSlicerCache = this.aSlicers[i] && this.aSlicers[i].getTableSlicerCache();
			if (tableSlicerCache && tableSlicerCache.tableId === tableName && tableSlicerCache.column === colName) {
				res.push(this.aSlicers[i]);
			}
		}
		return res.length ? res : null;
	};
	Worksheet.prototype.handleDrawings = function (fCallback) {
		for(var nIndex = 0; nIndex < this.Drawings.length; ++nIndex) {
			this.Drawings[nIndex].handleObject(fCallback);
		}
	};

	Worksheet.prototype.changeTableColName = function (tableName, oldVal, newVal) {
		if (this.workbook.bUndoChanges || this.workbook.bRedoChanges) {
			return;
		}

		var slicers = this.getSlicersByTableColName(tableName, oldVal);
		if (slicers) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();

			for (var i = 0; i < slicers.length; i++) {
				slicers[i].setTableColName(oldVal, newVal);
			}

			AscCommon.History.EndTransaction();
		}
	};

	Worksheet.prototype.changeTableName = function (oldVal, newVal) {
		if (this.workbook.bUndoChanges || this.workbook.bRedoChanges) {
			return;
		}

		var slicers = this.workbook.getSlicersByTableName(oldVal);
		if (slicers) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();

			for (var i = 0; i < slicers.length; i++) {
				slicers[i].setTableName(newVal);
			}

			AscCommon.History.EndTransaction();
		}
	};

	Worksheet.prototype.setSlicerTableName = function (tableName, newTableName) {
		//TODO history
		for (var i = 0; i < this.aSlicers.length; i++) {
			var tableSlicerCache = this.aSlicers[i].getTableSlicerCache();
			if (tableSlicerCache && tableSlicerCache.tableId === tableName) {
				tableSlicerCache.setTableName(newTableName);
			}
		}
	};
	Worksheet.prototype.onSlicerUpdate = function (sName) {
		var bRet = false;
		for(var i = 0; i < this.Drawings.length; ++i) {
			bRet = bRet || this.Drawings[i].onSlicerUpdate(sName);
		}
		return bRet;
	};

	Worksheet.prototype.onSlicerDelete = function (sName) {
		var bRet = false;
		for(var i = 0; i < this.Drawings.length; ++i) {
			bRet = bRet || this.Drawings[i].onSlicerDelete(sName);
		}
		return bRet;
	};
	Worksheet.prototype.onSlicerLock = function (sName, bLock) {
		for(var i = 0; i < this.Drawings.length; ++i) {
			this.Drawings[i].onSlicerLock(sName, bLock);
		}
	};
	Worksheet.prototype.onSlicerChangeName = function (sName, sNewName) {
		for(var i = 0; i < this.Drawings.length; ++i) {
			this.Drawings[i].onSlicerChangeName(sName, sNewName);
		}
	};
	Worksheet.prototype.onTimeSlicerDelete = function (sName) {
		for(var i = 0; i < this.Drawings.length; ++i) {
			this.Drawings[i].onTimeSlicerDelete(sName);
		}
	};

	Worksheet.prototype.deleteSlicersByTable = function (tableName) {
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		var slicers = this.workbook.getSlicersByTableName(tableName);
		if (slicers) {
			for (var i = 0; i < slicers.length; i++) {
				this.deleteSlicer(slicers[i].name);
			}
		}

		AscCommon.History.EndTransaction();
	};

	Worksheet.prototype.deleteSlicersByTableCol = function (tableName, colMap) {
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		for (var j in colMap) {
			var slicers = this.getSlicersByTableColName(tableName, j);
			if (slicers) {
				for (var i = 0; i < slicers.length; i++) {
					this.deleteSlicer(slicers[i].name);
				}
			}
		}

		AscCommon.History.EndTransaction();
	};

	Worksheet.prototype.changeSlicerCacheName = function (oldVal, newVal) {
		if (this.workbook.bUndoChanges || this.workbook.bRedoChanges) {
			return;
		}

		var slicers = this.getSlicersByCacheName(oldVal);
		if (slicers) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();

			for (var i = 0; i < slicers.length; i++) {
				slicers[i].setCacheName(newVal);
			}

			AscCommon.History.EndTransaction();
		}
	};

	Worksheet.prototype.checkChangeTablesContent = function (arn) {
		this.autoFilters.renameTableColumn(arn);
		var tables = this.autoFilters.getTablesIntersectionRange(arn);
		if (tables) {
			for (var i = 0; i < tables.length; i++) {
				this.workbook.slicersUpdateAfterChangeTable(tables[i].DisplayName);
			}
		}
	};

	Worksheet.prototype.updateSlicersByRange = function (range) {
		var tables = this.autoFilters.getTablesIntersectionRange(range);
		if (tables) {
			for (var i = 0; i < tables.length; i++) {
				this.workbook.slicersUpdateAfterChangeTable(tables[i].DisplayName);
			}
		}
		var pivot = false;
		if (pivot) {

		}
	};


	Worksheet.prototype.getActiveNamedSheetViewId = function () {
		return this.activeNamedSheetViewId;
	};

	Worksheet.prototype.setActiveNamedSheetView = function (id) {
		this.activeNamedSheetViewId = id;
	};

	Worksheet.prototype.getNvsFilterByTableName = function (val, opt_name, viewId) {
		var activeNamedSheetViewId;
		if (viewId) {
			activeNamedSheetViewId = viewId;
		} else {
			activeNamedSheetViewId = opt_name ? this.getIdNamedSheetViewByName(opt_name) : this.getActiveNamedSheetViewId();
		}

		if (activeNamedSheetViewId === null) {
			return;
		}

		if (Asc.CT_NamedSheetView.prototype.getNsvFiltersByTableId) {
			var sheetView = this.getNamedSheetViewById(activeNamedSheetViewId);
			return sheetView ? sheetView.getNsvFiltersByTableId(val) : null;
		}

		return null;
	};

	Worksheet.prototype.getNamedSheetViewFilterColumns = function (val) {
		var nsvFilter = this.getNvsFilterByTableName(val);
		if (nsvFilter && nsvFilter.columnsFilter) {
			return nsvFilter.columnsFilter;
		}
		return null;
	};

	Worksheet.prototype.addNamedSheetView = function (sheetView, isDuplicate) {

		if (!sheetView) {
			sheetView = new Asc.CT_NamedSheetView();
			sheetView.name = sheetView.generateName();
		}

		var _filter;
		if (!isDuplicate) {
			for (var i = 0; i < this.TableParts.length; i++) {
				_filter = new Asc.CT_NsvFilter();
				_filter.init(this.TableParts[i]);
				sheetView.nsvFilters.push(_filter);
			}
			if (this.AutoFilter) {
				_filter = new Asc.CT_NsvFilter();
				_filter.init(this.AutoFilter);
				sheetView.nsvFilters.push(_filter);
			}
		}

		this.insertNamedSheetView(sheetView, true);
	};

	Worksheet.prototype.deleteNamedSheetViews = function (arr, doNotAddHistory) {
		var namedSheetViews = this.aNamedSheetViews;
		if (namedSheetViews && arr) {
			var diff = 0;
			for (var i = 0; i < arr.length; i++) {
				if (!arr[i]) {
					continue;
				}
				var index = this.getIndexNamedSheetViewByName(arr[i].name);
				var namedSheetView = this.getNamedSheetViewByName(arr[i].name);
				if (namedSheetView.Id === this.getActiveNamedSheetViewId()) {
					if (this.workbook.oApi.asc_setActiveNamedSheetView) {
						this.workbook.oApi.asc_setActiveNamedSheetView(null);
					}
				}

				if (!doNotAddHistory) {
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SheetViewDelete, this.getId(), null,
						new AscCommonExcel.UndoRedoData_NamedSheetViewRedo(namedSheetView.Get_Id(), namedSheetView, null));
				}

				namedSheetViews.splice(index - diff, 1);
				diff++;
			}
		}
	};

	Worksheet.prototype.insertNamedSheetView = function (sheetView, addToHistory) {
		sheetView.ws = this;
		this.aNamedSheetViews.push(sheetView);
		if (addToHistory) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SheetViewAdd, this.getId(), null,
				new AscCommonExcel.UndoRedoData_BinaryWrapper(sheetView));
		}
	};

	Worksheet.prototype.getNamedSheetViews = function () {
		var namedSheetViews = this.aNamedSheetViews;
		if (namedSheetViews) {
			var res = [], ascSheetView;
			for (var i = 0; i < namedSheetViews.length; i++) {
				ascSheetView = namedSheetViews[i];
				ascSheetView._isActive = ascSheetView.Id === this.getActiveNamedSheetViewId();
				res.push(ascSheetView);
			}
			return res;
		}
		return null;
	};

	Worksheet.prototype.getNamedSheetViewByName = function (name) {
		var namedSheetViews = this.aNamedSheetViews;
		if (namedSheetViews) {
			for (var i = 0; i < namedSheetViews.length; i++) {
				if (name === namedSheetViews[i].name) {
					return namedSheetViews[i];
				}
			}
		}
		return null;
	};

	Worksheet.prototype.getNamedSheetViewById = function (id) {
		var namedSheetViews = this.aNamedSheetViews;
		if (namedSheetViews) {
			for (var i = 0; i < namedSheetViews.length; i++) {
				if (id === namedSheetViews[i].Get_Id()) {
					return namedSheetViews[i];
				}
			}
		}
		return null;
	};

	Worksheet.prototype.getIndexNamedSheetViewByName = function (name) {
		var namedSheetViews = this.aNamedSheetViews;
		if (namedSheetViews) {
			for (var i = 0; i < namedSheetViews.length; i++) {
				if (name === namedSheetViews[i].name) {
					return i;
				}
			}
		}
		return null;
	};

	Worksheet.prototype.getIdNamedSheetViewByName = function (name) {
		var namedSheetViews = this.aNamedSheetViews;
		if (namedSheetViews) {
			for (var i = 0; i < namedSheetViews.length; i++) {
				if (name === namedSheetViews[i].name) {
					return namedSheetViews[i].Id;
				}
			}
		}
		return null;
	};

	Worksheet.prototype.forEachView = function (actionView) {
		var namedSheetViews = this.aNamedSheetViews;
		if (namedSheetViews) {
			var activeId = this.getActiveNamedSheetViewId();
			for (var i = 0; i < namedSheetViews.length; i++) {
				actionView(namedSheetViews[i], activeId === namedSheetViews[i].Id);
			}
		}
	};

	Worksheet.prototype.checkCorrectTables = function () {
		for (let i = 0; i < this.TableParts.length; ++i) {
			let table = this.TableParts[i];
			if (table.isHeaderRow()) {

				this.workbook.dependencyFormulas.lockRecal();
				for (let j = 0; j < table.TableColumns.length; j++) {
					let tableColName = table.TableColumns[j].getTableColumnName();
					this._getCell(table.Ref.r1, table.Ref.c1 + j, function(cell) {
						let valueData = cell.getValueData();
						let val = valueData && valueData.value && valueData.value.text;
						if (val !== tableColName){
							cell.setValueData(
								new AscCommonExcel.UndoRedoData_CellValueData(null, new AscCommonExcel.CCellValue({
									text: tableColName,
									type: CellValueType.String
								})));
						}
					});
				}
				this.workbook.dependencyFormulas.unlockRecal(true);

				if (table.TableColumns.length < (table.Ref.c2 - table.Ref.c1 + 1)) {
					table.Ref.c2 = table.Ref.c1 + table.TableColumns.length - 1;
				}
			}
		}
	};

	Worksheet.prototype.getDataValidationProps = function (doExtend) {
		var _selection = this.getSelection();
		
		if (!this.dataValidations) {
			var newDataValidation = new window['AscCommonExcel'].CDataValidation();
			newDataValidation.showErrorMessage = true;
			newDataValidation.showInputMessage = true;
			newDataValidation.allowBlank = true;
			return newDataValidation;
		} else {
			return this.dataValidations.getProps(_selection.ranges, doExtend, this);
		}
	};

	Worksheet.prototype.setDataValidationProps = function (props) {
		var _selection = this.getSelection();

		if (!this.dataValidations) {
			this.dataValidations = new window['AscCommonExcel'].CDataValidations();
		}

		this.dataValidations.setProps(this, _selection.ranges, props);
	};

	Worksheet.prototype.addDataValidation = function (dataValidation, addToHistory) {
		if (!this.dataValidations) {
			this.dataValidations = new window['AscCommonExcel'].CDataValidations();
		}
		this.dataValidations.add(this, dataValidation, addToHistory);
	};

	Worksheet.prototype.deleteDataValidationById = function (id) {
		if (this.dataValidations) {
			this.dataValidations.delete(this, id);
		}
	};

	Worksheet.prototype.getDataValidationById = function (id) {
		if (this.dataValidations) {
			return this.dataValidations.getById(id);
		}
	};

	Worksheet.prototype.shiftDataValidation = function (bInsert, operType, updateRange, addToHistory) {
		if (this.dataValidations) {
			this.dataValidations.shift(this, bInsert, operType, updateRange, addToHistory);
		}
	};

	Worksheet.prototype.clearDataValidation = function (ranges, addToHistory) {
		if (this.dataValidations) {
			this.dataValidations.clear(this, ranges, addToHistory);
		}
	};

	Worksheet.prototype._moveDataValidation = function (oBBoxFrom, oBBoxTo, copyRange, offset, wsTo) {
		if (!wsTo) {
			wsTo = this;
		}
		if (false === this.workbook.bUndoChanges && false === this.workbook.bRedoChanges) {
			wsTo.clearDataValidation([oBBoxTo], true);

			if (this === wsTo) {
				if (this.dataValidations) {
					this.dataValidations.move(this, oBBoxFrom, oBBoxTo, copyRange, offset);
				}
			} else {
				var aDataValidations = this._getCopyDataValidationByRange(oBBoxFrom, offset);
				if (aDataValidations) {
					if (!copyRange) {
						this.clearDataValidation([oBBoxFrom], true);
					}

					//далее необходимо создать новые объекты на новом листе
					for (var i = 0; i < aDataValidations.length; i++) {
						wsTo.addDataValidation(aDataValidations[i], true);
					}
				}
			}
		}
	};

	Worksheet.prototype._getCopyDataValidationByRange = function (range, offset) {
		if (this.dataValidations) {
			return this.dataValidations.getCopyByRange(range, offset);
		}
		return null;
	};


	Worksheet.prototype.setCFRule = function (val) {
		if (!val) {
			return;
		}
		var changedRule = this.getCFRuleById(val.id);
		if (changedRule) {
			this.changeCFRule(changedRule, val, true);
		} else {
			this.addCFRule(val, true);
		}
	};

	Worksheet.prototype.getCFRuleById = function (id) {
		if (this.isConditionalFormattingRules()) {
			return this.aConditionalFormattingRules[id];
		}
		return null;
	};

	Worksheet.prototype.isConditionalFormattingRules = function () {
		return this.aConditionalFormattingRules != null;
	};

	Worksheet.prototype.getConditionalFormattingRules = function () {
		return this.aConditionalFormattingRules;
	};

	Worksheet.prototype.addConditionalFormattingRule = function (rule) {
		if (!this.isConditionalFormattingRules()) {
			this.aConditionalFormattingRules = {};
		}
		this.aConditionalFormattingRules[rule.id] = rule;
	};

	Worksheet.prototype.changeCFRule = function (from, to, addToHistory) {
		if (!from) {
			return;
		}

		var updateRange;
		if (from) {
			updateRange = from.ranges;
		}
		if (to) {
			if (updateRange) {
				updateRange = updateRange.concat(to.ranges);
			} else {
				updateRange = to.ranges;
			}
		}
		if (updateRange) {
			this.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange(updateRange));
		}

		to.recalcInterfaceFormula(this);
		from.set(to, addToHistory, this);
	};

	Worksheet.prototype.addCFRule = function (val, addToHistory) {
		if (!val) {
			return;
		}

		if (!val.ranges) {
			val.ranges = [];
			for (var i = 0; i < this.selectionRange.ranges.length; i++) {
				val.ranges.push(this.selectionRange.ranges[i].clone());
			}
		}


		val.recalcInterfaceFormula(this);
		this.addConditionalFormattingRule(val);
		this.cleanConditionalFormattingRangeIterator();
		if (addToHistory) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_CFRuleAdd, this.getId(), val.getUnionRange(), new AscCommonExcel.UndoRedoData_CF(val.id, null, val.clone ? val.clone() : val));
		}
	};

	Worksheet.prototype.tryClearCFRule = function (rule, ranges) {
		if (!rule) {
			return;
		}

		var isDel = false;
		if (ranges) {
			var _newRanges = [];

			var ruleRanges = rule.ranges;
			for (var i = 0; i < ruleRanges.length; i++) {

				var tempRanges = [];
				for (var j = 0; j < ranges.length; j++) {
					if (tempRanges.length) {
						var tempRanges2 = [];
						for (var k = 0; k < tempRanges.length; k++) {
							tempRanges2 = tempRanges2.concat(ranges[j].intersection(tempRanges[k]) ? ranges[j].difference(tempRanges[k]) : tempRanges[k]);
						}
						tempRanges = tempRanges2;
					} else {
						tempRanges = ranges[j].intersection(ruleRanges[i]) ? ranges[j].difference(ruleRanges[i]) : ruleRanges[i];
					}
				}
				_newRanges = _newRanges.concat(tempRanges);
			}

			if (!_newRanges.length) {
				this.deleteCFRule(rule.id, true);
				isDel = true;
			} else {
				var newRule = rule.clone();
				newRule.ranges = _newRanges;
				this.changeCFRule(rule, newRule, true);
			}
		} else {
			this.deleteCFRule(rule.id, true);
			isDel = true;
		}

		return isDel;
	};

	Worksheet.prototype.deleteCFRule = function (id, addToHistory) {
		let oRule = this.getCFRuleById(id);
		if (oRule) {
			delete this.aConditionalFormattingRules[id];
			this.cleanConditionalFormattingRangeIterator();
			if (addToHistory) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_CFRuleDelete, this.getId(), oRule.getUnionRange(),
					new AscCommonExcel.UndoRedoData_CF(id, oRule));
			}
			if (oRule.ranges) {
				this.setDirtyConditionalFormatting(new AscCommonExcel.MultiplyRange(oRule.ranges));
			}
		}
	};

	Worksheet.prototype.generateCFRuleFromPreset = function (presetId) {
		if (!presetId || !presetId.length) {
			return null;
		}

		var oRule = new AscCommonExcel.CConditionalFormattingRule();
		oRule.applyPreset(presetId);
		return oRule;
	};

	Worksheet.prototype.updateConditionalFormattingOffset = function (range, offset) {
		let t = this;
		if (offset.row < 0 || offset.col < 0) {
			this.forEachConditionalFormattingRules(function (val) {
				if (val.containsIntoRange(range)) {
					t.deleteCFRule(val.id, true);
				}
			});
		}
		this.setConditionalFormattingOffset(range, offset);
	};
	Worksheet.prototype.setConditionalFormattingOffset = function (range, offset) {
		let t = this;
		this.forEachConditionalFormattingRules(function (oRule) {
			oRule.setOffset(offset, range, t, true);
		});
	};
	Worksheet.prototype.moveConditionalFormatting = function (oBBoxFrom, oBBoxTo, copyRange, offset, wsTo, wsFrom, bTransposeTo) {
		var t = this;
		if (!wsTo) {
			wsTo = this;
		}
		if (false === this.workbook.bUndoChanges && false === this.workbook.bRedoChanges) {
			//чистим ту область, куда переносим
			wsTo.clearConditionalFormattingRulesByRanges([oBBoxTo], oBBoxFrom);

			if (!wsFrom) {
				wsFrom = this;
			}
			wsFrom.forEachConditionalFormattingRules(function (_rule) {
				//если клонируем - то добавляем новое правило со смещенным диапазоном пересечения
				//если нет + если в пределах одного листа - меняем диапазона у текущего правила
				//если на другой лист - меняем диапазон у текущего правила + создаём новое со смещенным диапазоном пересечения

				var isChanged = null;
				var ruleRanges = _rule.ranges;
				var constantPart, movePart;
				var _moveRanges = [];
				var _constantRanges = [];
				for (var i = 0; i < ruleRanges.length; i++) {
					movePart = ruleRanges[i].intersection(oBBoxFrom);
					if (movePart) {
						if (!copyRange) {
							constantPart = oBBoxFrom.difference(ruleRanges[i]);
							_constantRanges = _constantRanges.concat(constantPart);
						}
						if (bTransposeTo) {
							movePart = movePart.transpose(oBBoxFrom.c1, oBBoxFrom.r1);
						}
						movePart.setOffset(offset);
						_moveRanges.push(movePart);
						isChanged = true;
					} else if (!copyRange) {
						_constantRanges.push(ruleRanges[i]);
					}
				}
				if (isChanged) {
					//в случае клонирования фрагмента - создаём новое правило
					var _newRule;
					if (copyRange) {
						_newRule = _rule.clone();
						_newRule.ranges = _moveRanges;
						wsTo.addCFRule(_newRule, true);
					} else {
						if (t !== wsTo) {
							if (_moveRanges.length) {
								_newRule = _rule.clone();
								_newRule.ranges = _moveRanges;
								wsTo.addCFRule(_newRule, true);
							}
							if (_constantRanges.length) {
								_rule.setLocation(_constantRanges, t, true);
							}
						} else {
							_rule.setLocation(_constantRanges.concat(_moveRanges), t, true);
						}
					}
				}
			});
		}
		wsTo && wsTo.forEachConditionalFormattingRules(function (_rule) {
			_rule.updateFormulas(wsTo);
		});
		wsFrom && wsFrom.forEachConditionalFormattingRules(function (_rule) {
			_rule.updateFormulas(wsFrom);
		});
	};

	Worksheet.prototype.clearConditionalFormattingRulesByRanges = function (ranges, exceptionRange) {
		let t = this;
		this.forEachConditionalFormattingRules(function (oRule) {
			var isExcept = exceptionRange && oRule.getIntersections(exceptionRange);
			if (!isExcept) {
				t.tryClearCFRule(oRule, ranges);
			}
		});
	};

	Worksheet.prototype.forEachConditionalFormattingRules = function (callback) {
		if (!this.isConditionalFormattingRules()) {
			return;
		}
		for (let i in this.aConditionalFormattingRules) {
			if (callback(this.aConditionalFormattingRules[i], i)) {
				break;
			}
		}
	};
	Worksheet.prototype.cleanConditionalFormattingRangeIterator = function() {
		this.conditionalFormattingRangeIterator = null;
	};
	Worksheet.prototype.getConditionalFormattingRangeIterator = function() {
		if (!this.conditionalFormattingRangeIterator) {
			let aRules = [];
			if (this.isConditionalFormattingRules()) {
				this.forEachConditionalFormattingRules(function (val) {
					aRules.push(val);
				})
			}
			aRules.sort(function(v1, v2) {
				return v2.priority - v1.priority;
			});
			this.conditionalFormattingRangeIterator = new AscCommon.RangeTopBottomIterator();
			this.conditionalFormattingRangeIterator.init(aRules, function(rule) {
				return rule.ranges || [];
			});
		}
		return this.conditionalFormattingRangeIterator;
	};
	Worksheet.prototype.updateTopLeftCell = function(range) {
		this.setTopLeftCell(this.generateTopLeftCellFromRange(range), true);
	};
	Worksheet.prototype.generateTopLeftCellFromRange = function(range) {
		if (!range) {
			return null;
		}

		//todo закрепленные области
		var newVal;
		if (range.c1 === 0 && range.r1 === 0) {
			newVal = null;
		} else {
			var topLeftFrozenCell = this.sheetViews[0] && this.sheetViews[0].pane && this.sheetViews[0].pane.topLeftFrozenCell;
			if (topLeftFrozenCell && topLeftFrozenCell.col > 1 && topLeftFrozenCell.row > 1) {
				newVal = null;
			} else {
				newVal = new Asc.Range(0, 0, 0, 0);
				if (!topLeftFrozenCell || topLeftFrozenCell.col <= 1) {
					newVal.c1 = newVal.c2 = range.c1;
				}
				if (!topLeftFrozenCell || topLeftFrozenCell.row <= 1) {
					newVal.r1 = newVal.r2 = range.r1;
				}
			}
		}

		return newVal;
	};
	Worksheet.prototype.setTopLeftCell = function(val, addToHistory) {
		var view = this.sheetViews[0];
		
		if (!view) {
			return;
		}
		
		if ((!val && view.topLeftCell) || (val && !view.topLeftCell) || (val && view.topLeftCell && !val.isEqual(view.topLeftCell))) {
			var oldValue = view.topLeftCell ? view.topLeftCell.clone() : null;
			view.topLeftCell = val;
			if (addToHistory) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetTopLeftCell,
					this.getId(), null, new UndoRedoData_FromTo(oldValue ? new UndoRedoData_BBox(oldValue) : oldValue, val ? new UndoRedoData_BBox(val) : val));
			}
		}
	};

	Worksheet.prototype.getTopLeftCell = function() {
		var view = this.sheetViews && this.sheetViews[0];
		return view && view.topLeftCell;
	};

	Worksheet.prototype.setCustomSort = function(props, obj, doNotSortRange, cellCommentator, opt_range) {
		//формируем sortState из настроек
		var t = this;
		var selection = opt_range ? opt_range : this.selectionRange.getLast();
		var sortState = new AscCommonExcel.SortState();

		//? activeRange
		sortState.Ref = new Asc.Range(selection.c1, selection.r1, selection.c2, selection.r2);

		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		var columnSort = props.columnSort;
		sortState.ColumnSort = !columnSort;
		sortState.CaseSensitive = props.caseSensitive;
		for(var i = 0; i < props.levels.length; i++) {
			var sortCondition = new AscCommonExcel.SortCondition();
			var level = props.levels[i];
			var r1 = columnSort ? selection.r1 : level.index + selection.r1;
			var c1 = columnSort ? selection.c1 + level.index : selection.c1;
			var r2 = columnSort ? selection.r2 : level.index + selection.r1;
			var c2  = columnSort ? selection.c1 + level.index : selection.c2;
			sortCondition.Ref = new Asc.Range(c1, r1, c2, r2);
			sortCondition.ConditionSortBy = null;
			sortCondition.ConditionDescending = Asc.c_oAscSortOptions.Descending === level.descending;

			var conditionSortBy = level.sortBy;
			var sortColor = null, newDxf, isRgbColor;
			switch (conditionSortBy) {
				case Asc.c_oAscSortOptions.ByColorFill: {
					sortCondition.ConditionSortBy = Asc.ESortBy.sortbyCellColor;
					sortColor = level.color;
					isRgbColor = sortColor && sortColor.getType && sortColor.getType() === AscCommonExcel.UndoRedoDataTypes.RgbColor;
					sortColor = sortColor && !isRgbColor ? new AscCommonExcel.RgbColor((sortColor.asc_getR() << 16) + (sortColor.asc_getG() << 8) + sortColor.asc_getB()) : null;

					newDxf = new AscCommonExcel.CellXfs();
					newDxf.fill = new AscCommonExcel.Fill();
					newDxf.fill.fromColor(sortColor);

					break;
				}
				case Asc.c_oAscSortOptions.ByColorFont: {
					sortCondition.ConditionSortBy = Asc.ESortBy.sortbyFontColor;
					sortColor = level.color;
					isRgbColor = sortColor && sortColor.getType && sortColor.getType() === AscCommonExcel.UndoRedoDataTypes.RgbColor;
					sortColor = sortColor && !isRgbColor ? new AscCommonExcel.RgbColor((sortColor.asc_getR() << 16) + (sortColor.asc_getG() << 8) + sortColor.asc_getB()) : null;

					newDxf = new AscCommonExcel.CellXfs();
					newDxf.font = new AscCommonExcel.Font();
					newDxf.font.setColor(sortColor);

					break;
				}
				case Asc.c_oAscSortOptions.ByIcon: {
					sortCondition.ConditionSortBy = Asc.ESortBy.sortbyIcon;
					break;
				}
				default: {
					sortCondition.ConditionSortBy = Asc.ESortBy.sortbyValue;
					break;
				}
			}

			if(newDxf) {
				sortCondition.dxf = AscCommonExcel.g_StyleCache.addXf(newDxf);
			}


			if(!sortState.SortConditions) {
				sortState.SortConditions = [];
			}

			sortState.SortConditions.push(sortCondition);
		}

		if(obj) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoSortState, AscCH.historyitem_SortState_Add, t.getId(), null,
				new AscCommonExcel.UndoRedoData_SortState(obj.sortState ? obj.sortState.clone() : null, sortState ? sortState.clone() : null, true, obj.DisplayName));

			obj.SortState = sortState;
		} else {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoSortState, AscCH.historyitem_SortState_Add, t.getId(), null,
				new AscCommonExcel.UndoRedoData_SortState(t.sortState ? t.sortState.clone() : null, sortState ? sortState.clone() : null));

			sortState._hasHeaders = props.hasHeaders;
			t.sortState = sortState;
		}

		if(!doNotSortRange) {
			var range = t.getRange3(selection.r1, selection.c1, selection.r2, selection.c2);
			var oSort = t._doSort(range, null, null, null, null, !columnSort, sortState);
			if (cellCommentator) {
				cellCommentator.sortComments(oSort);
			}
		}

		AscCommon.History.EndTransaction();
	};

	Worksheet.prototype._doSort = function (range, nOption, nStartRowCol, sortColor, opt_guessHeader, opt_by_row, opt_custom_sort) {
		var res;

		var bordersArr = [];
		range._foreachNoEmpty(function(cell, row, col) {
			var style = cell ? cell.getStyle() : null;
			if(style && style.border) {
				if(!bordersArr[row]) {
					bordersArr[row] = [];
				}
				bordersArr[row][col] = style.border;
				cell.setBorder(null);
			}
		});
		res = range.sort(nOption, nStartRowCol, sortColor, opt_guessHeader, opt_by_row, opt_custom_sort);
		for(var i = 0; i < bordersArr.length; i++) {
			if(bordersArr[i]) {
				for(var j = 0; j < bordersArr[i].length; j++) {
					if(bordersArr[i][j]) {
						var curBorder = bordersArr[i][j];
						this._getCell(i, j, function(cell) {
							cell.setBorder(curBorder);
						});

					}
				}
			}
		}

		return res;
	};

	Worksheet.prototype.getIntersectionRules = function (range) {
		var t = this;
		var res = [];
		this.forEachConditionalFormattingRules(function (_rule) {
			var changedRanges = _rule.getIntersections(range);
			if (changedRanges) {
				res.push({ranges: changedRanges, id: _rule.id});
			}
		});
		return res.length ? res : null;
	};

	Worksheet.prototype.getRuleById = function (id) {
		var t = this;
		var res = null;
		this.forEachConditionalFormattingRules(function (_rule, index) {
			if (_rule.id === id) {
				res = {data: _rule, index: index};
				return true;
			}
		});
		return res;
	};


	Worksheet.prototype.setProtectedSheet = function (props, addToHistory) {
		if (!this.sheetProtection) {
			this.sheetProtection = new window["Asc"].CSheetProtection();
			this.sheetProtection.setDefaultInterface();
		}
		this.cleanTempProtectedRanges();
		this.sheetProtection.set(props, addToHistory, this);
		if (this.sheetProtection.isDefault()) {
			this.sheetProtection = null;
		}
		return true;
	};

	Worksheet.prototype.cleanTempProtectedRanges = function () {
		if (this.aProtectedRanges && this.aProtectedRanges.length) {
			for (var i = 0; i < this.aProtectedRanges.length; i++) {
				this.aProtectedRanges[i].cleanTemp();
			}
		}
	};

	Worksheet.prototype.getSheetProtection = function (type) {
		var res = false;

		if (this.sheetProtection && this.sheetProtection.getSheet()) {
			//если тип не определен возвращаем защищен лист или нет
			res = true;
			switch (type) {
				case Asc.c_oAscSheetProtectType.objects:
					res = this.sheetProtection.getObjects();
					break;
				case Asc.c_oAscSheetProtectType.scenarios:
					res = this.sheetProtection.getScenarios();
					break;
				case Asc.c_oAscSheetProtectType.formatCells:
					res = this.sheetProtection.getFormatCells();
					break;
				case Asc.c_oAscSheetProtectType.formatColumns:
					res = this.sheetProtection.getFormatColumns();
					break;
				case Asc.c_oAscSheetProtectType.formatRows:
					res = this.sheetProtection.getFormatRows();
					break;
				case Asc.c_oAscSheetProtectType.insertColumns:
					res = this.sheetProtection.getInsertColumns();
					break;
				case Asc.c_oAscSheetProtectType.insertRows:
					res = this.sheetProtection.getInsertRows();
					break;
				case Asc.c_oAscSheetProtectType.insertHyperlinks:
					res = this.sheetProtection.getInsertHyperlinks();
					break;
				case Asc.c_oAscSheetProtectType.deleteColumns:
					res = this.sheetProtection.getDeleteColumns();
					break;
				case Asc.c_oAscSheetProtectType.deleteRows:
					res = this.sheetProtection.getDeleteRows();
					break;
				case Asc.c_oAscSheetProtectType.selectLockedCells:
					res = this.sheetProtection.getSelectLockedCells();
					break;
				case Asc.c_oAscSheetProtectType.sort:
					res = this.sheetProtection.getSort();
					break;
				case Asc.c_oAscSheetProtectType.autoFilter:
					res = this.sheetProtection.getAutoFilter();
					break;
				case Asc.c_oAscSheetProtectType.pivotTables:
					res = this.sheetProtection.getPivotTables();
					break;
				case Asc.c_oAscSheetProtectType.selectUnlockedCells:
					res = this.sheetProtection.getSelectUnlockedCells();
					break;
			}
		}

		return res;
	};

	Worksheet.prototype.getLockedCell = function (c, r) {
		var cellTo =  this.getCell3(r, c);
		if (cellTo) {
			var cellxfs = cellTo.getXfs(false);
			return cellxfs && cellxfs.asc_getLocked();
		}
	};

	Worksheet.prototype.getUnlockedCellInRange = function (range) {
		if (!range) {
			range = new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0);
		}
		var res = null;
		this.getRange3(range.r1, range.c1, range.r2, range.c2)._foreachNoEmpty(function(cell) {
			if (cell) {
				var cellxfs = cell.xfs;
				if (cellxfs && !cellxfs.asc_getLocked()) {
					res = cell;
					return true;
				}
			}
		});
		return res;
	};

	Worksheet.prototype.protectedRangesContains = function (c, r) {
		var res = [];
		if (this.aProtectedRanges && this.aProtectedRanges.length) {
			for (var i = 0; i < this.aProtectedRanges.length; i++) {
				if (this.aProtectedRanges[i].contains(c, r)) {
					res.push(this.aProtectedRanges[i]);
				}
			}
		}
		return res.length ? res : null;
	};

	Worksheet.prototype.protectedRangesContainsRanges = function (ranges) {
		for (var i = 0; i < ranges.length; i++) {
			if (!this.protectedRangesContainsRange(ranges[i])) {
				return false;
			}
		}
		return true;
	};

	Worksheet.prototype.protectedRangesContainsRange = function (range, ignoreWithoutPassword) {
		var res = [];
		if (this.aProtectedRanges && this.aProtectedRanges.length) {
			for (var i = 0; i < this.aProtectedRanges.length; i++) {
				if (this.aProtectedRanges[i].containsRange(range)) {
					var isPassword = this.aProtectedRanges[i].asc_isPassword();
					if (!ignoreWithoutPassword || (ignoreWithoutPassword && isPassword && !this.aProtectedRanges[i].isUserEnteredPassword())) {
						res.push(this.aProtectedRanges[i]);
					}
				}
			}
		}
		return res.length ? res : null;
	};

	Worksheet.prototype.checkProtectedRangesPassword = function (val, data, callback) {
		//здесь првоеряем в тч при попытке ввода в ячейку
		var t = this;
		if (this.aProtectedRanges && this.aProtectedRanges.length) {
			var aCheckHash = [];
			var checkRanges = [];
			for (var i = 0; i < this.aProtectedRanges.length; i++) {
				if (!this.aProtectedRanges[i].contains(data.col, data.row)) {
					continue;
				}
				if (!this.aProtectedRanges[i].asc_isPassword() || this.aProtectedRanges[i]._isEnterPassword) {
					callback && callback(true);
					return;
				} else if (this.aProtectedRanges[i].asc_isPassword()) {
					checkRanges.push(this.aProtectedRanges[i]);
					aCheckHash.push({
						password: val,
						salt: this.aProtectedRanges[i].saltValue,
						spinCount: this.aProtectedRanges[i].spinCount,
						alg: AscCommon.fromModelAlgorithmName(this.aProtectedRanges[i].algorithmName)
					});
				}
			}
			if (aCheckHash && aCheckHash.length) {
				AscCommon.calculateProtectHash(aCheckHash, function (aHash) {
					for (var i = 0; i < aHash.length; i++) {
						if (aHash[i] === checkRanges[i].hashValue) {
							checkRanges[i]._isEnterPassword = true;
							callback && callback(true);
							return;
						}
					}
					callback && callback(false);
				});
			} else {
				callback && callback(true);
			}
		} else {
			callback && callback(true);
		}
	};

	Worksheet.prototype.checkProtectedRangeName = function (name) {
		var res = c_oAscDefinedNameReason.OK;
		//TODO пересмотреть проверку на rx_defName
		if (!AscCommon.rx_protectedRangeName.test(name.toLowerCase()) || name.length > g_nDefNameMaxLength) {
			return c_oAscDefinedNameReason.WrongName;
		}

		var pR = this.getProtectedRangeByName(name);
		if (pR) {
			res = c_oAscDefinedNameReason.Existed;
		}
		return res;
	};

	Worksheet.prototype.isIntersectLockedRanges = function (ranges) {
		for (var i = 0; i < ranges.length; i++) {
			if (this.isLockedRange(ranges[i])) {
				return true;
			}
		}
		return false;
	};

	Worksheet.prototype.isLockedRange = function (range) {
		let oRange = this.getRange3(range.r1, range.c1, range.r2, range.c2);
		let res = true;

		let _getLocked = function (_xfs) {
			let _res = null;
			if (_xfs/* && _xfs.applyProtection*/) {
				_res = true;//null/true
				if (_xfs.getLocked() === false) {
					_res = false;
				}
			}
			return _res;
		};

		//1. init default value
		//all cells
		if (null != this.oAllCol && this.oAllCol.xfs && this.oAllCol.xfs.locked != null) {
			res = _getLocked(this.oAllCol.xfs);
		}

		//all selected rows
		let unLockedRowIndex = range.r1 - 1;
		oRange._foreachRowNoEmpty(function(row){
			if(null != row && row.xfs && !res === _getLocked(row.xfs)) {
				if (unLockedRowIndex + 1 === row.index) {
					unLockedRowIndex++;
				}
			}
		});
		if (unLockedRowIndex === range.r2) {
			res = !res;
		}

		//all selected cols
		let unLockedColIndex = range.c1 - 1;
		oRange._foreachColNoEmpty(function(col){
			if(null != col && col.xfs && !res === _getLocked(col.xfs)) {
				if (unLockedColIndex + 1 === col.index) {
					unLockedColIndex++;
				}
			}
		});
		if (unLockedColIndex === range.c2) {
			res = !res;
		}

		let counter = 0;
		let bbox = oRange.getBBox0 && oRange.getBBox0();
		let sizeRange = bbox ? bbox.getHeight() * bbox.getWidth() : null;
		oRange._foreachNoEmpty(function(cell){
			if (!cell) {
				return;
			}

			var isLocked = _getLocked(cell.xfs);
			if (isLocked === true) {
				res = true;
				return true;
			} else if (isLocked === false) {
				counter++;
				if (res === true && sizeRange) {
					if (counter === sizeRange) {
						res = false;
					}
				} else {
					res = false;
				}
			}
		});

		return res;
	};

	Worksheet.prototype.getLockedRanges = function (range) {
		var res = [range.clone()];
		this.getRange3(range.r1, range.c1, range.r2, range.c2)._foreachNoEmpty(function(cell){
			if (!cell) {
				return;
			}
			var cellxfs = cell && cell.xfs;
			var isLocked = cellxfs && cellxfs.asc_getLocked();
			if (isLocked === false) {
				//исключаем ячейку из общего диапазона
				var _range = new Asc.Range(cell.nCol, cell.nRow, cell.nCol, cell.nRow);
				var newRange = [];
				for (var i = 0; i < res.length; i++) {
					var _difference = _range.difference(res[i]);
					if (_difference && _difference.length) {
						newRange = newRange.concat(_difference);
					}
				}
				res = newRange;
			}
		});
		return res;
	};


	Worksheet.prototype.getProtectedRanges = function (needClone) {
		var protectedRanges = this.aProtectedRanges;
		var res = null;
		if (needClone && protectedRanges && protectedRanges.length) {
			res = [];
			for (var i = 0; i < protectedRanges.length; i++) {
				var id = protectedRanges[i].Id;
				var cloneRange = protectedRanges[i].clone();
				cloneRange.Id = id;
				cloneRange.isLock = protectedRanges[i].isLock;
				res.push(cloneRange);
			}
		}
		return !res ? protectedRanges : res;
	};

	Worksheet.prototype.setProtectedRange = function (val) {
		if (!val) {
			return;
		}
		var pR = this.getProtectedRangeById(val.Id);
		if (pR) {
			this.changeProtectedRange(pR.val, val, true);
		} else {
			this.addProtectedRange(val, true);
		}
	};

	Worksheet.prototype.getProtectedRangeById = function (id) {
		if (this.aProtectedRanges) {
			for (var i = 0; i < this.aProtectedRanges.length; i++) {
				if (this.aProtectedRanges[i].Id === id) {
					return {val: this.aProtectedRanges[i], index: i};
				}
			}
		}
		return null;
	};

	Worksheet.prototype.getProtectedRangeByRange = function (range) {
		if (this.aProtectedRanges) {
			for (var i = 0; i < this.aProtectedRanges.length; i++) {
				if (this.aProtectedRanges[i].intersection(range)) {
					return {val: this.aProtectedRanges[i], index: i};
				}
			}
		}
		return null;
	};

	Worksheet.prototype.getProtectedRangeByName = function (name) {
		if (this.aProtectedRanges) {
			for (var i = 0; i < this.aProtectedRanges.length; i++) {
				if (this.aProtectedRanges[i].name === name) {
					return {val: this.aProtectedRanges[i], index: i};
				}
			}
		}
		return null;
	};

	Worksheet.prototype.changeProtectedRange = function (from, to, addToHistory) {
		if (!from) {
			return;
		}
		from.set(to, addToHistory, this);
	};

	Worksheet.prototype.addProtectedRange = function (val, addToHistory) {
		if (!val) {
			return;
		}

		this.aProtectedRanges.push(val);
		if (addToHistory) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_AddProtectedRange,
				this.getId(), /*val.getUnionRange()*/null,
				new AscCommonExcel.UndoRedoData_ProtectedRange(val.Id, null, val));
		}
	};

	Worksheet.prototype.deleteProtectedRange = function (id, addToHistory) {
		var protectedRange = this.getProtectedRangeById(id);
		if (protectedRange) {
			this.aProtectedRanges.splice(protectedRange.index, 1);
			if (addToHistory) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_DelProtectedRange,
					this.getId(), /*oRule.val.getUnionRange()*/null,
					new AscCommonExcel.UndoRedoData_ProtectedRange(id, protectedRange.val));
			}
		}
	};

	Worksheet.prototype.updateProtectedRangeOffset = function (range, offset) {
		if (offset.row < 0 || offset.col < 0) {
			for (var i = 0; i < this.aProtectedRanges.length; ++i) {
				if (this.aProtectedRanges[i].containsIntoRange(range)) {
					this.deleteProtectedRange(this.aProtectedRanges[i].Id, true);
				}
			}
		}
		this.setProtectedRangeOffset(range, offset);
	};

	Worksheet.prototype.setProtectedRangeOffset = function (range, offset) {
		var protectedRange;
		for (var i = 0; i < this.aProtectedRanges.length; ++i) {
			protectedRange = this.aProtectedRanges[i];
			protectedRange.setOffset(offset, range, this, true);
		}
	};

	Worksheet.prototype.moveProtectedRange = function (oBBoxFrom, oBBoxTo, copyRange, offset, wsTo, wsFrom, bTransposeTo) {
		var t = this;
		if (!wsTo) {
			wsTo = this;
		}

		if (wsTo.isUserProtectedRangesIntersection(oBBoxTo) || (wsFrom && wsFrom.isUserProtectedRangesIntersection(oBBoxFrom))) {
			wsTo.workbook.handlers && wsTo.workbook.handlers.trigger("asc_onError", c_oAscError.ID.ProtectedRangeByOtherUser, c_oAscError.Level.NoCritical);
			return;
		}

		if (wsTo.getSheetProtection()) {
			return;
		}
		if (false === this.workbook.bUndoChanges && false === this.workbook.bRedoChanges) {
			//чистим ту область, куда переносим
			if (wsTo.aProtectedRanges && wsTo.aProtectedRanges.length) {
				wsTo.aProtectedRanges.forEach(function (_protectedRange) {
					t.tryClearProtectedRange(_protectedRange, [oBBoxTo]);
				});
			}

			if (!wsFrom) {
				wsFrom = this;
			}

			if (wsFrom.aProtectedRanges && wsFrom.aProtectedRanges.length) {
				wsFrom.aProtectedRanges.forEach(function (pR) {
					//если клонируем - то добавляем новый диапазон со смещенным диапазоном пересечения
					//если нет + если в пределах одного листа - меняем диапазона у текущего правила
					//если на другой лист - меняем диапазон у текущего + создаём новый со смещенным диапазоном пересечения

					var isChanged = null;
					var _protectedSqref = pR.sqref;
					var constantPart, movePart;
					var _moveRanges = [];
					var _constantRanges = [];
					for (var i = 0; i < _protectedSqref.length; i++) {
						movePart = _protectedSqref[i].intersection(oBBoxFrom);
						if (movePart) {
							if (!copyRange) {
								constantPart = oBBoxFrom.difference(_protectedSqref[i]);
								_constantRanges = _constantRanges.concat(constantPart);
							}

							if (bTransposeTo) {
								movePart = movePart.transpose(oBBoxFrom.c1, oBBoxFrom.r1);
							}

							movePart.setOffset(offset);
							_moveRanges.push(movePart);
							isChanged = true;
						} else if (!copyRange) {
							_constantRanges.push(_protectedSqref[i]);
						}
					}
					if (isChanged) {
						//в случае клонирования фрагмента - создаём новый
						var _newPr;
						if (copyRange) {
							if (wsFrom === wsTo) {
								pR.setSqref(pR.sqref.concat(_moveRanges), t, true);
							} else {
								_newPr = pR.clone();
								_newPr.sqref = _moveRanges;
								_newPr.generateNewName(wsTo.aProtectedRanges);
								wsTo.addProtectedRange(_newPr, true);
							}
						} else {
							if (t !== wsTo) {
								if (_moveRanges.length) {
									_newPr = pR.clone();
									_newPr.sqref = _moveRanges;
									wsTo.addProtectedRange(_newPr, true);
								}
								if (_constantRanges.length) {
									pR.setSqref(_constantRanges, t, true);
								}
							} else {
								pR.setSqref(_constantRanges.concat(_moveRanges), t, true);
							}
						}
					}
				});
			}
		}
	};

	Worksheet.prototype.tryClearProtectedRange = function (protectedRange, ranges) {
		if (!protectedRange) {
			return;
		}

		if (ranges) {
			var _newRanges = [];

			var protectedSqref = protectedRange.sqref;
			for (var i = 0; i < protectedSqref.length; i++) {

				var tempRanges = [];
				for (var j = 0; j < ranges.length; j++) {
					if (tempRanges.length) {
						var tempRanges2 = [];
						for (var k = 0; k < tempRanges.length; k++) {
							tempRanges2 = tempRanges2.concat(ranges[j].intersection(tempRanges[k]) ? ranges[j].difference(tempRanges[k]) : tempRanges[k]);
						}
						tempRanges = tempRanges2;
					} else {
						tempRanges = ranges[j].intersection(protectedSqref[i]) ? ranges[j].difference(protectedSqref[i]) : protectedSqref[i];
					}
				}
				_newRanges = _newRanges.concat(tempRanges);
			}

			if (!_newRanges.length) {
				this.deleteProtectedRange(protectedRange.Id, true)
			} else {
				var newProtectedRange = protectedRange.clone();
				newProtectedRange.sqref = _newRanges;
				this.changeProtectedRange(protectedRange, newProtectedRange, true);
			}
		} else {
			this.deleteProtectedRange(protectedRange.Id, true);
		}
	};

	Worksheet.prototype.getProtectedRangesByActiveCell = function () {
		var activeCell = this.selectionRange.activeCell;
		var res = [];
		for (var i = 0; i < this.aProtectedRanges.length; i++) {
			if (this.aProtectedRanges[i].contains(activeCell.col, activeCell.row)) {
				res.push(this.aProtectedRanges[i]);
			}
		}
		return res.length ? res : null;
	};

	Worksheet.prototype.getProtectedRangesByActiveRange = function () {
		var res = [];
		var activeRanges = this.selectionRange.ranges;

		if (this.aProtectedRanges) {
			for (var i = 0; i < this.aProtectedRanges.length; i++) {
				for (var j = 0; j < activeRanges.length; j++) {
					if (this.aProtectedRanges[i].intersection(activeRanges[j])) {
						res.push(this.aProtectedRanges[i]);
					}
				}
			}
		}

		return res.length ? res : null;
	};

	Worksheet.prototype.isLockedActiveCell = function () {
		var activeCell = this.selectionRange.activeCell;
		return this.getLockedCell(activeCell.col, activeCell.row);
	};

	Worksheet.prototype.getPrintOptionsJson = function () {
		var printProps = this.PagePrintOptions;
		printProps.initPrintTitles();
		printProps = printProps.clone();
		printProps.pageSetup.headerFooter = this && this.headerFooter && this.headerFooter.getForInterface();
		var printArea = this.workbook.getDefinesNames("Print_Area", this.getId());
		printProps.pageSetup.printArea = printArea ? printArea.clone() : false;

		printProps.printTitlesHeight = this.PagePrintOptions.printTitlesHeight;
		printProps.printTitlesWidth = this.PagePrintOptions.printTitlesWidth;

		if (this.PagePrintOptions && this.PagePrintOptions.pageSetup) {
			printProps.pageSetup.fitToHeight = this.PagePrintOptions.pageSetup.asc_getFitToHeight();
			printProps.pageSetup.fitToWidth = this.PagePrintOptions.pageSetup.asc_getFitToWidth();
		}

		return printProps.getJson(this);
	};

	Worksheet.prototype.changeRowColBreaks = function (from, to, range, byCol, addToHistory) {
		let min = null;
		let max = !byCol ? gc_nMaxCol0 : gc_nMaxRow0;
		let man = true, pt = null;

		let printArea = this.workbook.getDefinesNames("Print_Area", this.getId());
		if (printArea && range) {
			if (byCol) {
				min = range.r1;
				max = range.r2;
			} else {
				min = range.c1;
				max = range.c2;
			}
		}

		this._changeRowColBreaks(from, to, min, max, man, pt, byCol, addToHistory);
	};

	Worksheet.prototype._changeRowColBreaks = function (from, to, min, max, man, pt, byCol, addToHistory) {
		let t = this;
		let rowColBreaks = !byCol ? t.rowBreaks : t.colBreaks;

		let checkInit = function () {
			if (!byCol) {
				if (!t.rowBreaks) {
					t.rowBreaks = new AscCommonExcel.CRowColBreaks();
					rowColBreaks = t.rowBreaks;
				}
			} else {
				if (!t.colBreaks) {
					t.colBreaks = new AscCommonExcel.CRowColBreaks();
					rowColBreaks = t.colBreaks;
				}
			}
		};

		let isChanged = false;
		let oFromBreak = rowColBreaks && rowColBreaks.getBreak(from, min, max);
		if (oFromBreak) {
			if (to) {
				//change
				checkInit();
				isChanged = rowColBreaks.changeBreak(from, to, min, max, true);
			} else {
				//delete
				isChanged = rowColBreaks && rowColBreaks.removeBreak(from, min, max);
			}
		} else if (to) {
			//add
			checkInit();
			isChanged = rowColBreaks.addBreak(to, min, max, true, null)
		}

		if (isChanged && addToHistory) {
			let fromData = oFromBreak && new AscCommonExcel.UndoRedoData_RowColBreaks(from, oFromBreak.min, oFromBreak.max, oFromBreak.man, oFromBreak.pt, byCol);
			let toData = new AscCommonExcel.UndoRedoData_RowColBreaks(to, min, max, man, pt, byCol);

			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ChangeRowColBreaks, this.getId(),
				null, new AscCommonExcel.UndoRedoData_FromTo(fromData, toData));
		}

		this.workbook.handlers.trigger("onChangePageSetupProps", this.getId());
	};

	Worksheet.prototype.isBreak = function (index, range, byCol) {
		let min = null;
		let max = !byCol ? gc_nMaxCol0 : gc_nMaxRow0;

		let printArea = this.workbook.getDefinesNames("Print_Area", this.getId());
		if (printArea && range) {
			if (byCol) {
				min = range.r1;
				max = range.r2;
			} else {
				min = range.c1;
				max = range.c2;
			}
		}

		let rowColBreaks = !byCol ? this.rowBreaks : this.colBreaks;
		return rowColBreaks && rowColBreaks.isBreak(index, min, max);
	};

	Worksheet.prototype.resetAllPageBreaks = function () {
		let t = this;

		let doRemoveBreaks = function(_breaks, byCol) {
			if (!_breaks) {
				return;
			}
			let aBreaks = _breaks.getBreaks();
			for (let i = 0; i < aBreaks.length; i++) {
				let fromData = new AscCommonExcel.UndoRedoData_RowColBreaks(aBreaks[i].id, aBreaks[i].min, aBreaks[i].max, aBreaks[i].man, aBreaks[i].pt, byCol);
				if (_breaks.removeBreak(aBreaks[i].id)) {
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ChangeRowColBreaks, t.getId(),
						null, new AscCommonExcel.UndoRedoData_FromTo(fromData, null));
					i--;
				}
			}
		};

		if (this.rowBreaks) {
			doRemoveBreaks(this.rowBreaks);
		}
		if (this.colBreaks) {
			doRemoveBreaks(this.colBreaks, true);
		}
		this.workbook.handlers.trigger("onChangePageSetupProps", this.getId());
	};

	Worksheet.prototype.getPrintAreaRangeByRowCol = function (row, col) {
		let printArea = this.workbook.getDefinesNames("Print_Area", this.getId());
		let t = this;

		if (printArea) {
			let ranges;
			AscCommonExcel.executeInR1C1Mode(false, function () {
				ranges = AscCommonExcel.getRangeByRef(printArea.ref, t, true, true)
			});
			if (ranges) {
				for (let i = 0; i < ranges.length; i++) {
					if (ranges[i].bbox.contains(col, row)) {
						return ranges[i].bbox;
					}
				}
			}
		}

		return null;
	};

	Worksheet.prototype.addCellWatch = function (ref, addToHistory) {
		for (var i = 0; i < this.aCellWatches.length; i++) {
			if (ref.isEqual(this.aCellWatches[i].r)) {
				return;
			}
		}

		var cellWatch = new AscCommonExcel.CCellWatch(this);
		cellWatch.setRef(ref);
		this.aCellWatches.push(cellWatch);

		if (addToHistory) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_AddCellWatch, this.getId(),
				null, new AscCommonExcel.UndoRedoData_FromTo(null, new AscCommonExcel.UndoRedoData_BBox(ref)));
		}

		this.workbook.handlers && this.workbook.handlers.trigger("changeCellWatches", this.getIndex());
	};

	Worksheet.prototype.deleteCellWatch = function (ref, addToHistory) {
		for (var i = 0; i < this.aCellWatches.length; i++) {
			if (ref.isEqual(this.aCellWatches[i].r)) {
				this.aCellWatches.splice(i, 1);
				if (addToHistory) {
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_DelCellWatch,
						this.getId(), null,
						new AscCommonExcel.UndoRedoData_FromTo(new AscCommonExcel.UndoRedoData_BBox(ref), null));
				}
				this.workbook.handlers && this.workbook.handlers.trigger("changeCellWatches", this.getIndex());
				break;
			}
		}
	};

	Worksheet.prototype.deleteCellWatches = function (addToHistory) {
		if (addToHistory) {
			for (var i = 0; i < this.aCellWatches.length; i++) {
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_DelCellWatch,
					this.getId(), null,
					new AscCommonExcel.UndoRedoData_FromTo(new AscCommonExcel.UndoRedoData_BBox(this.aCellWatches[i].r), null));
			}
		}
		this.aCellWatches = [];
		this.workbook.handlers && this.workbook.handlers.trigger("changeCellWatches", this.getIndex());
	};

	Worksheet.prototype.deleteCellWatchesByRange = function (range, addToHistory) {
		for (var i = 0; i < this.aCellWatches.length; i++) {
			if (range.containsRange(this.aCellWatches[i].r)) {
				this.deleteCellWatch(this.aCellWatches[i].r, addToHistory);
				i--;
			}
		}
	};

	Worksheet.prototype.getCellWatchesByRange = function (range) {
		let res = [];
		for (var i = 0; i < this.aCellWatches.length; i++) {
			if (range.containsRange(this.aCellWatches[i].r)) {
				res.push(this.aCellWatches[i]);
			}
		}
		return res;
	};

	Worksheet.prototype.checkImportXmlLocationForError = function(ranges) {
		for (var i = 0; i < ranges.length; ++i) {
			var range = ranges[i];
			if (this.autoFilters.isIntersectionTable(range)) {
				return c_oAscError.ID.PivotOverlap;
			}
			if (this.inPivotTable(range)) {
				return c_oAscError.ID.PivotOverlap;
			}
		}
		return c_oAscError.ID.No;
	};

	//*****user range protect*****
	Worksheet.prototype.editUserProtectedRanges = function(oldObj, newObj, addToHistory) {
		//if undo/redo - need apply changes
		if (addToHistory && ((newObj && this.isIntersectionOtherUserProtectedRanges(newObj.ref)) ||
			(oldObj && this.isIntersectionOtherUserProtectedRanges(oldObj.ref)))) {
			return false;
		}

		let newRes;
		if (oldObj || newObj) {
			let cloneNewOnj = newObj && newObj.clone();
			newRes = cloneNewOnj;
			if (cloneNewOnj) {
				cloneNewOnj.Id = newObj.Id;
				cloneNewOnj._ws = this;
			}
			if (oldObj) {
				let modelUserRange = this.getUserProtectedRangeById(oldObj.Id);
				if (modelUserRange) {
					//TODO clone
					if (cloneNewOnj) {
						this.userProtectedRanges[modelUserRange.index] = cloneNewOnj;
					} else {
						this.userProtectedRanges.splice(modelUserRange.index, 1);
					}
				}
			} else if (cloneNewOnj) {
				this.userProtectedRanges.push(cloneNewOnj);
			}

			if (addToHistory) {
				let updateRange = null;
				if (oldObj && newObj && oldObj.ref && newObj.ref) {
					updateRange = oldObj.ref.clone();
					updateRange.union2(newObj.ref);
				} else if (oldObj && oldObj.ref) {
					updateRange = oldObj.ref.clone();
				} else if (newObj && newObj.ref) {
					updateRange = newObj.ref.clone();
				}

				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ChangeUserProtectedRange, this.getId(), updateRange,
					new UndoRedoData_FromTo(oldObj, newObj));
			}
		}

		return newRes ? newRes : true;
	};

	Worksheet.prototype.deleteUserProtectedRanges = function (range) {
		for (var i = this.userProtectedRanges.length - 1; i >= 0; --i) {
			var userProtectedRange = this.userProtectedRanges[i];
			if (range.containsRange(userProtectedRange.ref)) {
				this.editUserProtectedRanges(userProtectedRange, null, true);
			}
		}
		return true;
	};

	Worksheet.prototype.getUserProtectedRangeById = function(id) {
		var res = null;
		if(!this.userProtectedRanges)
			return res;

		for(var i = 0; i < this.userProtectedRanges.length; i++)
		{
			if(this.userProtectedRanges[i].Id === id)
			{
				res = {obj: this.userProtectedRanges[i], index: i};
				break;
			}
		}
		return res;
	};

	Worksheet.prototype.getUserProtectedRangeByName = function(name) {
		var res = null;
		if(!this.userProtectedRanges)
			return res;

		for(var i = 0; i < this.userProtectedRanges.length; i++)
		{
			if(this.userProtectedRanges[i].name === name)
			{
				res = {obj: this.userProtectedRanges[i], index: i};
				break;
			}
		}
		return res;
	};

	Worksheet.prototype.unlockUserProtectedRanges = function(){
		if (this.userProtectedRanges) {
			for (let i = 0; i < this.userProtectedRanges.length; i++) {
				this.userProtectedRanges[i].isLock = null;
			}
		}
	};

	Worksheet.prototype.isIntersectionOtherUserProtectedRanges = function(range){
		let res = false;
		if (this.userProtectedRanges) {
			let oApi = Asc.editor;
			let sUserId = oApi.DocInfo && oApi.DocInfo.get_UserId();
			if (sUserId) {
				for (let i = 0; i < this.userProtectedRanges.length; i++) {
					let curUserProtectedRange = this.userProtectedRanges[i];
					if (curUserProtectedRange.intersection(range) && !curUserProtectedRange.isUserCanEdit(sUserId)) {
						res = true;
						break;
					}
				}
			}
		}
		return res;
	};

	Worksheet.prototype.isUserProtectedRangesIntersection = function (range, userId, notCheckUser, type) {
		//range - array of ranges or range
		let res = false;
		if (!this.userProtectedRanges || !this.userProtectedRanges.length) {
			return res;
		}

		if (!userId) {
			let oApi = Asc.editor;
			userId = oApi.DocInfo && oApi.DocInfo.get_UserId();
		}
		if (this.userProtectedRanges && userId) {
			for (let i = 0; i < this.userProtectedRanges.length; i++) {
				let curUserProtectedRange = this.userProtectedRanges[i];
				if (range && range.length) {
					for (let j = 0; j < range.length; j++) {
						if (curUserProtectedRange.intersection(range[j]) && (notCheckUser || !curUserProtectedRange.isUserCanDoByType(userId, type))) {
							return true;
						}
					}
				} else if ((!range || curUserProtectedRange.intersection(range)) && (notCheckUser || !curUserProtectedRange.isUserCanDoByType(userId, type))) {
					res = true;
					break;
				}
			}
		}
		return res;
	};

	Worksheet.prototype.isUserProtectedRangesIntersectionCell = function(oCell, userId, notCheckUser, type){
		//range - array of ranges or range
		let res = false;
		if (!this.userProtectedRanges || !this.userProtectedRanges.length || !oCell) {
			return res;
		}
		if (!userId) {
			let oApi = Asc.editor;
			userId = oApi.DocInfo && oApi.DocInfo.get_UserId();
		}

		oCell = {col: oCell.nCol, row: oCell.nRow};

		if (this.userProtectedRanges && userId) {
			for (let i = 0; i < this.userProtectedRanges.length; i++) {
				let curUserProtectedRange = this.userProtectedRanges[i];
				if (curUserProtectedRange.contains2(oCell) && (notCheckUser || !curUserProtectedRange.isUserCanDoByType(userId, type))) {
					return true;
				}
			}
		}
		return res;
	};

	Worksheet.prototype.isUserProtectedRangesCanView = function(oCell, userId){
		//range - array of ranges or range
		let res = true;
		if (!this.userProtectedRanges || !this.userProtectedRanges.length || !oCell) {
			return res;
		}
		if (!userId) {
			let oApi = Asc.editor;
			userId = oApi.DocInfo && oApi.DocInfo.get_UserId();
		}

		oCell = {col: oCell.nCol, row: oCell.nRow};

		if (this.userProtectedRanges && userId) {
			for (let i = 0; i < this.userProtectedRanges.length; i++) {
				let curUserProtectedRange = this.userProtectedRanges[i];
				if (curUserProtectedRange.contains2(oCell) && !curUserProtectedRange.isUserCanView(userId)) {
					return false;
				}
			}
		}
		return res;
	};

	Worksheet.prototype.updateUserProtectedRangesOffset = function (range, offset) {
		if (offset.row < 0 || offset.col < 0) {
			this.deleteUserProtectedRanges(range);
		}

		this.shiftUserProtectedRanges(range, offset);
	};
	Worksheet.prototype.moveUserProtectedRangesOffset = function (range, offset) {
		var userProtectedRange;
		for (var i = 0; i < this.userProtectedRanges.length; ++i) {
			userProtectedRange = this.userProtectedRanges[i];
			if (userProtectedRange.isInRange(range)) {
				userProtectedRange.setOffset(offset, true);
			}
		}
	};

	Worksheet.prototype._moveUserProtectedRange = function (oBBoxFrom, oBBoxTo, copyRange, offset, wsTo) {
		if (!wsTo) {
			wsTo = this;
		}
		if (false == this.workbook.bUndoChanges && false == this.workbook.bRedoChanges) {
			if (copyRange) {
				wsTo.deleteUserProtectedRanges(oBBoxTo);
				this.copyUserProtectedRanges(oBBoxFrom, offset, wsTo);
			} else {
				if (this === wsTo) {
					this.moveUserProtectedRangesOffset(oBBoxFrom, offset);
				} else {
					this.copyUserProtectedRanges(oBBoxFrom, offset, wsTo);
					this.deleteUserProtectedRanges(oBBoxFrom);
				}
			}
		}
	};

	Worksheet.prototype.copyUserProtectedRanges = function (range, offset, wsTo) {
		var t = this;
		var userProtectedRange;
		for (var i = this.userProtectedRanges.length - 1; i >= 0; --i) {
			userProtectedRange = this.userProtectedRanges[i];
			if (userProtectedRange.isInRange(range)) {
				var newUserProtectedRange = userProtectedRange.clone(wsTo);
				newUserProtectedRange.setOffset(offset, false);
				wsTo.editUserProtectedRanges(null, newUserProtectedRange, true);
			}
		}
	};

	Worksheet.prototype.shiftUserProtectedRanges = function (range, offset, wsTo) {
		var t = this;
		var userProtectedRange;
		for (var i = this.userProtectedRanges.length - 1; i >= 0; --i) {
			userProtectedRange = this.userProtectedRanges[i];

			if (range.isIntersectForShift(userProtectedRange.ref, offset)) {
				let newRef = userProtectedRange.ref.clone();
				newRef.forShift(range, offset);
				if (!userProtectedRange.ref.isEqual(newRef)) {
					userProtectedRange.setLocation(newRef, true);
				}
			}
		}
	};


	Worksheet.prototype.setSheetViewType = function(val, addToHistory) {
		var sheetView = this.sheetViews[0];

		if (!sheetView) {
			return;
		}

		var oldValue = sheetView.view;
		if (oldValue !== val) {
			sheetView.view = val;
			if (addToHistory) {
				AscCommon.History.Create_NewPoint();
				AscCommon.History.StartTransaction();

				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetSheetViewType,
					this.getId(), null, new UndoRedoData_FromTo(oldValue, val));

				AscCommon.History.EndTransaction();
			}

			this.workbook.handlers && this.workbook.handlers.trigger("asc_updateSheetViewType", this.index);
			return true;
		}
	};

	Worksheet.prototype.setColsCount = function(val) {
		this.nColsCount = val;
	};

	Worksheet.prototype.changeLegacyDrawingHFPictures = function(picturesMap) {
		if (!this.legacyDrawingHF) {
			this.legacyDrawingHF = new AscCommonExcel.CLegacyDrawingHF(this);
		}
		this.legacyDrawingHF.addPictures(picturesMap);
	};

	Worksheet.prototype.removeLegacyDrawingHFPictures = function(aPictures) {
		if (!this.legacyDrawingHF) {
			this.legacyDrawingHF = new AscCommonExcel.CLegacyDrawingHF(this);
		}
		this.legacyDrawingHF.removePictures(aPictures);
	};

	Worksheet.prototype.getLegacyDrawingHFById = function(val) {
		if (!this.legacyDrawingHF) {
			return null;
		}
		return this.legacyDrawingHF.getDrawingById(val);
	};

	Worksheet.prototype.getCountNoEmptyCells = function() {
		if (this.nRowsCount === 0 || this.nColsCount === 0) {
			return 0;
		}
		let count = 0;
		let range = this.getRange3(0, 0, this.nRowsCount - 1, this.nColsCount - 1);
		range._foreachNoEmptyByCol(function () {
			count++;
		});
		return count;
	};
	/**
	 * The method returns a flag that recognizes fill handle was called from right click
	 * @returns {boolean}
	 */
	Worksheet.prototype.getFillHandleRightClick = function () {
		return this.bFillHandleRightClick;
	};
	/**
	 * The method sets a flag that recognizes fill handle was called from right click
	 * @param {boolean} bFillHandleRightClick
	 */
	Worksheet.prototype.setFillHandleRightClick = function (bFillHandleRightClick) {
		this.bFillHandleRightClick = bFillHandleRightClick;
	};
	/**
	 * The method returns the value of chosen property from context menu for autofill.
	 * @memberof Worksheet
	 * @returns {Asc.c_oAscFillType}
	 */
	Worksheet.prototype.getFillMenuChosenProp = function () {
		return this.nFillMenuChosenProp;
	};
	/**
	 * The method sets the value of chosen property from context menu for autofill.
	 * @memberof Worksheet
	 * @param {Asc.c_oAscFillType} nFillMenuChosenProp
	 */
	Worksheet.prototype.setFillMenuChosenProp = function (nFillMenuChosenProp) {
		this.nFillMenuChosenProp = nFillMenuChosenProp;
	}
	Worksheet.prototype.getActiveFillType = function () {
		return this.activeFillType;
	};
	Worksheet.prototype.setActiveFillType = function (val) {
		this.activeFillType = val;
	};
	Worksheet.prototype.onTimelinesDelete = function(name, addToHistory) {
		if (this.timelines && this.timelines.length) {
			for (let i = 0; i < this.timelines.length; i++) {
				if (this.timelines[i].name === name) {
					if (addToHistory) {
						AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_TimelineDelete,
							this.getId(), null, new UndoRedoData_FromTo(this.timelines[i], null));
					}

					this.timelines.splice(i, 1);
					i--;
				}
			}
		}
	};
	Worksheet.prototype.isSelectedTab = function () {
		for (let i = 0; i < this.sheetViews.length; i++) {
			if (this.sheetViews[i].tabSelected) {
				return true;
			}
		}
		return false;
	};
	
	Worksheet.prototype.addChangedArray = function (rangeName, arrayInfo) {
		if (!rangeName || !arrayInfo) {
			return
		}

		if (!this.changedArrays) {
			this.changedArrays = {};
		}

		if (!this.changedArrays[rangeName]) {
			this.changedArrays[rangeName] = arrayInfo
		}
	};
	Worksheet.prototype.getChangedArrayList = function () {
		if (!this.changedArrays) {
			return
		}

		return this.changedArrays
	};
	Worksheet.prototype.removeFromChangedArrayList = function (name) {
		if (!this.changedArrays) {
			return
		}

		return this.changedArrays[name]
	};
	Worksheet.prototype.clearChangedArrayList = function () {
		this.changedArrays = null;
	};
	Worksheet.prototype.recalculateVolatileArrays = function () {
		if (AscCommonExcel.bIsSupportDynamicArrays) {
			const ws = this;

			// recalculate all volatile arrays on page
			if (ws && ws.workbook && ws.workbook.dependencyFormulas) {
				let depGraph = ws.workbook.dependencyFormulas;
				// get volatileArraysList
				let volatileArrayList = depGraph.getVolatileArrays();
				if (volatileArrayList) {
					for (let listenerId in volatileArrayList) {
						let formula = volatileArrayList[listenerId];
						let formulaResult = formula.calculate();
						let firstCellRef = formula.parent && new Asc.Range(formula.parent.nCol, formula.parent.nRow, formula.parent.nCol, formula.parent.nRow);
						if (!(formula.aca && formula.ca)) {
							// array can expand, setValue for each cell except first
							let dimensions = formulaResult.getDimensions();
							let newRef = new Asc.Range(firstCellRef.c1, firstCellRef.r1, firstCellRef.c1 + dimensions.col - 1, firstCellRef.r1 + dimensions.row - 1);
							formula.setDynamicRef(newRef);
							if (newRef) {
								for (let row = newRef.r1; row <= newRef.r2; row++) {
									for (let col = newRef.c1; col <= newRef.c2; col++) {
										// if (row === newRef.r1 && col === newRef.c1) {
										// 	continue
										// }
										// get cell and setPF to it
										ws._getCell(row, col, function(cell) {
											cell && cell.setFormulaInternal(formula);
										});
									}
								}
							}
							depGraph.addToChangedRange2(formula.getWs().getId(), formula.getDynamicRef());
							depGraph.endListeningVolatileArray(listenerId);
						} else {
							formula.setDynamicRef(firstCellRef);
							ws._getCell(firstCellRef.r1, firstCellRef.c1, function(cell) {
								cell && cell.setFormulaInternal(formula);
							});

							depGraph.addToChangedRange2(formula.getWs().getId(), formula.getArrayFormulaRef());
						}
					}
				}
			}
		}
	};
	Worksheet.prototype.getRefDynamicInfo = function (formula, calculateResult) {
		if (formula) {
			let applyByArray = true, ctrlKey = true, dynamicRange = null, cannotChangeFormulaArray = false;

			if (formula.ref) {
				dynamicRange = formula.ref;
			} else {
				let tempRef = new Asc.Range(formula.parent.nCol, formula.parent.nRow, formula.parent.nCol, formula.parent.nRow);
				formula.ref = tempRef;

				let formulaResult = formula.calculate(null, null, null, null, calculateResult);
				let arraySize = formulaResult.getDimensions(true);
				let newR2 = (formula.parent.nRow + arraySize.row) > AscCommon.gc_nMaxRow ? AscCommon.gc_nMaxRow - 1 : (formula.parent.nRow + arraySize.row - 1);
				let newC2 = (formula.parent.nCol + arraySize.col) > AscCommon.gc_nMaxCol ? AscCommon.gc_nMaxCol - 1 : (formula.parent.nCol + arraySize.col - 1);

				if (formulaResult.type !== cElementType.array) {
					return false;
				}

				let tempDynamicSelectionRange = this.getRange3(formula.parent.nRow, formula.parent.nCol, newR2, newC2);
				tempDynamicSelectionRange._foreachNoEmpty(function (cell) {
					let ref = cell.formulaParsed && cell.formulaParsed.ref ? cell.formulaParsed.ref : null;

					if (ref && !tempDynamicSelectionRange.bbox.containsRange(ref)) {
						cannotChangeFormulaArray = true;
						return false;
					}
				});

				if (tempDynamicSelectionRange.bbox.isOneCell() && (formulaResult.type !== cElementType.array && formulaResult.type !== cElementType.cellsRange && formulaResult.type !== cElementType.cellsRange3D)) {
					applyByArray = false;
					ctrlKey = false;
				}

				dynamicRange = tempDynamicSelectionRange.bbox;
			}

			return {applyByArray: applyByArray, ctrlKey: ctrlKey, dynamicRange: dynamicRange, cannotChangeFormulaArray: cannotChangeFormulaArray};
		}
	};

	Worksheet.prototype.findEOT = function (bCheckStyles) {
		var maxCols = this.getColsCount();
		var maxRows = this.getRowsCount();
		var lastC = -1, lastR = -1;

		let t = this;
		let _cell = null;
		let _getCell = function (_col, _row) {
			if (_col < 0 || _col > gc_nMaxCol0 || _row < 0 || _row > gc_nMaxRow0) {
				return null;
			}

			return t.getCell3(_row, _col);
		};
		let _isCellNullText = function (_col, _row) {
			let c = _col;
			_cell = null;
			if (row !== undefined) {
				c = _getCell(_col, _row);
				_cell = c;
			}
			return null === c || c.isNullText();
		};

		let isEmptyCell = function () {
			if (_cell) {
				if (_cell.hasMerged()) {
					return false;
				} else {
					var ws = _cell.worksheet;
					var nRow = _cell.bbox.r1;
					var nCol = _cell.bbox.c1;
					var xfs;
					ws._getCellNoEmpty(nRow, nCol, function (oCell) {
						xfs = ws.getCompiledStyle(nRow, nCol, oCell);
					});
					if (xfs) {
						return false;
					}
				}
			}
			return true;
		};

		for (var col = 0; col < maxCols; ++col) {
			for (var row = 0; row < maxRows; ++row) {
				if (!_isCellNullText(col, row) || (bCheckStyles && !isEmptyCell())) {
					lastC = Math.max(lastC, col);
					lastR = Math.max(lastR, row);
				}
			}
		}

		return new AscCommon.CellBase(lastR, lastC);
	};

	Worksheet.prototype.getAllInks = function(arrInks) {
		arrInks = arrInks || [];
		for (let i = 0; i < this.Drawings.length; i++) {
			const oGraphicObject = this.Drawings[i].graphicObject;
			if (oGraphicObject.isInk() || oGraphicObject.isHaveOnlyInks()) {
				arrInks.push(oGraphicObject);
			} else {
				oGraphicObject.getAllInks(arrInks);
			}
		}
		return arrInks;
	};


//-------------------------------------------------------------------------------------------------
	var g_nCellStructSize = 4 + 4 + 8;

	var g_nCellFlag_empty = 0;
	var g_nCellFlag_init = 1;
	var g_nCellFlag_typeMask = 6;
	var g_nCellFlag_valueMask = 24;
	var g_nCellFlag_isDirtyMask = 32;
	var g_nCellFlag_isCalcMask = 64;
	/**
	 * @constructor
	 */
	function Cell(worksheet){
		this.ws = worksheet;
		this.nRow = -1;
		this.nCol = -1;
		this.xfs = null;
		this.formulaParsed = null;

		this.type = CellValueType.Number;
		this.number = null;
		this.text = null;
		this.multiText = null;
		this.textIndex = null;

		this.isDirty = false;
		this.isCalc = false;

		this.cm = null;
		this.vm = null;

		this._hasChanged = false;
	}
	Cell.prototype.clear = function(keepIndex) {
			this.nRow = -1;
			this.nCol = -1;
		this.clearData();

		this._hasChanged = false;
	};
	Cell.prototype.clearData = function() {
		this.xfs = null;
		this.formulaParsed = null;

		this.type = CellValueType.Number;
		this.number = null;
		this.text = null;
		this.multiText = null;
		this.textIndex = null;

		this.isDirty = false;
		this.isCalc = false;

		this.cm = null;
		this.vm = null;

		this._hasChanged = true;
	};
	Cell.prototype.clearDataKeepXf = function(border) {
		var xfs = this.xfs;
		this.clearData();
		this.xfs = xfs;
		AscCommon.History.TurnOff();
		this.setBorder(border);
		AscCommon.History.TurnOn();
	};
	Cell.prototype.saveContent = function(opt_inCaseOfChange) {
		if (this.hasRowCol() && (!opt_inCaseOfChange || this._hasChanged)) {
			this._hasChanged = false;
			var wb = this.ws.workbook;
			var sheetMemory = this.ws.getColData(this.nCol);
			sheetMemory.checkIndex(this.nRow);
			var xfSave = this.xfs ? this.xfs.getIndexNumber() : 0;
			var numberSave = 0;
			var formulaSave = this.formulaParsed ? wb.workbookFormulas.add(this.formulaParsed).getIndexNumber() : 0;
			var flagValue = 0;
			if (null != this.number) {
				flagValue = 1;
				const flags = this._toFlags(flagValue);
				sheetMemory.setInt32(this.nRow, 0, xfSave | (flags << 24));
				sheetMemory.setInt32(this.nRow, 4, formulaSave);
				sheetMemory.setFloat64(this.nRow, 8, this.number);
			} else if (null != this.text) {
				flagValue = 2;
				const flags = this._toFlags(flagValue);
				sheetMemory.setInt32(this.nRow, 0, xfSave | (flags << 24));
				sheetMemory.setInt32(this.nRow, 4, formulaSave);
				numberSave = this.getTextIndex();
				sheetMemory.setInt32(this.nRow, 8, numberSave);
			} else if (null != this.multiText) {
				flagValue = 3;
				const flags = this._toFlags(flagValue);
				sheetMemory.setInt32(this.nRow, 0, xfSave | (flags << 24));
				sheetMemory.setInt32(this.nRow, 4, formulaSave);
				numberSave = this.getTextIndex();
				sheetMemory.setInt32(this.nRow, 8, numberSave);
			} else {
				const flags = this._toFlags(flagValue);
				sheetMemory.setInt32(this.nRow, 0, xfSave | (flags << 24));
				sheetMemory.setInt32(this.nRow, 4, formulaSave);
			}
		}
	};
	Cell.prototype.loadContent = function(row, col, opt_sheetMemory) {
		var res = false;
		this.clear();
		this.nRow = row;
		this.nCol = col;
		let sheetMemory = opt_sheetMemory;
		if (!sheetMemory) {
			sheetMemory = this.ws.getColDataNoEmpty(this.nCol);
			if (!sheetMemory) {
				return res;
			}
		}
		if (sheetMemory.hasIndex(this.nRow)) {
			const mix = sheetMemory.getInt32(this.nRow, 0);
			const flags = (mix >> 24) & 0xff;
			const xfIndex = mix & 0xffffff;
			if (0 !== (g_nCellFlag_init & flags)) {
				var wb = this.ws.workbook;
				var flagValue = this._fromFlags(flags);
				if (xfIndex > 0) {
					this.xfs = g_StyleCache.getXf(xfIndex);
				}
				const formulaIndex = sheetMemory.getInt32(this.nRow, 4)
				if (formulaIndex > 0) {
					this.formulaParsed = wb.workbookFormulas.get(formulaIndex);
				}
				if (1 === flagValue) {
					this.number = sheetMemory.getFloat64(this.nRow, 8);
				} else if (2 === flagValue) {
					this.textIndex = sheetMemory.getInt32(this.nRow, 8);
					this.text = wb.sharedStrings.get(this.textIndex);
				} else if (3 === flagValue) {
					this.textIndex = sheetMemory.getInt32(this.nRow, 8);
					this.multiText = wb.sharedStrings.get(this.textIndex);
				}
				res = true;
			}
		}
		return res;
	};
	Cell.prototype._toFlags = function(flagValue) {
		var flags = g_nCellFlag_init | (this.type << 1) | (flagValue << 3);
		if(this.isDirty){
			flags |= g_nCellFlag_isDirtyMask;
		}
		if(this.isCalc){
			flags |= g_nCellFlag_isCalcMask;
		}
		return flags;
	};
	Cell.prototype._fromFlags = function(flags) {
		this.type = (flags & g_nCellFlag_typeMask) >>> 1;
		this.isDirty = 0 != (flags & g_nCellFlag_isDirtyMask);
		this.isCalc = 0 != (flags & g_nCellFlag_isCalcMask);
		return (flags & g_nCellFlag_valueMask) >>> 3;
	};
	Cell.prototype.processFormula = function(callback) {
		if (this.formulaParsed) {
			var shared = this.formulaParsed.getShared();
			var offsetRow, offsetCol;
			if (shared) {
				offsetRow = this.nRow - shared.base.nRow;
				offsetCol = this.nCol - shared.base.nCol;
				if (0 !== offsetRow || 0 !== offsetCol) {
					var oldRow = this.formulaParsed.parent.nRow;
					var oldCol = this.formulaParsed.parent.nCol;

					//todo assemble by param
					var old = AscCommonExcel.g_ProcessShared;
					AscCommonExcel.g_ProcessShared = true;
					var offsetShared = new AscCommon.CellBase(offsetRow, offsetCol);
					this.formulaParsed.changeOffset(offsetShared, false);
					this.formulaParsed.parent.nRow = this.nRow;
					this.formulaParsed.parent.nCol = this.nCol;
					callback(this.formulaParsed);
					offsetShared.row = -offsetRow;
					offsetShared.col = -offsetCol;
					this.formulaParsed.changeOffset(offsetShared, false);
					this.formulaParsed.parent.nRow = oldRow;
					this.formulaParsed.parent.nCol = oldCol;
					AscCommonExcel.g_ProcessShared = old;
				} else {
					callback(this.formulaParsed);
				}
			} else {
				callback(this.formulaParsed);
			}
		}
	};
	Cell.prototype.setChanged = function(val) {
		this._hasChanged = val;
	};
	Cell.prototype.getStyle=function(){
		return this.xfs;
	};
	Cell.prototype.getCompiledStyle = function (opt_styleComponents, opt_AffectingText) {
		return this.ws.getCompiledStyle(this.nRow, this.nCol, this, opt_styleComponents, opt_AffectingText);
	};
	Cell.prototype.getCompiledStyleCustom = function(needTable, needCell, needConditional) {
		return this.ws.getCompiledStyleCustom(this.nRow, this.nCol, needTable, needCell, needConditional, this);
	};
	Cell.prototype.getTableStyle = function () {
		var hiddenManager = this.ws.hiddenManager;
		var sheetMergedStyles = this.ws.sheetMergedStyles;
		var styleComponents = sheetMergedStyles.getStyle(hiddenManager, this.nRow, this.nCol, this.ws);
		return getCompiledStyleFromArray(null, styleComponents.table);
	};
	Cell.prototype.duplicate=function(){
		var t = this;
		var oNewCell = new Cell(this.ws);
		oNewCell.nRow = this.nRow;
		oNewCell.nCol = this.nCol;
		oNewCell.xfs = this.xfs;
		oNewCell.type = this.type;
		oNewCell.number = this.number;
		oNewCell.text = this.text;
		oNewCell.multiText = this.multiText;
		this.processFormula(function(parsed) {
			//todo without parse
			var newFormula = new parserFormula(parsed.getFormula(), oNewCell, t.ws);
			AscCommonExcel.executeInR1C1Mode(false, function () {
				newFormula.parse();
			});
			if (parsed.ca) { // For recursive formulas
				newFormula.ca = parsed.ca;
			}
			var arrayFormulaRef = parsed.getArrayFormulaRef();
			if(arrayFormulaRef) {
				newFormula.setArrayFormulaRef(arrayFormulaRef);
			}
			oNewCell.setFormulaInternal(newFormula);
		});
		return oNewCell;
	};
	Cell.prototype.clone=function(oNewWs, renameParams){
		if(!oNewWs)
			oNewWs = this.ws;
		let oThis = this;
		var oNewCell = new Cell(oNewWs);
		oNewCell.nRow = this.nRow;
		oNewCell.nCol = this.nCol;
		if(null != this.xfs)
			oNewCell.xfs = this.xfs;
		oNewCell.type = this.type;
		oNewCell.number = this.number;
		oNewCell.text = this.text;
		oNewCell.multiText = this.multiText;
		this.processFormula(function(parsed) {
			var newFormula;
			if (oNewWs != oThis.ws && renameParams) {
				var formula = parsed.clone(null, null, this.ws);
				formula.renameSheetCopy(renameParams);
				newFormula = formula.assemble(true);
			} else {
				newFormula = parsed.getFormula();
			}
			oNewCell.setFormulaInternal(new parserFormula(newFormula, oNewCell, oNewWs));
			oNewWs.workbook.dependencyFormulas.addToBuildDependencyCell(oNewCell);
		});
		return oNewCell;
	};
	Cell.prototype.setRowCol=function(nRow, nCol){
		this.nRow = nRow;
		this.nCol = nCol;
	};
	Cell.prototype.hasRowCol = function() {
		return this.nRow >= 0 && this.nCol >= 0;
	};
	Cell.prototype.isEqual = function (options) {
		var cellText;
		cellText = (options.lookIn === Asc.c_oAscFindLookIn.Formulas) ? this.getValueForEdit() : this.getValue();
		if (true !== options.isMatchCase) {
			cellText = cellText.toLowerCase();
		}
		var isWordEnter = cellText.indexOf(options.findWhat);
		if (options.findWhat instanceof RegExp && options.isWholeWord) {
			isWordEnter = cellText.search(options.findWhat);
		}
		return (0 <= isWordEnter) &&
		(true !== options.isWholeCell || options.findWhat.length === cellText.length);

	};
	Cell.prototype.isNullText=function(){
		return this.isNullTextString() && !this.formulaParsed;
	};
	Cell.prototype.isEmptyTextString = function() {
		this._checkDirty();
		if(null != this.number || (null != this.text && "" != this.text))
			return false;
		if(null != this.multiText && "" != AscCommonExcel.getStringFromMultiText(this.multiText))
			return false;
		return true;
	};
	Cell.prototype.isNullTextString = function() {
		this._checkDirty();
		return null === this.number && null === this.text && null === this.multiText;
	};
	Cell.prototype.isEmpty=function(){
		if(false == this.isNullText())
			return false;
		if(null != this.xfs)
			return false;
		return true;
	};
	Cell.prototype.isFormula=function(){
		return this.formulaParsed ? true : false;
	};
	Cell.prototype.getName=function(){
		return g_oCellAddressUtils.getCellId(this.nRow, this.nCol);
	};
	Cell.prototype.setTypeInternal=function(val) {
		this.type = val;
		this._hasChanged = true;
	};
	Cell.prototype.correctValueByType = function() {
		//todo implemented only Number->String. other is stub
		switch (this.type) {
			case CellValueType.Number:
				if (null !== this.text) {
					this.setValueNumberInternal(parseInt(this.text) || 0);
				} else if (null !== this.multiText) {
					this.setValueNumberInternal(0);
				}
				break;
			case CellValueType.Bool:
				if (null !== this.text || null !== this.multiText) {
					this.setValueNumberInternal(0);
				}
				break;
			case CellValueType.String:
				if (null !== this.number) {
					this.setValueTextInternal(this.number.toString());
				}
				break;
			case CellValueType.Error:
				if (null !== this.number) {
					this.setValueTextInternal(cErrorOrigin["nil"]);
				}
				break;
		}
	};
	Cell.prototype.setValueNumberInternal=function(val) {
		this.number = val;
		this.text = null;
		this.multiText = null;
		this.textIndex = null;
		this._hasChanged = true;
	};
	Cell.prototype.setValueTextInternal=function(val) {
		this.number = null;
		this.text = val;
		this.multiText = null;
		this.textIndex = null;
		this._hasChanged = true;
	};
	Cell.prototype.setValueMultiTextInternal=function(val) {
		this.number = null;
		this.text = null;
		this.multiText = val;
		this.textIndex = null;
		this._hasChanged = true;
	};
	Cell.prototype.setValue=function(val,callback, isCopyPaste, byRef, ignoreHyperlink, dynamicRange, skipExternalCheck) {
		var ws = this.ws;
		var wb = ws.workbook;
		var DataOld = null;
		if (AscCommon.History.Is_On()) {
			DataOld = this.getValueData();
		}
		var isFirstArrayFormulaCell = byRef && this.nCol === byRef.c1 && this.nRow === byRef.r1;
		var newFP = this.setValueGetParsed(val, callback, isCopyPaste, byRef, dynamicRange);
		if (undefined === newFP) {
			return;
		}

		var oldFP = this.formulaParsed;
		//удаляем старые значения
		this.cleanText();
		this.setFormulaInternal(null);

		if (newFP) {
			this.setFormulaInternal(newFP);
			if(byRef) {
				if(isFirstArrayFormulaCell) {
					if (dynamicRange) {
						this.setDynamicArrayFlags();
					}
					wb.dependencyFormulas.addToBuildDependencyArray(newFP);
				}
			} else {
				wb.dependencyFormulas.addToBuildDependencyCell(this);
			}
			if (this.ws.workbook.handlers) {
				this.ws.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.cellValue, this, null, this.ws.getId());
			}
		} else if (val) {
			this._setValue(val, ignoreHyperlink);
			if (!ignoreHyperlink && window['AscCommonExcel'].g_AutoCorrectHyperlinks) {
				this._autoformatHyperlink(val);
			}
			wb.dependencyFormulas.addToChangedCell(this);
		} else {
			// when deleting a cell value we get here
			if (oldFP && oldFP.dynamicRange) {
				isFirstArrayFormulaCell = this.nCol === oldFP.dynamicRange.c1 && this.nRow === oldFP.dynamicRange.r1;
				// we check here and add the range with daf to the dependency sheet
				if (isFirstArrayFormulaCell) {
					wb.dependencyFormulas.addToChangedRange2(oldFP.getWs().getId(), oldFP.getDynamicRef());
				}
			} else {
				wb.dependencyFormulas.addToChangedCell(this);
			}
			if (this.ws.workbook.handlers) {
				this.ws.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.cellValue, this, null, this.ws.getId());
			}
		}

		var DataNew = null;
		if (g_cCalcRecursion.getCellPasteValue() != null) {
			this.setValueNumberInternal(g_cCalcRecursion.getCellPasteValue());
			g_cCalcRecursion.setCellPasteValue(null);
		}
		if (AscCommon.History.Is_On()) {
			DataNew = this.getValueData();
		}
		if (AscCommon.History.Is_On() && false == DataOld.isEqual(DataNew)) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(),
						new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow),
						new UndoRedoData_CellSimpleData(this.nRow, this.nCol, DataOld, DataNew));
		}
		//sortDependency вызывается ниже AscCommon.History.Add(AscCH.historyitem_Cell_ChangeValue, потому что в ней может быть выставлен формат ячейки(если это текстовый, то принимая изменения формула станет текстом)
		this.ws.workbook.sortDependency();
		if (!this.ws.workbook.dependencyFormulas.isLockRecal()) {
			this._adjustCellFormat();
		}

		//todo не должны удаляться ссылки, если сделать merge ее части.
		if (this.isNullTextString() && !this.isFormula()) {
			var cell = this.ws.getCell3(this.nRow, this.nCol);
			cell.removeHyperlink();
		}

		if (!skipExternalCheck) {
			this.checkRemoveExternalReferences(newFP, oldFP);
		}
	};

	Cell.prototype.checkRemoveExternalReferences=function(fNew, fOld) {
		//1.проверяем, были ли ссылки на внешние данные
		let externalLinks;
		let i;
		if (fOld && (!fNew || fNew.Formula !== fOld.Formula) && fOld.outStack) {
			let needCheckImportFuncs = false;
			for (i = 0; i < fOld.outStack.length; i++) {
				if (fOld.outStack[i].externalLink) {
					if (!externalLinks) {
						externalLinks = {};
					}
					externalLinks[fOld.outStack[i].externalLink] = fOld.outStack[i].getWsId();
				} else if (fOld.outStack[i].type === cElementType.func && fOld.outStack[i].name === "IMPORTRANGE") {
					needCheckImportFuncs = true;
				}
			}

			if (needCheckImportFuncs) {
				AscCommonExcel.importRangeLinksState.startBuildImportRangeLinks = true;
				fOld.calculate();
				AscCommonExcel.importRangeLinksState.startBuildImportRangeLinks = null;

				if (AscCommonExcel.importRangeLinksState.importRangeLinks) {
					for (i in AscCommonExcel.importRangeLinksState.importRangeLinks) {
						for (let j = 0; j < AscCommonExcel.importRangeLinksState.importRangeLinks[i].length; j++) {
							if (AscCommonExcel.importRangeLinksState.importRangeLinks[i][j]) {
								let ws = this.ws.workbook.getExternalWorksheet(i, AscCommonExcel.importRangeLinksState.importRangeLinks[i][j].sheet);
								if (ws) {
									if (!externalLinks) {
										externalLinks = {};
									}
									externalLinks[i] = ws.getId();
								}
							}

						}
					}
				}
				AscCommonExcel.importRangeLinksState.importRangeLinks = null;
			}
		}


		//2. проверяем, каких ссылок не осталось в новых данных
		if (externalLinks && fNew && fNew.outStack) {
			for (i = 0; i < fNew.outStack.length; i++) {
				if (fNew.outStack[i].externalLink) {
					if (externalLinks[fNew.outStack[i].externalLink]) {
						delete externalLinks[fNew.outStack[i].externalLink];
					}
				}
			}
			if (fNew.importFunctionsRangeLinks) {
				for (i in fNew.importFunctionsRangeLinks) {
					if (externalLinks[i]) {
						delete externalLinks[i];
					}
				}
			}
		}

		//3. проверям, не ссылаются ли на эти ссылки кто-то другой
		if (externalLinks && fOld) {
			let listenerId = fOld.getListenerId();

			for (let link in externalLinks) {
				// check all external link listeners
				// let erIndex = this.getExternalLinkIndexBySheetId(sheetId);
				// let eR = this.ws.workbook.externalReferences[+index - 1];
				let index = this.ws.workbook.getExternalLinkIndexBySheetId(externalLinks[link]);
				let eR = this.ws.workbook.externalReferences[index - 1];

				if (eR && null != listenerId) {
					let hasListeners;
					for (let ws in eR.worksheets) {
						// if listeners are found, then we interrupt the cycle and move on to the next external link
						if (hasListeners) {
							break
						}

						let wsId = eR.worksheets[ws].getId();
						let sheetContainer = fOld.wb && fOld.wb.dependencyFormulas && fOld.wb.dependencyFormulas.sheetListeners && fOld.wb.dependencyFormulas.sheetListeners[wsId];

						if (sheetContainer) {
							if (Object.keys(sheetContainer.cellMap).length === 0 && Object.keys(sheetContainer.areaMap).length === 0 && Object.keys(sheetContainer.defName3d).length === 0) {
								hasListeners = false;
							} else {
								hasListeners = true;
							}
						}
					}

					if (!hasListeners) {
						this.ws && this.ws.workbook && this.ws.workbook.removeExternalReference(index, true);
					}
				}
			}
		}

		// //3. проверям, не ссылаются ли на эти ссылки кто-то другой
		// if (externalLinks && fOld) {
		// 	let listenerId = fOld.getListenerId();
		// 	for (i in externalLinks) {
		// 		if (null != listenerId) {
		// 			let sheetId = externalLinks[i];
		// 			let sheetContainer = fOld.wb && fOld.wb.dependencyFormulas && fOld.wb.dependencyFormulas.sheetListeners && fOld.wb.dependencyFormulas.sheetListeners[sheetId];
		// 			if (sheetContainer && Object.keys(sheetContainer.cellMap).length === 0) {
		// 				//если есть ссылки на внешние источники, необходимо их удалить
		// 				this.ws && this.ws.workbook && this.ws.workbook.removeExternalReferenceBySheet(sheetId);
		// 			}
		// 		}
		// 	}
		// }
	};

	Cell.prototype.setValueGetParsed=function(val,callback, isCopyPaste, byRef, dynamicRange) {
		var ws = this.ws;
		var wb = ws.workbook;
		var bIsTextFormat = false;
		if (!isCopyPaste) {
			var sNumFormat;
			if (null != this.xfs && null != this.xfs.num) {
				sNumFormat = this.xfs.num.getFormat();
			} else {
				sNumFormat = g_oDefaultFormat.Num.getFormat();
			}
			var numFormat = oNumFormatCache.get(sNumFormat);
			bIsTextFormat = numFormat.isTextFormat();
		}

		var isFirstArrayFormulaCell = byRef && this.nCol === byRef.c1 && this.nRow === byRef.r1;
		var newFP = null;
		if (false == bIsTextFormat || (byRef && !isFirstArrayFormulaCell)) {
			/*
			 Устанавливаем значение в Range ячеек. При этом происходит проверка значения на формулу.
			 Если значение является формулой, то проверяем содержиться ли в ячейке формула или нет, если "да" - то очищаем в графе зависимостей список, от которых зависит формула(masterNodes), позже будет построен новый. Затем выставляем флаг о необходимости дальнейшего пересчета, и заносим ячейку в список пересчитываемых ячеек.
			 */
			if (null != val && val[0] == "=" && val.length > 1) {
				//***array-formula***
				if(byRef && !isFirstArrayFormulaCell) {
					newFP = this.ws.formulaArrayLink;
					if(newFP === null) {
						return;
					}
					if(this.nCol === byRef.c2 && this.nRow === byRef.r2) {
						this.ws.formulaArrayLink = null;
					}
				} else {
					var cellWithFormula = new CCellWithFormula(this.ws, this.nRow, this.nCol);
					newFP = new parserFormula(val.substring(1), cellWithFormula, this.ws);

					var formulaLocaleParse = isCopyPaste === true ? false : oFormulaLocaleInfo.Parse;
					var formulaLocaleDigetSep = isCopyPaste === true ? false : oFormulaLocaleInfo.DigitSep;
					var parseResult = new AscCommonExcel.ParseResult();
					if (!newFP.parse(formulaLocaleParse, formulaLocaleDigetSep, parseResult)) {
						switch (parseResult.error) {
							case c_oAscError.ID.FrmlWrongFunctionName:
								break;
							case c_oAscError.ID.FrmlParenthesesCorrectCount:
								return this.setValueGetParsed("=" + newFP.getFormula(), callback, isCopyPaste, byRef);
							default :
							{
								var _pasteHelper = window['AscCommon'].g_specialPasteHelper;
								var isPaste = _pasteHelper && _pasteHelper.pasteStart;
								if (isPaste) {
									if (!_pasteHelper._formulaError) {
										_pasteHelper._formulaError = true;
										wb.handlers && wb.handlers.trigger("asc_onError", parseResult.error, c_oAscError.Level.NoCritical);
									}
								} else {
									wb.handlers && wb.handlers.trigger("asc_onError", parseResult.error, c_oAscError.Level.NoCritical);
								}
								if (callback) {
									callback(false);
								}
								return;
							}
						}
					} else {
						newFP.setFormulaString(newFP.assemble());
						//***dynamic-array-formula***
						if(byRef && dynamicRange) {
							newFP.dynamicRange = dynamicRange;
							newFP.ref = byRef;
							this.ws.formulaArrayLink = newFP;
						}
						//***array-formula***
						else if(byRef) {
							newFP.ref = byRef;
							this.ws.formulaArrayLink = newFP;
						}
					}
				}
			}
		}
		return newFP;
	};
	Cell.prototype.setValue2=function(array){
		var DataOld = null;
		if(AscCommon.History.Is_On())
			DataOld = this.getValueData();
		//[{text:"",format:TextFormat},{}...]
		var xfTableAndCond = this.getCompiledStyleCustom(true, false, true);
		var oldFP = this.formulaParsed;
		this.setFormulaInternal(null);
		this.cleanText();
		this._setValue2(array, undefined, xfTableAndCond);
		this.ws.workbook.dependencyFormulas.addToChangedCell(this);
		this.ws.workbook.sortDependency();
		var DataNew = null;
		if(AscCommon.History.Is_On())
			DataNew = this.getValueData();
		if(AscCommon.History.Is_On() && false == DataOld.isEqual(DataNew)) {
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, DataOld, DataNew));
		}
		//todo не должны удаляться ссылки, если сделать merge ее части.
		if(this.isNullTextString())
		{
			var cell = this.ws.getCell3(this.nRow, this.nCol);
			cell.removeHyperlink();
		}
		this.checkRemoveExternalReferences(this.formulaParsed, oldFP);
	};
	Cell.prototype.setValueForValidation = function(array, formula, callback, isCopyPaste, byRef) {
		this.setFormulaInternal(null, true);
		this.cleanText();
		if (formula) {
			var newFP = this.setValueGetParsed(formula, callback, isCopyPaste, byRef);
			this.ws.formulaArrayLink = null;//todo
			if (newFP) {
				this.setFormulaInternal(newFP, true);
				newFP.calculate();
				this._updateCellValue();
			} else if (undefined !== newFP) {
				this._setValue(formula);
			}
		} else {
			this._setValue2(array, true);
		}
	};
	Cell.prototype.setFormulaTemplate = function(bHistoryUndo, action) {
		const formulaParsed = this.getFormulaParsed();
		let DataOld = null;
		let DataNew = null;
		if (AscCommon.History.Is_On()) {
			DataOld = this.getValueData();
		}

		if (!formulaParsed || (formulaParsed && !formulaParsed.ca)) {
			this.cleanText();
		}
		action(this);

		if (AscCommon.History.Is_On()) {
			DataNew = this.getValueData();
			if (DataOld.isEqual(DataNew) === false){
				let typeHistory = bHistoryUndo ? AscCH.historyitem_Cell_ChangeValueUndo : AscCH.historyitem_Cell_ChangeValue;
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, typeHistory, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, DataOld, DataNew), bHistoryUndo);}
		}
	};
	Cell.prototype.setFormula = function(formula, bHistoryUndo, formulaRef, caProps) {
		var cellWithFormula = new CCellWithFormula(this.ws, this.nRow, this.nCol);
		var parser = new parserFormula(formula, cellWithFormula, this.ws);
		if (caProps && caProps.ca) {
			parser.ca = caProps.ca;
		}
		if(formulaRef) {
			/* when merging, the formulaRef is taken with the old data, we need to shift ref to topLeftCell */
			// parser.setArrayFormulaRef(formulaRef);
			// this.ws.getRange3(formulaRef.r1, formulaRef.c1, formulaRef.r2, formulaRef.c2)._foreachNoEmpty(function(cell){
			// 	cell.setFormulaParsed(parser, bHistoryUndo);
			// });
			if ((this.nCol !== formulaRef.c1 || this.nRow !== formulaRef.r1) && !formulaRef.isOneCell()) {
				return;
			}
			let newFormulaRef = formulaRef.isOneCell() ? new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow) : formulaRef;
			parser.setArrayFormulaRef(newFormulaRef);
			this.ws.getRange3(newFormulaRef.r1, newFormulaRef.c1, newFormulaRef.r2, newFormulaRef.c2)._foreach2(function(cell){
				cell && cell.setFormulaParsed(parser, bHistoryUndo);
			});
		} else {
			this.setFormulaParsed(parser, bHistoryUndo, caProps);
		}
	};
	Cell.prototype.setFormulaParsed = function(parsed, bHistoryUndo, caProps) {
		this.setFormulaTemplate(bHistoryUndo, function(cell){
			cell.setFormulaInternal(parsed);
			cell.ws.workbook.dependencyFormulas.addToBuildDependencyCell(cell);
			if (caProps) {
				if (caProps.oldValue != null) {
					cell.setValueNumberInternal(caProps.oldValue);
				} else if (caProps.oldValueText != null) {
					cell.setTypeInternal(CellValueType.String);
					cell.setValueTextInternal(caProps.oldValueText);
				}
			}
		});
	};
	Cell.prototype.setFormulaInternal = function(formula, dontTouchPrev) {
		if (!dontTouchPrev && this.formulaParsed) {
			let shared = this.formulaParsed.getShared();
			let arrayFormula = this.formulaParsed.getArrayFormulaRef();
			let dynamicRange = this.formulaParsed.getDynamicRef();
			if (shared) {
				if (shared.ref.isOnTheEdge(this.nCol, this.nRow)) {
					this.ws.workbook.dependencyFormulas.addToChangedShared(this.formulaParsed);
				}
				var index = this.ws.workbook.workbookFormulas.add(this.formulaParsed).getIndexNumber();
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_RemoveSharedFormula, this.ws.getId(),
					new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, index, null), true);
			} else if(arrayFormula && this.formulaParsed.checkFirstCellArray(this)) {
				//***array-formula***
				var fText = "=" + this.formulaParsed.getFormula();
				this.formulaParsed.removeDependencies();
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoArrayFormula, AscCH.historyitem_ArrayFromula_DeleteFormula, this.ws.getId(),
					new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new AscCommonExcel.UndoRedoData_ArrayFormula(arrayFormula, fText), true);
			} else {
				this.formulaParsed.removeDependencies();
			}
		}
		this.formulaParsed = formula;
		this._hasChanged = true;
	};
	Cell.prototype.transformSharedFormula = function() {
		var res = false;
		var parsed = this.formulaParsed;
		if (parsed) {
			var shared = parsed.getShared();
			if (shared) {
				var offsetShared = new AscCommon.CellBase(this.nRow - shared.base.nRow, this.nCol - shared.base.nCol);
				var cellWithFormula = new AscCommonExcel.CCellWithFormula(this.ws, this.nRow, this.nCol);
				var newFormula = parsed.clone(undefined, cellWithFormula);
				newFormula.removeShared();
				newFormula.changeOffset(offsetShared, false);
				newFormula.setFormulaString(newFormula.assemble(true));
				this.setFormulaInternal(newFormula);
				res = true;
			}
		}
		return res;
	};
	Cell.prototype.changeOffset = function(offset, canResize, bHistoryUndo, notOffset3d) {
		this.setFormulaTemplate(bHistoryUndo, function(cell){
			cell.transformSharedFormula();
			var parsed = cell.getFormulaParsed();
			parsed.removeDependencies();
			parsed.changeOffset(offset, canResize, null, notOffset3d);
			parsed.setFormulaString(parsed.assemble(true));
			parsed.buildDependencies();
		});
	};
	Cell.prototype.setType=function(type){
		if(type != this.type){
			var DataOld = this.getValueData();
			this.setTypeInternal(type);
			this.correctValueByType();
			this._hasChanged = true;
			var DataNew = this.getValueData();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, DataOld, DataNew));
		}
		return type;
	};
	Cell.prototype.changeTextCase=function(type){
		if (this.isEmpty() || this.isFormula()) {
			return;
		}

		if(null != this.multiText) {
			let dataOld = this._cloneMultiText();
			let changedText = changeTextCase(this.multiText, type);

			if (changedText) {
				let dataNew = this._cloneMultiText();
				AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeArrayValueFormat,
					this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow),
					new UndoRedoData_CellSimpleData(this.nRow, this.nCol, dataOld, dataNew));
			}
		} else if (null != this.text) {
			let changedText = changeTextCase([{text: this.text}], type);
			if (changedText && changedText.text) {
				this.setValue(changedText.text);
			}
		}
	};
	Cell.prototype.getType=function(){
		this._checkDirty();
		return this.type;
	};
	Cell.prototype.setCellStyle=function(val){
		var oStyle;
		var newVal = this.ws.workbook.CellStyles._prepareCellStyle(val);
		var oRes = this.ws.workbook.oStyleManager.setCellStyle(this, newVal);
		if(AscCommon.History.Is_On()) {
			var oldStyleName = this.ws.workbook.CellStyles.getStyleNameByXfId(oRes.oldVal);
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Style, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oldStyleName, val));

			// Выставляем стиль
			oStyle = this.ws.workbook.CellStyles.getStyleByXfId(oRes.newVal);
			if (oStyle.ApplyFont)
				this.setFont(oStyle.getFont());
			if (oStyle.ApplyFill)
				this.setFill(oStyle.getFill());
			if (oStyle.ApplyBorder)
				this.setBorder(oStyle.getBorder());
			if (oStyle.ApplyNumberFormat)
				this.setNumFormat(oStyle.getNumFormatStr());
		}
		return oStyle;
	};
	Cell.prototype.setNumFormat=function(val){
		this.setNum(new AscCommonExcel.Num({f:val}));
	};
	Cell.prototype.setNum=function(val){
		var oRes = this.ws.workbook.oStyleManager.setNum(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Num, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.getShiftedNumFormat=function(nShift, dDigitsCount){
		let newNumFormat;
		var sNumFormat;
		if(null != this.xfs && null != this.xfs.num)
			sNumFormat = this.xfs.num.getFormat();
		else
			sNumFormat = g_oDefaultFormat.Num.getFormat();
		var type = this.getType();
		var oCurNumFormat = oNumFormatCache.get(sNumFormat);
		if (null != oCurNumFormat && !oCurNumFormat.isGeneralFormat()) {
			let output = {};
			if (oCurNumFormat.shiftFormat(output, nShift)) {
				newNumFormat = output.format;
			}
		} else if (CellValueType.Number == type) {
			var sGeneral = AscCommon.DecodeGeneralFormat(this.number, type, dDigitsCount);
			var oGeneral = oNumFormatCache.get(sGeneral);
			if (null != oGeneral && !oGeneral.isGeneralFormat()) {
				let output = {};
				if (oGeneral.shiftFormat(output, nShift)) {
					newNumFormat = output.format;
				}
			}
		}
		return newNumFormat;
	};
	Cell.prototype.setFont=function(val, bModifyValue){
		if(false != bModifyValue)
		{
			//убираем комплексные строки
			if(null != this.multiText && false == this.ws.workbook.bUndoChanges && false == this.ws.workbook.bRedoChanges)
			{
				var oldVal = null;
				if(AscCommon.History.Is_On())
					oldVal = this.getValueData();
				this.setValueTextInternal(AscCommonExcel.getStringFromMultiText(this.multiText));
				if(AscCommon.History.Is_On())
				{
					var newVal = this.getValueData();
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oldVal, newVal));
				}
			}
		}
		var oRes = this.ws.workbook.oStyleManager.setFont(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
		{
			var oldVal = null;
			if(null != oRes.oldVal)
				oldVal = oRes.oldVal.clone();
			var newVal = null;
			if(null != oRes.newVal)
				newVal = oRes.newVal.clone();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetFont, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oldVal, newVal));
		}
	};
	Cell.prototype.setFontname=function(val){
		var oRes = this.ws.workbook.oStyleManager.setFontname(this, val);
		this._setFontProp(function(format){return val != format.getName();}, function(format){format.setName(val);});
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Fontname, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setFontsize=function(val){
		var oRes = this.ws.workbook.oStyleManager.setFontsize(this, val);
		this._setFontProp(function(format){return val != format.getSize();}, function(format){format.setSize(val);});
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Fontsize, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setFontcolor=function(val){
		var oRes = this.ws.workbook.oStyleManager.setFontcolor(this, val);
		this._setFontProp(function(format){return val != format.getColor();}, function(format){format.setColor(val);});
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Fontcolor, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setBold=function(val){
		var oRes = this.ws.workbook.oStyleManager.setBold(this, val);
		this._setFontProp(function(format){return val != format.getBold();}, function(format){format.setBold(val);});
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Bold, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setItalic=function(val){
		var oRes = this.ws.workbook.oStyleManager.setItalic(this, val);
		this._setFontProp(function(format){return val != format.getItalic();}, function(format){format.setItalic(val);});
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Italic, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setUnderline=function(val){
		var oRes = this.ws.workbook.oStyleManager.setUnderline(this, val);
		this._setFontProp(function(format){return val != format.getUnderline();}, function(format){format.setUnderline(val);});
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Underline, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setStrikeout=function(val){
		var oRes = this.ws.workbook.oStyleManager.setStrikeout(this, val);
		this._setFontProp(function(format){return val != format.getStrikeout();}, function(format){format.setStrikeout(val);});
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Strikeout, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setFontAlign=function(val){
		var oRes = this.ws.workbook.oStyleManager.setFontAlign(this, val);
		this._setFontProp(function(format){return val != format.getVerticalAlign();}, function(format){format.setVerticalAlign(val);});
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_FontAlign, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setAlignVertical=function(val){
		var oRes = this.ws.workbook.oStyleManager.setAlignVertical(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_AlignVertical, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setAlignHorizontal=function(val){
		var oRes = this.ws.workbook.oStyleManager.setAlignHorizontal(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_AlignHorizontal, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setFill=function(val){
		var oRes = this.ws.workbook.oStyleManager.setFill(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Fill, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setFillColor=function(val){
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(val);
		return this.setFill(fill);
	};
	Cell.prototype.setBorder=function(val){
		var oRes = this.ws.workbook.oStyleManager.setBorder(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal){
			var oldVal = null;
			if(null != oRes.oldVal)
				oldVal = oRes.oldVal.clone();
			var newVal = null;
			if(null != oRes.newVal)
				newVal = oRes.newVal.clone();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Border, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oldVal, newVal));
		}
	};
	Cell.prototype.setShrinkToFit=function(val){
		var oRes = this.ws.workbook.oStyleManager.setShrinkToFit(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ShrinkToFit, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setWrap=function(val){
		var oRes = this.ws.workbook.oStyleManager.setWrap(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Wrap, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setAngle=function(val){
		var oRes = this.ws.workbook.oStyleManager.setAngle(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Angle, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setIndent=function(val){
		var oRes = this.ws.workbook.oStyleManager.setIndent(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_Indent, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setQuotePrefix=function(val){
		var oRes = this.ws.workbook.oStyleManager.setQuotePrefix(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetQuotePrefix, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setApplyProtection=function(val){
		var oRes = this.ws.workbook.oStyleManager.setApplyProtection(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetApplyProtection, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setLocked=function(val){
		var oRes = this.ws.workbook.oStyleManager.setLocked(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetLocked, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setHiddenFormulas=function(val){
		var oRes = this.ws.workbook.oStyleManager.setHiddenFormulas(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetHidden, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setPivotButton=function(val){
		var oRes = this.ws.workbook.oStyleManager.setPivotButton(this, val);
		if(AscCommon.History.Is_On() && oRes.oldVal != oRes.newVal)
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetPivotButton, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oRes.oldVal, oRes.newVal));
	};
	Cell.prototype.setStyle=function(xfs){
		var oldVal = this.xfs;
		this.setStyleInternal(xfs);
		if(AscCommon.History.Is_On() && oldVal != this.xfs)
		{
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_SetStyle, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, oldVal, this.xfs));
		}
	};
	Cell.prototype.setStyleInternal = function(xfs) {
		this.xfs = g_StyleCache.addXf(xfs);
		this._hasChanged = true;
	};
	Cell.prototype.getFormula=function(){
		var res = "";
		this.processFormula(function(parsed) {
			res = parsed.getFormula();
		});
		return res;
	};
	Cell.prototype.getFormulaParsed=function(){
		return this.formulaParsed;
	};
	Cell.prototype.getValueForEdit = function(checkFormulaArray) {
		this._checkDirty();
		//todo
		// if (CellValueType.Error == this.getType()) {
		// 	return this._getValueTypeError(textValueForEdit);
		// }
		if ((this.ws && this.ws.getSheetProtection() && this.xfs && this.xfs.getHidden()) ||
			(this.ws.isUserProtectedRangesIntersectionCell(this, null, null, Asc.c_oSerUserProtectedRangeType.View))
			|| !this.ws.isUserProtectedRangesCanView(this, null)) {
			return "";
		}

		var res = AscCommonExcel.getStringFromMultiText(this.getValueForEdit2());
		return this.formulaParsed && this.formulaParsed.ref && checkFormulaArray ? "{" + res + "}" : res;
	};
	Cell.prototype.getValueForEdit2 = function() {
		this._checkDirty();
		var cultureInfo = AscCommon.g_oDefaultCultureInfo;
		//todo проблема точности. в excel у чисел только 15 значащих цифр у нас больше.
		//применяем форматирование
		var oValueText = null;
		var oValueArray = null;
		var xfs = this.getCompiledStyle();
		if (this.isFormula()) {
			if (!((this.ws && this.ws.getSheetProtection() && this.xfs && this.xfs.getHidden() ||
				(this.ws.isUserProtectedRangesIntersectionCell(this, null, null, Asc.c_oSerUserProtectedRangeType.View))))) {
				this.processFormula(function (parsed) {
					// ToDo если будет притормаживать, то завести переменную и не рассчитывать каждый раз!
					oValueText = "=" + parsed.assembleLocale(AscCommonExcel.cFormulaFunctionToLocale, true);
				});
			}
		} else {
			if(null != this.text || null != this.number)
			{
				if (CellValueType.Bool === this.type && null != this.number)
					oValueText = (this.number == 1) ? cBoolLocal.t : cBoolLocal.f;
				else
				{
					if(null != this.text)
						oValueText = this.text;
					if(CellValueType.Number === this.type || CellValueType.String === this.type)
					{
						var oNumFormat;
						if(null != xfs && null != xfs.num)
							oNumFormat = oNumFormatCache.get(xfs.num.getFormat());
						else
							oNumFormat = oNumFormatCache.get(g_oDefaultFormat.Num.getFormat());
						if(CellValueType.String != this.type && null != oNumFormat && null != this.number)
						{
							var nValue = this.number;
							var oTargetFormat = oNumFormat.getFormatByValue(nValue);
							if(oTargetFormat)
							{
								if(1 == oTargetFormat.nPercent)
								{
									//prercent
									oValueText = AscCommon.oGeneralEditFormatCache.format(nValue * 100) + "%";
								}
								else if(oTargetFormat.bDateTime)
								{
									//Если число не подходит под формат даты возвращаем само число
									if(false == oTargetFormat.isInvalidDateValue(nValue))
									{
										var bDate = oTargetFormat.bDate;
										var bTime = oTargetFormat.bTime;
										if(false == bDate && nValue >= 1)
											bDate = true;
										if(false == bTime && Math.floor(nValue) != nValue)
											bTime = true;
										var sDateFormat = "";
										if (bDate) {
											sDateFormat = AscCommon.getShortDateFormat(cultureInfo);
										}
										var sTimeFormat = 'h:mm:ss';

										if (AscCommon.is12HourTimeFormat(cultureInfo)){
											sTimeFormat += ' AM/PM';
										}
										if(bDate && bTime)
											oNumFormat = oNumFormatCache.get(sDateFormat + ' ' + sTimeFormat);
										else if(bTime)
											oNumFormat = oNumFormatCache.get(sTimeFormat);
										else
											oNumFormat = oNumFormatCache.get(sDateFormat);

										var aFormatedValue = oNumFormat.format(nValue, CellValueType.Number, AscCommon.gc_nMaxDigCount);
										oValueText = "";
										for(var i = 0, length = aFormatedValue.length; i < length; ++i)
											oValueText += aFormatedValue[i].text;
									}
									else
										oValueText = AscCommon.oGeneralEditFormatCache.format(nValue);
								}
								else
									oValueText = AscCommon.oGeneralEditFormatCache.format(nValue);
							}
						}
					}
				}
			}
			else if(this.multiText)
				oValueArray = this.multiText;
		}
		if(null != xfs && true == xfs.QuotePrefix && CellValueType.String == this.type && false == this.isFormula())
		{
			if(null != oValueText)
				oValueText = "'" + oValueText;
			else if(null != oValueArray)
				oValueArray = [{text:"'"}].concat(oValueArray);
		}
		return this._getValue2Result(oValueText, oValueArray, true);
	};
	Cell.prototype.getValueForExample = function(numFormat, cultureInfo) {
		var aText = this._getValue2(AscCommon.gc_nMaxDigCountView, function(){return true;}, numFormat, cultureInfo, true);
		return AscCommonExcel.getStringFromMultiText(aText);
	};
	Cell.prototype.getValueWithoutFormat = function() {
		this._checkDirty();
		var sResult = "";
		if(null != this.number)
		{
			if(CellValueType.Bool === this.type)
				sResult = this.number == 1 ? cBoolLocal.t : cBoolLocal.f;
			else
				sResult = this.number.toString();
		}
		else if(null != this.text)
			sResult = this.text;
		else if(null != this.multiText)
			sResult = AscCommonExcel.getStringFromMultiText(this.multiText);
		return sResult;
	};
	Cell.prototype.getValue = function() {
		this._checkDirty();
		let aTextValue2 = this._getValue2(AscCommon.gc_nMaxDigCountView, function() {return true;},
			undefined, undefined, true);
		return AscCommonExcel.getStringFromMultiText(aTextValue2);
	};
	Cell.prototype.getValueSkipToSpace = function() {
		this._checkDirty();
		let aTextValue2 = this._getValue2(AscCommon.gc_nMaxDigCountView, function() {return true;},
			undefined, undefined, true);
		return AscCommonExcel.getStringFromMultiTextSkipToSpace(aTextValue2);
	};
	Cell.prototype.getValue2 = function(dDigitsCount, fIsFitMeasurer) {
		this._checkDirty();
		if(null == fIsFitMeasurer)
			fIsFitMeasurer = function(aText){return true;};
		if(null == dDigitsCount)
			dDigitsCount = AscCommon.gc_nMaxDigCountView;
		return this._getValue2(dDigitsCount, fIsFitMeasurer);
	};
	Cell.prototype.getNumberValue = function() {
		this._checkDirty();
		return this.number;
	};
	Cell.prototype.getBoolValue = function() {
		this._checkDirty();
		return this.number == 1;
	};
	Cell.prototype.getErrorValue = function() {
		this._checkDirty();
		return AscCommonExcel.cError.prototype.getErrorTypeFromString(this.text);
	};
	Cell.prototype.getValueText = function() {
		this._checkDirty();
		return this.text;
	};
	Cell.prototype.getTextIndex = function() {
		if (null === this.textIndex) {
			var wb = this.ws.workbook;
			if (null != this.text) {
				this.textIndex = wb.sharedStrings.addText(this.text);
			} else if (null != this.multiText) {
				this.textIndex = wb.sharedStrings.addMultiText(this.multiText);
			}
		}
		return this.textIndex;
	};
	Cell.prototype.getValueMultiText = function() {
		this._checkDirty();
		return this.multiText;
	};
	Cell.prototype.getNumFormatStr=function(){
		if(null != this.xfs && null != this.xfs.num)
			return this.xfs.num.getFormat();
		return g_oDefaultFormat.Num.getFormat();
	};
	Cell.prototype.getNumFormat=function(){
		return oNumFormatCache.get(this.getNumFormatStr());
	};
	Cell.prototype.getNumFormatType=function(){
		return this.getNumFormat().getType();
	};
	Cell.prototype.getOffset=function(cell){
		return this.getOffset3(cell.nCol + 1, cell.nRow + 1);
	};
	Cell.prototype.getOffset2=function(cellId){
		var cAddr2 = new CellAddress(cellId);
		return this.getOffset3(cAddr2.col, cAddr2.row);
	};
	Cell.prototype.getOffset3=function(col, row){
		return new AscCommon.CellBase(this.nRow - row + 1, this.nCol - col + 1);
	};
	Cell.prototype.getValueData = function(){
		this._checkDirty();
		var formula = this.isFormula() ? this.getFormula() : null;
		var formulaRef;
		let ca;
		if(formula) {
			var parser = this.getFormulaParsed();
			if(parser) {
				formulaRef = parser.getArrayFormulaRef();
				ca = parser.ca;
			}
		}
		return new UndoRedoData_CellValueData(formula, new AscCommonExcel.CCellValue(this), formulaRef, ca);
	};
	Cell.prototype.setValueData = function(Val){
		//значения устанавляваются через setValue, чтобы пересчитались формулы
		if (null != Val.formula) {
			let caProps = {
				ca: Val.ca,
				oldValue: Val.value.number,
				oldValueText: Val.value.text
			};
			this.setFormula(Val.formula, null, Val.formulaRef, caProps);
		} else if (null != Val.value) {
			var DataOld = null;
			var DataNew = null;
			if (AscCommon.History.Is_On())
				DataOld = this.getValueData();
			if (this.formulaParsed && this.formulaParsed.getDynamicRef()) {
				this.ws.workbook.dependencyFormulas.addToChangedRange2(this.formulaParsed.getWs().getId(), this.formulaParsed.getDynamicRef());
				if (Val.value.getTextValue()) {
					// non empty val, delete PF from cell
					this.setFormulaInternal(null);
				}
			} else {
				this.setFormulaInternal(null);
				this.ws.workbook.dependencyFormulas.addToChangedCell(this);
			}

			this._setValueData(Val.value);
			this.ws.workbook.sortDependency();
			if (AscCommon.History.Is_On()) {
				DataNew = this.getValueData();
				if (false == DataOld.isEqual(DataNew))
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, DataOld, DataNew));
			}
		}
		else
			this.setValue("");

		if (this.ws.workbook.handlers) {
			this.ws.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.cellValue, this, null, this.ws.getId());
		}
	};
	Cell.prototype.getQuotePrefix = function () {
		return this.xfs && this.xfs.getQuotePrefix();
	};

	/**
	 * Checks whether the formula is an exclude formula, which doesn't need recursive check.
	 * @param {[]} aOutStack
	 * @param {string} sFunctionName
	 * @returns {boolean}
	 * @private
	 */
	function _isExcludeFormula(aOutStack, sFunctionName) {
		const aExcludeFormulas = AscCommonExcel.aExcludeRecursiveFormulas;
		const aUnOperators = ['un_minus', 'un_plus'];
		let nFuncCount = 0;

		if (!sFunctionName) {
			return false;
		}
		if (!aExcludeFormulas.includes(sFunctionName)) {
			return false;
		}
		aOutStack.forEach(function (oElem) {
			if (oElem.type === cElementType.func || (oElem.type === cElementType.operator && !aUnOperators.includes(oElem.name))) {
				nFuncCount++;
			}
		});
		if (nFuncCount === 1 && aExcludeFormulas.includes(sFunctionName)) {
			if (sFunctionName === 'CELL') {
				let oStringElem, nCountArgs = null;
				aOutStack.forEach(function (oElem) {
					if (oElem.type === cElementType.string && !oStringElem) {
						oStringElem = oElem;
					} else if (nCountArgs === null && typeof oElem === 'number') {
						nCountArgs = oElem;
					}
				});
				const sValue = oStringElem && oStringElem.toString().toLowerCase();
				if ((sValue && sValue === 'contents') && (nCountArgs &&  nCountArgs === 2)) {
					return false;
				}
			}
			return true;
		}

		return false;
	}
	/**
	 * Checks "areaMap" belongs formula that ignores recursion rules
	 * @param {[]} aOutStack
	 * @param {object} oAreaMap
	 * @returns {boolean}
	 * @private
	 */
	function _areaMapIsExcludeFormula(aOutStack, oAreaMap) {
		const aExcludeFormulas = AscCommonExcel.aExcludeRecursiveFormulas;
		for (let i = 0, length = aOutStack.length; i < length; i++) {
			const oElem = aOutStack[i];
			if (oElem.type === cElementType.func && aExcludeFormulas.includes(oElem.name)) {
				let nIndexOfArgCount = i - 1;
				let nArgCount = aOutStack[nIndexOfArgCount];
				let nStartArgIndex = nIndexOfArgCount - nArgCount;
				let bBelongToFormula = false;
				for (let j = nStartArgIndex; j < nIndexOfArgCount; j++) {
					let oArgElem = aOutStack[j];
					let oBbox = null;
					let aRef = [cElementType.cell, cElementType.cell3D, cElementType.cellsRange, cElementType.cellsRange3D];
					if (!aRef.includes(oArgElem.type)) {
						continue;
					}
					if (oArgElem.type === cElementType.name || oArgElem.type === cElementType.name3D) {
						const oElemValue = oArgElem.getValue();
						if (!aRef.includes(oElemValue.type)) {
							continue;
						}
						oBbox = oElemValue.getBBox0();
					}  else {
						oBbox = oArgElem.getBBox0();
					}
					bBelongToFormula = oAreaMap.bbox.isIntersect(oBbox);
					if (bBelongToFormula) {
						return true;
					}
				}
			}
		}
		return false;
	}
	/**
	 * Method returns listeners of cell
	 * @returns {object}
	 */
	Cell.prototype.getListeners = function () {
		const ws = this.ws;
		const oDepFormulas = ws.workbook.dependencyFormulas;

		if (!oDepFormulas || !oDepFormulas.sheetListeners.hasOwnProperty(ws.Id)) {
			return null;
		}

		const nCellIndex = getCellIndex(this.nRow, this.nCol);
		const oFormulaParsed = this.getFormulaParsed();
		const sFunctionName = oFormulaParsed && oFormulaParsed.getFunctionName();
		if (oFormulaParsed && !oFormulaParsed.ca && !g_cCalcRecursion.getFunctionsResult().length) {
			return null;
		}
		const aOutStack = oFormulaParsed && oFormulaParsed.outStack;
		const oSheetListeners = oDepFormulas.sheetListeners[ws.Id];

		if (oSheetListeners.cellMap.hasOwnProperty(nCellIndex) && !_isExcludeFormula(aOutStack, sFunctionName)) {
			return oSheetListeners.cellMap[nCellIndex];
		} else if (aOutStack && aOutStack.length && sFunctionName) {
			for (let nIndex in oSheetListeners.areaMap) {
				if (oSheetListeners.areaMap[nIndex].bbox.contains(this.nCol, this.nRow)) {
					return _areaMapIsExcludeFormula(aOutStack, oSheetListeners.areaMap[nIndex]) ? null : oSheetListeners.areaMap[nIndex];
				}
			}
		}
	};

	/**
	 * Iterative extract listeners of a cell
	 * @param fAction {Function} Action on the listener.
	 * @param oCell {Cell}
	 * @param oCellListeners {object}
	 * @private
	 */
	function _foreachListeners(fAction, oCell, oCellListeners) {
		const oListeners = oCellListeners.listeners;
		for (let i in oListeners) {
			let oListenerCell = oListeners[i].getParent();
			if (!oListenerCell) {
				continue;
			}
			let nListenerCellIndex = null;
			if (oListenerCell instanceof DefName) {
				if (!oListeners[i].ca) {
					continue;
				}
				let oParserRef = oListenerCell.parsedRef;
				let aRef = [cElementType.cell, cElementType.cell3D, cElementType.cellsRange, cElementType.cellsRange3D];
				let oOutStackElem = oParserRef.outStack.find(function (oElem) {
					return aRef.includes(oElem.type);
				});
				let oRange = oOutStackElem && oOutStackElem.getRange();
				if (oRange && oRange.isOneCell()) {
					nListenerCellIndex = getCellIndex(oRange.bbox.r1, oRange.bbox.c1);
					let oWs = oRange.worksheet;
					oWs._getCell(oRange.bbox.r1, oRange.bbox.c1, function (oCell) {
						oListenerCell = oCell;
					});
				} else if (oRange) {
					oRange._foreachNoEmpty(function (oCell) {
						if (oCell.isFormula()) {
							nListenerCellIndex = getCellIndex(oCell.nRow, oCell.nCol);
							oListenerCell = oCell;
							return true; // break loop;
						}
					})
				}
			} else if (oListeners[i].getShared() && !oListeners[i].is3D && oCellListeners.bbox) {
				const oShared = oListeners[i].getShared();
				const oCurrentRange = new Asc.Range(oCell.nCol, oCell.nRow, oCell.nCol, oCell.nRow);
				const oListenerSharedRange = oCellListeners.bbox.getSharedIntersect(oShared.ref, oCurrentRange);
				oListenerCell.ws._getCell(oListenerSharedRange.r1, oListenerSharedRange.c1, function (oCell) {
					oListenerCell = oCell;
				});
				nListenerCellIndex = getCellIndex(oListenerCell.nRow, oListenerCell.nCol);
			} else {
				nListenerCellIndex = getCellIndex(oListenerCell.nRow, oListenerCell.nCol);
			}
			if (nListenerCellIndex == null || isNaN(nListenerCellIndex)) {
				continue;
			}
			if (oListenerCell instanceof Asc.CT_WorksheetSource) {
				continue;
			}
			if (oListeners[i].getFunctionName() && oCell.compareCellIndex(oListenerCell)) {
				const aOutStack = oListeners[i].outStack;
				const aExcludeFormulas = AscCommonExcel.aExcludeRecursiveFormulas;
				const oCycleCells = new Map();

				foreachRefElements(function (oElem, nIndex, nLastIndex) {
					if (oElem.containCell2(oListenerCell)) {
						const nIndexWithFunction = nIndex + 2;
						if (nIndexWithFunction > nLastIndex) {
							return;
						}
						const oElemWithFunction = aOutStack[nIndexWithFunction];
						if (oElemWithFunction && oElemWithFunction.type === cElementType.func && aExcludeFormulas.includes(oElemWithFunction.name)) {
							oCycleCells.set(true, nIndex);
							return;
						}
						oCycleCells.set(false, nIndex);
					}
				}, aOutStack);
				if (oCycleCells.size > 0 && oCycleCells.get(false) === undefined) {
					continue;
				}
			}
			let oRes = fAction(oListenerCell, oCell, nListenerCellIndex);
			if (oRes != null) {
				return;
			}
		}
	}

	/**
	 * Compares cell index of two cells.
	 * @memberof Cell
	 * @param {Cell|CCellWithFormula} oCompareCell
	 * @returns {boolean}
	 */
	Cell.prototype.compareCellIndex = function (oCompareCell) {
		const nCellIndex = getCellIndex(this.nRow, this.nCol);
		const nCompareCellIndex = getCellIndex(oCompareCell.nRow, oCompareCell.nCol);
		const sWsName = this.ws.getName();
		const sCompareWsName = oCompareCell.ws.getName();

		return nCellIndex === nCompareCellIndex && sWsName === sCompareWsName;
	};
	/**
	 * Checks cell contains in formula.
	 * @memberof Cell
	 * @returns {boolean}
	 */
	Cell.prototype.containInFormula = function () {
		const oThis = this;
		const formulaParsed = this.getFormulaParsed();
		const aOutStack = formulaParsed && formulaParsed.outStack;
		if (!aOutStack) {
			return false;
		}
		let bContainsInFormula = false;
		foreachRefElements(function (oRange) {
			if (oRange.containCell2(oThis)) {
				bContainsInFormula = true;
				return true;
			}
		}, aOutStack);

		return bContainsInFormula;
	};
	/**
	 * Method finds a start cell index of recursion formula.
	 * Uses for iterative calculations feature.
	 * @param {Cell} [oPrevCell] - Previous cell in recursion.
	 * @param {Cell} [oFirstCell] - First cell in recursion.
	 * @param {Cell[]} [aPassedCells] - Passed cells in recursion.
	 * @memberof Cell
	 */
	Cell.prototype.initStartCellForIterCalc = function (oPrevCell, oFirstCell, aPassedCells) {
		if (g_cCalcRecursion.checkRecursionCounter()) {
			g_cCalcRecursion.resetRecursionCounter();
			return;
		}

		const oParserFormula = this.getFormulaParsed();
		const sFunctionName = oParserFormula && oParserFormula.getFunctionName();
		if (oParserFormula && oParserFormula._isConditionalFormula(sFunctionName) && this.containInFormula() && !oParserFormula.isRecursiveCondFormula(sFunctionName)) {
			g_cCalcRecursion.resetRecursionCounter();
			return;
		}
		const oCellListeners = this.getListeners();
		if (!oCellListeners) {
			g_cCalcRecursion.resetRecursionCounter();
			return;
		}
		const nCellIndex = getCellIndex(this.nRow, this.nCol);
		let oCellFromListener = null;
		let nPrevCellIndex = oPrevCell ? getCellIndex(oPrevCell.nRow, oPrevCell.nCol) : null;
		oFirstCell = oFirstCell || this;
		let nFirstCellIndex = getCellIndex(oFirstCell.nRow, oFirstCell.nCol);
		aPassedCells = aPassedCells || [];
		let bBreakFunction = false
		_foreachListeners(function (oListenerCell, oThis) {
			if (oThis.compareCellIndex(oListenerCell) && oThis.containInFormula()) {
				g_cCalcRecursion.setStartCellIndex({cellId: nCellIndex, wsName: oThis.ws.getName().toLowerCase()});
				g_cCalcRecursion.resetRecursionCounter();
				return true;
			} else if (nPrevCellIndex != null && oPrevCell.compareCellIndex(oListenerCell)) {
				let sPrevCellWsName = oPrevCell.ws.getName().toLowerCase();
				let sThisCellWsName = oThis.ws.getName().toLowerCase();
				let sCellProp = sPrevCellWsName !== sThisCellWsName ? sThisCellWsName : nCellIndex.toString();
				let sPrevCellProp = sPrevCellWsName !== sThisCellWsName ? sPrevCellWsName : nPrevCellIndex.toString();
				if (sPrevCellProp < sCellProp) {
					g_cCalcRecursion.setStartCellIndex({cellId: nPrevCellIndex, wsName: sPrevCellWsName});
				} else {
					g_cCalcRecursion.setStartCellIndex({cellId: nCellIndex, wsName: sThisCellWsName});
				}
				g_cCalcRecursion.resetRecursionCounter();
				return true;
			} else if (oFirstCell.compareCellIndex(oListenerCell) && g_cCalcRecursion.getRecursionCounter() > 0) {
				g_cCalcRecursion.setStartCellIndex({cellId: nFirstCellIndex, wsName: oFirstCell.ws.getName().toLowerCase()});
				g_cCalcRecursion.resetRecursionCounter();
				return true;
			} else {
				let oWs = oListenerCell.ws;
				oWs._getCell(oListenerCell.nRow, oListenerCell.nCol, function (oCell) {
					oCellFromListener = oCell;
				});
				if (oCellFromListener.compareCellIndex(oThis)) {
					g_cCalcRecursion.resetRecursionCounter();
					bBreakFunction = true;
					return;
				}
				let bPassedCell = aPassedCells.some(function (oPassedCell) {
					let nPassedCellIndex = getCellIndex(oPassedCell.nRow, oPassedCell.nCol);
					let nCellIndex = getCellIndex(oCellFromListener.nRow, oCellFromListener.nCol);
					return nPassedCellIndex === nCellIndex && oPassedCell.ws.getName() === oCellFromListener.ws.getName();
				});
				if (bPassedCell) {
					g_cCalcRecursion.resetRecursionCounter();
					return;
				}
				aPassedCells.push(oCellFromListener);
				g_cCalcRecursion.incRecursionCounter();
				oCellFromListener.initStartCellForIterCalc(oThis, oFirstCell, aPassedCells);
				g_cCalcRecursion.resetRecursionCounter();
				if (g_cCalcRecursion.getStartCellIndex() != null) {
					return true;
				}
			}
		}, this, oCellListeners);
		if (!g_cCalcRecursion.getStartCellIndex() && !bBreakFunction) {
			let oTableStructOperand = oParserFormula && oParserFormula.outStack.find(function (oOperand) {
				return oOperand.type === cElementType.table;
			});
			if (oTableStructOperand) {
				let oTableOpRange = oTableStructOperand.getRange();
				if (oTableOpRange.bbox && !oTableOpRange.bbox.contains(this.nCol, this.nRow)) {
					return;
				}
				let bContainRange = oCellListeners.bbox && oCellListeners.bbox.containsRange(oTableOpRange.bbox);
				if (bContainRange && oCellListeners.bbox.contains(this.nCol, this.nRow)) {
					g_cCalcRecursion.setStartCellIndex({cellId: nCellIndex, wsName: this.ws.getName().toLowerCase()});
				}
			}
		}
	};

	/**
	 * Function returns an array of reference cells from outStack attribute of parserFormula class.
	 * @param {parserFormula} oFormulaParsed
	 * @returns {[]}
	 * @private
	 */
	function _getRefElements(oFormulaParsed) {
		const aRefElements = [];
		for (let i = 0, length = oFormulaParsed.getOutStackSize(); i < length; i++) {
			const oOutStackElem = oFormulaParsed.getOutStackElem(i);
			const nElemType = oOutStackElem.type;
			const b3D = nElemType === cElementType.cell3D || nElemType === cElementType.cellsRange3D || nElemType === cElementType.name3D;
			const bArea = nElemType === cElementType.cellsRange || nElemType === cElementType.name;
			const bDefName = nElemType === cElementType.name || nElemType === cElementType.name3D;
			const bTable = nElemType === cElementType.table;

			if (nElemType === cElementType.cell || b3D || bArea || bDefName || bTable) {
				aRefElements.push(oOutStackElem);
			}
		}

		return aRefElements;
	}

	/**
	 * Iterative extract reference elements from outStack attribute of parserFormula class.
	 * @param {Function} fAction - Action with reference element
	 * @param {[]} aRefElements
	 */
	function foreachRefElements (fAction, aRefElements) {
		if (!aRefElements.length) {
			return;
		}
		let nLastIndex = aRefElements.length - 1;
		for (let i = 0; i < aRefElements.length; i++) {
			const oRefElement = aRefElements[i];
			const nRefType = oRefElement.type;
			const b3D = nRefType === cElementType.cell3D || nRefType === cElementType.cellsRange3D;
			const bArea = nRefType === cElementType.cellsRange;
			const bDefName = nRefType === cElementType.name || nRefType === cElementType.name3D;
			const bTable = nRefType === cElementType.table;
			let oRange = null;

			if (nRefType !== cElementType.cell && !b3D && !bArea && !bDefName && !bTable) {
				continue;
			}
			if (nRefType === cElementType.name || nRefType === cElementType.name3D) {
				const oElemValue = oRefElement.getValue();
				const oElemType = oElemValue.type;
				let aRef = [cElementType.cell, cElementType.cell3D, cElementType.cellsRange, cElementType.cellsRange3D];
				if (!aRef.includes(oElemType)) {
					return;
				}
				oRange = oElemValue.getRange();
			} else if (nRefType === cElementType.table) {
				let oRefElem = oRefElement.toRef();
				oRange = oRefElem.getRange();
			} else {
				oRange = oRefElement.getRange();
			}
			let oRes = fAction(oRange, i, nLastIndex);
			if (oRes != null) {
				return;
			}
		}
	}

	/**
	 * Changes ca flag of cell from listener.
	 * Function is recursive.
	 * @param {Cell} oSourceCell
	 * @private
	 */
	function _changeCellsFromListener (oSourceCell) {
		if (g_cCalcRecursion.checkRecursionCounter()) {
			g_cCalcRecursion.resetRecursionCounter();
			return;
		}
		const oCellListeners = oSourceCell.getListeners();
		if (!oCellListeners) {
			g_cCalcRecursion.resetRecursionCounter();
			return;
		}
		_foreachListeners(function (oListenerCellWithFormula, oCell) {
			const oWs = oCell.ws;
			let oListenerCell = null;
			oWs._getCell(oListenerCellWithFormula.nRow, oListenerCellWithFormula.nCol, function (oCell) {
				oListenerCell = oCell;
			})
			if (!oListenerCell.isFormula()) {
				return;
			}
			const oFormulaParsed = oListenerCell.getFormulaParsed();
			if (oFormulaParsed.ca) {
				return;
			}
			oFormulaParsed.ca = true;
			if (oListenerCell.getListeners()) {
			g_cCalcRecursion.incRecursionCounter();
			_changeCellsFromListener(oListenerCell);
			}
		}, oSourceCell, oCellListeners);
		g_cCalcRecursion.resetRecursionCounter();
	}

	/**
	 * Method checks is cell has recursive formula.
	 * Recursion method
	 * @param {CCellWithFormula} oCellWithFormula
	 * @param {Cell[]} [aPassedCell]
	 * @returns {boolean}
	 */
	Cell.prototype.checkRecursiveFormula = function (oCellWithFormula, aPassedCell) {
		if (g_cCalcRecursion.checkRecursionCounter()) {
			g_cCalcRecursion.resetRecursionCounter();
			return false;
		}
		if (oCellWithFormula == null) {
			g_cCalcRecursion.resetRecursionCounter();
			return false;
		}

		const nCellIndexFromFormula = getCellIndex(oCellWithFormula.nRow, oCellWithFormula.nCol);
		if (isNaN(nCellIndexFromFormula)) {
			g_cCalcRecursion.resetRecursionCounter();
			return false;
		}
		if (this.compareCellIndex(oCellWithFormula)) {
			g_cCalcRecursion.resetRecursionCounter();
			return true;
		}

		if (!this.isFormula()) {
			g_cCalcRecursion.resetRecursionCounter();
			return false;
		}
		const oFormulaParsed = this.getFormulaParsed();
		const aRefElements = _getRefElements(oFormulaParsed);
		const oThis = this;
		let bRecursiveFormula = false;
		aPassedCell = aPassedCell || [];
		foreachRefElements(function (oRange) {
			oRange._foreachNoEmpty(function (oCell) {
				let bCellIsPassed = aPassedCell.some(function (oElem) {
					return oCell.compareCellIndex(oElem);
				});
				if (oCell.compareCellIndex(oCellWithFormula)) {
					bRecursiveFormula = true;
					return true; // break loop
				} else if (!bRecursiveFormula && oCell.isFormula() && !bCellIsPassed) {
					g_cCalcRecursion.incRecursionCounter();
					aPassedCell.push(oCell);
					bRecursiveFormula = oCell.checkRecursiveFormula(oCellWithFormula, aPassedCell);
				}
				if (bRecursiveFormula) {
					return true; // break loop
				}
			});
			if (bRecursiveFormula) {
				if (g_cCalcRecursion.getRecursionCounter() === 0 && aPassedCell.length) {
					const oWs = oThis.ws;
					let oSourceCell = null;
					oWs._getCell(oCellWithFormula.nRow, oCellWithFormula.nCol, function (oCell) {
						oSourceCell = oCell;
					});
					_changeCellsFromListener(oSourceCell);
				}
				oFormulaParsed.ca = true;
				return true;
			}
		}, aRefElements);

		g_cCalcRecursion.resetRecursionCounter();
		return bRecursiveFormula;
	};
	/**
	 * Method modifies linked cell.
	 * @param {function} fAction - Callback function with logic for modifying cell.
	 * @param {boolean} bRecursiveCall - Changed all linked cells or only nearby cell.
	 * @memberof Cell
	 */
	Cell.prototype.changeLinkedCell = function (fAction, bRecursiveCall) {
		const oThis = this;
		const nThisIndex = getCellIndex(this.nRow, this.nCol);
		const oFormulaParsed = this.getFormulaParsed();

		if (!oFormulaParsed) {
			g_cCalcRecursion.resetRecursionCounter();
			return;
		}
		if (g_cCalcRecursion.checkRecursionCounter()) {
			g_cCalcRecursion.resetRecursionCounter();
			return;
		}

		const aRefElements = _getRefElements(oFormulaParsed);
		foreachRefElements(function (oRange, nIndex, nLastRefElemIndex) {
			oRange._foreachNoEmpty(function(oCell) {
				let nCellIndex = getCellIndex(oCell.nRow, oCell.nCol);
				let sCellWsName = oCell.ws.getName().toLowerCase();
				let sThisWsName = oThis.ws.getName().toLowerCase();
				if (!oCell.isFormula()) {
					return;
				}
				if ((nCellIndex !== nThisIndex && nCellIndex > nThisIndex) || sCellWsName > sThisWsName) {
					let res = fAction(oCell);
					if (res != null) {
						return true;
					}
					if (bRecursiveCall) {
						g_cCalcRecursion.incRecursionCounter();
						oCell.changeLinkedCell(fAction, true);
					}
				} else if (nIndex === nLastRefElemIndex && nCellIndex === nThisIndex && sCellWsName === sThisWsName) {
					// Try to find listener with linked cell.
					const oCellListeners = oCell.getListeners();
					if (!oCellListeners) {
						return;
					}
					_foreachListeners(function (oListenerCell, oThis, nListenerCellIndex) {
						let sThisWsName = oThis.ws.getName().toLowerCase();
						let sListenerWsName = oListenerCell.ws.getName().toLowerCase();
						if (nListenerCellIndex !== nCellIndex || sThisWsName !== sListenerWsName) {
							let oWs = oListenerCell.ws;
							oWs._getCell(oListenerCell.nRow, oListenerCell.nCol, function (oCell) {
								if (oCell.isFormula()) {
									let res = fAction(oCell);
									if (res != null) {
										return;
									}
									if (bRecursiveCall) {
										g_cCalcRecursion.incRecursionCounter();
										oCell.changeLinkedCell(fAction, true);
									}
								}
							});
							return true;
						}
					}, oCell, oCellListeners);
				}
			});
		}, aRefElements);
		g_cCalcRecursion.resetRecursionCounter();
	};
	/**
	 * Fills array of recursive cells.
	 * Uses for Iterative calculation.
	 * @memberof Cell
	 * @param {{cellId:number, wsName:string}[]}aRecursiveCells
	 * @param {{cellId:number, wsName:string}}oStartCellIndex
	 */
	Cell.prototype.fillRecursiveCells = function (aRecursiveCells, oStartCellIndex) {
		aRecursiveCells.unshift(oStartCellIndex); // The first element of the array is always a start cell.
		this.changeLinkedCell(function (oCell) {
			let oCellIndex = {
				cellId: getCellIndex(oCell.nRow, oCell.nCol),
				wsName: oCell.ws.getName().toLowerCase()
			};
			let aPrevRecursiveCell = g_cCalcRecursion.getRecursiveCells(oCell);
			if (aPrevRecursiveCell.length) {
				const oPrevStartCellIndex = aPrevRecursiveCell[0];
				aPrevRecursiveCell = aPrevRecursiveCell.filter(function (oPrevCellIndex) {
					return oPrevCellIndex.cellId !== oCellIndex.cellId || oPrevCellIndex.wsName !== oCellIndex.wsName;
				})
				g_cCalcRecursion.updateRecursiveCells(oPrevStartCellIndex, aPrevRecursiveCell);
			}
			let bDuplicateElem = aRecursiveCells.some(function (oElem) {
				return oElem.cellId === oCellIndex.cellId && oElem.wsName === oCellIndex.wsName;
			});
			if (bDuplicateElem) {
				return true;
			}
			aRecursiveCells.push(oCellIndex);
		}, true);
		g_cCalcRecursion.addRecursiveCells(this, aRecursiveCells);
	};
	/**
	 * Method calculates cells with formula aren't calculated yet.
	 * @memberof Cell
	 * @private
	 */
	Cell.prototype._checkDirty = function() {
		const t = this;
		// Checks cell contains formula or formula is not calculated yet
		if (this.getIsDirty()) {
			if (g_cCalcRecursion.checkLevel()) {
				g_cCalcRecursion.incLevel();
				const isCalc = this.getIsCalc();
				this.setIsCalc(true);
				const calculatedArrayFormulas = [];
				this.processFormula(function(parsed) {
					if (!isCalc) {
						//***array-formula***
						//добавлен последний параметр для обработки формулы массива
						if(parsed.getArrayFormulaRef()) {
							const listenerId = parsed.getListenerId();
							if(parsed.checkFirstCellArray(t) && !calculatedArrayFormulas[listenerId]) {
								parsed.calculate();
								calculatedArrayFormulas[listenerId] = 1;
							} else {
								if(null === parsed.value && !calculatedArrayFormulas[listenerId]) {
									parsed.calculate();
									calculatedArrayFormulas[listenerId] = 1;
								}

								let oldParent = parsed.parent;
								parsed.parent = new AscCommonExcel.CCellWithFormula(t.ws, t.nRow, t.nCol);
								parsed._endCalculate();
								parsed.parent = oldParent;
							}
						} else {
							parsed.calculate();
						}
						if (g_cCalcRecursion.getFunctionsResult().length && g_cCalcRecursion.getIterStep() === 1 && g_cCalcRecursion.getLevel() === 1) {
							const functionsResult = g_cCalcRecursion.getFunctionsResult();
							const isEnabledRecursion = g_cCalcRecursion.getIsEnabledRecursion();
							let isCycleCell = false;

							foreachRefElements(function (range) {
								if (range.bbox.contains(t.nCol, t.nRow)) {
									if (!parsed.ca) {
										parsed.ca = true;
									}
									isCycleCell = true;
									!isEnabledRecursion && g_cCalcRecursion.addCycleCell(t);
								} else {
									range._foreachNoEmpty(function (cell) {
										if (cell.isFormula()) {
											const isRecursiveFormula = cell.checkRecursiveFormula(t);
											!isRecursiveFormula && cell.initStartCellForIterCalc();
											if (g_cCalcRecursion.getStartCellIndex() != null || isRecursiveFormula) {
												if (!cell.getFormulaParsed().ca) {
													cell.getFormulaParsed().ca = true;
												}
												isCycleCell = true;
												!isEnabledRecursion && g_cCalcRecursion.addCycleCell(cell);
												g_cCalcRecursion.getStartCellIndex() && g_cCalcRecursion.setStartCellIndex(null);
												return true;
											}
										}
									});
								}
								if (isCycleCell) {
									return true;
								}
							}, functionsResult);
							if (isCycleCell) {
								if (isEnabledRecursion) {
									const recursiveCells = g_cCalcRecursion.getRecursiveCells(t);
									if (!recursiveCells.length) {
										const cellIndex = {
											cellId: getCellIndex(t.nRow, t.nCol),
											wsName: t.ws.getName().toLowerCase()
										};
										t.fillRecursiveCells(recursiveCells, cellIndex);
									}

								} else {
									parsed.value = null;
									t.setIsDirty(false);
									if (t.getType() !== CellValueType.Number) {
										t.setTypeInternal(CellValueType.Number);
									}
									t.setValueNumberInternal(0);
								}
							}
						}
						if (g_cCalcRecursion.getFunctionsResult().length) {
							g_cCalcRecursion.clearFunctionsResult();
						}
					}
				});

				g_cCalcRecursion.decLevel();
				if (g_cCalcRecursion.getIsForceBacktracking()) {
					g_cCalcRecursion.insert({ws: this.ws, nRow: this.nRow, nCol: this.nCol});
					if (g_cCalcRecursion.getLevel() === 0 && !g_cCalcRecursion.getIsProcessRecursion()) {
						g_cCalcRecursion.setIsProcessRecursion(true);
						do {
							g_cCalcRecursion.setIsForceBacktracking(false);
							g_cCalcRecursion.foreachInReverse(function(oElem) {
								oElem.ws._getCellNoEmpty(oElem.nRow, oElem.nCol, function(oCell) {
									if(oCell && oCell.getIsDirty()) {
										oCell.setIsCalc(false);
										oCell._checkDirty();
									}
								});
							});
						} while (g_cCalcRecursion.getIsForceBacktracking());
						g_cCalcRecursion.setIsProcessRecursion(false);
					}
				} else {
					this.setIsCalc(false);
					this.setIsDirty(false);
				}
			}
		}
	};
	Cell.prototype.getFont=function(){
		var xfs = this.getCompiledStyle();
		if(null != xfs && null != xfs.font)
			return xfs.font;
		return g_oDefaultFormat.Font;
	};
	Cell.prototype.getCellFont = function() {
		return (this.xfs && this.xfs.font) || g_oDefaultFormat.Font;
	};
	Cell.prototype.getFillColor = function () {
		return this.getFill().bg();
	};
	Cell.prototype.getFill = function () {
		var xfs = this.getCompiledStyle();
		if(null != xfs && null != xfs.fill)
			return xfs.fill;
		return g_oDefaultFormat.Fill;
	};
	Cell.prototype.getBorderSrc = function () {
		var xfs = this.getCompiledStyle();
		if(null != xfs && null != xfs.border)
			return xfs.border;
		return g_oDefaultFormat.Border;
	};
	Cell.prototype.getBorder = function () {
		return this.getBorderSrc();
	};
	Cell.prototype.getAlign=function(){
		var xfs = this.getCompiledStyle();
		if (xfs) {
			return xfs.align || g_oDefaultFormat.AlignAbs;
		}
		return g_oDefaultFormat.Align;
	};
	Cell.prototype._adjustCellFormat = function() {
		var t = this;
		this.processFormula(function(parsed) {
			var valueCalc = parsed.value;
			if (valueCalc) {
				if (0 <= valueCalc.numFormat) {
					if (aStandartNumFormatsId[t.getNumFormatStr()] == 0) {
						t.setNum(new AscCommonExcel.Num({id: valueCalc.numFormat}));
					}
				} else if (typeof valueCalc.numFormat === "string") {
					t.setNum(new AscCommonExcel.Num({f: valueCalc.numFormat}));
				} else if (AscCommonExcel.cNumFormatFirstCell === valueCalc.numFormat) {
					//ищет в формуле первый рэндж и устанавливает формат ячейки как формат первой ячейки в рэндже
					//принимают формат первой ячейки в рейндже только функции с inheritFormat = true
					//причём это не касается внутренних функий в формуле. если одна из внешних функций принимает формат, тогда выставляем формат у ячейки

					var r = parsed.getFirstRange();
					if (r && r.getNumFormatStr) {
						var sCurFormat = t.getNumFormatStr();
						if (g_oDefaultFormat.Num.getFormat() == sCurFormat) {
							var sNewFormat = r.getNumFormatStr();
							if (sCurFormat != sNewFormat) {
								t.setNumFormat(sNewFormat);
							}
						}
					}
				}
			}
		});
	};
	Cell.prototype._updateCellValue = function() {
		if (!this.isFormula()) {
			return;
		}
		var parsed = this.getFormulaParsed();
		var res = parsed.simplifyRefType(parsed.value, this.ws, this.nRow, this.nCol);

		let DataOld, DataNew;
		if (AscCommon.History.Is_On()) {
			DataOld = this.getValueData();
		}

		if (res) {
			this.cleanText();
			switch (res.type) {
				case cElementType.number:
					this.setTypeInternal(CellValueType.Number);
					this.setValueNumberInternal(res.getValue());
					break;
				case cElementType.bool:
					this.setTypeInternal(CellValueType.Bool);
					this.setValueNumberInternal(res.value ? 1 : 0);
					break;
				case cElementType.error:
					this.setTypeInternal(CellValueType.Error);
					this.setValueTextInternal(res.getValue().toString());
					break;
				case cElementType.name:
					this.setTypeInternal(CellValueType.Error);
					this.setValueTextInternal(res.getValue().toString());
					break;
				case cElementType.empty:
					//=A1:A3
					this.setTypeInternal(CellValueType.Number);
					this.setValueNumberInternal(res.tocNumber().getValue());
					break;
				default:
					this.setTypeInternal(CellValueType.String);
					this.setValueTextInternal(res.getValue().toString());
			}

			//обработка для функции hyperlink. необходимо проставить формат.
			var isArray = parsed.getArrayFormulaRef();
			var firstArrayRef = parsed.checkFirstCellArray(this);
			if(res.getHyperlink && null !== res.getHyperlink() && (!isArray || (isArray && firstArrayRef))) {
				var oHyperlinkFont = new AscCommonExcel.Font();
				oHyperlinkFont.setName(this.ws.workbook.getDefaultFont());
				oHyperlinkFont.setSize(this.ws.workbook.getDefaultSize());
				oHyperlinkFont.setUnderline(Asc.EUnderline.underlineSingle);
				oHyperlinkFont.setColor(AscCommonExcel.g_oColorManager.getThemeColor(AscCommonExcel.g_nColorHyperlink));
				AscFormat.ExecuteNoHistory(function () {
					this.setFont(oHyperlinkFont);
				}, this, []);
			}

			if (AscCommon.History.Is_On()) {
				DataNew = this.getValueData();
			}

			let _replaceFormulaType = parsed.replaceFormulaAfterCalc;
			if (AscCommon.History.Is_On() && (parsed.unknownOrCustomFunction || _replaceFormulaType === AscCommonExcel.cReplaceFormulaType.formula)
				&& !(DataNew.value && DataNew.value.type === CellValueType.Error && DataNew.value.text === AscCommon.cErrorOrigin['busy'])) {

				let wb = Asc["editor"] && Asc["editor"].wb;
				let currentFunc = wb && wb.customFunctionEngine && wb.customFunctionEngine.getFunc(parsed.unknownOrCustomFunction);

				if ((currentFunc && currentFunc.replaceFormulaToVal) || _replaceFormulaType) {
					if (_replaceFormulaType === AscCommonExcel.cReplaceFormulaType.formula) {
						AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(),
							new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow),
							new UndoRedoData_CellSimpleData(this.nRow, this.nCol, DataOld, this.getValueData()));
					} else {
						this.setFormulaInternal(null);
						AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(),
							new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow),
							new UndoRedoData_CellSimpleData(this.nRow, this.nCol, DataOld, this.getValueData()));
					}
				} else if (!DataOld.isEqual(DataNew)) {
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(),
						new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow),
						new UndoRedoData_CellSimpleData(this.nRow, this.nCol, DataOld, DataNew));
				}
				parsed.replaceFormulaAfterCalc = null;
			}

			this.ws.workbook.dependencyFormulas.addToCleanCellCache(this.ws.getId(), this.nRow, this.nCol);
			AscCommonExcel.g_oLOOKUPCache.remove(this);
			AscCommonExcel.g_oVLOOKUPCache.remove(this);
			AscCommonExcel.g_oHLOOKUPCache.remove(this);
			AscCommonExcel.g_oMatchCache.remove(this);
			AscCommonExcel.g_oSUMIFSCache.remove(this);
			AscCommonExcel.g_oFormulaRangesCache.remove(this);
			AscCommonExcel.g_oCountIfCache.remove(this);
		}
	};
	Cell.prototype.cleanText = function() {
		this.number = null;
		this.text = null;
		this.multiText = null;
		this.textIndex = null;
		this.type = CellValueType.Number;
		this._hasChanged = true;
	};
	Cell.prototype._BuildDependencies = function(parse, opt_dirty) {
		var parsed = this.getFormulaParsed();
		if (parsed) {
			if (parse) {
				parsed.parse();
			}
			parsed.buildDependencies();
			if (opt_dirty || parsed.ca || this.isNullTextString()) {
				this.ws.workbook.dependencyFormulas.addToChangedCell(this);
			}
		}
	};
	Cell.prototype._setValueData = function(val){
		this.number = val.number;
		this.text = val.text;
		this.multiText = val.multiText;
		this.textIndex = null;
		this.type = val.type;
		this._hasChanged = true;
	};
	Cell.prototype._getValue2 = function(dDigitsCount, fIsFitMeasurer, opt_numFormat, opt_cultureInfo, opt_AffectingText) {
		var aRes = null;
		var bNeedMeasure = true;
		var sText = null;
		var aText = null;
		var isMultyText = false;
		if (CellValueType.Number == this.type || CellValueType.String == this.type) {
			if (null != this.text) {
				sText = this.text;
			} else if (null != this.multiText) {
				aText = this.multiText;
				isMultyText = true;
			}

			if (CellValueType.String == this.type) {
				bNeedMeasure = false;
			}
			var oNumFormat;
			if (opt_numFormat) {
				oNumFormat = opt_numFormat;
			} else {
				var xfs = this.getCompiledStyle(undefined, opt_AffectingText);
				if (null != xfs && null != xfs.num) {
					oNumFormat = oNumFormatCache.get(xfs.num.getFormat());
				} else {
					oNumFormat = oNumFormatCache.get(g_oDefaultFormat.Num.getFormat());
				}
			}

			if (false == oNumFormat.isGeneralFormat()) {
				if (null != this.number) {
					aText = oNumFormat.format(this.number, this.type, dDigitsCount, false, opt_cultureInfo);
					isMultyText = false;
					sText = null;
				} else if (CellValueType.String == this.type) {
					var oTextFormat = oNumFormat.getTextFormat();
					if (null != oTextFormat && "@" != oTextFormat.formatString) {
						if (null != this.text) {
							aText = oNumFormat.format(this.text, this.type, dDigitsCount, false, opt_cultureInfo);
							isMultyText = false;
							sText = null;
						} else if (null != this.multiText) {
							var sSimpleString = AscCommonExcel.getStringFromMultiText(this.multiText);
							aText = oNumFormat.format(sSimpleString, this.type, dDigitsCount, false, opt_cultureInfo);
							isMultyText = false;
							sText = null;
						}
					}
				}
			} else if (CellValueType.Number == this.type && null != this.number) {
				bNeedMeasure = false;
				var bFindResult = false;
				//варируем dDigitsCount чтобы результат влез в ячейку
				var nTempDigCount = Math.ceil(dDigitsCount);
				var sOriginText = this.number;
				while (nTempDigCount >= 1) {
					//Строим подходящий general format
					var sGeneral = AscCommon.DecodeGeneralFormat(sOriginText, this.type, nTempDigCount);
					if (null != sGeneral) {
						oNumFormat = oNumFormatCache.get(sGeneral);
					}

					if (null != oNumFormat) {
						sText = null;
						isMultyText = false;
						aText = oNumFormat.format(sOriginText, this.type, dDigitsCount, false, opt_cultureInfo);
						if (true == oNumFormat.isTextFormat()) {
							break;
						} else {
							aRes = this._getValue2Result(sText, aText, isMultyText, opt_AffectingText);
							//Проверяем влезает ли текст
							if (true == fIsFitMeasurer(aRes)) {
								bFindResult = true;
								break;
							}
							aRes = null;
						}
					}
					nTempDigCount--;
				}
				if (false == bFindResult) {
					aRes = null;
					sText = null;
					isMultyText = false;
					var font = new AscCommonExcel.Font();
					if (dDigitsCount > 1) {
						font.setRepeat(true);
						aText = [{text: "#", format: font}];
					} else {
						aText = [{text: "", format: font}];
					}
				}
			}
		} else if (CellValueType.Bool === this.type) {
			if (null != this.number) {
				sText = (0 != this.number) ? cBoolLocal.t : cBoolLocal.f;
			}
		} else if (CellValueType.Error === this.type) {
			if (null != this.text) {
				sText = this._getValueTypeError(this.text);
			}
		}
		if (bNeedMeasure) {
			aRes = this._getValue2Result(sText, aText, isMultyText, opt_AffectingText);
			//Проверяем влезает ли текст
			if (false == fIsFitMeasurer(aRes)) {
				aRes = null;
				sText = null;
				isMultyText = false;
				var font = new AscCommonExcel.Font();
				font.setRepeat(true);
				aText = [{text: "#", format: font}];
			}
		}
		if (null == aRes) {
			aRes = this._getValue2Result(sText, aText, isMultyText, opt_AffectingText);
		}
		return aRes;
	};
	Cell.prototype._setValue = function(val)
	{
		this.cleanText();

		function checkCellValueTypeError(sUpText){
			switch (sUpText){
				case cErrorLocal["nil"]:
					return cErrorOrigin["nil"];
					break;
				case cErrorLocal["div"]:
					return cErrorOrigin["div"];
					break;
				case cErrorLocal["value"]:
					return cErrorOrigin["value"];
					break;
				case cErrorLocal["ref"]:
					return cErrorOrigin["ref"];
					break;
				case cErrorLocal["name"]:
				case cErrorLocal["name"].replace('\\', ''): // ToDo это неправильная правка для бага 32463 (нужно переделать parse формул)
					return cErrorOrigin["name"];
					break;
				case cErrorLocal["num"]:
					return cErrorOrigin["num"];
					break;
				case cErrorLocal["na"]:
					return cErrorOrigin["na"];
					break;
				case cErrorLocal["getdata"]:
					return cErrorOrigin["getdata"];
					break;
				case cErrorLocal["uf"]:
					return cErrorOrigin["uf"];
					break;
			}
			return false;
		}



		if("" == val) {
			if (this.ws.workbook.handlers) {
				this.ws.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.cellValue, this, null, this.ws.getId());
			}
			return;
		}

		var oNumFormat;
		var xfs = this.getCompiledStyle(undefined, true);
		if(null != xfs && null != xfs.num)
			oNumFormat = oNumFormatCache.get(xfs.num.getFormat());
		else
			oNumFormat = oNumFormatCache.get(g_oDefaultFormat.Num.getFormat());
		if(oNumFormat.isTextFormat())
		{
			this.setTypeInternal(CellValueType.String);
			this.setValueTextInternal(val);
		}
		else
		{
			if (AscCommon.g_oFormatParser.isLocaleNumber(val))
			{
				this.setTypeInternal(CellValueType.Number);
				this.setValueNumberInternal(AscCommon.g_oFormatParser.parseLocaleNumber(val));
				if (/E/i.test(val)) {
					this.setNumFormat('0.00E+00');
				}
			}
			else
			{
				var sUpText = val.toUpperCase();
				if(cBoolLocal.t === sUpText || cBoolLocal.f === sUpText)
				{
					this.setTypeInternal(CellValueType.Bool);
					this.setValueNumberInternal((cBoolLocal.t === sUpText) ? 1 : 0);
				}
				else if(checkCellValueTypeError(sUpText))
				{
					this.setTypeInternal(CellValueType.Error);
					this.setValueTextInternal(checkCellValueTypeError(sUpText));
				}
				else
				{
					//распознаем формат
					var res = AscCommon.g_oFormatParser.parse(val);
					if(null != res)
					{
						//Сравниваем с текущим форматом, если типы совпадают - меняем только значение ячейки
						var nFormatType = oNumFormat.getType();
						if(!((c_oAscNumFormatType.Percent == nFormatType && res.bPercent) ||
							(c_oAscNumFormatType.Currency == nFormatType && res.bCurrency) ||
							(c_oAscNumFormatType.Date == nFormatType && res.bDate) ||
							(c_oAscNumFormatType.Time == nFormatType && res.bTime)) && res.format != oNumFormat.sFormat) {
							this.setNumFormat(res.format);
						}
						this.setTypeInternal(CellValueType.Number);
						this.setValueNumberInternal(res.value);
					}
					else
					{
						this.setTypeInternal(CellValueType.String);
						//проверяем QuotePrefix
						if(val.length > 0 && "'" == val[0])
						{
							this.setQuotePrefix(true);
							val = val.substring(1);
						}
						this.setValueTextInternal(val);
					}
				}
			}
		}
		if (this.ws.workbook.handlers) {
			this.ws.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.cellValue, this, null, this.ws.getId());
		}
	};
	Cell.prototype._autoformatHyperlink = function(val){
		if (AscCommon.rx_allowedProtocols.test(val) || /^(www.)|@/i.test(val)) {
			// Удаляем концевые пробелы и переводы строки перед проверкой гиперссылок
			val = val.replace(/\s+$/, '');
			var typeHyp = AscCommon.getUrlType(val);
			if (AscCommon.c_oAscUrlType.Invalid != typeHyp) {
				val = AscCommon.prepareUrl(val, typeHyp);

				var oNewHyperlink = new AscCommonExcel.Hyperlink();
				oNewHyperlink.Ref = this.ws.getCell3(this.nRow, this.nCol);
				oNewHyperlink.Hyperlink = val;
				oNewHyperlink.Ref.setHyperlink(oNewHyperlink);
			}
		}
	}
	Cell.prototype._setValue2 = function(aVal, ignoreHyperlink, xfTableAndCond)
	{
		var sSimpleText = "";
		for(var i = 0, length = aVal.length; i < length; ++i)
			sSimpleText += aVal[i].getFragmentText ? aVal[i].getFragmentText() : aVal[i].text;
		this._setValue(sSimpleText);
		if (!ignoreHyperlink && window['AscCommonExcel'].g_AutoCorrectHyperlinks) {
			this._autoformatHyperlink(sSimpleText);
		}
		var nRow = this.nRow;
		var nCol = this.nCol;
		if(CellValueType.String == this.type && null == this.ws.hyperlinkManager.getByCell(nRow, nCol))
		{
			this.cleanText();
			this.setTypeInternal(CellValueType.String);
			//проверяем можно ли перевести массив в простую строку.
			if(aVal.length > 0)
			{
				this.multiText = [];
				for(var i = 0, length = aVal.length; i < length; i++){
					var item = aVal[i];
					var oNewElem = new AscCommonExcel.CMultiTextElem();
					oNewElem.text = item.getFragmentText ? item.getFragmentText() : item.text;
					if (null != item.format) {
						oNewElem.format = new AscCommonExcel.Font();
						oNewElem.format.assign(item.format);
					}
					this.multiText.push(oNewElem);
				}
				this._subtractFontFromMultiText(xfTableAndCond && xfTableAndCond.font);
				this._minimizeMultiText(true);
			}
			//обрабатываем QuotePrefix
			if(null != this.text)
			{
				if(this.text.length > 0 && "'" == this.text[0])
				{
					this.setQuotePrefix(true);
					this.setValueTextInternal(this.text.substring(1));
				}
			}
			else if(null != this.multiText)
			{
				if(this.multiText.length > 0)
				{
					var oFirstItem = this.multiText[0];
					if(null != oFirstItem.text && oFirstItem.text.length > 0 && "'" == oFirstItem.text[0])
					{
						this.setQuotePrefix(true);
						if(1 != oFirstItem.text.length)
							oFirstItem.text = oFirstItem.text.substring(1);
						else
						{
							this.multiText.shift();
							if(0 == this.multiText.length)
							{
								this.setValueTextInternal("");
							}
						}
					}
				}
			}
		}
	};
	Cell.prototype._setFontProp = function(fCheck, fAction)
	{
		var bRes = false;
		if(null != this.multiText)
		{
			//проверяем поменяются ли свойства
			var bChange = false;
			for(var i = 0, length = this.multiText.length; i < length; ++i)
			{
				var elem = this.multiText[i];
				if (null != elem.format && true == fCheck(elem.format))
				{
					bChange = true;
					break;
				}
			}
			if(bChange)
			{
				var backupObj = this.getValueData();
				for (var i = 0, length = this.multiText.length; i < length; ++i) {
					var elem = this.multiText[i];
					if (null != elem.format)
						fAction(elem.format)
				}
				//пробуем преобразовать в простую строку
				if(this._minimizeMultiText(false))
				{
					var DataNew = this.getValueData();
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeValue, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow,this.nCol, backupObj, DataNew));
				}
				else
				{
					var DataNew = this._cloneMultiText();
					AscCommon.History.Add(AscCommonExcel.g_oUndoRedoCell, AscCH.historyitem_Cell_ChangeArrayValueFormat, this.ws.getId(), new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow), new UndoRedoData_CellSimpleData(this.nRow, this.nCol, backupObj.value.multiText, DataNew));
				}
			}
			bRes = true;
		}
		return bRes;
	};
	Cell.prototype._getValue2Result = function(sText, aText, isMultyText, opt_AffectingText)
	{
		var aResult = [];
		if(null == sText && null == aText)
			sText = "";
		var oNewItem, cellfont = g_oDefaultFormat.Font;
		var cellSelfFont = this.xfs && this.xfs.font;
		var xfs = this.getCompiledStyle(undefined, opt_AffectingText);
		if(null != xfs && null != xfs.font)
			cellfont = xfs.font;
		if(null != sText){
			oNewItem = new AscCommonExcel.Fragment();
			oNewItem.setFragmentText(sText);
			oNewItem.format = cellfont.clone();
			oNewItem.checkVisitedHyperlink(this.nRow, this.nCol, this.ws.hyperlinkManager);
			oNewItem.format.setSkip(false);
			oNewItem.format.setRepeat(false);
			aResult.push(oNewItem);
		} else if(null != aText){
			for(var i = 0; i < aText.length; i++){
				oNewItem = new AscCommonExcel.Fragment();
				var oCurtext = aText[i];
				if(null != oCurtext.text)
				{
					oNewItem.setFragmentText(oCurtext.text);
					var oCurFormat = new AscCommonExcel.Font();
					if (isMultyText && xfs && xfs.font) {
						if (null != oCurtext.format &&  !(cellSelfFont && cellSelfFont.isEqual(oCurtext.format))) {
							//MultyText format equals to cell font
							if (this.xfs && !this.xfs.isNormalFont()) {
								//cell font is not default
								oCurFormat.assign(oCurtext.format);
								if (cellSelfFont && cellSelfFont.c && oCurtext.format.c && oCurtext.format.c.isEqual(cellSelfFont.c)) {
									oCurFormat.c = xfs.font.c;
								}
							} else {
								//like in CellXfs.prototype.merge
								var isTableColor = oCurtext.format.isNormalXfColor();
								let mergedFont = xfs._mergeProperty(g_StyleCache.addFont, oCurtext.format, xfs.font, true, isTableColor);
								oCurFormat.assign(mergedFont);
							}
						} else {
							oCurFormat.assign(cellfont);
							oCurFormat.setSkip(false);
							oCurFormat.setRepeat(false);
						}
					} else {
						oCurFormat.assign(cellfont);
						oCurFormat.setSkip(false);
						oCurFormat.setRepeat(false);
						if (null != oCurtext.format) {
							oCurFormat.assignFromObject(oCurtext.format);
						}
					}
					oNewItem.format = oCurFormat;
					oNewItem.checkVisitedHyperlink(this.nRow, this.nCol, this.ws.hyperlinkManager);
					aResult.push(oNewItem);
				}
			}
		}
		return aResult;
	};
	Cell.prototype._getValueTypeError = function(text) {
		switch (text){
			case cErrorOrigin["nil"]:
				return cErrorLocal["nil"];
				break;
			case cErrorOrigin["div"]:
				return cErrorLocal["div"];
				break;
			case cErrorOrigin["value"]:
				return cErrorLocal["value"];
				break;
			case cErrorOrigin["ref"]:
				return cErrorLocal["ref"];
				break;
			case cErrorOrigin["name"]:
				return cErrorLocal["name"].replace('\\', ''); // ToDo это неправильная правка для бага 32463 (нужно переделать parse формул)
				break;
			case cErrorOrigin["num"]:
				return cErrorLocal["num"];
				break;
			case cErrorOrigin["na"]:
				return cErrorLocal["na"];
				break;
			case cErrorOrigin["getdata"]:
				return cErrorLocal["getdata"];
				break;
			case cErrorOrigin["uf"]:
				return cErrorLocal["uf"];
				break;
			case cErrorOrigin["calc"]:
				return cErrorLocal["calc"];
				break;
			case cErrorOrigin["spill"]:
				return cErrorLocal["spill"];
				break;
			case cErrorOrigin["busy"]:
				return cErrorLocal["busy"];
				break;
		}
		return cErrorLocal["nil"];
	};
	Cell.prototype._subtractFontFromMultiText = function(font) {
		if (!font) {
			return;
		}
		var cellFont = this.getCellFont();
		for (var i = 0; i < this.multiText.length; i++) {
			var elem = this.multiText[i];
			if (null != elem.format) {
				elem.format = elem.format.clone();
				elem.format.subtractEqual(font, cellFont);
			}
		}
	};
	Cell.prototype._minimizeMultiText = function(bSetCellFont) {
		var bRes = false;
		if(null != this.multiText && this.multiText.length > 0)
		{
			var cellFont = this.getCellFont();
			var oIntersectFont = null;
			for (var i = 0, length = this.multiText.length; i < length; i++) {
				var elem = this.multiText[i];
				if (null != elem.format) {
					if (null == oIntersectFont)
						oIntersectFont = elem.format.clone();
					oIntersectFont.intersect(elem.format, cellFont);
				}
				else {
					oIntersectFont = cellFont;
					break;
				}
			}

			if(bSetCellFont)
			{
				if (oIntersectFont.isEqual(g_oDefaultFormat.Font))
					this.setFont(null, false);
				else
					this.setFont(oIntersectFont, false);
			}
			//если у всех элементов один формат, то сохраняем только текст
			var bIsEqual = true;
			for (var i = 0, length = this.multiText.length; i < length; i++)
			{
				var elem = this.multiText[i];
				if (null != elem.format && false == oIntersectFont.isEqual(elem.format))
				{
					bIsEqual = false;
					break;
				}
			}
			if(bIsEqual)
			{
				this.setValueTextInternal(AscCommonExcel.getStringFromMultiText(this.multiText));
				bRes = true;
			}
		}
		return bRes;
	};
	Cell.prototype._cloneMultiText = function() {
		var oRes = [];
		for(var i = 0, length = this.multiText.length; i < length; ++i)
			oRes.push(this.multiText[i].clone());
		return oRes;
	};
	Cell.prototype.getIsDirty = function() {
		return this.isDirty;
	};
	Cell.prototype.setIsDirty = function(val) {
		this.isDirty = val;
		this._hasChanged = true;
	};
	Cell.prototype.getIsCalc = function() {
		return this.isCalc;
	};
	Cell.prototype.setIsCalc = function(val) {
		this.isCalc = val;
		this._hasChanged = true;
	};
	Cell.prototype.fromXLSB = function(stream, type, row, aCellXfs, aSharedStrings, tmp, formula) {
		var end = stream.XlsbReadRecordLength() + stream.GetCurPos();

		this.setRowCol(row, stream.GetULongLE() & 0x3FFF);
		var nFlags2 = stream.GetULongLE();
		if (0 !== (nFlags2 & 0xFFFFFF)) {
			var xf = aCellXfs[nFlags2 & 0xFFFFF];
			if (xf) {
				this.setStyle(xf);
			}
		}
		//todo rt_CELL_RK
		if (AscCommonExcel.XLSB.rt_CELL_REAL === type || AscCommonExcel.XLSB.rt_FMLA_NUM === type) {
			this.setTypeInternal(CellValueType.Number);
			this.setValueNumberInternal(stream.GetDoubleLE());
		} else if (AscCommonExcel.XLSB.rt_CELL_ISST === type) {
			this.setTypeInternal(CellValueType.String);
			//todo textIndex
			var ss = aSharedStrings[stream.GetULongLE()];
			if (undefined !== ss) {
				if (typeof ss === 'string') {
					this.setValueTextInternal(ss);
				} else {
					this.setValueMultiTextInternal(ss);
				}
			}
		} else if (AscCommonExcel.XLSB.rt_CELL_ST === type || AscCommonExcel.XLSB.rt_FMLA_STRING === type) {
			this.setTypeInternal(CellValueType.String);
			this.setValueTextInternal(stream.GetString());
		} else if (AscCommonExcel.XLSB.rt_CELL_ERROR === type || AscCommonExcel.XLSB.rt_FMLA_ERROR === type) {
			this.setTypeInternal(CellValueType.Error);
			switch (stream.GetByte()) {
				case 0x00:
					this.setValueTextInternal("#NULL!");
					break;
				case 0x07:
					this.setValueTextInternal("#DIV/0!");
					break;
				case 0x0F:
					this.setValueTextInternal("#VALUE!");
					break;
				case 0x17:
					this.setValueTextInternal("#REF!");
					break;
				case 0x1D:
					this.setValueTextInternal("#NAME?");
					break;
				case 0x24:
					this.setValueTextInternal("#NUM!");
					break;
				case 0x2A:
					this.setValueTextInternal("#N/A");
					break;
				case 0x2B:
					this.setValueTextInternal("#GETTING_DATA");
					break;
			}
		} else if (AscCommonExcel.XLSB.rt_CELL_BOOL === type || AscCommonExcel.XLSB.rt_FMLA_BOOL === type) {
			this.setTypeInternal(CellValueType.Bool);
			this.setValueNumberInternal(stream.GetByte());
		} else {
			this.setTypeInternal(CellValueType.Number);
		}
		if (AscCommonExcel.XLSB.rt_FMLA_STRING <= type && type <= AscCommonExcel.XLSB.rt_FMLA_ERROR) {
			this.fromXLSBFormula(stream, tmp.formula);
		}
		var flags = stream.GetUShortLE();
		if (0 !== (flags & 0x4)) {
			this.fromXLSBFormulaExt(stream, tmp.formula, flags);
			if (0 !== (flags & 0x4000)) {
				this.setTypeInternal(CellValueType.Number);
				this.setValueNumberInternal(null);
			}
		}
		if (0 !== (flags & 0x2000)) {
			this.setTypeInternal(CellValueType.String);
			var multiText = [];
			if (this.fromXLSBRichText(stream, multiText)) {
				this.setValueMultiTextInternal(multiText);
			} else {
				var text = multiText.reduce(function(accumulator, currentValue) {
					return accumulator + currentValue.text;
				}, '');
				this.setValueTextInternal(text);
			}
		}
		if (0 !== (nFlags2 & 0x2000000))
		{
			this.setCm(stream.GetULong());
		}
		if (0 !== (nFlags2 & 0x4000000))
		{
			this.setVm(stream.GetULong());
		}

		stream.Seek2(end);
	};
	Cell.prototype.fromXLSBFormula = function(stream, formula) {
		var flags = stream.GetUChar();
		stream.Skip2(9);
		if (0 !== (flags & 0x2)) {
			formula.ca = true;
		}
	};
	Cell.prototype.fromXLSBFormulaExt = function(stream, formula, flags) {
		formula.t = flags & 0x3;
		formula.v = stream.GetString();
		if (0 !== (flags & 0x8)) {
			formula.aca = true;
		}
		if (0 !== (flags & 0x10)) {
			formula.bx = true;
		}
		if (0 !== (flags & 0x20)) {
			formula.del1 = true;
		}
		if (0 !== (flags & 0x40)) {
			formula.del2 = true;
		}
		if (0 !== (flags & 0x80)) {
			formula.dt2d = true;
		}
		if (0 !== (flags & 0x100)) {
			formula.dtr = true;
		}
		if (0 !== (flags & 0x200)) {
			formula.r1 = stream.GetString();
		}
		if (0 !== (flags & 0x400)) {
			formula.r2 = stream.GetString();
		}
		if (0 !== (flags & 0x800)) {
			formula.ref = stream.GetString();
		}
		if (0 !== (flags & 0x1000)) {
			formula.si = stream.GetULongLE();
		}
	};
	Cell.prototype.fromXLSBRichText = function(stream, richText) {
		var hasFormat = false;
		var count = stream.GetULongLE();
		while (count-- > 0) {
			var typeRun = stream.GetUChar();
			var run;
			if (0x1 === typeRun) {
				run = new AscCommonExcel.CMultiTextElem();
				run.text = "";
				if (stream.GetBool()) {
					hasFormat = true;
					run.format = new AscCommonExcel.Font();
					run.format.fromXLSB(stream);
					run.format.checkSchemeFont(this.ws.workbook.theme);
				}
				var textCount = stream.GetULongLE();
				while (textCount-- > 0) {
					run.text += stream.GetString();
				}
				richText.push(run);
			}
			else if (0x2 === typeRun) {
				run = new AscCommonExcel.CMultiTextElem();
				run.text = stream.GetString();
				richText.push(run);
			}
		}
		return hasFormat;
	};
	Cell.prototype.toXLSB = function(stream, nXfsId, formulaToWrite, oSharedStrings) {
		var len = 4 + 4 + 2;
		var type = AscCommonExcel.XLSB.rt_CELL_BLANK;
		if (formulaToWrite) {
			len += this.getXLSBSizeFormula(formulaToWrite);
		}
		var textToWrite;
		var isBlankFormula = false;
		if (formulaToWrite) {
			if (!this.isNullTextString()) {
				switch (this.type) {
					case CellValueType.Number:
						type = AscCommonExcel.XLSB.rt_FMLA_NUM;
						len += 8;
						break;
					case CellValueType.String:
						type = AscCommonExcel.XLSB.rt_FMLA_STRING;
						textToWrite = this.text ? this.text : AscCommonExcel.getStringFromMultiText(this.multiText);
						len += 4 + 2 * textToWrite.length;
						break;
					case CellValueType.Error:
						type = AscCommonExcel.XLSB.rt_FMLA_ERROR;
						len += 1;
						break;
					case CellValueType.Bool:
						type = AscCommonExcel.XLSB.rt_FMLA_BOOL;
						len += 1;
						break;
				}
			} else {
				type = AscCommonExcel.XLSB.rt_FMLA_STRING;
				textToWrite = "";
				len += 4 + 2 * textToWrite.length;
				isBlankFormula = true;
			}
		} else {
			if (!this.isNullTextString()) {
				switch (this.type) {
					case CellValueType.Number:
						type = AscCommonExcel.XLSB.rt_CELL_REAL;
						len += 8;
						break;
					case CellValueType.String:
						type = AscCommonExcel.XLSB.rt_CELL_ISST;
						len += 4;
						break;
					case CellValueType.Error:
						type = AscCommonExcel.XLSB.rt_CELL_ERROR;
						len += 1;
						break;
					case CellValueType.Bool:
						type = AscCommonExcel.XLSB.rt_CELL_BOOL;
						len += 1;
						break;
				}
			}
		}

		stream.XlsbStartRecord(type, len);
		stream.WriteULong(this.nCol & 0x3FFF);
		var nFlags2 = 0;
		if (null !== nXfsId) {
			nFlags2 = nXfsId;
		}
		if (this.cm != null) {
			nFlags2 |= 0x2000000;
		}
		if (this.vm != null) {
			nFlags2 |= 0x4000000;
		}
		stream.WriteULong(nFlags2);

		//todo RkNumber
		switch (type) {
			case AscCommonExcel.XLSB.rt_CELL_REAL:
			case AscCommonExcel.XLSB.rt_FMLA_NUM:
				stream.WriteDouble2(this.number);
				break;
			case AscCommonExcel.XLSB.rt_CELL_ISST:
				var index;
				var textIndex = this.getTextIndex();
				if (null !== textIndex) {
					index = oSharedStrings.strings[textIndex];
					if (undefined === index) {
						index = oSharedStrings.index++;
						oSharedStrings.strings[textIndex] = index;
					}
				}
				stream.WriteULong(index);
				break;
			case AscCommonExcel.XLSB.rt_CELL_ST:
			case AscCommonExcel.XLSB.rt_FMLA_STRING:
				stream.WriteString4(textToWrite);
				break;
			case AscCommonExcel.XLSB.rt_CELL_ERROR:
			case AscCommonExcel.XLSB.rt_FMLA_ERROR:
				if ("#NULL!" == this.text) {
					stream.WriteByte(0x00);
				}
				else if ("#DIV/0!" == this.text) {
					stream.WriteByte(0x07);
				}
				else if ("#VALUE!" == this.text) {
					stream.WriteByte(0x0F);
				}
				else if ("#REF!" == this.text) {
					stream.WriteByte(0x17);
				}
				else if ("#NAME?" == this.text) {
					stream.WriteByte(0x1D);
				}
				else if ("#NUM!" == this.text) {
					stream.WriteByte(0x24);
				}
				else if ("#N/A" == this.text) {
					stream.WriteByte(0x2A);
				}
				else {
					stream.WriteByte(0x2B);
				}
				break;
			case AscCommonExcel.XLSB.rt_CELL_BOOL:
			case AscCommonExcel.XLSB.rt_FMLA_BOOL:
				stream.WriteByte(this.number);
				break;
		}
		var flags = 0;
		if (formulaToWrite) {
			flags = this.toXLSBFormula(stream, formulaToWrite, isBlankFormula);
		}


		stream.WriteUShort(flags);

		if (formulaToWrite) {
			flags = this.toXLSBFormulaExt(stream, formulaToWrite);
		}

		if (this.cm != null) {
			stream.WriteULong(this.cm);
		}
		if (this.vm != null) {
			stream.WriteULong(this.vm);
		}

		stream.XlsbEndRecord();
	};
	Cell.prototype.readAttributes = function(attr, uq) {
		if (attr()) {
			this.parseAttributes(attr(), uq);
		}
	};
	Cell.prototype.parseAttributes = function(vals, uq) {
		var val;
		val = vals["r"];
		if (undefined !== val) {
			var oCellAddress = AscCommon.g_oCellAddressUtils.getCellAddress(val);
			this.setRowCol(oCellAddress.getRow0(), oCellAddress.getCol0());
			this.ws.nRowsCount = Math.max(this.ws.nRowsCount, this.nRow);
			this.ws.nColsCount = Math.max(this.ws.nColsCount, this.nCol);
			this.ws.cellsByColRowsCount = Math.max(this.ws.cellsByColRowsCount, this.nCol);
		}
		val = vals["t"];
		if (undefined !== val) {
			if("s" === val) {
				this.type = CellValueType.String;
			}
		}
	};
	Cell.prototype.onStartNode = function(elem, attr, uq, tagend, getStrNode) {
		var attrVals;
		if ('v' === elem) {
			return this;
		}
		return this;
	};
	Cell.prototype.onTextNode = function(text, uq) {
		if(CellValueType.String === this.type) {
			this.text = AscCommon.unleakString(uq(text));
		} else if(CellValueType.Number === this.type) {
			this.number = parseInt(text);
		}
	};
	Cell.prototype.onEndNode = function(prevContext, elem) {
		var res = true;
		if ('v' === elem) {
		} else {
			res = false;
		}
		return res;
	};
	Cell.prototype.getXLSBSizeFormula = function(formulaToWrite) {
		var len = 2 + 4 + 4;
		if (formulaToWrite.formula) {
			len += 4 + 2 * formulaToWrite.formula.length;
		} else {
			len += 4;
		}
		if (undefined !== formulaToWrite.ref) {
			len += 4 + 2 * formulaToWrite.ref.getName().length;
		}
		if (undefined !== formulaToWrite.si) {
			len += 4;
		}
		return len;
	};
	Cell.prototype.toXLSBFormula = function(stream, formulaToWrite, isBlankFormula) {
		var flags = 0;
		if (formulaToWrite.ca) {
			flags |= 0x2;
		}
		stream.WriteUShort(flags);
		stream.WriteULong(0);//cce
		stream.WriteULong(0);//cb

		var flagsExt = 0;
		if (undefined !== formulaToWrite.type) {
			flagsExt |= formulaToWrite.type;
		}
		else {
			flagsExt |= Asc.ECellFormulaType.cellformulatypeNormal;
		}
		flagsExt |= 0x4;
		if (undefined !== formulaToWrite.ref) {
			flagsExt |= 0x800;
		}
		if (undefined !== formulaToWrite.si) {
			flagsExt |= 0x1000;
		}
		if (isBlankFormula) {
			flagsExt |= 0x4000;
		}
		return flagsExt;
	};
	Cell.prototype.toXLSBFormulaExt = function(stream, formulaToWrite) {
		if (undefined !== formulaToWrite.formula) {
			stream.WriteString4(formulaToWrite.formula);
		} else {
			stream.WriteString4("");
		}
		if (undefined !== formulaToWrite.ref) {
			stream.WriteString4(formulaToWrite.ref.getName());
		}
		if (undefined !== formulaToWrite.si) {
			stream.WriteULong(formulaToWrite.si);
		}
	};
	Cell.prototype.setDynamicArrayFlags = function () {
		this.cm = 1;
		// this.vm = 1;
		// todo prop vm - determines the order in which dynamic arrays are expanded?
		// presumably the metadata it refers to contains the estimated size of the array
		// the size of the array is unique and has its own number in the order of initialization on the page
		// for example array 3x1 initialized before than array 2x1 will have the number 1 in vm prop
	};
	Cell.prototype.setCm = function(val) {
		this.cm = val;
	};
	Cell.prototype.setVm = function(val) {
		this.vm = val;
	};
//-------------------------------------------------------------------------------------------------

	function CCellWithFormula(ws, row, col) {
		this.ws = ws;
		this.nRow = row;
		this.nCol = col;
	}
	CCellWithFormula.prototype.onFormulaEvent = function(type, eventData) {
		if (AscCommon.c_oNotifyParentType.GetRangeCell === type) {
			return new Asc.Range(this.nCol, this.nRow, this.nCol, this.nRow);
		} else if (AscCommon.c_oNotifyParentType.Change === type) {
			this._onChange(eventData);
		} else if (AscCommon.c_oNotifyParentType.ChangeFormula === type) {
			this._onChangeFormula(eventData);
		} else if (AscCommon.c_oNotifyParentType.EndCalculate === type) {
			this.ws._getCell(this.nRow, this.nCol, function(cell) {
				cell._updateCellValue();
			});
		} else if (AscCommon.c_oNotifyParentType.Shared === type) {
			return this._onShared(eventData);
		}
	};
	CCellWithFormula.prototype._onChange = function(eventData) {
		var areaData = eventData.notifyData.areaData;
		var shared = eventData.formula.getShared();
		var arrayF = eventData.formula.getArrayFormulaRef();
		var dependencyFormulas = this.ws.workbook.dependencyFormulas;
		if (shared) {
			if (areaData) {
				var bbox = areaData.bbox;
				var changedRange = bbox.getSharedIntersect(shared.ref, areaData.changedBBox);
				dependencyFormulas.addToChangedRange2(this.ws.getId(), changedRange);
			} else {
				dependencyFormulas.addToChangedRange2(this.ws.getId(), shared.ref);
			}
		} else if(arrayF) {
			dependencyFormulas.addToChangedRange2(this.ws.getId(), arrayF);
		} else {
			this.ws.workbook.dependencyFormulas.addToChangedCell(this);
		}
	};
	CCellWithFormula.prototype._onChangeFormula = function(eventData) {
		var t = this;
		var wb = this.ws.workbook;
		var parsed = eventData.formula;
		var shared = parsed.getShared();
		if (shared) {
			var index = wb.workbookFormulas.add(parsed).getIndexNumber();
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoSharedFormula, AscCH.historyitem_SharedFormula_ChangeFormula, null, null, new UndoRedoData_IndexSimpleProp(index, false, parsed.getFormulaRaw(), eventData.assemble), true);
			wb.dependencyFormulas.addToChangedRange2(this.ws.getId(), shared.ref);
		} else {
			this.ws._getCell(this.nRow, this.nCol, function(cell) {
				if (parsed === cell.formulaParsed) {
					cell.setFormulaTemplate(true, function(cell) {
						cell.formulaParsed.setFormulaString(eventData.assemble);
						wb.dependencyFormulas.addToChangedCell(cell);
			});
		}
			});
		}
	};
	CCellWithFormula.prototype._onShared = function(eventData) {
		var res = false;
		var data = eventData.notifyData;
		var parsed = eventData.formula;
		var forceTransform = false;
		var sharedShift;//affected shared
		var ranges;//ranges than needed to be transform
		var i;
		var shared = parsed.getShared();
		if (shared) {
			if (c_oNotifyType.Shift === data.type) {
				var bHor = 0 !== data.offset.col;
				var toDelete = data.offset.col < 0 || data.offset.row < 0;
				var bboxShift = AscCommonExcel.shiftGetBBox(data.bbox, bHor);
				sharedShift = parsed.getSharedIntersect(data.sheetId, bboxShift);
				if (sharedShift) {
					//try to remain as many shared formulas as possible
					//that shared can only be at intersection with bboxShift
					var sharedIntersect;
					if (parsed.canShiftShared(bHor) && (sharedIntersect = bboxShift.intersectionSimple(shared.ref))) {
						//collect relative complement of sharedShift and sharedIntersect
						ranges = sharedIntersect.difference(sharedShift);
						ranges = ranges.concat(sharedShift.difference(sharedIntersect));
						//cut off shared than affected relative complement of bboxShift
						var cantBeShared;
						var diff = bboxShift.difference(new Asc.Range(0, 0, gc_nMaxCol0, gc_nMaxRow0));
						if (toDelete) {
							//cut off shared than affected delete range
							diff.push(data.bbox);
						}
						for (i = 0; i < diff.length; ++i) {
							var elem = parsed.getSharedIntersect(data.sheetId, diff[i]);
							if (elem) {
								if (cantBeShared) {
									cantBeShared.union2(elem)
								} else {
									cantBeShared = elem;
								}
							}
						}
						if (cantBeShared) {
							var intersection = sharedIntersect.intersectionSimple(cantBeShared);
							if (intersection) {
								ranges.push(intersection);
							}
						}
						forceTransform = true;
						data.shiftedShared[parsed.getListenerId()] = c_oSharedShiftType.PreProcessed;
					} else {
						ranges = [sharedShift];
					}
				}
				res = true;
			} else if (c_oNotifyType.Move === data.type || c_oNotifyType.Delete === data.type) {
				sharedShift = parsed.getSharedIntersect(data.sheetId, data.bbox);
				if (sharedShift) {
					ranges = [sharedShift];
				}
				res = true;
			} else if (AscCommon.c_oNotifyType.ChangeDefName === data.type && data.bConvertTableFormulaToRef) {
				this._processShared(shared, shared.ref, data, parsed, true, function(newFormula) {
					return newFormula.processNotify(data);
				});
				res = true;
			}
			if (ranges) {
				for (i = 0; i < ranges.length; ++i) {
					this._processShared(shared, ranges[i], data, parsed, forceTransform, function(newFormula) {
						return newFormula.shiftCells(data.type, data.sheetId, data.bbox, data.offset, data.sheetIdTo);
					});
				}
			}
		}
		return res;
	};
	CCellWithFormula.prototype._processShared = function(shared, ref, data, parsed, forceTransform, action) {
		var t = this;
		var cellWithFormula;
		var cellOffset = new AscCommon.CellBase();
		var newFormula;
		this.ws.getRange3(ref.r1, ref.c1, ref.r2, ref.c2)._foreachNoEmpty(function(cell) {
			if (parsed === cell.getFormulaParsed()) {
				if (!cellWithFormula) {
					cellWithFormula = new AscCommonExcel.CCellWithFormula(cell.ws, cell.nRow, cell.nCol);
					newFormula = parsed.clone(undefined, cellWithFormula);
					newFormula.removeShared();
					cellOffset.row = cell.nRow - shared.base.nRow;
					cellOffset.col = cell.nCol - shared.base.nCol;
				} else {
					cellOffset.row = cell.nRow - cellWithFormula.nRow;
					cellOffset.col = cell.nCol - cellWithFormula.nCol;
					cellWithFormula.nRow = cell.nRow;
					cellWithFormula.nCol = cell.nCol;
				}
				newFormula.changeOffset(cellOffset, false);
				if (action(newFormula) || forceTransform) {
					cellWithFormula = undefined;
					newFormula.setFormulaString(newFormula.assemble(true));
					cell.setFormulaTemplate(true, function(cell) {
						cell.setFormulaInternal(newFormula);
						newFormula.buildDependencies();
					});
				}
				t.ws.workbook.dependencyFormulas.addToChangedCell(cell);
			}
		});
	};

	function CellTypeAndValue(type, v) {
		this.type = type;
		this.v = v;
	}
	CellTypeAndValue.prototype.valueOf = function() {
		return this.v;
	};

	function ignoreFirstRowSort(worksheet, bbox) {
		var res = false;

		var oldExcludeVal = worksheet.bExcludeHiddenRows;
		worksheet.bExcludeHiddenRows = false;
		if(bbox.r1 < bbox.r2) {
			var rowFirst = worksheet.getRange3(bbox.r1, bbox.c1, bbox.r1, bbox.c2);
			var rowSecond = worksheet.getRange3(bbox.r1 + 1, bbox.c1, bbox.r1 + 1, bbox.c2);
			var typesFirst = [];
			var typesSecond = [];
			rowFirst._setPropertyNoEmpty(null, null, function(cell, row, col) {
				if (cell && !cell.isNullTextString()) {
					typesFirst.push({col: col, type: cell.getType()});
				}
			});
			rowSecond._setPropertyNoEmpty(null, null, function(cell, row, col) {
				if (cell && !cell.isNullTextString()) {
					typesSecond.push({col: col, type: cell.getType()});
				}
			});
			var indexFirst = 0;
			var indexSecond = 0;
			while (indexFirst < typesFirst.length && indexSecond < typesSecond.length) {
				var curFirst = typesFirst[indexFirst];
				var curSecond = typesSecond[indexSecond];
				if (curFirst.col < curSecond.col) {
					indexFirst++;
				} else if (curFirst.col > curSecond.col) {
					indexSecond++;
				} else {
					if (curFirst.type != curSecond.type) {
						//has head
						res = true;
						break;
					}
					indexFirst++;
					indexSecond++;
				}
			}
		}
		worksheet.bExcludeHiddenRows = oldExcludeVal;

		return res;
	}

	/**
	 * @constructor
	 */
	function Range(worksheet, r1, c1, r2, c2){
		this.worksheet = worksheet;
		this.bbox = new Asc.Range(c1, r1, c2, r2);
	}
	Range.prototype.createFromBBox=function(worksheet, bbox){
		var oRes = new Range(worksheet, bbox.r1, bbox.c1, bbox.r2, bbox.c2);
		oRes.bbox = bbox.clone();
		return oRes;
	};
	Range.prototype.clone=function(oNewWs){
		if(!oNewWs)
			oNewWs = this.worksheet;
		return this.createFromBBox(oNewWs, this.bbox);
	};
	Range.prototype.isIntersect=function(oRange){
		if(this.worksheet === oRange.worksheet) {
			return this.bbox.isIntersect(oRange.bbox);
		}
		return false;
	};
	Range.prototype.containsRange=function(oRange){
		if(this.worksheet === oRange.worksheet) {
			return this.bbox.containsRange(oRange.bbox);
		}
		return false;
	};
	Range.prototype.isIntersectForInsertColRow=function(oRange, isInsertCol){
		if(this.worksheet === oRange.worksheet) {
			return this.bbox.isIntersectForInsertColRow(oRange.bbox, isInsertCol);
		}
		return false;
	};
	Range.prototype._foreach = function(action) {
		if (null != action) {
			var wb = this.worksheet.workbook;
			var tempCell = new Cell(this.worksheet);
			wb.loadCells.push(tempCell);
			var oBBox = this.bbox;
			for (var i = oBBox.r1; i <= oBBox.r2; i++) {
				if (this.worksheet.bExcludeHiddenRows && this.worksheet.getRowHidden(i)) {
					continue;
				}
				for (var j = oBBox.c1; j <= oBBox.c2; j++) {
					var targetCell = null;
					for (var k = 0; k < wb.loadCells.length - 1; ++k) {
						var elem = wb.loadCells[k];
						if (elem.nRow == i && elem.nCol == j && this.worksheet === elem.ws) {
							targetCell = elem;
							break;
						}
					}
					if (null === targetCell) {
						if (!tempCell.loadContent(i, j)) {
							this.worksheet._initCell(tempCell, i, j);
						}
						action(tempCell, i, j, oBBox.r1, oBBox.c1);
						tempCell.saveContent(true);
					} else {
						action(targetCell, i, j, oBBox.r1, oBBox.c1);
					}
				}
			}
			wb.loadCells.pop();
		}
	};
	Range.prototype._foreach2=function(action){
		if(null != action)
		{
			var wb = this.worksheet.workbook;
			var oRes;
			var tempCell = new Cell(this.worksheet);
			wb.loadCells.push(tempCell);
			var oBBox = this.bbox, minC = Math.min( this.worksheet.getColDataLength(), oBBox.c2 ), minR = Math.min( this.worksheet.cellsByColRowsCount - 1, oBBox.r2 );
			for(var i = oBBox.r1; i <= minR; i++){
				if (this.worksheet.bExcludeHiddenRows && this.worksheet.getRowHidden(i)) {
					continue;
				}
				for(var j = oBBox.c1; j <= minC; j++){
					var targetCell = null;
					for (var k = 0; k < wb.loadCells.length - 1; ++k) {
						var elem = wb.loadCells[k];
						if (elem.nRow == i && elem.nCol == j && this.worksheet === elem.ws) {
							targetCell = elem;
							break;
						}
					}
					if (null === targetCell) {
						if (tempCell.loadContent(i, j)) {
							oRes = action(tempCell, i, j, oBBox.r1, oBBox.c1);
							tempCell.saveContent(true);
						} else {
							oRes = action(null, i, j, oBBox.r1, oBBox.c1);
						}
					} else {
						oRes = action(targetCell, i, j, oBBox.r1, oBBox.c1);
					}

					if(null != oRes){
						wb.loadCells.pop();
						return oRes;
					}
				}
			}
			wb.loadCells.pop();
		}
	};
	Range.prototype._foreachNoEmpty = function(actionCell, actionRow, excludeHiddenRows) {
		var oRes, i, oBBox = this.bbox, minR = Math.max(this.worksheet.cellsByColRowsCount - 1, this.worksheet.rowsData.getMaxIndex());
		minR = Math.min(minR, oBBox.r2);
		if (actionCell || actionRow) {
			var itRow = new RowIterator();
			if (actionCell) {
				itRow.init(this.worksheet, this.bbox.r1, this.bbox.c1, this.bbox.c2);
			}
			var bExcludeHiddenRows = (this.worksheet.bExcludeHiddenRows || excludeHiddenRows);
			var excludedCount = 0;
			var tempCell;
			var tempRow = new AscCommonExcel.Row(this.worksheet);
			var allRow = this.worksheet.getAllRow();
			var allRowHidden = allRow && allRow.getHidden();
			for (i = oBBox.r1; i <= minR; i++) {
				if (actionRow) {
					if (tempRow.loadContent(i)) {
						if (bExcludeHiddenRows && tempRow.getHidden()) {
							excludedCount++;
							continue;
						}
						oRes = actionRow(tempRow, excludedCount);
						tempRow.saveContent(true);
						if (null != oRes) {
							if (actionCell) {
								itRow.release();
							}
							return oRes;
						}
					} else if (bExcludeHiddenRows && allRowHidden) {
						excludedCount++;
						continue;
					}
				} else if (bExcludeHiddenRows && this.worksheet.getRowHidden(i)) {
					excludedCount++;
					continue;
				}
				if (actionCell) {
					itRow.setRow(i);
					while (tempCell = itRow.next()) {
						oRes = actionCell(tempCell, i, tempCell.nCol, oBBox.r1, oBBox.c1, excludedCount);
						if (null != oRes) {
							if (actionCell) {
								itRow.release();
							}
							return oRes;
						}
					}
				}
			}
			if (actionCell) {
				itRow.release();
			}
		}
	};
	Range.prototype._foreachNoEmptyByCol = function(actionCell, excludeHiddenRows) {
		var oRes, i, j, colData;
		var wb = this.worksheet.workbook;
		var oBBox = this.bbox, minR = Math.min(this.worksheet.cellsByColRowsCount - 1, oBBox.r2);
		var minC = Math.min(this.worksheet.getColDataLength() - 1, oBBox.c2);
		if (actionCell && oBBox.c1 <= minC && oBBox.r1 <= minR) {
			var bExcludeHiddenRows = (this.worksheet.bExcludeHiddenRows || excludeHiddenRows);
			var excludedCount = 0;
			var tempCell = new Cell(this.worksheet);
			wb.loadCells.push(tempCell);
			for (j = oBBox.c1; j <= minC; ++j) {
				colData = this.worksheet.getColDataNoEmpty(j);
				if (colData) {
					for (i = oBBox.r1; i <= Math.min(minR, colData.getMaxIndex()); i++) {
						if (bExcludeHiddenRows && this.worksheet.getRowHidden(i)) {
							excludedCount++;
							continue;
						}
						var targetCell = null;
						for (var k = 0; k < wb.loadCells.length - 1; ++k) {
							var elem = wb.loadCells[k];
							if (elem.nRow == i && elem.nCol == j && this.worksheet === elem.ws) {
								targetCell = elem;
								break;
							}
						}
						if (null === targetCell) {
							if (tempCell.loadContent(i, j, colData)) {
								oRes = actionCell(tempCell, i, j, oBBox.r1, oBBox.c1, excludedCount);
								tempCell.saveContent(true);
							}
						} else {
							oRes = actionCell(targetCell, i, j, oBBox.r1, oBBox.c1, excludedCount);
						}
						if (null != oRes) {
							wb.loadCells.pop();
							return oRes;
						}
					}
				}
			}
			wb.loadCells.pop();
		}
	};
	Range.prototype._foreachRow = function(actionRow, actionCell){
		var oBBox = this.bbox;
		if (null != actionRow) {
			var tempRow = new AscCommonExcel.Row(this.worksheet);
			for (var i = oBBox.r1; i <= oBBox.r2; i++) {
				if (!tempRow.loadContent(i)) {
					this.worksheet._initRow(tempRow, i);
				}
				if (this.worksheet.bExcludeHiddenRows && tempRow.getHidden()) {
					continue;
				}
				actionRow(tempRow);
				tempRow.saveContent(true);
			}
		}
		if (null != actionCell) {
			return this._foreachNoEmpty(actionCell);
		}
	};
	Range.prototype._foreachRowNoEmpty = function(actionRow, actionCell, excludeHiddenRows) {
		return this._foreachNoEmpty(actionCell, actionRow, excludeHiddenRows);
	};
	Range.prototype._foreachCol = function(actionCol, actionCell){
		var oBBox = this.bbox;
		if(null != actionCol)
		{
			for(var i = oBBox.c1; i <= oBBox.c2; ++i)
			{
				var col = this.worksheet._getCol(i);
				if(null != col)
					actionCol(col);
			}
		}
		if(null != actionCell) {
			this._foreachNoEmpty(actionCell);
						}
	};
	Range.prototype._foreachColNoEmpty=function(actionCol, actionCell){
		var oBBox = this.bbox;
		var minC = Math.min(oBBox.c2, this.worksheet.aCols.length);
			if(null != actionCol)
			{
				for(var i = oBBox.c1; i <= minC; ++i)
				{
					var col = this.worksheet._getColNoEmpty(i);
					if(null != col)
					{
						var oRes = actionCol(col);
						if(null != oRes)
							return oRes;
					}
				}
			}
		if (null != actionCell) {
			return this._foreachNoEmpty(actionCell);
		}
	};
	Range.prototype._getRangeType=function(oBBox){
		if(null == oBBox)
			oBBox = this.bbox;
		return getRangeType(oBBox);
	};
	Range.prototype._getValues = function (numbers) {
		var res = [];
		var fAction = numbers ? function (c) {
			var v = c.getNumberValue();
			if (null !== v) {
				res.push(v);
			}
		} : function (c) {
			res.push(new CellTypeAndValue(c.getType(), c.getValueWithoutFormat()));
		};
		this._setPropertyNoEmpty(null, null, fAction);
		return res;
	};
	Range.prototype._getValuesAndMap = function (withEmpty) {
		var v, arrRes = [], mapRes = {};
		var fAction = function(c) {
			v = c.getValueWithoutFormat();
			arrRes.push(new CellTypeAndValue(c.getType(), v));
			mapRes[v.toLowerCase()] = true;
		};
		if (withEmpty) {
			this._foreach(fAction);
		} else {
			this._setPropertyNoEmpty(null, null, fAction);
		}
		return {values: arrRes, map: mapRes};
	};
	Range.prototype._setProperty=function(actionRow, actionCol, actionCell){
		var nRangeType = this._getRangeType();
		if(c_oRangeType.Range == nRangeType)
			this._foreach(actionCell);
		else if(c_oRangeType.Row == nRangeType)
			this._foreachRow(actionRow, actionCell);
		else if(c_oRangeType.Col == nRangeType)
			this._foreachCol(actionCol, actionCell);
		else
		{
			//сюда не должны заходить вообще
			// this._foreachRow(actionRow, actionCell);
			// if(null != actionCol)
			// this._foreachCol(actionCol, null);
		}
	};
	Range.prototype._setPropertyNoEmpty=function(actionRow, actionCol, actionCell){
		var nRangeType = this._getRangeType();
		if(c_oRangeType.Range == nRangeType)
			return this._foreachNoEmpty(actionCell);
		else if(c_oRangeType.Row == nRangeType)
			return this._foreachRowNoEmpty(actionRow, actionCell);
		else if(c_oRangeType.Col == nRangeType)
			return this._foreachColNoEmpty(actionCol, actionCell);
		else
		{
			var oRes = this._foreachRowNoEmpty(actionRow, actionCell);
			if(null != oRes)
				return oRes;
			if(null != actionCol)
				oRes = this._foreachColNoEmpty(actionCol, null);
			return oRes;
		}
	};
	Range.prototype.containCell=function(cellId){
		var cellAddress = cellId;
		return 	cellAddress.getRow0() >= this.bbox.r1 && cellAddress.getCol0() >= this.bbox.c1 &&
			cellAddress.getRow0() <= this.bbox.r2 && cellAddress.getCol0() <= this.bbox.c2;
	};
	Range.prototype.containCell2=function(cell){
		return 	cell.nRow >= this.bbox.r1 && cell.nCol >= this.bbox.c1 &&
			cell.nRow <= this.bbox.r2 && cell.nCol <= this.bbox.c2;
	};
	Range.prototype.cross = function(bbox){
		if( bbox.r1 >= this.bbox.r1 && bbox.r1 <= this.bbox.r2 && this.bbox.c1 == this.bbox.c2)
			return {r:bbox.r1};
		if( bbox.c1 >= this.bbox.c1 && bbox.c1 <= this.bbox.c2 && this.bbox.r1 == this.bbox.r2)
			return {c:bbox.c1};

		return undefined;
	};
	Range.prototype.getWorksheet=function(){
		return this.worksheet;
	};
	Range.prototype.isFormula = function(){
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var isFormula;
		this.worksheet._getCellNoEmpty(nRow, nCol, function(cell) {
			isFormula = cell.isFormula();
		});
		return isFormula;
	};
	Range.prototype.isOneCell=function(){
		var oBBox = this.bbox;
		return oBBox.r1 === oBBox.r2 && oBBox.c1 === oBBox.c2;
	};
	Range.prototype.getBBox0=function(){
		//0 - based
		return this.bbox;
	};
	Range.prototype.getName=function(){
		return this.bbox.getName();
	};
	Range.prototype.getMinimalCellsRange = function() {
		var nType = this._getRangeType();
		var oMinRange;
		if(nType === c_oRangeType.Range) {
			return this;
		} else {
			oMinRange = this.worksheet.getMinimalRange();
			if(!oMinRange) {
				return null;
			}
			var oBB = this.bbox;
			var r1 = oBB.r1, c1 = oBB.c1, r2 = oBB.r2, c2 = oBB.c2;
			if(nType === c_oRangeType.Col) {
				r1 = oMinRange.r1;
				r2 = oMinRange.r2;
			} else if(nType === c_oRangeType.Row) {
				c1 = oMinRange.c1;
				c2 = oMinRange.c2;
			} else {
				r1 = oMinRange.r1;
				r2 = oMinRange.r2;
				c1 = oMinRange.c1;
				c2 = oMinRange.c2;
			}
			return new Range(this.worksheet, r1, c1, r2, c2);
		}
	};
	Range.prototype.setValue=function(val, callback, isCopyPaste, byRef, ignoreHyperlink, dynamicRange){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		//не хотелось бы вводить дополнительный параметр, поэтому если byRef == null
		//то в качестве значения придёт формула, которой необходимо сделать offset в зависимости от range
		//при вызове данной функции обратить внимание на  параметр byRef
		var _formula, t = this, activeCell;
		if (byRef === null) {
			_formula = new AscCommonExcel.parserFormula(val.substr(1), null, this.worksheet);
			if (!_formula.parse(true)) {
				_formula = null;
			} else {
				var _selection = this.worksheet.getSelection();
				activeCell = _selection.activeCell;
			}
		}
		this._foreach(function(cell){
			var _val = val;
			if (_formula) {
				_formula.isParsed = false;
				_formula.outStack = [];
				_formula.parse(true);
				var offset = new AscCommon.CellBase(cell.nRow - activeCell.row, cell.nCol - activeCell.col);
				_val = "=" + _formula.changeOffset(offset, null, true).assembleLocale(AscCommonExcel.cFormulaFunctionToLocale, true, true);
			}
			// when creating new array formulas, it is necessary to transfer information about the dynamic array because regular ref is used
			cell.setValue(_val, callback, isCopyPaste, byRef, ignoreHyperlink, dynamicRange);
		});

		//***array-formula***
		let _sFormula = this.getFormula();
		if(_sFormula && byRef) {
			History.Add(AscCommonExcel.g_oUndoRedoArrayFormula, AscCH.historyitem_ArrayFromula_AddFormula, this.worksheet.getId(),
				new Asc.Range(this.bbox.c1, this.bbox.r1, this.bbox.c2, this.bbox.r2), new AscCommonExcel.UndoRedoData_ArrayFormula(this.bbox, "=" + _sFormula));
		}

		AscCommon.History.EndTransaction();
	};
	Range.prototype.setValue2=function(array, pushOnlyFirstMergedCell){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		//[{"text":"qwe","format":{"b":true, "i":false, "u":Asc.EUnderline.underlineNone, "s":false, "fn":"Arial", "fs": 12, "c": 0xff00ff, "va": "subscript"  }},{}...]
		/*
		 Устанавливаем значение в Range ячеек. В отличае от setValue, сюда мы попадаем только в случае ввода значения отличного от формулы. Таким образом, если в ячейке была формула, то для нее в графе очищается список ячеек от которых зависела. После чего выставляем флаг о необходимости пересчета.
		 */
		let t = this;
		this._foreach(function(cell){
			if (pushOnlyFirstMergedCell) {
				let merged = t.worksheet.getMergedByCell(cell.nRow, cell.nCol);
				if (!merged || (merged && merged.r1 === cell.nRow && merged.c1 === cell.nCol)) {
					cell.setValue2(array);
				}
			} else {
				cell.setValue2(array);
			}
			// if(cell.isEmpty())
			// cell.Remove();
		});
		AscCommon.History.EndTransaction();
		this.worksheet.workbook.oApi.onWorksheetChange(this.bbox);
	};
	Range.prototype.setValueData = function(val){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		this._foreach(function(cell){
			cell.setValueData(val);
		});
		AscCommon.History.EndTransaction();
	};
	Range.prototype.setCellStyle=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var oStyle;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			oStyle = this.worksheet.getAllCol().setCellStyle(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							if(c_oRangeType.All == nRangeType && null == row.xfs)
								return;
							oStyle = row.setCellStyle(val);
						},
						function(col){
							oStyle = col.setCellStyle(val);
						},
						function(cell){
							oStyle = cell.setCellStyle(val);
						});
		if (oStyle && oStyle.ApplyNumberFormat)
			this.setNumFormatPivot(oStyle.getNumFormatStr());
	};
	Range.prototype.setStyle=function(val){
		if (val === null) {
			return;
		}
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setStyle(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
				if(c_oRangeType.All == nRangeType && null == row.xfs)
					return;
				row.setStyle(val);
			},
			function(col){
				col.setStyle(val);
			},
			function(cell){
				cell.setStyle(val);
			});
	};
	Range.prototype.clearTableStyle = function() {
		this.worksheet.sheetMergedStyles.clearTablePivotStyle(this.bbox);
	};
	Range.prototype.setTableStyle = function(xf, stripe) {
		if (xf) {
			this.worksheet.sheetMergedStyles.setTablePivotStyle(this.bbox, xf, stripe);
		}
	};
	Range.prototype.setNumFormatPivot=function(val){
		const pivotTables = this.worksheet.getPivotTablesIntersectingRange(this.bbox);
		pivotTables.forEach(function (pivotTable) {
			pivotTable.formatsManager.setNum(this.bbox, val);
		}, this);
	}
	Range.prototype.setNumFormat=function(val){
		this.setNum(new AscCommonExcel.Num({f:val}));
		this.setNumFormatPivot(val);
	};
	Range.prototype.setNum = function(val) {
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if (c_oRangeType.All == nRangeType) {
			this.worksheet.getAllCol().setNum(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row) {
							  if (c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setNum(val);
						  },
						  function(col) {
							  col.setNum(val);
						  },
						  function(cell) {
							  cell.setNum(val);
						  });
	};
	Range.prototype.getShiftedNumFormat=function(nShift, dDigitsCount){
		AscCommon.History.Create_NewPoint();
		let newNumFormat;
		this._setPropertyNoEmpty(null, null, function(cell, nRow0, nCol0, nRowStart, nColStart){
			newNumFormat = cell.getShiftedNumFormat(nShift, dDigitsCount);
		});
		return newNumFormat;
	};
	Range.prototype.setFont=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setFont(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setFont(val);
						  },
						  function(col){
							  col.setFont(val);
						  },
						  function(cell){
							  cell.setFont(val);
						  });
	};
	Range.prototype.setFontname=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setFontname(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setFontname(val);
						  },
						  function(col){
							  col.setFontname(val);
						  },
						  function(cell){
							  cell.setFontname(val);
						  });
	};
	Range.prototype.setFontsize=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setFontsize(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setFontsize(val);
						  },
						  function(col){
							  col.setFontsize(val);
						  },
						  function(cell){
							  cell.setFontsize(val);
						  });
	};
	Range.prototype.setFontcolor=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setFontcolor(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setFontcolor(val);
						  },
						  function(col){
							  col.setFontcolor(val);
						  },
						  function(cell){
							  cell.setFontcolor(val);
						  });
	};
	Range.prototype.setBold=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setBold(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setBold(val);
						  },
						  function(col){
							  col.setBold(val);
						  },
						  function(cell){
							  cell.setBold(val);
						  });
	};
	Range.prototype.setItalic=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setItalic(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setItalic(val);
						  },
						  function(col){
							  col.setItalic(val);
						  },
						  function(cell){
							  cell.setItalic(val);
						  });
	};
	Range.prototype.setUnderline=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setUnderline(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setUnderline(val);
						  },
						  function(col){
							  col.setUnderline(val);
						  },
						  function(cell){
							  cell.setUnderline(val);
						  });
	};
	Range.prototype.setStrikeout=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setStrikeout(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setStrikeout(val);
						  },
						  function(col){
							  col.setStrikeout(val);
						  },
						  function(cell){
							  cell.setStrikeout(val);
						  });
	};
	Range.prototype.setFontAlign=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setFontAlign(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setFontAlign(val);
						  },
						  function(col){
							  col.setFontAlign(val);
						  },
						  function(cell){
							  cell.setFontAlign(val);
						  });
	};
	Range.prototype.setAlignVertical=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setAlignVertical(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setAlignVertical(val);
						  },
						  function(col){
							  col.setAlignVertical(val);
						  },
						  function(cell){
							  cell.setAlignVertical(val);
						  });
	};
	Range.prototype.setAlignHorizontal=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setAlignHorizontal(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setAlignHorizontal(val);
						  },
						  function(col){
							  col.setAlignHorizontal(val);
						  },
						  function(cell){
							  cell.setAlignHorizontal(val);
						  });
	};
	Range.prototype.setFill=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setFill(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setFill(val);
						  },
						  function(col){
							  col.setFill(val);
						  },
						  function(cell){
							  cell.setFill(val);
						  });
	};
	Range.prototype.setFillColor=function(val){
		var fill = new AscCommonExcel.Fill();
		fill.fromColor(val);
		return this.setFill(fill);
	};
	Range.prototype.setBorderSrc=function(border){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		if (null == border)
			border = new Border();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setBorder(border.clone());
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setBorder(border.clone());
						  },
						  function(col){
							  col.setBorder(border.clone());
						  },
						  function(cell){
							  cell.setBorder(border.clone());
						  });
		AscCommon.History.EndTransaction();
	};
	Range.prototype._setBorderMerge=function(bLeft, bTop, bRight, bBottom, oNewBorder, oCurBorder){
		var oTargetBorder = new Border();
		//не делаем clone для свойств потому у нас нельзя поменять свойство отдельное свойство border можно только применить border целиком
		if(bLeft)
			oTargetBorder.l = oNewBorder.l;
		else
			oTargetBorder.l = oNewBorder.iv;
		if(bTop)
			oTargetBorder.t = oNewBorder.t;
		else
			oTargetBorder.t = oNewBorder.ih;
		if(bRight)
			oTargetBorder.r = oNewBorder.r;
		else
			oTargetBorder.r = oNewBorder.iv;
		if(bBottom)
			oTargetBorder.b = oNewBorder.b;
		else
			oTargetBorder.b = oNewBorder.ih;
		oTargetBorder.d = oNewBorder.d;
		oTargetBorder.dd = oNewBorder.dd;
		oTargetBorder.du = oNewBorder.du;
		var oRes = null;
		if(null != oCurBorder)
		{
			oCurBorder.mergeInner(oTargetBorder);
			oRes = oCurBorder;
		}
		else
			oRes = oTargetBorder;
		return oRes;
	};
	Range.prototype._setCellBorder=function(bbox, cell, oNewBorder){
		if(null == oNewBorder)
			cell.setBorder(oNewBorder);
		else
		{
			var oCurBorder = null;
			if(null != cell.xfs && null != cell.xfs.border)
				oCurBorder = cell.xfs.border.clone();
			else
				oCurBorder = g_oDefaultFormat.Border.clone();
			var nRow = cell.nRow;
			var nCol = cell.nCol;
			cell.setBorder(this._setBorderMerge(nCol == bbox.c1, nRow == bbox.r1, nCol == bbox.c2, nRow == bbox.r2, oNewBorder, oCurBorder));
		}
	};
	Range.prototype._setRowColBorder=function(bbox, rowcol, bRow, oNewBorder){
		if(null == oNewBorder)
			rowcol.setBorder(oNewBorder);
		else
		{
			var oCurBorder = null;
			if(null != rowcol.xfs && null != rowcol.xfs.border)
				oCurBorder = rowcol.xfs.border.clone();
			var bLeft, bTop, bRight, bBottom = false;
			if(bRow)
			{
				bTop = rowcol.index == bbox.r1;
				bBottom = rowcol.index == bbox.r2;
			}
			else
			{
				bLeft = rowcol.index == bbox.c1;
				bRight = rowcol.index == bbox.c2;
			}
			rowcol.setBorder(this._setBorderMerge(bLeft, bTop, bRight, bBottom, oNewBorder, oCurBorder));
		}
	};
	Range.prototype._setBorderEdge=function(bbox, oItemWithXfs, nRow, nCol, oNewBorder){
		var oCurBorder = null;
		if(null != oItemWithXfs.xfs && null != oItemWithXfs.xfs.border)
			oCurBorder = oItemWithXfs.xfs.border;
		if(null != oCurBorder)
		{
			var oCurBorderProp = null;
			if(nCol == bbox.c1 - 1)
				oCurBorderProp = oCurBorder.r;
			else if(nRow == bbox.r1 - 1)
				oCurBorderProp = oCurBorder.b;
			else if(nCol == bbox.c2 + 1)
				oCurBorderProp = oCurBorder.l;
			else if(nRow == bbox.r2 + 1)
				oCurBorderProp = oCurBorder.t;
			var oNewBorderProp = null;
			if(null == oNewBorder)
				oNewBorderProp = new AscCommonExcel.BorderProp();
			else
			{
				if(nCol == bbox.c1 - 1)
					oNewBorderProp = oNewBorder.l;
				else if(nRow == bbox.r1 - 1)
					oNewBorderProp = oNewBorder.t;
				else if(nCol == bbox.c2 + 1)
					oNewBorderProp = oNewBorder.r;
				else if(nRow == bbox.r2 + 1)
					oNewBorderProp = oNewBorder.b;
			}

			if(null != oNewBorderProp && null != oCurBorderProp && c_oAscBorderStyles.None != oCurBorderProp.s && (null == oNewBorder || c_oAscBorderStyles.None != oNewBorderProp.s) &&
				(oNewBorderProp.s != oCurBorderProp.s || oNewBorderProp.getRgbOrNull() != oCurBorderProp.getRgbOrNull())){
				var oTargetBorder = oCurBorder.clone();
				if(nCol == bbox.c1 - 1)
					oTargetBorder.r = new AscCommonExcel.BorderProp();
				else if(nRow == bbox.r1 - 1)
					oTargetBorder.b = new AscCommonExcel.BorderProp();
				else if(nCol == bbox.c2 + 1)
					oTargetBorder.l = new AscCommonExcel.BorderProp();
				else if(nRow == bbox.r2 + 1)
					oTargetBorder.t = new AscCommonExcel.BorderProp();
				oItemWithXfs.setBorder(oTargetBorder);
			}
		}
	};
	Range.prototype.setBorder=function(border){
		//border = null очисть border
		//"ih" - внутренние горизонтальные, "iv" - внутренние вертикальные
		AscCommon.History.Create_NewPoint();
		var _this = this;
		var oBBox = this.bbox;
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			var oAllCol = this.worksheet.getAllCol();
			_this._setRowColBorder(oBBox, oAllCol, false, border);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  _this._setRowColBorder(oBBox, row, true, border);
						  },
						  function(col){
							  _this._setRowColBorder(oBBox, col, false, border);
						  },
						  function(cell){
							  _this._setCellBorder(oBBox, cell, border);
						  });
		//убираем граничные border
		var aEdgeBorders = [];
		if(oBBox.c1 > 0 && (null == border || (border.l && !border.l.isEmpty())))
			aEdgeBorders.push(this.worksheet.getRange3(oBBox.r1, oBBox.c1 - 1, oBBox.r2, oBBox.c1 - 1));
		if(oBBox.r1 > 0 && (null == border || (border.t && !border.t.isEmpty())))
			aEdgeBorders.push(this.worksheet.getRange3(oBBox.r1 - 1, oBBox.c1, oBBox.r1 - 1, oBBox.c2));
		if(oBBox.c2 < gc_nMaxCol0 && (null == border || (border.r && !border.r.isEmpty())))
			aEdgeBorders.push(this.worksheet.getRange3(oBBox.r1, oBBox.c2 + 1, oBBox.r2, oBBox.c2 + 1));
		if(oBBox.r2 < gc_nMaxRow0 && (null == border || (border.b && !border.b.isEmpty())))
			aEdgeBorders.push(this.worksheet.getRange3(oBBox.r2 + 1, oBBox.c1, oBBox.r2 + 1, oBBox.c2));
		for(var i = 0, length = aEdgeBorders.length; i < length; i++)
		{
			var range = aEdgeBorders[i];
			range._setPropertyNoEmpty(function(row){
										  if(c_oRangeType.All == nRangeType && null == row.xfs)
											  return;
										  _this._setBorderEdge(oBBox, row, row.index, 0, border);
									  },
									  function(col){
										  _this._setBorderEdge(oBBox, col, 0, col.index, border);
									  },
									  function(cell){
										  _this._setBorderEdge(oBBox, cell, cell.nRow, cell.nCol, border);
									  });
		}
	};
	Range.prototype.setShrinkToFit=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setShrinkToFit(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setShrinkToFit(val);
						  },
						  function(col){
							  col.setShrinkToFit(val);
						  },
						  function(cell){
							  cell.setShrinkToFit(val);
						  });
	};
	Range.prototype.setWrap=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setWrap(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setWrap(val);
						  },
						  function(col){
							  col.setWrap(val);
						  },
						  function(cell){
							  cell.setWrap(val);
						  });
	};
	Range.prototype.setAngle=function(val){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			this.worksheet.getAllCol().setAngle(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function(row){
							  if(c_oRangeType.All == nRangeType && null == row.xfs)
								  return;
							  row.setAngle(val);
						  },
						  function(col){
							  col.setAngle(val);
						  },
						  function(cell){
							  cell.setAngle(val);
						  });
	};
	Range.prototype.setIndent = function (val) {
		if (val < 0) {
			return;
		}
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if (c_oRangeType.All == nRangeType) {
			this.worksheet.getAllCol().setAngle(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function (row) {
			if (c_oRangeType.All == nRangeType && null == row.xfs) {
				return;
			}
			row.setIndent(val);
		}, function (col) {
			col.setIndent(val);
		}, function (cell) {
			cell.setIndent(val);
		});
	};
	Range.prototype.setApplyProtection = function (val) {
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if (c_oRangeType.All == nRangeType) {
			this.worksheet.getAllCol().setApplyProtection(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function (row) {
			if (c_oRangeType.All == nRangeType && null == row.xfs) {
				return;
			}
			row.setApplyProtection(val);
		}, function (col) {
			col.setApplyProtection(val);
		}, function (cell) {
			cell.setApplyProtection(val);
		});
	};
	Range.prototype.setLocked = function (val) {
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if (c_oRangeType.All == nRangeType) {
			this.worksheet.getAllCol().setLocked(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function (row) {
			if (c_oRangeType.All == nRangeType && null == row.xfs) {
				return;
			}
			row.setLocked(val);
		}, function (col) {
			col.setLocked(val);
		}, function (cell) {
			cell.setLocked(val);
		});
	};
	Range.prototype.setHiddenFormulas = function (val) {
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if (c_oRangeType.All == nRangeType) {
			this.worksheet.getAllCol().setHiddenFormulas(val);
			fSetProperty = this._setPropertyNoEmpty;
		}
		fSetProperty.call(this, function (row) {
			if (c_oRangeType.All == nRangeType && null == row.xfs) {
				return;
			}
			row.setHiddenFormulas(val);
		}, function (col) {
			col.setHiddenFormulas(val);
		}, function (cell) {
			cell.setHiddenFormulas(val);
		});
	};
	Range.prototype.setType=function(type){
		AscCommon.History.Create_NewPoint();
		this.createCellOnRowColCross();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
			fSetProperty = this._setPropertyNoEmpty;
		fSetProperty.call(this, null, null,
						  function(cell){
							  cell.setType(type);
						  });
	};
	Range.prototype.changeTextCase=function(type){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();

		this._setPropertyNoEmpty(null, null,function(cell){
			cell.changeTextCase(type);
		});
		AscCommon.History.EndTransaction();
	};
	Range.prototype.getType=function(){
		var type;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			if(null != cell)
				type = cell.getType();
			else
				type = null;
		});
		return type;
	};
	Range.prototype.isNullText=function(){
		var isNullText;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			isNullText = (null != cell) ? cell.isNullText() : true;
		});
		return isNullText;
	};
	Range.prototype.isEmptyTextString=function(){
		var isEmptyTextString;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			isEmptyTextString = (null != cell) ? cell.isEmptyTextString() : true;
		});
		return isEmptyTextString;
	};
	Range.prototype.isNullTextString=function(){
		var isNullTextString;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			isNullTextString = (null != cell) ? cell.isNullTextString() : true;
		});
		return isNullTextString;
	};
	Range.prototype.isFormula=function(){
		var isFormula;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			isFormula = (null != cell) ? cell.isFormula() : false;
		});
		return isFormula;
	};
	Range.prototype.isFormulaContains=function(){
		var isFormula;
		this._foreachNoEmpty(function(cell) {
			isFormula = (null != cell) ? cell.isFormula() : false;
			return isFormula ? isFormula : null;
		});
		return isFormula;
	};
	Range.prototype.getFormula=function(){
		var formula;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			if(null != cell)
				formula = cell.getFormula();
			else
				formula = "";
		});
		return formula;
	};
	Range.prototype.getValueForEdit=function(checkFormulaArray){
		var t = this;
		var valueForEdit;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			if(null != cell)
			{
				valueForEdit = cell.getValueForEdit(checkFormulaArray);
			}
			else
				valueForEdit = "";
		});
		return valueForEdit;
	};
	Range.prototype.getValueForEdit2=function(){
		var t = this;
		var valueForEdit2;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			if(null != cell)
			{
				valueForEdit2 = cell.getValueForEdit2();
			}
			else
			{
				var xfs = null;
				t.worksheet._getRowNoEmpty(t.bbox.r1, function(row){
					var oCol = t.worksheet._getColNoEmptyWithAll(t.bbox.c1);
					if(row && null != row.xfs)
						xfs = row.xfs.clone();
					else if(null != oCol && null != oCol.xfs)
						xfs = oCol.xfs.clone();
				});
				var oTempCell = new Cell(t.worksheet);
				oTempCell.setRowCol(t.bbox.r1, t.bbox.c1);
				oTempCell.setStyleInternal(xfs);
				valueForEdit2 = oTempCell.getValueForEdit2();
			}
		});
		return valueForEdit2;
	};
	Range.prototype.getValueWithoutFormat=function(){
		var valueWithoutFormat;
		this.worksheet._getCellNoEmpty(this.bbox.r1, this.bbox.c1, function(cell) {
			if(null != cell)
				valueWithoutFormat = cell.getValueWithoutFormat();
			else
				valueWithoutFormat = "";
		});
		return valueWithoutFormat;
	};
	Range.prototype.getQuotePrefix=function(){
		let res = false;
		this.worksheet._getCellNoEmpty(this.bbox.r1, this.bbox.c1, function(cell) {
			if(null != cell && cell.getQuotePrefix())
				res = true;
		});
		return res;
	};
	Range.prototype.getValue=function(){
		return this.getValueWithoutFormat();
	};
	Range.prototype.getValueWithFormat=function(){
		var value;
		this.worksheet._getCellNoEmpty(this.bbox.r1, this.bbox.c1, function(cell) {
			if(null != cell)
				value = cell.getValue();
			else
				value = "";
		});
		return value;
	};
	Range.prototype.getValueWithFormatSkipToSpace=function(){
		var value;
		this.worksheet._getCellNoEmpty(this.bbox.r1, this.bbox.c1, function(cell) {
			if(null != cell)
				value = cell.getValueSkipToSpace();
			else
				value = "";
		});
		return value;
	};
	Range.prototype.getValue2=function(dDigitsCount, fIsFitMeasurer){
		//[{"text":"qwe","format":{"b":true, "i":false, "u":Asc.EUnderline.underlineNone, "s":false, "fn":"Arial", "fs": 12, "c": 0xff00ff, "va": "subscript"  }},{}...]
		var t = this;
		var value2;
		this.worksheet._getCellNoEmpty(this.bbox.r1, this.bbox.c1, function(cell){
			if(null != cell)
				value2 = cell.getValue2(dDigitsCount, fIsFitMeasurer);
			else
			{
				var xfs = null;
				t.worksheet._getRowNoEmpty(t.bbox.r1, function(row){
					var oCol = t.worksheet._getColNoEmptyWithAll(t.bbox.c1);

					if(row && null != row.xfs)
						xfs = row.xfs.clone();
					else if(null != oCol && null != oCol.xfs)
						xfs = oCol.xfs.clone();
				});
				var oTempCell = new Cell(t.worksheet);
				oTempCell.setRowCol(t.bbox.r1, t.bbox.c1);
				oTempCell.setStyleInternal(xfs);
				value2 = oTempCell.getValue2(dDigitsCount, fIsFitMeasurer);
			}
		});
		return value2;
	};
	Range.prototype.getNumberValue = function() {
		var numberValue;
		this.worksheet._getCellNoEmpty(this.bbox.r1, this.bbox.c1, function(cell) {
			numberValue = null != cell ? cell.getNumberValue() : null;
		});
		return numberValue;
	};
	Range.prototype.getValueData=function(){
		var res = null;
		this.worksheet._getCellNoEmpty(this.bbox.r1, this.bbox.c1, function(cell) {
			if(null != cell)
				res = cell.getValueData();
		});
		return res;
	};
	Range.prototype.getXfs = function (compiled) {
		var ws = this.worksheet;
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var xfs;
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			xfs = ws.getCompiledStyle(nRow, nCol, cell, compiled ? null : emptyStyleComponents);
		});
		return xfs || g_oDefaultFormat.xfs;
	};
	Range.prototype.getStyle = Range.prototype.getXfs;
	Range.prototype.getXfId = function () {
		var t = this;
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var XfId = g_oDefaultFormat.XfId;
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			var xfs = cell ? cell.getCompiledStyle() : t.worksheet.getCompiledStyle(nRow, nCol);
			if (xfs && null !== xfs.XfId) {
				XfId = xfs.XfId;
			}
		});
		return XfId;
	};
	Range.prototype.getStyleName=function(){
		var res = this.worksheet.workbook.CellStyles.getStyleNameByXfId(this.getXfId());

		// ToDo убрать эту заглушку (нужно делать на открытии) в InitStyleManager
		return res || this.worksheet.workbook.CellStyles.getStyleNameByXfId(g_oDefaultFormat.XfId);
	};
	Range.prototype.getTableStyle=function(){
		var tableStyle;
		this.worksheet._getCellNoEmpty(this.bbox.r1,this.bbox.c1, function(cell) {
			tableStyle = cell ? cell.getTableStyle() : null;
		});
		return tableStyle;
	};
	Range.prototype.getCompiledStyleCustom = function(needTable, needCell, needConditional) {
		return this.worksheet.getCompiledStyleCustom(this.bbox.r1,this.bbox.c1, needTable, needCell, needConditional);
	};
	Range.prototype.getNumFormat=function(){
		return oNumFormatCache.get(this.getNumFormatStr());
	};
	Range.prototype.getNumFormatStr = function () {
		var t = this;
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var numFormatStr = g_oDefaultFormat.Num.getFormat();
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			var xfs = cell ? cell.getCompiledStyle() : t.worksheet.getCompiledStyle(nRow, nCol);
			if (xfs && xfs.num) {
				numFormatStr = xfs.num.getFormat();
			}
		});
		return numFormatStr;
	};
	Range.prototype.getNumFormatType=function(){
		return this.getNumFormat().getType();
	};
// Узнаем отличается ли шрифт (размер и гарнитура) в ячейке от шрифта в строке
	Range.prototype.isNotDefaultFont = function () {
		// Получаем фонт ячейки
		var t = this;
		var cellFont = this.getFont();
		var rowFont = g_oDefaultFormat.Font;
		this.worksheet._getRowNoEmpty(this.bbox.r1, function(row) {
			if (row && null != row.xfs && null != row.xfs.font)
				rowFont = row.xfs.font;
			else if (null != t.worksheet.oAllCol && t.worksheet.oAllCol.xfs && t.worksheet.oAllCol.xfs.font)
				rowFont = t.worksheet.oAllCol.xfs.font;
		});

		return (cellFont.getName() !== rowFont.getName() || cellFont.getSize() !== rowFont.getSize());
	};
	Range.prototype.getFont = function (original) {
		// ToDo разобраться. Есть предположение, что эта функция не нужна и она работает не верно
		//  при выставлении стиля всему столбцу и тексте в ячейке
		var t = this;
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var font = g_oDefaultFormat.Font;
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			var xfs;
			if (cell) {
				xfs = original ? cell.getStyle() : cell.getCompiledStyle();
			} else {
				xfs = t.worksheet.getCompiledStyle(nRow, nCol, null, original ? emptyStyleComponents : null);
			}
			if (xfs && xfs.font) {
				font = xfs.font;
			}
		});
		return font;
	};
	Range.prototype.getAlignHorizontalByValue=function(align){
		//возвращает Align в зависимости от значния в ячейке
		//values:  none, center, justify, left , right, null
		var t = this;
		if(null == align){
			//пытаемся определить по значению
			var nRow = this.bbox.r1;
			var nCol = this.bbox.c1;
			this.worksheet._getCellNoEmpty(nRow, nCol, function(cell) {
				if(cell){
					switch(cell.getType()){
						case CellValueType.String:align = AscCommon.align_Left;break;
						case CellValueType.Bool:
						case CellValueType.Error:align = AscCommon.align_Center;break;
						default:
							//Если есть value и не проставлен тип значит это число, у всех остальных типов значение не null
							if(t.getValueWithoutFormat())
							{
								//смотрим
								var oNumFmt = t.getNumFormat();
								if(true == oNumFmt.isTextFormat())
									align = AscCommon.align_Left;
								else
									align = AscCommon.align_Right;
							}
							else
								align = AscCommon.align_Left;
							break;
					}
				}
			});
			if(null == align)
				align = AscCommon.align_Left;
		}
		return align;
	};

	Range.prototype.getAngle = function () {
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var angle;
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			var align = cell.getAlign();
			angle = align.getAngle();
		});
		return angle;
	}

	Range.prototype.getFillColor = function () {
		return this.getFill().bg();
	};
	Range.prototype.getFill = function () {
		var t = this;
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var fill = g_oDefaultFormat.Fill;
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			var xfs = cell ? cell.getCompiledStyle() : t.worksheet.getCompiledStyle(nRow, nCol);
			if (xfs && xfs.fill) {
				fill = xfs.fill;
			}
		});
		return fill;
	};
	Range.prototype.getLocked = function () {
		var t = this;
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var isLocked = true;
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			var xfs = cell ? cell.getCompiledStyle() : t.worksheet.getCompiledStyle(nRow, nCol);
			if (xfs) {
				isLocked = xfs.getLocked();
			}
		});
		return isLocked;
	};
	Range.prototype.getHidden = function () {
		var t = this;
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var isHidden = false;
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			var xfs = cell ? cell.getCompiledStyle() : t.worksheet.getCompiledStyle(nRow, nCol);
			if (xfs) {
				isHidden = xfs.getHidden();
			}
		});
		return isHidden;
	};
	Range.prototype.getBorderSrc = function (opt_row, opt_col) {
		//Возвращает как записано в файле, не проверяя бордеры соседних ячеек
		//формат
		//\{"l": {"s": "solid", "c": 0xff0000}, "t": {} ,"r": {} ,"b": {} ,"d": {},"dd": false ,"du": false }
		//"s" values: none, thick, thin, medium, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot
		//"dd" diagonal line, starting at the top left corner of the cell and moving down to the bottom right corner of the cell
		//"du" diagonal line, starting at the bottom left corner of the cell and moving up to the top right corner of the cell
		var t = this;
		var nRow = null != opt_row ? opt_row : this.bbox.r1;
		var nCol = null != opt_col ? opt_col : this.bbox.c1;
		var border = g_oDefaultFormat.Border;
		this.worksheet._getCellNoEmpty(nRow, nCol, function (cell) {
			var xfs = cell ? cell.getCompiledStyle() : t.worksheet.getCompiledStyle(nRow, nCol);
			if (xfs && xfs.border) {
				border = xfs.border;
			}
		});
		return border;
	};
	Range.prototype.getBorder = function (opt_row, opt_col) {
		//Возвращает как записано в файле, не проверяя бордеры соседних ячеек
		//формат
		//\{"l": {"s": "solid", "c": 0xff0000}, "t": {} ,"r": {} ,"b": {} ,"d": {},"dd": false ,"du": false }
		//"s" values: none, thick, thin, medium, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot
		//"dd" diagonal line, starting at the top left corner of the cell and moving down to the bottom right corner of the cell
		//"du" diagonal line, starting at the bottom left corner of the cell and moving up to the top right corner of the cell
		return this.getBorderSrc(opt_row, opt_col) || g_oDefaultFormat.Border;
	};
	Range.prototype.getBorderFull=function(){
		//Возвращает как excel, т.е. проверяет бордеры соседних ячеек
		//
		//\{"l": {"s": "solid", "c": 0xff0000}, "t": {} ,"r": {} ,"b": {} ,"d": {},"dd": false ,"du": false }
		//"s" values: none, thick, thin, medium, dashDot, dashDotDot, dashed, dotted, double, hair, mediumDashDot, mediumDashDotDot, mediumDashed, slantDashDot
		//
		//"dd" diagonal line, starting at the top left corner of the cell and moving down to the bottom right corner of the cell
		//"du" diagonal line, starting at the bottom left corner of the cell and moving up to the top right corner of the cell
		var borders = this.getBorder(this.bbox.r1, this.bbox.c1).clone();
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		if(!borders.l || borders.l.isEmpty()){
			if(nCol > 1){
				var left = this.getBorder(nRow, nCol - 1);
				if(left.r && !left.r.isEmpty())
					borders.l = left.r;
			}
		}
		if(!borders.t || borders.t.isEmpty()){
			if(nRow > 1){
				var top = this.getBorder(nRow - 1, nCol);
				if(top.b && !top.b.isEmpty())
					borders.t = top.b;
			}
		}
		if(!borders.r || borders.r.isEmpty()){
			var right = this.getBorder(nRow, nCol + 1);
			if(right.l && !right.l.isEmpty())
				borders.r = right.l;
		}
		if(!borders.b || borders.b.isEmpty()){
			var bottom = this.getBorder(nRow + 1, nCol);
			if(bottom.t && !bottom.t.isEmpty())
				borders.b = bottom.t;
		}
		return borders;
	};
	Range.prototype.getAlign=function(){
		//угол от -90 до 90 против часовой стрелки от оси OX
		var t = this;
		var nRow = this.bbox.r1;
		var nCol = this.bbox.c1;
		var align = g_oDefaultFormat.Align;
		this.worksheet._getCellNoEmpty(nRow, nCol, function(cell) {
			var xfs = cell ? cell.getCompiledStyle() : t.worksheet.getCompiledStyle(nRow, nCol);
			if (xfs) {
				align = xfs.align || g_oDefaultFormat.AlignAbs;
			}
		});
		return align;
	};
	Range.prototype.hasMerged=function(){
		var res = this.worksheet.mergeManager.getAny(this.bbox);
		return res ? res.bbox : null;
	};
	Range.prototype.mergeOpen=function(){
		this.worksheet.mergeManager.add(this.bbox, 1);
	};
	Range.prototype.merge=function(type){
		const t = this;
		if(null == type)
			type = Asc.c_oAscMergeOptions.Merge;
		var oBBox = this.bbox;
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		if(oBBox.r1 == oBBox.r2 && oBBox.c1 == oBBox.c2){
			if(type == Asc.c_oAscMergeOptions.MergeCenter)
				this.setAlignHorizontal(AscCommon.align_Center);
			AscCommon.History.EndTransaction();
			return;
		}
		if(this.hasMerged())
		{
			this.unmerge();
			if(type == Asc.c_oAscMergeOptions.MergeCenter)
			{
				//сбрасываем AlignHorizontal
				this.setAlignHorizontal(null);
				AscCommon.History.EndTransaction();
				return;
			}
		}

		//пробегаемся по границе диапазона, чтобы посмотреть какие границы нужно оставлять
		var oLeftBorder = null;
		var oTopBorder = null;
		var oRightBorder = null;
		var oBottomBorder = null;
		var nRangeType = this._getRangeType(oBBox);
		if(c_oRangeType.Range == nRangeType)
		{
			var oThis = this;
			var fGetBorder = function(bRow, v1, v2, v3, type)
			{
				var oRes = null;
				for(var i = v1; i <= v2; ++i)
				{
					var bNeedDelete = true;
					oThis.worksheet._getCellNoEmpty(bRow ? v3 : i, bRow ? i : v3, function(cell) {
						if(null != cell && null != cell.xfs && null != cell.xfs.border)
						{
							var border = cell.xfs.border;
							var oBorderProp;
							switch(type)
							{
								case 1: oBorderProp = border.l;break;
								case 2: oBorderProp = border.t;break;
								case 3: oBorderProp = border.r;break;
								case 4: oBorderProp = border.b;break;
							}
							if(oBorderProp && !oBorderProp.isEmpty())
							{
								if(null == oRes)
								{
									oRes = oBorderProp;
									bNeedDelete = false;
								}
								else if(true == oRes.isEqual(oBorderProp))
									bNeedDelete = false;
							}
						}
					});
					if(bNeedDelete)
					{
						oRes = null;
						break;
					}
				}
				return oRes;
			};
			oLeftBorder = fGetBorder(false, oBBox.r1, oBBox.r2, oBBox.c1, 1);
			oTopBorder = fGetBorder(true, oBBox.c1, oBBox.c2, oBBox.r1, 2);
			oRightBorder = fGetBorder(false, oBBox.r1, oBBox.r2, oBBox.c2, 3);
			oBottomBorder = fGetBorder(true, oBBox.c1, oBBox.c2, oBBox.r2, 4);
		}
		else if(c_oRangeType.Row == nRangeType)
		{
			this.worksheet._getRowNoEmpty(oBBox.r1, function(row){
				if(row && null != row.xfs && null != row.xfs.border && row.xfs.border.t && !row.xfs.border.t.isEmpty())
					oTopBorder = row.xfs.border.t;
			});
			this.worksheet._getRowNoEmpty(oBBox.r2, function(row){
				if(row && null != row.xfs && null != row.xfs.border &&  row.xfs.border.b && !row.xfs.border.b.isEmpty())
					oBottomBorder = row.xfs.border.b;
			});
		}
		else
		{
			var oLeftCol = this.worksheet._getColNoEmptyWithAll(oBBox.c1);
			if(null != oLeftCol && null != oLeftCol.xfs && null != oLeftCol.xfs.border && oLeftCol.xfs.border.l && !oLeftCol.xfs.border.l.isEmpty())
				oLeftBorder = oLeftCol.xfs.border.l;
			var oRightCol = this.worksheet._getColNoEmptyWithAll(oBBox.c2);
			if(null != oRightCol && null != oRightCol.xfs && null != oRightCol.xfs.border && oRightCol.xfs.border.r && !oRightCol.xfs.border.r.isEmpty())
				oRightBorder = oRightCol.xfs.border.r;
		}

		var bFirst = true;
		var oLeftTopCellStyle = null;
		var oFirstCellStyle = null;
		var oFirstCellValue = null;
		var oFirstCellRow = null;
		var oFirstCellCol = null;
		var oFirstCellHyperlink = null;

		let error = false;
		let firstMergeCellIndex;

		this._setPropertyNoEmpty(null, null,
			function(cell, nRow0, nCol0, nRowStart, nColStart) {
				if (!error) {
					let formulaRef = cell.formulaParsed && cell.formulaParsed.getArrayFormulaRef();

					if (bFirst && false == cell.isNullText()) {
						if (formulaRef && !formulaRef.isOneCell()) {
							/* if an array larger than 1x1 cell is encountered, then no value is written in ms */
							bFirst = false;
							oFirstCellValue = null;
						} else {
							bFirst = false;
							oFirstCellStyle = cell.getStyle();
							oFirstCellValue = cell.getValueData();
							oFirstCellRow = cell.nRow;
							oFirstCellCol = cell.nCol;

							/* we write down the coordinates of the first cell with data, that we will skip when clearing dependencies */
							firstMergeCellIndex = AscCommonExcel.getCellIndex(oFirstCellRow, oFirstCellCol);
						}
					}

					/* checking whether the ref is included in the selected merge area (if not, then we return an error) */
					if (formulaRef && !t.bbox.containsRange(formulaRef)) {
						error = c_oAscError.ID.CannotChangeFormulaArray;
						return false;
					} else if (nRow0 == nRowStart && nCol0 == nColStart) {
						if (formulaRef && !formulaRef.isOneCell()) {
							error = c_oAscError.ID.CannotChangeFormulaArray;
							return false;
						} else {
							oLeftTopCellStyle = cell.getStyle();
						}
					}
				}
			});

		if (error) {
			/* Complete the transaction in history before exiting the function merge */
			AscCommon.History.EndTransaction();
			return {errorType: error};
		}

		this.worksheet.workbook.handlers.trigger(
			"changeDocument", 
			AscCommonExcel.docChangedType.mergeRange, 
			true, /* arg1 */
			this.bbox, /* arg2 */
			this.worksheet.getId(), 
			firstMergeCellIndex
		);
		//правила работы с гиперссылками во время merge(отличются от Excel в случаем областей, например hyperlink: C3:D3 мержим C2:C3)
		// 1)оставляем все ссылки, которые не полностью лежат в merge области
		// 2)оставляем многоклеточные ссылки, top граница которых совпадает с top границей merge области, а высота merge > 1 или совпадает с высотой области merge
		// 3)оставляем и переносим в первую ячейку одну одноклеточную ссылку, если она находится в первой ячейке с данными
		var aHyperlinks = this.worksheet.hyperlinkManager.get(oBBox);
		var aHyperlinksToRestore = [];
		for(var i = 0, length = aHyperlinks.inner.length; i < length; i++)
		{
			var elem = aHyperlinks.inner[i];
			if(oFirstCellRow == elem.bbox.r1 && oFirstCellCol == elem.bbox.c1 && elem.bbox.r1 == elem.bbox.r2 && elem.bbox.c1 == elem.bbox.c2)
			{
				var oNewHyperlink = elem.data.clone();
				oNewHyperlink.Ref.setOffset(new AscCommon.CellBase(oBBox.r1 - oFirstCellRow, oBBox.c1 - oFirstCellCol));
				aHyperlinksToRestore.push(oNewHyperlink);
			}
			else if( oBBox.r1 == elem.bbox.r1 && (elem.bbox.r1 != elem.bbox.r2 || (elem.bbox.c1 != elem.bbox.c2 && oBBox.r1 == oBBox.r2)))
				aHyperlinksToRestore.push(elem.data);
		}
		this.cleanAll();
		//восстанавливаем hyperlink
		for(var i = 0, length = aHyperlinksToRestore.length; i < length; i++)
		{
			var elem = aHyperlinksToRestore[i];
			this.worksheet.hyperlinkManager.add(elem.Ref.getBBox0(), elem);
		}
		var oTargetStyle = null;
		if(null != oFirstCellValue && null != oFirstCellRow && null != oFirstCellCol)
		{
			if(null != oFirstCellStyle)
				oTargetStyle = oFirstCellStyle.clone();
			this.worksheet._getCell(oBBox.r1, oBBox.c1, function(cell){
				cell.setValueData(oFirstCellValue);
			});
			if(null != oFirstCellHyperlink)
			{
				var oLeftTopRange = this.worksheet.getCell3(oBBox.r1, oBBox.c1);
				oLeftTopRange.setHyperlink(oFirstCellHyperlink, true);
			}
		}
		else if(null != oLeftTopCellStyle)
			oTargetStyle = oLeftTopCellStyle.clone();

		//убираем бордеры
		if(null != oTargetStyle)
		{
			if(null != oTargetStyle.border)
				oTargetStyle.border = null;
		}
		else if(null != oLeftBorder || null != oTopBorder || null != oRightBorder || null != oBottomBorder)
			oTargetStyle = new AscCommonExcel.CellXfs();
		var fSetProperty = this._setProperty;
		var nRangeType = this._getRangeType();
		if(c_oRangeType.All == nRangeType)
		{
			fSetProperty = this._setPropertyNoEmpty;
			oTargetStyle = null
		}
		fSetProperty.call(this, function(row){
							  if(null == oTargetStyle)
								  row.setStyle(null);
							  else
							  {
								  var oNewStyle = oTargetStyle.clone();
								  if(row.index == oBBox.r1 && null != oTopBorder)
								  {
									  oNewStyle.border = new Border();
									  oNewStyle.border.t = oTopBorder.clone();
									  if(row.index == oBBox.r2 && null != oBottomBorder) {
										  oNewStyle.border.b = oBottomBorder.clone();
									  }
								  }
								  else if(row.index == oBBox.r2 && null != oBottomBorder)
								  {
									  oNewStyle.border = new Border();
									  oNewStyle.border.b = oBottomBorder.clone();
								  }
								  row.setStyle(oNewStyle);
							  }
						  },function(col){
							  if(null == oTargetStyle)
								  col.setStyle(null);
							  else
							  {
								  var oNewStyle = oTargetStyle.clone();
								  if(col.index == oBBox.c1 && null != oLeftBorder)
								  {
									  oNewStyle.border = new Border();
									  oNewStyle.border.l = oLeftBorder.clone();
									  if(col.index == oBBox.c2 && null != oRightBorder) {
										  oNewStyle.border.r = oRightBorder.clone();
									  }
								  }
								  else if(col.index == oBBox.c2 && null != oRightBorder)
								  {
									  oNewStyle.border = new Border();
									  oNewStyle.border.r = oRightBorder.clone();
								  }
								  col.setStyle(oNewStyle);
							  }
						  },
						  function(cell, nRow, nCol, nRowStart, nColStart){
							  //важно установить именно здесь, чтобы ячейка не удалилась после применения стилей.
							  if(null == oTargetStyle)
								  cell.setStyle(null);
							  else
							  {
								  var oNewStyle = oTargetStyle.clone();
								  if(oBBox.r1 == nRow && oBBox.c1 == nCol)
								  {
									  if(null != oLeftBorder || null != oTopBorder || (oBBox.r1 == oBBox.r2 && null != oBottomBorder) || (oBBox.c1 == oBBox.c2 && null != oRightBorder))
									  {
										  oNewStyle.border = new Border();
										  if(null != oLeftBorder)
											  oNewStyle.border.l = oLeftBorder.clone();
										  if(null != oTopBorder)
											  oNewStyle.border.t = oTopBorder.clone();
										  if(oBBox.r1 == oBBox.r2 && null != oBottomBorder)
											  oNewStyle.border.b = oBottomBorder.clone();
										  if(oBBox.c1 == oBBox.c2 && null != oRightBorder)
											  oNewStyle.border.r = oRightBorder.clone();
									  }
								  }
								  else if(oBBox.r1 == nRow && oBBox.c2 == nCol)
								  {
									  if(null != oRightBorder || null != oTopBorder || (oBBox.r1 == oBBox.r2 && null != oBottomBorder))
									  {
										  oNewStyle.border = new Border();
										  if(null != oRightBorder)
											  oNewStyle.border.r = oRightBorder.clone();
										  if(null != oTopBorder)
											  oNewStyle.border.t = oTopBorder.clone();
										  if(oBBox.r1 == oBBox.r2 && null != oBottomBorder)
											  oNewStyle.border.b = oBottomBorder.clone();
									  }
								  }
								  else if(oBBox.r2 == nRow && oBBox.c1 == nCol)
								  {
									  if(null != oLeftBorder || null != oBottomBorder || (oBBox.c1 == oBBox.c2 && null != oRightBorder))
									  {
										  oNewStyle.border = new Border();
										  if(null != oLeftBorder)
											  oNewStyle.border.l = oLeftBorder.clone();
										  if(null != oBottomBorder)
											  oNewStyle.border.b = oBottomBorder.clone();
										  if(oBBox.c1 == oBBox.c2 && null != oRightBorder)
											  oNewStyle.border.r = oRightBorder.clone();
									  }
								  }
								  else if(oBBox.r2 == nRow && oBBox.c2 == nCol)
								  {
									  if(null != oRightBorder || null != oBottomBorder)
									  {
										  oNewStyle.border = new Border();
										  if(null != oRightBorder)
											  oNewStyle.border.r = oRightBorder.clone();
										  if(null != oBottomBorder)
											  oNewStyle.border.b = oBottomBorder.clone();
									  }
								  }
								  else if(oBBox.r1 == nRow)
								  {
									  if(null != oTopBorder || (oBBox.r1 == oBBox.r2 && null != oBottomBorder))
									  {
										  oNewStyle.border = new Border();
										  if(null != oTopBorder)
											  oNewStyle.border.t = oTopBorder.clone();
										  if(oBBox.r1 == oBBox.r2 && null != oBottomBorder)
											  oNewStyle.border.b = oBottomBorder.clone();
									  }
								  }
								  else if(oBBox.r2 == nRow)
								  {
									  if(null != oBottomBorder)
									  {
										  oNewStyle.border = new Border();
										  oNewStyle.border.b = oBottomBorder.clone();
									  }
								  }
								  else if(oBBox.c1 == nCol)
								  {
									  if(null != oLeftBorder || (oBBox.c1 == oBBox.c2 && null != oRightBorder))
									  {
										  oNewStyle.border = new Border();
										  if(null != oLeftBorder)
											  oNewStyle.border.l = oLeftBorder.clone();
										  if(oBBox.c1 == oBBox.c2 && null != oRightBorder)
											  oNewStyle.border.r = oRightBorder.clone();
									  }
								  }
								  else if(oBBox.c2 == nCol)
								  {
									  if(null != oRightBorder)
									  {
										  oNewStyle.border = new Border();
										  oNewStyle.border.r = oRightBorder.clone();
									  }
								  }
								  cell.setStyle(oNewStyle);
							  }
						  });
		if(type == Asc.c_oAscMergeOptions.MergeCenter)
			this.setAlignHorizontal(AscCommon.align_Center);
		if(false == this.worksheet.workbook.bUndoChanges && false == this.worksheet.workbook.bRedoChanges)
			this.worksheet.mergeManager.add(this.bbox, 1);

		//сбрасываем dataValidation кроме 1 ячейки
		var dataValidationRanges = Asc.Range(this.bbox.c1, this.bbox.r1, this.bbox.c1, this.bbox.r1).difference(this.bbox);
		if (dataValidationRanges) {
			this.worksheet.clearDataValidation(dataValidationRanges, true);
		}
		this.worksheet.workbook.handlers.trigger(
			"changeDocument", 
			AscCommonExcel.docChangedType.mergeRange, 
			null, /* arg1 */
			this.bbox, /* arg2 */
			this.worksheet.getId(), 
			firstMergeCellIndex
		);

		AscCommon.History.EndTransaction();
	};
	Range.prototype.unmerge=function(bOnlyInRange){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		if(false == this.worksheet.workbook.bUndoChanges && false == this.worksheet.workbook.bRedoChanges)
			this.worksheet.mergeManager.remove(this.bbox);
		AscCommon.History.EndTransaction();
	};
	Range.prototype._getHyperlinks=function(){
		var nRangeType = this._getRangeType();
		var result = [];
		var oThis = this;
		if(c_oRangeType.Range == nRangeType)
		{
			var oTempRows = {};
			var fAddToTempRows = function(oTempRows, bbox, data){
				if(null != bbox)
				{
					for(var i = bbox.r1; i <= bbox.r2; i++)
					{
						var row = oTempRows[i];
						if(null == row)
						{
							row = {};
							oTempRows[i] = row;
						}
						for(var j = bbox.c1; j <= bbox.c2; j++)
						{
							var cell = row[j];
							if(null == cell)
								row[j] = data;
						}
					}
				}
			};
			//todo возможно надо сделать оптимизацию для скрытых строк
			var aHyperlinks = this.worksheet.hyperlinkManager.get(this.bbox);
			for(var i = 0, length = aHyperlinks.all.length; i < length; i++)
			{
				var hyp = aHyperlinks.all[i];
				var hypBBox = hyp.bbox.intersectionSimple(this.bbox);
				fAddToTempRows(oTempRows, hypBBox, hyp.data);
				//расширяем гиперссылки на merge ячейках
				var aMerged = this.worksheet.mergeManager.get(hyp.bbox);
				for(var j = 0, length2 = aMerged.all.length; j < length2; j++)
				{
					var merge = aMerged.all[j];
					var mergeBBox = merge.bbox.intersectionSimple(this.bbox);
					fAddToTempRows(oTempRows, mergeBBox, hyp.data);
				}
			}
			//формируем результат
			for(var i in oTempRows)
			{
				var nRowIndex = i - 0;
				var row = oTempRows[i];
				for(var j in row)
				{
					var nColIndex = j - 0;
					var oCurHyp = row[j];
					result.push({hyperlink: oCurHyp, col: nColIndex, row: nRowIndex});
				}
			}
		}
		return result;
	};
	Range.prototype.getHyperlink=function(){
		var aHyperlinks = this._getHyperlinks();
		if(null != aHyperlinks && aHyperlinks.length > 0)
			return aHyperlinks[0].hyperlink;
		return null;
	};
	Range.prototype.getHyperlinks=function(){
		return this._getHyperlinks();
	};
	Range.prototype.setHyperlinkOpen=function(val){
		if(null != val && false == val.isValid())
			return;
		this.worksheet.hyperlinkManager.add(val.Ref.getBBox0(), val);
	};
	Range.prototype.setHyperlink=function(val, bWithoutStyle){
		if(null != val && false == val.isValid())
			return;

		//проверяем, может эта ссылка уже существует
		var i, length, hyp;
		var bExist = false;
		var aHyperlinks = this.worksheet.hyperlinkManager.get(this.bbox);
		for(i = 0, length = aHyperlinks.all.length; i < length; i++)
		{
			hyp = aHyperlinks.all[i];
			if(hyp.data.isEqual(val))
			{
				bExist = true;
				break;
			}
		}
		if(false == bExist)
		{
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();
			if(false == this.worksheet.workbook.bUndoChanges && false == this.worksheet.workbook.bRedoChanges)
			{
				//удаляем ссылки с тем же адресом
				for(i = 0, length = aHyperlinks.all.length; i < length; i++)
				{
					hyp = aHyperlinks.all[i];
					if(hyp.bbox.isEqual(this.bbox))
						this.worksheet.hyperlinkManager.removeElement(hyp);
				}
			}
			//todo перейти на CellStyle
			if(true != bWithoutStyle)
			{
				var oHyperlinkFont = new AscCommonExcel.Font();
				oHyperlinkFont.setName(this.worksheet.workbook.getDefaultFont());
				oHyperlinkFont.setSize(this.worksheet.workbook.getDefaultSize());
				oHyperlinkFont.setUnderline(Asc.EUnderline.underlineSingle);
				oHyperlinkFont.setColor(AscCommonExcel.g_oColorManager.getThemeColor(AscCommonExcel.g_nColorHyperlink));
				this.setFont(oHyperlinkFont);
			}
			if(false == this.worksheet.workbook.bUndoChanges && false == this.worksheet.workbook.bRedoChanges)
				this.worksheet.hyperlinkManager.add(val.Ref.getBBox0(), val);
			AscCommon.History.EndTransaction();
		}
	};
	Range.prototype.removeHyperlink = function (elem, removeStyle) {
		var bbox = this.bbox;
		if(false == this.worksheet.workbook.bUndoChanges && false == this.worksheet.workbook.bRedoChanges)
		{
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();
			var oChangeParam = { removeStyle: removeStyle };
			if(elem)
				this.worksheet.hyperlinkManager.removeElement(elem, oChangeParam);
			else
				this.worksheet.hyperlinkManager.remove(bbox, !bbox.isOneCell(), oChangeParam);
			AscCommon.History.EndTransaction();
		}
	};
	Range.prototype.deleteCellsShiftUp=function(preDeleteAction){
		return this._shiftUpDown(true, preDeleteAction);
	};
	Range.prototype.addCellsShiftBottom=function(displayNameFormatTable){
		return this._shiftUpDown(false, null, displayNameFormatTable);
	};
	Range.prototype.addCellsShiftRight=function(displayNameFormatTable){
		return this._shiftLeftRight(false, null,displayNameFormatTable);
	};
	Range.prototype.deleteCellsShiftLeft=function(preDeleteAction){
		return this._shiftLeftRight(true, preDeleteAction);
	};
	Range.prototype._canShiftLeftRight=function(bLeft){
		var aColsToDelete = [], aCellsToDelete = [];
		var oBBox = this.bbox;
		var nRangeType = this._getRangeType(oBBox);
		if(c_oRangeType.Range != nRangeType && c_oRangeType.Col != nRangeType)
			return null;

		var nWidth = oBBox.c2 - oBBox.c1 + 1;
		if(!bLeft && !this.worksheet.workbook.bUndoChanges && !this.worksheet.workbook.bRedoChanges){
			var rangeEdge = this.worksheet.getRange3(oBBox.r1, gc_nMaxCol0 - nWidth + 1, oBBox.r2, gc_nMaxCol0);
			var aMerged = this.worksheet.mergeManager.get(rangeEdge.bbox);
			if(aMerged.all.length > 0)
				return null;
			var aHyperlink = this.worksheet.hyperlinkManager.get(rangeEdge.bbox);
			if(aHyperlink.all.length > 0)
				return null;

			var bError = rangeEdge._setPropertyNoEmpty(null, function(col){
				if(null != col){
					if(null != col && null != col.xfs && null != col.xfs.fill && col.xfs.fill.notEmpty())
						return true;
					aColsToDelete.push(col);
				}
			}, function(cell){
				if(null != cell){
					if(null != cell.xfs && null != cell.xfs.fill && cell.xfs.fill.notEmpty())
						return true;
					if(!cell.isNullText())
						return true;
					aCellsToDelete.push(cell.nRow, cell.nCol);
				}
			});
			if(bError)
				return null;
		}
		return {aColsToDelete: aColsToDelete, aCellsToDelete: aCellsToDelete};
	};
	Range.prototype._shiftLeftRight=function(bLeft, preDeleteAction, displayNameFormatTable){
		var canShiftRes = this._canShiftLeftRight(bLeft);
		if(null === canShiftRes)
			return false;

		if (preDeleteAction)
			preDeleteAction();

		//удаляем крайние колонки и ячейки
		var i, length, colIndex;
		for(i = 0, length = canShiftRes.aColsToDelete.length; i < length; ++i){
			colIndex = canShiftRes.aColsToDelete[i].index;
			this.worksheet._removeCols(colIndex, colIndex);
		}
		for(i = 0; i < canShiftRes.aCellsToDelete.length; i+=2)
			this.worksheet._removeCell(canShiftRes.aCellsToDelete[i], canShiftRes.aCellsToDelete[i + 1]);

		var oBBox = this.bbox;
		var nWidth = oBBox.c2 - oBBox.c1 + 1;
		var nRangeType = this._getRangeType(oBBox);
		var mergeManager = this.worksheet.mergeManager;
		this.worksheet.workbook.dependencyFormulas.lockRecal();
		//todo вставить предупреждение, что будет unmerge
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		var oShiftGet = null;
		if (false == this.worksheet.workbook.bUndoChanges && (false == this.worksheet.workbook.bRedoChanges || this.worksheet.workbook.bCollaborativeChanges))
		{
			AscCommon.History.LocalChange = true;
			oShiftGet = mergeManager.shiftGet(this.bbox, true);
			var aMerged = oShiftGet.elems;
			if(null != aMerged.outer && aMerged.outer.length > 0)
			{
				var bChanged = false;
				for(i = 0, length = aMerged.outer.length; i < length; i++)
				{
					var elem = aMerged.outer[i];
					if(!(elem.bbox.c1 < oShiftGet.bbox.c1 && oShiftGet.bbox.r1 <= elem.bbox.r1 && elem.bbox.r2 <= oShiftGet.bbox.r2))
					{
						mergeManager.removeElement(elem);
						bChanged = true;
					}
				}
				if(bChanged)
					oShiftGet = null;
			}
			AscCommon.History.LocalChange = false;
		}
		//сдвигаем ячейки
		if(bLeft)
		{
			if(c_oRangeType.Range == nRangeType)
				this.worksheet._shiftCellsLeft(oBBox);
			else
				this.worksheet._removeCols(oBBox.c1, oBBox.c2);
		}
		else
		{
			if(c_oRangeType.Range == nRangeType)
				this.worksheet._shiftCellsRight(oBBox, displayNameFormatTable);
			else
				this.worksheet._insertColsBefore(oBBox.c1, nWidth);
		}
		if (false == this.worksheet.workbook.bUndoChanges && (false == this.worksheet.workbook.bRedoChanges || this.worksheet.workbook.bCollaborativeChanges))
		{
			AscCommon.History.LocalChange = true;
			mergeManager.shift(this.bbox, !bLeft, true, oShiftGet);
			this.worksheet.hyperlinkManager.shift(this.bbox, !bLeft, true);
			AscCommon.History.LocalChange = false;
		}
		AscCommon.History.EndTransaction();
		this.worksheet.workbook.dependencyFormulas.unlockRecal();
		return true;
	};
	Range.prototype._canShiftUpDown=function(bUp){
		var aRowsToDelete = [], aCellsToDelete = [];
		var oBBox = this.bbox;
		var nRangeType = this._getRangeType(oBBox);
		if(c_oRangeType.Range != nRangeType && c_oRangeType.Row != nRangeType)
			return null;

		var nHeight = oBBox.r2 - oBBox.r1 + 1;
		if(!bUp && !this.worksheet.workbook.bUndoChanges && !this.worksheet.workbook.bRedoChanges){
			var rangeEdge = this.worksheet.getRange3(gc_nMaxRow0 - nHeight + 1, oBBox.c1, gc_nMaxRow0, oBBox.c2);
			var aMerged = this.worksheet.mergeManager.get(rangeEdge.bbox);
			if(aMerged.all.length > 0)
				return null;
			var aHyperlink = this.worksheet.hyperlinkManager.get(rangeEdge.bbox);
			if(aHyperlink.all.length > 0)
				return null;

			var bError = rangeEdge._setPropertyNoEmpty(function(row){
				if(null != row){
					if(null != row.xfs && null != row.xfs.fill && row.xfs.fill.notEmpty())
						return true;
					aRowsToDelete.push(row.index);
				}
			}, null,  function(cell){
				if(null != cell){
					if(null != cell.xfs && null != cell.xfs.fill && cell.xfs.fill.notEmpty())
						return true;
					if(!cell.isNullText())
						return true;
					aCellsToDelete.push(cell.nRow, cell.nCol);
				}
			});
			if(bError)
				return null;
		}
		return {aRowsToDelete: aRowsToDelete, aCellsToDelete: aCellsToDelete};
	};
	Range.prototype._shiftUpDown = function (bUp, preDeleteAction, displayNameFormatTable) {
		var canShiftRes = this._canShiftUpDown(bUp);
		if(null === canShiftRes)
			return false;

		if (preDeleteAction)
			preDeleteAction();

		//удаляем крайние колонки и ячейки
		var i, length, rowIndex;
		for(i = 0, length = canShiftRes.aRowsToDelete.length; i < length; ++i){
			rowIndex = canShiftRes.aRowsToDelete[i];
			this.worksheet._removeRows(rowIndex, rowIndex);
		}
		for(i = 0; i < canShiftRes.aCellsToDelete.length; i+=2)
			this.worksheet._removeCell(canShiftRes.aCellsToDelete[i], canShiftRes.aCellsToDelete[i + 1]);

		var oBBox = this.bbox;
		var nHeight = oBBox.r2 - oBBox.r1 + 1;
		var nRangeType = this._getRangeType(oBBox);
		var mergeManager = this.worksheet.mergeManager;
		this.worksheet.workbook.dependencyFormulas.lockRecal();
		//todo вставить предупреждение, что будет unmerge
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		var oShiftGet = null;
		if (false == this.worksheet.workbook.bUndoChanges && (false == this.worksheet.workbook.bRedoChanges || this.worksheet.workbook.bCollaborativeChanges))
		{
			AscCommon.History.LocalChange = true;
			oShiftGet = mergeManager.shiftGet(this.bbox, false);
			var aMerged = oShiftGet.elems;
			if(null != aMerged.outer && aMerged.outer.length > 0)
			{
				var bChanged = false;
				for(i = 0, length = aMerged.outer.length; i < length; i++)
				{
					var elem = aMerged.outer[i];
					if(!(elem.bbox.r1 < oShiftGet.bbox.r1 && oShiftGet.bbox.c1 <= elem.bbox.c1 && elem.bbox.c2 <= oShiftGet.bbox.c2))
					{
						mergeManager.removeElement(elem);
						bChanged = true;
					}
				}
				if(bChanged)
					oShiftGet = null;
			}
			AscCommon.History.LocalChange = false;
		}
		//сдвигаем ячейки
		if(bUp)
		{
			if(c_oRangeType.Range == nRangeType)
				this.worksheet._shiftCellsUp(oBBox);
			else
				this.worksheet._removeRows(oBBox.r1, oBBox.r2);
		}
		else
		{
			if(c_oRangeType.Range == nRangeType)
				this.worksheet._shiftCellsBottom(oBBox, displayNameFormatTable);
			else
				this.worksheet._insertRowsBefore(oBBox.r1, nHeight);
		}
		if (false == this.worksheet.workbook.bUndoChanges && (false == this.worksheet.workbook.bRedoChanges || this.worksheet.workbook.bCollaborativeChanges))
		{
			AscCommon.History.LocalChange = true;
			mergeManager.shift(this.bbox, !bUp, false, oShiftGet);
			this.worksheet.hyperlinkManager.shift(this.bbox, !bUp, false);
			AscCommon.History.LocalChange = false;
		}
		AscCommon.History.EndTransaction();
		this.worksheet.workbook.dependencyFormulas.unlockRecal();
		return true;
	};
	Range.prototype.setOffset=function(offset){
		this.bbox.setOffset(offset);
	};
	Range.prototype.setOffsetFirst=function(offset){
		this.bbox.setOffsetFirst(offset);
	};
	Range.prototype.setOffsetLast=function(offset){
		this.bbox.setOffsetLast(offset);
	};
	Range.prototype.setOffsetWithAbs = function() {
		this.bbox.setOffsetWithAbs.apply(this.bbox, arguments);
	};
	Range.prototype.intersect=function(range){
		var oBBox1 = this.bbox;
		var oBBox2 = range.bbox;
		var r1 = Math.max(oBBox1.r1, oBBox2.r1);
		var c1 = Math.max(oBBox1.c1, oBBox2.c1);
		var r2 = Math.min(oBBox1.r2, oBBox2.r2);
		var c2 = Math.min(oBBox1.c2, oBBox2.c2);
		if(r1 <= r2 && c1 <= c2)
			return this.worksheet.getRange3(r1, c1, r2, c2);
		return null;
	};
	Range.prototype.cleanFormat=function(ignoreNoEmpty){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		this.unmerge();

		let fSetProperty = ignoreNoEmpty ? this._setProperty : this._setPropertyNoEmpty;
		fSetProperty.call(this, function(row){
			row.setStyle(null);
			// if(row.isEmpty())
			// row.Remove();
		},function(col){
			col.setStyle(null);
			// if(col.isEmpty())
			// col.Remove();
		},function(cell, nRow0, nCol0, nRowStart, nColStart){
			cell.setStyle(null);
			// if(cell.isEmpty())
			// cell.Remove();
		});
		AscCommon.History.EndTransaction();
	};
	Range.prototype.cleanText = function () {
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		this._setPropertyNoEmpty(null, null,
			function (cell, nRow0, nCol0, nRowStart, nColStart) {
				cell.setValue("");
				// if(cell.isEmpty())
				// cell.Remove();
			});
		AscCommon.History.EndTransaction();
	};
	Range.prototype.cleanTextExceptFirst = function () {
		History.Create_NewPoint();
		History.StartTransaction();
		this._setPropertyNoEmpty(null, null,
			function (cell, nRow0, nCol0, nRowStart, nColStart) {
				if (!(cell.nRow === nRowStart && cell.nCol === nColStart)) {
					cell.setValue("");
				}
			});
		History.EndTransaction();
	};
	Range.prototype.cleanAll=function(ignoreNoEmpty){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		this.unmerge();
		//удаляем только гиперссылки, которые полностью лежат в области
		var aHyperlinks = this.worksheet.hyperlinkManager.get(this.bbox);
		for(var i = 0, length = aHyperlinks.inner.length; i < length; ++i)
			this.removeHyperlink(aHyperlinks.inner[i]);
		var oThis = this;
		let fSetProperty = ignoreNoEmpty ? this._setProperty : this._setPropertyNoEmpty;
		fSetProperty.call(this, function(row){
			row.setStyle(null);
			// if(row.isEmpty())
			// row.Remove();
		},function(col){
			col.setStyle(null);
			// if(col.isEmpty())
			// col.Remove();
		},function(cell, nRow0, nCol0, nRowStart, nColStart){
			oThis.worksheet._removeCell(nRow0, nCol0, cell);
		});

		this.worksheet.workbook.dependencyFormulas.calcTree();
		AscCommon.History.EndTransaction();
	};
	Range.prototype.cleanHyperlinks=function(){
		AscCommon.History.Create_NewPoint();
		AscCommon.History.StartTransaction();
		//удаляем только гиперссылки, которые полностью лежат в области
		var aHyperlinks = this.worksheet.hyperlinkManager.get(this.bbox);
		for(var i = 0, length = aHyperlinks.inner.length; i < length; ++i)
			this.removeHyperlink(aHyperlinks.inner[i]);
		AscCommon.History.EndTransaction();
	};
	Range.prototype.sort = function (nOption, nStartRowCol, sortColor, opt_guessHeader, opt_by_row, opt_custom_sort) {
		var bbox = this.bbox;
		if (opt_guessHeader) {
			//если тип ячеек первого и второго row попарно совпадает, то считаем первую строку заголовком
			//todo рассмотреть замерженые ячейки. стили тоже влияют, но непонятно как сравнивать border
			var bIgnoreFirstRow = ignoreFirstRowSort(this.worksheet, bbox);

			if (bIgnoreFirstRow) {
				bbox = bbox.clone();
				bbox.r1++;
			}
		}

		//todo горизонтальная сортировка
		var aMerged = this.worksheet.mergeManager.get(bbox);
		if (aMerged.outer.length > 0 || (aMerged.inner.length > 0 && null == _isSameSizeMerged(bbox, aMerged.inner, true))) {
			return null;
		}

		var nMergedWidth = 1;
		var nMergedHeight = 1;
		if (aMerged.inner.length > 0) {
			var merged = aMerged.inner[0];
			if (opt_by_row) {
				nMergedWidth = merged.bbox.c2 - merged.bbox.c1 + 1;
				//меняем nStartCol, потому что приходит колонка той ячейки, на которой начали выделение
				nStartRowCol = merged.bbox.r1;
			} else {
				nMergedHeight = merged.bbox.r2 - merged.bbox.r1 + 1;
				//меняем nStartCol, потому что приходит колонка той ячейки, на которой начали выделение
				nStartRowCol = merged.bbox.c1;
			}

		}

		this.worksheet.workbook.dependencyFormulas.lockRecal();

		var oRes = null;
		var oThis = this;
		var bAscent = false;
		if (nOption == Asc.c_oAscSortOptions.Ascending) {
			bAscent = true;
		}

		//get sort elems
		//_getSortElems - for split big function
		var elemObj = this._getSortElems(bbox, nStartRowCol, nOption, sortColor, opt_by_row, opt_custom_sort);
		var aSortElems = elemObj.aSortElems;
		var nColFirst0 = elemObj.nColFirst0;
		var nRowFirst0 = elemObj.nRowFirst0;
		var nLastRow0 = elemObj.nLastRow0;
		var nLastCol0 = elemObj.nLastCol0;


		//проверяем что это не пустая операция
		var aSortData = [];
		var nHiddenCount = 0;
		var oFromArray = {};

		var nNewIndex, oNewElem, i, length;
		var nColMax = 0, nRowMax = 0;
		var nColMin = gc_nMaxCol0, nRowMin = gc_nMaxRow0;
		var nToMax = 0;
		for (i = 0, length = aSortElems.length; i < length; ++i) {
			var item = aSortElems[i];
			if (opt_by_row) {
				nNewIndex = i * nMergedWidth + nColFirst0 + nHiddenCount;
				while (false != oThis.worksheet.getColHidden(nNewIndex)) {
					nHiddenCount++;
					nNewIndex = i * nMergedWidth + nColFirst0 + nHiddenCount;
				}
				oNewElem = new UndoRedoData_FromToRowCol(false, item.col, nNewIndex);
				oFromArray[item.col] = 1;
				if (nColMax < item.col) {
					nColMax = item.col;
				}
				if (nColMax < nNewIndex) {
					nColMax = nNewIndex;
				}
				if (nColMin > item.col) {
					nColMin = item.col;
				}
				if (nColMin > nNewIndex) {
					nColMin = nNewIndex;
				}
			} else {
				nNewIndex = i * nMergedHeight + nRowFirst0 + nHiddenCount;
				while (false != oThis.worksheet.getRowHidden(nNewIndex)) {
					nHiddenCount++;
					nNewIndex = i * nMergedHeight + nRowFirst0 + nHiddenCount;
				}
				oNewElem = new UndoRedoData_FromToRowCol(true, item.row, nNewIndex);
				oFromArray[item.row] = 1;
				if (nRowMax < item.row) {
					nRowMax = item.row;
				}
				if (nRowMax < nNewIndex) {
					nRowMax = nNewIndex;
				}
				if (nRowMin > item.row) {
					nRowMin = item.row;
				}
				if (nRowMin > nNewIndex) {
					nRowMin = nNewIndex;
				}
			}
			if (nToMax < nNewIndex) {
				nToMax = nNewIndex;
			}
			if (oNewElem.from != oNewElem.to) {
				aSortData.push(oNewElem);
			}
		}


		if (aSortData.length > 0) {
			//добавляем индексы перехода пустых ячеек(нужно для сортировки комментариев)
			var nRowColMin = opt_by_row ? nColMin : nRowMin;
			var nRowColMax = opt_by_row ? nColMax : nRowMax;
			var hiddenFunc = opt_by_row ? oThis.worksheet.getColHidden : oThis.worksheet.getRowHidden;
			for (i = nRowColMin; i <= nRowColMax; ++i) {
				if (null == oFromArray[i] && false == hiddenFunc.apply(oThis.worksheet, [i])) {
					var nFrom = i;
					var nTo = ++nToMax;
					while (false != hiddenFunc.apply(oThis.worksheet, [nTo]))
						nTo = ++nToMax;
					if (nFrom != nTo) {
						oNewElem = new UndoRedoData_FromToRowCol(false, nFrom, nTo);
						aSortData.push(oNewElem);
					}
				}
			}

			AscCommon.History.Create_NewPoint();
			var oSelection = AscCommon.History.GetSelection();
			if (null != oSelection) {
				oSelection = oSelection.clone();
				oSelection.assign(nColFirst0, nRowFirst0, nLastCol0, nLastRow0);
				AscCommon.History.SetSelection(oSelection);
				AscCommon.History.SetSelectionRedo(oSelection);
			}
			var oUndoRedoBBox = new UndoRedoData_BBox({r1: nRowFirst0, c1: nColFirst0, r2: nLastRow0, c2: nLastCol0});
			oRes = new AscCommonExcel.UndoRedoData_SortData(oUndoRedoBBox, aSortData, opt_by_row);
			this._sortByArray(oUndoRedoBBox, aSortData, null, opt_by_row);

			var range = opt_by_row ? new Asc.Range(nColFirst0, 0, nLastCol0, gc_nMaxRow0) : new Asc.Range(0, nRowFirst0, gc_nMaxCol0, nLastRow0);
			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_Sort, this.worksheet.getId(), range, oRes)
		}
		this.worksheet.workbook.dependencyFormulas.unlockRecal();
		return oRes;
	};
	Range.prototype._getSortElems = function (bbox, nStartRowCol, nOption, sortColor, opt_by_row, opt_custom_sort) {
		var oThis = this;

		var bAscent = false;
		if (nOption == Asc.c_oAscSortOptions.Ascending) {
			bAscent = true;
		}

		var colorFill = nOption === Asc.c_oAscSortOptions.ByColorFill;
		var colorText = nOption === Asc.c_oAscSortOptions.ByColorFont;
		var isSortColor = colorFill || colorText;

		var nRowFirst0 = bbox.r1;
		var nRowLast0 = bbox.r2;
		var nColFirst0 = bbox.c1;
		var nColLast0 = bbox.c2;

		var bWholeCol = false;
		var bWholeRow = false;
		if (0 == nRowFirst0 && gc_nMaxRow0 == nRowLast0) {
			bWholeCol = true;
		}
		if (0 == nColFirst0 && gc_nMaxCol0 == nColLast0) {
			bWholeRow = true;
		}

		var sortConditions, caseSensitive;
		if(opt_custom_sort && opt_custom_sort.SortConditions) {
			sortConditions = opt_custom_sort.SortConditions;
			nStartRowCol = opt_by_row ? sortConditions[0].Ref.r1 : sortConditions[0].Ref.c1;
			bAscent = !sortConditions[0].ConditionDescending;
		}
		if(opt_custom_sort) {
			//caseSensitive = opt_custom_sort.CaseSensitive;
			//пока игнорируем данный флаг, поскольку сравнения строк в excel при сортировке работает иначе(например - "Green" > "green")
			//возможно, стоит воспользоваться функцией localeCompare - но для этого необходимо проверить грамотное ли сравнение будет
		}

		var nLastRow0, nLastCol0;
		if (opt_by_row) {
			var oRangeRow = this.worksheet.getRange(new CellAddress(nStartRowCol, nColFirst0, 0), new CellAddress(nStartRowCol, nColLast0, 0));
			nLastRow0 = nRowLast0;
			nLastCol0 = 0;
			if (true == bWholeCol) {
				nLastRow0 = 0;
				this._foreachColNoEmpty(null, function (cell) {
					var nCurRow0 = cell.nRow;
					if (nCurRow0 > nLastRow0) {
						nLastRow0 = nCurRow0;
					}
				});
			}
		} else {
			var oRangeCol = this.worksheet.getRange(new CellAddress(nRowFirst0, nStartRowCol, 0), new CellAddress(nRowLast0, nStartRowCol, 0));
			nLastRow0 = 0;
			nLastCol0 = nColLast0;
			if (true == bWholeRow) {
				nLastCol0 = 0;
				this._foreachRowNoEmpty(null, function (cell) {
					var nCurCol0 = cell.nCol;
					if (nCurCol0 > nLastCol0) {
						nLastCol0 = nCurCol0;
					}
				});
			}
		}

		var aMerged = this.worksheet.mergeManager.get(this.bbox);
		var checkMerged = function(_cell) {
			var res = null;

			if(aMerged && aMerged.inner && aMerged.inner.length > 0) {
				for(var i = 0; i < aMerged.inner.length; i++) {
					if(aMerged.inner[i].bbox.contains(_cell.nCol, _cell.nRow)) {
						if(aMerged.inner[i].bbox.r1 === _cell.nRow && aMerged.inner[i].bbox.c1 === _cell.nCol) {
							return true;
						} else {
							return false;
						}
					}
				}
			}

			return res;
		};

		//собираем массив обьектов для сортировки
		var aSortElems = [];
		var putElem = false;
		var fAddSortElems = function (oCell, nRow0, nCol0) {
			//не сортируем сткрытие строки
			if ((opt_by_row && !oThis.worksheet.getColHidden(nCol0)) || (!opt_by_row && !oThis.worksheet.getRowHidden(nRow0))) {
				if (!opt_by_row && nLastRow0 < nRow0) {
					nLastRow0 = nRow0;
				}
				if (opt_by_row && nLastCol0 < nCol0) {
					nLastCol0 = nCol0;
				}
				var val = oCell.getValueWithoutFormat();

				if(opt_custom_sort && false === checkMerged(oCell)) {
					return;
				}

				//for sort color
				var colorFillCell, colorsTextCell = null;
				if (colorFill || opt_custom_sort) {
					var styleCell = oCell.getCompiledStyleCustom(false, true, true);
					colorFillCell = styleCell !== null && styleCell.fill ? styleCell.fill.bg() : null;
				}
				if (colorText || opt_custom_sort) {
					var value2 = oCell.getValue2();
					for (var n = 0; n < value2.length; n++) {
						if (null === colorsTextCell) {
							colorsTextCell = [];
						}

						colorsTextCell.push(value2[n].format.getColor());
					}
				}

				var nNumber = null;
				var sText = null;
				var res;
				if ("" != val) {
					var nVal = val - 0;
					if (nVal == val) {
						nNumber = nVal;
					} else {
						sText = val;
					}
					if (opt_by_row) {
						res = {col: nCol0, num: nNumber, text: sText, colorFill: colorFillCell, colorsText: colorsTextCell};
					} else {
						res = {row: nRow0, num: nNumber, text: sText, colorFill: colorFillCell, colorsText: colorsTextCell};
					}
				} else if (isSortColor || (opt_custom_sort && (colorFillCell || colorsTextCell))) {
					if (opt_by_row) {
						res = {col: nCol0, num: nNumber, text: sText, colorFill: colorFillCell, colorsText: colorsTextCell};
					} else {
						res = {row: nRow0, num: nNumber, text: sText, colorFill: colorFillCell, colorsText: colorsTextCell};
					}
				}

				if(!putElem) {
					return res;
				} else if(res) {
					aSortElems.push(res);
				}
			}
		};

		putElem = true;
		if (opt_by_row) {
			if (nRowFirst0 == nStartRowCol) {
				while (0 == aSortElems.length && nStartRowCol <= nLastRow0) {
					if (false == bWholeRow) {
						oRangeRow._foreachNoEmptyByCol(fAddSortElems);
					} else {
						oRangeRow._foreachRowNoEmpty(null, fAddSortElems);
					}
					if (0 == aSortElems.length) {
						nStartRowCol++;
						oRangeRow = this.worksheet.getRange(new CellAddress(nStartRowCol, nColFirst0, 0), new CellAddress(nStartRowCol, nColLast0, 0));
					}
				}
			} else {
				if (false == bWholeRow) {
					oRangeRow._foreachNoEmptyByCol(fAddSortElems);
				} else {
					oRangeRow._foreachRowNoEmpty(null, fAddSortElems);
				}
			}
		} else {
			if (nColFirst0 == nStartRowCol) {
				while (0 == aSortElems.length && nStartRowCol <= nLastCol0) {
					if (false == bWholeCol) {
						oRangeCol._foreachNoEmpty(fAddSortElems);
					} else {
						oRangeCol._foreachColNoEmpty(null, fAddSortElems);
					}
					if (0 == aSortElems.length) {
						nStartRowCol++;
						oRangeCol = this.worksheet.getRange(new CellAddress(nRowFirst0, nStartRowCol, 0), new CellAddress(nRowLast0, nStartRowCol, 0));
					}
				}
			} else {
				if (false == bWholeCol) {
					oRangeCol._foreachNoEmpty(fAddSortElems);
				} else {
					oRangeCol._foreachColNoEmpty(null, fAddSortElems);
				}
			}
		}

		function strcmp(str1, str2) {
			if (str1 === str2) {
				return 0;
			}
			return str1.localeCompare(str2);
		}


		//color sort
		var colorFillCmp = function (color1, color2, _customCellColor) {
			var res = false;
			//TODO возможно так сравнивать не правильно, позже пересмотреть
			if (colorFill || _customCellColor === true) {
				res = (color1 !== null && color2 !== null && color1.rgb === color2.rgb) || (color1 === color2);
			} else if ((colorText || _customCellColor === false) && color1 && color1.length) {
				for (var n = 0; n < color1.length; n++) {
					if (color1[n] && color2 !== null && color1[n].rgb === color2.rgb) {
						res = true;
						break;
					}
				}
			}

			return res;
		};

		var getSortElem = function(row, col) {
			var oCell;
			oThis.worksheet._getCellNoEmpty(row, col, function(cell) {
				oCell = cell;
			});
			return oCell ? fAddSortElems(oCell, row, col) : null;
		};

		putElem = false;
		if (isSortColor) {
			var newArrayNeedColor = [];
			var newArrayAnotherColor = [];

			for (var i = 0; i < aSortElems.length; i++) {
				var color = colorFill ? aSortElems[i].colorFill : aSortElems[i].colorsText;
				if (colorFillCmp(color, sortColor)) {
					newArrayNeedColor.push(aSortElems[i]);
				} else {
					newArrayAnotherColor.push(aSortElems[i]);
				}
			}

			aSortElems = newArrayNeedColor.concat(newArrayAnotherColor);
		} else {
			aSortElems.sort(function (a, b) {
				var res = 0;
				var nullVal = false;
				var compare = function(_a, _b, _sortCondition) {
					//если есть opt_custom_sort(->sortConditions) - тогда может быть несколько условий сортировки
					//в данном случае идём по отдельной ветке и по-другому обрабатываем сортировку по цвету
					//TODO стоит рассмотреть вариант одной обработки для разных вариантов сортировки
					if(_sortCondition && (_sortCondition.ConditionSortBy === Asc.ESortBy.sortbyCellColor || _sortCondition.ConditionSortBy === Asc.ESortBy.sortbyFontColor)) {
						if(!_a || !_b) {
							return res;
						}
						var _isCellColor = _sortCondition.ConditionSortBy === Asc.ESortBy.sortbyCellColor;
						var _color1 = _isCellColor ? _a.colorFill : _a.colorsText;
						var _color2 = _isCellColor ? _b.colorFill : _b.colorsText;
						var _sortColor = _isCellColor ? _sortCondition.dxf.fill.bg() : _sortCondition.dxf.font.getColor();
						var cmp1 = colorFillCmp(_color1, _sortColor, _isCellColor);
						var cmp2 = colorFillCmp(_color2, _sortColor, _isCellColor);

						if(cmp1 === cmp2) {
							res = 0;
						} else if(cmp1 && !cmp2) {
							res = -1;
						} else if(!cmp1 && cmp2) {
							res = 1;
						}
					} else {
						if (_a && null != _a.text) {
							if (_b && null != _b.text) {
								var val1 = caseSensitive ? _a.text : _a.text.toUpperCase();
								var val2 = caseSensitive ? _b.text : _b.text.toUpperCase();
								res = strcmp(val1, val2);
							} else if(_b && null != _b.num) {
								res = 1;
							} else {
								res = -1;
								nullVal = true;
							}
						} else if (_a && null != _a.num) {
							if (_b && null != _b.num) {
								res = _a.num - _b.num;
							} else if(_b && null != _b.text) {
								res = -1;
							} else {
								res = -1;
								nullVal = true;
							}
						} else if(_b && (null != _b.num || null != _b.text)){
							res = 1;
							nullVal = true;
						}
					}
				};

				compare(a, b, sortConditions ? sortConditions[0] : null);
				if (0 == res) {
					if(sortConditions) {
						for(var i = 1; i < sortConditions.length; i++) {
							var row = sortConditions[i].Ref.r1;
							var col = sortConditions[i].Ref.c1;
							var row1 = opt_by_row ? row : a.row;
							var col1 = !opt_by_row ? col : a.col;
							var row2 = opt_by_row ? row : b.row;
							var col2 = !opt_by_row ? col : b.col;
							var tempA = getSortElem(row1, col1);
							var tempB = getSortElem(row2, col2);
							compare(tempA, tempB, sortConditions[i]);
							var tempAscent = !sortConditions[i].ConditionDescending;
							if(res != 0) {
								if(!tempAscent) {
									res = -res;
								}
								break;
							} else if(i === sortConditions.length - 1 && tempA && tempB) {
								res = opt_by_row ? tempA.col - tempB.col : tempA.row - tempB.row;
							}
						}
					} else {
						res = opt_by_row ? a.col - b.col : a.row - b.row;
					}
				} else if (!bAscent && !nullVal) {
					res = -res;
				}
				return res;
			});
		}

		return {aSortElems: aSortElems, nRowFirst0: nRowFirst0, nColFirst0: nColFirst0, nLastRow0: nLastRow0, nLastCol0: nLastCol0};
	};
	Range.prototype._sortByArray = function (oBBox, aSortData, bUndo, opt_by_row) {
		var t = this;
		var height = oBBox.r2 - oBBox.r1 + 1;
		var oSortedIndexes = {};
		var nFrom, nTo, length, i;
		for (i = 0, length = aSortData.length; i < length; ++i) {
			var item = aSortData[i];
			nFrom = item.from;
			nTo = item.to;
			if (true == this.worksheet.workbook.bUndoChanges) {
				nFrom = item.to;
				nTo = item.from;
			}
			oSortedIndexes[nFrom] = nTo;
		}
		//сортируются только одинарные гиперссылки, все неодинарные оставляем
		var aSortedHyperlinks = [], hyp;
		if (false == this.worksheet.workbook.bUndoChanges && (false == this.worksheet.workbook.bRedoChanges || this.worksheet.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			var aHyperlinks = this.worksheet.hyperlinkManager.get(this.bbox);
			for (i = 0, length = aHyperlinks.inner.length; i < length; i++) {
				var elem = aHyperlinks.inner[i];
				hyp = elem.data;
				if (hyp.Ref.isOneCell()) {
					nFrom = opt_by_row ? elem.bbox.c1 : elem.bbox.r1;
					nTo = oSortedIndexes[nFrom];
					if (null != nTo) {
						//удаляем ссылки, а не перемещаем, чтобы не было конфликтов(например в случае если все ячейки имеют ссылки
						// и их надо передвинуть)
						this.worksheet.hyperlinkManager.removeElement(elem);
						var oNewHyp = hyp.clone();
						if(opt_by_row) {
							oNewHyp.Ref.setOffset(new AscCommon.CellBase(0, nTo - nFrom));
						} else {
							oNewHyp.Ref.setOffset(new AscCommon.CellBase(nTo - nFrom, 0));
						}
						aSortedHyperlinks.push(oNewHyp);
					}
				}
			}
			AscCommon.History.LocalChange = false;
		}

		var tempRange = this.worksheet.getRange3(oBBox.r1, oBBox.c1, oBBox.r2, oBBox.c2);
		var func = opt_by_row ? tempRange._foreachNoEmptyByCol : tempRange._foreachNoEmpty;
		func.apply(tempRange, [(function (cell) {
			var ws = t.worksheet;
			var formula = cell.getFormulaParsed();
			if (formula) {
				var cellWithFormula = formula.getParent();
				var nFrom = opt_by_row ? cell.nCol : cell.nRow;
				var nTo = oSortedIndexes[nFrom];
				if (null != nTo) {
					if (!cell.formulaParsed.isParsed) {
						cell.formulaParsed.parse();
					}
					if (opt_by_row) {
						cell.changeOffset(new AscCommon.CellBase(0, nTo - nFrom), true, true, true);
					} else {
						cell.changeOffset(new AscCommon.CellBase(nTo - nFrom, 0), true, true, true);
					}
					formula = cell.getFormulaParsed();
					cellWithFormula = formula.getParent();
					if (opt_by_row) {
						cellWithFormula.nCol = nTo;
					} else {
						cellWithFormula.nRow = nTo;
					}
				}
				ws.workbook.dependencyFormulas.addToChangedCell(cellWithFormula);
			}
		})]);


		var tempSheetMemory, nIndexFrom, nIndexTo, j;
		if (opt_by_row) {
			tempSheetMemory = [];
			for (j in oSortedIndexes) {
				nIndexFrom = j - 0;
				nIndexTo = oSortedIndexes[j];
				var sheetMemoryFrom = this.worksheet.getColData(nIndexFrom);
				var sheetMemoryTo = this.worksheet.getColData(nIndexTo);

				tempSheetMemory[nIndexTo] = new AscCommonExcel.SheetMemory(g_nCellStructSize, height);
				tempSheetMemory[nIndexTo].copyRange(sheetMemoryTo, oBBox.r1, 0, height);

				if (tempSheetMemory[nIndexFrom]) {
					sheetMemoryTo.copyRange(tempSheetMemory[nIndexFrom], 0, oBBox.r1, height);
				} else {
					sheetMemoryTo.copyRange(sheetMemoryFrom, oBBox.r1, oBBox.r1, height);
				}
			}
		} else {
			tempSheetMemory = new AscCommonExcel.SheetMemory(g_nCellStructSize, height);
			for (i = oBBox.c1; i <= oBBox.c2; ++i) {
				var sheetMemory = this.worksheet.getColDataNoEmpty(i);
				if (sheetMemory) {
					tempSheetMemory.copyRange(sheetMemory, oBBox.r1, 0, height);
					for (j in oSortedIndexes) {
						nIndexFrom = j - 0;
						nIndexTo = oSortedIndexes[j];
						tempSheetMemory.copyRange(sheetMemory, nIndexFrom, nIndexTo - oBBox.r1, 1);
					}
					sheetMemory.copyRange(tempSheetMemory, 0, oBBox.r1, height);
				}
			}
		}

		this.worksheet.workbook.dependencyFormulas.addToChangedRange(this.worksheet.getId(), new Asc.Range(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r2));
		this.worksheet.workbook.dependencyFormulas.calcTree();
		if (this.worksheet.workbook.handlers) {
			this.worksheet.workbook.handlers.trigger("changeDocument", AscCommonExcel.docChangedType.sheetContent, this.worksheet, null, this.worksheet.getId());
		}

		if (false == this.worksheet.workbook.bUndoChanges && (false == this.worksheet.workbook.bRedoChanges || this.worksheet.workbook.bCollaborativeChanges)) {
			AscCommon.History.LocalChange = true;
			//восстанавливаем удаленые гиперссылки
			if (aSortedHyperlinks.length > 0) {
				for (i = 0, length = aSortedHyperlinks.length; i < length; i++) {
					hyp = aSortedHyperlinks[i];
					this.worksheet.hyperlinkManager.add(hyp.Ref.getBBox0(), hyp);
				}
			}
			AscCommon.History.LocalChange = false;
		}
	};
	Range.prototype.fillData=function(data){
		for (var i = 0; i < data.length; ++i) {
			var row = data[i];
			for (var j = 0; j < row.length; ++j) {
				this.setOffset(new AscCommon.CellBase(i, j));
				var val = row[j];
				if ("string" === typeof val) {
					this.setValue(val);
				} else {
					if (val.value) {
						this.setValueData(new AscCommonExcel.UndoRedoData_CellValueData(null, val.value));
					}
					if (val.format) {
						this.setNumFormat(val.format);
					}
				}
				this.setOffset(new AscCommon.CellBase(-i, -j));
			}
		}
	};


	function _isSameSizeMerged(bbox, aMerged, checkProportion) {
		var oRes = null;
		var nWidth = null;
		var nHeight = null;
		for (var i = 0, length = aMerged.length; i < length; i++) {
			var mergedBBox = aMerged[i].bbox;
			var nCurWidth = mergedBBox.c2 - mergedBBox.c1 + 1;
			var nCurHeight = mergedBBox.r2 - mergedBBox.r1 + 1;
			if (null == nWidth || null == nHeight) {
				nWidth = nCurWidth;
				nHeight = nCurHeight;
			} else if (nCurWidth != nWidth || nCurHeight != nHeight) {
				nWidth = null;
				nHeight = null;
				break;
			}
		}
		if (null != nWidth && null != nHeight) {
			var getRowColArr = function (byCol) {
				var _aRowColTest = byCol ? new Array(nBBoxWidth) : new Array(nBBoxHeight);
				for (var i = 0, length = aMerged.length; i < length; i++) {
					var merged = aMerged[i];
					var j;
					if (byCol) {
						for (j = merged.bbox.c1; j <= merged.bbox.c2; j++) {
							_aRowColTest[j - bbox.c1] = 1;
						}
					} else {
						for (j = merged.bbox.r1; j <= merged.bbox.r2; j++) {
							_aRowColTest[j - bbox.r1] = 1;
						}
					}

				}
				return _aRowColTest;
			};
			var checkArr = function (_aRowColTest) {
				var _res = null;
				var bExistNull = false;
				for (var i = 0, length = _aRowColTest.length; i < length; i++) {
					if (null == _aRowColTest[i]) {
						bExistNull = true;
						break;
					}
				}
				if (!bExistNull) {
					_res = true;
				}

				return _res;
			};

			//проверяем что merge ячеки полностью заполняют область
			var nBBoxWidth = bbox.c2 - bbox.c1 + 1;
			var nBBoxHeight = bbox.r2 - bbox.r1 + 1;
			if (checkProportion && nBBoxWidth % nWidth === 0 && nBBoxHeight % nHeight === 0) {
				aRowColTest = getRowColArr();
				bRes = checkArr(aRowColTest);
				if (bRes) {
					aRowColTest = getRowColArr(true);
					bRes = checkArr(aRowColTest);
				}
				if (bRes) {
					oRes = {width: nWidth, height: nHeight};
				}
			} else if (nBBoxWidth == nWidth || nBBoxHeight == nHeight) {
				var bRes = false;
				var aRowColTest = null;
				if (nBBoxWidth == nWidth && nBBoxHeight == nHeight) {
					bRes = true;
				} else if (nBBoxWidth == nWidth) {
					aRowColTest = getRowColArr();
				} else if (nBBoxHeight == nHeight) {
					aRowColTest = getRowColArr(true);
				}
				if (null != aRowColTest) {
					bRes = checkArr(aRowColTest);
				}
				if (bRes) {
					oRes = {width: nWidth, height: nHeight};
				}
			}
		}
		return oRes;
	}
	function _canPromote(from, wsFrom, to, wsTo, bIsPromote, nWidth, nHeight, bVertical, nIndex) {
		var oRes = {oMergedFrom: null, oMergedTo: null, to: to};
		//если надо только удалить внутреннее содержимое не смотрим на замерженость
		if(!bIsPromote || !((true == bVertical && nIndex >= 0 && nIndex < nHeight) || (false == bVertical && nIndex >= 0 && nIndex < nWidth)))
		{
			if(null != to){
				var oMergedTo = wsTo.mergeManager.get(to);
				if(oMergedTo.outer.length > 0)
					oRes = null;
				else
				{
					var oMergedFrom = wsFrom.mergeManager.get(from);
					oRes.oMergedFrom = oMergedFrom;
					if(oMergedTo.inner.length > 0)
					{
						oRes.oMergedTo = oMergedTo;
						if (bIsPromote) {
							if (oMergedFrom.inner.length > 0) {
								//merge области должны иметь одинаковый размер
								var oSizeFrom = _isSameSizeMerged(from, oMergedFrom.inner);
								var oSizeTo = _isSameSizeMerged(to, oMergedTo.inner);
								if (!(null != oSizeFrom && null != oSizeTo && oSizeTo.width == oSizeFrom.width && oSizeTo.height == oSizeFrom.height))
									oRes = null;
							}
							else
								oRes = null;
						}
					}
				}
			}
		}
		return oRes;
	}
// Подготовка Copy Style
	function preparePromoteFromTo(from, to) {
		var bSuccess = true;
		if (to.isOneCell())
			to.setOffsetLast(new AscCommon.CellBase((from.r2 - from.r1) - (to.r2 - to.r1), (from.c2 - from.c1) - (to.c2 - to.c1)));

		if(!from.isIntersect(to)) {
			var bFromWholeCol = (0 == from.c1 && gc_nMaxCol0 == from.c2);
			var bFromWholeRow = (0 == from.r1 && gc_nMaxRow0 == from.r2);
			var bToWholeCol = (0 == to.c1 && gc_nMaxCol0 == to.c2);
			var bToWholeRow = (0 == to.r1 && gc_nMaxRow0 == to.r2);
			bSuccess = (bFromWholeCol === bToWholeCol && bFromWholeRow === bToWholeRow);
		} else
			bSuccess = false;
		return bSuccess;
	}
// Перед promoteFromTo обязательно должна быть вызывана функция preparePromoteFromTo
	function promoteFromTo(from, wsFrom, to, wsTo) {
		var bVertical = true;
		var nIndex = 1;
		//проверяем можно ли осуществить promote
		var oCanPromote = _canPromote(from, wsFrom, to, wsTo, false, 1, 1, bVertical, nIndex);
		if(null != oCanPromote)
		{
			AscCommon.History.Create_NewPoint();
			var oSelection = AscCommon.History.GetSelection();
			if(null != oSelection)
			{
				oSelection = oSelection.clone();
				oSelection.assign(from.c1, from.r1, from.c2, from.r2);
				AscCommon.History.SetSelection(oSelection);
			}
			var oSelectionRedo = AscCommon.History.GetSelectionRedo();
			if(null != oSelectionRedo)
			{
				oSelectionRedo = oSelectionRedo.clone();
				oSelectionRedo.assign(to.c1, to.r1, to.c2, to.r2);
				AscCommon.History.SetSelectionRedo(oSelectionRedo);
			}
			//удаляем merge ячейки в to(после _canPromote должны остаться только inner)
			wsTo.mergeManager.remove(to, true);
			_promoteFromTo(from, wsFrom, to, wsTo, false, oCanPromote, false, bVertical, nIndex);
		}
	}
	Range.prototype.canPromote=function(bCtrl, bVertical, nIndex){
		var oBBox = this.bbox;
		var nWidth = oBBox.c2 - oBBox.c1 + 1;
		var nHeight = oBBox.r2 - oBBox.r1 + 1;
		var bWholeCol = false;	var bWholeRow = false;
		if(0 == oBBox.r1 && gc_nMaxRow0 == oBBox.r2)
			bWholeCol = true;
		if(0 == oBBox.c1 && gc_nMaxCol0 == oBBox.c2)
			bWholeRow = true;
		if((bWholeCol && bWholeRow) || (true == bVertical && bWholeCol) || (false == bVertical && bWholeRow))
			return null;
		var oPromoteAscRange = null;
		if(0 == nIndex)
			oPromoteAscRange = new Asc.Range(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r2);
		else
		{
			if(bVertical)
			{
				if(nIndex > 0)
				{
					if(nIndex >= nHeight)
						oPromoteAscRange = new Asc.Range(oBBox.c1, oBBox.r2 + 1, oBBox.c2, oBBox.r1 + nIndex);
					else
						oPromoteAscRange = new Asc.Range(oBBox.c1, oBBox.r1 + nIndex, oBBox.c2, oBBox.r2);
				}
				else
					oPromoteAscRange = new Asc.Range(oBBox.c1, oBBox.r1 + nIndex, oBBox.c2, oBBox.r1 - 1);
			}
			else
			{
				if(nIndex > 0)
				{
					if(nIndex >= nWidth)
						oPromoteAscRange = new Asc.Range(oBBox.c2 + 1, oBBox.r1, oBBox.c1 + nIndex, oBBox.r2);
					else
						oPromoteAscRange = new Asc.Range(oBBox.c1 + nIndex, oBBox.r1, oBBox.c2, oBBox.r2);
				}
				else
					oPromoteAscRange = new Asc.Range(oBBox.c1 + nIndex, oBBox.r1, oBBox.c1 - 1, oBBox.r2);
			}
		}
		//проверяем можно ли осуществить promote
		return _canPromote(oBBox, this.worksheet, oPromoteAscRange, this.worksheet, true, nWidth, nHeight, bVertical, nIndex);
	};
	Range.prototype.promote=function(bCtrl, bVertical, nIndex, oCanPromote){
		//todo отдельный метод для promote в таблицах и merge в таблицах
		if (!oCanPromote) {
			oCanPromote = this.canPromote(bCtrl, bVertical, nIndex);
		}
		var oBBox = this.bbox;
		var nWidth = oBBox.c2 - oBBox.c1 + 1;
		var nHeight = oBBox.r2 - oBBox.r1 + 1;

		AscCommon.History.Create_NewPoint();
		var oSelection = AscCommon.History.GetSelection();
		if(null != oSelection)
		{
			oSelection = oSelection.clone();
			oSelection.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r2);
			AscCommon.History.SetSelection(oSelection);
		}
		var oSelectionRedo = AscCommon.History.GetSelectionRedo();
		if(null != oSelectionRedo)
		{
			oSelectionRedo = oSelectionRedo.clone();
			if(0 == nIndex)
				oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r2);
			else
			{
				if(bVertical)
				{
					if(nIndex > 0)
					{
						if(nIndex >= nHeight)
							oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r1 + nIndex);
						else
							oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r1 + nIndex - 1);
					}
					else
						oSelectionRedo.assign(oBBox.c1, oBBox.r1 + nIndex, oBBox.c2, oBBox.r2);
				}
				else
				{
					if(nIndex > 0)
					{
						if(nIndex >= nWidth)
							oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c1 + nIndex, oBBox.r2);
						else
							oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c1 + nIndex - 1, oBBox.r2);
					}
					else
						oSelectionRedo.assign(oBBox.c1 + nIndex, oBBox.r1, oBBox.c2, oBBox.r2);
				}
			}
			AscCommon.History.SetSelectionRedo(oSelectionRedo);
		}
		_promoteFromTo(oBBox, this.worksheet, oCanPromote.to, this.worksheet, true, oCanPromote, bCtrl, bVertical, nIndex);
	};
	function _addAInputTimePeriod(aTimePeriodName, bNeedSwapFirstLastElem) {
		let aTimePeriod = aTimePeriodName.map(function (sTimePeriod) {
			return sTimePeriod.toLowerCase();
		});
		const nLastIndex = aTimePeriod.length - 1;

		if(!aTimePeriod[nLastIndex]) {
			aTimePeriod.pop();
		}

		if (bNeedSwapFirstLastElem) {
			let sLastElem = aTimePeriod.pop();
			aTimePeriod.unshift(sLastElem);
		}

		return aTimePeriod;
	}
	function _updateATimePeriod (aTimePeriods, sValue) {
		let sFirstSymbols = sValue.slice(0,2);

		return aTimePeriods.map(function (sTimePeriod) {
			if (sFirstSymbols === sFirstSymbols.toUpperCase()) {
					return sTimePeriod.toUpperCase();
			}
			if (sFirstSymbols === sFirstSymbols.toLowerCase()) {
				// Because source array already has elements in lowercase
				return sTimePeriod
			}
			// For cases like Monday, MoNdAy or mOnDaY. If first symbol is lowercase then values in lowercase else values start with capitalized
			if (sFirstSymbols[0] === sFirstSymbols[0].toUpperCase()) {
					return sTimePeriod[0].toUpperCase() + sTimePeriod.slice(1);
			}

			return sTimePeriod;
		});
	}
	function _getIndexATimePeriods(aTimePeriods, sValue) {
		let sFirstSymbols = sValue.slice(0,2);
		let sFormatedValue = '';
		let sSlicedValue = sValue.slice(1);

		if (sValue === sValue.toLowerCase() || sValue === sValue.toUpperCase()) {
			return  aTimePeriods.indexOf(sValue);
		}

		// For cases like Monday, MoNdAy, MOnDaY or mOnDaY. If first symbol is lowercase then sValue convert in lowercase else sValue convert to start with capitalized
		if (sFirstSymbols[0] === sFirstSymbols[0].toLowerCase()) {
			sFormatedValue = sValue.toLowerCase();
			return aTimePeriods.indexOf(sFormatedValue);
		}
		if (sFirstSymbols === sFirstSymbols.toUpperCase()) {
			sFormatedValue = sValue.toUpperCase();
			return aTimePeriods.indexOf(sFormatedValue);
		}
		sFormatedValue = sSlicedValue === sSlicedValue.toLowerCase() ? sValue : sValue[0] + sSlicedValue.toLowerCase();
		return aTimePeriods.indexOf(sFormatedValue);
	}
	function _getRepeatTimePeriod(nRepeat, nPreviousVal, nIndex, nMaxTimePeriod) {
		let bIsSequence = false;
		let nDiff = 0;
		let nCurrentVal = nIndex;
		let nLastElement = nMaxTimePeriod - 1;

		if (nPreviousVal) {
			// Defining is datas has asc sequence like Monday, Tuesday etc or even\odd sequence?
			if (nIndex === 0 && nPreviousVal === nLastElement) {
				bIsSequence = true;
			} else {
				if (nRepeat > 0) nCurrentVal += nMaxTimePeriod * nRepeat;
				nDiff = nCurrentVal - nPreviousVal;
				if (nDiff === 1) {
					bIsSequence = true;
				}
			}
			if (bIsSequence) {
				return nIndex === 0 ? nRepeat + 1 : nRepeat;
			} else if (nIndex === 0 || (nIndex === 1 && nPreviousVal >= nLastElement)) {
				return nRepeat + 1;
			}
		}

		return nRepeat;
	}
	function _findIndexAInputTimePeriod(aTimePeriodsList, sValue) {
		return aTimePeriodsList.findIndex(function (aTimePeriod) {
			return aTimePeriod.includes(sValue.toLowerCase());
		});
	}
	function _getAInputTimePeriod(aInputTimePeriodList, nIndex, nPrevIndex, sValue, sNextValue) {
		let aInputTimePeriod = aInputTimePeriodList[nIndex];
		// Checking nearby cells if value has in both arrays (short and full names)
		let sValueLowReg = sValue.toLowerCase();
		let aPrevInputTimePeriod = nPrevIndex ? aInputTimePeriodList[nPrevIndex] : null;

		if (aPrevInputTimePeriod && nIndex !== nPrevIndex && aPrevInputTimePeriod.includes(sValueLowReg)) {
			aInputTimePeriod = aPrevInputTimePeriod;
		} else if (!aPrevInputTimePeriod && sNextValue) { // doesn't check next cell if previous cell exist
			let nNextIndex = _findIndexAInputTimePeriod(aInputTimePeriodList, sNextValue);
			let aNextInputTimePeriod = aInputTimePeriodList[nNextIndex];
			if (aNextInputTimePeriod && nIndex !== nNextIndex && aNextInputTimePeriod.includes(sValueLowReg)) {
				aInputTimePeriod = aNextInputTimePeriod;
			}
		}

		return aInputTimePeriod
	}
	function _getNextValue(wsFrom, oCell, bVertical, from) {
		if (bVertical) {
			let nNextItRow = oCell.nRow + 1;
			if (nNextItRow <= from.r2) {
				return wsFrom.getCell3(nNextItRow, oCell.nCol).getValueWithoutFormat();
			}
		} else {
			let nNextItCol = oCell.nCol + 1;
			if (nNextItCol <= from.c2) {
				return wsFrom.getCell3(oCell.nRow, nNextItCol).getValueWithoutFormat();
			}
		}

		return null;
	}
	function _promoteFromTo(from, wsFrom, to, wsTo, bIsPromote, oCanPromote, bCtrl, bVertical, nIndex) {
		var wb = wsFrom.workbook;
		const oDefaultCultureInfo = AscCommon.g_oDefaultCultureInfo;

		wb.dependencyFormulas.lockRecal();
		AscCommon.History.StartTransaction();

		var oldExcludeHiddenRows = wsFrom.bExcludeHiddenRows;
		if(wsFrom.autoFilters.bIsExcludeHiddenRows(from, wsFrom.selectionRange.activeCell)){
			wsFrom.excludeHiddenRows(true);
		}

		var toRange = wsTo.getRange3(to.r1, to.c1, to.r2, to.c2);
		var fromRange = wsFrom.getRange3(from.r1, from.c1, from.r2, from.c2);
		var bChangeRowColProp = false;
		var nLastCol = from.c2;
		if (0 == from.c1 && gc_nMaxCol0 == from.c2)
		{
			var aRowProperties = [];
			nLastCol = 0;
			fromRange._foreachRowNoEmpty(function(row){
				if(!row.isEmptyProp())
					aRowProperties.push({index: row.index - from.r1, prop: row.getHeightProp(), style: row.getStyle()});
			}, function(cell){
				var nCurCol0 = cell.nCol;
				if(nCurCol0 > nLastCol)
					nLastCol = nCurCol0;
			});
			if(aRowProperties.length > 0)
			{
				bChangeRowColProp = true;
				var nCurCount = 0;
				var nCurIndex = 0;
				while (true) {
					for (var i = 0, length = aRowProperties.length; i < length; ++i) {
						var propElem = aRowProperties[i];
						nCurIndex = to.r1 + nCurCount * (from.r2 - from.r1 + 1) + propElem.index;
						if (nCurIndex > to.r2)
							break;
						else{
							wsTo._getRow(nCurIndex, function(row) {
								if (null != propElem.style)
									row.setStyle(propElem.style);
								if (null != propElem.prop) {
									var oNewProps = propElem.prop;
									var oOldProps = row.getHeightProp();
									if (false === oOldProps.isEqual(oNewProps)) {
										row.setHeightProp(oNewProps);
										AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_RowProp, wsTo.getId(), row._getUpdateRange(), new UndoRedoData_IndexSimpleProp(nCurIndex, true, oOldProps, oNewProps));
									}
								}
							});
						}
					}
					nCurCount++;
					if (nCurIndex > to.r2)
						break;
				}
			}
		}
		var nLastRow = from.r2;
		if (0 == from.r1 && gc_nMaxRow0 == from.r2)
		{
			var aColProperties = [];
			nLastRow = 0;
			fromRange._foreachColNoEmpty(function(col){
				if(!col.isEmpty())
					aColProperties.push({ index: col.index - from.c1, prop: col.getWidthProp(), style: col.getStyle() });
			}, function(cell){
				var nCurRow0 = cell.nRow;
				if(nCurRow0 > nLastRow)
					nLastRow = nCurRow0;
			});
			if (aColProperties.length > 0)
			{
				bChangeRowColProp = true;
				var nCurCount = 0;
				var nCurIndex = 0;
				while (true) {
					for (var i = 0, length = aColProperties.length; i < length; ++i) {
						var propElem = aColProperties[i];
						nCurIndex = to.c1 + nCurCount * (from.c2 - from.c1 + 1) + propElem.index;
						if (nCurIndex > to.c2)
							break;
						else{
							var col = wsTo._getCol(nCurIndex);
							if (null != propElem.style)
								col.setStyle(propElem.style);
							if (null != propElem.prop) {
								var oNewProps = propElem.prop;
								var oOldProps = col.getWidthProp();
								if (false == oOldProps.isEqual(oNewProps)) {
									col.setWidthProp(oNewProps);
									wsTo.initColumn(col);
									AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_ColProp, wsTo.getId(), new Asc.Range(nCurIndex, 0, nCurIndex, gc_nMaxRow0), new UndoRedoData_IndexSimpleProp(nCurIndex, false, oOldProps, oNewProps));
								}
							}
						}
					}
					nCurCount++;
					if (nCurIndex > to.c2)
						break;
				}
			}
		}
		if (bChangeRowColProp) {
			wb.handlers && wb.handlers.trigger("changeWorksheetUpdate", wsTo.getId());
		}
		if(nLastCol != from.c2 || nLastRow != from.r2)
		{
			var offset = new AscCommon.CellBase(nLastRow - from.r2, nLastCol - from.c2);
			toRange.setOffsetLast(offset);
			to = toRange.getBBox0();
			fromRange.setOffsetLast(offset);
			from = fromRange.getBBox0();
		}
		var nWidth = from.c2 - from.c1 + 1;
		var nHeight = from.r2 - from.r1 + 1;
		//удаляем текст или все в области для заполнения
		if(bIsPromote && nIndex >= 0 && ((true == bVertical && nHeight > nIndex) || (false == bVertical && nWidth > nIndex)))
		{
			//удаляем только текст в области для заполнения
			toRange.cleanText();
		}
		else
		{
			//удаляем все в области для заполнения
			if(bIsPromote)
				toRange.cleanAll();
			else
				toRange.cleanFormat();
			//собираем все данные
			var bReverse = false;
			if(nIndex < 0)
				bReverse = true;
			let oPromoteHelper = new PromoteHelper(bVertical, bReverse, from);
			oPromoteHelper.setFillHandleRightClick(wsFrom.getFillHandleRightClick());
			let aInputDaysOfWeek = _addAInputTimePeriod(oDefaultCultureInfo.DayNames, false);
			let aInputShortDaysOfWeek = _addAInputTimePeriod(oDefaultCultureInfo.AbbreviatedDayNames, false);
			let aInputMonths = _addAInputTimePeriod(oDefaultCultureInfo.MonthNames, true);
			let aInputMonthShort = _addAInputTimePeriod(oDefaultCultureInfo.AbbreviatedMonthNames, true);
			let aInputTimePeriodList = [aInputDaysOfWeek, aInputShortDaysOfWeek, aInputMonths, aInputMonthShort];
			let nPreviousVal = null;
			let nPrevInputTimePeriod = null;
			let nRepeat = 0;
			fromRange._foreachNoEmpty(function(oCell, nRow0, nCol0, nRowStart0, nColStart0){
				if(null != oCell)
				{
					function calcTimePeriodValues() {
						// Update array of a time period based on sValue
						aTimePeriods = _updateATimePeriod(aInputTimePeriod, sValue);
						let nMaxTimePeriod = aTimePeriods.length;
						let nIndex = _getIndexATimePeriods(aTimePeriods, sValue);
						nRepeat = _getRepeatTimePeriod(nRepeat, nPreviousVal, nIndex, nMaxTimePeriod);
						// In nVal stores number of a time period.
						// e.g. "index from array days of the week + 7 (count days in week) * count of repeat"
						nVal = nIndex + nMaxTimePeriod * nRepeat;
						bDate = true;
						nPreviousVal = nVal;
						nPrevInputTimePeriod = nInputTimePeriod;
					}
					var nVal = null;
					var bDelimiter = false;
					var sPrefix = null;
					var padding = 0;
					var bDate = false;
					let aTimePeriods = null;
					let aInputTimePeriod = null;
					let sNextValue = null;
					let nInputTimePeriod = null;
					if(bIsPromote)
					{
						if (!oCell.isFormula())
						{
							var sValue = oCell.getValueWithoutFormat();
							if("" != sValue)
							{
								bDelimiter = true;
								var nType = oCell.getType();
								if(CellValueType.Number == nType || CellValueType.String == nType)
								{
									if(CellValueType.Number == nType)
										nVal = sValue - 0;
									else
									{
										//если текст заканчивается на цифру тоже используем ее
										var nEndIndex = sValue.length;
										for(var k = sValue.length - 1; k >= 0; --k)
										{
											var sCurChart = sValue[k];
											if('0' <= sCurChart && sCurChart <= '9')
												nEndIndex--;
											else
												break;
										}
										if(sValue.length != nEndIndex)
										{
											sPrefix = sValue.substring(0, nEndIndex);
											var sNumber = sValue.substring(nEndIndex);
											//sNumber have no decimal point, so can use simple parseInt
											nVal = sNumber - 0;
											padding = sNumber[0] === '0' ? sNumber.length : 0;
										}
										// Value of cell is it a time period?
										sValue = sValue.replace(/[.]$/, '').trim();
										sNextValue = _getNextValue(wsFrom, oCell, bVertical, from);
										nInputTimePeriod = _findIndexAInputTimePeriod(aInputTimePeriodList, sValue);
										aInputTimePeriod = _getAInputTimePeriod(aInputTimePeriodList, nInputTimePeriod, nPrevInputTimePeriod, sValue, sNextValue);
										if (aInputTimePeriod) {
											calcTimePeriodValues();
										}
									}
								}
								if(null != oCell.xfs && null != oCell.xfs.num && null != oCell.xfs.num.getFormat()){
									var numFormat = oNumFormatCache.get(oCell.xfs.num.getFormat());
									if(numFormat.isDateTimeFormat())
										bDate = true;
								}
								if(null != nVal)
									bDelimiter = false;
							}
						}
						else
							bDelimiter = true;
					}
					oPromoteHelper.add(nRow0 - nRowStart0, nCol0 - nColStart0, nVal, bDelimiter, sPrefix, padding, bDate, oCell.duplicate(), aTimePeriods);
				}
			});
			var bCopy = false;
			if(bCtrl)
				bCopy = true;
			//в случае одной ячейки с числом меняется смысл bCtrl
			if(1 == nWidth && 1 == nHeight && oPromoteHelper.isOnlyIntegerSequence()) {
				let fillType = wsTo && wsTo.getActiveFillType();
				if (fillType !== Asc.c_oAscFillType.fillDown && fillType !== Asc.c_oAscFillType.fillUp
				&& fillType !== Asc.c_oAscFillType.fillLeft && fillType !== Asc.c_oAscFillType.fillRight) {
					bCopy = !bCopy;
				}
			}
			oPromoteHelper.setFillMenuChosenProp(wsFrom.getFillMenuChosenProp());
			oPromoteHelper.finishAdd(bCopy);
			//заполняем ячейки данными
			var nStartRow, nEndRow, nStartCol, nEndCol, nColDx, bRowFirst;
			if(bVertical)
			{
				nStartRow = to.c1;
				nEndRow = to.c2;
				bRowFirst = false;
				if(bReverse)
				{
					nStartCol = to.r2;
					nEndCol = to.r1;
					nColDx = -1;
				}
				else
				{
					nStartCol = to.r1;
					nEndCol = to.r2;
					nColDx = 1;
				}
			}
			else
			{
				nStartRow = to.r1;
				nEndRow = to.r2;
				bRowFirst = true;
				if(bReverse)
				{
					nStartCol = to.c2;
					nEndCol = to.c1;
					nColDx = -1;
				}
				else
				{
					nStartCol = to.c1;
					nEndCol = to.c2;
					nColDx = 1;
				}
			}
			var addedFormulasArray = [];
			var isAddFormulaArray = function(arrayRef) {
				var res = false;
				for(var n = 0; n < addedFormulasArray.length; n++) {
					if(addedFormulasArray[n].isEqual(arrayRef)) {
						res = true;
						break;
					}
				}
				return res;
			};
			for(var i = nStartRow; i <= nEndRow; i ++)
			{
				oPromoteHelper.setIndex(i - nStartRow);
				for(var j = nStartCol; (nStartCol - j) * (nEndCol - j) <= 0; j += nColDx)
				{
					if (bVertical && wsTo.bExcludeHiddenRows && wsTo.getRowHidden(j))
					{
						continue;
					}

					var data = oPromoteHelper.getNext();
					if(null != data && (data.getAdditional() || (false == bCopy && null != data.getCurValue())))
					{
						var oFromCell = data.getAdditional();
						var nRow = bRowFirst ? i : j;
						var nCol = bRowFirst ? j : i;
						let sPrefix = data.getPrefix();
						let nCurValue = data.getCurValue();
						let nPadding = data.getPadding();
						let aTimePeriods = data.getTimePeriods();
						let aReverseTimePeriods = [];

						if (aTimePeriods && (bReverse || nCurValue < 0)) {
							aReverseTimePeriods = aTimePeriods.slice();
							aReverseTimePeriods.reverse();
						}
						wsTo._getCell(nRow, nCol, function(oCopyCell){
							if(bIsPromote)
							{
								if(false === bCopy && null != nCurValue)
								{
									var oCellValue = new AscCommonExcel.CCellValue();
									if (null != sPrefix) {
										var sVal = sPrefix;
										//toString enough, because nCurValue nave not decimal part
										var sNumber = nCurValue.toString();
										if (sNumber.length < nPadding) {
											sNumber = '0'.repeat(nPadding - sNumber.length) + sNumber;
										}
										sVal += sNumber;
										oCellValue.text = sVal;
										oCellValue.type = CellValueType.String;
									} else if (aTimePeriods) {
										let nIndexDay = nCurValue % aTimePeriods.length;
										if (nIndexDay < 0) {
											oCellValue.text = aReverseTimePeriods[~nIndexDay];
										} else {
											oCellValue.text = aTimePeriods[nIndexDay];
										}
										oCellValue.type = CellValueType.String;
									} else {
										oCellValue.number = nCurValue;
										oCellValue.type = CellValueType.Number;
									}
									oCopyCell.setValueData(new UndoRedoData_CellValueData(null, oCellValue));
								}
								else if(null != oFromCell)
								{
									//копируем полностью
									if(!oFromCell.isFormula()){
										oCopyCell.setValueData(oFromCell.getValueData());
										//todo
										// if(oCopyCell.isEmptyTextString())
										// wsTo._getHyperlink().remove({r1: oCopyCell.nRow, c1: oCopyCell.nCol, r2: oCopyCell.nRow, c2: oCopyCell.nCol});
									} else {
										var fromFormulaParsed = oFromCell.getFormulaParsed();
										var formulaArrayRef = fromFormulaParsed.getArrayFormulaRef();

										var _p_,offset, offsetArray, assemb;
										if(formulaArrayRef) {
											var intersectionFrom = from.intersection(formulaArrayRef);
											if(intersectionFrom) {
												if((intersectionFrom.c1 === oFromCell.nCol && intersectionFrom.r1 === oFromCell.nRow) || (intersectionFrom.c2 === oFromCell.nCol && intersectionFrom.r2 === oFromCell.nRow)) {

													offsetArray = oCopyCell.getOffset2(oFromCell.getName());
													intersectionFrom.setOffset(offsetArray);

													var intersectionTo = intersectionFrom.intersection(to);
													if(intersectionTo && !isAddFormulaArray(intersectionTo)) {
														addedFormulasArray.push(intersectionTo);
														//offset = oCopyCell.getOffset3(formulaArrayRef.c1 + 1, formulaArrayRef.r1 + 1);
														offset = new AscCommon.CellBase(intersectionTo.r1 - formulaArrayRef.r1, intersectionTo.c1 - formulaArrayRef.c1);
														_p_ = oFromCell.getFormulaParsed().clone(null, oFromCell, this);
														_p_.changeOffset(offset);

														var rangeFormulaArray = oCopyCell.ws.getRange3(intersectionTo.r1, intersectionTo.c1, intersectionTo.r2, intersectionTo.c2);
														rangeFormulaArray.setValue("=" + _p_.assemble(true), function (r) {}, null, intersectionTo);

														AscCommon.History.Add(AscCommonExcel.g_oUndoRedoArrayFormula,
															AscCH.historyitem_ArrayFromula_AddFormula, oCopyCell.ws.getId(),
															new Asc.Range(intersectionTo.c1, intersectionTo.r1, intersectionTo.c2, intersectionTo.r2),
															new AscCommonExcel.UndoRedoData_ArrayFormula(intersectionTo, "=" + oCopyCell.getFormulaParsed().assemble(true)));
													}

												}
											}
										} else {
											_p_ = oFromCell.getFormulaParsed().clone(null, oFromCell, this);
											offset = oCopyCell.getOffset2(oFromCell.getName());
											assemb = _p_.changeOffset(offset).assemble(true);
											let oCaProps = {};
											if (_p_.ca) { // for recursion formulas
												oCaProps.ca = _p_.ca;
												oCaProps.oldValue = oFromCell.getNumberValue();
											}
											oCopyCell.setFormula(assemb, null, null, oCaProps);
										}

									}
								}
							}
							//выставляем стиль после текста, потому что если выставить числовой стиль ячейки 'text', то после этого не применится формула
							if (null != oFromCell) {
								oCopyCell.setStyle(oFromCell.getStyle());
								if (bIsPromote)
									oCopyCell.setType(oFromCell.getType());
							}
						});
					}
				}
			}
			if(bIsPromote) {
				wb.dependencyFormulas.addToChangedRange( wsTo.Id, to );
			}
			//добавляем замерженые области
			var nDx = from.c2 - from.c1 + 1;
			var nDy = from.r2 - from.r1 + 1;
			var oMergedFrom = oCanPromote.oMergedFrom;
			if(null != oMergedFrom && oMergedFrom.all.length > 0)
			{
				for (var i = to.c1; i <= to.c2; i += nDx) {
					for (var j = to.r1; j <= to.r2; j += nDy) {
						for (var k = 0, length3 = oMergedFrom.all.length; k < length3; k++) {
							var oMergedBBox = oMergedFrom.all[k].bbox;
							var oNewMerged = new Asc.Range(i + oMergedBBox.c1 - from.c1, j + oMergedBBox.r1 - from.r1, i + oMergedBBox.c2 - from.c1, j + oMergedBBox.r2 - from.r1);
							if(to.contains(oNewMerged.c1, oNewMerged.r1)) {
								if(to.c2 < oNewMerged.c2)
									oNewMerged.c2 = to.c2;
								if(to.r2 < oNewMerged.r2)
									oNewMerged.r2 = to.r2;
								if(!oNewMerged.isOneCell())
									wsTo.mergeManager.add(oNewMerged, 1);
							}
						}
					}
				}
			}
			if(bIsPromote)
			{
				//добавляем ссылки
				//не как в Excel поддерживаются ссылки на диапазоны
				var oHyperlinks = wsFrom.hyperlinkManager.get(from);
				if(oHyperlinks.inner.length > 0)
				{
					for (var i = to.c1; i <= to.c2; i += nDx) {
						for (var j = to.r1; j <= to.r2; j += nDy) {
							for(var k = 0, length3 = oHyperlinks.inner.length; k < length3; k++){
								var oHyperlink = oHyperlinks.inner[k];
								var oHyperlinkBBox = oHyperlink.bbox;
								var oNewHyperlink = new Asc.Range(i + oHyperlinkBBox.c1 - from.c1, j + oHyperlinkBBox.r1 - from.r1, i + oHyperlinkBBox.c2 - from.c1, j + oHyperlinkBBox.r2 - from.r1);
								if (to.containsRange(oNewHyperlink))
									wsTo.hyperlinkManager.add(oNewHyperlink, oHyperlink.data.clone());
							}
						}
					}
				}


				var aDataValidations;
				if (wsFrom.dataValidations && wsTo.dataValidations) {
					wsTo.clearDataValidation([to], true);
					aDataValidations = wsFrom.dataValidations.getIntersectionByRange(from);
				}
				if(aDataValidations && aDataValidations.length > 0) {
					var newDataValidations = [];
					for (var i = to.c1; i <= to.c2; i += nDx) {
						for (var j = to.r1; j <= to.r2; j += nDy) {
							for(var k = 0, length3 = aDataValidations.length; k < length3; k++){
								var oDataValidation = aDataValidations[k];
								var ranges = oDataValidation.ranges;
								for (var n = 0; n < ranges.length; n++) {
									var _newRange = new Asc.Range(i + ranges[n].c1 - from.c1, j + ranges[n].r1 - from.r1, i + ranges[n].c2 - from.c1, j + ranges[n].r2 - from.r1);
									if (to.containsRange(_newRange)) {
										if (!newDataValidations[k]) {
											newDataValidations[k] = [];
										}
										newDataValidations[k].push(_newRange);
									}
								}
							}
						}
					}
					if (newDataValidations && newDataValidations.length) {
						for (var i = 0; i < newDataValidations.length; i++) {
							if (newDataValidations[i] && newDataValidations[i].length) {
								var fromDataValidation = wsFrom.getDataValidationById(aDataValidations[i].id);
								if (fromDataValidation) {
									fromDataValidation = fromDataValidation.data;
									var toDataValidation = fromDataValidation.clone();
									toDataValidation.ranges = fromDataValidation.ranges.concat(newDataValidations[i]);
									wsTo.dataValidations.change(wsTo, fromDataValidation, toDataValidation, true);
								}
							}
						}
					}
				}

				if (wsTo.aSparklineGroups && wsTo.aSparklineGroups.length) {
					wsTo.removeSparklines(to);
				}
				if(wsFrom === wsTo && wsFrom.aSparklineGroups && wsFrom.aSparklineGroups.length) {
					for (var i = to.c1; i <= to.c2; i += nDx) {
						for (var j = to.r1; j <= to.r2; j += nDy) {
							for(var k = 0, length3 = wsFrom.aSparklineGroups.length; k < length3; k++){
								var _arrSparklines = wsFrom.aSparklineGroups[k].getModifySparklinesForPromote(from, to, new AscCommon.CellBase(j - from.r1, i - from.c1));
								if (_arrSparklines) {
									wsFrom.aSparklineGroups[k].setSparklines(_arrSparklines, true, true);
								}
							}
						}
					}
				}
			}

			var aRules;
			if (wsTo.isConditionalFormattingRules()) {
				wsTo.clearConditionalFormattingRulesByRanges([to])
		}
			if (wsFrom.isConditionalFormattingRules()) {
				aRules = wsFrom.getIntersectionRules(from);
			}
			if(aRules && aRules.length > 0) {
				//TODO сделать объединение диапазонов! (допустим, при копировании стиля на несколько ячеек)
				var newRules = [];
				for (var i = to.c1; i <= to.c2; i += nDx) {
					for (var j = to.r1; j <= to.r2; j += nDy) {
						for(var k = 0, length3 = aRules.length; k < length3; k++){
							var oRule = aRules[k];
							var ranges = oRule.ranges;
							for (var n = 0; n < ranges.length; n++) {
								var _newRange = new Asc.Range(i + ranges[n].c1 - from.c1, j + ranges[n].r1 - from.r1, i + ranges[n].c2 - from.c1, j + ranges[n].r2 - from.r1);
								if (!to.containsRange(_newRange)) {
									_newRange = to.intersection(_newRange);
								}
								if (_newRange) {
									if (!newRules[k]) {
										newRules[k] = [];
									}
									newRules[k].push(_newRange);
								}
							}
						}
					}
				}
				if (newRules && newRules.length) {
					for (var i = 0; i < newRules.length; i++) {
						if (newRules[i] && newRules[i].length) {
							var fromRule = wsFrom.getRuleById(aRules[i].id);
							if (fromRule) {
								fromRule = fromRule.data;
								var toRule = fromRule.clone();
								//toRule.id = fromRule.id;
								if (bIsPromote) {
									toRule.id = fromRule.id;
									toRule.ranges = fromRule.ranges.concat(newRules[i]);
									toRule.combineRangesToSingle();
								} else {
									toRule.ranges = newRules[i];
								}
								wsTo.setCFRule(toRule);
							}
						}
					}
				}
			}
		}

		wsFrom.excludeHiddenRows(oldExcludeHiddenRows);
		AscCommon.History.EndTransaction();
		wb.dependencyFormulas.unlockRecal();
	}
	Range.prototype.createCellOnRowColCross=function(){
		var oThis = this;
		var bbox = this.bbox;
		var nRangeType = this._getRangeType(bbox);
		if(c_oRangeType.Row == nRangeType)
		{
			this._foreachColNoEmpty(function(col){
				if(null != col.xfs)
				{
					for(var i = bbox.r1; i <= bbox.r2; ++i)
						oThis.worksheet._getCell(i, col.index, function(){});
				}
			}, null);
		}
		else if(c_oRangeType.Col == nRangeType)
		{
			this._foreachRowNoEmpty(function(row){
				if(null != row.xfs)
				{
					for(var i = bbox.c1; i <= bbox.c2; ++i)
						oThis.worksheet._getCell(row.index, i, function(){});
				}
			}, null);
		}
	};

	Range.prototype.getLeftTopCell = function(fAction) {
		return this.worksheet._getCell(this.bbox.r1, this.bbox.c1, fAction);
	};

	Range.prototype.getLeftTopCellNoEmpty = function(fAction) {
		return this.worksheet._getCellNoEmpty(this.bbox.r1, this.bbox.c1, fAction);
	};

	Range.prototype.move = function (oBBoxTo, copyRange, wsTo) {
		this.worksheet._moveRange(this.bbox, oBBoxTo, copyRange, wsTo);
	};

	function SweepLineRowIterator() {
		this.cellsByCol = null;
		this.row = -1;
		this.col = -1;
		this.colData = null;

		this.colDatas = [];
		this.colDatasCol = [];
		this.colDatasIndex = 0;

		this.toInsert = [];
		this.toInsertIndex = 0;
		this.toDelete = [];
		this.toDeleteIndex = 0;

		this.r1 = [];
		this.r1Index = 0;
		this.r2 = [];
		this.r2Index = 0;
	}

	SweepLineRowIterator.prototype.init = function (cellsByCol, r1, c1, c2) {
		this.cellsByCol = cellsByCol;
		this.row = r1 - 1;
		this.col = -1;
		this.colData = null;

		this.colDatas = [];
		this.colDatasCol = [];
		this.colDatasIndex = 0;

		this.toInsert = [];
		this.toInsertIndex = 0;
		this.toDelete = [];
		this.toDeleteIndex = 0;
		this.r1 = [];
		this.r1Index = 0;
		this.r2 = [];
		this.r2Index = 0;

		c2 = Math.min(c2, this.cellsByCol.length - 1);
		for (let i = c1; i <= c2; i++) {
			let colData = this.cellsByCol[i];
			if (colData && r1 <= colData.getMaxIndex()) {
				//minIndex is max of
				let elem = {col: i, minIndex: Math.max(r1, colData.getMinIndex()), maxIndex: colData.getMaxIndex(), isInserted: false}
				this.r1.push(elem);
				this.r2.push(elem);
			}
		}
		//Array.sort is stable in all browsers
		//https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array/sort#browser_compatibility
		this.r1.sort(function (a, b) {
			return a.minIndex - b.minIndex
		});
		this.r2.sort(function (a, b) {
			return a.maxIndex - b.maxIndex
		});
	};
	SweepLineRowIterator.prototype.setRow = function (row) {
		while (this.row < row) {
			//structure of r1 and r2 forces rows to be changed with fixed step
			this.nextRow();
			if (this.row < row) {
				while (this.nextCol()) {
				}
			}
		}
	}
	SweepLineRowIterator.prototype.nextRow = function () {
		this.row++;
		this.col = -1;
		this.colData = null;
		this.colDatasIndex = 0;
		while (this.r1Index < this.r1.length && this.row >= this.r1[this.r1Index].minIndex) {
			let elem = this.r1[this.r1Index];
			if (this.row <= elem.maxIndex) {
				elem.isInserted = true;
				this.toInsert.push(elem.col);
			}
			this.r1Index++;
		}
		while (this.r2Index < this.r2.length && this.row > this.r2[this.r2Index].maxIndex) {
			let elem = this.r2[this.r2Index];
			if (elem.isInserted) {
				this.toDelete.push(elem.col);
			}
			this.r2Index++;
		}
	}
	SweepLineRowIterator.prototype.nextCol = function () {
		this.col = -1;
		this.colData = null;
		if (this.colDatasIndex < this.colDatasCol.length) {
			this.col = this.colDatasCol[this.colDatasIndex];
			this.colData = this.colDatas[this.colDatasIndex];
			this.colDatasIndex++;
			//delete
			let toDeleteOld = this.toDeleteIndex;
			while (this.toDeleteIndex < this.toDelete.length && this.toDelete[this.toDeleteIndex] === this.col) {
				this.toDeleteIndex++;
				if (this.colDatasIndex < this.colDatasCol.length) {
					this.col = this.colDatasCol[this.colDatasIndex];
					this.colData = this.colDatas[this.colDatasIndex];
					this.colDatasIndex++;
				} else {
					this.col = -1;
					this.colData = undefined;
				}
			}
			if (toDeleteOld !== this.toDeleteIndex) {
				let deleteCount = this.toDeleteIndex - toDeleteOld;
				this.colDatasIndex -= -1 === this.col ? deleteCount : deleteCount + 1;
				this.colDatas.splice(this.colDatasIndex, deleteCount);
				this.colDatasCol.splice(this.colDatasIndex, deleteCount);
				if (this.colDatasIndex < this.colDatasCol.length) {
					this.col = this.colDatasCol[this.colDatasIndex];
					this.colData = this.colDatas[this.colDatasIndex];
					this.colDatasIndex++;
				} else {
					this.col = -1;
					this.colData = undefined;
				}
			}
		}
		//insert
		let toInsertOld = this.toInsertIndex;
		if (-1 !== this.col) {
			while (this.toInsertIndex < this.toInsert.length && this.toInsert[this.toInsertIndex] < this.col) {
				this.toInsertIndex++;
			}
		} else {
			this.toInsertIndex = this.toInsert.length;
		}
		if (toInsertOld !== this.toInsertIndex) {
			let cols = [];
			let datas = [];
			for (let i = toInsertOld; i < this.toInsertIndex; ++i) {
				let curCol = this.toInsert[i];
				let colData = this.cellsByCol[curCol];
				if (colData && this.row <= colData.getMaxIndex()) {
					cols.push(curCol);
					datas.push(colData);
				}
			}
			if (-1 !== this.col) {
				this.colDatasIndex--;
			}
			this.colDatas.splice.apply(this.colDatas, [this.colDatasIndex, 0].concat(datas));
			this.colDatasCol.splice.apply(this.colDatasCol, [this.colDatasIndex, 0].concat(cols));

			if (this.colDatasIndex < this.colDatasCol.length) {
				this.col = this.colDatasCol[this.colDatasIndex];
				this.colData = this.colDatas[this.colDatasIndex];
				this.colDatasIndex++;
			} else {
				this.col = -1;
				this.colData = undefined;
			}
		}
		return -1 !== this.col;
	};

	function RowIterator() {
	}

	RowIterator.prototype.init = function (ws, r1, c1, c2) {
		this.ws = ws;
		this.cell = new Cell(ws);
		this.ws.workbook.loadCells.push(this.cell);

		this.iter = new SweepLineRowIterator();
		this.iter.init(ws.cellsByCol, r1, c1, c2);
	};
	RowIterator.prototype.release = function () {
		this.cell.saveContent(true);
		this.ws.workbook.loadCells.pop();
	};
	RowIterator.prototype.setRow = function (index) {
		this.iter.setRow(index);
	};
	RowIterator.prototype.next = function () {
		var wb = this.ws.workbook;
		this.cell.saveContent(true);
		while (this.iter.nextCol()) {
			var nRow = this.iter.row;
			var nCol = this.iter.col;
			var colData = this.iter.colData;
			var targetCell = null;
			for (var k = 0; k < wb.loadCells.length - 1; ++k) {
				var elem = wb.loadCells[k];
				if (elem.nRow === nRow && elem.nCol === nCol && this.ws === elem.ws) {
					targetCell = elem;
					break;
				}
			}
			if (null === targetCell) {
				if (this.cell.loadContent(nRow, nCol, colData)) {
					return this.cell;
				}
			} else {
				return targetCell;
			}
		}
	};
//-------------------------------------------------------------------------------------------------
	/**
	 * @constructor
	 */
	function PromoteHelper(bVerical, bReverse, bbox){
		//автозаполнение происходит всегда в правую сторону, поэтому менются индексы в методе add, и это надо учитывать при вызове getNext
		this.bVerical = bVerical;
		this.bReverse = bReverse;
		this.bbox = bbox;
		this.oDataRow = {};
		//для get
		this.oCurRow = null;
		this.nCurColIndex = null;
		this.nRowLength = 0;
		this.nColLength = 0;
		if(this.bVerical)
		{
			this.nRowLength = this.bbox.c2 - this.bbox.c1 + 1;
			this.nColLength = this.bbox.r2 - this.bbox.r1 + 1;
		}
		else
		{
			this.nRowLength = this.bbox.r2 - this.bbox.r1 + 1;
			this.nColLength = this.bbox.c2 - this.bbox.c1 + 1;
		}
		// Attributes for work with context menu of autofill.
		this.bFillHandleRightClick = false;
		this.nFillMenuChosenProp = null;
		this.nLastDayInMonth = null;
		this.nDaysInYear = null;
		this.nPrevDateValue = null;
		this.nPrevIntDateValue = null;
		this.bDiffDirectionDateTime = null;
		this.nCurrentDayValue = null; // Using for Date & Time format for month step mode when day changed.

		this.bOneSelectedCell = false;
	}
	PromoteHelper.prototype = {
		add: function(nRow, nCol, nVal, bDelimiter, sPrefix, padding, bDate, oAdditional, aTimePeriods){
			if(this.bVerical) {
				//транспонируем для удобства
				var temp = nRow;
				nRow = nCol;
				nCol = temp;
			}
			if(this.bReverse)
				nCol = this.nColLength - nCol - 1;
			var row = this.oDataRow[nRow];
			if(null == row) {
				row = {};
				this.oDataRow[nRow] = row;
			}
			row[nCol] = new cDataRow(nCol, nVal, bDelimiter, sPrefix, padding, bDate, oAdditional, aTimePeriods);
			// Checking Date format has Date & Time
			if (bDate && oAdditional.xfs && oAdditional.xfs.num && oAdditional.xfs.num.getFormat()) {
				let oNumFormat = oNumFormatCache.get(oAdditional.xfs.num.getFormat());
				if (oNumFormat.isTimeFormat() && oNumFormat.isDateFormat()) {
					row[nCol].setIsDateTime(true);
				} else if (oNumFormat.isTimeFormat() && !oNumFormat.isDateFormat()) {
					row[nCol].setIsTime(true);
				}
			}
			row[nCol].setStartValue(nVal);
		},
		isOnlyIntegerSequence: function(){
			var bRes = true;
			var bEmpty = true;
			for(var i in this.oDataRow) {
				var row = this.oDataRow[i];
				for(var j in row) {
					var data = row[j];
					bEmpty = false;
					if(!(null != data.getVal() && true != data.getIsDate() && null == data.getPrefix())) {
						bRes = false;
						break;
					}
				}
				if(!bRes)
					break;
			}
			if(bEmpty)
				bRes = false;
			return bRes;
		},
		_promoteSequence: function(aDigits){
			// Это коэффициенты линейного приближения (http://office.microsoft.com/ru-ru/excel-help/HP010072685.aspx)
			// y=a1*x+a0 (где: x=0,1....; y=значения в ячейках; a0 и a1 - это решения приближения функции методом наименьших квадратов
			// (n+1)*a0        + (x0+x1+....)      *a1=(y0+y1+...)
			// (x0+x1+....)*a0 + (x0*x0+x1*x1+....)*a1=(y0*x0+y1*x1+...)
			// http://www.exponenta.ru/educat/class/courses/vvm/theme_7/theory.asp
			var a0 = 0.0;
			var a1 = 0.0;
			// Индекс X
			var nX = 0;
			if(1 == aDigits.length)
			{
				nX = 1;
				a1 = 1;
				a0 = aDigits[0].y;
			}
			else
			{
				// (n+1)
				var nN = aDigits.length;
				// (x0+x1+....)
				var nXi = 0;
				// (x0*x0+x1*x1+....)
				var nXiXi = 0;
				// (y0+y1+...)
				var dYi = 0.0;
				// (y0*x0+y1*x1+...)
				var dYiXi = 0.0;

				// Цикл по всем строкам
				for (var i = 0, length = aDigits.length; i < length; ++i)
				{
					var data = aDigits[i];
					nX = data.x;
					var dValue = data.y;

					// Вычисляем значения
					nXi += nX;
					nXiXi += nX * nX;
					dYi += dValue;
					dYiXi += dValue * nX;
				}
				nX++;

				// Теперь решаем систему уравнений
				// Общий детерминант
				var dD = nN * nXiXi - nXi * nXi;
				// Детерминант первого корня
				var dD1 = dYi * nXiXi - nXi * dYiXi;
				// Детерминант второго корня
				var dD2 = nN * dYiXi - dYi * nXi;

				a0 = dD1 / dD;
				a1 = dD2 / dD;
			}
			return {a0: a0, a1: a1, nX: nX};
		},
		_addSequenceToRow : function(nRowIndex, aSortRowIndex, row, aCurSequence){
			if(aCurSequence.length > 0) {
				var oFirstData = aCurSequence[0];
				var bCanPromote = true;
				const bFillHandleRightClick = this.getFillHandleRightClick();
				//если последовательность состоит из одного числа и той же колонке есть еще последовательности, то надо копировать, а не автозаполнять
				if(aCurSequence.length === 1 && !bFillHandleRightClick) {
					var bVisitRowIndex = false;
					var oVisitData = null;
					for(var i = 0, length = aSortRowIndex.length; i < length; i++) {
						var nCurRowIndex = aSortRowIndex[i];
						if(nRowIndex == nCurRowIndex) {
							bVisitRowIndex = true;
							if(oVisitData && oFirstData.prefixDataCompare(oVisitData)) {
								bCanPromote = false;
								break;
							}
						} else {
							var oCurRow = this.oDataRow[nCurRowIndex];
							if(oCurRow) {
								var data = oCurRow[oFirstData.getCol()];
								if(null != data) {
									if(null != data.getVal()) {
										oVisitData = data;
										if(bVisitRowIndex) {
											if(oFirstData.prefixDataCompare(oVisitData))
												bCanPromote = false;
											break;
										}
									}
									else if(data.getDelimiter()) {
										oVisitData = null;
										if(bVisitRowIndex)
											break;
									}
								}
							}
						}
					}
				}
				if(bCanPromote) {
					var nMinIndex = null;
					var nMaxIndex = null;
					var bValidIndexDif = true;
					var nPrevX = null;
					var nPrevVal = null;
					var nIndexDif = null;
					var nValueDif = null;
					var nMaxPadding = 0;
					//анализируем последовательность, если числа расположены не на одинаковом расстоянии, то считаем их сплошной последовательностью
					//последовательность с промежутками может быть только целочисленной
					for(let i = 0, length = aCurSequence.length; i < length; i++) {
						var data = aCurSequence[i];
						var nCurX = data.getCol();
						let nCurVal = data.getVal();
						if(null == nMinIndex || null == nMaxIndex)
							nMinIndex = nMaxIndex = nCurX;
						else {
							if(nCurX < nMinIndex)
								nMinIndex = nCurX;
							if(nCurX > nMaxIndex)
								nMaxIndex = nCurX;
						}
						if(bValidIndexDif) {
							if(null != nPrevX && null != nPrevVal) {
								var nCurDif = nCurX - nPrevX;
								var nCurValDif = nCurVal - nPrevVal;
								if(null == nIndexDif || null == nCurValDif) {
									nIndexDif = nCurDif;
									nValueDif = nCurValDif;
								}
								else if(nIndexDif != nCurDif || nValueDif != nCurValDif) {
									nIndexDif = null;
									bValidIndexDif = false;
								}
							}
						}
						nMaxPadding = Math.max(nMaxPadding, data.getPadding());
						nPrevX = nCurX;
						nPrevVal = nCurVal;
					}
					var bWithSpace = false;
					if(null != nIndexDif) {
						nIndexDif = Math.abs(nIndexDif);
						if(nIndexDif > 1)
							bWithSpace = true;
					}
					//заполняем массив с координатами
					var bExistSpace = false;
					nPrevX = null;
					var aDigits = [];
					for (let i = 0, length = aCurSequence.length; i < length; i++) {
						var data = aCurSequence[i];
						data.setPadding(nMaxPadding);
						var nCurX = data.getCol();
						var x = nCurX - nMinIndex;
						if(null != nIndexDif && nIndexDif > 0)
							x /= nIndexDif;
						if(null != nPrevX && nCurX - nPrevX > 1)
							bExistSpace = true;
						var y = data.getVal();
						//даты автозаполняем только по целой части
						if(data.getIsDate() && !data.getIsDateTime() && !data.getIsMixedDateFormat() && !data.getIsTime()) {
							y = parseInt(y);
						}
						aDigits.push({x: x, y: y});
						nPrevX = nCurX;
					}
					if(aDigits.length > 0) {
						let bMixedDateFormat = oFirstData.getIsMixedDateFormat();
						let bTime = oFirstData.getIsTime();
						let bDate = oFirstData.getIsDate();
						let bDateTime = oFirstData.getIsDateTime();
						let oSequence = null;
						const aMenuPropsForDate = [oFillType.fillMonths, oFillType.fillYears];

						// For activating month or year calculate mode in default autofill. Using for data of date format.
						if (bDate && this.getFillMenuChosenProp() === oFillType.fillSeries) {
							this.setFillMenuChosenProp(null);
						}
						if (this.getFillMenuChosenProp() == null && bDate && aDigits.length > 1) {
							this.initFillMenuChosenProp(aDigits, oFirstData.getVal());
						}
						let nMenuChosenProp = this.getFillMenuChosenProp();
						let bIsCalcDateMode = bDate && nMenuChosenProp != null && aMenuPropsForDate.includes(nMenuChosenProp);
						if (bDate && nMenuChosenProp != null) {
							let nLastValFromRange = aDigits[aDigits.length - 1].y;
							this.setPrevDateValue(nLastValFromRange);
							this.setPrevIntDateValue(nLastValFromRange);
						}
						if (bMixedDateFormat || (bIsCalcDateMode && bDateTime && aDigits.length > 2)) {
							/* In a mixed date format, the promotion sequence should be calculated similarly to the Date format.
							But, the original array should be saved to use the time portion of the cell with the Date & Time format
							 as a constant value for the correct sequence. */
							let aIntDigits = aDigits.map(function (elem) {
								return {x: elem.x, y: parseInt(elem.y)};
							});
							oSequence = this._promoteSequence(aIntDigits);
						} else {
							oSequence = this._promoteSequence(aDigits);
						}
						if (aDigits.length === 1) {
							this.setIsOneSelectedCell(true);
							if (bTime) {
								// For one selected cell, step for Time format must be 1 hour.
								// The time component of a serial value increases by 1/86,400 each second.
								oSequence.a1 = (1 / 86400) * 3600;
							}
							if (this.bReverse) {
								//меняем коэффициенты для случая одного числа в последовательности, иначе она в любую сторону будет возрастающей
								oSequence.a1 *= -1;
								if (bTime && oSequence.a0 === 0) {
									oSequence.a0 = 1;
								}
							}
						}
						if (aDigits.length === 2 && bDateTime) {
							this.initDiffDirectionDateTime(aDigits);
						}
						//для дат и чисел с префиксом автозаполняются только целочисленные последовательности
						let nFirstVal = oFirstData.getVal();
						let bIsNotIntegerSequence = oSequence.a1 !== parseInt(oSequence.a1);
						let bDateCalcModeHasDiffDay = aDigits.length > 1 && bIsCalcDateMode && this.isDiffDaysForCalcMode(aDigits, nFirstVal);
						let bDateCalcModeCorrectSeq = bIsCalcDateMode && !bDateCalcModeHasDiffDay && parseInt(nFirstVal) === Math.round(oSequence.a0);
						let sPrefix = oFirstData.getPrefix();
						let bDelimiter = oFirstData.getDelimiter();
						// For Date format
						let bStepForDateIsZero = bDate && oSequence.a1 === 0;
						let bIntStartValue = bDate && oSequence.a0 === parseInt(oSequence.a0);
						// For Date & Time format
						let bDayOfDateDiff = aDigits.length > 1 ? Math.floor(aDigits[0].y) !== Math.floor(aDigits[1].y) : true;
						let bDateTimeSeqIsCorrect = aDigits[0].y.toFixed(3) === oSequence.a0.toFixed(3);
						// If Date & Time format hasn't diff in days of date, but sequence by time is not correct we use step = 0
						// and start sequence point is equal first element of aDigits or for reverse sequence the last element of aDigit.
						if (bDateTime && !bDayOfDateDiff && !bDateTimeSeqIsCorrect && (oSequence.a1 < 1 && oSequence.a1 > -1)) {
							oSequence.a0 = this.bReverse ? aDigits[aDigits.length - 1].y : aDigits[0].y;
							oSequence.a1 = parseInt(oSequence.a1);
						}
						if (bMixedDateFormat || (bIsCalcDateMode && bDateTime && aDigits.length > 2)) {
							// The start element needs to add part of the time from the first (or last for reverse sequences) cell
							// with Date & Time format from the selected range.
							let aDigitsWithDateTime = aDigits.filter(function(item) {
								return item.y !== parseInt(item.y);
							});
							let nDateTimeVal = this.bReverse ? aDigitsWithDateTime[aDigitsWithDateTime.length - 1].y : aDigitsWithDateTime[0].y;
							let nTimePart = nDateTimeVal - parseInt(nDateTimeVal);
							oSequence.a0 = parseInt(oSequence.a0) + nTimePart;
						}
						if (bTime && bDayOfDateDiff && aDigits.length > 1) {
							if (this.bReverse && aDigits[0].y - aDigits[aDigits.length - 1].y >= 1) {
								let nDateVal = parseInt(aDigits[aDigits.length - 1].y)
								oSequence.a0 = (oSequence.a0 - parseInt(oSequence.a0)) + nDateVal;
							}
							oSequence.a1 = oSequence.a1 - parseInt(oSequence.a1);
						}
						if(!((sPrefix != null || (bDate && !bDateTime && !bMixedDateFormat && !bTime)) && bIsNotIntegerSequence && !bDateCalcModeCorrectSeq)) {
							// If for Date or Date & Time format the sequence is not correct or sequence step is 0, skip work with oSequence
							if (bDate) {
								if (bStepForDateIsZero && !bDateTime && !bMixedDateFormat) {
									return;
								}
								if (!bIntStartValue && !bDateCalcModeCorrectSeq && !bDateTime && !bTime) {
									return;
								}
								if (bDateTime && bDayOfDateDiff && !bDateTimeSeqIsCorrect && !bDateCalcModeCorrectSeq) {
									return;
								}
							}
							if(!bWithSpace && bExistSpace) {
								for (let i = nMinIndex; i <= nMaxIndex; i++) {
									var data = row[i];
									if(null == data) {
										data = new cDataRow(i, null, bDelimiter, sPrefix, null, bDate);
										row[i] = data;
									}
									data.setSequence(oSequence);
								}
							}
							else {
								for(let i = 0, length = aCurSequence.length; i < length; i++) {
									var nCurX = aCurSequence[i].nCol;
									if(null != nCurX)
										row[nCurX].setSequence(oSequence);
								}
							}
						}
					}
				}
			}
		},
		finishAdd : function(bCopy) {
			if(true != bCopy) {
				var aSortRowIndex = [];
				for(var i in this.oDataRow)
					aSortRowIndex.push(i - 0);
				aSortRowIndex.sort(fSortAscending);
				for(var i = 0, length = aSortRowIndex.length; i < length; i++) {
					var nRowIndex = aSortRowIndex[i];
					var row = this.oDataRow[nRowIndex];
					//собираем информация о последовательностях в row
					var aSortIndex = [];
					for(var j in row)
						aSortIndex.push(j - 0);
					aSortIndex.sort(fSortAscending);
					var aCurSequence = [];
					var oPrevRowData = null;
					for(var j = 0, length2 = aSortIndex.length; j < length2; j++) {
						var nColIndex = aSortIndex[j];
						var rowData = row[nColIndex];
						let rowDataNext = nColIndex + 1 < length2 && row[nColIndex + 1];
						let bDate = rowData.getIsDate();
						let bDateTime = rowData.getIsDateTime();
						let bDateNextRow = rowDataNext && rowDataNext.getIsDate();
						let bDateTimeNextRow = rowDataNext && rowDataNext.getIsDateTime();
						if ((bDate && bDateTime && bDateNextRow && !bDateTimeNextRow) || (bDate && !bDateTime && bDateNextRow && bDateTimeNextRow)) {
							let nFirstColIndex = aSortIndex[0];
							let oFirstRowData = row[nFirstColIndex];
							oFirstRowData.setIsMixedDateFormat(true);
							oFirstRowData.setIsDateTime(false);
							if (rowDataNext) {
								rowDataNext.setIsMixedDateFormat(true);
								rowDataNext.setIsDateTime(false);
							}
						}
						var bAddToSequence = false;
						let nVal = rowData.getVal();
						let bDelimiter = rowData.getDelimiter();
						if(null != nVal) {
							bAddToSequence = true;
							if(null != oPrevRowData && !rowData.compare(oPrevRowData)) {
								this._addSequenceToRow(nRowIndex, aSortRowIndex, row, aCurSequence);
								aCurSequence = [];
								oPrevRowData = null;
							}
							oPrevRowData = rowData;
						}
						else if(bDelimiter) {
							this._addSequenceToRow(nRowIndex, aSortRowIndex, row, aCurSequence);
							aCurSequence = [];
							oPrevRowData = null;
						}
						if(bAddToSequence)
							aCurSequence.push(rowData);
					}
					this._addSequenceToRow(nRowIndex, aSortRowIndex, row, aCurSequence);
				}
			}
		},
		setIndex: function(index){
			if(0 != this.nRowLength && index >= this.nRowLength)
				index = index % (this.nRowLength);
			this.oCurRow = this.oDataRow[index];
			this.nCurColIndex = 0;
		},
		getNext: function() {
			let oRes = null;
			if(this.oCurRow) {
				oRes = this.oCurRow[this.nCurColIndex];
				if(oRes != null) {
					oRes.setCurValue(null);
					if(oRes.getSequence() != null) {
						let sequence = oRes.getSequence();
						// Define modes fill: "By Month", " By Week" or "By Years" for values in Date format.
						let nFillMenuChosenProp = this.getFillMenuChosenProp();
						const aMenuPropsForDate = [oFillType.fillWeekdays, oFillType.fillMonths, oFillType.fillYears];
						if(oRes.getPrefix() != null) {
							oRes.setCurValue(Math.abs(sequence.a1 * sequence.nX + sequence.a0));
						} else if (oRes.getIsDate() && !oRes.getTimePeriods()) {
							let nCurValue;
							if (nFillMenuChosenProp && aMenuPropsForDate.includes(nFillMenuChosenProp)) {
								nCurValue = this.calculateDate(oRes);
							} else {
								nCurValue = sequence.a1 * sequence.nX + sequence.a0;
							}
							if (nCurValue >= 0 || (oRes.getIsDateTime() && parseInt(sequence.a1) !== sequence.a1 && sequence.a1 > -1)) {
								oRes.setCurValue(nCurValue);
							} else if (oRes.getIsTime() && nCurValue < 0) {
								// Reset nX and start from beginning
								sequence.nX = this.getIsOneSelectedCell() ? 1 : 0;
								oRes.setCurValue(sequence.a1 * sequence.nX + sequence.a0);
							}
						} else {
							oRes.setCurValue(sequence.a1 * sequence.nX + sequence.a0);
						}
						sequence.nX++;
					}
				}
				this.nCurColIndex++;
				if(this.nCurColIndex >= this.nColLength)
					this.nCurColIndex = 0;
			}
			return oRes;
		},
		getFillHandleRightClick: function () {
			return this.bFillHandleRightClick;
		},
		setFillHandleRightClick: function (bFillHandleRightClick) {
			this.bFillHandleRightClick = bFillHandleRightClick;
		},
		/**
		 * Returns chosen property from context menu.
		 * @returns {Asc.c_oAscFillType}
		 */
		getFillMenuChosenProp: function () {
			return this.nFillMenuChosenProp;
		},
		/**
		 * Sets chosen property from context menu.
		 * @param {Asc.c_oAscFillType} nFillMenuChosenProp
		 */
		setFillMenuChosenProp: function (nFillMenuChosenProp) {
			this.nFillMenuChosenProp = nFillMenuChosenProp;
		},
		/**
		 * Returns the last day of a month.
		 * @returns {number}
		 */
		getLastDayOfMonth: function () {
			return this.nLastDayInMonth;
		},
		/**
		 * Sets the last day of a month taken from ExcelDateValue.
		 * @param {number} nExcelDateValue
		 */
		setLastDayOfMonth: function (nExcelDateValue) {
			const CURRENT_MONTH = 0;
			let dtLastDayInMonth = AscCommonExcel.getLastDayInMonth(nExcelDateValue, CURRENT_MONTH);

			this.nLastDayInMonth = dtLastDayInMonth.getDate();
		},
		/**
		 * Returns days in year.
		 * @returns {number}
		 */
		getDaysInYear: function () {
			return this.nDaysInYear;
		},
		/**
		 * Sets days in year taken from ExcelDateValue.
		 * @param {number} nExcelDateValue
		 */
		setDaysInYear: function (nExcelDateValue) {
			const oDayCountBasis = AscCommonExcel.DayCountBasis;

			this.nDaysInYear = AscCommonExcel.daysInYear(nExcelDateValue, oDayCountBasis.ActualActual).getValue();
		},
		/**
		 * Returns previous date value while calculating for modes: "Fill months", "Fill years" and "Fill weekdays".
		 * @returns {number}
		 */
		getPrevDateValue: function () {
			return this.nPrevDateValue;
		},
		/**
		 * Sets previous date value while calculating for modes: "Fill months", "Fill years" and "Fill weekdays".
		 * @param {number} nPrevDateValue
		 */
		setPrevDateValue: function (nPrevDateValue) {
			this.nPrevDateValue = nPrevDateValue;
		},
		/**
		 * Returns previous integer date value while calculating for modes: "Fill months", "Fill years".
		 * Using when step is fractional number.
		 * @returns {number}
		 */
		getPrevIntDateValue: function () {
			return this.nPrevIntDateValue;
		},
		/**
		 * Sets previous integer date value while calculating for modes: "Fill months", "Fill years".
		 * Using when step is fractional number.
		 * @param {number} nPrevIntDateValue
		 */
		setPrevIntDateValue: function (nPrevIntDateValue) {
			this.nPrevIntDateValue = nPrevIntDateValue;
		},
		/**
		 * Returns the current state of the flag recognizing different direction at date and time.
		 * @returns {boolean}
		 */
		getIsDiffDirectionDateTime: function () {
			return this.bDiffDirectionDateTime;
		},
		/**
		 * Sets the current state of the flag recognizing different direction at date and time.
		 * @param {boolean} bDiffDirectionDateTime
		 */
		setIsDiffDirectionDateTime: function (bDiffDirectionDateTime) {
			this.bDiffDirectionDateTime = bDiffDirectionDateTime;
		},
		/**
		 * Returns current day value for Date & Time format.
		 * Using for month step mode when day changed.
		 * @returns {number}
		 */
		getCurrentDayValue: function () {
			return this.nCurrentDayValue;
		},
		/**
		 * Sets current day value for Date & Time format.
		 * Using for month step mode when day changed.
		 * @param {number} nCurrentDayValue
		 */
		setCurrentDayValue: function (nCurrentDayValue) {
			this.nCurrentDayValue = nCurrentDayValue;
		},
		/**
		 * Initializes flag recognizing that date and time have different direction between each other.
		 * true - date and time have different direction.
		 * false - date and time have same direction
		 * @param {{x:number, y:number}[]} aDigits
		 */
		initDiffDirectionDateTime: function (aDigits) {
			let nFirstValueTimePart = aDigits[0].y - parseInt(aDigits[0].y);
			let nSecondValueTimePart = aDigits[1].y - parseInt(aDigits[1].y);
			let nFirstValueDatePart = parseInt(aDigits[0].y);
			let nSecondValueDatePart = parseInt(aDigits[1].y);
			let bAscTimePart = nSecondValueTimePart > nFirstValueTimePart;
			let bAscDatePart = nSecondValueDatePart > nFirstValueDatePart;

			this.setIsDiffDirectionDateTime(bAscTimePart !== bAscDatePart);
		},
		/**
		 * Initializes property of autofill context menu if in default autofill, has the required step to needed mode.
		 * Modes: "Fill months", "Fill years".
		 * @param {{x:number, y:number}[]} aDigits
		 * @param {number} nFirstValue
		 */
		initFillMenuChosenProp: function (aDigits, nFirstValue) {
			const SECOND_VALUE_INDEX = 1;
			const oFillType = Asc.c_oAscFillType;
			const nSecondValue = aDigits[SECOND_VALUE_INDEX].y;
			const dtFirstValue = new Asc.cDate().getDateFromExcel(nFirstValue < 60 ? nFirstValue + 1 : nFirstValue);
			const dtSecondValue = new Asc.cDate().getDateFromExcel(nSecondValue < 60 ? nSecondValue + 1 : nSecondValue);

			let nDayFirstValue = dtFirstValue.getDate();
			let nMonthFirstValue = dtFirstValue.getMonth();
			let nYearFirstValue = dtFirstValue.getFullYear();

			let nDaySecondValue = dtSecondValue.getDate();
			let nMonthSecondValue = dtSecondValue.getMonth();
			let nYearSecondValue = dtSecondValue.getFullYear();

			if (nDayFirstValue === nDaySecondValue && nMonthFirstValue !== nMonthSecondValue) {
				this.setFillMenuChosenProp(oFillType.fillMonths);
			} else if (nDayFirstValue === nDaySecondValue && nMonthFirstValue === nMonthSecondValue && nYearFirstValue !== nYearSecondValue) {
				this.setFillMenuChosenProp(oFillType.fillYears);
			}
		},
		/**
		 * Checks has different days in the date of the selected range for date calculate modes from the context menu.
		 * Applies for modes : "Fill months", "Fill years".
		 * @param {{x:number, y:number}[]} aDigits
		 * @param {number} nFirstValue
		 * @returns {boolean}
		 */
		isDiffDaysForCalcMode: function (aDigits, nFirstValue) {
			const SECOND_VALUE_INDEX = 1;
			const nSecondValue = aDigits[SECOND_VALUE_INDEX].y;
			const dtFirstValue = new Asc.cDate().getDateFromExcel(nFirstValue < 60 ? nFirstValue + 1 : nFirstValue);
			const dtSecondValue = new Asc.cDate().getDateFromExcel(nSecondValue < 60 ? nSecondValue + 1 : nSecondValue);

			return dtFirstValue.getDate() !== dtSecondValue.getDate();
		},
		/**
		 * Calculates date based on step and modes: "Fill months", "Fill years" or "Fill weekdays".
		 * @param {cDataRow} oDataRow
		 * @returns {number}
		 */
		calculateDate: function (oDataRow) {
			const DEFAULT_STEP = this.bReverse ? -1 : 1;
			const oSequence = oDataRow.getSequence();
			const nChosenMenuItem = this.getFillMenuChosenProp();
			let nStep =  null;
			let nStepDivider = null;
			let nUnitDate = null;
			// Flag for "Fill Month" and "Fill years". If the step is less than the last day of the month or days in the year,
			// we change nVal of DataRow and work with it as previous value for every cell of selected range.
			let bUseValFromDataRow = false;

			// Find last day of month and days in year for correct work with month and year mode.
			if (this.getLastDayOfMonth() == null && this.getDaysInYear() == null) {
				let nFirstVal = oDataRow.getVal();
				this.setLastDayOfMonth(nFirstVal);
				this.setDaysInYear(nFirstVal);
			}
			switch (nChosenMenuItem) {
				case Asc.c_oAscFillType.fillWeekdays:
					nUnitDate = oSeriesDateUnitType.weekday;
					nStepDivider = 1;
					break;
				case Asc.c_oAscFillType.fillMonths:
					nUnitDate = oSeriesDateUnitType.month
					nStepDivider = this.getLastDayOfMonth();
					break;
				case Asc.c_oAscFillType.fillYears:
					nUnitDate = oSeriesDateUnitType.year
					nStepDivider = this.getDaysInYear();
					break;
			}
			if (!this.getIsOneSelectedCell() && nChosenMenuItem !== Asc.c_oAscFillType.fillWeekdays) {
				bUseValFromDataRow =  Math.abs(oSequence.a1) < nStepDivider;
			}
			if (bUseValFromDataRow) {
				nStep = DEFAULT_STEP;
				const oCellInfo = {
					step: nStep,
					dateUnit: nUnitDate,
					previousValue: oDataRow.getVal(),
					previousIntValue: oDataRow.getVal(),
					expectedDayValue: oDataRow.getStartValue()
				};
				const oResult = _calculateDate(oCellInfo, true);
				oDataRow.setVal(oResult.previousValue);

				return oResult.currentValue;
			}
			nStep = this.getIsOneSelectedCell() ? oSequence.a1 : Math.round(oSequence.a1 / nStepDivider);
			const oCellInfo = {
				step: nStep,
				dateUnit: nUnitDate,
				previousValue: this.getPrevDateValue(),
				previousIntValue: this.getPrevIntDateValue(),
				expectedDayValue: oDataRow.getVal()
			};
			if (oDataRow.getIsDateTime() && this.nColLength === 2) {
				if (oSequence.nX === this.nColLength) { // For first iteration using element from selected range.
					this.setCurrentDayValue(oDataRow.getVal());
				}
				oCellInfo.expectedDayValue = this.getCurrentDayValue();
			}
			const oResult = _calculateDate(oCellInfo, true);
			let bFloatDate = oDataRow.getIsDateTime() || oDataRow.getIsMixedDateFormat();
			// Logic for time.
			if (bFloatDate && !this.getFillHandleRightClick()) {
				let nTimePart = oSequence.a0 - parseInt(oSequence.a0);
				if (oDataRow.getIsDateTime() && this.nColLength === 2) {
					nTimePart = oCellInfo.previousValue - parseInt(oCellInfo.previousValue);
					nTimePart += (oSequence.a1 - parseInt(oSequence.a1));
					if (this.getIsDiffDirectionDateTime()) {
						if (this.bReverse) {
							if (Math.sign(nTimePart) < 0) {
								nTimePart = nTimePart + 1;
							} else {
								nTimePart += 1;
							}
						} else {
							if (nTimePart > 1) {
								nTimePart -= parseInt(nTimePart);
							} else {
								nTimePart = nTimePart - 1;
							}
						}
					}
				}
				oResult.previousValue += nTimePart;
				if (oResult.previousIntValue) {
					oResult.previousIntValue += nTimePart;
					if (nTimePart > 1 || Math.sign(nTimePart) < 0) {
						this.setCurrentDayValue(oResult.previousIntValue);
					}
				}
				oResult.currentValue += nTimePart;
			}
			this.setPrevDateValue(oResult.previousValue);
			if (oResult.previousIntValue) {
				this.setPrevIntDateValue(oResult.previousIntValue);
			}

			return oResult.currentValue;
		},
		/**
		 * Returns flag that recognizes it as one selected cell.
		 * @returns {boolean}
		 */
		getIsOneSelectedCell: function () {
			return this.bOneSelectedCell;
		},
		/**
		 * Sets flag that recognizes it as one selected cell.
		 * @param {boolean} bOneSelectedCell
		 */
		setIsOneSelectedCell: function (bOneSelectedCell) {
			this.bOneSelectedCell = bOneSelectedCell;
		}
	};

//-------------------------------------------------------------------------------------------------

	/**
	 * @constructor
	 */
	function HiddenManager(ws) {
		this.ws = ws;
		this.hiddenRowsSum = [];
		this.hiddenColsSum = [];
		this.dirty = true;
		this.recalcHiddenRows = {};
		this.recalcHiddenCols = {};
		this.hiddenRowMin = gc_nMaxRow0;
		this.hiddenRowMax = 0;
	}
	HiddenManager.prototype.initPostOpen = function () {
		this.hiddenRowMin = gc_nMaxRow0;
		this.hiddenRowMax = 0;
	};
	HiddenManager.prototype.addHidden = function (isRow, index) {
		(isRow ? this.recalcHiddenRows : this.recalcHiddenCols)[index] = true;
		if (isRow) {
			this.hiddenRowMin = Math.min(this.hiddenRowMin, index);
			this.hiddenRowMax = Math.max(this.hiddenRowMax, index);
		}
		this.setDirty(true);
	};
	HiddenManager.prototype.getRecalcHidden = function () {
		var res = [];
		var i;
		for (i in this.recalcHiddenRows) {
			i = +i;
			res.push(new Asc.Range(0, i, gc_nMaxCol0, i));
		}
		for (i in this.recalcHiddenCols) {
			i = +i;
			res.push(new Asc.Range(i, 0, i, gc_nMaxRow0));
		}
		this.recalcHiddenRows = {};
		this.recalcHiddenCols = {};
		return res;
	};
	HiddenManager.prototype.getHiddenRowsRange = function() {
		var res;
		if (this.hiddenRowMin <= this.hiddenRowMax) {
			res = {r1: this.hiddenRowMin, r2: this.hiddenRowMax};
			this.hiddenRowMin = gc_nMaxRow0;
			this.hiddenRowMax = 0;
		}
		return res;
	};
	HiddenManager.prototype.setDirty = function(val) {
		this.dirty = val;
	};
	HiddenManager.prototype.getHiddenRowsCount = function(from, to) {
		return this._getHiddenCount(true, from, to);
	};
	HiddenManager.prototype.getHiddenColsCount = function(from, to) {
		return this._getHiddenCount(false, from, to);
	};
	HiddenManager.prototype._getHiddenCount = function(isRow, from, to) {
		//todo wrong result if 'to' is hidden
		if (this.dirty) {
			this.dirty = false;
			this._init();
		}
		var hiddenSum = isRow ? this.hiddenRowsSum : this.hiddenColsSum;
		var toCount = to < hiddenSum.length ? hiddenSum[to] : 0;
		var fromCount = from < hiddenSum.length ? hiddenSum[from] : 0;
		return fromCount - toCount;
	};
	HiddenManager.prototype._init = function() {
		if (this.ws) {
			this.hiddenColsSum = this._initHiddenSumCol(this.ws.aCols);
			this.hiddenRowsSum = this._initHiddenSumRow();
		}
	};
	HiddenManager.prototype._initHiddenSumCol = function(elems) {
		var hiddenSum = [];
		if (this.ws) {
			var i;
			var hiddenFlags = [];
			for (i in elems) {
				if (elems.hasOwnProperty(i)) {
					var elem = elems[i];
					if (null != elem && elem.getHidden()) {
						hiddenFlags[i] = 1;
					}
				}
			}
			var sum = 0;
			for (i = hiddenFlags.length - 1; i >= 0; --i) {
				if (hiddenFlags[i] > 0) {
					sum++;
				}
				hiddenSum[i] = sum;
			}
		}
		return hiddenSum;
	};
	HiddenManager.prototype._initHiddenSumRow = function() {
		var hiddenSum = [];
		if (this.ws) {
			var i;
			var hiddenFlags = [];
			this.ws._forEachRow(function(row) {
				if (row.getHidden()) {
					hiddenFlags[row.getIndex()] = 1;
				}
			});
			var sum = 0;
			for (i = hiddenFlags.length - 1; i >= 0; --i) {
				if (hiddenFlags[i] > 0) {
					sum++;
				}
				hiddenSum[i] = sum;
			}
		}
		return hiddenSum;
	};


	function tryTranslateToPrintArea(val) {
		const printAreaStr = "Print_Area";
		const printAreaStrLocale = AscCommon.translateManager && AscCommon.translateManager.getValue(printAreaStr);
		if(printAreaStrLocale && printAreaStrLocale.toLowerCase() === val.toLowerCase()) {
			return printAreaStr;
		}
		return null;
	}
	//-------------------------------------------------------------------------------------------------
	/**
	 * Class represents data using for calculating a sequence of filling cells in autofill.
	 * @constructor
	 */
	function cDataRow(nCol, nVal, bDelimiter, sPrefix, nPadding, bDate, oAdditional, aTimePeriods) {
		this.nCol = nCol;
		this.nVal = nVal;
		this.nStartVal = null;
		this.bDelimiter = bDelimiter;
		this.sPrefix = sPrefix;
		this.nPadding = nPadding;
		this.bDate = bDate;
		this.bDateTime = false;
		this.bTime = false;
		this.bMixedDateFormat = false;
		this.oAdditional = oAdditional;
		this.aTimePeriods = aTimePeriods;
		this.oSequence = null;
		this.nCurValue = null;
	}

	/**
	 * Method returns column number
	 * @memberof cDataRow
	 * @returns {number}
	 */
	cDataRow.prototype.getCol = function() {
		return this.nCol;
	};
	/**
	 * Method returns value of current cell
	 * @memberof cDataRow
	 * @returns {number}
	 */
	cDataRow.prototype.getVal = function() {
		return this.nVal;
	};
	/**
	 * Method sets value of current cell
	 * @memberof cDataRow
	 * @param {number} nVal
	 */
	cDataRow.prototype.setVal = function(nVal) {
		this.nVal = nVal;
	};
	/**
	 * Method returns initial value of current cell.
	 * Uses to autofill by context menu for  "Fill months",  "Fill years".
	 * @memberof cDataRow
	 * @returns {number}
	 */
	cDataRow.prototype.getStartValue = function() {
		return this.nStartVal;
	};
	/**
	 * Method sets initial value of current cell.
	 * Uses to autofill by context menu for  "Fill months",  "Fill years".
	 * @memberOf cDataRow
	 * @param {number} nVal
	 */
	cDataRow.prototype.setStartValue = function(nVal) {
		this.nStartVal = nVal;
	};
	/**
	 * Method returns flag which checks cell has delimiter.
	 * @memberof cDataRow
	 * @returns {boolean}
	 */
	cDataRow.prototype.getDelimiter = function() {
		return this.bDelimiter;
	};
	/**
	 * Method returns prefix of current cell
	 * @memberof cDataRow
	 * @returns {string}
	 */
	cDataRow.prototype.getPrefix = function () {
		return this.sPrefix;
	};
	/**
	 * Method returns padding of current cell
	 * @memberof cDataRow
	 * @returns {number}
	 */
	cDataRow.prototype.getPadding = function() {
		return this.nPadding;
	};
	/**
	 * Method sets padding of current cell
	 * @memberof cDataRow
	 * @param {number} nPadding
	 */
	cDataRow.prototype.setPadding = function(nPadding) {
		this.nPadding = nPadding;
	};
	/**
	 * Method returns flag which recognizes a current cell has format Date or not.
	 * @memberof cDataRow
	 * @returns {boolean}
	 */
	cDataRow.prototype.getIsDate = function() {
		return this.bDate;
	};
	/**
	 * Method returns flag which recognizes a current cell has format Date & Time or not.
	 * @memberof cDataRow
	 * @returns {boolean}
	 */
	cDataRow.prototype.getIsDateTime = function() {
		return this.bDateTime;
	};
	/**
	 * Method sets flag which recognizes a current cell has format Date & Time or not.
	 * @memberof cDataRow
	 * @param {boolean} bDateTime
	 */
	cDataRow.prototype.setIsDateTime = function(bDateTime) {
		this.bDateTime = bDateTime;
	};
	/**
	 * Method returns flag which recognizes a current cell has mixed format of Dates or not.
	 * @memberof cDataRow
	 * @returns {boolean}
	 */
	cDataRow.prototype.getIsMixedDateFormat = function () {
		return this.bMixedDateFormat;
	};
	/**
	 * Method sets flag which recognizes a current cell has mixed format of Dates or not.
	 * @memberof cDataRow
	 * @param {boolean} bMixedDateFormat
	 */
	cDataRow.prototype.setIsMixedDateFormat = function (bMixedDateFormat) {
		this.bMixedDateFormat = bMixedDateFormat;
	};
	/**
	 * Method returns flag which recognizes a current cell has format Time or not.
	 * @memberof cDataRow
	 * @returns {boolean}
	 */
	cDataRow.prototype.getIsTime = function() {
		return this.bTime;
	};
	/**
	 * Method sets flag which recognizes a current cell has format Time or not.
	 * @param {boolean} bTime
	 */
	cDataRow.prototype.setIsTime = function(bTime) {
		this.bTime = bTime;
	};
	/**
	 * Method returns additional data
	 * @memberof cDataRow
	 * @returns {Cell}
	 */
	cDataRow.prototype.getAdditional = function() {
		return this.oAdditional;
	};
	/**
	 * Method returns array of time periods
	 * @memberof cDataRow
	 * @returns {[]}
	 */
	cDataRow.prototype.getTimePeriods = function () {
		return this.aTimePeriods;
	};
	/**
	 * Method returns sequence of calculated step and start cell. For fill next cells in autofill.
	 * @memberof cDataRow
	 * @returns {null|{a0:number, a1:number, nX:number}}
	 */
	cDataRow.prototype.getSequence = function() {
		return this.oSequence;
	};
	/**
	 * Method sets sequence of calculated step and start cell. For fill next cells in autofill.
	 * @memberof cDataRow
	 * @param {{a0:number, a1:number, nX:number}} oSequence
	 */
	cDataRow.prototype.setSequence = function (oSequence) {
		this.oSequence = oSequence;
	};
	/**
	 * Method returns the current calculated value for cell who need to be filled.
	 * @memberof cDataRow
	 * @returns {number}
	 */
	cDataRow.prototype.getCurValue = function() {
		return this.nCurValue;
	};
	/**
	 * Method sets the current calculated value for cell who need to be filled.
	 * @memberof cDataRow
	 * @param {number} nCurValue
	 */
	cDataRow.prototype.setCurValue = function(nCurValue) {
		this.nCurValue = nCurValue;
	};
	/**
	 * Method compares 2 DataRow objects for equality
	 * @memberof cDataRow
	 * @param {cDataRow} oComparedRowData
	 * @returns {boolean}
	 */
	cDataRow.prototype.compare = function(oComparedRowData) {
		let sComparedTimePeriods =  oComparedRowData.getTimePeriods() ? oComparedRowData.getTimePeriods().join() : null;
		let sTimePeriods = this.getTimePeriods() ? this.getTimePeriods().join() : null;
		let bComparedDelimiter = oComparedRowData.getDelimiter();
		let bDelimiter = this.getDelimiter();
		let sComparedPrefix = oComparedRowData.getPrefix();
		let sPrefix = this.getPrefix();
		let bComparedDate = oComparedRowData.getIsDate();
		let bDate = this.getIsDate();

		return bComparedDelimiter === bDelimiter && sComparedPrefix === sPrefix && bComparedDate === bDate && sComparedTimePeriods === sTimePeriods;
	};
	/**
	 * Method compares bDate and Prefix for equality between two DataRow objects.
	 * @memberof cDataRow
	 * @param {cDataRow} oComparedRowData
	 * @returns {boolean}
	 */
	cDataRow.prototype.prefixDataCompare = function(oComparedRowData) {
		let sComparedPrefix = oComparedRowData.getPrefix();
		let sPrefix = this.getPrefix();
		let bComparedDate = oComparedRowData.getIsDate();
		let bDate = this.getIsDate();

		return sComparedPrefix === sPrefix && bComparedDate === bDate;
	};
	//-------------------------------------------------------------------------------------------------
	/**
	 * Class representing a Series feature
	 * @param {asc_CSeriesSettings} settings - Data from dialog window
	 * @constructor
	 */
	function CSerial (settings) {
		this.bVertical = settings.seriesIn === oSeriesInType.columns;
		this.nType = settings.type;
		this.nDateUnit = settings.dateUnit;
		this.nStep = settings.stepValue;
		this.nStopValue = settings.stopValue;
		this.bTrend = settings.trend;
		this.nChosenContextMenuProp = settings.contextMenuChosenProperty;

		this.oFromRange = null;
		this.oToRange = null;
		this.oActiveFillHandle = null;
		this.ws = null;
		this.nPrevValue = null;
		this.nIndex = null;

		// Iteration coordinates for range who needs to be fill
		this.nStartIndex = null;
		this.nEndIndex = null;
		this.nDirectionStep = null;
	}
	/**
	 * Returns the range direction flag. Vertical or horizontal direction
	 * @memberof CSerial
	 * @returns {boolean}
	 */
	CSerial.prototype.getVertical = function () {
		return this.bVertical;
	};
	/**
	 * Returns modes of progressions.
	 * @memberof CSerial
	 * @returns {c_oAscSeriesType}
	 */
	CSerial.prototype.getType = function () {
		return this.nType;
	};
	/**
	 * Returns modes for work with date
	 * @memberof CSerial
	 * @returns {c_oAscDateUnitType}
	 */
	CSerial.prototype.getDateUnit = function () {
		return this.nDateUnit;
	};
	/**
	 * Returns a step of progression
	 * @memberof CSerial
	 * @returns {number}
	 */
	CSerial.prototype.getStep = function () {
		return this.nStep;
	};
	/**
	 * Returns a stop value. Fills a progression to value inputted in "Stop value"
	 * @memberof CSerial
	 * @returns {number}
	 */
	CSerial.prototype.getStopValue = function () {
		return this.nStopValue;
	};
	/**
	 * Returns a flag of trend mode. true - Fills cells in trend mode. Step defines automatically. false - Fills cells according step value
	 * @memberof CSerial
	 * @returns {boolean}
	 */
	CSerial.prototype.getTrend = function () {
		return this.bTrend;
	};
	/**
	 * Returns a chosen context menu property
	 * @returns {asc_CSeriesSettings.contextMenuChosenProperty|c_oAscFillType}
	 */
	CSerial.prototype.getChosenContextMenuProp = function () {
		return this.nChosenContextMenuProp;
	};
	/**
	 * Sets a range of selected cells
	 * @memberof CSerial
	 * @param {Range} oFromRange
	 */
	CSerial.prototype.setFromRange = function (oFromRange) {
		this.oFromRange = oFromRange;
		this.setWs(oFromRange.worksheet);
	};
	/**
	 * Returns a range of selected cells
	 * @memberof CSerial
	 * @returns {Range}
	 */
	CSerial.prototype.getFromRange = function () {
		return this.oFromRange;
	};
	/**
	 * Sets a range of cells who will be fill
	 * @memberof CSerial
	 * @param {Range} oRange
	 */
	CSerial.prototype.setToRange = function (oRange) {
		this.oToRange = oRange;
	};
	/**
	 * Returns a range of cells who will be fill
	 * @memberof CSerial
	 * @returns {Range}
	 */
	CSerial.prototype.getToRange = function () {
		return this.oToRange;
	};
	/**
	 * Sets an active fill handle.
	 * @param {Asc.Range} oActiveFillHandle
	 */
	CSerial.prototype.setActiveFillHandle = function (oActiveFillHandle) {
		this.oActiveFillHandle = oActiveFillHandle;
	};
	/**
	 * Returns an active fill handle
	 * @returns {Asc.Range}
	 */
	CSerial.prototype.getActiveFillHandle = function () {
		return this.oActiveFillHandle;
	};
	/**
	 * Sets a worksheet of selected range
	 * @param {Worksheet} ws
	 */
	CSerial.prototype.setWs = function (ws) {
		this.ws = ws;
	};
	/**
	 * Returns a worksheet of selected range
	 * @memberof CSerial
	 * @returns {Worksheet}
	 */
	CSerial.prototype.getWs = function () {
		return this.ws;
	};
	/**
	 * Sets a previous value of progression
	 * @memberof CSerial
	 * @param {number} nValue
	 */
	CSerial.prototype.setPrevValue = function (nValue) {
		this.nPrevValue = nValue;
	};
	/**
	 * Returns a previous value of progression
	 * @memberof CSerial
	 * @returns {number}
	 */
	CSerial.prototype.getPrevValue = function () {
		return this.nPrevValue;
	};
	/**
	 * Sets a last index of selected range
	 * @memberof CSerial
	 * @param {number} nIndex
	 */
	CSerial.prototype.setIndex = function (nIndex) {
		this.nIndex = nIndex;
	};
	/**
	 * Returns a last index of selected range
	 * @memberof CSerial
	 * @returns {number}
	 */
	CSerial.prototype.getIndex = function () {
		return this.nIndex;
	};
	/**
	 * Initialize a last index of selected range
	 * @memberof CSerial
	 */
	CSerial.prototype.initIndex = function () {
		let ws = this.getWs();
		let oActiveFillHandle = this.getActiveFillHandle();
		let oRange = ws.selectionRange.getLast();

		if (oActiveFillHandle) {
			this.setIndex(this.getVertical() ? oActiveFillHandle.r2 - oRange.r1 : oActiveFillHandle.c2 - oRange.c1);
		} else {
			this.setIndex(this.getVertical() ? oRange.r2 - oRange.r1 : oRange.c2 - oRange.c1);
		}
	};
	/**
	 * Sets a start point of a range needs to be fill
	 * @memberof CSerial
	 * @param {number} nStartIndex
	 */
	CSerial.prototype.setStartIndex = function (nStartIndex) {
		this.nStartIndex = nStartIndex;
	};
	/**
	 * Returns a start point of a range needs to be fill
	 * @memberof CSerial
	 * @returns {number}
	 */
	CSerial.prototype.getStartIndex = function () {
		return this.nStartIndex;
	};
	/**
	 * Sets a finish point of a range who needs to be fill
	 * @param {number} nEndIndex
	 */
	CSerial.prototype.setEndIndex = function (nEndIndex) {
		this.nEndIndex = nEndIndex;
	};
	/**
	 * Returns a finish point of a range who needs to be fill
	 * @returns {number}
	 */
	CSerial.prototype.getEndIndex = function () {
		return this.nEndIndex;
	};
	/**
	 * Sets a direction step
	 * @param {number} nDirectionStep
	 */
	CSerial.prototype.setDirectionStep = function (nDirectionStep) {
		this.nDirectionStep = nDirectionStep;
	};
	/**
	 * Returns a direction step
	 * @returns {number}
	 */
	CSerial.prototype.getDirectionStep = function () {
		return this.nDirectionStep;
	};
	/**
	 * Initializes iteration coordinates for promote cells in trend mode
	 * @param {Object} oFilledLine - Line of cells which need to fill and first filled cell
	 * @param {number} oFilledLine.nValue - Value of first cell in line
	 * @param {Range} oFilledLine.oToRange - Range of cells which will be fill
	 * @param {Cell} oFilledLine.oCell - First cell of line
	 * @param {Range} oFilledLine.oFilledRange - Range with filled cells
	 * @param {number} oFilledLine.nIndex - Index of shift
	 */
	CSerial.prototype.initTrendIterCoords = function (oFilledLine) {
		const bReverse = this.getIndex() < 0;
		const bActiveFillHandleExists = !!this.getActiveFillHandle();
		const bVertical = this.getVertical();
		const oFrom = this.getFromRange().bbox;
		const oFilled = oFilledLine.oFilledRange.bbox;
		const nIndex = oFilledLine.nIndex;

		if (bActiveFillHandleExists && this.getChosenContextMenuProp() !== Asc.c_oAscFillType.series) {
			this.initToRange(this.getFromRange());
			const oTo = this.getToRange().bbox;
			if (bReverse) {
				this.setStartIndex(bVertical ? oTo.r2 : oTo.c2);
				this.setEndIndex(bVertical ? oTo.r1 - 1 : oTo.c1 - 1);
				this.setDirectionStep(-1);
			} else {
				this.setStartIndex(bVertical ? oTo.r1 : oTo.c1);
				this.setEndIndex(bVertical ? oTo.r2 + 1 : oTo.c2 + 1);
				this.setDirectionStep(1);
			}
		} else {
			if (bReverse) {
				let nEndIndexFilledRange = bVertical ? oFrom.r1 : oFrom.c1;
				this.setStartIndex(bVertical ? oFrom.r2 : oFrom.c2);
				this.setEndIndex((nIndex + nEndIndexFilledRange) - 1);
				this.setDirectionStep(-1);
			} else {
				let nStartIndexFilled = bVertical ? oFilled.r1 : oFilled.c1;
				this.setStartIndex(bVertical ? oFrom.r1 : oFrom.c1);
				this.setEndIndex((nIndex + nStartIndexFilled) + 1);
				this.setDirectionStep(1);
			}
		}
	};
	/**
	 * Creates and returns a range with filled cells
	 * @memberof CSerial
	 * @param {number} [nIndexLine] - Line index of filled cells. Not required param
	 * @returns {Range}
	 */
	CSerial.prototype.getFilledRange = function (nIndexLine) {
		let oFilledRange = this.getFromRange().clone();
		let bVertical = this.getVertical();
		let bTrend = this.getTrend();
		let nRow = oFilledRange.bbox.r1;
		let nCol = oFilledRange.bbox.c1;
		let sFirstValue = null;
		let ws = oFilledRange.worksheet;
		let oBBox = oFilledRange.bbox;
		let nType = this.getType();

		if (nIndexLine != null) {
			if (bVertical) {
				oFilledRange.bbox.c1 = nIndexLine;
				oFilledRange.bbox.c2 = nIndexLine;
			} else {
				oFilledRange.bbox.r1 = nIndexLine;
				oFilledRange.bbox.r2 = nIndexLine;
			}
		}
		if (bTrend) {
			oFilledRange._foreach(function (oCell, nCurRow, nCurCol) {
				if (oCell && oCell.getValueWithoutFormat()) {
					nRow = nCurRow;
					nCol = nCurCol;
					if (nCurRow === oFilledRange.bbox.r1 && nCurCol === oFilledRange.bbox.c1) {
						sFirstValue = oCell.getValueWithoutFormat();
					}
				} else {
					let sNextValue = _getNextValue(ws, oCell, bVertical, oBBox);
					if (!sFirstValue && sNextValue) {
						if (bVertical) {
							oFilledRange.bbox.r1 = nCurRow + 1;
							oFilledRange.bbox.c1 = nCurCol;
						} else {
							oFilledRange.bbox.c1 = nCurCol + 1;
							oFilledRange.bbox.r1 = nCurRow;
						}
					}
				}
			});
		} else {
			oFilledRange._foreachNoEmpty(function (oCell, nCurRow, nCurCol) {
				if (oCell && oCell.getValueWithoutFormat()) {
					if (nType === oSeriesType.autoFill) {
						nRow = nCurRow;
						nCol = nCurCol;
					} else { // for other modes use only first filled cell in range
						nRow = nCurRow;
						nCol = nCurCol;
						return true; // break loop
					}
				}
			});
		}

		// Check if the filled range has merged cells
		if (!bTrend || this.getChosenContextMenuProp() !== Asc.c_oAscFillType.series) {
			let aMergedFilled = oFilledRange.worksheet.mergeManager.get(oFilledRange.bbox).all;
			if (aMergedFilled.length) {
				// Get the last merged range as the last points of the filled range for Growth and Linear trend of context menu
				let oMergedRange = bTrend ? aMergedFilled[aMergedFilled.length - 1] : aMergedFilled[0];
				nRow = oMergedRange.bbox.r2;
				nCol = oMergedRange.bbox.c2;
			}
		}

		oFilledRange.bbox.r2 = nRow;
		oFilledRange.bbox.c2 = nCol;

		return oFilledRange;
	};
	/**
	 * Creates a range to be filled
	 * @memberof CSerial
	 * @param {Range} oFilledRange - Range of filled line of cells.
	 */
	CSerial.prototype.initToRange = function (oFilledRange) {
		let ws = this.getWs();
		let oTo = null;
		let oFrom = this.getActiveFillHandle() ? this.getActiveFillHandle(): this.getFromRange().bbox;
		let oFilled = oFilledRange.bbox;

		this.initIndex();
		let bFilledRangeHasShift = this.getVertical() ? oFilled.r1 !== oFrom.r1 : oFilled.c1 !== oFrom.c1;
		if (this.getIndex() < 0) {
			if (!this.getTrend()) {
				return;
			}
			bFilledRangeHasShift = this.getVertical() ? oFilled.r2 !== oFrom.r1 : oFilled.c2 !== oFrom.c1;
			let oActiveFillHandle = this.getActiveFillHandle();
			oFrom = new Asc.Range(oActiveFillHandle.c2, oActiveFillHandle.r2, oActiveFillHandle.c1, oActiveFillHandle.r1);
		}
		// For a filled range, need to correct the index according to a shift in the start point in the filled range
		if (bFilledRangeHasShift) {
			let nIndex = this.getIndex();
			this.setIndex(this.getVertical() ? nIndex - oFilled.r1 : nIndex - oFilled.c1);
		}
		if (this.getIndex() === 0) {
			let oCanPromote = oFilledRange.canPromote(false, this.getVertical(), 1);
			if (oCanPromote) {
				oTo = oCanPromote.to;
			}
		} else {
			let oCanPromote = oFilledRange.canPromote(false, this.getVertical(), this.getIndex());
			if (oCanPromote) {
				oTo = oCanPromote.to;
			}
		}
		if (oTo) {
			if (!oFrom.containsRange(oTo) && this.getStopValue() == null) {
				return;
			}
			this.setToRange(ws.getRange3(oTo.r1, oTo.c1, oTo.r2, oTo.c2));
		}
	};
	/**
	 * Returns an array of Objects that contains: the value of first cell, the range which will be fill data and object of the cell
	 * @memberof CSerial
	 * @returns {Object[]}
	 */
	CSerial.prototype.getFilledCells = function () {
		let oSerial = this;
		let oFromRange = this.getFromRange();
		let bVertical = this.getVertical();
		let nType = this.getType();
		let aFilledCells = [];

		oFromRange._foreachNoEmpty(function (oCell, nRow, nCol) {
			if (oCell && oCell.getValueWithoutFormat() && nType !== oSeriesType.autoFill) {
				let nTypeCell = oCell.getType();
				let oFilledRange = oSerial.getFilledRange(bVertical ? nCol : nRow);
				let bFirstCellInRange = bVertical ? nRow === oFilledRange.bbox.r1 : nCol === oFilledRange.bbox.c1;
				if (bFirstCellInRange && nTypeCell === CellValueType.Number) {
					oSerial.initToRange(oFilledRange);
					let oToRange = oSerial.getToRange();
					if (!oToRange) {
						return true; // break loop
					} else {
						aFilledCells.push({
							nValue: oCell.getNumberValue(),
							oToRange: oToRange,
							oCell: oCell.duplicate(),
							oFilledRange: oFilledRange,
							nIndex: oSerial.getIndex()
						});
					}
				}
			}
		});

		return aFilledCells;
	};
	/**
	 * Creates a history point for "Series" feature
	 * @memberof CSerial
	 */
	CSerial.prototype.initHistoryPoint = function () {
		let nIndex = this.getIndex();
		let oBBox = this.getFromRange().bbox;
		let nWidth = oBBox.c2 - oBBox.c1 + 1;
		let nHeight = oBBox.r2 - oBBox.r1 + 1;

		AscCommon.History.Create_NewPoint();
		let oSelection = AscCommon.History.GetSelection();
		if (null != oSelection) {
			oSelection = oSelection.clone();
			oSelection.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r2);
			AscCommon.History.SetSelection(oSelection);
		}
		let oSelectionRedo = AscCommon.History.GetSelectionRedo();
		if (null != oSelectionRedo) {
			oSelectionRedo = oSelectionRedo.clone();
			if (this.getVertical()) {
				if (nIndex >= nHeight) {
					oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r1 + nIndex);
				} else {
					oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c2, oBBox.r1 + nIndex - 1);
				}
			} else {
				if (nIndex >= nWidth) {
					oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c1 + nIndex, oBBox.r2);
				} else {
					oSelectionRedo.assign(oBBox.c1, oBBox.r1, oBBox.c1 + nIndex - 1, oBBox.r2);
				}
			}
			AscCommon.History.SetSelectionRedo(oSelectionRedo);
		}
	};

	/**
	 * Rounds a float number for correct calculation.
	 * @param {number} nValue - The value which will be rounded.
	 * @param {number} nFloatNumTemplate - The number template will be used for counting numbers after the decimal point.
	 * @returns {number} - Rounded number
	 * @private
	 */
	function _smartRound(nValue, nFloatNumTemplate) {
		if (!Number.isInteger(nFloatNumTemplate)) {
			let nCountNumsAfterDec = nFloatNumTemplate.toString().split(".")[1].length;
			return nValue.toFixed(nCountNumsAfterDec) - 0;
		}

		return nValue;
	}
	/**
	 * Checks if the next month is February and the day of the current date is more than the last day of February
	 * or isn't the last day of the current month.
	 * Using for fixed problem with February month, when the day of the current date that needs to be increased to Feb
	 * is more than the last day of Feb, the Date function may skip this month.
	 * @param {number} nCurrentExcelDate
	 * @param {number} nStep
	 * @return {{nextMonthIsFeb:boolean, lastDayOfFeb:number}}
	 */
	function checkNextMonthIsFeb (nCurrentExcelDate, nStep) {
		const CURRENT_MONTH = 0;
		const FEB_MONTH  = 1;

		let dtTempCurrentDate = new Asc.cDate().getDateFromExcel(nCurrentExcelDate);
		let nMonthIdOfCurrentDate = dtTempCurrentDate.getMonth();
		let nNextMonthId = (nMonthIdOfCurrentDate + nStep) % 12;
		let nDayOfCurrentDate = dtTempCurrentDate.getDate();
		let nLastDayOfCurrentMonth = AscCommonExcel.getLastDayInMonth(nCurrentExcelDate, CURRENT_MONTH).getDate();
		dtTempCurrentDate.addMonths(nStep) // For recognize correct year for next month
		let nCurrentYear = dtTempCurrentDate.getFullYear();
		let nLastDayOfFeb = AscCommonExcel.getLastDayInMonth(new Asc.cDate(nCurrentYear, FEB_MONTH, 1, 12).getExcelDate(), CURRENT_MONTH).getDate();

		return {
			nextMonthIsFeb: nNextMonthId === FEB_MONTH && nDayOfCurrentDate > nLastDayOfFeb && nDayOfCurrentDate < nLastDayOfCurrentMonth,
			lastDayOfFeb: nLastDayOfFeb
		};
	}

	/**
	 * Calculates date for autofill and Series.
	 * @param {{step:number, dateUnit:c_oAscDateUnitType, previousValue:number, previousIntValue:number, expectedDayValue:number}} oCellInfo
	 * @param {boolean} bAutofill
	 * @returns {{previousValue:number, currentValue:number, previousIntValue:number}}
	 * @private
	 */
	function _calculateDate(oCellInfo, bAutofill) {
		const nStep = oCellInfo.step;
		const nDateUnit = oCellInfo.dateUnit;
		const bDate1904 = AscCommon.bDate1904;
		let nPrevValue = oCellInfo.previousValue;
		let nPrevIntValue = oCellInfo.previousIntValue;
		let dtExpectedDayValue = new Asc.cDate().getDateFromExcel(oCellInfo.expectedDayValue < 59 && !bDate1904 ? oCellInfo.expectedDayValue + 1 : oCellInfo.expectedDayValue);
		let oReturn = {};

		// Condition: nPrevVal < 60 is temporary solution for "01/01/1900 - 01/03/1900" dates
		/* TODO Need make system solution for cDate class for case when excelDate is 1 (01/01/1900).
		    For now if try convert "1" to Date using getDateFromExcel method result is 31/12/1899
		    by this reason of next methods addDays and getExcelDate work incorrect. Result of function is always "-30"
		    If for the "1" value add 1 day and try to use getDateFromExcel method, it returns 01/01/1900 =>
		    Temporary solution for "01/01/1900 - 01/03/1900" dates - always need add 1 day for getDateFromExcel.
		*/
		if (nDateUnit === oSeriesDateUnitType.weekday) {
			const aWeekdays = [1, 2, 3, 4, 5];
			const MAX_LIMIT_STEP = 2147483647;
			let oCurrentValDate = null;
			let nCurrentVal = nPrevValue;
			let i = 0;
			let nIntStep = Math.floor(nStep);
			let nFinalStep = Math.sign(nIntStep) + _smartRound(nStep - nIntStep, nStep);
			do {
				if (nFinalStep !== 0) {
					nCurrentVal = _smartRound(nCurrentVal + nFinalStep, nStep);
				}
				// Convert number to cDate object
				oCurrentValDate = new Asc.cDate().getDateFromExcel(nPrevValue < 59 && !bDate1904 ? nCurrentVal + 1 : nCurrentVal);
				let nDayOfWeek = nPrevValue < 60 ? oCurrentValDate.getDay() - 1 : oCurrentValDate.getDay();
				if (!aWeekdays.includes(nDayOfWeek)) {
					while (true) {
						nCurrentVal += Math.sign(nStep);
						oCurrentValDate = new Asc.cDate().getDateFromExcel(nCurrentVal < 59 && !bDate1904 ? nCurrentVal + 1 : nCurrentVal);
						nDayOfWeek = nCurrentVal < 60 && !bDate1904 ? oCurrentValDate.getDay() - 1 : oCurrentValDate.getDay();
						if (aWeekdays.includes(nDayOfWeek)) {
							i++;
							break;
						}
					}
				} else {
					i++;
				}
			} while (i < Math.abs(nIntStep) && Math.abs(nIntStep) <= MAX_LIMIT_STEP);

			oReturn.previousValue = nCurrentVal;
			oReturn.currentValue = nCurrentVal < 0 ? nCurrentVal : oCurrentValDate.getExcelDate();
			return oReturn;
		}

		let nCurrentVal = _smartRound(nPrevValue + nStep, nStep);
		oReturn.previousValue = nCurrentVal;
		if (nDateUnit === oSeriesDateUnitType.day) {
			if (nStep > -1 && nStep < 0 && nCurrentVal < 0)  {
				oReturn.previousValue = nCurrentVal + 1;
				oReturn.currentValue = nCurrentVal + 1;
				return oReturn;
			}
			oReturn.currentValue = nCurrentVal;
			return oReturn;
		}
		let nIntegerVal = nPrevIntValue;
		let oCurrentValDate = new Asc.cDate().getDateFromExcel(nIntegerVal < 59 && !bDate1904 ? nIntegerVal + 1 : nIntegerVal);
		let nFinalStep = _smartRound(nCurrentVal - nIntegerVal, nStep);
		if (nFinalStep < 0 && !bAutofill) {
			oReturn.currentValue = NaN;
			return oReturn;
		}
		if (nDateUnit === oSeriesDateUnitType.month) {
			let oRes = checkNextMonthIsFeb(nIntegerVal, nFinalStep);
			let bNextMonthIsFeb = oRes.nextMonthIsFeb;
			if(bNextMonthIsFeb) {
				// Change day to the last day of february for don't skip Feb month.
				oCurrentValDate.setUTCDate(oRes.lastDayOfFeb);
			}
			if (Number.isInteger(nFinalStep)) {
				oCurrentValDate.addMonths(nFinalStep);
				let nLastDayOfCurrentMonth = AscCommonExcel.getLastDayInMonth(oCurrentValDate.getExcelDate(), 0).getDate();
				if (!bNextMonthIsFeb &&  nLastDayOfCurrentMonth >= dtExpectedDayValue.getDate() && dtExpectedDayValue.getDate() !== oCurrentValDate.getDate()) {
					oCurrentValDate.setUTCDate(dtExpectedDayValue.getDate());
				}
				oReturn.currentValue = oCurrentValDate.getExcelDate();
				oReturn.previousIntValue = oReturn.currentValue;
				oReturn.previousValue = oReturn.currentValue;
				return oReturn;
			}
			oCurrentValDate.addMonths(nFinalStep);
			let nLastDayOfCurrentMonth = AscCommonExcel.getLastDayInMonth(oCurrentValDate.getExcelDate(), 0).getDate();
			if (!bNextMonthIsFeb && nLastDayOfCurrentMonth >= dtExpectedDayValue.getDate() && dtExpectedDayValue.getDate() !== oCurrentValDate.getDate()) {
				oCurrentValDate.setUTCDate(dtExpectedDayValue.getDate());
			}
			oReturn.currentValue = oCurrentValDate.getExcelDate();
			return oReturn;
		}
		if (nDateUnit === oSeriesDateUnitType.year) {
			if (Number.isInteger(nFinalStep)) {
				oCurrentValDate.addYears(nFinalStep);
				oReturn.currentValue = oCurrentValDate.getExcelDate();
				oReturn.previousIntValue = oReturn.currentValue;
				oReturn.previousValue = oReturn.currentValue;
				return oReturn;
			}
			oCurrentValDate.addYears(nFinalStep);
			oReturn.currentValue = oCurrentValDate.getExcelDate();
			return oReturn;
		}
	}

	/**
	 * Fills current value for Date type
	 * @memberof CSerial
	 * @param {object} oFilledLine - Value of first cell in line
	 * @param {number} oFilledLine.nValue - Value of first cell in line
	 * @param {Range} oFilledLine.oToRange - Range of cells which will be fill
	 * @param {Cell} oFilledLine.oCell - First cell of line
	 * @param {Range} oFilledLine.oFilledRange - Range with filled cells
	 * @param {number} oFilledLine.nIndex - Index of shift
	 * @returns {number} Current value in ExcelDate format
	 * @private
	 */
	CSerial.prototype._fillExcelDate = function (oFilledLine) {
		const oCellInfo = {
			step: this.getStep(),
			dateUnit: this.getDateUnit(),
			previousValue: this.getPrevValue(),
			previousIntValue: oFilledLine.nValue,
			expectedDayValue: oFilledLine.oCell.getNumberValue()
		}

		const oResult = _calculateDate(oCellInfo, false);
		this.setPrevValue(oResult.previousValue);
		if (oResult.previousIntValue) {
			oFilledLine.nValue = oResult.previousIntValue;
		}

		return oResult.currentValue;
	};
	/**
	 * Fills cells in Linear and Growth regression except Trend mode. Works with:
	 * - Type: Linear, Growth, Date.
	 * - Stop value
	 * - Step value
	 * @memberof CSerial
	 * @param {Object} oFilledLine - Line of cells which need to fill and first filled cell
	 * @param {number} oFilledLine.nValue - Value of first cell in line
	 * @param {Range} oFilledLine.oToRange - Range of cells which will be fill
	 * @param {Cell} oFilledLine.oCell - First cell of line
	 * @param {Range} oFilledLine.oFilledRange - Range with filled cells
	 * @param {number} oFilledLine.nIndex - Index of shift
	 */
	CSerial.prototype.promoteCells = function (oFilledLine) {
		function fillCell(nRow, nCol) {
			let bStopLoop = false
			oWsTo._getCell(nRow, nCol, function (oCopyCell) {
				if (oSerial.getPrevValue() != null) {
					let oCellValue = new AscCommonExcel.CCellValue();
					let nCurrentValue = oProgressionCalc[oSerial.getType()]();
					let bNeedStopLoop = false;
					if (nStopValue != null) {
						if (isNaN(nCurrentValue)) {
							bNeedStopLoop = true;
						} else if (oSerial.getType() === oSeriesType.growth) {
							bNeedStopLoop = Number.isInteger(nStep) ? nCurrentValue	> nStopValue : nCurrentValue < nStopValue;
						} else if (oSerial.getDateUnit() === oSeriesDateUnitType.weekday) {
							bNeedStopLoop = nCurrentValue <= 0;
						} else {
							bNeedStopLoop = nStep > 0 ? nCurrentValue > nStopValue : nCurrentValue < nStopValue;
						}
					}
					if (bNeedStopLoop || (nRow === gc_nMaxRow0 || nCol === gc_nMaxCol0)) {
						bStopLoop = true;
						return;
					}
					if (isNaN(nCurrentValue)) {
						let oErrorValue = new cError(cErrorType.not_numeric);
						oCellValue.type = CellValueType.Error
						oCellValue.text = oErrorValue.value;
					} else {
						oCellValue.type = CellValueType.Number;
						oCellValue.number = nCurrentValue;
					}
					oCopyCell.setValueData(new UndoRedoData_CellValueData(null, oCellValue));
					// Add styles from firstCell
					if (null != oFromCell) {
						oCopyCell.setStyle(oFromCell.getStyle());
					}
				}
			});

			return bStopLoop;
		}

		let oSerial = this;
		let nStep = this.getStep();
		let nStopValue = this.getStopValue() != null ? this.getStopValue() : null;
		let nIndexFilledLine = this.getVertical() ? oFilledLine.oCell.nCol : oFilledLine.oCell.nRow;
		let oTo = oFilledLine.oToRange.bbox;
		let oWsTo = oFilledLine.oToRange.worksheet;
		let oFromCell = oFilledLine.oCell;
		let oFilledRange = oFilledLine.oFilledRange;
		let oFrom = oFilledRange.bbox;
		this.setPrevValue(oFilledLine.nValue);
		// Init variables for filling cells
		let nStartIndex = this.getVertical() ? oTo.r1 : oTo.c1;
		let nEndIndex = this.getVertical() ? oTo.r2 : oTo.c2;
		let nDirectStep = 1;
		// Fill range cells for i row or col
		let oProgressionCalc = {
			0: function () { // linear
				let nCurrentVal = _smartRound(oSerial.getPrevValue() + nStep, nStep);
				oSerial.setPrevValue(nCurrentVal);
				return nCurrentVal;
			}, 1: function () { // growth
				let nCurrentVal = oSerial.getPrevValue() * nStep
				oSerial.setPrevValue(nCurrentVal);
				return nCurrentVal;
			}, 2: function () { // date
				return oSerial._fillExcelDate(oFilledLine);
			}
		};
		let bStopLoop = false;
		let bIncorrectStopValue = false;

		if (nStopValue != null) {
			if (this.getType() === oSeriesType.growth) {
				bIncorrectStopValue = (nStep <= 0 || nStep === 1) || nStopValue <= 0;
			} else {
				let nStartVal = oFilledLine.nValue;
				bIncorrectStopValue = nStep < 0 ? nStartVal <= nStopValue : nStartVal >= nStopValue;
			}
		}
		if (bIncorrectStopValue) {
			return;
		}
		if (oFilledRange.hasMerged()) {
			nDirectStep = this.getVertical() ? oFrom.r2 - oFrom.r1 + 1 : oFrom.c2 - oFrom.c1 + 1;
		}
		if (nStopValue != null && nStartIndex === nEndIndex) {
			let nIndex = nStartIndex;
			do {
				let nRow = this.getVertical() ? nIndex : nIndexFilledLine;
				let nCol = this.getVertical() ? nIndexFilledLine : nIndex;
				bStopLoop = fillCell(nRow, nCol);
				nIndex += nDirectStep;
			} while (!bStopLoop);
		} else {
			for (let j = nStartIndex; j <= nEndIndex; j += nDirectStep) {
				let nRow = this.getVertical() ? j : nIndexFilledLine;
				let nCol = this.getVertical() ? nIndexFilledLine : j;
				bStopLoop = fillCell(nRow, nCol);
				if (bStopLoop) {
					break;
				}
			}
		}
	};
	/**
	 * Gets value and index filled cell in loop for calculate sum in fAction. Uses for linear regression formula
	 * in promoteCellsWithTrend method. Trend mode
	 * @memberof CSerial
	 * @param {number} nStartIndex - Index for start loop
	 * @param {number} nEndIndex - Index for finish loop
	 * @param {number} nFilledCellIter - Iteration of filled cell. Uses for define nRow or nCol at getting cell
	 * @param fAction - Callback function for calculate sum with parameters: nValue, nIndexCell
	 */
	CSerial.prototype._calcSum = function (nStartIndex, nEndIndex, nFilledCellIter, fAction) {
		if (fAction) {
			let oSerial = this;
			let oWsFrom = this.getWs();
			let oRes = null;
			for (let i = nStartIndex; i < nEndIndex; i++) {
				let nRow = this.getVertical() ? i : nFilledCellIter;
				let nCol = this.getVertical() ? nFilledCellIter : i;
				oWsFrom._getCell(nRow, nCol, function (oCell) {
					let nValue = oCell.getNumberValue();
					if (nValue !== null) {
						// For Growth type take ln from value
						if (oSerial.getType() === oSeriesType.growth) {
							nValue = Math.log(nValue);
						}
						oRes = fAction(nValue, i);
					}
				});
				if (oRes != null) {
					return oRes;
				}
			}
		}
	};
	/**
	 * Fills cells in Trend mode. For fill cells Trend mode uses linear regression formula - y = a + bx, when:
	 * a - intercept,
	 * b - slop,
	 * x - index of cell who must be fill.
	 * Works only for Linear and Growth types and without "Step value" and "Stop value" data.
	 * @memberof CSerial
	 * @param {Object} oFilledLine - Line of filled cells and cells which need to fill
	 * @param {number} oFilledLine.nValue - Value of first cell in line
	 * @param {Range} oFilledLine.oToRange - Range of cells which will be fill
	 * @param {Cell} oFilledLine.oCell - First cell of line
	 * @param {Range} oFilledLine.oFilledRange - Range with filled cells
	 * @param {number} oFilledLine.nIndex - Index of shift
	 * @param {number} nFilledLineLength - Length line of filled cells
	 */
	CSerial.prototype.promoteCellsWithTrend = function (oFilledLine, nFilledLineLength) {
		this.initTrendIterCoords(oFilledLine);
		let oSerial = this;
		let oFilledRange = oFilledLine.oFilledRange;
		let nIndexFilledLine = this.getVertical() ? oFilledLine.oCell.nCol : oFilledLine.oCell.nRow;
		let nFilledStartIndex = this.getVertical() ? oFilledRange.bbox.r1 : oFilledRange.bbox.c1;
		// If the first cell after calculating the linear regression for "growth trend" is Infinity, then the next cells in range fill 0.
		// This flag is only uses for the Growth trend with an active fill handle
		let bFirstCellValueInf = false;
		// Define variables for calculate linear regression
		let nSumX = 0;
		let nSumY = 0;
		let nNumeratorOfSlope = 0;
		let nDenominatorOfSlope = 0;
		let nSlope = null;
		// Calculate nEndIndex for Filled line with taking into account the shift first element in the filled range
		let nEndIndexFilledLine = nFilledLineLength + nFilledStartIndex;
		// Getting filled cells and calculating sum of index cell (x) and values (y) for calculate intercept and slop
		// for linear regression formula: y = intercept + slop * x
		let nFilledRangeLength = 0;
		this._calcSum(nFilledStartIndex, nEndIndexFilledLine, nIndexFilledLine, function (nValue, nCellIndex) {
			nFilledRangeLength++;
			if ((nValue === -Infinity || isNaN(nValue)) && nFilledStartIndex === nCellIndex && oSerial.getActiveFillHandle()) {
				bFirstCellValueInf = true;
				return true; // break loop
			}
			nSumX += nCellIndex;
			nSumY += nValue;
		});
		let nXAvg = nSumX / nFilledRangeLength;
		let nYAvg = nSumY / nFilledRangeLength;
		if (nFilledLineLength > 1 && !bFirstCellValueInf) {
			this._calcSum(nFilledStartIndex, nEndIndexFilledLine, nIndexFilledLine, function (nValue, nCellIndex) {
				nNumeratorOfSlope += (nCellIndex - nXAvg) * (nValue - nYAvg);
				nDenominatorOfSlope += Math.pow((nCellIndex - nXAvg), 2);
			});
			nSlope = nNumeratorOfSlope / nDenominatorOfSlope;
		} else {
			nSlope = this.getType() === oSeriesType.growth ? 0 : 1;
		}
		let nIntercept = nYAvg - nSlope * nXAvg;
		if (isNaN(nSlope) || isNaN(nIntercept)) {
			return;
		}
		// Fill cells in selection range
		let nStartIndex = this.getStartIndex();
		let nEndIndex = this.getEndIndex();
		let nDirectionStep = this.getDirectionStep();
		let oFromCell = oFilledLine.oCell;
		let oWsTo = oFilledLine.oToRange.worksheet;
		for (let x = nStartIndex; x !== nEndIndex; x += nDirectionStep) {
			let nRow = this.getVertical() ? x : nIndexFilledLine;
			let nCol = this.getVertical() ? nIndexFilledLine : x;
			oWsTo._getCell(nRow, nCol, function (oCopyCell) {
				let oCellValue = new AscCommonExcel.CCellValue();
				let nCellValue = nIntercept + nSlope * x;
				if (oSerial.getType() === oSeriesType.growth) {
					nCellValue = bFirstCellValueInf ? nCellValue : Math.exp(nCellValue);
				}
				oCellValue.type = CellValueType.Number;
				oCellValue.number = nCellValue;
				oCopyCell.setValueData(new UndoRedoData_CellValueData(null, oCellValue));
				// Add styles from firstCell
				if (null != oFromCell) {
					oCopyCell.setStyle(oFromCell.getStyle());
				}
			});
		}
	};

	/**
	 * Gets length of filled line
	 * @param {Range} oToRange - Range of cells which will be fill
	 * @param {Range} oFilledRange - Range with filled cells
	 * @param {boolean} bVertical - Flag recognizes is vertical or horizontal direction
	 * @param {number} nIndex - Index of cell who must be fill
	 * @returns {number}
	 * @private
	 */
	function _getFilledLineLength(oToRange, oFilledRange, bVertical, nIndex) {
		if (nIndex > 0) {
			return bVertical ? oToRange.bbox.r1 - oFilledRange.bbox.r1 : oToRange.bbox.c1 - oFilledRange.bbox.c1;
		}
		// Reverse direction
		return bVertical ? oFilledRange.bbox.r2 - oToRange.bbox.r2 : oFilledRange.bbox.c2 - oToRange.bbox.c2;
	}

	/**
	 * Creates merged cells
	 * @param {Range} oFromRange - Filled range.
	 * @param {Range} oToRange - The range that needs to be filled.
	 * @private
	 */
	function _mergeCells(oFromRange, oToRange) {
		const oFrom = oFromRange.bbox;
		const oTo = oToRange.bbox;
		const oWsFrom = oFromRange.worksheet;
		const oWsTo = oToRange.worksheet;
		const oMergedFrom = oWsFrom.mergeManager.get(oFrom);

		let nDx = oFrom.c2 - oFrom.c1 + 1;
		let nDy = oFrom.r2 - oFrom.r1 + 1;

		if(oMergedFrom && oMergedFrom.all.length > 0) {
			for (let i = oTo.c1; i <= oTo.c2; i += nDx) {
				for (let j = oTo.r1; j <= oTo.r2; j += nDy) {
					for (let k = 0, length = oMergedFrom.all.length; k < length; k++) {
						const oMergedBBox = oMergedFrom.all[k].bbox;
						let nC1 = i + oMergedBBox.c1 - oFrom.c1;
						let nR1 = j + oMergedBBox.r1 - oFrom.r1;
						let nC2 = i + oMergedBBox.c2 - oFrom.c1;
						let nR2 = j + oMergedBBox.r2 - oFrom.r1;
						const oNewMerged = new Asc.Range(nC1, nR1, nC2, nR2);
						if(oTo.contains(oNewMerged.c1, oNewMerged.r1)) {
							if(oTo.c2 < oNewMerged.c2) {
								oNewMerged.c2 = oTo.c2;
							}
							if(oTo.r2 < oNewMerged.r2) {
								oNewMerged.r2 = oTo.r2;
							}
							if(!oNewMerged.isOneCell()) {
								oWsTo.mergeManager.add(oNewMerged, 1);
							}
						}
					}
				}
			}
		}
	}
	/**
	 * Main method runs "Series" feature according to the specified parameters
	 * @memberof CSerial
	 */
	CSerial.prototype.exec = function () {
		let bVertical = this.getVertical();
		let nType = this.getType();
		let oFromRange = this.getFromRange();
		let oWs = this.getWs();
		let oMergedFrom = oWs.mergeManager.get(oFromRange.bbox);

		if (nType === oSeriesType.autoFill) {
			let oFilledRange = this.getFilledRange();
			let oFilled = oFilledRange.bbox;
			this.initIndex();
			let oCanPromote = oFilledRange.canPromote(false, bVertical, this.getIndex());
			if (!oFilled.containsRange(oCanPromote.to)) {
				oFilledRange.promote(false, this.getVertical(), this.getIndex(), oCanPromote);
			}
		} else {
			let aFilledCells = this.getFilledCells();
			if (aFilledCells.length) {
				this.initHistoryPoint();
				AscCommon.History.StartTransaction();
				for (let i = 0, length = aFilledCells.length; i < length; i++) {
					let oToRange = aFilledCells[i].oToRange;
					let oFilledRange = aFilledCells[i].oFilledRange;
					let nFilledLineLength = _getFilledLineLength(oToRange, oFilledRange, bVertical, this.getIndex());
					if (this.getTrend()) {
						this.promoteCellsWithTrend(aFilledCells[i], nFilledLineLength);
					} else {
						this.promoteCells(aFilledCells[i]);
					}
					_mergeCells(oFilledRange, oToRange);
				}
				AscCommon.History.EndTransaction();
			} else if (oMergedFrom.all.length) {
				let oApi = Asc.editor;
				oApi.sendEvent("asc_onError", c_oAscError.ID.CannotFillRange, c_oAscError.Level.NoCritical);
			}
		}
	};

	function CExternalReferenceHelper(wb) {
		this.wb = wb;
	}

	CExternalReferenceHelper.prototype.getExternalLinkStr = function (nExternalLinkIndex, locale, isShortLink) {
		let index = nExternalLinkIndex;

		let sameFile;
		let fileName = window["Asc"]["editor"].DocInfo && window["Asc"]["editor"].DocInfo.get_Title();
		if ((fileName != null && index === fileName) || index == "0") {
			sameFile = true;
		}

		let wbModel = this.wb;
		let oExternalLink = nExternalLinkIndex && wbModel && wbModel.getExternalLinkByIndex(index - 1, true);

		if (oExternalLink && !locale) {
			return "[" + index + "]";
		}
		let path = oExternalLink && oExternalLink.path;
		let name = oExternalLink && oExternalLink.name;
		let res = "";
		if (path || name) {
			if (path) {
				res += path;
			}
			if (name) {
				res += isShortLink ? name : "[" + name + "]";
			}
		} else if (oExternalLink) {
			res = oExternalLink;
		} else if (sameFile) {
			res += locale ? fileName : "[0]";
		}
		return res;
	};

	CExternalReferenceHelper.prototype.check3dRef = function (_3DRefTmp, local) {
		let t = this;
		let externalLink = _3DRefTmp[3];
		let externalDefName, externalSheetName, receivedDefName, receivedLink, isShortLink, isCurrentFile;

		// this argument contain shortLink object with full formula and two parts of it
		let receivedShortLink = _3DRefTmp[4];
		if (receivedShortLink) {
			// receivedShortLink - the received, short formula that needs to be checked for an internal link (they have the same in structure)
			// the link can be either to another sheet or to another book - they have the same entry
			receivedLink = receivedShortLink.externalLink;
			externalSheetName = receivedShortLink.externalLink;
			receivedDefName = receivedShortLink.defname;
			isCurrentFile = receivedShortLink.currentFile;
		}

		// This check of short links is performed only when opening/reading, manual input is processed differently
		if (receivedShortLink && !local) {
			let eReference = t.wb.getExternalLinkByIndex(externalLink - 1);
			if (eReference && eReference.DefinedNames) {
				for (let i = 0; i < eReference.DefinedNames.length; i++) {
					if (eReference.DefinedNames[i].Name === receivedDefName) {
						externalDefName = eReference.DefinedNames[i];
						if (externalDefName.SheetId !== null) {
							externalSheetName = eReference.SheetNames[eReference.DefinedNames[i].SheetId];
						} 
						else if (!externalDefName.SheetId && externalDefName.RefersTo) {
							// RefersTo differs from the wsName and may contain several exclamation marks,
							// so we use a condition and a regular expression to get the correct name

							let refString = externalDefName.RefersTo;
							let exclamationMarkIndex = refString.lastIndexOf("!");

							refString = refString.slice(0, exclamationMarkIndex);
							refString = refString[0] === "=" ? refString.substring(1) : refString;
							externalSheetName = refString;

							// regex to find string enclosed in single qoutes
							let regex = /^'(.*)'$/;
							let match = regex.exec(refString);
							if (match && match[1]) {
								externalSheetName = match[1];
							} else if (refString) {
								let externalWB = eReference.getWb();
								let depFormulas = externalWB && externalWB.dependencyFormulas;
								if (depFormulas && depFormulas.defNames && depFormulas.defNames.wb && depFormulas.defNames.wb[receivedDefName]) {
									externalSheetName = refString;
								}
							}
						}
						break;
					}
				}
			}
		}

		let sheetName = _3DRefTmp[1] ? _3DRefTmp[1] : externalSheetName;
		if (!sheetName && local) {
			return false;
		}

		let externalName = _3DRefTmp[3];
		//check on add to this document
		let thisTitle = externalLink && window["Asc"]["editor"] && window["Asc"]["editor"].DocInfo && window["Asc"]["editor"].DocInfo.get_Title();
		if (thisTitle === externalLink) {
			externalLink = null;
		}

		// we check whether sheetName is part of the document or is it a short link to external data
		if (local && !externalLink && receivedShortLink) {
			// if there is a sheet with the same name, we link to it, otherwise we create an external link (shortened)
			let innerSheet = t.wb.getWorksheetByName(sheetName);
			if (!innerSheet) {
				let eReference = t.wb.getExternalLinkByName(sheetName);
				if (eReference && eReference.DefinedNames) {
					for (let i = 0; i < eReference.DefinedNames.length; i++) {
						if (eReference.DefinedNames[i].Name === receivedDefName) {
							externalDefName = eReference.DefinedNames[i];
							if (externalDefName.SheetId !== null) {
								externalSheetName = eReference.SheetNames[eReference.DefinedNames[i].SheetId];
							} else if (!externalDefName.SheetId && externalDefName.RefersTo) {
								// RefersTo differs from the wsName and may contain several exclamation marks,
								// so we use a condition and a regular expression to get the correct name
	
								let refString = externalDefName.RefersTo;
								let exclamationMarkIndex = refString.lastIndexOf("!");
	
								refString = refString.slice(0, exclamationMarkIndex);
								refString = refString[0] === "=" ? refString.substring(1) : refString;
								externalSheetName = refString;
	
								// regex to find string enclosed in single qoutes
								let regex = /^'(.*)'$/;
								let match = regex.exec(refString);
								if (match && match[1]) {
									externalSheetName = match[1];
								}
							}
							break;
						}
					}
				}

				sheetName = externalSheetName ? externalSheetName : receivedLink.split(".")[0];
				externalName = receivedLink;
				externalLink = receivedLink;
				isShortLink = true;
			}
		}

		return {
			sheetName: sheetName, 
			externalLink: externalLink, 
			receivedLink: receivedLink, 
			externalName: externalName, 
			isShortLink: isCurrentFile ? true : isShortLink,
			isCurrentFile: isCurrentFile,
			currentFileDefname: isCurrentFile ? receivedDefName : null
		};
	};

	// Export
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};
	window['AscCommonExcel'].g_nVerticalTextAngle = g_nVerticalTextAngle;
	window['AscCommonExcel'].g_nSheetNameMaxLength = g_nSheetNameMaxLength;
	window['AscCommonExcel'].oDefaultMetrics = oDefaultMetrics;
	window['AscCommonExcel'].g_nAllColIndex = g_nAllColIndex;
	window['AscCommonExcel'].g_nAllRowIndex = g_nAllRowIndex;
	window['AscCommonExcel'].g_DefNameWorksheet = null;
	window['AscCommonExcel'].aStandartNumFormats = aStandartNumFormats;
	window['AscCommonExcel'].aStandartNumFormatsId = aStandartNumFormatsId;
	window['AscCommonExcel'].oFormulaLocaleInfo = oFormulaLocaleInfo;
	window['AscCommonExcel'].getDefNameIndex = getDefNameIndex;
	window['AscCommonExcel'].getCellIndex = getCellIndex;
	window['AscCommonExcel'].getFromCellIndex = getFromCellIndex;
	window['AscCommonExcel'].angleFormatToInterface2 = angleFormatToInterface2;
	window['AscCommonExcel'].angleInterfaceToFormat = angleInterfaceToFormat;
	window['AscCommonExcel'].Workbook = Workbook;
	window['AscCommonExcel'].Worksheet = Worksheet;
	window['AscCommonExcel'].Cell = Cell;
	window['AscCommonExcel'].Range = Range;
	window['AscCommonExcel'].DefName = DefName;
	window['AscCommonExcel'].DependencyGraph = DependencyGraph;
	window['AscCommonExcel'].HiddenManager = HiddenManager;
	window['AscCommonExcel'].CCellWithFormula = CCellWithFormula;
	window['AscCommonExcel'].preparePromoteFromTo = preparePromoteFromTo;
	window['AscCommonExcel'].promoteFromTo = promoteFromTo;
	window['AscCommonExcel'].getCompiledStyle = getCompiledStyle;
	window['AscCommonExcel'].getCompiledStyleFromArray = getCompiledStyleFromArray;
	window['AscCommonExcel'].ignoreFirstRowSort = ignoreFirstRowSort;
	window['AscCommonExcel'].tryTranslateToPrintArea = tryTranslateToPrintArea;
	window['AscCommonExcel']._isSameSizeMerged = _isSameSizeMerged;
	window['AscCommonExcel'].g_nDefNameMaxLength = g_nDefNameMaxLength;
	window['AscCommonExcel'].changeTextCase = changeTextCase;
	window['AscCommonExcel'].g_sNewSheetNamePattern = g_sNewSheetNamePattern;
	window['AscCommonExcel'].CSerial = CSerial;
	window['AscCommonExcel'].SweepLineRowIterator = SweepLineRowIterator;
	window['AscCommonExcel'].BroadcastHelper = BroadcastHelper;
	window['AscCommonExcel'].foreachRefElements = foreachRefElements;

})(window);
