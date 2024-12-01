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
(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
function (window, undefined) {

	/*
	 * Import
	 * -----------------------------------------------------------------------------
	 */
	let asc = window["Asc"];

	let asc_Range = asc.Range;
	let cElementType = AscCommonExcel.cElementType;

	function TraceDependentsManager(ws) {
		this.ws = ws;
		this.precedents = null;
		this.precedentsExternal = null;
		this.dependentsExternal = null;
		this.dependents = null;
		this.isDependetsCall = null;
		this.inLoop = null;
		this.isPrecedentsCall = null;
		this.precedentsAreas = null;
		this.precedentsAreasHeaders = null;
		this.data = {
			referenceMaxLevel: Math.pow(100, 10),
			lastHeaderIndex: -1,
			prevIndex: -1,
			recLevel: 0,
			maxRecLevel: 0,
			indices: {
				// cellIndex[From]: cellIndex[To]: recLevel
			}
		};
		this.currentCalculatedPrecedentAreas = {
			// rangeName: {
			// inProgress: null,
			// isCalculated: null
			// }
		};
		this.aPassedPrecedents = null;
		this.aPassedDependents = null;

		this._lockChangeDocument = null;

		/* an array with the coordinates of the start and end of all drawn lines */
		this.tracesCoords = null;
	}

	TraceDependentsManager.prototype.setPrecedentsCall = function () {
		this.isPrecedentsCall = true;
		this.isDependetsCall = false;
	};
	TraceDependentsManager.prototype.setDependentsCall = function () {
		this.isDependetsCall = true;
		this.isPrecedentsCall = false;
	};
	TraceDependentsManager.prototype.setPrecedentExternal = function (from, to, elemRange, elemWs, externalLink) {
		if (!this.precedentsExternal) {
			this.precedentsExternal = {};
		}

		if (!this.precedentsExternal[from]) {
			this.precedentsExternal[from] = {};
		}

		let docInfo = window["Asc"]["editor"].DocInfo;
		let externalInfo = {range: elemRange, ws: elemWs, fullPath: null, isCurrentWorkbook: null};

		let rangeName = elemRange.getName();
		let wsName = elemWs.getName();
		let isCurrentWorkbook = true;
		let eR, fileName;

		if (externalLink != null) {
			isCurrentWorkbook = false;
			eR = this.ws && this.ws.workbook && this.ws.workbook.model && this.ws.workbook.model.getExternalLinkByIndex(externalLink - 1);
			fileName = eR ? eR.Id : "";
		}

		fileName = isCurrentWorkbook ? docInfo && docInfo.get_Title() : fileName;
		// in fullPath we write a line with full information about the name of the book, sheet and range
		let fullPath = "[" + fileName + "]" + wsName + "!" + rangeName;

		externalInfo.isCurrentWorkbook = isCurrentWorkbook;
		externalInfo.fullPath = fullPath;

		this.precedentsExternal[from][to] = externalInfo;
	};
	TraceDependentsManager.prototype.checkPrecedentExternal = function (cellIndex) {
		if (!this.precedentsExternal) {
			return false;
		}
		return this.precedentsExternal[cellIndex];
	};
	// dependentsExternal
	TraceDependentsManager.prototype.setDependentsExternal = function (from, to, elemRange, elemWs) {
		if (!this.dependentsExternal) {
			this.dependentsExternal = {};
		}

		if (!this.dependentsExternal[from]) {
			this.dependentsExternal[from] = {};
		}

		let docInfo = window["Asc"]["editor"].DocInfo;
		let externalInfo = {range: elemRange, ws: elemWs, fullPath: null};

		let rangeName = elemRange.getName();
		let wsName = elemWs.getName();
		let fileName = docInfo ? docInfo.get_Title() : "";
		// in fullPath we write a line with complete information about the name of the book, sheet and range
		let fullPath = "[" + fileName + "]" + wsName + "!" + rangeName;

		externalInfo.fullPath = fullPath;

		this.dependentsExternal[from][to] = externalInfo;
	};
	TraceDependentsManager.prototype.checkDependentExternal = function (cellIndex) {
		if (!this.dependentsExternal) {
			return false;
		}
		return this.dependentsExternal[cellIndex];
	};

	TraceDependentsManager.prototype.checkCircularReference = function (cellIndex, isDependentCall) {
		if (this.dependents && this.dependents[cellIndex] && this.precedents && this.precedents[cellIndex]) {
			if (isDependentCall) {
				for (let i in this.dependents[cellIndex]) {
					if (this._getDependents(i, cellIndex) && this._getDependents(cellIndex, i)) {
						return true;
					}
				}
			} else {
				for (let i in this.precedents) {
					if (this._getPrecedents(i, cellIndex) && this._getPrecedents(cellIndex, i)) {
						return true;
					}
				}
			}
		}
	};
	TraceDependentsManager.prototype.clearLastDependent = function (row, col) {
		let ws = this.ws && this.ws.model;
		if (!ws || !this.dependents) {
			return;
		}
		if (Object.keys(this.dependents).length === 0) {
			return;
		}

		const t = this;
		if (row == null || col == null) {
			let selection = ws.getSelection();
			let activeCell = selection.activeCell;
			row = activeCell.row;
			col = activeCell.col;
			let mergedRange = ws.getMergedByCell(row, col);
			if (mergedRange) {
				row = mergedRange.r1;
				col = mergedRange.c1;
			}
		}

		const findMaxNesting = function (row, col) {
			let currentCellIndex = AscCommonExcel.getCellIndex(row, col);

			if (t.data.prevIndex !== -1 && t.data.indices[t.data.prevIndex] && t.data.indices[t.data.prevIndex][currentCellIndex]) {
				return;
			}
			if (t.dependents[currentCellIndex]) {
				if (checkCircularReference(currentCellIndex)) {
					return;
				}

				let interLevel, fork;
				if (Object.keys(t.dependents[currentCellIndex]).length > 1) {
					fork = true;
				}

				t.data.recLevel++;
				t.data.maxRecLevel = t.data.recLevel > t.data.maxRecLevel ? t.data.recLevel : t.data.maxRecLevel;
				interLevel = t.data.recLevel;
				for (let j in t.dependents[currentCellIndex]) {
					t.data.prevIndex = currentCellIndex;
					if (j.includes(";")) {
						// [fromCurrent][toExternal]
						if (!t.data.indices[currentCellIndex]) {
							t.data.indices[currentCellIndex] = {};
						}
						t.data.indices[currentCellIndex][j] = t.data.recLevel;
						continue;
					}
					let coords = AscCommonExcel.getFromCellIndex(j, true);
					findMaxNesting(coords.row, coords.col);
					t.data.recLevel = fork ? interLevel : t.data.recLevel;
				}
			} else {
				if (!t.data.indices[t.data.prevIndex]) {
					t.data.indices[t.data.prevIndex] = {};
				}
				t.data.indices[t.data.prevIndex][currentCellIndex] = t.data.recLevel;	// [from][to]
			}
		};
		const checkCircularReference = function (index) {
			for (let i in t.dependents[index]) {
				if (t._getDependents(index, i) && t._getDependents(i, index)) {
					let related = index + "|" + i;
					t.data.recLevel = t.data.referenceMaxLevel;
					t.data.maxRecLevel = t.data.recLevel;
					t.data.indices[related] = t.data.recLevel;
					return true;
				}
			}
		};

		findMaxNesting(row, col);
		const maxLevel = this.data.maxRecLevel;
		if (maxLevel === 0) {
			this._setDefaultData();
			return;
		} else if (maxLevel === t.data.referenceMaxLevel) {
			// TODO improve check of cyclic references
			// temporary solution: now, when finding cyclic dependencies, the maximum nesting number(100^10) is set for them and only they are will be initially removed
			for (let i in this.data.indices) {
				if (this.data.indices[i] === maxLevel) {
					let val = i.split("|");
					this._deleteDependent(val[0], val[1]);
					this._deletePrecedent(val[0], val[1]);
					this._deleteDependent(val[1], val[0]);
					this._deletePrecedent(val[1], val[0]);
				}
			}
		}

		for (let index in this.data.indices) {
			for (let i in this.data.indices[index]) {
				if (this.data.indices[index][i] === maxLevel) {
					this._deletePrecedent(i, index);
					this._deleteDependent(index, i);
				}
			}
		}

		this._setDefaultData();
	};
	TraceDependentsManager.prototype.calculateDependents = function (row, col) {
		let ws = this.ws && this.ws.model;
		if (!ws) {
			return;
		}
		if (row == null || col == null) {
			let selection = ws.getSelection();
			let activeCell = selection.activeCell;
			row = activeCell.row;
			col = activeCell.col;
			let mergedRange = ws.getMergedByCell(row, col);
			if (mergedRange) {
				row = mergedRange.r1;
				col = mergedRange.c1;
			}
		}

		let depFormulas = ws.workbook.dependencyFormulas;
		if (depFormulas && depFormulas.sheetListeners) {
			if (!this.dependents) {
				this.dependents = {};
			}

			let sheetListeners = depFormulas.sheetListeners;
			let curListener = sheetListeners[ws.Id];
			let cellIndex = AscCommonExcel.getCellIndex(row, col);
			this._calculateDependents(cellIndex, curListener);
			this.setDependentsCall();
		}
	};
	TraceDependentsManager.prototype._calculateDependents = function (cellIndex, curListener, isSecondCall) {
		let t = this;
		let ws = this.ws.model;
		let wb = this.ws.model.workbook;
		let dependencyFormulas = wb.dependencyFormulas;
		let allDefNamesListeners = dependencyFormulas.defNameListeners;
		let cellAddress = AscCommonExcel.getFromCellIndex(cellIndex, true);

		const findCellListeners = function () {
			const listeners = {};
			if (curListener && curListener.areaMap) {
				for (let j in curListener.areaMap) {
					if (curListener.areaMap.hasOwnProperty(j)) {
						if (curListener.areaMap[j] && curListener.areaMap[j].bbox.contains(cellAddress.col, cellAddress.row)) {
							Object.assign(listeners, curListener.areaMap[j].listeners);
						}
					}
				}
			}
			if (curListener && curListener.cellMap && curListener.cellMap[cellIndex]) {
				if (Object.keys(curListener.cellMap[cellIndex]).length > 0) {
					Object.assign(listeners, curListener.cellMap[cellIndex].listeners);
				}
			}
			if (curListener && curListener.defName3d) {
				Object.assign(listeners, curListener.defName3d);
			}
			return listeners;
		};
		const checkIfHeader = function (tableHeader) {
			if (!tableHeader) {
				return false;
			}

			return tableHeader.col === cellAddress.col && tableHeader.row === cellAddress.row;
		};
		const getTableHeader = function (table) {
			if (!table.Ref) {
				return false;
			}

			return {col: table.Ref.c1, row: table.Ref.r1};
		};
		const setDefNameIndexes = function (defName, isTable, defNameRange) {
			let tableHeader = isTable ? getTableHeader(ws.getTableByName(defName)) : false;
			let isCurrentCellHeader = isTable ? checkIfHeader(tableHeader) : false;
			for (let i in allDefNamesListeners) {
				if (allDefNamesListeners.hasOwnProperty(i) && i.toLowerCase() === defName.toLowerCase()) {
					for (let listener in allDefNamesListeners[i].listeners) {
						// TODO maybe add all listeners in 'curListener' at once
						// listener can be: cell, range, table, named range - there will be unique processing for each case
						let elem = allDefNamesListeners[i].listeners[listener];
						let isArea = elem.ref ? true : false;
						let is3D = elem.ws.Id ? elem.ws.Id !== ws.Id : false;
						let isIntersect;

						if (isArea && !is3D && !isCurrentCellHeader) {
							if (defNameRange) {
								let defBBox = defNameRange.getBBox0();
								// check clicked cell for entry into dependent areas
								// if the cell is not included, then the dependency will not be drawn
								let colShift = defBBox.c1 - elem.ref.c1,
									rowShift = defBBox.r1 - elem.ref.r1;

								isIntersect = elem.ref.contains(cellAddress.col - colShift, cellAddress.row - rowShift);
							}

							if (isIntersect) {
								// decompose all elements into dependencies
								let areaIndexes = getAllAreaIndexes(elem);
								if (areaIndexes) {
									for (let index in areaIndexes) {
										if (areaIndexes.hasOwnProperty(index)) {
											t._setDependents(cellIndex, areaIndexes[index]);
											t._setPrecedents(areaIndexes[index], cellIndex);
										}
									}
								}
							}
							continue;
						}

						let parentCellIndex = getParentIndex(elem.parent);
						if (parentCellIndex === null) {
							continue;
						}

						if (isTable) {
							// check Headers
							// if current header and listener is header, make trace only with header
							// check if current cell header or not
							if (elem.Formula.includes("Headers")) {
								if (isCurrentCellHeader) {
									t._setDependents(cellIndex, parentCellIndex);
									t._setPrecedents(parentCellIndex, cellIndex);
								} else {
									continue;
								}
							} else if (!elem.Formula.includes("Headers") && isCurrentCellHeader) {
								continue;
							}
							// ?additional check if the listener is in the same table, need to check if it is a listener of the main cell
							if (elem.outStack) {
								let arr = [];
								// check each element of the stack for an occurrence in the original cell
								for (let table in elem.outStack) {
									if (elem.outStack[table].type !== cElementType.table) {
										continue;
									}

									let bbox = elem.outStack[table].area.bbox ? elem.outStack[table].area.bbox : (elem.outStack[table].area.range.bbox ? elem.outStack[table].area.range.bbox : null);

									if (bbox) {
										arr.push(bbox.contains2(cellAddress));
									}
								}
								if (!arr.includes(true)) {
									continue;
								}
							}

							// shared checks
							if (elem.shared !== null && !is3D) {
								let currentCellRange = ws.getCell3(cellAddress.row, cellAddress.col);
								setSharedTableIntersection(ws.getTableByName(defName).getRangeWithoutHeaderFooter(), currentCellRange, elem.shared);
								continue;
							}
						}
						t._setDependents(cellIndex, parentCellIndex);
						t._setPrecedents(parentCellIndex, cellIndex);
					}
				}
			}
		};
		const getAllAreaIndexes = function (parserFormula) {
			const indexes = [], range = parserFormula.ref;
			if (!range) {
				return;
			}
			for (let c = range.c1; c <= range.c2; c++) {
				for (let r = range.r1; r <= range.r2; r++) {
					let index = AscCommonExcel.getCellIndex(r, c);
					indexes.push(index);
				}
			}

			return indexes;
		};
		const getParentIndex = function (_parent) {
			if (!_parent || _parent.nCol == null || _parent.nRow == null) {
				return null;
			}
			let _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, _parent.nCol);
			//parent -> cell
			if (_parent.ws !== t.ws.model) {
				_parentCellIndex += ";" + _parent.ws.index;
			}
			return _parentCellIndex;
		};
		const setSharedIntersection = function (currentRange, shared) {
			// get the cell is contained in one of the areaMap
			// if contain, call getSharedIntersect with currentRange whom contain cell and sharedRange
			if (curListener && curListener.areaMap) {
				for (let j in curListener.areaMap) {
					if (curListener.areaMap.hasOwnProperty(j)) {
						if (curListener.areaMap[j] && curListener.areaMap[j].bbox.contains(cellAddress.col, cellAddress.row)) {
							let isNotSharedRange;
							for (let listener in curListener.areaMap[j].listeners) {
								if (curListener.areaMap[j].listeners[listener].shared === null) {
									isNotSharedRange = true;
								}
								break;
							}
							let res = isNotSharedRange ? null : curListener.areaMap[j].bbox.getSharedIntersect(shared.ref, currentRange.bbox);
							// draw dependents to coords from res
							if (res) {
								let index = AscCommonExcel.getCellIndex(res.r1, res.c1);
								if (res.r1 === res.r2 && res.c1 !== res.c2) {
									index = res.containsCol(currentRange.bbox.c1) ? AscCommonExcel.getCellIndex(res.r1, currentRange.bbox.c1) : AscCommonExcel.getCellIndex(res.r1, res.c1);
								} else if (res.c1 === res.c2 && res.r1 !== res.r2) {
									index = res.containsRow(currentRange.bbox.r1) ? AscCommonExcel.getCellIndex(currentRange.bbox.r1, res.c1) : AscCommonExcel.getCellIndex(res.r1, res.c1);
								}
								t._setDependents(cellIndex, index);
								t._setPrecedents(index, cellIndex);
							}
						}
					}
				}
			}
		};
		const setSharedTableIntersection = function (currentRange, currentCellRange, shared) {
			// row mode || col mode
			let isRowMode = currentRange.r1 === currentRange.r2,
				isColumnMode = currentRange.c1 === currentRange.c2, res, tempRange;

			if (isColumnMode && currentRange.r2 > shared.ref.r2) {
				if (!shared.ref.containsRow(currentCellRange.bbox.r2)) {
					return
				}
				if (currentCellRange.r2 > shared.ref.r2) {
					return;
				}
				// do check with rest of the currentRange
				tempRange = new asc_Range(currentRange.c1, currentRange.r1, currentRange.c2, shared.ref.r2);
			} else if (isRowMode && currentRange.c2 > shared.ref.c2) {
				// contains
				if (!shared.ref.containsCol(currentCellRange.bbox.c2)) {
					return
				}
				if (currentCellRange.c2 > shared.ref.c2) {
					return;
				}
				tempRange = new asc_Range(currentRange.c1, currentRange.r1, shared.ref.c2, currentRange.r2);
			}

			if (tempRange) {
				res = tempRange.getSharedIntersect(shared.ref, currentCellRange.bbox);
			}

			res = !res ? currentRange.getSharedIntersect(shared.ref, currentCellRange.bbox) : res;

			if (res && (res.r1 === res.r2 && res.c1 === res.c2)) {
				let index = AscCommonExcel.getCellIndex(res.r1, res.c1);
				t._setDependents(cellIndex, index);
				t._setPrecedents(index, cellIndex);
			} else {
				// split shared range on two parts
				let split = currentRange.difference(shared.ref);

				if (split.length > 1) {
					// first part
					res = currentRange.getSharedIntersect(split[0], currentCellRange.bbox);
					if (res && (res.r1 === res.r2 && res.c1 === res.c2)) {
						let index = AscCommonExcel.getCellIndex(res.r1, res.c1);
						t._setDependents(cellIndex, index);
						t._setPrecedents(index, cellIndex);
					}

					// second part
					if (split[1]) {
						let range = split[1], indexes = [];
						for (let col = range.c1; col <= range.c2; col++) {
							for (let row = range.r1; row <= range.r2; row++) {
								let index = AscCommonExcel.getCellIndex(row, col);
								indexes.push(index);
							}
						}
						if (indexes.length > 0) {
							for (let index in indexes) {
								if (indexes.hasOwnProperty(index)) {
									t._setDependents(cellIndex, indexes[index]);
									t._setPrecedents(indexes[index], cellIndex);
								}
							}
						}
					}
				}
			}
		};

		const cellListeners = findCellListeners();
		if (cellListeners && Object.keys(cellListeners).length > 0) {
			if (!this.dependents) {
				this.dependents = {};
			}
			if (!this.dependents[cellIndex]) {
				// if dependents by cellIndex didn't exist, create it
				this.dependents[cellIndex] = {};
				let parentCellIndex = null;
				for (let i in cellListeners) {
					if (cellListeners.hasOwnProperty(i)) {
						let parent = cellListeners[i].parent;
						let parentWsId = parent.ws ? parent.ws.Id : null;
						let isTable = parent.parsedRef ? parent.parsedRef.isTable : false;
						let isDefName = !!parent.name;
						let formula = cellListeners[i].Formula;
						let is3D = false;

						//todo slow operation. parent not have type
						if (parent instanceof Asc.CT_WorksheetSource) {
							// if the listener is a pivot table, skip the iteration
							continue;
						}

						if (isDefName) {
							let parentInnerElementType = parent.parsedRef.outStack[0] ? parent.parsedRef.outStack[0].type : false,
								defNameRange;
							if (parentInnerElementType === cElementType.cellsRange || parentInnerElementType === cElementType.cellsRange3D || parentInnerElementType === cElementType.cell3D) {
								defNameRange = parent.parsedRef.outStack[0].getRange();
							}
							// TODO check external table ref
							setDefNameIndexes(parent.name, isTable, defNameRange);
							continue;
						} else if (cellListeners[i].is3D) {
							is3D = true;
						}

						if (cellListeners[i].shared !== null && !is3D) {
							// can be shared ref in otheer sheet
							let shared = cellListeners[i].getShared();
							let currentCellRange = ws.getCell3(cellAddress.row, cellAddress.col);
							setSharedIntersection(currentCellRange, shared);
							continue;
						}

						if (formula.includes(":") && !is3D) {
							// call splitAreaListeners which return cellIndexes of each element(this will be parentCellIndex)
							// go through the values and set dependents for each
							let areaIndexes = getAllAreaIndexes(cellListeners[i]);
							if (areaIndexes) {
								for (let index in areaIndexes) {
									if (areaIndexes.hasOwnProperty(index)) {
										this._setDependents(cellIndex, areaIndexes[index]);
										this._setPrecedents(areaIndexes[index], cellIndex);
									}
								}
								continue;
							}
						}

						parentCellIndex = getParentIndex(parent);
						if (parentCellIndex === null) {
							//if (parentCellIndex === null || (typeof(parentCellIndex) === "number" && isNaN(parentCellIndex))) {
							continue;
						}

						if (is3D && typeof parentCellIndex === "string") {
							// the object dependentsExternal is only needed to handle a double click on an arrow, and is not used in drawing
							let elemRange = cellListeners[i].ref ? cellListeners[i].ref : new Asc.Range(parent.col, parent.row, parent.col, parent.row);
							let elemWs = parent.ws;
							this.setDependentsExternal(cellIndex, parentCellIndex, elemRange, elemWs);
						}
						this._setDependents(cellIndex, parentCellIndex);
						this._setPrecedents(parentCellIndex, cellIndex, true);
					}
				}
				if (Object.keys(this.dependents[cellIndex]).length === 0 && cellIndex !== parentCellIndex) {
					delete this.dependents[cellIndex];
					this.ws.workbook.handlers.trigger("asc_onError", c_oAscError.ID.TraceDependentsNoFormulas, c_oAscError.Level.NoCritical);
				}
			} else {
				if (this.checkCircularReference(cellIndex, true)) {
					return;
				}
				if (this.checkPassedDependents(cellIndex)) {
					return;
				}
				// if dependents by cellIndex aldready exist, check current tree
				let currentIndex = Object.keys(this.dependents[cellIndex])[0];
				let isUpdated = false;
				let bCellHasNotTrace = false;
				if (Object.keys(this.dependents[cellIndex]).length === 0) {
					bCellHasNotTrace = true;
				}
				for (let i in cellListeners) {
					if (cellListeners.hasOwnProperty(i)) {
						let parent = cellListeners[i].parent;
						//todo slow operation. parent not have type
						if (parent instanceof AscCommonExcel.DefName || parent instanceof Asc.CT_WorksheetSource) {
							// if the listener is a pivot table, skip the iteration
							continue;
						}

						let elemCellIndex = cellListeners[i].shared !== null ? currentIndex : getParentIndex(parent),
							formula = cellListeners[i].Formula;
						if (formula.includes(":") && !cellListeners[i].is3D) {
							// call getAllAreaIndexes which return cellIndexes of each element(this will be parentCellIndex)
							let areaIndexes = getAllAreaIndexes(cellListeners[i]);
							if (areaIndexes) {
								// go through the values and set dependents for each
								for (let index in areaIndexes) {
									if (areaIndexes.hasOwnProperty(index)) {
										this._setDependents(cellIndex, areaIndexes[index]);
										this._setPrecedents(areaIndexes[index], cellIndex, true);
									}
								}
								continue;
							}
						}

						// if the child cell does not yet have a dependency with listeners, create it
						if (!this._getDependents(cellIndex, elemCellIndex) && cellIndex !== elemCellIndex) {
							this._setDependents(cellIndex, elemCellIndex);
							this._setPrecedents(elemCellIndex, cellIndex, true);
							isUpdated = true;
						}
					}
				}

				if (!isUpdated) {
					this.setPassedDependents(cellIndex);
					for (let i in this.dependents[cellIndex]) {
						if (this.dependents[cellIndex].hasOwnProperty(i)) {
							this._calculateDependents(+i, curListener, true);
						}
					}
					if (!isSecondCall) {
						this.clearPassedDependents();
					}
				}
				if (Object.keys(this.dependents[cellIndex]).length === 0 && bCellHasNotTrace) {
					this.ws.workbook.handlers.trigger("asc_onError", c_oAscError.ID.TraceDependentsNoFormulas, c_oAscError.Level.NoCritical);
				}
			}
		} else if (!isSecondCall) {
			this.ws.workbook.handlers.trigger("asc_onError", c_oAscError.ID.TraceDependentsNoFormulas, c_oAscError.Level.NoCritical);
		}
	};
	TraceDependentsManager.prototype._getDependents = function (from, to) {
		return this.dependents[from] && this.dependents[from][to];
	};
	TraceDependentsManager.prototype._setDependents = function (from, to) {
		if (!this.dependents) {
			this.dependents = {};
		}
		if (!this.dependents[from]) {
			this.dependents[from] = {};
		}
		if (from === to) {
			return;
		}
		this.dependents[from][to] = 1;
	};
	TraceDependentsManager.prototype._setDefaultData = function () {
		this.data = {
			referenceMaxLevel: Math.pow(100, 10),
			recLevel: 0,
			maxRecLevel: 0,
			lastHeaderIndex: -1,
			prevIndex: -1,
			indices: {}
		};
	};
	TraceDependentsManager.prototype.clearLastPrecedent = function (row, col) {
		let ws = this.ws && this.ws.model;
		if (!ws || !this.precedents) {
			return;
		}
		if (Object.keys(this.precedents).length === 0) {
			return;
		}

		const t = this;
		if (row == null || col == null) {
			let selection = ws.getSelection();
			let activeCell = selection.activeCell;
			row = activeCell.row;
			col = activeCell.col;
			let mergedRange = ws.getMergedByCell(row, col);
			if (mergedRange) {
				row = mergedRange.r1;
				col = mergedRange.c1;
			}
		}

		const checkCircularReference = function (index) {
			for (let i in t.precedents[index]) {
				if (t._getPrecedents(index, i) && t._getPrecedents(i, index)) {
					let related = index + "|" + i;
					t.data.recLevel = t.data.referenceMaxLevel;
					t.data.maxRecLevel = t.data.recLevel;
					t.data.indices[related] = t.data.recLevel;
					return true;
				}
			}
		};
		const checkIfHeader = function (cellIndex) {
			if (!t.precedentsAreas || !t.precedentsAreasHeaders) {
				return;
			}
			if (t.precedentsAreasHeaders[cellIndex]) {
				return true;
			}
		};
		const getAllAreaIndexes = function (areas, areaName) {
			const indexes = [];
			if (!areas) {
				return;
			}

			let area = areas[areaName];
			for (let r = area.range.r1; r <= area.range.r2; r++) {
				for (let c = area.range.c1; c <= area.range.c2; c++) {
					let index = AscCommonExcel.getCellIndex(r, c);
					indexes.push(index);
				}
			}
			return indexes;
		};
		const findMaxNesting = function (row, col, callFromArea) {
			let currentCellIndex = AscCommonExcel.getCellIndex(row, col);
			if (t.data.indices[currentCellIndex] && t.data.indices[currentCellIndex][t.data.prevIndex]) {
				t.data.indices[currentCellIndex][t.data.prevIndex] = t.data.recLevel;
				return;
			}

			let ifHeader, interLevel, fork;
			if (t.data.recLevel > 0 && t.data.lastHeaderIndex !== currentCellIndex) {
				// checking if a cell is a table header
				ifHeader = callFromArea ? false : checkIfHeader(currentCellIndex);

				if (!t.precedents[currentCellIndex] && !ifHeader) {
					if (!t.data.indices[t.data.prevIndex]) {
						t.data.indices[t.data.prevIndex] = {};
					}
					t.data.indices[t.data.prevIndex][currentCellIndex] = t.data.recLevel;	// [from][to] format 
					return;
				}
			}

			if (ifHeader) {
				// go through area
				let areaName = t.precedentsAreasHeaders[currentCellIndex];
				let areaIndexes = getAllAreaIndexes(t.precedentsAreas, areaName);

				if (areaIndexes.length > 0) {
					fork = true;
					interLevel = t.data.recLevel;
					for (let index in areaIndexes) {
						if (areaIndexes.hasOwnProperty(index)) {
							let _index = areaIndexes[index];
							let cellAddress = AscCommonExcel.getFromCellIndex(_index, true);
							if (_index === currentCellIndex) {
								t.data.lastHeaderIndex = _index;
							}
							if (!t.precedents[_index] && _index !== currentCellIndex) {
								continue;
							}
							findMaxNesting(cellAddress.row, cellAddress.col, true);
							t.data.recLevel = fork ? interLevel : t.data.recLevel;
						}
					}
				}
			} else if (t.precedents[currentCellIndex]) {
				if (checkCircularReference(currentCellIndex)) {
					return;
				}
				if (Object.keys(t.precedents[currentCellIndex]).length > 1) {
					fork = true;
				}

				t.data.recLevel++;
				t.data.maxRecLevel = t.data.recLevel > t.data.maxRecLevel ? t.data.recLevel : t.data.maxRecLevel;
				interLevel = t.data.recLevel;
				for (let j in t.precedents[currentCellIndex]) {
					t.data.prevIndex = currentCellIndex;
					if (j.includes(";")) {
						// [fromCurrent][toExternal]
						if (!t.data.indices[currentCellIndex]) {
							t.data.indices[currentCellIndex] = {};
						}
						t.data.indices[currentCellIndex][j] = t.data.recLevel;
						continue;
					}
					let coords = AscCommonExcel.getFromCellIndex(j, true);
					findMaxNesting(coords.row, coords.col);
					t.data.recLevel = fork ? interLevel : t.data.recLevel;
				}
			} else {
				if (!t.data.indices[t.data.prevIndex]) {
					t.data.indices[t.data.prevIndex] = {};
				}
				// [from][to]
				t.data.indices[t.data.prevIndex][currentCellIndex] = t.data.recLevel;
			}
		};
		findMaxNesting(row, col);
		const maxLevel = this.data.maxRecLevel;
		if (maxLevel === 0) {
			this._setDefaultData();
			return;
		}
		// TODO improve check of cyclic references
		// temporary solution: now, when finding cyclic dependencies, the maximum nesting number(100^10) is set for them and only they are will be initially removed
		else if (maxLevel === t.data.referenceMaxLevel) {
			for (let i in this.data.indices) {
				if (this.data.indices[i] === maxLevel) {
					let val = i.split("|");
					this._deleteDependent(val[0], val[1]);
					this._deletePrecedent(val[0], val[1]);
					this._deleteDependent(val[1], val[0]);
					this._deletePrecedent(val[1], val[0]);
				}
			}
		}

		for (let index in this.data.indices) {
			for (let i in this.data.indices[index]) {
				if (this.data.indices[index][i] === maxLevel) {
					this._deletePrecedent(index, i);
					this._deleteDependent(i, index);
				}
			}
		}
		this.checkAreas();
		this._setDefaultData();
	};
	TraceDependentsManager.prototype.calculatePrecedents = function (row, col, isSecondCall, callFromArea) {
		//depend from row/col cell
		let ws = this.ws && this.ws.model;
		if (!ws) {
			return;
		}
		const t = this;
		if (row == null || col == null) {
			// if first call, create/clear object with calculated areas 
			this.currentCalculatedPrecedentAreas = {};

			let selection = ws.getSelection();
			let activeCell = selection.activeCell;
			row = activeCell.row;
			col = activeCell.col;
			let mergedRange = ws.getMergedByCell(row, col);
			if (mergedRange) {
				row = mergedRange.r1;
				col = mergedRange.c1;
			}
		}

		const cellIndex = AscCommonExcel.getCellIndex(row, col);
		const getAllAreaIndexesWithFormula = function (areas, areaName) {
			const indexes = [];
			if (!areas) {
				return;
			}

			let area = areas[areaName];
			for (let i = area.range.r1; i <= area.range.r2; i++) {
				for (let j = area.range.c1; j <= area.range.c2; j++) {
					// ??? check parserFormula and return indexes only with it
					if (!ws.getCell3(i, j).isFormula()) {
						continue;
					}
					let index = AscCommonExcel.getCellIndex(i, j);
					indexes.push(index);
				}
			}

			area.isCalculated = true;
			return indexes;
		};
		const isCellAreaHeader = function (cellIndex) {
			if (!t.precedentsAreas || !t.precedentsAreasHeaders) {
				return;
			}
			if (t.precedentsAreasHeaders[cellIndex]) {
				return true;
			}
		};

		let formulaParsed;
		ws.getCell3(row, col)._foreachNoEmpty(function (cell) {
			formulaParsed = cell.formulaParsed;
		});

		// TODO another way to check table
		let isAreaHeader = callFromArea ? false : isCellAreaHeader(cellIndex);
		if (this.precedentsAreas && isSecondCall && isAreaHeader) {
			// calculate all precedents in areas
			let areaName = this.precedentsAreasHeaders[cellIndex];
			let areaIndexes = getAllAreaIndexesWithFormula(this.precedentsAreas, areaName);

			if (!this.currentCalculatedPrecedentAreas[areaName]) {
				this.currentCalculatedPrecedentAreas[areaName] = {};
				// go through the values and check precedents for each
				for (let index in areaIndexes) {
					if (areaIndexes.hasOwnProperty(index)) {
						let cellAddress = AscCommonExcel.getFromCellIndex(areaIndexes[index], true);
						this.calculatePrecedents(cellAddress.row, cellAddress.col, true, true);
					}
				}
			}
		} else if (formulaParsed) {
			this._calculatePrecedents(formulaParsed, row, col, isSecondCall);
			this.setPrecedentsCall();
		} else if (!isSecondCall) {
			this.ws.workbook.handlers.trigger("asc_onError", c_oAscError.ID.TracePrecedentsNoValidReference, c_oAscError.Level.NoCritical);
		}
	};
	TraceDependentsManager.prototype._calculatePrecedents = function (formulaParsed, row, col, isSecondCall) {
		if (!this.precedents) {
			this.precedents = {};
		}
		let t = this;
		let currentCellIndex = AscCommonExcel.getCellIndex(row, col);
		let formulaInfoObject = this.checkUnrecordedAndFormNewStack(currentCellIndex, formulaParsed), isHaveUnrecorded,
			newOutStack;
		let bCellHasNotTrace = false;

		if (this.precedents[currentCellIndex] && Object.keys(this.precedents[currentCellIndex]).length === 0) {
			bCellHasNotTrace = true;
		}
		if (formulaInfoObject) {
			isHaveUnrecorded = formulaInfoObject.isHaveUnrecorded;
			newOutStack = formulaInfoObject.newOutStack;
		}

		if (isHaveUnrecorded) {
			let shared, base;
			if (formulaParsed.shared !== null) {
				shared = formulaParsed.getShared();
				base = shared.base;		// base index - where shared formula start
			}

			if (newOutStack) {
				let currentWsIndex = formulaParsed.ws.index;
				let ref = formulaParsed.ref;
				// iterate through the elements and find all reference
				for (let index in newOutStack) {
					if (!newOutStack.hasOwnProperty(index)) {
						continue;
					}
					let elem = newOutStack[index].element;
					let elemType = elem.type ? elem.type : null;

					let is3D = elemType === cElementType.cell3D || elemType === cElementType.cellsRange3D || elemType === cElementType.name3D,
						isArea = elemType === cElementType.cellsRange || elemType === cElementType.name,
						isDefName = elemType === cElementType.name || elemType === cElementType.name3D,
						isTable = elemType === cElementType.table, areaName;
					let bPinCell = false;
					let sValue = null;

					if (elemType === cElementType.cell || is3D || isArea || isDefName || isTable) {
						let cellRange = new asc_Range(col, row, col, row), elemRange, elemCellIndex;
						if (isDefName) {
							let elemDefName = elem.getDefName();
							let elemValue = elem.getValue();
							if (!elemDefName) {
								continue
							}
							let defNameParentWsIndex = elemDefName.parsedRef.outStack[0].wsFrom ? elemDefName.parsedRef.outStack[0].wsFrom.index : (elemDefName.parsedRef.outStack[0].ws ? elemDefName.parsedRef.outStack[0].ws.index : null);
							elemRange = elemValue.range.bbox ? elemValue.range.bbox : elemValue.bbox;

							if (defNameParentWsIndex && defNameParentWsIndex !== currentWsIndex) {
								// 3D
								is3D = true;
								isArea = false;
							} else if (elemRange.isOneCell()) {
								isArea = false;
							}
							sValue = elemValue.value;
							bPinCell = sValue.includes("$");
						} else if (isTable) {
							let currentWsId = elem.ws.Id,
								elemWsId = elem.area.ws ? elem.area.ws.Id : elem.area.wsFrom.Id;
							// elem.area can be cRef and cArea
							is3D = currentWsId !== elemWsId;
							elemRange = elem.area.bbox ? elem.area.bbox : (elem.area.range ? elem.area.range.bbox : null);
							isArea = ref ? true : !elemRange.isOneCell();
						} else {
							sValue = elem.value;
							bPinCell = sValue.includes("$");
							elemRange = elem.range.bbox ? elem.range.bbox : elem.bbox;
						}

						if (!elemRange) {
							continue;
						}

						if (shared) {
							if (isTable) {
								let isRowMode = shared.ref.r1 === shared.ref.r2,
									isColumnMode = shared.ref.c1 === shared.ref.c2,
									diff = [];

								if ((isRowMode && (cellRange.c2 > elemRange.c2)) || (isColumnMode && (cellRange.r2 > elemRange.r2))) {
									// regular link to main table
									elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);
								} else {
									diff = elemRange.difference(shared.ref);
									if (diff.length > 0) {
										let res = diff[0].getSharedIntersect(elemRange, cellRange);
										if (res && (res.r1 === res.r2 && res.c1 === res.c2)) {
											elemCellIndex = AscCommonExcel.getCellIndex(res.r1, res.c1);
										} else {
											elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1 + (row - base.nRow), elemRange.c1 + (col - base.nCol));
										}
									}
								}
							} else if (bPinCell) {
								const FIRST_INDEX_VALUE = 0;
								let nIndex = sValue.indexOf("$");
								if (nIndex === FIRST_INDEX_VALUE) {
									sValue = sValue.slice(nIndex + 1);
									let bStaticCell = sValue.includes("$");
									if (bStaticCell) {
										elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);
									} else {
										elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1 + (row - base.nRow), elemRange.c1);
									}
								} else {
									elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1 + (col - base.nCol));
								}
							} else {
								elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1 + (row - base.nRow), elemRange.c1 + (col - base.nCol));
							}
						} else {
							elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);
						}

						// cross check for element:
						// if the element is Area and it doesn't have reference to another sheet, 
						// and the element is not in the function, 
						// and the formula is not CSE - we are checking for cross
						if (isArea && !ref && !is3D && !newOutStack[index].inFormulaRef) {
							if (elemRange.getWidth() > 1 && elemRange.getHeight() <= 1) {
								// check cols
								if (elemRange.containsCol(col)) {
									elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, col);
									isArea = false;
								}
							} else if (elemRange.getWidth() <= 1 && elemRange.getHeight() > 1) {
								// check rows
								if (elemRange.containsRow(row)) {
									elemCellIndex = AscCommonExcel.getCellIndex(row, elemRange.c1);
									isArea = false;
								}
							} else {
								isArea = true;
							}
						}

						// if the area is on the same sheet - write to the array of areas for drawing
						if (isArea && !is3D) {
							let copyRange = elemRange.clone();
							if (shared && !isTable) {
								const offset = {
									row: row - base.nRow,
									col: col - base.nCol
								};
								// set offset according to base shift
								copyRange.setOffset(offset);
							}
							const areaRange = {};
							areaName = copyRange.getName();			// areaName - unique key for areaRange
							areaRange[areaName] = {};
							areaRange[areaName].range = copyRange;
							areaRange[areaName].isCalculated = null;
							areaRange[areaName].areaHeader = elemCellIndex;

							this._setPrecedentsAreas(areaRange);
						}

						if (is3D) {
							// external dependencies are stored in a separate object and rendered in a separate loop 
							let elemIndex = elem.wsTo ? elem.wsTo.index : elem.ws.index;
							let elemWs = elem.wsFrom ? elem.wsTo : elem.ws;			
							if (currentWsIndex !== elemIndex) {
								if (elem.externalLink != null) {
									elemIndex = "[" + elem.externalLink + "]";
								}

								elemCellIndex += ";" + elemIndex;
								this.setPrecedentExternal(currentCellIndex, elemCellIndex, elemRange, elemWs, elem.externalLink);
								continue
							}
							this._setDependents(elemCellIndex, currentCellIndex);
							this._setPrecedents(currentCellIndex, elemCellIndex);
						} else {
							this._setPrecedents(currentCellIndex, elemCellIndex, false, areaName ? areaName : false);
							this._setDependents(elemCellIndex, currentCellIndex);
							if (areaName) {
								this._setPrecedentsAreaHeader(elemCellIndex, areaName);
							}
						}
					}
				}
			}
		} else {
			if (this.checkCircularReference(currentCellIndex, false)) {
				return;
			}
			if (this.checkPassedPrecedents(currentCellIndex)) {
				return;
			}
			this.setPrecedentsLoop(true);
			this.setPassedPrecedents(currentCellIndex);
			let isHavePrecedents = false;
			// check first level, then if function return false, check second, third and so on
			for (let i in this.precedents[currentCellIndex]) {
				let coords = AscCommonExcel.getFromCellIndex(i, true);
				this.calculatePrecedents(coords.row, coords.col, true);
				isHavePrecedents = true;
			}

			if (!isHavePrecedents && !isSecondCall) {
				this.ws.workbook.handlers.trigger("asc_onError", c_oAscError.ID.TracePrecedentsNoValidReference, c_oAscError.Level.NoCritical);
			}

			this.setPrecedentsLoop(false);
			if (!isSecondCall) {
				this.clearPassedPrecedents();
			}
		}
		if (this.precedents && this.precedents[currentCellIndex] && Object.keys(this.precedents[currentCellIndex]).length === 0 && bCellHasNotTrace) {
			this.ws.workbook.handlers.trigger("asc_onError", c_oAscError.ID.TracePrecedentsNoValidReference, c_oAscError.Level.NoCritical);
		}
	};
	TraceDependentsManager.prototype.addLineCoordinates = function (from, to, /*fromXY, toXY*/x1, y1, x2, y2) {
		/*	
			from - cellIndex
			to - cellIndex
			fromXY - from coordinates
			toXY - to coordinates

			we are working with the precedents object, where
			[from] - this is the position of the line with the arrow pointing to the cell,
			[to] - this is the position of the line with a dot at the beginning
		*/

		if (from === undefined || to === undefined) {
			return
		}

		let traceLineInfo = {from : {x: x1, y: y1, cellRange: null, areaRange: null}, to: {x: x2, y: y2, cellRange: null, areaRange: null}};

		let fromRowCol = AscCommonExcel.getFromCellIndex(from, true);
		let toRowCol = AscCommonExcel.getFromCellIndex(to, true);

		let toAreInfo = this.precedents[from] && this.precedents[from][to];
		let fromAreaInfo = this.precedents[to] && this.precedents[to][from];

		traceLineInfo.from.cellRange = new Asc.Range(fromRowCol.col, fromRowCol.row, fromRowCol.col, fromRowCol.row);
		traceLineInfo.to.cellRange = new Asc.Range(toRowCol.col, toRowCol.row, toRowCol.col, toRowCol.row);

		// areaCheck
		if (this.precedentsAreas) {
			if (toAreInfo && toAreInfo !== 1 && this.precedentsAreas[toAreInfo]) {
				traceLineInfo.to.areaRange = this.precedentsAreas[toAreInfo].range;
			}
			if (fromAreaInfo && fromAreaInfo !== 1 && this.precedentsAreas[fromAreaInfo]) {
				traceLineInfo.from.areaRange = this.precedentsAreas[fromAreaInfo].range;
			}
		}

		if (!this.tracesCoords) {
			this.tracesCoords = [];
		}
		this.tracesCoords.push(traceLineInfo);
	};
	TraceDependentsManager.prototype.addExternalLineCoordinates = function (from, x1, y1, x2, y2) {
		/*	
			from - cellIndex
			we are working with the precedentsExternal and dependentsExternal object, where
			[from] - this is the position of the line with an arrow indicating external dependencies
		*/

		if (from === undefined) {
			return
		}

		// in external we pass an array of strings like "[BookName.xlsx]SheetName!$A$1:$A$3" which we should pass to the goto window
		let traceLineInfo = {from : {x: x1, y: y1}, to: {x: x2, y: y2}, external: null};

		// we go through all external dependencies and fill array of external dependencies for chosen line
		if (this.precedentsExternal && this.precedentsExternal[from]) {
			for (let i in this.precedentsExternal[from]) {
				let externalInfo = this.precedentsExternal[from][i];

				if (!traceLineInfo.external) {
					traceLineInfo.external = [];
				}

				traceLineInfo.external.push(externalInfo);
			}
		}
		if (this.dependentsExternal && this.dependentsExternal[from]) {
			for (let i in this.dependentsExternal[from]) {
				let externalInfo = this.dependentsExternal[from][i];

				if (!traceLineInfo.external) {
					traceLineInfo.external = [];
				}

				traceLineInfo.external.push(externalInfo);
			}
		}

		if (!this.tracesCoords) {
			this.tracesCoords = [];
		}
		this.tracesCoords.push(traceLineInfo);
	};
	TraceDependentsManager.prototype.checkUnrecordedAndFormNewStack = function (cellIndex, formulaParsed) {
		let newOutStack = [], isHaveUnrecorded;
		if (formulaParsed && formulaParsed.outStack) {
			let currentWsIndex = formulaParsed.ws.index,
				ref = formulaParsed.ref,
				coords = AscCommonExcel.getFromCellIndex(cellIndex, true),
				row = coords.row, col = coords.col, shared, base,
				length = formulaParsed.outStack.length, isPartOfFunc, numberOfArgs, funcReturnType, funcArrayIndexes;

			if (formulaParsed.shared !== null) {
				shared = formulaParsed.getShared();
				base = shared.base;
			}

			for (let i = length - 1; i >= 0; i--) {
				let elem = formulaParsed.outStack[i];
				if (!elem) {
					continue;
				}
				if (numberOfArgs <= 0) {
					funcArrayIndexes = null;
					isPartOfFunc = null;
				}

				let elemTypeExist = elem.type !== undefined;
				let elemType = elem.type, inFormulaRef;
				if (isPartOfFunc && numberOfArgs > 0 && elemTypeExist) {
					if (cElementType.cellsRange === elemType || cElementType.name === elemType) {
						if (funcReturnType === AscCommonExcel.cReturnFormulaType.array) {
							// range refers to formula, add property inFormulaRef = true
							inFormulaRef = true;
						} else if (funcArrayIndexes) {
							// if have no returnType check for arrayIndexes and if element pass in raw form(as array, range) to argument
							if (funcArrayIndexes[numberOfArgs - 1]) {
								inFormulaRef = true;
							}
						}
					}
					numberOfArgs--;
				}
				if (elemType === cElementType.func) {
					isPartOfFunc = true;
					numberOfArgs = formulaParsed.outStack[i - 1];
					funcReturnType = elem.returnValueType;
					funcArrayIndexes = elem.arrayIndexes;
				}

				let is3D = elemType === cElementType.cell3D || elemType === cElementType.cellsRange3D || elemType === cElementType.name3D,
					isArea = elemType === cElementType.cellsRange || elemType === cElementType.name,
					isDefName = elemType === cElementType.name || elemType === cElementType.name3D,
					isTable = elemType === cElementType.table;

				if (elemType === cElementType.cell || isArea || is3D || isDefName || isTable) {
					// in any case, add the element to the array
					newOutStack.push({element: elem, inFormulaRef: inFormulaRef});

					// if already know about unrealized entries, skip all checks
					if (isHaveUnrecorded) {
						continue
					}

					let elemRange, elemCellIndex;
					if (isDefName) {
						let elemDefName = elem.getDefName(),
							elemValue = elem.getValue();
						if (!elemDefName) {
							continue
						}

						let defNameParentWsIndex = elemDefName.parsedRef.outStack[0].wsFrom ? elemDefName.parsedRef.outStack[0].wsFrom.index : (elemDefName.parsedRef.outStack[0].ws ? elemDefName.parsedRef.outStack[0].ws.index : null);
						elemRange = elemValue.range.bbox ? elemValue.range.bbox : elemValue.bbox;

						if (defNameParentWsIndex && defNameParentWsIndex !== currentWsIndex) {
							is3D = true;
							isArea = false;
						} else if (elemRange.isOneCell()) {
							isArea = false;
						}
					} else if (isTable) {
						let currentWsId = elem.ws.Id,
							elemWsId = elem.area.ws ? elem.area.ws.Id : elem.area.wsFrom.Id;
						is3D = currentWsId !== elemWsId;
						elemRange = elem.area.bbox ? elem.area.bbox : (elem.area.range ? elem.area.range.bbox : null);
					} else {
						elemRange = elem.range.bbox ? elem.range.bbox : elem.bbox;
					}

					if (!elemRange) {
						continue;
					}

					if (shared) {
						elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1 + (row - base.nRow), elemRange.c1 + (col - base.nCol));
					} else {
						elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);
					}

					// cross check for cell
					if (isArea && !ref && !is3D) {
						if (elemRange.getWidth() > 1 && elemRange.getHeight() <= 1) {
							// check cols
							if (elemRange.containsCol(col)) {
								elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, col);
								isArea = false;
							}
						} else if (elemRange.getWidth() <= 1 && elemRange.getHeight() > 1) {
							// check rows
							if (elemRange.containsRow(row)) {
								elemCellIndex = AscCommonExcel.getCellIndex(row, elemRange.c1);
								isArea = false;
							}
						} else {
							isArea = true;
						}
					}

					if (is3D) {
						// elemCellIndex += ";" + (elem.wsTo ? elem.wsTo.index : elem.ws.index);

						let elemIndex = elem.wsTo ? elem.wsTo.index : elem.ws.index;
						let elemWs = elem.wsFrom ? elem.wsTo : elem.ws;
						if (currentWsIndex !== elemIndex) {
							let hasExternalPrecedent = this.checkPrecedentExternal(cellIndex);
							elemCellIndex += ";" + elemIndex;
							if (!hasExternalPrecedent || (hasExternalPrecedent && !this.precedentsExternal[cellIndex][elemCellIndex])) {
								isHaveUnrecorded = true;
							}
							// this.setPrecedentExternal(currentCellIndex, elemCellIndex, elemRange, elemWs);
							continue
						}
					}
					if (!this._getPrecedents(cellIndex, elemCellIndex)) {
						isHaveUnrecorded = true;
					}
				}
			}

			return {
				isHaveUnrecorded: isHaveUnrecorded,
				newOutStack: newOutStack
			};
		}
		return false;
	};
	TraceDependentsManager.prototype.setPrecedentsLoop = function (inLoop) {
		this.inLoop = inLoop;
	};
	TraceDependentsManager.prototype.getPrecedentsLoop = function () {
		return this.inLoop;
	};
	TraceDependentsManager.prototype._getPrecedents = function (from, to) {
		return this.precedents[from] && this.precedents[from][to];
	};
	TraceDependentsManager.prototype._deleteDependent = function (from, to) {
		if (this.dependents[from] && this.dependents[from][to]) {
			delete this.dependents[from][to];
			if (Object.keys(this.dependents[from]).length === 0) {
				delete this.dependents[from];
			}
		}
	};
	TraceDependentsManager.prototype._deletePrecedentFromArea = function (index) {
		if (this.dependents[index] && this.precedents) {
			for (let precedentIndex in this.precedents) {
				for (let i in this.precedents[precedentIndex]) {
					if (i == index) {
						this._deleteDependent(index, precedentIndex);
						this._deletePrecedent(precedentIndex, index);
					}
				}
			}
		}
	};
	TraceDependentsManager.prototype._deletePrecedent = function (from, to) {
		if (this.precedents[from] && this.precedents[from][to]) {
			delete this.precedents[from][to];
			if (Object.keys(this.precedents[from]).length === 0) {
				delete this.precedents[from];
			}
		}
	};
	TraceDependentsManager.prototype._setPrecedents = function (from, to, isDependent, areaName) {
		if (!this.precedents) {
			this.precedents = {};
		}
		if (!this.precedents[from]) {
			this.precedents[from] = {};
		}
		if (from === to) {
			return;
		}
		// TODO calculated: 1, not_calculated: 2
		// TODO isAreaHeader: "A3:B4"
		// this.precedents[from][to] = isDependent ? 2 : 1;
		this.precedents[from][to] = areaName ? areaName : 1;
	};
	TraceDependentsManager.prototype._setPrecedentsAreaHeader = function (headerCellIndex, areaName) {
		if (!this.precedentsAreasHeaders) {
			this.precedentsAreasHeaders = {};
		}
		this.precedentsAreasHeaders[headerCellIndex] = areaName;
	};
	TraceDependentsManager.prototype._setPrecedentsAreas = function (area) {
		if (!this.precedentsAreas) {
			this.precedentsAreas = {};
		}
		Object.assign(this.precedentsAreas, area);
	};
	TraceDependentsManager.prototype._getPrecedentsAreas = function () {
		return this.precedentsAreas;
	};
	TraceDependentsManager.prototype.isHaveData = function () {
		return this.isHaveDependents() || this.isHavePrecedents();
	};
	TraceDependentsManager.prototype.isHaveDependents = function () {
		return !!this.dependents;
	};
	TraceDependentsManager.prototype.isHavePrecedents = function () {
		return !!this.precedents;
	};
	TraceDependentsManager.prototype.isHaveExternalPrecedents = function () {
		return !!this.precedentsExternal;
	};
	TraceDependentsManager.prototype.isHaveExternalDependents = function () {
		return !!this.dependentsExternal;
	};
	TraceDependentsManager.prototype.forEachDependents = function (callback) {
		for (let i in this.dependents) {
			callback(i, this.dependents[i], this.isPrecedentsCall);
		}
	};
	TraceDependentsManager.prototype.forEachExternalPrecedent = function (callback) {
		for (let i in this.precedentsExternal) {
			callback(i);
		}
	};
	TraceDependentsManager.prototype.clear = function (type) {
		if (Asc.c_oAscRemoveArrowsType.all === type) {
			this.clearAll();
		}
		if (Asc.c_oAscRemoveArrowsType.dependent === type) {
			this.clearLastDependent();
		}
		if (Asc.c_oAscRemoveArrowsType.precedent === type) {
			this.clearLastPrecedent();
		}
	};
	TraceDependentsManager.prototype.clearAll = function (needDraw) {
		if (needDraw && this.ws) {
			// need to call cleanSelection before any data is removed from the traceManager, because otherwise, isHaveData inside cleanSelection will return false, and the cleaning won't occur
			this.ws.cleanSelection();
		}
		this.precedents = null;
		this.precedentsExternal = null;
		this.dependentsExternal = null;
		this.dependents = null;
		this.isDependetsCall = null;
		this.inLoop = null;
		this.isPrecedentsCall = null;
		this.precedentsAreas = null;
		this.currentCalculatedPrecedentAreas = null;
		this.precedentsAreasHeaders = null;
		this._setDefaultData();
		this.clearCoordsData();

		if (needDraw) {
			if (this.ws && this.ws.overlayCtx) {
				// on the other hand, drawSelection should be called after removing data from traceManager because there is a dependency drawing call inside it
				this.ws._drawSelection();
			}
		}
	};
	TraceDependentsManager.prototype.changeDocument = function (prop, arg1, arg2) {
		switch (prop) {
			case AscCommonExcel.docChangedType.cellValue:
				if (this._lockChangeDocument) {
					return;
				}
				if (arg1) {
					this.clearCellTraces(arg1.nRow, arg1.nCol);
				}
				break;
			case AscCommonExcel.docChangedType.rangeValues:
				break;
			case AscCommonExcel.docChangedType.sheetContent:
				if (this._lockChangeDocument) {
					return;
				}
				this.clearAll();
				break;
			case AscCommonExcel.docChangedType.sheetRemove:
				break;
			case AscCommonExcel.docChangedType.sheetRename:
				break;
			case AscCommonExcel.docChangedType.sheetChangeIndex:
				break;
			case AscCommonExcel.docChangedType.markModifiedSearch:
				break;
			case AscCommonExcel.docChangedType.mergeRange:
				if (arg1 === true) {
					this._lockChangeDocument = true;
				} else {
					this._lockChangeDocument = null;
					let t = this;
					if (t.isHaveData() && arg2) {
						if (Asc.c_oAscSelectionType.RangeMax === arg2.getType()) {
							t.clearAll();
							break;
						}

						let maxRowToClear = t.ws.nRowsCount ? Math.min(t.ws.nRowsCount, arg2.r2) : arg2.r2,
							maxColToClear = t.ws.nColsCount ? Math.min(t.ws.nColsCount, arg2.c2) : arg2.c2;

						for (let col = arg2.c1; col <= maxColToClear; col++) {
							for (let row = arg2.r1; row <= maxRowToClear; row++) {
								if (!(arg2.c1 === col && arg2.r1 === row)) {
									t.clearCellTraces(row, col);
								}
							}
						}
					}
				}
				break;
		}
	};
	TraceDependentsManager.prototype.clearCellTraces = function (row, col) {
		let ws = this.ws && this.ws.model;
		if (!ws || row == null || col == null || !this.precedents || !this.dependents) {
			return;
		}

		let cellIndex = AscCommonExcel.getCellIndex(row, col);
		if (this.precedents[cellIndex]) {
			for (let i in this.precedents[cellIndex]) {
				this._deleteDependent(i, cellIndex);
			}
			delete this.precedents[cellIndex];
		}

		// check the ranges for existence of dependencies on it
		this.checkAreas();
	};
	TraceDependentsManager.prototype.checkAreas = function () {
		if (!this.precedentsAreas) {
			return
		}
		for (let i in this.precedentsAreas) {
			let areaHeader = this.precedentsAreas[i].areaHeader;
			if (!this.dependents[areaHeader]) {
				delete this.precedentsAreas[i];
				delete this.precedentsAreasHeaders[areaHeader];
			}
		}
	};
	TraceDependentsManager.prototype.clearCoordsData = function () {
		/* clear coordinatess data */
		this.tracesCoords = null;
	};
	/**
	 * Sets passed precedents cells index for recursive calls
	 * @memberof TraceDependentsManager
	 * @param {number} currentCellIndex
	 */
	TraceDependentsManager.prototype.setPassedPrecedents = function (currentCellIndex) {
		if (!this.aPassedPrecedents) {
			this.aPassedPrecedents = [];
		}
		this.aPassedPrecedents.push(currentCellIndex);
	};
	/**
	 * Checks current cell index is already passed in recursive calls
	 * @memberof TraceDependentsManager
	 * @param {number} currentCellIndex
	 * @returns {boolean}
	 */
	TraceDependentsManager.prototype.checkPassedPrecedents = function (currentCellIndex) {
		return !!(this.aPassedPrecedents && this.aPassedPrecedents.includes(currentCellIndex));
	};
	/**
	 * Clears attribute of passed precedents
	 * @memberof TraceDependentsManager
	 */
	TraceDependentsManager.prototype.clearPassedPrecedents = function () {
		this.aPassedPrecedents = null;
	};
	/**
	 * Sets passed dependents cells index for recursive calls
	 * @memberof TraceDependentsManager
	 * @param {number} currentCellIndex
	 */
	TraceDependentsManager.prototype.setPassedDependents = function (currentCellIndex) {
		if (!this.aPassedDependents) {
			this.aPassedDependents = [];
		}
		this.aPassedDependents.push(currentCellIndex);
	};
	/**
	 * Checks current cell index is already passed in recursive calls
	 * @memberof TraceDependentsManager
	 * @param {number} currentCellIndex
	 * @returns {boolean}
	 */
	TraceDependentsManager.prototype.checkPassedDependents = function (currentCellIndex) {
		return !!(this.aPassedDependents && this.aPassedDependents.includes(currentCellIndex));
	};
	/**
	 * Clears attribute of passed dependents
	 */
	TraceDependentsManager.prototype.clearPassedDependents = function () {
		this.aPassedDependents = null;
	};


	//------------------------------------------------------------export---------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["AscCommonExcel"].TraceDependentsManager = TraceDependentsManager;


})(window);
