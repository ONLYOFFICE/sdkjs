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

	function TraceDependentsManager(ws) {
		this.ws = ws;
		this.precedents = null;
		this.precedentsExternal = null;
		this.currentPrecedents = null;
		this.dependents = null;
		this.isDependetsCall = null;
		this.inLoop = null;
		this.isPrecedentsCall = null;
		this.precedentsAreas = null;
		this.inPrecedentsAreasLoop = null;
		this.data = {
			recLevel: 0,
			maxRecLevel: 0,
			indices: {
				// cellIndex: level
			}
		};
		this.currentPrecedentsAreas = null;
	}

	TraceDependentsManager.prototype.setPrecedentsCall = function () {
		this.isPrecedentsCall = true;
		this.isDependetsCall = false;
	};
	TraceDependentsManager.prototype.setDependentsCall = function () {
		this.isDependetsCall = true;
		this.isPrecedentsCall = false;
	};
	TraceDependentsManager.prototype.setPrecedentExternal = function (cellIndex) {
		if (!this.precedentsExternal) {
			this.precedentsExternal = new Set();
		}
		this.precedentsExternal.add(cellIndex);
	};
	TraceDependentsManager.prototype.checkPrecedentExternal = function (cellIndex) {
		if (!this.precedentsExternal) {
			return false;
		}
		return this.precedentsExternal.has(cellIndex);
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
		}

		const findMaxNesting = function (row, col) {
			let currentCellIndex = AscCommonExcel.getCellIndex(row, col);

			if (t.dependents[currentCellIndex]) {
				let interLevel, fork;
				if (Object.keys(t.dependents[currentCellIndex]).length > 1) {
					fork = true;
				}

				t.data.recLevel++;
				t.data.maxRecLevel = t.data.recLevel > t.data.maxRecLevel ? t.data.recLevel : t.data.maxRecLevel;
				interLevel = t.data.recLevel;
				for (let j in t.dependents[currentCellIndex]) {
					if (j.includes(";")) {
						let uniqueIndex = j + "|" + currentCellIndex;
						t.data.indices[uniqueIndex] = t.data.recLevel;
						continue;
					}
					let coords = AscCommonExcel.getFromCellIndex(j, true);
					findMaxNesting(coords.row, coords.col);
					t.data.recLevel = fork ? interLevel : t.data.recLevel;
				}
			} else {
				t.data.indices[currentCellIndex] = t.data.recLevel;
				return;
			}
		}

		findMaxNesting(row, col);

		const maxLevel = this.data.maxRecLevel;

		if (maxLevel === 0) {
			this._setDefaultData();
			return;
		}

		for (let [index, indexLevel] of Object.entries(this.data.indices)) {
			if (indexLevel == maxLevel) {
				let fromIndex;
				if (index.includes(";")) {
					let parts = index.split("|");
					index = parts[0];
					fromIndex = parts[1];
				} else {
					// fromIndex = this.precedents && this.precedents[index] ?  Object.keys(this.precedents[index])[0] : null;
					fromIndex = this.precedents && this.precedents[index] ?  this.precedents[index] : null;
				}
				if (fromIndex) {
					if (typeof fromIndex === 'string' && index.includes(";")) {
						this._deleteDependent(fromIndex, index);
						this._deletePrecedent(index, fromIndex);
						continue;
					} else {
						for (let i in fromIndex) {
							this._deleteDependent(i, index);
							this._deletePrecedent(index, i);
						}
					}
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
	TraceDependentsManager.prototype._calculateDependents = function (cellIndex, curListener) {
		let t = this;
		let ws = this.ws.model;
		let wb = this.ws.model.workbook;
		let dependencyFormulas = wb.dependencyFormulas;
		let allDefNamesListeners = dependencyFormulas.defNameListeners;
		let cellAddress = AscCommonExcel.getFromCellIndex(cellIndex, true);
		const currentCellInfo = {};

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
		const getHeader = function (table) {
			if (!table.Ref) {
				return false;
			}

			return {col: table.Ref.c1, row: table.Ref.r1};
		};
		const setDefNameIndexes = function (defName, isTable) {
			let tableHeader = isTable ? getHeader(ws.getTableByName(defName)) : false;
			let isCurrentCellHeader = isTable ? checkIfHeader(tableHeader) : false;
			for (const i in allDefNamesListeners) {
				if (allDefNamesListeners.hasOwnProperty(i) && i.toLowerCase() === defName.toLowerCase()) {
					for (const listener in allDefNamesListeners[i].listeners) {
						// TODO возможно стоить добавить все слушатели сразу в curListener
						let elem = allDefNamesListeners[i].listeners[listener];
						let isArea = elem.ref ? !elem.ref.isOneCell() : false;
						let is3D = elem.ws.Id ? elem.ws.Id !== ws.Id : false;
						if (isArea && !is3D && !isCurrentCellHeader) {
							// decompose all elements into dependencies
							let areaIndexes = getAllAreaIndexes(elem);
							if (areaIndexes) {
								for (let index of areaIndexes) {
									t._setDependents(cellIndex, index);
								}
								continue;
							}
						}
						let parentCellIndex = getParentIndex(elem.parent);
						if (!parentCellIndex) {
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
								// continue;
							} else if (!elem.Formula.includes("Headers") && isCurrentCellHeader) {
								continue;
							}
							// ?additional check if the listener is in the same table, need to check if it is a listener of the main cell
							if (elem.outStack) {
								let arr = [];
								// check each element of the stack for an occurrence in the original cell
								for (let table in elem.outStack) {
									if (elem.outStack[table].type !== AscCommonExcel.cElementType.table) {
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
							t._setDependents(cellIndex, parentCellIndex);
							t._setPrecedents(parentCellIndex, cellIndex);
							continue;
						} else {
							t._setDependents(cellIndex, parentCellIndex);
							t._setPrecedents(parentCellIndex, cellIndex);
						}
					}
				}
			}
		};
		const getAllAreaIndexes = function (parserFormula) {
			const indexes = [], range = parserFormula.ref;
			if (!range) {
				return;
			}
			for (let i = range.c1; i <= range.c2; i++) {
				for (let j = range.r1; j <= range.r2; j++) {
					let index = AscCommonExcel.getCellIndex(j, i);
					indexes.push(index);
				}
			}

			return indexes;
		};
		const getParentIndex = function (_parent) {
			let _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, _parent.nCol);
			//parent -> cell/defname
			if (_parent.parsedRef/*parent instanceof AscCommonExcel.DefName*/) {
				_parentCellIndex = null;
			} else if (_parent.ws !== t.ws.model) {
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
							let res = curListener.areaMap[j].bbox.getSharedIntersect(shared.ref, currentRange.bbox);
							// draw dependents to coords from res
							if (res && (res.r1 === res.r2 && res.c1 === res.c2)) {
								let index = AscCommonExcel.getCellIndex(res.r1, res.c1);
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
							for (let index of indexes) {
								t._setDependents(cellIndex, index);
								t._setPrecedents(index, cellIndex);
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
				for (let i in cellListeners) {
					if (cellListeners.hasOwnProperty(i)) {
						let parent = cellListeners[i].parent;
						let parentWsId = parent.ws ? parent.ws.Id : null;
						let isTable = parent.parsedRef ? parent.parsedRef.isTable : false;
						let isDefName = parent.name ? true : false;
						let formula = cellListeners[i].Formula;
						let is3D = false;
						const parentInfo = {
							parent, 
							parentWsId,
							isTable,
							isDefName
						};

						if (isDefName) {
							// TODO check external table ref
							setDefNameIndexes(parent.name, isTable);
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
								for (let index of areaIndexes) {
									this._setDependents(cellIndex, index);
									this._setPrecedents(index, cellIndex);
								}
								continue;
							}
						}
						let parentCellIndex = getParentIndex(parent);
						
						if (parentCellIndex === null) {
							continue;
						}
						this._setDependents(cellIndex, parentCellIndex);
						this._setPrecedents(parentCellIndex, cellIndex, true);
					}
				}
			} else {
				// if dependents by cellIndex aldready exist, check current tree
				let currentIndex = Object.keys(this.dependents[cellIndex])[0];
				let isUpdated = false;
				for (let i in cellListeners) {
					if (cellListeners.hasOwnProperty(i)) {
						let parent = cellListeners[i].parent;
						let elemCellIndex = cellListeners[i].shared !== null ? currentIndex : getParentIndex(parent);
						let formula = cellListeners[i].Formula;

						if (parent.name) {
							continue;
						}

						if (formula.includes(":") && !cellListeners[i].is3D) {
							// call getAllAreaIndexes which return cellIndexes of each element(this will be parentCellIndex)
							let areaIndexes = getAllAreaIndexes(cellListeners[i]);
							if (areaIndexes) {
								// go through the values and set dependents for each
								for (let index of areaIndexes) {
									this._setDependents(cellIndex, index);
								}
								continue;
							}
						}

						// if the child cell does not yet have a dependency with listeners, create it
						if (!this._getDependents(cellIndex, elemCellIndex)) {
							this._setDependents(cellIndex, elemCellIndex);
							isUpdated = true;
						}
					}
				}
				
				if (!isUpdated) {
					for (let i in this.dependents[cellIndex]) {
						if (this.dependents[cellIndex].hasOwnProperty(i)) {
							this._calculateDependents(i, curListener);
						}
					}
				}
			}
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
		this.dependents[from][to] = 1;
	};
	TraceDependentsManager.prototype._setDefaultData = function () {
		this.data = {
			recLevel: 0,
			maxRecLevel: 0,
			indices: {}
		};
	}
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
		}

		const findMaxNesting = function (row, col) {
			let currentCellIndex = AscCommonExcel.getCellIndex(row, col);

			let formulaParsed;
			ws.getCell3(row, col)._foreachNoEmpty(function (cell) {
				formulaParsed = cell.formulaParsed;
			});
	
			if (!formulaParsed) {
				t.data.indices[currentCellIndex] = t.data.recLevel;
				return;
			}
	
			if (t.precedents[currentCellIndex]) {
				let interLevel, fork;
				if (Object.keys(t.precedents[currentCellIndex]).length > 1) {
					fork = true;
				}

				t.data.recLevel++;
				t.data.maxRecLevel = t.data.recLevel > t.data.maxRecLevel ? t.data.recLevel : t.data.maxRecLevel;
				interLevel = t.data.recLevel;
				for (let j in t.precedents[currentCellIndex]) {
					if (j.includes(";")) {
						let uniqueIndex = j + "|" + currentCellIndex;
						t.data.indices[uniqueIndex] = t.data.recLevel;
						continue;
					}
					let coords = AscCommonExcel.getFromCellIndex(j, true);
					findMaxNesting(coords.row, coords.col);
					t.data.recLevel = fork ? interLevel : t.data.recLevel;
				}
			} else {
				t.data.indices[currentCellIndex] = t.data.recLevel;
				return;
			}
		}

		const filterAreas = function (areas, cell) {
			let result = {}; 
			for (let area in areas) {
				if (areas.hasOwnProperty(area)) {
					if (!areas[area].range.contains2(cell)) {
						result[area] = Object.assign({}, areas[area]);
					}
				}
			}
			return result; 
		}

		findMaxNesting(row, col);

		const maxLevel = this.data.maxRecLevel;

		if (maxLevel === 0) {
			this._setDefaultData();
			return;
		}

		for (let [index, indexLevel] of Object.entries(this.data.indices)) {
			// ? add is3D flag
			if (indexLevel == maxLevel) {
				let fromIndex;
				if (index.includes(";")) {
					let parts = index.split("|");
					index = parts[0];
					fromIndex = parts[1];
				} else {
					fromIndex = this.dependents && this.dependents[index] ?  Object.keys(this.dependents[index])[0] : null;
				}
				if (fromIndex) {
					if (index.includes(";")) {
						this._deleteDependent(index, fromIndex);
						this._deletePrecedent(fromIndex, index);
						continue;
					}
					// TODO check if index in area
					let indexCoords = AscCommonExcel.getFromCellIndex(index, true);
					if (this.precedentsAreas) {
						// check all areas
						let res = filterAreas(this.precedentsAreas, indexCoords);
						this.precedentsAreas = res;
					}

					this._deleteDependent(index, fromIndex);
					this._deletePrecedent(fromIndex, index);
				}
			}
		}

		this._setDefaultData();
	};
	TraceDependentsManager.prototype.calculatePrecedents = function (row, col, isSecondCall) {
		//depend from row/col cell
		let ws = this.ws && this.ws.model;
		if (!ws) {
			return;
		}
		const t = this;
		const cellIndex = AscCommonExcel.getCellIndex(row, col);
		if (row == null || col == null) {
			let selection = ws.getSelection();
			let activeCell = selection.activeCell;
			row = activeCell.row;
			col = activeCell.col;
		}

		// const getAllAreaIndexesOld = function (areas) {
		// 	const indexes = [];
		// 	if (!areas) {
		// 		return;
		// 	}
		// 	for (const area in areas) {
		// 		if (areas[area].isCalculated) {
		// 			continue;
		// 		}
		// 		for (let i = areas[area].range.r1; i <= areas[area].range.r2; i++) {
		// 			for (let j = areas[area].range.c1; j <= areas[area].range.c2; j++) {
		// 				let index = AscCommonExcel.getCellIndex(i, j);
		// 				indexes.push(index);
		// 			}
		// 		}
		// 		areas[area].isCalculated = true;
		// 	}

		// 	return indexes;
		// };

		const getAllAreaIndexes = function (areas, header) {
			const indexes = [];
			if (!areas) {
				return;
			}
			for (const area in areas) {
				if (areas[area].isCalculated || areas[area].areaHeader !== header) {
					continue;
				}
				for (let i = areas[area].range.r1; i <= areas[area].range.r2; i++) {
					for (let j = areas[area].range.c1; j <= areas[area].range.c2; j++) {
						let index = AscCommonExcel.getCellIndex(i, j);
						indexes.push(index);
					}
				}
				areas[area].isCalculated = true;
			}

			return indexes;
		};

		const isCellTableHeader = function (cellIndex) {
			if (!t.currentPrecedentsAreas) {
				return;
			}
			for (let area in t.currentPrecedentsAreas) {
				if (t.currentPrecedentsAreas[area].areaHeader === cellIndex) {
					return true;
				} 
			}
		}

		let formulaParsed;
		ws.getCell3(row, col)._foreachNoEmpty(function (cell) {
			formulaParsed = cell.formulaParsed;
		});

		if (formulaParsed) {
			this._calculatePrecedents(formulaParsed, row, col, isSecondCall);
			this.setPrecedentsCall();
		}

		// TODO another way to check table
		let isCellHeader = isCellTableHeader(cellIndex);

		if (this.currentPrecedentsAreas && !this.inPrecedentsAreasLoop && isSecondCall && isCellHeader) {
			// calculate all precedents in areas
			this.setPrecedentsAreasLoop(true);
			let areaIndexes = getAllAreaIndexes(this.currentPrecedentsAreas, cellIndex);
			if (areaIndexes) {
				// go through the values and check precedents for each
				for (let index of areaIndexes) {
					let cellAddress = AscCommonExcel.getFromCellIndex(index, true), formula;
					ws.getCell3(cellAddress.row, cellAddress.col)._foreachNoEmpty(function (cell) {
						formula = cell.formulaParsed;
					});
					// let res = isCellHaveUnrecordedTraces(index, formula);
					let res = this.isCellHaveUnrecordedTraces(index, formula);
					if (!res) {
						continue;
					}
					
					this.calculatePrecedents(cellAddress.row, cellAddress.col);
				}
			}
			this.setPrecedentsAreasLoop(false);
			// return;
		}
		// else {
			// let formulaParsed;
			// ws.getCell3(row, col)._foreachNoEmpty(function (cell) {
			// 	formulaParsed = cell.formulaParsed;
			// });

			// if (formulaParsed) {
			// 	this._calculatePrecedents(formulaParsed, row, col, isSecondCall);
			// 	this.setPrecedentsCall();
			// }
		// }
	};
	TraceDependentsManager.prototype._calculatePrecedents = function (formulaParsed, row, col, isSecondCall) {
		if (!this.precedents) {
			this.precedents = {};
		}
		let t = this;
		let currentCellIndex = AscCommonExcel.getCellIndex(row, col);
		let isHaveUnrecorded = this.isCellHaveUnrecordedTraces(currentCellIndex, formulaParsed);

		if (isHaveUnrecorded) {
		// if (!this.precedents[currentCellIndex]) {
			let shared, base;
			if (formulaParsed.shared !== null) {
				shared = formulaParsed.getShared();
				base = shared.base;		// base index - where shared formula start
			}
			
			if (formulaParsed.outStack) {
				let currentWsIndex = formulaParsed.ws.index;
				let ref = formulaParsed.ref;
				// iterate and find all reference
				for (const elem of formulaParsed.outStack) {
					let elemType = elem.type ? elem.type : null;
					if (elemType === AscCommonExcel.cElementType.cell || elemType === AscCommonExcel.cElementType.cellsRange || 
						elemType === AscCommonExcel.cElementType.cell3D || elemType === AscCommonExcel.cElementType.cellsRange3D ||
						elemType === AscCommonExcel.cElementType.name || elemType === AscCommonExcel.cElementType.name3D || 
						elemType === AscCommonExcel.cElementType.table) {
						let is3D = elemType === AscCommonExcel.cElementType.cell3D || elemType === AscCommonExcel.cElementType.cellsRange3D || elemType === AscCommonExcel.cElementType.name3D;
						let isArea = elemType === AscCommonExcel.cElementType.cellsRange || elemType === AscCommonExcel.cElementType.name;
						let isDefName = elemType === AscCommonExcel.cElementType.name || elemType === AscCommonExcel.cElementType.name3D;
						let isTable = elemType === AscCommonExcel.cElementType.table;
						let elemRange, elemCellIndex, 
							cellRange = new asc_Range(col, row, col, row);

						if (isDefName) {
							let elemDefName = elem.getDefName();
							let elemValue = elem.getValue();
							let defNameParentWsIndex = elemDefName.parsedRef.outStack[0].wsFrom ? elemDefName.parsedRef.outStack[0].wsFrom.index : (elemDefName.parsedRef.outStack[0].ws ? elemDefName.parsedRef.outStack[0].ws.index : null);
							elemRange = elemValue.range.bbox ? elemValue.range.bbox : elemValue.bbox;

							if (defNameParentWsIndex && defNameParentWsIndex !== currentWsIndex) {
								// 3D
								is3D = true;
								isArea = false;
							} else if (elemRange.isOneCell()) {
								isArea = false;
							}
						} else if (isTable) {
							let currentWsId = elem.ws.Id,
								elemWsId = elem.area.ws ? elem.area.ws.Id : elem.area.wsFrom.Id;
							// elem.area can be cRef and cArea
							is3D = currentWsId !== elemWsId ? true : false;
							elemRange = elem.area.bbox ? elem.area.bbox : (elem.area.range ? elem.area.range.bbox : null);
							isArea = ref ? true : (!elemRange.isOneCell() ? true : false);

						} else {
							elemRange = elem.range.bbox ? elem.range.bbox : elem.bbox;
						}

						if (!elemRange) {
							return;
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
							} else {
								elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1 + (row - base.nRow), elemRange.c1 + (col - base.nCol));
							}
						} else {
							elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);
						}
						
						if (isArea && !ref) {
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

						if (isArea) {
							const areaRange = {};
							const areaName = elem.value;	// areaName - unique key for areaRange
							areaRange[areaName] = {};
							areaRange[areaName].range = elemRange;
							areaRange[areaName].isCalculated = null;
							areaRange[areaName].areaHeader = elemCellIndex;

							this._setPrecedentsAreas(areaRange);
						}

						let elemDependentsIndex = elemCellIndex;
						let currentDependentsCellIndex = currentCellIndex;

						if (is3D) {
							elemCellIndex += ";" + (elem.wsTo ? elem.wsTo.index : elem.ws.index);
							// currentDependentsCellIndex += ";" + currentWsIndex
							this._setDependents(elemCellIndex, currentCellIndex);
							this._setPrecedents(currentCellIndex, elemCellIndex);
							this.setPrecedentExternal(currentCellIndex);
						} else {
							this._setPrecedents(currentCellIndex, elemCellIndex);
							// change links to another sheet
							this._setDependents(elemDependentsIndex, currentDependentsCellIndex);
						}
					}
				}
			}
		} 
		else {
		// else if (!this.getPrecedentsLoop()) {
			this.currentPrecedents = Object.assign({}, this.precedents);
			this.currentPrecedentsAreas = Object.assign({}, this._getPrecedentsAreas());
			this.setPrecedentsLoop(true);

			// check all current precedents randomly (not in order of building dependencies)
			// for (let i in currentPrecedent) {
			// 	if (currentPrecedent.hasOwnProperty(i)) {
			// 		for (let j in currentPrecedent[i]) {
			// 			if (currentPrecedent[i].hasOwnProperty(j)) {
			// 				let coords = AscCommonExcel.getFromCellIndex(j, true);
			// 				this.calculatePrecedents(coords.row, coords.col, true);
			// 			}
			// 		}
			// 	}
			// }

			// ??? choose the method for currentPrecedent

			// check first level, then if function return false, check second, third and so on
			// for (let i in this.precedents[currentCellIndex]) {
			for (let i in this.currentPrecedents[currentCellIndex]) {
				let coords = AscCommonExcel.getFromCellIndex(i, true);
				this.calculatePrecedents(coords.row, coords.col, true);	
			}

			this.setPrecedentsLoop(false);
		}
	};
	TraceDependentsManager.prototype.isCellHaveUnrecordedTraces = function (cellIndex, formulaParsed) {
		if (formulaParsed && formulaParsed.outStack) {
			let currentWsIndex = formulaParsed.ws.index;
			for (const elem of formulaParsed.outStack) {
				let elemType = elem.type ? elem.type : null;
				if (elemType === AscCommonExcel.cElementType.cell || elemType === AscCommonExcel.cElementType.cellsRange || 
					elemType === AscCommonExcel.cElementType.cell3D || elemType === AscCommonExcel.cElementType.cellsRange3D ||
					elemType === AscCommonExcel.cElementType.name || elemType === AscCommonExcel.cElementType.name3D || 
					elemType === AscCommonExcel.cElementType.table) {
					let is3D = elemType === AscCommonExcel.cElementType.cell3D || elemType === AscCommonExcel.cElementType.cellsRange3D || elemType === AscCommonExcel.cElementType.name3D;
					let isArea = elemType === AscCommonExcel.cElementType.cellsRange || elemType === AscCommonExcel.cElementType.name;
					let isDefName = elemType === AscCommonExcel.cElementType.name || elemType === AscCommonExcel.cElementType.name3D;
					let isTable = elemType === AscCommonExcel.cElementType.table;
					let elemRange, elemCellIndex;

					if (isDefName) {
						let elemDefName = elem.getDefName();
						let elemValue = elem.getValue();
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
						is3D = currentWsId !== elemWsId ? true : false;
						elemRange = elem.area.bbox ? elem.area.bbox : (elem.area.range ? elem.area.range.bbox : null);
					} else {
						elemRange = elem.range.bbox ? elem.range.bbox : elem.bbox;
					}

					if (!elemRange) {
						return;
					}
					
					elemCellIndex = AscCommonExcel.getCellIndex(elemRange.r1, elemRange.c1);

					if (is3D) {
						elemCellIndex += ";" + (elem.wsTo ? elem.wsTo.index : elem.ws.index);
						
					} 
					if (!this._getPrecedents(cellIndex, elemCellIndex)) {
						return true;
					}
				}
			}
		}
	}
	TraceDependentsManager.prototype.setPrecedentsLoop = function (inLoop) {
		this.inLoop = inLoop;
	};
	TraceDependentsManager.prototype.setPrecedentsAreasLoop = function (inLoop) {
		this.inPrecedentsAreasLoop = inLoop;
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
	TraceDependentsManager.prototype._deletePrecedent = function (from, to) {
		if (this.precedents[from] && this.precedents[from][to]) {
			delete this.precedents[from][to];
			if (Object.keys(this.precedents[from]).length === 0) {
				delete this.precedents[from];
			}
		}
	};
	TraceDependentsManager.prototype._setPrecedents = function (from, to, isDependent) {
		if (!this.precedents) {
			this.precedents = {};
		}
		if (!this.precedents[from]) {
			this.precedents[from] = {};
		}
		// calculated: 1, not_calculated: 2
		this.precedents[from][to] = isDependent ? 2 : 1;
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
	TraceDependentsManager.prototype.forEachDependents = function (callback) {
		for (let i in this.dependents) {
			callback(i, this.dependents[i], this.isPrecedentsCall);
		}
	};
	TraceDependentsManager.prototype.forEachExternalPrecedent = function (callback) {
		for (let i in this.precedents) {
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
	TraceDependentsManager.prototype.clearAll = function () {
		this.precedents = null;
		this.precedentsExternal = null;
		this.currentPrecedents = null;
		this.dependents = null;
		this.isDependetsCall = null;
		this.inLoop = null;
		this.isPrecedentsCall = null;
		this.precedentsAreas = null;
		this.currentPrecedentsAreas = null;
		this.inPrecedentsAreasLoop = null;
		this._setDefaultData();
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
	}







	//------------------------------------------------------------export---------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["AscCommonExcel"].TraceDependentsManager = TraceDependentsManager;


})(window);
