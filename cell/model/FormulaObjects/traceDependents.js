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

	function TraceDependentsManager(ws) {
		this.ws = ws;
		this.precedents = null;
		this.dependents = null
	}
	TraceDependentsManager.prototype.calculateDependents = function (row, col) {
		//depend from row/col cell
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
		}
	};
	TraceDependentsManager.prototype._calculateDependents = function (cellIndex, curListener) {
		if (!this.dependents) {
			this.dependents = {};
		}

		let t = this;
		let wb = this.ws.model.workbook;
		let dependencyFormulas = wb.dependencyFormulas;
		let cellAddress = AscCommonExcel.getFromCellIndex(cellIndex, true);
		let findCellListeners = function () {
			if (curListener && curListener.areaMap) {
				for (let j  in curListener.areaMap) {
					if (curListener.areaMap.hasOwnProperty(j)) {
						if (curListener.areaMap[j] && curListener.areaMap[j].bbox.contains(cellAddress.col, cellAddress.row)) {
							return curListener.areaMap[j];
						}
					}
				}
			}
			return curListener.cellMap[cellIndex];
		};

		let getParentIndex = function (_parent) {
			let _parentCellIndex = AscCommonExcel.getCellIndex(_parent.nRow, _parent.nCol);
			//parent -> cell/defname
			if (_parent.parsedRef/*parent instanceof AscCommonExcel.DefName*/) {
				_parentCellIndex = null;
			} else if (_parent.ws !== t.ws.model) {
				_parentCellIndex += ";" + _parent.ws.index;
			}
			return _parentCellIndex;
		};

		let getListenersMap = function (_cellListeners) {
			if (!_cellListeners) {
				return;
			}
			let _listeners = {};
			for (let j in _cellListeners.listeners) {
				let pushListeners = function (_defNameListeners) {
					if (_defNameListeners) {
						for (let n in _defNameListeners.listeners) {
							if (_defNameListeners.listeners.hasOwnProperty(n)) {
								_listeners[n] = _defNameListeners.listeners[n];
							}
						}
					}
				};

				if (cellListeners.listeners.hasOwnProperty(j)) {
					if (_cellListeners.listeners[j].parent.parsedRef) {
						//def name
						let _name = _cellListeners.listeners[j].parent.name;
						let defNameWs = dependencyFormulas.getDefNameByName(_name, t.ws.model.Id,true);
						if (defNameWs) {
							let _range = defNameWs && defNameWs.parsedRef && defNameWs.parsedRef.getFirstRange();
							if (_range.bbox.r1 === cellAddress.row && _range.bbox.c1 === cellAddress.col) {
								let defNameListeners = dependencyFormulas.defNameListeners[_name];
								pushListeners(defNameListeners);
							}
						} else {
							let defNameWb = dependencyFormulas.getDefNameByName(_name);
							if (defNameWb) {
								let defNameListeners = dependencyFormulas.defNameListeners[_name];
								pushListeners(defNameListeners);
							}
						}
					} else {
						_listeners[j] = _cellListeners.listeners[j];
					}
				}
			}
			return _listeners;
		};

		let cellListeners = findCellListeners();
		let listenersMap = getListenersMap(cellListeners);
		if (listenersMap) {
			if (!this.dependents[cellIndex]) {
				this.dependents[cellIndex] = {};
				for (let i in listenersMap) {
					if (listenersMap.hasOwnProperty(i)) {
						let parent = listenersMap[i].parent;
						let parentCellIndex = getParentIndex(parent);
						if (parentCellIndex === null) {
							continue;
						}
						this.dependents[cellIndex][parentCellIndex] = 1;
					}
				}
			} else {
				//if change formulas and add new sheetListeners
				//check current tree
				let isUpdated = false;
				for (let i in listenersMap) {
					if (listenersMap.hasOwnProperty(i)) {
						let parent = listenersMap[i].parent;
						let parentCellIndex = getParentIndex(parent);
						if (parentCellIndex === null) {
							continue;
						}
						if (!this.dependents[cellIndex][parentCellIndex]) {
							this.dependents[cellIndex][parentCellIndex] = 1;
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
	TraceDependentsManager.prototype._updateDependents = function (cellIndex, curListener) {
		if (this.dependents[cellIndex]) {
			let cellListeners = curListener.cellMap[cellIndex];
			if (cellListeners && cellListeners.listeners) {
				if (this.dependents[cellIndex]) {
					this.dependents[cellIndex] = {};
					for (let i in cellListeners.listeners) {
						let parentCellIndex = AscCommonExcel.getCellIndex(cellListeners.listeners[i].parent.nRow, cellListeners.listeners[i].parent.nCol);
						this.dependents[cellIndex][parentCellIndex] = 1;
					}
				}
			}
		} else {
			for (let i in this.dependents[cellIndex]) {
				this._calculateDependents(i, curListener);
			}
		}

	};
	TraceDependentsManager.prototype.calculatePrecedents = function (row, col) {

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
		for (var i in this.dependents) {
			callback(i, this.dependents[i]);
		}
	};
	TraceDependentsManager.prototype.clear = function (type) {
		if (Asc.c_oAscRemoveArrowsType.all === type || Asc.c_oAscRemoveArrowsType.precedent === type) {
			this.precedents = null;
		}
		if (Asc.c_oAscRemoveArrowsType.all === type || Asc.c_oAscRemoveArrowsType.dependent === type) {
			this.dependents = null;
		}
	};







	//------------------------------------------------------------export---------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["AscCommonExcel"].TraceDependentsManager = TraceDependentsManager;


})(window);
