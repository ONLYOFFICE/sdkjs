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

	function TraceDependentsCellManager(ws) {
		this.ws = ws;
		this.precedents = null;
		this.dependents = null
	}
	TraceDependentsCellManager.prototype.calculateDependents = function (row, col) {
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
			return;




			let cellListeners = curListener.cellMap[cellIndex];
			if (cellListeners && cellListeners.listeners) {
				if (!this.dependents[cellIndex]) {
					this.dependents[cellIndex] = [];
				}

				let nextLevel = this.dependents[cellIndex].length;
				if (!this.dependents[cellIndex][nextLevel]) {
					this.dependents[cellIndex][nextLevel] = {};
				}
				if (nextLevel === 0) {
					for (let i in cellListeners.listeners) {
						let parentCellIndex = AscCommonExcel.getCellIndex(cellListeners.listeners[i].parent.nRow, cellListeners.listeners[i].parent.nCol);
						this.dependents[cellIndex][nextLevel][parentCellIndex] = 1;
					}
				} else {
					for (let i in this.dependents[cellIndex][nextLevel - 1]) {

					}
				}
			}
		}
	};
	TraceDependentsCellManager.prototype._calculateDependents = function (cellIndex, curListener) {
		if (!this.dependents) {
			this.dependents = {};
		}

		if (!this.dependents[cellIndex]) {
			let cellListeners = curListener.cellMap[cellIndex];
			if (cellListeners && cellListeners.listeners) {
				if (!this.dependents[cellIndex]) {
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
	TraceDependentsCellManager.prototype.calculatePrecedents = function (row, col) {

	};
	TraceDependentsCellManager.prototype.draw = function (visibleRange, offsetX, offsetY, args) {
		if (this.dependents) {

		}
	};
	TraceDependentsCellManager.prototype.isHaveDependents = function () {
		return !!this.dependents;
	};
	TraceDependentsCellManager.prototype.isHavePrecedents = function () {
		return !!this.precedents;
	};




	//------------------------------------------------------------export---------------------------------------------------
	window['AscCommonExcel'] = window['AscCommonExcel'] || {};

	window["AscCommonExcel"].TraceDependentsCellManager = TraceDependentsCellManager;


})(window);
