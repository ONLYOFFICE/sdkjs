/*
 * Copyright (C) Ascensio System SIA, 2009-2026
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation, together with the
 * additional terms provided in the LICENSE file.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. For
 * details, see the GNU AGPL at: https://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA by email at info@onlyoffice.com
 * or by postal mail at 20A-6 Ernesta Birznieka-Upisha Street, Riga,
 * LV-1050, Latvia, European Union.
 *
 * The interactive user interfaces in modified versions of the Program
 * are required to display Appropriate Legal Notices in accordance with
 * Section 5 of the GNU AGPL version 3.
 *
 * No trademark rights are granted under this License.
 *
 * All non-code elements of the Product, including illustrations,
 * icon sets, and technical writing content, are licensed under the
 * Creative Commons Attribution-ShareAlike 4.0 International License:
 * https://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
 * This license applies only to such non-code elements and does not
 * modify or replace the licensing terms applicable to the Program's
 * source code, which remains licensed under the GNU Affero General
 * Public License v3.
 *
 * SPDX-License-Identifier: AGPL-3.0-only
 */

"use strict";

(function (window, undefined) {
		/*
		 * Import
		 * -----------------------------------------------------------------------------
		 */
		var asc_applyFunction = AscCommonExcel.applyFunction;

		var c_oAscLockTypes = AscCommon.c_oAscLockTypes;

		var c_oAscRecalcIndexTypes = AscCommonExcel.c_oAscRecalcIndexTypes;
		var c_oAscLockTypeElemSubType = AscCommonExcel.c_oAscLockTypeElemSubType;
		var c_oAscLockTypeElem = AscCommonExcel.c_oAscLockTypeElem;

		/**
		 * Responsible for collaborative editing
		 * -----------------------------------------------------------------------------
		 *
		 * @constructor
		 * @memberOf AscCommonExcel
		 */
		function CCollaborativeEditing(handlers, isViewerMode) {
			if (!(this instanceof CCollaborativeEditing)) {
				return new CCollaborativeEditing();
			}
			AscCommon.CCollaborativeEditingBase.call(this);

			this.m_nUseType					= 1;  // 1 - single client and we save history, -1 - multiple clients, 0 - transition from -1 to 1

			this.handlers					= new AscCommonExcel.asc_CHandlersList(handlers);
			this.m_bIsViewerMode			= !!isViewerMode; // Viewer mode
			this.m_bGlobalLock				= 0; // Global lock
			this.m_bGlobalLockEditCell		= false; // Global lock (for cell editing) - disable selection change, but allow immediate input
			this.m_arrCheckLocks			= [];    // Array for checking lock status of objects we intend to modify

			this.m_arrNeedUnlock			= []; // Array with list of locked objects (locked by other users)
			this.m_arrNeedUnlock2			= []; // Array with list of locked objects (locked on this client)

			this.m_arrChanges				= []; // Array with changes from other users

			this.m_oRecalcIndexColumns		= {};
			this.m_oRecalcIndexRows			= {};

			this.m_oInsertColumns			= {}; // Map of sheets with arrays of added columns lists
			this.m_oInsertRows				= {}; // Map of sheets with arrays of added rows lists

			this.m_bFast = false;

			this.m_aForeignCursorsData = {};
			this.WaitImages = {};//cell/model/DrawingObjects/GlobalCounters.js

			this.init();

			return this;
		}

		CCollaborativeEditing.prototype = Object.create(AscCommon.CCollaborativeEditingBase.prototype);
		CCollaborativeEditing.prototype.init = function () {
		};

		// Clear recalculation indexes (necessary when opening)
		CCollaborativeEditing.prototype.clearRecalcIndex = function () {
			this.m_oRecalcIndexColumns = {};
			this.m_oRecalcIndexRows = {};
		};

		// Start collaborative editing
		CCollaborativeEditing.prototype.startCollaborationEditing = function () {
			this.m_nUseType = -1;
		};

		// Temporary end of collaborative editing
		CCollaborativeEditing.prototype.endCollaborationEditing = function () {
			if (this.m_nUseType <= 0)
				this.m_nUseType = 0;
		};

		// Setting view mode
		CCollaborativeEditing.prototype.setViewerMode = function (isViewerMode) {
			this.m_bIsViewerMode = isViewerMode;
		};

		CCollaborativeEditing.prototype.setFast = function (bFast) {
			if (false === bFast) {
				this.Remove_AllForeignCursors();
			}
			return this.m_bFast = bFast;
		};
		CCollaborativeEditing.prototype.getFast = function () {
			return this.m_bFast;
		};
		CCollaborativeEditing.prototype.Is_SingleUser = function () {
			return !this.getCollaborativeEditing();
		};
		CCollaborativeEditing.prototype.getCollaborativeEditing = function () {
			if (this.m_bIsViewerMode)
				return false;
			return 1 !== this.m_nUseType;
		};

		CCollaborativeEditing.prototype.haveOtherChanges = function () {
			return 0 < this.m_arrChanges.length;
		};

		CCollaborativeEditing.prototype.getOwnLocksLength = function () {
			return this.m_arrNeedUnlock2.length;
		};

		//-----------------------------------------------------------------------------------
		// Functions for checking object lock status
		//-----------------------------------------------------------------------------------
		CCollaborativeEditing.prototype.getGlobalLock = function () {
			return this.m_bGlobalLock;
		};
		CCollaborativeEditing.prototype.getGlobalLockEditCell = function () {
			return this.m_bGlobalLockEditCell;
		};
		CCollaborativeEditing.prototype.onStartEditCell = function () {
			// Call this function only when editing a cell and if we are not the only one editing!!!
			if (this.getCollaborativeEditing())
				this.m_bGlobalLockEditCell = true;
		};
		CCollaborativeEditing.prototype.onStopEditCell = function () {
			// Call this function only when finishing cell editing!!!
			this.m_bGlobalLockEditCell = false;
		};
		CCollaborativeEditing.prototype.lock = function (arrLocks, callback) {
			var type;
			callback = this._checkCollaborative(callback);

			this.onStartCheckLock();
			for (var i = 0; i < arrLocks.length; ++i) {
				type = this._addCheckLock(arrLocks[i], callback);
				if (c_oAscLockTypes.kLockTypeNone !== type) {
					// Release global lock (for cell editing)
					this.m_bGlobalLockEditCell = false;
					return c_oAscLockTypes.kLockTypeMine === type;
				}
			}
			this.onEndCheckLock(callback);
			return true;
		};
		CCollaborativeEditing.prototype._checkCollaborative = function (callback) {
			if (false === this.getCollaborativeEditing()) {
				// User is editing alone: don't wait for response, continue editing immediately
				AscCommonExcel.applyFunction(callback, true);
				callback = undefined;
			}
			return callback;
		};
		CCollaborativeEditing.prototype._addCheckLock = function (lockInfo, callback) {
			if (false !== this.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeMine, false)) {
				// Self editing
				AscCommonExcel.applyFunction(callback, true);
				return c_oAscLockTypes.kLockTypeMine;
			} else if (false !== this.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeOther, false)) {
				// Someone is already editing the cell
				AscCommonExcel.applyFunction(callback, false);
				return c_oAscLockTypes.kLockTypeOther;
			}

			this.m_arrCheckLocks.push(lockInfo);
			return c_oAscLockTypes.kLockTypeNone;
		};
		CCollaborativeEditing.prototype.onStartCheckLock = function () {
			this.m_arrCheckLocks.length = 0;
		};
		CCollaborativeEditing.prototype.addCheckLock = function (oItem) {
			this.m_arrCheckLocks.push(oItem);
		};
		CCollaborativeEditing.prototype.onEndCheckLock = function (callback) {
			var t = this;
			if (this.m_arrCheckLocks.length > 0) {
				// Send request to server with list of elements
				this.handlers.trigger("askLock", this.m_arrCheckLocks, function (result) {
					t.onCallbackAskLock(result, callback);
				});

				if (undefined !== callback) {
					// Set global lock (only if we are not alone and waiting for response!)
					this.Set_GlobalLock(true);
				}
			} else {
				asc_applyFunction(callback, true);

				// Release global lock (for cell editing)
				this.m_bGlobalLockEditCell = false;
			}
		};

		CCollaborativeEditing.prototype.onCallbackAskLock = function (result, callback) {
			// Release global lock
			this.Set_GlobalLock(false);
			// Release global lock (for cell editing)
			this.m_bGlobalLockEditCell = false;

			if (result["lock"]) {
				// Iterate through the array and mark as locked by us
				var count = this.m_arrCheckLocks.length;
				for (var i = 0; i < count; ++i) {
					var oItem = this.m_arrCheckLocks[i];

					if (true !== oItem && false !== oItem) // compare by value and type strictly
					{
						var oNewLock = new CLock(oItem);
						oNewLock.setType(c_oAscLockTypes.kLockTypeMine);
						this.addUnlock2(oNewLock);
					}
				}

				asc_applyFunction(callback, true);
			} else if (result["error"]) {
				if (Asc.editor.frameManager.isLoadingOleEditor) {
					Asc.editor.sync_closeOleEditor();
				}

				asc_applyFunction(callback, false);
			}

			Asc.editor.frameManager.endLoadOleEditor();
		};
		CCollaborativeEditing.prototype.addUnlock = function (LockClass) {
			this.m_arrNeedUnlock.push(LockClass);
		};
		CCollaborativeEditing.prototype.addUnlock2 = function (Lock) {
			this.m_arrNeedUnlock2.push(Lock);
			this.handlers.trigger("updateDocumentCanSave");
		};

		CCollaborativeEditing.prototype.removeUnlock = function (Lock) {
			for (var i = 0; i < this.m_arrNeedUnlock.length; ++i)
				if (Lock.Element["guid"] === this.m_arrNeedUnlock[i].Element["guid"]) {
					this.m_arrNeedUnlock.splice(i, 1);
					return true;
				}
			return false;
		};

		CCollaborativeEditing.prototype.addChanges = function (oChanges) {
			this.m_arrChanges.push(oChanges);
		};

		// Returns - whether to send end action
		CCollaborativeEditing.prototype.applyChanges = function (oColor) {
			var t = this;
			var length = this.m_arrChanges.length;
			var oApi = Asc.editor;
			// Apply changes
			if (0 < length) {
				//splice to prevent double apply other changes in case of load fonts
				oApi.sendEvent("asc_onBeforeApplyChanges");
				var changes = t.m_arrChanges.splice(0, length);
				this.handlers.trigger("applyChanges", changes, function () {
					t.handlers.trigger("updateAfterApplyChanges");
				}, oColor);
				oApi.sendEvent("asc_onApplyChanges");

				return false;
			}

			return true;
		};

		CCollaborativeEditing.prototype.sendChanges = function (IsUserSave, isAfterAskSave, changesToSend) {
			// When not collaborative editing, nothing needs to be cleared, but sending is required.
			var bIsCollaborative = this.getCollaborativeEditing();

			var bCheckRedraw = false, bRedrawGraphicObjects = false, bUnlockDefName = false;
			var oLock = null;
			if (bIsCollaborative) {
				if (0 < this.m_arrNeedUnlock.length || 0 < this.m_arrNeedUnlock2.length) {
					bCheckRedraw = true;
					this.handlers.trigger("cleanSelection");
				}

				// Clear our changes
				while (0 < this.m_arrNeedUnlock2.length) {
					oLock = this.m_arrNeedUnlock2.shift();
					oLock.setType(c_oAscLockTypes.kLockTypeNone, false);

					var drawing = AscCommon.g_oTableId.Get_ById(oLock.Element["rangeOrObjectId"]);
					if (drawing && drawing.lockType !== c_oAscLockTypes.kLockTypeNone) {
						var bLocked = drawing.lockType !== c_oAscLockTypes.kLockTypeNone && drawing.lockType !== c_oAscLockTypes.kLockTypeMine;
						drawing.lockType = c_oAscLockTypes.kLockTypeNone;
						bRedrawGraphicObjects = true;
						if (drawing instanceof AscCommon.CCore) {
							if (bLocked) {
								Asc.editor && Asc.editor.sendEvent("asc_onLockCore", false);
							}
						}
					}
					if (!bUnlockDefName) {
						bUnlockDefName = this.handlers.trigger("checkDefNameLock", oLock);
					}

					this.handlers.trigger("releaseLocks", oLock.Element["guid"]);
				}

				// Clear applied changes from other users
				var nIndex = 0;
				var nCount = this.m_arrNeedUnlock.length;
				for (; nIndex < nCount; ++nIndex) {
					oLock = this.m_arrNeedUnlock[nIndex];
					if (c_oAscLockTypes.kLockTypeOther2 === oLock.getType()) {
						if (!this.handlers.trigger("checkCommentRemoveLock", oLock.Element)
							&& !this.handlers.trigger("checkCFRemoveLock", oLock.Element) &&
							!this.handlers.trigger("checkProtectedRangeRemoveLock", oLock.Element)) {
							drawing = AscCommon.g_oTableId.Get_ById(oLock.Element["rangeOrObjectId"]);
							if (drawing && drawing.lockType !== c_oAscLockTypes.kLockTypeNone) {
								var bLocked = drawing.lockType !== c_oAscLockTypes.kLockTypeNone && drawing.lockType !== c_oAscLockTypes.kLockTypeMine;
								drawing.lockType = c_oAscLockTypes.kLockTypeNone;
								bRedrawGraphicObjects = true;
								if (drawing instanceof AscCommon.CCore) {
									if (bLocked) {
										Asc.editor && Asc.editor.sendEvent("asc_onLockCore", false);
									}
								}
							}
							if (!bUnlockDefName) {
								bUnlockDefName = this.handlers.trigger("checkDefNameLock", oLock);
							}
						}

						this.m_arrNeedUnlock.splice(nIndex, 1);
						--nIndex;
						--nCount;
					}
				}
			}

			// Send changes to server
			this.handlers.trigger("sendChanges", this.getRecalcIndexSave(this.m_oRecalcIndexColumns), this.getRecalcIndexSave(this.m_oRecalcIndexRows), isAfterAskSave, changesToSend);

			if (bIsCollaborative) {
				// Recalculate locks from other users
				this._recalcLockArrayOthers();

				// Clear our changes (remove array of added rows/columns)
				delete this.m_oInsertColumns;
				delete this.m_oInsertRows;
				this.m_oInsertColumns = {};
				this.m_oInsertRows = {};
				// Clear our recalculation indexes
				this.clearRecalcIndex();

				// Clear Undo/Redo
				AscCommon.History.Clear_Redo();
				AscCommon.History.Clear();

				// Redraw
				if (bCheckRedraw) {
					this.handlers.trigger("drawSelection");
					this.handlers.trigger("drawFrozenPaneLines");
					this.handlers.trigger("updateAllSheetsLock");
					this.handlers.trigger("showComments");
				}

				if (bCheckRedraw || bRedrawGraphicObjects)
					this.handlers.trigger("showDrawingObjects");

//                if(bUnlockDefName){
				this.handlers.trigger("unlockDefName");
//                }

				this.handlers.trigger("updateAllLayoutsLock");
				this.handlers.trigger("asc_onLockPrintArea");
				this.handlers.trigger("updateAllHeaderFooterLock");
				this.handlers.trigger("onUpdateAllPrintScaleLock");
				this.handlers.trigger("updateAllSheetViewLock");

				this.handlers.trigger("unlockCF");
				this.handlers.trigger("unlockProtectedRange");

				this.handlers.trigger("unlockUserProtectedRanges");

				if (0 === this.m_nUseType)
					this.m_nUseType = 1;
			} else {
				// Update the last save point in history
				AscCommon.History.Reset_SavedIndex(IsUserSave);
			}

			Asc.editor.wb.Document_UpdateUndoRedoState();
		};

		CCollaborativeEditing.prototype.getRecalcIndexSave = function (oRecalcIndex) {
			var bHasIndex = false;
			var result = {};
			var element = null;
			for (var sheetId in oRecalcIndex) {
				if (!oRecalcIndex.hasOwnProperty(sheetId))
					continue;
				result[sheetId] = {"_arrElements": []};
				for (var i = 0, length = oRecalcIndex[sheetId]._arrElements.length; i < length; ++i) {
					bHasIndex = true;
					element = oRecalcIndex[sheetId]._arrElements[i];
					result[sheetId]["_arrElements"].push({
						"_recalcType": element._recalcType,
						"_position": element._position, "_count": element._count,
						"m_bIsSaveIndex": element.m_bIsSaveIndex
					});
				}
			}

			return bHasIndex ? result : null;
		};

		CCollaborativeEditing.prototype.S4 = function () {
			return (((1 + Math.random()) * 0x10000) | 0).toString(16).substring(1);
		};
		CCollaborativeEditing.prototype.createGUID = function () {
			return (this.S4() + this.S4() + "-" + this.S4() + "-" + this.S4() + "-" + this.S4() + "-" + this.S4() + this.S4() + this.S4());
		};

		CCollaborativeEditing.prototype.getLockInfo = function (typeElem, subType, sheetId, info) {
			var oLockInfo = new AscCommonExcel.asc_CLockInfo();
			oLockInfo["sheetId"] = sheetId;
			oLockInfo["type"] = typeElem;
			oLockInfo["subType"] = subType;
			oLockInfo["guid"] = this.createGUID();
			oLockInfo["rangeOrObjectId"] = info;
			return oLockInfo;
		};

		CCollaborativeEditing.prototype.getLockByElem = function (element, type) {
			var arrayElements = (c_oAscLockTypes.kLockTypeMine === type) ? this.m_arrNeedUnlock2 : this.m_arrNeedUnlock;
			for (var i = 0; i < arrayElements.length; ++i)
				if (element["guid"] === arrayElements[i].Element["guid"])
					return arrayElements[i];
			return null;
		};

		/**
		 * Check lock for element
		 * @param {asc_CLockInfo} element  element to check lock
		 * @param {c_oAscLockTypes} type ourselves(kLockTypeMine) or someone else
		 * @param {Boolean} bCheckOnlyLockAll check only lock for entire sheet properties (or only check if sheet is deleted, not just locked)
		 */
		CCollaborativeEditing.prototype.getLockIntersection = function (element, type, bCheckOnlyLockAll) {
			var arrayElements = (c_oAscLockTypes.kLockTypeMine === type) ? this.m_arrNeedUnlock2 : this.m_arrNeedUnlock;
			var oUnlockElement = null, rangeTmp1, rangeTmp2;
			for (var i = 0; i < arrayElements.length; ++i) {
				oUnlockElement = arrayElements[i].Element;
				if (c_oAscLockTypeElem.Sheet === element["type"] && element["type"] === oUnlockElement["type"]) {
					// Check only for sheet deletion (if checking for ourselves, don't exit immediately since we need to check locks from other elements)
					if ((c_oAscLockTypes.kLockTypeMine !== type && false === bCheckOnlyLockAll) ||
						element["sheetId"] === oUnlockElement["sheetId"]) {
						// If someone locked a sheet, no one else can lock sheets (otherwise all sheets could be deleted)
						return arrayElements[i];
					}
				}
				if (element["sheetId"] !== oUnlockElement["sheetId"])
					continue;

				if (null !== element["subType"] && null !== oUnlockElement["subType"])
					return arrayElements[i];

				// Don't consider lock from ChangeProperties (only if it's not a sheet lock)
				if (true === bCheckOnlyLockAll ||
					(c_oAscLockTypeElemSubType.ChangeProperties === oUnlockElement["subType"]
						&& c_oAscLockTypeElem.Sheet !== element["type"]))
					continue;

				if (element["type"] === oUnlockElement["type"]) {
					if (element["type"] === c_oAscLockTypeElem.Object) {
						if (element["rangeOrObjectId"] === oUnlockElement["rangeOrObjectId"])
							return arrayElements[i];
					} else if (element["type"] === c_oAscLockTypeElem.Range) {
						// Don't consider lock from Insert
						if (c_oAscLockTypes.kLockTypeMine === type || c_oAscLockTypeElemSubType.InsertRows === oUnlockElement["subType"] || c_oAscLockTypeElemSubType.InsertColumns === oUnlockElement["subType"])
							continue;
						rangeTmp1 = oUnlockElement["rangeOrObjectId"];
						rangeTmp2 = element["rangeOrObjectId"];
						if (rangeTmp2["c1"] > rangeTmp1["c2"] || rangeTmp2["c2"] < rangeTmp1["c1"] || rangeTmp2["r1"] > rangeTmp1["r2"] || rangeTmp2["r2"] < rangeTmp1["r1"])
							continue;
						return arrayElements[i];
					}
				} else if (oUnlockElement["type"] === c_oAscLockTypeElem.Sheet ||
					(element["type"] === c_oAscLockTypeElem.Sheet && c_oAscLockTypes.kLockTypeMine !== type)) {
					// If someone already locked the sheet or we're trying to lock it ourselves and checking for other locks
					return arrayElements[i];
				}
			}
			return false;
		};

		CCollaborativeEditing.prototype.getLockElem = function (typeElem, type, sheetId) {
			var arrayElements = (c_oAscLockTypes.kLockTypeMine === type) ? this.m_arrNeedUnlock2 : this.m_arrNeedUnlock;
			var count = arrayElements.length;
			var element = null, oRangeOrObjectId = null;
			var result = [];
			var c1, c2, r1, r2;

			if (!this.m_oRecalcIndexColumns.hasOwnProperty(sheetId)) {
				this.m_oRecalcIndexColumns[sheetId] = new CRecalcIndex();
			}
			if (!this.m_oRecalcIndexRows.hasOwnProperty(sheetId)) {
				this.m_oRecalcIndexRows[sheetId] = new CRecalcIndex();
			}

			for (var i = 0; i < count; ++i) {
				element = arrayElements[i].Element;
				if (element["sheetId"] !== sheetId || element["type"] !== typeElem)
					continue;

				// Don't display lock status for rows/columns deleted by current user (nothing to display anymore)
				if (c_oAscLockTypes.kLockTypeMine === type && c_oAscLockTypeElem.Range === typeElem &&
					(c_oAscLockTypeElemSubType.DeleteColumns === element["subType"] ||
						c_oAscLockTypeElemSubType.DeleteRows === element["subType"]))
					continue;
				// Don't display lock status for rows/columns added by another user (nothing to display yet)
				if (c_oAscLockTypeElem.Range === typeElem &&
					(c_oAscLockTypeElemSubType.InsertColumns === element["subType"] ||
						c_oAscLockTypeElemSubType.InsertRows === element["subType"]))
					continue;
				// Don't display lock range for lockAll (entire sheet)
				if (c_oAscLockTypeElemSubType.ChangeProperties === element["subType"])
					continue;

				oRangeOrObjectId = element["rangeOrObjectId"];
				// For range, need to recalculate considering deleted or added rows/columns
				if (c_oAscLockTypeElem.Range === typeElem) {
					// Don't recalculate for rows/columns deleted by another user
					if (c_oAscLockTypes.kLockTypeMine !== type && c_oAscLockTypeElem.Range === typeElem &&
						(c_oAscLockTypeElemSubType.DeleteColumns === element["subType"] ||
							c_oAscLockTypeElemSubType.DeleteRows === element["subType"])) {
						c1 = oRangeOrObjectId["c1"];
						c2 = oRangeOrObjectId["c2"];
						r1 = oRangeOrObjectId["r1"];
						r2 = oRangeOrObjectId["r2"];
					} else {
						c1 = this.m_oRecalcIndexColumns[sheetId].getLockOther(oRangeOrObjectId["c1"], type);
						c2 = this.m_oRecalcIndexColumns[sheetId].getLockOther(oRangeOrObjectId["c2"], type);
						r1 = this.m_oRecalcIndexRows[sheetId].getLockOther(oRangeOrObjectId["r1"], type);
						r2 = this.m_oRecalcIndexRows[sheetId].getLockOther(oRangeOrObjectId["r2"], type);
					}
					if (null === c1 || null === c2 || null === r1 || null === r2)
						continue;

					oRangeOrObjectId = new Asc.Range(c1, r1, c2, r2);
				}

				result.push(oRangeOrObjectId);
			}

			return result;
		};

		CCollaborativeEditing.prototype.getLockCellsMe = function (sheetId) {
			return this.getLockElem(c_oAscLockTypeElem.Range, c_oAscLockTypes.kLockTypeMine, sheetId);
		};
		CCollaborativeEditing.prototype.getLockCellsOther = function (sheetId) {
			return this.getLockElem(c_oAscLockTypeElem.Range, c_oAscLockTypes.kLockTypeOther, sheetId);
		};
		CCollaborativeEditing.prototype.getLockObjectsMe = function (sheetId) {
			return this.getLockElem(c_oAscLockTypeElem.Object, c_oAscLockTypes.kLockTypeMine, sheetId);
		};
		CCollaborativeEditing.prototype.getLockObjectsOther = function (sheetId) {
			return this.getLockElem(c_oAscLockTypeElem.Object, c_oAscLockTypes.kLockTypeOther, sheetId);
		};
		/**
		 * Check lock for entire sheet
		 * @param {Number} sheetId  element to check lock
		 * @return {Asc.c_oAscMouseMoveLockedObjectType} oLockedObjectType
		 */
		CCollaborativeEditing.prototype.isLockAllOther = function (sheetId) {
			var arrayElements = this.m_arrNeedUnlock;
			var count = arrayElements.length;
			var element = null;
			var oLockedObjectType = Asc.c_oAscMouseMoveLockedObjectType.None;

			for (var i = 0; i < count; ++i) {
				element = arrayElements[i].Element;
				if (element["sheetId"] === sheetId) {
					if (element["type"] === c_oAscLockTypeElem.Sheet) {
						oLockedObjectType = Asc.c_oAscMouseMoveLockedObjectType.Sheet;
						break;
					} else if (element["type"] === c_oAscLockTypeElem.Range && null !== element["subType"])
						oLockedObjectType = Asc.c_oAscMouseMoveLockedObjectType.TableProperties;
				}
			}
			return oLockedObjectType;
		};

		CCollaborativeEditing.prototype._recalcLockArray = function (typeLock, oRecalcIndexColumns, oRecalcIndexRows) {
			var arrayElements = (c_oAscLockTypes.kLockTypeMine === typeLock) ? this.m_arrNeedUnlock2 : this.m_arrNeedUnlock;
			var count = arrayElements.length;
			var element = null, oRangeOrObjectId = null;
			var i;
			var sheetId = -1;

			for (i = 0; i < count; ++i) {
				element = arrayElements[i].Element;
				// No need to recalculate indexes for deletion
				if (c_oAscLockTypeElem.Range !== element["type"] ||
					c_oAscLockTypeElemSubType.InsertColumns === element["subType"] ||
					c_oAscLockTypeElemSubType.InsertRows === element["subType"] ||
					c_oAscLockTypeElemSubType.DeleteColumns === element["subType"] ||
					c_oAscLockTypeElemSubType.DeleteRows === element["subType"])
					continue;
				sheetId = element["sheetId"];

				oRangeOrObjectId = element["rangeOrObjectId"];

				if (oRecalcIndexColumns && oRecalcIndexColumns.hasOwnProperty(sheetId)) {
					// Column recalculation
					oRangeOrObjectId["c1"] = oRecalcIndexColumns[sheetId].getLockMe(oRangeOrObjectId["c1"]);
					oRangeOrObjectId["c2"] = oRecalcIndexColumns[sheetId].getLockMe(oRangeOrObjectId["c2"]);
				}
				if (oRecalcIndexRows && oRecalcIndexRows.hasOwnProperty(sheetId)) {
					// Row recalculation
					oRangeOrObjectId["r1"] = oRecalcIndexRows[sheetId].getLockMe(oRangeOrObjectId["r1"]);
					oRangeOrObjectId["r2"] = oRecalcIndexRows[sheetId].getLockMe(oRangeOrObjectId["r2"]);
				}
			}
		};
		// Recalculation only for other users' Locks when saving on client that added/deleted rows or columns
		CCollaborativeEditing.prototype._recalcLockArrayOthers = function () {
			var typeLock = c_oAscLockTypes.kLockTypeOther;
			var arrayElements = (c_oAscLockTypes.kLockTypeMine === typeLock) ? this.m_arrNeedUnlock2 : this.m_arrNeedUnlock;
			var count = arrayElements.length;
			var element = null, oRangeOrObjectId = null;
			var i;
			var sheetId = -1;

			for (i = 0; i < count; ++i) {
				element = arrayElements[i].Element;
				if (c_oAscLockTypeElem.Range !== element["type"] ||
					c_oAscLockTypeElemSubType.InsertColumns === element["subType"] ||
					c_oAscLockTypeElemSubType.InsertRows === element["subType"])
					continue;
				sheetId = element["sheetId"];

				oRangeOrObjectId = element["rangeOrObjectId"];

				if (this.m_oRecalcIndexColumns.hasOwnProperty(sheetId)) {
					// Column recalculation
					oRangeOrObjectId["c1"] = this.m_oRecalcIndexColumns[sheetId].getLockOther(oRangeOrObjectId["c1"]);
					oRangeOrObjectId["c2"] = this.m_oRecalcIndexColumns[sheetId].getLockOther(oRangeOrObjectId["c2"]);
				}
				if (this.m_oRecalcIndexRows.hasOwnProperty(sheetId)) {
					// Row recalculation
					oRangeOrObjectId["r1"] = this.m_oRecalcIndexRows[sheetId].getLockOther(oRangeOrObjectId["r1"]);
					oRangeOrObjectId["r2"] = this.m_oRecalcIndexRows[sheetId].getLockOther(oRangeOrObjectId["r2"]);
				}
			}
		};

		CCollaborativeEditing.prototype.addRecalcIndex = function (type, oRecalcIndex) {
			if (null == oRecalcIndex)
				return null;
			var nIndex = 0;
			var nRecalcType = c_oAscRecalcIndexTypes.RecalcIndexAdd;
			var oRecalcIndexElement = null;
			var oRecalcIndexResult = {};

			var oRecalcIndexTmp = ("0" === type) ? this.m_oRecalcIndexColumns : this.m_oRecalcIndexRows;
			for (var sheetId in oRecalcIndex) {
				if (oRecalcIndex.hasOwnProperty(sheetId)) {
					if (!oRecalcIndexTmp.hasOwnProperty(sheetId)) {
						oRecalcIndexTmp[sheetId] = new CRecalcIndex();
					}
					if (!oRecalcIndexResult.hasOwnProperty(sheetId)) {
						oRecalcIndexResult[sheetId] = new CRecalcIndex();
					}
					for (; nIndex < oRecalcIndex[sheetId]["_arrElements"].length; ++nIndex) {
						oRecalcIndexElement = oRecalcIndex[sheetId]["_arrElements"][nIndex];
						if (true === oRecalcIndexElement["m_bIsSaveIndex"])
							continue;
						nRecalcType = (c_oAscRecalcIndexTypes.RecalcIndexAdd === oRecalcIndexElement["_recalcType"]) ?
							c_oAscRecalcIndexTypes.RecalcIndexRemove : c_oAscRecalcIndexTypes.RecalcIndexAdd;
						oRecalcIndexTmp[sheetId].add(nRecalcType, oRecalcIndexElement["_position"],
							oRecalcIndexElement["_count"], /*bIsSaveIndex*/true);
						// Duplicate for returning result (we need to recalculate only by the last index
						oRecalcIndexResult[sheetId].add(nRecalcType, oRecalcIndexElement["_position"],
							oRecalcIndexElement["_count"], /*bIsSaveIndex*/true);
					}
				}
			}

			return oRecalcIndexResult;
		};

		// Undo for adding/deleting columns
		CCollaborativeEditing.prototype.undoCols = function (sheetId, count) {
			if (!this.m_oRecalcIndexColumns.hasOwnProperty(sheetId))
				return;
			this.m_oRecalcIndexColumns[sheetId].remove(count);
		};
		// Undo for adding/deleting rows
		CCollaborativeEditing.prototype.undoRows = function (sheetId, count) {
			if (!this.m_oRecalcIndexRows.hasOwnProperty(sheetId))
				return;
			this.m_oRecalcIndexRows[sheetId].remove(count);
		};

		CCollaborativeEditing.prototype.removeCols = function (sheetId, position, count) {
			if (!this.m_oRecalcIndexColumns.hasOwnProperty(sheetId)) {
				this.m_oRecalcIndexColumns[sheetId] = new CRecalcIndex();
			}
			this.m_oRecalcIndexColumns[sheetId].add(c_oAscRecalcIndexTypes.RecalcIndexRemove, position,
				count, /*bIsSaveIndex*/false);
		};
		CCollaborativeEditing.prototype.addCols = function (sheetId, position, count) {
			if (!this.m_oRecalcIndexColumns.hasOwnProperty(sheetId)) {
				this.m_oRecalcIndexColumns[sheetId] = new CRecalcIndex();
			}
			this.m_oRecalcIndexColumns[sheetId].add(c_oAscRecalcIndexTypes.RecalcIndexAdd, position,
				count, /*bIsSaveIndex*/false);
		};
		CCollaborativeEditing.prototype.removeRows = function (sheetId, position, count) {
			if (!this.m_oRecalcIndexRows.hasOwnProperty(sheetId)) {
				this.m_oRecalcIndexRows[sheetId] = new CRecalcIndex();
			}
			this.m_oRecalcIndexRows[sheetId].add(c_oAscRecalcIndexTypes.RecalcIndexRemove, position,
				count, /*bIsSaveIndex*/false);
		};
		CCollaborativeEditing.prototype.addRows = function (sheetId, position, count) {
			if (!this.m_oRecalcIndexRows.hasOwnProperty(sheetId)) {
				this.m_oRecalcIndexRows[sheetId] = new CRecalcIndex();
			}
			this.m_oRecalcIndexRows[sheetId].add(c_oAscRecalcIndexTypes.RecalcIndexAdd, position,
				count, /*bIsSaveIndex*/false);
		};
		CCollaborativeEditing.prototype.addColsRange = function (sheetId, range) {
			if (!this.m_oInsertColumns.hasOwnProperty(sheetId)) {
				this.m_oInsertColumns[sheetId] = [];
			}
			var arrInsertColumns = this.m_oInsertColumns[sheetId];
			// Before adding, need to shift existing ones
			var countCols = range.c2 - range.c1 + 1;
			var isAddNewRange = true;
			for (var i = 0; i < arrInsertColumns.length; ++i) {
				if (arrInsertColumns[i].c1 > range.c1) {
					arrInsertColumns[i].c1 += countCols;
					arrInsertColumns[i].c2 += countCols;
				} else if (arrInsertColumns[i].c1 <= range.c1 && arrInsertColumns[i].c2 >= range.c1) {
					arrInsertColumns[i].c2 += countCols;
					isAddNewRange = false;
				}
			}
			if (isAddNewRange)
				arrInsertColumns.push(range);
		};
		CCollaborativeEditing.prototype.addRowsRange = function (sheetId, range) {
			if (!this.m_oInsertRows.hasOwnProperty(sheetId)) {
				this.m_oInsertRows[sheetId] = [];
			}
			var arrInsertRows = this.m_oInsertRows[sheetId];
			// Before adding, need to shift existing ones
			var countRows = range.r2 - range.r1 + 1;
			var isAddNewRange = true;
			for (var i = 0; i < arrInsertRows.length; ++i) {
				if (arrInsertRows[i].r1 > range.r1) {
					arrInsertRows[i].r1 += countRows;
					arrInsertRows[i].r2 += countRows;
				} else if (arrInsertRows[i].r1 <= range.r1 && arrInsertRows[i].r2 >= range.r1) {
					arrInsertRows[i].r2 += countRows;
					isAddNewRange = false;
				}
			}
			if (isAddNewRange)
				arrInsertRows.push(range);
		};
		CCollaborativeEditing.prototype.removeColsRange = function (sheetId, range) {
			if (!this.m_oInsertColumns.hasOwnProperty(sheetId))
				return;
			var arrInsertColumns = this.m_oInsertColumns[sheetId];
			// Need to remove columns that fall within the range
			var countCols = range.c2 - range.c1 + 1;
			for (var i = 0; i < arrInsertColumns.length; ++i) {
				if (arrInsertColumns[i].c1 > range.c2) {
					// To the right of the range being deleted
					arrInsertColumns[i].c1 -= countCols;
					arrInsertColumns[i].c2 -= countCols;
				} else if (arrInsertColumns[i].c1 >= range.c1 && arrInsertColumns[i].c2 <= range.c2) {
					// Completely included in the range being deleted
					arrInsertColumns.splice(i, 1);
					i -= 1;
				} else if (arrInsertColumns[i].c1 >= range.c1 && arrInsertColumns[i].c1 <= range.c2 && arrInsertColumns[i].c2 > range.c2) {
					// Partial inclusion of range start
					arrInsertColumns[i].c1 = range.c2 + 1;
					arrInsertColumns[i].c1 -= countCols;
					arrInsertColumns[i].c2 -= countCols;
				} else if (arrInsertColumns[i].c1 < range.c1 && arrInsertColumns[i].c2 >= range.c1 && arrInsertColumns[i].c2 <= range.c2) {
					// Partial inclusion of range end
					arrInsertColumns[i].c2 = range.c1 - 1;
				} else if (arrInsertColumns[i].c1 < range.c1 && arrInsertColumns[i].c2 > range.c2) {
					// Range being deleted is inside our range
					arrInsertColumns[i].c2 -= countCols;
				}
			}
		};
		CCollaborativeEditing.prototype.removeRowsRange = function (sheetId, range) {
			if (!this.m_oInsertRows.hasOwnProperty(sheetId))
				return;
			var arrInsertRows = this.m_oInsertRows[sheetId];
			// Need to remove rows that fall within the range
			var countRows = range.r2 - range.r1 + 1;
			for (var i = 0; i < arrInsertRows.length; ++i) {
				if (arrInsertRows[i].r1 > range.r2) {
					// Below the range being deleted
					arrInsertRows[i].r1 -= countRows;
					arrInsertRows[i].r2 -= countRows;
				} else if (arrInsertRows[i].r1 >= range.r1 && arrInsertRows[i].r2 <= range.r2) {
					// Completely included in the range being deleted
					arrInsertRows.splice(i, 1);
					i -= 1;
				} else if (arrInsertRows[i].r1 >= range.r1 && arrInsertRows[i].r1 <= range.r2 && arrInsertRows[i].r2 > range.r2) {
					// Partial inclusion of range start
					arrInsertRows[i].r1 = range.r2 + 1;
					arrInsertRows[i].r1 -= countRows;
					arrInsertRows[i].r2 -= countRows;
				} else if (arrInsertRows[i].r1 < range.r1 && arrInsertRows[i].r2 >= range.r1 && arrInsertRows[i].r2 <= range.r2) {
					// Partial inclusion of range end
					arrInsertRows[i].r2 = range.r1 - 1;
				} else if (arrInsertRows[i].r1 < range.r1 && arrInsertRows[i].r2 > range.r2) {
					// Range being deleted is inside our range
					arrInsertRows[i].r2 -= countRows;
				}
			}
		};
		CCollaborativeEditing.prototype.isIntersectionInCols = function (sheetId, col) {
			if (!this.m_oInsertColumns.hasOwnProperty(sheetId))
				return false;
			var arrInsertColumns = this.m_oInsertColumns[sheetId];
			for (var i = 0; i < arrInsertColumns.length; ++i) {
				if (arrInsertColumns[i].c1 <= col && col <= arrInsertColumns[i].c2)
					return true;
			}
			return false;
		};
		CCollaborativeEditing.prototype.isIntersectionInRows = function (sheetId, row) {
			if (!this.m_oInsertRows.hasOwnProperty(sheetId))
				return false;
			var arrInsertRows = this.m_oInsertRows[sheetId];
			for (var i = 0; i < arrInsertRows.length; ++i) {
				if (arrInsertRows[i].r1 <= row && row <= arrInsertRows[i].r2)
					return true;
			}
			return false;
		};
		CCollaborativeEditing.prototype.getArrayInsertColumnsBySheetId = function (sheetId) {
			if (!this.m_oInsertColumns.hasOwnProperty(sheetId))
				return [];

			return this.m_oInsertColumns[sheetId];
		};
		CCollaborativeEditing.prototype.getArrayInsertRowsBySheetId = function (sheetId) {
			if (!this.m_oInsertRows.hasOwnProperty(sheetId))
				return [];

			return this.m_oInsertRows[sheetId];
		};
		CCollaborativeEditing.prototype.getLockMeColumn = function (sheetId, col) {
			if (!this.m_oRecalcIndexColumns.hasOwnProperty(sheetId))
				return col;
			return this.m_oRecalcIndexColumns[sheetId].getLockMe(col);
		};
		CCollaborativeEditing.prototype.getLockMeRow = function (sheetId, row) {
			if (!this.m_oRecalcIndexRows.hasOwnProperty(sheetId))
				return row;
			return this.m_oRecalcIndexRows[sheetId].getLockMe(row);
		};
		// Only when column changes from other users (for recalculation)
		CCollaborativeEditing.prototype.getLockMeColumn2 = function (sheetId, col) {
			if (!this.m_oRecalcIndexColumns.hasOwnProperty(sheetId))
				return col;
			return this.m_oRecalcIndexColumns[sheetId].getLockMe2(col);
		};
		// Only when row changes from other users (for recalculation)
		CCollaborativeEditing.prototype.getLockMeRow2 = function (sheetId, row) {
			if (!this.m_oRecalcIndexRows.hasOwnProperty(sheetId))
				return row;
			return this.m_oRecalcIndexRows[sheetId].getLockMe2(row);
		};
		// Only for accepting changes from other users! (for recalculation only when saving)
		CCollaborativeEditing.prototype.getLockOtherColumn2 = function (sheetId, col) {
			if (!this.m_oRecalcIndexColumns.hasOwnProperty(sheetId))
				return col;
			return this.m_oRecalcIndexColumns[sheetId].getLockSaveOther(col);
		};
		// Only for accepting changes from other users! (for recalculation only when saving)
		CCollaborativeEditing.prototype.getLockOtherRow2 = function (sheetId, row) {
			if (!this.m_oRecalcIndexRows.hasOwnProperty(sheetId))
				return row;
			return this.m_oRecalcIndexRows[sheetId].getLockSaveOther(row);
		};
		CCollaborativeEditing.prototype.checkObjectsLock = function (aObjectId, callback) {
			var oApi = Asc.editor;
			var bCanNotEdit = this.getGlobalLock() || oApi && !oApi.canEdit();
			if (bCanNotEdit) {
				if (callback) {
					callback(false, true);
				}
				return false;
			}
			if (oApi) {
				oApi.incrementCounterLongAction();
			}
			const oInkDrawerState = oApi.inkDrawer.getState();
			var callbackEx = function (result, sync) {
				if (oApi) {
					oApi.decrementCounterLongAction();
				}
				if (callback) {
					oApi.inkDrawer.restoreState(oInkDrawerState);
					callback(result, sync);
					oApi.inkDrawer.restoreState(oInkDrawerState);
				}
			};
			var bRet = true;
			if (!aObjectId.length) {
				asc_applyFunction(callbackEx, true, true);
				return bRet;
			}
			this.onStartCheckLock();
			for (var i = 0; i < aObjectId.length; i++) {
				var lockInfo = this.getLockInfo(AscCommonExcel.c_oAscLockTypeElem.Object, null, "", aObjectId[i]);
				if (false === this.getCollaborativeEditing()) {
					asc_applyFunction(callbackEx, true, true);
					callbackEx = undefined;
				}
				if (false !== this.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeMine, false)) {
					continue;
				} else if (false !== this.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeOther, false)) {
					asc_applyFunction(callbackEx, false);
					return false;
				}
				this.addCheckLock(lockInfo);
			}
			this.onEndCheckLock(callbackEx);
			return bRet;
		};
		CCollaborativeEditing.prototype.Add_ForeignCursor = function (UserId, DocumentPos, UserShortId) {
			var isEqual = function (val1, val2) {
				var res = false;
				if (val1.isEdit === val2.isEdit && val1.sheetId === val2.sheetId) {
					if (val1.ranges.length === val2.ranges.length) {
						res = true;
						for (var i = 0; i < val1.ranges.length; i++) {
							if (!val1.ranges[i].isEqual(val2.ranges[i])) {
								res = false;
								break;
							}
						}
					}
				}
				return res;
			};

			if (this.m_aForeignCursorsData[UserId] && isEqual(this.m_aForeignCursorsData[UserId], DocumentPos)) {
				return false;
			}

			if (DocumentPos) {
				DocumentPos.shortId = UserShortId;
			}
			this.m_aForeignCursorsData[UserId] = DocumentPos;

			return true;
		};
		CCollaborativeEditing.prototype.Remove_ForeignCursor = function (UserId) {
			delete this.m_aForeignCursorsData[UserId];
		};
		CCollaborativeEditing.prototype.Remove_AllForeignCursors = function () {
			this.handlers.trigger("cleanSelection");
			for (var UserId in this.m_aForeignCursorsData) {
				this.Remove_ForeignCursor(UserId);
			}
			this.handlers.trigger("drawSelection");
		};
		CCollaborativeEditing.prototype.getForeignSelectIntersection = function (col, row) {
			var res = null;
			for (var i in this.m_aForeignCursorsData) {
				if (this.m_aForeignCursorsData[i]) {
					for (var j = 0; j < this.m_aForeignCursorsData[i].ranges.length; j++) {
						if (this.m_aForeignCursorsData[i].ranges[j].contains(col, row)) {
							res = this.m_aForeignCursorsData[i];
							res.userId = i;
						}
					}
				}
			}
			return res;
		};
		CCollaborativeEditing.prototype.PreUndo = function () {
			let res = AscCommon.CCollaborativeEditingBase.prototype.PreUndo.apply(this);

			this.oRedoObjectParam = new AscCommonExcel.RedoObjectParam();
			AscCommon.History.UndoRedoPrepare(this.oRedoObjectParam, false, true);
			//todo integrate into GetReverseOwnChanges
			if (this.CoHistory.OwnRanges.length > 0) {
				let range     = this.CoHistory.OwnRanges[this.CoHistory.OwnRanges.length - 1];
				let change = this.CoHistory.Changes[range.Position];
				if (change && change.oData && change.oData.snapshot) {
					this.oRedoObjectParam.snapshot = change.oData.snapshot
				}
			}
			return res;
		}
		CCollaborativeEditing.prototype.PostUndo = function (state, changes) {
			let Point = {Items: []}
			if (changes.length > 0) {
				//change is not the last because more may be added during adjustment
				let elem = changes.find(function(elem){
					if(elem && elem.oData && elem.oData.Point) {
						return true;
					}
				});
				if (elem) {
					Point = elem.oData.Point;
				}
			}
			//todo Apply_LinkData inside UndoRedoEnd
			AscCommon.CollaborativeEditing.Apply_LinkData();
			AscCommon.History.UndoRedoEnd(Point, this.oRedoObjectParam, false);

			AscCommon.CCollaborativeEditingBase.prototype.PostUndo.apply(this, arguments);
		}

		AscCommon.CCollaborativeHistory.prototype.CommuteRelated = function(oClass, oChange, nStartPosition)
		{
			//todo inherit because SDK merge is planned
			var arrChangesForProceed = this.Changes;
			for (var nIndex = nStartPosition, nOverallCount = arrChangesForProceed.length; nIndex < nOverallCount; ++nIndex) {
				var oOtherAction = arrChangesForProceed[nIndex];
				if (!oOtherAction) {
					continue;
				}
				if (true !== oOtherAction.IsReverted() && oChange.CommuteRelated && false === oChange.CommuteRelated(oChange, oOtherAction)) {
					return false;
				}
			}
			return true;
		};

		AscCommon.CCollaborativeHistory.prototype.saveChanges = function(changesToSend)
		{
			this.CoEditing.sendChanges(false, true, changesToSend);
		};

		/**
		 * Responsible for lock in collaborative editing
		 * -----------------------------------------------------------------------------
		 *
		 * @constructor
		 * @memberOf Asc
		 */
		function CLock(element) {
			this.Type = c_oAscLockTypes.kLockTypeNone;
			this.UserId = null;
			this.Element = element;

			this.init();

			return this;
		}

		CLock.prototype.init = function () {
		};
		CLock.prototype.getType = function () {
			return this.Type;
		};
		CLock.prototype.setType = function (newType) {
			if (newType === c_oAscLockTypes.kLockTypeNone)
				this.UserId = null;

			this.Type = newType;
		};

		CLock.prototype.Lock = function (bMine) {
			if (c_oAscLockTypes.kLockTypeNone === this.Type) {
				if (true === bMine)
					this.Type = c_oAscLockTypes.kLockTypeMine;
				else
					this.Type = c_oAscLockTypes.kLockTypeOther;
			}
		};

		CLock.prototype.setUserId = function (UserId) {
			this.UserId = UserId;
		};

		function CRecalcIndexElement(recalcType, position, bIsSaveIndex) {
			if (!(this instanceof CRecalcIndexElement)) {
				return new CRecalcIndexElement(recalcType, position, bIsSaveIndex);
			}

			this._recalcType	= recalcType;		// Type of changes (deletion or addition)
			this._position		= position;			// Position where changes occurred
			this._count			= 1;				// Consider all changes as elementary
			this.m_bIsSaveIndex	= !!bIsSaveIndex;	// These are indexes from changes of other users (which we haven't applied yet)

			return this;
		}

		// Recalculation for others
		CRecalcIndexElement.prototype.getLockOther = function (position, type) {
			var inc = (c_oAscRecalcIndexTypes.RecalcIndexAdd === this._recalcType) ? +1 : -1;
			if (position === this._position && c_oAscRecalcIndexTypes.RecalcIndexRemove === this._recalcType &&
				true === this.m_bIsSaveIndex) {
				// We haven't applied other users' changes yet (so no need to draw for insert)
				// RecalcIndexRemove (because we reverse for correct processing, RecalcIndexAdd
				// came from another user
				return null;
			} else if (position === this._position &&
				c_oAscRecalcIndexTypes.RecalcIndexRemove === this._recalcType &&
				c_oAscLockTypes.kLockTypeMine === type && false === this.m_bIsSaveIndex) {
				// For the user who deleted the column, no need to draw cells
				// that were previously locked in this column
				return null;
			} else if (position < this._position)
				return position;
			else
				return (position + inc);
		};
		// Recalculation for others (only for saving)
		CRecalcIndexElement.prototype.getLockSaveOther = function (position, type) {
			if (this.m_bIsSaveIndex)
				return position;

			var inc = (c_oAscRecalcIndexTypes.RecalcIndexAdd === this._recalcType) ? +1 : -1;
			if (position === this._position && c_oAscRecalcIndexTypes.RecalcIndexRemove === this._recalcType &&
				true === this.m_bIsSaveIndex) {
				// We haven't applied other users' changes yet (so no need to draw for insert)
				// RecalcIndexRemove (because we reverse for correct processing, RecalcIndexAdd
				// came from another user
				return null;
			} else if (position === this._position &&
				c_oAscRecalcIndexTypes.RecalcIndexRemove === this._recalcType &&
				c_oAscLockTypes.kLockTypeMine === type && false === this.m_bIsSaveIndex) {
				// For the user who deleted the column, no need to draw cells
				// that were previously locked in this column
				return null;
			} else if (position < this._position)
				return position;
			else
				return (position + inc);
		};
		// Recalculation for me
		CRecalcIndexElement.prototype.getLockMe = function (position) {
			var inc = (c_oAscRecalcIndexTypes.RecalcIndexAdd === this._recalcType) ? -1 : +1;
			if (position < this._position)
				return position;
			else
				return (position + inc);
		};
		// Only when changes from other users (for recalculation)
		CRecalcIndexElement.prototype.getLockMe2 = function (position) {
			var inc = (c_oAscRecalcIndexTypes.RecalcIndexAdd === this._recalcType) ? -1 : +1;
			if (true !== this.m_bIsSaveIndex || position < this._position)
				return position;
			else
				return (position + inc);
		};

		function CRecalcIndex() {
			if (!(this instanceof CRecalcIndex)) {
				return new CRecalcIndex();
			}

			this._arrElements = [];		// Array of CRecalcIndexElement

			return this;
		}

		CRecalcIndex.prototype.add = function (recalcType, position, count, bIsSaveIndex) {
			for (var i = 0; i < count; ++i)
				this._arrElements.push(new CRecalcIndexElement(recalcType, position, bIsSaveIndex));
		};
		// Removes from recalculation, for undo
		CRecalcIndex.prototype.remove = function (count) {
			for (var i = 0; i < count; ++i)
				this._arrElements.pop();
		};
		CRecalcIndex.prototype.clear = function () {
			this._arrElements.length = 0;
		};

		// Recalculation for others
		CRecalcIndex.prototype.getLockOther = function (position, type) {
			var newPosition = position;
			/*var count = this._arrElements.length;
			 for (var i = 0; i < count; ++i) {
			 newPosition = this._arrElements[i].getLockOther(newPosition, type);
			 if (null === newPosition)
			 break;
			 }*/

			var count = this._arrElements.length;
			if (0 >= count)
				return newPosition;
			// For recalculation when added by self - reverse order
			// For recalculation when added by someone else - direct order
			var bIsDirect = !this._arrElements[0].m_bIsSaveIndex;
			var i;
			if (bIsDirect) {
				for (i = 0; i < count; ++i) {
					newPosition = this._arrElements[i].getLockOther(newPosition, type);
					if (null === newPosition)
						break;
				}
			} else {
				for (i = count - 1; i >= 0; --i) {
					newPosition = this._arrElements[i].getLockOther(newPosition, type);
					if (null === newPosition)
						break;
				}
			}

			return newPosition;
		};
		// Recalculation for others (only for saving)
		CRecalcIndex.prototype.getLockSaveOther = function (position, type) {
			var newPosition = position;
			var count = this._arrElements.length;
			for (var i = 0; i < count; ++i) {
				newPosition = this._arrElements[i].getLockSaveOther(newPosition, type);
				if (null === newPosition)
					break;
			}

			return newPosition;
		};
		// Recalculation for me
		CRecalcIndex.prototype.getLockMe = function (position) {
			var newPosition = position;
			var count = this._arrElements.length;
			if (0 >= count)
				return newPosition;
			// For recalculation when added by self - reverse order
			// For recalculation when added by someone else - direct order
			var bIsDirect = this._arrElements[0].m_bIsSaveIndex;
			var i;
			if (bIsDirect) {
				for (i = 0; i < count; ++i) {
					newPosition = this._arrElements[i].getLockMe(newPosition);
					if (null === newPosition)
						break;
				}
			} else {
				for (i = count - 1; i >= 0; --i) {
					newPosition = this._arrElements[i].getLockMe(newPosition);
					if (null === newPosition)
						break;
				}
			}

			return newPosition;
		};
		// Only when changes from other users (for recalculation)
		CRecalcIndex.prototype.getLockMe2 = function (position) {
			var newPosition = position;
			var count = this._arrElements.length;
			if (0 >= count)
				return newPosition;
			// For recalculation when added by self - reverse order
			// For recalculation when added by someone else - direct order
			var bIsDirect = this._arrElements[0].m_bIsSaveIndex;
			var i;
			if (bIsDirect) {
				for (i = 0; i < count; ++i) {
					newPosition = this._arrElements[i].getLockMe2(newPosition);
					if (null === newPosition)
						break;
				}
			} else {
				for (i = count - 1; i >= 0; --i) {
					newPosition = this._arrElements[i].getLockMe2(newPosition);
					if (null === newPosition)
						break;
				}
			}

			return newPosition;
		};

		//----------------------------------------------------------export----------------------------------------------------
		window['Asc'] = window['Asc'] || {};
		window['AscCommonExcel'] = window['AscCommonExcel'] || {};
		window['AscCommonExcel'].CLock = CLock;

		window['AscCommonExcel'].CCollaborativeEditing = CCollaborativeEditing;
		window['Asc'].CRecalcIndexElement = CRecalcIndexElement;
		window['Asc'].CRecalcIndex = CRecalcIndex;
	}
)(window);
