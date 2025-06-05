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
	 * @param {jQuery} $
	 * @param {Window} window
	 * @param {undefined} undefined
	 */
	function ($, window, undefined) {

		var AscBrowser = AscCommon.AscBrowser;

		var asc = window["Asc"] ? window["Asc"] : (window["Asc"] = {});
		var asc_applyFunction = AscCommonExcel.applyFunction;
		var c_oTargetType = AscCommonExcel.c_oTargetType;

		/**
		 * Desktop event controller for WorkbookView
		 * -----------------------------------------------------------------------------
		 * @constructor
		 * @memberOf Asc
		 */
		function asc_CEventsController() {
			//----- declaration -----
			this.settings = {
				vscrollStep: 10,
				hscrollStep: 10,
				scrollTimeout: 20,
				wheelScrollLinesV: 3
			};

			this.view = undefined;
			this.widget = undefined;
			this.element = undefined;
			this.handlers = undefined;
			this.vsb = undefined;
			this.vsbApi = undefined;
			this.vsbMax = undefined;
			this.hsb = undefined;
			this.hsbApi = undefined;
			this.hsbMax = undefined;

			this.resizeTimerId = undefined;
			this.scrollTimerId = undefined;
			this.moveRangeTimerId = undefined;
			this.moveResizeRangeTimerId = undefined;
			this.fillHandleModeTimerId = undefined;
			this.enableKeyEvents = true;
			this.isSelectMode = false;
			this.hasCursor = false;
			this.hasFocus = false;
			this.skipKeyPress = undefined;
			this.targetInfo = undefined;
			this.isResizeMode = false;
			this.isResizeModeMove = false;

			// Autofill mode
			this.isFillHandleMode = false;
			this.isMoveRangeMode = false;
			this.isMoveResizeRange = false;
			// Pinned Areas Setup Mode
			this.frozenAnchorMode = false;

			// Click handler for graph objects
			this.clickCounter = new AscFormat.ClickCounter();
			this.isMousePressed = false;
			this.isShapeAction = false;
			this.isUpOnCanvas = false;

			// Was DblClick handled in onMouseDown event
			this.isDblClickInMouseDown = false;
			// Should I handle the browser event dblClick
			this.isDoBrowserDblClick = false;
			// Last coordinates, on MouseDown (for IE)
			this.mouseDownLastCord = null;
			//-----------------------

			this.vsbApiLockMouse = false;
			this.hsbApiLockMouse = false;

			// when you click on the collapse/expand button for a group of rows
			this.isRowGroup = false;

			this.smoothWheelCorrector = null;
			if (AscCommon.AscBrowser.isMacOs) {
				this.smoothWheelCorrector = new AscCommon.CMouseSmoothWheelCorrector(this, function (deltaX, deltaY) {

					if (deltaX) {
						deltaX = Math.sign(deltaX) * Math.ceil(Math.abs(deltaX / 3));
						this.scrollHorizontal(deltaX, null);
					}
					if (deltaY) {
						deltaY = Math.sign(deltaY) * Math.ceil(Math.abs(deltaY * this.settings.wheelScrollLinesV / 3));
						this.scrollVertical(deltaY, null);
					}

				});

				this.smoothWheelCorrector.setNormalDeltaActive(3);
			}

			this.lastTab = null;

			this.skipCellEditor = false;

			return this;
		}

		/**
		 * @param {AscCommonExcel.WorkbookView} view
		 * @param {Element} widgetElem
		 * @param {Element} canvasElem
		 * @param {Object} handlers  Event handlers (resize, scrollY, scrollX, changeSelection, ...)
		 */
		asc_CEventsController.prototype.init = function (view, widgetElem, canvasElem, handlers) {
			var self = this;
			this.view = view;
			this.widget = widgetElem;
			this.element = canvasElem;
			this.handlers = new AscCommonExcel.asc_CHandlersList(handlers);
			this._createScrollBars();

			if (Asc.editor.isEditOleMode) {
				return;
			}

			if (this.view.Api.isUseOldMobileVersion()) {
				/*Previously, resize events were called from the menu via the controller. Now the controller is not available in the menu, for resize we subscribe to the global resize from window.*/
				window.addEventListener("resize", function () {
					self._onWindowResize.apply(self, arguments);
				}, false);
				return this;
			}

			// initialize events
			if (window.addEventListener) {
				window.addEventListener("resize", function () {
					self._onWindowResize.apply(self, arguments);
				}, false);
				window.addEventListener(AscCommon.getPtrEvtName("move"), function () {
					return self._onWindowMouseMove.apply(self, arguments);
				}, false);
				window.addEventListener(AscCommon.getPtrEvtName("up"), function () {
					return self._onWindowMouseUp.apply(self, arguments);
				}, false);
				window.addEventListener(AscCommon.getPtrEvtName("leave"), function () {
					return self._onWindowMouseLeaveOut.apply(self, arguments);
				}, false);
				window.addEventListener(AscCommon.getPtrEvtName("out"), function () {
					return self._onWindowMouseLeaveOut.apply(self, arguments);
				}, false);
			}

			// prevent changing mouse cursor when 'mousedown' is occurred
			if (this.element.onselectstart) {
				this.element.onselectstart = function () {
					return false;
				};
			}

			if (this.element.addEventListener) {
				this.element.addEventListener(AscCommon.getPtrEvtName("down"), function () {
					return self._onMouseDown.apply(self, arguments);
				}, false);
				this.element.addEventListener(AscCommon.getPtrEvtName("up"), function () {
					return self._onMouseUp.apply(self, arguments);
				}, false);
				this.element.addEventListener(AscCommon.getPtrEvtName("move"), function () {
					return self._onMouseMove.apply(self, arguments);
				}, false);
				this.element.addEventListener(AscCommon.getPtrEvtName("leave"), function () {
					return self._onMouseLeave.apply(self, arguments);
				}, false);
				this.element.addEventListener("dblclick", function () {
					return self._onMouseDblClick.apply(self, arguments);
				}, false);
			}
			if (this.widget.addEventListener) {
				// https://developer.mozilla.org/en-US/docs/Web/Reference/Events/wheel
				// detect available wheel event
				var nameWheelEvent = (!AscCommon.AscBrowser.isMacOs && "onwheel" in document.createElement("div")) ? "wheel" :	// Modern browsers support "wheel"
					document.onmousewheel !== undefined ? "mousewheel" : 				// Webkit and IE support at least "mousewheel"
						"DOMMouseScroll";												// let's assume that remaining browsers are older Firefox

				this.widget.addEventListener(nameWheelEvent, function () {
					return self._onMouseWheel.apply(self, arguments);
				}, false);
				this.widget.addEventListener('contextmenu', function (e) {
					e.stopPropagation();
					e.preventDefault();
					return false;
				}, false)
			}

			// Cursor for graphic objects. We define mousedown and mouseup for text selection.
			var oShapeCursor = document.getElementById("id_target_cursor");
			if (null != oShapeCursor && oShapeCursor.addEventListener) {
				oShapeCursor.addEventListener(AscCommon.getPtrEvtName("down"), function () {
					return self._onMouseDown.apply(self, arguments);
				}, false);
				oShapeCursor.addEventListener(AscCommon.getPtrEvtName("up"), function () {
					return self._onMouseUp.apply(self, arguments);
				}, false);
				oShapeCursor.addEventListener(AscCommon.getPtrEvtName("move"), function () {
					return self._onMouseMove.apply(self, arguments);
				}, false);
				oShapeCursor.addEventListener(AscCommon.getPtrEvtName("leave"), function () {
					return self._onMouseLeave.apply(self, arguments);
				}, false);
			}

			return this;
		};

		asc_CEventsController.prototype.destroy = function () {
			return this;
		};

		/** @param flag {Boolean} */
		asc_CEventsController.prototype.enableKeyEventsHandler = function (flag) {
			this.enableKeyEvents = !!flag;
		};

		/** @return {Boolean} */
		asc_CEventsController.prototype.canEdit = function () {
			return this.view.canEdit();
		};

		/** @return {Boolean} */
		asc_CEventsController.prototype.getCellEditMode = function () {
			return this.view.isCellEditMode;
		};

		asc_CEventsController.prototype.setFocus = function (hasFocus) {
			this.hasFocus = !!hasFocus;
		};

		asc_CEventsController.prototype.gotFocus = function (hasFocus) {
			this.setFocus(hasFocus);
			this.handlers.trigger('gotFocus', this.hasFocus);
		};

		/** @return {Boolean} */
		asc_CEventsController.prototype.getFormulaEditMode = function () {
			return this.view.isFormulaEditMode;
		};

		asc_CEventsController.prototype.getSelectionDialogMode = function () {
			return this.view.selectionDialogMode;
		};

		asc_CEventsController.prototype.reinitScrollX = function (settings, pos, max, max2) {
			var step = this.settings.hscrollStep;
			this.hsbMax = Math.max(max * step, 1);
			settings.contentW = this.hsbMax - 1;
			this.hsbApi.Repos(settings, false, false, pos * step);
			this.hsbApi.maxScrollX2 = Math.max(max2 * step, 1);
		};
		asc_CEventsController.prototype.reinitScrollY = function (settings, pos, max, max2) {
			var step = this.settings.vscrollStep;
			this.vsbMax = Math.max(max * step, 1);
			settings.contentH = this.vsbMax - 1;
			this.vsbApi.Repos(settings, false, false, pos * step);
			this.vsbApi.maxScrollY2 = Math.max(max2 * step, 1);
		};

		/**
		 * @param {AscCommon.CellBase} delta
		 */
		asc_CEventsController.prototype.scroll = function (delta) {
			if (delta) {
				if (delta.col) {
					this.scrollHorizontal(delta.col);
				}
				if (delta.row) {
					this.scrollVertical(delta.row);
				}
			}
		};

		/**
		 * @param delta {Number}
		 * @param [event] {MouseEvent}
		 */
		asc_CEventsController.prototype.scrollVertical = function (delta, event) {
			if (window["NATIVE_EDITOR_ENJINE"])
				return;

			if (event && event.preventDefault)
				event.preventDefault();
			this.vsbApi.scrollByY(this.settings.vscrollStep * delta);
			return true;
		};

		/**
		 * @param delta {Number}
		 * @param [event] {MouseEvent}
		 */
		asc_CEventsController.prototype.scrollHorizontal = function (delta, event) {
			if (window["NATIVE_EDITOR_ENJINE"])
				return;

			if (event && event.preventDefault)
				event.preventDefault();
			this.hsbApi.scrollByX(this.settings.hscrollStep * delta);
			return true;
		};

		// We will do dblClick as in Excel
		asc_CEventsController.prototype.doMouseDblClick = function (event) {
			var t = this;
			var ctrlKey = !AscCommon.getAltGr(event) && (event.metaKey || event.ctrlKey);

			// The formula does not require exiting cell editing
			if (this.getFormulaEditMode() || this.getSelectionDialogMode()) {
				return true;
			}

			if (this.view.Api.isEditVisibleAreaOleEditor) {
				return true;
			}

			if (this.targetInfo && (this.targetInfo.target === AscCommonExcel.c_oTargetType.GroupRow ||
				this.targetInfo.target === AscCommonExcel.c_oTargetType.GroupCol)) {
				return false;
			}

			if (this.targetInfo && (this.targetInfo.target === c_oTargetType.MoveResizeRange ||
				this.targetInfo.target === c_oTargetType.MoveRange ||
				this.targetInfo.target === c_oTargetType.FilterObject ||
				this.targetInfo.target === c_oTargetType.TableSelectionChange))
				return true;

			if (t.getCellEditMode()) {
				if (!t.handlers.trigger("stopCellEditing")) {
					return true;
				}
			}

			var coord = t._getCoordinates(event);
			var graphicsInfo = t.handlers.trigger("getGraphicsInfo", coord.x, coord.y);
			if (graphicsInfo)
				return;

			setTimeout(function () {
				var coord = t._getCoordinates(event);
				t.handlers.trigger("mouseDblClick", coord.x, coord.y, event, function () {
					// We resized the column/row, not edited the cell. Update the cursor state
					t.handlers.trigger("updateWorksheet", coord.x, coord.y, ctrlKey,
						function (info) {
							t.targetInfo = info;
						});
				});
			}, 100);

			return true;
		};

		// We will show the cursor near the cell editor (only for dblClick)
		asc_CEventsController.prototype.showCellEditorCursor = function () {
			if (this.getCellEditMode()) {
				if (this.isDoBrowserDblClick) {
					this.isDoBrowserDblClick = false;
					this.handlers.trigger("showCellEditorCursor");
				}
			}
		};

		asc_CEventsController.prototype.createScrollSettings = function () {
			var settings = new AscCommon.ScrollSettings();

			settings.scrollBackgroundColor = GlobalSkin.ScrollBackgroundColor;
			settings.scrollBackgroundColorHover = GlobalSkin.ScrollBackgroundColor;
			settings.scrollBackgroundColorActive = GlobalSkin.ScrollBackgroundColor;

			settings.scrollerColor = GlobalSkin.ScrollerColor;
			settings.scrollerHoverColor = GlobalSkin.ScrollerHoverColor;
			settings.scrollerActiveColor = GlobalSkin.ScrollerActiveColor;

			settings.arrowColor = GlobalSkin.ScrollArrowColor;
			settings.arrowHoverColor = GlobalSkin.ScrollArrowHoverColor;
			settings.arrowActiveColor = GlobalSkin.ScrollArrowActiveColor;

			settings.strokeStyleNone = GlobalSkin.ScrollOutlineColor;
			settings.strokeStyleOver = GlobalSkin.ScrollOutlineHoverColor;
			settings.strokeStyleActive = GlobalSkin.ScrollOutlineActiveColor;

			settings.targetColor = GlobalSkin.ScrollerTargetColor;
			settings.targetHoverColor = GlobalSkin.ScrollerTargetHoverColor;
			settings.targetActiveColor = GlobalSkin.ScrollerTargetActiveColor;

			return settings;
		};

		asc_CEventsController.prototype.updateScrollSettings = function () {
			var opt = this.settings, settings;
			var ws = window["Asc"]["editor"].wb.getWorksheet();
			if (this.vsbApi) {
				settings = this.createScrollSettings();

				settings.vscrollStep = opt.vscrollStep;
				settings.hscrollStep = opt.hscrollStep;
				settings.wheelScrollLines = opt.wheelScrollLinesV;
				settings.isVerticalScroll = true;
				settings.isHorizontalScroll = false;
				this.vsbApi.canvasH = null;
				this.reinitScrollY(settings, ws.workbook.getSmoothScrolling() ? ws.getFirstVisibleRowSmoothScroll(true) : ws.getFirstVisibleRow(true), ws.getVerticalScrollRange(), ws.getVerticalScrollMax());
				this.vsbApi.settings = settings;
			}
			if (this.hsbApi) {
				settings = this.createScrollSettings();
				settings.vscrollStep = opt.vscrollStep;
				settings.hscrollStep = opt.hscrollStep;
				settings.isVerticalScroll = false;
				settings.isHorizontalScroll = true;
				this.hsbApi.canvasW = null;
				this.reinitScrollX(settings, ws.workbook.getSmoothScrolling() ? ws.getFirstVisibleColSmoothScroll(true) : ws.getFirstVisibleCol(true), ws.getHorizontalScrollRange(), ws.getHorizontalScrollMax());
				this.hsbApi.settings = settings;
			}
		};

		asc_CEventsController.prototype._createScrollBars = function () {
			var self = this, settings, opt = this.settings;

			// vertical scroll bar
			this.vsb = document.createElement('div');
			this.vsb.id = "ws-v-scrollbar";
			this.vsb.style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;

			//TODO test rtl
			/*if (window.rightToleft) {
				this.vsb.style.left = "0px";
				this.widget.prepend(this.vsb);
				this.widget.children[1].style.left = this.vsb.clientWidth + "px";
				this.widget.children[1].style.overflow = "visible"
			} else {*/
				this.widget.appendChild(this.vsb);
			//}

			if (!this.vsbApi) {
				settings = this.createScrollSettings();
				settings.vscrollStep = opt.vscrollStep;
				settings.hscrollStep = opt.hscrollStep;
				settings.wheelScrollLines = opt.wheelScrollLinesV;
				settings.isVerticalScroll = true;
				settings.isHorizontalScroll = false;

				this.vsbApi = new AscCommon.ScrollObject(this.vsb.id, settings);
				this.vsbApi.bind("scrollvertical", function (evt) {
					self.handlers.trigger("scrollY", evt.scrollPositionY / self.settings.vscrollStep, !self.vsbApi.scrollerMouseDown);
				});
				this.vsbApi.bind(AscCommon.getPtrEvtName("up"), function (evt) {
					if (self.vsbApi.scrollerMouseDown) {
						self.handlers.trigger('initRowsCount');
					}
				});
				this.vsbApi.onLockMouse = function (evt) {
					self.vsbApiLockMouse = true;
				};
				this.vsbApi.offLockMouse = function () {
					self.vsbApiLockMouse = false;
				};
			}

			// horizontal scroll bar
			this.hsb = document.createElement('div');
			this.hsb.id = "ws-h-scrollbar";
			this.hsb.style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
			this.widget.appendChild(this.hsb);

			if (!this.hsbApi) {
				settings = this.createScrollSettings();
				settings.vscrollStep = opt.vscrollStep;
				settings.hscrollStep = opt.hscrollStep;
				settings.isVerticalScroll = false;
				settings.isHorizontalScroll = true;

				this.hsbApi = new AscCommon.ScrollObject(this.hsb.id, settings);
				this.hsbApi.bind("scrollhorizontal", function (evt) {
					self.handlers.trigger("scrollX", evt.scrollPositionX / self.settings.hscrollStep, !self.hsbApi.scrollerMouseDown);
				});
				this.hsbApi.bind(AscCommon.getPtrEvtName("up"), function (evt) {
					if (self.hsbApi.scrollerMouseDown) {
						self.handlers.trigger('initColsCount');
					}
				});
				this.hsbApi.onLockMouse = function () {
					self.hsbApiLockMouse = true;
				};
				this.hsbApi.offLockMouse = function () {
					self.hsbApiLockMouse = false;
				};
			}

			if (!this.view.Api.isMobileVersion) {
				// right bottom corner
				var corner = document.createElement('div');
				corner.id = "ws-scrollbar-corner";
				corner.style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor
				this.widget.appendChild(corner);
			} else {
				this.hsb.style.zIndex = -10;
				this.hsb.style.right = 0;
				this.hsb.style.display = "none";
				this.vsb.style.zIndex = -10;
				this.vsb.style.bottom = 0;
				this.vsb.style.display = "none";
			}

			this.showVerticalScroll(this.view.getShowVerticalScroll());
			this.showHorizontalScroll(this.view.getShowHorizontalScroll());
		};

		/**
		 * @param event {MouseEvent}
		 * @param callback {Function}
		 * @private
		 */
		asc_CEventsController.prototype._continueChangeVisibleArea = function (event, callback) {
			var t = this;
			var coord = this._getCoordinates(event);

			this.handlers.trigger("changeVisibleArea", /*isStartPoint*/false, coord.x, coord.y, false,
				function (d) {
					t.scroll(d);

					asc_applyFunction(callback);
				});
		};

		/**
		 *
		 * @param event {MouseEvent}
		 */
		asc_CEventsController.prototype._continueChangeVisibleArea2 = function (event) {
			var t = this;
			var fn = function () {
				t._continueChangeVisibleArea2(event);
			};

			var callback = function () {
				if (t.isChangeVisibleAreaMode && !t.hasCursor) {
					t.visibleAreaSelectionTimerId = window.setTimeout(fn, t.settings.scrollTimeout);
				}
			}
			window.clearTimeout(t.visibleAreaSelectionTimerId);
			window.setTimeout(function () {
				if (t.isChangeVisibleAreaMode && !t.hasCursor) {
					t._continueChangeVisibleArea(event, callback);
				}
			}, 0);

		};

		/**
		 *
		 * @param event {MouseEvent}
		 * @param callback {Function}
		 */
		asc_CEventsController.prototype._startChangeVisibleArea = function (event, callback) {
			var coord = this._getCoordinates(event);
			this.handlers.trigger("changeVisibleArea", /*isStartPoint*/true, coord.x, coord.y, false, callback);
		};

		/**
		 * @param event {MouseEvent}
		 * @param callback {Function}
		 */
		asc_CEventsController.prototype._changeSelection = function (event, callback) {
			var t = this;
			var coord = this._getCoordinates(event);

			this.handlers.trigger("changeSelection", /*isStartPoint*/false, coord.x, coord.y, /*isCoord*/true, false,
				function (d) {
					t.scroll(d);

					asc_applyFunction(callback);
				});
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._changeSelection2 = function (event) {
			var t = this;

			var fn = function () {
				t._changeSelection2(event);
			};
			var callback = function () {
				if (t.isSelectMode && !t.hasCursor) {
					t.scrollTimerId = window.setTimeout(fn, t.settings.scrollTimeout);
				}
			};

			window.clearTimeout(t.scrollTimerId);
			t.scrollTimerId = window.setTimeout(function () {
				if (t.isSelectMode && !t.hasCursor) {
					t._changeSelection(event, callback);
				}
			}, 0);
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._moveRangeHandle2 = function (event) {
			var t = this;

			var fn = function () {
				t._moveRangeHandle2(event);
			};

			var callback = function () {
				if (t.isMoveRangeMode && !t.hasCursor) {
					t.moveRangeTimerId = window.setTimeout(fn, t.settings.scrollTimeout);
				}
			};

			window.clearTimeout(t.moveRangeTimerId);
			t.moveRangeTimerId = window.setTimeout(function () {
				if (t.isMoveRangeMode && !t.hasCursor) {
					t._moveRangeHandle(event, callback);
				}
			}, 0);
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._moveResizeRangeHandle2 = function (event) {
			var t = this;

			var fn = function () {
				t._moveResizeRangeHandle2(event);
			};

			var callback = function () {
				if (t.isMoveResizeRange && !t.hasCursor) {
					t.moveResizeRangeTimerId = window.setTimeout(fn, t.settings.scrollTimeout);
				}
			};

			window.clearTimeout(t.moveResizeRangeTimerId);
			t.moveResizeRangeTimerId = window.setTimeout(function () {
				if (t.isMoveResizeRange && !t.hasCursor) {
					t._moveResizeRangeHandle(event, t.targetInfo, callback);
				}
			}, 0);
		};

		/**
		 * End of selection
		 * @param event {MouseEvent}
		 */
		asc_CEventsController.prototype._changeSelectionDone = function (event) {
			var coord = this._getCoordinates(event);
			var ctrlKey = !AscCommon.getAltGr(event) && (event.metaKey || event.ctrlKey);
			if (false !== ctrlKey) {
				coord.x = -1;
				coord.y = -1;
			}
			this.handlers.trigger("changeSelectionDone", coord.x, coord.y, event);
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._resizeElement = function (event) {
			var coord = this._getCoordinates(event);
			this.handlers.trigger("resizeElement", this.targetInfo, coord.x, coord.y);
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._resizeElementDone = function (event) {
			var coord = this._getCoordinates(event);
			this.handlers.trigger("resizeElementDone", this.targetInfo, coord.x, coord.y, this.isResizeModeMove);
			this.isResizeModeMove = false;
		};

		/**
		 * @param event {MouseEvent}
		 * @param callback {Function}
		 */
		asc_CEventsController.prototype._changeFillHandle = function (event, callback, tableIndex) {
			var t = this;
			// Update in autofill mode
			var coord = this._getCoordinates(event);
			this.handlers.trigger("changeFillHandle", coord.x, coord.y,
				function (d) {
					if (!d) return;
					t.scroll(d);
					asc_applyFunction(callback);
				}, tableIndex);
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._changeFillHandle2 = function (event) {
			var t = this;

			var fn = function () {
				t._changeFillHandle2(event);
			};

			var callback = function () {
				if (t.isFillHandleMode && !t.hasCursor) {
					t.fillHandleModeTimerId = window.setTimeout(fn, t.settings.scrollTimeout);
				}
			};

			window.clearTimeout(t.fillHandleModeTimerId);
			t.fillHandleModeTimerId = window.setTimeout(function () {
				if (t.isFillHandleMode && !t.hasCursor) {
					t._changeFillHandle(event, callback);
				}
			}, 0);
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._changeFillHandleDone = function (event) {
			// We've finished autofilling, let's recalculate
			var coord = this._getCoordinates(event);
			var ctrlKey = !AscCommon.getAltGr(event) && (event.metaKey || event.ctrlKey);
			this.handlers.trigger("changeFillHandleDone", coord.x, coord.y, ctrlKey);
		};

		/**
		 * @param event {MouseEvent}
		 * @param callback {Function}
		 */
		asc_CEventsController.prototype._moveRangeHandle = function (event, callback, colRowMoveProps) {
			var t = this;
			// Updating in range moving mode
			var coord = this._getCoordinates(event);
			this.handlers.trigger("moveRangeHandle", coord.x, coord.y,
				function (d) {
					if (!d) return;
					t.scroll(d);
					asc_applyFunction(callback);
				}, colRowMoveProps);
		};

		/**
		 * @param event {MouseEvent}
		 * @param target
		 */
		asc_CEventsController.prototype._moveFrozenAnchorHandle = function (event, target) {
			var t = this;
			var coord = t._getCoordinates(event);
			t.handlers.trigger("moveFrozenAnchorHandle", coord.x, coord.y, target);
		};

		/**
		 * @param event {MouseEvent}
		 * @param target
		 */
		asc_CEventsController.prototype._moveFrozenAnchorHandleDone = function (event, target) {
			// Fixing the area
			var t = this;
			var coord = t._getCoordinates(event);
			t.handlers.trigger("moveFrozenAnchorHandleDone", coord.x, coord.y, target);
		};

		/**
		 * @param event {MouseEvent}
		 * @param target
		 * @param callback {Function}
		 */
		asc_CEventsController.prototype._moveResizeRangeHandle = function (event, target, callback) {
			var t = this;
			// Updating in range moving mode
			var coord = this._getCoordinates(event);
			this.handlers.trigger("moveResizeRangeHandle", coord.x, coord.y, target,
				function (d) {
					if (!d) return;
					t.scroll(d);
					asc_applyFunction(callback);
				});
		};

		asc_CEventsController.prototype._groupRowClick = function (event, target) {
			// Updating in range moving mode
			var coord = this._getCoordinates(event);
			return this.handlers.trigger("groupRowClick", coord.x, coord.y, target, event.type);
		};

		asc_CEventsController.prototype._commentCellClick = function (event) {
			// ToDo delete this function!
			var coord = this._getCoordinates(event);
			this.handlers.trigger("commentCellClick", coord.x, coord.y);
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._moveRangeHandleDone = function (event) {
			// We've finished moving the range, let's recalculate
			var ctrlKey = !AscCommon.getAltGr(event) && (event.metaKey || event.ctrlKey);
			this.handlers.trigger("moveRangeHandleDone", ctrlKey);
		};

		asc_CEventsController.prototype._moveResizeRangeHandleDone = function (isPageBreakPreview) {
			// We've finished moving the range, let's recalculate
			this.handlers.trigger("moveResizeRangeHandleDone", isPageBreakPreview);
		};

		/** @param event {jQuery.Event} */
		asc_CEventsController.prototype._onWindowResize = function (event) {
			var self = this;
			window.clearTimeout(this.resizeTimerId);
			this.resizeTimerId = window.setTimeout(function () {
				self.handlers.trigger("resize", event);
			}, 150);
		};
		asc_CEventsController.prototype.executeShortcut = function(nShortcutAction) {
			let oRet = {keyResult: keydownresult_PreventAll};
			const bSelectionDialogMode = this.getSelectionDialogMode();
			const bCanEdit = this.canEdit();
			switch (nShortcutAction) {
				case Asc.c_oAscSpreadsheetShortcutType.RecalculateAll: {
					this.handlers.trigger("calculate", Asc.c_oAscCalculateType.All);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.RecalculateActiveSheet: {
					this.handlers.trigger("calculate", Asc.c_oAscCalculateType.ActiveSheet);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots:
				case Asc.c_oAscSpreadsheetShortcutType.RefreshSelectedPivots: {
					if (bCanEdit && !this.getCellEditMode() && !bSelectionDialogMode) {
						this.handlers.trigger("refreshConnections", nShortcutAction === Asc.c_oAscSpreadsheetShortcutType.RefreshAllPivots);
					}
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.EditSelectAll: {
					if (!this.getCellEditMode()) {
						this.handlers.trigger("selectAllByRange");
					}
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellInsertDate:
				case Asc.c_oAscSpreadsheetShortcutType.CellInsertTime: {
					if (!bCanEdit || this.getCellEditMode() || bSelectionDialogMode) {
						break;
					}
					// When pressing the symbol, we do not set the focus. We clear the contents of the cell
					const oEnterOptions = new AscCommonExcel.CEditorEnterOptions();
					oEnterOptions.newText = '';
					oEnterOptions.quickInput = true;
					this.handlers.trigger("editCell", oEnterOptions);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.NextWorksheet: {
					this.handlers.trigger("showNextPrevWorksheet", +1);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.PreviousWorksheet: {
					this.handlers.trigger("showNextPrevWorksheet", -1);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.Strikeout: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setFontAttributes", "s");
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.Italic: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setFontAttributes", "i");
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.Bold: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setFontAttributes", "b");
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.Underline: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setFontAttributes", "u");
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.EditRedo: {
					if (!(bCanEdit || this.handlers.trigger('isRestrictionComments')) || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("redo");
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.EditUndo: {
					if (!(bCanEdit || this.handlers.trigger('isRestrictionComments')) || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("undo");
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.SpeechWorker: {
					if (this.getCellEditMode() || bSelectionDialogMode) {
						break;
					}
					AscCommon.EditorActionSpeaker.toggle();
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellInsertSumFunction: {
					if (!bCanEdit || this.getCellEditMode() || bSelectionDialogMode) {
						break;
					}
					this.handlers.trigger('addFunction',
						AscCommonExcel.cFormulaFunctionToLocale ? AscCommonExcel.cFormulaFunctionToLocale['SUM'] :
							'SUM', Asc.c_oAscPopUpSelectorType.Func, true);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.Print: {
					if (this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("print");
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.EditOpenCellEditor: {
					if (!bCanEdit || this.getCellEditMode() || bSelectionDialogMode) {
						break;
					}
					if (!AscBrowser.isOpera) {
						oRet.keyResult = keydownresult_PreventNothing;
					}
					// With F2 we set the focus in the editor
					const oEnterOptions = new AscCommonExcel.CEditorEnterOptions();
					oEnterOptions.focus = true;
					this.handlers.trigger("editCell", oEnterOptions);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellAddSeparator: {
					if (!bCanEdit || this.getCellEditMode() || bSelectionDialogMode) {
						break;
					}
					window["Asc"]["editor"].wb.EnterText(this.view.Api.asc_getDecimalSeparator().charCodeAt(0), true);
					//stop to prevent double enter
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellNumberFormat: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setCellFormat", Asc.c_oAscNumFormatType.Number);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellTimeFormat: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setCellFormat", Asc.c_oAscNumFormatType.Time);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellDateFormat: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setCellFormat", Asc.c_oAscNumFormatType.Date);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellCurrencyFormat: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setCellFormat", Asc.c_oAscNumFormatType.Currency);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellPercentFormat: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setCellFormat", Asc.c_oAscNumFormatType.Percent);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellExponentialFormat: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setCellFormat", Asc.c_oAscNumFormatType.Scientific);
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.CellGeneralFormat: {
					if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("setCellFormat", Asc.c_oAscNumFormatType.General);
					break;
				}

				case Asc.c_oAscSpreadsheetShortcutType.ShowFormulas: {
					if (bSelectionDialogMode || this.getCellEditMode()) {
						break;
					}
					this.handlers.trigger("showFormulas");
					break;
				}
				case Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize:
				case Asc.c_oAscSpreadsheetShortcutType.DecreaseFontSize: {
					if (this.getCellEditMode() || !bCanEdit || bSelectionDialogMode) {
						return true;
					}
					this.view.setFontAttributes("changeFontSize", nShortcutAction === Asc.c_oAscSpreadsheetShortcutType.IncreaseFontSize);
					break;
				}
				default: {
					const oCustom = this.view.Api.getCustomShortcutAction(nShortcutAction);
					if (oCustom) {
						if (!bCanEdit || bSelectionDialogMode || this.getCellEditMode()) {
							break;
						}
						if (AscCommon.c_oAscCustomShortcutType.Symbol === oCustom.Type) {
							const oEnterOptions = new AscCommonExcel.CEditorEnterOptions();
							oEnterOptions.newText = '';
							oEnterOptions.quickInput = true;
							this.handlers.trigger("editCell", oEnterOptions);
						}
					} else {
						oRet = null;
					}
					break;
				}
			}
			return oRet;
		};
		/** @param oEvent {AscCommon.CKeyboardEvent} */
		asc_CEventsController.prototype._onWindowKeyDown = function (oEvent) {
			const oThis = this;
			let nRetValue = keydownresult_PreventKeyPress;
			const nShortcutAction = oThis.view.Api.getShortcut(oEvent);

			// For those browsers that do not send the left mouse button release for double click, when exiting
			// the editor window and releasing the button, we will handle exiting the window (only Chrome sends the MouseUp event even when exiting the browser)
			this.showCellEditorCursor();

			if (oThis.getCellEditMode() && !oThis.hasFocus || oThis.isSelectMode ||
				oThis.isFillHandleMode || oThis.isMoveRangeMode || oThis.isMoveResizeRange) {
				// For some reason, I really want to process extra conditions in our code, instead of processing them at the top...
				if (oThis.enableKeyEvents || (nShortcutAction !== Asc.c_oAscSpreadsheetShortcutType.Print)) {
					// Only if events are disabled and Ctrl+S or Ctrl+P is pressed we will process them
					return nRetValue;
				}
			}

			let bIsSelect = oEvent.IsShift();
			const bIsMacOs = AscCommon.AscBrowser.isMacOs;
			const bSelectionDialogMode = this.getSelectionDialogMode();
			const bIsFormulaEditMode = this.getFormulaEditMode();
			const bCanEdit = this.canEdit();
			let nDeltaColumn = 0;
			let nDeltaRow = 0;
			//Are we moving in the selected area?
			let bSelectionActivePointChanged = false;
			let bIsNeedCheckActiveCellChanged = null;
			let oActiveCell;

			const oWb = window["Asc"]["editor"].wb;
			const oWs = oWb.getWorksheet();
			const isRightToLeft = oWs && oWs.getRightToLeft();

			const oShortcutRes = this.executeShortcut(nShortcutAction);
			if (oShortcutRes) {
				nRetValue = oShortcutRes.keyResult;
			} else {
				switch (oEvent.GetKeyCode()) {
					case 82:
						if (oEvent.CtrlKey && bIsSelect) {
							if (bCanEdit && !oThis.getCellEditMode() && !bSelectionDialogMode) {
								oThis.handlers.trigger("changeFormatTableInfo");
							}
							nRetValue = keydownresult_PreventAll;
						}
						break;
					case 8: // backspace
						if (!bCanEdit || oThis.getCellEditMode() || bSelectionDialogMode) {
							break;
						}
						nRetValue = keydownresult_PreventAll;

						// When backspace is pressed, the focus is not in the editor (we erase the contents)
						const oEnterOptions = new AscCommonExcel.CEditorEnterOptions();
						oEnterOptions.newText = '';
						oThis.handlers.trigger("editCell", oEnterOptions);
						break;

					case 46: // Del
						if (!bCanEdit || this.getCellEditMode() || bSelectionDialogMode || bIsSelect) {
							break;
						}
						// Delete the contents
						this.handlers.trigger("empty");
						break;

					case 9: // tab
						if (oThis.getCellEditMode() || bSelectionDialogMode) {
							break;
						}
						// Disable the browser's standard handling of tab pressing
						nRetValue = keydownresult_PreventAll;

						// Special case (movement in the selected area is possible)
						bSelectionActivePointChanged = true;
						if (bIsSelect) {
							nDeltaColumn = -1;			// (shift + tab) - move cells to the left by 1 column
							bIsSelect = false;	// Let's drop shift because we're not selecting
						} else {
							oActiveCell = oThis.handlers.trigger("getActiveCell");
							if (oThis.lastTab === null) {
								if (oActiveCell) {
									oThis.lastTab = oActiveCell.c2;
								}
							} else if (!oActiveCell) {
								oThis.lastTab = null;
							}
							nDeltaColumn = +1;			// (tab) - move cells to the right by 1 column
						}
						break;

					case 13:  // "enter"
						if (oThis.getCellEditMode() || bSelectionDialogMode) {
							break;
						}

						if (bCanEdit && (oEvent.CtrlKey || oEvent.AltKey)) {
							oThis.setSkipCellEditor(true);
							const oEnterOptions = new AscCommonExcel.CEditorEnterOptions();
							oEnterOptions.focus = true;
							oThis.handlers.trigger("editCell", oEnterOptions);
							break;
						}
						// Special case (movement in the selected area is possible)
						bSelectionActivePointChanged = true;
						if (bIsSelect) {
							nDeltaRow = -1;			// (shift + enter) - move up cells by 1 row
							bIsSelect = false;	// Let's drop shift because we're not selecting
							oThis.lastTab = null;
						} else {
							if (oThis.lastTab !== null) {
								oActiveCell = oThis.handlers.trigger("getActiveCell");
								if (oActiveCell) {
									nDeltaColumn = oThis.lastTab - oActiveCell.c2;
								} else {
									oThis.lastTab = null;
								}
							}
							nDeltaRow = +1;			// (enter) - move down cells by 1 row
						}
						break;

					case 27: // Esc
						oThis.handlers.trigger("stopFormatPainter");
						oThis.handlers.trigger("stopAddShape");
						oThis.handlers.trigger("cleanCutData", true, true);
						oThis.handlers.trigger("cleanCopyData", true, true);
						oThis.view.Api.cancelEyedropper();
						window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
						break;

					case 144: //Num Lock
					case 145: //Scroll Lock
						if (AscBrowser.isOpera) {
							nRetValue = keydownresult_PreventAll;
						}
						break;

					case 32: // Spacebar
						if (oThis.getCellEditMode()) {
							break;
						}
						const bIsSelectColumns = oEvent.IsShortcutCtrl() || oEvent.IsMacCmd();
						if (bIsSelectColumns && bIsSelect && bIsMacOs) {
							break;
						}
						// Disable the browser's standard click handling
						// Ctrl+Shift+Spacebar, Ctrl+Spacebar, Shift+Spacebar
						if (bIsSelectColumns) {
							oThis.handlers.trigger("selectColumnsByRange");
							nRetValue = keydownresult_PreventAll;
						}
						if (bIsSelect) {
							oThis.handlers.trigger("selectRowsByRange");
							nRetValue = keydownresult_PreventAll;
						}
						break;
					case 33: // PageUp
						nDeltaRow = -0.5;
						bIsNeedCheckActiveCellChanged = true;
						nRetValue = keydownresult_PreventAll;
						break;

					case 34: // PageDown
						nDeltaRow = +0.5;
						bIsNeedCheckActiveCellChanged = true;
						nRetValue = keydownresult_PreventAll;
						break;

					case 37: // left
						nDeltaColumn = oEvent.CtrlKey ? -1.5 : -1;  // Movement with arrows (left-right, up-down)
						bIsNeedCheckActiveCellChanged = true;
						nRetValue = keydownresult_PreventAll;                          // Disable the browser's standard handling of pressing left
						if (isRightToLeft) {
							nDeltaColumn = -nDeltaColumn;
						}
						break;

					case 38: // up
						if (bCanEdit && !oThis.getCellEditMode() && !bSelectionDialogMode && oEvent.IsAlt() && oThis.handlers.trigger("onDataValidation")) {
							break;
						}
						nDeltaRow = oEvent.CtrlKey ? -1.5 : -1;  // Movement with arrows (left-right, up-down)
						bIsNeedCheckActiveCellChanged = true;
						nRetValue = keydownresult_PreventAll;                           // Disable the browser's standard handling of pressing up
						break;

					case 39: // right
						nDeltaColumn = oEvent.CtrlKey ? +1.5 : +1;  // Movement with arrows (left-right, up-down)
						bIsNeedCheckActiveCellChanged = true;
						nRetValue = keydownresult_PreventAll;                           // Disable the browser's standard handling of pressing right
						if (isRightToLeft) {
							nDeltaColumn = -nDeltaColumn;
						}
						break;

					case 40: // down
						nRetValue = keydownresult_PreventAll;                           // Disable the browser's standard handling of pressing down
						// Alt + down processing
						if (bCanEdit && !oThis.getCellEditMode() && !bSelectionDialogMode && oEvent.IsAlt()) {
							if (oThis.handlers.trigger("onShowFilterOptionsActiveCell")) {
								break;
							}
							if (oThis.handlers.trigger("onDataValidation")) {
								break;
							}
							oThis.handlers.trigger("showAutoComplete");
							break;
						}
						nDeltaRow = oEvent.CtrlKey ? +1.5 : +1;  // Movement with arrows (left-right, up-down)
						bIsNeedCheckActiveCellChanged = true;
						break;

					case 36: // home
						nRetValue = keydownresult_PreventAll;                           // Disable the browser's standard handling of pressing home
						if (bIsFormulaEditMode) {
							break;
						}
						nDeltaColumn = -2.5;
						if (oEvent.CtrlKey) {
							nDeltaRow = -2.5;
						}
						bIsNeedCheckActiveCellChanged = true;
						break;

					case 35: // end
						nRetValue = keydownresult_PreventAll;                           // Disable the browser's standard handling of pressing end
						if (bIsFormulaEditMode) {
							break;
						}
						nDeltaColumn = 2.5;
						if (oEvent.CtrlKey) {
							nDeltaRow = 2.5;
						}
						bIsNeedCheckActiveCellChanged = true;
						break;
					case 93:
						if (!oEvent.MacCmdKey) {
							nRetValue = keydownresult_PreventAll;
							this.handlers.trigger('onContextMenu', oEvent);
						}
						break;
					default:
						nRetValue = keydownresult_PreventNothing;
						break;
				}
			}

			let oActiveCellBefore;
			if (bIsNeedCheckActiveCellChanged) {
				oActiveCellBefore = oThis.handlers.trigger("getActiveCell");
			}
			const CheckLastTab = function () {
				if (bIsNeedCheckActiveCellChanged) {
					const oActiveCellAfter = oThis.handlers.trigger("getActiveCell");
					if (!oActiveCellBefore || !oActiveCellAfter || !oActiveCellAfter.isEqual(oActiveCellBefore)) {
						oThis.lastTab = null;
					}
				}
			};

			if ((nDeltaColumn !== 0 || nDeltaRow !== 0) && false === oThis.handlers.trigger("isGlobalLockEditCell")) {
				const bIsChangeVisibleAreaMode = this.view.Api.isEditVisibleAreaOleEditor;
				if (bIsChangeVisibleAreaMode) {
					oThis.handlers.trigger("changeVisibleArea", !bIsSelect, nDeltaColumn, nDeltaRow, false, function (d) {
						if (oThis.targetInfo) {
							oWb._onUpdateWorksheet(oThis.targetInfo.coordX, oThis.targetInfo.coordY, false);
						}
						oThis.scroll(d);
						const oOleSize = oWb.getOleSize();
						oOleSize.addPointToLocalHistory();
						CheckLastTab();
					}, true);
				} else if (bSelectionActivePointChanged) { // Check for movement in a selected area
					oThis.handlers.trigger("selectionActivePointChanged", nDeltaColumn, nDeltaRow, function (d) {
						oThis.scroll(d);
						CheckLastTab();
					});
				} else {
					oThis.handlers.trigger("changeSelection", /*isStartPoint*/!bIsSelect, nDeltaColumn, nDeltaRow, /*isCoord*/false, false,
						function (d) {
							const oWb = window["Asc"]["editor"].wb;
							if (oThis.targetInfo) {
								oWb._onUpdateWorksheet(oThis.targetInfo.coordX, oThis.targetInfo.coordY, false);
							}
							oThis.scroll(oWs.convertOffsetToSmooth(d));
							CheckLastTab();
						});
				}
			}

			if (nRetValue & keydownresult_PreventKeyPress) {
				oThis._setSkipKeyPress(true);
			}
			return nRetValue;
		};


		/** @param event {AscCommon.CKeyboardEvent} */
		asc_CEventsController.prototype._onWindowKeyPress = function (event) {
			// It is not allowed to return false when events are disabled (this also applies to ViewerMode)
			if (!this.enableKeyEvents) {
				return true;
			}

			// do not enter text in view mode
			// if false is returned in FF, then further processing of the keydown -> keypress -> keyup series is cancelled
			// and then we will not process ctrl+c and similar events
			if (!this.canEdit() || this.getSelectionDialogMode() || this.view.Api.isEditVisibleAreaOleEditor) {
				return true;
			}

			// For those browsers that do not send the left mouse button release for double click, when exiting the editor
			// window and releasing the button, we will handle exiting the window (only Chrome sends the MouseUp event even when exiting the browser)
			this.showCellEditorCursor();

			// Cannot enter when selecting or when performing actions on an object
			if (this.getCellEditMode() && !this.hasFocus || this.isSelectMode ||
				!this.handlers.trigger('canReceiveKeyPress')) {
				return true;
			}

			if (this.skipKeyPress || event.KeyCode < 32) {
				this._setSkipKeyPress(true);
				return true;
			}

			if (!this.getCellEditMode()) {
				if (this.handlers.trigger("graphicObjectWindowKeyPress", event)) {
					return true;
				}

				// When a symbol is pressed, we do not set the focus and clear the contents of the cell
				var enterOptions = new AscCommonExcel.CEditorEnterOptions();
				enterOptions.newText = '';
				enterOptions.quickInput = true;
				this.handlers.trigger("editCell", enterOptions);
			}
			return true;
		};

		asc_CEventsController.prototype.EnterText = function (codePoints) {
			//TODO is almost a copy of _onWindowKeyPress - after the EnterText function is enabled - check and merge the functions
			// It is impossible to return false when events are disabled (this also applies to ViewerMode)
			if (!this.enableKeyEvents) {
				return true;
			}

			// do not enter text in view mode
			// if false is returned in FF, then further processing of the keydown -> keypress -> keyup series is cancelled
			// and then we will not process ctrl+c and similar events
			if (!this.canEdit() || this.getSelectionDialogMode() || this.view.Api.isEditVisibleAreaOleEditor) {
				return true;
			}

			// For those browsers that do not send the left mouse button release for double click, when exiting the editor window and releasing the button,
			// we will handle exiting the window (only Chrome sends the MouseUp event even when exiting the browser)
			this.showCellEditorCursor();

			// Cannot enter when selecting or when performing actions on an object
			if (this.getCellEditMode() && !this.hasFocus || this.isSelectMode ||
				!this.handlers.trigger('canReceiveKeyPress')) {
				return true;
			}

			/*if (this.skipKeyPress) {
				this._setSkipKeyPress(true);
				return true;
			}*/

			if (!this.getCellEditMode()) {
				if (this.handlers.trigger("graphicObjectWindowEnterText", codePoints)) {
					return true;
				}

				// When a symbol is pressed, we do not set the focus and clear the contents of the cell
				var enterOptions = new AscCommonExcel.CEditorEnterOptions();
				enterOptions.newText = '';
				enterOptions.quickInput = true;
				this.handlers.trigger("editCell", enterOptions);
			}
			return true;
		};

		/** @param event {AscCommon.CKeyboardEvent} */
		asc_CEventsController.prototype._onWindowKeyUp = function (event) {
			// When you release shift, you need to send the selection information
			if (16 === event.KeyCode) {
				this.handlers.trigger("updateSelectionName");
			}
			this.handlers.trigger("graphicObjectWindowKeyUp", event);

			return true;
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._onWindowMouseMove = function (event) {
			var coord = this._getCoordinates(event);

			if (this.isSelectMode && !this.hasCursor) {
				this._changeSelection2(event);
			}
			if (this.isChangeVisibleAreaMode && !this.hasCursor) {
				this._continueChangeVisibleArea2(event);
			}
			if (this.isResizeMode && !this.hasCursor) {
				this.isResizeModeMove = true;
				this._resizeElement(event);
			}
			if (this.hsbApiLockMouse)
				this.hsbApi.mouseDown ? this.hsbApi.evt_mousemove.call(this.hsbApi, event) : false;
			else if (this.vsbApiLockMouse)
				this.vsbApi.mouseDown ? this.vsbApi.evt_mousemove.call(this.vsbApi, event) : false;

			// Pinned Areas Setup Mode
			if (this.frozenAnchorMode) {
				this._moveFrozenAnchorHandle(event, this.frozenAnchorMode);
				return true;
			}

			if (this.isShapeAction) {
				event.isLocked = this.isMousePressed;
				this.handlers.trigger("graphicObjectMouseMove", event, coord.x, coord.y);
			}

			if (this.isRowGroup) {
				if (!this._groupRowClick(event, this.targetInfo)) {
					this.isRowGroup = false;
				}
			}

			return true;
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._onWindowMouseUp = function (event) {
			AscCommon.global_mouseEvent.UnLockMouse();

			var button = AscCommon.getMouseButton(event);
			var coord = this._getCoordinates(event);
			if (this.hsbApiLockMouse)
				this.hsbApi.mouseDown ? this.hsbApi.evt_mouseup.call(this.hsbApi, event) : false;
			else if (this.vsbApiLockMouse)
				this.vsbApi.mouseDown ? this.vsbApi.evt_mouseup.call(this.vsbApi, event) : false;

			this.isMousePressed = false;
			// Shapes
			if (this.isShapeAction) {
				if (!this.isUpOnCanvas) {
					let oDrawingsController = this.view.getWorksheet().objectRender.controller;
					if (oDrawingsController.haveTrackedObjects()) {
						event.isLocked = this.isMousePressed;
						event.ClickCount = this.clickCounter.clickCount;
						event.fromWindow = true;
						this.handlers.trigger("graphicObjectMouseUp", event, coord.x, coord.y);
						this._changeSelectionDone(event);
					}
				}
				this.isUpOnCanvas = false;
				return true;
			}

			if (this.isChangeVisibleAreaMode) {
				this.isChangeVisibleAreaMode = false;
				const oOleSize = this.view.getOleSize();
				oOleSize.addPointToLocalHistory();
			}

			if (this.isSelectMode) {
				this.isSelectMode = false;
				this._changeSelectionDone(event);
			}

			if (this.isResizeMode) {
				this.isResizeMode = false;
				this._resizeElementDone(event);
			}

			// Autofill mode
			if (this.isFillHandleMode) {
				// Finished autofilling
				this.isFillHandleMode = false;
				if (2 === button && this.handlers.trigger('isRightClickFill')) {
					this.handlers.trigger('onContextMenu', event, Asc.c_oAscContextMenuTypes.changeSeries);
				} else {
					this._changeFillHandleDone(event);
				}
			}

			// Range Shift Mode
			if (this.isMoveRangeMode) {
				// Finished moving the range
				this.isMoveRangeMode = false;
				this._moveRangeHandleDone(event);
			}

			if (this.isMoveResizeRange) {
				this.isMoveResizeRange = false;
				this._moveResizeRangeHandleDone(this.targetInfo && this.targetInfo.isPageBreakPreview);
			}
			// Pinned Areas Setup Mode
			if (this.frozenAnchorMode) {
				this._moveFrozenAnchorHandleDone(event, this.frozenAnchorMode);
				this.frozenAnchorMode = false;
			}

			// We may not work on dblClick if we left the area and released the mouse button, we need to work on it
			this.showCellEditorCursor();


			return true;
		};

		/**
		 *
		 * @param event
		 * @param x
		 * @param y
		 */
		asc_CEventsController.prototype._onWindowMouseUpExternal = function (event, x, y) {
			// ToDo should be reworked into a normal scheme, for now we will write it directly in the event
			if (null != x && null != y)
				event.coord = {x: x, y: y};
			this._onWindowMouseUp(event);

			if (window.g_asc_plugins)
				window.g_asc_plugins.onExternalMouseUp();
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._onWindowMouseLeaveOut = function (event) {
			// When there is nothing to process, we leave.
			if (!this.isDoBrowserDblClick)
				return true;

			var relatedTarget = event.relatedTarget || event.fromElement;
			// If we move along the cell editor, we don’t remove anything.
			if (relatedTarget && ("ce-canvas-outer" === relatedTarget.id ||
				"ce-canvas" === relatedTarget.id || "ce-canvas-overlay" === relatedTarget.id ||
				"ce-cursor" === relatedTarget.id || "ws-canvas-overlay" === relatedTarget.id))
				return true;

			// For those browsers that do not send the left mouse button release for double click, when exiting the editor window
			// and releasing the button, we will handle exiting the window (only Chrome sends the MouseUp event even when exiting the browser)
			this.showCellEditorCursor();
			return true;
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._onMouseDown = function (event) {
			let touchManager = this.view.Api.wb.MobileTouchManager;
			if (touchManager && touchManager.checkTouchEvent(event))
			{
				touchManager.startTouchingInProcess();
				let res = touchManager.mainOnTouchStart(event);
				touchManager.stopTouchingInProcess();
				return res;
			}
			if (touchManager)
				touchManager.checkMouseFocus(event);

			var t = this;
			asc["editor"].checkInterfaceElementBlur();
			var ctrlKey = !AscCommon.getAltGr(event) && (event.metaKey || event.ctrlKey);
			var coord = t._getCoordinates(event);
			var button = AscCommon.getMouseButton(event);
			event.isLocked = t.isMousePressed = true;
			this.isShapeAction = false;
			// Shapes
			var graphicsInfo = t.handlers.trigger("getGraphicsInfo", coord.x, coord.y);
			if (!graphicsInfo) {
				// Update state for device without cursor
				this._onMouseMove(event);
			}

			if (this.view.Api.isEditVisibleAreaOleEditor) {
				if (button === 0 && this.view.isInnerOfWorksheet(coord.x, coord.y)) {
					if (this.targetInfo && this.targetInfo.target === c_oTargetType.MoveResizeRange) {
						t._moveResizeRangeHandle(event, t.targetInfo);
						t.isMoveResizeRange = true;
					} else {
						this._startChangeVisibleArea(event);
						this.isChangeVisibleAreaMode = true;
					}
				}
				return;
			}

			if (AscCommon.g_inputContext) {
				AscCommon.g_inputContext.externalChangeFocus();
			}

			AscCommon.global_mouseEvent.LockMouse();
			AscCommon.capturePointer(event, this.element);

			if (t.view.Api.isEyedropperStarted()) {
				return;
			}
			if (t.handlers.trigger("isGlobalLockEditCell")) {
				return;
			}

			if (!this.enableKeyEvents) {
				t.handlers.trigger("canvasClick");
			}

			if (!(asc["editor"].isStartAddShape || asc["editor"].isInkDrawerOn() || this.getSelectionDialogMode() || this.getCellEditMode() && !this.handlers.trigger("stopCellEditing"))) {
				const isPlaceholder = t.handlers.trigger("onPointerDownPlaceholder", coord.x, coord.y);
				if (isPlaceholder) {
					return;
				}
			}

			// do not work with drawings in selection dialog mode
			if (!this.getSelectionDialogMode()) {
				if (asc["editor"].isStartAddShape || asc["editor"].isInkDrawerOn() || graphicsInfo) {


					if (this.getCellEditMode() && !this.handlers.trigger("stopCellEditing")) {
						return;
					}

					t.isShapeAction = true;
					t.isUpOnCanvas = false;


					t.clickCounter.mouseDownEvent(coord.x, coord.y, button);
					event.ClickCount = t.clickCounter.clickCount;
					if (0 === event.ClickCount % 2) {
						t.isDblClickInMouseDown = true;
					}

					t.handlers.trigger("graphicObjectMouseDown", event, coord.x, coord.y);
					t.handlers.trigger("updateSelectionShape", /*isSelectOnShape*/true);
					return;
				}
			}


			if (2 === event.detail) {
				// This means that it is MouseDown for dblClick event (it does not need to be processed)
				// The order of events for dblClick is http://javascript.ru/tutorial/events/mouse#dvoynoy-levyy-klik

				// Check for IE, because it sends DblClick when the mouse moves...
				if (this.mouseDownLastCord && coord.x === this.mouseDownLastCord.x && coord.y === this.mouseDownLastCord.y &&
					0 === button && !this.handlers.trigger('isFormatPainter')) {
					// We show that we have already done dblClick (otherwise the browser suddenly does not support the detail property)
					this.isDblClickInMouseDown = true;
					// We need to handle the browser event about dblClick (if we edit a cell, then we will show the cursor, if not, then nothing will happen)
					this.isDoBrowserDblClick = true;
					this.doMouseDblClick(event);
					// We reset the coordinates
					this.mouseDownLastCord = null;
					return;
				}
			}
			// For IE preventDefault is not necessary
			if (!(AscBrowser.isIE || AscBrowser.isOpera)) {
				if (event.preventDefault) {
					event.preventDefault();
				} else {
					event.returnValue = false;
				}
			}

			// Remember the coordinates of the click
			this.mouseDownLastCord = coord;

			if (!t.getCellEditMode() && !t.getSelectionDialogMode()) {
				this.gotFocus(true);
				if (event.shiftKey && !(t.targetInfo.target === c_oTargetType.ColumnRowHeaderMove && this.canEdit())) {
					t.isSelectMode = true;
					t._changeSelection(event);
					return;
				}
				if (t.targetInfo) {
					if ((t.targetInfo.target === c_oTargetType.ColumnResize ||
						t.targetInfo.target === c_oTargetType.RowResize) && 0 === button) {
						t.isResizeMode = true;
						t._resizeElement(event);
						return;
					} else if (t.targetInfo.target === c_oTargetType.FillHandle && this.canEdit()) {
						// In autofill mode
						this.isFillHandleMode = true;
						t._changeFillHandle(event, null, t.targetInfo.tableIndex);
						return;
					} else if (t.targetInfo.target === c_oTargetType.MoveRange && this.canEdit()) {
						// In range moving mode
						this.isMoveRangeMode = true;
						t._moveRangeHandle(event);
						return;
					} else if (t.targetInfo.target === c_oTargetType.FilterObject) {
						if (0 === button) {
							if (t.targetInfo.isDataValidation) {
								this.handlers.trigger('onDataValidation');
							} else if (t.targetInfo.idPivot) {
								this.handlers.trigger("pivotFiltersClick", t.targetInfo.idPivot);
							} else if (t.targetInfo.idPivotCollapse) {
								this.handlers.trigger("pivotCollapseClick", t.targetInfo.idPivotCollapse);
							} else if (t.targetInfo.idTableTotal) {
								this.handlers.trigger("tableTotalClick", t.targetInfo.idTableTotal);
							} else {
								this.handlers.trigger("autoFiltersClick", t.targetInfo.idFilter);
							}
						}
						event.preventDefault && event.preventDefault();
						event.stopPropagation && event.stopPropagation();
						return;
					} else if (t.targetInfo.commentIndexes) {
						t._commentCellClick(event);
					} else if (t.targetInfo.target === c_oTargetType.MoveResizeRange && this.canEdit()) {
						this.isMoveResizeRange = true;
						t._moveResizeRangeHandle(event, t.targetInfo);
						return;
					} else if ((t.targetInfo.target === c_oTargetType.FrozenAnchorV ||
						t.targetInfo.target === c_oTargetType.FrozenAnchorH) && this.canEdit()) {
						// Pinned Areas Setup Mode
						this.frozenAnchorMode = t.targetInfo.target;
						t._moveFrozenAnchorHandle(event, this.frozenAnchorMode);
						return;
					} else if (t.targetInfo.target === c_oTargetType.GroupRow && 0 === button) {
						if (t._groupRowClick(event, t.targetInfo)) {
							t.isRowGroup = true;
						}
						return;
					} else if (t.targetInfo.target === c_oTargetType.GroupCol && 0 === button) {
						if (t._groupRowClick(event, t.targetInfo)) {
							t.isRowGroup = true;
						}
						return;
					} else if ((t.targetInfo.target === c_oTargetType.GroupCol || t.targetInfo.target === c_oTargetType.GroupRow) && 2 === button) {
						this.handlers.trigger('onContextMenu', null);
						return;
					} else if (t.targetInfo.target === c_oTargetType.TableSelectionChange) {
						this.handlers.trigger('onChangeTableSelection', t.targetInfo);
						return;
					} else if (t.targetInfo.target === c_oTargetType.TraceDependents) {
						// if we do a single click on traces, do nothing
						return;
					}
				}
			} else {
				if (this.getFormulaEditMode()) {
					if (this.targetInfo && this.targetInfo.target === c_oTargetType.MoveResizeRange) {
						this.isMoveResizeRange = true;
						this._moveResizeRangeHandle(event, this.targetInfo);
						return;
					} else if (this.targetInfo && this.targetInfo.target === c_oTargetType.FillHandle) {
						return;
					}

					if (2 === button) {
						return;
					}
				} else {
					if (!this.handlers.trigger("stopCellEditing")) {
						return;
					}
				}

				this.gotFocus(true);

				if (event.shiftKey) {
					this.isSelectMode = true;
					this._changeSelection(event);
					return;
				}
			}

			// If you press the right mouse button, then we will change the selection only if we are not in the selected area
			if (2 === button) {
				this.handlers.trigger("changeSelectionRightClick", coord.x, coord.y, this.targetInfo && this.targetInfo.target);
				this.handlers.trigger('onContextMenu', event);
			} else {
				if (this.targetInfo && this.targetInfo.target === c_oTargetType.FillHandle && this.canEdit()) {
					// In autofill mode
					this.isFillHandleMode = true;
					this._changeFillHandle(event, null, t.targetInfo.tableIndex);
				} else {
					if (this.targetInfo && this.targetInfo.target === c_oTargetType.ColumnRowHeaderMove) {
						this.isMoveRangeMode = true;
						t._moveRangeHandle(event, null, {ctrlKey: ctrlKey, shiftKey: event.shiftKey});
						t.handlers.trigger("updateWorksheet", coord.x, coord.y);
						//t.handlers.trigger("updateWorksheet", coord.x, coord.y, ctrlKey, function(info){t.targetInfo = info;});
					} else {
						this.isSelectMode = true;
						this.handlers.trigger("changeSelection", /*isStartPoint*/true, coord.x, coord.y, /*isCoord*/true,
							ctrlKey);
					}
				}
			}
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._onMouseUp = function (event) {
			let touchManager = this.view.Api.wb.MobileTouchManager;
			if (touchManager && touchManager.checkTouchEvent(event))
			{
				touchManager.startTouchingInProcess();
				let res = touchManager.mainOnTouchEnd(event);
				touchManager.stopTouchingInProcess();
				return res;
			}

			var button = AscCommon.getMouseButton(event);
			AscCommon.global_mouseEvent.UnLockMouse();

			if (this.isChangeVisibleAreaMode) {
				this.isChangeVisibleAreaMode = false;
				const oOleSize = this.view.getOleSize();
				oOleSize.addPointToLocalHistory();
			}

			var coord = this._getCoordinates(event);
			// Shapes
			event.isLocked = this.isMousePressed = false;
			if (2 === button) {
				if (this.isShapeAction) {
					event.ClickCount = this.clickCounter.clickCount;
					this.handlers.trigger("graphicObjectMouseUp", event, coord.x, coord.y);
					this.handlers.trigger('onContextMenu', event);
				}
				return true;
			}

			if (this.view.Api.isEyedropperStarted()) {
				this.view.Api.finishEyedropper();
				var t = this;
				t.handlers.trigger("updateWorksheet", coord.x, coord.y, false, function (info) {
					t.targetInfo = info;
				});
				return true;
			}

			if (this.isShapeAction) {
				event.ClickCount = this.clickCounter.clickCount;
				this.handlers.trigger("graphicObjectMouseUp", event, coord.x, coord.y);
				this._changeSelectionDone(event);
				if (asc["editor"].isStartAddShape || asc["editor"].isInkDrawerOn()) {
					event.preventDefault && event.preventDefault();
					event.stopPropagation && event.stopPropagation();
				} else {
					this.isUpOnCanvas = true;
				}

				return true;
			}

			if (this.isSelectMode) {
				this.isSelectMode = false;
				this._changeSelectionDone(event);
			}

			if (this.isResizeMode) {
				this.isResizeMode = false;
				this._resizeElementDone(event);
			}
			// Autofill mode
			if (this.isFillHandleMode) {
				// Finished autofilling
				this.isFillHandleMode = false;
				this._changeFillHandleDone(event);
			}
			// Range Shift Mode
			if (this.isMoveRangeMode) {
				this.isMoveRangeMode = false;
				this._moveRangeHandleDone(event);
			}

			if (this.isMoveResizeRange) {
				this.isMoveResizeRange = false;
				this._moveResizeRangeHandleDone(this.targetInfo && this.targetInfo.pageBreakSelectionType);
				return true;
			}
			// Pinned Areas Setup Mode
			if (this.frozenAnchorMode) {
				this._moveFrozenAnchorHandleDone(event, this.frozenAnchorMode);
				this.frozenAnchorMode = false;
			}

			if (this.isRowGroup/* && this.targetInfo && this.targetInfo.target === c_oTargetType.GroupRow && 0 === event.button*/) {
				this._groupRowClick(event, this.targetInfo);
				this.isRowGroup = false;
				return;
			}

			if (this.targetInfo && this.targetInfo.target === c_oTargetType.ColumnHeader) {
				this._onMouseMove(event);
			}

			// We may not work on dblClick if we left the area and released the mouse button, we need to work on it
			this.showCellEditorCursor();
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._onMouseMove = function (event) {
			let touchManager = this.view.Api.wb.MobileTouchManager;
			if (touchManager && touchManager.checkTouchEvent(event))
			{
				touchManager.startTouchingInProcess();
				let res = touchManager.mainOnTouchMove(event);
				touchManager.stopTouchingInProcess();
				return res;
			}

			var t = this;
			var ctrlKey = !AscCommon.getAltGr(event) && (event.metaKey || event.ctrlKey);
			var coord = t._getCoordinates(event);

			t.hasCursor = true;

			if (t.view.Api.isEyedropperStarted()) {
				t.view.Api.checkEyedropperColor(coord.x, coord.y);
				t.handlers.trigger("updateWorksheet", coord.x, coord.y, ctrlKey, function (info) {
					t.targetInfo = info;
				});
				return true;
			}
			// Shapes
			var graphicsInfo = t.handlers.trigger("getGraphicsInfo", coord.x, coord.y);
			if (graphicsInfo)
				this.clickCounter.mouseMoveEvent(coord.x, coord.y);

			if (t.isSelectMode) {
				t._changeSelection(event);
				return true;
			}

			if (t.isChangeVisibleAreaMode) {
				t._continueChangeVisibleArea(event);
				return true;
			}

			if (t.isResizeMode) {
				t._resizeElement(event);
				this.isResizeModeMove = true;
				return true;
			}

			// Autofill mode
			if (t.isFillHandleMode) {
				t._changeFillHandle(event);
				return true;
			}

			// Range Shift Mode
			if (t.isMoveRangeMode) {
				t._moveRangeHandle(event);
				return true;
			}

			if (t.isMoveResizeRange) {
				t._moveResizeRangeHandle(event, t.targetInfo);
				return true;
			}

			// Pinned Areas Setup Mode
			if (t.frozenAnchorMode) {
				t._moveFrozenAnchorHandle(event, this.frozenAnchorMode);
				return true;
			}

			if (t.isShapeAction || graphicsInfo || asc["editor"].isInkDrawerOn()) {
				event.isLocked = t.isMousePressed;
				t.handlers.trigger("graphicObjectMouseMove", event, coord.x, coord.y);
				t.handlers.trigger("updateWorksheet", coord.x, coord.y, ctrlKey, function (info) {
					t.targetInfo = info;
				});
				return true;
			}

			t.handlers.trigger("updateWorksheet", coord.x, coord.y, ctrlKey, function (info) {
				t.targetInfo = info;
			});
			return true;
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._onMouseLeave = function (event) {
			var t = this;
			var coord = t._getCoordinates(event);
			this.hasCursor = false;
			if (!this.isSelectMode && !this.isResizeMode && !this.isMoveResizeRange) {
				this.targetInfo = undefined;
				this.handlers.trigger("updateWorksheet", coord.x, coord.y);
			}
			if (this.isMoveRangeMode) {
				t.moveRangeTimerId = window.setTimeout(function () {
					t._moveRangeHandle2(event)
				}, 0);
			}
			if (this.isMoveResizeRange) {
				t.moveResizeRangeTimerId = window.setTimeout(function () {
					t._moveResizeRangeHandle2(event)
				}, 0);
			}
			if (this.isFillHandleMode) {
				t.fillHandleModeTimerId = window.setTimeout(function () {
					t._changeFillHandle2(event)
				}, 0);
			}
			if (t.view.Api.isEyedropperStarted()) {
				this.view.Api.sendEvent("asc_onHideEyedropper");
			}
			return true;
		};

		/** @param event {MouseEvent} */
		asc_CEventsController.prototype._onMouseWheel = function (event) {
			var ctrlKey = !AscCommon.getAltGr(event) && (event.metaKey || event.ctrlKey);
			if (ctrlKey) {
				if (event.preventDefault) {
					event.preventDefault();
				} else {
					event.returnValue = false;
				}

				return false;
			}
			if (this.isFillHandleMode || this.isMoveRangeMode || this.isMoveResizeRange || this.isChangeVisibleAreaMode) {
				return true;
			}

			if (undefined !== window["AscDesktopEditor"]) {
				if (false === window["AscDesktopEditor"]["CheckNeedWheel"]())
					return true;
			}

			var self = this;
			var deltaX = 0, deltaY = 0;

			//TODO!!! while only check direction. need refactor, and replace up code on checkMouseWhell function
			let values = AscCommon.checkMouseWhell(event, {
				isSupportBidirectional : false,
				isAllowHorizontal : true,
				isUseMaximumDelta : true
			});

			const wb = window["Asc"]["editor"].wb;
			//TODO for mac touchpads. need review
			if (wb.smoothScroll && AscCommon.AscBrowser.isMacOs) {
				deltaX = (values.x / wb.getWorksheet().getHScrollStep()) * AscCommon.AscBrowser.retinaPixelRatio;
				deltaY = (values.y / wb.getWorksheet().getVScrollStep()) * AscCommon.AscBrowser.retinaPixelRatio;
			} else {
				if (undefined !== event.wheelDelta && 0 !== event.wheelDelta) {
					deltaY = -1 * event.wheelDelta / 40;
				} else if (undefined !== event.detail && 0 !== event.detail) {
					// FF
					deltaY = event.detail;
				} else if (undefined !== event.deltaY && 0 !== event.deltaY) {
					// FF
					//limiting step due to incorrect deltaY value after FF update
					//TODO needs to be revised. correct values ​​are needed and taking into account the system step.
					var _maxDelta = 3;
					if (AscCommon.AscBrowser.isMozilla && Math.abs(event.deltaY) > _maxDelta) {
						deltaY = Math.sign(event.deltaY) * _maxDelta;
					} else {
						deltaY = event.deltaY;
					}
				}
				if (undefined !== event.deltaX && 0 !== event.deltaX) {
					deltaX = event.deltaX;
				}
				if (event.axis !== undefined && event.axis === event.HORIZONTAL_AXIS) {
					deltaX = deltaY;
					deltaY = 0;
				}

				if (undefined !== event.wheelDeltaX && 0 !== event.wheelDeltaX) {
					// Webkit
					deltaX = -1 * event.wheelDeltaX / 40;
				}
				if (undefined !== event.wheelDeltaY && 0 !== event.wheelDeltaY) {
					// Webkit
					deltaY = -1 * event.wheelDeltaY / 40;
				}
			}

			if (values.x === 0) {
				deltaX = 0;
			}
			if (values.y === 0) {
				deltaY = 0;
			}

			if (!AscCommon.AscBrowser.isMacOs && event.shiftKey) {
				deltaX = deltaY;
				deltaY = 0;
			}

			if (this.smoothWheelCorrector && !wb.smoothScroll) {
				deltaX = this.smoothWheelCorrector.get_DeltaX(deltaX);
				deltaY = this.smoothWheelCorrector.get_DeltaY(deltaY);
			}
			if (this.handlers.trigger("graphicObjectMouseWheel", deltaX, deltaY)) {
				self._onMouseMove(event);
				AscCommon.stopEvent(event);
				return true;
			}

			this.handlers.trigger("updateWorksheet", /*x*/undefined, /*y*/undefined, /*ctrlKey*/undefined,
				function () {
					if (deltaX && (!self.smoothWheelCorrector || !self.smoothWheelCorrector.isBreakX())) {
						if (!wb.smoothScroll) {
							deltaX = Math.sign(deltaX) * Math.ceil(Math.abs(deltaX / 3));
						}
						self.scrollHorizontal(deltaX, event);
					}
					if (deltaY && (!self.smoothWheelCorrector || !self.smoothWheelCorrector.isBreakY())) {
						if (!wb.smoothScroll) {
							deltaY = Math.sign(deltaY) * Math.ceil(Math.abs(deltaY * self.settings.wheelScrollLinesV / 3));
						}
						self.scrollVertical(deltaY, event);
					}
					self._onMouseMove(event);
				});

			this.smoothWheelCorrector && this.smoothWheelCorrector.checkBreak();
			AscCommon.stopEvent(event);
			return true;
		};

		/** @param event {KeyboardEvent} */
		asc_CEventsController.prototype._onMouseDblClick = function (event) {
			if (this.handlers.trigger('isGlobalLockEditCell') || this.handlers.trigger('isFormatPainter')) {
				return false;
			}

			// The browser does not support the detail property (we will do it by coordinates)
			if (false === this.isDblClickInMouseDown) {
				return this.doMouseDblClick(event);
			}

			this.isDblClickInMouseDown = false;

			// Need to handle cursor display if dblClick was handled in MouseDown
			this.showCellEditorCursor();
			return true;
		};

		/** @param event */
		asc_CEventsController.prototype._getCoordinates = function (event) {
			// ToDo needs to be redone
			if (event.coord) {
				return event.coord;
			}

			var offs = AscCommon.UI.getBoundingClientRect(this.element);
			var x = ((event.pageX * AscBrowser.zoom) >> 0) - offs.left;
			var y = ((event.pageY * AscBrowser.zoom) >> 0) - offs.top;

			x *= AscCommon.AscBrowser.retinaPixelRatio;
			y *= AscCommon.AscBrowser.retinaPixelRatio;

			return {x: x, y: y};
		};

		asc_CEventsController.prototype._setSkipKeyPress = function (val) {
			this.skipKeyPress = val;
		};

		asc_CEventsController.prototype.showHorizontalScroll = function (val) {
			if (!this.hsb || !this.hsb.style) {
				return;
			}
			let toVisibility = val ? "visible" : "hidden";
			let res;
			if (this.hsb.style.visibility === toVisibility) {
				res = false;
			} else {
				this.hsb.style.visibility = toVisibility;
				res = true;
			}

			let isVisibleVerScroll = this.view.getShowVerticalScroll();
			let scrollWidth = this.view && this.view.defaults && this.view.defaults.scroll ? this.view.defaults.scroll.widthPx : 14;

			this.hsb.style.right = isVisibleVerScroll ? scrollWidth + "px" : "0px";

			if (!this.view.Api.isMobileVersion) {
				let cornerStyle = val && isVisibleVerScroll ? "visible" : "hidden";
				let corner = document.getElementById("ws-scrollbar-corner");
				if (corner) {
					corner.style.visibility = cornerStyle;
				}
			}
			return res;
		};

		asc_CEventsController.prototype.showVerticalScroll = function (val) {
			if (!this.vsb || !this.vsb.style) {
				return;
			}

			let toVisibility = val ? "visible" : "hidden";
			let res;
			if (this.vsb.style.visibility === toVisibility) {
				res = false;
			} else {
				this.vsb.style.visibility = toVisibility;
				res = true;
			}
			this.vsb.style.visibility = toVisibility;
			let isVisibleHorScroll = this.view.getShowHorizontalScroll();
			let scrollWidth = this.view && this.view.defaults && this.view.defaults.scroll ? this.view.defaults.scroll.heightPx : 14;

			this.vsb.style.bottom = isVisibleHorScroll ? scrollWidth + "px" : "0px";

			if (!this.view.Api.isMobileVersion) {
				let cornerStyle = val && isVisibleHorScroll ? "visible" : "hidden";
				let corner = document.getElementById("ws-scrollbar-corner");
				if (corner) {
					corner.style.visibility = cornerStyle;
				}
			}
			return res;
		};

		asc_CEventsController.prototype.setSkipCellEditor = function (val) {
			this.skipCellEditor = val;
		};

		//------------------------------------------------------------export---------------------------------------------------
		window['AscCommonExcel'] = window['AscCommonExcel'] || {};
		window["AscCommonExcel"].asc_CEventsController = asc_CEventsController;
	}
)(jQuery, window);
