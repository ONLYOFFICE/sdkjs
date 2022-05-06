/*
 * (c) Copyright Ascensio System SIA 2010-2019
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
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
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
  function(window, undefined) {


  /*
   * Import
   * -----------------------------------------------------------------------------
   */
  var c_oAscFormatPainterState = AscCommon.c_oAscFormatPainterState;
  var AscBrowser = AscCommon.AscBrowser;
  var CColor = AscCommon.CColor;
  var cBoolLocal = AscCommon.cBoolLocal;
  var History = AscCommon.History;

  var asc = window["Asc"];
  var asc_applyFunction = AscCommonExcel.applyFunction;
  var asc_round = asc.round;
  var asc_typeof = asc.typeOf;
  var asc_CMM = AscCommonExcel.asc_CMouseMoveData;
  var asc_CPrintPagesData = AscCommonExcel.CPrintPagesData;
  var c_oTargetType = AscCommonExcel.c_oTargetType;
  var c_oAscError = asc.c_oAscError;
  var c_oAscCleanOptions = asc.c_oAscCleanOptions;
  var c_oAscSelectionDialogType = asc.c_oAscSelectionDialogType;
  var c_oAscMouseMoveType = asc.c_oAscMouseMoveType;
  var c_oAscPopUpSelectorType = asc.c_oAscPopUpSelectorType;
  var c_oAscAsyncAction = asc.c_oAscAsyncAction;
  var c_oAscFontRenderingModeType = asc.c_oAscFontRenderingModeType;
  var c_oAscAsyncActionType = asc.c_oAscAsyncActionType;

  var g_clipboardExcel = AscCommonExcel.g_clipboardExcel;


  function WorkbookCommentsModel(handlers, aComments) {
    this.workbook = {handlers: handlers};
    this.aComments = aComments;
  }

  WorkbookCommentsModel.prototype.getId = function() {
    return null;
  };
  WorkbookCommentsModel.prototype.getMergedByCell = function() {
    return null;
  };

	function WorksheetViewSettings() {
		//TODO темные цвета необходимо скорректировать
		this.getCColor = function (_color) {
			var rgb = parseInt(_color.split('#')[1], 16);
			return new CColor((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
		};
		this.updateStyle = function () {
			this.header.style = this._generateStyle();
			this.header.groupDataBorder = this.getCColor(AscCommon.GlobalSkin.GroupDataBorder);
			this.header.editorBorder = this.getCColor(AscCommon.GlobalSkin.EditorBorder);
		};
		this._generateStyle = function () {
			return [// Header colors
				{ // kHeaderDefault
					background: this.getCColor(AscCommon.GlobalSkin.Background),
					border: this.getCColor(AscCommon.GlobalSkin.Border),
					color: this.getCColor(AscCommon.GlobalSkin.Color),
					backgroundDark: this.getCColor(AscCommon.GlobalSkin.BackgroundDark),
					colorDark: this.getCColor(AscCommon.GlobalSkin.ColorDark)
				}, { // kHeaderActive
					background: this.getCColor(AscCommon.GlobalSkin.BackgroundActive),
					border: this.getCColor(AscCommon.GlobalSkin.BorderActive),
					color: this.getCColor(AscCommon.GlobalSkin.ColorActive),
					backgroundDark: this.getCColor(AscCommon.GlobalSkin.BackgroundDarkActive),
					colorDark: this.getCColor(AscCommon.GlobalSkin.ColorDarkActive)
				}, { // kHeaderHighlighted
					background: this.getCColor(AscCommon.GlobalSkin.BackgroundHighlighted),
					border: this.getCColor(AscCommon.GlobalSkin.BorderHighlighted),
					color: this.getCColor(AscCommon.GlobalSkin.ColorHighlighted),
					backgroundDark: this.getCColor(AscCommon.GlobalSkin.BackgroundDarkHighlighted),
					colorDark: this.getCColor(AscCommon.GlobalSkin.ColorDarkHighlighted)
				}];
		};
		this.header = {
			style: this._generateStyle(),
			cornerColor: new CColor(193, 193, 193),
			groupDataBorder: this.getCColor(AscCommon.GlobalSkin.GroupDataBorder),
			editorBorder: this.getCColor(AscCommon.GlobalSkin.EditorBorder),
			printBackground: new CColor(238, 238, 238),
			printBorder: new CColor(216, 216, 216),
			printColor: new CColor(0, 0, 0)
		};
		this.cells = {
			defaultState: {
				background: new CColor(255, 255, 255), border: new CColor(202, 202, 202)
			}, padding: -1 /*px horizontal padding*/
		};
		this.activeCellBorderColor = new CColor(72, 121, 92);
		this.activeCellBorderColor2 = new CColor(255, 255, 255, 1);
		this.findFillColor = new CColor(255, 238, 128, 1);

		// Цвет закрепленных областей
		this.frozenColor = new CColor(105, 119, 62, 1);

		// Число знаков для математической информации
		this.mathMaxDigCount = 9;

		var cnv = document.createElement("canvas");
		cnv.width = 2;
		cnv.height = 2;
		var ctx = cnv.getContext("2d");
		ctx.clearRect(0, 0, 2, 2);
		ctx.fillStyle = "#000";
		ctx.fillRect(0, 0, 1, 1);
		ctx.fillRect(1, 1, 1, 1);
		this.ptrnLineDotted1 = ctx.createPattern(cnv, "repeat");

		this.halfSelection = false;

		return this;
	}


  /**
   * Widget for displaying and editing Workbook object
   * -----------------------------------------------------------------------------
   * @param {AscCommonExcel.Workbook} model                        Workbook
   * @param {AscCommonExcel.asc_CEventsController} controller          Events controller
   * @param {HandlersList} handlers                  Events handlers for WorkbookView events
   * @param {Element} elem                          Container element
   * @param {Element} inputElem                      Input element for top line editor
   * @param {Object} Api
   * @param {CCollaborativeEditing} collaborativeEditing
   * @param {c_oAscFontRenderingModeType} fontRenderingMode
   *
   * @constructor
   * @memberOf Asc
   */
  function WorkbookView(model, controller, handlers, elem, inputElem, Api, collaborativeEditing, fontRenderingMode) {
    this.defaults = {
      scroll: {
        widthPx: 14, heightPx: 14
      }, worksheetView: new WorksheetViewSettings()
    };

    this.model = model;
    this.enableKeyEvents = true;
    this.controller = controller;
    this.handlers = handlers;
    this.wsViewHandlers = null;
    this.element = elem;
    this.input = inputElem;
    this.Api = Api;
    this.collaborativeEditing = collaborativeEditing;
    this.lastSendInfoRange = null;
    this.oSelectionInfo = null;
    this.canUpdateAfterShiftUp = false;	// Нужно ли обновлять информацию после отпускания Shift
    this.keepType = false;
    this.timerId = null;
    this.timerEnd = false;

    //----- declaration -----
    this.isInit = false;
    this.canvas = undefined;
    this.canvasOverlay = undefined;
    this.canvasGraphic = undefined;
    this.canvasGraphicOverlay = undefined;
    this.wsActive = -1;
    this.wsMustDraw = false; // Означает, что мы выставили активный, но не отрисовали его
    this.wsViews = [];
    this.cellEditor = undefined;
    this.fontRenderingMode = null;
    this.lockDraw = false;		// Lock отрисовки на некоторое время

    this.isCellEditMode = false;
    this.isFormulaEditMode = false;
    this.isWizardMode = false;

	  this.isShowComments = true;
	  this.isShowSolved = true;

    this.formulasList = [];		// Список всех формул
    this.lastFPos = -1; 		// Последняя позиция формулы
    this.lastFNameLength = '';		// Последний кусок формулы
    this.skipHelpSelector = false;	// Пока true - не показываем окно подсказки
    // Константы для подстановке формулы (что не нужно добавлять скобки)
    this.arrExcludeFormulas = [];

    this.fReplaceCallback = null;	// Callback для замены текста

    // Фонт, который выставлен в DrawingContext, он должен быть один на все DrawingContext-ы
    this.m_oFont = AscCommonExcel.g_oDefaultFormat.Font.clone();

    // Теперь у нас 2 FontManager-а на весь документ + 1 для автофигур (а не на каждом листе свой)
    this.fmgrGraphics = [];						// FontManager for draw (1 для обычного + 1 для поворотного текста)
    this.fmgrGraphics.push(new AscFonts.CFontManager({mode:"cell"}));	// Для обычного
    this.fmgrGraphics.push(new AscFonts.CFontManager({mode:"cell"}));	// Для поворотного
    this.fmgrGraphics.push(new AscFonts.CFontManager());	// Для автофигур
    this.fmgrGraphics.push(new AscFonts.CFontManager({mode:"cell"}));	// Для измерений

    this.fmgrGraphics[0].Initialize(true); // IE memory enable
    this.fmgrGraphics[1].Initialize(true); // IE memory enable
    this.fmgrGraphics[2].Initialize(true); // IE memory enable
    this.fmgrGraphics[3].Initialize(true); // IE memory enable

    this.buffers = {};
    this.drawingCtx = undefined;
    this.overlayCtx = undefined;
    this.drawingGraphicCtx = undefined;
    this.overlayGraphicCtx = undefined;
    this.shapeCtx = null;
    this.shapeOverlayCtx = null;
    this.mainGraphics = undefined;
    this.stringRender = undefined;
    this.trackOverlay = null;
    this.mainOverlay = null;
    this.autoShapeTrack = null;

    this.formatPainterState = c_oAscFormatPainterState.kOff;
    this.formatPainterRange = null;
    this.formatPainterSheet = null;

    this.selectionDialogMode = false;
    this.dialogAbsName = false;
    this.dialogSheetName = false;
    this.copyActiveSheet = -1;

    // Комментарии для всего документа
    this.cellCommentator = null;

    // Флаг о подписке на эвенты о смене позиции документа (скролл) для меню
    this.isDocumentPlaceChangedEnabled = false;

    // Максимальная ширина числа из 0,1,2...,9, померенная в нормальном шрифте(дефалтовый для книги) в px(целое)
    // Ecma-376 Office Open XML Part 1, пункт 18.3.1.13
    this.maxDigitWidth = 0;
    //-----------------------

    this.MobileTouchManager = null;

    this.defNameAllowCreate = true;

    this._init(fontRenderingMode);

    this.autoCorrectStore = null;//объект для хранения параметров иконки авторазвертывания таблиц
	this.cutIdSheet = null;

    this.NeedUpdateTargetForCollaboration = true;
    this.LastUpdateTargetTime = 0;

	this.printPreviewState = new AscCommonExcel.CPrintPreviewState(this);
	this.printOptionsJson = null;

    return this;
  }

  WorkbookView.prototype._init = function(fontRenderingMode) {
    var self = this;

    // Init font managers rendering
    // Изначально мы инициализируем c_oAscFontRenderingModeType.hintingAndSubpixeling
    this.setFontRenderingMode(fontRenderingMode, /*isInit*/true);

    // add style
    var _head = document.getElementsByTagName('head')[0];
    var style0 = document.createElement('style');
    style0.type = 'text/css';
    style0.innerHTML = ".block_elem { position:absolute;padding:0;margin:0; }";
    _head.appendChild(style0);

    // create canvas
    if (null != this.element) {
		if (!this.Api.VersionHistory) {
			this.element.innerHTML = '<div id="ws-canvas-outer">\
											<canvas id="ws-canvas"></canvas>\
											<canvas id="ws-canvas-overlay"></canvas>\
											<canvas id="ws-canvas-graphic"></canvas>\
											<canvas id="ws-canvas-graphic-overlay"></canvas>\
											<div id="id_target_cursor" class="block_elem" width="1" height="1"\
												style="width:2px;height:13px;display:none;z-index:9;"></div>\
										</div>';
		}

      this.canvas = document.getElementById("ws-canvas");
      this.canvasOverlay = document.getElementById("ws-canvas-overlay");
      this.canvasGraphic = document.getElementById("ws-canvas-graphic");
      this.canvasGraphicOverlay = document.getElementById("ws-canvas-graphic-overlay");
    }

    this.buffers.main = new asc.DrawingContext({
      canvas: this.canvas, units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont
    });
    this.buffers.overlay = new asc.DrawingContext({
      canvas: this.canvasOverlay, units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont
    });

    this.buffers.mainGraphic = new asc.DrawingContext({
      canvas: this.canvasGraphic, units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont
    });
    this.buffers.overlayGraphic = new asc.DrawingContext({
      canvas: this.canvasGraphicOverlay, units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont
    });

    this.drawingCtx = this.buffers.main;
    this.overlayCtx = this.buffers.overlay;
    this.drawingGraphicCtx = this.buffers.mainGraphic;
    this.overlayGraphicCtx = this.buffers.overlayGraphic;

    this.shapeCtx = new AscCommon.CGraphics();
    this.shapeOverlayCtx = new AscCommon.CGraphics();
    this.mainGraphics = new AscCommon.CGraphics();
    this.trackOverlay = new AscCommon.COverlay();
    this.trackOverlay.IsCellEditor = true;
    if(this.Api.isMobileVersion) {
        this.mainOverlay = new AscCommon.COverlay();
        this.mainOverlay.IsCellEditor = true;
    }
    this.autoShapeTrack = new AscCommon.CAutoshapeTrack();
    this.shapeCtx.m_oAutoShapesTrack = this.autoShapeTrack;

    this.shapeCtx.m_oFontManager = this.fmgrGraphics[2];
    this.shapeOverlayCtx.m_oFontManager = this.fmgrGraphics[2];
    this.mainGraphics.m_oFontManager = this.fmgrGraphics[0];

    // Обновляем размеры (чуть ниже, потому что должны быть проинициализированы ctx)
    this._canResize();

    this.stringRender = new AscCommonExcel.StringRender(this.buffers.main);

    // Мерить нужно только со 100% и один раз для всего документа
    this._calcMaxDigitWidth();

	  if (!window["NATIVE_EDITOR_ENJINE"]) {
		  // initialize events controller
		  this.controller.init(this, this.element, /*this.canvasOverlay*/ this.canvasGraphicOverlay, /*handlers*/{
			  "resize": function () {
				  self.resize.apply(self, arguments);
			  }, "gotFocus": function (hasFocus) {
				  if (self.isCellEditMode) {
					  self.cellEditor.setFocus(!hasFocus);
				  }
			  }, "initRowsCount": function () {
				  self._onInitRowsCount.apply(self, arguments);
			  }, "initColsCount": function () {
				  self._onInitColsCount.apply(self, arguments);
			  }, "scrollY": function () {
				  self._onScrollY.apply(self, arguments);
			  }, "scrollX": function () {
				  self._onScrollX.apply(self, arguments);
			  }, "changeSelection": function () {
				  self._onChangeSelection.apply(self, arguments);
			  }, "changeSelectionDone": function () {
				  self._onChangeSelectionDone.apply(self, arguments);
			  }, "changeSelectionRightClick": function () {
				  self._onChangeSelectionRightClick.apply(self, arguments);
			  }, "selectionActivePointChanged": function () {
				  self._onSelectionActivePointChanged.apply(self, arguments);
			  }, "updateWorksheet": function () {
				  self._onUpdateWorksheet.apply(self, arguments);
			  }, "resizeElement": function () {
				  self._onResizeElement.apply(self, arguments);
			  }, "resizeElementDone": function () {
				  self._onResizeElementDone.apply(self, arguments);
			  }, "changeFillHandle": function () {
				  self._onChangeFillHandle.apply(self, arguments);
			  }, "changeFillHandleDone": function () {
				  self._onChangeFillHandleDone.apply(self, arguments);
			  }, "moveRangeHandle": function () {
				  self._onMoveRangeHandle.apply(self, arguments);
			  }, "moveRangeHandleDone": function () {
				  self._onMoveRangeHandleDone.apply(self, arguments);
			  }, "moveResizeRangeHandle": function () {
				  self._onMoveResizeRangeHandle.apply(self, arguments);
			  }, "moveResizeRangeHandleDone": function () {
				  self._onMoveResizeRangeHandleDone.apply(self, arguments);
			  }, "editCell": function () {
				  self._onEditCell.apply(self, arguments);
			  }, "stopCellEditing": function () {
				  return self.closeCellEditor.apply(self, arguments);
			  }, "isRestrictionComments": function () {
				  return self.Api.isRestrictionComments();
			  }, "empty": function () {
				  self._onEmpty.apply(self, arguments);
			  }, "undo": function () {
				  self.undo.apply(self, arguments);
			  }, "redo": function () {
				  self.redo.apply(self, arguments);
			  }, "mouseDblClick": function () {
				  self._onMouseDblClick.apply(self, arguments);
			  }, "showNextPrevWorksheet": function () {
				  self._onShowNextPrevWorksheet.apply(self, arguments);
			  }, "setFontAttributes": function () {
				  self._onSetFontAttributes.apply(self, arguments);
			  }, "setCellFormat": function () {
				  self._onSetCellFormat.apply(self, arguments);
			  }, "selectColumnsByRange": function () {
				  self._onSelectColumnsByRange.apply(self, arguments);
			  }, "selectRowsByRange": function () {
				  self._onSelectRowsByRange.apply(self, arguments);
			  }, "save": function () {
				  self.Api.asc_Save();
			  }, "showCellEditorCursor": function () {
				  self._onShowCellEditorCursor.apply(self, arguments);
			  }, "print": function () {
				  self.Api.onPrint();
			  }, "addFunction": function () {
				  self.insertInCellEditor.apply(self, arguments);
			  }, "canvasClick": function () {
				  self.enableKeyEventsHandler(true);
			  }, "autoFiltersClick": function () {
				  self._onAutoFiltersClick.apply(self, arguments);
			  }, "tableTotalClick": function () {
				  self._onTableTotalClick.apply(self, arguments);
			  }, "pivotFiltersClick": function () {
				  self._onPivotFiltersClick.apply(self, arguments);
			  }, "pivotCollapseClick": function () {
				  self._onPivotCollapseClick.apply(self, arguments);
			  }, "commentCellClick": function () {
				  self._onCommentCellClick.apply(self, arguments);
			  }, "isGlobalLockEditCell": function () {
				  return self.collaborativeEditing.getGlobalLockEditCell();
			  }, "updateSelectionName": function () {
				  self._onUpdateSelectionName.apply(self, arguments);
			  }, "stopFormatPainter": function () {
				  self._onStopFormatPainter.apply(self, arguments);
			  }, "groupRowClick": function () {
				  return self._onGroupRowClick.apply(self, arguments);
			  }, "onChangeTableSelection": function () {
				  return self._onChangeTableSelection.apply(self, arguments);
			  }, "getActiveCell": function () {
				  var ws = self.getWorksheet();
				  if (ws) {
				  	var selectionRanges = ws.getSelectedRanges();
				  	if (selectionRanges && selectionRanges.length === 1 && selectionRanges[0].bbox && selectionRanges[0].bbox.isOneCell()) {
						  return ws.getActiveCell(0, 0, false)
					  }
				  }
				  return null;
			  },


			  // Shapes
			  "graphicObjectMouseDown": function () {
				  self._onGraphicObjectMouseDown.apply(self, arguments);
			  }, "graphicObjectMouseMove": function () {
				  self._onGraphicObjectMouseMove.apply(self, arguments);
			  }, "graphicObjectMouseUp": function () {
				  self._onGraphicObjectMouseUp.apply(self, arguments);
			  }, "graphicObjectMouseUpEx": function () {
				  self._onGraphicObjectMouseUpEx.apply(self, arguments);
			  }, "graphicObjectWindowKeyDown": function () {
				  return self._onGraphicObjectWindowKeyDown.apply(self, arguments);
			  }, "graphicObjectWindowKeyUp": function () {
				  return self._onGraphicObjectWindowKeyUp.apply(self, arguments);
			  }, "graphicObjectWindowKeyPress": function () {
				  return self._onGraphicObjectWindowKeyPress.apply(self, arguments);
			  }, "graphicObjectMouseWheel": function () {
				  return self._onGraphicObjecMouseWheel.apply(self, arguments);
			  }, "getGraphicsInfo": function () {
				  return self._onGetGraphicsInfo.apply(self, arguments);
			  }, "updateSelectionShape": function () {
				  return self._onUpdateSelectionShape.apply(self, arguments);
			  }, "canReceiveKeyPress": function () {
				  return self.getWorksheet().objectRender.controller.canReceiveKeyPress();
			  }, "stopAddShape": function () {
				  self.getWorksheet().objectRender.controller.checkEndAddShape();
			  },

			  // Frozen anchor
			  "moveFrozenAnchorHandle": function () {
				  self._onMoveFrozenAnchorHandle.apply(self, arguments);
			  }, "moveFrozenAnchorHandleDone": function () {
				  self._onMoveFrozenAnchorHandleDone.apply(self, arguments);
			  },

			  // AutoComplete
			  "showAutoComplete": function () {
				  self.showAutoComplete.apply(self, arguments);
			  }, "onContextMenu": function (event) {
				  self.handlers.trigger("asc_onContextMenu", event);
			  },

			  // DataValidation
			  "onDataValidation": function () {
				  if (self.oSelectionInfo && self.oSelectionInfo.dataValidation) {
					  var list = self.oSelectionInfo.dataValidation.getListValues(self.model.getActiveWs());
					  if (list) {
						 self.handlers.trigger("asc_onValidationListMenu", list[0], list[1]);
					  }
					  return !!list;
				  }
			  },

			  // FormatPainter
			  'isFormatPainter': function () {
				  return self.formatPainterState;
			  },

			  //calculate
			  'calculate': function () {
			  	self.calculate.apply(self, arguments);
			  },

			  'changeFormatTableInfo': function () {
				  var table = self.getSelectionInfo().formatTableInfo;
				  return table && self.changeFormatTableInfo(table.tableName, Asc.c_oAscChangeTableStyleInfo.rowTotal, !table.lastRow);
			  },

			  //special paste
			  "hideSpecialPasteOptions": function () {
				  self.handlers.trigger("hideSpecialPasteOptions");
			  }
		  });

      if (this.input && this.input.addEventListener) {
        this.input.addEventListener("focus", function () {
          self.input.isFocused = true;
          if (!self.canEdit()) {
            return;
          }
		  if (self.isProtectActiveCell()) {
			  self.input.blur();
			  self.handlers.trigger("asc_onError", c_oAscError.ID.ChangeOnProtectedSheet, c_oAscError.Level.NoCritical);
			  return;
		  }
          self._onStopFormatPainter();
          self.cellEditor.callTopLineMouseup = true;
          if (!self.getCellEditMode() && !self.controller.isFillHandleMode) {
            var enterOptions = new AscCommonExcel.CEditorEnterOptions();
            enterOptions.focus = true;
            self._onEditCell(enterOptions);
          }
        }, false);

        this.input.addEventListener('keydown', function (event) {
          if (self.isCellEditMode) {
            self.handlers.trigger('asc_onInputKeyDown', event);
            if (!event.defaultPrevented) {
              self.cellEditor._onWindowKeyDown(event, true);
            }
          }
        }, false);
      }

      this.Api.onKeyDown = function (event) {
        self.controller._onWindowKeyDown(event);
        if (self.isCellEditMode) {
          self.cellEditor._onWindowKeyDown(event, false);
        }
      };
      this.Api.onKeyPress = function (event) {
        self.controller._onWindowKeyPress(event);
        if (self.isCellEditMode) {
          self.cellEditor._onWindowKeyPress(event);
        }
      };
      this.Api.onKeyUp = function (event) {
        self.controller._onWindowKeyUp(event);
        if (self.isCellEditMode) {
          self.cellEditor._onWindowKeyUp(event);
        }
      };
      this.Api.Begin_CompositeInput = function () {
        var oWSView = self.getWorksheet();
        if (oWSView && oWSView.isSelectOnShape) {
          if (oWSView.objectRender) {
            oWSView.objectRender.Begin_CompositeInput();
          }
          return;
        }

        if (!self.isCellEditMode) {
          var enterOptions = new AscCommonExcel.CEditorEnterOptions();
          enterOptions.newText = '';
          enterOptions.quickInput = true;
          self._onEditCell(enterOptions, function () {
            self.cellEditor.Begin_CompositeInput();
          });
        } else {
          self.cellEditor.Begin_CompositeInput();
        }
      };
      this.Api.Replace_CompositeText = function (arrCharCodes) {
        var oWSView = self.getWorksheet();
        if(oWSView && oWSView.isSelectOnShape){
          if(oWSView.objectRender){
            oWSView.objectRender.Replace_CompositeText(arrCharCodes);
          }
          return;
        }
        if (self.isCellEditMode) {
          self.cellEditor.Replace_CompositeText(arrCharCodes);
        }
      };
      this.Api.End_CompositeInput = function () {
        var oWSView = self.getWorksheet();
        if(oWSView && oWSView.isSelectOnShape){
          if(oWSView.objectRender){
            oWSView.objectRender.End_CompositeInput();
          }
          return;
        }
        if (self.isCellEditMode) {
          self.cellEditor.End_CompositeInput();
        }
      };
      this.Api.Set_CursorPosInCompositeText = function (nPos) {
        var oWSView = self.getWorksheet();
        if(oWSView && oWSView.isSelectOnShape){
          if(oWSView.objectRender){
            oWSView.objectRender.Set_CursorPosInCompositeText(nPos);
          }
          return;
        }
        if (self.isCellEditMode) {
          self.cellEditor.Set_CursorPosInCompositeText(nPos);
        }
      };
      this.Api.Get_CursorPosInCompositeText = function () {
        var res = 0;
        var oWSView = self.getWorksheet();
        if(oWSView && oWSView.isSelectOnShape){
          if(oWSView.objectRender){
            res = oWSView.objectRender.Get_CursorPosInCompositeText();
          }
        }
        else if (self.isCellEditMode) {
          res = self.cellEditor.Get_CursorPosInCompositeText();
        }
        return res;
      };
      this.Api.Get_MaxCursorPosInCompositeText = function () {
        var res = 0; var oWSView = self.getWorksheet();
        if(oWSView && oWSView.isSelectOnShape){
          if(oWSView.objectRender){
            res = oWSView.objectRender.Get_CursorPosInCompositeText();
          }
        }
        else if (self.isCellEditMode) {
          res = self.cellEditor.Get_MaxCursorPosInCompositeText();
        }
        return res;
      };
      this.Api.AddTextWithPr = function (familyName, arrCharCodes) {
      	var ws = self.getWorksheet();
      	if (ws && ws.isSelectOnShape) {
      		var textPr = new CTextPr();
			textPr.RFonts = new CRFonts();
			textPr.RFonts.SetAll(familyName, -1);
      		ws.objectRender.controller.addTextWithPr(new AscCommon.CUnicodeStringEmulator(arrCharCodes), textPr, true);
      		return;
      	}

      	if (!self.isCellEditMode) {
      		self._onEditCell(new AscCommonExcel.CEditorEnterOptions(), function () {
				self.cellEditor.setTextStyle('fn', familyName);
				self.cellEditor._addCharCodes(arrCharCodes);
			});
		} else {
			self.cellEditor.setTextStyle('fn', familyName);
			self.cellEditor._addCharCodes(arrCharCodes);
		}
	  };
      this.Api.beginInlineDropTarget = function (event) {
      	if (!self.controller.isMoveRangeMode) {
      		self.controller.isMoveRangeMode = true;
			self.getWorksheet().dragAndDropRange = new Asc.Range(0, 0, 0, 0);
		}
      	self.controller._onMouseMove(event);
	  };
      this.Api.endInlineDropTarget = function (event) {
      	var ws = self.getWorksheet();
      	if (!ws.activeMoveRange) {
      		return;
      	}
      	self.controller.isMoveRangeMode = false;
      	var newSelection = ws.activeMoveRange.clone();
      	ws._cleanSelectionMoveRange();
      	ws.dragAndDropRange = null;
      	self._onSetSelection(newSelection);
	  };
      this.Api.isEnabledDropTarget = function () {
      	return !self.isCellEditMode;
	  };


      AscCommon.InitBrowserInputContext(this.Api, "id_target_cursor");
      AscCommonExcel.executeInR1C1Mode(false, function () {
          self.model.dependencyFormulas.calcTree();
      });
    }

	  this.cellEditor =
		  new AscCommonExcel.CellEditor(this.element, this.input, this.fmgrGraphics, this.m_oFont, /*handlers*/{
			  "closed": function () {
				  self._onCloseCellEditor.apply(self, arguments);
			  }, "updated": function () {
				  self.Api.checkLastWork();
				  self._onUpdateCellEditor.apply(self, arguments);
			  }, "gotFocus": function (hasFocus) {
				  self.controller.setFocus(!hasFocus);
			  }, "updateFormulaEditMod": function (val) {
                  self.setFormulaEditMode(val);
			  }, "updateEditorState": function (state) {
				  self.handlers.trigger("asc_onEditCell", state);
			  }, "updateTopLine": function (state) {
                  // Implemented through asc_onEditCell
                  self.handlers.trigger("asc_onEditCell", state);
              }, "isGlobalLockEditCell": function () {
				  return self.collaborativeEditing.getGlobalLockEditCell();
			  }, "onMouseDown": function (event) {
			  	return self.controller._onMouseDown(event);
			  }, "newRanges": function (ranges) {
			      if (self.isActive()) {
			          var ws = self.getWorksheet();
			          ws.cleanSelection();
			          ws.oOtherRanges = ranges;
			          ws._drawSelection();
				  }
			  }, "cleanSelectRange": function () {
			      self._onCleanSelectRange();
              }, "updateUndoRedoChanged": function (bCanUndo, bCanRedo) {
				  self.handlers.trigger("asc_onCanUndoChanged", bCanUndo);
				  self.handlers.trigger("asc_onCanRedoChanged", bCanRedo);
			  }, "applyCloseEvent": function () {
				  self.controller._onWindowKeyDown.apply(self.controller, arguments);
			  }, "canEdit": function () {
				  return self.canEdit();
			  }, "isProtectActiveCell": function () {
				  return self.isProtectActiveCell();
			  }, "getFormulaRanges": function () {
			      return self.isActive() ? self.getWorksheet().oOtherRanges : null;
			  }, "isActive": function () {
				  return self.isActive();
			  }, "getWizard": function () {
			      return self.isWizardMode;
              }, "getActiveWS": function () {
			      return self.getActiveWS();
			  }, "updateEditorSelectionInfo": function (xfs) {
				  self.handlers.trigger("asc_onEditorSelectionChanged", xfs);
			  }, "onContextMenu": function (event) {
				  self.handlers.trigger("asc_onContextMenu", event);
			  }, "updatedEditableFunction": function (fName, pos) {
				  self.handlers.trigger("asc_onFormulaInfo", fName, pos);
			  }, "onSelectionEnd" : function () {
				  self.handlers.trigger("asc_onSelectionEnd");
        }
		  }, this.defaults.worksheetView.cells.padding);

	  this.wsViewHandlers = new AscCommonExcel.asc_CHandlersList(/*handlers*/{
		  "getViewMode": function () {
			  return self.Api.getViewMode();
		  }, "isRestrictionComments": function () {
			  return self.Api.isRestrictionComments();
		  }, "reinitializeScroll": function (type) {
			  self._onScrollReinitialize(type);
		  }, "selectionChanged": function () {
			  self._onWSSelectionChanged();
		  }, "selectionNameChanged": function () {
			  self._onSelectionNameChanged.apply(self, arguments);
		  }, "selectionMathInfoChanged": function () {
			  self._onSelectionMathInfoChanged.apply(self, arguments);
		  }, 'onFilterInfo': function (countFilter, countRecords) {
			  self.handlers.trigger("asc_onFilterInfo", countFilter, countRecords);
		  }, "onErrorEvent": function (errorId, level) {
			  self.handlers.trigger("asc_onError", errorId, level);
		  }, "slowOperation": function (isStart) {
			  (isStart ? self.Api.sync_StartAction : self.Api.sync_EndAction).call(self.Api,
				  c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.SlowOperation);
		  }, "setAutoFiltersDialog": function (arrVal) {
			  self.handlers.trigger("asc_onSetAFDialog", arrVal);
		  }, "selectionRangeChanged": function (val) {
		      self._onSelectionRangeChanged(val);
		  }, "onRenameCellTextEnd": function (countFind, countReplace) {
			  self.handlers.trigger("asc_onRenameCellTextEnd", countFind, countReplace);
		  }, 'onStopFormatPainter': function () {
              self._onStopFormatPainter.apply(self, arguments);
		  }, "onDocumentPlaceChanged": function () {
			  self._onDocumentPlaceChanged();
		  }, "updateSheetViewSettings": function () {
			  self.handlers.trigger("asc_onUpdateSheetViewSettings");
		  }, "onScroll": function (d) {
			  self.controller.scroll(d);
		  }, "getLockDefNameManagerStatus": function () {
			  return self.defNameAllowCreate;
		  }, 'isActive': function () {
			  return self.isActive();
		  }, "drawMobileSelection": function (oOverlay, oColor) {
			  if (self.MobileTouchManager) {
				  self.MobileTouchManager.CheckSelect(oOverlay, oColor);
			  }
		  }, "showSpecialPasteOptions": function (val) {
			  self.handlers.trigger("asc_onShowSpecialPasteOptions", val);
			  if (!window['AscCommon'].g_specialPasteHelper.showSpecialPasteButton) {
				  window['AscCommon'].g_specialPasteHelper.showSpecialPasteButton = true;
			  }
		  }, 'checkLastWork': function () {
			  self.Api.checkLastWork();
		  }, "toggleAutoCorrectOptions": function (bIsShow, val) {
		      self.toggleAutoCorrectOptions(bIsShow, val);
		  }, "selectSearchingResults": function () {
			  return self.Api.selectSearchingResults;
		  }, "getMainGraphics": function () {
			  return self.mainGraphics;
		  }, "cleanCutData": function (bDrawSelection, bCleanBuffer) {
			  self.cleanCutData(bDrawSelection, bCleanBuffer);
		  }
	  });

	  this.model.handlers.add("changeSheetViewSettings", function (wsId, type) {
		  var ws = self.getWorksheetById(wsId, true);
		  if (ws) {
			  ws._onChangeSheetViewSettings(type);
		  }
	  });

    this.model.handlers.add("cleanCellCache", function(wsId, oRanges, skipHeight, needResetCache) {
      var ws = self.getWorksheetById(wsId, true);
      if (ws) {
        if (needResetCache) {
          ws.cache && ws.cache.reset();
        }
        ws.updateRanges(oRanges, skipHeight);
      }
    });
    this.model.handlers.add("changeWorksheetUpdate", function(wsId, val) {
      var ws = self.getWorksheetById(wsId);
      if (ws) {
        ws.changeWorksheet("update", val);
      }
    });
    this.model.handlers.add("showWorksheet", function(wsId) {
      var wsModel = self.model.getWorksheetById(wsId), index;
      if (wsModel) {
        index = wsModel.getIndex();
        self.showWorksheet(index, true);
      }
    });
    this.model.handlers.add("setSelection", function() {
      self._onSetSelection.apply(self, arguments);
    });
    this.model.handlers.add("getSelectionState", function() {
      return self._onGetSelectionState.apply(self);
    });
    this.model.handlers.add("setSelectionState", function() {
      self._onSetSelectionState.apply(self, arguments);
    });
    this.model.handlers.add("drawWS", function() {
      self.drawWS.apply(self, arguments);
    });
    this.model.handlers.add("scrollToTopLeftCell", function() {
      self.scrollToTopLeftCell.apply(self, arguments);
    });
    this.model.handlers.add("showDrawingObjects", function() {
      self.onShowDrawingObjects.apply(self, arguments);
    });
    this.model.handlers.add("setCanUndo", function(bCanUndo) {
      self.handlers.trigger("asc_onCanUndoChanged", bCanUndo);
    });
    this.model.handlers.add("setCanRedo", function(bCanRedo) {
      self.handlers.trigger("asc_onCanRedoChanged", bCanRedo);
    });
    this.model.handlers.add("setDocumentModified", function(bIsModified) {
      self.Api.onUpdateDocumentModified(bIsModified);
    });
    this.model.handlers.add("updateWorksheetByModel", function() {
      self.updateWorksheetByModel.apply(self, arguments);
    });
    this.model.handlers.add("undoRedoAddRemoveRowCols", function(sheetId, type, range, bUndo) {
      if (true === bUndo) {
        if (AscCH.historyitem_Worksheet_AddRows === type) {
          self.collaborativeEditing.removeRowsRange(sheetId, range.clone(true));
          self.collaborativeEditing.undoRows(sheetId, range.r2 - range.r1 + 1);
        } else if (AscCH.historyitem_Worksheet_RemoveRows === type) {
          self.collaborativeEditing.addRowsRange(sheetId, range.clone(true));
          self.collaborativeEditing.undoRows(sheetId, range.r2 - range.r1 + 1);
        } else if (AscCH.historyitem_Worksheet_AddCols === type) {
          self.collaborativeEditing.removeColsRange(sheetId, range.clone(true));
          self.collaborativeEditing.undoCols(sheetId, range.c2 - range.c1 + 1);
        } else if (AscCH.historyitem_Worksheet_RemoveCols === type) {
          self.collaborativeEditing.addColsRange(sheetId, range.clone(true));
          self.collaborativeEditing.undoCols(sheetId, range.c2 - range.c1 + 1);
        }
      } else {
        if (AscCH.historyitem_Worksheet_AddRows === type) {
          self.collaborativeEditing.addRowsRange(sheetId, range.clone(true));
          self.collaborativeEditing.addRows(sheetId, range.r1, range.r2 - range.r1 + 1);
        } else if (AscCH.historyitem_Worksheet_RemoveRows === type) {
          self.collaborativeEditing.removeRowsRange(sheetId, range.clone(true));
          self.collaborativeEditing.removeRows(sheetId, range.r1, range.r2 - range.r1 + 1);
        } else if (AscCH.historyitem_Worksheet_AddCols === type) {
          self.collaborativeEditing.addColsRange(sheetId, range.clone(true));
          self.collaborativeEditing.addCols(sheetId, range.c1, range.c2 - range.c1 + 1);
        } else if (AscCH.historyitem_Worksheet_RemoveCols === type) {
          self.collaborativeEditing.removeColsRange(sheetId, range.clone(true));
          self.collaborativeEditing.removeCols(sheetId, range.c1, range.c2 - range.c1 + 1);
        }
      }
    });
    this.model.handlers.add("undoRedoHideSheet", function(sheetId) {
      self.showWorksheet(sheetId);
      // Посылаем callback об изменении списка листов
      self.Api.sheetsChanged();
    });
    this.model.handlers.add("updateSelection", function () {
		if (!self.lockDraw) {
			self.getWorksheet().updateSelection();
		}
	});

    this.handlers.add("asc_onLockDefNameManager", function(reason) {
      self.defNameAllowCreate = !(reason == Asc.c_oAscDefinedNameReason.LockDefNameManager);
    });
    this.handlers.add('addComment', function (id, data) {
      self._onWSSelectionChanged();
      self.handlers.trigger('asc_onAddComment', id, data);
    });
    this.handlers.add('removeComment', function (id) {
      self._onWSSelectionChanged();
      self.handlers.trigger('asc_onRemoveComment', id);
    });
    this.handlers.add('hiddenComments', function () {
      return !self.isShowComments;
    });
	  this.handlers.add('showSolved', function () {
		  return self.isShowSolved;
	  });
	this.model.handlers.add("hideSpecialPasteOptions", function() {
		self.handlers.trigger("asc_onHideSpecialPasteOptions");
    });
	this.model.handlers.add("toggleAutoCorrectOptions", function(bIsShow, val) {
		self.toggleAutoCorrectOptions(bIsShow, val);
	});
	this.model.handlers.add("cleanCutData", function(bDrawSelection, bCleanBuffer) {
	  self.cleanCutData(bDrawSelection, bCleanBuffer);
	});
	this.model.handlers.add("updateGroupData", function() {
	  self.updateGroupData();
	});
	this.model.handlers.add("updatePrintPreview", function() {
	  self.updatePrintPreview();
	});
    this.cellCommentator = new AscCommonExcel.CCellCommentator({
      model: new WorkbookCommentsModel(this.handlers, this.model.aComments),
      collaborativeEditing: this.collaborativeEditing,
      draw: function() {
      },
      handlers: {
        trigger: function() {
          return false;
        }
      }
    });
    if (0 < this.model.aComments.length) {
      this.handlers.trigger("asc_onAddComments", this.model.aComments);
    }

    this.initFormulasList();

    this.fReplaceCallback = function() {
      self._replaceCellTextCallback.apply(self, arguments);
    };

    return this;
  };

  WorkbookView.prototype.destroy = function() {
    this.controller.destroy();
    this.cellEditor.destroy();
    return this;
  };

  WorkbookView.prototype._createWorksheetView = function(wsModel) {
    return new AscCommonExcel.WorksheetView(this, wsModel, this.wsViewHandlers, this.buffers, this.stringRender, this.maxDigitWidth, this.collaborativeEditing, this.defaults.worksheetView);
  };

  WorkbookView.prototype._onSelectionNameChanged = function (name) {
    this.handlers.trigger("asc_onSelectionNameChanged", name);
  };

  WorkbookView.prototype._onSelectionMathInfoChanged = function (info) {
    this.handlers.trigger("asc_onSelectionMathChanged", info);
  };

  WorkbookView.prototype._onSelectionRangeChanged = function (val) {
      if (this.isFormulaEditMode && !this.isWizardMode) {
          this.skipHelpSelector = true;
          this.cellEditor.changeCellText(val);
          this.skipHelpSelector = false;
      }
      this.handlers.trigger("asc_onSelectionRangeChanged", val);
  };

  WorkbookView.prototype._onCleanSelectRange = function (force) {
      if (this.selectionDialogMode) {
          var ws = this.getWorksheet();
          if (ws.model.selectionRange || force) {
              ws.cleanSelection();
              ws.model.selectionRange = null;
			  if (this.isActive()) {
				  ws._drawSelection();
			  }
          }
      }
  };

  // Проверяет, сменили ли мы диапазон (для того, чтобы не отправлять одинаковую информацию о диапазоне)
  WorkbookView.prototype._isEqualRange = function(range, isSelectOnShape) {
    if (null === this.lastSendInfoRange) {
      return false;
    }

    return this.lastSendInfoRange.isEqual(range) && this.lastSendInfoRangeIsSelectOnShape === isSelectOnShape;
  };

  WorkbookView.prototype._updateSelectionInfo = function () {
    if (this.selectionDialogMode) {
      return false;
    }
	if(this.Api && this.Api.noCreatePoint) {
		return false;
	}
    var ws = this.getWorksheet();
    this.oSelectionInfo = ws.getSelectionInfo();
    this.lastSendInfoRange = ws.model.selectionRange.clone();
    this.lastSendInfoRangeIsSelectOnShape = ws.getSelectionShape();
    this.updateTargetForCollaboration();

    return true;
  };
  WorkbookView.prototype._onWSSelectionChanged = function(isSaving) {
    if (!this._updateSelectionInfo()) {
      return;
    }

    // При редактировании ячейки не нужно пересылать изменения
    if (this.input && !this.getCellEditMode()) {
      // Сами запретим заходить в строку формул, когда выделен shape
      if (this.lastSendInfoRangeIsSelectOnShape) {
        this.input.disabled = true;
        this.input.value = '';
      } else {
        this.input.disabled = false;
        this.input.value = this.oSelectionInfo.text;
      }
    }
    this.handlers.trigger("asc_onSelectionChanged", this.oSelectionInfo);
    this.handlers.trigger("asc_onSelectionEnd");
    this._onInputMessage();
    if (!isSaving) {
      this.Api.cleanSpelling();
    }
  };

  WorkbookView.prototype._onInputMessage = function () {
  	var title = null, message = null;
  	var dataValidation = this.oSelectionInfo && this.oSelectionInfo.dataValidation;
  	if (dataValidation && dataValidation.showInputMessage && !this.model.getActiveWs().getDisablePrompts()) {
  		title = dataValidation.promptTitle;
		message = dataValidation.prompt;
	}
  	this.handlers.trigger("asc_onInputMessage", title, message);
  };

	WorkbookView.prototype._onScrollReinitialize = function (type) {
		if (window["NATIVE_EDITOR_ENJINE"] || !type) {
			return;
		}

		var ws = this.getWorksheet();
		if (AscCommonExcel.c_oAscScrollType.ScrollHorizontal & type) {
			this.controller.reinitScrollX(this.controller.hsbApi.settings, ws.getFirstVisibleCol(true), ws.getHorizontalScrollRange(), ws.getHorizontalScrollMax());
		}
		if (AscCommonExcel.c_oAscScrollType.ScrollVertical & type) {
			this.controller.reinitScrollY(this.controller.vsbApi.settings, ws.getFirstVisibleRow(true), ws.getVerticalScrollRange(), ws.getVerticalScrollMax());
		}

		if (this.Api.isMobileVersion) {
			this.MobileTouchManager.Resize();
		}
	};

	WorkbookView.prototype._onInitRowsCount = function () {
		var ws = this.getWorksheet();
		if (ws._initRowsCount()) {
			this._onScrollReinitialize(AscCommonExcel.c_oAscScrollType.ScrollVertical);
		}
	};

	WorkbookView.prototype._onInitColsCount = function () {
		var ws = this.getWorksheet();
		if (ws._initColsCount()) {
			this._onScrollReinitialize(AscCommonExcel.c_oAscScrollType.ScrollHorizontal);
		}
	};

  WorkbookView.prototype._onScrollY = function(pos, initRowsCount) {
    var ws = this.getWorksheet();
    var delta = asc_round(pos - ws.getFirstVisibleRow(true));
    if (delta !== 0) {
      ws.scrollVertical(delta, this.cellEditor, initRowsCount);
    }
  };

  WorkbookView.prototype._onScrollX = function(pos, initColsCount) {
    var ws = this.getWorksheet();
    var delta = asc_round(pos - ws.getFirstVisibleCol(true));
    if (delta !== 0) {
      ws.scrollHorizontal(delta, this.cellEditor, initColsCount);
    }
  };

  WorkbookView.prototype._onSetSelection = function(range) {
    var ws = this.getWorksheet();
    ws._endSelectionShape();
    ws.setSelection(range);
  };

  WorkbookView.prototype._onGetSelectionState = function() {
    var res = null;
    var ws = this.getWorksheet(null, true);
    if (ws && AscCommon.isRealObject(ws.objectRender) && AscCommon.isRealObject(ws.objectRender.controller)) {
      res = ws.objectRender.controller.getSelectionState();
    }
    return (res && res[0] && res[0].focus) ? res : null;
  };

  WorkbookView.prototype._onSetSelectionState = function(state) {
    if (null !== state) {
      var ws = this.getWorksheetById(state[0].worksheetId);
      if (ws && ws.objectRender && ws.objectRender.controller) {
        ws.objectRender.controller.setSelectionState(state);
        ws.setSelectionShape(true);
        ws._scrollToRange(ws.objectRender.getSelectedDrawingsRange());
        ws.objectRender.showDrawingObjectsEx();
        ws.objectRender.controller.updateOverlay();
        ws.objectRender.controller.updateSelectionState();
      }
      // Селектим после выставления состояния
    }
  };

    WorkbookView.prototype._onChangeSelection = function (isStartPoint, dc, dr, isCoord, isCtrl, callback) {
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
        if (isCoord && !this.timerEnd && this.timerId === null) {
            this.timerId = setTimeout(function () {
                var arrClose = [];
                arrClose.push(new asc_CMM({type: c_oAscMouseMoveType.None}));
                t.handlers.trigger("asc_onMouseMove", arrClose);
                t._onUpdateCursor(AscCommonExcel.kCurCells);
                t.timerId = null;
                t.timerEnd = true;
            }, 1000);
        }

        asc_applyFunction(callback, d);
    };

  // Окончание выделения
  WorkbookView.prototype._onChangeSelectionDone = function(x, y, event) {
  	if (this.timerId !== null) {
		clearTimeout(this.timerId);
		this.timerId = null;
	}
  	this.keepType = false;
    if (this.selectionDialogMode) {
      return;
    }
    var formatPainterState = this.formatPainterState;
    var ws = this.getWorksheet();
    ws.changeSelectionDone();
    this._onSelectionNameChanged(ws.getSelectionName(/*bRangeText*/false));
    // Проверим, нужно ли отсылать информацию о ячейке
    var ar = ws.model.selectionRange.getLast();
    var isSelectOnShape = ws.getSelectionShape();
    if (!this._isEqualRange(ws.model.selectionRange, isSelectOnShape)) {
      this._onWSSelectionChanged();
      this._onSelectionMathInfoChanged(ws.getSelectionMathInfo());
      this.controller.lastTab = null;
    }

    // Нужно очистить поиск
    this.model.cleanFindResults();

    var ct = ws.getCursorTypeFromXY(x, y);

    if(!this.isFormulaEditMode && !formatPainterState) {
        if (c_oTargetType.Hyperlink === ct.target) {
            // Проверим замерженность
            var isHyperlinkClick = false;
            if(isSelectOnShape) {
                var button = 0;
                if(event) {
                    button = AscCommon.getMouseButton(event);
                }
                if(button === 0) {
                    isHyperlinkClick = true;
                }
            }
            else if(ar.isOneCell()) {
                isHyperlinkClick = true;
            }
            else {
                var mergedRange = ws.model.getMergedByCell(ar.r1, ar.c1);
                if (mergedRange && ar.isEqual(mergedRange)) {
                    isHyperlinkClick = true;
                }
            }
            if (isHyperlinkClick && !this.timerEnd) {
                if (false === ct.hyperlink.hyperlinkModel.getVisited() && !isSelectOnShape) {
                    ct.hyperlink.hyperlinkModel.setVisited(true);
                    if (ct.hyperlink.hyperlinkModel.Ref) {
                        ws._updateRange(ct.hyperlink.hyperlinkModel.Ref.getBBox0());
                        ws.draw();
                    }
                }
                switch (ct.hyperlink.asc_getType()) {
                    case Asc.c_oAscHyperlinkType.WebLink:
                        this.handlers.trigger("asc_onHyperlinkClick", ct.hyperlink.asc_getHyperlinkUrl());
                        break;
                    case Asc.c_oAscHyperlinkType.RangeLink:
                        // ToDo надо поправить отрисовку комментария для данной ячейки (с которой уходим)
                        this.handlers.trigger("asc_onHideComment");
                        this.Api._asc_setWorksheetRange(ct.hyperlink);
                        break;
                }
            }
        }
        else if(isSelectOnShape && c_oTargetType.Shape === ct.target &&  ct.macro) {
            var button = 0;
            if(event) {
                button = AscCommon.getMouseButton(event);
            }
            if(button === 0) {
                if(!this.timerEnd) {
                    this.Api.asc_runMacros(ct.macro);
                }
            }
        }
    }
    this.timerEnd = false;
  };

  // Обработка нажатия правой кнопки мыши
  WorkbookView.prototype._onChangeSelectionRightClick = function(dc, dr, target) {
    var ws = this.getWorksheet();
    ws.changeSelectionStartPointRightClick(dc, dr, target);
  };

  // Обработка движения в выделенной области
  WorkbookView.prototype._onSelectionActivePointChanged = function(dc, dr, callback) {
    var ws = this.getWorksheet();
    var d = ws.changeSelectionActivePoint(dc, dr);
    asc_applyFunction(callback, d);
  };

  WorkbookView.prototype._onUpdateWorksheet = function(x, y, ctrlKey, callback) {
    var ws = this.getWorksheet(), ct = undefined;
    var arrMouseMoveObjects = [];					// Теперь это массив из объектов, над которыми курсор
	var t = this;
    //ToDo: включить определение target, если находимся в режиме редактирования ячейки.
    if (x === undefined && y === undefined) {
    	ws.cleanHighlightedHeaders();
		if (this.timerId === null) {
    	this.timerId = setTimeout(function () {
    			var arrClose = [];
				arrClose.push(new asc_CMM({type: c_oAscMouseMoveType.None}));
				t.handlers.trigger("asc_onMouseMove", arrClose);
				t.timerId = null;
			}, 1000);
		}
    } else {
      ct = ws.getCursorTypeFromXY(x, y);
      ct.coordX = x;
      ct.coordY = y;
      if (this.timerId !== null) {
      	clearTimeout(this.timerId);
      	this.timerId = null;
      }

      // Отправление эвента об удалении всего листа (именно удалении, т.к. если просто залочен, то не рисуем рамку вокруг)
      if (undefined !== ct.userIdAllSheet) {
        arrMouseMoveObjects.push(new asc_CMM({
          type: c_oAscMouseMoveType.LockedObject,
          x: AscCommon.AscBrowser.convertToRetinaValue(ct.lockAllPosLeft),
          y: AscCommon.AscBrowser.convertToRetinaValue(ct.lockAllPosTop),
          userId: ct.userIdAllSheet,
          lockedObjectType: Asc.c_oAscMouseMoveLockedObjectType.Sheet
        }));
      } else {
        // Отправление эвента о залоченности свойств всего листа (только если не удален весь лист)
        if (undefined !== ct.userIdAllProps) {
          arrMouseMoveObjects.push(new asc_CMM({
            type: c_oAscMouseMoveType.LockedObject,
            x: AscCommon.AscBrowser.convertToRetinaValue(ct.lockAllPosLeft),
            y: AscCommon.AscBrowser.convertToRetinaValue(ct.lockAllPosTop),
            userId: ct.userIdAllProps,
            lockedObjectType: Asc.c_oAscMouseMoveLockedObjectType.TableProperties
          }));
        }
      }
      // Отправление эвента о наведении на залоченный объект
      if (undefined !== ct.userId) {
        arrMouseMoveObjects.push(new asc_CMM({
          type: c_oAscMouseMoveType.LockedObject,
          x: AscCommon.AscBrowser.convertToRetinaValue(ct.lockRangePosLeft),
          y: AscCommon.AscBrowser.convertToRetinaValue(ct.lockRangePosTop),
          userId: ct.userId,
          lockedObjectType: Asc.c_oAscMouseMoveLockedObjectType.Range
        }));
      }

		// Отправление эвента о наведении на залоченный объект
		if (undefined !== ct.userIdForeignSelect) {
			arrMouseMoveObjects.push(new asc_CMM({
				type: c_oAscMouseMoveType.ForeignSelect,
				x: AscCommon.AscBrowser.convertToRetinaValue(ct.foreignSelectPosLeft),
				y: AscCommon.AscBrowser.convertToRetinaValue(ct.foreignSelectPosTop),
				userId: ct.userIdForeignSelect,
				color: AscCommon.getUserColorById(ct.shortIdForeignSelect, null, true)
			}));
		}


      // Проверяем комментарии ячейки
      if (ct.commentIndexes) {
        arrMouseMoveObjects.push(new asc_CMM({
          type: c_oAscMouseMoveType.Comment,
          x: ct.commentCoords.dLeftPX,
          reverseX: ct.commentCoords.dReverseLeftPX,
          y: ct.commentCoords.dTopPX,
          aCommentIndexes: ct.commentIndexes
        }));
      }
      // Проверяем гиперссылку
      if (ct.target === c_oTargetType.Hyperlink) {
		  if (!ctrlKey) {
			  arrMouseMoveObjects.push(new asc_CMM({
				  type: c_oAscMouseMoveType.Hyperlink,
				  x: AscCommon.AscBrowser.convertToRetinaValue(x),
				  y: AscCommon.AscBrowser.convertToRetinaValue(y),
				  hyperlink: ct.hyperlink
			  }));
		  } else {
			  ct.cursor = AscCommonExcel.kCurCells;
		  }
	  }

		// проверяем фильтр
		if (ct.target === c_oTargetType.FilterObject) {
			var filterObj;
			if (ct.idPivot) {
				filterObj = ws.pivot_setDialogProp(ct.idPivot);
			} else {
				filterObj = ws.af_setDialogProp(ct.idFilter, true);
			}
			if(filterObj) {
				arrMouseMoveObjects.push(new asc_CMM({
					type: c_oAscMouseMoveType.Filter,
					x: AscCommon.AscBrowser.convertToRetinaValue(x),
					y: AscCommon.AscBrowser.convertToRetinaValue(y),
					filter: filterObj
				}));
			}
		}

		// check shape
		if (ct.target === c_oTargetType.Shape) {
		    if(typeof ct.tooltip === "string" && ct.tooltip.length > 0) {
                arrMouseMoveObjects.push(new asc_CMM({
                    type: c_oAscMouseMoveType.Tooltip,
                    x: AscCommon.AscBrowser.convertToRetinaValue(x),
                    y: AscCommon.AscBrowser.convertToRetinaValue(y),
                    tooltip: ct.tooltip
                }));
            }
		}

      /* Проверяем, может мы на никаком объекте (такая схема оказалась приемлимой
       * для отдела разработки приложений)
       */
      if (0 === arrMouseMoveObjects.length) {
        // Отправляем эвент, что мы ни на какой области
        arrMouseMoveObjects.push(new asc_CMM({type: c_oAscMouseMoveType.None}));
      }
      // Отсылаем эвент с объектами
      this.handlers.trigger("asc_onMouseMove", arrMouseMoveObjects);

      if (ct.target === c_oTargetType.MoveRange && ctrlKey && ct.cursor === "move") {
        ct.cursor = "copy";
      }

      this._onUpdateCursor(ct.cursor);
      if (ct.target === c_oTargetType.ColumnHeader || ct.target === c_oTargetType.RowHeader) {
        ws.drawHighlightedHeaders(ct.col, ct.row);
      } else {
        ws.cleanHighlightedHeaders();
      }
    }
    asc_applyFunction(callback, ct);
  };

  WorkbookView.prototype._onUpdateCursor = function (cursor) {
  	var newHtmlCursor = AscCommon.g_oHtmlCursor.value(cursor);
  	if (this.element.style.cursor !== newHtmlCursor) {
		this.element.style.cursor = newHtmlCursor;
  	}
  };

  WorkbookView.prototype._onResizeElement = function(target, x, y) {
    var arrMouseMoveObjects = [];
    if (target.target === c_oTargetType.ColumnResize) {
      arrMouseMoveObjects.push(this.getWorksheet().drawColumnGuides(target.col, x, y, target.mouseX));
    } else if (target.target === c_oTargetType.RowResize) {
      arrMouseMoveObjects.push(this.getWorksheet().drawRowGuides(target.row, x, y, target.mouseY));
    }

    /* Проверяем, может мы на никаком объекте (такая схема оказалась приемлимой
     * для отдела разработки приложений)
     */
    if (0 === arrMouseMoveObjects.length) {
      // Отправляем эвент, что мы ни на какой области
      arrMouseMoveObjects.push(new asc_CMM({type: c_oAscMouseMoveType.None}));
    }
    // Отсылаем эвент с объектами
    this.handlers.trigger("asc_onMouseMove", arrMouseMoveObjects);
  };

  WorkbookView.prototype._onResizeElementDone = function(target, x, y, isResizeModeMove) {
    var ws = this.getWorksheet();
    if (isResizeModeMove) {
      if (target.target === c_oTargetType.ColumnResize) {
        ws.changeColumnWidth(target.col, x, target.mouseX);
      } else if (target.target === c_oTargetType.RowResize) {
        ws.changeRowHeight(target.row, y, target.mouseY);
      }
      window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Update_Position();
      this._onDocumentPlaceChanged();
    }
    ws.draw();

    // Отсылаем окончание смены размеров (в FF не срабатывало обычное движение)
    this.handlers.trigger("asc_onMouseMove", [new asc_CMM({type: c_oAscMouseMoveType.None})]);
  };

  // Обработка автозаполнения
  WorkbookView.prototype._onChangeFillHandle = function(x, y, callback, tableIndex) {
    var ws = this.getWorksheet();
    var d = ws.changeSelectionFillHandle(x, y, tableIndex);
    asc_applyFunction(callback, d);
  };

  // Обработка окончания автозаполнения
  WorkbookView.prototype._onChangeFillHandleDone = function(x, y, ctrlPress) {
    var ws = this.getWorksheet();
    ws.applyFillHandle(x, y, ctrlPress);
  };

  // Обработка перемещения диапазона
  WorkbookView.prototype._onMoveRangeHandle = function(x, y, callback) {
    var ws = this.getWorksheet();
    var d = ws.changeSelectionMoveRangeHandle(x, y);
    asc_applyFunction(callback, d);
  };

  // Обработка окончания перемещения диапазона
  WorkbookView.prototype._onMoveRangeHandleDone = function(ctrlKey) {
    var ws = this.getWorksheet();
    ws.applyMoveRangeHandle(ctrlKey);
  };

  WorkbookView.prototype._onMoveResizeRangeHandle = function(x, y, target, callback) {
    var ws = this.getWorksheet();
    var d = ws.changeSelectionMoveResizeRangeHandle(x, y, target, this.cellEditor);
    asc_applyFunction(callback, d);
  };

  WorkbookView.prototype._onMoveResizeRangeHandleDone = function() {
    var ws = this.getWorksheet();
    ws.applyMoveResizeRangeHandle();
  };

  // Frozen anchor
  WorkbookView.prototype._onMoveFrozenAnchorHandle = function(x, y, target) {
    var ws = this.getWorksheet();
    ws.drawFrozenGuides(x, y, target);
  };

  WorkbookView.prototype._onMoveFrozenAnchorHandleDone = function(x, y, target) {
    // Закрепляем область
    var ws = this.getWorksheet();
    ws.applyFrozenAnchor(x, y, target);
  };

  WorkbookView.prototype.showAutoComplete = function() {
    var ws = this.getWorksheet();
    var arrValues = ws.getCellAutoCompleteValues(ws.model.selectionRange.activeCell);
    this.handlers.trigger('asc_onEntriesListMenu', arrValues);
  };

  WorkbookView.prototype._onAutoFiltersClick = function(idFilter) {
    this.getWorksheet().af_setDialogProp(idFilter);
  };

  WorkbookView.prototype._onTableTotalClick = function(idTableTotal) {
      var ws = this.getWorksheet();
      if (idTableTotal) {
          var _table = ws.model.TableParts ? ws.model.TableParts[idTableTotal.id] : null;
          if (_table) {
            var _tableColumn = _table.TableColumns[idTableTotal.colId];
            if (_tableColumn) {
                var val = _tableColumn.TotalsRowFunction;
                if (null === val) {
                    val = Asc.ETotalsRowFunction.totalrowfunctionNone;
                }
                this.handlers.trigger("asc_onTableTotalMenu", val);
            }
          }
      }
  };

  WorkbookView.prototype._onPivotFiltersClick = function(idPivot) {
    var filterObj = this.getWorksheet().pivot_setDialogProp(idPivot);
    if (filterObj) {
      this.handlers.trigger("asc_onSetAFDialog", filterObj);
    }
  };
  WorkbookView.prototype._onPivotCollapseClick = function(idPivotCollapse) {
    if (!idPivotCollapse || idPivotCollapse.hidden) {
      return;
    }
    var pivotTable = this.model.getPivotTableById(idPivotCollapse.id);
    if (!pivotTable) {
      return;
    }
    pivotTable.setVisibleFieldItem(this.Api, !idPivotCollapse.sd, idPivotCollapse.fld, idPivotCollapse.index);
  };

  WorkbookView.prototype._onGroupRowClick = function(x, y, target, type) {
  	return this.getWorksheet().groupRowClick(x, y, target, type);
  };

  WorkbookView.prototype._onChangeTableSelection = function(target) {
      var ws = this.getWorksheet();
      if (ws && ws.model && target) {
          var table = ws.model.TableParts[target.tableIndex];
          if (table) {
              ws.changeTableSelection(table.DisplayName, target.type, target.row, target.col);
          }
      }
  };

  WorkbookView.prototype._onCommentCellClick = function(x, y) {
    this.getWorksheet().cellCommentator.showCommentByXY(x, y);
  };

  WorkbookView.prototype._onUpdateSelectionName = function (forcibly) {
    if (this.canUpdateAfterShiftUp || forcibly) {
      this.canUpdateAfterShiftUp = false;
      var ws = this.getWorksheet();
      this._onSelectionNameChanged(ws.getSelectionName(/*bRangeText*/false));
    }
  };

  WorkbookView.prototype._onStopFormatPainter = function (bLockDraw) {
    if (this.formatPainterState) {
      this.formatPainter(c_oAscFormatPainterState.kOff, bLockDraw);
    }
  };

  // Shapes
  WorkbookView.prototype._onGraphicObjectMouseDown = function(e, x, y) {
    var ws = this.getWorksheet();
    ws.objectRender.graphicObjectMouseDown(e, x, y);
  };

  WorkbookView.prototype._onGraphicObjectMouseMove = function(e, x, y) {
    var ws = this.getWorksheet();
    ws.objectRender.graphicObjectMouseMove(e, x, y);
  };

  WorkbookView.prototype._onGraphicObjectMouseUp = function(e, x, y) {
    var ws = this.getWorksheet();
    ws.objectRender.graphicObjectMouseUp(e, x, y);
  };

  WorkbookView.prototype._onGraphicObjectMouseUpEx = function(e, x, y) {
    //var ws = this.getWorksheet();
    //ws.objectRender.calculateCell(x, y);
  };

  WorkbookView.prototype._onGraphicObjectWindowKeyDown = function(e) {
    var objectRender = this.getWorksheet().objectRender;
    return (0 < objectRender.getSelectedGraphicObjects().length) ? objectRender.graphicObjectKeyDown(e) : false;
  };
  WorkbookView.prototype._onGraphicObjectWindowKeyUp = function(e) {
    var oWS, bRet = false;
    for(var i in this.wsViews) {
      oWS = this.wsViews[i];
      if(oWS && oWS.objectRender) {
          bRet |= oWS.objectRender.graphicObjectKeyUp(e);
      }
    }
    return bRet;
  };

  WorkbookView.prototype._onGraphicObjectWindowKeyPress = function(e) {
    var objectRender = this.getWorksheet().objectRender;
    return (0 < objectRender.getSelectedGraphicObjects().length) ? objectRender.graphicObjectKeyPress(e) : false;
  };
  WorkbookView.prototype._onGraphicObjecMouseWheel = function(deltaX, deltaY) {
    var objectRender = this.getWorksheet().objectRender;
      if(objectRender && objectRender.controller) {
          return objectRender.controller.onMouseWheel(deltaX, deltaY);
      }
    return false;
  };

  WorkbookView.prototype._onGetGraphicsInfo = function(x, y) {
    var ws = this.getWorksheet();
    return ws.objectRender.checkCursorDrawingObject(x, y);
  };

  WorkbookView.prototype._onUpdateSelectionShape = function(isSelectOnShape) {
    var ws = this.getWorksheet();
    return ws.setSelectionShape(isSelectOnShape);
  };

  // Double click
  WorkbookView.prototype._onMouseDblClick = function(x, y, event, callback) {
    var ws = this.getWorksheet();
    var ct = ws.getCursorTypeFromXY(x, y);

    if (ct.target === c_oTargetType.ColumnResize || ct.target === c_oTargetType.RowResize) {
      if (ct.target === c_oTargetType.ColumnResize) {
      	ws.autoFitColumnsWidth(ct.col);
      } else {
      	ws.autoFitRowHeight(ct.row, ct.row);
      }
      asc_applyFunction(callback);
    } else {
      // In view mode or click on column | row | all | frozenMove | drawing object do not process
      if (!this.canEdit() || c_oTargetType.ColumnHeader === ct.target || c_oTargetType.RowHeader === ct.target ||
		  c_oTargetType.Corner === ct.target || c_oTargetType.FrozenAnchorH === ct.target ||
		  c_oTargetType.FrozenAnchorV === ct.target || ws.objectRender.checkCursorDrawingObject(x, y)) {
        asc_applyFunction(callback);
        return;
      }

      // При dbl клике фокус выставляем в зависимости от наличия текста в ячейке
      var enterOptions = new AscCommonExcel.CEditorEnterOptions();
      enterOptions.focus = null;
      enterOptions.eventPos = event;
      this._onEditCell(enterOptions);
    }
  };

	WorkbookView.prototype._onWindowMouseUpExternal = function (event, x, y) {
		this.controller._onWindowMouseUpExternal(event, x, y);
		if (this.isCellEditMode) {
			this.cellEditor._onWindowMouseUp(event, x, y);
		}
	};

  WorkbookView.prototype._onEditCell = function (enterOptions, callback) {
    var t = this;

    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || this.controller.isResizeMode || !this.canEdit()) {
      return;
    }

    var ws = t.getWorksheet();
    var activeCellRange = ws.getActiveCell(0, 0, false);
    var selectionRange = ws.model.selectionRange.clone();

    var activeWsModel = this.model.getActiveWs();
    if (activeWsModel.inPivotTable(activeCellRange)) {
		if (t.input.isFocused) {
			t.input.blur();
		}
		this.handlers.trigger("asc_onError", c_oAscError.ID.LockedCellPivot, c_oAscError.Level.NoCritical);
		return;
	}

    var editFunction = function() {
      if (needBlur) {
		  t.input && t.input.focus();
	  }
      t.setCellEditMode(true);
      t.hideSpecialPasteButton();
      ws.openCellEditor(t.cellEditor, enterOptions, selectionRange);
      t.input.disabled = false;

      t.Api.cleanSpelling(true);

      t.updateTargetForCollaboration();
      t.sendCursor(true);

      asc_applyFunction(callback, true);
    };

    var editLockCallback = function(res) {
      if (!res) {
        t.setCellEditMode(false);
        t.input.disabled = true;

        // Выключаем lock для редактирования ячейки
        t.collaborativeEditing.onStopEditCell();
        t.cellEditor.close(false);
        t._onWSSelectionChanged();
      }
    };

    var doEdit = function (success) {
		if (!success) {
			return;
		}
    	// Стартуем редактировать ячейку
		activeCellRange = ws.expandActiveCellByFormulaArray(activeCellRange);
		if (ws._isLockedCells(activeCellRange, /*subType*/null, editLockCallback)) {
			editFunction();
		}
	};

	var needBlur = false;
	var needBlurFunc;
	if (t.input && t.input.isFocused) {
		needBlurFunc = function () {
			t.input && t.input.blur();
			needBlur = true;
		}
	}
    ws.checkProtectRangeOnEdit([new Asc.Range(activeCellRange.c1, activeCellRange.r1, activeCellRange.c1, activeCellRange.r1)], doEdit, null, needBlurFunc);
  };

  WorkbookView.prototype._checkStopCellEditorInFormulas = function() {
	  if (this.isCellEditMode && this.isFormulaEditMode && this.cellEditor.canEnterCellRange()) {
		  return true;
	  }
      return this.closeCellEditor();
  };

  WorkbookView.prototype._onCloseCellEditor = function() {
    var isCellEditMode = this.getCellEditMode();
    this.setCellEditMode(false);

    if (isCellEditMode) {
      if (window['IS_NATIVE_EDITOR']) {
          window["native"]["closeCellEditor"]();
      }
    }

    // Обновляем состояние Undo/Redo
    History._sendCanUndoRedo();

    var ws = this.getWorksheet();
    ws.cleanSelection();
    ws.updateSelectionWithSparklines();
    // Обновляем состояние информации
    this._onWSSelectionChanged();

    // Закрываем подбор формулы
    if (-1 !== this.lastFPos) {
      this.handlers.trigger('asc_onFormulaCompleteMenu', null);
      this.lastFPos = -1;
      this.lastFNameLength = 0;
    }
    this.handlers.trigger('asc_onFormulaInfo', null);

    this.updateTargetForCollaboration();
    this.sendCursor();
  };

  WorkbookView.prototype._onEmpty = function() {
	  var ws = this.getWorksheet();
	  if (ws) {
		  var doDelete = function (success) {
			  if (success) {
				  ws.emptySelection(c_oAscCleanOptions.Text);
			  }
		  };
		  var selection = ws._getSelection();
		  var ranges = selection && selection.ranges;
		  if (!ws.objectRender.selectedGraphicObjectsExists()) {
			  ws.checkProtectRangeOnEdit(ranges, doDelete);
		  } else {
			  doDelete(true);
		  }
	  }
  };

  WorkbookView.prototype._onShowNextPrevWorksheet = function(direction) {
    // Колличество листов
    var countWorksheets = this.model.getWorksheetCount();
    // Покажем следующий лист или предыдущий (если больше нет)
    var i = this.wsActive + direction, ws;
    while (i !== this.wsActive) {
      if (0 > i) {
        i = countWorksheets - 1;
      } else  if (i >= countWorksheets) {
        i = 0;
      }

      ws = this.model.getWorksheet(i);
      if (!ws.getHidden()) {
        this.showWorksheet(i);
        return true;
      }

      i += direction;
    }
    return false;
  };

  WorkbookView.prototype._onSetFontAttributes = function(prop) {
    var val;
    var xfs = this.getSelectionInfo().asc_getXfs();
    switch (prop) {
      case "b":
        val = !(xfs.asc_getFontBold());
        break;
      case "i":
        val = !(xfs.asc_getFontItalic());
        break;
      case "u":
        // ToDo для двойного подчеркивания нужно будет немного переделать схему
        val = !(xfs.asc_getFontUnderline());
        val = val ? Asc.EUnderline.underlineSingle : Asc.EUnderline.underlineNone;
        break;
      case "s":
        val = !(xfs.asc_getFontStrikeout());
        break;
    }
    return this.setFontAttributes(prop, val);
  };

	WorkbookView.prototype._onSetCellFormat = function (prop) {
	  var info = new Asc.asc_CFormatCellsInfo();
	  info.asc_setSymbol(AscCommon.g_oDefaultCultureInfo.LCID);
	  info.asc_setType(Asc.c_oAscNumFormatType.None);
	  var formats = AscCommon.getFormatCells(info);
	  this.setCellFormat(formats[prop]);
	};

  WorkbookView.prototype._onSelectColumnsByRange = function() {
    this.getWorksheet()._selectColumnsByRange();
  };

  WorkbookView.prototype._onSelectRowsByRange = function() {
    this.getWorksheet()._selectRowsByRange();
  };

  WorkbookView.prototype._onShowCellEditorCursor = function() {
    // Показываем курсор
    if (this.getCellEditMode()) {
      this.cellEditor.showCursor();
    }
  };

  WorkbookView.prototype._onDocumentPlaceChanged = function() {
    if (this.isDocumentPlaceChangedEnabled) {
      this.handlers.trigger("asc_onDocumentPlaceChanged");
    }
  };

  WorkbookView.prototype.getCellStyles = function(width, height) {
  	return AscCommonExcel.generateCellStyles(width, height, this);
  };
  WorkbookView.prototype.getSlicerStyles = function() {
      return AscCommonExcel.generateSlicerStyles(36, 48, this);
  };

  WorkbookView.prototype.getWorksheetById = function(id, onlyExist) {
    var wsModel = this.model.getWorksheetById(id);
    if (wsModel) {
      return this.getWorksheet(wsModel.getIndex(), onlyExist);
    }
    return null;
  };

  WorkbookView.prototype.getActiveWS = function () {
      return this.model.getWorksheet(-1 === this.copyActiveSheet ? this.wsActive : this.copyActiveSheet)
  };

  /**
   * @param {Number} [index]
   * @param {Boolean} [onlyExist]
   * @return {AscCommonExcel.WorksheetView}
   */
  WorkbookView.prototype.getWorksheet = function(index, onlyExist) {
    var wb = this.model;
    var i = asc_typeof(index) === "number" && index >= 0 ? index : wb.getActive();
    var ws = this.wsViews[i];
    if (!ws && !onlyExist) {
      ws = this.wsViews[i] = this._createWorksheetView(wb.getWorksheet(i));
      ws._prepareComments();
      ws._prepareDrawingObjects();
    }
    return ws;
  };

	WorkbookView.prototype.drawWorksheet = function () {
		if (-1 === this.wsActive) {
			return this.showWorksheet();
		}
		var ws = this.getWorksheet();
		ws.draw();
		ws.objectRender.controller.updateSelectionState();
		ws.objectRender.controller.updateOverlay();
		this._onScrollReinitialize(AscCommonExcel.c_oAscScrollType.ScrollVertical | AscCommonExcel.c_oAscScrollType.ScrollHorizontal);
	};

  /**
   *
   * @param index
   * @param [bLockDraw]
   * @returns {WorkbookView}
   */
  WorkbookView.prototype.showWorksheet = function (index, bLockDraw) {
  	if (AscCommon.isFileBuild()) {
		return this;
	}
    // ToDo disable method for assembly
	var ws, wb = this.model;
	if (asc_typeof(index) !== "number" || 0 > index) {
      index = wb.getActive();
	}
    if (index === this.wsActive) {
		if (!bLockDraw) {
			this.drawWorksheet();
		}
      return this;
    }

    var selectionRange = null;
    // Только если есть активный
    if (-1 !== this.wsActive) {
      ws = this.getWorksheet();
      // Останавливаем ввод данных в редакторе ввода. Если в режиме ввода формул, то продолжаем работать с cellEditor'ом, чтобы можно было
      // выбирать ячейки для формулы
      if (!this._checkStopCellEditorInFormulas()) {
          index = this.copyActiveSheet;
      }
      // Делаем очистку селекта
      ws.cleanSelection();
      this.stopTarget(ws);

      if (this.selectionDialogMode) {
          // Copy selection to set on new sheet
          if (ws.model.selectionRange) {
              selectionRange = ws.model.selectionRange.getLast().clone(true);
          }
          ws.cloneSelection(false);
      }
    }

    if (index !== wb.getActive()) {
      wb.setActive(index);
    }
    this.wsActive = index;
    this.wsMustDraw = bLockDraw;

    // Посылаем эвент о смене активного листа
    this.handlers.trigger("asc_onActiveSheetChanged", this.wsActive);
    this.handlers.trigger("asc_onHideComment");

    ws = this.getWorksheet(index);
    if (this.selectionDialogMode) {
      // Когда идет выбор диапазона, то на показываемом листе должны выставить нужный режим
      ws.cloneSelection(true, selectionRange);
      this._onSelectionRangeChanged(ws.getSelectionRangeValue());
    }

    // Мы делали resize или меняли zoom, но не перерисовывали данный лист (он был не активный)
    if (ws.updateResize && ws.updateZoom) {
      ws.changeZoomResize();
    } else if (ws.updateResize) {
      ws.resize(true, this.cellEditor);
    } else if (ws.updateZoom) {
      ws.changeZoom(true);
    }

    this.updateGroupData();

    if (this.cellEditor && this.isFormulaEditMode) {
      if (this.isActive()) {
        this.cellEditor._showCanvas();
      } else {
        this.cellEditor._hideCanvas();
      }
    }

    if (!bLockDraw) {
      ws.draw();
      ws.objectRender.controller.updateSelectionState();
      ws.objectRender.controller.updateOverlay();
    }

    if (!window["NATIVE_EDITOR_ENJINE"] || window["IS_NATIVE_EDITOR"]) {
      this._onSelectionNameChanged(ws.getSelectionName(/*bRangeText*/false));
      this._onWSSelectionChanged();
      this._onSelectionMathInfoChanged(ws.getSelectionMathInfo());
    }
    this._onScrollReinitialize(AscCommonExcel.c_oAscScrollType.ScrollVertical | AscCommonExcel.c_oAscScrollType.ScrollHorizontal);
    // Zoom теперь на каждом листе одинаковый, не отправляем смену

    //TODO при добавлении любого действия в историю (например добавление нового листа), мы можем его потом отменить с повощью опции авторазвертывания
    this.toggleAutoCorrectOptions(null, true);
    window['AscCommon'].g_specialPasteHelper.SpecialPasteButton_Hide();
    return this;
  };

  WorkbookView.prototype.stopTarget = function(ws) {
    if (null === ws && -1 !== this.wsActive) {
      ws = this.getWorksheet(this.wsActive);
    }
    if (null !== ws && ws.objectRender && ws.objectRender.drawingDocument) {
      ws.objectRender.drawingDocument.TargetEnd();
    }
  };

  WorkbookView.prototype.updateWorksheetByModel = function() {
    // ToDo Сделал небольшую заглушку с показом листа. Нужно как мне кажется перейти от wsViews на wsViewsId (хранить по id)
    var oldActiveWs;
    if (-1 !== this.wsActive) {
      oldActiveWs = this.wsViews[this.wsActive];
    }

    //расставляем ws так как они идут в модели.
    var oNewWsViews = [];
    for (var i in this.wsViews) {
      var item = this.wsViews[i];
      if (null != item && null != this.model.getWorksheetById(item.model.getId())) {
        oNewWsViews[item.model.getIndex()] = item;
      }
    }
    this.wsViews = oNewWsViews;
    var wsActive = this.model.getActive();

    var newActiveWs = this.wsViews[wsActive];
    if (undefined === newActiveWs || oldActiveWs !== newActiveWs) {
      // Если сменили, то покажем
      this.wsActive = -1;
      this.showWorksheet(wsActive, true);
    } else {
      this.wsActive = wsActive;
    }
  };

  WorkbookView.prototype._canResize = function() {
    var styleWidth, styleHeight;
    styleWidth = this.element.offsetWidth - (this.Api.isMobileVersion ? 0 : this.defaults.scroll.widthPx);
    styleHeight = this.element.offsetHeight - (this.Api.isMobileVersion ? 0 : this.defaults.scroll.heightPx);

    this.isInit = true;

    this.canvas.style.width = this.canvasOverlay.style.width = this.canvasGraphic.style.width = this.canvasGraphicOverlay.style.width = styleWidth + 'px';
    this.canvas.style.height = this.canvasOverlay.style.height = this.canvasGraphic.style.height = this.canvasGraphicOverlay.style.height = styleHeight + 'px';

    AscCommon.calculateCanvasSize(this.canvas);
    AscCommon.calculateCanvasSize(this.canvasOverlay);
    AscCommon.calculateCanvasSize(this.canvasGraphic);
    AscCommon.calculateCanvasSize(this.canvasGraphicOverlay);

    this._reInitGraphics();

    return true;
  };

  WorkbookView.prototype._reInitGraphics = function () {
  	var canvasWidth = this.canvasGraphic.width;
  	var canvasHeight = this.canvasGraphic.height;
  	this.shapeCtx.init(this.drawingGraphicCtx.ctx, canvasWidth, canvasHeight, canvasWidth * 25.4 / this.drawingGraphicCtx.ppiX, canvasHeight * 25.4 / this.drawingGraphicCtx.ppiY);
  	this.shapeCtx.CalculateFullTransform();

  	var overlayWidth = this.canvasGraphicOverlay.width;
  	var overlayHeight = this.canvasGraphicOverlay.height;
  	this.shapeOverlayCtx.init(this.overlayGraphicCtx.ctx, overlayWidth, overlayHeight, overlayWidth * 25.4 / this.overlayGraphicCtx.ppiX, overlayHeight * 25.4 / this.overlayGraphicCtx.ppiY);
  	this.shapeOverlayCtx.CalculateFullTransform();

  	this.mainGraphics.init(this.drawingCtx.ctx, canvasWidth, canvasHeight, canvasWidth * 25.4 / this.drawingCtx.ppiX, canvasHeight * 25.4 / this.drawingCtx.ppiY);

  	this.trackOverlay.init(this.shapeOverlayCtx.m_oContext, "ws-canvas-graphic-overlay", 0, 0, overlayWidth, overlayHeight, (overlayWidth * 25.4 / this.overlayGraphicCtx.ppiX), (overlayHeight * 25.4 / this.overlayGraphicCtx.ppiY));
  	this.autoShapeTrack.init(this.trackOverlay, 0, 0, overlayWidth, overlayHeight, overlayWidth * 25.4 / this.overlayGraphicCtx.ppiX, overlayHeight * 25.4 / this.overlayGraphicCtx.ppiY);
  	this.autoShapeTrack.Graphics.CalculateFullTransform();
    if(this.mainOverlay) {
        overlayWidth = this.canvasOverlay.width;
        overlayHeight = this.canvasOverlay.height;
        this.mainOverlay.init(this.overlayCtx.ctx, "ws-canvas-overlay", 0, 0, overlayWidth, overlayHeight, (overlayWidth * 25.4 / this.overlayCtx.ppiX), (overlayHeight * 25.4 / this.overlayCtx.ppiY))
    }
  };

  /** @param event {jQuery.Event} */
  WorkbookView.prototype.resize = function(event) {
    if (this._canResize()) {
      var item;
      var activeIndex = this.model.getActive();
      for (var i in this.wsViews) {
        item = this.wsViews[i];
        // Делаем resize (для не активных сменим как только сделаем его активным)
        item.resize(/*isDraw*/i == activeIndex, this.cellEditor);
      }
      this.drawWorksheet();
    } else {
      // ToDo не должно происходить ничего, но нам приходит resize сверху, поэтому проверим отрисовывали ли мы
      if (-1 === this.wsActive || this.wsMustDraw) {
        this.drawWorksheet();
      }
    }
    this.wsMustDraw = false;
  };

  WorkbookView.prototype.getSelectionInfo = function () {
    if (!this.oSelectionInfo) {
      this._updateSelectionInfo();
    }
    return this.oSelectionInfo;
  };

  // Получаем свойство: редактируем мы сейчас или нет
  WorkbookView.prototype.getCellEditMode = function() {
	  return this.isCellEditMode;
  };

	WorkbookView.prototype.canEdit = function() {
		return this.Api.canEdit();
	};

	WorkbookView.prototype.isProtectActiveCell = function() {
		var ws = this.getWorksheet();
		if (!ws) {
			return false;
		}
		return ws.isProtectActiveCell();
	};

    WorkbookView.prototype.getDialogSheetName = function () {
        return this.dialogSheetName || !this.isActive();
    };

	WorkbookView.prototype.setCellEditMode = function(mode) {
		this.isCellEditMode = !!mode;
		if (!this.isCellEditMode) {
			this.setWizardMode(false);
            this.setFormulaEditMode(false);
        }
	};

    WorkbookView.prototype.setFormulaEditMode = function (mode) {
        this.isFormulaEditMode = mode;
        this.setSelectionDialogMode(mode ? c_oAscSelectionDialogType.Function : c_oAscSelectionDialogType.None, '');
    };

	WorkbookView.prototype.setWizardMode = function (mode) {
		if (mode !== this.isWizardMode) {
			this.isWizardMode = mode;
			if (this.isWizardMode) {
				this.cellEditor.updateWizardMode(this.isWizardMode);
				this._onCleanSelectRange(true);
			}
			this.input.disabled = this.isWizardMode;
		}
	};

	WorkbookView.prototype.isActive = function () {
		return (-1 === this.copyActiveSheet || this.wsActive === this.copyActiveSheet);
	};

	WorkbookView.prototype.isDrawFormatPainter = function () {
	    return this.formatPainterState && this.formatPainterSheet === this.wsActive;
    };
    WorkbookView.prototype.getFormatPainterSheet = function () {
        return this.formatPainterState && this.model.getWorksheet(this.formatPainterSheet);
    };

  WorkbookView.prototype.getIsTrackShape = function() {
    var ws = this.getWorksheet();
    if (!ws) {
      return false;
    }
    if (ws.objectRender && ws.objectRender.controller) {
      return ws.objectRender.controller.checkTrackDrawings();
    }
  };

  WorkbookView.prototype.getZoom = function() {
    return this.drawingCtx.getZoom();
  };

  WorkbookView.prototype.changeZoom = function(factor, changeZoomOnPrint, doNotDraw) {
  	if (factor === this.getZoom()) {
      return;
    }

    this.buffers.main.changeZoom(factor);
    this.buffers.overlay.changeZoom(factor);
    this.buffers.mainGraphic.changeZoom(factor);
    this.buffers.overlayGraphic.changeZoom(factor);
    if (!factor) {
      this.cellEditor.changeZoom(factor);
    }

    // Нужно сбросить кэш букв
    var i, length;
    for (i = 0, length = this.fmgrGraphics.length; i < length; ++i)
      this.fmgrGraphics[i].ClearFontsRasterCache();

    if (AscCommon.g_fontManager) {
        AscCommon.g_fontManager.ClearFontsRasterCache();
        AscCommon.g_fontManager.m_pFont = null;
    }
    if (AscCommon.g_fontManager2) {
        AscCommon.g_fontManager2.ClearFontsRasterCache();
        AscCommon.g_fontManager2.m_pFont = null;
    }

    if (!factor) {
		this.wsMustDraw = true;
		//this._calcMaxDigitWidth();
	}

    var item;
    var activeIndex = this.model.getActive();
    for (i in this.wsViews) {
      item = this.wsViews[i];
      // Меняем zoom (для не активных сменим как только сделаем его активным)
      if (!factor) {
      	item._initWorksheetDefaultWidth();
	  }
	  //changeZoomOnPrint - добавляю, поскольку при конвертации файлов есть проблемы. проверить, возможно отрисовка тоже не нужна
      item.changeZoom(/*isDraw*/i == activeIndex, changeZoomOnPrint);
      this._reInitGraphics();
      item.objectRender.changeZoom(this.drawingCtx.scaleFactor);
      if (i == activeIndex && factor && !doNotDraw) {
        item.draw();
        //ToDo item.drawDepCells();
      }
    }

    this._onScrollReinitialize(AscCommonExcel.c_oAscScrollType.ScrollVertical | AscCommonExcel.c_oAscScrollType.ScrollHorizontal);
    this.handlers.trigger("asc_onZoomChanged", this.getZoom());
  };

  WorkbookView.prototype.getEnableKeyEventsHandler = function(bIsNaturalFocus) {
    var res = this.enableKeyEvents;
    if (res && bIsNaturalFocus && this.getCellEditMode() && this.input.isFocused) {
      res = false;
    }
    return res;
  };
  WorkbookView.prototype.enableKeyEventsHandler = function(f) {
    this.enableKeyEvents = !!f;
    this.controller.enableKeyEventsHandler(this.enableKeyEvents);
    if (this.cellEditor) {
      this.cellEditor.enableKeyEventsHandler(this.enableKeyEvents);
    }
  };

	// Останавливаем ввод данных в редакторе ввода
	WorkbookView.prototype.closeCellEditor = function (cancel) {
		return this.getCellEditMode() ? this.cellEditor.close(!cancel) : true;
	};

	WorkbookView.prototype.restoreFocus = function () {
		if (window["NATIVE_EDITOR_ENJINE"]) {
			return;
		}

		if (this.cellEditor.hasFocus) {
			this.cellEditor.restoreFocus();
		}
	};

	WorkbookView.prototype._onUpdateCellEditor = function (text, cursorPosition, fPos, fName) {
		if (this.skipHelpSelector) {
			return;
		}

		var _getInnerTable = function (_sFullTable, tableName) {
			var _bracketCount = 0;
			var res = null;
			for (var j = tableName.length; j < _sFullTable.length; j++) {
				if (_sFullTable[j] === "[") {
					_bracketCount++;
					if (!res) {
						res = "";
					}
				} else if (_sFullTable[j] === "]") {
					_bracketCount--;
					res = null;
					if (_bracketCount <= 0) {
						break;
					}
				} else {
					if (_bracketCount > 0) {
						if (!res) {
							res = "";
						}
						res += _sFullTable[j];
					} else {
						res = null;
					}
				}
			}
			return res;
		};

		// ToDo для ускорения можно завести объект, куда класть результаты поиска по формулам и второй раз не искать.
		var i, arrResult = [], defNamesList, defName, defNameStr, _lastFNameLength, _type;
		if (fName) {
			fName = fName.toUpperCase();
			for (i = 0; i < this.formulasList.length; ++i) {
				if (0 === this.formulasList[i].indexOf(fName)) {
					arrResult.push(new AscCommonExcel.asc_CCompleteMenu(this.formulasList[i], c_oAscPopUpSelectorType.Func));
				}
			}
			defNamesList = this.getDefinedNames(Asc.c_oAscGetDefinedNamesList.WorksheetWorkbook);
			fName = fName.toLowerCase();
			for (i = 0; i < defNamesList.length; ++i) {
				defName = defNamesList[i];
				defNameStr = defName.Name;
				if (null !== defName.LocalSheetId && defNameStr.toLowerCase() === "print_area") {
					defNameStr = AscCommon.translateManager.getValue("Print_Area");
				}

				if (0 === defNameStr.toLowerCase().indexOf(fName)) {
					_type = c_oAscPopUpSelectorType.Range;
					if (defName.type === Asc.c_oAscDefNameType.slicer) {
						_type = c_oAscPopUpSelectorType.Slicer;
					} else if (defName.type === Asc.c_oAscDefNameType.table) {
						_type = c_oAscPopUpSelectorType.Table;
					}
					arrResult.push(new AscCommonExcel.asc_CCompleteMenu(defNameStr, _type));
				} else if (defName.type === Asc.c_oAscDefNameType.table && 0 === fName.indexOf(defNameStr.toLowerCase())) {
					if (-1 !== fName.indexOf("[")) {
						var tableNameParse = fName.split("[");
						if (tableNameParse[0] && 0 === defNameStr.toLowerCase().indexOf(tableNameParse[0])) {
							//ищем совпадения по названию столбцов
							var table = this.model.getTableByName(defNameStr);
							if (table) {
								var sTableInner = _getInnerTable(fName, defNameStr.toLowerCase());
								if (null !== sTableInner) {
									var _str, j;
									for (j = 0; j < table.TableColumns.length; j++) {
										_str = table.TableColumns[j].Name;
										_type = c_oAscPopUpSelectorType.TableColumnName;
										if (sTableInner === "" || 0 === _str.toLowerCase().indexOf(sTableInner)) {
											arrResult.push(new AscCommonExcel.asc_CCompleteMenu(_str, _type));
										}
									}

									if (AscCommon.cStrucTableLocalColumns) {
										for (j in AscCommon.cStrucTableLocalColumns) {
											_str = AscCommon.cStrucTableLocalColumns[j];
											if (j === "h") {
												_str = "#" + _str;
												_type = c_oAscPopUpSelectorType.TableHeaders;
											} else if (j === "d") {
												_str = "#" + _str;
												_type = c_oAscPopUpSelectorType.TableData;
											} else if (j === "a") {
												_str = "#" + _str;
												_type = c_oAscPopUpSelectorType.TableAll;
											} else if (j === "tr") {
												_str = "@" + " - " + _str;
												_type = c_oAscPopUpSelectorType.TableThisRow;
											} else if (j === "t") {
												_str = "#" + _str;
												_type = c_oAscPopUpSelectorType.TableTotals;
											}
											if (sTableInner === "" || (0 === _str.toLocaleLowerCase().indexOf(sTableInner) && _type !== c_oAscPopUpSelectorType.TableThisRow)) {
												arrResult.push(new AscCommonExcel.asc_CCompleteMenu(_str, _type));
											}
										}
									}
									fPos += defNameStr.length + (fName.length - defNameStr.length - sTableInner.length);
									_lastFNameLength = sTableInner.length;
								}
							}
						}
					}
				}
			}
		}
		if (0 < arrResult.length) {
			this.handlers.trigger('asc_onFormulaCompleteMenu', arrResult, this.cellEditor.calculateOffset(fPos));

			this.lastFPos = fPos;
			this.lastFNameLength = _lastFNameLength !== undefined ? _lastFNameLength : fName.length;
		} else {
			this.handlers.trigger('asc_onFormulaCompleteMenu', null);

			this.lastFPos = -1;
			this.lastFNameLength = 0;
		}
	};

    // Вставка формулы в редактор
    WorkbookView.prototype.insertInCellEditor = function (name, type, autoComplete) {
        var t = this, ws = this.getWorksheet(), cursorPos, tmp;

        var isNotFunction = c_oAscPopUpSelectorType.Func !== type;

        // Проверяем, открыт ли редактор
        if (this.getCellEditMode()) {
            if (isNotFunction) {
                this.skipHelpSelector = true;
            }

            if (-1 !== this.lastFPos) {
                if (-1 === this.arrExcludeFormulas.indexOf(name) && !isNotFunction) {
                    //если следующий символ скобка - не добавляем ещё одну
                    if('(' !== this.cellEditor.textRender.getChars(this.cellEditor.cursorPos, 1)) {
                        name += '('; // ToDo сделать проверки при добавлении, чтобы не вызывать постоянно окно
                    }
                } else {
                    this.skipHelpSelector = true;
                }
                tmp = this.cellEditor.skipTLUpdate;
                this.cellEditor.skipTLUpdate = false;
                this.cellEditor.replaceText(this.lastFPos, this.lastFNameLength, type === c_oAscPopUpSelectorType.TableThisRow ? "@" : name);
                this.cellEditor.skipTLUpdate = tmp;
            } else if (false === this.cellEditor.insertFormula(name, isNotFunction)) {
                // Не смогли вставить формулу, закроем редактор, с сохранением текста
                this.cellEditor.close(true);
            }

            this.skipHelpSelector = false;
        } else {
            if (c_oAscPopUpSelectorType.None === type) {
				ws.executeWithFirstActiveCellInMerge(function () {
					ws.setSelectionInfo("value", name, /*onlyActive*/true);
				})
                return;
            } else if (c_oAscPopUpSelectorType.TotalRowFunc === type) {
                ws.setSelectionInfo("totalRowFunc", name, /*onlyActive*/true);
                return;
            }

            var callback = function (success) {
                // ToDo ?
                if (isNotFunction) {
                    t.skipHelpSelector = false;
                }
            };

            // Редактор закрыт
            var cellRange = {};
            // Если нужно сделать автозаполнение формулы, то ищем ячейки)
            if (autoComplete) {
                cellRange = ws.autoCompleteFormula(name);
            }
            if (isNotFunction) {
                name = "=" + name;
            } else {
                if (cellRange.notEditCell) {
                    // Мы уже ввели все что нужно, редактор открывать не нужно
                    return;
                }
                if (cellRange.text) {
                    // Меняем значение ячейки
                    name = ws.generateAutoCompleteFormula(name, cellRange.text);
                } else {
                    // Меняем значение ячейки
                    name = ws.generateAutoCompleteFormula(name, "")
                }
                // Вычисляем позицию курсора (он должен быть в функции)
                cursorPos = name.length - 1;
            }

            if (isNotFunction) {
                t.skipHelpSelector = true;
            }

            var enterOptions = new AscCommonExcel.CEditorEnterOptions();
            // Открываем, с выставлением позиции курсора
            enterOptions.newText = name;
            enterOptions.cursorPos = cursorPos;

            this._onEditCell(enterOptions, callback);
        }
    };

    WorkbookView.prototype.startWizard = function (name, doCleanCellContent) {
        var t = this;
        var callback = function (success) {
            if (success) {
                addFunction(name);
            }
        };

        var addFunction = function (name) {
        	t.setWizardMode(true);
			if (doCleanCellContent || !t.cellEditor.isFormula()) {
                t.cellEditor.selectionBegin = 0;
                t.cellEditor.selectionEnd = t.cellEditor.textRender.getEndOfText();
            }
			t.cellEditor.insertFormula(name);
			// ToDo send info from selection
			var res = name ? new AscCommonExcel.CFunctionInfo(name) : null;
			t.handlers.trigger("asc_onSendFunctionWizardInfo", res);
        };

        if (!this.getCellEditMode()) {
			if (!name) {
				this.cellEditor.needFindFirstFunction = true;
			}
            this._onEditCell(new AscCommonExcel.CEditorEnterOptions(), callback);
            return;
        }

        addFunction(name);
    };

    WorkbookView.prototype.canEnterWizardRange = function (char) {
        return this.getCellEditMode() && this.cellEditor.checkSymbolBeforeRange(char);
    };

	WorkbookView.prototype.insertArgumentsInFormula = function (args, argNum, argType, name) {
		if (this.getCellEditMode()) {
			var sArguments = args.join(AscCommon.FormulaSeparators.functionArgumentSeparator);
			this.cellEditor.changeCellText(sArguments);

			if (name) {
				var ws = this.getActiveWS();

				var res = new AscCommonExcel.CFunctionInfo(name);
				res.argumentsResult = [];
				var argCalc = ws.calculateWizardFormula(args[argNum], argType);
				res.argumentsResult[argNum] = argCalc.str;
				if (argCalc.obj && argCalc.obj.type !== AscCommonExcel.cElementType.error) {
					var funcCalc = ws.calculateWizardFormula(name + '(' + sArguments + ')');
					res.functionResult = funcCalc.str;
					if (funcCalc.obj && funcCalc.obj.type !== AscCommonExcel.cElementType.error) {
						res.formulaResult = ws.calculateWizardFormula(this.cellEditor.getText().substring(1)).str;
					}
				}

				return res;
			}
		}

		return null;
	};

  WorkbookView.prototype.bIsEmptyClipboard = function() {
    return g_clipboardExcel.bIsEmptyClipboard(this.getCellEditMode());
  };

   WorkbookView.prototype.checkCopyToClipboard = function(_clipboard, _formats) {
    if(this.isProtectedFromCut()) {
      return false;
    }
    var ws = this.getWorksheet();
    g_clipboardExcel.checkCopyToClipboard(ws, _clipboard, _formats);
  };

  WorkbookView.prototype.pasteData = function(_format, data1, data2, text_data, doNotShowButton) {
    var t = this, ws;
    ws = t.getWorksheet();
    g_clipboardExcel.pasteData(ws, _format, data1, data2, text_data, null, doNotShowButton);
  };

  WorkbookView.prototype.specialPasteData = function(props) {
    if (!this.getCellEditMode()) {
		this.getWorksheet().specialPaste(props);
	}
  };

  WorkbookView.prototype.showSpecialPasteButton = function(props) {
	if (!this.getCellEditMode()) {
		this.getWorksheet().showSpecialPasteOptions(props);
	}
  };

  WorkbookView.prototype.updateSpecialPasteButton = function(props) {
  	if (!this.getCellEditMode()) {
  		this.getWorksheet().updateSpecialPasteButton(props);
	}
  };

	WorkbookView.prototype.hideSpecialPasteButton = function() {
		//TODO пересмотреть!
		//сейчас сначала добавляются данные в историю, потом идет закрытие редактора ячейки
		//следовательно убираю проверку на редактирование ячейки
		/*if (!this.getCellEditMode()) {*/
			this.handlers.trigger("hideSpecialPasteOptions");
		//}
	};
  WorkbookView.prototype.isProtectedFromCut = function() {
    var ws = this.getWorksheet();
    if (ws && ws.model && ws.model.getSheetProtection() && AscCommon.g_clipboardBase.bCut) {
      if(ws.isSelectOnShape) {
        if(ws.objectRender && ws.objectRender.controller) {
          var oController = ws.objectRender.controller;
          if(oController.isProtectedFromCut()) {
            return true;
          }
        }
      } else {
        var selection = ws._getSelection();
        var selectionRanges = selection ? selection.ranges : null;
        if (selectionRanges && ws.model.isIntersectLockedRanges(selectionRanges)) {
          return true;
        }
      }
    }
    return false;
  };
	WorkbookView.prototype.selectionCut = function () {
		if (this.getCellEditMode()) {
			this.cellEditor.cutSelection();
		} else {
			var ws = this.getWorksheet();
      if(this.isProtectedFromCut()) {
        this.handlers.trigger("asc_onError", c_oAscError.ID.ChangeOnProtectedSheet, c_oAscError.Level.NoCritical);
        return false;
      }

			if (ws.isNeedSelectionCut()) {
				ws.emptySelection(c_oAscCleanOptions.All, true);
			} else {
				//в данном случае не вырезаем, а записываем
				if(false === ws.isMultiSelect()) {
					this.cutIdSheet = ws.model.Id;
					ws.cutRange = ws.model.selectionRange.getLast();
				}
			}
		}
	};

  WorkbookView.prototype.undo = function(Options) {
    var oFormulaLocaleInfo = AscCommonExcel.oFormulaLocaleInfo;
    oFormulaLocaleInfo.Parse = false;
    oFormulaLocaleInfo.DigitSep = false;
    if (!this.getCellEditMode()) {
      if (!History.Undo(Options) && this.collaborativeEditing.getFast() && this.collaborativeEditing.getCollaborativeEditing()) {
        this.Api.sync_TryUndoInFastCollaborative();
      }
    } else {
      this.cellEditor.undo();
    }
    oFormulaLocaleInfo.Parse = true;
    oFormulaLocaleInfo.DigitSep = true;
  };

  WorkbookView.prototype.redo = function() {
    if (!this.getCellEditMode()) {
      History.Redo();
    } else {
      this.cellEditor.redo();
    }
  };

  WorkbookView.prototype.setFontAttributes = function(prop, val) {
    if (!this.getCellEditMode()) {
      this.getWorksheet().setSelectionInfo(prop, val);
    } else {
      this.cellEditor.setTextStyle(prop, val);
    }
  };

  WorkbookView.prototype.changeFontSize = function(prop, val) {
    if (!this.getCellEditMode()) {
      this.getWorksheet().setSelectionInfo(prop, val);
    } else {
      this.cellEditor.setTextStyle(prop, val);
    }
  };

	WorkbookView.prototype.setCellFormat = function (format) {
		this.getWorksheet().setSelectionInfo("format", format);
	};

	WorkbookView.prototype.emptyCells = function (options, isMineComments) {
		if (!this.getCellEditMode()) {
			if (Asc.c_oAscCleanOptions.Comments === options) {
				this.removeAllComments(false, true);
			} else {
				//TODO isMineComments - временный флаг, как только в сдк появится класс для групп, добавить этот флаг туда
				this.getWorksheet().emptySelection(options, null, isMineComments);
			}
			this.restoreFocus();
		} else {
			this.cellEditor.empty(options);
		}
	};

  WorkbookView.prototype._setSelectionDialogType = function (selectionDialogType) {
      this.dialogSheetName = (c_oAscSelectionDialogType.Chart === selectionDialogType ||
          c_oAscSelectionDialogType.PivotTableData === selectionDialogType ||
          c_oAscSelectionDialogType.PivotTableReport === selectionDialogType);
      this.dialogAbsName = (c_oAscSelectionDialogType.None !== selectionDialogType &&
          c_oAscSelectionDialogType.Function !== selectionDialogType);
  };
  WorkbookView.prototype.setSelectionDialogMode = function (selectionDialogType, selectRange) {
      var newSelectionDialogMode = c_oAscSelectionDialogType.None !== selectionDialogType;

      if (newSelectionDialogMode === this.selectionDialogMode) {
          return;
      }

      this._setSelectionDialogType(selectionDialogType);
      var drawSelection = false;
      var index;
      if (newSelectionDialogMode) {
          this.copyActiveSheet = this.wsActive;
          if(c_oAscSelectionDialogType.Chart === selectionDialogType) {
              var sRef = selectRange;
              if(sRef.charAt(0) === '=') {
                  sRef = sRef.slice(1);
              }
              var aRefs = sRef.split(',');
              var aSelectRanges = [];
              var oRange, oAscRange;
              var sWSName = null;
              for(var nRef = 0; nRef < aRefs.length; ++nRef) {
                  oRange = AscCommon.parserHelp.parse3DRef(aRefs[nRef]);
                  if(!oRange) {
                      aSelectRanges.length = 0;
                      break;
                  }
                  if(sWSName === null) {
                      sWSName = oRange.sheet;
                  }
                  else {
                      if(sWSName !== oRange.sheet) {
                          aSelectRanges.length = 0;
                          break;
                      }
                  }
                  oAscRange = AscCommonExcel.g_oRangeCache.getAscRange(oRange.range);
                  if(oAscRange) {
                      aSelectRanges.push(oAscRange);
                  }
                  else {
                      aSelectRanges.length = 0;
                      break;
                  }
              }
              if(aSelectRanges.length > 0) {
                  var oWS = this.model.getWorksheetByName(sWSName);
                  if (!oWS || oWS.getHidden()) {
                      this.getWorksheet().cloneSelection(true, null);
                  } else {
                      index = oWS.getIndex();
                      if (index !== this.wsActive) {
                          this.showWorksheet(index);
                      }
                      this.getWorksheet().cloneSelection(true, new Asc.Range(0, 0, 0, 0));
                      oWS.selectionRange.assign2(aSelectRanges[0]);
                      for(nRef = 1; nRef < aSelectRanges.length; ++nRef) {
                          oWS.selectionRange.addRange();
                          oWS.selectionRange.getLast().assign2(aSelectRanges[nRef]);
                      }
                      oWS.selectionRange.update();
                  }
              }
              else {
                  this.getWorksheet().cloneSelection(true, null);
              }
          }
          else {
              var tmpSelectRange = AscCommon.parserHelp.parse3DRef(selectRange);
              if (tmpSelectRange) {
                  var ws = this.model.getWorksheetByName(tmpSelectRange.sheet);
                  if (!ws || ws.getHidden()) {
                      tmpSelectRange = null;
                  } else {
                      index = ws.getIndex();
                      if (index !== this.wsActive) {
                          this.showWorksheet(index);
                      }

                      tmpSelectRange = tmpSelectRange.range;
                  }
              } else {
                  tmpSelectRange = selectRange;
              }
              this.getWorksheet().cloneSelection(true, tmpSelectRange && AscCommonExcel.g_oRangeCache.getAscRange(tmpSelectRange));
          }
          this.selectionDialogMode = newSelectionDialogMode;
          this.input.disabled = !this.isFormulaEditMode || this.isWizardMode;
          drawSelection = true;
      } else {
          this.selectionDialogMode = newSelectionDialogMode;
          this.getWorksheet().cloneSelection(false);
          if (this.copyActiveSheet !== this.wsActive) {
              this.showWorksheet(this.copyActiveSheet);
          } else {
              drawSelection = true;
          }

          this.copyActiveSheet = -1;
          this.input.disabled = false;
      }

      if (drawSelection) {
          this.getWorksheet()._drawSelection();
      }
  };

  WorkbookView.prototype.formatPainter = function(formatPainterState, bLockDraw) {
    var ws = this.getWorksheet();
    if (!bLockDraw) {
        ws.cleanSelection();
    }

    // Если передали состояние, то выставляем его. Если нет - то меняем на противоположное.
    this.formatPainterState = (null != formatPainterState) ? formatPainterState :
        ((c_oAscFormatPainterState.kOff !== this.formatPainterState) ? c_oAscFormatPainterState.kOff : c_oAscFormatPainterState.kOn);

    if (this.formatPainterState) {
      this.formatPainterSheet = this.wsActive;
      this.formatPainterRange = ws.model.selectionRange.clone();
    } else {
      this.formatPainterSheet = this.formatPainterRange = null;
      this.handlers.trigger('asc_onStopFormatPainter');
    }
    if (!bLockDraw) {
      ws.updateSelectionWithSparklines();
    }
  };

  // Поиск текста в листе
  WorkbookView.prototype.findCellText = function (options) {
    this.closeCellEditor();
    // Для поиска эта переменная не нужна (но она может остаться от replace)
    options.selectionRange = null;

    var result = this.model.findCellText(options);
    if (result) {
      var ws = this.getWorksheet();
      var range = new Asc.Range(result.col, result.row, result.col, result.row);
      options.findInSelection ? ws.setActiveCell(result) : ws.setSelection(range);
      return true;
    }
    return null;
  };

  // Замена текста в листе
  WorkbookView.prototype.replaceCellText = function (options) {
  	this.closeCellEditor();
  	if (!options.isMatchCase) {
  		options.findWhat = options.findWhat.toLowerCase();
  	}
  	options.findRegExp = AscCommonExcel.getFindRegExp(options.findWhat, options, true);

    History.Create_NewPoint();
    History.StartTransaction();

    options.clearFindAll();
    if (options.isReplaceAll) {
      // На ReplaceAll ставим медленную операцию
      this.Api.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.SlowOperation);
    }

    var ws = this.getWorksheet();
    ws.replaceCellText(options, false, this.fReplaceCallback);
  };
  WorkbookView.prototype._replaceCellTextCallback = function(options) {
    if (!options.error) {
		options.updateFindAll();
		if (!options.scanOnOnlySheet && options.isReplaceAll) {
			// Замена на всей книге
			var i = ++options.sheetIndex;
			if (this.model.getActive() === i) {
				i = ++options.sheetIndex;
			}

			if (i < this.model.getWorksheetCount()) {
				var ws = this.getWorksheet(i);
				ws.replaceCellText(options, true, this.fReplaceCallback);
				return;
			}
      options.sheetIndex = -1;
		}

		this.handlers.trigger("asc_onRenameCellTextEnd", options.countFindAll, options.countReplaceAll);
    }

    History.EndTransaction();
    if (options.isReplaceAll) {
      // Заканчиваем медленную операцию
      this.Api.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.SlowOperation);
    }
  };

  WorkbookView.prototype.getDefinedNames = function(defNameListId, excludeErrorRefNames) {
    return this.model.getDefinedNamesWB(defNameListId, true, excludeErrorRefNames);
  };

  WorkbookView.prototype.setDefinedNames = function(defName) {
    //ToDo проверка defName.ref на знак "=" в начале ссылки. знака нет тогда это либо число либо строка, так делает Excel.

    this.model.setDefinesNames(defName.Name, defName.Ref, defName.Scope);
    this.handlers.trigger("asc_onDefName");

  };

  WorkbookView.prototype.editDefinedNames = function(oldName, newName) {
    //ToDo проверка defName.ref на знак "=" в начале ссылки. знака нет тогда это либо число либо строка, так делает Excel.
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    var ws = this.getWorksheet(), t = this;

    var editDefinedNamesCallback = function(res) {
      if (res) {
        if (oldName && oldName.asc_getType() === Asc.c_oAscDefNameType.table) {
          ws.model.autoFilters.changeDisplayNameTable(oldName.asc_getName(), newName.asc_getName());
        } else {
          t.model.editDefinesNames(oldName, newName);
        }
        if (oldName && newName) {
			ws.model.changeSlicerCacheName(oldName.asc_getName(), newName.asc_getName());
		}
        t.handlers.trigger("asc_onEditDefName", oldName, newName);
        //условие исключает второй вызов asc_onRefreshDefNameList(первый в unlockDefName)
        if(!(t.collaborativeEditing.getCollaborativeEditing() && t.collaborativeEditing.getFast()))
        {
          t.handlers.trigger("asc_onRefreshDefNameList");
        }
      } else {
        t.handlers.trigger("asc_onError", c_oAscError.ID.LockCreateDefName, c_oAscError.Level.NoCritical);
      }
      t._onSelectionNameChanged(ws.getSelectionName(/*bRangeText*/false));
      if(ws.viewPrintLines) {
		  ws.updateSelection();
	  }
    };
    var defNameId;
    if (oldName) {
      defNameId = t.model.getDefinedName(oldName);
      defNameId = defNameId ? defNameId.getNodeId() : null;
    }

    var callback = function() {
      ws._isLockedDefNames(editDefinedNamesCallback, defNameId);
    };

    var tableRange;
    if(oldName && oldName.type === Asc.c_oAscDefNameType.table)
    {
      var table = ws.model.autoFilters._getFilterByDisplayName(oldName.Name);
      if(table)
      {
        tableRange = table.Ref;
      }
    } else if (oldName) {
      var slicerCache = ws.model.getSlicerCacheByCacheName(oldName.Name);
      if (slicerCache) {
        tableRange = slicerCache.getRange();
      }
    }

    if (tableRange) {
      ws._isLockedCells( tableRange, null, callback );
    } else {
      callback();
    }
  };

  WorkbookView.prototype.delDefinedNames = function(oldName) {
    //ToDo проверка defName.ref на знак "=" в начале ссылки. знака нет тогда это либо число либо строка, так делает Excel.
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    var ws = this.getWorksheet(), t = this;

    if (oldName) {

      var delDefinedNamesCallback = function(res) {
        if (res) {
          t.model.delDefinesNames(oldName);
          t.handlers.trigger("asc_onRefreshDefNameList");
        } else {
          t.handlers.trigger("asc_onError", c_oAscError.ID.LockCreateDefName, c_oAscError.Level.NoCritical);
        }
        t._onSelectionNameChanged(ws.getSelectionName(/*bRangeText*/false));
		if(ws.viewPrintLines) {
		  ws.updateSelection();
		}
      };
      var defNameId = t.model.getDefinedName(oldName).getNodeId();

      ws._isLockedDefNames(delDefinedNamesCallback, defNameId);

    }
  };

  WorkbookView.prototype.getDefaultDefinedName = function() {
    //ToDo проверка defName.ref на знак "=" в начале ссылки. знака нет тогда это либо число либо строка, так делает Excel.
    var val = this.getWorksheet().getDefaultDefinedNameText();
    return new Asc.asc_CDefName(val, this.getWorksheet().getSelectionRangeValue(true, true), null);

  };
  WorkbookView.prototype.getDefaultTableStyle = function() {
  	return this.model.TableStyles.DefaultTableStyle;
  };
  WorkbookView.prototype.unlockDefName = function() {
    this.model.unlockDefName();
    this.handlers.trigger("asc_onRefreshDefNameList");
    this.handlers.trigger("asc_onLockDefNameManager", Asc.c_oAscDefinedNameReason.OK);
  };
  WorkbookView.prototype.unlockCurrentDefName = function(name, sheetId) {
    this.model.unlockCurrentDefName(name, sheetId);
    //this.handlers.trigger("asc_onRefreshDefNameList");
    //this.handlers.trigger("asc_onLockDefNameManager", Asc.c_oAscDefinedNameReason.OK);
  };

  WorkbookView.prototype._onCheckDefNameLock = function() {
    return this.model.checkDefNameLock();
  };

  // Печать
  WorkbookView.prototype.printSheets = function(printPagesData, pdfDocRenderer) {
	  var pdfPrinter;
	  var t = this;
	  this._executeWithoutZoom(function () {
		  pdfPrinter = new AscCommonExcel.CPdfPrinter(t.fmgrGraphics[3], t.m_oFont);
		  if (pdfDocRenderer) {
			  pdfPrinter.DocumentRenderer = pdfDocRenderer;
		  }
		  var ws;
		  if (0 === printPagesData.arrPages.length) {
			  // Печать пустой страницы
			  ws = t.getWorksheet();
			  ws.drawForPrint(pdfPrinter, null);
		  } else {
			  var indexWorksheet = -1;
			  var indexWorksheetTmp = -1;
			  for (var i = 0; i < printPagesData.arrPages.length; ++i) {
				  indexWorksheetTmp = printPagesData.arrPages[i].indexWorksheet;
				  if (indexWorksheetTmp !== indexWorksheet) {
					  ws = t.getWorksheet(indexWorksheetTmp);
					  indexWorksheet = indexWorksheetTmp;
				  }
				  ws.drawForPrint(pdfPrinter, printPagesData.arrPages[i], i, printPagesData.arrPages.length);
			  }
		  }
	  });

	  return pdfPrinter;
  };

	WorkbookView.prototype._executeWithoutZoom = function (runFunction) {
		//TODO есть проблемы при отрисовке(не связано с печатью). сначала меняем zoom редактора,
		// потом системный(открываем при системном зуме != 100%)

		//change zoom on default
		var trueRetinaPixelRatio = AscCommon.AscBrowser.retinaPixelRatio;
		var viewZoom = this.getZoom();
		if (viewZoom === 1 && trueRetinaPixelRatio === 1) {
			runFunction();
			return;
		}

		AscCommon.AscBrowser.retinaPixelRatio = 1;
		//приходится несколько раз выполнять действия, чтобы ppi выставился правильно
		//если не делать init, то не сбросится ppi от системного зума - смотри функцию DrawingContext.prototype.changeZoom
		if (viewZoom !== 1) {
			this.changeZoom(1, true, true);
		}
		this.changeZoom(null, true, true);

		runFunction();

		AscCommon.AscBrowser.retinaPixelRatio = trueRetinaPixelRatio;
		this.changeZoom(null, true, true);
		this.changeZoom(viewZoom, true, true);
	};

	WorkbookView.prototype.printSheetPrintPreview = function(index) {
		var printPreviewState = this.printPreviewState;
		var page = printPreviewState.getPage(index);
		var printPreviewContext = printPreviewState.getCtx();

		if (page) {
			page = page.clone();
		}

		var kF = printPreviewState.pageZoom * AscCommon.AscBrowser.retinaPixelRatio;
		//TODO 1 -
		if (page) {
			page.leftFieldInPx = Math.floor(page.leftFieldInPx * kF)  - 1;
			page.pageClipRectHeight = Math.ceil(page.pageClipRectHeight * kF);
			page.pageClipRectLeft = Math.floor(page.pageClipRectLeft * kF);
			page.pageClipRectTop = Math.floor(page.pageClipRectTop * kF);
			page.pageClipRectWidth = Math.ceil(page.pageClipRectWidth * kF);
			page.topFieldInPx = Math.floor(page.topFieldInPx * kF) - 1;

			page.titleWidth = Math.floor(page.titleWidth * kF);
			page.titleHeight = Math.floor(page.titleHeight * kF);
		}

		printPreviewContext.clear();
		printPreviewContext.setFillStyle( this.defaults.worksheetView.cells.defaultState.background )
			.fillRect( 0, 0, printPreviewContext.getWidth(), printPreviewContext.getHeight() );

		var ws;
		if (!page) {
			// Печать пустой страницы
			ws = this.getWorksheet();
			ws.drawForPrint(printPreviewContext, null);
		} else {
			ws = this.getWorksheet(page.indexWorksheet);
			ws.drawForPrint(printPreviewContext, page, index, printPreviewState.getPagesLength());
		}
	};

  WorkbookView.prototype._calcPagesPrintSheet = function (index, printPagesData, onlySelection, adjustPrint) {
  	var ws = this.model.getWorksheet(index);
  	var wsView = this.getWorksheet(index);
  	if (!ws.getHidden()) {
		var pageOptionsMap = adjustPrint ? adjustPrint.asc_getPageOptionsMap() : null;
  		var pagePrintOptions = pageOptionsMap && pageOptionsMap[index] ? pageOptionsMap[index] : ws.PagePrintOptions;
  		wsView.calcPagesPrint(pagePrintOptions, onlySelection, index, printPagesData.arrPages, null, adjustPrint);
  	}
  };
  WorkbookView.prototype.calcPagesPrint = function (adjustPrint) {
  	if (!adjustPrint) {
  		adjustPrint = new Asc.asc_CAdjustPrint();
	}

    var viewZoom = this.getZoom();
    this.changeZoom(1, true, true);

	var isPrintPreview = this.printPreviewState.isStart();
	var nActive = isPrintPreview && null !== this.printPreviewState.realActiveSheet ? this.printPreviewState.realActiveSheet : this.model.getActive();

    var printPagesData = new asc_CPrintPagesData();
    var printType = adjustPrint.asc_getPrintType();
    if (printType === Asc.c_oAscPrintType.ActiveSheets) {
      this._calcPagesPrintSheet(nActive, printPagesData, false, adjustPrint);
    } else if (printType === Asc.c_oAscPrintType.EntireWorkbook) {
      // Колличество листов
      var countWorksheets = this.model.getWorksheetCount();
      for (var i = 0; i < countWorksheets; ++i) {
      	if(adjustPrint.isOnlyFirstPage && i !== 0) {
      		break;
		}
      	this._calcPagesPrintSheet(i, printPagesData, false, adjustPrint);
      }
    } else if (printType === Asc.c_oAscPrintType.Selection) {
      this._calcPagesPrintSheet(nActive, printPagesData, true, adjustPrint);
    }

    if (this.printPreviewState.isNeedShowError(AscCommonExcel.c_kMaxPrintPages === printPagesData.arrPages.length)) {
		this.handlers.trigger("asc_onError", c_oAscError.ID.PrintMaxPagesCount, c_oAscError.Level.NoCritical);
    }

    this.changeZoom(viewZoom, true, true);

    return printPagesData;
  };

  // Вызывать только для нативной печати
  WorkbookView.prototype._nativeCalculate = function() {
    var item;
    for (var i in this.wsViews) {
      item = this.wsViews[i];
      item._cleanCellsTextMetricsCache();
      item._prepareDrawingObjects();
    }
  };

  WorkbookView.prototype.calculate = function (type) {
  	this.model.calculate(type);
  	this.drawWS();
  };

  WorkbookView.prototype.drawWS = function() {
    this.getWorksheet().draw();
  };
  WorkbookView.prototype.scrollToTopLeftCell = function() {
    this.getWorksheet().scrollToTopLeftCell();
  };
  WorkbookView.prototype.onShowDrawingObjects = function() {
      var oWSView = this.getWorksheet();
      var oDrawingsRender;
      if(oWSView) {
          oDrawingsRender = oWSView.objectRender;
          if(oDrawingsRender) {
              oDrawingsRender.showDrawingObjects(null);
          }
      }
  };
    WorkbookView.prototype.handleChartsOnWorkbookChange = function (aRanges) {
        if(!Array.isArray(aRanges) || aRanges.length === 0) {
            return;
        }
        var aRefsToChange = [];
        var aCharts = [];
        this.model.handleDrawings(function(oDrawing) {
            if(oDrawing.getObjectType() === AscDFH.historyitem_type_ChartSpace) {
                var nPrevLength = aRefsToChange.length;
                oDrawing.collectIntersectionRefs(aRanges, aRefsToChange);
                if(aRefsToChange.length > nPrevLength) {
                    aCharts.push(oDrawing);
                }
            }
        });
        if(aRefsToChange.length > 0) {
            for(var nRef = 0; nRef < aRefsToChange.length; ++nRef) {
                aRefsToChange[nRef].updateCacheAndCat();
            }
            for(var nChart = 0; nChart < aCharts.length; ++nChart) {
                aCharts[nChart].recalculate();
            }
            this.onShowDrawingObjects();
        }
    };
    WorkbookView.prototype.handleChartsOnChangeSheetName = function (oWorksheet, sOldName, sNewName) {
        //change sheet name in chart references
        var oWorkbook = this.model;
        var oRenameData = oWorkbook.getChartSheetRenameData(oWorksheet, sOldName);
        var oThis = this;
        if(oRenameData.refs.length > 0) {
            oWorkbook.checkObjectsLock(oRenameData.ids, function(bNoLock) {
                if(bNoLock) {
                    oWorkbook.changeSheetNameInRefs(oRenameData.refs, sOldName, sNewName);
                }
                //recalculate in any case. Some charts might depend on new chart name
                oThis.recalculateDrawingObjects(null, false);
            });
        }
        else {
            //recalculate in any case. Some charts might depend on new chart name
            oThis.recalculateDrawingObjects(null, false);
        }
    };
    WorkbookView.prototype.recalculateDrawingObjects = function(oHistoryPoint, bAll) {
        var aWSVies = this.wsViews;
        var oWSView, oDrawingsRender;
        History.Get_RecalcData(oHistoryPoint);
        for (var i = 0; i < aWSVies.length; ++i) {
            oWSView = aWSVies[i];
            if(oWSView) {
                oDrawingsRender = oWSView.objectRender;
                if(oDrawingsRender) {
                    oDrawingsRender.recalculate(bAll);
                }
            }
        }
        this.onShowDrawingObjects();
    };

  WorkbookView.prototype.insertHyperlink = function(options) {
    var ws = this.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists()) {
      if (ws.objectRender.controller.canAddHyperlink()) {
        ws.objectRender.controller.insertHyperlink(options);
      }
    } else {
      ws.setSelectionInfo("hyperlink", options);
      this.restoreFocus();
    }
  };
  WorkbookView.prototype.removeHyperlink = function() {
    var ws = this.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists()) {
      ws.objectRender.controller.removeHyperlink();
    } else {
      ws.setSelectionInfo("rh");
    }
  };

  WorkbookView.prototype.setDocumentPlaceChangedEnabled = function(val) {
    this.isDocumentPlaceChangedEnabled = val;
  };

	WorkbookView.prototype.showComments = function (val, isShowSolved) {
		if (this.isShowComments !== val || this.isShowSolved !== isShowSolved) {
			this.isShowComments = val;
			this.isShowSolved = isShowSolved;
			this.drawWS();
		}
	};
	WorkbookView.prototype.removeComment = function (id) {
		var ws = this.getWorksheet();
		ws.cellCommentator.removeComment(id);
		this.cellCommentator.removeComment(id);
	};
	WorkbookView.prototype.removeAllComments = function (isMine, isCurrent) {
		var range;
		var ws = this.getWorksheet();
		isMine = isMine ? (this.Api.DocInfo && this.Api.DocInfo.get_UserId()) : null;
		History.Create_NewPoint();
		History.StartTransaction();
		if (isCurrent) {
			ws._getSelection().ranges.forEach(function (item) {
				ws.cellCommentator.deleteCommentsRange(item, isMine);
			});
		} else {
			range = new Asc.Range(0, 0, AscCommon.gc_nMaxCol0, AscCommon.gc_nMaxRow0);
			this.cellCommentator.deleteCommentsRange(range, isMine);
			ws.cellCommentator.deleteCommentsRange(range, isMine);
		}
		History.EndTransaction();
	};

	WorkbookView.prototype.resolveAllComments = function (isMine, isCurrent) {
		var range;
		var ws = this.getWorksheet();
		isMine = isMine ? (this.Api.DocInfo && this.Api.DocInfo.get_UserId()) : null;
		History.Create_NewPoint();
		History.StartTransaction();
		if (isCurrent) {
			ws._getSelection().ranges.forEach(function (item) {
				ws.cellCommentator.resolveCommentsRange(item, isMine);
			});
		} else {
			range = new Asc.Range(0, 0, AscCommon.gc_nMaxCol0, AscCommon.gc_nMaxRow0);
			this.cellCommentator.resolveCommentsRange(range, isMine);
			ws.cellCommentator.resolveCommentsRange(range, isMine);
		}
		History.EndTransaction();
	};

  /*
   * @param {c_oAscRenderingModeType} mode Режим отрисовки
   * @param {Boolean} isInit инициализация или нет
   */
  WorkbookView.prototype.setFontRenderingMode = function(mode, isInit) {
    if (mode !== this.fontRenderingMode) {
      this.fontRenderingMode = mode;
      if (c_oAscFontRenderingModeType.noHinting === mode) {
        this._setHintsProps(false, false);
      } else if (c_oAscFontRenderingModeType.hinting === mode) {
        this._setHintsProps(true, false);
      } else if (c_oAscFontRenderingModeType.hintingAndSubpixeling === mode) {
        this._setHintsProps(true, true);
      }

      if (!isInit) {
        this.drawWS();
        this.cellEditor.setFontRenderingMode(mode);
      }
    }
  };

  WorkbookView.prototype.initFormulasList = function() {
    this.formulasList = [];
    var oFormulaList = AscCommonExcel.cFormulaFunctionLocalized ? AscCommonExcel.cFormulaFunctionLocalized :
      AscCommonExcel.cFormulaFunction;
    for (var f in oFormulaList) {
      this.formulasList.push(f);
    }
    this.arrExcludeFormulas = [cBoolLocal.t, cBoolLocal.f];
  };

  WorkbookView.prototype._setHintsProps = function(bIsHinting, bIsSubpixHinting) {
    var manager;
    for (var i = 0, length = this.fmgrGraphics.length; i < length; ++i) {
      manager = this.fmgrGraphics[i];
      // Последний без хинтования (только для измерения)
      if (i === length - 1) {
        bIsHinting = bIsSubpixHinting = false;
      }

      manager.SetHintsProps(bIsHinting, bIsSubpixHinting);
    }
  };

  WorkbookView.prototype._calcMaxDigitWidth = function () {
	  var t = this;
	  this.executeDefaultDpi(function () {
		  // set default worksheet header font for calculations
		  t.buffers.main.setFont(AscCommonExcel.g_oDefaultFormat.Font);
		  // Измеряем в pt
		  t.stringRender.measureString("0123456789", new AscCommonExcel.CellFlags());

		  // Переводим в px и приводим к целому (int)
		  t.model.maxDigitWidth = t.maxDigitWidth = t.stringRender.getWidestCharWidth();
		  // Проверка для Calibri 11 должно быть this.maxDigitWidth = 7

		  if (!t.maxDigitWidth) {
			  throw new Error("Error: can't measure text string");
		  }

		  // Padding рассчитывается исходя из maxDigitWidth (http://social.msdn.microsoft.com/Forums/en-US/9a6a9785-66ad-4b6b-bb9f-74429381bd72/margin-padding-in-cell-excel?forum=os_binaryfile)
		  t.defaults.worksheetView.cells.padding = Math.max(asc.ceil(t.maxDigitWidth / 4), 2);
		  t.model.paddingPlusBorder = 2 * t.defaults.worksheetView.cells.padding + 1;
	  })
  };

	WorkbookView.prototype.executeDefaultDpi = function (runFunction) {
		var truePPIX = this.drawingCtx.ppiX;
		var truePPIY = this.drawingCtx.ppiY;
		var retinaPixelRatio = AscBrowser.retinaPixelRatio;

		this.drawingCtx.ppiY = 96;
		this.drawingCtx.ppiX = 96;
		AscBrowser.retinaPixelRatio = 1;

		runFunction();

		this.drawingCtx.ppiY = truePPIY;
		this.drawingCtx.ppiX = truePPIX;
		AscBrowser.retinaPixelRatio = retinaPixelRatio;
	};


	WorkbookView.prototype.getPivotMergeStyle = function (sheetMergedStyles, range, style, pivot) {
		var styleInfo = pivot.asc_getStyleInfo();
		var i, r, dxf, stripe1, stripe2, emptyStripe = new Asc.CTableStyleElement();
		if (style) {
			dxf = style.wholeTable && style.wholeTable.dxf;
			if (dxf) {
				sheetMergedStyles.setTablePivotStyle(range, dxf);
			}

			if (styleInfo.showColStripes) {
				stripe1 = style.firstColumnStripe || emptyStripe;
				stripe2 = style.secondColumnStripe || emptyStripe;
				if (stripe1.dxf) {
					sheetMergedStyles.setTablePivotStyle(range, stripe1.dxf,
						new Asc.CTableStyleStripe(stripe1.size, stripe2.size));
				}
				if (stripe2.dxf && range.c1 + stripe1.size <= range.c2) {
					sheetMergedStyles.setTablePivotStyle(
						new Asc.Range(range.c1 + stripe1.size, range.r1, range.c2, range.r2), stripe2.dxf,
						new Asc.CTableStyleStripe(stripe2.size, stripe1.size));
				}
			}
			if (styleInfo.showRowStripes) {
				stripe1 = style.firstRowStripe || emptyStripe;
				stripe2 = style.secondRowStripe || emptyStripe;
				if (stripe1.dxf) {
					sheetMergedStyles.setTablePivotStyle(range, stripe1.dxf,
						new Asc.CTableStyleStripe(stripe1.size, stripe2.size, true));
				}
				if (stripe2.dxf && range.r1 + stripe1.size <= range.r2) {
					sheetMergedStyles.setTablePivotStyle(
						new Asc.Range(range.c1, range.r1 + stripe1.size, range.c2, range.r2), stripe2.dxf,
						new Asc.CTableStyleStripe(stripe2.size, stripe1.size, true));
				}
			}

			dxf = style.firstColumn && style.firstColumn.dxf;
			if (styleInfo.showRowHeaders && dxf) {
				sheetMergedStyles.setTablePivotStyle(new Asc.Range(range.c1, range.r1, range.c1, range.r2), dxf);
			}

			dxf = style.headerRow && style.headerRow.dxf;
			if (styleInfo.showColHeaders && dxf) {
				sheetMergedStyles.setTablePivotStyle(new Asc.Range(range.c1, range.r1, range.c2, range.r1), dxf);
			}

			dxf = style.firstHeaderCell && style.firstHeaderCell.dxf;
			if (styleInfo.showColHeaders && styleInfo.showRowHeaders && dxf) {
				sheetMergedStyles.setTablePivotStyle(new Asc.Range(range.c1, range.r1, range.c1, range.r1), dxf);
			}

			if (pivot.asc_getColGrandTotals()) {
				dxf = style.lastColumn && style.lastColumn.dxf;
				if (dxf) {
					sheetMergedStyles.setTablePivotStyle(new Asc.Range(range.c2, range.r1, range.c2, range.r2), dxf);
				}
			}

			if (styleInfo.showRowHeaders) {
				for (i = range.r1 + 1; i < range.r2; ++i) {
					r = i - (range.r1 + 1);
					if (0 === r % 3) {
						dxf = style.firstRowSubheading;
					}
					if (dxf = (dxf && dxf.dxf)) {
						sheetMergedStyles.setTablePivotStyle(new Asc.Range(range.c1, i, range.c2, i), dxf);
					}
				}
			}

			if (pivot.asc_getRowGrandTotals()) {
				dxf = style.totalRow && style.totalRow.dxf;
				if (dxf) {
					sheetMergedStyles.setTablePivotStyle(new Asc.Range(range.c1, range.r2, range.c2, range.r2), dxf);
				}
			}
		}
	};

	WorkbookView.prototype.getTableStyles = function (props, bPivotTable) {
		var wb = this.model;
		var t = this;

		var result = [];
		var canvas = document.createElement('canvas');
		var tableStyleInfo;
		var pivotStyleInfo;

		var defaultStyles, styleThumbnailHeight, row, col = 5;
		var styleThumbnailWidth = window["IS_NATIVE_EDITOR"] ? 90 : 60;
		if(bPivotTable)
		{
			styleThumbnailHeight = 49;
			row = 8;
			defaultStyles =  wb.TableStyles.DefaultStylesPivot;
			pivotStyleInfo = props;
		}
		else
		{
			styleThumbnailHeight = window["IS_NATIVE_EDITOR"] ? 48 : 44;
			row = 5;
			defaultStyles = wb.TableStyles.DefaultStyles;
			tableStyleInfo = new AscCommonExcel.TableStyleInfo();
			if (props) {
				tableStyleInfo.ShowColumnStripes = props.asc_getBandVer();
				tableStyleInfo.ShowFirstColumn = props.asc_getFirstCol();
				tableStyleInfo.ShowLastColumn = props.asc_getLastCol();
				tableStyleInfo.ShowRowStripes = props.asc_getBandHor();
				tableStyleInfo.HeaderRowCount = props.asc_getFirstRow();
				tableStyleInfo.TotalsRowCount = props.asc_getLastRow();
			} else {
				tableStyleInfo.ShowColumnStripes = false;
				tableStyleInfo.ShowFirstColumn = false;
				tableStyleInfo.ShowLastColumn = false;
				tableStyleInfo.ShowRowStripes = true;
				tableStyleInfo.HeaderRowCount = true;
				tableStyleInfo.TotalsRowCount = false;
			}
		}

		styleThumbnailWidth = AscCommon.AscBrowser.convertToRetinaValue(styleThumbnailWidth, true);
		styleThumbnailHeight = AscCommon.AscBrowser.convertToRetinaValue(styleThumbnailHeight, true);

		canvas.width = styleThumbnailWidth;
		canvas.height = styleThumbnailHeight;
		var sizeInfo = {w: styleThumbnailWidth, h: styleThumbnailHeight, row: row, col: col};

		var ctx = new Asc.DrawingContext({canvas: canvas, units: 0/*px*/, fmgrGraphics: this.fmgrGraphics, font: this.m_oFont});

		var addStyles = function(styles, type, bEmptyStyle)
		{
			var style;
			for (var i in styles)
			{
				if ((bPivotTable && styles[i].pivot) || (!bPivotTable && styles[i].table))
				{
					if (window["IS_NATIVE_EDITOR"]) {
						//TODO empty style?
						window["native"]["BeginDrawStyle"](type, i);
					}
					t._drawTableStyle(ctx, styles[i], tableStyleInfo, pivotStyleInfo, sizeInfo);
					if (window["IS_NATIVE_EDITOR"]) {
						window["native"]["EndDrawStyle"]();
					} else {
						style = new AscCommon.CStyleImage();
						style.name = bEmptyStyle ? null : i;
						style.displayName = styles[i].displayName;
						style.type = type;
						style.image = canvas.toDataURL("image/png");
						result.push(style);
					}
				}
			}
		};

		addStyles(wb.TableStyles.CustomStyles, AscCommon.c_oAscStyleImage.Document);
		if (props) {
			//None style
			var emptyStyle = new Asc.CTableStyle();
			emptyStyle.displayName = "None";
			emptyStyle.pivot = bPivotTable;
			addStyles({null: emptyStyle}, AscCommon.c_oAscStyleImage.Default, true);
		}
		addStyles(defaultStyles, AscCommon.c_oAscStyleImage.Default);

		return result;
	};

  WorkbookView.prototype._drawTableStyle = function (ctx, style, tableStyleInfo, pivotStyleInfo, size) {
  	ctx.clear();

	var w = size.w;
	var h = size.h;
	var row = size.row;
	var col = size.col;

	var startX = 1;
	var startY = 1;

	var ySize = (h - 1) - 2 * startY;
	var xSize = w - 2 * startX;

	var stepY = (ySize) / row;
	var stepX = (xSize) / col;
	var lineStepX = (xSize - 1) / 5;

	var whiteColor = new CColor(255, 255, 255);
	var blackColor = new CColor(0, 0, 0);

	var defaultColor;
	if (!style || !style.wholeTable || !style.wholeTable.dxf.font) {
		defaultColor = blackColor;
	} else {
		defaultColor = style.wholeTable.dxf.font.getColor();
	}

	ctx.setFillStyle(whiteColor);
	ctx.fillRect(0, 0, xSize + 2 * startX, ySize + 2 * startY);

	var calculateLineVer = function(color, x, y1, y2)
	{
		ctx.beginPath();
		ctx.setStrokeStyle(color);

		ctx.lineVer(x + startX, y1 + startY, y2 + startY);

		ctx.stroke();
		ctx.closePath();
	};

	var calculateLineHor = function(color, x1, y, x2)
	{
		ctx.beginPath();
		ctx.setStrokeStyle(color);

		ctx.lineHor(x1 + startX, y + startY, x2 + startX);

		ctx.stroke();
		ctx.closePath();
	};

	var calculateRect = function(color, x1, y1, w, h)
	{
		ctx.beginPath();
		ctx.setFillStyle(color);
		ctx.fillRect(x1 + startX, y1 + startY, w, h);
		ctx.closePath();
	};

	var bbox = new Asc.Range(0, 0, col - 1, row - 1);
	var sheetMergedStyles = new AscCommonExcel.SheetMergedStyles();
	var hiddenManager = new AscCommonExcel.HiddenManager(null);

	if(pivotStyleInfo)
	{
		this.getPivotMergeStyle(sheetMergedStyles, bbox, style, pivotStyleInfo);
	}
	else if(tableStyleInfo)
	{
		style.initStyle(sheetMergedStyles, bbox, tableStyleInfo,
			null !== tableStyleInfo.HeaderRowCount ? tableStyleInfo.HeaderRowCount : 1,
			null !== tableStyleInfo.TotalsRowCount ? tableStyleInfo.TotalsRowCount : 0);
	}

	var compiledStylesArr = [];
	for (var i = 0; i < row; i++)
	{
		for (var j = 0; j < col; j++) {
			var color = null, prevStyle;
			var curStyle = AscCommonExcel.getCompiledStyle(sheetMergedStyles, hiddenManager, i, j);

			if(!compiledStylesArr[i])
			{
				compiledStylesArr[i] = [];
			}
			compiledStylesArr[i][j] = curStyle;

			//fill
			color = curStyle && curStyle.fill && curStyle.fill.bg();
			if(color)
			{
				calculateRect(color, j * stepX, i * stepY, stepX, stepY);
			}

			//borders
			//left
			prevStyle = (j - 1 >= 0) ? compiledStylesArr[i][j - 1] : null;
			color = AscCommonExcel.getMatchingBorder(prevStyle && prevStyle.border && prevStyle.border.r, curStyle && curStyle.border && curStyle.border.l);
			if(color && color.w > 0)
			{
				calculateLineVer(color.getColorOrDefault(), j * lineStepX, i * stepY, (i + 1) * stepY);
			}
			//right
			color = curStyle && curStyle.border && curStyle.border.r;
			if(color && color.w > 0)
			{
				calculateLineVer(color.getColorOrDefault(), (j + 1) * lineStepX, i * stepY, (i + 1) * stepY);
			}
			//top
			prevStyle = (i - 1 >= 0) ? compiledStylesArr[i - 1][j] : null;
			color = AscCommonExcel.getMatchingBorder(prevStyle && prevStyle.border && prevStyle.border.b, curStyle && curStyle.border && curStyle.border.t);
			if(color && color.w > 0)
			{
				calculateLineHor(color.getColorOrDefault(), j * stepX, i * stepY, (j + 1) * stepX);
			}
			//bottom
			color = curStyle && curStyle.border && curStyle.border.b;
			if(color && color.w > 0)
			{
				calculateLineHor(color.getColorOrDefault(), j * stepX, (i + 1) * stepY, (j + 1) * stepX);
			}

			//marks
			color = (curStyle && curStyle.font && curStyle.font.c) || defaultColor;
			calculateLineHor(color, j * lineStepX + 3, (i + 1) * stepY - stepY / 2, (j + 1) * lineStepX - 2);
		}
	}
  };

	WorkbookView.prototype.IsSelectionUse = function () {
        return !this.getWorksheet().getSelectionShape();
    };
	WorkbookView.prototype.GetSelectionRectsBounds = function () {
	    var ws = this.getWorksheet();
		if (ws.getSelectionShape()) {
            return null;
        }

		var range = ws.model.getSelection().getLast();
		var type = range.getType();
		var l = ws.getCellLeft(range.c1, 3);
		var t = ws.getCellTop(range.r1, 3);

		var offset = ws.getCellsOffset(3);

		return {
			X: asc.c_oAscSelectionType.RangeRow === type ? -offset.left : l - offset.left,
			Y: asc.c_oAscSelectionType.RangeCol === type ? -offset.top : t - offset.top,
			W: asc.c_oAscSelectionType.RangeRow === type ? offset.left :
				ws.getCellLeft(range.c2, 3) - l + ws.getColumnWidth(range.c2, 3),
			H: asc.c_oAscSelectionType.RangeCol === type ? offset.top :
				ws.getCellTop(range.r2, 3) - t + ws.getRowHeight(range.r2, 3),
			T: type
		};
	};
	WorkbookView.prototype.GetCaptionSize = function()
	{
		var offset = this.getWorksheet().getCellsOffset(3);
		return {
			W: offset.left,
			H: offset.top
		};
	};
	WorkbookView.prototype.ConvertXYToLogic = function (x, y) {
	  return this.getWorksheet().ConvertXYToLogic(x, y);
	};
	WorkbookView.prototype.ConvertLogicToXY = function (xL, yL) {
		return this.getWorksheet().ConvertLogicToXY(xL, yL)
	};
	WorkbookView.prototype.changeFormatTableInfo = function (tableName, optionType, val) {
		var ws = this.getWorksheet();
		return ws.af_changeFormatTableInfo(tableName, optionType, val);
	};

	WorkbookView.prototype.applyAutoCorrectOptions = function (val) {

		var api = window["Asc"]["editor"];
		var prevProps;
		switch (val) {
			case Asc.c_oAscAutoCorrectOptions.UndoTableAutoExpansion: {
				prevProps = {
					props: this.autoCorrectStore.props,
					cell: this.autoCorrectStore.cell,
					wsId: this.autoCorrectStore.wsId
				};
				api.asc_Undo();
				this.autoCorrectStore = prevProps;
				this.autoCorrectStore.props[0] = Asc.c_oAscAutoCorrectOptions.RedoTableAutoExpansion;
				this.toggleAutoCorrectOptions(true);
				break;
			}
			case Asc.c_oAscAutoCorrectOptions.RedoTableAutoExpansion: {
				prevProps = {
					props: this.autoCorrectStore.props,
					cell: this.autoCorrectStore.cell,
					wsId: this.autoCorrectStore.wsId
				};
				api.asc_Redo();
				this.autoCorrectStore = prevProps;
				this.autoCorrectStore.props[0] = Asc.c_oAscAutoCorrectOptions.UndoTableAutoExpansion;
				this.toggleAutoCorrectOptions(true);
				break;
			}
		}

		return true;
	};

	WorkbookView.prototype.toggleAutoCorrectOptions = function (isSwitch, val) {
		if (isSwitch) {
			if (val) {
				this.autoCorrectStore = val;
				var options = new Asc.asc_CAutoCorrectOptions();
				options.asc_setOptions(this.autoCorrectStore.props);
				options.asc_setCellCoord(
					this.getWorksheet().getCellCoord(this.autoCorrectStore.cell.c1, this.autoCorrectStore.cell.r1));

				this.handlers.trigger("asc_onToggleAutoCorrectOptions", options);
			} else if (this.autoCorrectStore) {
				if (this.autoCorrectStore.wsId === this.model.getActiveWs().getId()) {
					var options = new Asc.asc_CAutoCorrectOptions();
					options.asc_setOptions(this.autoCorrectStore.props);
					options.asc_setCellCoord(
						this.getWorksheet().getCellCoord(this.autoCorrectStore.cell.c1, this.autoCorrectStore.cell.r1));

					this.handlers.trigger("asc_onToggleAutoCorrectOptions", options);
				} else {
					this.handlers.trigger("asc_onToggleAutoCorrectOptions");
				}
			}
		} else {
			if(this.autoCorrectStore) {
				if (val) {
					this.autoCorrectStore = null;
				}
				this.handlers.trigger("asc_onToggleAutoCorrectOptions");
			}
		}
	};

	WorkbookView.prototype.savePagePrintOptions = function (arrPagesPrint) {
		var t = this;
		var viewMode = !this.canEdit();

		if(!arrPagesPrint) {
			return;
		}

		var callback = function (isSuccess) {
			if (false === isSuccess) {
				return;
			}

			for(var i in arrPagesPrint) {
				var ws = t.getWorksheet(parseInt(i));
				ws.savePageOptions(arrPagesPrint[i], viewMode);
				window["Asc"]["editor"]._onUpdateLayoutMenu(ws.model.Id);
			}
		};

		//разом лочу и настройки и колонтитулы. будут проблема - можно раъединить
		var lockInfoArr = [];
		var lockInfo;
		for(var i in arrPagesPrint) {
			lockInfo = this.getWorksheet(parseInt(i)).getLayoutLockInfo();
			lockInfoArr.push(lockInfo);
			if (arrPagesPrint[i].pageSetup && arrPagesPrint[i].pageSetup.headerFooter) {
				lockInfo = this.getWorksheet(parseInt(i)).getHeaderFooterLockInfo();
				lockInfoArr.push(lockInfo);
			}
		}

		if(viewMode) {
			callback();
		} else {
			this.collaborativeEditing.lock(lockInfoArr, callback);
		}
	};

	WorkbookView.prototype.cleanCutData = function (bDrawSelection, bCleanBuffer) {
		if(this.cutIdSheet) {
			var activeWs = this.wsViews[this.wsActive];
			var ws = this.getWorksheetById(this.cutIdSheet);

			if(bDrawSelection && activeWs && ws && activeWs.model.Id === ws.model.Id) {
				activeWs.cleanSelection();
			}

			if(ws) {
				ws.cutRange = null;
			}
			this.cutIdSheet = null;

			if(bDrawSelection && activeWs && ws && activeWs.model.Id === ws.model.Id) {
				activeWs.updateSelection();
			}
			//ms чистит буфер, если происходят какие-то изменения в документе с посвеченной областью вырезать
			//мы в данном случае не можем так делать, поскольку не можем прочесть информацию из буфера и убедиться, что
			//там именно тот вырезанный фрагмент. тем самым можем затереть какую-то важную информацию
			if(bCleanBuffer) {
				AscCommon.g_clipboardBase.ClearBuffer();
			}
		}
	};
	WorkbookView.prototype.updateGroupData = function () {
		this.getWorksheet()._updateGroups(true);
		this.getWorksheet()._updateGroups(null);
	};
	WorkbookView.prototype.pasteSheet = function (base64, insertBefore, name, callback) {

		var t = this;
		var tempWorkbook = new AscCommonExcel.Workbook();
        tempWorkbook.DrawingDocument = Asc.editor.wbModel.DrawingDocument;
		tempWorkbook.setCommonIndexObjectsFrom(this.model);
		var pasteProcessor = AscCommonExcel.g_clipboardExcel.pasteProcessor;
		var aPastedImages = pasteProcessor._readExcelBinary(base64.split('xslData;')[1], tempWorkbook, true);
		var pastedWs = tempWorkbook.aWorksheets[0];
		var newFonts = {};
		newFonts = tempWorkbook.generateFontMap2();
		newFonts = pasteProcessor._convertFonts(newFonts);
		for (var i = 0; i < pastedWs.Drawings.length; i++) {
			pastedWs.Drawings[i].graphicObject.getAllFonts(newFonts);
		}

		var doCopy = function() {
			History.Create_NewPoint();
			var renameParams = t.model.copyWorksheet(0, insertBefore, name, undefined, undefined, undefined, pastedWs, base64);
			//TODO ошибку по срезам добавил в renameParams. необходимо пересмотреть
			//переименовать эту переменную, либо не добавлять copySlicerError и посылать ошибку в другом месте
			if (renameParams && renameParams.copySlicerError) {
				t.handlers.trigger("asc_onError", c_oAscError.ID.MoveSlicerError, c_oAscError.Level.NoCritical);
			}
			callback(renameParams);
		};

		var api = window["Asc"]["editor"];
		api._loadFonts(newFonts, function () {
			if (aPastedImages && aPastedImages.length) {
				pasteProcessor._loadImagesOnServer(aPastedImages, function () {
					doCopy();
				});
			} else {
				doCopy();
			}
		});
	};

	WorkbookView.prototype.beforeInsertSlicer = function () {
		return this.getWorksheet().beforeInsertSlicer();
	};

	WorkbookView.prototype.insertSlicers = function (arr) {
		return this.getWorksheet().insertSlicers(arr);
	};

	WorkbookView.prototype.setFilterValuesFromSlicer = function (name, val) {
		var slicer = this.model.getSlicerByName(name);
		//нам нужно получить индекс листа где находится кэш данного среза
		var sheetIndex = slicer.getIndexSheetCache();
		if (sheetIndex !== null) {
			var ws = this.getWorksheet(sheetIndex);
			if (ws) {
				ws.setFilterValuesFromSlicer(slicer, val);
			}
		}
	};

	WorkbookView.prototype.deleteSlicer = function (name) {
		for(var i in this.wsViews) {
			this.wsViews[i].deleteSlicer(name);
		}
	};

	WorkbookView.prototype.deleteSlicers = function (names) {
		var slicers = [];
		for(var j = 0; j < names.length; j++) {
			for (var i in this.wsViews) {
				var slicer = this.wsViews[i].model.getSlicerByName(names[j]);
				if (slicer) {
					slicers.push({ws: this.wsViews[i], slicer: slicer});
					break;
				}
			}
		}

		var oThis = this;
		var callback = function (success) {
			var oWSView = oThis.getWorksheet();
			if (!success) {
                History.EndTransaction();
				oWSView.handlers.trigger("selectionChanged");
				return;
			}
			History.StartTransaction();

			for (var i = 0; i < slicers.length; i++) {
				slicers[i].ws.model.deleteSlicer(slicers[i].slicer.name);
			}

			History.EndTransaction();
            History.EndTransaction();
			oWSView.handlers.trigger("selectionChanged");
		};

		if (slicers && slicers.length) {
			this.checkLockSlicers(slicers, true, callback);
		}
	};

	WorkbookView.prototype.setSlicer = function (name, obj) {
		for(var i in this.wsViews) {
			var slicer = this.wsViews[i].model.getSlicerByName(name);
			if (slicer) {
				this.wsViews[i].setSlicer(slicer, obj);
				break;
			}
		}
	};

	WorkbookView.prototype.setSlicers = function (names, obj) {
		var slicers = [];
		for(var j = 0; j < names.length; j++) {
			for (var i in this.wsViews) {
				var slicer = this.wsViews[i].model.getSlicerByName(names[j]);
				if (slicer) {
					slicers.push({ws: this.wsViews[i], slicer: slicer});
					break;
				}
			}
		}

		var oThis = this;
		var callback = function (success) {
		    var oWSView = oThis.getWorksheet();
			if (!success) {
                oWSView.handlers.trigger("selectionChanged");
                //Transaction was started in applyDrawingProps in order to prevent save between applyDrawingProps and asc_setSlicers
                History.EndTransaction();
				return;
			}
            History.StartTransaction();
			for (var i = 0; i < slicers.length; i++) {
				slicers[i].slicer.set(obj);
			}
            History.EndTransaction();
            //Transaction was started in applyDrawingProps in order to prevent save between applyDrawingProps and asc_setSlicers
            History.EndTransaction();
            oWSView.handlers.trigger("selectionChanged");
		};

		if (slicers && slicers.length) {
			this.checkLockSlicers(slicers, true, callback);
		}
	};

	WorkbookView.prototype.checkLockSlicers = function (slicers, doLockRange, callback) {
		var t = this;
		var _lockMap = [];
		var lockInfoArr = [];
		var lockRanges = [];
		var cache, defNameId, lockInfo, lockRange, sheetId;
		for (var i = 0; i < slicers.length; i++) {
			cache = slicers[i].slicer.getCacheDefinition();
			sheetId =  slicers[i].ws.model.getId();
			if (!_lockMap[cache.name]) {
				_lockMap[cache.name] = 1;
				defNameId = this.model.dependencyFormulas.getDefNameByName(cache.name, sheetId);
				defNameId = defNameId ? defNameId.getNodeId() : null;
				lockInfo = this.collaborativeEditing.getLockInfo(AscCommonExcel.c_oAscLockTypeElem.Object, null, -1, defNameId);
				lockInfoArr.push(lockInfo);
				if (doLockRange) {
					lockRange = cache.getRange();
					if (lockRange) {
						lockRanges.push(this.collaborativeEditing.getLockInfo(AscCommonExcel.c_oAscLockTypeElem.Range, null, sheetId,
							new AscCommonExcel.asc_CCollaborativeRange(lockRange.c1, lockRange.r1, lockRange.c2, lockRange.r2)));
					}
				}
			}
		}

		var _callback = function (success) {
			if (!success && callback) {
				callback(false);
			}

			if (lockRanges && lockRanges.length) {
				t.collaborativeEditing.lock(lockRanges, callback);
			} else {
				callback(true);
			}
		};

		if (lockInfoArr && lockInfoArr.length) {
			this.collaborativeEditing.lock(lockInfoArr, _callback);
		} else {
			_callback(true);
		}
	};

	WorkbookView.prototype.updateSkin = function () {
		this.defaults.worksheetView.updateStyle();
	};

	WorkbookView.prototype.executeWithCurrentTopLeftCell = function (runFunction) {
		var i, oWS;
		var aTrueTopLeftCell = {};
		for(i in this.wsViews) {
			oWS = this.wsViews[i];
			if (oWS) {
				aTrueTopLeftCell[i] = oWS.model.getTopLeftCell();
				oWS.model.setTopLeftCell(oWS.getCurrentTopLeftCell());
			}
		}
		runFunction();
		for(i in this.wsViews) {
			oWS = this.wsViews[i];
			if (oWS) {
				oWS.model.setTopLeftCell(aTrueTopLeftCell[i]);
			}
		}
	};
	WorkbookView.prototype.convertEquationToMath = function (oEquation, isAll) {
		this.model.convertEquationToMath(oEquation, isAll);
	};
	WorkbookView.prototype.removeHandlersList = function () {
		var eventList = ["changeSheetViewSettings", "cleanCellCache", "changeWorksheetUpdate", "showWorksheet",
			"setSelection", "getSelectionState", "setSelectionState", "drawWS", "showDrawingObjects", "setCanUndo",
			"setCanRedo", "setDocumentModified", "updateWorksheetByModel", "undoRedoAddRemoveRowCols",
			"undoRedoHideSheet", "updateSelection", "asc_onLockDefNameManager", 'addComment', 'removeComment',
			'hiddenComments', 'showSolved', "hideSpecialPasteOptions", "toggleAutoCorrectOptions", "cleanCutData",
			"updateGroupData"];

		for (var i = 0; i < eventList.length; i++) {
			this.handlers.remove(eventList[i]);
		}
	};

	WorkbookView.prototype.sendCursor = function (needSend) {
		var CurTime = new Date().getTime();
		if (needSend || (true === this.NeedUpdateTargetForCollaboration && (CurTime - this.LastUpdateTargetTime > 1000)))
		{
			this.NeedUpdateTargetForCollaboration = false;
			var HaveChanges = History.Have_Changes(true);
			if (true !== HaveChanges)
			{
				var CursorInfo = this.getCursorInfo();
				if (null !== CursorInfo)
				{
					this.Api.CoAuthoringApi.sendCursor(CursorInfo);
					this.LastUpdateTargetTime = CurTime;
				}
			}
			else
			{
				this.LastUpdateTargetTime = CurTime;
			}
		}
	};
	WorkbookView.prototype.updateTargetForCollaboration = function () {
		this.NeedUpdateTargetForCollaboration = true;
	};

	WorkbookView.prototype.Update_ForeignCursor = function (CursorInfo, UserId, Show, UserShortId) {
		if (UserId === this.Api.CoAuthoringApi.getUserConnectionId())
			return;

		// "" - это означает, что курсор нужно удалить
		if (!CursorInfo || "" === CursorInfo) {
			this.Remove_ForeignCursor(UserId);
			return;
		}


		//AscCommon.getUserColorById(this.ShortId, null, true)

		//var CursorPos = [{Class : Run, Position : InRunPos}];
		var aCursorInfo = CursorInfo.split(",");
		var sDrawingData = aCursorInfo[0];
		var oWsView = this.getWorksheet();
		var oDrawingsController = null;
		if (oWsView && oWsView.objectRender) {
			oDrawingsController = oWsView.objectRender.controller;
		}
		AscFormat.drawingsUpdateForeignCursor(oDrawingsController, Asc.editor.wbModel.DrawingDocument, sDrawingData, UserId, Show, UserShortId);

		var selectionInfo = aCursorInfo[1];
		if (sDrawingData || !selectionInfo) {
			this.getWorksheet().cleanSelection();
			this.collaborativeEditing.Remove_ForeignCursor(UserId);
			this.getWorksheet()._drawSelection();
			return;
		}

		var Changes = new AscCommon.CCollaborativeChanges();
		var Reader = Changes.GetStream(selectionInfo);

		var sheetId = Reader.GetString2();
		var isEdit = Reader.GetBool();
		var sRanges = Reader.GetString2();

		var newCursorInfo = {sheetId: sheetId, isEdit: isEdit};
		var i = 0;
		var ranges = sRanges.split(",");
		while (i < ranges.length) {
			if (!newCursorInfo.ranges) {
				newCursorInfo.ranges = [];
			}
			if (i + 4 < ranges.length) {
				var _c1 = ranges[i] - 0;
				var _r1 = ranges[i + 1] - 0;
				var _c2 = ranges[i + 2] - 0;
				var _r2 = ranges[i + 3] - 0;

				newCursorInfo.ranges.push(new Asc.Range(_c1, _r1, _c2, _r2));
			}
			i += 4;
		}

		this.getWorksheet().cleanSelection();
		if (this.collaborativeEditing.Add_ForeignCursor(UserId, newCursorInfo, UserShortId)) {
			newCursorInfo.needDrawLabel = true;
		}
		this.getWorksheet()._drawSelection();

		//if (true === Show)
		//this.CollaborativeEditing.Update_ForeignCursorPosition(UserId, Run, InRunPos, true);
	};
	WorkbookView.prototype.Remove_ForeignCursor = function (UserId) {
		this.model.DrawingDocument.Collaborative_RemoveTarget(UserId);
		AscCommon.CollaborativeEditing.Remove_ForeignCursor(UserId);

		this.getWorksheet().cleanSelection();
		this.collaborativeEditing.Remove_ForeignCursor(UserId);
		this.Api.hideForeignSelectLabel(UserId);
		this.getWorksheet()._drawSelection();
	};
	WorkbookView.prototype.getCursorInfo = function () {
		var sSelectionInfo = "";
		var sDrawingData = "";
		var oWsView = this.getWorksheet();
		if (oWsView && oWsView.isSelectOnShape) {
			if (oWsView.objectRender) {
				sDrawingData = oWsView.objectRender.getDocumentPositionBinary();
			}
		} else {
			sSelectionInfo = this.getCursorInfoBinary();
		}
		return sDrawingData + "," + sSelectionInfo;
	};

	WorkbookView.prototype.getCursorInfoBinary = function () {

		var oWs = this.getActiveWS();
		var id = oWs.getId();
		var selection = oWs.getSelection();
		var isEdit = this.getCellEditMode();

		var rangeStr = "";
		for (var i = 0; i < selection.ranges.length; i++) {
			var _range = selection.ranges[i];
			rangeStr += _range.c1 + "," + _range.r1 + "," + _range.c2 + "," + _range.r2 + ",";
		}

		var oWriter = new AscCommon.CMemory(true);
		oWriter.CheckSize(50);

		var BinaryPos = oWriter.GetCurPosition();
		oWriter.WriteString2(id);
		oWriter.WriteBool(isEdit);
		oWriter.WriteString2(rangeStr);
		var BinaryLen = oWriter.GetCurPosition() - BinaryPos;
		return (BinaryLen + ";" + oWriter.GetBase64Memory2(BinaryPos, BinaryLen));
	};

	WorkbookView.prototype.updatePrintPreview = function () {
		if (!this.printPreviewState || !this.printPreviewState.isStart()) {
			return;
		}
		for(var i in this.wsViews) {
			this.wsViews[i]._recalculate();
		}
		if (!this.printPreviewState.isDrawPrintPreview) {
			var needUpdate;
			//if (this.workbook.printPreviewState.isNeedUpdate(this.model, this.getMaxRowColWithData())) {
			//возможно стоит добавить эвент об изменении количетсва страниц
			needUpdate = true;
			//}
			this.model.handlers.trigger("asc_onPrintPreviewSheetDataChanged", needUpdate);
		}
	};

	WorkbookView.prototype.setPrintOptionsJson = function (val) {
		//протаскиваю временные опции для печати
		//из данных опций пока испольуются только колонтитулы
		this.printOptionsJson = val;
	};

	WorkbookView.prototype.getPrintHeaderFooterFromJson = function (index) {
		var res = null;
		if (this.printOptionsJson) {
			var ws = this.model.getWorksheet(index);
			res = new Asc.CHeaderFooter(ws);
			if (this.printOptionsJson[index] && this.printOptionsJson[index]["pageSetup"] && this.printOptionsJson[index]["pageSetup"]["headerFooter"]) {
				res.initFromJson(this.printOptionsJson[index]["pageSetup"]["headerFooter"]);
			}
		}
		return res;
	};

	//------------------------------------------------------------export---------------------------------------------------
  window['AscCommonExcel'] = window['AscCommonExcel'] || {};
  window["AscCommonExcel"].WorkbookView = WorkbookView;
})(window);
