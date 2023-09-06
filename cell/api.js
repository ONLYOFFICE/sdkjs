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

var editor;
(/**
 * @param {Window} window
 * @param {undefined} undefined
 */
  function(window, undefined) {
  var asc = window["Asc"];
  var prot;

  var c_oAscAdvancedOptionsAction = AscCommon.c_oAscAdvancedOptionsAction;
  var c_oAscLockTypes = AscCommon.c_oAscLockTypes;
  var CColor = AscCommon.CColor;
  var g_oDocumentUrls = AscCommon.g_oDocumentUrls;
  var sendCommand = AscCommon.sendCommand;
  var parserHelp = AscCommon.parserHelp;
  var g_oIdCounter = AscCommon.g_oIdCounter;
  var g_oTableId = AscCommon.g_oTableId;

  var c_oAscLockTypeElem = AscCommonExcel.c_oAscLockTypeElem;

  var c_oAscError = asc.c_oAscError;
  var c_oAscFileType = asc.c_oAscFileType;
  var c_oAscAsyncAction = asc.c_oAscAsyncAction;
  var c_oAscAdvancedOptionsID = asc.c_oAscAdvancedOptionsID;
  var c_oAscAsyncActionType = asc.c_oAscAsyncActionType;

  var History = null;


  /**
   *
   * @param config
   * @constructor
   * @returns {spreadsheet_api}
   * @extends {AscCommon.baseEditorsApi}
   */
  function spreadsheet_api(config) {
    AscCommon.baseEditorsApi.call(this, config, AscCommon.c_oEditorId.Spreadsheet);

    /************ private!!! **************/
    this.topLineEditorName = config['id-input'] || '';
    this.topLineEditorElement = null;

    this.controller = null;

    this.handlers = new AscCommonExcel.asc_CHandlersList();

    this.fontRenderingMode = Asc.c_oAscFontRenderingModeType.hintingAndSubpixeling;
    this.wb = null;
    this.wbModel = null;
    this.tmpLCID = null;
    this.tmpDecimalSeparator = null;
    this.tmpGroupSeparator = null;
    this.tmpLocalization = null;

    // spellcheck
    this.defaultLanguage = 1033;
    this.spellcheckState = new AscCommonExcel.CSpellcheckState();

    this.documentFormatSave = c_oAscFileType.XLSX;

    // объекты, нужные для отправки в тулбар (шрифты, стили)
    this._gui_control_colors = null;
    this.GuiControlColorsMap = null;
    this.IsSendStandartColors = false;

    this.asyncMethodCallback = undefined;

    // Переменная отвечает, загрузились ли фонты
    this.FontLoadWaitComplete = false;
    //текущий обьект куда записываются информация для update, когда принимаются изменения в native редакторе
    this.oRedoObjectParamNative = null;

    this.collaborativeEditing = null;

    // AutoSave
    this.autoSaveGapRealTime = 30;	  // Интервал быстрого автосохранения (когда выставлен флаг realtime) - 30 мс.

    // Shapes
    this.isStartAddShape = false;
    this.shapeElementId = null;
    this.textArtElementId = null;

    //frozen pane border type
    this.frozenPaneBorderType = Asc.c_oAscFrozenPaneBorderType.shadow;

	  // Styles sizes
      this.styleThumbnailWidth = 100;
	  this.styleThumbnailHeight = 20;

    this.formulasList = null;	// Список всех формул

	this.openingEnd = {bin: false, xlsxStart: false, xlsx: false, data: null, perfStart: 0};

	this.tmpR1C1mode = null;

    this.isEditVisibleAreaOleEditor = false;

    this.insertDocumentUrlsData = null;

    this._init();
    return this;
  }
  spreadsheet_api.prototype = Object.create(AscCommon.baseEditorsApi.prototype);
  spreadsheet_api.prototype.constructor = spreadsheet_api;
  spreadsheet_api.prototype.sendEvent = function() {
    this.sendInternalEvent.apply(this, arguments);
    this.handlers.trigger.apply(this.handlers, arguments);
  };

  spreadsheet_api.prototype._init = function() {
    AscCommon.baseEditorsApi.prototype._init.call(this);
    this.topLineEditorElement = document.getElementById(this.topLineEditorName);
    // ToDo нужно ли это
    asc['editor'] = ( asc['editor'] || this );
  };

  spreadsheet_api.prototype._loadSdkImages = function () {
    var aImages = AscCommonExcel.getIconsForLoad();
    aImages.push(AscCommonExcel.sFrozenImageUrl, AscCommonExcel.sFrozenImageRotUrl);
    this.ImageLoader.bIsAsyncLoadDocumentImages = false;
    this.ImageLoader.LoadDocumentImages(aImages);
    this.ImageLoader.bIsAsyncLoadDocumentImages = true;
  };

  spreadsheet_api.prototype.asc_CheckGuiControlColors = function() {
    // потом реализовать проверку на то, что нужно ли посылать

    var arr_colors = new Array(10);
	var array_colors_types = [6, 15, 7, 16, 0, 1, 2, 3, 4, 5];
    var _count = arr_colors.length;
    for (var i = 0; i < _count; ++i) {
      var color = AscCommonExcel.g_oColorManager.getThemeColor(i);
      arr_colors[i] = new Asc.asc_CColor(color.getR(), color.getG(), color.getB());
      arr_colors[i].setColorSchemeId(array_colors_types[i])
    }

    // теперь проверим
    var bIsSend = false;
    if (this.GuiControlColorsMap != null) {
      for (var i = 0; i < _count; ++i) {
        var _color1 = this.GuiControlColorsMap[i];
        var _color2 = arr_colors[i];

        if ((_color1.r !== _color2.r) || (_color1.g !== _color2.g) || (_color1.b !== _color2.b)) {
          bIsSend = true;
          break;
        }
      }
    } else {
      this.GuiControlColorsMap = new Array(_count);
      bIsSend = true;
    }

    if (bIsSend) {
      for (var i = 0; i < _count; ++i) {
        this.GuiControlColorsMap[i] = arr_colors[i];
      }

      this.asc_SendControlColors();
    }
  };

  spreadsheet_api.prototype.asc_SendControlColors = function() {
    let standart_colors = null;
    if (!this.IsSendStandartColors) {
      let standartColors = AscCommon.g_oStandartColors;
      let _c_s = standartColors.length;
      standart_colors = new Array(_c_s);

      for (let i = 0; i < _c_s; ++i) {
        standart_colors[i] = new Asc.asc_CColor(standartColors[i].R, standartColors[i].G, standartColors[i].B);
      }

      this.IsSendStandartColors = true;
    }

    let _count = this.GuiControlColorsMap.length;

    let _ret_array = new Array(_count * 6);
    let _cur_index = 0;

    var array_colors_types = [6, 15, 7, 16, 0, 1, 2, 3, 4, 5];
    for (let i = 0; i < _count; ++i) {
      let basecolor = AscCommonExcel.g_oColorManager.getThemeColor(i);
      let aTints = AscCommonExcel.g_oThemeColorsDefaultModsSpreadsheet[AscCommon.GetDefaultColorModsIndex(basecolor.getR(), basecolor.getG(), basecolor.getB())];
      for (let j = 0, length = aTints.length; j < length; ++j) {
	    let tint = aTints[j];
	    let color = AscCommonExcel.g_oColorManager.getThemeColor(i, tint);
	    let oColor = new Asc.asc_CColor(color.getR(), color.getG(), color.getB());
		oColor.setColorSchemeId(array_colors_types[i]);
		oColor.put_effectValue(tint);
        _ret_array[_cur_index] = oColor;
        _cur_index++;
      }
    }

    this.asc_SendThemeColors(_ret_array, standart_colors);
  };

  spreadsheet_api.prototype.asc_getFunctionArgumentSeparator = function () {
    return AscCommon.FormulaSeparators.functionArgumentSeparator;
  };
  spreadsheet_api.prototype.asc_getCurrencySymbols = function () {
		var result = {};
		for (var key in AscCommon.g_aCultureInfos) {
			result[key] = AscCommon.g_aCultureInfos[key].CurrencySymbol;
		}
		return result;
	};
	spreadsheet_api.prototype.asc_getAdditionalCurrencySymbols = function () {
		return AscCommon.g_aAdditionalCurrencySymbols;
	};
	spreadsheet_api.prototype.asc_getLocaleExample = function(format, value, culture) {
		var cultureInfo = AscCommon.g_aCultureInfos[culture] || AscCommon.g_oDefaultCultureInfo;
		var numFormat = AscCommon.oNumFormatCache.get(format || "General");
		var res;
		if (null == value) {
			var ws = this.wbModel.getActiveWs();
			var activeCell = ws.selectionRange.activeCell;
			ws._getCellNoEmpty(activeCell.row, activeCell.col, function(cell) {
              if (cell) {
                res = cell.getValueForExample(numFormat, cultureInfo);
              } else {
                res = '';
              }
            });
		} else {
			res = numFormat.formatToChart(value, undefined, cultureInfo);
		}
		return res;
	};
	spreadsheet_api.prototype.asc_convertNumFormatLocal2NumFormat = function(format) {
		var oFormat = new AscCommon.CellFormat(format,undefined,true);
		return oFormat.toString(0, false);
	};
	spreadsheet_api.prototype.asc_convertNumFormat2NumFormatLocal = function(format) {
		var oFormat = new AscCommon.CellFormat(format,undefined,false);
		return oFormat.toString(0, true);
	};
	spreadsheet_api.prototype.asc_getFormatCells = function(info) {
		var res = AscCommon.getFormatCells(info);
		if (Asc.c_oAscNumFormatType.Custom === info.type) {
			var unique = {};
			res.forEach(function(elem) {
				unique[elem] = 1;
			});
			//todo delete button
			var wbNums = AscCommonExcel.g_StyleCache.getNumFormatStrings();
			wbNums.forEach(function(elem) {
				if (!unique[elem]) {
					unique[elem] = 1;
					res.push(elem);
				}
			});
		}
		return res;
	};
  spreadsheet_api.prototype.asc_getLocaleCurrency = function(val) {
    var cultureInfo = AscCommon.g_aCultureInfos[val];
    if (!cultureInfo) {
      cultureInfo = AscCommon.g_aCultureInfos[1033];
    }
    return AscCommonExcel.getCurrencyFormat(cultureInfo, 2, true, true, null);
  };


  spreadsheet_api.prototype.asc_getCurrentListType = function(){
      var ws = this.wb.getWorksheet();
      var oParaPr;
      if (ws && ws.objectRender && ws.objectRender.controller) {
          oParaPr = ws.objectRender.controller.getParagraphParaPr();
      }
      return new AscCommon.asc_CListType(AscFormat.fGetListTypeFromBullet(oParaPr && oParaPr.Bullet));
  };

  spreadsheet_api.prototype.asc_setLocale = function (LCID, decimalSeparator, groupSeparator) {
    if (!this.isLoadFullApi) {
      this.tmpLCID = LCID;
      this.tmpDecimalSeparator = decimalSeparator;
      this.tmpGroupSeparator = groupSeparator;
      return;
    }
    if (AscCommon.setCurrentCultureInfo(LCID, decimalSeparator, groupSeparator)) {
      parserHelp.setDigitSeparator(AscCommon.g_oDefaultCultureInfo.NumberDecimalSeparator);
      if (this.wbModel) {
        AscCommon.oGeneralEditFormatCache.cleanCache();
        AscCommon.oNumFormatCache.cleanCache();
        this.wbModel.rebuildColors();
        if (this.isDocumentLoadComplete) {
          AscCommon.checkCultureInfoFontPicker();
          this.wb && this.wb.cleanCache();
          this._loadFonts([], function () {
            this._onUpdateAfterApplyChanges();
          });
        }
      }
    }
  };
  spreadsheet_api.prototype.asc_getLocale = function () {
    return this.isLoadFullApi ? AscCommon.g_oDefaultCultureInfo.LCID : this.tmpLCID;
  };
  spreadsheet_api.prototype.asc_getDecimalSeparator = function (culture) {
  	var cultureInfo = AscCommon.g_aCultureInfos[culture] || AscCommon.g_oDefaultCultureInfo;
    return cultureInfo.NumberDecimalSeparator;
  };
  spreadsheet_api.prototype.asc_getGroupSeparator = function (culture) {
  	var cultureInfo = AscCommon.g_aCultureInfos[culture] || AscCommon.g_oDefaultCultureInfo;
  	return cultureInfo.NumberGroupSeparator;
  };
  spreadsheet_api.prototype.asc_getFrozenPaneBorderType = function() {
    return this.frozenPaneBorderType;
  };
  spreadsheet_api.prototype.asc_setFrozenPaneBorderType = function(nType) {
    if(this.frozenPaneBorderType !== nType) {
      this.frozenPaneBorderType = nType;
      if(this.wbModel) {
        this.asc_showWorksheet(this.asc_getActiveWorksheetIndex());
      }
    }
  };
  spreadsheet_api.prototype._openDocument = function(data) {
    this.wbModel = new AscCommonExcel.Workbook(this.handlers, this);
    this.initGlobalObjects(this.wbModel);
	  AscFonts.IsCheckSymbols = true;
	  if(this.isOpenOOXInBrowser) {
		  this.openingEnd.xlsx = true;
		  this.openingEnd.xlsxStart = true;
		  this.openingEnd.data = data;
	  } else {
		  var oBinaryFileReader = new AscCommonExcel.BinaryFileReader();
		  oBinaryFileReader.Read(data, this.wbModel);
	  }
	  AscFonts.IsCheckSymbols = false;
    this.openingEnd.bin = true;
    this._onEndOpen();
  };

  spreadsheet_api.prototype.initGlobalObjects = function(wbModel) {
    // History & global counters
    History.init(wbModel);

    AscCommonExcel.g_oUndoRedoCell = new AscCommonExcel.UndoRedoCell(wbModel);
    AscCommonExcel.g_oUndoRedoWorksheet = new AscCommonExcel.UndoRedoWoorksheet(wbModel);
    AscCommonExcel.g_oUndoRedoWorkbook = new AscCommonExcel.UndoRedoWorkbook(wbModel);
    AscCommonExcel.g_oUndoRedoCol = new AscCommonExcel.UndoRedoRowCol(wbModel, false);
    AscCommonExcel.g_oUndoRedoRow = new AscCommonExcel.UndoRedoRowCol(wbModel, true);
    AscCommonExcel.g_oUndoRedoComment = new AscCommonExcel.UndoRedoComment(wbModel);
    AscCommonExcel.g_oUndoRedoAutoFilters = new AscCommonExcel.UndoRedoAutoFilters(wbModel);
    AscCommonExcel.g_oUndoRedoSparklines = new AscCommonExcel.UndoRedoSparklines(wbModel);
    AscCommonExcel.g_DefNameWorksheet = new AscCommonExcel.Worksheet(wbModel, -1);
    AscCommonExcel.g_oUndoRedoSharedFormula = new AscCommonExcel.UndoRedoSharedFormula(wbModel);
    AscCommonExcel.g_oUndoRedoLayout = new AscCommonExcel.UndoRedoRedoLayout(wbModel);
    AscCommonExcel.g_oUndoRedoHeaderFooter = new AscCommonExcel.UndoRedoHeaderFooter(wbModel);
    AscCommonExcel.g_oUndoRedoArrayFormula = new AscCommonExcel.UndoRedoArrayFormula(wbModel);
    AscCommonExcel.g_oUndoRedoSortState = new AscCommonExcel.UndoRedoSortState(wbModel);
    AscCommonExcel.g_oUndoRedoSlicer = new AscCommonExcel.UndoRedoSlicer(wbModel);
    AscCommonExcel.g_oUndoRedoPivotTables = new AscCommonExcel.UndoRedoPivotTables(wbModel);
    AscCommonExcel.g_oUndoRedoPivotFields = new AscCommonExcel.UndoRedoPivotFields(wbModel);
    AscCommonExcel.g_oUndoRedoCF = new AscCommonExcel.UndoRedoCF(wbModel);
    AscCommonExcel.g_oUndoRedoProtectedRange = new AscCommonExcel.UndoRedoProtectedRange(wbModel);
    AscCommonExcel.g_oUndoRedoProtectedSheet = new AscCommonExcel.UndoRedoProtectedSheet(wbModel);
    AscCommonExcel.g_oUndoRedoProtectedWorkbook = new AscCommonExcel.UndoRedoProtectedWorkbook(wbModel);
    AscCommonExcel.g_oUndoRedoNamedSheetViews = new AscCommonExcel.UndoRedoNamedSheetViews(wbModel);
    AscCommonExcel.g_oUndoRedoUserProtectedRange = new AscCommonExcel.UndoRedoUserProtectedRange(wbModel);
  };

  spreadsheet_api.prototype.asc_DownloadAs = function (options) {
    if (!this.canSave || this.isFrameEditor() || c_oAscAdvancedOptionsAction.None !== this.advancedOptionsAction) {
      return;
    }
    if (this.isLongAction()) {
      return;
    }

    this.downloadAs(c_oAscAsyncAction.DownloadAs, options);
  };
	spreadsheet_api.prototype._saveCheck = function() {
		return !this.isFrameEditor() && c_oAscAdvancedOptionsAction.None === this.advancedOptionsAction &&
			!this.isLongAction() && !this.asc_getIsTrackShape() && !this.isOpenedChartFrame &&
			History.IsEndTransaction();
	};
	spreadsheet_api.prototype._haveOtherChanges = function () {
	  return this.collaborativeEditing.haveOtherChanges();
    };
	spreadsheet_api.prototype._prepareSave = function (isIdle) {
		var tmpHandlers;
		if (isIdle) {
			tmpHandlers = this.wbModel.handlers.handlers['asc_onError'];
			this.wbModel.handlers.handlers['asc_onError'] = null;
		}

      /* Нужно закрыть редактор (до выставления флага canSave, т.к. мы должны успеть отправить
       asc_onDocumentModifiedChanged для подписки на сборку) Баг http://bugzilla.onlyoffice.com/show_bug.cgi?id=28331 */
		if (!this.asc_closeCellEditor()) {
			if (isIdle) {
				this.asc_closeCellEditor(true);
			} else {
				return false;
			}
		}

		if (isIdle) {
			this.wbModel.handlers.handlers['asc_onError'] = tmpHandlers;
		}
		return true;
    };

	spreadsheet_api.prototype._printDesktop = function (options) {
		let advOpt = options && options.advancedOptions;
		if (advOpt && advOpt.asc_getNativeOptions() && advOpt.asc_getNativeOptions()["quickPrint"]) {
			advOpt.ignorePrintArea = true;
			advOpt.printType = Asc.c_oAscPrintType.EntireWorkbook;
		}

		window.AscDesktopEditor_PrintOptions = options;

		let desktopOptions = {};
		if (advOpt) {
			desktopOptions["nativeOptions"] = advOpt.asc_getNativeOptions();
		}

		window["AscDesktopEditor"]["Print"](JSON.stringify(desktopOptions));
		return true;
	};

  spreadsheet_api.prototype.asc_ChangePrintArea = function(type) {
	var ws = this.wb.getWorksheet();
	return ws.changePrintArea(type);
  };

  spreadsheet_api.prototype.asc_CanAddPrintArea = function() {
      var ws = this.wb.getWorksheet();
      return ws.canAddPrintArea();
  };

  spreadsheet_api.prototype.asc_SetPrintScale = function(width, height, scale) {
	  if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
		  return false;
	  }
	  var ws = this.wb.getWorksheet();
	  return ws.setPrintScale(width, height, scale);
  };

  spreadsheet_api.prototype.asc_Copy = function() {
    if (window["AscDesktopEditor"])
    {
      window["asc_desktop_copypaste"](this, "Copy");
      return true;
    }
    return AscCommon.g_clipboardBase.Button_Copy();
  };

  spreadsheet_api.prototype.asc_Paste = function() {
    if (window["AscDesktopEditor"])
    {
      window["asc_desktop_copypaste"](this, "Paste");
      return true;
    }
    if (!AscCommon.g_clipboardBase.IsWorking()) {
      return AscCommon.g_clipboardBase.Button_Paste();
    }
    return false;
  };

  spreadsheet_api.prototype.asc_SpecialPaste = function(props) {
    return AscCommon.g_specialPasteHelper.Special_Paste(props);
  };

  spreadsheet_api.prototype.asc_SpecialPasteData = function(props) {
	if (this.canEdit()) {
      this.wb.specialPasteData(props);
    }
  };

  spreadsheet_api.prototype.asc_TextImport = function(options, callback, bPaste) {
    //return this.asc_TextFromUrl(null, options, callback);
    //return this.asc_TextFromFile(options, callback);
    if (this.canEdit()) {
      var text;
      if(bPaste) {
        text = AscCommon.g_specialPasteHelper.GetPastedData(true);
      } else {
        var ws = this.wb.getWorksheet();
        text = ws.getRangeText();
      }
      if(!text) {
        //error
        //no data was selected to parse
        this.sendEvent('asc_onError', c_oAscError.ID.NoDataToParse, c_oAscError.Level.NoCritical);
        callback(false);
        return;
      }
      callback(AscCommon.parseText(text, options, true));
    }
  };

	spreadsheet_api.prototype.asc_TextFromFileOrUrl = function (options, callback, url) {
		if (this.canEdit()) {
			if (url) {
				this._getTextFromUrl(url, options, callback);
			} else {
				this._getTextFromFile(options, callback);
			}
		}
	};

	spreadsheet_api.prototype._getTextFromUrl = function (url, options, callback) {
		var t = this;
		if (this.canEdit()) {
			var document = {url: url, format: "TXT"};
			this.insertDocumentUrlsData = {
				imageMap: null, documents: [document], convertCallback: function (_api, url) {
					_api.insertDocumentUrlsData.imageMap = url;
					if (!url['output.txt']) {
						_api.endInsertDocumentUrls();
						_api.sendEvent("asc_onError", Asc.c_oAscError.ID.DirectUrl, Asc.c_oAscError.Level.NoCritical);
						return;
					}

					if (typeof Blob !== 'undefined' && typeof FileReader !== 'undefined') {
						AscCommon.loadFileContent(url['output.txt'], function (httpRequest) {
							var cp = {
								'codepage': AscCommon.c_oAscCodePageUtf8, "delimiter": AscCommon.c_oAscCsvDelimiter.Comma,
								'encodings': AscCommon.getEncodingParams()
							};

							if (httpRequest && httpRequest.response) {
								let data = httpRequest.response;
								var dataUint = new Uint8Array(data);
								var bom = AscCommon.getEncodingByBOM(dataUint);
								if (AscCommon.c_oAscCodePageNone !== bom.encoding) {
									cp['codepage'] = bom.encoding;
									data = dataUint.subarray(bom.size);
								}
								cp['data'] = data;
								callback(new AscCommon.asc_CAdvancedOptions(cp));
							} else {
								t.handlers.trigger("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.Critical);
							}
							_api.endInsertDocumentUrls();
						}, "arraybuffer");
					}
				}, endCallback: function (_api) {
				}
			};

			var _options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.TXT);
			_options.isNaturalDownload = true;
			_options.isGetTextFromUrl = true;
			if (document.url) {
				_options.errorDirect = Asc.c_oAscError.ID.DirectUrl;
			}
			this.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, _options);
		}
	};

	spreadsheet_api.prototype._getTextFromFile = function (options, callback) {
		let t = this;

		function wrapper_callback(data) {
			let bom = AscCommon.getEncodingByBOM(data);
			let cp = {
				'codepage': AscCommon.c_oAscCodePageNone !== bom.encoding ? bom.encoding : AscCommon.c_oAscCodePageUtf8,
				"delimiter": AscCommon.c_oAscCsvDelimiter.Comma,
				'encodings': AscCommon.getEncodingParams(),
				'data': AscCommon.c_oAscCodePageNone !== bom.encoding ? data.subarray(bom.size) : data
			};
			callback(new AscCommon.asc_CAdvancedOptions(cp));
		}

		if (window["AscDesktopEditor"]) {
			// TODO: add translations
			window["AscDesktopEditor"]["OpenFilenameDialog"]("csv/txt", false, function (_file) {
				let file = _file;
				if (Array.isArray(file))
					file = file[0];
				if (!file)
					return;

				window["AscDesktopEditor"]["loadLocalFile"](file, function(uint8Array) {
					if (!uint8Array)
						return;

					wrapper_callback(uint8Array);
				});
			});
			return;
		}

		AscCommon.ShowTextFileDialog(function (error, files) {
			if (Asc.c_oAscError.ID.No !== error) {
				t.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
				return;
			}

			let reader = new FileReader();
			reader.onload = function () {
				wrapper_callback(new Uint8Array(reader.result));
			};

			reader.onerror = function () {
				t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
			};

			reader.readAsArrayBuffer(files[0]);
		});
	};

	spreadsheet_api.prototype.asc_ImportXmlStart = function (callback) {
		let t = this;

		if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]()) {
			// TODO: add translations
			window["AscDesktopEditor"]["OpenFilenameDialog"]("(*.xml)", false, function (_file) {
				let file = _file;
				if (Array.isArray(file)) {
					file = file[0];
				}
				if (!file) {
					callback(null);
					return;
				}

				window["AscDesktopEditor"]["convertFile"](file, 0x2002, function (convertedFile) {
					let stream = null;
					if (convertedFile) {
						stream = convertedFile["get"](/*Editor.bin*/);
						convertedFile["close"]();
					}
					callback(stream ? new Uint8Array(stream) : null);
				});
			});
		} else if (!window["NATIVE_EDITOR_ENJINE"]) {
			AscCommon.ShowXmlFileDialog(function (error, files) {
				if (Asc.c_oAscError.ID.No !== error) {
					t.sendEvent("asc_onError", error, Asc.c_oAscError.Level.NoCritical);
					callback(null);
					return;
				}

				let format = AscCommon.GetFileExtension(files[0].name);
				let reader = new FileReader();
				reader.onload = function () {
					t._convertFromXml({data: new Uint8Array(reader.result), format: format}, callback);
				};
				reader.onerror = function () {
					t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
					callback(null);
				};

				reader.readAsArrayBuffer(files[0]);
			});
		} else {
			callback(null);
		}
	};

	spreadsheet_api.prototype.asc_ImportXmlEnd = function (stream, dataRef, newSheetName) {
		let t = this;


		let doInsertXml = function (_dataRef, _ws) {
			let binaryData = stream;
			if (!window["AscDesktopEditor"]) {
				//xlst
				binaryData = null;
				let jsZlib = new AscCommon.ZLib();
				if (!stream || !jsZlib.open(stream)) {
					//t.model.handlers.trigger("asc_onErrorUpdateExternalReference", eR.Id);
					return false;
				}

				if (jsZlib.files && jsZlib.files.length) {
					binaryData = jsZlib.getFile(jsZlib.files[0]);
				}
			}

			if (!binaryData) {
				return;
			}

			//заполняем через банарник
			let oBinaryFileReader = new AscCommonExcel.BinaryFileReader(true);
			//чтобы лишнего не читать, проставляю флаг копипаст
			oBinaryFileReader.InitOpenManager.copyPasteObj = {
				isCopyPaste: true, activeRange: null, selectAllSheet: true
			};

			let wb = new AscCommonExcel.Workbook();
			wb.DrawingDocument = Asc.editor.wbModel.DrawingDocument;

			AscFormat.ExecuteNoHistory(function () {
				AscCommonExcel.executeInR1C1Mode(false, function () {
					oBinaryFileReader.Read(binaryData, wb);
				});
			});

			if (wb.aWorksheets) {
				let arrSheets = wb.aWorksheets;
				let pastedSheet = arrSheets && arrSheets[0];
				if (pastedSheet) {
					if (_ws) {
						t.wbModel.setActive(_ws.index);
						t.wb.updateWorksheetByModel();
						t.wb.showWorksheet();
					}
					let ws = t.wb.getWorksheet();
					if (_dataRef) {
						ws.setSelection(_dataRef);
					}
					AscCommonExcel.g_clipboardExcel.pasteProcessor.activeRange = new Asc.Range(0, 0, Math.max(pastedSheet.nColsCount - 1, 0), Math.max(pastedSheet.nRowsCount - 1, 0)).getName();
					t.wb.getWorksheet().setSelectionInfo('paste', {data: pastedSheet, fromBinary: true, fontsNew: [], pasteAllSheet: true, wb: wb});
				}
			}
		};

		let doCheckRange = function (_sDataRange) {
			let result = parserHelp.parse3DRef(_sDataRange);
			let _range, sheetModel;
			if (result)
			{
				sheetModel = t.wb.model.getWorksheetByName(result.sheet);
				if (sheetModel)
				{
					_range = AscCommonExcel.g_oRangeCache.getAscRange(result.range);
				}
			} else {
				_range = AscCommonExcel.g_oRangeCache.getAscRange(_sDataRange);
			}
			if (!_range) {
				_range = AscCommon.rx_defName.test(_sDataRange);
			}
			if (!_range) {
				_range = parserHelp.isTable(_sDataRange, 0, true);
			}

			return _range ? {range: _range, sheetModel: sheetModel} : false;
		};

		let wb = this.wbModel;
		let alreadyAddedSheet = newSheetName && t.wb.model.getWorksheetByName(newSheetName);
		if (newSheetName && !alreadyAddedSheet) {
			this._isLockedAddWorksheets(function(res) {
				if (res) {
					History.Create_NewPoint();
					History.StartTransaction();
					t._addWorksheetsWithoutLock([newSheetName], wb.getActive());
					doInsertXml();
					History.EndTransaction();
				} else {
					//todo
					t.sendEvent('asc_onError', c_oAscError.ID.LockedCellPivot, c_oAscError.Level.NoCritical);
				}
			});
		} else {
			let _checkRange = doCheckRange(dataRef);
			if (_checkRange) {
				History.Create_NewPoint();
				History.StartTransaction();
				doInsertXml(_checkRange.range, _checkRange.sheetModel);
				History.EndTransaction();
			} else {
				this.sendEvent('asc_onError', c_oAscError.ID.PivotLabledColumns, c_oAscError.Level.NoCritical);
			}
		}
	};


	spreadsheet_api.prototype._convertFromXml = function (document, callback) {
		let stream = null;
		this.insertDocumentUrlsData = {
			imageMap: null, documents: [document], convertCallback: function (_api, url) {
				_api.insertDocumentUrlsData.imageMap = url;
				if (!url['output.xlst']) {
					_api.endInsertDocumentUrls();
					_api.sendEvent("asc_onError", Asc.c_oAscError.ID.DirectUrl,
						Asc.c_oAscError.Level.NoCritical);
					callback(null);
					return;
				}
				AscCommon.loadFileContent(url['output.xlst'], function (httpRequest) {
					if (null === httpRequest || !(stream = AscCommon.initStreamFromResponse(httpRequest))) {
						_api.endInsertDocumentUrls();
						_api.sendEvent("asc_onError", Asc.c_oAscError.ID.DirectUrl,
							Asc.c_oAscError.Level.NoCritical);
						callback(null);
						return;
					}
					_api.endInsertDocumentUrls();
				}, "arraybuffer");
			}, endCallback: function (_api) {

				if (stream) {
					callback(stream);
				} else {
					callback(null);
				}
			}
		};

		let options = new Asc.asc_CDownloadOptions(Asc.c_oAscFileType.XLSY);
		options.isNaturalDownload = true;
		options.isGetTextFromUrl = true;
		if (document.url) {
			options.errorDirect = Asc.c_oAscError.ID.DirectUrl;
		}
		this.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, options);
	};

	spreadsheet_api.prototype.asc_TextToColumns = function (options, opt_text, opt_activeRange) {
		if (this.canEdit()) {
			var ws = this.wb.getWorksheet();
			var text = opt_text ? opt_text : ws.getRangeText();
			var specialPasteHelper = window['AscCommon'].g_specialPasteHelper;
			if (!specialPasteHelper.specialPasteProps) {
				specialPasteHelper.specialPasteProps = new Asc.SpecialPasteProps();
			}
			specialPasteHelper.specialPasteProps.property = Asc.c_oSpecialPasteProps.useTextImport;
			specialPasteHelper.specialPasteProps.asc_setAdvancedOptions(options);

			//remove last empty string
			if (opt_text && opt_text.length > 1 && opt_text[opt_text.length - 1] &&
				opt_text[opt_text.length - 1].length === 1 && opt_text[opt_text.length - 1][0] === "") {
				opt_text.splice(opt_text.length - 1, 1);
			}

			var selectionRange;
			var activeSheet = null;
			if (opt_activeRange) {
				var is3dRef = parserHelp.parse3DRef(opt_activeRange);
				//TODO вставка на другой лист
				var range, sheetModel;
				if (is3dRef) {
					sheetModel = this.wb.model.getWorksheetByName(is3dRef.sheet);
					if (sheetModel) {
						range = AscCommonExcel.g_oRangeCache.getAscRange(is3dRef.range);
					}
				} else {
					range = AscCommonExcel.g_oRangeCache.getAscRange(opt_activeRange);
				}
				if (sheetModel && sheetModel !== ws.model) {
					activeSheet = sheetModel;
					this.wb.model.nActive = sheetModel.index;
				}
				if (range) {
					if (activeSheet) {
						selectionRange = activeSheet.selectionRange.clone();
						activeSheet.selectionRange.ranges = [range.clone()];
					} else {
						selectionRange = ws.model.selectionRange.clone();
						ws.model.selectionRange.ranges = [range.clone()];
					}
				}
			}

			this.wb.pasteData(AscCommon.c_oAscClipboardDataFormat.Text, text, null, null, true);

			if (selectionRange) {
				if (activeSheet !== null) {
					activeSheet.selectionRange = selectionRange;
					this.wb.model.nActive = ws.model.index;
					ws.draw();
				} else {
					ws.cleanSelection();
					ws.model.selectionRange = selectionRange;
					ws._drawSelection();
				}
			}
		}
	};

  spreadsheet_api.prototype.asc_ShowSpecialPasteButton = function(props) {
      if (this.canEdit()) {
          this.wb.showSpecialPasteButton(props);
      }
  };

  spreadsheet_api.prototype.asc_UpdateSpecialPasteButton = function(props) {
      if (this.canEdit()) {
          this.wb.updateSpecialPasteButton(props);
      }
  };

  spreadsheet_api.prototype.asc_HideSpecialPasteButton = function() {
      if (this.canEdit()) {
          this.wb.hideSpecialPasteButton();
      }
 };

  spreadsheet_api.prototype.asc_Cut = function() {
    if (window["AscDesktopEditor"])
    {
      window["asc_desktop_copypaste"](this, "Cut");
      return true;
    }
    return AscCommon.g_clipboardBase.Button_Cut();
  };

  spreadsheet_api.prototype.asc_PasteData = function (_format, data1, data2, text_data) {
    if (this.canEdit()) {
      this.wb.pasteData(_format, data1, data2, text_data, arguments[5]);
      //this.asc_EndMoveSheet2(data1, 1, "test2");
    }
  };

  spreadsheet_api.prototype.asc_CheckCopy = function (_clipboard /* CClipboardData */, _formats) {
    return this.wb.checkCopyToClipboard(_clipboard, _formats);
  };

  spreadsheet_api.prototype.asc_SelectionCut = function () {
    if (this.canEdit()) {
      this.wb.selectionCut();
    }
  };

  spreadsheet_api.prototype.asc_bIsEmptyClipboard = function() {
    var result = this.wb.bIsEmptyClipboard();
    this.wb.restoreFocus();
    return result;
  };

  spreadsheet_api.prototype.asc_Undo = function() {
    if (!this.canUndoRedoByRestrictions())
      return;
    this.wb.undo();
    this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_Redo = function() {
    if (!this.canUndoRedoByRestrictions())
      return;
    this.wb.redo();
    this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_Resize = function () {
    var oldScale = AscCommon.AscBrowser.retinaPixelRatio;
    AscCommon.AscBrowser.checkZoom();
    if (this.wb) {
      if (Math.abs(oldScale - AscCommon.AscBrowser.retinaPixelRatio) > 0.001) {
        this.wb.changeZoom(null);
        this._sendWorkbookStyles();
      }
      this.wb.resize();

      if (AscCommon.g_inputContext) {
        AscCommon.g_inputContext.onResize("ws-canvas-outer");
      }
    }
  };

  spreadsheet_api.prototype.asc_addAutoFilter = function(styleName, addFormatTableOptionsObj) {
    var ws = this.wb.getWorksheet();
    return ws.addAutoFilter(styleName, addFormatTableOptionsObj);
  };

  spreadsheet_api.prototype.asc_changeAutoFilter = function(tableName, optionType, val) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    return ws.changeAutoFilter(tableName, optionType, val);
  };

  spreadsheet_api.prototype.asc_applyAutoFilter = function(autoFilterObject) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    ws.applyAutoFilter(autoFilterObject);
  };

  spreadsheet_api.prototype.asc_applyAutoFilterByType = function(autoFilterObject) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    ws.applyAutoFilterByType(autoFilterObject);
  };

  spreadsheet_api.prototype.asc_reapplyAutoFilter = function(displayName) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    ws.reapplyAutoFilter(displayName);
  };

  spreadsheet_api.prototype.asc_sortColFilter = function(type, cellId, displayName, color, bIsExpandRange) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    ws.sortRange(type, cellId, displayName, color, bIsExpandRange);
  };

  spreadsheet_api.prototype.asc_getAddFormatTableOptions = function(range) {
    var ws = this.wb.getWorksheet();
    return ws.getAddFormatTableOptions(range);
  };

  spreadsheet_api.prototype.asc_clearFilter = function() {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
  	var ws = this.wb.getWorksheet();
    return ws.clearFilter();
  };

  spreadsheet_api.prototype.asc_clearFilterColumn = function(cellId, displayName) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    return ws.clearFilterColumn(cellId, displayName);
  };

  spreadsheet_api.prototype.asc_changeSelectionFormatTable = function(tableName, optionType) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    return ws.changeTableSelection(tableName, optionType);
  };

  spreadsheet_api.prototype.asc_changeFormatTableInfo = function(tableName, optionType, val) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    return this.wb.changeFormatTableInfo(tableName, optionType, val);
  };

  spreadsheet_api.prototype.asc_applyAutoCorrectOptions = function(val) {
      this.wb.applyAutoCorrectOptions(val);
  };

  spreadsheet_api.prototype.asc_insertCellsInTable = function(tableName, optionType) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    return ws.af_insertCellsInTable(tableName, optionType);
  };

  spreadsheet_api.prototype.asc_deleteCellsInTable = function(tableName, optionType) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    return ws.af_deleteCellsInTable(tableName, optionType);
  };

  spreadsheet_api.prototype.asc_changeDisplayNameTable = function(tableName, newName) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    return ws.af_changeDisplayNameTable(tableName, newName);
  };

  spreadsheet_api.prototype.asc_changeTableRange = function(tableName, range) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    return ws.af_changeTableRange(tableName, range);
  };

  spreadsheet_api.prototype.asc_convertTableToRange = function(tableName) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    return ws.af_convertTableToRange(tableName);
  };

	spreadsheet_api.prototype.asc_getTablePictures = function (props, pivot) {
		return this.wb.getTableStyles(props, pivot);
	};
	spreadsheet_api.prototype.asc_getSlicerPictures = function () {
		return this.wb.getSlicerStyles();
	};

  spreadsheet_api.prototype.asc_setViewMode = function (isViewMode) {
    this.isViewMode = !!isViewMode;
    if (!this.isLoadFullApi) {
      return;
    }
    if (this.collaborativeEditing) {
      this.collaborativeEditing.setViewerMode(isViewMode);
    }
    if(this.isViewMode) {
        this.turnOffSpecialModes();
    }
  };

	  spreadsheet_api.prototype.asc_setFilteringMode = function (mode) {
		  window['AscCommonExcel'].filteringMode = !!mode;
	  };

  /*
   idOption идентификатор дополнительного параметра, пока c_oAscAdvancedOptionsID.CSV.
   option - какие свойства применить, пока массив. для CSV объект asc_CTextOptions(codepage, delimiter)
   exp:	asc_setAdvancedOptions(c_oAscAdvancedOptionsID.CSV, new Asc.asc_CTextOptions(1200, c_oAscCsvDelimiter.Comma) );
   */
  spreadsheet_api.prototype.asc_setAdvancedOptions = function(idOption, option) {
    // Проверяем тип состояния в данный момент
    if (this.advancedOptionsAction !== c_oAscAdvancedOptionsAction.Open) {
      return;
    }
    if (AscCommon.EncryptionWorker.asc_setAdvancedOptions(this, idOption, option)) {
      return;
    }

    var v;
    switch (idOption) {
      case c_oAscAdvancedOptionsID.CSV:
        v = {
          "id": this.documentId,
          "userid": this.documentUserId,
          "format": this.documentFormat,
          "c": "reopen",
          "title": this.documentTitle,
          "delimiter": option.asc_getDelimiter(),
          "delimiterChar": option.asc_getDelimiterChar(),
          "codepage": option.asc_getCodePage(),
          "nobase64": true
        };
        sendCommand(this, null, v);
        break;
      case c_oAscAdvancedOptionsID.DRM:
        this.currentPassword = option.asc_getPassword();
        v = {
          "id": this.documentId,
          "userid": this.documentUserId,
          "format": this.documentFormat,
          "c": "reopen",
          "title": this.documentTitle,
          "password": option.asc_getPassword(),
          "nobase64": true
        };
        sendCommand(this, null, v);
        break;
    }
  };
  // Опции страницы (для печати)
  spreadsheet_api.prototype.asc_setPageOptions = function(options, index) {
    var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
    this.wb.getWorksheet(sheetIndex).setPageOptions(options);
  };

  spreadsheet_api.prototype.asc_savePagePrintOptions = function(arrPagesPrint) {
      this.wb.savePagePrintOptions(arrPagesPrint);
  };

    spreadsheet_api.prototype.getDrawingObjects = function () {
        var oController = this.getGraphicController();
        if (oController) {
            return oController.drawingObjects;
        }
    };

    spreadsheet_api.prototype.getDrawingDocument = function () {
        return this.wbModel && this.wbModel.DrawingDocument
    };

  spreadsheet_api.prototype.asc_getPageOptions = function(index, initPrintTitles, opt_copy) {
    var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
    var ws = this.wbModel.getWorksheet(sheetIndex);
    var printOptions = ws.PagePrintOptions;
	//TODO похожая инициализация в getPrintOptionsJson - сделать общую
	if(initPrintTitles && printOptions) {
		printOptions.initPrintTitles();
    }
    if(printOptions && opt_copy) {
		printOptions = ws.PagePrintOptions.clone();
		printOptions.pageSetup.headerFooter = ws && ws.headerFooter && ws.headerFooter.getForInterface();
		var printArea = this.wbModel.getDefinesNames("Print_Area", ws.getId());
		printOptions.pageSetup.printArea = printArea ? printArea.clone() : false;

		printOptions.printTitlesHeight = ws.PagePrintOptions.printTitlesHeight;
		printOptions.printTitlesWidth = ws.PagePrintOptions.printTitlesWidth;

		if (ws.PagePrintOptions && ws.PagePrintOptions.pageSetup) {
			printOptions.pageSetup.fitToHeight = ws.PagePrintOptions.pageSetup.asc_getFitToHeight();
			printOptions.pageSetup.fitToWidth = ws.PagePrintOptions.pageSetup.asc_getFitToWidth();
		}
    }

    return printOptions;
  };

  spreadsheet_api.prototype.asc_setPageOption = function (func, val, index) {
	  if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
		  return false;
	  }
      var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
      var ws = this.wb.getWorksheet(sheetIndex);
      ws.setPageOption(func, val);
  };

  spreadsheet_api.prototype.asc_changeDocSize = function (width, height, index) {
	if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
	  return false;
	}
    var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
    var ws = this.wb.getWorksheet(sheetIndex);
    ws.changeDocSize(width, height);
  };

  spreadsheet_api.prototype.asc_changePageMargins = function (oMargins, bHorCentered, bVerCentered, nHeader, nFooter, index) {
	  if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
		  return false;
	  }
	  var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
      var ws = this.wb.getWorksheet(sheetIndex);
      ws.changePageMargins(oMargins, bHorCentered, bVerCentered, nHeader, nFooter);
  };

  spreadsheet_api.prototype.asc_changePageOrient = function (isPortrait, index) {
	if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
	  return false;
	}
  	var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
      var ws = this.wb.getWorksheet(sheetIndex);
      if (isPortrait) {
          ws.changePageOrient(Asc.c_oAscPageOrientation.PagePortrait);
      } else {
          ws.changePageOrient(Asc.c_oAscPageOrientation.PageLandscape);
      }
  };

	spreadsheet_api.prototype.asc_SetPrintHeadings = function(val, index) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}
		var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
		var ws = this.wb.getWorksheet(sheetIndex);
		ws.setPrintHeadings(val);
	};

	spreadsheet_api.prototype.asc_SetPrintGridlines = function(val, index) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}

		var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
		var ws = this.wb.getWorksheet(sheetIndex);
		ws.setGridLines(val);
	};

  spreadsheet_api.prototype.asc_changePrintTitles = function (cols, rows, index) {
	  if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
		  return false;
	  }
	  var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
	  var ws = this.wb.getWorksheet(sheetIndex);
	  ws.changePrintTitles(cols, rows);
  };

  spreadsheet_api.prototype.asc_getPrintTitlesRange = function (prop, byHeight, index) {
      var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
      var ws = this.wb.getWorksheet(sheetIndex);
      return ws.getPrintTitlesRange(prop, byHeight);
  };

  spreadsheet_api.prototype._onNeedParams = function(data, opt_isPassword) {
    var t = this;
    // Проверяем, возможно нам пришли опции для CSV
    if (this.documentOpenOptions && !opt_isPassword) {
      var codePageCsv = AscCommon.c_oAscEncodingsMap[this.documentOpenOptions["codePage"]] || AscCommon.c_oAscCodePageUtf8, delimiterCsv = this.documentOpenOptions["delimiter"],
		  delimiterCharCsv = this.documentOpenOptions["delimiterChar"];
      if (null != codePageCsv && (null != delimiterCsv || null != delimiterCharCsv)) {
        this.asc_setAdvancedOptions(c_oAscAdvancedOptionsID.CSV, new asc.asc_CTextOptions(codePageCsv, delimiterCsv, delimiterCharCsv));
        return;
      }
    }
	if (opt_isPassword) {
		if (t.handlers.hasTrigger("asc_onAdvancedOptions")) {
			asc["editor"].sendEvent("asc_onDocumentPassword", true);
			t.handlers.trigger("asc_onAdvancedOptions", c_oAscAdvancedOptionsID.DRM);
		} else {
			t.handlers.trigger("asc_onError", c_oAscError.ID.ConvertationPassword, c_oAscError.Level.Critical);
		}
	} else {
		if (t.handlers.hasTrigger("asc_onAdvancedOptions")) {
			// ToDo разделитель пока только "," http://bugzilla.onlyoffice.com/show_bug.cgi?id=31009
			var cp = {
				'codepage': AscCommon.c_oAscCodePageUtf8, "delimiter": AscCommon.c_oAscCsvDelimiter.Comma,
				'encodings': AscCommon.getEncodingParams()
			};
			if (data && typeof Blob !== 'undefined' && typeof FileReader !== 'undefined') {
				AscCommon.loadFileContent(data, function(httpRequest) {
					if (httpRequest && httpRequest.response) {
						let data = httpRequest.response;
						var dataUint = new Uint8Array(data);
						var bom = AscCommon.getEncodingByBOM(dataUint);
						if (AscCommon.c_oAscCodePageNone !== bom.encoding) {
							cp['codepage'] = bom.encoding;
							data = dataUint.subarray(bom.size);
						}
						cp['data'] = data;
						t.handlers.trigger("asc_onAdvancedOptions", c_oAscAdvancedOptionsID.CSV, new AscCommon.asc_CAdvancedOptions(cp));
					} else {
						t.handlers.trigger("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.Critical);
					}
				}, "arraybuffer");
			} else {
				t.handlers.trigger("asc_onAdvancedOptions", c_oAscAdvancedOptionsID.CSV, new AscCommon.asc_CAdvancedOptions(cp));
			}
		} else {
			this.asc_setAdvancedOptions(c_oAscAdvancedOptionsID.CSV, new asc.asc_CTextOptions(AscCommon.c_oAscCodePageUtf8, AscCommon.c_oAscCsvDelimiter.Comma));
		}
    }
  };
	spreadsheet_api.prototype._onEndOpen = function() {
		var t = this;
		if (this.openingEnd.bin && this.openingEnd.xlsx && this.openDocumentFromZip(t.wbModel, this.openingEnd.data)) {
			if (this.openingEnd.perfStart > 0) {
				let perfEnd = performance.now();
				AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onOpenDocument", perfEnd - this.openingEnd.perfStart), this);
				this.openingEnd.perfStart = 0;
			}
			//opening xlsx depends on getBinaryOtherTableGVar(called after Editor.bin)
			Asc.ReadDefTableStyles(t.wbModel);
			g_oIdCounter.Set_Load(false);
			AscCommon.checkCultureInfoFontPicker();
			AscCommonExcel.checkStylesNames(t.wbModel.CellStyles);
			t.FontLoader.LoadDocumentFonts(t.wbModel.generateFontMap2());

			// Какая-то непонятная заглушка, чтобы не падало в ipad
			if (t.isMobileVersion) {
				AscCommon.AscBrowser.isSafariMacOs = false;
				AscCommon.PasteElementsId.PASTE_ELEMENT_ID = "wrd_pastebin";
				AscCommon.PasteElementsId.ELEMENT_DISPAY_STYLE = "none";
			}
		}
	};
	spreadsheet_api.prototype._openOnClient = function() {
		if(this.isOpenOOXInBrowser) {
			return;
		}
		var t = this;
		if (this.openingEnd.xlsxStart) {
			return;
		}
		this.openingEnd.xlsxStart = true;
		var url = AscCommon.g_oDocumentUrls.getUrl('Editor.xlsx');
		if (url) {
			AscCommon.loadFileContent(url, function(httpRequest) {
				if (httpRequest && httpRequest.response) {
					t.openingEnd.xlsx = true;
					t.openingEnd.data = httpRequest.response;
					t._onEndOpen();
				} else {
					t.sendEvent('asc_onError', c_oAscError.ID.Unknown, c_oAscError.Level.Critical);
				}
			}, "arraybuffer");
		} else {
			t.openingEnd.xlsx = true;
			t._onEndOpen();
		}
	};

  spreadsheet_api.prototype._downloadAs = function(actionType, options, oAdditionalData, dataContainer, downloadType) {
    var fileType = options.fileType;

	if (this.isCloudSaveAsLocalToDrawingFormat(actionType, fileType)) {
	  var printPagesData, pdfPrinterMemory, t = this;
      this.wb._executeWithoutZoom(function () {
        printPagesData = t.wb.calcPagesPrint(options.advancedOptions);
        pdfPrinterMemory = t.wb.printSheets(printPagesData, null, options.advancedOptions).DocumentRenderer.Memory;
	  });
      this.localSaveToDrawingFormat(pdfPrinterMemory.GetBase64Memory(), fileType);
	  return true;
	}

    if (c_oAscFileType.PDF === fileType || c_oAscFileType.PDFA === fileType) {
      var printPagesData, pdfPrinterMemory, t = this;
      this.wb._executeWithoutZoom(function () {
      	t.wb.printPreviewState.advancedOptions = options.advancedOptions;
        printPagesData = t.wb.calcPagesPrint(options.advancedOptions);
        pdfPrinterMemory = t.wb.printSheets(printPagesData, null, options.advancedOptions).DocumentRenderer.Memory;
      	t.wb.printPreviewState.advancedOptions = null;
      });
      dataContainer.data = oAdditionalData["nobase64"] ? pdfPrinterMemory.GetData() : pdfPrinterMemory.GetBase64Memory();
    } else if (this.insertDocumentUrlsData) {
      var last = this.insertDocumentUrlsData.documents.shift();
      oAdditionalData['url'] = last.url;
      oAdditionalData['format'] = last.format;
      if (last.token) {
        oAdditionalData['tokenDownload'] = last.token;
        //remove to reduce message size
        oAdditionalData['tokenSession'] = undefined;
      }
      oAdditionalData['outputurls']= true;
      // ToDo select txt params
      oAdditionalData["codepage"] = AscCommon.c_oAscCodePageUtf8;
      dataContainer.data = last.data;
    } else if(this.isOpenOOXInBrowser && this.saveDocumentToZip) {
		var t = this;
		var title = this.documentTitle;
		AscCommonExcel.executeInR1C1Mode(false, function () {
			t.saveDocumentToZip(t.wb.model, AscCommon.c_oEditorId.Spreadsheet, function(data) {
				if (data) {
					if (c_oAscFileType.XLSX === fileType && !window.isCloudCryptoDownloadAs) {
						AscCommon.DownloadFileFromBytes(data, title, AscCommon.openXml.GetMimeType("xlsx"));
					} else {
						dataContainer.data = data;
						if (window.isCloudCryptoDownloadAs)
						{
							var sParamXml = ("<m_nCsvTxtEncoding>" + oAdditionalData["codepage"] + "</m_nCsvTxtEncoding>");
							sParamXml += ("<m_nCsvDelimiter>" + oAdditionalData["delimiter"] + "</m_nCsvDelimiter>");
							window["AscDesktopEditor"]["CryptoDownloadAs"](dataContainer.data, fileType, sParamXml);
							return true;
						}
						t._downloadAsUsingServer(actionType, options, oAdditionalData, dataContainer, downloadType);
						return;
					}
				} else {
					t.sendEvent("asc_onError", Asc.c_oAscError.ID.Unknown, Asc.c_oAscError.Level.NoCritical);
				}
				if (actionType)
				{
					t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, actionType);
				}
			});

		});
		return true;
	} else {
      var oBinaryFileWriter = new AscCommonExcel.BinaryFileWriter(this.wbModel);
      if (c_oAscFileType.CSV === fileType) {
        if (options.advancedOptions instanceof asc.asc_CTextOptions) {
          oAdditionalData["codepage"] = options.advancedOptions.asc_getCodePage();
          oAdditionalData["delimiter"] = options.advancedOptions.asc_getDelimiter();
          oAdditionalData["delimiterChar"] = options.advancedOptions.asc_getDelimiterChar();
        }
      }
      //перед записью подменяю topLeftCell на тот, который видим
        this.wb.executeWithCurrentTopLeftCell(function () {
          dataContainer.data = oBinaryFileWriter.Write(oAdditionalData["nobase64"]);
        });
    }

    if (window.isCloudCryptoDownloadAs) {
      var sParamXml = ("<m_nCsvTxtEncoding>" + oAdditionalData["codepage"] + "</m_nCsvTxtEncoding>");
      sParamXml += ("<m_nCsvDelimiter>" + oAdditionalData["delimiter"] + "</m_nCsvDelimiter>");
      window["AscDesktopEditor"]["CryptoDownloadAs"](dataContainer.data, fileType, sParamXml);
      return true;
    }
  };

	spreadsheet_api.prototype.asc_initPrintPreview = function (containerId, options) {
		var curElem = document.getElementById(containerId);
		if (curElem) {
			var isInitCanvas = false;
			var canvasId = containerId + "-canvas";
			var canvas = document.getElementById(canvasId);
			if (!canvas) {
				canvas = document.createElement('canvas');
				canvas.id = canvasId;
				canvas.width = curElem.clientWidth;
				canvas.height = curElem.clientHeight;
				//obj.canvas.style.width = AscCommon.AscBrowser.convertToRetinaValue(t.parentWidth) + "px";
				canvas.style.border = "1px solid";
				canvas.style.borderColor = "#" + this.wb.defaults.worksheetView.cells.defaultState.border.get_hex();
				curElem.appendChild(canvas);
				isInitCanvas = true;
			}

			if (!this.wb.printPreviewState.getCtx() || isInitCanvas) {
				this.wb.printPreviewState.setCtx(new asc.DrawingContext({
					canvas: canvas, units: 0/*px*/, fmgrGraphics: this.wb.fmgrGraphics, font: this.wb.m_oFont
				}));
			}
		}

		this.wb.printPreviewState.isDrawPrintPreview = true;
		this.wb.printPreviewState.init();
		var pages = this.wb.calcPagesPrint(options ? options.advancedOptions : null);
		this.wb.printPreviewState.setPages(pages);
		this.wb.printPreviewState.setAdvancedOptions(options && options.advancedOptions);
		this.wb.printPreviewState.isDrawPrintPreview = false;

		if (pages.arrPages.length) {
			this.asc_drawPrintPreview(0);
		}
		return pages.arrPages.length;
	};

	spreadsheet_api.prototype.asc_updatePrintPreview = function (options) {
		this.wb.printPreviewState.isDrawPrintPreview = true;
		var pages = this.wb.calcPagesPrint(options.advancedOptions);
		this.wb.printPreviewState.setPages(pages);
		this.wb.printPreviewState.setAdvancedOptions(options && options.advancedOptions);
		this.wb.printPreviewState.isDrawPrintPreview = false;
		return pages.arrPages.length;
	};

	spreadsheet_api.prototype.asc_drawPrintPreview = function (index, indexSheet) {
		if (this.wb.printPreviewState.isDrawPrintPreview) {
			return;
		}
		this.wb.printPreviewState.isDrawPrintPreview = true;
		if (indexSheet != null) {
			index = this.wb.printPreviewState.getIndexPageByIndexSheet(indexSheet);
		}
		if (index == null) {
			index = this.wb.printPreviewState.activePage;
		}
		this.wb.printPreviewState.setPage(index, true);
		this.wb.printSheetPrintPreview(index);
		var curPage = this.wb.printPreviewState.getPage(index);
		//возвращаю инфомарцию об активном листе, который печатаем
		var indexActiveWs = curPage && curPage.indexWorksheet;
		if (indexActiveWs === undefined) {
			indexActiveWs = this.wbModel.getActive();
		}
		this.handlers.trigger("asc_onPrintPreviewSheetChanged", indexActiveWs);
		this.handlers.trigger("asc_onPrintPreviewPageChanged", index);
		this.wb.printPreviewState.isDrawPrintPreview = false;
	};

	spreadsheet_api.prototype.asc_closePrintPreview = function () {
		this.wb.printPreviewState.clean(true);
	};

  spreadsheet_api.prototype.processSavedFile             = function(url, downloadType, filetype)
  {
    if (this.insertDocumentUrlsData && this.insertDocumentUrlsData.convertCallback)
    {
      this.insertDocumentUrlsData.convertCallback(this, url);
    }
    else
    {
      AscCommon.baseEditorsApi.prototype.processSavedFile.call(this, url, downloadType, filetype);
    }
  };

  spreadsheet_api.prototype.asc_isDocumentModified = function() {
    if (!this.canSave || this.asc_getCellEditMode()) {
      // Пока идет сохранение или редактирование ячейки, мы не закрываем документ
      return true;
    } else if (History && History.Have_Changes) {
      return History.Have_Changes();
    }
    return false;
  };
  spreadsheet_api.prototype.isDocumentModified = function() {
    return this.asc_isDocumentModified();
  };

  // Actions and callbacks interface

  /*
   * asc_onStartAction			(type, id)
   * asc_onEndAction				(type, id)
   * asc_onInitEditorFonts		(gui_fonts)
   * asc_onInitEditorStyles		(gui_styles)
   * asc_onOpenDocumentProgress	(AscCommon.COpenProgress)
   * asc_onAdvancedOptions		(c_oAscAdvancedOptionsID, asc_CAdvancedOptions)			- эвент на получение дополнительных опций (открытие/сохранение CSV)
   * asc_onError				(c_oAscError.ID, c_oAscError.Level)						- эвент об ошибке
   * asc_onEditCell				(Asc.c_oAscCellEditorState)								- эвент на редактирование ячейки с состоянием (переходами из формулы и обратно)
   * asc_onEditorSelectionChanged	(CellXfs)											- эвент на смену информации о выделении в редакторе ячейки
   * asc_onSelectionChanged		(asc_CCellInfo)										- эвент на смену информации о выделении
   * asc_onSelectionNameChanged	(sName)												- эвент на смену имени выделения (Id-ячейки, число выделенных столбцов/строк, имя диаграммы и др.)
   * asc_onSelection
   *
   * Changed	(asc_CSelectionMathInfo)							- эвент на смену математической информации о выделении
   * asc_onZoomChanged			(zoom)
   * asc_onSheetsChanged			()													- эвент на обновление списка листов
   * asc_onActiveSheetChanged		(indexActiveSheet)									- эвент на обновление активного листа
   * asc_onCanUndoChanged			(bCanUndo)											- эвент на обновление возможности undo
   * asc_onCanRedoChanged			(bCanRedo)											- эвент на обновление возможности redo
   * asc_onSaveUrl				(sUrl, callback(hasError))							- эвент на сохранение файла на сервер по url
   * asc_onDocumentModifiedChanged(bIsModified)										- эвент на обновление статуса "изменен ли файл"
   * asc_onMouseMove				(asc_CMouseMoveData)								- эвент на наведение мышкой на гиперлинк или комментарий
   * asc_onHyperlinkClick			(sUrl)												- эвент на нажатие гиперлинка
   * asc_onCoAuthoringDisconnect	()													- эвент об отключении от сервера без попытки reconnect
   * asc_onSelectionRangeChanged	(selectRange)										- эвент о выборе диапазона для диаграммы (после нажатия кнопки выбора)
   * asc_onRenameCellTextEnd		(countCellsFind, countCellsReplace)					- эвент об окончании замены текста в ячейках (мы не можем сразу прислать ответ)
   * asc_onWorkbookLocked			(result)											- эвент залочена ли работа с листами или нет
   * asc_onWorksheetLocked		(index, result)										- эвент залочен ли лист или нет
   * asc_onGetEditorPermissions	(permission)										- эвент о правах редактора
   * asc_onStopFormatPainter		()													- эвент об окончании форматирования по образцу
   * asc_onUpdateSheetViewSettings	()													- эвент об обновлении свойств листа (закрепленная область, показывать сетку/заголовки)
   * asc_onUpdateTabColor			(index)												- эвент об обновлении цвета иконки листа
   * asc_onDocumentCanSaveChanged	(bIsCanSave)										- эвент об обновлении статуса "можно ли сохранять файл"
   * asc_onDocumentUpdateVersion	(callback)											- эвент о том, что файл собрался и не может больше редактироваться
   * asc_onContextMenu			(event)												- эвент на контекстное меню
   * asc_onDocumentContentReady ()                        - эвент об окончании загрузки документа
   * asc_onFilterInfo	        (countFilter, countRecords)								- send count filtered and all records
   * asc_onLockDocumentProps/asc_onUnLockDocumentProps    - эвент о том, что залочены опции layout
   * asc_onUpdateDocumentProps                            - эвент о том, что необходимо обновить данные во вкладке layout
   * asc_onLockPrintArea/asc_onUnLockPrintArea            - эвент о локе в меню опции print area во вкладке layout
   * asc_onPrintPreviewSheetChanged                 - эвент о смене печатаемого листа при предварительной печати
   * asc_onPrintPreviewPageChanged                  - эвент о смене печатаемой страницы при предварительной печати
   */

  spreadsheet_api.prototype.asc_registerCallback = function(name, callback, replaceOldCallback) {
    this.handlers.add(name, callback, replaceOldCallback);
  };

  spreadsheet_api.prototype.asc_unregisterCallback = function(name, callback) {
    this.handlers.remove(name, callback);
  };

  spreadsheet_api.prototype.asc_SetDocumentPlaceChangedEnabled = function(val) {
    this.wb.setDocumentPlaceChangedEnabled(val);
  };

  spreadsheet_api.prototype.asc_SetFastCollaborative = function(bFast) {
    if (this.collaborativeEditing) {
      AscCommon.CollaborativeEditing.Set_Fast(bFast);
      this.collaborativeEditing.setFast(bFast);
    }
  };

	spreadsheet_api.prototype.asc_setThumbnailStylesSizes = function (width, height) {
		this.styleThumbnailWidth = width;
		this.styleThumbnailHeight = height;
	};

  // Посылает эвент о том, что обновились листы
  spreadsheet_api.prototype.sheetsChanged = function() {
    this.handlers.trigger("asc_onSheetsChanged");
  };

  spreadsheet_api.prototype.asyncFontsDocumentStartLoaded = function(blockType) {
    this.OpenDocumentProgress.Type = c_oAscAsyncAction.LoadDocumentFonts;
    this.OpenDocumentProgress.FontsCount = this.FontLoader.fonts_loading.length;
    this.OpenDocumentProgress.CurrentFont = 0;
    this.sync_StartAction(undefined === blockType ? c_oAscAsyncActionType.BlockInteraction : blockType, c_oAscAsyncAction.LoadDocumentFonts);
  };

  spreadsheet_api.prototype.asyncFontsDocumentEndLoaded = function(blockType) {
    this.sync_EndAction(undefined === blockType ? c_oAscAsyncActionType.BlockInteraction : blockType, c_oAscAsyncAction.LoadDocumentFonts);

    if (this.asyncMethodCallback !== undefined) {
      this.asyncMethodCallback();
      this.asyncMethodCallback = undefined;
    } else {
      // Шрифты загрузились, возможно стоит подождать совместное редактирование
      this.FontLoadWaitComplete = true;
        this._openDocumentEndCallback();
        if (this.fAfterLoad) {
          this.fAfterLoad();
        }
      }
  };

  spreadsheet_api.prototype.asyncFontEndLoaded = function(font) {
    this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);
  };

  spreadsheet_api.prototype._loadFonts = function(fonts, callback) {
    if (window["NATIVE_EDITOR_ENJINE"]) {
      return callback.call(this);
    }
    this.asyncMethodCallback = callback;
    var arrLoadFonts = [];
    for (var i in fonts)
      arrLoadFonts.push(new AscFonts.CFont(i, 0, "", 0));
    AscFonts.FontPickerByCharacter.extendFonts(arrLoadFonts);
    this.FontLoader.LoadDocumentFonts2(arrLoadFonts);
  };

  spreadsheet_api.prototype.openDocument = function(file) {
	//todo native.js -> openDocument
	this.openingEnd.perfStart = performance.now();
	if (file.changes && this.VersionHistory) {
  		this.VersionHistory.changes = file.changes;
	}
	this.isOpenOOXInBrowser = this["asc_isSupportFeature"]("ooxml") && AscCommon.checkOOXMLSignature(file.data);
	if (this.isOpenOOXInBrowser) {
		this.openOOXInBrowserZip = file.data;
	}
	this._openDocument(file.data);
	this._openOnClient();
  };

	spreadsheet_api.prototype.asc_CloseFile = function()
	{
		History.Clear();
		g_oIdCounter.Clear();
		g_oTableId.Clear();
		AscCommonExcel.g_StyleCache.Clear();
		AscCommon.CollaborativeEditing.Clear();
		AscCommon.g_oDocumentUrls.Clear();
		this.openingEnd = {bin: false, xlsxStart: false, xlsx: false, data: null};
		this.isApplyChangesOnOpenEnabled = true;
		this.isDocumentLoadComplete = false;
        this.turnOffSpecialModes();

		//удаляю весь handlersList, добавленный при инициализации wbView
		//потому что старый при открытии использовать нельзя(в случае с истрией версий при повторном открытии файла там остаются старые функции от предыдущего workbookview)
		//по идее нужно делать его полное зануление, а при открытии создавать заново. но есть функции, которые
		//добавляются в интерфейсе и в случае с историей версий заново не добавляются
		this.wb.removeHandlersList();
		if (this.isEditOleMode) {
			this.wb.removeEventListeners();
			var cellEditor = this.wb.cellEditor;
			if (this.wb.cellEditor) {
				cellEditor.removeEventListeners();
			}
		}

		if (this.wbModel.DrawingDocument) {
			this.wbModel.DrawingDocument.CloseFile();
		}

        if(this.wb.MobileTouchManager) {
            this.wb.MobileTouchManager.Destroy();
        }
	};

	spreadsheet_api.prototype.openDocumentFromZip = function (wb, data) {
		var t = this;
		if (!data) {
			return true;
		}

		var openXml = AscCommon.openXml;
		var StaxParser = AscCommon.StaxParser;
		var pivotCaches = {};
		var xmlParserContext = new AscCommon.XmlParserContext();
		xmlParserContext.DrawingDocument = this.wbModel.DrawingDocument;
		var initOpenManager = xmlParserContext.InitOpenManager = AscCommonExcel.InitOpenManager ? new AscCommonExcel.InitOpenManager(null, wb) : null;
		var wbPart = null;
		var wbXml = null;
		let jsZlib = new AscCommon.ZLib();
		if (!jsZlib.open(data)) {
			return false;
		}
		xmlParserContext.zip = jsZlib;

		//check fonts inside
		AscFonts.IsCheckSymbols = true;

		AscCommonExcel.executeInR1C1Mode(false, function () {
			var doc = new openXml.OpenXmlPackage(jsZlib, null);
			var reader, i, j;

			//core
			var coreXmlPart = doc.getPartByRelationshipType(openXml.Types.coreFileProperties.relationType);
			if (coreXmlPart) {
				var contentCore = coreXmlPart.getDocumentContent();
				if (contentCore) {
					wb.Core = new AscCommon.CCore();
					reader = new StaxParser(contentCore, coreXmlPart, xmlParserContext);
					wb.Core.fromXml(reader, true);
				}
			}

			//app
			var appXmlPart = doc.getPartByRelationshipType(openXml.Types.extendedFileProperties.relationType);
			if (appXmlPart) {
				var contentApp = appXmlPart.getDocumentContent();
				if (contentApp) {
					wb.App = new AscCommon.CApp();
					reader = new StaxParser(contentApp, appXmlPart, xmlParserContext);
					wb.App.fromXml(reader, true);
				}
			}

			//workbook
			wbPart = doc.getPartByRelationshipType(openXml.Types.workbook.relationType);
			if (wbPart) {
				var contentWorkbook = wbPart.getDocumentContent();
				wbXml = new AscCommonExcel.CT_Workbook(wb);
				reader = new StaxParser(contentWorkbook, wbPart, xmlParserContext);
				wbXml.fromXml(reader);
			}


			if (t.isOpenOOXInBrowser) {

				//theme
				var workbookThemePart = wbPart.getPartByRelationshipType(openXml.Types.theme.relationType);
				if (workbookThemePart) {
					var contentWorkbookTheme = workbookThemePart.getDocumentContent();
					var oTheme = new AscFormat.CTheme();
					reader = new StaxParser(contentWorkbookTheme, workbookThemePart, xmlParserContext);
					oTheme.fromXml(reader, true);
					wb.theme = oTheme;
				}
				xmlParserContext.InitOpenManager.initSchemeAndTheme(wb);

				//TODO oMediaArray

				//external reference
				if (wbXml && wbXml.externalReferences) {
					wbXml.externalReferences.forEach(function (externalReference) {
						if (null !== externalReference) {
							var externalWorkbookPart = wbPart.getPartById(externalReference);
							if (externalWorkbookPart) {
								var contentExternalWorkbook = externalWorkbookPart.getDocumentContent();
								if (contentExternalWorkbook) {
									var oExternalReference = new AscCommonExcel.CT_ExternalReference(wb);
									var reader = new StaxParser(contentExternalWorkbook, externalWorkbookPart, xmlParserContext);
									oExternalReference.fromXml(reader);

									if (oExternalReference.val) {
										if (oExternalReference.val.externalBook) {
											var relationship = externalWorkbookPart.getRelationship(oExternalReference.val.externalBook.Id);
											//подменяем id на target
											if (relationship && relationship.targetFullName) {
												oExternalReference.val.externalBook.Id = AscCommonExcel.decodeXmlPath(relationship.targetFullName);
											}
											wb.externalReferences.push(oExternalReference.val.externalBook);

										}
									}
								}
							}
						}
					});
				}

				//extLxt(slicercache inside)
				if (wbXml && wbXml.extLst) {
					wbXml.extLst.forEach(function (ext) {
						if (ext.slicerCachesIds) {
							ext.slicerCachesIds.forEach(function (slicerCacheId) {
								if (null !== slicerCacheId) {
									var slicerCacheWorkbookPart = wbPart.getPartById(slicerCacheId);
									if (slicerCacheWorkbookPart) {
										var contentSlicerCache = slicerCacheWorkbookPart.getDocumentContent();
										if (contentSlicerCache) {
											var oSlicerCacheDefinition = new Asc.CT_slicerCacheDefinition();
											var reader = new StaxParser(contentSlicerCache, slicerCacheWorkbookPart, xmlParserContext);
											oSlicerCacheDefinition.fromXml(reader);

											xmlParserContext.InitOpenManager.oReadResult.slicerCaches[oSlicerCacheDefinition.name] = oSlicerCacheDefinition;
										}
									}
								}
							});
						}
					});
				}

				//not ext slicer caches
				if (wbXml && wbXml.slicerCachesIds) {
					wbXml.slicerCachesIds.forEach(function (slicerCacheId) {
						if (null !== slicerCacheId) {
							var slicerCacheWorkbookPart = wbPart.getPartById(slicerCacheId);
							if (slicerCacheWorkbookPart) {
								var contentSlicerCache = slicerCacheWorkbookPart.getDocumentContent();
								if (contentSlicerCache) {
									var oSlicerCacheDefinition = new Asc.CT_slicerCacheDefinition();
									var reader = new StaxParser(contentSlicerCache, slicerCacheWorkbookPart, xmlParserContext);
									oSlicerCacheDefinition.fromXml(reader);

									xmlParserContext.InitOpenManager.oReadResult.slicerCaches[oSlicerCacheDefinition.name] = oSlicerCacheDefinition;
								}
							}
						}
					});
				}

				//connection
				//пока читаю в строку connections. в serialize сейчас аналогично не парсим структуру, а храним в виде массива байтов
				var connectionsPart = wbPart.getPartByRelationshipType(openXml.Types.connections.relationType);
				if (connectionsPart) {
					wb.connections = connectionsPart.getDocumentContent();
				}

				//styles
				var dxfs = [];
				var aCellXfs = [];
				var stylesPart = wbPart.getPartByRelationshipType(openXml.Types.styles.relationType);
				if (stylesPart) {
					var contentStyles = stylesPart.getDocumentContent();
					if (contentStyles) {
						var styleSheet = new AscCommonExcel.CT_Stylesheet(new Asc.CTableStyles());
						reader = new StaxParser(contentStyles, stylesPart, xmlParserContext);
						styleSheet.fromXml(reader);


						var oStyleObject = {
							aBorders: styleSheet.borders,
							aFills: styleSheet.fills,
							aFonts: styleSheet.fonts,
							oNumFmts: styleSheet.numFmts,
							aCellStyleXfs: styleSheet.cellStyleXfs,
							aCellXfs: styleSheet.cellXfs,
							aDxfs: styleSheet.dxfs,
							aExtDxfs: styleSheet.aExtDxfs,
							aCellStyles: styleSheet.cellStyles,
							oCustomTableStyles: styleSheet.tableStyles.CustomStyles,
							oCustomSlicerStyles: styleSheet.oCustomSlicerStyles
						};

						xmlParserContext.InitOpenManager.InitStyleManager(oStyleObject, aCellXfs);
						dxfs = oStyleObject.aDxfs;
						wb.oNumFmtsOpen = oStyleObject.oNumFmts;
					}
				}
				xmlParserContext.InitOpenManager.aCellXfs = aCellXfs;
				xmlParserContext.InitOpenManager.Dxfs = dxfs;

				//jsaProject
				var jsaProjectPart = wbPart.getPartByRelationshipType(openXml.Types.jsaProject.relationType);
				if (jsaProjectPart) {
					var contentJsaProject = jsaProjectPart.getDocumentContent();
					if (contentJsaProject) {
						xmlParserContext.InitOpenManager.oReadResult.macros = contentJsaProject;
					}
				}

				//vbaProject
				var vbaProjectPart = wbPart.getPartByRelationshipType(openXml.Types.vbaProject.relationType);
				if (vbaProjectPart) {
					var contentVbaProject = vbaProjectPart.getDocumentContent(true);
					if (contentVbaProject) {
						xmlParserContext.InitOpenManager.oReadResult.vbaMacros = new Uint8Array(contentVbaProject);
					}
				}

				//person list
				var personListPart = wbPart.getPartByRelationshipType(openXml.Types.person.relationType);
				var personList;
				if (personListPart) {
					var contentPersonList = personListPart.getDocumentContent();
					if (contentPersonList) {
						personList = new AscCommonExcel.CT_PersonList();
						reader = new StaxParser(contentPersonList, personListPart, xmlParserContext);
						personList.fromXml(reader);
					}
				}

				//wb comments
				//лежит в виде бинарника, читаем через serialize
				var workbookComment = wbPart.getPartByRelationshipType(openXml.Types.workbookComment.relationType);
				if (workbookComment) {
					var contentWorkbookComment = workbookComment.getDocumentContent(true);
					if (contentWorkbookComment) {
						AscCommonExcel.ReadWbComments(wb, contentWorkbookComment, xmlParserContext.InitOpenManager);
					}
				}
			}

			//pivotCaches
			if (wbXml && wbXml.pivotCaches) {
				wbXml.pivotCaches.forEach(function (wbPivotCacheXml) {
					var pivotTableCacheDefinitionPart;
					if (null !== wbPivotCacheXml.cacheId && null !== wbPivotCacheXml.id) {
						pivotTableCacheDefinitionPart = wbPart.getPartById(wbPivotCacheXml.id);
						var contentCacheDefinition = pivotTableCacheDefinitionPart.getDocumentContent();
						if (contentCacheDefinition) {
							var pivotTableCacheDefinition = new Asc.CT_PivotCacheDefinition();

							new openXml.SaxParserBase().parse(contentCacheDefinition, pivotTableCacheDefinition);

							if (pivotTableCacheDefinition.isValidCacheSource()) {
								pivotCaches[wbPivotCacheXml.cacheId] = pivotTableCacheDefinition;
								if (pivotTableCacheDefinition.id) {
									var partPivotTableCacheRecords = pivotTableCacheDefinitionPart.getPartById(pivotTableCacheDefinition.id);
									var contentCacheRecords = partPivotTableCacheRecords.getDocumentContent();
									if (contentCacheRecords) {
										var pivotTableCacheRecords = new Asc.CT_PivotCacheRecords();
										new openXml.SaxParserBase().parse(contentCacheRecords, pivotTableCacheRecords);
										pivotTableCacheDefinition.cacheRecords = pivotTableCacheRecords;
									}
								}
							}
						}
					}
				});
			}

			//sharedString
			if (t.isOpenOOXInBrowser) {
				var sharedStringPart = wbPart.getPartByRelationshipType(openXml.Types.sharedStringTable.relationType);
				if (sharedStringPart) {
					var contentSharedStrings = sharedStringPart.getDocumentContent();
					if (contentSharedStrings) {
						var sharedStrings = new AscCommonExcel.CT_SharedStrings();
						reader = new StaxParser(contentSharedStrings, sharedStringPart, xmlParserContext);
						sharedStrings.fromXml(reader);
					}
				}
			}

			//TODO CalcChain - из бинарника не читается, и не пишется в бинарник. реализовать позже

			//Custom xml
			if (t.isOpenOOXInBrowser) {
				//папка customXml, в неё лежат item[n].xml, itemProps[n].xml + rels

				//в Content_Types пишется только ссылка на itemProps в слудующем виде:
				//<Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>

				//rels(которые внутри customXml) лежит ссылка на itemProps  в следующем виде:
				//<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship  Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>

				//workbook.xml.rels лежит ссылка на item  в следующем виде:
				//<Relationship  Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/item1.xml"/>

				//TODO проверить когда несколько ссылок на customXml
				var customXmlParts = wbPart.getPartsByRelationshipType(openXml.Types.customXml.relationType);
				if (customXmlParts) {
					for (i = 0; i < customXmlParts.length; i++) {
						var customXmlPart = customXmlParts[i];
						var customXml = customXmlPart.getDocumentContent("string");
						var customXmlPropsPart = customXmlPart.getPartByRelationshipType(openXml.Types.customXmlProps.relationType);
						var customXmlProps = customXmlPropsPart && customXmlPropsPart.getDocumentContent("string");

						//в бинарник не будем писать, для совместимости оставляю поля, добавляю ещё новые
						var custom = {Uri: [], ItemId: null, Content: null, item: customXml, itemProps: customXmlProps};
						if (!wb.customXmls) {
							wb.customXmls = [];
						}
						wb.customXmls.push(custom);
					}
				}
			}

			//sheets
			var wsParts = [];
			if (t.isOpenOOXInBrowser && wbXml && wbXml.sheets) {
				//вначале беру все листы, потом запрашиваю контент каждого из них.
				//связано с проблемой внтури парсера, на примере файла Read_Only_part_of_lists.xlsx
				wbXml.sheets.forEach(function (wbSheetXml) {
					if (null !== wbSheetXml.id && wbSheetXml.name) {
						var wsPart = wbPart.getPartById(wbSheetXml.id);
						wsParts.push({wsPart: wsPart, id: wbSheetXml.id, name: wbSheetXml.name, bHidden: wbSheetXml.bHidden, sheetId: wbSheetXml.sheetId});
					}
				});

				wsParts.forEach(function (wbSheetXml) {
					if (null !== wbSheetXml.id && wbSheetXml.name) {
						var wsPart = wbSheetXml.wsPart;
						var contentSheetXml = wsPart && wsPart.getDocumentContent();
						if (contentSheetXml) {
							var ws = new AscCommonExcel.Worksheet(wb, wb.aWorksheets.length);
							ws.sName = wbSheetXml.name;
							if (null !== wbSheetXml.bHidden) {
								ws.bHidden = wbSheetXml.bHidden;
							}
							//var wsView = new AscCommonExcel.asc_CSheetViewSettings();
							//wsView.pane = new AscCommonExcel.asc_CPane();
							//ws.sheetViews.push(wsView);
							if (contentSheetXml) {
								var reader = new StaxParser(contentSheetXml, wsPart, xmlParserContext);
								ws.fromXml(reader);
							}

							xmlParserContext.InitOpenManager.putSheetAfterRead(wb, ws);
							xmlParserContext.InitOpenManager.oReadResult.sheetIds[wbSheetXml.sheetId] = ws;

							var drawingPart = wsPart.getPartById(xmlParserContext.drawingId);
							if (drawingPart) {
								var drawingWS = new AscCommonExcel.CT_DrawingWS(ws);
								var contentDrawing = drawingPart.getDocumentContent();
								reader = new StaxParser(contentDrawing, drawingPart, xmlParserContext);
								drawingWS.fromXml(reader);
								let aSpTree = [];
								for(let nDrawing = 0; nDrawing < ws.Drawings.length; ++nDrawing) {
									aSpTree.push(ws.Drawings[nDrawing].graphicObject);
								}
								reader.context.assignConnectors(aSpTree);
							}
							if (wsPart) {
								//pivot
								var pivotParts = wsPart.getPartsByRelationshipType(openXml.Types.pivotTable.relationType);
								for (i = 0; i < pivotParts.length; ++i) {
									var contentPivotTable = pivotParts[i].getDocumentContent();
									var pivotTable = new Asc.CT_pivotTableDefinition(true);
									new openXml.SaxParserBase().parse(contentPivotTable, pivotTable);
									var cacheDefinition = pivotCaches[pivotTable.cacheId];
									if (cacheDefinition) {
										pivotTable.cacheDefinition = cacheDefinition;
										ws.insertPivotTable(pivotTable);
									}
								}

								//tables
								var tableParts = wsPart.getPartsByRelationshipType(openXml.Types.tableDefinition.relationType);
								for (i = 0; i < tableParts.length; ++i) {
									var contentTable = tableParts[i].getDocumentContent();
									var oNewTable = ws.createTablePart();
									reader = new StaxParser(contentTable, oNewTable, xmlParserContext);
									oNewTable.fromXml(reader);

									var queryTables = tableParts[i].getPartsByRelationshipType(openXml.Types.queryTable.relationType);
									for (j = 0; j < queryTables.length; ++j) {
										var contentQueryTable = queryTables[j].getDocumentContent();
										var oNewQueryTable = new AscCommonExcel.QueryTable();
										reader = new StaxParser(contentQueryTable, oNewQueryTable, xmlParserContext);
										oNewQueryTable.fromXml(reader);
										oNewTable.QueryTable = oNewQueryTable;
									}

									if (null != oNewTable.Ref && null != oNewTable.DisplayName) {
										ws.workbook.dependencyFormulas.addTableName(ws, oNewTable, true);
									}
									ws.TableParts.push(oNewTable);
								}

								//namedSheetViews
								var namedSheetViews = wsPart.getPartsByRelationshipType(openXml.Types.namedSheetViews.relationType);
								for (i = 0; i < namedSheetViews.length; ++i) {
									var contentSheetView = namedSheetViews[i].getDocumentContent();
									var namedSheetView = new Asc.CT_NamedSheetViews();
									reader = new StaxParser(contentSheetView, namedSheetView, xmlParserContext);
									namedSheetView.fromXml(reader);
									//связь с таблицыми по id осуществляется через tableIdOpen, который потом в методе initPostOpen преобразуется в tableId
									ws.aNamedSheetViews = namedSheetView.namedSheetView;
								}

								//slicers
								var slicers = wsPart.getPartsByRelationshipType(openXml.Types.slicers.relationType);
								for (i = 0; i < slicers.length; ++i) {
									var contentSlicers = slicers[i].getDocumentContent();
									var oSlicers = new Asc.CT_slicers(ws);
									oSlicers.slicer = ws.aSlicers;
									reader = new StaxParser(contentSlicers, oSlicers, xmlParserContext);
									oSlicers.fromXml(reader);
								}

								//COMMENTS
								//буду читать по формату, далее преобразовывать
								var comments, pThreadedComments;
								var commentsFile = wsPart.getPartsByRelationshipType(openXml.Types.worksheetComments.relationType);
								for (i = 0; i < commentsFile.length; ++i) {
									var contentComment = commentsFile[i].getDocumentContent();
									comments = new AscCommonExcel.CT_CComments();
									reader = new StaxParser(contentComment, comments, xmlParserContext);
									comments.fromXml(reader);
								}

								var threadedCommentsFile = wsPart.getPartsByRelationshipType(openXml.Types.threadedComment.relationType);
								for (i = 0; i < threadedCommentsFile.length; ++i) {
									var threadedComment = threadedCommentsFile[i].getDocumentContent();
									pThreadedComments = new AscCommonExcel.CT_CThreadedComments();
									reader = new StaxParser(threadedComment, pThreadedComments, xmlParserContext);
									pThreadedComments.fromXml(reader);
								}

								AscCommonExcel.PrepareComments(ws, xmlParserContext, comments, pThreadedComments, personList);
							}
						}
					}
				});
			} else if(wbXml && wbXml.sheets) {
				wsParts = [];

				//вначале беру все листы, потом запрашиваю контент каждого из них.
				//связано с проблемой внтури парсера, на примере файла Read_Only_part_of_lists.xlsx
				wbXml.sheets.forEach(function (wbSheetXml) {
					if (null !== wbSheetXml.id && wbSheetXml.name) {
						var wsPart = wbPart.getPartById(wbSheetXml.id);
						wsParts.push({wsPart: wsPart, id: wbSheetXml.id, name: wbSheetXml.name, bHidden: wbSheetXml.bHidden});
					}
				});

				wsParts.forEach(function(wbSheetXml, wsIndex) {
					var ws = t.wbModel.getWorksheet(wsIndex);
					if (ws && null !== wbSheetXml.id && wbSheetXml.name) {
						var wsPart = wbSheetXml.wsPart;
						if (wsPart) {
							//pivot
							var pivotParts = wsPart.getPartsByRelationshipType(openXml.Types.pivotTable.relationType);
							for (i = 0; i < pivotParts.length; ++i) {
								var contentPivotTable = pivotParts[i].getDocumentContent();
								var pivotTable = new Asc.CT_pivotTableDefinition(true);
								new openXml.SaxParserBase().parse(contentPivotTable, pivotTable);
								var cacheDefinition = pivotCaches[pivotTable.cacheId];
								if (cacheDefinition) {
									pivotTable.cacheDefinition = cacheDefinition;
									ws.insertPivotTable(pivotTable);
								}
							}
						}
					}
				});
			}

			if (t.isOpenOOXInBrowser) {
				//defined names
				if (wbXml && wbXml.newDefinedNames) {
					xmlParserContext.InitOpenManager.oReadResult.defNames = wbXml.newDefinedNames;
					xmlParserContext.InitOpenManager.PostLoadPrepareDefNames(wb);
				}
			}

			var readSheetDataExternal = function (bNoBuildDep) {
				for (var i = 0; i < xmlParserContext.InitOpenManager.oReadResult.sheetData.length; ++i) {
					var sheetDataElem = xmlParserContext.InitOpenManager.oReadResult.sheetData[i];
					var ws = sheetDataElem.ws;

					var tmp = {
						pos: null,
						len: null,
						bNoBuildDep: bNoBuildDep,
						ws: ws,
						row: new AscCommonExcel.Row(ws),
						cell: new AscCommonExcel.Cell(ws),
						formula: new AscCommonExcel.OpenFormula(),
						sharedFormulas: {},
						prevFormulas: {},
						siFormulas: {},
						prevRow: -1,
						prevCol: -1,
						formulaArray: []
					};


					var sheetData = new AscCommonExcel.CT_SheetData();
					xmlParserContext.InitOpenManager.tmp = tmp;

					sheetDataElem.reader.setState(sheetDataElem.state);
					//TODO пересмотреть фунцию fromXml
					sheetData.fromXml2(sheetDataElem.reader);

					if (!bNoBuildDep) {
						//TODO возможно стоит делать это в worksheet после полного чтения
						//***array-formula***
						//добавление ко всем ячейкам массива головной формулы
						for (var j = 0; j < tmp.formulaArray.length; j++) {
							var curFormula = tmp.formulaArray[j];
							var ref = curFormula.ref;
							if (ref) {
								var rangeFormulaArray = tmp.ws.getRange3(ref.r1, ref.c1, ref.r2, ref.c2);
								rangeFormulaArray._foreach(function (cell) {
									cell.setFormulaInternal(curFormula);
									if (curFormula.ca || cell.isNullTextString()) {
										tmp.ws.workbook.dependencyFormulas.addToChangedCell(cell);
									}
								});
							}
						}
						for (var nCol in tmp.prevFormulas) {
							if (tmp.prevFormulas.hasOwnProperty(nCol)) {
								var prevFormula = tmp.prevFormulas[nCol];
								if (!tmp.siFormulas[prevFormula.parsed.getListenerId()]) {
									prevFormula.parsed.buildDependencies();
								}
							}
						}
						for (var listenerId in tmp.siFormulas) {
							if (tmp.siFormulas.hasOwnProperty(listenerId)) {
								tmp.siFormulas[listenerId].buildDependencies();
							}
						}
					}
				}
			};

			//TODO общий код с serialize
			//ReadSheetDataExternal
			if (t.isOpenOOXInBrowser) {
				if (!initOpenManager.copyPasteObj.isCopyPaste || initOpenManager.copyPasteObj.selectAllSheet) {
					readSheetDataExternal(false);
					if (!initOpenManager.copyPasteObj.isCopyPaste) {
						initOpenManager.PostLoadPrepare(wb);
					}
					wb.init(initOpenManager.oReadResult.tableCustomFunc, initOpenManager.oReadResult.tableIds, initOpenManager.oReadResult.sheetIds, false, true);
				} else {
					readSheetDataExternal(true);
					if (Asc["editor"] && Asc["editor"].wb) {
						wb.init(initOpenManager.oReadResult.tableCustomFunc, initOpenManager.oReadResult.tableIds, initOpenManager.oReadResult.sheetIds, true);
					}
				}
			}

			if (t.isOpenOOXInBrowser) {
				initOpenManager.readDefStyles(wb, wb.CellStyles.DefaultStyles);
			}

			wb.initPostOpenZip(pivotCaches, xmlParserContext);
		});

		AscFonts.IsCheckSymbols = false;

		jsZlib.close();
		//clean up
		openXml.SaxParserDataTransfer = {};
		return true;
	};


	spreadsheet_api.prototype.openDocumentFromZip2 = function (wb, data) {
		//TODO зачитать sharedStrings
		if (!data || !this["asc_isSupportFeature"]("ooxml")) {
			return null;
		}

		var openXml = AscCommon.openXml;
		var StaxParser = AscCommon.StaxParser;
		var res = [];

		AscFormat.ExecuteNoHistory(function() {
			var xmlParserContext = new AscCommon.XmlParserContext();
			xmlParserContext.DrawingDocument = this.wbModel.DrawingDocument;
			xmlParserContext.InitOpenManager = AscCommonExcel.InitOpenManager ? new AscCommonExcel.InitOpenManager(null, wb) : null;
			if (!xmlParserContext.InitOpenManager) {
				xmlParserContext.InitOpenManager = {};
			}
			xmlParserContext.InitOpenManager.aCellXfs = [];
			xmlParserContext.InitOpenManager.Dxfs = [];
			
			var wbPart = null;
			var wbXml = null;
			if (!window.nativeZlibEngine || !window.nativeZlibEngine.open(data)) {
				return false;
			}
			xmlParserContext.zip = window.nativeZlibEngine;

			var doc = new openXml.OpenXmlPackage(window.nativeZlibEngine, null);

			//workbook
			wbPart = doc.getPartByRelationshipType(openXml.Types.workbook.relationType);
			if (wbPart) {
				var contentWorkbook = wbPart.getDocumentContent();
				AscCommonExcel.executeInR1C1Mode(false, function () {
					wbXml = new AscCommonExcel.CT_Workbook(wb);
					var reader = new StaxParser(contentWorkbook, wbPart, xmlParserContext);
					wbXml.fromXml(reader, {"sheets" : 1});
				});
			}

			//sharedString
			var sharedStringPart = wbPart.getPartByRelationshipType(openXml.Types.sharedStringTable.relationType);
			if (sharedStringPart) {
				var contentSharedStrings = sharedStringPart.getDocumentContent();
				if (contentSharedStrings) {
					var sharedStrings = new AscCommonExcel.CT_SharedStrings();
					var reader = new StaxParser(contentSharedStrings, sharedStringPart, xmlParserContext);
					sharedStrings.fromXml(reader);
				}
			}

			//sheets
			if (wbXml && wbXml.sheets) {
				var wsParts = [];

				wbXml.sheets.forEach(function (wbSheetXml) {
					//TODO нужно отсеять только необходимые листы
					if (null !== wbSheetXml.id && wbSheetXml.name) {
						var wsPart = wbPart.getPartById(wbSheetXml.id);
						wsParts.push({wsPart: wsPart, id: wbSheetXml.id, name: wbSheetXml.name, bHidden: wbSheetXml.bHidden, sheetId: wbSheetXml.sheetId});
					}
				});

				wsParts.forEach(function (wbSheetXml) {
					//TODO нужно отсеять только необходимые дипапазоны
					if (null !== wbSheetXml.id && wbSheetXml.name) {
						var wsPart = wbSheetXml.wsPart;
						var contentSheetXml = wsPart && wsPart.getDocumentContent();
						if (contentSheetXml) {
							var ws = new AscCommonExcel.Worksheet(wb, wb.aWorksheets.length);
							ws.sName = wbSheetXml.name;
							if (null !== wbSheetXml.bHidden) {
								ws.bHidden = wbSheetXml.bHidden;
							}
							//var wsView = new AscCommonExcel.asc_CSheetViewSettings();
							//wsView.pane = new AscCommonExcel.asc_CPane();
							//ws.sheetViews.push(wsView);
							if (contentSheetXml) {
								AscCommonExcel.executeInR1C1Mode(false, function () {
									var reader = new StaxParser(contentSheetXml, wsPart, xmlParserContext);
									ws.fromXml(reader);
								});
							}
							res.push(ws);
						}
					}
				});
			}

			var readSheetDataExternal = function (bNoBuildDep) {
				for (var i = 0; i < xmlParserContext.InitOpenManager.oReadResult.sheetData.length; ++i) {
					var sheetDataElem = xmlParserContext.InitOpenManager.oReadResult.sheetData[i];
					var ws = sheetDataElem.ws;

					var tmp = {
						pos: null,
						len: null,
						bNoBuildDep: bNoBuildDep,
						ws: ws,
						row: new AscCommonExcel.Row(ws),
						cell: new AscCommonExcel.Cell(ws),
						formula: new AscCommonExcel.OpenFormula(),
						sharedFormulas: {},
						prevFormulas: {},
						siFormulas: {},
						prevRow: -1,
						prevCol: -1,
						formulaArray: []
					};


					var sheetData = new AscCommonExcel.CT_SheetData();
					xmlParserContext.InitOpenManager.tmp = tmp;

					sheetDataElem.reader.setState(sheetDataElem.state);
					//TODO пересмотреть фунцию fromXml
					sheetData.fromXml2(sheetDataElem.reader);

					if (!bNoBuildDep) {
						//TODO возможно стоит делать это в worksheet после полного чтения
						//***array-formula***
						//добавление ко всем ячейкам массива головной формулы
						for (var j = 0; j < tmp.formulaArray.length; j++) {
							var curFormula = tmp.formulaArray[j];
							var ref = curFormula.ref;
							if (ref) {
								var rangeFormulaArray = tmp.ws.getRange3(ref.r1, ref.c1, ref.r2, ref.c2);
								rangeFormulaArray._foreach(function (cell) {
									cell.setFormulaInternal(curFormula);
									if (curFormula.ca || cell.isNullTextString()) {
										tmp.ws.workbook.dependencyFormulas.addToChangedCell(cell);
									}
								});
							}
						}
						for (var nCol in tmp.prevFormulas) {
							if (tmp.prevFormulas.hasOwnProperty(nCol)) {
								var prevFormula = tmp.prevFormulas[nCol];
								if (!tmp.siFormulas[prevFormula.parsed.getListenerId()]) {
									prevFormula.parsed.buildDependencies();
								}
							}
						}
						for (var listenerId in tmp.siFormulas) {
							if (tmp.siFormulas.hasOwnProperty(listenerId)) {
								tmp.siFormulas[listenerId].buildDependencies();
							}
						}
					}
				}
			};

			xmlParserContext.InitOpenManager.aCellXfs = [];
			xmlParserContext.InitOpenManager.Dxfs = [];

			readSheetDataExternal(true);
		}, this, []);

		return res;
	};

  // Эвент о пришедщих изменениях
  spreadsheet_api.prototype.syncCollaborativeChanges = function() {
    // Для быстрого сохранения уведомлять не нужно.
    if (!this.collaborativeEditing.getFast()) {
      this.handlers.trigger("asc_onCollaborativeChanges");
    }
  };

  // Применение изменений документа, пришедших при открытии
  // Их нужно применять после того, как мы создали WorkbookView
  // т.к. автофильтры, диаграммы, изображения и комментарии завязаны на WorksheetView (ToDo переделать)
  spreadsheet_api.prototype._applyFirstLoadChanges = function() {
    if (this.isDocumentLoadComplete) {
      return;
    }
    if (this.collaborativeEditing.applyChanges()) {
      // Изменений не было
      this.onDocumentContentReady();
    }
    // Пересылаем свои изменения (просто стираем чужие lock-и, т.к. своих изменений нет)
    this.collaborativeEditing.sendChanges();
  };

  // GoTo
  spreadsheet_api.prototype._goToComment = function(data) {
    var comment = this.wbModel.getComment(data);
    if (comment) {
      this.asc_showWorksheet(this.wbModel.getWorksheetById(comment.wsId).getIndex());
      this.asc_selectComment(comment.nId);
      this.asc_showComment(comment.nId);
    }
  };
  spreadsheet_api.prototype._goToBookmark = function(data) {
    // Disable edit because if there is no name, we will try to create it
    var tmpRestrictions = this.restrictions;
    this.restrictions = Asc.c_oAscRestrictionType.View;
    // Set A1-mode because all bookmarks in A1-mode
    var tmpR1C1 = AscCommonExcel.g_R1C1Mode;
    AscCommonExcel.g_R1C1Mode = false;
    this.asc_findCell(data);
    // Restore variables
    this.restrictions = tmpRestrictions;
    AscCommonExcel.g_R1C1Mode = tmpR1C1;
  };

  /////////////////////////////////////////////////////////////////////////
  ///////////////////CoAuthoring and Chat api//////////////////////////////
  /////////////////////////////////////////////////////////////////////////
  spreadsheet_api.prototype._coAuthoringInitEnd = function() {
    var t = this;
    this.collaborativeEditing = new AscCommonExcel.CCollaborativeEditing(/*handlers*/{
      "askLock": function() {
        t.CoAuthoringApi.askLock.apply(t.CoAuthoringApi, arguments);
      },
      "releaseLocks": function() {
        t.CoAuthoringApi.releaseLocks.apply(t.CoAuthoringApi, arguments);
      },
      "sendChanges": function() {
        t._onSaveChanges.apply(t, arguments);
      },
      "applyChanges": function() {
        t._onApplyChanges.apply(t, arguments);
      },
      "updateAfterApplyChanges": function() {
        t._onUpdateAfterApplyChanges.apply(t, arguments);
      },
      "drawSelection": function() {
        t._onDrawSelection.apply(t, arguments);
      },
      "drawFrozenPaneLines": function() {
        t._onDrawFrozenPaneLines.apply(t, arguments);
      },
      "updateAllSheetsLock": function() {
        t._onUpdateAllSheetsLock.apply(t, arguments);
      },
      "showDrawingObjects": function() {
        t._onShowDrawingObjects.apply(t, arguments);
      },
      "showComments": function() {
        t._onShowComments.apply(t, arguments);
      },
      "cleanSelection": function() {
        t._onCleanSelection.apply(t, arguments);
      },
      "updateDocumentCanSave": function() {
        t._onUpdateDocumentCanSave();
      },
      "checkCommentRemoveLock": function(lockElem) {
        return t._onCheckCommentRemoveLock(lockElem);
      },
      "unlockDefName": function() {
        t._onUnlockDefName.apply(t, arguments);
      },
      "checkDefNameLock": function(lockElem) {
        return t._onCheckDefNameLock(lockElem);
      },
      "updateAllLayoutsLock": function() {
          t._onUpdateAllLayoutsLock.apply(t, arguments);
      },
      "updateAllHeaderFooterLock": function() {
          t._onUpdateAllHeaderFooterLock.apply(t, arguments);
      },
      "updateAllPrintScaleLock": function() {
          t._onUpdateAllPrintScaleLock.apply(t, arguments);
      },
      "updateAllSheetViewLock": function() {
          if (t._onUpdateAllSheetViewLock) {
            t._onUpdateAllSheetViewLock.apply(t, arguments);
          }
      },
      "unlockCF": function() {
        t._onUnlockCF.apply(t, arguments);
      },
      "checkCFRemoveLock": function(lockElem) {
        return t._onCheckCFRemoveLock(lockElem);
      },
      "unlockProtectedRange": function() {
      	t._onUnlockProtectedRange.apply(t, arguments);
      },
      "checkProtectedRangeRemoveLock": function(lockElem) {
      	return t._onCheckProtectedRangeRemoveLock(lockElem);
      },
      "unlockUserProtectedRanges": function() {
      	t._onUnlockUserProtectedRanges.apply(t, arguments);
      }
    }, this.getViewMode());

    this.CoAuthoringApi.onConnectionStateChanged = function(e) {
		if (true === AscCommon.CollaborativeEditing.Is_Fast() && false === e['state']) {
			t.wb.Remove_ForeignCursor(e['id']);
		}
    	t.handlers.trigger("asc_onConnectionStateChanged", e);
    };
    this.CoAuthoringApi.onLocksAcquired = function(e) {
      if (t._coAuthoringCheckEndOpenDocument(t.CoAuthoringApi.onLocksAcquired, e)) {
        return;
      }

      if (2 != e["state"]) {
        var elementValue = e["blockValue"];
        var lockElem = t.collaborativeEditing.getLockByElem(elementValue, c_oAscLockTypes.kLockTypeOther);
        if (null === lockElem) {
          lockElem = new AscCommonExcel.CLock(elementValue);
          t.collaborativeEditing.addUnlock(lockElem);
        }

        var drawing, lockType = lockElem.Element["type"];
        var oldType = lockElem.getType();
        if (c_oAscLockTypes.kLockTypeOther2 === oldType || c_oAscLockTypes.kLockTypeOther3 === oldType) {
          lockElem.setType(c_oAscLockTypes.kLockTypeOther3, true);
        } else {
          lockElem.setType(c_oAscLockTypes.kLockTypeOther, true);
        }

        // Выставляем ID пользователя, залочившего данный элемент
        lockElem.setUserId(e["user"]);

        if (lockType === c_oAscLockTypeElem.Object) {
          drawing = g_oTableId.Get_ById(lockElem.Element["rangeOrObjectId"]);
          if (drawing) {
            var bLocked, bLocked2;
            bLocked = drawing.lockType !== c_oAscLockTypes.kLockTypeNone && drawing.lockType !== c_oAscLockTypes.kLockTypeMine;

            drawing.lockType = lockElem.Type;

            if(drawing instanceof AscCommon.CCore) {
              bLocked2 = drawing.lockType !== c_oAscLockTypes.kLockTypeNone && drawing.lockType !== c_oAscLockTypes.kLockTypeMine;
              if(bLocked2 !== bLocked) {
                t.sendEvent("asc_onLockCore", bLocked2);
              }
            }
          }
        }

        if (t.wb) {
          // Шлем update для toolbar-а, т.к. когда select в lock ячейке нужно заблокировать toolbar
          t.wb._onWSSelectionChanged(true);

          // Шлем update для листов
          t._onUpdateSheetsLock(lockElem);

          t._onUpdateDefinedNames(lockElem);

          //эвент о локе в меню вкладки layout(кроме print area)
          t._onUpdateLayoutLock(lockElem);
          //эвент о локе в меню опции print area во вкладке layout
          t._onUpdatePrintAreaLock(lockElem);
          //эвент о локе в меню опции headers/footers во вкладке layout
          t._onUpdateHeaderFooterLock(lockElem);
          //эвент о локе в меню опции scale во вкладке layout
          t._onUpdatePrintScaleLock(lockElem);
          //эвент о локе представлений
          if (t._onUpdateNamedSheetViewLock) {
            t._onUpdateNamedSheetViewLock(lockElem);
          }

          t._onUpdateCFLock(lockElem);
          t._onUpdateProtectedRangesLock(lockElem);
          t._onUpdateUserProtectedRange(lockElem);


          var ws = t.wb.getWorksheet();
          var lockSheetId = lockElem.Element["sheetId"];
          if (lockSheetId === ws.model.getId()) {
            if (lockType === c_oAscLockTypeElem.Object) {
              // Нужно ли обновлять закрепление областей
              if (t._onUpdateFrozenPane(lockElem)) {
                ws.draw();
              } else if (drawing && ws.model === drawing.worksheet) {
                if (ws.objectRender) {
                  ws.objectRender.showDrawingObjects();
                }
              }
            } else if (lockType === c_oAscLockTypeElem.Range || lockType === c_oAscLockTypeElem.Sheet) {
              ws.updateSelection();
            }
          } else if (-1 !== lockSheetId && 0 === lockSheetId.indexOf(AscCommonExcel.CCellCommentator.sStartCommentId)) {
            // Коммментарий
            t.handlers.trigger("asc_onLockComment", lockElem.Element["rangeOrObjectId"], e["user"]);
          }
        }
      }
    };
    this.CoAuthoringApi.onLocksReleased = function(e, bChanges) {
      if (t._coAuthoringCheckEndOpenDocument(t.CoAuthoringApi.onLocksReleased, e, bChanges)) {
        return;
      }

      var element = e["block"];
      var lockElem = t.collaborativeEditing.getLockByElem(element, c_oAscLockTypes.kLockTypeOther);
      if (null != lockElem) {
        var curType = lockElem.getType();

        var newType = c_oAscLockTypes.kLockTypeNone;
        if (curType === c_oAscLockTypes.kLockTypeOther) {
          if (true != bChanges) {
            newType = c_oAscLockTypes.kLockTypeNone;
          } else {
            newType = c_oAscLockTypes.kLockTypeOther2;
          }
        } else if (curType === c_oAscLockTypes.kLockTypeMine) {
          // Такого быть не должно
          newType = c_oAscLockTypes.kLockTypeMine;
        } else if (curType === c_oAscLockTypes.kLockTypeOther2 || curType === c_oAscLockTypes.kLockTypeOther3) {
          newType = c_oAscLockTypes.kLockTypeOther2;
        }

        if (t.wb) {
          t.wb.getWorksheet().cleanSelection();
        }

        var drawing;
        if (c_oAscLockTypes.kLockTypeNone !== newType) {
          lockElem.setType(newType, true);
        } else {
          // Удаляем из lock-ов, тот, кто правил ушел и не сохранил
          t.collaborativeEditing.removeUnlock(lockElem);
          if (!t._onCheckCommentRemoveLock(lockElem.Element)) {
            if (lockElem.Element["type"] === c_oAscLockTypeElem.Object) {
              drawing = g_oTableId.Get_ById(lockElem.Element["rangeOrObjectId"]);
              if (drawing) {
                var bLocked = drawing.lockType !== c_oAscLockTypes.kLockTypeNone && drawing.lockType !== c_oAscLockTypes.kLockTypeMine;
                drawing.lockType = c_oAscLockTypes.kLockTypeNone;
                if(drawing instanceof AscCommon.CCore) {
                  if(bLocked){
                    t.sendEvent("asc_onLockCore", false);
                  }
                }
              }
            }
          }
        }
        if (t.wb) {
          // Шлем update для листов
          t._onUpdateSheetsLock(lockElem);
          /*снимаем лок для DefName*/
          t.handlers.trigger("asc_onLockDefNameManager",Asc.c_oAscDefinedNameReason.OK);
          //эвент о локе в меню вкладки layout
          t._onUpdateLayoutLock(lockElem);
          //эвент о локе в меню опции print area во вкладке layout
          t._onUpdatePrintAreaLock(lockElem);
          //эвент о локе в меню опции headers/footers во вкладке layout
          t._onUpdateHeaderFooterLock(lockElem);
          //эвент о локе в меню опции scale во вкладке layout
          t._onUpdatePrintScaleLock(lockElem);
          /*снимаем лок c защищенных юзером диапазонов*/
          t.handlers.trigger("asc_onLockUserProtectedManager");
        }
      }
    };
    this.CoAuthoringApi.onLocksReleasedEnd = function() {
      if (!t.isDocumentLoadComplete) {
        // Пока документ еще не загружен ничего не делаем
        return;
      }

      if (t.wb) {
        // Шлем update для toolbar-а, т.к. когда select в lock ячейке нужно сбросить блокировку toolbar
        t.wb._onWSSelectionChanged(true);

        var worksheet = t.wb.getWorksheet();
        worksheet.cleanSelection();
        worksheet._drawSelection();
        worksheet._drawFrozenPaneLines();
        if (worksheet.objectRender) {
          worksheet.objectRender.showDrawingObjects();
        }
      }
    };
    this.CoAuthoringApi.onSaveChanges = function(e, userId, bFirstLoad) {
      t.collaborativeEditing.addChanges(e);
      if (!bFirstLoad && t.isDocumentLoadComplete) {
        t.syncCollaborativeChanges();
      }
    };
	this.CoAuthoringApi.onChangesIndex = function(changesIndex)
	{
		if (t.isLiveViewer() && changesIndex >= 0) {
			//todo
		}
	};
    this.CoAuthoringApi.onRecalcLocks = function(excelAdditionalInfo) {
      if (!excelAdditionalInfo) {
        return;
      }

      var tmpAdditionalInfo = JSON.parse(excelAdditionalInfo);
      // Это мы получили recalcIndexColumns и recalcIndexRows
      var oRecalcIndexColumns = t.collaborativeEditing.addRecalcIndex('0', tmpAdditionalInfo['indexCols']);
      var oRecalcIndexRows = t.collaborativeEditing.addRecalcIndex('1', tmpAdditionalInfo['indexRows']);

      // Теперь нужно пересчитать индексы для lock-элементов
      if (null !== oRecalcIndexColumns || null !== oRecalcIndexRows) {
        t.collaborativeEditing._recalcLockArray(c_oAscLockTypes.kLockTypeMine, oRecalcIndexColumns, oRecalcIndexRows);
        t.collaborativeEditing._recalcLockArray(c_oAscLockTypes.kLockTypeOther, oRecalcIndexColumns, oRecalcIndexRows);
      }
        if (true === AscCommon.CollaborativeEditing.Is_Fast()) {
            AscCommon.CollaborativeEditing.UpdateForeignCursorByAdditionalInfo(tmpAdditionalInfo);
        }
    };
    this.CoAuthoringApi.onCursor = function(e) {
      if (AscCommon.CollaborativeEditing && true === AscCommon.CollaborativeEditing.Is_Fast() && e && e[e.length - 1]) {
        t.wb.Update_ForeignCursor(e[e.length - 1]['cursor'], e[e.length - 1]['user'], true, e[e.length - 1]['useridoriginal']);
      }
    };
	  this.CoAuthoringApi.onParticipantsChangedOrigin     = function(users)
	  {
		  let m_bIsCollaborativeWithLiveViewer = users && -1 !== users.findIndex(function(element) {
				  return !!element['isLiveViewer'];
			  });
		  t.collaborativeEditing.m_bIsCollaborativeWithLiveViewer = m_bIsCollaborativeWithLiveViewer;
		  if (t.isDocumentLoadComplete && m_bIsCollaborativeWithLiveViewer) {
			  AscCommon.History.Clear();
		  }
	  };
  };

  spreadsheet_api.prototype._onSaveChanges = function(recalcIndexColumns, recalcIndexRows, isAfterAskSave) {
    if (this.isDocumentLoadComplete) {
      var arrChanges = this.wbModel.SerializeHistory();
      var deleteIndex = History.GetDeleteIndex();
      var excelAdditionalInfo = null;
      var bCollaborative = this.collaborativeEditing.getCollaborativeEditing();
      if (bCollaborative) {
        // Пересчетные индексы добавляем только если мы не одни
        if (recalcIndexColumns || recalcIndexRows) {
          excelAdditionalInfo = {"indexCols": recalcIndexColumns, "indexRows": recalcIndexRows};
        }
      }
      if (0 < arrChanges.length || null !== deleteIndex || null !== excelAdditionalInfo) {
          var oWs = this.wb.getWorksheet();
          var sCursorBinary = "";
          if (oWs && oWs.objectRender) {
              sCursorBinary = oWs.objectRender.getDocumentPositionBinary();
          }
          if(typeof sCursorBinary === "string" && sCursorBinary.length > 0) {
              if(!AscCommon.isRealObject(excelAdditionalInfo)) {
                  excelAdditionalInfo = {};
              }
              excelAdditionalInfo["UserId"] = this.CoAuthoringApi.getUserConnectionId();
              excelAdditionalInfo["UserShortId"] = this.DocInfo.get_UserId();
              excelAdditionalInfo["CursorInfo"] = this.wb.getCursorInfo();
          }
        this.CoAuthoringApi.saveChanges(arrChanges, deleteIndex, excelAdditionalInfo, this.canUnlockDocument2, bCollaborative);
        History.CanNotAddChanges = true;
      } else {
        this.CoAuthoringApi.unLockDocument(!!isAfterAskSave, this.canUnlockDocument2, null, bCollaborative);
      }
      this.canUnlockDocument2 = false;
    }
  };

  spreadsheet_api.prototype._onApplyChanges = function(changes, fCallback) {
	  let t = this;
	  let perfStart = performance.now();
	  let callback = fCallback;
	  if (!this.isDocumentLoadComplete) {
		  callback = function() {
			  let perfEnd = performance.now();
			  AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onApplyChanges", perfEnd - perfStart), t);
			  fCallback();
		  };

	  }

    this.inkDrawer.startSilentMode();
    this.wbModel.DeserializeHistory(changes, callback);
    this.inkDrawer.endSilentMode();
  };

  spreadsheet_api.prototype._onUpdateAfterApplyChanges = function() {
    if (!this.isDocumentLoadComplete) {
      // При открытии после принятия изменений мы должны сбросить пересчетные индексы
      this.collaborativeEditing.clearRecalcIndex();
      this.onDocumentContentReady();
    } else if (this.wb && !window["NATIVE_EDITOR_ENJINE"]) {
      // Нужно послать 'обновить свойства' (иначе для удаления данных не обновится строка формул).
      // ToDo Возможно стоит обновлять только строку формул
      AscCommon.CollaborativeEditing.Load_Images();
      if(AscCommon.CollaborativeEditing.Is_Fast()) {
          AscCommon.CollaborativeEditing.Refresh_ForeignCursors();
      }
      this.wb._onWSSelectionChanged();
      History.TurnOff();
      this.wb.drawWorksheet();
      History.TurnOn();
    }
    if (this.isApplyChangesOnVersionHistory)
    {
      this.isApplyChangesOnVersionHistory = false;
      this._openVersionHistoryEndCallback();
    }
  };

  spreadsheet_api.prototype._onCleanSelection = function() {
    if (this.wb) {
      this.wb.getWorksheet().cleanSelection();
    }
  };

  spreadsheet_api.prototype._onDrawSelection = function() {
    if (this.wb) {
      this.wb.getWorksheet()._drawSelection();
    }
  };

  spreadsheet_api.prototype._onDrawFrozenPaneLines = function() {
    if (this.wb) {
      this.wb.getWorksheet()._drawFrozenPaneLines();
    }
  };

	spreadsheet_api.prototype._onUpdateAllSheetsLock = function () {
		if (this.wbModel) {
			// Шлем update для листов
			this.handlers.trigger("asc_onWorkbookLocked", this.asc_isWorkbookLocked());
			var i, length, wsModel, wsIndex;
			for (i = 0, length = this.wbModel.getWorksheetCount(); i < length; ++i) {
				wsModel = this.wbModel.getWorksheet(i);
				wsIndex = wsModel.getIndex();
				this.handlers.trigger("asc_onWorksheetLocked", wsIndex, this.asc_isWorksheetLockedOrDeleted(wsIndex));
			}
		}
	};

  spreadsheet_api.prototype._onUpdateAllLayoutsLock = function () {
      var t = this;
      if (t.wbModel) {
          var i, length, wsModel, wsIndex;
          for (i = 0, length = t.wbModel.getWorksheetCount(); i < length; ++i) {
              wsModel = t.wbModel.getWorksheet(i);
              wsIndex = wsModel.getIndex();

              var isLocked = t.asc_isLayoutLocked(wsIndex);
              if (isLocked) {
                  t.handlers.trigger("asc_onLockDocumentProps", wsIndex);
              } else {
                  t.handlers.trigger("asc_onUnLockDocumentProps", wsIndex);
              }
          }
      }
  };

  spreadsheet_api.prototype._onUpdateAllHeaderFooterLock = function () {
      var t = this;
      if (t.wbModel) {
          var i, length, wsModel, wsIndex;
          for (i = 0, length = t.wbModel.getWorksheetCount(); i < length; ++i) {
              wsModel = t.wbModel.getWorksheet(i);
              wsIndex = wsModel.getIndex();

              var isLocked = t.asc_isLayoutLocked(wsIndex);
              if (isLocked) {
                  t.handlers.trigger("asc_onLockHeaderFooter", wsIndex);
              } else {
                  t.handlers.trigger("asc_onUnLockHeaderFooter", wsIndex);
              }
          }
      }
  };

	spreadsheet_api.prototype._onUpdateAllPrintScaleLock = function () {
		var t = this;
		if (t.wbModel) {
			var i, length, wsModel, wsIndex;
			for (i = 0, length = t.wbModel.getWorksheetCount(); i < length; ++i) {
				wsModel = t.wbModel.getWorksheet(i);
				wsIndex = wsModel.getIndex();

				var isLocked = t.asc_isLayoutLocked(wsIndex);
				if (isLocked) {
					t.handlers.trigger("asc_onLockPrintScale", wsIndex);
				} else {
					t.handlers.trigger("asc_onUnLockPrintScale", wsIndex);
				}
			}
		}
	};

  spreadsheet_api.prototype._onUpdateLayoutMenu = function (nSheetId) {
      var t = this;
      if (t.wbModel) {
          var wsModel = t.wbModel.getWorksheetById(nSheetId);
          if (wsModel) {
              var wsIndex = wsModel.getIndex();
              t.handlers.trigger("asc_onUpdateDocumentProps", wsIndex);
          }
      }
  };

  spreadsheet_api.prototype._onShowDrawingObjects = function() {
    if (this.wb) {
      var ws = this.wb.getWorksheet();
      if (ws && ws.objectRender) {
        ws.objectRender.showDrawingObjects();
      }
    }
  };

  spreadsheet_api.prototype._onShowComments = function() {
    if (this.wb) {
      this.wb.getWorksheet().cellCommentator.drawCommentCells();
    }
  };

	spreadsheet_api.prototype._onUpdateSheetsLock = function (lockElem) {
		// Шлем update для листов, т.к. нужно залочить лист
		if (c_oAscLockTypeElem.Sheet === lockElem.Element["type"]) {
			this._onUpdateAllSheetsLock();
		} else {
			// Шлем update для листа
			var wsModel = this.wbModel.getWorksheetById(lockElem.Element["sheetId"]);
			if (wsModel) {
				var wsIndex = wsModel.getIndex();
				this.handlers.trigger("asc_onWorksheetLocked", wsIndex, this.asc_isWorksheetLockedOrDeleted(wsIndex));
			}
		}
	};

  spreadsheet_api.prototype._onUpdateLayoutLock = function(lockElem) {
      var t = this;

      var wsModel = t.wbModel.getWorksheetById(lockElem.Element["sheetId"]);
      if (wsModel) {
          var wsIndex = wsModel.getIndex();

          var isLocked = t.asc_isLayoutLocked(wsIndex);
          if(isLocked) {
			  t.handlers.trigger("asc_onLockDocumentProps", wsIndex);
          } else {
			  t.handlers.trigger("asc_onUnLockDocumentProps", wsIndex);
          }
      }
  };

  spreadsheet_api.prototype._onUpdatePrintAreaLock = function(lockElem) {
      var t = this;

	  var wsModel = t.wbModel.getWorksheetById(lockElem.Element["sheetId"]);
	  var wsIndex = wsModel? wsModel.getIndex() : undefined;

      var isLocked = t.asc_isPrintAreaLocked(wsIndex);
      if(isLocked) {
          t.handlers.trigger("asc_onLockPrintArea");
      } else {
          t.handlers.trigger("asc_onUnLockPrintArea");
      }
  };

  spreadsheet_api.prototype._onUpdateHeaderFooterLock = function(lockElem) {
      var t = this;

      var wsModel = t.wbModel.getWorksheetById(lockElem.Element["sheetId"]);
      if (wsModel) {
          var wsIndex = wsModel.getIndex();

          var isLocked = t.asc_isHeaderFooterLocked();
          if(isLocked) {
              t.handlers.trigger("asc_onLockHeaderFooter");
          } else {
              t.handlers.trigger("asc_onUnLockHeaderFooter");
          }
      }
  };

  spreadsheet_api.prototype._onUpdatePrintScaleLock = function(lockElem) {
      var t = this;

      var wsModel = t.wbModel.getWorksheetById(lockElem.Element["sheetId"]);
      if (wsModel) {
          var wsIndex = wsModel.getIndex();

          var isLocked = t.asc_isPrintScaleLocked();
          if(isLocked) {
              t.handlers.trigger("asc_onLockPrintScale");
          } else {
              t.handlers.trigger("asc_onUnLockPrintScale");
          }
      }
  };

  spreadsheet_api.prototype._onUpdateFrozenPane = function(lockElem) {
    return (c_oAscLockTypeElem.Object === lockElem.Element["type"] && lockElem.Element["rangeOrObjectId"] === AscCommonExcel.c_oAscLockNameFrozenPane);
  };

  spreadsheet_api.prototype._sendWorkbookStyles = function () {
    // Для нативной версии (сборка и приложение) не генерируем стили
    if (this.wbModel && !window["NATIVE_EDITOR_ENJINE"]) {
      // Отправка стилей ячеек
      this.handlers.trigger("asc_onInitEditorStyles", this.wb.getCellStyles(this.styleThumbnailWidth, this.styleThumbnailHeight));
    }
  };

  spreadsheet_api.prototype.startCollaborationEditing = function() {
    // Начинаем совместное редактирование
    this.collaborativeEditing.startCollaborationEditing();

	  if (this.isDocumentLoadComplete) {
		  var worksheet = this.wb.getWorksheet();
		  worksheet.cleanSelection();
		  worksheet._drawSelection();
		  worksheet._drawFrozenPaneLines();
		  if (worksheet.objectRender) {
			  worksheet.objectRender.showDrawingObjects();
		  }
	  }
  };

  spreadsheet_api.prototype.endCollaborationEditing = function() {
    // Временно заканчиваем совместное редактирование
    this.collaborativeEditing.endCollaborationEditing();
  };

	// End Load document
	spreadsheet_api.prototype._openDocumentEndCallback = function () {
		// Не инициализируем дважды
		if (this.isDocumentLoadComplete || !this.ServerIdWaitComplete || !this.FontLoadWaitComplete) {
			return;
		}

        if (AscCommon.EncryptionWorker)
        {
            AscCommon.EncryptionWorker.init();
            if (!AscCommon.EncryptionWorker.isChangesHandled)
                return AscCommon.EncryptionWorker.handleChanges(this.collaborativeEditing.m_arrChanges, this, this._openDocumentEndCallback);
        }

		if (0 === this.wbModel.getWorksheetCount()) {
			this.sendEvent("asc_onError", c_oAscError.ID.ConvertationOpenError, c_oAscError.Level.Critical);
			return;
        }

		//история версий - возможно стоит грамотно чистить wbview, но не пересоздавать
		var previousVersionZoom;
		if ((this.VersionHistory || this.isEditOleMode) && this.controller) {
			var elem = document.getElementById("ws-v-scrollbar");
			if (elem) {
				elem.parentNode.removeChild(elem);
			}
			elem = document.getElementById("ws-h-scrollbar");
			if (elem) {
				elem.parentNode.removeChild(elem);
			}
			this.controller.vsbApi = null;
			this.controller.hsbApi = null;
			previousVersionZoom = this.wb && this.wb.getZoom();
		}

		this.wb = new AscCommonExcel.WorkbookView(this.wbModel, this.controller, this.handlers, this.HtmlElement,
			this.topLineEditorElement, this, this.collaborativeEditing, this.fontRenderingMode);

		if (this.isCopyOutEnabled && this.topLineEditorElement) {
			if (this.isCopyOutEnabled() === false) {
				this.topLineEditorElement.oncopy = function () {
					return false;
				};
				this.topLineEditorElement.oncut = function () {
					return false;
				};
			}
		}

		if (this.isMobileVersion) {
			this.wb.defaults.worksheetView.halfSelection = true;
			this.wb.defaults.worksheetView.activeCellBorderColor = new CColor(79, 158, 79);
			var _container = document.getElementById(this.HtmlElementName);
			if (_container) {
				_container.style.overflow = "hidden";
			}
			this.wb.MobileTouchManager = new AscCommonExcel.CMobileTouchManager({eventsElement: "cell_mobile_element"});
			this.wb.MobileTouchManager.Init(this);

			// input context must be created!!!
			var _areaId = AscCommon.g_inputContext.HtmlArea.id;
			var _element = document.getElementById(_areaId);
			_element.parentNode.parentNode.style.zIndex = 10;

			this.wb.MobileTouchManager.initEvents(AscCommon.g_inputContext.HtmlArea.id);
		}

		this.asc_CheckGuiControlColors();
		this.sendColorThemes(this.wbModel.theme);
		this.asc_ApplyColorScheme(false);

		this.sendStandartTextures();
		this.sendMathToMenu();

		this._applyPreOpenLocks();
		// Применяем пришедшие при открытии изменения
		this._applyFirstLoadChanges();
		// Go to if sent options
		this.goTo();

		// Меняем тип состояния (на никакое)
		this.advancedOptionsAction = c_oAscAdvancedOptionsAction.None;

		// Были ошибки при открытии, посылаем предупреждение
		if (0 < this.wbModel.openErrors.length) {
			this.sendEvent('asc_onError', c_oAscError.ID.OpenWarning, c_oAscError.Level.NoCritical);
		}

		if (this.VersionHistory || this.isEditOleMode) {
			if (this.VersionHistory && this.VersionHistory.changes) {
				this.VersionHistory.applyChanges(this);
			}
			this.sheetsChanged();
			if (previousVersionZoom) {
				this.asc_setZoom(previousVersionZoom);
			}
			this.asc_Resize();
		}
		
		if (this.canEdit() && this.asc_getExternalReferences()) {
			this.handlers.trigger("asc_onNeedUpdateExternalReferenceOnOpen");
		}
		//this.asc_Resize(); // Убрал, т.к. сверху приходит resize (http://bugzilla.onlyoffice.com/show_bug.cgi?id=14680)
	};

	// Переход на диапазон в листе
	spreadsheet_api.prototype._asc_setWorksheetRange = function (val) {
		// Получаем sheet по имени
		var ranges = null, ws;
        var sheet = val.asc_getSheet();
        if (!sheet) {
			ranges = AscCommonExcel.getRangeByRef(val.asc_getLocation(), this.wbModel.getActiveWs(), true);
			if (ranges = ranges[0]) {
				ws = ranges.worksheet;
            }
        } else {
        	//при переходе к диапазону внутри документа, игнорируется регистр у имени листа -> Sheet1===SHEEt1
			ws = this.wbModel.getWorksheetByName(sheet, true);
        }
		if (!ws) {
			this.handlers.trigger("asc_onHyperlinkClick", null);
			return;
		} else if (ws.getHidden()) {
			return;
		}
		// Индекс листа
		var sheetIndex = ws.getIndex();
		// Если не совпали индекс листа и индекс текущего, то нужно сменить
		if (this.asc_getActiveWorksheetIndex() !== sheetIndex) {
			// Меняем активный лист
			this.asc_showWorksheet(sheetIndex);
		}
		var range;
		if (ranges) {
			range = ranges.bbox;
        } else {
			range = ws.getRange2(val.asc_getRange());
			if (range) {
				range = range.getBBox0();
            }
        }
		this.wb._onSetSelection(range, /*validRange*/ true);
	};

	spreadsheet_api.prototype.asc_setWorksheetRange = function (val) {
        this._asc_setWorksheetRange(val);
    };

	spreadsheet_api.prototype._onSaveCallbackInner = function () {
		var t = this;

		AscCommon.CollaborativeEditing.Clear_CollaborativeMarks();
		// Принимаем чужие изменения
		this.collaborativeEditing.applyChanges();

		this.CoAuthoringApi.onUnSaveLock = function () {
			t.CoAuthoringApi.onUnSaveLock = null;
			if (t.isForceSaveOnUserSave && t.IsUserSave) {
				t.forceSaveButtonContinue = t.forceSave();
			}
			if (t.forceSaveForm) {
				t.forceSaveForm();
			}

			if (t.collaborativeEditing.getCollaborativeEditing()) {
				// Шлем update для toolbar-а, т.к. когда select в lock ячейке нужно заблокировать toolbar
				t.wb._onWSSelectionChanged(true);
			}

			t.canSave = true;
			t.IsUserSave = false;
			t.lastSaveTime = null;

			if (!t.forceSaveButtonContinue) {
				t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.Save);
			}
			// Обновляем состояние возможности сохранения документа
			t.onUpdateDocumentModified(History.Have_Changes());

			if (undefined !== window["AscDesktopEditor"]) {
				window["AscDesktopEditor"]["OnSave"]();
			}
			if (t.disconnectOnSave) {
				t.CoAuthoringApi.disconnect(t.disconnectOnSave.code, t.disconnectOnSave.reason);
				t.disconnectOnSave = null;
			}

			if (t.canUnlockDocument) {
				t._unlockDocument();
			}
		};
		// Пересылаем свои изменения
		this.collaborativeEditing.sendChanges(this.IsUserSave, true);
	};

  spreadsheet_api.prototype._isLockedSparkline = function (id, callback) {
    var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Object, /*subType*/null,
        this.asc_getActiveWorksheetId(), id);
    this.collaborativeEditing.lock([lockInfo], callback);
  };

	spreadsheet_api.prototype._isLockedAddWorksheets = function(callback) {
		// Проверка глобального лока
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}
		var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null,
			AscCommonExcel.c_oAscLockAddSheet, AscCommonExcel.c_oAscLockAddSheet);
		this.collaborativeEditing.lock([lockInfo], callback);
	};
	spreadsheet_api.prototype._addWorksheetsCheck = function(callback) {
		let t = this;
		if (this.asc_isProtectedWorkbook()) {
			callback.call(t, false);
		}
		this._isLockedAddWorksheets(function (res) {
			callback.call(t, res);
		});
	};
	spreadsheet_api.prototype._addWorksheetsWithoutLock = function (arrNames, where) {
		var res = [];
		this.inkDrawer.startSilentMode();
		History.Create_NewPoint();
		History.StartTransaction();
		for (var i = arrNames.length - 1; i >= 0; --i) {
			res.push(this.wbModel.createWorksheet(where, arrNames[i]));
		}
		this.wbModel.setActive(where);
		this.wb.updateWorksheetByModel();
		this.wb.showWorksheet();
		History.EndTransaction();
		// Посылаем callback об изменении списка листов
		this.sheetsChanged();
		this.inkDrawer.endSilentMode();
		return res;
	};

  // Workbook interface

  spreadsheet_api.prototype.asc_getWorksheetsCount = function() {
    return this.wbModel.getWorksheetCount();
  };

  spreadsheet_api.prototype.asc_getWorksheetName = function(index) {
    return this.wbModel.getWorksheet(index).getName();
  };

  spreadsheet_api.prototype.asc_getWorksheetTabColor = function(index) {
    return this.wbModel.getWorksheet(index).getTabColor();
  };
  spreadsheet_api.prototype.asc_setWorksheetTabColor = function(color, arrSheets) {
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }

    if (this.asc_isProtectedWorkbook()) {
      return false;
    }

    if (!arrSheets) {
      arrSheets = [this.wbModel.getActive()];
    }

    var sheet, arrLocks = [];
    for (var i = 0; i < arrSheets.length; ++i) {
      sheet = this.wbModel.getWorksheet(arrSheets[i]);
      arrLocks.push(this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Object, /*subType*/null, sheet.getId(), AscCommonExcel.c_oAscLockNameTabColor));
    }

    var t = this;
    var changeTabColorCallback = function(res) {
      if (res) {
        color = AscCommonExcel.CorrectAscColor(color);
        History.Create_NewPoint();
        History.StartTransaction();
        for (var i = 0; i < arrSheets.length; ++i) {
          t.wbModel.getWorksheet(arrSheets[i]).setTabColor(color);
        }
        History.EndTransaction();
      }
    };
    this.collaborativeEditing.lock(arrLocks, changeTabColorCallback);
  };

  spreadsheet_api.prototype.asc_getActiveWorksheetIndex = function() {
    return this.wbModel.getActive();
  };

  spreadsheet_api.prototype.asc_getActiveWorksheetId = function() {
    var activeIndex = this.wbModel.getActive();
    return this.wbModel.getWorksheet(activeIndex).getId();
  };

  spreadsheet_api.prototype.asc_getWorksheetId = function(index) {
    return this.wbModel.getWorksheet(index).getId();
  };

  spreadsheet_api.prototype.asc_isWorksheetHidden = function(index) {
    return this.wbModel.getWorksheet(index).getHidden();
  };

  spreadsheet_api.prototype.asc_getDefinedNames = function(defNameListId, excludeErrorRefNames) {
    return this.wb.getDefinedNames(defNameListId,excludeErrorRefNames);
  };

  spreadsheet_api.prototype.asc_setDefinedNames = function(defName) {
//            return this.wb.setDefinedNames(defName);
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }
    return this.wb.editDefinedNames(null, defName);
  };

  spreadsheet_api.prototype.asc_editDefinedNames = function(oldName, newName) {
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    return this.wb.editDefinedNames(oldName, newName);
  };

  spreadsheet_api.prototype.asc_delDefinedNames = function(oldName) {
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }
    return this.wb.delDefinedNames(oldName);
  };

  spreadsheet_api.prototype.asc_checkDefinedName = function(checkName, scope) {
    return this.wbModel.checkDefName(checkName, scope);
  };

  spreadsheet_api.prototype.asc_getDefaultDefinedName = function() {
    return this.wb.getDefaultDefinedName();
  };

  spreadsheet_api.prototype.asc_getDefaultTableStyle = function() {
      return this.wb.getDefaultTableStyle();
  };

  spreadsheet_api.prototype._onUpdateDefinedNames = function(lockElem) {
//      if( lockElem.Element["subType"] == AscCommonExcel.c_oAscLockTypeElemSubType.DefinedNames ){
      if( lockElem.Element["sheetId"] == -1 && lockElem.Element["rangeOrObjectId"] != -1 && !this.collaborativeEditing.getFast() ){
          var dN = this.wbModel.dependencyFormulas.getDefNameByNodeId(lockElem.Element["rangeOrObjectId"]);
          if (dN) {
              dN.isLock = lockElem.UserId;
              this.handlers.trigger("asc_onRefreshDefNameList",dN.getAscCDefName());
          }
          this.handlers.trigger("asc_onLockDefNameManager",Asc.c_oAscDefinedNameReason.LockDefNameManager);
      }
  };

  spreadsheet_api.prototype._onUnlockDefName = function() {
    this.wb.unlockDefName();
  };

  spreadsheet_api.prototype._onCheckDefNameLock = function() {
    return this.wb._onCheckDefNameLock();
  };

  // Залочена ли работа с листом
  spreadsheet_api.prototype.asc_isWorksheetLockedOrDeleted = function(index) {
    var ws = this.wbModel.getWorksheet(index);
    var sheetId = null;
    if (null === ws || undefined === ws) {
      sheetId = this.asc_getActiveWorksheetId();
    } else {
      sheetId = ws.getId();
    }

    var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheetId, sheetId);
    // Проверим, редактирует ли кто-то лист
    return (false !== this.collaborativeEditing.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeOther, /*bCheckOnlyLockAll*/false));
  };

  // Залочена ли работа с листами
  spreadsheet_api.prototype.asc_isWorkbookLocked = function() {
    var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, null, null);
    // Проверим, редактирует ли кто-то лист
    return (false !== this.collaborativeEditing.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeOther, /*bCheckOnlyLockAll*/false));
  };

  // Залочена ли работа с листом
  spreadsheet_api.prototype.asc_isLayoutLocked = function(index) {
      var ws = this.wbModel.getWorksheet(index);
      var sheetId = null;
      if (null === ws || undefined === ws) {
          sheetId = this.asc_getActiveWorksheetId();
      } else {
          sheetId = ws.getId();
      }

      var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Object, /*subType*/null, sheetId, "layoutOptions");
      // Проверим, редактирует ли кто-то лист
      return (false !== this.collaborativeEditing.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeOther, /*bCheckOnlyLockAll*/false));
  };

	// Залочена ли работа с листом
  spreadsheet_api.prototype.asc_isHeaderFooterLocked = function(index) {
      var ws = this.wbModel.getWorksheet(index);
      var sheetId = null;
      if (null === ws || undefined === ws) {
          sheetId = this.asc_getActiveWorksheetId();
      } else {
          sheetId = ws.getId();
      }

      var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Object, /*subType*/null, sheetId, "headerFooter");
      // Проверим, редактирует ли кто-то лист
      return (false !== this.collaborativeEditing.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeOther, /*bCheckOnlyLockAll*/false));
  };

  spreadsheet_api.prototype.asc_isPrintScaleLocked = function(index) {
      var ws = this.wbModel.getWorksheet(index);
      var sheetId = null;
      if (null === ws || undefined === ws) {
          sheetId = this.asc_getActiveWorksheetId();
      } else {
          sheetId = ws.getId();
      }

      var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Object, /*subType*/null, sheetId, "printScaleOptions");
      return (false !== this.collaborativeEditing.getLockIntersection(lockInfo, c_oAscLockTypes.kLockTypeOther, /*bCheckOnlyLockAll*/false));
  };

	spreadsheet_api.prototype.asc_isPrintAreaLocked = function(index) {
		//проверка лока для именованного диапазона -  области печати
		//локи для именованных диапазонов устроены следующем образом: если изменяется хоть один именованный диапазон
		//добавлять новые/редактировать старые/удалять измененные нельзя, но зато разрешается удалять не измененные
		//в данном случае для области печати логика следующая - если был добавлен именованный диапазон(в тч и область печати)
		//не даём возможность добавлять/редактировть область печати на всех листах, но разрешаем удалять не измененные области печати
		//если вторым пользователем добавлена область печати, первым удалять разершается по общей схеме
		//но поскольку изменения не приняты, этой области печати ещё нет у первого пользователя - удаления не произойдёт

		var res = false;
		if(index !== undefined) {
			var sheetId = this.wbModel.getWorksheet(index).getId();
			var dN = this.wbModel.dependencyFormulas.getDefNameByName("Print_Area", sheetId);
			if(dN) {
				res = !!dN.isLock;
			}
		}

		return res;
	};

  spreadsheet_api.prototype.asc_getHiddenWorksheets = function() {
    var model = this.wbModel;
    var len = model.getWorksheetCount();
    var i, ws, res = [];

    for (i = 0; i < len; ++i) {
      ws = model.getWorksheet(i);
      if (ws.getHidden()) {
        res.push({"index": i, "name": ws.getName()});
      }
    }
    return res;
  };

  spreadsheet_api.prototype.asc_showWorksheet = function(index) {
  	if (typeof index === "number") {
      var t = this;
      var ws = this.wbModel.getWorksheet(index);
      var isHidden = ws.getHidden();
      var showWorksheetCallback = function(res) {
        if (res) {
          t.wbModel.getWorksheet(index).setHidden(false);
          t.wb.showWorksheet(index);
        }
      };
      if (isHidden) {
        if (this.asc_isProtectedWorkbook()) {
          return false;
        }
        var sheetId = this.wbModel.getWorksheet(index).getId();
        var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheetId, sheetId);
        this.collaborativeEditing.lock([lockInfo], showWorksheetCallback);
      } else {
        showWorksheetCallback(true);
      }
    }
  };

  spreadsheet_api.prototype.asc_hideWorksheet = function (arrSheets) {
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }

    if (this.asc_isProtectedWorkbook()) {
      return false;
    }

    if (!arrSheets) {
      arrSheets = [this.wbModel.getActive()];
    }

    // Вдруг остался один лист
    if (this.asc_getWorksheetsCount() <= this.asc_getHiddenWorksheets().length + arrSheets.length) {
      return false;
    }

    var sheet, arrLocks = [];
    for (var i = 0; i < arrSheets.length; ++i) {
      sheet = this.wbModel.getWorksheet(arrSheets[i]);
      arrLocks.push(this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheet.getId(), sheet.getId()));
    }

    var t = this;
    var hideWorksheetCallback = function(res) {
      if (res) {
        History.Create_NewPoint();
        History.StartTransaction();
        for (var i = 0; i < arrSheets.length; ++i) {
          t.wbModel.getWorksheet(arrSheets[i]).setHidden(true);
        }
        History.EndTransaction();
      }
    };

    this.collaborativeEditing.lock(arrLocks, hideWorksheetCallback);
    return true;
  };

  spreadsheet_api.prototype.asc_renameWorksheet = function(name) {
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }

    if (this.asc_isProtectedWorkbook()) {
      return false;
    }

    var i = this.wbModel.getActive();
    var sheetId = this.wbModel.getWorksheet(i).getId();
    var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheetId, sheetId);

    var t = this;
    var renameCallback = function(res) {
      if (res) {
        AscFonts.FontPickerByCharacter.getFontsByString(name);
        t._loadFonts([], function() {
            var oWorkbook = t.wbModel;
            var oWorksheet = oWorkbook.getWorksheet(i);
            var sOldName = oWorksheet.getName();
            oWorksheet.setName(name);
            t.sheetsChanged();
            if(t.wb) {
                //change sheet name in chart references
                t.wb.handleChartsOnChangeSheetName(oWorksheet, sOldName, name);
            }
        });
      } else {
        t.handlers.trigger("asc_onError", c_oAscError.ID.LockedWorksheetRename, c_oAscError.Level.NoCritical);
      }
    };

    this.collaborativeEditing.lock([lockInfo], renameCallback);
    return true;
  };

  spreadsheet_api.prototype.asc_addWorksheet = function (name) {
    var i = this.wbModel.getActive();
    this._addWorksheetsCheck(function (res) {
        if (res) {
            this._addWorksheetsWithoutLock([name], i + 1);
        }
    });
  };

  spreadsheet_api.prototype.asc_insertWorksheet = function (arrNames) {
  	// Support old versions
    if (!Array.isArray(arrNames)) {
      arrNames = [arrNames];
    }
    var i = this.wbModel.getActive();
    this._addWorksheetsCheck(function (res) {
        if (res) {
            this._addWorksheetsWithoutLock(arrNames, i);
        }
    });
  };

  // Удаление листа
  spreadsheet_api.prototype.asc_deleteWorksheet = function (arrSheets) {
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }

    if (this.asc_isProtectedWorkbook()) {
      return false;
    }

    if (!arrSheets) {
      arrSheets = [this.wbModel.getActive()];
    }

    // Check delete all
    if (this.wbModel.getWorksheetCount() === arrSheets.length) {
      return false;
    }

    var sheet, arrLocks = [];
    for (var i = 0; i < arrSheets.length; ++i) {
      sheet = arrSheets[i] = this.wbModel.getWorksheet(arrSheets[i]);
      arrLocks.push(this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheet.getId(), sheet.getId()));
    }

    var t = this;
    var deleteCallback = function(res) {
      if (res) {
        History.Create_NewPoint();
        History.StartTransaction();
        for (var i = 0; i < arrSheets.length; ++i) {
          t.wbModel.removeWorksheet(arrSheets[i].getIndex());
        }
        t.wbModel.handleChartsOnWorksheetsRemove(arrSheets);
        t.wb.updateWorksheetByModel();
        t.wb.showWorksheet();
        History.EndTransaction();
        // Посылаем callback об изменении списка листов
        t.sheetsChanged();
      }
    };

    this.collaborativeEditing.lock(arrLocks, deleteCallback);
    return true;
  };

  spreadsheet_api.prototype.asc_moveWorksheet = function (where, arrSheets) {
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }

    if (this.asc_isProtectedWorkbook()) {
      return false;
    }


    if (!arrSheets) {
      arrSheets = [this.wbModel.getActive()];
    }

    var active, sheet, i, index, _where, arrLocks = [];
    var arrSheetsLeft = [], arrSheetsRight = [];
    for (i = 0; i < arrSheets.length; ++i) {
      index = arrSheets[i];
      sheet = this.wbModel.getWorksheet(index);
      ((index < where) ? arrSheetsLeft : arrSheetsRight).push(sheet);
      if (!active) {
        active = sheet;
      }
      arrLocks.push(this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheet.getId(), sheet.getId()));
    }

    var t = this;
    var moveCallback = function (res) {
      if (res) {
        History.Create_NewPoint();
        History.StartTransaction();
        for (i = 0, _where = where; i < arrSheetsRight.length; ++i, ++_where) {
          index = arrSheetsRight[i].getIndex();
          if (index !== _where) {
            t.wbModel.replaceWorksheet(index, _where);
          }
        }
        for (i = arrSheetsLeft.length - 1, _where = where - 1; i >= 0; --i, --_where) {
          index = arrSheetsLeft[i].getIndex();
          if (index !== _where) {
            t.wbModel.replaceWorksheet(index, _where);
          }
        }
        // Обновим текущий номер
        t.wbModel.setActive(active.getIndex());
        t.wb.updateWorksheetByModel();
        t.wb.showWorksheet();
        History.EndTransaction();
        // Посылаем callback об изменении списка листов
        t.sheetsChanged();
      }
    };

    this.collaborativeEditing.lock(arrLocks, moveCallback);
    return true;
  };

  spreadsheet_api.prototype.asc_copyWorksheet = function (where, arrNames, arrSheets) {
    // Проверка глобального лока
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }

    if (this.asc_isProtectedWorkbook()) {
      return false;
    }

    // Support old versions
    if (!Array.isArray(arrNames)) {
      arrNames = [arrNames];
    }
    if (0 === arrNames.length) {
      return false;
    }
    if (!arrSheets) {
      arrSheets = [this.wbModel.getActive()];
    }

    var scale = this.asc_getZoom();

    // ToDo уйти от lock для листа при копировании
    var sheet, arrLocks = [];
    for (var i = 0; i < arrSheets.length; ++i) {
      sheet = arrSheets[i] = this.wbModel.getWorksheet(arrSheets[i]);
      arrLocks.push(this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheet.getId(), sheet.getId()));
    }

    var t = this;
    var copyWorksheet = function(res) {
      if (res) {
        // ToDo перейти от wsViews на wsViewsId
        History.Create_NewPoint();
        History.StartTransaction();
        var index;
        for (var i = arrSheets.length - 1; i >= 0; --i) {
          index = arrSheets[i].getIndex();
          t.wbModel.copyWorksheet(index, where, arrNames[i]);
        }
        // Делаем активным скопированный
        t.wbModel.setActive(where);
        t.wb.updateWorksheetByModel();
        t.wb.showWorksheet();
        History.EndTransaction();
        // Посылаем callback об изменении списка листов
        t.sheetsChanged();
      }
    };

    this.collaborativeEditing.lock(arrLocks, copyWorksheet);
  };

  spreadsheet_api.prototype.asc_StartMoveSheet = function (arrSheets) {
	  // Проверка глобального лока
	  // Лок каждого листа необходимо проверять в интерфейсе. если что-то залочено - не переносим
      if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
		  return false;
	  }

	  if (this.asc_isProtectedWorkbook()) {
		  return false;
	  }

	  //если выделены все - не перенесим(проверка в интерфейсе)
	  var sheet, sBinarySheet, res = [];
      var activeIndex = this.wbModel.nActive;
	  for (var i = 0; i < arrSheets.length; ++i) {
		  sheet = this.wbModel.getWorksheet(arrSheets[i]);
		  this.wbModel.nActive = sheet.getIndex();
		  sBinarySheet = AscCommonExcel.g_clipboardExcel.copyProcessor.getBinaryForCopy(sheet, null, null, true, true);
          res.push(sBinarySheet);
	  }
	  this.wbModel.nActive = activeIndex;

      return res;
  };

  spreadsheet_api.prototype.asc_EndMoveSheet = function(where, arrNames, arrSheets) {
	  // Проверка глобального лока
	  if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
		  return false;
	  }

	  if (this.asc_isProtectedWorkbook()) {
		  return false;
	  }

	  // Support old versions
	  if (!Array.isArray(arrNames)) {
		  arrNames = [arrNames];
	  }
	  if (0 === arrNames.length) {
		  return false;
	  }
	  if (!arrSheets) {
		  return false;
	  }

	  var scale = this.asc_getZoom();
	  var t = this;
	  var addWorksheet = function(res) {
		  if (res) {
			  // ToDo перейти от wsViews на wsViewsId
			  History.Create_NewPoint();
			  History.StartTransaction();

			  var renameParamsArr = [], renameSheetMap = {};
			  for (var i = arrSheets.length - 1; i >= 0; --i) {
				  t.wb.pasteSheet(arrSheets[i], where, arrNames[i], function(renameParams) {
					  // Делаем активным скопированный
					  renameParamsArr.push(renameParams);
					  renameSheetMap[renameParams.lastName] =  renameParams.newName;
					  t.asc_showWorksheet(where);
					  t.asc_setZoom(scale);
					  // Посылаем callback об изменении списка листов
					  t.sheetsChanged();
				  });

			  }
			  //парсинг формул после вставки всех листов, поскольку внутри одного листа может быть ссылка в формуле на другой лист который ещё не вставился
			  //поэтому дожидаемся вставку всех листов
			  for(var j = 0; j < renameParamsArr.length; j++) {
				var newSheet = t.wb.model.getWorksheetByName(renameParamsArr[j].newName);
			    newSheet.copyFromFormulas(renameParamsArr[j], renameSheetMap);
			  }

			  // Делаем активным скопированный
			  t.wbModel.setActive(where);
			  t.wb.updateWorksheetByModel();
			  t.wb.showWorksheet();
			  History.EndTransaction();
			  // Посылаем callback об изменении списка листов
			  t.sheetsChanged();
		  }
	  };

	  //TODO нужно лочить все листы
	  addWorksheet(true);
	  //this.collaborativeEditing.lock([], addWorksheet);
  };

  spreadsheet_api.prototype.asc_cleanSelection = function() {
    this.wb.getWorksheet().cleanSelection();
  };

  spreadsheet_api.prototype.asc_getZoom = function() {
    return this.wb.getZoom();
  };

  spreadsheet_api.prototype.asc_setZoom = function(scale) {
	  this.wb && this.wb.changeZoom(scale);
  };

  spreadsheet_api.prototype.asc_enableKeyEvents = function(isEnabled, isFromInput) {
    if (!this.isLoadFullApi) {
      this.tmpFocus = isEnabled;
      return;
    }

    if (this.wb) {
      this.wb.enableKeyEventsHandler(isEnabled);
    }

    if (isFromInput !== true && AscCommon.g_inputContext)
      AscCommon.g_inputContext.setInterfaceEnableKeyEvents(isEnabled);
  };

  spreadsheet_api.prototype.asc_IsFocus = function(bIsNaturalFocus) {
    var res = true;
	if(this.wb.cellEditor.isTopLineActive)
	{
		res = false;
	}
    else if (this.wb)
	{
      res = this.wb.getEnableKeyEventsHandler(bIsNaturalFocus);
    }

    return res;
  };

  spreadsheet_api.prototype.asc_searchEnabled = function(bIsEnabled) {
  };

  spreadsheet_api.prototype.asc_findText = function(options, callback) {
    var result = null;

	//***searchEngine
	var SearchEngine = this.wb.Search(options);
	var Id = this.wb.GetSearchElementId(!options || options.scanForward);

	if (null != Id) {
		this.wb.SelectSearchElement(Id);
	}

	if (window["NATIVE_EDITOR_ENJINE"]) {
		var ws = this.wb.getWorksheet();
		var activeCell = this.wbModel.getActiveWs().selectionRange.activeCell;
		result = [ws.getCellLeftRelative(activeCell.col, 0), ws.getCellTopRelative(activeCell.row, 0)];
	} else {
		result = SearchEngine.Count;
	}

    if (callback)
      callback(result);
    return result;
  };

  spreadsheet_api.prototype.asc_replaceText = function(options) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    var wb = this.wb;
    var ws = wb.getWorksheet();
    if (this.wb) {
    	//check user protected range. if sheet at least one protected for this user range
    	if (options.isReplaceAll) {
			if (options.scanOnOnlySheet === Asc.c_oAscSearchBy.Workbook && wb.model.isUserProtectedRangesIntersection()) {
				this.handlers.trigger("asc_onError", c_oAscError.ID.ProtectedRangeByOtherUser, c_oAscError.Level.NoCritical);
				return;
			} else if (options.scanOnOnlySheet === Asc.c_oAscSearchBy.Sheet && ws.model.isUserProtectedRangesIntersection()) {
				this.handlers.trigger("asc_onError", c_oAscError.ID.ProtectedRangeByOtherUser, c_oAscError.Level.NoCritical);
				return;
			}
		}
	}

    if (ws.model.getSheetProtection()) {
        this.handlers.trigger("asc_onError", c_oAscError.ID.CannotUseCommandProtectedSheet, c_oAscError.Level.NoCritical);
        return;
    }

    options.lookIn = Asc.c_oAscFindLookIn.Formulas; // При замене поиск только в формулах
	  wb.replaceCellText(options);
  };

  spreadsheet_api.prototype.asc_endFindText = function() {
    // Нужно очистить поиск
    this.wb._cleanFindResults();
  };

	//***searchEngine
  spreadsheet_api.prototype.sync_setSearchCurrent = function (nCurrent, nOverallCount) {
    this.sendEvent("asc_onSetSearchCurrent", nCurrent, nOverallCount);
  };

	spreadsheet_api.prototype.sync_startTextAroundSearch = function () {
		this.sendEvent("asc_onStartTextAroundSearch");
	};
	spreadsheet_api.prototype.sync_endTextAroundSearch = function () {
		this.sendEvent("asc_onEndTextAroundSearch");
	};
	spreadsheet_api.prototype.sync_getTextAroundSearchPack = function (arrElements) {
		this.sendEvent("asc_onGetTextAroundSearchPack", arrElements);
	};
	spreadsheet_api.prototype.sync_removeTextAroundSearch = function (sId) {
		this.sendEvent("asc_onRemoveTextAroundSearch", [sId]);
	};
	spreadsheet_api.prototype.sync_SearchEndCallback = function () {
		this.sendEvent("asc_onSearchEnd");
	};
    spreadsheet_api.prototype.sync_closeOleEditor = function() {
        this.sendEvent("asc_onCloseOleEditor");
    };
	spreadsheet_api.prototype.sync_changedElements = function (arr) {
		this.sendEvent("asc_onUpdateSearchElem", arr);
	};

	spreadsheet_api.prototype.asc_StartTextAroundSearch = function()
	{
		let wb = this.wb;
		if (!wb || !wb.SearchEngine)
			return;

		wb.SearchEngine.StartTextAround();
	};

	spreadsheet_api.prototype.asc_SelectSearchElement = function(sId)
	{
		let wb = this.wb;
		if (!wb || !wb.SearchEngine)
			return;

		wb.SelectSearchElement(sId);
	};

  /**
   * Делает активной указанную ячейку
   * @param {String} reference  Ссылка на ячейку вида A1 или R1C1
   */
  spreadsheet_api.prototype.asc_findCell = function (reference) {
    if (this.asc_getCellEditMode()) {
      return;
    }
    var ws = this.wb.getWorksheet();
    var d = ws.findCell(reference);
    if (0 === d.length) {
      return;
    }

    // Получаем sheet по имени
    ws = d[0].getWorksheet();
    if (!ws || ws.getHidden()) {
      return;
    }
    // Индекс листа
    var sheetIndex = ws.getIndex();
    // Если не совпали индекс листа и индекс текущего, то нужно сменить
    if (this.asc_getActiveWorksheetIndex() !== sheetIndex) {
      // Меняем активный лист
      this.asc_showWorksheet(sheetIndex);
    }

    ws = this.wb.getWorksheet();
    ws.setSelection(1 === d.length ? d[0].getBBox0() : d);
  };

	spreadsheet_api.prototype.asc_closeCellEditor = function (cancel) {
		var result = true;
		if (this.wb) {
			this.wb.setWizardMode(false);
			result = this.wb.closeCellEditor(cancel);
		}
		return result;
	};

	spreadsheet_api.prototype.asc_setR1C1Mode = function (value) {
		AscCommonExcel.g_R1C1Mode = value;
		if (this.wbModel) {
			var trueNeedUpdateTarget = this.wb.NeedUpdateTargetForCollaboration;
			this._onUpdateAfterApplyChanges();
			this.wb._onUpdateSelectionName(true);
			this.wb.NeedUpdateTargetForCollaboration = trueNeedUpdateTarget;
        }
	};

	spreadsheet_api.prototype.asc_SetAutoCorrectHyperlinks = function (value) {
		window['AscCommonExcel'].g_AutoCorrectHyperlinks = value;
	};

    spreadsheet_api.prototype.asc_setIncludeNewRowColTable = function (value) {
      window['AscCommonExcel'].g_IncludeNewRowColInTable = value;
    };

  // Spreadsheet interface

  spreadsheet_api.prototype.asc_getColumnWidth = function() {
    var ws = this.wb.getWorksheet();
    return ws.getSelectedColumnWidthInSymbols();
  };

  spreadsheet_api.prototype.asc_setColumnWidth = function(width) {
    this.wb.getWorksheet().changeWorksheet("colWidth", width);
  };

  spreadsheet_api.prototype.asc_showColumns = function() {
    this.wb.getWorksheet().changeWorksheet("showCols");
  };

  spreadsheet_api.prototype.asc_hideColumns = function() {
    this.wb.getWorksheet().changeWorksheet("hideCols");
  };

  spreadsheet_api.prototype.asc_autoFitColumnWidth = function() {
    this.wb.getWorksheet().autoFitColumnsWidth(null);
  };

  spreadsheet_api.prototype.asc_getRowHeight = function() {
    var ws = this.wb.getWorksheet();
    return ws.getSelectedRowHeight();
  };

  spreadsheet_api.prototype.asc_setRowHeight = function(height) {
    this.wb.getWorksheet().changeWorksheet("rowHeight", height);
  };
  spreadsheet_api.prototype.asc_autoFitRowHeight = function() {
    this.wb.getWorksheet().autoFitRowHeight(null);
  };

  spreadsheet_api.prototype.asc_showRows = function() {
    this.wb.getWorksheet().changeWorksheet("showRows");
  };

  spreadsheet_api.prototype.asc_hideRows = function() {
    this.wb.getWorksheet().changeWorksheet("hideRows");
  };

  spreadsheet_api.prototype.asc_group = function(val) {
    if(val) {
		this.wb.getWorksheet().changeWorksheet("groupRows");
    } else {
		this.wb.getWorksheet().changeWorksheet("groupCols");
    }
  };

  spreadsheet_api.prototype._canGroupPivot = function () {
    var ws = this.wbModel.getActiveWs();
    var activeCell = ws.selectionRange.activeCell;
    var pivotTable = ws.getPivotTable(activeCell.col, activeCell.row);
    if (pivotTable && ws.selectionRange.inContains(pivotTable.getReportRanges())) {
      var layout = pivotTable.getLayoutsForGroup(ws.selectionRange);
      if (null !== layout.fld) {
        return {pivotTable: pivotTable, layout: layout};
      }
    }
    return null;
  };
  spreadsheet_api.prototype.asc_canGroupPivot = function () {
    return !!this._canGroupPivot();
  };
  spreadsheet_api.prototype._groupPivot = function (confirmation, onRepeat, opt_rangePr, opt_dateTypes) {
    var canGroupRes = this._canGroupPivot();
    if(canGroupRes) {
      canGroupRes.pivotTable.groupPivot(this, canGroupRes.layout, confirmation, onRepeat, opt_rangePr, opt_dateTypes);
    }
  };
  spreadsheet_api.prototype.asc_groupPivot = function (opt_rangePr, opt_dateTypes) {
    var t = this;
    var onRepeat = function(){
      t._groupPivot(true, onRepeat, opt_rangePr, opt_dateTypes);
    };
    this._groupPivot(false, onRepeat, opt_rangePr, opt_dateTypes);
  };
  spreadsheet_api.prototype._ungroupPivot = function (confirmation, onRepeat) {
    var canGroupRes = this._canGroupPivot();
    if(canGroupRes) {
      canGroupRes.pivotTable.ungroupPivot(this, canGroupRes.layout, confirmation, onRepeat);
    }
  };
  spreadsheet_api.prototype.asc_ungroupPivot = function () {
    var t = this;
    var onRepeat = function(){
      t._ungroupPivot(true, onRepeat);
    };
    this._ungroupPivot(false, onRepeat);
  };

  spreadsheet_api.prototype.asc_ungroup = function(val) {
    if(val) {
        this.wb.getWorksheet().changeWorksheet("groupRows", true);
    } else {
        this.wb.getWorksheet().changeWorksheet("groupCols", true);
    }
  };

  spreadsheet_api.prototype.asc_checkAddGroup = function(bUngroup) {
	  //true - rows, false - columns, null - show dialog, undefined - error
    return this.wb.getWorksheet().checkAddGroup(bUngroup);
  };

  spreadsheet_api.prototype.asc_clearOutline = function() {
	  this.wb.getWorksheet().changeWorksheet("clearOutline");
  };

  spreadsheet_api.prototype.asc_changeGroupDetails = function(bExpand) {
	  this.wb.getWorksheet().changeGroupDetails(bExpand);
  };

  spreadsheet_api.prototype.asc_insertCells = function(options) {
    this.wb.getWorksheet().changeWorksheet("insCell", options);
  };

  spreadsheet_api.prototype.asc_deleteCells = function(options) {
    this.wb.getWorksheet().changeWorksheet("delCell", options);
  };

  spreadsheet_api.prototype.asc_mergeCells = function(options) {
    this.wb.getWorksheet().setSelectionInfo("merge", options);
  };

  spreadsheet_api.prototype.asc_sortCells = function(options) {
    this.wb.getWorksheet().setSelectionInfo("sort", options);
  };

  spreadsheet_api.prototype.asc_emptyCells = function(options, isMineComments) {
    //TODO isMineComments - временный флаг, как только в сдк появится класс для групп, добавить этот флаг туда
  	this.wb.emptyCells(options, isMineComments);
  };

  spreadsheet_api.prototype.asc_drawDepCells = function(se) {
    /* ToDo
     if( se != AscCommonExcel.c_oAscDrawDepOptions.Clear )
     this.wb.getWorksheet().prepareDepCells(se);
     else
     this.wb.getWorksheet().cleanDepCells();*/
  };

  // Потеряем ли мы что-то при merge ячеек
  spreadsheet_api.prototype.asc_mergeCellsDataLost = function(options) {
    return this.wb.getWorksheet().getSelectionMergeInfo(options);
  };

  //нужно ли спрашивать пользователя о расширении диапазона
  spreadsheet_api.prototype.asc_sortCellsRangeExpand = function() {
    return this.wb.getWorksheet().getSelectionSortInfo();
  };

  spreadsheet_api.prototype.asc_getSheetViewSettings = function() {
    return this.wb.getWorksheet().getSheetViewSettings();
  };

	spreadsheet_api.prototype.asc_setDisplayGridlines = function (value) {
		this.wb.getWorksheet().changeSheetViewSettings(AscCH.historyitem_Worksheet_SetDisplayGridlines, value);
	};

	spreadsheet_api.prototype.asc_setDisplayHeadings = function (value) {
		this.wb.getWorksheet().changeSheetViewSettings(AscCH.historyitem_Worksheet_SetDisplayHeadings, value);
	};

	spreadsheet_api.prototype.asc_setShowZeros = function (value) {
		this.wb.getWorksheet().changeSheetViewSettings(AscCH.historyitem_Worksheet_SetShowZeros, value);
	};

	spreadsheet_api.prototype.asc_setShowFormulas = function (value) {
		this.wb.getWorksheet().changeSheetViewSettings(AscCH.historyitem_Worksheet_SetShowFormulas, value);
	};

	spreadsheet_api.prototype.asc_getShowFormulas = function () {
		let ws = this.wb.getWorksheet();
		return ws.model && ws.model.getShowFormulas();
	};

	spreadsheet_api.prototype.asc_setDate1904 = function (value) {
		this.wb.setDate1904(value);
	};

	spreadsheet_api.prototype.asc_getDate1904 = function () {
		return AscCommon.bDate1904;
	};

  // Images & Charts

  spreadsheet_api.prototype.asc_drawingObjectsExist = function() {
    for (var i = 0; i < this.wb.model.aWorksheets.length; i++) {
      if (this.wb.model.aWorksheets[i].Drawings && this.wb.model.aWorksheets[i].Drawings.length) {
        return true;
      }
    }
    return false;
  };

	spreadsheet_api.prototype._createSmartArt = function (oSmartArt, oPlaceholder)
	{
		const oController = this.getGraphicController();
		const oDrawingObjects = this.getDrawingObjects();
		if (!oDrawingObjects || !oController)
		{
			return;
		}
		oSmartArt.setDrawingObjects(oDrawingObjects);
		const oWS = oDrawingObjects.getWorksheet();
		const oWSModel = oWS.model;
		oSmartArt.setWorksheet(oWSModel);
		oSmartArt.addToDrawingObjects(undefined, AscCommon.c_oAscCellAnchorType.cellanchorTwoCell);
		oSmartArt.checkDrawingBaseCoords();
		oSmartArt.fitFontSize();
		oController.checkChartTextSelection();
		oController.resetSelection();
		oSmartArt.select(oController, 0);
		oWS.setSelectionShape(true);
		oController.startRecalculate();
		oDrawingObjects.sendGraphicObjectProps();
	};

  spreadsheet_api.prototype.asc_getChartObject = function(bNoLock) {		// Return new or existing chart. For image return null

    if(bNoLock !== true){
     this.asc_onOpenChartFrame();
   }
    var ws = this.wb.getWorksheet();
    return ws.objectRender.getAscChartObject(bNoLock);
  };

  spreadsheet_api.prototype.asc_addChartDrawingObject = function(chart) {
    var ws = this.wb.getWorksheet();
    if (ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.objects)) {
      this.asc_onCloseChartFrame();
      return false;
    }

    AscFonts.IsCheckSymbols = true;
    var ret = ws.objectRender.addChartDrawingObject(chart);
    AscFonts.IsCheckSymbols = false;
    this.asc_onCloseChartFrame();
    return ret;
  };

    spreadsheet_api.prototype.getScaleCoefficientsForOleTableImage = function (nImageWidth, nImageHeight) {
      const oThis = this;
      return this.wb._executeWithoutZoom(function () {
        const oWorksheet = oThis.wb.getWorksheet();
        const oSingleChart = oWorksheet.isHaveOnlyOneChart(true);
        let oRangeSizes = {};
        if (oSingleChart) {
          oRangeSizes = {
            width: oSingleChart.extX * AscCommon.g_dKoef_mm_to_pix,
            height: oSingleChart.extY * AscCommon.g_dKoef_mm_to_pix
          };
        } else {
          const oOleSize = oThis.wbModel.getOleSize().getLast();
          if (oOleSize) {
            oRangeSizes = oWorksheet.getRangePosition(oOleSize);
          }
        }
        if (oRangeSizes.width && oRangeSizes.height) {
          return {
            widthCoefficient: nImageWidth / oRangeSizes.width,
            heightCoefficient: nImageHeight / oRangeSizes.height
          };
        }
        return {widthCoefficient: 1, heightCoefficient: 1};
      });
    };

  /**
   * Loading ole editor
   * @param {{}} [oOleObjectInfo] info from oleObject
   */
  spreadsheet_api.prototype.asc_addTableOleObjectInOleEditor = function(oOleObjectInfo) {
      this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
      // на случай, если изображение поставили на загрузку, закрыли редактор, и потом опять открыли
      this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.LoadImage);
      this.sendFromFrameToGeneralEditor({
          "type": AscCommon.c_oAscFrameDataType.OpenFrame
      });
			let bIsCreatingOleObject = false;
			if (!oOleObjectInfo)
			{
				oOleObjectInfo = {"binary": AscCommon.getEmpty()};
				bIsCreatingOleObject = true;
			}

      const sStream = oOleObjectInfo["binary"];
      const oThis = this;
      const oFile = new AscCommon.OpenFileResult();
      oFile.bSerFormat = AscCommon.checkStreamSignature(sStream, AscCommon.c_oSerFormat.Signature);
      oFile.data = sStream;
      this.isEditOleMode = true;
      this.isChartEditor = false;
      this.isFromSheetEditor = oOleObjectInfo["isFromSheetEditor"];
      const oDocumentImageUrls = oOleObjectInfo["documentImageUrls"];
      this.asc_CloseFile();
      this.fAfterLoad = function () {
          const nImageWidth = oOleObjectInfo["imageWidth"];
          const nImageHeight = oOleObjectInfo["imageHeight"];
          if (nImageWidth && nImageHeight) {
              oThis.saveImageCoefficients = oThis.getScaleCoefficientsForOleTableImage(nImageWidth, nImageHeight);
          }
					if (bIsCreatingOleObject)
					{
						AscFormat.ExecuteNoHistory(function ()
						{
							const oFirstWorksheet = oThis.wbModel.getWorksheet(0);
							if (oFirstWorksheet)
							{
								oFirstWorksheet.sName = '';
								const sName = oThis.wbModel.getUniqueSheetNameFrom(AscCommon.translateManager.getValue(AscCommonExcel.g_sNewSheetNamePattern), false);
								oFirstWorksheet.setName(sName);
								oThis.sheetsChanged();
							}
						}, oThis, []);
					}
          oThis.wb.scrollToOleSize();
          // добавляем первый поинт после загрузки, чтобы в локальную историю добавился либо стандартный oleSize, либо заданный пользователем
          const oleSize = oThis.wb.getOleSize();
          oleSize.addPointToLocalHistory();

          oThis.wb.onOleEditorReady();
          oThis.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Open);
      }

      this.imagesFromGeneralEditor = oDocumentImageUrls || {};
      this.openDocument(oFile);
    };
  /**
   * get binary info about changed ole object
   * @returns {{}} binary info about oleObject
   */
  spreadsheet_api.prototype.asc_getBinaryInfoOleObject = function () {
    const sDataUrl = this.wb.getImageFromTableOleObject();
    const oBinaryFileWriter = new AscCommonExcel.BinaryFileWriter(this.wbModel);
    const arrBinaryData = oBinaryFileWriter.Write().split(';');
    const sCleanBinaryData = arrBinaryData[arrBinaryData.length - 1];
    const oBinaryInfo = {};
    const arrRasterImageIds = [];
    const arrWorksheetLength = this.wbModel.aWorksheets.length;
    for (let i = 0; i < arrWorksheetLength; i += 1) {
      const oWorksheet = this.wbModel.aWorksheets[i];
      const arrDrawings = oWorksheet.Drawings;
      if (arrDrawings) {
        for (let j = 0; j < arrDrawings.length; j += 1) {
          const oDrawing = arrDrawings[j];
          oDrawing.graphicObject.getAllRasterImages(arrRasterImageIds);
        }
      }
    }
    const urlsForAddToHistory = [];
    for (let i = 0; i < arrRasterImageIds.length; i += 1) {
      const url = AscCommon.g_oDocumentUrls.mediaPrefix + arrRasterImageIds[i];
      if (!(this.imagesFromGeneralEditor && this.imagesFromGeneralEditor[url] && this.imagesFromGeneralEditor[url] === AscCommon.g_oDocumentUrls.getUrls()[url])) {
        urlsForAddToHistory.push(arrRasterImageIds[i]);
      }
    }


    oBinaryInfo["binary"] = sCleanBinaryData;
    oBinaryInfo["base64Image"] = sDataUrl;
    oBinaryInfo["isFromSheetEditor"] = this.isFromSheetEditor;
    oBinaryInfo["imagesForAddToHistory"] = urlsForAddToHistory;
    if (this.saveImageCoefficients) {
      oBinaryInfo["widthCoefficient"] = this.saveImageCoefficients.widthCoefficient;
      oBinaryInfo["heightCoefficient"] = this.saveImageCoefficients.heightCoefficient;
      delete this.saveImageCoefficients;
    }

    return oBinaryInfo;
  }

	spreadsheet_api.prototype.asc_toggleChangeVisibleAreaOleEditor = function (bForceValue) {
		const ws = this.wb.getWorksheet();
		ws.cleanSelection();
		ws.endEditChart();
		ws._endSelectionShape();
        const previousValue = this.isEditVisibleAreaOleEditor;
		if (typeof bForceValue === 'boolean') {
			this.isEditVisibleAreaOleEditor = bForceValue;
		} else {
			this.isEditVisibleAreaOleEditor = !this.isEditVisibleAreaOleEditor;
		}
		if (this.isEditVisibleAreaOleEditor) {
			this.wb.setCellEditMode(false);
		}
        const currentValue = this.isEditVisibleAreaOleEditor;
        if (previousValue === true && currentValue === false) {
            const oOleSize = this.wbModel.getOleSize();
            oOleSize.addToGlobalHistory();
        }
		ws._drawSelection();
	};

	spreadsheet_api.prototype.asc_toggleShowVisibleAreaOleEditor = function (bForceValue) {
		var newValue;
		if (typeof bForceValue === 'boolean') {
			newValue = bForceValue;
		} else {
			newValue = !this.isShowVisibleAreaOleEditor;
		}
		var ws = this.wb.getWorksheet();
		ws.cleanSelection();
		if (newValue) {
			this.isShowVisibleAreaOleEditor = newValue;
			ws.cleanSelection();
			ws._drawSelection();
		} else {
			ws.cleanSelection();
			this.isShowVisibleAreaOleEditor = newValue;
			ws._drawSelection();
		}
	};

  spreadsheet_api.prototype.asc_editChartDrawingObject = function(chart) {
    var ws = this.wb.getWorksheet();
    var ret = ws.objectRender.editChartDrawingObject(chart);
    this.asc_onCloseChartFrame();
    return ret;
  };

  spreadsheet_api.prototype.asc_addImageDrawingObject = function (urls, imgProp, token) {

    var t = this;
    var ws = t.wb.getWorksheet();
    if (ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.objects)) {
      return false;
    }
    this.AddImageUrl(urls, imgProp, token, null);
  };


  spreadsheet_api.prototype.asc_AddMath = function(Type)
  {
    var t = this, fonts = {};
    fonts["Cambria Math"] = 1;
    t._loadFonts(fonts, function() {t.asc_AddMath2(Type);});
  };

  spreadsheet_api.prototype.asc_AddMath2 = function(Type)
  {
    var ws = this.wb.getWorksheet();
    ws.objectRender.addMath(Type);
  };
  spreadsheet_api.prototype.asc_ConvertMathView = function(isToLinear, isAll)
  {
    var ws = this.wb.getWorksheet();
    ws.objectRender.convertMathView(isToLinear, isAll);
  };

  spreadsheet_api.prototype.asc_SetMathProps = function(MathProps)
  {
    var ws = this.wb.getWorksheet();
    ws.objectRender.setMathProps(MathProps);
  };

  spreadsheet_api.prototype.asc_showImageFileDialog = function() {
    // ToDo заменить на общую функцию для всех
    this.asc_addImage();
  };
  spreadsheet_api.prototype._addImageUrl = function(arrUrls, oOptionObject) {
    if (oOptionObject && oOptionObject.sendUrlsToFrameEditor && this.isOpenedChartFrame) {
      this.addImageUrlsFromGeneralToFrameEditor(arrUrls);
      return;
    }
    const oWS = this.wb.getWorksheet();
    if (oWS) {
      if (oOptionObject) {
        if (oOptionObject.isImageChangeUrl || oOptionObject.isShapeImageChangeUrl || oOptionObject.isTextArtChangeUrl || oOptionObject.fAfterUploadOleObjectImage) {
          oWS.objectRender.editImageDrawingObject(arrUrls[0], oOptionObject);
        } else {
          if (this.ImageLoader) {
            const oApi = this;
            this.ImageLoader.LoadImagesWithCallback(arrUrls, function() {
               const arrImages = [];
              for(let i = 0; i < arrUrls.length; ++i) {
                const oImage = oApi.ImageLoader.LoadImage(arrUrls[i], 1);
                if(oImage){
                  arrImages.push(oImage);
                }
              }
              oApi.wbModel.addImages(arrImages, oOptionObject);
            }, []);
          }
        }
      } else {
        oWS.objectRender.addImageDrawingObject(arrUrls, null);
      }
    }
  };
    // signatures
    spreadsheet_api.prototype.asc_addSignatureLine = function (oPr, Width, Height, sImgUrl) {
      var ws = this.wb && this.wb.getWorksheet();
      if (ws && ws.model && ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.objects)) {
        return false;
      }

      if(ws && ws.objectRender){
          ws.objectRender.addSignatureLine(oPr, Width, Height, sImgUrl);
      }
    };

    spreadsheet_api.prototype.asc_getAllSignatures = function(){
      var ret = [];
      var aSpTree = [];
	  if (this.wbModel) {
		  this.wbModel.forEach(function (ws) {
			  for (var j = 0; j < ws.Drawings.length; ++j) {
				  aSpTree.push(ws.Drawings[j].graphicObject);
			  }
		  });
	  }
      AscFormat.DrawingObjectsController.prototype.getAllSignatures2(ret, aSpTree);
      return ret;
    };
    spreadsheet_api.prototype.getSignatureLineSp = function(sGuid) {
        var ret = [];
        var aSpTree = [];
		if (this.wbModel) {
			this.wbModel.forEach(function (ws) {
				for (var j = 0; j < ws.Drawings.length; ++j) {
					aSpTree.push(ws.Drawings[j].graphicObject);
				}
			});
		}
        AscFormat.DrawingObjectsController.prototype.getAllSignatures2(ret, aSpTree);
        for(var i = 0; i < aSpTree.length; ++i){
            if(aSpTree[i].signatureLine && aSpTree[i].signatureLine.id === sGuid){
                return aSpTree[i];
            }
        }
        return null;
    };
    spreadsheet_api.prototype.asc_CallSignatureDblClickEvent = function(sGuid){
        var oSp = this.getSignatureLineSp(sGuid);
        if(oSp){
            this.sendEvent("asc_onSignatureDblClick", sGuid, oSp.extX, oSp.extY);
        }
    };
    spreadsheet_api.prototype.gotoSignatureInternal = function(sGuid){
        var oSp = this.getSignatureLineSp(sGuid);
        if(oSp && !oSp.group){
            var oWorksheet = oSp.worksheet;
            if(oWorksheet) {
                var nSheetIdx = oWorksheet.getIndex();
                if (this.asc_getActiveWorksheetIndex() !== nSheetIdx) {
                    this.asc_showWorksheet(nSheetIdx);
                }
                var oWSView = this.wb.getWorksheet();
                if(oWSView) {
                    var oRender = oWSView.objectRender;
                    if(oRender) {
                        var oController = oRender.controller;
                        if(oController) {
                            oSp.Set_CurrentElement(false);
                            oController.selection.textSelection = null;
                            oWSView.setSelectionShape(true);
                            oController.updateSelectionState();
                            oController.updateOverlay();
                            oRender.sendGraphicObjectProps();
                        }
                    }
                }
            }
        }
    };
    //-------------------------------------------------------

    spreadsheet_api.prototype.asc_getCurrentDrawingMacrosName = function() {
        var ws = this.wb.getWorksheet();
        return ws.objectRender.getCurrentDrawingMacrosName();
    };
    spreadsheet_api.prototype.asc_assignMacrosToCurrentDrawing = function(sName) {
        var ws = this.wb.getWorksheet();
        var sGuid = this.asc_getMacrosGuidByName(sName);
        return ws.objectRender.assignMacrosToCurrentDrawing(sGuid);
    };
    spreadsheet_api.prototype.asc_setSelectedDrawingObjectLayer = function(layerType) {
        var ws = this.wb.getWorksheet();
        return ws.objectRender.setGraphicObjectLayer(layerType);
    };

  spreadsheet_api.prototype.asc_setSelectedDrawingObjectAlign = function(alignType) {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.setGraphicObjectAlign(alignType);
  };
  spreadsheet_api.prototype.asc_DistributeSelectedDrawingObjectHor = function() {
      var ws = this.wb.getWorksheet();
      return ws.objectRender.distributeGraphicObjectHor();
  };

  spreadsheet_api.prototype.asc_DistributeSelectedDrawingObjectVer = function() {
      var ws = this.wb.getWorksheet();
      return ws.objectRender.distributeGraphicObjectVer();
  };

  spreadsheet_api.prototype.asc_getSelectedDrawingObjectsCount = function()
  {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.getSelectedDrawingObjectsCount();
  };

  spreadsheet_api.prototype.asc_canEditCrop = function()
  {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.controller.canStartImageCrop();
  };

  spreadsheet_api.prototype.asc_startEditCrop = function()
  {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.controller.startImageCrop();
  };

  spreadsheet_api.prototype.asc_endEditCrop = function()
  {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.controller.endImageCrop();
  };

  spreadsheet_api.prototype.asc_cropFit = function()
  {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.controller.cropFit();
  };

  spreadsheet_api.prototype.asc_cropFill = function()
  {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.controller.cropFill();
  };


  spreadsheet_api.prototype.asc_addTextArt = function(nStyle) {
    var ws = this.wb.getWorksheet();
    if (ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.objects)) {
      return false;
    }

    return ws.objectRender.addTextArt(nStyle);
  };

  spreadsheet_api.prototype.asc_checkDataRange = function(dialogType, dataRange, fullCheck, isRows, chartType) {
    return parserHelp.checkDataRange(this.wbModel, this.wb, dialogType, dataRange, fullCheck, isRows, chartType);
  };

  // Для вставки диаграмм в Word
  spreadsheet_api.prototype.asc_getBinaryFileWriter = function() {
    return new AscCommonExcel.BinaryFileWriter(this.wbModel);
  };

  spreadsheet_api.prototype.asc_getWordChartObject = function() {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.getWordChartObject();
  };

  spreadsheet_api.prototype.asc_cleanWorksheet = function() {
    var ws = this.wb.getWorksheet();	// Для удаления данных листа и диаграмм
    if (ws.objectRender) {
      ws.objectRender.cleanWorksheet();
    }
  };

  // Выставление данных (пока используется только для MailMerge)
  spreadsheet_api.prototype.asc_setData = function(oData) {
    this.wb.getWorksheet().setData(oData);
  };
  // Получение данных
  spreadsheet_api.prototype.asc_getData = function() {
    this.asc_closeCellEditor();
    return this.wb.getWorksheet().getData();
  };

  // Cell comment interface
	spreadsheet_api.prototype.asc_addComment = function (oComment) {
		if (this.collaborativeEditing.getGlobalLock() || (!this.canEdit() && !this.isRestrictionComments())) {
			return false;
		}
		let oPlace = oComment.bDocument ? this.wb : (this.wb && this.wb.getWorksheet());
		oPlace && oPlace.cellCommentator.addComment(oComment);
	};

	spreadsheet_api.prototype.asc_changeComment = function (id, oComment) {
		if (this.wb) {
			if (oComment.bDocument) {
				this.wb.cellCommentator.changeComment(id, oComment);
			} else {
				var ws = this.wb.getWorksheet();
				ws.cellCommentator.changeComment(id, oComment);
			}
		}
	};

	spreadsheet_api.prototype.asc_selectComment = function (id) {
		this.wb && this.wb.getWorksheet().cellCommentator.selectComment(id);
	};

	spreadsheet_api.prototype.asc_showComment = function (id, bNew) {
		if (this.wb) {
			let ws = this.wb.getWorksheet();
			ws.cellCommentator.showCommentById(id, bNew);
		}
	};

	spreadsheet_api.prototype.asc_findComment = function (id) {
		let ws = this.wb && this.wb.getWorksheet();
		return ws ? ws.cellCommentator.findComment(id) : null;
	};

	spreadsheet_api.prototype.asc_removeComment = function (id) {
		if (this.wb) {
			this.wb.removeComment(id);
		}
	};

	spreadsheet_api.prototype.asc_RemoveAllComments = function (isMine, isCurrent) {
		if (this.collaborativeEditing.getGlobalLock() || (!this.canEdit() && !this.isRestrictionComments())) {
			return;
		}
		if (this.wb) {
			this.wb.removeAllComments(isMine, isCurrent);
		}
	};

	spreadsheet_api.prototype.asc_GetCommentLogicPositionv = function (sId) {
		return -1;
	};

	spreadsheet_api.prototype.asc_ResolveAllComments = function (isMine, isCurrent, arrIds) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		if (this.wb) {
			this.wb.resolveAllComments(isMine, isCurrent);
		}
	};

	spreadsheet_api.prototype.asc_showComments = function (isShowSolved) {
		if (this.wb) {
			this.wb.showComments(true, isShowSolved);
		}
	};
	spreadsheet_api.prototype.asc_hideComments = function () {
		if (this.wb) {
			this.wb.showComments(false, false);
		}
	};

  // Shapes
  spreadsheet_api.prototype.setStartPointHistory = function() {
    this.noCreatePoint = true;
    this.exucuteHistory = true;
    this.asc_stopSaving();
  };

  spreadsheet_api.prototype.setEndPointHistory = function() {
    this.noCreatePoint = false;
    this.exucuteHistoryEnd = true;
    this.asc_continueSaving();
  };

  spreadsheet_api.prototype.asc_startAddShape = function(sPreset) {
    var ws = this.wb.getWorksheet();
    if (ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.objects)) {
      this.asc_endAddShape();
      return false;
    }
	this.stopInkDrawer();
	this.cancelEyedropper();
    this.isStartAddShape = this.controller.isShapeAction = true;
    ws.objectRender.controller.startTrackNewShape(sPreset);
  };

  spreadsheet_api.prototype.asc_endAddShape = function() {
    this.isStartAddShape = false;
    this.handlers.trigger("asc_onEndAddShape");
  };

  spreadsheet_api.prototype.asc_doubleClickOnTableOleObject = function (obj) {
    this.isOleEditor = true;	// Для совместного редактирования
    this.asc_onOpenChartFrame();
    // console.log(editor.WordControl)
    // if(!window['IS_NATIVE_EDITOR']) {
    //   this.WordControl.onMouseUpMainSimple();
    // }
    if(this.handlers.hasTrigger("asc_doubleClickOnTableOleObject"))
    {
      this.sendEvent("asc_doubleClickOnTableOleObject", obj);
    }
    else
    {
      this.sendEvent("asc_doubleClickOnChart", obj); // TODO: change event type
    }
  };

    spreadsheet_api.prototype.asc_canEditGeometry = function () {
        var ws = this.wb.getWorksheet();
        if(ws && ws.objectRender && ws.objectRender.controller) {
            return ws.objectRender.controller.canEditGeometry();
        }
        return false;
    };

    spreadsheet_api.prototype.asc_editPointsGeometry = function() {
        var ws = this.wb.getWorksheet();
        if(ws && ws.objectRender && ws.objectRender.controller) {
            ws.objectRender.controller.startEditGeometry();
        }
    };

    spreadsheet_api.prototype.asc_addShapeOnSheet = function(sPreset) {
        if(this.wb){
          var ws = this.wb.getWorksheet();
          if (ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.objects)) {
            return false;
          }

          if(ws && ws.objectRender){
            ws.objectRender.addShapeOnSheet(sPreset);
          }
        }
    };

  spreadsheet_api.prototype.asc_addOleObjectAction = function(sLocalUrl, sData, sApplicationId, fWidth, fHeight, nWidthPix, nHeightPix, bSelect, arrImagesForAddToHistory)
  {
    var _image = this.ImageLoader.LoadImage(AscCommon.getFullImageSrc2(sLocalUrl), 1);
    if (null != _image){
        var ws = this.wb.getWorksheet();
        if (ws.model.getSheetProtection(Asc.c_oAscSheetProtectType.objects)) {
          return false;
        }

      if(ws.objectRender){
        this.asc_canPaste();
        ws.objectRender.addOleObject(fWidth, fHeight, nWidthPix, nHeightPix, sLocalUrl, sData, sApplicationId, bSelect, arrImagesForAddToHistory);
        this.asc_endPaste();
      }
    }
  };

  spreadsheet_api.prototype.asc_editOleObjectAction = function(oOleObject, sImageUrl, sData, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory)
  {
    if (oOleObject)
    {
      var ws = this.wb.getWorksheet();
      if(ws.objectRender){
        this.asc_canPaste();
        ws.objectRender.editOleObject(oOleObject, sData, sImageUrl, fWidth, fHeight, nPixWidth, nPixHeight, arrImagesForAddToHistory);
        this.asc_endPaste();
      }
    }
  };


    spreadsheet_api.prototype.asc_startEditCurrentOleObject = function(){
        var ws = this.wb.getWorksheet();
        if(ws && ws.objectRender){
            ws.objectRender.startEditCurrentOleObject();
        }
    };

  spreadsheet_api.prototype.asc_isAddAutoshape = function() {
    return this.isStartAddShape;
  };

	spreadsheet_api.prototype.onInkDrawerChangeState = function() {
		const oController = this.getGraphicController();
		if(!oController) {
			return;
		}
        if (this.isStartAddShape) {
            this.asc_endAddShape();
        }
        this.cancelEyedropper();

        if (this.isFormatPainterOn())
        {
            this.formatPainter.putState(AscCommon.c_oAscFormatPainterState.kOff);
            if (this.wb) {
                this.wb.formatPainter(AscCommon.c_oAscFormatPainterState.kOff, undefined);
            }
        }
		oController.onInkDrawerChangeState();
	};

  spreadsheet_api.prototype.asc_canAddShapeHyperlink = function() {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.controller.canAddHyperlink();
  };

  spreadsheet_api.prototype.asc_canGroupGraphicsObjects = function() {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.controller.canGroup();
  };

  spreadsheet_api.prototype.asc_groupGraphicsObjects = function() {
    var ws = this.wb.getWorksheet();
    ws.objectRender.groupGraphicObjects();
  };

  spreadsheet_api.prototype.asc_canUnGroupGraphicsObjects = function() {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.controller.canUnGroup();
  };

  spreadsheet_api.prototype.asc_unGroupGraphicsObjects = function() {
    var ws = this.wb.getWorksheet();
    ws.objectRender.unGroupGraphicObjects();
  };

  spreadsheet_api.prototype.asc_changeShapeType = function(value) {
    this.asc_setGraphicObjectProps(new Asc.asc_CImgProperty({ShapeProperties: {type: value}}));
  };

  spreadsheet_api.prototype.asc_getGraphicObjectProps = function() {
    var ws = this.wb.getWorksheet();
    if (ws && ws.objectRender && ws.objectRender.controller) {
      return ws.objectRender.controller.getGraphicObjectProps();
    }
    return null;
  };
  spreadsheet_api.prototype.asc_GetSelectedText = function(bClearText, select_Pr) {
    bClearText = typeof(bClearText) === "boolean" ? bClearText : false;
    var ws = this.wb.getWorksheet();
    if (this.wb.getCellEditMode()) {
      var fragments = this.wb.cellEditor.copySelection();
      if (null !== fragments) {
        return AscCommonExcel.getFragmentsText(fragments);
      }
    }
    if (ws && ws.objectRender && ws.objectRender.controller) {
      return ws.objectRender.controller.GetSelectedText(bClearText, select_Pr);
    }
    return "";
  };

  spreadsheet_api.prototype.asc_setGraphicObjectProps = function(props) {
	if (!this.canEdit()) {
	  return;
	}
    var ws = this.wb.getWorksheet();
    var fReplaceCallback = null, sImageUrl = null, sToken = undefined;
    if(!AscCommon.isNullOrEmptyString(props.ImageUrl)){
      if(!g_oDocumentUrls.getImageLocal(props.ImageUrl)){
        sImageUrl = props.ImageUrl;
        sToken = props.Token;
        fReplaceCallback = function(sLocalUrl){
          props.ImageUrl = sLocalUrl;
        }
      }
    }
    else if(props.ShapeProperties && props.ShapeProperties.fill && props.ShapeProperties.fill.fill &&
    !AscCommon.isNullOrEmptyString(props.ShapeProperties.fill.fill.url)){
      if(!g_oDocumentUrls.getImageLocal(props.ShapeProperties.fill.fill.url)){
        sImageUrl = props.ShapeProperties.fill.fill.url;
        sToken = props.ShapeProperties.fill.fill.token;
        fReplaceCallback = function(sLocalUrl){
          props.ShapeProperties.fill.fill.url = sLocalUrl;
        }
      }
    }
    else if(props.ShapeProperties && props.ShapeProperties.textArtProperties &&
        props.ShapeProperties.textArtProperties.Fill && props.ShapeProperties.textArtProperties.Fill.fill &&
        !AscCommon.isNullOrEmptyString(props.ShapeProperties.textArtProperties.Fill.fill.url)){
      if(!g_oDocumentUrls.getImageLocal(props.ShapeProperties.textArtProperties.Fill.fill.url)){
        sImageUrl = props.ShapeProperties.textArtProperties.Fill.fill.url;
        sToken = props.ShapeProperties.textArtProperties.Fill.fill.token;
        fReplaceCallback = function(sLocalUrl){
          props.ShapeProperties.textArtProperties.Fill.fill.url = sLocalUrl;
        }
      }
    }
    if(fReplaceCallback) {

      if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]()) {
        var firstUrl = window["AscDesktopEditor"]["LocalFileGetImageUrl"](sImageUrl);
        firstUrl = g_oDocumentUrls.getImageUrl(firstUrl);
        fReplaceCallback(firstUrl);
        ws.objectRender.setGraphicObjectProps(props);
        return;
      }
      AscCommon.sendImgUrls(this, [sImageUrl], function (data) {

        if (data && data[0] && data[0].url !== "error") {
          fReplaceCallback(data[0].url);
          ws.objectRender.setGraphicObjectProps(props);
        }

      }, undefined, sToken);
    }
    else{
      var sBulletSymbol = props.asc_getBulletSymbol && props.asc_getBulletSymbol();
      var sBulletFont = props.asc_getBulletFont && props.asc_getBulletFont();
      if(typeof sBulletSymbol === "string" && sBulletSymbol.length > 0
      && typeof sBulletFont === "string" && sBulletFont.length > 0) {
        var t = this;
          var fonts = {};
          fonts[sBulletFont] = 1;
          AscFonts.FontPickerByCharacter.checkTextLight(sBulletSymbol);
          t._loadFonts(fonts, function() {
            ws.objectRender.setGraphicObjectProps(props);
          });
      }
      else {
        var oSlicerPr = props.SlicerProperties;
        var sForCheck = null;
        if(oSlicerPr && typeof oSlicerPr.caption === "string" && oSlicerPr.caption.length > 0) {
          sForCheck = oSlicerPr.caption;
        }
        if(typeof sForCheck === "string" && sForCheck.length > 0) {
          AscFonts.FontPickerByCharacter.checkText(sForCheck, this, function () {
            ws.objectRender.setGraphicObjectProps(props);
          });
        }
        else {
          ws.objectRender.setGraphicObjectProps(props);
        }
      }
    }
  };

  spreadsheet_api.prototype.asc_getOriginalImageSize = function() {
    var ws = this.wb.getWorksheet();
    return ws.objectRender.getOriginalImageSize();
  };

  spreadsheet_api.prototype.asc_setInterfaceDrawImagePlaceTextArt = function(elementId) {
    this.textArtElementId = elementId;
  };

  spreadsheet_api.prototype.asc_changeImageFromFile = function() {
    this.asc_addImage({isImageChangeUrl: true});
  };

  spreadsheet_api.prototype.asc_changeShapeImageFromFile = function(type) {
    this.asc_addImage({isShapeImageChangeUrl: true, textureType: type});
  };

  spreadsheet_api.prototype.asc_changeArtImageFromFile = function(type) {
    this.asc_addImage({isTextArtChangeUrl: true, textureType: type});
  };

  spreadsheet_api.prototype.getImageDataFromSelection = function() {
      var ws = this.wb.getWorksheet();
      return ws.objectRender.controller.getImageDataFromSelection();
  };
  spreadsheet_api.prototype.putImageToSelection = function(sImageSrc, nWidth, nHeight, replaceMode) {
      var ws = this.wb.getWorksheet();
      return ws.objectRender.controller.putImageToSelection(sImageSrc, nWidth, nHeight, replaceMode);
  };


	spreadsheet_api.prototype.getPluginContextMenuInfo = function () {
		const oWorksheet = this.wb.getWorksheet();
		if(oWorksheet){
			if(oWorksheet.isSelectOnShape){
				return oWorksheet.objectRender.controller.getPluginSelectionInfo();
			}
			else{
				return new AscCommon.CPluginCtxMenuInfo(Asc.c_oPluginContextMenuTypes.Selection);
			}
		}
		return new AscCommon.CPluginCtxMenuInfo();
	};

  spreadsheet_api.prototype.asc_putPrLineSpacing = function(type, value) {
    var ws = this.wb.getWorksheet();
    ws.objectRender.controller.putPrLineSpacing(type, value);
  };

  spreadsheet_api.prototype.asc_putLineSpacingBeforeAfter = function(type, value) { // "type == 0" means "Before", "type == 1" means "After"
    var ws = this.wb.getWorksheet();
    ws.objectRender.controller.putLineSpacingBeforeAfter(type, value);
  };

  spreadsheet_api.prototype.asc_setDrawImagePlaceParagraph = function(element_id, props) {
    var ws = this.wb.getWorksheet();
    ws.objectRender.setDrawImagePlaceParagraph(element_id, props);
  };

    spreadsheet_api.prototype.asc_replaceLoadImageCallback = function(fCallback){
        if(this.wb){
            var ws = this.wb.getWorksheet();
            if(ws.objectRender){
                ws.objectRender.asyncImageEndLoaded = fCallback;
            }
        }
    };

  spreadsheet_api.prototype.asyncImageEndLoaded = function(_image) {
    if (this.wb) {
      var ws = this.wb.getWorksheet();
      if (ws.objectRender.asyncImageEndLoaded) {
        ws.objectRender.asyncImageEndLoaded(_image);
      }
    }
  };

  spreadsheet_api.prototype.asyncImagesDocumentEndLoaded = function() {
    if (c_oAscAdvancedOptionsAction.None === this.advancedOptionsAction && this.wb && !window["NATIVE_EDITOR_ENJINE"]) {
      var ws = this.wb.getWorksheet();
      ws.objectRender.showDrawingObjects();
      ws.objectRender.controller.getGraphicObjectProps();
    }
  };

  spreadsheet_api.prototype.asyncImageEndLoadedBackground = function() {
    var worksheet = this.wb.getWorksheet();
    if (worksheet && worksheet.objectRender) {
      var drawing_area = worksheet.objectRender.drawingArea;
      if (drawing_area) {
        for (var i = 0; i < drawing_area.frozenPlaces.length; ++i) {
          worksheet.objectRender.showDrawingObjects();
            worksheet.objectRender.controller && worksheet.objectRender.controller.getGraphicObjectProps();
        }
      }
    }
  };

    // spellCheck
    spreadsheet_api.prototype.cleanSpelling = function (isCellEditing) {
      if (!this.spellcheckState.lockSpell) {
        if (this.spellcheckState.startCell) {
          var cellsChange = this.spellcheckState.cellsChange;
          var lastFindOptions = this.spellcheckState.lastFindOptions;
          if (cellsChange.length !== 0 && lastFindOptions && !isCellEditing) {
            this.asc_replaceMisspelledWords(lastFindOptions);
          }

          this.handlers.trigger("asc_onSpellCheckVariantsFound", null);
          this.spellcheckState.clean();
        } else {
          var ws = this.wb.getWorksheet();
          if (ws) {
            var maxC = ws.model.getColsCount() - 1;
            var maxR = ws.model.getRowsCount() - 1;
            if (-1 !== maxC || -1 !== maxR) {
              this.handlers.trigger("asc_onSpellCheckVariantsFound", null);
            }
          }
        }
      }
    };

    spreadsheet_api.prototype.SpellCheck_CallBack = function (e) {
      this.spellcheckState.lockSpell = false;
      var ws = this.wb.getWorksheet();
      var type = e["type"];
      var usrWords = e["usrWords"];
      var cellsInfo = e["cellsInfo"];
      var changeWords = this.spellcheckState.changeWords;
      var ignoreWords = this.spellcheckState.ignoreWords;
      var lastOptions = this.spellcheckState.lastFindOptions;
      var cellsChange = this.spellcheckState.cellsChange;
      var isStart = this.spellcheckState.isStart;

      if (type === "spell") {
        var usrCorrect = e["usrCorrect"];
        this.spellcheckState.wordsIndex = e["wordsIndex"];
        var lastIndex = this.spellcheckState.lastIndex;
        var activeCell = ws.model.selectionRange.activeCell;
        var isIgnoreUppercase = this.spellcheckState.isIgnoreUppercase;

        for (var i = 0; i < usrWords.length; i++) {
          if (this.spellcheckState.isIgnoreNumbers) {
            var isNumberInStr = /\d+/;
            if (usrWords[i].match(isNumberInStr)) {
              usrCorrect[i] = true;
            }
          }

          if (ignoreWords[usrWords[i]] || changeWords[usrWords[i]] || usrWords[i].length === 1
            || (isIgnoreUppercase && AscCommon.IsAbbreviation(usrWords[i]))) {
            usrCorrect[i] = true;
          }
        }

        while (!isStart && usrCorrect[lastIndex]) {
          ++lastIndex;
        }

        this.spellcheckState.lockSpell = false;
        this.spellcheckState.lastIndex = lastIndex;
        var currIndex = cellsInfo[lastIndex];
        if (currIndex && (currIndex.col !== activeCell.col || currIndex.row !== activeCell.row)) {
          var currentCellIsInactive = true;
        }
        while (isStart && currentCellIsInactive && usrCorrect[lastIndex]) {
          if (!cellsInfo[lastIndex + 1] || cellsInfo[lastIndex + 1].col !== cellsInfo[lastIndex].col) {
            var cell = cellsInfo[lastIndex];
            cellsChange.push(new Asc.Range(cell.col, cell.row, cell.col, cell.row));
          }
          ++lastIndex;
        }

        if (undefined === cellsInfo[lastIndex]) {
          this.spellcheckState.nextRow();
          this.asc_nextWord();
          return;
        }

        var cellStartIndex = 0;
        while (cellsInfo[cellStartIndex].col !== cellsInfo[lastIndex].col) {
          cellStartIndex++;
        }
        e["usrWords"].splice(0, cellStartIndex);
        e["usrCorrect"].splice(0, cellStartIndex);
        e["usrLang"].splice(0, cellStartIndex);
        e["cellsInfo"].splice(0, cellStartIndex);
        e["wordsIndex"].splice(0, cellStartIndex);

        lastIndex = lastIndex - cellStartIndex;
        this.spellcheckState.lastIndex = lastIndex;
        this.spellcheckState.lockSpell = true;
        this.spellcheckState.lastSpellInfo = e;

        this.SpellCheckApi.spellCheck({
          "type": "suggest",
          "usrWords": [e["usrWords"][lastIndex]],
          "usrLang": [e["usrLang"][lastIndex]],
          "cellInfo": e["cellsInfo"][lastIndex],
          "wordsIndex": e["wordsIndex"][lastIndex]
        });
      } else if (type === "suggest") {
        this.handlers.trigger("asc_onSpellCheckVariantsFound", new AscCommon.asc_CSpellCheckProperty(e["usrWords"][0], null, e["usrSuggest"][0], null, null));
        var cellInfo = e["cellInfo"];
        this.spellcheckState.lockSpell = true;

        if (!ws.model.selectionRange.activeCell.isEqual(cellInfo) && isStart && lastOptions) {
          this.asc_replaceMisspelledWords(lastOptions);
        } else {
          ws.setSelection(new Asc.Range(cellInfo.col, cellInfo.row, cellInfo.col, cellInfo.row));
          this.spellcheckState.lockSpell = false;
          if(this.spellcheckState.afterReplace) {
            History.EndTransaction();
            this.spellcheckState.afterReplace = false;
          }
        }
        this.spellcheckState.isStart = true;
        this.spellcheckState.wordsIndex = e["wordsIndex"];
      }
    };
  spreadsheet_api.prototype._spellCheckDisconnect = function () {
    this.cleanSpelling();
  };
  spreadsheet_api.prototype._spellCheckRestart = function (word) {
    var lastSpellInfo;
    if ((lastSpellInfo = this.spellcheckState.lastSpellInfo)) {
      var lastIndex = this.spellcheckState.lastIndex;
      this.spellcheckState.lastIndex = -1;

      var usrLang = [];
      for (var i = lastIndex; i < lastSpellInfo["usrWords"].length; ++i) {
        usrLang.push(this.defaultLanguage);
      }

      this.spellcheckState.lockSpell = true;
      this.SpellCheckApi.spellCheck({
        "type": "spell",
        "usrWords": lastSpellInfo["usrWords"].slice(lastIndex),
        "usrLang": usrLang,
        "cellsInfo": lastSpellInfo["cellsInfo"].slice(lastIndex),
        "wordsIndex": lastSpellInfo["wordsIndex"].slice(lastIndex)
      });
    } else {
      this.cleanSpelling();
    }
  };
  spreadsheet_api.prototype.asc_setDefaultLanguage = function (val) {
    if (this.spellcheckState.lockSpell || this.defaultLanguage === val) {
      return;
    }
    this.defaultLanguage = val;
    this._spellCheckRestart();
  };
  spreadsheet_api.prototype.asc_nextWord = function () {
    if (this.spellcheckState.lockSpell) {
      return;
    }

    var ws = this.wb.getWorksheet();
    var activeCell = ws.model.selectionRange.activeCell;
    var lastSpell = this.spellcheckState.lastSpellInfo;
    var cellsInfo;

    if (lastSpell) {
      cellsInfo = lastSpell["cellsInfo"];
      var usrWords = lastSpell["usrWords"];
      var usrCorrect = lastSpell["usrCorrect"];
      var wordsIndex = lastSpell["wordsIndex"];
      var ignoreWords = this.spellcheckState.ignoreWords;
      var changeWords = this.spellcheckState.changeWords;
      var lastIndex = this.spellcheckState.lastIndex;
      this.spellcheckState.cellText = this.asc_getCellInfo().text;
      var cellText = this.spellcheckState.newCellText || this.spellcheckState.cellText;
      var afterReplace = this.spellcheckState.afterReplace;

      for (var i = 0; i < usrWords.length; i++) {
        var usrWord = usrWords[i];

        if (ignoreWords[usrWord] || changeWords[usrWord]) {
          usrCorrect[i] = true;
        }
      }

      while (cellsInfo[lastIndex].col === activeCell.col && cellsInfo[lastIndex].row === activeCell.row) {
        var letterDifference = null;
        var word = usrWords[lastIndex];
        var newWord = this.spellcheckState.newWord;

        if (newWord) {
          letterDifference = newWord.length - word.length;
        } else if (changeWords[word]) {
          letterDifference = changeWords[word].length - word.length;
        }

        if (letterDifference !== null) {
          var replaceWith = newWord || changeWords[word];
          var valueForSearching = new RegExp(usrWords[lastIndex], "y");
          valueForSearching.lastIndex = wordsIndex[lastIndex];
          cellText = cellText.replace(valueForSearching, replaceWith);
          if (letterDifference !== 0) {
            var j = lastIndex + 1;
            while (j < usrWords.length && cellsInfo[j].col === cellsInfo[lastIndex].col) {
              wordsIndex[j] += letterDifference;
              j++;
            }
          }
          this.spellcheckState.newCellText = cellText;
          this.spellcheckState.newWord = null;
        }

        if (this.spellcheckState.newCellText) {
          this.spellcheckState.cellsChange = [];
          var cell = cellsInfo[lastIndex];
          this.spellcheckState.cellsChange.push(new Asc.Range(cell.col, cell.row, cell.col, cell.row));
        }

        if (afterReplace && !usrCorrect[lastIndex]) {
          afterReplace = false;
          break;
        }

        ++lastIndex;

        if (!afterReplace && !usrCorrect[lastIndex]) {
          afterReplace = false;
          break;
        }
      }
      this.spellcheckState.lastIndex = lastIndex;
      this.SpellCheck_CallBack(this.spellcheckState.lastSpellInfo);
      return;
    }

    var maxC = ws.model.getColsCount() - 1;
    var maxR = ws.model.getRowsCount() - 1;
    if (-1 === maxC || -1 === maxR) {
      this.handlers.trigger("asc_onSpellCheckVariantsFound", new AscCommon.asc_CSpellCheckProperty());
      return;
    }

    if (!this.spellcheckState.startCell) {
      this.spellcheckState.startCell = activeCell.clone();
      if (this.spellcheckState.startCell.col > maxC) {
        this.spellcheckState.startCell.row += 1;
        this.spellcheckState.startCell.col = 0;
      }
      if (this.spellcheckState.startCell.row > maxR) {
        this.spellcheckState.startCell.row = 0;
        this.spellcheckState.startCell.col = 0;
      }
      this.spellcheckState.currentCell = this.spellcheckState.startCell.clone();
    }

    var startCell = this.spellcheckState.startCell;
    var currentCell = this.spellcheckState.currentCell;
    var lang = this.defaultLanguage;

    var langArray = [];
    var wordsArray = [];
    cellsInfo = [];
    var wordsIndexArray = [];
    var isEnd = false;

    do {
      if (this.spellcheckState.iteration && currentCell.row > startCell.row) {
        break;
      }
      if (currentCell.row > maxR) {
        currentCell.row = 0;
        currentCell.col = 0;
        this.spellcheckState.iteration = true;
      }
      if (this.spellcheckState.iteration && currentCell.row === startCell.row) {
        maxC = startCell.col - 1;
      }
      if (currentCell.col > maxC) {
        this.spellcheckState.nextRow();
        continue;
      }
      ws.model.getRange3(currentCell.row, currentCell.col, currentCell.row, maxC)._foreachNoEmpty(function (cell, r, c) {
        if (cell.text !== null && !cell.isFormula()) {
          var cellInfo = new AscCommon.CellBase(r, c);
          var wordsObject = AscCommonExcel.WordSplitting(cell.text);
          var words = wordsObject.wordsArray;
          var wordsIndex = wordsObject.wordsIndex;
          for (var i = 0; i < words.length; ++i) {
            wordsArray.push(words[i]);
            wordsIndexArray.push(wordsIndex[i]);
            langArray.push(lang);
            cellsInfo.push(cellInfo);
          }
          isEnd = true;
        }
      });
      if (isEnd) {
        break;
      }
      this.spellcheckState.nextRow();
    } while (true);

    if (0 < wordsArray.length) {
      this.spellcheckState.lockSpell = true;
      this.SpellCheckApi.spellCheck({
        "type": "spell",
        "usrWords": wordsArray,
        "usrLang": langArray,
        "cellsInfo": cellsInfo,
        "wordsIndex": wordsIndexArray
      });
    } else {
      this.handlers.trigger("asc_onSpellCheckVariantsFound", new AscCommon.asc_CSpellCheckProperty());
      var lastFindOptions = this.spellcheckState.lastFindOptions;
      var cellsChange = this.spellcheckState.cellsChange;
      if (cellsChange.length !== 0 && lastFindOptions) {
        this.spellcheckState.lockSpell = true;
        this.asc_replaceMisspelledWords(lastFindOptions);
      }
      this.spellcheckState.isStart = false;
    }
  };

  spreadsheet_api.prototype.asc_replaceMisspelledWord = function(newWord, variantsFound, replaceAll) {
    var options = new Asc.asc_CFindOptions();
    options.isSpellCheck = true;
    options.wordsIndex = this.spellcheckState.wordsIndex;
    this.spellcheckState.newWord = newWord;
    options.isMatchCase = true;

    if (replaceAll === true) {
      if (!this.spellcheckState.isIgnoreUppercase) {
        this.spellcheckState.changeWords[variantsFound.Word] = newWord;
      } else {
        this.spellcheckState.changeWords[variantsFound.Word.toLowerCase()] = newWord;
      }
      options.isReplaceAll = true;
    }
      this.spellcheckState.lockSpell = false;
      this.spellcheckState.lastFindOptions = options;
      this.asc_nextWord();
  };

    spreadsheet_api.prototype.asc_replaceMisspelledWords = function (options) {
      var t = this;
      var ws = this.wb.getWorksheet();
      var cellsChange = this.spellcheckState.cellsChange;
      var changeWords = this.spellcheckState.changeWords;
      var cellText = this.spellcheckState.cellText;
      var newCellText = this.spellcheckState.newCellText;

      if (!newCellText) {
        cellText = null;
        newCellText = null;
      }

      var replaceWords = [];
      options.isWholeWord = true;
      for (var key in changeWords) {
        replaceWords.push([AscCommonExcel.getFindRegExp(key, options), changeWords[key]]);
      }
      options.findWhat = cellText;
      options.isMatchCase = true;
      options.replaceWith = newCellText;
      options.replaceWords = replaceWords;

      History.Create_NewPoint();
      History.StartTransaction();

      ws._replaceCellsText(cellsChange, options, false, function () {
        t.spellcheckState.cellsChange = [];
        options.indexInArray = 0;
        var lastSpell = t.spellcheckState.lastSpellInfo;
        if (lastSpell) {
          var lastIndex = t.spellcheckState.lastIndex;
          var cellInfo = lastSpell["cellsInfo"][lastIndex];
          t.spellcheckState.lockSpell = true;
          ws.setSelection(new Asc.Range(cellInfo.col, cellInfo.row, cellInfo.col, cellInfo.row));
          t.spellcheckState.lockSpell = false;
          t.spellcheckState.newWord = null;
          t.spellcheckState.newCellText = null;
          t.spellcheckState.afterReplace = true;
          t.spellcheckState.lastIndex = 0;
          t.spellcheckState.cellText = t.asc_getCellInfo().text;
          t.asc_nextWord();
          return;
        }
        t.spellcheckState.lockSpell = false;
        History.EndTransaction();
      });
    };

  spreadsheet_api.prototype.asc_ignoreMisspelledWord = function(spellCheckProperty, ignoreAll) {
    if (ignoreAll) {
      var word = spellCheckProperty.Word;
        this.spellcheckState.ignoreWords[word] = word;
    }
    this.asc_nextWord();
  };

    spreadsheet_api.prototype.asc_ignoreNumbers = function (isIgnore) {
      this.spellcheckState.isIgnoreNumbers = isIgnore;
    };

    spreadsheet_api.prototype.asc_ignoreUppercase = function (isIgnore) {
      this.spellcheckState.isIgnoreUppercase = isIgnore;
    };

  spreadsheet_api.prototype.asc_cancelSpellCheck = function() {
    this.cleanSpelling();
  };

  // Frozen pane
  spreadsheet_api.prototype.asc_freezePane = function (type) {
    if (this.canEdit()) {
      this.wb.getWorksheet().freezePane(type);
    }
  };

	spreadsheet_api.prototype.asc_setSparklineGroup = function (id, oSparklineGroup) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		var t = this;
		var changeSparkline = function (res) {
			if (res) {
				var changedSparkline = g_oTableId.Get_ById(id);
				if (changedSparkline) {
					History.Create_NewPoint();
					History.StartTransaction();
					changedSparkline.set(oSparklineGroup);
					History.EndTransaction();
					t.wb._onWSSelectionChanged();
					t.wb.getWorksheet().draw();
				}
			}
		};
		this._isLockedSparkline(id, changeSparkline);
	};

	spreadsheet_api.prototype.asc_addSparklineGroup = function (type, sDataRange, sLocationRange) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}

		var wsView = this.wb.getWorksheet();
		return wsView.addSparklineGroup(type, sDataRange, sLocationRange);
	};

    spreadsheet_api.prototype.asc_setListType = function (type, subtype, custom) {
      var t = this;
        var sNeedFont = AscFormat.fGetFontByNumInfo(type, subtype, custom);
      if(typeof sNeedFont === "string" && sNeedFont.length > 0){
          var t = this, fonts = {};
          fonts[sNeedFont] = 1;
          t._loadFonts(fonts, function() {t.asc_setListType2(type, subtype, custom);});
      }
      else{
          t.asc_setListType2(type, subtype);
      }
    };
    spreadsheet_api.prototype.asc_setListType2 = function (type, subtype, custom) {
        var oWorksheet = this.wb.getWorksheet();
        if(oWorksheet){
            if(oWorksheet.isSelectOnShape){
                return oWorksheet.objectRender.setListType(type, subtype, custom);
            }
        }
    };

  // Cell interface
  spreadsheet_api.prototype.asc_getCellInfo = function() {
    return this.wb && this.wb.getSelectionInfo();
  };

  // Получить координаты активной ячейки
  spreadsheet_api.prototype.asc_getActiveCellCoord = function(useUpRightMerge) {
    var oWorksheet = this.wb.getWorksheet();
    if(oWorksheet){
      if(oWorksheet.isSelectOnShape){
        return oWorksheet.objectRender.getContextMenuPosition();
      }
      else{
          return oWorksheet.getActiveCellCoord(useUpRightMerge);
      }
    }
  };

	// Получить координаты активной ячейки
	spreadsheet_api.prototype.asc_getActiveCell = function() {
		var oWorksheet = this.wb.getWorksheet();
		if(oWorksheet){
			if(oWorksheet.isSelectOnShape){
				return null;
			}
			else{
				return oWorksheet.getActiveCell();
			}
		}

	};

  // Получить координаты для каких-либо действий (для общей схемы)
  spreadsheet_api.prototype.asc_getAnchorPosition = function() {
    return this.asc_getActiveCellCoord();
  };

  // Получаем свойство: редактируем мы сейчас или нет
  spreadsheet_api.prototype.asc_getCellEditMode = function() {
    return this.wb ? this.wb.getCellEditMode() : false;
  };

  spreadsheet_api.prototype.asc_getHeaderFooterMode = function() {
	  return this.wb && this.wb.getCellEditMode() && this.wb.cellEditor && this.wb.cellEditor.options &&
		  this.wb.cellEditor.options.menuEditor;
  };

  spreadsheet_api.prototype.asc_getActiveRangeStr = function(referenceType, opt_getActiveCell, opt_ignore_r1c1) {
  	var ws = this.wb.getWorksheet();
  	var res = null;
  	var tmpR1C1;
  	if (opt_ignore_r1c1) {
		tmpR1C1 = AscCommonExcel.g_R1C1Mode;
		AscCommonExcel.g_R1C1Mode = false;
  	}
  	if (ws && ws.model && ws.model.selectionRange) {
  		var range;
  		if (opt_getActiveCell) {
			var activeCell = ws.model.selectionRange.activeCell;
			range = new Asc.Range(activeCell.col, activeCell.row, activeCell.col, activeCell.row);
		} else {
			var lastRange = ws.model.selectionRange.getLast();
			range = new Asc.Range(lastRange.c1, lastRange.r1, lastRange.c2, lastRange.r2);
		}

		res = range.getName(referenceType);
	}
  	if (opt_ignore_r1c1) {
		AscCommonExcel.g_R1C1Mode = tmpR1C1;
  	}
	return res;
  };

  spreadsheet_api.prototype.asc_getIsTrackShape = function()  {
    return this.wb ? this.wb.getIsTrackShape() : false;
  };

  spreadsheet_api.prototype.asc_setCellFontName = function(fontName) {
    var t = this, fonts = {};
    fonts[fontName] = 1;
    t._loadFonts(fonts, function() {
      var ws = t.wb.getWorksheet();
      if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellFontName) {
        ws.objectRender.controller.setCellFontName(fontName);
      } else {
        t.wb.setFontAttributes("fn", fontName);
        t.wb.restoreFocus();
      }
    });
  };

  spreadsheet_api.prototype.asc_setCellFontSize = function(fontSize) {
    var ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellFontSize) {
      ws.objectRender.controller.setCellFontSize(fontSize);
    } else {
      this.wb.setFontAttributes("fs", fontSize);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellBold = function(isBold) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
     return;
    }

  	let ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellBold) {
      ws.objectRender.controller.setCellBold(isBold);
    } else {
      this.wb.setFontAttributes("b", isBold);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellItalic = function(isItalic) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    let ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellItalic) {
      ws.objectRender.controller.setCellItalic(isItalic);
    } else {
      this.wb.setFontAttributes("i", isItalic);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellUnderline = function(isUnderline) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    let ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellUnderline) {
      ws.objectRender.controller.setCellUnderline(isUnderline);
    } else {
      this.wb.setFontAttributes("u", isUnderline ? Asc.EUnderline.underlineSingle : Asc.EUnderline.underlineNone);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellStrikeout = function(isStrikeout) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    let ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellStrikeout) {
      ws.objectRender.controller.setCellStrikeout(isStrikeout);
    } else {
      this.wb.setFontAttributes("s", isStrikeout);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellSubscript = function(isSubscript) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    let ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellSubscript) {
      ws.objectRender.controller.setCellSubscript(isSubscript);
    } else {
      this.wb.setFontAttributes("fa", isSubscript ? AscCommon.vertalign_SubScript : null);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellSuperscript = function(isSuperscript) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    let ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellSuperscript) {
      ws.objectRender.controller.setCellSuperscript(isSuperscript);
    } else {
      this.wb.setFontAttributes("fa", isSuperscript ? AscCommon.vertalign_SuperScript : null);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellAlign = function(align) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    let ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellAlign) {
      ws.objectRender.controller.setCellAlign(align);
    } else {
      this.wb.getWorksheet().setSelectionInfo("a", align);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellVertAlign = function(align) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    let ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellVertAlign) {
      ws.objectRender.controller.setCellVertAlign(align);
    } else {
      this.wb.getWorksheet().setSelectionInfo("va", align);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellTextWrap = function(isWrapped) {
    var ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellTextWrap) {
      ws.objectRender.controller.setCellTextWrap(isWrapped);
    } else {
      this.wb.getWorksheet().setSelectionInfo("wrap", isWrapped);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellTextShrink = function(isShrinked) {
    var ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellTextShrink) {
      ws.objectRender.controller.setCellTextShrink(isShrinked);
    } else {
      this.wb.getWorksheet().setSelectionInfo("shrink", isShrinked);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellTextColor = function(color) {
    var ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellTextColor) {
      ws.objectRender.controller.setCellTextColor(color);
    } else {
      if (color instanceof Asc.asc_CColor) {
        color = AscCommonExcel.CorrectAscColor(color);
        this.wb.setFontAttributes("c", color);
        this.wb.restoreFocus();
      }
    }

  };

  spreadsheet_api.prototype.asc_setCellFill = function (fill) {
    var ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellBackgroundColor) {
      ws.objectRender.controller.setCellBackgroundColor(fill);
    } else {
      this.wb.getWorksheet().setSelectionInfo("f", fill);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellBackgroundColor = function(color) {
    var ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellBackgroundColor) {
      ws.objectRender.controller.setCellBackgroundColor(color);
    } else {
      if (color instanceof Asc.asc_CColor || null == color) {
        if (null != color) {
          color = AscCommonExcel.CorrectAscColor(color);
        }
        this.wb.getWorksheet().setSelectionInfo("bc", color);
        this.wb.restoreFocus();
      }
    }
  };

  spreadsheet_api.prototype.asc_setCellBorders = function(borders) {
    this.wb.getWorksheet().setSelectionInfo("border", borders);
    this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_setCellFormat = function(format) {
    var t = this;
    //todo split setCellFormat into set->_loadFonts->draw and remove checkCultureInfoFontPicker(checkCultureInfoFontPicker is called inside StyleManager.setNum)
    var numFormat = AscCommon.oNumFormatCache.get(format);
    numFormat.checkCultureInfoFontPicker();
    this._loadFonts([], function () {
      t.wb.setCellFormat(format);
      t.wb.restoreFocus();
    });
  };

  spreadsheet_api.prototype.asc_setCellAngle = function(angle) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    if (this.wb) {
      var ws = this.wb.getWorksheet();
      if (ws) {
        if (ws.objectRender && ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.setCellAngle) {
            ws.objectRender.controller.setCellAngle(angle);
        } else {
            ws.setSelectionInfo("angle", angle);
            this.wb.restoreFocus();
      	}
      }
    }
  };

  spreadsheet_api.prototype.asc_setCellStyle = function(name) {
    this.wb.getWorksheet().setSelectionInfo("style", name);
    this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_ChangeTextCase = function(nType) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    if (this.wb) {
      var ws = this.wb && this.wb.getWorksheet();
      if (ws && ws.objectRender && ws.objectRender.selectedGraphicObjectsExists()) {
      	ws.objectRender.controller.changeTextCase(nType);
      } else {
      	this.wb.changeTextCase(nType);
      	this.wb.restoreFocus();
      }
	}
  };

  spreadsheet_api.prototype.asc_increaseCellDigitNumbers = function() {
    this.wb.getWorksheet().setSelectionInfo("changeDigNum", +1);
    this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_decreaseCellDigitNumbers = function() {
    this.wb.getWorksheet().setSelectionInfo("changeDigNum", -1);
    this.wb.restoreFocus();
  };

  // Увеличение размера шрифта
  spreadsheet_api.prototype.asc_increaseFontSize = function() {
    var ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.increaseFontSize) {
      ws.objectRender.controller.increaseFontSize();
    } else {
      this.wb.changeFontSize("changeFontSize", true);
      this.wb.restoreFocus();
    }
  };

  // Уменьшение размера шрифта
  spreadsheet_api.prototype.asc_decreaseFontSize = function() {
    var ws = this.wb.getWorksheet();
    if (ws.objectRender.selectedGraphicObjectsExists() && ws.objectRender.controller.decreaseFontSize) {
      ws.objectRender.controller.decreaseFontSize();
    } else {
      this.wb.changeFontSize("changeFontSize", false);
      this.wb.restoreFocus();
    }
  };

  spreadsheet_api.prototype.asc_setCellIndent = function(val) {
	  this.wb.getWorksheet().setSelectionInfo("indent", val);
	  this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_setCellProtection = function (val) {
    this.wb.getWorksheet().setSelectionInfo("applyProtection", val);
    this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_setCellLocked = function (val) {
    this.wb.getWorksheet().setSelectionInfo("locked", val);
    this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_setCellHiddenFormulas = function (val) {
    this.wb.getWorksheet().setSelectionInfo("hiddenFormulas", val);
    this.wb.restoreFocus();
  };

	spreadsheet_api.prototype.asc_checkProtectedRange = function () {
		var ws = this.wbModel.getActiveWs();
		/*if (!ws.isLockedActiveCell()) {
			return false;
		}*/

		var protectedRanges = ws.getProtectedRangesByActiveRange();

		//входит ли в зону защищенных диапазонов ячейка - null(не входит)/true(входит и защищена паролем)/false(входит и не защищена паролем или была защищена и уже не защищена)
		var res = null;
		if (protectedRanges) {
			for (var i = 0; i < protectedRanges.length; i++) {
				if (protectedRanges[i].asc_isPassword()) {
					if (protectedRanges[i].isUserEnteredPassword()) {
						return false;
					} else if (res !== false) {
						res = true;
					}
				} else {
					return false;
				}
			}
		}

		return res;
	};

	spreadsheet_api.prototype.asc_checkLockedCells = function () {
		var ws = this.wbModel.getActiveWs();
		var res = null;
		if (ws) {
			var _selection = ws.getSelection();
			if (_selection && _selection.ranges) {
				res = ws.isIntersectLockedRanges(_selection.ranges);
			}
		}
		return res;
	};

	spreadsheet_api.prototype.asc_checkActiveCellPassword = function (val, callback) {
		var ws = this.wbModel.getActiveWs();
		var activeCell = ws && ws.selectionRange && ws.selectionRange.activeCell;
		if (activeCell) {
			ws.checkProtectedRangesPassword(val, activeCell, callback);
		}
	};

  // Формат по образцу
  spreadsheet_api.prototype.asc_formatPainter = function(formatPainterState) {
	this.changeFormatPainterState(formatPainterState, undefined);
  };


	spreadsheet_api.prototype.changeFormatPainterState = function(formatPainterState, bLockDraw) {
        if (this.isStartAddShape) {
            this.asc_endAddShape();
        }
        this.stopInkDrawer();
        this.cancelEyedropper();
		this.formatPainter.putState(formatPainterState);
		if (this.wb) {
			this.wb.formatPainter(formatPainterState, bLockDraw);
		}
	};
	spreadsheet_api.prototype.retrieveFormatPainterData = function()
	{
		let oWSView = this.wb.getWorksheet();
		if(!oWSView || !oWSView.model) {
			return null;
		}
		return new AscCommonExcel.CCellFormatPasteData(oWSView);
	};

  spreadsheet_api.prototype.asc_showAutoComplete = function() {
    this.wb.showAutoComplete();
  };

  spreadsheet_api.prototype.asc_onMouseUp = function(event, x, y) {
    if (this.wb) {
      this.wb._onWindowMouseUpExternal(event, x, y);
    }
  };

  //

  spreadsheet_api.prototype.asc_selectFunction = function() {

  };

	spreadsheet_api.prototype.asc_insertHyperlink = function (options) {
		AscFonts.FontPickerByCharacter.checkText(options.text, this, function () {
			this.wb.insertHyperlink(options);
		});
	};

  spreadsheet_api.prototype.asc_removeHyperlink = function() {
    this.wb.removeHyperlink();
  };

  spreadsheet_api.prototype.asc_getFullHyperlinkLength = function(str) {
    return window["AscCommonExcel"].getFullHyperlinkLength(str);
  };

    spreadsheet_api.prototype.asc_cleanSelectRange = function () {
        this.wb._onCleanSelectRange();
    };

  spreadsheet_api.prototype.asc_insertInCell = function(functionName, type, autoComplete) {
    this.wb.insertInCellEditor(functionName, type, autoComplete);
    this.wb.restoreFocus();
  };

  spreadsheet_api.prototype.asc_startWizard = function (name, doCleanCellContent) {
    this.wb.startWizard(name, doCleanCellContent);
    this.wb.restoreFocus();
  };
  spreadsheet_api.prototype.asc_canEnterWizardRange = function(char) {
    return this.wb.canEnterWizardRange(char);
  };
  spreadsheet_api.prototype.asc_insertArgumentsInFormula = function(val, argNum, argType, name) {
    var res = this.wb.insertArgumentsInFormula(val, argNum, argType, name);
    this.wb.restoreFocus();
    return res;
  };

  spreadsheet_api.prototype.asc_getFormulasInfo = function() {
    return this.formulasList;
  };
  spreadsheet_api.prototype.asc_getFormulaLocaleName = function(name) {
    return AscCommonExcel.cFormulaFunctionToLocale ? AscCommonExcel.cFormulaFunctionToLocale[name] : name;
  };
  spreadsheet_api.prototype.asc_getFormulaNameByLocale = function (name) {
    var f = AscCommonExcel.cFormulaFunctionLocalized && AscCommonExcel.cFormulaFunctionLocalized[name];
    return f ? f.prototype.name : name;
  };
  spreadsheet_api.prototype.asc_calculate = function(type) {
    if (this.canEdit()) {
        this.wb.calculate(type);
    }
  };

  spreadsheet_api.prototype.asc_setFontRenderingMode = function(mode) {
    if (mode !== this.fontRenderingMode) {
      this.fontRenderingMode = mode;
      if (this.wb) {
        this.wb.setFontRenderingMode(mode, /*isInit*/false);
      }
    }
  };

  /**
   * Режим выбора диапазона
   * @param {Asc.c_oAscSelectionDialogType} selectionDialogType
   * @param selectRange
   */
  spreadsheet_api.prototype.asc_setSelectionDialogMode = function(selectionDialogType, selectRange) {
    if (this.wb) {
      this.wb._onStopFormatPainter();
      this.wb.setSelectionDialogMode(selectionDialogType, selectRange);
    }
  };

  spreadsheet_api.prototype.asc_SendThemeColors = function(colors, standart_colors) {
    this._gui_control_colors = { Colors: colors, StandartColors: standart_colors };
    var ret = this.handlers.trigger("asc_onSendThemeColors", colors, standart_colors);
    if (false !== ret) {
      this._gui_control_colors = null;
    }
  };

  spreadsheet_api.prototype.getCurrentTheme = function()
  {
     return this.wbModel && this.wbModel.theme;
  };

  spreadsheet_api.prototype.getGraphicController = function () {
      var ws = this.wb.getWorksheet();
      return ws && ws.objectRender && ws.objectRender.controller;
  };

    spreadsheet_api.prototype.asc_ChangeColorScheme = function (sSchemeName) {
		if (this.wbModel) {
			for (var i = 0; i < this.wbModel.aWorksheets.length; ++i) {
				var sheet = this.wbModel.aWorksheets[i];
				if (sheet.getSheetProtection()) {
					//TODO error
					return;
				}
			}
		}

		var t = this;
		var onChangeColorScheme = function (res) {
			if (res) {
				if (t.wbModel.changeColorScheme(sSchemeName)) {
					t.asc_AfterChangeColorScheme();
				}
			}
		};
		// ToDo поправить заглушку, сделать новый тип lock element-а
		var sheetId = -1; // Делаем не существующий лист и не существующий объект
		var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Object, /*subType*/null, sheetId,
			sheetId);
		this.collaborativeEditing.lock([lockInfo], onChangeColorScheme);
	};

  spreadsheet_api.prototype.asc_ChangeColorSchemeByIdx = function (nIdx) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return;
    }

    if (this.wbModel) {
      for (var i = 0; i < this.wbModel.aWorksheets.length; ++i) {
      	var sheet = this.wbModel.aWorksheets[i];
      	if (sheet.getSheetProtection()) {
      		//TODO error
     		return;
     	 }
      }
    }

    var t = this;
    var onChangeColorScheme = function (res) {
      if (res) {
        if (t.wbModel.changeColorSchemeByIdx(nIdx)) {
          t.asc_AfterChangeColorScheme();
        }
      }
    };
    // ToDo поправить заглушку, сделать новый тип lock element-а
    var sheetId = -1; // Делаем не существующий лист и не существующий объект
    var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Object, /*subType*/null, sheetId,
        sheetId);
    this.collaborativeEditing.lock([lockInfo], onChangeColorScheme);
  };


  spreadsheet_api.prototype.asc_AfterChangeColorScheme = function() {
    this.asc_CheckGuiControlColors();
    this.asc_ApplyColorScheme(true);
  };
  spreadsheet_api.prototype.asc_ApplyColorScheme = function(bRedraw) {

    if (window['IS_NATIVE_EDITOR'] || !window["NATIVE_EDITOR_ENJINE"]) {
      Asc["editor"].wb.recalculateDrawingObjects();
      this.chartPreviewManager.clearPreviews();
      this.textArtPreviewManager.clear();
    }

    // На view-режиме не нужно отправлять стили
    if (!this.getViewMode()) {
      // Отправка стилей
      this._sendWorkbookStyles();
    }

    if (bRedraw) {
      this.handlers.trigger("asc_onUpdateChartStyles");
      this.handlers.trigger("asc_onSelectionChanged", this.asc_getCellInfo());
      this.wb.drawWS();
    }
  };

  /////////////////////////////////////////////////////////////////////////
  ////////////////////////////AutoSave api/////////////////////////////////
  /////////////////////////////////////////////////////////////////////////
	spreadsheet_api.prototype._autoSaveInner = function () {
		if (this.asc_getCellEditMode() || this.asc_getIsTrackShape()) {
		  return;
        }

		if (this.isLiveViewer()) {
			if (this.collaborativeEditing.haveOtherChanges()) {
				this.collaborativeEditing.applyChanges();
			}
			return;
		} else if (!History.Have_Changes(true) && !(this.collaborativeEditing.getCollaborativeEditing() &&
			0 !== this.collaborativeEditing.getOwnLocksLength())) {
			if (this.collaborativeEditing.getFast() && this.collaborativeEditing.haveOtherChanges()) {
				AscCommon.CollaborativeEditing.Clear_CollaborativeMarks();

				// Принимаем чужие изменения
				this.collaborativeEditing.applyChanges();
				// Пересылаем свои изменения (просто стираем чужие lock-и, т.к. своих изменений нет)
				this.collaborativeEditing.sendChanges();
				// Шлем update для toolbar-а, т.к. когда select в lock ячейке нужно заблокировать toolbar
				this.wb._onWSSelectionChanged();
			}
            if (AscCommon.CollaborativeEditing.Is_Fast() /*&& !AscCommon.CollaborativeEditing.Is_SingleUser()*/) {
                this.wb.sendCursor();
            }
			return;
		}
		if (null === this.lastSaveTime) {
			this.lastSaveTime = new Date();
			return;
		}

        var _bIsWaitScheme = false;
        var _curTime =  new Date();
        if((this.collaborativeEditing.Is_SingleUser() || !this.collaborativeEditing.getFast()) && History.Points[History.Index]) {
            if ((_curTime - History.Points[History.Index].Time) < this.intervalWaitAutoSave) {
                _bIsWaitScheme = true;
            }
        }
        if(!_bIsWaitScheme) {
            var saveGap = this.collaborativeEditing.getFast() ? this.autoSaveGapRealTime :
                (this.collaborativeEditing.getCollaborativeEditing() ? this.autoSaveGapSlow : this.autoSaveGapFast);
            var gap = _curTime - this.lastSaveTime - saveGap;
            if (0 <= gap) {
                this.asc_Save(true);
            }
            if (AscCommon.CollaborativeEditing.Is_Fast() /*&& !AscCommon.CollaborativeEditing.Is_SingleUser()*/) {
                this.wb.sendCursor();
            }
        }
	};

	spreadsheet_api.prototype._onUpdateDocumentCanSave = function () {
		// Можно модифицировать это условие на более быстрое (менять самим состояние в аргументах, а не запрашивать каждый раз)
		var tmp = History.Have_Changes() || (this.collaborativeEditing.getCollaborativeEditing() &&
			0 !== this.collaborativeEditing.getOwnLocksLength()) || this.asc_getCellEditMode();
		if (tmp !== this.isDocumentCanSave) {
			this.isDocumentCanSave = tmp;
			this.handlers.trigger('asc_onDocumentCanSaveChanged', this.isDocumentCanSave);
		}
	};
	spreadsheet_api.prototype._onUpdateDocumentCanUndoRedo = function () {
		AscCommon.History && AscCommon.History._sendCanUndoRedo();
	};

  spreadsheet_api.prototype._onCheckCommentRemoveLock = function(lockElem) {
    var res = false;
    var sheetId = lockElem["sheetId"];
    if (-1 !== sheetId && 0 === sheetId.indexOf(AscCommonExcel.CCellCommentator.sStartCommentId)) {
      // Коммментарий
      res = true;
      this.handlers.trigger("asc_onUnLockComment", lockElem["rangeOrObjectId"]);
    }
    return res;
  };

  spreadsheet_api.prototype.onUpdateDocumentModified = function(bIsModified) {
    // Обновляем только после окончания сохранения
    if (this.canSave) {
      this.handlers.trigger("asc_onDocumentModifiedChanged", bIsModified);
      this._onUpdateDocumentCanSave();

      if (undefined !== window["AscDesktopEditor"]) {
        window["AscDesktopEditor"]["onDocumentModifiedChanged"](bIsModified);
      }
    }
  };

	// Выставление локали
	spreadsheet_api.prototype.asc_setLocalization = function (oLocalizedData) {
		if (!this.isLoadFullApi) {
			this.tmpLocalization = oLocalizedData;
			return;
		}

		if (null == oLocalizedData) {
			AscCommonExcel.cFormulaFunctionLocalized = null;
			AscCommonExcel.cFormulaFunctionToLocale = null;
		} else {
			AscCommonExcel.cFormulaFunctionLocalized = {};
			AscCommonExcel.cFormulaFunctionToLocale = {};
			var localName;
			for (var i in AscCommonExcel.cFormulaFunction) {
				localName = oLocalizedData[i] ? oLocalizedData[i] : null;
				localName = localName ? localName : i;
				AscCommonExcel.cFormulaFunctionLocalized[localName] = AscCommonExcel.cFormulaFunction[i];
				AscCommonExcel.cFormulaFunctionToLocale[i] = localName;
			}
		}
		AscCommon.build_local_rx(oLocalizedData ? oLocalizedData["LocalFormulaOperands"] : null);
		if (this.wb) {
			this.wb.initFormulasList();
			this.wb._onWSSelectionChanged();
		}
		if (this.wbModel) {
			this.wbModel.rebuildColors();
		}

		//on change formula language need load fonts
		const constError = oLocalizedData &&  oLocalizedData["LocalFormulaOperands"] && oLocalizedData["LocalFormulaOperands"]["CONST_ERROR"];
		if (constError && constError["nil"]) {
			if (AscFonts.FontPickerByCharacter.getFontsByString(constError["nil"])) {
				this._loadFonts([], function() {});
			}
		}
	};

  spreadsheet_api.prototype.asc_nativeOpenFile = function(base64File, version, isUser, xlsxPath) {
	var t = this;
    asc["editor"] = this;

    this.SpellCheckUrl = '';

    if (undefined == isUser) {
        this.User = new AscCommon.asc_CUser();
        this.User.setId("TM");
        this.User.setUserName("native");
    }
    if (undefined !== version) {
      AscCommon.CurFileVersion = version;
    }

	  let xlsxData;
	  this.isOpenOOXInBrowser = this["asc_isSupportFeature"]("ooxml") && AscCommon.checkOOXMLSignature(base64File);
	  if (this.isOpenOOXInBrowser) {
		  //slice because array contains garbage after end of function
		  this.openOOXInBrowserZip = base64File.slice();
	  } else if (xlsxPath && window["native"]["GetFileBinary"]) {
		  xlsxData = xlsxPath && window["native"]["GetFileBinary"] && window["native"]["GetFileBinary"](xlsxPath);
	  }
	  this._openDocument(base64File);
	  if (this.openDocumentFromZip(t.wbModel, xlsxData)) {
		  Asc.ReadDefTableStyles(t.wbModel);
		  g_oIdCounter.Set_Load(false);
		  AscCommon.checkCultureInfoFontPicker();
		  AscCommonExcel.checkStylesNames(t.wbModel.CellStyles);
		  t._coAuthoringInit();
		  t.wb = new AscCommonExcel.WorkbookView(t.wbModel, t.controller, t.handlers, window["_null_object"], window["_null_object"], t, t.collaborativeEditing, t.fontRenderingMode);
	  }
  };

  spreadsheet_api.prototype.asc_nativeCalculateFile = function() {
    window['DoctRendererMode'] = true;
    this.wb._nativeCalculate();
  };

  spreadsheet_api.prototype.asc_nativeApplyChanges = function(changes) {
    for (var i = 0, l = changes.length; i < l; ++i) {
      this.CoAuthoringApi.onSaveChanges(changes[i], null, true);
    }
    this.collaborativeEditing.applyChanges();
  };

  spreadsheet_api.prototype.asc_nativeApplyChanges2 = function(data, isFull) {
    if (null != this.wbModel) {
      this.oRedoObjectParamNative = this.wbModel.DeserializeHistoryNative(this.oRedoObjectParamNative, data, isFull);
    }
    if (isFull) {
      this._onUpdateAfterApplyChanges();

      if (window["NATIVE_EDITOR_ENJINE"] === true && window["native"]["AddImageInChanges"] && window["NATIVE_EDITOR_ENJINE_NEW_IMAGES"])
      {
        var _new_images     = window["NATIVE_EDITOR_ENJINE_NEW_IMAGES"];
        var _new_images_len = _new_images.length;

        for (var nImage = 0; nImage < _new_images_len; nImage++)
          window["native"]["AddImageInChanges"](_new_images[nImage]);
      }
    }
  };

	spreadsheet_api.prototype._coAuthoringSetChanges = function(e, oColor)
	{
		var Count = e.length;
		for (var Index = 0; Index < Count; ++Index) {
			this.CoAuthoringApi.onSaveChanges(e[Index], null, true);
		}
		this.collaborativeEditing.applyChanges();
		//this._onUpdateAfterApplyChanges();
	};

  spreadsheet_api.prototype.asc_nativeGetFile = function() {
	  var oBinaryFileWriter = new AscCommonExcel.BinaryFileWriter(this.wbModel);
    return oBinaryFileWriter.Write();
  };
  spreadsheet_api.prototype.asc_nativeGetFile3 = function()
  {
      var oBinaryFileWriter = new AscCommonExcel.BinaryFileWriter(this.wbModel);
      oBinaryFileWriter.Write(true, true);
      return { data: oBinaryFileWriter.Write(true, true), header: oBinaryFileWriter.WriteFileHeader(oBinaryFileWriter.Memory.GetCurPosition(), Asc.c_nVersionNoBase64) };
  };
  spreadsheet_api.prototype.asc_nativeGetFileData = function() {
	  //calc to fix case where file has formulas with no cache values and no changes
	  this.wbModel.dependencyFormulas.calcTree();
	  if (this.isOpenOOXInBrowser && this.saveDocumentToZip) {
		  let res;
		  this.saveDocumentToZip(this.wb.model, this.editorId, function(data) {
			  res = data;
		  });
		  if (res) {
			  window["native"]["Save_End"](";v10;", res.length);
			  return res;
		  }
		  return new Uint8Array(0);
	  } else {
		  var oBinaryFileWriter = new AscCommonExcel.BinaryFileWriter(this.wbModel);
		  oBinaryFileWriter.Write(true);

		  var _header = oBinaryFileWriter.WriteFileHeader(oBinaryFileWriter.Memory.GetCurPosition(), Asc.c_nVersionNoBase64);
		  window["native"]["Save_End"](_header, oBinaryFileWriter.Memory.GetCurPosition());

		  return oBinaryFileWriter.Memory.ImData.data;
	  }
  };
  spreadsheet_api.prototype.asc_nativeCalculate = function() {
  };

	spreadsheet_api.prototype.getPrintOptionsJson = function() {
		var res = {};
		//отдельного флага для дополнительных настроек не делаю, общие опции тоже читаются, но пока не сделан мерж опций - верхние настройки перекроют внутренниие(если они заданы)
		res["spreadsheetLayout"] = {};
		res["spreadsheetLayout"]["ignorePrintArea"] = false;
		res["spreadsheetLayout"]["sheetsProps"] = this.wbModel && this.wbModel.getPrintOptionsJson();
		return res;
	};

  spreadsheet_api.prototype.asc_nativePrint = function (_printer, _page, _options) {
    //calc to fix case where file has formulas with no cache values and no changes
    this.wbModel.dependencyFormulas.calcTree();
    var _adjustPrint = (window.AscDesktopEditor_PrintOptions && window.AscDesktopEditor_PrintOptions.advancedOptions) || new Asc.asc_CAdjustPrint();
    window.AscDesktopEditor_PrintOptions = undefined;

    var pageSetup;
    var countWorksheets = this.wbModel.getWorksheetCount();
    var t = this;

    if(_options) {
      //печатаем только 1 страницу первой книги
      var isOnlyFirstPage = _options["printOptions"] && _options["printOptions"]["onlyFirstPage"];
       if(isOnlyFirstPage) {
		   _adjustPrint.isOnlyFirstPage = true;
       }
	   var spreadsheetLayout = _options["spreadsheetLayout"];
	   var _ignorePrintArea = spreadsheetLayout && (true === spreadsheetLayout["ignorePrintArea"] || false === spreadsheetLayout["ignorePrintArea"])? spreadsheetLayout["ignorePrintArea"] : null;
	   if(null !== _ignorePrintArea) {
		   _adjustPrint.asc_setIgnorePrintArea(_ignorePrintArea);
       } else {
		   _adjustPrint.asc_setIgnorePrintArea(true);
       }

	   _adjustPrint.asc_setPrintType(Asc.c_oAscPrintType.EntireWorkbook);

       var ws, newPrintOptions;
	   var _orientation = spreadsheetLayout ? spreadsheetLayout["orientation"] : null;
	   if(_orientation === "portrait") {
		   _orientation = Asc.c_oAscPageOrientation.PagePortrait;
       } else if(_orientation === "landscape") {
		   _orientation = Asc.c_oAscPageOrientation.PageLandscape;
       } else {
		   _orientation = null;
       }
       //need number
	   var _fitToWidth = spreadsheetLayout && AscCommon.isNumber( spreadsheetLayout["fitToWidth"]) ? spreadsheetLayout["fitToWidth"] : null;
	   var _fitToHeight = spreadsheetLayout && AscCommon.isNumber( spreadsheetLayout["fitToHeight"]) ? spreadsheetLayout["fitToHeight"] : null;
	   var _scale = spreadsheetLayout && AscCommon.isNumber( spreadsheetLayout["scale"]) ? spreadsheetLayout["scale"] : null;
	   //need true/false
	   var _headings = spreadsheetLayout && (true === spreadsheetLayout["headings"] || false === spreadsheetLayout["headings"])? spreadsheetLayout["headings"] : null;
	   var _gridLines = spreadsheetLayout && (true === spreadsheetLayout["gridLines"] || false === spreadsheetLayout["gridLines"])? spreadsheetLayout["gridLines"] : null;
       //convert to mm
	   var _pageSize = spreadsheetLayout ? spreadsheetLayout["pageSize"] : null;
	   if(_pageSize) {
	     var width = AscCommon.valueToMm(_pageSize["width"]);
	     var height = AscCommon.valueToMm(_pageSize["height"]);
	     _pageSize = null !== width && null !== height ? {width: width, height: height} : null;
       }
	   var _margins = spreadsheetLayout ? spreadsheetLayout["margins"] : null;
	   if(_margins) {
         var left = AscCommon.valueToMm(_margins["left"]);
         var right = AscCommon.valueToMm(_margins["right"]);
         var top = AscCommon.valueToMm(_margins["top"]);
         var bottom = AscCommon.valueToMm(_margins["bottom"]);
         _margins = null !== left && null !== right && null !== top && null !== bottom ? {left: left, right: right, top: top, bottom: bottom} : null;
       }

	   for (var index = 0; index < this.wbModel.getWorksheetCount(); ++index) {
           ws = this.wbModel.getWorksheet(index);
		   if (spreadsheetLayout && spreadsheetLayout["sheetsProps"] && spreadsheetLayout["sheetsProps"][index]) {
			   newPrintOptions = new Asc.asc_CPageOptions();
			   newPrintOptions.setJson(spreadsheetLayout["sheetsProps"][index]);
		   } else {
			   newPrintOptions = ws.PagePrintOptions.clone();
		   }
           //regionalSettings ?

           var _pageSetup = newPrintOptions.pageSetup;
           if (null !== _orientation) {
               _pageSetup.orientation = _orientation;
           }
           if (_fitToWidth || _fitToHeight) {
               _pageSetup.fitToWidth = _fitToWidth;
               _pageSetup.fitToHeight = _fitToHeight;
           } else if (_scale) {
               _pageSetup.scale = _scale;
               _pageSetup.fitToWidth = 0;
               _pageSetup.fitToHeight = 0;
           }
           if (null !== _headings) {
               newPrintOptions.headings = _headings;
           }
           if (null !== _gridLines) {
               newPrintOptions.gridLines = _gridLines;
           }
           if (_pageSize) {
               _pageSetup.width = _pageSize.width;
               _pageSetup.height = _pageSize.height;
           }
           var pageMargins = newPrintOptions.pageMargins;
           if (_margins) {
               pageMargins.left = _margins.left;
               pageMargins.right = _margins.right;
               pageMargins.top = _margins.top;
               pageMargins.bottom = _margins.bottom;
           }

           if (!_adjustPrint.pageOptionsMap) {
               _adjustPrint.pageOptionsMap = [];
           }
           _adjustPrint.pageOptionsMap[index] = newPrintOptions;
       }
    }

    this.wb.setPrintOptionsJson(spreadsheetLayout && spreadsheetLayout["sheetsProps"]);

    var _printPagesData = this.wb.calcPagesPrint(_adjustPrint);

    if (undefined === _printer && _page === undefined) {
      // ПУСТОЙ вызов, так как он должен быть ДО команд печати (картинки). А реальзый вызов - после (pagescount)
      window["AscDesktopEditor"] && window["AscDesktopEditor"]["Print_Start"]();
      this.wb.executeWithoutPreview(function () {
        _printer = t.wb.printSheets(_printPagesData, null, _adjustPrint).DocumentRenderer;
      });
      if (undefined !== window["AscDesktopEditor"]) {
        var pagescount = _printer.m_lPagesCount;

        window["AscDesktopEditor"]["Print_Start"](this.documentId + "/", pagescount, "", -1);

        for (var i = 0; i < pagescount; i++) {
          var _start = _printer.m_arrayPages[i].StartOffset;
          var _end = _printer.Memory.pos;
          if (i != (pagescount - 1)) {
            _end = _printer.m_arrayPages[i + 1].StartOffset;
          }

          window["AscDesktopEditor"]["Print_Page"](
			  _printer.Memory.GetBase64Memory2(_start, _end - _start),
			  _printer.m_arrayPages[i].Width, _printer.m_arrayPages[i].Height);
        }

        // detect fit flag
        var paramEnd = 0;
        if (_adjustPrint && _adjustPrint.asc_getPrintType() == Asc.c_oAscPrintType.EntireWorkbook) {
          for (var j = 0; j < countWorksheets; ++j) {
            if (_adjustPrint && _adjustPrint.pageOptionsMap && _adjustPrint.pageOptionsMap[j]) {
              pageSetup = _adjustPrint.pageOptionsMap[j].pageSetup;
            } else {
              pageSetup = this.wbModel.getWorksheet(j).PagePrintOptions.asc_getPageSetup();
            }
            if (pageSetup.asc_getFitToWidth() || pageSetup.asc_getFitToHeight())
              paramEnd |= 1;
          }
        }

        if (_adjustPrint && _adjustPrint.asc_getPrintType() == Asc.c_oAscPrintType.ActiveSheets) {
          var activeSheet = this.wbModel.getActive();
          if (_adjustPrint && _adjustPrint.pageOptionsMap && _adjustPrint.pageOptionsMap[activeSheet]) {
            pageSetup = _adjustPrint.pageOptionsMap[activeSheet].pageSetup;
          } else {
            pageSetup = this.wbModel.getWorksheet(activeSheet).PagePrintOptions.asc_getPageSetup();
          }
          if (pageSetup.asc_getFitToWidth() || pageSetup.asc_getFitToHeight())
            paramEnd |= 1;
        }

        window["AscDesktopEditor"]["Print_End"](paramEnd);
      }
    } else {
      this.wb.printSheets(_printPagesData, _printer, _adjustPrint);
    }

    this.wb.setPrintOptionsJson(null);

    return _printer.Memory;
  };

  spreadsheet_api.prototype.asc_nativePrintPagesCount = function() {
    return 1;
  };

  spreadsheet_api.prototype.asc_nativeGetPDF = function(options) {
    var _ret = this.asc_nativePrint(undefined, undefined, options);

    window["native"]["Save_End"]("", _ret.GetCurPosition());
    return _ret.data;
  };

  spreadsheet_api.prototype.asc_canPaste = function () {
    History.Create_NewPoint();
    History.StartTransaction();
    return true;
  };
  spreadsheet_api.prototype.asc_endPaste = function () {
    History.EndTransaction();
  };
  spreadsheet_api.prototype.asc_Recalculate = function () {
      History.EndTransaction();

	  var ws = this.wbModel.getActiveWs();
	  this.wbModel.handlers.trigger('showWorksheet', ws.getId());

      //в _onUpdateAfterApplyChanges нет очистки кэша, добавляю -
	  var lastPointIndex = History.Points && History.Points.length - 1;
	  var lastPoint = History.Points[lastPointIndex];
	  if (lastPoint && lastPoint.UpdateRigions) {
		  for (var i in lastPoint.UpdateRigions) {
			  this.wb.handlers.trigger("cleanCellCache", i, [lastPoint.UpdateRigions[i]], null, true);
		  }
	  }
      this._onUpdateAfterApplyChanges();
  };


  spreadsheet_api.prototype.pre_Paste = function(_fonts, _images, callback)
  {
    AscFonts.FontPickerByCharacter.extendFonts(_fonts);

    var oFontMap = {};
    for(var i = 0; i < _fonts.length; ++i){
      oFontMap[_fonts[i].name] = 1;
    }
    this._loadFonts(oFontMap, function() {

      var aImages = [];
      for(var key in _images){
        if(_images.hasOwnProperty(key)){
          aImages.push(_images[key])
        }
      }
      if(aImages.length > 0)      {
         window["Asc"]["editor"].ImageLoader.LoadDocumentImages(aImages, true);
      }
      callback();
    });
  };

  spreadsheet_api.prototype.getDefaultFontFamily = function () {
      return this.wbModel.getDefaultFont();
  };

  spreadsheet_api.prototype.getDefaultFontSize = function () {
      return this.wbModel.getDefaultSize();
  };

	spreadsheet_api.prototype._onEndLoadSdk = function () {
		AscCommon.baseEditorsApi.prototype._onEndLoadSdk.call(this);

		History = AscCommon.History;

		this.controller = new AscCommonExcel.asc_CEventsController();

		this.formulasList = AscCommonExcel.getFormulasInfo();
		this.asc_setLocale(this.tmpLCID, this.tmpDecimalSeparator, this.tmpGroupSeparator);
		this.asc_setLocalization(this.tmpLocalization);
		this.asc_setViewMode(this.isViewMode);

        if (this.openFileCryptBinary)
        {
            this.openFileCryptCallback(this.openFileCryptBinary);
        }
	};

	spreadsheet_api.prototype.asc_OnShowContextMenu = function() {
	  this.asc_closeCellEditor();
    };

	spreadsheet_api.prototype._changePivotSimple = function(pivot, isInsert, needUpdateView, callback) {
		var wsModel = pivot.GetWS();
		var ws = this.wb.getWorksheet(wsModel.getIndex());
		if (isInsert) {
			pivot.stashEmptyReportRange();
		} else {
			pivot.stashCurReportRange();
		}

		callback(wsModel);

		var dataRow;
		var pivotChanged = pivot.getAndCleanChanged();
		if (pivotChanged.data) {
			var updateRes = pivot.updateAfterEdit();
			dataRow = updateRes.dataRow;
		}
		this._updatePivotTable(pivot, pivotChanged, wsModel, ws, dataRow, needUpdateView, false);
	};
	spreadsheet_api.prototype.updatePivotTables = function() {
		var t = this;
		this.wbModel.forEach(function(wsModel) {
			var ws = t.wb.getWorksheet(wsModel.getIndex());
			for (var i = 0; i < wsModel.pivotTables.length; ++i) {
				var pivot = wsModel.pivotTables[i];
				t._updatePivotTable(pivot, pivot.getAndCleanChanged(), wsModel, ws, undefined, false, true);
			}
		});
	};
	spreadsheet_api.prototype._updatePivotTable = function(pivot, changed, wsModel, ws, dataRow, needUpdateView, canModifyDocument, doNotAutoFitColumnsWidth) {
		var ranges = wsModel.updatePivotTable(pivot, changed, dataRow, canModifyDocument);
		if (needUpdateView) {
			if (changed.oldRanges) {
				ws.updateRanges(changed.oldRanges);
			}
			if (ranges) {
				ws.updateRanges(ranges);
				if (pivot.useAutoFormatting && !doNotAutoFitColumnsWidth) {
					ws._autoFitColumnsWidth(ranges);
				}
			}
			//ws can be inactive in case of slicer on other sheet
			if (this.wbModel.getActive() === wsModel.getIndex()) {
				ws.draw();
			}
		}
	};

	spreadsheet_api.prototype.asc_getPivotInfo = function(opt_pivotTable) {
    var ws = this.wbModel.getActiveWs();
    var activeCell = ws.selectionRange.activeCell;
    var pivotTable = opt_pivotTable || ws.getPivotTable(activeCell.col, activeCell.row);
    if (pivotTable) {
      return pivotTable.getContextMenuInfo(ws.selectionRange);
    }
    return null;
	};
  // Uses for % of, difference from, % difference from, running total in, % running total in, % of parent
  spreadsheet_api.prototype.asc_getPivotShowValueAsInfo = function(showAs, opt_pivotTable) {
    function correctInfoPercentOfParent(layout) {
      if (layout.rows) {
        if (layout.rows.length == 1 && AscCommonExcel.st_VALUES !== layout.rows[layout.rows.length - 1].fld) {
          return layout.rows[layout.rows.length - 1];
        } else if (layout.rows.length > 1){
          for (let i = layout.rows.length - 2; i >= 0; i -= 1) {
            if (layout.rows[i].fld !== AscCommonExcel.st_VALUES) {
              return layout.rows[i];
            }
          }
          if (layout.rows[layout.rows.length - 1].fld !== AscCommonExcel.st_VALUES) {
            return layout.rows[layout.rows.length - 1];
          }
        }
      } else if (layout.cols) {
        if (layout.cols.length == 1 && AscCommonExcel.st_VALUES !== layout.cols[layout.cols.length - 1].fld) {
          return layout.cols[layout.cols.length - 1];
        } else if (layout.cols.length > 1){
          for (let i = layout.cols.length - 2; i >= 0; i -= 1) {
            if (layout.cols[i].fld !== AscCommonExcel.st_VALUES) {
              return layout.cols[i];
            }
          }
          if (layout.cols[layout.cols.length - 1].fld !== AscCommonExcel.st_VALUES) {
            return layout.cols[layout.cols.length - 1];
          }
        }
      }
      return null;
    }
    var res = null;
    var ws = this.wbModel.getActiveWs();
    var activeCell = ws.selectionRange.activeCell;
    var pivotTable = opt_pivotTable || ws.getPivotTable(activeCell.col, activeCell.row);
    if (pivotTable) {
      var layout = pivotTable.getLayoutByCell(activeCell.row, activeCell.col);
      var cellLayout = null;
      if (layout) {
        if (showAs === Asc.c_oAscShowDataAs.PercentOfParent) {
          cellLayout = correctInfoPercentOfParent(layout);
        } else {
          cellLayout = layout.getHeaderCellLayoutExceptValue();
        }
      }
      if (cellLayout !== null) {
        if (cellLayout.fld === null) {
          let rowFields = pivotTable.asc_getRowFields();
          let colFields = pivotTable.asc_getColumnFields();
          let rowFieldsFirstIndex = rowFields && rowFields[0].asc_getIndex();
          let firstIndex = rowFieldsFirstIndex;
          if (rowFieldsFirstIndex === null) {
            firstIndex = colFields && colFields[0].asc_getIndex();
          }
          cellLayout.fld = firstIndex;
        }
        res = new Asc.CT_DataField();
        res.baseField = cellLayout.fld;
        res.baseItem = cellLayout.v;
      }
    }
    return res;
  }
	spreadsheet_api.prototype.asc_createSheetName = function () {
		let sheetNames = [];
		let wc = this.asc_getWorksheetsCount();
		for(let i = 0; i < wc; i += 1) {
			sheetNames.push(this.asc_getWorksheetName(i).toLowerCase());
		}
		let name = 'Sheet';
		let i = 0;
		while (i += 1) {
			name = AscCommon.translateManager.getValue("Sheet") + i;
			if (sheetNames.indexOf(name.toLowerCase()) < 0) {
				break;
			}
		}
		return name;
	};
	/**
	* @param {CT_pivotTableDefinition} opt_pivotTable 
	* @return {boolean} Success
	*/
	spreadsheet_api.prototype.asc_pivotShowDetails = function(opt_pivotTable) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}
		let t = this;
		let ws = this.wbModel.getActiveWs();
		let activeCell = ws.selectionRange.activeCell;
		let pivotTable = opt_pivotTable || ws.getPivotTable(activeCell.col, activeCell.row);
		if (!pivotTable) {
			return false;
		}
		this._addWorksheetsCheck(function(res){
			if (!res) {
				return;
			}
			History.Create_NewPoint();
			History.StartTransaction();
			const indexes = pivotTable.getItemsIndexesByActiveCell(activeCell.row, activeCell.col);
			const itemMapArray = pivotTable.getNoFilterItemFieldsMapArray(indexes.rowItemIndex, indexes.colItemIndex)
			let worksheets = t._addWorksheetsWithoutLock([pivotTable.getShowDetailsSheetName(itemMapArray)], this.wbModel.getActive());
			let ws = worksheets[0];			

			let lengths = pivotTable.showDetails(ws, itemMapArray);

			let range = new Asc.Range(0, 0, lengths.colLength, lengths.rowLength);
			let ref = range.getAbsName();
			let options = t.asc_getAddFormatTableOptions(ref);
			let tableStyle = t.asc_getDefaultTableStyle();
			t.asc_addAutoFilter(tableStyle, options);

			History.EndTransaction();

			t.handlers.trigger("setSelection", range.clone());
		});
		return true;
	};
	/**
	 * @param {number} fld pivotFields index
	 * @param {boolean} isAll expand/collapse all
	 */
	spreadsheet_api.prototype.asc_pivotShowDetailsHeader = function(fld, isAll) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}
		let t = this;
		let ws = this.wbModel.getActiveWs();
		let activeCell = ws.selectionRange.activeCell;
		let pivotTable = ws.getPivotTable(activeCell.col, activeCell.row);
		if (!pivotTable) {
			return false;
		}
		pivotTable.showDetailsHeaderByCell(this, activeCell.row, activeCell.col, fld,  isAll);
	};

	spreadsheet_api.prototype.asc_getAddPivotTableOptions = function(range) {
		var ws = this.wb.getWorksheet();
		if (ws.model.getSelection().isSingleRange()) {
			return ws.getAddFormatTableOptions(range, true);
		} else {
			//todo move to getAddFormatTableOptions
			var res = new AscCommonExcel.AddFormatTableOptions();
			res.asc_setIsTitle(false);
			res.asc_setRange("");
		}
		return res;
	};
	spreadsheet_api.prototype.asc_insertPivotNewWorksheet = function(dataRef, newSheetName) {
		var t = this;
		if (Asc.CT_pivotTableDefinition.prototype.isValidDataRef(dataRef)) {
			var wb = this.wbModel;
			this._isLockedAddWorksheets(function(res) {
				if (res) {
					History.Create_NewPoint();
					History.StartTransaction();
					var worksheets = t._addWorksheetsWithoutLock([newSheetName], wb.getActive());
					var ws = worksheets[0];
					var range = new Asc.Range(AscCommonExcel.NEW_PIVOT_COL, AscCommonExcel.NEW_PIVOT_ROW, AscCommonExcel.NEW_PIVOT_COL, AscCommonExcel.NEW_PIVOT_ROW);
					t._asc_insertPivot(wb, dataRef, ws, range, false);
					History.EndTransaction();
				} else {
					//todo
					t.sendEvent('asc_onError', c_oAscError.ID.LockedCellPivot, c_oAscError.Level.NoCritical);
				}
			});
		} else {
			this.sendEvent('asc_onError', c_oAscError.ID.PivotLabledColumns, c_oAscError.Level.NoCritical);
		}
	};
	spreadsheet_api.prototype.asc_insertPivotExistingWorksheet = function(dataRef, pivotRef) {
		if (!Asc.CT_pivotTableDefinition.prototype.isValidDataRef(dataRef)) {
			this.sendEvent('asc_onError', c_oAscError.ID.PivotLabledColumns, c_oAscError.Level.NoCritical);
			return;
		}
		var location = AscFormat.ExecuteNoHistory(function() {
			return Asc.CT_pivotTableDefinition.prototype.parseDataRef(pivotRef);
		}, this, []);
		if (location) {
			var wb = this.wbModel;
			var ws = location.ws;
			if (ws) {
				this.wbModel.setActiveById(ws.getId());
			} else {
				ws = wb.getActiveWs();
			}
			this.wb.updateWorksheetByModel();
			this.wb.showWorksheet();
			this._asc_insertPivot(wb, dataRef, ws, location.bbox, false);
		}
	};
	spreadsheet_api.prototype._asc_insertPivot = function(wb, dataRef, ws, bbox, confirmation) {
		var t = this;
		History.Create_NewPoint();
		History.StartTransaction();
		var pivotName = wb.dependencyFormulas.getNextPivotName();
		var pivot = new Asc.CT_pivotTableDefinition(true);
		var dataLocation = AscFormat.ExecuteNoHistory(function() {
			return Asc.CT_pivotTableDefinition.prototype.parseDataRef(dataRef);
		}, this, []);
		var cacheDefinition = wb.getPivotCacheByDataLocation(dataLocation);
		if (!cacheDefinition) {
			cacheDefinition = new Asc.CT_PivotCacheDefinition();
			cacheDefinition.asc_create();
			cacheDefinition.fromDataRef(dataRef);
		}
		pivot.asc_create(ws, pivotName, cacheDefinition, bbox);
		pivot.stashEmptyReportRange();
		this._changePivotWithLockExt(pivot, confirmation, true, function(ws, pivot) {
			ws.insertPivotTable(pivot, true, false);
			pivot.setChanged(true);
		});
		History.EndTransaction();
		return pivot;
	};
	spreadsheet_api.prototype.asc_refreshAllPivots = function(opt_confirmation) {
		var t = this;
		let pivotTables = [];
		this.wbModel.forEach(function(ws) {
			for (var i = 0; i < ws.pivotTables.length; ++i) {
				pivotTables.push(ws.pivotTables[i]);
			}
		});
		if (0 === pivotTables.length) {
			return;
		}
		this._isLockedPivotAndConnectedByPivotCache(pivotTables, function(res) {
			if (!res) {
				t.sendEvent('asc_onError', c_oAscError.ID.PivotOverlap, c_oAscError.Level.NoCritical);
				return;
			}
			History.Create_NewPoint();
			History.StartTransaction();
			t.wbModel.dependencyFormulas.lockRecal();

			let changeRes;
			for (let i = pivotTables.length - 1; i >= 0; --i) {
				let checkRefresh = pivotTables[i].checkRefresh();
				if (c_oAscError.ID.No === checkRefresh) {
					changeRes = t._changePivot(pivotTables[i], opt_confirmation, false, function(ws, pivot) {
						let error = pivot.refresh();
					});
				} else {
					changeRes = {error: checkRefresh, warning: c_oAscError.ID.No, updateRes: undefined};
				}
				if (c_oAscError.ID.No !== changeRes.error || c_oAscError.ID.No !== changeRes.warning) {
					break;
				}
			}
			t.wbModel.dependencyFormulas.unlockRecal();
			History.EndTransaction();
			t._changePivotEndCheckError(changeRes, function(){
				t.asc_refreshAllPivots(true);
			});
		});

	};
	spreadsheet_api.prototype._isLockedPivot = function (pivot, callback) {
		var lockInfos = [];
		pivot.fillLockInfo(lockInfos, this.collaborativeEditing);
		this.collaborativeEditing.lock(lockInfos, callback);
	};
	spreadsheet_api.prototype._isLockedPivotAndConnectedBySlicer = function (pivot, flds, callback) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			callback(false);
			return;
		}
		var t = this;
		var lockInfos = [];
		pivot.fillLockInfo(lockInfos, this.collaborativeEditing);
		flds.forEach(function(fld){
			var pivotTables = pivot.getPivotTablesConnectedBySlicer(fld);
			pivotTables.forEach(function(pivotTable){
				pivotTable.fillLockInfo(lockInfos, t.collaborativeEditing);
			});
		});
		this.collaborativeEditing.lock(lockInfos, callback);
	};
	spreadsheet_api.prototype._isLockedPivotAndConnectedByPivotCache = function (pivotTables, callback) {
		if (this.collaborativeEditing.getGlobalLock()) {
			callback(false);
			return;
		}
		var t = this;
		var lockInfos = [];
		pivotTables.forEach(function(pivotTable) {
			pivotTable.fillLockInfo(lockInfos, t.collaborativeEditing);
		});
		this.collaborativeEditing.lock(lockInfos, callback);
	};
	spreadsheet_api.prototype._changePivotWithLock = function (pivot, onAction, doNotCheckUnderlyingData, doNotAutoFitColumnsWidth) {
		this._changePivotWithLockExt(pivot, false, true, onAction, doNotCheckUnderlyingData, doNotAutoFitColumnsWidth);
	};
	spreadsheet_api.prototype._changePivotWithLockExt = function (pivot, confirmation, updateSelection, onAction, doNotCheckUnderlyingData, doNotAutoFitColumnsWidth) {
		// Проверка глобального лока
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		var t = this;
		this._isLockedPivot(pivot, function (res) {
			if (!res) {
				t.sendEvent('asc_onError', c_oAscError.ID.PivotOverlap, c_oAscError.Level.NoCritical);
				return;
			}
			History.Create_NewPoint();
			History.StartTransaction();
			t.wbModel.dependencyFormulas.lockRecal();
			var changeRes = t._changePivot(pivot, confirmation, updateSelection, onAction, doNotCheckUnderlyingData, doNotAutoFitColumnsWidth);
			t.wbModel.dependencyFormulas.unlockRecal();
			History.EndTransaction();
			t._changePivotEndCheckError(changeRes, function () {
				//undo can replace pivot complitly. note: getPivotTableById returns nothing while insert operation
				var pivotAfterUndo = t.wbModel.getPivotTableById(pivot.Get_Id()) || pivot;
				t._changePivotWithLockExt(pivotAfterUndo, true, updateSelection, onAction);
			});
		});
	};
	spreadsheet_api.prototype._changePivotAndConnectedBySlicerWithLock = function (pivot, flds, onAction) {
		// Проверка глобального лока

		var t = this;
		this._isLockedPivotAndConnectedBySlicer(pivot, flds, function(res) {
			if (!res) {
				t.sendEvent('asc_onError', c_oAscError.ID.PivotOverlap, c_oAscError.Level.NoCritical);
				return;
			}
			onAction();
		});
	};
	spreadsheet_api.prototype._changePivotAndConnectedByPivotCacheWithLock = function (pivot, confirmation, onAction, onRepeat) {
		// Проверка глобального лока

		var t = this;
		var pivotTables = pivot.getPivotTablesConnectedByPivotCache();
		this._isLockedPivotAndConnectedByPivotCache(pivotTables, function(res) {
			if (!res) {
				t.sendEvent('asc_onError', c_oAscError.ID.PivotOverlap, c_oAscError.Level.NoCritical);
				return;
			}
			History.Create_NewPoint();
			History.StartTransaction();
			t.wbModel.dependencyFormulas.lockRecal();
			var changeRes = onAction(confirmation, pivotTables);
			t.wbModel.dependencyFormulas.unlockRecal();
			History.EndTransaction();
			t._changePivotEndCheckError(changeRes, onRepeat);
		});
	};
	spreadsheet_api.prototype._changePivot = function(pivot, confirmation, updateSelection, onAction, doNotCheckUnderlyingData, doNotAutoFitColumnsWidth) {
		if (!doNotCheckUnderlyingData && !pivot.checkPivotUnderlyingData()) {
			return {error: c_oAscError.ID.PivotWithoutUnderlyingData, warning: c_oAscError.ID.No, updateRes: undefined};
		}
		var wsModel = pivot.GetWS();
		pivot.stashCurReportRange();

		onAction(wsModel, pivot);

		var pivotChanged = pivot.getAndCleanChanged();
		var error = c_oAscError.ID.No;
		var warning = c_oAscError.ID.No;
		var updateRes, reportRanges;
		if (pivotChanged.data) {
			updateRes = pivot.updateAfterEdit();
			reportRanges = pivot.getReportRanges();
			error = wsModel.checkPivotReportLocationForError(reportRanges, pivot);
			if (c_oAscError.ID.No === error) {
				//todo remove cleanAll from checkPivotReportLocationForConfirm
				warning = wsModel.checkPivotReportLocationForConfirm(reportRanges, pivotChanged);
				if (confirmation) {
					warning = c_oAscError.ID.No;
				}
			}
		}
		var isSuccess = c_oAscError.ID.No === error && c_oAscError.ID.No === warning;
		if (isSuccess) {
			var ws = this.wb.getWorksheet(wsModel.getIndex());
			this._updatePivotTable(pivot, pivotChanged, wsModel, ws, updateRes && updateRes.dataRow, true, true, doNotAutoFitColumnsWidth);
			if (updateSelection) {
				pivot.updateSelection(ws);
			}
			//ws can be inactive in case of slicer on other sheet
			if (this.wbModel.getActive() === wsModel.getIndex()) {
				ws.draw();
			}
		} else {
			pivot.stashEmptyReportRange();//to prevent clearTableStyle while undo
		}
		return {error: error, warning: warning, updateRes: updateRes};
	};
	spreadsheet_api.prototype._changePivotRevert = function () {
		History.Undo();
		History.Clear_Redo();
		this._onUpdateDocumentCanUndoRedo();
	};
	spreadsheet_api.prototype._changePivotEndCheckError = function (changeRes, onRepeat) {
		if (!changeRes) {
			return true;
		}
		if (c_oAscError.ID.No !== changeRes.error) {
			this._changePivotRevert();
			this.sendEvent('asc_onError', changeRes.error, c_oAscError.Level.NoCritical);
		} else if (c_oAscError.ID.No !== changeRes.warning) {
			this._changePivotRevert();
			this.handlers.trigger("asc_onConfirmAction", Asc.c_oAscConfirm.ConfirmReplaceRange, function (can) {
				if (can) {
					//repeate with whole checks because of collaboration changes
					onRepeat();
				}
			});
		} else {
			return true;
		}
		return false;
	};

	spreadsheet_api.prototype._selectSearchingResults = function (bShow) {
	  var ws = this.wbModel.getActiveWs();
	  if (!bShow) {
		  this.wb.drawWS();
	  } else if (ws && ws.lastFindOptions) {
	    this.wb.drawWS();
      }
	};
	spreadsheet_api.prototype.asc_getAppProps = function()
	{
		return this.wbModel && this.wbModel.App || null;
	};

    spreadsheet_api.prototype.checkObjectsLock = function(aObjectsId, fCallback) {
      if(!this.collaborativeEditing) {
        fCallback(true, true);
        return;
      }
      this.collaborativeEditing.checkObjectsLock(aObjectsId, fCallback);
    };
    spreadsheet_api.prototype.asc_setCoreProps = function(oProps)
    {
      var oCore = this.getInternalCoreProps();
      if(!oCore) {
        return;
      }
      this.checkObjectsLock([oCore.Get_Id()], function(bNoLock) {
        if(bNoLock) {
          History.Create_NewPoint();
          oCore.setProps(oProps);
        }
      });
      return null;
    };

	spreadsheet_api.prototype.getInternalCoreProps = function()
	{
      return this.wbModel && this.wbModel.Core;
	};

	spreadsheet_api.prototype.asc_setGroupSummary = function (val, bCol) {
		var ws = this.wb && this.wb.getWorksheet();
		if(ws) {
			ws.asc_setGroupSummary(val, bCol);
		}
	};

	spreadsheet_api.prototype.asc_getGroupSummaryRight = function () {
		var ws = this.wbModel.getActiveWs();
		return ws && ws.sheetPr ? ws.sheetPr.SummaryRight : true;
	};

	spreadsheet_api.prototype.asc_getGroupSummaryBelow = function () {
		var ws = this.wbModel.getActiveWs();
		return ws && ws.sheetPr ? ws.sheetPr.SummaryBelow : true;
	};

	spreadsheet_api.prototype.asc_getSortProps = function (bExpand) {
		var ws = this.wb && this.wb.getWorksheet();
		if(ws) {
			return ws.getSortProps(bExpand);
		}
	};

	spreadsheet_api.prototype.asc_setSortProps = function (props, bCancel) {
		var ws = this.wb && this.wb.getWorksheet();
		if(ws) {
		  if(bCancel) {
		    ws.setSortProps(props, null, true);
          }	else {
		    ws.setSelectionInfo("customSort", props);
          }
		}
	};
	spreadsheet_api.prototype.asc_validSheetName = function (val) {
		return window["AscCommon"].rx_test_ws_name.isValidName(val);
	};
    spreadsheet_api.prototype.asc_getRemoveDuplicates = function (bExpand) {
      var ws = this.wb && this.wb.getWorksheet();
      if(ws) {
        return ws.getRemoveDuplicates(bExpand);
      }
    };
    spreadsheet_api.prototype.asc_setRemoveDuplicates = function (props, bCancel) {
      if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
        return false;
      }
      var ws = this.wb && this.wb.getWorksheet();
      if(ws) {
        return ws.setRemoveDuplicates(props, bCancel);
      }
    };

	spreadsheet_api.prototype.asc_getCF = function (type, id) {
		var sheet;
		var rules = this.wbModel.getRulesByType(type, id, true);
		var aSheet = type === Asc.c_oAscSelectionForCFType.selection ? sheet : this.wbModel.getActiveWs();
		var activeRanges = aSheet.selectionRange.ranges;
		var sActiveRanges = [];
		if (activeRanges) {
			activeRanges.forEach(function (item) {
				sActiveRanges.push(item.getAbsName());
			});
		}

		return [rules, "=" + sActiveRanges.join(AscCommon.FormulaSeparators.functionArgumentSeparator)];
	};

	spreadsheet_api.prototype.asc_getPreviewCF = function (id, props, text) {
		if (!props && text) {
			props = new AscCommonExcel.CellXfs();
		}
		if (props) {
			props.asc_getPreview2(this, id, text);
		}
	};

	spreadsheet_api.prototype.asc_setCF = function (arr, deleteIdArr, presetId) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}

		var ws = this.wb.getWorksheet();
		ws.setCF(arr, deleteIdArr, presetId);
	};

	spreadsheet_api.prototype.asc_clearCF = function (type, id) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}

		var rules = this.wbModel.getRulesByType(type, id);
		if (rules && rules.length) {
			var ws = this.wb.getWorksheet();
			ws.deleteCF(rules, type);
		}
	};

	spreadsheet_api.prototype._onUpdateCFLock = function (lockElem) {
		var t = this;
		var sheetId = lockElem.Element["sheetId"];
		if (-1 !== sheetId && 0 === sheetId.indexOf(AscCommonExcel.CConditionalFormattingRule.sStartLockCFId)) {
			sheetId = sheetId.split(AscCommonExcel.CConditionalFormattingRule.sStartLockCFId)[1];
			var wsModel = this.wbModel.getWorksheetById(sheetId);
			if (wsModel) {
				var wsIndex = wsModel.getIndex();
				var cFRule = wsModel.getCFRuleById(lockElem.Element["rangeOrObjectId"]);
				if (cFRule && cFRule.val) {
					cFRule = cFRule.val;
					cFRule.isLock = lockElem.UserId;
					this.handlers.trigger("asc_onLockCFRule", wsIndex, cFRule.id, lockElem.UserId);
				} else {
					var wsView = this.wb.getWorksheetById(sheetId);
					wsView._lockAddNewRule = true;
				}
				this.handlers.trigger("asc_onLockCFManager", wsModel.index);
			}
		}
	};

	spreadsheet_api.prototype._onUnlockCF = function () {
		var t = this;
		if (t.wbModel) {
			var i, length, wsModel, wsIndex;
			for (i = 0, length = t.wbModel.getWorksheetCount(); i < length; ++i) {
				wsModel = t.wbModel.getWorksheet(i);
				wsIndex = wsModel.getIndex();
				//TODO необходимо добавить инофрмацию о локе нового добавленного правила!!!

				var isLockedRules = false;
				if (wsModel.aConditionalFormattingRules && wsModel.aConditionalFormattingRules.length) {
					wsModel.aConditionalFormattingRules.forEach(function (_rule) {
						if (_rule.isLock) {
							isLockedRules = true;
						}
					});
					if (!isLockedRules) {
						var wsView = this.wb.getWorksheetById(wsModel.Id);
						if (wsView._lockAddNewRule) {
							isLockedRules = true;
						}
					}
				}
				if (!isLockedRules) {
					t.handlers.trigger("asc_onUnLockCFManager", wsIndex);
				}
			}
		}
	};

	spreadsheet_api.prototype.asc_getFullCFIcons = function () {
		return AscCommonExcel.getFullCFIcons();
	};

	spreadsheet_api.prototype.asc_getCFPresets = function () {
		return AscCommonExcel.getFullCFPresets();
	};

	spreadsheet_api.prototype.asc_getCFIconsByType = function () {
		return AscCommonExcel.getCFIconsByType();
	};

	spreadsheet_api.prototype._onCheckCFRemoveLock = function (lockElem) {
		//лок правила - с правилом делать ничего нельзя
		//лок менеджера - незалоченное правило можно удалять и редактировать. новые правила добавлять нельзя.
		//так же нельзя перемещать местами правила

		//лочим правило как объект. в лок кладём id и лист с префиксом CConditionalFormattingRule.sStartLockCFId
		//на принятии изменений удаляем локи с соответсвующих элементов
		//разлочиваем менеджер если нет залоченных элементов(т.е. проверяем все на лок)
		//+ проверяем нет ли нового добавленного правила другим юзером
		//всего для передачи в интерфейс 4 события - asc_onLockCFRule/asc_onUnLockCFRule; asc_onLockCFManager/asc_onUnLockCFManager

		var res = false;
		var t = this;
		var sheetId = lockElem["sheetId"];
		if (-1 !== sheetId && 0 === sheetId.indexOf(AscCommonExcel.CConditionalFormattingRule.sStartLockCFId)) {
			res = true;
			if (t.wbModel) {
				sheetId = sheetId.split(AscCommonExcel.CConditionalFormattingRule.sStartLockCFId)[1];
				var wsModel = t.wbModel.getWorksheetById(sheetId);
				if (wsModel) {
					var wsIndex = wsModel.getIndex();
					var wsView = this.wb.getWorksheetById(sheetId);
					var cFRule = wsModel.getCFRuleById(lockElem["rangeOrObjectId"]);
					if (cFRule) {
						if (cFRule.val.isLock) {
							cFRule.val.isLock = null;
						} else {
							wsView._lockAddNewRule = null;
						}
						this.handlers.trigger("asc_onUnLockCFRule", wsIndex, lockElem["rangeOrObjectId"]);
					} else {
						wsView._lockAddNewRule = null;
					}
				}
			}
		}
		return res;
	};

	spreadsheet_api.prototype.asc_isValidDataRefCf = function (type, props) {
		return AscCommonExcel.isValidDataRefCf(type, props);
	};

  spreadsheet_api.prototype.asc_beforeInsertSlicer = function () {
    //пока возвращаю только данные о ф/т
    return this.wb.beforeInsertSlicer();
  };
  spreadsheet_api.prototype.asc_insertSlicer = function (arr) {
    //пока возвращаю только данные о ф/т
    return this.wb.insertSlicers(arr);
  };

  spreadsheet_api.prototype.asc_setSlicers = function (names, obj) {
    return this.wb.setSlicers(names, obj);
  };
  spreadsheet_api.prototype.asc_Remove = function() {
    var ws = this.wb.getWorksheet();
    if (ws && ws.objectRender && ws.objectRender.controller) {
      ws.objectRender.controller.resetChartElementsSelection();
    }
    AscCommon.baseEditorsApi.prototype.asc_Remove.call(this);
  };

  spreadsheet_api.prototype.asc_getDataValidationProps = function(extend) {
    //если активная область затрагивает частично ячейку с date validation, частично без - выдаем предупреждение
    //второе предупреждение - если выделено несколько разных ячеек с разными data validation
    //возвращаем либо id ошибки, либо объект для диалога
    var ws = this.wbModel.getActiveWs();
    if (ws) {
      return ws.getDataValidationProps(extend);
    }
  };

  spreadsheet_api.prototype.asc_setDataValidation = function(props, switchApplySameSettings) {
    if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
      return false;
    }
    var ws = this.wb.getWorksheet();
    if (ws) {
      if (undefined === switchApplySameSettings) {
        return ws.setDataValidationProps(props);
      } else {
		  return ws.setDataValidationSameSettings(props, switchApplySameSettings);
      }
    }
  };

  spreadsheet_api.prototype.asc_getProtectedRanges = function () {
    var ws = this.wbModel.getActiveWs();
    if (ws) {
      return ws.getProtectedRanges(true);
    }
  };

	spreadsheet_api.prototype.asc_setProtectedRanges = function (arr, deleteArr) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}

		if (this.wb && this.wb.getCellEditMode()) {
			return;
			//this.asc_closeCellEditor(true);
		}

		var ws = this.wb.getWorksheet();
		var t = this;

		var checkPassword = function (hash, doNotCheckPassword) {
			if (doNotCheckPassword) {
				ws.setProtectedRanges(arr, deleteArr);
			} else {
				var j = 0;
				for (var i = 0; i < arr.length; i++) {
					if (arr[i].temporaryPassword) {
						arr[i].hashValue = hash[j];
						arr[i].temporaryPassword = null;
						j++;
					}
				}

				ws.setProtectedRanges(arr, deleteArr);
			}

			t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction);
		};

		var aCheckHash = [];
		for (var i = 0; i < arr.length; i++) {
			if (arr[i].temporaryPassword) {
				aCheckHash.push({
					password: arr[i].temporaryPassword,
					salt: arr[i].saltValue,
					spinCount: arr[i].spinCount,
					alg: AscCommon.fromModelAlgorithmName(arr[i].algorithmName)
				});
			}
		}

		this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction);
		if (aCheckHash.length) {
			AscCommon.calculateProtectHash(aCheckHash, checkPassword);
		} else {
			checkPassword(null, true);
		}
	};

  spreadsheet_api.prototype._onUpdateProtectedRangesLock = function (lockElem) {
    var t = this;
    var sheetId = lockElem.Element["sheetId"];
    if (-1 !== sheetId && 0 === sheetId.indexOf(Asc.CProtectedRange.sStartLock)) {
      sheetId = sheetId.split(Asc.CProtectedRange.sStartLock)[1];
      var wsModel = this.wbModel.getWorksheetById(sheetId);
      if (wsModel) {
        var wsIndex = wsModel.getIndex();
        var protectedRange = wsModel.getProtectedRangeById(lockElem.Element["rangeOrObjectId"]);
        if (protectedRange && protectedRange.val) {
          protectedRange = protectedRange.val;
          protectedRange.isLock = lockElem.UserId;
          this.handlers.trigger("asc_onLockProtectedRange", wsIndex, protectedRange.Id, lockElem.UserId);
        } else {
          var wsView = this.wb.getWorksheetById(sheetId);
          wsView._lockAddProtectedRange = true;
        }
        this.handlers.trigger("asc_onLockProtectedRangeManager", wsModel.index);
      }
    }
  };

  spreadsheet_api.prototype._onUnlockProtectedRange = function () {
    var t = this;
    if (t.wbModel) {
      var i, length, wsModel, wsIndex;
      for (i = 0, length = t.wbModel.getWorksheetCount(); i < length; ++i) {
        wsModel = t.wbModel.getWorksheet(i);
        wsIndex = wsModel.getIndex();

        var isLocked = false;
        if (wsModel.aProtectedRanges && wsModel.aProtectedRanges.length) {
          wsModel.aProtectedRanges.forEach(function (pR) {
            if (pR.isLock) {
              isLocked = true;
            }
          });
          if (!isLocked) {
            var wsView = this.wb.getWorksheetById(wsModel.Id);
            if (wsView._lockAddProtectedRange) {
              isLocked = true;
            }
          }
        }
        if (!isLocked) {
          t.handlers.trigger("asc_onUnLockProtectedRangeManager", wsIndex);
        }
      }
    }
  };

  spreadsheet_api.prototype.asc_checkProtectedRangesPassword = function (val, data, callback) {
	  var ws = this.wb.getWorksheet();
	  return ws.model.checkProtectedRangesPassword(val, data, callback);
  };

  spreadsheet_api.prototype._onCheckProtectedRangeRemoveLock = function (lockElem) {
    //лок правила - с правилом делать ничего нельзя
    //лок менеджера - незалоченное правило можно удалять и редактировать. новые правила добавлять нельзя.
    //так же нельзя перемещать местами правила

    //лочим правило как объект. в лок кладём id и лист с префиксом Asc.CProtectedRange.sStartLock
    //на принятии изменений удаляем локи с соответсвующих элементов
    //разлочиваем менеджер если нет залоченных элементов(т.е. проверяем все на лок)
    //+ проверяем нет ли нового добавленного правила другим юзером
    //всего для передачи в интерфейс 4 события - asc_onLockProtectedRange/asc_onUnLockProtectedRange; asc_onLockProtectedRangeManager/asc_onUnLockProtectedRangeManager

    var res = false;
    var t = this;
    var sheetId = lockElem["sheetId"];
    if (-1 !== sheetId && 0 === sheetId.indexOf(Asc.CProtectedRange.sStartLock)) {
      res = true;
      if (t.wbModel) {
        sheetId = sheetId.split(Asc.CProtectedRange.sStartLock)[1];
        var wsModel = t.wbModel.getWorksheetById(sheetId);
        if (wsModel) {
          var wsIndex = wsModel.getIndex();
          var wsView = this.wb.getWorksheetById(sheetId);
          var protectedRange = wsModel.getProtectedRangeById(lockElem["rangeOrObjectId"]);
          if (protectedRange) {
            if (protectedRange.val.isLock) {
              protectedRange.val.isLock = null;
            } else {
              wsView._lockAddProtectedRange = null;
            }
            this.handlers.trigger("asc_onUnLockProtectedRange", wsIndex, lockElem["rangeOrObjectId"]);
          } else {
            wsView._lockAddProtectedRange = null;
          }
        }
      }
    }
    return res;
  };

  spreadsheet_api.prototype.asc_checkProtectedRangeName = function(checkName) {
  	if (!this.wbModel) {
		return;
	}
	var ws = this.wbModel.getActiveWs();
  	return ws.checkProtectedRangeName(checkName);
  };

  spreadsheet_api.prototype.asc_getProtectedSheet = function () {
    if (!this.wbModel) {
    	return;
    }
    var ws = this.wbModel.getActiveWs();
    var res = null;
    if (ws) {
      if (ws.sheetProtection) {
        res = ws.sheetProtection.clone();
      } else {
        res = new window["Asc"].CSheetProtection();
        res.setDefaultInterface();
      }
    }
    return res;
  };

  spreadsheet_api.prototype.asc_isProtectedSheet = function (index) {
    if (!this.wbModel) {
    	return;
    }
    var sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
    var ws = this.wb.getWorksheet(sheetIndex);
    var res = null;
    if (ws) {
      if (ws.model.getSheetProtection()) {
        res = true;
      }
    }
    return res;
  };

	spreadsheet_api.prototype.asc_setProtectedSheet = function (props) {
		// Проверка глобального лока
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit() || !this.wb) {
			return false;
		}

		var wsView = this.wb.getWorksheet();
		var ws = wsView && wsView.model;
		if (ws && ws.isUserProtectedRangesIntersection(null, null, true)) {
			this.handlers.trigger("asc_onError", c_oAscError.ID.ProtectedRangeByOtherUser, c_oAscError.Level.NoCritical);
			return;
		}

		if (!props) {
			this.handlers.trigger("asc_onError", c_oAscError.ID.PasswordIsNotCorrect, c_oAscError.Level.NoCritical);
			this.handlers.trigger("asc_onChangeProtectWorksheet", i);
			return;
		}

		if (this.wb && this.wb.getCellEditMode()) {
			return;
			//this.asc_closeCellEditor(true);
		}

		var i = this.wbModel.getActive();
		var sheetId = this.wbModel.getWorksheet(i).getId();
		var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheetId,
			sheetId);
		var t = this;

		var callback = function (res) {
			t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction);

			if (res) {
				History.Create_NewPoint();
				History.StartTransaction();
				if (!t.wbModel.getWorksheet(i).setProtectedSheet(props, true)) {
					t.handlers.trigger("asc_onError", c_oAscError.ID.LockedWorksheetRename,
						c_oAscError.Level.NoCritical);
				} else if (wsView) {
					wsView.updateAfterChangeSheetProtection();
				}
				t.handlers.trigger("asc_onChangeProtectWorksheet", i);

				History.EndTransaction();
			} else {
				//t.handlers.trigger("asc_onError", c_oAscError.ID.LockedWorksheetRename, c_oAscError.Level.NoCritical);
			}
		};

		var checkPassword = function (hash, doNotCheckPassword) {
			if (doNotCheckPassword) {
				t.collaborativeEditing.lock([lockInfo], callback);
			} else {
				if (props.sheet) {
					props.hashValue = hash && hash[0] ? hash[0] : null;
					t.collaborativeEditing.lock([lockInfo], callback);
				} else {
					if (props.isPasswordXL() && hash && hash[0] && hash[0].toLowerCase() === props.password.toLowerCase()) {
						props.password = null;
						t.collaborativeEditing.lock([lockInfo], callback);
					} else if (!props.isPasswordXL() && hash && hash[0] === props.hashValue) {
						props.hashValue = null;
						props.saltValue = null;
						props.spinCount = null;
						props.algorithmName = null;
						t.collaborativeEditing.lock([lockInfo], callback);
					} else {
						//неверный пароль
						t.handlers.trigger("asc_onError", c_oAscError.ID.PasswordIsNotCorrect,
							c_oAscError.Level.NoCritical);
						t.handlers.trigger("asc_onChangeProtectWorksheet", i);
						t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction);
					}
				}
				props.temporaryPassword = null;
			}
		};

		this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction);
		if (props && props.temporaryPassword != null) {
			if (props.temporaryPassword === "") {
				checkPassword([""]);
			} else if (props.isPasswordXL()) {
				checkPassword([AscCommonExcel.getPasswordHash(props.temporaryPassword, true)]);
			} else {
				var checkHash = {password: props.temporaryPassword, salt: props.saltValue, spinCount: props.spinCount,
					alg: AscCommon.fromModelAlgorithmName(props.algorithmName)};
				AscCommon.calculateProtectHash([checkHash], checkPassword);
			}
		} else {
			checkPassword(null, true);
		}

		return true;
	};

	spreadsheet_api.prototype.asc_getProtectedWorkbook = function () {
		var wb = this.wbModel;
		var res = null;
		if (wb) {
			if (wb.workbookProtection) {
				res = wb.workbookProtection.clone();
			} else {
				res = new asc.CWorkbookProtection();
			}
		}
		return res;
	};

	spreadsheet_api.prototype.asc_isProtectedWorkbook = function (type) {
		if (!this.wbModel) {
			return;
		}
		var wb = this.wbModel;
		return wb && wb.getWorkbookProtection(type);
	};

	spreadsheet_api.prototype.asc_setProtectedWorkbook = function (props) {
		// Проверка глобального лока
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}

		if (!props) {
			this.handlers.trigger("asc_onError", c_oAscError.ID.PasswordIsNotCorrect, c_oAscError.Level.NoCritical);
			this.handlers.trigger("asc_onChangeProtectWorkbook");
			return;
		}

		if (this.wb && this.wb.getCellEditMode()) {
			return;
			//this.asc_closeCellEditor(true);
		}

		var wb = this.wbModel;
		var sheet, i, arrLocks = [];
		for (i = 0; i < wb.aWorksheets.length; ++i) {
			sheet = wb.aWorksheets[i];
			arrLocks.push(
				this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Sheet, /*subType*/null, sheet.getId(),
					sheet.getId()));
		}

		var t = this;
		var callback = function (res) {
			t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction);
			if (res) {
				History.Create_NewPoint();
				History.StartTransaction();
				if (!t.wbModel.setProtectedWorkbook(props, true)) {
					t.handlers.trigger("asc_onError", c_oAscError.ID.LockedWorksheetRename,
						c_oAscError.Level.NoCritical);
				}
				t.handlers.trigger("asc_onChangeProtectWorkbook");
				History.EndTransaction();
			} else {
				//t.handlers.trigger("asc_onError", c_oAscError.ID.LockedWorksheetRename, c_oAscError.Level.NoCritical);
			}
		};


		var checkPassword = function (hash, doNotCheckPassword) {
			if (doNotCheckPassword) {
				//TODO проверить, может быть нужен глобальный лок?
				t.collaborativeEditing.lock(arrLocks, callback);
			} else {
				if (props.lockStructure) {
					props.workbookHashValue = hash && hash[0] ? hash[0] : null;
					t.collaborativeEditing.lock(arrLocks, callback);
				} else {
					if (props.isPasswordXL() && hash && hash[0] && hash[0].toLowerCase() === props.workbookPassword.toLowerCase()) {
						props.workbookPassword = null;
						t.collaborativeEditing.lock(arrLocks, callback);
					} else if (!props.isPasswordXL() && hash && hash[0] === props.workbookHashValue) {
						props.workbookHashValue = null;
						props.workbookSaltValue = null;
						props.workbookSpinCount = null;
						props.workbookAlgorithmName = null;
						//TODO проверить, может быть нужен глобальный лок?
						t.collaborativeEditing.lock(arrLocks, callback);
					} else {
						//неверный пароль
						t.handlers.trigger("asc_onError", c_oAscError.ID.PasswordIsNotCorrect,
							c_oAscError.Level.NoCritical);
						t.handlers.trigger("asc_onChangeProtectWorkbook");
						t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction);
					}
				}
				props.temporaryPassword = null;
			}
		};

		//only lockStructure
		this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction);
		if (props && props.temporaryPassword != null) {
			if (props.temporaryPassword === "") {
				checkPassword([""]);
			} else if (props.isPasswordXL()) {
				checkPassword([AscCommonExcel.getPasswordHash(props.temporaryPassword, true)]);
			} else {
				var checkHash = {password: props.temporaryPassword, salt: props.workbookSaltValue, spinCount: props.workbookSpinCount,
					alg: AscCommon.fromModelAlgorithmName(props.algorithmName)};
				AscCommon.calculateProtectHash([checkHash], checkPassword);
			}
		} else {
			checkPassword(null, true);
		}

		return true;
	};

  spreadsheet_api.prototype.asc_setSkin = function (theme) {
    AscCommon.updateGlobalSkin(theme);

    var elem = document.getElementById("ws-v-scrollbar");
    if (elem) {
      elem.style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
    }
    elem = document.getElementById("ws-h-scrollbar");
    if (elem) {
      elem.style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
    }
    elem = document.getElementById("ws-scrollbar-corner");
    if (elem) {
      elem.style.backgroundColor = AscCommon.GlobalSkin.ScrollBackgroundColor;
    }

    if (this.wb) {
      this.wb.updateSkin();
      var ws = this.wb.getWorksheet();
      if (ws) {
          this.controller.updateScrollSettings();
		  ws.draw();
      }
    }
  };

  spreadsheet_api.prototype.turnOffSpecialModes = function() {
      let bResult = false;
      if (this.isStartAddShape) {
          this.asc_endAddShape();
          bResult = true;
      }
      if(this.isEyedropperStarted())
      {
          this.cancelEyedropper();
          bResult = true;
      }
      if (this.isFormatPainterOn()) {
          this.formatPainter.putState(AscCommon.c_oAscFormatPainterState.kOff);
          if (this.wb) {
              this.wb.formatPainter(AscCommon.c_oAscFormatPainterState.kOff, undefined);
          }
          bResult = true;
      }
      if(this.isInkDrawerOn()) {
          this.stopInkDrawer();
          bResult = true;
      }
      return bResult;
  };
  spreadsheet_api.prototype.onUpdateRestrictions = function () {
    this._onUpdateDocumentCanUndoRedo();
    this.turnOffSpecialModes();
  };
  spreadsheet_api.prototype.isShowShapeAdjustments = function()
  {
    return this.canEdit();
  };
  spreadsheet_api.prototype.isShowTableAdjustments = function()
  {
      return this.canEdit();
  };
  spreadsheet_api.prototype.isShowEquationTrack = function()
  {
      return this.canEdit();
  };

  spreadsheet_api.prototype.asc_getEscapeSheetName = function(sheet)
  {
      return AscCommon.parserHelp.getEscapeSheetName(sheet)
  };

  spreadsheet_api.prototype.asc_undoAllChanges = function() {
  	if (this.wb.getCellEditMode()) {
		this.asc_closeCellEditor();
	}
  	this.wb.undo({All : true});
  };
	
  spreadsheet_api.prototype.asc_restartCheckSpelling = function()
  {
  	if (this.wb /*&& !this.spellcheckState.lockSpell*/) {
		this._spellCheckRestart();
  	}
  };
  spreadsheet_api.prototype.asc_ConvertEquationToMath = function(oEquation, isAll)
  {
      // TODO: Вообще здесь нужно запрашивать шрифты, которые использовались в старой формуле,
      //      но пока это только 1 шрифт "Cambria Math".
      var oWorkbook = this.wb;
      var loader   = AscCommon.g_font_loader;
      var fontinfo = AscFonts.g_fontApplication.GetFontInfo("Cambria Math");
      var isasync  = loader.LoadFont(fontinfo, function()
      {
          oWorkbook.convertEquationToMath(oEquation, isAll);
      }, this);

      if (false === isasync)
      {
          oWorkbook.convertEquationToMath(oEquation, isAll);
      }
  };



	/*отправляем инфомарцию, инфомарция в виде строки(id + ";" + isEdit + ";" + rangeStr;)
		_autoSaveInner -> wb.sendCursor -> CDocsCoApi.prototype.sendCursor
		NeedUpdateTargetForCollaboration  - флаг высталяем в true, когда поменялся селект, потом предыдущая функция отсылает инфу на сервер

	храним инф. о курсорах в
		CCollaborativeEditing->m_aForeignCursorsData, добавляем/удаляем с помощью методов Add_ForeignCursor/Remove_ForeignCursor

	удаляем инф. на
		t.handlers.trigger("asc_onConnectionStateChanged", e) -> Remove_ForeignCursor

	принимаем инфомарцию о курсорах
		this.CoAuthoringApi.onCursor -> WorkbookView.prototype.Update_ForeignCursor

	эвенты в интерфейс: asc_onShowForeignSelectLabel/asc_onHideForeignSelectLabel*/

	
	spreadsheet_api.prototype.showForeignSelectLabel = function (UserId, X, Y, Color, isEdit) {
		this.sendEvent("asc_onShowForeignCursorLabel", UserId, X, Y, new AscCommon.CColor(Color.r, Color.g, Color.b, 255), isEdit);
	};
	spreadsheet_api.prototype.hideForeignSelectLabel = function (UserId) {
		this.sendEvent("asc_onHideForeignCursorLabel", UserId);
	};

	//TODO временно положил в прототип. перенести!
	spreadsheet_api.prototype.sheetViewManagerLocks = [];
	spreadsheet_api.prototype.asc_addNamedSheetView = function (duplicateNamedSheetView, setActive) {
		var t = this;
		var ws = this.wb && this.wb.getWorksheet();
		var wsModel = ws ? ws.model : null;
		if (!wsModel) {
			return;
		}

		if (this.isNamedSheetViewManagerLocked(wsModel.Id)) {
			t.handlers.trigger("asc_onError", c_oAscError.ID.LockedEditView, c_oAscError.Level.NoCritical);
			return;
		}

		var namedSheetView;
		if (duplicateNamedSheetView) {
			namedSheetView = duplicateNamedSheetView.clone();
		} else {
			//если создаём новый вью когда находимся на другом вью, клонируем аквтиный
			var activeNamedSheetViewId = wsModel.getActiveNamedSheetViewId();
			if (activeNamedSheetViewId !== null) {
				duplicateNamedSheetView = true;
				namedSheetView = wsModel.getNamedSheetViewById(activeNamedSheetViewId).clone();
				namedSheetView.name = null;
			} else {
				namedSheetView = new Asc.CT_NamedSheetView();
			}
		}
		namedSheetView.ws = wsModel;
		namedSheetView.name = namedSheetView.generateName();

		this._isLockedNamedSheetView([namedSheetView], function(success) {
			if (!success) {
				t.handlers.trigger("asc_onError", c_oAscError.ID.LockedEditView, c_oAscError.Level.NoCritical);
				return;
			}

			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();
			wsModel.addNamedSheetView(namedSheetView, !!duplicateNamedSheetView);

			if (setActive) {
				t.asc_setActiveNamedSheetView(namedSheetView.name);
			}

			AscCommon.History.EndTransaction();

			if (!setActive) {
				t.handlers.trigger("asc_onRefreshNamedSheetViewList", wsModel.index);
			}
		});
	};

	spreadsheet_api.prototype.asc_getNamedSheetViews = function () {
		var ws = this.wb && this.wb.getWorksheet();
		var wsModel = ws ? ws.model : null;
		if (!wsModel) {
			return null;
		}

		return wsModel.getNamedSheetViews();
	};

	spreadsheet_api.prototype.asc_getActiveNamedSheetView = function (index) {
		var ws = this.wbModel.getWorksheet(index);
		if (!ws) {
			return null;
		}

		var activeNamedSheetViewId = ws.getActiveNamedSheetViewId();
		if (activeNamedSheetViewId !== null) {
			var activeNamedSheetView = ws.getNamedSheetViewById(activeNamedSheetViewId);
			if (activeNamedSheetView) {
				return activeNamedSheetView.name;
			}
		}

		return null;
	};

	spreadsheet_api.prototype.asc_deleteNamedSheetViews = function (namedSheetViews) {
		var t = this;
		var ws = this.wb && this.wb.getWorksheet();
		var wsModel = ws ? ws.model : null;
		if (!wsModel) {
			return;
		}

		this._isLockedNamedSheetView(namedSheetViews, function(success) {
			if (!success) {
				t.handlers.trigger("asc_onError", c_oAscError.ID.LockedEditView, c_oAscError.Level.NoCritical);
				return;
			}

			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();
			wsModel.deleteNamedSheetViews(namedSheetViews);
			AscCommon.History.EndTransaction();

			t.handlers.trigger("asc_onRefreshNamedSheetViewList", wsModel.index);
		});
	};

	spreadsheet_api.prototype._isLockedNamedSheetView = function (namedSheetViews, callback) {
		if (!namedSheetViews || !namedSheetViews.length) {
			callback(false);
		}
		var lockInfoArr =  [];
		for (var i = 0; i < namedSheetViews.length; i++) {
			var namedSheetView = namedSheetViews[i];
			var lockInfo = this.collaborativeEditing.getLockInfo(c_oAscLockTypeElem.Object, null,
				this.asc_getActiveWorksheetId(), namedSheetView.Get_Id());
			lockInfoArr.push(lockInfo);
		}
		this.collaborativeEditing.lock(lockInfoArr, callback);
	};

	spreadsheet_api.prototype._onUpdateNamedSheetViewLock = function(lockElem) {
		var t = this;

		if (c_oAscLockTypeElem.Object === lockElem.Element["type"]) {
			var wsModel = t.wbModel.getWorksheetById(lockElem.Element["sheetId"]);
			if (wsModel) {
				var wsIndex = wsModel.getIndex();
				var sheetView = wsModel.getNamedSheetViewById(lockElem.Element["rangeOrObjectId"]);
				if (sheetView) {
					sheetView.isLock = lockElem.UserId;
					this.handlers.trigger("asc_onRefreshNamedSheetViewList", wsIndex);
				}

				this.sheetViewManagerLocks[wsModel.Id] = true;
			}
		}
	};

	spreadsheet_api.prototype._onUpdateAllSheetViewLock = function () {
		var t = this;
		if (t.wbModel) {
			var i, length, wsModel, wsIndex;
			for (i = 0, length = t.wbModel.getWorksheetCount(); i < length; ++i) {
				wsModel = t.wbModel.getWorksheet(i);
				wsIndex = wsModel.getIndex();

				if (wsModel.aNamedSheetViews) {
					for (var j = 0; j < wsModel.aNamedSheetViews.length; j++) {
						var sheetView = wsModel.aNamedSheetViews[j];
						sheetView.isLock = null;
					}
				}
				this.handlers.trigger("asc_onRefreshNamedSheetViewList", wsIndex);
				this.sheetViewManagerLocks[wsModel.Id] = false;
			}
		}
	};

	spreadsheet_api.prototype.isNamedSheetViewManagerLocked = function (id) {
		return this.sheetViewManagerLocks[id];
	};

	spreadsheet_api.prototype.asc_setActiveNamedSheetView = function(name, index) {
		if (index === undefined) {
			index = this.wbModel.getActive();
		}
		var ws = this.wbModel.getWorksheet(index);

		//при переходе между вью - hidden manager не обновляется.
		var changedHiddenRowsArr = [];
		var historyUpdateRange = new asc.Range(0, 0, 0, 0);
		var i;

		ws.autoFilters.forEachTables(function (table) {
			historyUpdateRange.union2(table.Ref);
			for (var i = table.Ref.r1; i < table.Ref.r2; i++) {
				ws._getRowNoEmpty(i, function(row){
					if (row) {
						changedHiddenRowsArr[row.index] = row.getHidden();
					}
				});
			}
		});
		if (ws.AutoFilter && ws.AutoFilter.Ref) {
			for (i = ws.AutoFilter.Ref.r1; i < ws.AutoFilter.Ref.r2; i++) {
				ws._getRowNoEmpty(i, function(row){
					if (row) {
						changedHiddenRowsArr[row.index] = row.getHidden();
					}
				});
			}
		}

		var oldActiveId = ws.getActiveNamedSheetViewId();
		ws.setActiveNamedSheetView(null);
		for (i = 0; i < ws.aNamedSheetViews.length; i++) {
			if (name === ws.aNamedSheetViews[i].name) {
				ws.setActiveNamedSheetView(ws.aNamedSheetViews[i].Id);
				ws.aNamedSheetViews[i]._isActive = true;
			} else {
				ws.aNamedSheetViews[i]._isActive = false;
			}
		}
		if (oldActiveId !== ws.getActiveNamedSheetViewId()) {
			AscCommon.History.Create_NewPoint();
			AscCommon.History.StartTransaction();

			if (ws.AutoFilter && ws.AutoFilter.Ref) {
				historyUpdateRange.union2(ws.AutoFilter.Ref);
			}

			AscCommon.History.Add(AscCommonExcel.g_oUndoRedoWorksheet, AscCH.historyitem_Worksheet_SetActiveNamedSheetView,
				ws ? ws.getId() : null, historyUpdateRange,
				new AscCommonExcel.UndoRedoData_FromTo(oldActiveId, ws.getActiveNamedSheetViewId()), true);

			AscCommon.History.EndTransaction();

			//TODO нужно переприменять в дальнейшем сортировку

			//если переходим на вью, то необходимо открыть все строки и применить фильтры
			//если переходим на дефолт, то необходимо скрыть ещё те строки, которые в модели лежат
			//посколько при переходе во вью данные из модели удалились - их нужно получить
			//т.е. нужно где-то хранить!

			//при переходе во вью - переносим с дефолта все флаги о скрытии строчек
			//переприменяем все фильтры
			//применяем скрытие строчек внутрии а/ф - используя новый флаг о скрытии
			//все остальные строчки - используя старый флаг о скрытии строк
			//получение данных о скрытой строке: в режиме вью внутри а/ф используем новый флаг
			//вне а/ф - старый флаг
			//при переходе из дефолта внутри а/ф(к которому не применен фильтр) наследуем флаг об скрытии/открытии ячеек
			//для этого прохожусь по всем строкам - и наследую флаг

			if (ws.getActiveNamedSheetViewId() !== null) {
				//чтобы не усложнять логику решил не наследовать внутри а/ф скрытые строки от дефолта
				//просто отрываем все строки, а далее применяем те, что скрыты во вью
				ws.getRange3(0, 0, AscCommon.gc_nMaxRow0, 0)._foreachRowNoEmpty(function(row) {
					if (ws.autoFilters.containInFilter(row.index/*, true*/)) {
						row.setHidden(false, true);
					} /*else {
						//наследуем с дефолта, если в этих строчках нет применнного фильтра
						row.setHidden(row.getHidden(false), true);
					}*/
				});
			}

			var _changeHiddenManager = function (_row) {
				if (_row && _row.index >= 0 && (!_row.getHidden() !== !changedHiddenRowsArr[_row.index])) {
					ws.hiddenManager.addHidden(true, _row.index);
				}
			};

			ws.autoFilters.forEachTables(function (table) {
				for (var i = table.Ref.r1; i < table.Ref.r2; i++) {
					ws._getRowNoEmpty(i, function(row){
						_changeHiddenManager(row);
					});
				}
			});
			if (ws.AutoFilter && ws.AutoFilter.Ref) {
				for (i = ws.AutoFilter.Ref.r1; i < ws.AutoFilter.Ref.r2; i++) {
					ws._getRowNoEmpty(i, function(row){
						_changeHiddenManager(row);
					});
				}
			}

			var oRange = new AscCommonExcel.Range(ws, historyUpdateRange.r1, historyUpdateRange.c1, historyUpdateRange.r2, historyUpdateRange.c2);
			this.wb.handleChartsOnWorkbookChange([oRange]);
			ws.autoFilters.reapplyAllFilters(true, ws.getActiveNamedSheetViewId() !== null, null, true);
			this.updateAllFilters();
			this.handlers.trigger("asc_onRefreshNamedSheetViewList", index);
		}
	};

	spreadsheet_api.prototype.updateAllFilters = function() {
		var t = this;
		var wsModel = this.wbModel.getWorksheet(this.wbModel.getActive());
		var ws = t.wb.getWorksheet(wsModel.getIndex());

		var arrChangedRanges = [];
		for (var i = 0; i < wsModel.TableParts.length; ++i) {
			var table = wsModel.TableParts[i];
			arrChangedRanges.push(table.Ref);
		}

		if (wsModel.AutoFilter) {
			arrChangedRanges.push(wsModel.AutoFilter.Ref);
		}

		ws._updateGroups();
		//wsModel.autoFilters.reDrawFilter(arn);
		var oRecalcType = AscCommonExcel.recalcType.full;
		//reinitRanges = true;
		//updateDrawingObjectsInfo = {target: c_oTargetType.RowResize, row: arn.r1};

		ws._initCellsArea(oRecalcType);
		if (oRecalcType) {
			ws.cache.reset();
		}
		ws._cleanCellsTextMetricsCache();
		ws.objectRender.bUpdateMetrics = false;
		ws._prepareCellTextMetricsCache();
		ws.objectRender.bUpdateMetrics = true;

		//arrChangedRanges = arrChangedRanges.concat(t.model.hiddenManager.getRecalcHidden());

		ws.cellCommentator.updateAreaComments();
		ws.draw();

		ws._updateVisibleRowsCount();

		ws.handlers.trigger("selectionChanged");
		ws.getSelectionMathInfo(function (info) {
			ws.handlers.trigger("selectionMathInfoChanged", info);
		});
	};


	spreadsheet_api.prototype.asc_EditSelectAll = function() {
		if (this.wb) {
			this.wb.selectAll();
		}
	};

	spreadsheet_api.prototype.asc_addCellWatches = function (sRange) {
		var t = this;
		if (this.wb && this.wb.model) {
			var oRange = t.wb.model.getRangeAndSheetFromStr(sRange);
			if (oRange && oRange.sheet && oRange.range) {
				var maxCellsCount = 100;
				var countCells = oRange.range.getWidth() * oRange.range.getHeight();
				if (countCells > maxCellsCount) {
					this.handlers.trigger("asc_onConfirmAction", Asc.c_oAscConfirm.ConfirmAddCellWatches, function (can) {
						if (can) {
							t.wb.model.addCellWatches(oRange.sheet, oRange.range);
						}
					}, countCells);
				} else {
					t.wb.model.addCellWatches(oRange.sheet, oRange.range);
				}
			}
		}
	};

	spreadsheet_api.prototype.asc_deleteCellWatches = function (aCellWatches, opt_remove_all) {
		if (this.wb && this.wb.model) {
			this.wb.model.delCellWatches(aCellWatches, true, opt_remove_all);
		}
	};

	spreadsheet_api.prototype.asc_getCellWatches = function () {
		var res = null;
		if (this.wb && this.wb.model) {
			this.wb.model.recalculateCellWatches(true);

			for (var i = 0; i < this.wb.model.aWorksheets.length; i++) {
				var ws = this.wb.model.aWorksheets[i];
				if (ws && ws.aCellWatches.length) {
					if (!res) {
						res = [];
					}
					res = res.concat(ws.aCellWatches);
				}
			}
		}
		return res;
	};

	spreadsheet_api.prototype.setDrawGroupsRestriction = function() {
		if (this.wb) {
			//пока использую строку. будут другие ограничения на отрисовку - необходимо завести константы
			this.wb.setDrawRestriction("groups");
		}
	};

	spreadsheet_api.prototype.onWorksheetChange = function(props) {
		let ws = this.wbModel.getActiveWs();
		let range = null;
		let result = null;
		if (Array.isArray(props)) {
			// todo сделать получение листа ещё
			range = ws.getRange3(props[0].r1, props[0].c1, props[0].r2, props[0].c2);
			let arr = props.length <= 1 ? null : props.map(function(r){
				return ws.getRange3(r.r1, r.c1, r.r2, r.c2);
			})
			result = this.private_GetRange(range, arr);
		} else {
			// todo сделать получение листа ещё
			range = ws.getRange3(props.r1, props.c1, props.r2, props.c2);
			result = this.private_GetRange(range);
		}
		this.sendEvent('onWorksheetChange', result);
	};

	spreadsheet_api.prototype.asc_enterText = function(codePoints)
	{
		let wb = this.wb;
		if (!wb)
			return false;

		return wb.EnterText(codePoints);
	};

	spreadsheet_api.prototype.asc_correctEnterText = function(oldValue, newValue)
	{
		let wb = this.wb;
		if (!wb)
			return false;

		return wb.CorrectEnterText(oldValue, newValue);
	};

	spreadsheet_api.prototype.asc_getExternalReferences = function() {
		return this.wb.getExternalReferences();
	};

	spreadsheet_api.prototype.asc_updateExternalReferences = function(arr) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		this.wb.updateExternalReferences(arr);
	};

	spreadsheet_api.prototype.asc_openExternalReference = function(externalReference) {
		let isLocalDesktop = window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]();
		if (isLocalDesktop) {
			alert("NEED SUPPORT LOCAL OPEN FILE");
			return null;
		} else {
			return externalReference;
		}
	};

	spreadsheet_api.prototype.asc_removeExternalReferences = function(arr) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		this.wb.removeExternalReferences(arr);
	};

	spreadsheet_api.prototype.asc_changeExternalReference = function(eR, to) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		this.wb.changeExternalReference(eR, to);
	};

	spreadsheet_api.prototype.asc_fillHandleDone = function(range) {
		if (this.canEdit()) {
			let wb = this.wb;
			if (!wb) {
				return;
			}
			wb.fillHandleDone(range);
		}
	};

	spreadsheet_api.prototype.asc_canFillHandle = function(range) {
		if (this.canEdit()) {
			let wb = this.wb;
			if (!wb) {
				return;
			}
			return wb.canFillHandle(range);
		}
		return false;
	};

	spreadsheet_api.prototype.getEyedropperImgData = function() {
		const oWSViewerCanvas = document.getElementById("ws-canvas");
		const oWSOverlayCanvas = document.getElementById("ws-canvas-overlay");
		const oGrViewerCanvas = document.getElementById("ws-canvas-graphic");
		const oGrOverlayCanvas = document.getElementById("ws-canvas-graphic-overlay");
		if(!oWSViewerCanvas || !oWSOverlayCanvas
			|| !oGrViewerCanvas || !oGrOverlayCanvas) {
			return null;
		}
		let oCanvas = document.createElement("canvas");
		oCanvas.width = oWSViewerCanvas.width;
		oCanvas.height = oWSViewerCanvas.height;
		const oCtx = oCanvas.getContext("2d");
		oCtx.drawImage(oWSViewerCanvas, 0, 0);
		oCtx.drawImage(oWSOverlayCanvas, 0, 0);
		oCtx.drawImage(oGrViewerCanvas, 0, 0);
		oCtx.drawImage(oGrOverlayCanvas, 0, 0);
		return oCtx.getImageData(0, 0, oCanvas.width, oCanvas.height);
	};
	spreadsheet_api.prototype.asc_addUserProtectedRange = function (obj) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}

		if (this.wb && this.wb.getCellEditMode()) {
			return;
		}

		this.wb.changeUserProtectedRanges(null, obj);
	};

	spreadsheet_api.prototype.asc_changeUserProtectedRange = function (oldObj, newObj) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}

		if (this.wb && this.wb.getCellEditMode()) {
			return;
		}

		this.wb.changeUserProtectedRanges(oldObj, newObj);
	};

	spreadsheet_api.prototype.asc_deleteUserProtectedRange = function (arr) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return false;
		}

		if (this.wb && this.wb.getCellEditMode()) {
			return;
		}

		this.wb.deleteUserProtectedRanges(arr);
	};

	spreadsheet_api.prototype.asc_getUserProtectedRanges = function(sheetName) {
		if (!this.wb) {
			return null;
		}

		let res = [];
		if (sheetName) {
			let sheet = this.wb.model.getWorksheetByName(sheetName);
			for (let i = 0; i < sheet.userProtectedRanges.length; i++) {
				res.push(sheet.userProtectedRanges[i]);
			}
		} else {
			this.wb.model.forEach(function (ws) {
				for (let i = 0; i < ws.userProtectedRanges.length; i++) {
					res.push(ws.userProtectedRanges[i]);
				}
			});
		}

		res.sort(function (a, b) {
			return a.name > b.name ? 1 : -1;
		});

		return res;
	};

	spreadsheet_api.prototype._onUpdateUserProtectedRange = function (lockElem) {
		let _element = lockElem.Element;
		let isNeedObject = _element["sheetId"] !== -1 && _element["rangeOrObjectId"] !== -1;
		let isNeedType = isNeedObject && AscCommonExcel.c_oAscLockTypeElemSubType.UserProtectedRange === _element.subType &&
			c_oAscLockTypeElem.Object === _element["type"];

		if (isNeedType) {
			let sheet = this.wbModel.getWorksheetById(_element["sheetId"]);
			if (sheet) {
				var userRange = sheet.getUserProtectedRangeById(_element["rangeOrObjectId"]);
				if (userRange) {
					if (!this.collaborativeEditing.getFast()) {
						userRange.obj.isLock = lockElem.UserId;
					}
					this.handlers.trigger("asc_onRefreshUserProtectedRangesList");
				}
			}
			this.handlers.trigger("asc_onLockUserProtectedManager", true);
		}
	};

	spreadsheet_api.prototype._onUnlockUserProtectedRanges = function() {
		this.wb.unlockUserProtectedRanges();
	};

	spreadsheet_api.prototype.asc_checkUserProtectedRangeName = function(checkName) {
		if (!this.wbModel) {
			return;
		}
		return this.wbModel.checkUserProtectedRangeName(checkName);
	};

	spreadsheet_api.prototype.asc_SetSheetViewType = function(val) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		let wb = this.wb;
		if (!wb) {
			return;
		}
		var ws = this.wb.getWorksheet();
		//val -> window['Asc']['c_oAscESheetViewType']
		return ws.setSheetViewType(val);
	};

	spreadsheet_api.prototype.asc_GetSheetViewType = function(index) {
		let sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
		let ws = this.wbModel.getWorksheet(sheetIndex);
		let res = null;
		if (ws && ws.sheetViews) {
			var sheetView = ws.sheetViews[0];
			if (sheetView) {
				res = sheetView.view;
			}
		}
		return res == null ? AscCommonExcel.ESheetViewType.normal : res;
	};

	spreadsheet_api.prototype.asc_InsertPageBreak = function() {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		let wb = this.wb;
		if (!wb) {
			return;
		}
		var ws = this.wb.getWorksheet();
		return ws && ws.insertPageBreak();
	};

	spreadsheet_api.prototype.asc_RemovePageBreak = function() {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		let wb = this.wb;
		if (!wb) {
			return;
		}
		var ws = this.wb.getWorksheet();
		return ws.removePageBreak();
	};

	spreadsheet_api.prototype.asc_ResetAllPageBreaks = function() {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		let wb = this.wb;
		if (!wb) {
			return;
		}
		var ws = this.wb.getWorksheet();
		return ws.resetAllPageBreaks();
	};

	spreadsheet_api.prototype.asc_GetPageBreaksDisableType = function (index) {
		if (!this.wbModel) {
			return;
		}
		let sheetIndex = (undefined !== index && null !== index) ? index : this.wbModel.getActive();
		let ws = this.wbModel.getWorksheet(sheetIndex);
		let res = Asc.c_oAscPageBreaksDisableType.none;
		if (ws) {
			let isPageBreaks = ws && ((ws.colBreaks && ws.colBreaks.getCount()) || (ws.rowBreaks && ws.rowBreaks.getCount()));
			let activeCell = ws.selectionRange && ws.selectionRange.activeCell;
			let isFirstActiveCell = activeCell && activeCell.col === 0 && activeCell.row === 0;

			if (isFirstActiveCell && !isPageBreaks) {
				//disable all
				res = Asc.c_oAscPageBreaksDisableType.all;
			} else if (isFirstActiveCell && isPageBreaks) {
				//disable insert/remove
				res = Asc.c_oAscPageBreaksDisableType.insertRemove;
			} else if (!isFirstActiveCell && !isPageBreaks) {
				//disable reset
				res = Asc.c_oAscPageBreaksDisableType.reset;
			}
		}

		return res;
	};


	spreadsheet_api.prototype.asc_TracePrecedents = function() {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		let wb = this.wb;
		if (!wb) {
			return;
		}
		let ws = wb.getWorksheet();
		return ws.tracePrecedents();
	};

	spreadsheet_api.prototype.asc_TraceDependents = function() {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		let wb = this.wb;
		if (!wb) {
			return;
		}
		let ws = wb.getWorksheet();
		return ws.traceDependents();
	};

	spreadsheet_api.prototype.asc_RemoveTraceArrows = function(type) {
		if (this.collaborativeEditing.getGlobalLock() || !this.canEdit()) {
			return;
		}
		let wb = this.wb;
		if (!wb) {
			return;
		}
		let ws = wb.getWorksheet();
		return ws.removeTraceArrows(type);
	};



  /*
   * Export
   * -----------------------------------------------------------------------------
   */

  asc["spreadsheet_api"] = spreadsheet_api;
  prot = spreadsheet_api.prototype;

  prot["asc_GetFontThumbnailsPath"] = prot.asc_GetFontThumbnailsPath;
  prot["asc_setDocInfo"] = prot.asc_setDocInfo;
  prot["asc_changeDocInfo"] = prot.asc_changeDocInfo;
  prot['asc_getFunctionArgumentSeparator'] = prot.asc_getFunctionArgumentSeparator;
  prot['asc_getCurrencySymbols'] = prot.asc_getCurrencySymbols;
  prot['asc_getAdditionalCurrencySymbols'] = prot.asc_getAdditionalCurrencySymbols;
  prot['asc_getLocaleExample'] = prot.asc_getLocaleExample;
  prot['asc_convertNumFormatLocal2NumFormat'] = prot.asc_convertNumFormatLocal2NumFormat;
  prot['asc_convertNumFormat2NumFormatLocal'] = prot.asc_convertNumFormat2NumFormatLocal;
  prot['asc_getFormatCells'] = prot.asc_getFormatCells;
  prot["asc_getLocaleCurrency"] = prot.asc_getLocaleCurrency;
  prot["asc_setLocale"] = prot.asc_setLocale;
  prot["asc_getLocale"] = prot.asc_getLocale;
  prot["asc_getDecimalSeparator"] = prot.asc_getDecimalSeparator;
  prot["asc_getGroupSeparator"] = prot.asc_getGroupSeparator;
  prot["asc_getFrozenPaneBorderType"] = prot.asc_getFrozenPaneBorderType;
  prot["asc_setFrozenPaneBorderType"] = prot.asc_setFrozenPaneBorderType;
  prot["asc_getEditorPermissions"] = prot.asc_getEditorPermissions;
  prot["asc_LoadDocument"] = prot.asc_LoadDocument;
  prot["asc_DownloadAs"] = prot.asc_DownloadAs;
  prot["asc_Save"] = prot.asc_Save;
  prot["forceSave"] = prot.forceSave;
  prot["asc_setIsForceSaveOnUserSave"] = prot.asc_setIsForceSaveOnUserSave;
  prot["asc_Resize"] = prot.asc_Resize;
  prot["asc_Copy"] = prot.asc_Copy;
  prot["asc_Paste"] = prot.asc_Paste;
  prot["asc_SpecialPaste"] = prot.asc_SpecialPaste;
  prot["asc_Cut"] = prot.asc_Cut;
  prot["asc_Undo"] = prot.asc_Undo;
  prot["asc_Redo"] = prot.asc_Redo;
  prot["asc_TextImport"] = prot.asc_TextImport;
  prot["asc_TextToColumns"] = prot.asc_TextToColumns;
  prot["asc_TextFromFileOrUrl"] = prot.asc_TextFromFileOrUrl;

  prot["asc_initPrintPreview"] = prot.asc_initPrintPreview;
  prot["asc_updatePrintPreview"] = prot.asc_updatePrintPreview;
  prot["asc_drawPrintPreview"] = prot.asc_drawPrintPreview;
  prot["asc_closePrintPreview"] = prot.asc_closePrintPreview;


  prot["asc_getDocumentName"] = prot.asc_getDocumentName;
  prot["asc_getAppProps"] = prot.asc_getAppProps;
  prot["asc_getCoreProps"] = prot.asc_getCoreProps;
  prot["asc_setCoreProps"] = prot.asc_setCoreProps;
  prot["asc_isDocumentModified"] = prot.asc_isDocumentModified;
  prot["isDocumentModified"] = prot.isDocumentModified;
  prot["asc_isDocumentCanSave"] = prot.asc_isDocumentCanSave;
  prot["asc_getCanUndo"] = prot.asc_getCanUndo;
  prot["asc_getCanRedo"] = prot.asc_getCanRedo;
  prot["can_CopyCut"] = prot.can_CopyCut;

  prot["asc_setAutoSaveGap"] = prot.asc_setAutoSaveGap;

  prot["asc_setViewMode"] = prot.asc_setViewMode;
  prot["asc_setFilteringMode"] = prot.asc_setFilteringMode;
  prot["asc_setAdvancedOptions"] = prot.asc_setAdvancedOptions;
  prot["asc_setPageOptions"] = prot.asc_setPageOptions;
  prot["asc_savePagePrintOptions"] = prot.asc_savePagePrintOptions;
  prot["asc_getPageOptions"] = prot.asc_getPageOptions;
  prot["asc_changeDocSize"] = prot.asc_changeDocSize;
  prot["asc_changePageMargins"] = prot.asc_changePageMargins;
  prot["asc_setPageOption"] = prot.asc_setPageOption;
  prot["asc_changePageOrient"] = prot.asc_changePageOrient;
  prot["asc_changePrintTitles"] = prot.asc_changePrintTitles;
  prot["asc_getPrintTitlesRange"] = prot.asc_getPrintTitlesRange;
  prot["asc_SetPrintHeadings"] = prot.asc_SetPrintHeadings;
  prot["asc_SetPrintGridlines"] = prot.asc_SetPrintGridlines;


  prot["asc_ChangePrintArea"] = prot.asc_ChangePrintArea;
  prot["asc_CanAddPrintArea"] = prot.asc_CanAddPrintArea;
  prot["asc_SetPrintScale"] = prot.asc_SetPrintScale;


  prot["asc_decodeBuffer"] = prot.asc_decodeBuffer;

  prot["asc_registerCallback"] = prot.asc_registerCallback;
  prot["asc_unregisterCallback"] = prot.asc_unregisterCallback;

  prot["asc_changeArtImageFromFile"] = prot.asc_changeArtImageFromFile;

  prot["asc_SetDocumentPlaceChangedEnabled"] = prot.asc_SetDocumentPlaceChangedEnabled;
  prot["asc_SetFastCollaborative"] = prot.asc_SetFastCollaborative;
	prot["asc_setThumbnailStylesSizes"] = prot.asc_setThumbnailStylesSizes;

  // Workbook interface

  prot["asc_getWorksheetsCount"] = prot.asc_getWorksheetsCount;
  prot["asc_getWorksheetName"] = prot.asc_getWorksheetName;
  prot["asc_getWorksheetTabColor"] = prot.asc_getWorksheetTabColor;
  prot["asc_setWorksheetTabColor"] = prot.asc_setWorksheetTabColor;
  prot["asc_getActiveWorksheetIndex"] = prot.asc_getActiveWorksheetIndex;
  prot["asc_getActiveWorksheetId"] = prot.asc_getActiveWorksheetId;
  prot["asc_getWorksheetId"] = prot.asc_getWorksheetId;
  prot["asc_isWorksheetHidden"] = prot.asc_isWorksheetHidden;
  prot["asc_isWorksheetLockedOrDeleted"] = prot.asc_isWorksheetLockedOrDeleted;
  prot["asc_isWorkbookLocked"] = prot.asc_isWorkbookLocked;
  prot["asc_isLayoutLocked"] = prot.asc_isLayoutLocked;
  prot["asc_isPrintAreaLocked"] = prot.asc_isPrintAreaLocked;
  prot["asc_isHeaderFooterLocked"] = prot.asc_isHeaderFooterLocked;
  prot["asc_getHiddenWorksheets"] = prot.asc_getHiddenWorksheets;
  prot["asc_showWorksheet"] = prot.asc_showWorksheet;
  prot["asc_hideWorksheet"] = prot.asc_hideWorksheet;
  prot["asc_renameWorksheet"] = prot.asc_renameWorksheet;
  prot["asc_addWorksheet"] = prot.asc_addWorksheet;
  prot["asc_insertWorksheet"] = prot.asc_insertWorksheet;
  prot["asc_deleteWorksheet"] = prot.asc_deleteWorksheet;
  prot["asc_moveWorksheet"] = prot.asc_moveWorksheet;
  prot["asc_copyWorksheet"] = prot.asc_copyWorksheet;
  prot["asc_cleanSelection"] = prot.asc_cleanSelection;
  prot["asc_getZoom"] = prot.asc_getZoom;
  prot["asc_setZoom"] = prot.asc_setZoom;
  prot["asc_enableKeyEvents"] = prot.asc_enableKeyEvents;
  prot["asc_searchEnabled"] = prot.asc_searchEnabled;
  prot["asc_findText"] = prot.asc_findText;
  prot["asc_replaceText"] = prot.asc_replaceText;
  prot["asc_endFindText"] = prot.asc_endFindText;
  prot["asc_findCell"] = prot.asc_findCell;
  prot["asc_StartTextAroundSearch"] = prot.asc_StartTextAroundSearch;
  prot["asc_SelectSearchElement"] = prot.asc_SelectSearchElement;
  prot["asc_closeCellEditor"] = prot.asc_closeCellEditor;
  prot["asc_StartMoveSheet"] = prot.asc_StartMoveSheet;
  prot["asc_EndMoveSheet"] = prot.asc_EndMoveSheet;
  prot["asc_setWorksheetRange"] = prot.asc_setWorksheetRange;

  prot["asc_setR1C1Mode"] = prot.asc_setR1C1Mode;
  prot["asc_setIncludeNewRowColTable"] = prot.asc_setIncludeNewRowColTable;

  prot["asc_setShowZeroCellValues"] = prot.asc_setShowZeroCellValues;
  prot["asc_SetAutoCorrectHyperlinks"] = prot.asc_SetAutoCorrectHyperlinks;


  // Spreadsheet interface

  prot["asc_getColumnWidth"] = prot.asc_getColumnWidth;
  prot["asc_setColumnWidth"] = prot.asc_setColumnWidth;
  prot["asc_showColumns"] = prot.asc_showColumns;
  prot["asc_hideColumns"] = prot.asc_hideColumns;
  prot["asc_autoFitColumnWidth"] = prot.asc_autoFitColumnWidth;
  prot["asc_getRowHeight"] = prot.asc_getRowHeight;
  prot["asc_setRowHeight"] = prot.asc_setRowHeight;
  prot["asc_autoFitRowHeight"] = prot.asc_autoFitRowHeight;
  prot["asc_showRows"] = prot.asc_showRows;
  prot["asc_hideRows"] = prot.asc_hideRows;
  prot["asc_insertCells"] = prot.asc_insertCells;
  prot["asc_deleteCells"] = prot.asc_deleteCells;
  prot["asc_mergeCells"] = prot.asc_mergeCells;
  prot["asc_sortCells"] = prot.asc_sortCells;
  prot["asc_emptyCells"] = prot.asc_emptyCells;
  prot["asc_mergeCellsDataLost"] = prot.asc_mergeCellsDataLost;
  prot["asc_sortCellsRangeExpand"] = prot.asc_sortCellsRangeExpand;
  prot["asc_getSheetViewSettings"] = prot.asc_getSheetViewSettings;
  prot["asc_setDisplayGridlines"] = prot.asc_setDisplayGridlines;
  prot["asc_setDisplayHeadings"] = prot.asc_setDisplayHeadings;
  prot["asc_setShowZeros"] = prot.asc_setShowZeros;
  prot["asc_setShowFormulas"] = prot.asc_setShowFormulas;
  prot["asc_getShowFormulas"] = prot.asc_getShowFormulas;


  // Defined Names
  prot["asc_getDefinedNames"] = prot.asc_getDefinedNames;
  prot["asc_setDefinedNames"] = prot.asc_setDefinedNames;
  prot["asc_editDefinedNames"] = prot.asc_editDefinedNames;
  prot["asc_delDefinedNames"] = prot.asc_delDefinedNames;
  prot["asc_getDefaultDefinedName"] = prot.asc_getDefaultDefinedName;
  prot["asc_checkDefinedName"] = prot.asc_checkDefinedName;

  // Ole Editor
	prot["asc_toggleChangeVisibleAreaOleEditor"] = prot.asc_toggleChangeVisibleAreaOleEditor;
	prot["asc_toggleShowVisibleAreaOleEditor"] = prot.asc_toggleShowVisibleAreaOleEditor;
  prot["asc_addOleObjectAction"] = prot.asc_addOleObjectAction;
  prot["asc_addTableOleObjectInOleEditor"] = prot.asc_addTableOleObjectInOleEditor;
  prot["asc_getBinaryInfoOleObject"] = prot.asc_getBinaryInfoOleObject;
  prot["asc_editOleObjectAction"] = prot.asc_editOleObjectAction;
  prot["asc_doubleClickOnTableOleObject"] = prot.asc_doubleClickOnTableOleObject;
  prot["asc_startEditCurrentOleObject"] = prot.asc_startEditCurrentOleObject;

  // Auto filters interface + format as table
  prot["asc_addAutoFilter"] = prot.asc_addAutoFilter;
  prot["asc_changeAutoFilter"] = prot.asc_changeAutoFilter;
  prot["asc_applyAutoFilter"] = prot.asc_applyAutoFilter;
  prot["asc_applyAutoFilterByType"] = prot.asc_applyAutoFilterByType;
  prot["asc_reapplyAutoFilter"] = prot.asc_reapplyAutoFilter;
  prot["asc_sortColFilter"] = prot.asc_sortColFilter;
  prot["asc_getAddFormatTableOptions"] = prot.asc_getAddFormatTableOptions;
  prot["asc_clearFilter"] = prot.asc_clearFilter;
  prot["asc_clearFilterColumn"] = prot.asc_clearFilterColumn;
  prot["asc_changeSelectionFormatTable"] = prot.asc_changeSelectionFormatTable;
  prot["asc_changeFormatTableInfo"] = prot.asc_changeFormatTableInfo;
  prot["asc_insertCellsInTable"] = prot.asc_insertCellsInTable;
  prot["asc_deleteCellsInTable"] = prot.asc_deleteCellsInTable;
  prot["asc_changeDisplayNameTable"] = prot.asc_changeDisplayNameTable;
  prot["asc_changeTableRange"] = prot.asc_changeTableRange;
  prot["asc_convertTableToRange"] = prot.asc_convertTableToRange;
  prot["asc_getTablePictures"] = prot.asc_getTablePictures;
  prot["asc_getSlicerPictures"] = prot.asc_getSlicerPictures;
  prot["asc_getDefaultTableStyle"] = prot.asc_getDefaultTableStyle;


  prot["asc_applyAutoCorrectOptions"] = prot.asc_applyAutoCorrectOptions;

  //Group data
  prot["asc_group"] = prot.asc_group;
  prot["asc_ungroup"] = prot.asc_ungroup;
  prot["asc_canGroupPivot"] = prot.asc_canGroupPivot;
  prot["asc_groupPivot"] = prot.asc_groupPivot;
  prot["asc_ungroupPivot"] = prot.asc_ungroupPivot;
  prot["asc_clearOutline"] = prot.asc_clearOutline;
  prot["asc_changeGroupDetails"] = prot.asc_changeGroupDetails;
  prot["asc_checkAddGroup"] = prot.asc_checkAddGroup;
  prot["asc_setGroupSummary"] = prot.asc_setGroupSummary;
  prot["asc_getGroupSummaryRight"] = prot.asc_getGroupSummaryRight;
  prot["asc_getGroupSummaryBelow"] = prot.asc_getGroupSummaryBelow;


  // Drawing objects interface

  prot["asc_showDrawingObjects"] = prot.asc_showDrawingObjects;
  prot["asc_drawingObjectsExist"] = prot.asc_drawingObjectsExist;
  prot["asc_getChartObject"] = prot.asc_getChartObject;
  prot["asc_addChartDrawingObject"] = prot.asc_addChartDrawingObject;
  prot["asc_editChartDrawingObject"] = prot.asc_editChartDrawingObject;
  prot["asc_addImageDrawingObject"] = prot.asc_addImageDrawingObject;
  prot["asc_getCurrentDrawingMacrosName"] = prot.asc_getCurrentDrawingMacrosName;
  prot["asc_assignMacrosToCurrentDrawing"] = prot.asc_assignMacrosToCurrentDrawing;
  prot["asc_setSelectedDrawingObjectLayer"] = prot.asc_setSelectedDrawingObjectLayer;
  prot["asc_setSelectedDrawingObjectAlign"] = prot.asc_setSelectedDrawingObjectAlign;
  prot["asc_DistributeSelectedDrawingObjectHor"] = prot.asc_DistributeSelectedDrawingObjectHor;
  prot["asc_DistributeSelectedDrawingObjectVer"] = prot.asc_DistributeSelectedDrawingObjectVer;
  prot["asc_getSelectedDrawingObjectsCount"] = prot.asc_getSelectedDrawingObjectsCount;
  prot["SetDrawImagePreviewBulletForMenu"] = prot.SetDrawImagePreviewBulletForMenu;
  prot["asc_getChartPreviews"] = prot.asc_getChartPreviews;
  prot["asc_getTextArtPreviews"] = prot.asc_getTextArtPreviews;
  prot['asc_getPropertyEditorShapes'] = prot.asc_getPropertyEditorShapes;
  prot['asc_getPropertyEditorTextArts'] = prot.asc_getPropertyEditorTextArts;
  prot["asc_checkDataRange"] = prot.asc_checkDataRange;
  prot["asc_getBinaryFileWriter"] = prot.asc_getBinaryFileWriter;
  prot["asc_getWordChartObject"] = prot.asc_getWordChartObject;
  prot["asc_cleanWorksheet"] = prot.asc_cleanWorksheet;
  prot["asc_showImageFileDialog"] = prot.asc_showImageFileDialog;
  prot["asc_addImage"] = prot.asc_addImage;
  prot["asc_setData"] = prot.asc_setData;
  prot["asc_getData"] = prot.asc_getData;
  prot["asc_onCloseChartFrame"] = prot.asc_onCloseChartFrame;

  // Cell comment interface
  prot["asc_addComment"] = prot.asc_addComment;
  prot["asc_changeComment"] = prot.asc_changeComment;
  prot["asc_findComment"] = prot.asc_findComment;
  prot["asc_removeComment"] = prot.asc_removeComment;
  prot["asc_RemoveAllComments"] = prot.asc_RemoveAllComments;
  prot["asc_GetCommentLogicPosition"] = prot.asc_GetCommentLogicPosition;
  prot["asc_ResolveAllComments"] = prot.asc_ResolveAllComments;
  prot["asc_showComment"] = prot.asc_showComment;
  prot["asc_selectComment"] = prot.asc_selectComment;

  prot["asc_showComments"] = prot.asc_showComments;
  prot["asc_hideComments"] = prot.asc_hideComments;

  // Shapes
  prot["setStartPointHistory"] = prot.setStartPointHistory;
  prot["setEndPointHistory"] = prot.setEndPointHistory;
  prot["asc_startAddShape"] = prot.asc_startAddShape;
  prot["asc_endAddShape"] = prot.asc_endAddShape;
  prot["asc_addShapeOnSheet"] = prot.asc_addShapeOnSheet;
  prot["asc_canEditGeometry"] = prot.asc_canEditGeometry;
  prot["asc_editPointsGeometry"] = prot.asc_editPointsGeometry;
  prot["asc_isAddAutoshape"] = prot.asc_isAddAutoshape;
  prot["asc_canAddShapeHyperlink"] = prot.asc_canAddShapeHyperlink;
  prot["asc_canGroupGraphicsObjects"] = prot.asc_canGroupGraphicsObjects;
  prot["asc_groupGraphicsObjects"] = prot.asc_groupGraphicsObjects;
  prot["asc_canUnGroupGraphicsObjects"] = prot.asc_canUnGroupGraphicsObjects;
  prot["asc_unGroupGraphicsObjects"] = prot.asc_unGroupGraphicsObjects;
  prot["asc_getGraphicObjectProps"] = prot.asc_getGraphicObjectProps;
  prot["asc_GetSelectedText"] = prot.asc_GetSelectedText;
  prot["asc_setGraphicObjectProps"] = prot.asc_setGraphicObjectProps;
  prot["asc_getOriginalImageSize"] = prot.asc_getOriginalImageSize;
  prot["asc_changeShapeType"] = prot.asc_changeShapeType;
  prot["asc_setInterfaceDrawImagePlaceShape"] = prot.asc_setInterfaceDrawImagePlaceShape;
  prot["asc_setInterfaceDrawImagePlaceTextArt"] = prot.asc_setInterfaceDrawImagePlaceTextArt;
  prot["asc_changeImageFromFile"] = prot.asc_changeImageFromFile;
  prot["asc_putPrLineSpacing"] = prot.asc_putPrLineSpacing;
  prot["asc_addTextArt"] = prot.asc_addTextArt;
  prot["asc_canEditCrop"] = prot.asc_canEditCrop;
  prot["asc_startEditCrop"] = prot.asc_startEditCrop;
  prot["asc_endEditCrop"] = prot.asc_endEditCrop;
  prot["asc_cropFit"] = prot.asc_cropFit;
  prot["asc_cropFill"] = prot.asc_cropFill;
  prot["asc_putLineSpacingBeforeAfter"] = prot.asc_putLineSpacingBeforeAfter;
  prot["asc_setDrawImagePlaceParagraph"] = prot.asc_setDrawImagePlaceParagraph;
  prot["asc_changeShapeImageFromFile"] = prot.asc_changeShapeImageFromFile;
  prot["asc_AddMath"] = prot.asc_AddMath;
  prot["asc_ConvertMathView"] = prot.asc_ConvertMathView;
  prot["asc_SetMathProps"] = prot.asc_SetMathProps;
  //----------------------------------------------------------------------------------------------------------------------

  // Spellcheck
  prot["asc_setDefaultLanguage"] = prot.asc_setDefaultLanguage;
  prot["asc_nextWord"] = prot.asc_nextWord;
  prot["asc_replaceMisspelledWord"]= prot.asc_replaceMisspelledWord;
  prot["asc_ignoreMisspelledWord"] = prot.asc_ignoreMisspelledWord;
  prot["asc_spellCheckAddToDictionary"] = prot.asc_spellCheckAddToDictionary;
  prot["asc_spellCheckClearDictionary"] = prot.asc_spellCheckClearDictionary;
  prot["asc_cancelSpellCheck"] = prot.asc_cancelSpellCheck;
  prot["asc_ignoreNumbers"] = prot.asc_ignoreNumbers;
  prot["asc_ignoreUppercase"] = prot.asc_ignoreUppercase;

  // Frozen pane
  prot["asc_freezePane"] = prot.asc_freezePane;

  // Sparklines
  prot["asc_setSparklineGroup"] = prot.asc_setSparklineGroup;
  prot["asc_addSparklineGroup"] = prot.asc_addSparklineGroup;

  // Cell interface
  prot["asc_getCellInfo"] = prot.asc_getCellInfo;
  prot["asc_getActiveCellCoord"] = prot.asc_getActiveCellCoord;
  prot["asc_getAnchorPosition"] = prot.asc_getAnchorPosition;
  prot["asc_setCellFontName"] = prot.asc_setCellFontName;
  prot["asc_setCellFontSize"] = prot.asc_setCellFontSize;
  prot["asc_setCellBold"] = prot.asc_setCellBold;
  prot["asc_setCellItalic"] = prot.asc_setCellItalic;
  prot["asc_setCellUnderline"] = prot.asc_setCellUnderline;
  prot["asc_setCellStrikeout"] = prot.asc_setCellStrikeout;
  prot["asc_setCellSubscript"] = prot.asc_setCellSubscript;
  prot["asc_setCellSuperscript"] = prot.asc_setCellSuperscript;
  prot["asc_setCellAlign"] = prot.asc_setCellAlign;
  prot["asc_setCellVertAlign"] = prot.asc_setCellVertAlign;
  prot["asc_setCellTextWrap"] = prot.asc_setCellTextWrap;
  prot["asc_setCellTextShrink"] = prot.asc_setCellTextShrink;
  prot["asc_setCellTextColor"] = prot.asc_setCellTextColor;
  prot["asc_setCellFill"] = prot.asc_setCellFill;
  prot["asc_setCellBackgroundColor"] = prot.asc_setCellBackgroundColor;
  prot["asc_setCellBorders"] = prot.asc_setCellBorders;
  prot["asc_setCellFormat"] = prot.asc_setCellFormat;
  prot["asc_setCellAngle"] = prot.asc_setCellAngle;
  prot["asc_setCellStyle"] = prot.asc_setCellStyle;
  prot["asc_increaseCellDigitNumbers"] = prot.asc_increaseCellDigitNumbers;
  prot["asc_decreaseCellDigitNumbers"] = prot.asc_decreaseCellDigitNumbers;
  prot["asc_increaseFontSize"] = prot.asc_increaseFontSize;
  prot["asc_decreaseFontSize"] = prot.asc_decreaseFontSize;
  prot["asc_setCellIndent"] = prot.asc_setCellIndent;
  prot["asc_setCellProtection"] = prot.asc_setCellProtection;
  prot["asc_setCellLocked"] = prot.asc_setCellLocked;
  prot["asc_setCellHiddenFormulas"] = prot.asc_setCellHiddenFormulas;
  prot["asc_checkProtectedRange"] = prot.asc_checkProtectedRange;
  prot["asc_checkActiveCellPassword"] = prot.asc_checkActiveCellPassword;
  prot["asc_checkLockedCells"] = prot.asc_checkLockedCells;


  prot["asc_formatPainter"] = prot.asc_formatPainter;
  prot["asc_showAutoComplete"] = prot.asc_showAutoComplete;
  prot["asc_getHeaderFooterMode"] = prot.asc_getHeaderFooterMode;
  prot["asc_getActiveRangeStr"] = prot.asc_getActiveRangeStr;


  prot["asc_onMouseUp"] = prot.asc_onMouseUp;

  prot["asc_selectFunction"] = prot.asc_selectFunction;
  prot["asc_insertHyperlink"] = prot.asc_insertHyperlink;
  prot["asc_removeHyperlink"] = prot.asc_removeHyperlink;
  prot["asc_getFullHyperlinkLength"] = prot.asc_getFullHyperlinkLength;


  prot["asc_cleanSelectRange"] = prot.asc_cleanSelectRange;
  prot["asc_insertInCell"] = prot.asc_insertInCell;
  prot["asc_getFormulasInfo"] = prot.asc_getFormulasInfo;
  prot["asc_getFormulaLocaleName"] = prot.asc_getFormulaLocaleName;
  prot["asc_getFormulaNameByLocale"] = prot.asc_getFormulaNameByLocale;
  prot["asc_startWizard"] = prot.asc_startWizard;
  prot["asc_canEnterWizardRange"] = prot.asc_canEnterWizardRange;
  prot["asc_insertArgumentsInFormula"] = prot.asc_insertArgumentsInFormula;
  prot["asc_calculate"] = prot.asc_calculate;
  prot["asc_setFontRenderingMode"] = prot.asc_setFontRenderingMode;
  prot["asc_setSelectionDialogMode"] = prot.asc_setSelectionDialogMode;
  prot["asc_ChangeColorScheme"] = prot.asc_ChangeColorScheme;
  prot["asc_ChangeColorSchemeByIdx"] = prot.asc_ChangeColorSchemeByIdx;
  prot["asc_setListType"] = prot.asc_setListType;
  prot["asc_getCurrentListType"] = prot.asc_getCurrentListType;
  /////////////////////////////////////////////////////////////////////////
  ///////////////////CoAuthoring and Chat api//////////////////////////////
  /////////////////////////////////////////////////////////////////////////
  prot["asc_coAuthoringChatSendMessage"] = prot.asc_coAuthoringChatSendMessage;
  prot["asc_coAuthoringGetUsers"] = prot.asc_coAuthoringGetUsers;
  prot["asc_coAuthoringChatGetMessages"] = prot.asc_coAuthoringChatGetMessages;
  prot["asc_coAuthoringDisconnect"] = prot.asc_coAuthoringDisconnect;

  // other
  prot["asc_stopSaving"] = prot.asc_stopSaving;
  prot["asc_continueSaving"] = prot.asc_continueSaving;

  prot['sendEvent'] = prot.sendEvent;

  // Version History
  prot["asc_undoAllChanges"] = prot.asc_undoAllChanges;

  prot["asc_setLocalization"] = prot.asc_setLocalization;

  // native
  prot["asc_nativeOpenFile"] = prot.asc_nativeOpenFile;
  prot["asc_nativeCalculateFile"] = prot.asc_nativeCalculateFile;
  prot["asc_nativeApplyChanges"] = prot.asc_nativeApplyChanges;
  prot["asc_nativeApplyChanges2"] = prot.asc_nativeApplyChanges2;
  prot["asc_nativeGetFile"] = prot.asc_nativeGetFile;
  prot["asc_nativeGetFileData"] = prot.asc_nativeGetFileData;
  prot["asc_nativeCalculate"] = prot.asc_nativeCalculate;
  prot["asc_nativePrint"] = prot.asc_nativePrint;
  prot["asc_nativePrintPagesCount"] = prot.asc_nativePrintPagesCount;
  prot["asc_nativeGetPDF"] = prot.asc_nativeGetPDF;

  prot['asc_isOffline'] = prot.asc_isOffline;
  prot['asc_getUrlType'] = prot.asc_getUrlType;
  prot['asc_prepareUrl'] = prot.asc_prepareUrl;

  prot['asc_getSessionToken'] = prot.asc_getSessionToken;
  // Builder
  prot['asc_nativeInitBuilder'] = prot.asc_nativeInitBuilder;
  prot['asc_SetSilentMode'] = prot.asc_SetSilentMode;

  // plugins
  prot["asc_pluginsRegister"]       = prot.asc_pluginsRegister;
  prot["asc_pluginRun"]             = prot.asc_pluginRun;
  prot["asc_pluginStop"]            = prot.asc_pluginStop;
  prot["asc_pluginResize"]          = prot.asc_pluginResize;
  prot["asc_pluginButtonClick"]     = prot.asc_pluginButtonClick;
  prot["asc_startEditCurrentOleObject"]         = prot.asc_startEditCurrentOleObject;
  prot["asc_pluginEnableMouseEvents"] = prot.asc_pluginEnableMouseEvents;

  // system input
  prot["SetTextBoxInputMode"]       = prot.SetTextBoxInputMode;
  prot["GetTextBoxInputMode"]       = prot.GetTextBoxInputMode;

  prot["asc_InputClearKeyboardElement"] = prot.asc_InputClearKeyboardElement;

  prot["asc_OnHideContextMenu"] = prot.asc_OnHideContextMenu;
  prot["asc_OnShowContextMenu"] = prot.asc_OnShowContextMenu;

  // pivot
  prot["asc_getAddPivotTableOptions"] = prot.asc_getAddPivotTableOptions;
  prot["asc_insertPivotNewWorksheet"] = prot.asc_insertPivotNewWorksheet;
  prot["asc_insertPivotExistingWorksheet"] = prot.asc_insertPivotExistingWorksheet;
  prot["asc_refreshAllPivots"] = prot.asc_refreshAllPivots;
  prot["asc_getPivotInfo"] = prot.asc_getPivotInfo;
  prot["asc_getPivotShowValueAsInfo"] = prot.asc_getPivotShowValueAsInfo;
  prot["asc_pivotShowDetails"] = prot.asc_pivotShowDetails;
  prot["asc_pivotShowDetailsHeader"] = prot.asc_pivotShowDetailsHeader;
	// signatures
  prot["asc_addSignatureLine"] 		     = prot.asc_addSignatureLine;
  prot["asc_CallSignatureDblClickEvent"] = prot.asc_CallSignatureDblClickEvent;
  prot["asc_getRequestSignatures"] 	     = prot.asc_getRequestSignatures;
  prot["asc_AddSignatureLine2"]          = prot.asc_AddSignatureLine2;
  prot["asc_Sign"]             		     = prot.asc_Sign;
  prot["asc_RequestSign"]             	 = prot.asc_RequestSign;
  prot["asc_ViewCertificate"] 		     = prot.asc_ViewCertificate;
  prot["asc_SelectCertificate"] 	     = prot.asc_SelectCertificate;
  prot["asc_GetDefaultCertificate"]      = prot.asc_GetDefaultCertificate;
  prot["asc_getSignatures"] 		     = prot.asc_getSignatures;
  prot["asc_isSignaturesSupport"] 	     = prot.asc_isSignaturesSupport;
  prot["asc_isProtectionSupport"] 		 = prot.asc_isProtectionSupport;
  prot["asc_isAnonymousSupport"] 		 = prot.asc_isAnonymousSupport;
  prot["asc_RemoveSignature"] 		= prot.asc_RemoveSignature;
  prot["asc_RemoveAllSignatures"] 	= prot.asc_RemoveAllSignatures;
  prot["asc_gotoSignature"] 	    = prot.asc_gotoSignature;
  prot["asc_getSignatureSetup"] 	= prot.asc_getSignatureSetup;

  // password
  prot["asc_setCurrentPassword"]    = prot.asc_setCurrentPassword;
  prot["asc_resetPassword"] 		= prot.asc_resetPassword;

  // mobile
  prot["asc_Remove"] = prot.asc_Remove;

  prot["asc_getSortProps"] = prot.asc_getSortProps;
  prot["asc_setSortProps"] = prot.asc_setSortProps;

  prot["asc_validSheetName"] = prot.asc_validSheetName;

  prot["asc_getRemoveDuplicates"] = prot.asc_getRemoveDuplicates;
  prot["asc_setRemoveDuplicates"] = prot.asc_setRemoveDuplicates;

  //conditional formatting
  prot["asc_getCF"]            = prot.asc_getCF;
  prot["asc_setCF"]            = prot.asc_setCF;
  prot["asc_getPreviewCF"]     = prot.asc_getPreviewCF;
  prot["asc_clearCF"]          = prot.asc_clearCF;
  prot["asc_getCFIconsByType"] = prot.asc_getCFIconsByType;
  prot["asc_getCFPresets"]     = prot.asc_getCFPresets;
  prot["asc_getFullCFIcons"]   = prot.asc_getFullCFIcons;
  prot["asc_isValidDataRefCf"] = prot.asc_isValidDataRefCf;

  prot["asc_beforeInsertSlicer"] = prot.asc_beforeInsertSlicer;
  prot["asc_insertSlicer"] = prot.asc_insertSlicer;

  //data validation
  prot["asc_setDataValidation"] = prot.asc_setDataValidation;
  prot["asc_getDataValidationProps"] = prot.asc_getDataValidationProps;

  prot["asc_setSkin"] = prot.asc_setSkin;

  prot["asc_getEscapeSheetName"] = prot.asc_getEscapeSheetName;


  prot["asc_ConvertEquationToMath"] = prot.asc_ConvertEquationToMath;

  prot["asc_getProtectedRanges"]           = prot.asc_getProtectedRanges;
  prot["asc_setProtectedRanges"]           = prot.asc_setProtectedRanges;
  prot["asc_checkProtectedRangesPassword"] = prot.asc_checkProtectedRangesPassword;
  prot["asc_checkProtectedRangeName"]      = prot.asc_checkProtectedRangeName;

  prot["asc_getProtectedSheet"]            = prot.asc_getProtectedSheet;
  prot["asc_setProtectedSheet"]            = prot.asc_setProtectedSheet;
  prot["asc_isProtectedSheet"]             = prot.asc_isProtectedSheet;
  prot["asc_getProtectedWorkbook"]         = prot.asc_getProtectedWorkbook;
  prot["asc_setProtectedWorkbook"]         = prot.asc_setProtectedWorkbook;
  prot["asc_isProtectedWorkbook"]          = prot.asc_isProtectedWorkbook;

  //sheet-views
  prot["asc_addNamedSheetView"] = prot.asc_addNamedSheetView;
  prot["asc_getNamedSheetViews"] = prot.asc_getNamedSheetViews;
  prot["asc_deleteNamedSheetViews"] = prot.asc_deleteNamedSheetViews;
  prot["asc_setActiveNamedSheetView"] = prot.asc_setActiveNamedSheetView;
  prot["asc_getActiveNamedSheetView"] = prot.asc_getActiveNamedSheetView;

  prot["getPrintOptionsJson"] = prot.getPrintOptionsJson;

  prot["asc_EditSelectAll"] = prot.asc_EditSelectAll;

  prot["asc_setDate1904"] = prot.asc_setDate1904;
  prot["asc_getDate1904"] = prot.asc_getDate1904;

  prot["onWorksheetChange"] = prot.onWorksheetChange;  

  prot["asc_addCellWatches"]               = prot.asc_addCellWatches;
  prot["asc_deleteCellWatches"]            = prot.asc_deleteCellWatches;
  prot["asc_getCellWatches"]               = prot.asc_getCellWatches;

  prot["asc_getExternalReferences"] = prot.asc_getExternalReferences;
  prot["asc_updateExternalReferences"] = prot.asc_updateExternalReferences;
  prot["asc_removeExternalReferences"] = prot.asc_removeExternalReferences;
  prot["asc_openExternalReference"] = prot.asc_openExternalReference;
  prot["asc_changeExternalReference"] = prot.asc_changeExternalReference;


  prot["asc_fillHandleDone"] = prot.asc_fillHandleDone;
  prot["asc_canFillHandle"]  = prot.asc_canFillHandle;

  prot["asc_ImportXmlStart"] = prot.asc_ImportXmlStart;
  prot["asc_ImportXmlEnd"]   = prot.asc_ImportXmlEnd;

  prot["asc_addUserProtectedRange"]       = prot.asc_addUserProtectedRange;
  prot["asc_changeUserProtectedRange"]    = prot.asc_changeUserProtectedRange;
  prot["asc_deleteUserProtectedRange"]    = prot.asc_deleteUserProtectedRange;
  prot["asc_getUserProtectedRanges"]      = prot.asc_getUserProtectedRanges;
  prot["asc_checkUserProtectedRangeName"] = prot.asc_checkUserProtectedRangeName;
  prot["asc_SetSheetViewType"]   = prot.asc_SetSheetViewType;
  prot["asc_GetSheetViewType"]   = prot.asc_GetSheetViewType;

  prot["asc_ChangeTextCase"]   = prot.asc_ChangeTextCase;
  prot["asc_TracePrecedents"]     = prot.asc_TracePrecedents;
  prot["asc_TraceDependents"]     = prot.asc_TraceDependents;
  prot["asc_RemoveTraceArrows"]   = prot.asc_RemoveTraceArrows;

  prot["asc_InsertPageBreak"]         = prot.asc_InsertPageBreak;
  prot["asc_RemovePageBreak"]         = prot.asc_RemovePageBreak;
  prot["asc_ResetAllPageBreaks"]      = prot.asc_ResetAllPageBreaks;
  prot["asc_GetPageBreaksDisableType"]= prot.asc_GetPageBreaksDisableType;







})(window);
