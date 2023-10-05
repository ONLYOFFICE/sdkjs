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

(function(window, undefined)
{
	var prot;

	// Import
	var c_oEditorId = AscCommon.c_oEditorId;
	var c_oCloseCode = AscCommon.c_oCloseCode;
	var DownloadType = AscCommon.DownloadType;
	var c_oAscFrameDataType = AscCommon.c_oAscFrameDataType;

	var c_oAscError           = Asc.c_oAscError;
	var c_oAscAsyncAction     = Asc.c_oAscAsyncAction;
	var c_oAscAsyncActionType = Asc.c_oAscAsyncActionType;

	/** @constructor */
	function baseEditorsApi(config, editorId)
	{
		this.editorId      = editorId;
		this.isLoadFullApi = false;
        this.modulesCount  = 0;
		this.modulesLoaded = 0;
		this.openResult    = null;
		this.isOpenOOXInBrowser = false;
		this.openOOXInBrowserZip = null;

		this.HtmlElementName = config['id-view'] || '';
		this.HtmlElement     = null;

		this.isMobileVersion = (config['mobile'] === true);
		this.isEmbedVersion = (config['embedded'] === true);

		this.isViewMode = false;
		this.restrictions = Asc.c_oAscRestrictionType.None;

		this.FontLoader  = null;
		this.ImageLoader = null;

		this.LoadedObject        = null;
		this.DocumentType        = 0; // 0 - empty, 1 - test, 2 - document (from json)
		this.DocInfo             = null;
		this.documentId          = undefined;
		this.documentUserId      = undefined;
		this.documentUrl         = "null";
		this.documentUrlChanges  = null;
		this.documentTokenChanges  = null;
		this.documentCallbackUrl = undefined;		// Ссылка для отправления информации о документе
		this.documentFormat      = "null";
		this.documentTitle       = "null";
		this.documentFormatSave  = Asc.c_oAscFileType.UNKNOWN;

		this.documentOpenOptions = undefined;		// Опции при открытии (пока только опции для CSV)

		// Тип состояния на данный момент (сохранение, открытие или никакое)
		this.advancedOptionsAction = AscCommon.c_oAscAdvancedOptionsAction.None;
		this.OpenDocumentProgress  = new AscCommon.COpenProgress();
		var sProtocol              = window.location.protocol;
		this.documentOrigin        = ((sProtocol && '' !== sProtocol) ? sProtocol + '//' : '') + window.location.host; // for presentation theme url
		this.documentPathname      = window.location.pathname; // for presentation theme url

		// Переменная отвечает, получили ли мы ответ с сервера совместного редактирования
		this.ServerIdWaitComplete = false;

		// Long action
		this.IsLongActionCurrent       = 0;
		this.LongActionCallbacks       = [];
		this.LongActionCallbacksParams = [];

		// AutoSave
		this.autoSaveGap = 0;					// Интервал автосохранения (0 - означает, что автосохранения нет) в милесекундах
		this.lastSaveTime = null;				// Время последнего сохранения
		this.autoSaveGapFast = 2000;			// Интервал быстрого автосохранения (когда человек один) - 2 сек.
		this.autoSaveGapSlow = 10 * 60 * 1000;	// Интервал медленного автосохранения (когда совместно) - 10 минут
		this.intervalWaitAutoSave = 1000;

		// Unlock document
		this.canUnlockDocument = false;
		this.canUnlockDocument2 = false;		// Дублирующий флаг, только для saveChanges или unLockDocument
		this.canStartCoAuthoring = false;

		this.isDocumentCanSave = false;			// Флаг, говорит о возможности сохранять документ (активна кнопка save или нет)

		// translate manager
		this.translateManager = AscCommon.translateManager.init(config['translate']);

		//shape names map by preset. Set from interface
		this.shapeNames = {};

		// Chart
		this.chartPreviewManager   = null;
		this.textArtPreviewManager = null;
		this.shapeElementId        = null;
		// Режим вставки диаграмм в редакторе документов
		this.isChartEditor         = false;
		this.isOpenedChartFrame    = false;

		this.isOleEditor = false;

		this.MathMenuLoad          = false;

		this.mathInputType = Asc.c_oAscMathInputType.Unicode;

		// CoAuthoring and Chat
		this.User                   = undefined;
		this.CoAuthoringApi         = new AscCommon.CDocsCoApi();
		this.isCoAuthoringEnable    = true;
		// Массив lock-ов, которые были на открытии документа
		this.arrPreOpenLocksObjects = [];

		// Spell Checking
		this.SpellCheckUrl = '';    // Ссылка сервиса для проверки орфографии

		// Результат получения лицензии
		this.licenseResult       = null;
		// Получили ли лицензию
		this.isOnLoadLicense     = false;
		// Переменная, которая отвечает, послали ли мы окончание открытия документа
		this.isDocumentLoadComplete = false;
		// Переменная, которая отвечает, послали ли мы окончание открытия документа
		this.isPreOpenLocks = true;
		this.isApplyChangesOnOpenEnabled = true;
		this.isProtectionSupport = true;
		this.isAnonymousSupport = true;

		this.canSave    = true;        // Флаг нужен чтобы не происходило сохранение пока не завершится предыдущее сохранение
		this.IsUserSave = false;    // Флаг, контролирующий сохранение было сделано пользователем или нет (по умолчанию - нет)
		this.isForceSaveOnUserSave = false;
        this.forceSaveButtonTimeout = null;
        this.forceSaveButtonContinue = false;
        this.forceSaveTimeoutTimeout = null;
		this.forceSaveForm = null;
		this.disconnectOnSave = null;
		this.disconnectRestrictions = null;//to restore restrictions after disconnect
		this.forceSaveUndoRequest = false; // Флаг нужен, чтобы мы знали, что данное сохранение пришло по запросу Undo в совместке

		// Version History
		this.VersionHistory = null;				// Объект, который отвечает за точку в списке версий

		//Флаги для применения свойств через слайдеры
		this.noCreatePoint     = false;
		this.exucuteHistory    = false;
		this.exucuteHistoryEnd = false;

		this.selectSearchingResults = false;

		this.isSendStandartTextures = false;

		this.tmpFocus = null;

		this.fCurCallback = null;

		this.pluginsManager = null;

		this.isLockTargetUpdate   = false;
		this.isLockScrollToTarget = false;

		this.lastWorkTime = 0;

		this.signatures = [];

		this.currentPassword = "";

		this.disableAutostartMacros = false;
		this.macros = null;
		this.vbaMacros = null;
		this.vbaProject = null;

        this.openFileCryptBinary = null;

        this.copyOutEnabled = true;

		this.watermarkDraw = null;

		this.SaveAfterMacros = false;

		// Spell Checking
		this.SpellCheckApi = new AscCommon.CSpellCheckApi();
		this.isSpellCheckEnable = true;

		// macros & plugins events
		this.internalEvents = {};

		this.skinObject = config['skin'];
		this.isDarkMode = false;

		this.Shortcuts = new AscCommon.CShortcuts();
		this.initDefaultShortcuts();

		this.isUseNativeViewer = true;

		this.openedAt = undefined;
		this.maxChangesSize = 0;
		this.binaryChanges = false;

		this.isBlurEditor = false;


		this.formatPainter = new AscCommon.CFormatPainter(this);
		this.eyedropper = new AscCommon.CEyedropper(this);
		this.inkDrawer = new AscCommon.CInkDrawer(this);
		this._correctEmbeddedWork();

		return this;
	}

	baseEditorsApi.prototype._init                           = function()
	{
		if (window["AscDesktopEditor"])
			window["AscDesktopEditor"]["CreateEditorApi"](this);

		var t            = this;
		//Asc.editor = Asc['editor'] = AscCommon['editor'] = AscCommon.editor = this; // ToDo сделать это!
		this.HtmlElement = document.getElementById(this.HtmlElementName);
		if (this.HtmlElement)
		{
			// запрещаем действия браузера по умолчанию
			this.HtmlElement.style.touchAction = "none";
		}

		// init OnMessage
		AscCommon.InitOnMessage(function(error, url)
		{
			if (c_oAscError.ID.No !== error)
			{
				t.sendEvent("asc_onError", error, c_oAscError.Level.NoCritical);
			}
			else
			{
				t._addImageUrl([url]);
			}

			t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
		});

		AscCommon.loadSdk(this._editorNameById(), function()
		{
			t.isLoadFullApi = true;

			t._onEndLoadSdk();
			t.onEndLoadDocInfo();
		}, function(err) {
			t.sendEvent("asc_onError", Asc.c_oAscError.ID.LoadingScriptError, c_oAscError.Level.Critical);
		});

		this._loadModules();

		AscCommon.loadChartStyles(function() {}, function(err) {
			t.sendEvent("asc_onError", Asc.c_oAscError.ID.LoadingScriptError, c_oAscError.Level.NoCritical);
		});



		var oldOnError = window.onerror;
		window.onerror = function(errorMsg, url, lineNumber, column, errorObj) {
			//send only first error to reduce number of requests. also following error may be consequences of first
			window.onerror = oldOnError;
			let editorInfo = t.getEditorErrorInfo();
			let memoryInfo = AscCommon.getMemoryInfo();
			var msg = 'Error: ' + errorMsg + '\n Script: ' + url + '\n Line: ' + lineNumber + ':' + column +
				'\n userAgent: ' + (navigator.userAgent || navigator.vendor || window.opera) + '\n platform: ' +
				navigator.platform + '\n isLoadFullApi: ' + t.isLoadFullApi + '\n isDocumentLoadComplete: ' +
				t.isDocumentLoadComplete + (editorInfo ? '\n ' + editorInfo : "") +
				(memoryInfo ? '\n performance.memory: ' + memoryInfo : "") +
				'\n StackTrace: ' + (errorObj ? errorObj.stack : "");
			AscCommon.sendClientLog("error", "changesError: " + msg, t);
			if (t.isLoadFullApi ) {
				if(t.isDocumentLoadComplete) {
					//todo disconnect and downloadAs ability
					t.sendEvent("asc_onError", Asc.c_oAscError.ID.EditingError, c_oAscError.Level.NoCritical);
					t.asc_setViewMode(true);
				}
				else {
					t.sendEvent("asc_onError", Asc.c_oAscError.ID.ConvertationOpenError, c_oAscError.Level.Critical);
				}
			}
			if (oldOnError) {
				return oldOnError.apply(this, arguments);
			} else {
				return false;
			}
		};

		// disable mousewheel on macOS
		if (AscCommon.AscBrowser.isMacOs)
		{
			document.body.onmousewheel = function(e) {
				if (e.stopPropagation)
					e.stopPropagation();
				e.returnValue = false;
				return false;
			};
		}
	};
	baseEditorsApi.prototype._correctEmbeddedWork = function()
	{
		let body = document.getElementById("viewport");
		if (body && body.addEventListener)
		{
			body.addEventListener("mousewheel", function(e) {
				if (e.ctrlKey || e.altKey || e.shiftKey || e.metaKey)
					return;
				e.preventDefault && e.preventDefault();
				e.stopPropagation && e.stopPropagation();
				return false;
			}, false);
			body.addEventListener("DOMMouseScroll", function(e)
			{
				if (e.ctrlKey || e.altKey || e.shiftKey || e.metaKey)
					return;
				e.preventDefault && e.preventDefault();
				e.stopPropagation && e.stopPropagation();
				return false;
			}, false);
		}

		var _t = this;
		window.addEventListener && window.addEventListener("blur", function(){
			_t.isBlurEditor = true;
		}, false);
	};
	baseEditorsApi.prototype._editorNameById                 = function()
	{
		var res = '';
		switch (this.editorId)
		{
			case c_oEditorId.Word:
				res = 'word';
				break;
			case c_oEditorId.Spreadsheet:
				res = 'cell';
				break;
			case c_oEditorId.Presentation:
				res = 'slide';
				break;
		}
		return res;
	};
	baseEditorsApi.prototype.getEditorId                     = function()
	{
		return this.editorId;
	};
	baseEditorsApi.prototype.isPdfEditor = function()
	{
		return false;
	};
	baseEditorsApi.prototype.isDocumentRenderer = function()
	{
		return false;
	};
	baseEditorsApi.prototype.getEditorErrorInfo = function()
	{
		return "";
	};

	// modules
	baseEditorsApi.prototype._loadModules = function()
	{
		this.modulesCount = 1;
		AscFonts.load(this, this._onSuccessLoadModule.bind(this), this._onErrorLoadModule.bind(this));
		//AscCommon.zlib_load(this._onSuccessLoadModule.bind(this), this._onErrorLoadModule.bind(this));
	};
	baseEditorsApi.prototype._onSuccessLoadModule = function()
	{
		this.modulesLoaded++;
		if (this.modulesLoaded === this.modulesCount)
			this.onEndLoadFile(null);
	};
	baseEditorsApi.prototype._onErrorLoadModule = function()
	{
		this.sendEvent("asc_onError", Asc.c_oAscError.ID.LoadingScriptError, c_oAscError.Level.Critical);
	};
	baseEditorsApi.prototype._isLoadedModules = function()
	{
		return (this.modulesLoaded === this.modulesCount) ? true : false;
	};

	baseEditorsApi.prototype.asc_loadFontsFromServer             = function (fonts)
	{
		if (!fonts)
		{
			fonts = ["Arial", "Symbol", "Wingdings", "Courier New", "Times New Roman"];
		}
		this.FontLoader.LoadFontsFromServer(fonts);
	};
	baseEditorsApi.prototype.asc_GetFontThumbnailsPath       = function()
	{
		return '../Common/Images/';
	};
	baseEditorsApi.prototype.asc_getDocumentName             = function()
	{
		return this.documentTitle;
	};
	baseEditorsApi.prototype.asc_getAppProps                 = function()
	{
		return null;
	};
	baseEditorsApi.prototype.asc_getCoreProps                = function()
	{
		var oCore = this.getInternalCoreProps();
		if(oCore) {
			return oCore.copy();
		}
		return null;
	};
	baseEditorsApi.prototype.getInternalCoreProps                = function()
	{
		return null;
	};
	baseEditorsApi.prototype.isFrameEditor = function () {
		return !!(this.isChartEditor ||  this.isEditOleMode); // TODO: solve the confusion
	};
	baseEditorsApi.prototype.asc_setCoreProps                = function(oProps)
	{

	};
	baseEditorsApi.prototype.asc_setDocInfo                  = function(oDocInfo)
	{
		var oldInfo = this.DocInfo;
		if (oDocInfo)
		{
			this.DocInfo = oDocInfo;
		}

		if (this.DocInfo)
		{
			this.documentId          = this.DocInfo.get_Id();
			this.documentUserId      = this.DocInfo.get_UserId();
			this.documentUrl         = this.DocInfo.get_Url();
			this.documentTitle       = this.DocInfo.get_Title();
			this.documentFormat      = this.DocInfo.get_Format();
			this.documentCallbackUrl = this.DocInfo.get_CallbackUrl();

			this.documentOpenOptions = this.DocInfo.asc_getOptions();

			var permissions = this.DocInfo.asc_getPermissions();
			if (permissions && undefined !== permissions['copy'])
			{
				this.copyOutEnabled = permissions['copy'];
			}

			this.User = new AscCommon.asc_CUser();
			this.User.setId(this.DocInfo.get_UserId());
			this.User.setUserName(this.DocInfo.get_UserName());
			this.User.setFirstName(this.DocInfo.get_FirstName());
			this.User.setLastName(this.DocInfo.get_LastName());

			//чтобы в versionHistory был один documentId для auth и open
			this.CoAuthoringApi.setDocId(this.documentId);

			if (this.documentOpenOptions)
			{
				if (this.documentOpenOptions["watermark"])
				{
					this.watermarkDraw = new AscCommon.CWatermarkOnDraw(this.documentOpenOptions["watermark"], this);
					this.watermarkDraw.CheckParams(this);
				}

				if (false === this.documentOpenOptions["oform"] && AscCommon.IsSupportOFormFeature())
				{
					window["Asc"]["Addons"]["forms"] = false;
					AscCommon.g_oTableId.InitOFormClasses();
				}
			}
		}

		if (AscCommon.chartMode === this.documentUrl)
		{
			this.isChartEditor = true;
            AscCommon.EncryptionWorker.isChartEditor = true;
			this.DocInfo.put_OfflineApp(true);
		}
		else if (AscCommon.offlineMode === this.documentUrl)
		{
			this.DocInfo.put_OfflineApp(true);
		}

		if (this.DocInfo.get_EncryptedInfo())
		{
			if (undefined !== window["AscDesktopEditor"])
			{
				var obj = this.DocInfo.get_EncryptedInfo();
				obj["userId"] = this.documentUserId;
				window["AscDesktopEditor"]["execCommand"]("portal:cryptoinfo", JSON.stringify(obj));
			}
		}

		if (undefined !== window["AscDesktopEditor"] && !(this.DocInfo && this.DocInfo.get_OfflineApp()))
		{
			window["AscDesktopEditor"]["SetDocumentName"](this.documentTitle);
		}

		if (!this.isChartEditor && undefined !== window["AscDesktopEditor"] && undefined !== window["AscDesktopEditor"]["CryptoMode"])
		{
			this.DocInfo.put_Encrypted(0 < window["AscDesktopEditor"]["CryptoMode"]);
		}

		if (!oldInfo)
		{
			this.onEndLoadDocInfo();
		}
	};
	baseEditorsApi.prototype.asc_changeDocInfo = function(oDocInfo)
	{
		var rData = {
			"c": 'changedocinfo',
			"id": this.documentId,
			"username": oDocInfo.asc_getUserName()
		};
		var t            = this;
		t.fCurCallback   = function(input)
		{
			if (null != input && "changedocinfo" == input["type"])
			{
				if ('ok' === input["status"]) {
					t.DocInfo.asc_getUserInfo().asc_putFullName(oDocInfo.asc_getUserName());
					t.User.setUserName(oDocInfo.asc_getUserName());
				} else {
					t.sendEvent("asc_onError", AscCommon.mapAscServerErrorToAscError(parseInt(input["data"])),
						c_oAscError.Level.NoCritical);
				}
			}
			else
			{
				t.sendEvent("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
			}
		};
		AscCommon.sendCommand(this, null, rData);
	};
	baseEditorsApi.prototype.asc_isCrypto = function()
	{
		if (this.DocInfo && this.DocInfo.get_Encrypted() === true)
			return true;
		return false;
	};
	baseEditorsApi.prototype.asc_enableKeyEvents             = function(isEnabled, isFromInput)
	{
	};
	// Copy/Past/Cut
	baseEditorsApi.prototype.asc_IsFocus                     = function(bIsNaturalFocus)
	{
		var _ret = false;
		if (this.WordControl.IsFocus)
			_ret = true;
		if (_ret && bIsNaturalFocus && this.WordControl.TextBoxInputFocus)
			_ret = false;
		return _ret;
	};
	baseEditorsApi.prototype.isCopyOutEnabled                = function()
	{
		return this.copyOutEnabled;
	};
	baseEditorsApi.prototype.sync_CanCopyCutCallback = function (bCanCopyCut)
	{
		this.sendEvent("asc_onCanCopyCut", bCanCopyCut);
	};
	baseEditorsApi.prototype.can_CopyCut = function ()
	{
		return true;
	};
	// target pos
	baseEditorsApi.prototype.asc_LockTargetUpdate		     = function(isLock)
	{
		this.isLockTargetUpdate = isLock;
	};
	baseEditorsApi.prototype.asc_LockScrollToTarget          = function(isLock)
	{
		this.isLockScrollToTarget = isLock;
	};
	baseEditorsApi.prototype.isLiveViewer                     = function()
	{
		return this.isViewMode && AscCommon.CollaborativeEditing.Is_Fast() && !this.VersionHistory;
	};
	// Events
	baseEditorsApi.prototype.sendEvent                       = function()
	{
	};
	baseEditorsApi.prototype.SendOpenProgress                = function()
	{
		this.sendEvent("asc_onOpenDocumentProgress", this.OpenDocumentProgress);
	};
	baseEditorsApi.prototype.sync_InitEditorFonts            = function(gui_fonts)
	{
		if (!this.isViewMode) {
			// корректируем имена для текущего языка интерфейса
			var currentLang = this.asc_getLocale();
			if (typeof currentLang !== "string")
				currentLang = (AscCommon.g_oDefaultCultureInfo && (typeof AscCommon.g_oDefaultCultureInfo.Name === "string")) ? AscCommon.g_oDefaultCultureInfo.Name : "en";

			AscFonts.g_fontApplication && AscFonts.g_fontApplication.CheckNamesForInterface(currentLang);
			this.sendEvent("asc_onInitEditorFonts", gui_fonts);
		}
	};
	baseEditorsApi.prototype.sync_StartAction                = function(type, id)
	{
		if (type !== c_oAscAsyncActionType.Empty)
			this.sendEvent('asc_onStartAction', type, id);
		//console.log("asc_onStartAction: type = " + type + " id = " + id);

		if (c_oAscAsyncActionType.BlockInteraction === type)
		{
			this.incrementCounterLongAction();
		}
	};
	baseEditorsApi.prototype.sync_EndAction                  = function(type, id)
	{
		if (type !== c_oAscAsyncActionType.Empty)
			this.sendEvent('asc_onEndAction', type, id);
		//console.log("asc_onEndAction: type = " + type + " id = " + id);

		if (c_oAscAsyncActionType.BlockInteraction === type)
		{
			this.decrementCounterLongAction();
		}
	};
	baseEditorsApi.prototype.sync_TryUndoInFastCollaborative = function()
	{
		this.sendEvent("asc_OnTryUndoInFastCollaborative");
	};
	baseEditorsApi.prototype.asc_setViewMode                 = function()
	{
	};
	/**
	 * @param val
	 * @param additionalSettings {AscCommon.CRestrictionSettings}
	 */
	baseEditorsApi.prototype.asc_setRestriction              = function(val, additionalSettings)
	{
		// Если выставлен флаг OnlySignatures, то его нельзя перебить никак, кроме как явно снять через
		// editor.removeRestriction(Asc.c_oAscRestrictionType.OnlySignatures)
		if (this.restrictions & Asc.c_oAscRestrictionType.OnlySignatures)
			return;

		this.restrictions = val;
		this.onUpdateRestrictions(additionalSettings);
	};
	baseEditorsApi.prototype.getViewMode                     = function()
	{
		return this.isViewMode;
	};
	baseEditorsApi.prototype.asc_addRestriction              = function(val)
	{
		this.restrictions |= val;
		this.onUpdateRestrictions();
	};
	baseEditorsApi.prototype.asc_removeRestriction           = function(val)
	{
		this.restrictions &= ~val;
		this.onUpdateRestrictions();
	};
	baseEditorsApi.prototype.addTableOleObject = function(oleBinary)
	{
		var _this = this;
		if (oleBinary) {
			if (!oleBinary['imageUrl'])
			{
				var base64Image = oleBinary['base64Image'];
				var fAfterUploadOleObjectImage = function (url) {
					oleBinary['imageUrl'] = url;
					_this.addTableOleObject(oleBinary);
				}
				var obj = {
					fAfterUploadOleObjectImage: fAfterUploadOleObjectImage
				};
				AscCommon.sendImgUrls(_this, [base64Image], function(data) {
					if (data[0] && data[0].path != null && data[0].url !== "error")
					{
						oleBinary['imageUrl'] = data[0].url;
						_this._addImageUrl([data[0].url], obj);
					}
				});
				return;
			}
			var blipUrl = oleBinary['imageUrl'];
			var arrImagesForAddToHistory = oleBinary['imagesForAddToHistory'];
			var binaryDataOfSheet = AscCommon.Base64.decode(oleBinary['binary']);
			var sizes = AscCommon.getSourceImageSize(blipUrl);
			var mmExtX = sizes.width * AscCommon.g_dKoef_pix_to_mm;
			var mmExtY = sizes.height * AscCommon.g_dKoef_pix_to_mm;
			this.asc_addOleObjectAction(blipUrl, binaryDataOfSheet, 'Excel.Sheet.12', mmExtX, mmExtY, sizes.width, sizes.height, true, arrImagesForAddToHistory);
		}
	};
	baseEditorsApi.prototype.asc_addTableOleObject = function(oleBinary)
	{
		this.addTableOleObject(oleBinary);
	};

	baseEditorsApi.prototype.asc_editTableOleObject = function(oleBinary)
	{
		this.editTableOleObject(oleBinary);
	};

	baseEditorsApi.prototype.sendFromFrameToGeneralEditor = function (oData)
	{
		this.sendEvent("asc_sendFromFrameToGeneralEditor", oData);
	};

	baseEditorsApi.prototype.sendFromGeneralToFrameEditor = function (oData)
	{
		this.sendEvent("asc_sendFromGeneralToFrameEditor", oData);
	};


	baseEditorsApi.prototype.asc_getInformationBetweenFrameAndGeneralEditor = function (oData)
	{
		const nType = oData["type"];
		const oInformation = oData["information"]
		switch (nType)
		{
			case c_oAscFrameDataType.GetLoadedImages:
			{
				this.CoAuthoringApi.onDocumentOpen(oInformation["inputWrap"], true);
				break;
			}
			case c_oAscFrameDataType.SendImageUrls:
			{
				AscCommon.sendImgUrls(this, oInformation["images"], function () {}, true, oInformation["token"]);
				break;
			}
			case c_oAscFrameDataType.OpenFrame: // TODO: это нужно перенести в web-apps,
				// при открытии и закрытии фрейма метод должен вызываться там, в 7.2 это сделать не успели
			{
				this.asc_onOpenChartFrame();
				break;
			}
			case c_oAscFrameDataType.ShowImageDialogInFrame:
			{
				const oOptions = {
				sendUrlsToFrameEditor: true
				};
				this.asc_addImage(oOptions);
				break;
			}
			case c_oAscFrameDataType.GetUrlsFromImageDialog:
			{
				let oOptions;
				if (this.oSaveObjectForAddImage)
				{
					oOptions = this.oSaveObjectForAddImage;
					delete this.oSaveObjectForAddImage;
				}
				const arrUrlsForAddToDocumentUrls = oInformation;
				AscCommon.g_oDocumentUrls.addUrls(arrUrlsForAddToDocumentUrls);
				this._addImageUrl(Object.keys(arrUrlsForAddToDocumentUrls).map(function (localUrl) {
					return arrUrlsForAddToDocumentUrls[localUrl]
				}), oOptions);
				this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
				break;
			}
			case c_oAscFrameDataType.StartUploadImageAction:
			{
				this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
			}
			default:
			{
				break;
			}
		}
	};

	baseEditorsApi.prototype.sendStartUploadImageActionToFrameEditor = function () {
		this.sendFromGeneralToFrameEditor({
			"type": c_oAscFrameDataType.StartUploadImageAction,
		});
	}

	baseEditorsApi.prototype.addImageUrlsFromGeneralToFrameEditor = function (urls)
	{
		const urlsForAddToDocumentUrls = {};
		for (let i = 0; i < urls.length; i += 1)
		{
			const url = urls[i];
			urlsForAddToDocumentUrls[AscCommon.g_oDocumentUrls.getLocal(url)] = url;
		}
		this.sendFromGeneralToFrameEditor({
			"type": AscCommon.c_oAscFrameDataType.GetUrlsFromImageDialog,
			"information": urlsForAddToDocumentUrls
		});
	};

	baseEditorsApi.prototype.editTableOleObject = function(oOleBinaryInfo)
	{
		const oThis = this;
		if (oOleBinaryInfo)
		{
			if (!oOleBinaryInfo['imageUrl'])
			{
				const sBase64Image = oOleBinaryInfo['base64Image'];
				const fAfterUploadOleObjectImage = function (sUrl) {
					oOleBinaryInfo['imageUrl'] = sUrl;
					oThis.editTableOleObject(oOleBinaryInfo);
				}
				const oOptions = {
					fAfterUploadOleObjectImage: fAfterUploadOleObjectImage
				};
				AscCommon.sendImgUrls(oThis, [sBase64Image], function(data) {
					if (data[0] && data[0].path != null && data[0].url !== "error")
					{
						oOleBinaryInfo['imageUrl'] = data[0].url;
						oThis._addImageUrl([data[0].url], oOptions);
					}
				});
				return;
			}

			const oController = this.getGraphicController();
			if (oController)
			{
				const arrSelectedObjects = AscFormat.getObjectsByTypesFromArr(oController.selectedObjects);
				if (arrSelectedObjects.oleObjects.length === 1)
				{
					const arrImagesForAddToHistory = oOleBinaryInfo['imagesForAddToHistory'];
					const oSelectedOleObject = arrSelectedObjects.oleObjects[0];
					const sBlipUrl = oOleBinaryInfo['imageUrl'];
					const arrBinaryDataOfSheet = AscCommon.Base64.decode(oOleBinaryInfo['binary']);
					const oSizes = AscCommon.getSourceImageSize(sBlipUrl);
					const nImageWidthCoefficient = oOleBinaryInfo['widthCoefficient'] || 1;
					const nImageHeightCoefficient = oOleBinaryInfo['heightCoefficient'] || 1;

					let nMMExtX, nMMExtY, nAdaptSizeHeight, nAdaptSizeWidth;
					nAdaptSizeWidth = (oSizes.width || 0) * nImageWidthCoefficient;
					nAdaptSizeHeight = (oSizes.height || 0) * nImageHeightCoefficient;
					nMMExtY = nAdaptSizeHeight * AscCommon.g_dKoef_pix_to_mm;
					nMMExtX = nAdaptSizeWidth * AscCommon.g_dKoef_pix_to_mm;
					this.asc_editOleObjectAction(oSelectedOleObject, sBlipUrl, arrBinaryDataOfSheet, nMMExtX, nMMExtY, nAdaptSizeWidth, nAdaptSizeHeight, arrImagesForAddToHistory);
				}
			}
		}
	};

	baseEditorsApi.prototype.asc_SaveDrawingAsPicture = function(sFormat)
	{
		let oController = this.getGraphicController();
		if(oController)
		{
			let sImageFormat = "image/png";
			if(sFormat === "jpg" || sFormat === "jpeg") {
				sImageFormat = "image/jpeg";
			}
			let oImageData = oController.getImageDataForSaving(true, sImageFormat);
			if(oImageData)
			{
				let a = document.createElement("a");
				let sSrc = oImageData.src;
				a.href = sSrc;
				let sExt = sSrc.substring("data:image/".length, sSrc.indexOf(";base64"));
				a.download = AscCommon.translateManager.getValue("Picture") + "." + sExt;
				a.click();
			}
		}
	};
	baseEditorsApi.prototype.canEdit                         = function()
	{
		return !this.isViewMode && this.restrictions === Asc.c_oAscRestrictionType.None;
	};
	baseEditorsApi.prototype.isRestrictionForms              = function()
	{
		return !!(this.restrictions & Asc.c_oAscRestrictionType.OnlyForms);
	};
	baseEditorsApi.prototype.isRestrictionComments           = function()
	{
		return !!(this.restrictions & Asc.c_oAscRestrictionType.OnlyComments);
	};
	baseEditorsApi.prototype.isRestrictionSignatures         = function()
	{
		return !!(this.restrictions & Asc.c_oAscRestrictionType.OnlySignatures);
	};
	baseEditorsApi.prototype.isRestrictionView               = function()
	{
		return !!(this.restrictions & Asc.c_oAscRestrictionType.View);
	};
	baseEditorsApi.prototype.onUpdateRestrictions = function()
	{
	};
	baseEditorsApi.prototype.isLongActionBase                    = function()
	{
		//word api overrides isLongAction with additional checks
		return (0 !== this.IsLongActionCurrent);
	};
	baseEditorsApi.prototype.isLongAction                    = function()
	{
		return this.isLongActionBase();
	};
	baseEditorsApi.prototype.incrementCounterLongAction      = function()
	{
		++this.IsLongActionCurrent;
	};
	baseEditorsApi.prototype.decrementCounterLongAction      = function()
	{
		this.IsLongActionCurrent--;
		if (this.IsLongActionCurrent < 0)
		{
			this.IsLongActionCurrent = 0;
		}

		if (!this.isLongActionBase())
		{
			var _length = this.LongActionCallbacks.length;
			for (var i = 0; i < _length; i++)
			{
				this.LongActionCallbacks[i](this.LongActionCallbacksParams[i]);
			}
			this.LongActionCallbacks.splice(0, _length);
			this.LongActionCallbacksParams.splice(0, _length);
		}
	};
	baseEditorsApi.prototype.checkLongActionCallback         = function(_callback, _param)
	{
		if (this.isLongActionBase())
		{
			this.LongActionCallbacks[this.LongActionCallbacks.length]             = _callback;
			this.LongActionCallbacksParams[this.LongActionCallbacksParams.length] = _param;
			return false;
		}
		else
		{
			return true;
		}
	};
	baseEditorsApi.prototype.canUndoRedoByRestrictions = function()
	{
		return (this.canEdit() || this.isRestrictionComments() || this.isRestrictionForms());
	};
	/**
	 * Функция для загрузчика шрифтов (нужно ли грузить default шрифты). Для Excel всегда возвращаем false
	 * @returns {boolean}
	 */
	baseEditorsApi.prototype.IsNeedDefaultFonts = function()
	{
		var res = false;
		switch (this.editorId)
		{
			case c_oEditorId.Word:
				res = true;
				break;
			case c_oEditorId.Presentation:
				res = true;
				break;
		}
		return res;
	};
	baseEditorsApi.prototype.isShowShapeAdjustments = function()
	{
		return true;
	};
	baseEditorsApi.prototype.isShowTableAdjustments = function()
	{
		return true;
	};
	baseEditorsApi.prototype.isShowEquationTrack = function()
	{
		return true;
	};
	baseEditorsApi.prototype.onPrint                             = function()
	{
		this.sendEvent("asc_onPrint");
	};
	// Open
	baseEditorsApi.prototype.asc_LoadDocument                    = function(versionHistory, isRepeat)
	{
		// Меняем тип состояния (на открытие)
		this.advancedOptionsAction = AscCommon.c_oAscAdvancedOptionsAction.Open;
		var rData                  = null;
		if (!(this.DocInfo && this.DocInfo.get_OfflineApp()))
		{
			var locale = !window['NATIVE_EDITOR_ENJINE'] && this.asc_getLocale() || undefined;
			if (typeof locale === "string") {
				if (Asc.g_oLcidNameToIdMap) {
					locale = Asc.g_oLcidNameToIdMap[locale];
				} else {
					locale = undefined;
				}
			}
			let convertToOrigin = '';
			if (!!this.DocInfo.get_DirectUrl() && this["asc_isSupportFeature"]("ooxml")) {
				convertToOrigin = '.docx.xlsx.pptx';
			}

			rData = {
				"c"             : 'open',
				"id"            : this.documentId,
				"userid"        : this.documentUserId,
				"format"        : this.documentFormat,
				"url"           : this.documentUrl,
				"title"         : this.documentTitle,
				"lcid"          : locale,
				"nobase64"      : true,
				"convertToOrigin" : convertToOrigin
			};

			if (this.isUseNativeViewer)
			{
				rData["convertToOrigin"] += Asc.c_sNativeViewerFormats;
			}

			if (versionHistory)
			{
				rData["serverVersion"] = versionHistory.serverVersion;
                rData["closeonerror"] = versionHistory.isRequested;
				rData["tokenHistory"] = versionHistory.token;
				//чтобы результат пришел только этому соединению, а не всем кто в документе
				rData["userconnectionid"] = this.CoAuthoringApi.getUserConnectionId();
			}
		}
		if (versionHistory) {
			this.CoAuthoringApi.versionHistory(rData);
		} else {
			this.CoAuthoringApi.auth(this.getViewMode(), rData);
		}

		if (!isRepeat) {
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open);
		}

        if (this.DocInfo.get_Encrypted() && window["AscDesktopEditor"] && !window["AscDesktopEditor"]["IsLocalFile"](true))
        {
        	var t = this;
            window["AscDesktopEditor"]["OpenFileCrypt"](this.DocInfo.get_Title(), this.DocInfo.get_Url(), function () {t.openFileCryptCallback.apply(t, arguments);});
        }
	};
	baseEditorsApi.prototype._openChartOrLocalDocument           = function()
	{
		this._openEmptyDocument();
	};
	baseEditorsApi.prototype._openEmptyDocument           = function()
	{
		var file = new AscCommon.OpenFileResult();
		file.data = AscCommon.getEmpty();
		file.bSerFormat = AscCommon.checkStreamSignature(file.data, AscCommon.c_oSerFormat.Signature);
		this.onEndLoadFile(file);
	};
	baseEditorsApi.prototype._openDocumentEndCallback            = function()
	{
	};
	baseEditorsApi.prototype._openVersionHistoryEndCallback            = function()
	{
		this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open);
	};
	baseEditorsApi.prototype._openOnClient                       = function()
	{
	};
	baseEditorsApi.prototype._onOpenCommand                      = function(data)
	{
		var t = this;
		let perfStart = performance.now();
		AscCommon.openFileCommand(this.documentId, data, this.documentUrlChanges, this.documentTokenChanges, AscCommon.c_oSerFormat.Signature, function(error, result)
		{
			let perfEnd = performance.now();
			AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onDownloadFile", perfEnd - perfStart), t);
			let errorData;
			if (c_oAscError.ID.No === error) {
				let editorId = AscCommon.getEditorBySignature(result.data);
				//todo AscCommon.checkNativeViewerSignature(result.data);
				let isNativeViewerFormat = -1 !== Asc.c_sNativeViewerFormats.indexOf(t.documentFormat);
				if (!(!isNativeViewerFormat && t.editorId === editorId || (isNativeViewerFormat && editorId === null))) {
					error = c_oAscError.ID.ConvertationOpenFormat;
					switch(editorId) {
						case AscCommon.c_oEditorId.Word: errorData = 'docx';break;
						case AscCommon.c_oEditorId.Spreadsheet: errorData = 'xlsx';break;
						case AscCommon.c_oEditorId.Presentation: errorData = 'pptx';break;
						default: errorData = 'pdf';break;
					}
				}
			}
			if (c_oAscError.ID.No !== error)
			{
				var err = c_oAscError.ID.No !== error ? error : c_oAscError.ID.ConvertationOpenError;
				t.sendEvent("asc_onError",  err, c_oAscError.Level.Critical, errorData);
				return;
			}
			t.onEndLoadFile(result);
		});
		this._openOnClient();
	};
	baseEditorsApi.prototype.openFileCryptCallback               = function (stream)
	{
		if (!this.isLoadFullApi)
		{
			this.openFileCryptBinary = stream;
			return;
		}
		this.openFileCryptBinary = null;

		if (stream == null)
		{
			this.sendEvent("asc_onError", c_oAscError.ID.ConvertationOpenError, c_oAscError.Level.Critical);
			return;
		}

		var file = new AscCommon.OpenFileResult();
		file.bSerFormat = AscCommon.checkStreamSignature(stream, AscCommon.c_oSerFormat.Signature);
		file.data = stream;
		this.openDocument(file);
		this.sendEvent("asc_onDocumentPassword", ("" !== this.currentPassword));
	};
	baseEditorsApi.prototype._onNeedParams                       = function(data, opt_isPassword)
	{
	};
	baseEditorsApi.prototype.asyncServerIdEndLoaded              = function(openedAt)
	{
		this.setOpenedAt(openedAt);
		// С сервером соединились, возможно стоит подождать загрузку шрифтов
		this.ServerIdWaitComplete = true;
		this._openDocumentEndCallback();
	};
	baseEditorsApi.prototype.asyncFontStartLoaded                = function()
	{
		// здесь прокинуть евент о заморозке меню
		this.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);
	};
	baseEditorsApi.prototype.asyncImageStartLoaded               = function()
	{
		// здесь прокинуть евент о заморозке меню
	};
	baseEditorsApi.prototype.asyncImagesDocumentStartLoaded      = function()
	{
		// евент о заморозке не нужен... оно и так заморожено
		// просто нужно вывести информацию в статус бар (что началась загрузка картинок)
	};
	baseEditorsApi.prototype.onDocumentContentReady              = function()
	{
		var t = this;
		this.isDocumentLoadComplete = true;
		if (!window['IS_NATIVE_EDITOR']) {
			setInterval(function() {t._autoSave();}, 40);
		}
		this.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open);
		this.sendEvent('asc_onDocumentContentReady');

		let time = this.VersionHistory ? undefined : performance.now();//todo perfStart?
		AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("onDocumentContentReady", time, AscCommon.getMemoryInfo()), t);

		if (window.g_asc_plugins)
            window.g_asc_plugins.onPluginEvent("onDocumentContentReady");

        if (c_oEditorId.Spreadsheet === this.editorId) {
			this.onUpdateDocumentModified(this.isDocumentModified());
		}

		if (this.DocInfo)
			this["pluginMethod_SetProperties"](this.DocInfo.asc_getOptions());

		if (this.macros && !this.disableAutostartMacros && this.macros.isExistAuto())
			this.sendEvent("asc_onRunAutostartMacroses");

		if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["onDocumentContentReady"])
            window["AscDesktopEditor"]["onDocumentContentReady"]();

		// теперь на старте нельзя удалить бинарник для подбора - он может пригодиться в nativeViewer
		if (!this.disableRemoveFonts)
			delete window["g_fonts_selection_bin"];
	};
	// Save
	baseEditorsApi.prototype.processSavedFile                    = function(url, downloadType, filetype)
	{
		if (AscCommon.DownloadType.None !== downloadType)
		{
			this.sendEvent(downloadType, url, filetype);
		}
		else
		{
			AscCommon.getFile(url);
		}
	};
	baseEditorsApi.prototype.getDrawingObjects = function () {};
	baseEditorsApi.prototype.getDrawingDocument = function () {};
	baseEditorsApi.prototype.getLogicDocument = function () {};
	baseEditorsApi.prototype._createSmartArt = function () {};
	baseEditorsApi.prototype.asc_createSmartArt = function (nSmartArtType, oPlaceholderObject) {
		const oThis = this;
		AscCommon.g_oBinarySmartArts.checkLoadDrawing().then(function()
		{
			return AscCommon.g_oBinarySmartArts.checkLoadData(nSmartArtType);
		}).then(function()
		{
			AscCommon.History.Create_NewPoint(AscDFH.historydescription_Document_AddSmartArt);
			const oSmartArt = new AscFormat.SmartArt();
			oSmartArt.fillByPreset(nSmartArtType);
			const oController = oThis.getGraphicController();
			oThis._createSmartArt(oSmartArt, oPlaceholderObject);
			oController.clearTrackObjects();
			oController.clearPreTrackObjects();
			oController.updateOverlay();
			oController.changeCurrentState(new AscFormat.NullState(oController));
			oController.updateSelectionState();
			return oSmartArt;
		});
	};
	baseEditorsApi.prototype.forceSave = function()
	{
		return this.CoAuthoringApi.forceSave();
	};
	baseEditorsApi.prototype.saveFromChanges = function(data, timeout, callback) {
		var t = this;
		var fAfterSaveChanges = function() {
			t.forceSaveForm = null;
			if (!t.CoAuthoringApi.callPRC(data, timeout, callback)) {
				callback(false, undefined);
			}
		};
		if (this.asc_Save(true)) {
			this.forceSaveForm = fAfterSaveChanges;
		} else {
			fAfterSaveChanges();
		}
	};
	/**
	 * This callback is displayed as a global member.
	 * @callback saveRelativeFromChangesCallback
	 * @param {boolean} timeout
	 * @param {Object} data description
	 * @param {number} data.errorCode from AscCommon.c_oAscServerCommandErrors {NoError, NotModified, Token, UnknownError}
	 * @param {boolean} data.inProgress
	 * @param {string} data.url
	 */
	/**
	 * Сохранение документа с указанием относительного пути
	 * @param docId external document id
	 * @param token from intergation
	 * @param timeout in ms
	 * @param callback {saveRelativeFromChangesCallback}
	 */
	baseEditorsApi.prototype.saveRelativeFromChanges = function(docId, token, timeout, callback) {
		if (!this.CoAuthoringApi.callPRC({'type': 'saveRelativeFromChanges', 'docId': docId, 'token': token}, timeout, callback)) {
			callback(true, undefined);
		}
	};

	baseEditorsApi.prototype.asc_setIsForceSaveOnUserSave = function(val)
	{
		this.isForceSaveOnUserSave = val;
	};
	baseEditorsApi.prototype._onUpdateDocumentCanSave = function () {
	};
	baseEditorsApi.prototype._onUpdateDocumentCanUndoRedo = function () {
	};
	baseEditorsApi.prototype._saveCheck = function () {
		return false;
	};
	// Переопределяется во всех редакторах
	baseEditorsApi.prototype._haveOtherChanges = function () {
		return false;
	};
	baseEditorsApi.prototype._onSaveCallback = function (e) {
		var t = this;
		var nState;
		if (false == e["saveLock"]) {
			if (this.isLongAction()) {
				// Мы не можем в этот момент сохранять, т.к. попали в ситуацию, когда мы залочили сохранение и успели нажать вставку до ответа
				// Нужно снять lock с сохранения
				this.CoAuthoringApi.onUnSaveLock = function () {
					if (t.isForceSaveOnUserSave && t.IsUserSave) {
						t.forceSaveButtonContinue = t.forceSave();
					}
					if (t.forceSaveForm) {
						t.forceSaveForm();
					}

					t.canSave = true;
					t.IsUserSave = false;
					t.lastSaveTime = null;

					if (t.canUnlockDocument) {
						t._unlockDocument();
					}
				};
				this.CoAuthoringApi.unSaveLock();
				return;
			}

			this.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.Save);

			this.canUnlockDocument2 = this.canUnlockDocument;
			if (this.canUnlockDocument && this.canStartCoAuthoring) {
				this.CoAuthoringApi.onStartCoAuthoring(true);
			}
			this.canStartCoAuthoring = false;
			this.canUnlockDocument = false;

			this._onSaveCallbackInner();
		} else {
			nState = this.CoAuthoringApi.get_state();
			if (AscCommon.ConnectionState.ClosedCoAuth === nState || AscCommon.ConnectionState.ClosedAll === nState) {
				// Отключаемся от сохранения, соединение потеряно
				this.IsUserSave = false;
				this.canSave = true;
			} else {
				// Если автосохранение, то не будем ждать ответа, а просто перезапустим таймер на немного
				if (!this.IsUserSave) {
					this.canSave = true;
					if (this.canUnlockDocument) {
						this._unlockDocument();
					}
					return;
				}

				setTimeout(function() {
					t.CoAuthoringApi.askSaveChanges(function(event) {
						t._onSaveCallback(event);
					});
				}, 1000);
			}
		}
	};
	// Функция сохранения. Переопределяется во всех редакторах
	baseEditorsApi.prototype._onSaveCallbackInner = function () {
	};
	baseEditorsApi.prototype._autoSave = function () {
		if (this.canSave && (!this.isViewMode || this.isLiveViewer()) && (this.canUnlockDocument || 0 !== this.autoSaveGap)) {
			if (this.canUnlockDocument) {
				this.lastSaveTime = new Date();
				// Check edit mode after unlock document http://bugzilla.onlyoffice.com/show_bug.cgi?id=35971
				// Close cell edit without errors (isIdle = true)
				this.asc_Save(true, true);
			} else {
				this._autoSaveInner();
			}
		}
	};
	// Функция автосохранения. Переопределяется во всех редакторах
	baseEditorsApi.prototype._autoSaveInner = function () {
	};
	baseEditorsApi.prototype._prepareSave = function (isIdle) {
		return true;
	};
	// Unlock document when start co-authoring
	baseEditorsApi.prototype._unlockDocument = function (isWaitAuth) {
		if (isWaitAuth && this.isDocumentLoadComplete && !this.canSave) {
			var errorMsg = 'Error: connection state changed waitAuth' +
				';this.canSave:' + this.canSave;
			AscCommon.sendClientLog("error", "changesError: " + errorMsg, this);
		}
		if (this.isDocumentLoadComplete) {
			// Document is load
			this.canUnlockDocument = true;
			this.canStartCoAuthoring = true;
			if (this.canSave) {
				// We can only unlock with delete index
				this.CoAuthoringApi.unLockDocument(false, true, AscCommon.History.GetDeleteIndex());
				this.startCollaborationEditing();
				AscCommon.History.RemovePointsByDeleteIndex();
				this._onUpdateDocumentCanSave();
				this._onUpdateDocumentCanUndoRedo();
				this.canStartCoAuthoring = false;
				this.canUnlockDocument = false;
			} else {
				// ToDo !!!!
			}
		} else {
			// Когда документ еще не загружен, нужно отпустить lock (при быстром открытии 2-мя пользователями)
			this.startCollaborationEditing();
			this.CoAuthoringApi.unLockDocument(false, true);
		}
	};
	// Выставление интервала автосохранения (0 - означает, что автосохранения нет)
	baseEditorsApi.prototype.asc_setAutoSaveGap                  = function(autoSaveGap)
	{
		if (typeof autoSaveGap === "number")
		{
			this.autoSaveGap = autoSaveGap * 1000; // Нам выставляют в секундах
		}
	};
	baseEditorsApi.prototype.checkChangesSize = function() {
		if (0 === this.maxChangesSize) {
			return;
		}
		let api = this;
		let localSize = AscCommon.History.GetLocalChangesSize ? AscCommon.History.GetLocalChangesSize() : 0;
		let serverSize = api.CoAuthoringApi.get_serverChangesSize();
		if (localSize + serverSize > api.maxChangesSize) {
			api.asc_stopSaving();
			api.sendEvent("asc_onConfirmAction", Asc.c_oAscConfirm.ConfirmMaxChangesSize,
				function(can) {
					if (can) {
						api.asc_Undo ? api.asc_Undo() : api.Undo();
					} else {
						api.setViewModeDisconnect(true);
					}
					api.asc_continueSaving();
				});
		}
	};
	// send chart message
	baseEditorsApi.prototype.asc_coAuthoringChatSendMessage      = function(message)
	{
		this.CoAuthoringApi.sendMessage(message);
	};
	// get chart messages
	baseEditorsApi.prototype.asc_coAuthoringChatGetMessages      = function()
	{
		this.CoAuthoringApi.getMessages();
	};
	// get users, возвращается массив users
	baseEditorsApi.prototype.asc_coAuthoringGetUsers             = function()
	{
		this.CoAuthoringApi.getUsers();
	};
	// get permissions
	baseEditorsApi.prototype.asc_getEditorPermissions            = function()
	{
		AscCommon.sendClientLog("debug", AscCommon.getClientInfoString("getEditorPermissions", performance.now()), this);
		this._coAuthoringInit();
	};
	baseEditorsApi.prototype.getConvertedXLSXFileFromUrl  = function (sUrl, sFileType, sToken, nOutputFormat, fCallback) {
		if (this.canEdit()) {
			const oDocument = {url: sUrl, format: sFileType, token: sToken};
			this.insertDocumentUrlsData = {
				imageMap: null, documents: [oDocument], convertCallback: function (_api, url) {
					_api.insertDocumentUrlsData.imageMap = url;
					if (url['output.xlsx']) {
						fCallback(url['output.xlsx']);
					} else if (url['output.xlst']) {
						fCallback(url['output.xlst']);
					} else {
						fCallback(null);
					}
					_api.endInsertDocumentUrls();
				}, endCallback: function (_api) {
					fCallback(null);
				}
			};

			const oOptions = new Asc.asc_CDownloadOptions(nOutputFormat);
			oOptions.isNaturalDownload = true;
			oOptions.isGetTextFromUrl = true;
			this.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, oOptions);
		}
	};
	baseEditorsApi.prototype._onEndPermissions                   = function()
	{
		if (this.isOnLoadLicense) {
			var oResult = new AscCommon.asc_CAscEditorPermissions();
			if (null !== this.licenseResult) {
				var type = this.licenseResult['type'];
				oResult.setLicenseType(type);
				oResult.setCanBranding(this.licenseResult['branding']);
				oResult.setCustomization(this.licenseResult['customization']);
				oResult.setIsLight(this.licenseResult['light']);
				oResult.setLicenseMode(this.licenseResult['mode']);
				oResult.setRights(this.licenseResult['rights']);
				oResult.setBuildVersion(this.licenseResult['buildVersion']);
				oResult.setBuildNumber(this.licenseResult['buildNumber']);
				oResult.setLiveViewerSupport(this.licenseResult['liveViewerSupport']);

				if (undefined !== this.licenseResult['protectionSupport']) {
					this.isProtectionSupport = this.licenseResult['protectionSupport'];
				}
				if (undefined !== this.licenseResult['isAnonymousSupport']) {
					this.isAnonymousSupport = this.licenseResult['isAnonymousSupport'];
				}
			}
			this.sendEvent('asc_onGetEditorPermissions', oResult);
		}
	};
	// GoTo
	baseEditorsApi.prototype.goTo                                = function(options)
	{
		options = options || (this.DocInfo && this.DocInfo.asc_getOptions());
		options = options && options['action'];
		if (!options) {
			return;
		}
		switch (options['type']) {
			case 'bookmark':
			case 'internallink':
				this._goToBookmark(options['data']);
				break;
			case 'comment':
				this._goToComment(options['data']);
				break;
		}
	};
	baseEditorsApi.prototype._goToComment                        = function(data)
	{
	};
	baseEditorsApi.prototype._goToBookmark                       = function(data)
	{
	};
	// CoAuthoring
	baseEditorsApi.prototype._coAuthoringInit                    = function()
	{
		var t = this;
		//Если User не задан, отключаем коавторинг.
		if (null == this.User || null == this.User.asc_getId())
		{
			this.User = new AscCommon.asc_CUser();
			this.User.setId("Unknown");
			this.User.setUserName("Unknown");
		}
		//в обычном серверном режиме портим ссылку, потому что CoAuthoring теперь имеет встроенный адрес
		//todo надо использовать проверку get_OfflineApp
		if (!(window['NATIVE_EDITOR_ENJINE'] || (this.DocInfo && this.DocInfo.get_OfflineApp())) || window['IS_NATIVE_EDITOR'])
		{
			this.CoAuthoringApi.set_url(null);
		}

		this.CoAuthoringApi.onMessage                 = function(e, clear)
		{
			t.sendEvent('asc_onCoAuthoringChatReceiveMessage', e, clear);
		};
		this.CoAuthoringApi.onServerVersion = function (buildVersion, buildNumber) {
			if (null !== t.disconnectRestrictions) {
				t.sync_EndAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.Disconnect);
				t.asc_setRestriction(t.disconnectRestrictions);
				t.disconnectRestrictions = null;
			}

			t.sendEvent('asc_onServerVersion', buildVersion, buildNumber);
		};
		this.CoAuthoringApi.onAuthParticipantsChanged = function(users, userId)
		{
			t.sendEvent("asc_onAuthParticipantsChanged", users, userId);
		};
		this.CoAuthoringApi.onParticipantsChanged     = function(users)
		{
			t.sendEvent("asc_onParticipantsChanged", users);
		};
		this.CoAuthoringApi.onSpellCheckInit          = function(e)
		{
			t.SpellCheckUrl = e;
			t._coSpellCheckInit();
		};
		this.CoAuthoringApi.onSetIndexUser            = function(e)
		{
			AscCommon.g_oIdCounter.Set_UserId('' + e);
		};
		this.CoAuthoringApi.onFirstLoadChangesEnd     = function(openedAt)
		{
			t.asyncServerIdEndLoaded(openedAt);
		};
		this.CoAuthoringApi.onFirstConnect            = function()
		{
			if (!t.isOnLoadLicense) {
				t._onEndPermissions();
			} else {
				if (t.CoAuthoringApi.get_isAuth()) {
					t.CoAuthoringApi.auth(t.getViewMode(), undefined, t.isIdle());
				} else {
					//первый запрос или ответ не дошел надо повторить открытие
					t.asc_LoadDocument(undefined, true);
				}
			}
		};
		this.CoAuthoringApi.onLicense                 = function(res)
		{
			t.licenseResult   = res;
			t.isOnLoadLicense = true;
			t._onEndPermissions();
		};
		this.CoAuthoringApi.onLicenseChanged          = function(res)
		{
			if (res['settings']) {
				if (res['settings']['maxChangesSize']) {
					t.maxChangesSize = res['settings']['maxChangesSize'];
				}
				if (res['settings']['binaryChanges']) {
					t.binaryChanges = res['settings']['binaryChanges'];
					t.CoAuthoringApi.setBinaryChanges(t.binaryChanges);
				}
				if (res['settings']['limits_image_size']) {
					AscCommon.c_oAscImageUploadProp.MaxFileSize = res['settings']['limits_image_size']
				}
				if (res['settings']['limits_image_types_upload']) {
					AscCommon.c_oAscImageUploadProp.SupportedFormats = res['settings']['limits_image_types_upload'].split(";")
				}
			}
			let licenseType = res['licenseType'];
			if (t.licenseResult) {
				t.licenseResult['type'] = licenseType;
			}
			t.isOnLoadLicense = true;
			var oResult = new AscCommon.asc_CAscEditorPermissions();
			oResult.setLicenseType(licenseType);
			t.sendEvent('asc_onLicenseChanged', oResult);
		};
		this.CoAuthoringApi.onWarning                 = function(code)
		{
			t.sendEvent('asc_onError', code || c_oAscError.ID.Warning, c_oAscError.Level.NoCritical);
		};
		this.CoAuthoringApi.onMeta                    = function(data)
		{
			var newDocumentTitle = data["title"];
			if (newDocumentTitle) {
				t.documentTitle = newDocumentTitle;
				if (t.DocInfo) {
					t.DocInfo.asc_putTitle(newDocumentTitle);
				}
			}
			t.sendEvent('asc_onMeta', data);
		};
		this.CoAuthoringApi.onSession = function(data) {
			var code = data["code"];
			var reason = data["reason"];
			var interval = data["interval"];
			var extendSession = true;
			if (c_oCloseCode.sessionIdle == code) {
				var idleTime = t.isIdle();
				if (idleTime > interval) {
					extendSession = false;
				} else {
					t.CoAuthoringApi.extendSession(idleTime);
				}
			} else if (c_oCloseCode.sessionAbsolute == code) {
				extendSession = false;
			}
			if (!extendSession) {
				if (t.asc_Save(false, true)) {
					//enter view mode because save async
					t.setViewModeDisconnect(AscCommon.getEnableDownloadByCloseCode(code));
					t.disconnectOnSave = {code: code, reason: reason};
				} else {
					t.CoAuthoringApi.sendClientLog('debug', 'disconnect code:' + code + ';reason:' + reason);
					t.CoAuthoringApi.disconnect(code, reason);
				}
			}
		};
        this.CoAuthoringApi.onForceSave = function(data) {
            if (AscCommon.c_oAscForceSaveTypes.Button === data.type) {
                if (data.start) {
                    if (null === t.forceSaveButtonTimeout && !t.forceSaveButtonContinue) {
                        t.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.ForceSaveButton);
                    } else {
                        clearInterval(t.forceSaveButtonTimeout);
                    }
                    t.forceSaveButtonTimeout = setTimeout(function() {
                        t.forceSaveButtonTimeout = null;
                        if (t.forceSaveButtonContinue) {
                            t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.Save);
                        } else {
                            t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.ForceSaveButton);
                        }
                        t.forceSaveButtonContinue = false;
                        t.sendEvent('asc_onError', Asc.c_oAscError.ID.ForceSaveButton, c_oAscError.Level.NoCritical);
                    }, Asc.c_nMaxConversionTime);
                } else if (data.refuse) {
                    if (t.forceSaveButtonContinue) {
                        t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.Save);
                    }
                    t.forceSaveButtonContinue = false;
                } else {
                    if (null !== t.forceSaveButtonTimeout) {
                        clearInterval(t.forceSaveButtonTimeout);
                        t.forceSaveButtonTimeout = null;
                        if (t.forceSaveButtonContinue) {
                            t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.Save);
                        } else {
                            t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.ForceSaveButton);
                        }
                        t.forceSaveButtonContinue = false;
                        if (!data.success) {
                            t.sendEvent('asc_onError', Asc.c_oAscError.ID.ForceSaveButton, c_oAscError.Level.NoCritical);
                        }
                    }
                }
            } else {
                if (AscCommon.CollaborativeEditing.Is_Fast() || null !== t.forceSaveTimeoutTimeout) {
                    if (data.start) {
                        if (null === t.forceSaveTimeoutTimeout) {
                            t.sync_StartAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.ForceSaveTimeout);
                        } else {
                            clearInterval(t.forceSaveTimeoutTimeout);
                        }
                        t.forceSaveTimeoutTimeout = setTimeout(function() {
                            t.forceSaveTimeoutTimeout = null;
                            t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.ForceSaveTimeout);
                            t.sendEvent('asc_onError', Asc.c_oAscError.ID.ForceSaveTimeout, c_oAscError.Level.NoCritical);
                        }, Asc.c_nMaxConversionTime);
                    } else {
                        if (null !== t.forceSaveTimeoutTimeout) {
                            clearInterval(t.forceSaveTimeoutTimeout);
                            t.forceSaveTimeoutTimeout = null;
                            t.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.ForceSaveTimeout);
                            if (!data.success) {
                                t.sendEvent('asc_onError', Asc.c_oAscError.ID.ForceSaveTimeout, c_oAscError.Level.NoCritical);
                            }
                        }
                    }
                }
            }
        };
		this.CoAuthoringApi.onExpiredToken = function(data) {
			t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open);
			if (t.VersionHistory && t.VersionHistory.isRequested) {
				var error = AscCommon.getDisconnectErrorCode(t.isDocumentLoadComplete, data["code"]);
				var level = t.isDocumentLoadComplete ? Asc.c_oAscError.Level.NoCritical : Asc.c_oAscError.Level.Critical;
				if (Asc.c_oAscError.ID.UpdateVersion === error && !t.isDocumentModified()) {
					t.sendEvent("asc_onDocumentUpdateVersion", function() {});
				} else {
					t.sendEvent('asc_onError', error, level);
				}
			} else {
				t.VersionHistory = null;
				t.sendEvent('asc_onExpiredToken');
			}
		};
		this.CoAuthoringApi.onHasForgotten = function() {
			//todo very bad way, need rewrite
			var isDocumentCanSaveOld = t.isDocumentCanSave;
			var canSaveOld = t.canSave;
			t.isDocumentCanSave = true;
			t.canSave = false;
			t.sendEvent("asc_onDocumentModifiedChanged");
			t.isDocumentCanSave = isDocumentCanSaveOld;
			t.canSave = canSaveOld;
			t.sendEvent("asc_onDocumentModifiedChanged");
		};
		/**
		 * Event об отсоединении от сервера
		 * @param {jQuery} e  event об отсоединении с причиной
		 * @param {opt_closeCode: AscCommon.c_oCloseCode.drop} opt_closeCode
		 */
		this.CoAuthoringApi.onDisconnect = function(e, opt_closeCode)
		{
			if (AscCommon.ConnectionState.None === t.CoAuthoringApi.get_state())
			{
				t.asyncServerIdEndLoaded();
			}
			if (null != opt_closeCode) {
				if (null !== t.disconnectRestrictions) {
					t.sync_EndAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.Disconnect);
					t.asc_setRestriction(t.disconnectRestrictions);
					t.disconnectRestrictions = null;
				}
				var error = AscCommon.getDisconnectErrorCode(t.isDocumentLoadComplete, opt_closeCode);
				var level = t.isDocumentLoadComplete ? Asc.c_oAscError.Level.NoCritical : Asc.c_oAscError.Level.Critical;
				t.setViewModeDisconnect(AscCommon.getEnableDownloadByCloseCode(opt_closeCode));
				if (Asc.c_oAscError.ID.UpdateVersion === error && !t.isDocumentModified()) {
					t.sendEvent("asc_onDocumentUpdateVersion", function() {});
				} else {
					t.sendEvent('asc_onError', error, level);
				}
			} else if (null === t.disconnectRestrictions){
				t.disconnectRestrictions = t.restrictions;
				t.sync_StartAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.Disconnect);
				t.asc_setRestriction(Asc.c_oAscRestrictionType.View);
			}
		};
		this.CoAuthoringApi.onDocumentOpen = function (inputWrap) {
			if (t.isOpenedChartFrame) {
				const oSentInformation = {
					"type": c_oAscFrameDataType.GetLoadedImages,
					"information": {
						"inputWrap": inputWrap
					}
				};
				t.sendFromGeneralToFrameEditor(oSentInformation);
			}
			if (AscCommon.EncryptionWorker.isNeedCrypt())
			{
                if (t.fCurCallback) {
                	t.fCurCallback(inputWrap ? inputWrap["data"] : undefined);
                	t.fCurCallback = null;
                }
				return;
			}
			if (inputWrap["data"]) {
				var input = inputWrap["data"];
				switch (input["type"]) {
					case 'reopen':
					case 'open':
						switch (input["status"]) {
							case "updateversion":
							case "ok":
								//call setOpenedAt twice in case of waitAuth
								t.setOpenedAt(input["openedAt"]);
								var urls = input["data"];
								AscCommon.g_oDocumentUrls.init(urls);
								var documentUrl = urls['Editor.bin'] || urls['origin.' + t.documentFormat];
								if (t["asc_isSupportFeature"]("ooxml") && !documentUrl) {
									documentUrl = urls['origin.docx'] || urls['origin.xlsx'] || urls['origin.pptx'];
								}
								if (t.isUseNativeViewer && !documentUrl)
									documentUrl = urls['origin.pdf'] || urls['origin.xps'] || urls['origin.oxps'] || urls['origin.djvu'];
								if (null != documentUrl) {
									if ('ok' === input["status"] || t.getViewMode()) {
										t._onOpenCommand(documentUrl);
									} else {
										t.sendEvent("asc_onDocumentUpdateVersion", function () {
											if (t.isCoAuthoringEnable) {
												t.asc_coAuthoringDisconnect();
											}
											t._onOpenCommand(documentUrl);
										})
									}
								} else {
									t.sendEvent("asc_onError", c_oAscError.ID.ConvertationOpenError,
										c_oAscError.Level.Critical);
								}
								break;
							case "needparams":
								t._onNeedParams(input["data"]);
								break;
							case "needpassword":
								t._onNeedParams(null, true);
								break;
							case "err":
								t.sendEvent("asc_onError",
									AscCommon.mapAscServerErrorToAscError(parseInt(input["data"]),
										Asc.c_oAscError.ID.ConvertationOpenError), c_oAscError.Level.Critical);
								break;
						}
						break;
					default:
						if (t.fCurCallback) {
							t.fCurCallback(input);
							t.fCurCallback = null;
						} else {
							t.sendEvent("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
						}
						break;
				}
			}
		};
		this.CoAuthoringApi.onStartCoAuthoring = function (isStartEvent, isWaitAuth) {
			if (t.isViewMode) {
				return;
			}
			// На старте не нужно ничего делать
			if (isStartEvent) {
				t.startCollaborationEditing();
			} else {
				t._unlockDocument(isWaitAuth);
			}
		};
		this.CoAuthoringApi.onEndCoAuthoring = function (isStartEvent) {
			if (t.canUnlockDocument) {
				t.canStartCoAuthoring = false;
			} else {
				t.endCollaborationEditing();
			}
		};

		this._coAuthoringInitEnd();
		this.CoAuthoringApi.init(this.User, this.documentId, this.documentCallbackUrl, 'fghhfgsjdgfjs', this.editorId, this.documentFormatSave, this.DocInfo);
	};
	baseEditorsApi.prototype._coAuthoringInitEnd                 = function()
	{
	};
	baseEditorsApi.prototype.startCollaborationEditing           = function()
	{
	};
	baseEditorsApi.prototype.endCollaborationEditing             = function()
	{
	};
	baseEditorsApi.prototype._coAuthoringCheckEndOpenDocument    = function(f)
	{
		if (this.isPreOpenLocks)
		{
			var context = this.CoAuthoringApi;
			var args = Array.prototype.slice.call(arguments, 1);

			// Пока документ еще не загружен, будем сохранять функцию и аргументы
			this.arrPreOpenLocksObjects.push(function()
			{
				f.apply(context, args);
			});
			return true;
		}
		return false;
	};
	baseEditorsApi.prototype._applyPreOpenLocks                  = function()
	{
		this.isPreOpenLocks = false;
		// Применяем все lock-и (ToDo возможно стоит пересмотреть вообще Lock-и)
		for (var i = 0; i < this.arrPreOpenLocksObjects.length; ++i)
		{
			this.arrPreOpenLocksObjects[i]();
		}
		this.arrPreOpenLocksObjects = [];
	};
	// server disconnect
	baseEditorsApi.prototype.asc_coAuthoringDisconnect           = function()
	{
		this.CoAuthoringApi.disconnect();
		this.isCoAuthoringEnable = false;

		// Выставляем view-режим
		this.asc_setViewMode(true);
	};
	baseEditorsApi.prototype.asc_stopSaving                      = function()
	{
		this.incrementCounterLongAction();
	};
	baseEditorsApi.prototype.asc_continueSaving                  = function()
	{
		this.decrementCounterLongAction();
	};
	// SpellCheck
	baseEditorsApi.prototype.asc_SpellCheckDisconnect            = function()
	{
		if (!this.SpellCheckApi)
			return; // Error
		this.SpellCheckApi.disconnect();
		this.isSpellCheckEnable = false;
		this._spellCheckDisconnect();
	};
	baseEditorsApi.prototype._spellCheckRestart                  = function(word)
	{
	};
	baseEditorsApi.prototype._spellCheckDisconnect               = function()
	{
	};
	baseEditorsApi.prototype._coSpellCheckInit                   = function()
	{
		var t = this;

		if (!this.SpellCheckApi)
		{
			return; // Error
		}

		if (window["AscDesktopEditor"]) {

            window["asc_nativeOnSpellCheck"] = function(response) {
                var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
                if (_editor.SpellCheckApi) {
                	// поверяем на сообщение о полной очистке очереди задач для текущего
                	if ("clear" === response) {
						_editor.SpellCheckApi.isRestart = false;
						return;
					}
					if (_editor.SpellCheckApi.isRestart === true)
						return;
					_editor.SpellCheckApi.onSpellCheck(response);
				}
            };

			this.SpellCheckApi.spellCheck = function (spellData) {
				window["AscDesktopEditor"]["SpellCheck"](JSON.stringify(spellData));
			};
			this.SpellCheckApi.disconnect = function () {
			};
			this.SpellCheckApi.restart = function() {
				this.isRestart = true;
				window["AscDesktopEditor"]["SpellCheck"]("clear");
			};

			if (window["AscDesktopEditor"]["IsLocalFile"] && !window["AscDesktopEditor"]["IsLocalFile"]())
			{
				let langs = AscCommon.spellcheckGetLanguages();
				let langs_array = [];
				for (let item in langs)
				{
					if (!langs.hasOwnProperty(item))
						continue;
					langs_array.push(item);
				}

				this.sendEvent('asc_onSpellCheckInit', langs_array);
			}
		} else {
			if (!this.SpellCheckUrl && !window['NATIVE_EDITOR_ENJINE']) {
				this.SpellCheckApi = {};
				this.SpellCheckApi.log = false;
				this.SpellCheckApi.worker = new CSpellchecker({
					enginePath: "../../../../sdkjs/common/spell/spell",
					dictionariesPath: "./../../../../dictionaries"
				});
				this.SpellCheckApi.worker.restartCallback = function() {
					t.asc_restartCheckSpelling();
				};
				this.SpellCheckApi.checkDictionary = function (lang) {
					if (this.log) console.log("checkDictionary: " + lang + ": " + this.worker.checkDictionary(lang));
					return this.worker.checkDictionary(lang);
				};
				this.SpellCheckApi.spellCheck = function (spellData) {
					if (this.log) {
						console.log("spellCheck:");
						console.log(spellData);
					}
					this.worker.command(spellData);
				};
				this.SpellCheckApi.worker.oncommand = function (spellData) {
					if (t.SpellCheckApi.log) {
						console.log("onSpellCheck:");
						console.log(spellData);
					}
					t.SpellCheck_CallBack(spellData);
				};
				this.SpellCheckApi.disconnect = function ()
				{
				};
				this.SpellCheckApi.restart = function() {
					this.worker.restart();
				};

				this.sendEvent('asc_onSpellCheckInit', this.SpellCheckApi.worker.getLanguages());
				return;
			}
			
			// Deprecated old scheme with server
			if (this.SpellCheckUrl && this.isSpellCheckEnable) {
				this.SpellCheckApi.set_url(this.SpellCheckUrl);
			}
		}

		this.SpellCheckApi.onInit = function (e) {
			t.sendEvent('asc_onSpellCheckInit', e);
		};
		this.SpellCheckApi.onSpellCheck = function (e) {
			t.SpellCheck_CallBack(e);
		};
		this.SpellCheckApi.init(this.documentId);
	};
    baseEditorsApi.prototype.asc_spellCheckAddToDictionary       = function(SpellCheckProperty)
    {
		var word = (typeof SpellCheckProperty === "string") ? SpellCheckProperty : SpellCheckProperty.Word;
		if (window["AscDesktopEditor"])
		{
			window["AscDesktopEditor"]["SpellCheck"]("{\"type\":\"add\",\"usrWords\":[\"" + word + "\"]}");

			this._spellCheckRestart(word);
		}
    };
    baseEditorsApi.prototype.asc_spellCheckClearDictionary       = function()
    {
    };
	baseEditorsApi.prototype.asc_restartCheckSpelling            = function()
	{
	};
	// Print Desktop
	baseEditorsApi.prototype._waitPrint                          = function (actionType, options)
	{
		return false;
	};
	baseEditorsApi.prototype._printDesktop                       = function ()
	{
	};
	// Download
	baseEditorsApi.prototype.endInsertDocumentUrls               = function()
	{
		if (this.insertDocumentUrlsData) {
			this.insertDocumentUrlsData.endCallback(this);
			this.insertDocumentUrlsData = null;
		}
	};
	baseEditorsApi.prototype._downloadAs                        = function ()
	{
	};
	baseEditorsApi.prototype.downloadAs                         = function (actionType, options)
	{
		var isCloudCrypto = !!(window["AscDesktopEditor"] && (0 < window["AscDesktopEditor"]["CryptoMode"]));
		if (isCloudCrypto)
		{
			window.isCloudCryptoDownloadAs = true;
		}
		if (this._waitPrint(actionType, options))
		{
			return;
		}

		if (actionType)
		{
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, actionType);
		}

		var downloadType;
		if (options.isDownloadEvent) {
			downloadType = options.oDocumentMailMerge ? DownloadType.MailMerge : (actionType === c_oAscAsyncAction.Print ? DownloadType.Print : DownloadType.Download);
		} else {
			downloadType = DownloadType.None;
		}

		var isNoBase64 = (typeof ArrayBuffer !== 'undefined') && !isCloudCrypto;
		var dataContainer = {data : null, part : null, index : 0, count : 0};
		var oAdditionalData = {};
		oAdditionalData["c"] = 'save';
		oAdditionalData["id"] = this.documentId;
		oAdditionalData["userid"] = this.documentUserId;
		oAdditionalData["tokenSession"] = this.CoAuthoringApi.get_jwt();
		oAdditionalData["outputformat"] = options.fileType;
		oAdditionalData["title"] = AscCommon.changeFileExtention(this.documentTitle, AscCommon.getExtentionByFormat(options.fileType), Asc.c_nMaxDownloadTitleLen);
		oAdditionalData["nobase64"] = isNoBase64;
		let locale = this.asc_getLocale() || undefined;
		if (typeof locale === "string") {
			locale = Asc.g_oLcidNameToIdMap[locale];
		}
		oAdditionalData["lcid"] = locale;
		if (DownloadType.Print === downloadType)
		{
			oAdditionalData["withoutPassword"] = true;
			oAdditionalData["inline"] = 1;
		}
		if (Asc.c_oAscFileType.JPG === options.fileType || Asc.c_oAscFileType.TIFF === options.fileType
			|| Asc.c_oAscFileType.TGA === options.fileType || Asc.c_oAscFileType.GIF === options.fileType
			|| Asc.c_oAscFileType.PNG === options.fileType || Asc.c_oAscFileType.EMF === options.fileType
			|| Asc.c_oAscFileType.WMF === options.fileType || Asc.c_oAscFileType.BMP === options.fileType
			|| Asc.c_oAscFileType.CR2 === options.fileType || Asc.c_oAscFileType.PCX === options.fileType
			|| Asc.c_oAscFileType.RAS === options.fileType || Asc.c_oAscFileType.PSD === options.fileType
			|| Asc.c_oAscFileType.ICO === options.fileType) {
			oAdditionalData["thumbnail"] = {
				"aspect": 2,
				"first": false
			}
			switch (options.fileType) {
				case Asc.c_oAscFileType.PNG:
					oAdditionalData["thumbnail"]["format"] = 4;
					break;
				case Asc.c_oAscFileType.GIF:
					oAdditionalData["thumbnail"]["format"] = 2;
					break;
				case Asc.c_oAscFileType.BMP:
					oAdditionalData["thumbnail"]["format"] = 1;
					break;
				default:
					oAdditionalData["thumbnail"]["format"] = 3;
					break;
			}
			oAdditionalData["outputformat"] = Asc.c_oAscFileType.IMG;
			oAdditionalData["title"] = AscCommon.changeFileExtention(this.documentTitle, "zip", Asc.c_nMaxDownloadTitleLen);

			let jsonparams = {};
			//todo convert from asc_CAdjustPrint
			jsonparams["spreadsheetLayout"] = {"ignorePrintArea": true, "scale": 100};
			jsonparams["locale"] = this.asc_getLocale();
			//todo move cmd from header to body and uncomment
			// jsonparams["translate"] = AscCommon.translateManager.mapTranslate;
			jsonparams["documentLayout"] = { "openedAt" : this.openedAt};
			oAdditionalData["jsonparams"] = JSON.stringify(jsonparams);
		}
		if (options.textParams && undefined !== options.textParams.asc_getAssociation()) {
			oAdditionalData["textParams"] = {"association": options.textParams.asc_getAssociation()};
		}

		if (this._downloadAs(actionType, options, oAdditionalData, dataContainer, downloadType))
		{
			return;
		}
		this._downloadAsUsingServer(actionType, options, oAdditionalData, dataContainer, downloadType);
	};
	baseEditorsApi.prototype._downloadAsUsingServer                        = function (actionType, options, oAdditionalData, dataContainer, downloadType)
	{
		var t = this;
		this.fCurCallback = null;
		if (!options.callback)
		{
			this.fCurCallback = function(input, status)
			{
				var error = 403 === status ? c_oAscError.ID.AccessDeny : c_oAscError.ID.Unknown;
				//input = {'type': command, 'status': 'err', 'data': -80};
				if (null != input && oAdditionalData["c"] === input['type'])
				{
					if ('ok' === input['status'])
					{
						var url = input['data'];
						if (url)
						{
							error = c_oAscError.ID.No;
							t.processSavedFile(url, downloadType, input["filetype"]);
						}
					}
					else
					{
						error = AscCommon.mapAscServerErrorToAscError(parseInt(input["data"]),
							(options && options.isGetTextFromUrl) ? AscCommon.c_oAscAdvancedOptionsAction.Open : AscCommon.c_oAscAdvancedOptionsAction.Save);
					}
				}
				if (c_oAscError.ID.No !== error)
				{
					t.endInsertDocumentUrls();
					t.sendEvent('asc_onError', options.errorDirect || error, c_oAscError.Level.NoCritical);
				}
				if (actionType)
				{
					t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, actionType);
				}
			};
		}
		AscCommon.saveWithParts(function(fCallback1, oAdditionalData1, dataContainer1) {
			AscCommon.sendCommand(t, fCallback1, oAdditionalData1, dataContainer1);
		}, this.fCurCallback, options.callback, oAdditionalData, dataContainer);
	};
	baseEditorsApi.prototype._downloadOriginalFile = function (directUrl, url, fileType, token, callback)
	{
		if (directUrl) {
			AscCommon.loadFileContent(directUrl, function(resp) {
				if (resp) {
					callback(AscCommon.initStreamFromResponse(resp));
				} else {
					callback(null);
				}
			}, "arraybuffer");
		} else {
			AscCommon.DownloadOriginalFile(this.documentId, url, "", token, function(){
				callback(null);
			}, callback);
		}
	};
	baseEditorsApi.prototype.asc_generateSmartArtPreviews = function(nTypeOfSection)
	{
		return this.smartArtPreviewManager.Begin(nTypeOfSection);
	};
	// Images & Charts & TextArts
	baseEditorsApi.prototype.asc_getChartPreviews                = function(chartType, arrId, bEmpty)
	{
		return this.chartPreviewManager.getChartPreviews(chartType, arrId, bEmpty);
	};
	baseEditorsApi.prototype.asc_generateChartPreviews                = function(chartType, arrId)
	{
		return this.chartPreviewManager.Begin(chartType, arrId);
	};
	baseEditorsApi.prototype.asc_getTextArtPreviews              = function()
	{
		return this.textArtPreviewManager.getWordArtStyles();
	};
	baseEditorsApi.prototype.asc_onOpenChartFrame                = function()
	{
		if(this.isMobileVersion){
			return;
		}
		this.isOpenedChartFrame = true;
	};
	baseEditorsApi.prototype.asc_onCloseChartFrame               = function()
	{
		this.isOpenedChartFrame = false;
	};
	baseEditorsApi.prototype.asc_setInterfaceDrawImagePlaceShape = function(elementId)
	{
		this.shapeElementId = elementId;
	};
	baseEditorsApi.prototype.asc_getPropertyEditorShapes         = function()
	{
		return [AscCommon.g_oAutoShapesGroups, AscCommon.g_oAutoShapesTypes];
	};
	baseEditorsApi.prototype.asc_getPropertyEditorTextArts       = function()
	{
		return this.textArtPreviewManager.getWordArtPreviews();
	};
	// Add image
	baseEditorsApi.prototype.AddImageUrl       = function(urls, imgProp, token, obj)
	{
		if (this.isLongAction()) {
			return;
		}
		var t = this;
		var toSendUrls = [];
		var toSendIndex = [];
		for (var i = 0; i < urls.length; ++i) {
			if (!AscCommon.g_oDocumentUrls.getLocal(urls[i])) {
				toSendIndex.push(i);
				toSendUrls.push(urls[i]);
			}
		}
		var callback = function(urls) {
			t._addImageUrl(urls, obj);
		};
		if (toSendUrls.length > 0) {
			AscCommon.sendImgUrls(this, toSendUrls, function(data) {
				if (data) {
					data.forEach(function(currentValue, index) {
						urls[toSendIndex[index]] = currentValue.url;
					});
					callback.call(t, urls);
				}
			}, undefined, token);
		} else {
			callback.call(this, urls);
		}
	};
	baseEditorsApi.prototype._addImageUrl                        = function()
	{
	};
	baseEditorsApi.prototype.asc_addImage                        = function(obj)
	{
		if (this.isEditOleMode)
		{
			this.oSaveObjectForAddImage = obj;
			this.sendFromFrameToGeneralEditor({
				"type": AscCommon.c_oAscFrameDataType.ShowImageDialogInFrame,
			});
			return;
		}
		var t = this;
        if (this.WordControl) // после показа диалога может не прийти mouseUp
        	this.WordControl.m_bIsMouseLock = false;
		AscCommon.ShowImageFileDialog(this.documentId, this.documentUserId, this.CoAuthoringApi.get_jwt(), function(error, files)
		{
			t._uploadCallback(error, files, obj);
		}, function(error)
		{
			if (c_oAscError.ID.No !== error)
			{
				t.sendEvent("asc_onError", error, c_oAscError.Level.NoCritical);
			}
			if (obj && obj.sendUrlsToFrameEditor && t.isOpenedChartFrame)
			{
				t.sendStartUploadImageActionToFrameEditor();
			}
			t.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
		});
	};
	baseEditorsApi.prototype._uploadCallback                     = function(error, files, obj)
	{
		var t = this;
		if (c_oAscError.ID.No !== error)
		{
			this.sendEvent("asc_onError", error, c_oAscError.Level.NoCritical);
		}
		else
		{
			if (obj && obj.sendUrlsToFrameEditor && t.isOpenedChartFrame)
			{
				this.sendStartUploadImageActionToFrameEditor();
			}
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
			AscCommon.UploadImageFiles(files, this.documentId, this.documentUserId, this.CoAuthoringApi.get_jwt(), function(error, urls)
			{
				if (c_oAscError.ID.No !== error)
				{
					t.sendEvent("asc_onError", error, c_oAscError.Level.NoCritical);
				}
				else
				{
					t._addImageUrl(urls, obj);
				}
				t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
			});
		}
	};

	//метод, который подменяет callback загрузки в каждом редакторе, TODO: переделать, сделать одинаково в о всех редакторах
	baseEditorsApi.prototype.asc_replaceLoadImageCallback = function(fCallback)
	{
	};

	baseEditorsApi.prototype.asc_loadLocalImageAndAction = function(sLocalImage, fCallback)
	{
		var _loadedUrl = this.ImageLoader.LoadImage(AscCommon.getFullImageSrc2(sLocalImage), 1);
		if (_loadedUrl != null)
		    fCallback(_loadedUrl);
        else
        	this.asc_replaceLoadImageCallback(fCallback);
	};

	baseEditorsApi.prototype.asc_checkImageUrlAndAction = function(sImageUrl, fCallback)
	{
		var oThis = this;
		this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
		var fCallback2 = function()
		{
			oThis.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.UploadImage);
			fCallback.apply(oThis, arguments);
		};
		var sLocalImage = AscCommon.g_oDocumentUrls.getImageLocal(sImageUrl);
		if (sLocalImage)
		{
			this.asc_loadLocalImageAndAction(sLocalImage, fCallback2);
			return;
		}

		AscCommon.sendImgUrls(oThis, [sImageUrl], function(data)
		{
			if (data[0] && data[0].path != null && data[0].url !== "error")
			{
				oThis.asc_loadLocalImageAndAction(AscCommon.g_oDocumentUrls.imagePath2Local(data[0].path), fCallback2);
			}
		});
	};

	baseEditorsApi.prototype.asc_addOleObject = function(oPluginData)
	{
		if(this.isViewMode || this.isPdfEditor())
		{
			return;
		}
		let oThis      = this;
		let sImgSrc    = oPluginData["imgSrc"];
		let nWidthPix  = oPluginData["widthPix"];
		let nHeightPix = oPluginData["heightPix"];
		let fWidthMM   = oPluginData["width"];
		let fHeightMM  = oPluginData["height"];
		let sData      = oPluginData["data"];
		let sGuid      = oPluginData["guid"];
		let bSelect    = (oPluginData["select"] === true || oPluginData["select"] === false) ? oPluginData["select"] : true;
		let bPlugin    = oPluginData["plugin"] === true && !!window.g_asc_plugins;
		if (typeof sImgSrc === "string" && sImgSrc.length > 0 && typeof sData === "string"
			&& typeof sGuid === "string" && sGuid.length > 0
			/*&& AscFormat.isRealNumber(nWidthPix) && AscFormat.isRealNumber(nHeightPix)*/
			&& AscFormat.isRealNumber(fWidthMM) && AscFormat.isRealNumber(fHeightMM)
		)
		{
			if(bPlugin)
			{
				window.g_asc_plugins.setPluginMethodReturnAsync();
			}
			this.asc_checkImageUrlAndAction(sImgSrc, function(oImage)
			{
				oThis.asc_addOleObjectAction(AscCommon.g_oDocumentUrls.getImageLocal(oImage.src), sData, sGuid, fWidthMM, fHeightMM, nWidthPix, nHeightPix, bSelect);
				if(bPlugin)
				{
					window.g_asc_plugins.onPluginMethodReturn();
				}
			});
		}

	};

	baseEditorsApi.prototype.asc_editOleObject = function(oPluginData)
	{
		if(this.isViewMode || this.isPdfEditor())
		{
			return;
		}
		var oThis      = this;
		var bResize    = oPluginData["resize"];
		var sImgSrc    = oPluginData["imgSrc"];
		var oOleObject = AscCommon.g_oTableId.Get_ById(oPluginData["objectId"]);
		var nWidthPix  = oPluginData["widthPix"];
		var nHeightPix = oPluginData["heightPix"];
		var sData      = oPluginData["data"];
		var fWidthMM   = oPluginData["width"];
		var fHeightMM  = oPluginData["height"];
		if (typeof sImgSrc === "string" && sImgSrc.length > 0 && typeof sData === "string"
			&& oOleObject /*&& AscFormat.isRealNumber(nWidthPix) && AscFormat.isRealNumber(nHeightPix)*/
			&& AscFormat.isRealNumber(fWidthMM) && AscFormat.isRealNumber(fHeightMM))
		{
            this.asc_checkImageUrlAndAction(sImgSrc, function(oImage)
			{
				oThis.asc_editOleObjectAction(oOleObject, AscCommon.g_oDocumentUrls.getImageLocal(oImage.src), sData, fWidthMM, fHeightMM, nWidthPix, nHeightPix);
			});
		}
	};

	baseEditorsApi.prototype.asc_addOleObjectAction = function(sLocalUrl, sData, sApplicationId, fWidth, fHeight, nWidthPix, nHeightPix, bSelect, arrImagesForAddToHistory)
	{
	};

	baseEditorsApi.prototype.asc_editOleObjectAction = function(oOleObject, sImageUrl, sData, fWidthMM, fHeightMM, nPixWidth, nPixHeight, arrImagesForAddToHistory)
	{
	};


	baseEditorsApi.prototype.asc_startEditCurrentOleObject = function(){

	};
	baseEditorsApi.prototype.asc_canEditCrop = function()
	{
	};

	baseEditorsApi.prototype.asc_startEditCrop = function()
	{
	};

	baseEditorsApi.prototype.asc_endEditCrop = function()
	{
	};

	baseEditorsApi.prototype.asc_cropFit = function()
	{
	};

	baseEditorsApi.prototype.asc_cropFill = function()
	{
	};

	
	baseEditorsApi.prototype.asc_setShapeNames = function(oShapeNames)
	{
		if(oShapeNames !== null && typeof oShapeNames === "object")
		{
			this.shapeNames = oShapeNames;
		}
	};

	baseEditorsApi.prototype.getShapeName = function(sPreset)
	{
		var sShapeName = this.shapeNames[sPreset];
		if(typeof sShapeName !== "string" || sShapeName.length === 0)
		{
			sShapeName = "Shape";
		}
		return sShapeName;
	};

	baseEditorsApi.prototype.getGraphicController = function() {};

	baseEditorsApi.prototype.asc_canEditTableOleObject = function(bReturnOle) {
		var oController = this.getGraphicController();
		if(oController) {
			return oController.canEditTableOleObject(bReturnOle);
		}
		return bReturnOle ? null : false;
	};

	baseEditorsApi.prototype.turnOffSpecialModes = function()
	{};

	//Remove All comments
	baseEditorsApi.prototype.asc_RemoveAllComments = function(isMine, isCurrent)
	{};

	// Version History
	baseEditorsApi.prototype.asc_showRevision   = function(newObj)
	{
		if (!newObj.docId) {
			return;
		}
		if (this.isCoAuthoringEnable) {
			this.asc_coAuthoringDisconnect();
		}

		if (this.SpellCheckApi && this.SpellCheckApi.restart /* старый спеллчек (серверный - не поддеживает этот метод */)
			this.SpellCheckApi.restart();

		var bUpdate = true;
		if (null === this.VersionHistory) {
			this.VersionHistory = new window["Asc"].asc_CVersionHistory(newObj);
		} else {
			bUpdate = this.VersionHistory.update(newObj);
		}
		// ToDo должно быть все общее
		if (bUpdate) {
			this.asc_CloseFile();

			this.DocInfo.put_Id(this.VersionHistory.docId);
			this.DocInfo.put_Url(this.VersionHistory.url);
			this.documentUrlChanges = this.VersionHistory.urlChanges;
			this.documentTokenChanges = this.VersionHistory.token;
			this.asc_setDocInfo(this.DocInfo);
			this.asc_LoadDocument(this.VersionHistory);
		} else if (this.VersionHistory.currentChangeId < newObj.currentChangeId) {
			var oApi = Asc.editor || editor;
			this.isApplyChangesOnVersionHistory = true;
			this.sync_StartAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open);
			// Нужно только добавить некоторые изменения
			AscCommon.CollaborativeEditing.Clear_CollaborativeMarks();
			oApi.VersionHistory.applyChanges(oApi);
			AscCommon.CollaborativeEditing.Apply_Changes();
		}
	};
	baseEditorsApi.prototype.asc_undoAllChanges = function()
	{
	};
	baseEditorsApi.prototype.asc_getAdvancedOptions = function () {
		var cp            = {
			'codepage'  : AscCommon.c_oAscCodePageUtf8,
			'encodings' : AscCommon.getEncodingParams()
		};
		return new AscCommon.asc_CAdvancedOptions(cp);
	};
	baseEditorsApi.prototype.asc_Print = function (options) {
		if (window["AscDesktopEditor"] && this._printDesktop(options)) {
			return;
		}
		if (this.isLongAction()) {
			return;
		}

		if (!options) {
			options = new Asc.asc_CDownloadOptions();
		}
		options.fileType = Asc.c_oAscFileType.PDF;
		options.isPdfPrint = true;
		this.downloadAs(c_oAscAsyncAction.Print, options);
	};
	baseEditorsApi.prototype.asc_Save = function (isAutoSave, isIdle) {
		var t = this;
		var res = false;
		if (this.canSave && this._saveCheck()) {
			this.IsUserSave = !isAutoSave;

			if (this.asc_isDocumentCanSave() || AscCommon.History.Have_Changes() || this._haveOtherChanges() ||
				this.canUnlockDocument || this.forceSaveUndoRequest) {
				if (this._prepareSave(isIdle)) {
					// Не даем пользователю сохранять, пока не закончится сохранение (если оно началось)
					this.canSave = false;
					this.CoAuthoringApi.askSaveChanges(function (e) {
						t._onSaveCallback(e);
					});
					res = true;
				}
			} else if (this.isForceSaveOnUserSave && this.IsUserSave) {
				this.forceSave();
			}
		}
		return res;
	};
	/**
	 * Эта функция возвращает true, если есть изменения или есть lock-и в документе
	 */
	baseEditorsApi.prototype.asc_isDocumentCanSave = function()
	{
		return this.isDocumentCanSave;
	};
	baseEditorsApi.prototype.asc_getCanUndo = function()
	{
		if (!this.canUndoRedoByRestrictions())
			return false;

		return AscCommon.History.Can_Undo();
	};
	baseEditorsApi.prototype.asc_getCanRedo = function()
	{
		if (!this.canUndoRedoByRestrictions())
			return false;

		return AscCommon.History.Can_Redo();
	};
	// Offline mode
	baseEditorsApi.prototype.asc_isOffline  = function()
	{
		return (window.location.protocol.indexOf("file") == 0) ? true : false;
	};
	baseEditorsApi.prototype.asc_getUrlType = function(url)
	{
		let res = AscCommon.getUrlType(url);
		//check bugs after modification: 59753, 59780
		if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsLocalFile"]() &&
			(res === AscCommon.c_oAscUrlType.Invalid || res === AscCommon.c_oAscUrlType.Http) &&
			!(AscCommon.rx_allowedProtocols.test(url) || /^(www.)|@/i.test(url)) )
		{
			res = AscCommon.c_oAscUrlType.Unsafe;
		}
		return res;
	};
	baseEditorsApi.prototype.asc_prepareUrl = function(url)
	{
		return AscCommon.prepareUrl(url);
	};

	baseEditorsApi.prototype.openDocument  = function(file)
	{
	};
	baseEditorsApi.prototype.openDocumentFromZip  = function()
	{
	};
	baseEditorsApi.prototype.onEndLoadDocInfo = function()
	{
		if (this.isLoadFullApi && this.DocInfo)
		{
			if (this.DocInfo.get_OfflineApp())
			{
				this._openChartOrLocalDocument();
			}
			this.onEndLoadFile(null);
		}
	};
	baseEditorsApi.prototype.onEndLoadFile = function(result)
	{
		if (result)
		{
			this.openResult = result;
		}
		if (this.isLoadFullApi && this.DocInfo && this.openResult && this._isLoadedModules())
		{
			this.openDocument(this.openResult);
			this.sendEvent("asc_onDocumentPassword", ("" !== this.currentPassword));
			this.openResult = null;
		}

	};
	baseEditorsApi.prototype._onEndLoadSdk = function()
	{
		AscCommon.g_oTableId.init();

		// init drag&drop
		var t = this;
		AscCommon.InitDragAndDrop(this.HtmlElement, function(error, files)
		{
			t._uploadCallback(error, files);
		});

		AscFonts.g_fontApplication.Init();

		this.FontLoader  = AscCommon.g_font_loader;
		this.ImageLoader = AscCommon.g_image_loader;
		this.FontLoader.put_Api(this);
		this.ImageLoader.put_Api(this);
		this.FontLoader.SetStandartFonts();

		this.chartPreviewManager   = new AscCommon.ChartPreviewManager();
		this.smartArtPreviewManager   = new AscCommon.SmartArtPreviewDrawer();
		this.textArtPreviewManager = new AscCommon.TextArtPreviewManager();

		AscFormat.initStyleManager();

		if (null !== this.tmpFocus)
		{
			this.asc_enableKeyEvents(this.tmpFocus);
		}

		this.pluginsManager     = Asc.createPluginsManager(this);

		this.macros = new AscCommon.CDocumentMacros();

		this._loadSdkImages();

		if(AscFonts.FontPickerByCharacter && this.documentTitle) {
			AscFonts.FontPickerByCharacter.getFontsByString(this.documentTitle);
		}
	};
	baseEditorsApi.prototype._loadSdkImages = function ()
	{
	};

	baseEditorsApi.prototype.sendStandartTextures = function()
	{
	    if (this.isSendStandartTextures)
	        return;

	    this.isSendStandartTextures = true;

		var _count = AscCommon.g_oUserTexturePresets.length;
		var arr    = new Array(_count);
		var arrToDownload = [];
		for (var i = 0; i < _count; ++i)
		{
			arr[i]       = new AscCommon.asc_CTexture();
			arr[i].Id    = i;
			arr[i].Image = AscCommon.g_oUserTexturePresets[i];
			arrToDownload.push(AscCommon.g_oUserTexturePresets[i]);
		}
		if(this.editorId === c_oEditorId.Word)
		{
			arrToDownload.push(AscCommon.g_sWordPlaceholderImage);
		}
		this.ImageLoader.LoadImagesWithCallback(arrToDownload, function () {

		}, 0, true);

		this.sendEvent('asc_onInitStandartTextures', arr);
	};

	baseEditorsApi.prototype.sendMathToMenu = function ()
	{
		if (this.MathMenuLoad)
			return;
		// GENERATE_IMAGES
		//var _MathPainter = new CMathPainter(this.m_oWordControl.m_oApi);
		//_MathPainter.StartLoad();
		//return;
		var _MathPainter = new AscFormat.CMathPainter(this);
		_MathPainter.Generate();
		this.MathMenuLoad = true;
	};

	baseEditorsApi.prototype.sendMathTypesToMenu         = function(_math)
	{
		this.sendEvent("asc_onMathTypes", _math);
	};

	baseEditorsApi.prototype.asyncFontEndLoaded_MathDraw = function(Obj)
	{
		this.sync_EndAction(c_oAscAsyncActionType.Information, c_oAscAsyncAction.LoadFont);
		Obj.Generate2();
	};

	baseEditorsApi.prototype.getCurrentColorScheme = function()
	{
		var oTheme = this.getCurrentTheme();
		return oTheme && oTheme.themeElements && oTheme.themeElements.clrScheme;
	};


	baseEditorsApi.prototype.asc_GetCurrentColorSchemeName = function()
	{
		var oClrScheme = this.getCurrentColorScheme();
		if(oClrScheme && typeof oClrScheme.name === "string")
		{
			return oClrScheme.name;
		}
		return "";
	};

	baseEditorsApi.prototype.asc_GetCurrentColorSchemeIndex = function()
	{
		var oTheme = this.getCurrentTheme();
		if(!oTheme)
		{
			return -1;
		}
		return this.getColorSchemes(oTheme).index;
	};

	baseEditorsApi.prototype.getCurrentTheme = function()
	{
		return null;
	};

	baseEditorsApi.prototype.getColorSchemes = function(theme)
	{
		var result = AscCommon.g_oUserColorScheme.slice();
		// theme colors
		var asc_color_scheme, _scheme, i;
		var aCustomSchemes = theme.getExtraAscColorSchemes();
		_scheme = theme.themeElements && theme.themeElements.clrScheme;
		var nIndex = -1;
		if(_scheme) {
			asc_color_scheme = AscCommon.getAscColorScheme(_scheme, theme);
			nIndex = AscCommon.getIndexColorSchemeInArray(result, asc_color_scheme);
			if(nIndex === -1) {
				var nIdxInCustom = AscCommon.getIndexColorSchemeInArray(aCustomSchemes, asc_color_scheme);
				if(nIdxInCustom === -1) {
					aCustomSchemes.push(asc_color_scheme);
				}
			}
			aCustomSchemes.sort(function (a, b) {
				if(a.name === "" || a.name === null) return -1;
				if(b.name === "" || b.name === null) return 1;
				if(a.name > b.name)
				{
					return 1;
				}
				if(a.name < b.name)
				{
					return -1;
				}
				return 0;
			});

			result = result.concat(aCustomSchemes);

			if(nIndex === -1) {
				nIndex = AscCommon.getIndexColorSchemeInArray(result, asc_color_scheme);
			}
		}
		return {schemes: result, index: nIndex};
	};

	baseEditorsApi.prototype.getColorSchemeByIdx = function(nIdx)
	{
		var scheme = AscCommon.getColorSchemeByIdx(nIdx);
		if(!scheme) {
			var oSchemes = this.getColorSchemes(this.getCurrentTheme());
			var oAscScheme = oSchemes.schemes[nIdx];
			scheme = oAscScheme && oAscScheme.scheme;
		}
		return scheme;
	};


	baseEditorsApi.prototype.sendColorThemes = function (theme)
	{
		this.sendEvent("asc_onSendThemeColorSchemes", this.getColorSchemes(theme).schemes);
	};


	baseEditorsApi.prototype.showVideoControl = function(sMediaName, extX, extY, transform)
	{
		if (!window["AscDesktopEditor"] || !window["AscDesktopEditor"]["MediaStart"])
			return;

		switch (this.editorId)
		{
			case c_oEditorId.Word:
			{
				break;
			}
			case c_oEditorId.Presentation:
			{
                var manager = this.WordControl.DemonstrationManager;
                if (!manager.Mode)
                {
                    var pos = this.WordControl.m_oDrawingDocument.ConvertCoordsToCursorWR(0, 0, this.WordControl.m_oLogicDocument.CurPage, null, true);
                    pos.X += this.WordControl.X;
                    pos.Y += this.WordControl.Y;

					var zoom = AscCommon.g_dKoef_mm_to_pix * this.WordControl.m_nZoomValue / 100;

                    if (!transform)
                        window["AscDesktopEditor"]["MediaStart"](sMediaName, pos.X >> 0, pos.Y >> 0, extX, extY, zoom);
                    else
                        window["AscDesktopEditor"]["MediaStart"](sMediaName, pos.X >> 0, pos.Y >> 0, extX, extY, zoom, transform.sx, transform.shy, transform.shx, transform.sy, transform.tx, transform.ty);
                }
                else
				{
					var transition = this.WordControl.DemonstrationManager.Transition;
                    if ((manager.SlideNum >= 0 && manager.SlideNum < manager.SlidesCount) && (!transition || !transition.IsPlaying()))
                    {
						var _x = (transition.Rect.x / AscCommon.AscBrowser.retinaPixelRatio) >> 0;
						var _y = (transition.Rect.y / AscCommon.AscBrowser.retinaPixelRatio) >> 0;
                        var _w = transition.Rect.w / AscCommon.AscBrowser.retinaPixelRatio;
                        var _h = transition.Rect.h / AscCommon.AscBrowser.retinaPixelRatio;

                        var _w_mm = manager.HtmlPage.m_oLogicDocument.GetWidthMM();
                        var _h_mm = manager.HtmlPage.m_oLogicDocument.GetHeightMM();

                        if (this.isReporterMode)
						{
                            _x += ((this.WordControl.m_oMainParent.AbsolutePosition.L * AscCommon.g_dKoef_mm_to_pix) >> 0);
						}

                        var zoom = _w / _w_mm;

                        if (!transform)
                            window["AscDesktopEditor"]["MediaStart"](sMediaName, _x >> 0, _y >> 0, extX, extY, zoom);
                        else
                            window["AscDesktopEditor"]["MediaStart"](sMediaName, _x >> 0, _y >> 0, extX, extY, zoom, transform.sx, transform.shy, transform.shx, transform.sy, transform.tx, transform.ty);
                    }

				}
				break;
			}
            case c_oEditorId.Spreadsheet:
            {
                break;
            }
		}
	};
	baseEditorsApi.prototype.hideVideoControl = function()
	{
        if (!window["AscDesktopEditor"] || !window["AscDesktopEditor"]["MediaEnd"])
            return;
        window["AscDesktopEditor"]["MediaEnd"]();
	};
	// plugins
	baseEditorsApi.prototype._checkLicenseApiFunctions   = function()
	{
		return this.licenseResult && true === this.licenseResult['plugins'];
	};

	baseEditorsApi.prototype.asc_pluginsRegister   = function(basePath, plugins)
	{
		if (null != this.pluginsManager)
		{
			this.pluginsManager.register(basePath, plugins);
			this.checkInstalledPlugins();
		}
		else
		{
			this.preSetupPlugins = {
				path : basePath,
				plugins : plugins
			};
		}
	};
	baseEditorsApi.prototype.asc_pluginRun         = function(guid, variation, pluginData)
	{
		if (null != this.pluginsManager)
			this.pluginsManager.run(guid, variation, pluginData);
	};
	baseEditorsApi.prototype.asc_pluginStop        = function(guid)
	{
		if (null != this.pluginsManager)
			this.pluginsManager.close(guid);
	};
	baseEditorsApi.prototype.asc_pluginResize      = function(pluginData)
	{
		if (null != this.pluginsManager)
			this.pluginsManager.runResize(pluginData);
	};
	baseEditorsApi.prototype.asc_pluginButtonClick = function(id, guid, windowId)
	{
		if (null != this.pluginsManager)
			this.pluginsManager.buttonClick(id, guid, windowId);
	};

	baseEditorsApi.prototype.asc_pluginEnableMouseEvents = function(isEnable)
	{
		if (!this.pluginsManager)
			return;

		this.pluginsManager.onEnableMouseEvents(isEnable);
	};

    baseEditorsApi.prototype.isEnabledDropTarget = function()
    {
    	return true;
    };
    baseEditorsApi.prototype.beginInlineDropTarget = function(e)
    {
    };
    baseEditorsApi.prototype.endInlineDropTarget = function(e)
    {
    };
    baseEditorsApi.prototype.isSliderDragged = function()
    {
		return this.noCreatePoint || this.exucuteHistory || this.exucuteHistoryEnd;
    };

    baseEditorsApi.prototype["asc_insertSymbol"] = function(familyName, code, pr)
    {
		var arrCharCodes = [code];
        AscFonts.FontPickerByCharacter.checkTextLight(arrCharCodes, true);

        var fonts = [new AscFonts.CFont(AscFonts.g_fontApplication.GetFontInfoName(familyName), 0, "", 0, null)];
        AscFonts.FontPickerByCharacter.extendFonts(fonts);

        this.asyncMethodCallback = function() {

            switch (this.editorId)
            {
                case c_oEditorId.Word:
                case c_oEditorId.Presentation:
                {
                	if (pr && c_oEditorId.Word === this.editorId)
					{
						this.WordControl.m_oLogicDocument.AddSpecialSymbol(pr);
					}
                	else
					{
						var textPr = new AscCommonWord.CTextPr();
						textPr.SetFontFamily(familyName);

						let settings = new AscCommon.CAddTextSettings();
						settings.SetTextPr(textPr);
						settings.MoveCursorOutside(true);

						this.WordControl.m_oLogicDocument.AddTextWithPr(new AscCommon.CUnicodeStringEmulator(arrCharCodes), settings);
					}
                    break;
                }
                case c_oEditorId.Spreadsheet:
                {
                	this.AddTextWithPr(familyName, arrCharCodes);
                    break;
                }
            }
        };

        if (false === AscCommon.g_font_loader.CheckFontsNeedLoading(fonts))
        {
            this.asyncMethodCallback();
            this.asyncMethodCallback = undefined;
            return;
        }

        AscCommon.g_font_loader.LoadDocumentFonts2(fonts);
    };

    baseEditorsApi.prototype["asc_registerPlaceholderCallback"] = function(nType, fCallback)
    {
		const oDrawingDocument = this.getDrawingDocument();
    	if (oDrawingDocument && oDrawingDocument.placeholders)
		{
			oDrawingDocument.placeholders.registerCallback(nType, fCallback);
		}
    };
    baseEditorsApi.prototype["asc_uncheckPlaceholders"] = function()
    {
        if (this.WordControl && this.WordControl.m_oDrawingDocument && this.WordControl.m_oDrawingDocument.placeholders)
        {
            this.WordControl.m_oDrawingDocument.placeholders.closeAllActive();
        }
    };

    // Builder
	baseEditorsApi.prototype.asc_nativeInitBuilder = function()
	{
		AscCommon.History.TurnOffChanges();
		this.asc_setDocInfo(new Asc.asc_CDocInfo());
	};
	baseEditorsApi.prototype.asc_SetSilentMode     = function()
	{
	};
	baseEditorsApi.prototype.asc_canPaste          = function()
	{
		return false;
	};
	baseEditorsApi.prototype.asc_Recalculate       = function()
	{
	};

	// Native
	baseEditorsApi.prototype['asc_nativeCheckPdfRenderer'] = function (_memory1, _memory2) {
		if (true) {
			// pos не должен минимизироваться!!!

			_memory1.Copy = _memory1["Copy"];
			_memory1.ClearNoAttack = _memory1["ClearNoAttack"];
			_memory1.WriteByte = _memory1["WriteByte"];
			_memory1.WriteBool = _memory1["WriteBool"];
			_memory1.WriteLong = _memory1["WriteLong"];
			_memory1.WriteDouble = _memory1["WriteDouble"];
			_memory1.WriteString = _memory1["WriteString"];
			_memory1.WriteString2 = _memory1["WriteString2"];

			_memory2.Copy = _memory1["Copy"];
			_memory2.ClearNoAttack = _memory1["ClearNoAttack"];
			_memory2.WriteByte = _memory1["WriteByte"];
			_memory2.WriteBool = _memory1["WriteBool"];
			_memory2.WriteLong = _memory1["WriteLong"];
			_memory2.WriteDouble = _memory1["WriteDouble"];
			_memory2.WriteString = _memory1["WriteString"];
			_memory2.WriteString2 = _memory1["WriteString2"];
		}

		var _printer = new AscCommon.CDocumentRenderer();
		_printer.Memory = _memory1;
		_printer.VectorMemoryForPrint = _memory2;
		return _printer;
	};

	// input
	baseEditorsApi.prototype.Begin_CompositeInput = function()
	{
	};
	baseEditorsApi.prototype.Add_CompositeText = function(nCharCode)
	{
	};
	baseEditorsApi.prototype.Remove_CompositeText = function(nCount)
	{
	};
	baseEditorsApi.prototype.Replace_CompositeText = function(arrCharCodes)
	{
	};
	baseEditorsApi.prototype.Set_CursorPosInCompositeText = function(nPos)
	{
	};
	baseEditorsApi.prototype.Get_CursorPosInCompositeText = function()
	{
	};
	baseEditorsApi.prototype.End_CompositeInput = function()
	{
	};
	baseEditorsApi.prototype.Get_MaxCursorPosInCompositeText = function()
	{
	};
	baseEditorsApi.prototype.Input_UpdatePos = function()
	{
	};
	baseEditorsApi.prototype["setInputParams"] = function(_obj)
	{
		window["AscInputMethod"] = window["AscInputMethod"] || {};

		for (var _prop in _obj)
		{
			window["AscInputMethod"][_prop] = _obj[_prop];
		}
	};

	baseEditorsApi.prototype.asc_addSignatureLine = function (oPr, Width, Height, sImgUrl) {

    };
	baseEditorsApi.prototype.asc_getAllSignatures = function () {
		return [];
	};

	baseEditorsApi.prototype.asc_CallSignatureDblClickEvent = function(sGuid){

	};

	// signatures
	baseEditorsApi.prototype.asc_AddSignatureLine2 = function(_obj)
	{
		var _w = 50;
		var _h = 50;
		var _w_pix = (_w * AscCommon.g_dKoef_mm_to_pix) >> 0;
		var _h_pix = (_h * AscCommon.g_dKoef_mm_to_pix) >> 0;
		var _canvas = document.createElement("canvas");
		_canvas.width = _w_pix;
		_canvas.height = _h_pix;
		var _ctx = _canvas.getContext("2d");
		_ctx.fillStyle = "#000000";
		_ctx.strokeStyle = "#000000";
		_ctx.font = "10pt 'Courier New'";
		_ctx.lineWidth = 3;

		_ctx.beginPath();
		var _y_line = (_h_pix >> 1) + 0.5;
		_ctx.moveTo(0, _y_line);
		_ctx.lineTo(_w_pix, _y_line);
		_ctx.stroke();
		_ctx.beginPath();

		_ctx.lineWidth = 2;
		_y_line -= 10;
		_ctx.moveTo(10, _y_line);
		_ctx.lineTo(25, _y_line - 10);
		_ctx.lineTo(10, _y_line - 20);
		_ctx.stroke();
		_ctx.beginPath();

		_ctx.fillText(_obj.asc_getSigner1(), 10, _y_line + 25);
		_ctx.fillText(_obj.asc_getSigner2(), 10, _y_line + 40);
		_ctx.fillText(_obj.asc_getEmail(), 10, _y_line + 55);

		var _url = _canvas.toDataURL("image/png");
		_canvas = null;

		var _args = [];

		this.ImageLoader.LoadImagesWithCallback([_url], function() {
			this.asc_addSignatureLine(_obj, _w, _h, _url);
		}, _args);
	};

	baseEditorsApi.prototype.asc_getRequestSignatures = function()
	{
		var _sigs = this.asc_getAllSignatures();
		var _sigs_ret = [];

		var _found;
		for (var i = _sigs.length - 1; i >= 0; i--)
		{
			var _sig = _sigs[i];
			_found = false;

			for (var j = this.signatures.length - 1; j >= 0; j--)
			{
				if (this.signatures[j].guid == _sig.id)
				{
					_found = true;
					break;
				}
			}

			if (!_found)
			{
				var _add_sig = new AscCommon.asc_CSignatureLine();
				_add_sig.guid = _sig.id;
				_add_sig.signer1 = _sig.signer;
				_add_sig.signer2 = _sig.signer2;
				_add_sig.email = _sig.email;
				_add_sig.showDate = _sig.showDate;
				_add_sig.instructions = _sig.instructions;

				_sigs_ret.push(_add_sig);
			}
		}

		return _sigs_ret;
	};

	baseEditorsApi.prototype.asc_Sign = function(id, guid, url1, url2)
	{
		if (window["AscDesktopEditor"] && !this.isRestrictionView())
			window["AscDesktopEditor"]["Sign"](id, guid, url1, url2);
	};
	baseEditorsApi.prototype.asc_RequestSign = function(guid)
	{
		var signGuid = (guid == "unvisibleAdd") ? AscCommon.CreateGUID() : guid;

		if (window["asc_LocalRequestSign"])
			window["asc_LocalRequestSign"](signGuid);
	};

	baseEditorsApi.prototype.asc_ViewCertificate = function(id)
	{
		if (window["AscDesktopEditor"])
			window["AscDesktopEditor"]["ViewCertificate"](parseInt("" + id)); // integer or string!
	};

	baseEditorsApi.prototype.asc_SelectCertificate = function()
	{
		if (window["AscDesktopEditor"])
			window["AscDesktopEditor"]["SelectCertificate"]();
	};

	baseEditorsApi.prototype.asc_GetDefaultCertificate = function()
	{
		if (window["AscDesktopEditor"])
			window["AscDesktopEditor"]["GetDefaultCertificate"]();
	};

	baseEditorsApi.prototype.asc_getSignatures = function()
	{
		return this.signatures;
	};

	baseEditorsApi.prototype.asc_RemoveSignature = function(guid)
	{
		if (window["AscDesktopEditor"])
			window["AscDesktopEditor"]["RemoveSignature"](guid);
	};

	baseEditorsApi.prototype.asc_RemoveAllSignatures = function()
	{
		if (window["AscDesktopEditor"])
			window["AscDesktopEditor"]["RemoveAllSignatures"]();
	};

	baseEditorsApi.prototype.asc_isSignaturesSupport = function()
	{
		if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsSignaturesSupport"])
			return window["AscDesktopEditor"]["IsSignaturesSupport"]();
		return false;
	};
    baseEditorsApi.prototype.asc_isProtectionSupport = function()
    {
        if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["IsProtectionSupport"])
            return window["AscDesktopEditor"]["IsProtectionSupport"]();
        return !(this.DocInfo && this.DocInfo.get_OfflineApp()) && this.isProtectionSupport;
    };
	baseEditorsApi.prototype.asc_isAnonymousSupport = function()
	{
		return this.isAnonymousSupport;
	};

	baseEditorsApi.prototype.asc_gotoSignature = function(guid)
	{
		if (window["AscDesktopEditor"] && window["asc_IsVisibleSign"] && window["asc_IsVisibleSign"](guid))
		{
			this.gotoSignatureInternal(guid);
		}
	};
	baseEditorsApi.prototype.gotoSignatureInternal = function(guid)
	{
	};

	baseEditorsApi.prototype.asc_getSignatureSetup = function(guid)
	{
		var _sigs = this.asc_getAllSignatures();

		for (var i = _sigs.length - 1; i >= 0; i--)
		{
			var _sig = _sigs[i];
			if (_sig.id == guid)
			{
				var _add_sig = new AscCommon.asc_CSignatureLine();
				_add_sig.guid = _sig.id;
				_add_sig.signer1 = _sig.signer;
				_add_sig.signer2 = _sig.signer2;
				_add_sig.email = _sig.email;
				_add_sig.showDate = _sig.showDate;
				_add_sig.instructions = _sig.instructions;

				_add_sig.isrequested = true;
				for (var j = 0; j < this.signatures.length; j++)
				{
					var signDoc = this.signatures[j];
					if (signDoc.guid == _add_sig.guid)
					{
						_add_sig.valid = signDoc.valid;
						_add_sig.isrequested = false;
						break;
					}
				}

				return _add_sig;
			}
		}

		return null;
	};

	baseEditorsApi.prototype.asc_getSignatureImage = function (sGuid) {

		var count = this.signatures.length;
		for (var i = 0; i < count; i++)
		{
			if (this.signatures[i].guid == sGuid)
				return this.signatures[i].image;
		}
		return "";
    };

	baseEditorsApi.prototype.asc_getSessionToken = function () {
		return this.CoAuthoringApi.get_jwt()
	};

	baseEditorsApi.prototype.asc_InputClearKeyboardElement = function()
	{
		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext.nativeFocusElement = null;
	};

    // input helper
    baseEditorsApi.prototype.getTargetOnBodyCoords = function()
    {
        var ret = { X : 0, Y : 0, W : 0, H : 0, TargetH : 0 };
        ret.W = Math.max(document.documentElement.clientWidth, window.innerWidth || 0);
        ret.H = Math.max(document.documentElement.clientHeight, window.innerHeight || 0);

        switch (this.editorId)
        {
            case c_oEditorId.Word:
            {
                ret.X += this.WordControl.X;
                ret.Y += this.WordControl.Y;
                ret.X += (this.WordControl.m_oMainView.AbsolutePosition.L * AscCommon.g_dKoef_mm_to_pix);
                ret.Y += (this.WordControl.m_oMainView.AbsolutePosition.T * AscCommon.g_dKoef_mm_to_pix);
                ret.X += (this.WordControl.m_oDrawingDocument.TargetHtmlElementLeft);
                ret.Y += (this.WordControl.m_oDrawingDocument.TargetHtmlElementTop);

                ret.X >>= 0;
                ret.Y >>= 0;

                ret.TargetH = (this.WordControl.m_oDrawingDocument.m_dTargetSize * this.WordControl.m_nZoomValue * AscCommon.g_dKoef_mm_to_pix / 100) >> 0;
                break;
            }
            case c_oEditorId.Presentation:
            {
                ret.X += this.WordControl.X;
                ret.Y += this.WordControl.Y;

                ret.X += (this.WordControl.m_oMainParent.AbsolutePosition.L * AscCommon.g_dKoef_mm_to_pix);

                if (!this.WordControl.m_oLogicDocument.IsFocusOnNotes())
                {
                    ret.Y += (this.WordControl.m_oMainView.AbsolutePosition.T * AscCommon.g_dKoef_mm_to_pix);
                }
                else
                {
                    ret.Y += (this.WordControl.m_oNotesContainer.AbsolutePosition.T * AscCommon.g_dKoef_mm_to_pix);
                }

                ret.X += (this.WordControl.m_oDrawingDocument.TargetHtmlElementLeft);
                ret.Y += (this.WordControl.m_oDrawingDocument.TargetHtmlElementTop);

                ret.X >>= 0;
                ret.Y >>= 0;

                ret.TargetH = (this.WordControl.m_oDrawingDocument.m_dTargetSize * this.WordControl.m_nZoomValue * AscCommon.g_dKoef_mm_to_pix / 100) >> 0;
                break;
            }
            case c_oEditorId.Spreadsheet:
            {
                var off, selectionType = this.asc_getCellInfo().asc_getSelectionType();
                if (this.asc_getCellEditMode())
                {
                    // cell edit
                    var cellEditor = this.wb.cellEditor;
                    ret.X = cellEditor.curLeft;
                    ret.Y = cellEditor.curTop;
                    ret.TargetH = cellEditor.curHeight;
                    off = cellEditor.cursor;
                }
                else if (Asc.c_oAscSelectionType.RangeShapeText === selectionType ||
                    Asc.c_oAscSelectionType.RangeChartText === selectionType)
                {
                    // shape target
                    var drDoc = this.wb.getWorksheet().objectRender.controller.drawingDocument;
                    ret.X = drDoc.TargetHtmlElementLeft;
                    ret.Y = drDoc.TargetHtmlElementTop;
                    ret.TargetH = drDoc.m_dTargetSize * this.asc_getZoom() * AscCommon.g_dKoef_mm_to_pix;
                    off = this.HtmlElement;
                }

                if (off) {
                    off = jQuery(off).offset();
                    if (off)
                    {
                        ret.X += off.left;
                        ret.Y += off.top;
                    }
                }

                ret.X >>= 0;
                ret.Y >>= 0;
                ret.TargetH >>= 0;
                break;
            }
        }
        return ret;
    };

	baseEditorsApi.prototype.onKeyDown = function(e)
	{
	};
	baseEditorsApi.prototype.onKeyPress = function(e)
	{
	};
	baseEditorsApi.prototype.onKeyUp = function(e)
	{
	};
	/**
	 * Получаем текст (в виде массива юникодных значений), который будет добавлен на ивенте KeyDown
	 * @param e
	 * @returns {Number[]}
	 */
	baseEditorsApi.prototype.getAddedTextOnKeyDown = function(e)
	{
		return [];
	};
	baseEditorsApi.prototype.pre_Paste = function(_fonts, _images, callback)
	{
	};

	baseEditorsApi.prototype.asc_Remove = function()
	{
		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext.emulateKeyDownApi(46);
	};

	// System input
	baseEditorsApi.prototype.SetTextBoxInputMode = function(bIsEnable)
	{
		AscCommon.TextBoxInputMode = bIsEnable;
		if (AscCommon.g_inputContext)
			AscCommon.g_inputContext.systemInputEnable(AscCommon.TextBoxInputMode);
	};
	baseEditorsApi.prototype.GetTextBoxInputMode = function()
	{
		return AscCommon.TextBoxInputMode;
	};

	baseEditorsApi.prototype.asc_OnHideContextMenu = function()
	{
	};
	baseEditorsApi.prototype.asc_OnShowContextMenu = function()
	{
	};

	baseEditorsApi.prototype.isIdle = function()
	{
		// пока не стартовали - считаем работаем
		if (0 == this.lastWorkTime)
			return 0;

		// если плагин работает - то и мы тоже
		if (this.pluginsManager && this.pluginsManager.isWorked())
			return 0;

		if (this.isEmbedVersion)
			return 0;

		if (!this.canSave || !this._saveCheck())
			return 0;

		if (this.isPdfEditor())
			return 0;

		//viewer
		if (this.isViewMode)
			return 0;

		return new Date().getTime() - this.lastWorkTime;
	};

	baseEditorsApi.prototype.checkInterfaceElementBlur = function()
	{
		if (!document.activeElement || !document.createEvent || (document.activeElement.id === "area_id"))
			return;

		if (!document.activeElement.dispatchEvent)
			return;

		var e = document.createEvent("HTMLEvents");
		e.initEvent("blur", true, true);
		e.eventName = "blur";
		document.activeElement.dispatchEvent(e);
	};

	baseEditorsApi.prototype.checkLastWork = function()
	{
		this.lastWorkTime = new Date().getTime();
	};

	baseEditorsApi.prototype.setViewModeDisconnect = function(enableDownload)
	{
		// Посылаем наверх эвент об отключении от сервера
		this.sendEvent('asc_onCoAuthoringDisconnect', enableDownload);
		// И переходим в режим просмотра т.к. мы не можем сохранить файл
		this.asc_setViewMode(true);
	};

	baseEditorsApi.prototype.asc_setCurrentPassword = function(password)
	{
		this.currentPassword = password;
		this.asc_Save(false, undefined, true);
		if (!(this.DocInfo && this.DocInfo.get_OfflineApp())) {
			var rData = {
				"c": 'setpassword',
				"id": this.documentId,
				"password": password
			};
			var t            = this;
			t.fCurCallback   = function(input)
			{
				if (null != input && "setpassword" == input["type"])
				{
					if ('ok' === input["status"])
					{
						t.sendEvent("asc_onDocumentPassword", "" !== t.currentPassword);
					}
					else
					{
						t.sendEvent("asc_onError", AscCommon.mapAscServerErrorToAscError(parseInt(input["data"])),
							c_oAscError.Level.NoCritical);
					}
				}
				else
				{
					t.sendEvent("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
				}
			};
			AscCommon.sendCommand(this, null, rData);
		}
	};
	baseEditorsApi.prototype.asc_resetPassword = function()
	{
		this.asc_setCurrentPassword("");
	};

	baseEditorsApi.prototype.asc_GetPossibleNumberingLanguage = function () {};
	baseEditorsApi.prototype.CheckDeprecatedBulletPreviewInfo = function (arrDrawingInfo, nTypeOfPreview)
	{
		const arrAdaptedDrawingInfo = [];

		for (let i = 0; i < arrDrawingInfo.length; i += 1)
		{
			let sDivId;
			let oNumberingInfo;

			// Это верный тип передачи информации
			if (arrDrawingInfo[i]["numberingInfo"])
			{
				arrAdaptedDrawingInfo.push(arrDrawingInfo[i]);
			}
			else if (arrDrawingInfo[i]["divId"])
			{
				sDivId = arrDrawingInfo[i]["divId"];
				let sResult = null;
				switch (arrDrawingInfo[i]["type"])
				{
					case 0:
					{
						break;
					}
					case 1:
					{
						const oBullet = new AscFormat.CBullet();
						oBullet.fillBulletFromCharAndFont(arrDrawingInfo[i]["char"], arrDrawingInfo[i]["specialFont"] || "Arial");
						oNumberingInfo = oBullet.getJsonBullet();
						break;
					}
					case 2:
					{
						const oBullet = new AscFormat.CBullet();
						oBullet.fillBulletImage(arrDrawingInfo[i]["imageId"]);
						oNumberingInfo = oBullet.getJsonBullet();
						break;
					}
					case 3:
					{
						switch (arrDrawingInfo[i]["numberingType"])
						{
							case Asc.c_oAscNumberingLevel.UpperLetterDot_Left: sResult = '{"bulletTypeface":{"type":"bufont","typeface":"Arial"},"bulletType":{"type":"autonum","char":null,"autoNumType":"alphaUcPeriod","startAt":null}}'; break;
							case Asc.c_oAscNumberingLevel.LowerLetterBracket_Left: sResult = '{"bulletTypeface":{"type":"bufont","typeface":"Arial"},"bulletType":{"type":"autonum","char":null,"autoNumType":"alphaLcParenR","startAt":null}}'; break;
							case Asc.c_oAscNumberingLevel.LowerLetterDot_Left: sResult = '{"bulletTypeface":{"type":"bufont","typeface":"Arial"},"bulletType":{"type":"autonum","char":null,"autoNumType":"alphaLcPeriod","startAt":null}}'; break;
							case Asc.c_oAscNumberingLevel.DecimalDot_Right: sResult = '{"bulletTypeface":{"type":"bufont","typeface":"Arial"},"bulletType":{"type":"autonum","char":null,"autoNumType":"arabicPeriod","startAt":null}}'; break;
							case Asc.c_oAscNumberingLevel.DecimalBracket_Right: sResult = '{"bulletTypeface":{"type":"bufont","typeface":"Arial"},"bulletType":{"type":"autonum","char":null,"autoNumType":"arabicParenR","startAt":null}}'; break;
							case Asc.c_oAscNumberingLevel.UpperRomanDot_Right: sResult = '{"bulletTypeface":{"type":"bufont","typeface":"Arial"},"bulletType":{"type":"autonum","char":null,"autoNumType":"romanUcPeriod","startAt":null}}'; break;
							case Asc.c_oAscNumberingLevel.LowerRomanDot_Right: sResult = '{"bulletTypeface":{"type":"bufont","typeface":"Arial"},"bulletType":{"type":"autonum","char":null,"autoNumType":"romanLcPeriod","startAt":null}}'; break;
							default: break;
						}
						break;
					}
					default: break;
				}
				if (sResult !== null)
					oNumberingInfo = JSON.parse(sResult);
				arrAdaptedDrawingInfo.push({"divId": sDivId, "numberingInfo": {"bullet": oNumberingInfo}});
			}
		}

		return arrAdaptedDrawingInfo;
	};
	baseEditorsApi.prototype.ParseBulletPreviewInformation = function (arrDrawingInfo)
	{
		const arrNumberingLvls = [];
		AscFormat.ExecuteNoHistory(function ()
		{
			for (let i = 0; i < arrDrawingInfo.length; i += 1)
			{
				const oDrawInfo = arrDrawingInfo[i];
				const oNumberingInfo = oDrawInfo["numberingInfo"];
				if (!oNumberingInfo) continue;
				const sDivId = oDrawInfo["divId"];
				if (!oNumberingInfo["bullet"])
				{
					const oPresentationBullet = new AscCommonWord.CPresentationBullet();
					const oTextPr = new AscCommonWord.CTextPr();
					oPresentationBullet.m_sChar = AscCommon.translateManager.getValue("None");
					oPresentationBullet.m_nType = AscFormat.numbering_presentationnumfrmt_Char;
					oPresentationBullet.m_bFontTx = false;
					oPresentationBullet.m_sFont   = "Arial";
					oTextPr.Unifill = AscFormat.CreateSolidFillRGB(0, 0, 0);
					oTextPr.FontSize = oTextPr.FontSizeCS = 65;
					oPresentationBullet.MergeTextPr(oTextPr);
					arrNumberingLvls.push({divId: sDivId, arrLvls: [oPresentationBullet], isRemoving: true});
				}
				else
				{
					const oBullet = window['AscJsonConverter'].ReaderFromJSON.prototype.BulletFromJSON(oNumberingInfo["bullet"]);
					const oPresentationBullet = oBullet.getPresentationBullet(AscFormat.GetDefaultTheme(), AscFormat.GetDefaultColorMap());
					oPresentationBullet.m_bFontTx = false;
					const oTextPr = new AscCommonWord.CTextPr();
					oTextPr.Unifill = AscFormat.CreateSolidFillRGB(0, 0, 0);
					oTextPr.FontSize = oTextPr.FontSizeCS = 65;
					oPresentationBullet.MergeTextPr(oTextPr);
					arrNumberingLvls.push({divId: sDivId, arrLvls: [oPresentationBullet]});
				}
			}
		}, this);
		return arrNumberingLvls;
	};
	baseEditorsApi.prototype.SetDrawImagePreviewBulletForMenu = function(arrDrawingInfo, nType)
	{
		const arrAdaptedDrawingInfo = this.CheckDeprecatedBulletPreviewInfo(arrDrawingInfo, nType);
		const arrParsedInfo = this.ParseBulletPreviewInformation(arrAdaptedDrawingInfo);
		const oDrawer = new AscCommon.CBulletPreviewDrawer(arrParsedInfo, nType);
		oDrawer.checkFontsAndDraw();
	};
	baseEditorsApi.prototype.asc_setMacros = function(sData)
	{
		if (!this.macros)
			return true;

		if (true === AscCommon.CollaborativeEditing.Get_GlobalLock())
			return true;

		AscCommon.CollaborativeEditing.OnStart_CheckLock();
		this.macros.CheckLock();

		if (this.editorId == AscCommon.c_oEditorId.Spreadsheet)
		{
			var _this = this;
			Asc.editor.checkObjectsLock([this.macros.Get_Id()], function(bNoLock) {
				if (bNoLock)
				{
					AscCommon.History.Create_NewPoint(AscDFH.historydescription_DocumentMacros_Data);
					_this.macros.SetData(sData);
				}
			});
		}
		else
		{
			if (false === AscCommon.CollaborativeEditing.OnEnd_CheckLock(false))
			{
				AscCommon.History.Create_NewPoint(AscDFH.historydescription_DocumentMacros_Data);
				this.macros.SetData(sData);
			}
		}
	};
	baseEditorsApi.prototype.asc_getMacros = function()
	{
		return this.macros.GetData();
	};

	baseEditorsApi.prototype._beforeEvalCommand = function()
	{
		var oApi = this;
		switch (this.editorId)
		{
			case AscCommon.c_oEditorId.Word:
			{
				if (this.WordControl && this.WordControl.m_oLogicDocument)
					this.WordControl.m_oLogicDocument.LockPanelStyles();
				break;
			}
			case AscCommon.c_oEditorId.Spreadsheet:
			{
				if (AscCommonExcel)
				{
					oApi.tmpR1C1mode = AscCommonExcel.g_R1C1Mode;
					AscCommonExcel.g_R1C1Mode = false;
				}
				break;
			}
			default:
				break;
		}
	};

	baseEditorsApi.prototype._afterEvalCommand = function(endAction)
	{
		var oApi = this;
		switch (this.editorId)
		{
			case AscCommon.c_oEditorId.Word:
			case AscCommon.c_oEditorId.Presentation:
			{
				var oLogicDocument = this.WordControl.m_oLogicDocument;
				if (!oLogicDocument)
				{
					endAction && endAction();
					return;
				}

				var _imagesArray = oLogicDocument.Get_AllImageUrls();
				var _images = {};
				for (var i = 0; i < _imagesArray.length; i++)
					_images[_imagesArray[i]] = _imagesArray[i];

				AscCommon.Check_LoadingDataBeforePrepaste(this, oLogicDocument.Document_Get_AllFontNames(), _images, function() {
					if (oLogicDocument.Reassign_ImageUrls)
						oLogicDocument.Reassign_ImageUrls(_images);

					if (AscCommon.c_oEditorId.Word === oApi.editorId)
					{
						oLogicDocument.UnlockPanelStyles(true);
						oLogicDocument.OnEndLoadScript();
					}

					oApi.asc_Recalculate(true);
					oLogicDocument.FinalizeAction();

					if (oApi.SaveAfterMacros)
					{
						oApi.asc_Save();
						oApi.SaveAfterMacros = false;
					}

					endAction && endAction();
				});
				break;
			}
			case AscCommon.c_oEditorId.Spreadsheet:
			{
				var oModel = this.wbModel;
				var _imagesArray = oModel.getAllImageUrls();
				var _images = {};
				for (var i = 0; i < _imagesArray.length; i++)
					_images[_imagesArray[i]] = _imagesArray[i];

				AscCommon.Check_LoadingDataBeforePrepaste(this, oModel._generateFontMap(), _images, function() {
					oModel.reassignImageUrls(_images);
					oApi.asc_Recalculate(true);
					var wsView = oApi.wb && oApi.wb.getWorksheet();
					if (wsView && wsView.objectRender && wsView.objectRender.controller)
					{
						wsView.objectRender.controller.recalculate2(undefined);
					}

					if (oApi.SaveAfterMacros)
					{
						oApi.asc_Save();
						oApi.SaveAfterMacros = false;
					}

					endAction && endAction();
				});
				
				if (AscCommonExcel)
				{
					AscCommonExcel.g_R1C1Mode = oApi.tmpR1C1mode;
					oApi.tmpR1C1mode = null;
				}
				break;
			}
			default:
				break;
		}
	};


	baseEditorsApi.prototype.asc_runAutostartMacroses = function()
    {
    	if (!this.macros || this.disableAutostartMacros)
    		return;

		if (this.DocInfo && !this.DocInfo.asc_getIsEnabledMacroses())
			return;

    	if (!this.asc_canPaste())
    		return;

		if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["isSupportMacroses"] && !window["AscDesktopEditor"]["isSupportMacroses"]())
			return;

    	this._beforeEvalCommand();
		this.macros.runAuto();
		this._afterEvalCommand(undefined);
    };
	baseEditorsApi.prototype.asc_runMacros = function(sGuid)
    {
    	if (!this.macros)
    		return;

		if (this.DocInfo && !this.DocInfo.asc_getIsEnabledMacroses())
			return;

    	if (!this.asc_canPaste())
    		return;

		if (window["AscDesktopEditor"] && window["AscDesktopEditor"]["isSupportMacroses"] && !window["AscDesktopEditor"]["isSupportMacroses"]())
			return;

    	this._beforeEvalCommand();
		this.macros.run(sGuid);
		this._afterEvalCommand(undefined);
    };
	baseEditorsApi.prototype.asc_getAllMacrosNames = function()
    {
    	if (!this.macros)
    		return [];

		return this.macros.getAllNames();
    };
	baseEditorsApi.prototype.asc_getMacrosGuidByName = function(sName)
    {
    	if (!this.macros)
    		return "";

		return this.macros.getGuidByName(sName);
    };
	baseEditorsApi.prototype.asc_getMacrosByGuid = function(sGuid)
    {
    	if (!this.macros)
    		return "";

		return this.macros.getNameByGuid(sGuid);
    };
	baseEditorsApi.prototype.asc_getUserPermissionToMakeRequestFromMacros = function(url, callback)
	{
		var bNeedCallback = this.editorId === AscCommon.c_oEditorId.Spreadsheet ? this.handlers.hasTrigger('asc_onMacrosPermissionRequest') : this.asc_checkNeedCallback('asc_onMacrosPermissionRequest');
		if (bNeedCallback) {
			this.sendEvent('asc_onMacrosPermissionRequest', url, callback);
		} else {
			var permission = confirm("A macro makes a request to URL. Do you want to allow the request to the '" + url +"' ?");
			callback(permission);
		}
	};

	baseEditorsApi.prototype.asc_getSelectedDrawingObjectsCount = function()
	{
		return 0;
	};

	baseEditorsApi.prototype.asc_decodeBuffer = function(buffer, options, callback) {
		var reader = new FileReader();
		//todo onerror
		reader.onload = reader.onerror = function(e) {
			var text = e.target.result ? e.target.result : "";
			if (options instanceof Asc.asc_CTextOptions) {
				callback(AscCommon.parseText(text, options));
			} else {
				callback(text.match(/[^\r\n]+/g));
			}
		};

		var encoding = "UTF-8";
		var codePage = options.asc_getCodePage();
		var encodingsLen = AscCommon.c_oAscEncodings.length;
		for (var i = 0; i < encodingsLen; ++i)
		{
			if (AscCommon.c_oAscEncodings[i][0] == codePage)
			{
				encoding = AscCommon.c_oAscEncodings[i][2];
				break;
			}
		}

		reader.readAsText(new Blob([buffer]), encoding);
	};

	baseEditorsApi.prototype.asc_setVisiblePasteButton = function(val)
	{
		if (AscCommon.g_specialPasteHelper)
		{
			AscCommon.g_specialPasteHelper.setVisiblePasteButton(val);
		}
	};
	/**
	 * Return default array for Math AutoCorrect symbols for menu
	 * @returns {Array.<String, Number || Array.<Number>>}
	 */
	baseEditorsApi.prototype.asc_getAutoCorrectMathSymbols = function()
	{
        return window['AscCommonWord'].g_AutoCorrectMathsList.DefaultAutoCorrectMathSymbolsList;
	};
	/**
	 * Return default array for Math AutoCorrect functions for menu
	 * @returns {Array.<String>}
	 */
	baseEditorsApi.prototype.asc_getAutoCorrectMathFunctions = function()
	{
        return window['AscCommonWord'].g_AutoCorrectMathsList.DefaultAutoCorrectMathFuncs;
	};
	/**
	 * Reset to defaul math autocorrect symbols list
	 */
	baseEditorsApi.prototype.asc_resetToDefaultAutoCorrectMathSymbols = function()
	{
		window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathSymbols = JSON.parse(JSON.stringify(window['AscCommonWord'].g_AutoCorrectMathsList.DefaultAutoCorrectMathSymbolsList));
		AscMath.UpdateAutoCorrection();
	};
	/**
	 * Reset to default version math autocorrect functions list
	 */
	baseEditorsApi.prototype.asc_resetToDefaultAutoCorrectMathFunctions = function()
	{
		window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathFuncs = JSON.parse(JSON.stringify(window['AscCommonWord'].g_AutoCorrectMathsList.DefaultAutoCorrectMathFuncs));
		AscMath.UpdateFuncCorrection();
	};
	/**
	 * Delete item from g_AutoCorrectMathSymbols
	 * @param {string} element
	 */
	baseEditorsApi.prototype.asc_deleteFromAutoCorrectMathSymbols = function(element)
	{
        var remInd = window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathSymbols.findIndex(function(val, index){
			if (val[0] === element){
				return true;
			}
		});
		window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathSymbols.splice(remInd, 1);
		AscMath.UpdateAutoCorrection();
	};
	/**
	 * Delete item from g_AutoCorrectMathFuncs
	 * @param {string} element
	 */
	baseEditorsApi.prototype.asc_deleteFromAutoCorrectMathFunctions = function(element)
	{
        var remInd = window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathFuncs.findIndex(function(val, index){
			if (val === element){
				return true;
			}
		});
		window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathFuncs.splice(remInd, 1);
		AscMath.UpdateFuncCorrection();
	};
	/**
	 * Add or edit item from g_AutoCorrectMathSymbols
	 * @param {string} element
	 * @param {Array.<number> || number || string} repVal
	 */
	baseEditorsApi.prototype.asc_AddOrEditFromAutoCorrectMathSymbols = function(element, repVal)
	{
		if (typeof repVal === 'string') {
			repVal = AscCommon.convertUTF16toUnicode(repVal);
		}
		var changeInd = window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathSymbols.findIndex(function(val, index){
			if (val[0] === element){
				return true;
			}
		});
		if (changeInd >= 0) {
			window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathSymbols[changeInd][1] = repVal;
		} else {
			window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathSymbols.push([element, repVal]);
		}

		AscMath.UpdateAutoCorrection();
	};
	/**
	 * Add item from g_AutoCorrectMathFuncs
	 * @param {string} newEl
	 */
	baseEditorsApi.prototype.asc_AddFromAutoCorrectMathFunctions = function(newEl)
	{
		window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathFuncs.push(newEl);
		AscMath.UpdateFuncCorrection();
	};
	/**
	 * Refresh g_AutoCorrectMathSymbols on start
	 * @param {Array.<string>} remItems
	 * @param {Array.<string, Array.<number> || number || string>} addItems
	 * @param {boolean} flag
	 */
	baseEditorsApi.prototype.asc_refreshOnStartAutoCorrectMathSymbols = function(remItems, addItems, flag)
	{
		var me = this;
		this.asc_resetToDefaultAutoCorrectMathSymbols();
		if (remItems) {
			remItems.forEach(function(el) {
				me.asc_deleteFromAutoCorrectMathSymbols(el);
			});
		}
		if (addItems) {
			addItems.forEach(function(el) {
				me.asc_AddOrEditFromAutoCorrectMathSymbols(el[0], el[1]);
			});
		}
		this.asc_updateFlagAutoCorrectMathSymbols(flag);
	};
	/**
	 * Refresh g_AutoCorrectMathFuncs on start
	 * @param {Array.<string>} remItems
	 * @param {Array.<string>} addItems
	 */
	baseEditorsApi.prototype.asc_refreshOnStartAutoCorrectMathFunctions = function(remItems, addItems)
	{
		var me = this;
		this.asc_resetToDefaultAutoCorrectMathFunctions();
		if (remItems) {
			remItems.forEach(function(el) {
				me.asc_deleteFromAutoCorrectMathFunctions(el);
			});
		}
		if (addItems) {
			addItems.forEach(function(el) {
				me.asc_AddFromAutoCorrectMathFunctions(el);
			});
		}
	};
	baseEditorsApi.prototype.asc_updateFlagAutoCorrectMathSymbols = function(flag)
	{
		window['AscCommonWord'].b_DoAutoCorrectMathSymbols = flag;
	};
	baseEditorsApi.prototype.asc_AddMath = function(type)
	{};
	baseEditorsApi.prototype.asc_ConvertMathView = function(isToLinear, isAll)
	{
	};
	baseEditorsApi.prototype.getMathInputType = function()
	{
		return this.mathInputType;
	};
	baseEditorsApi.prototype.asc_GetMathInputType = function()
	{
		return this.getMathInputType();
	};
	baseEditorsApi.prototype.asc_SetMathInputType = function(type)
	{
		this.mathInputType = type;
	};

	baseEditorsApi.prototype.getFileAsFromChanges = function()
	{
		var func_before = null;
		var func_after = null;
		if (this.editorId === AscCommon.c_oEditorId.Word)
		{
			if (this.WordControl && this.WordControl.m_oLogicDocument && this.WordControl.m_oLogicDocument.IsViewModeInReview())
			{
				var isFinal = (this.WordControl.m_oLogicDocument.ViewModeInReview.mode === 1) ? true : false;

				func_before = function(api) {
					api.WordControl.m_oLogicDocument.Start_SilentMode();
					api.asc_EndViewModeInReview();
				};
				func_after = function(api) {
					api.asc_BeginViewModeInReview(isFinal);
					api.WordControl.m_oLogicDocument.End_SilentMode(false);
				};
			}
		}

		func_before && func_before(this);
		var ret = this.asc_nativeGetFile3();
		func_after && func_after(this);
		return ret;
	};

	baseEditorsApi.prototype.initShortcuts = function(arrShortcuts, isRemoveBeforeAdd)
	{
		// Массив
		// [[ActionType, KeyCode, Ctrl, Shift, Alt]]
		for (var nIndex = 0, nCount = arrShortcuts.length; nIndex < nCount; ++nIndex)
		{
			var s = arrShortcuts[nIndex];

			if (true === isRemoveBeforeAdd)
				this.Shortcuts.RemoveByType(s[0]);

			this.Shortcuts.Add(s[0], s[1], s[2], s[3], s[4]);
		}
	};
	baseEditorsApi.prototype.initDefaultShortcuts = function()
	{
	};
	baseEditorsApi.prototype.getShortcut = function(e)
	{
		if (e.GetKeyCode)
			return this.Shortcuts.Get(e.GetKeyCode(), e.IsCtrl(), e.IsShift(), e.IsAlt());
		else
			return this.Shortcuts.Get(e.KeyCode, e.CtrlKey, e.ShiftKey, e.AltKey);
	};
	baseEditorsApi.prototype.getCustomShortcutAction = function(nActionType)
	{
		return this.Shortcuts.GetCustomAction(nActionType);
	};
	baseEditorsApi.prototype.asc_initShortcuts = baseEditorsApi.prototype.initShortcuts;
	baseEditorsApi.prototype.asc_getShortcutAction = function(nKeyCode, isCtrl, isShift, isAlt)
	{
		return this.Shortcuts.Get(nKeyCode, isCtrl, isShift, isAlt);
	};
	baseEditorsApi.prototype.asc_removeShortcuts = function(arrShortcuts, arrActionTypes)
	{
		if (arrShortcuts)
		{
			for (var nIndex = 0, nCount = arrShortcuts.length; nIndex < nCount; ++nIndex)
			{
				var s = arrShortcuts[nIndex];
				this.Shortcuts.Remove(s[0], s[1], s[2], s[3]);
			}
		}

		if (arrActionTypes)
		{
			for (var nIndex = 0, nCount = arrActionTypes.length; nIndex < nCount; ++nIndex)
			{
				this.Shortcuts.RemoveByType(arrActionTypes[nIndex]);
			}
		}
	};
	baseEditorsApi.prototype.asc_addCustomShortcutInsertSymbol = function(nCharCode, sFont, sShortcut)
	{
		var nActionType = this.Shortcuts.AddCustomActionSymbol(nCharCode, sFont);
		this.Shortcuts.Add(nActionType, sShortcut[0], sShortcut[1], sShortcut[2], sShortcut[3]);
		return nActionType;
	};
	baseEditorsApi.prototype.asc_setSkin = baseEditorsApi.prototype["asc_setSkin"] = function(skin)
	{
		AscCommon.updateGlobalSkin(skin);
		this.updateSkin(skin);
	};
	baseEditorsApi.prototype.updateSkin = function()
	{
	};
	baseEditorsApi.prototype.isLocalMode = function()
	{
		if (window["AscDesktopEditor"])
		{
			if (window["AscDesktopEditor"]["IsLocalFile"]())
				return true;
			if (AscCommon.EncryptionWorker && AscCommon.EncryptionWorker.isNeedCrypt())
				return true;
		}
		return false;
	};
	baseEditorsApi.prototype.isCloudModeCrypt = function()
	{
		if (window["AscDesktopEditor"])
		{
			if (window["AscDesktopEditor"]["IsLocalFile"]())
				return false;
			if (AscCommon.EncryptionWorker && AscCommon.EncryptionWorker.isNeedCrypt())
				return true;
		}
		return false;
	};
	baseEditorsApi.prototype.asc_initPrintPreview                     = function()
	{
	};
	baseEditorsApi.prototype.asc_drawPrintPreview                     = function()
	{
	};
	baseEditorsApi.prototype.asc_EditSelectAll = function()
	{
	};
	baseEditorsApi.prototype.asc_enterText = function(value)
	{
	};
	baseEditorsApi.prototype.asc_correctEnterText = function(oldValue, newValue)
	{
	};
	baseEditorsApi.prototype.asc_setContentDarkMode = baseEditorsApi.prototype["asc_setContentDarkMode"] = function(isDarkMode)
	{
		if (this.isDarkMode === isDarkMode)
			return;
		
		this.isDarkMode = isDarkMode;
		
		this.updateDarkMode();
	};
	baseEditorsApi.prototype.updateDarkMode = function()
	{
	};
	//---------------------------------------------------------statistics-----------------------------------------------
	baseEditorsApi.prototype.startGetDocInfo = function()
	{
		this.sync_GetDocInfoStartCallback();
		this.sync_GetDocInfoEndCallback();
	};
	baseEditorsApi.prototype.stopGetDocInfo = function()
	{
		this.sync_GetDocInfoStopCallback();
	};
	baseEditorsApi.prototype.sync_DocInfoCallback = function(obj)
	{
		this.sendEvent("asc_onDocInfo", new AscCommon.CDocInfoProp(obj));
	};
	baseEditorsApi.prototype.sync_GetDocInfoStartCallback = function()
	{
		this.sendEvent("asc_onGetDocInfoStart");
	};
	baseEditorsApi.prototype.sync_GetDocInfoStopCallback = function()
	{
		this.sendEvent("asc_onGetDocInfoStop");
	};
	baseEditorsApi.prototype.sync_GetDocInfoEndCallback = function()
	{
		this.sendEvent("asc_onGetDocInfoEnd");
	};
	//---------------------------------------------------------search-----------------------------------------------------
	baseEditorsApi.prototype.asc_searchEnabled = function(isEnabled)
	{
	};
	baseEditorsApi.prototype.asc_findText = function(oProps, isNext, callback)
	{
		return 0;
	};
	baseEditorsApi.prototype.asc_endFindText = function()
	{
	};
	baseEditorsApi.prototype.asc_selectSearchingResults = function(isShow)
	{
		if (this.selectSearchingResults === isShow)
			return;
		
		this.selectSearchingResults = isShow;
		
		this._selectSearchingResults(isShow);
	};
	baseEditorsApi.prototype.asc_isSelectSearchingResults = function()
	{
		return this.selectSearchingResults;
	};
	baseEditorsApi.prototype._selectSearchingResults = function()
	{
	};
	baseEditorsApi.prototype.asc_StartTextAroundSearch = function()
	{
	};
	baseEditorsApi.prototype.sync_setSearchCurrent = function(nCurrent, nOverallCount)
	{
		this.sendEvent("asc_onSetSearchCurrent", nCurrent, nOverallCount);
	};
	baseEditorsApi.prototype.sync_startTextAroundSearch = function()
	{
		this.sendEvent("asc_onStartTextAroundSearch");
	};
	baseEditorsApi.prototype.sync_endTextAroundSearch = function()
	{
		this.sendEvent("asc_onEndTextAroundSearch");
	};
	baseEditorsApi.prototype.sync_getTextAroundSearchPack = function(arrElements)
	{
		this.sendEvent("asc_onGetTextAroundSearchPack", arrElements);
	};
	baseEditorsApi.prototype.sync_removeTextAroundSearch = function(sId)
	{
		this.sendEvent("asc_onRemoveTextAroundSearch", [sId]);
	};
	//---------------------------------------------------------version----------------------------------------------------
	baseEditorsApi.prototype["GetVersion"] = baseEditorsApi.prototype.GetVersion = function()
	{
		this.versionSdk = AscCommon.g_cProductVersion;
		return (this.versionSdk === "0.0.0" || this.versionSdk.substr(2) === "Version") ? "develop" : this.versionSdk;
	};
	//----------------------------------------------------------addons----------------------------------------------------
	baseEditorsApi.prototype["asc_isSupportFeature"] = function(type)
	{
		return (window["Asc"] && window["Asc"]["Addons"] && window["Asc"]["Addons"][type] === true) ? true : false;
	};

	baseEditorsApi.prototype["asc_setDefaultBlitMode"] = function(value)
	{
		AscFonts.setDefaultBlitting(value);
	};

    // ---------------------------------------------------- internal events ----------------------------------------------
    baseEditorsApi.prototype.attachEvent = function(name, callback, listenerId)
    {
    	if (!this.internalEvents.hasOwnProperty(name))
            this.internalEvents[name] = {};
        this.internalEvents[name]["" + ((undefined === listenerId) ? 0 : listenerId)] = callback;
    };
    baseEditorsApi.prototype.detachEvent = function(name, listenerId)
    {
        if (!this.internalEvents.hasOwnProperty(name))
        	return;
        var obj = this.internalEvents[name];
        var prop = "" + ((undefined === listenerId) ? 0 : listenerId);
        if (obj[prop])
        	delete obj[prop];
        if (0 === Object.getOwnPropertyNames(obj).length)
        	delete this.internalEvents[name];
    };
    baseEditorsApi.prototype.sendInternalEvent = function()
	{
        var name = arguments[0];
        if (this.internalEvents.hasOwnProperty(name))
        {
            var obj = this.internalEvents[name];
            for (var prop in obj)
			{
				obj[prop].apply(this || window, Array.prototype.slice.call(arguments, 1));
			}
        }
        return false;
	};

    // это для отложенных действий на клике (диалоги в мобильной версии)
	baseEditorsApi.prototype.setHandlerOnClick = function(handler)
	{
		this._handlerOnClick = handler;
	};
	baseEditorsApi.prototype.getHandlerOnClick = function()
	{
		return this._handlerOnClick;
	};
	// ---------------------------------------------------- interface events ---------------------------------------------
	baseEditorsApi.prototype["asc_onShowPopupWindow"] = function()
	{
		this.hideVideoControl();
	};


	baseEditorsApi.prototype.getPluginContextMenuInfo = function()
	{
		return new AscCommon.CPluginCtxMenuInfo();
	};

	// context menu items
	baseEditorsApi.prototype["onPluginContextMenuShow"] = function()
	{
		let contextMenuInfo = this.getPluginContextMenuInfo();
		let plugins = window.g_asc_plugins.onPluginEvent("onContextMenuShow", contextMenuInfo);
		if (0 === plugins.length)
		{
			if (this.contextMenuPlugins)
				delete this.contextMenuPlugins;
			return true;
		}

		if (!this.contextMenuPlugins)
		{
			this.contextMenuPlugins = {
				items : {},
				itemsCount : 0,
				processItems : {},
				processItemsCount : 0
			};
		}

		for (let i = 0, len = plugins.length; i < len; i++)
		{
			if (!this.contextMenuPlugins.processItems[plugins[i]])
			{
				this.contextMenuPlugins.processItems[plugins[i]] = 0;
				this.contextMenuPlugins.processItemsCount++;
			}
			++this.contextMenuPlugins.processItems[plugins[i]];
		}

		return false;
	};
	baseEditorsApi.prototype["onPluginContextMenuItemClick"] = function(guid, itemId)
	{
		let guids = {}; guids[guid] = true;
		window.g_asc_plugins.onPluginEvent2("onContextMenuClick", itemId, guids);
	};
	baseEditorsApi.prototype.onPluginCloseContextMenuItem = function(guid)
	{
		if (!this.contextMenuPlugins)
			return;

		let isUpdate = false;
		if (this.contextMenuPlugins.processItems[guid])
		{
			delete this.contextMenuPlugins.processItems[guid];
			this.contextMenuPlugins.processItemsCount--;
			isUpdate = true;
		}

		if (this.contextMenuPlugins.items[guid])
		{
			delete this.contextMenuPlugins.items[guid];
			this.contextMenuPlugins.itemsCount--;
			isUpdate = true;
		}

		if (isUpdate)
			this.onPluginCheckContextMenuItems();
	};
	baseEditorsApi.prototype.onPluginAddContextMenuItem = function(items)
	{
		if (!this.contextMenuPlugins)
			return;

		let guid = items["guid"];
		if (this.contextMenuPlugins.processItems[guid])
		{
			this.contextMenuPlugins.processItems[guid]--;
			// не копим.
			this.contextMenuPlugins.processItems[guid] = 0;
			if (this.contextMenuPlugins.processItems[guid] === 0)
			{
				delete this.contextMenuPlugins.processItems[guid];
				this.contextMenuPlugins.processItemsCount--;
			}
		}

		if (!this.contextMenuPlugins.items[guid])
			this.contextMenuPlugins.itemsCount++;

		this.contextMenuPlugins.items[guid] = items;

		this.onPluginCheckContextMenuItems();
	};
	baseEditorsApi.prototype.onPluginCheckContextMenuItems = function()
	{
		if (0 !== this.contextMenuPlugins.processItemsCount)
			return;

		let items = [];
		for (let item in this.contextMenuPlugins.items)
		{
			if (this.contextMenuPlugins.items.hasOwnProperty(item))
				items.push(this.contextMenuPlugins.items[item]);
		}

		delete this.contextMenuPlugins;
		this.sendEvent("onPluginContextMenu", items);
	};
	baseEditorsApi.prototype.onPluginUpdateContextMenuItem = function(items)
	{
		this.sendEvent("onPluginContextMenu", items);
	};

	// ---------------------------------------------------- wopi ---------------------------------------------
	baseEditorsApi.prototype.asc_wopi_renameFile = function(name) {
		var t = this;
		var callback = function(isTimeout, response) {
			if (response) {
				t.CoAuthoringApi.onMeta({'title': response['Name'] + '.' + AscCommon.GetFileExtension(t.documentTitle)});
			} else {
				t.sendEvent("asc_onError", c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
			}
		};
		if (!this.CoAuthoringApi.callPRC({'type': 'wopi_RenameFile', 'name': name}, Asc.c_nCommonRequestTime, callback)) {
			callback(false, undefined);
		}
	};

	baseEditorsApi.prototype.setOpenedAt = function(val)
	{
		this.openedAt = val;
	};

	baseEditorsApi.prototype.getImageDataFromSelection = function ()
	{
		return null;
	};

	baseEditorsApi.prototype.putImageToSelection = function (sImageSrc, nWidth, nHeight, replaceMode)
	{
	};

	// ---------------------------------------------------- oform ---------------------------------------------
	baseEditorsApi.prototype.signOform = function()
	{
		// TODO:
		// 1) делаем архив oform
		// 2) прогоняем архив через модуль для подписи. считаем хэши и делаем архив - и отдаем хмл для подписи
		// 3) вызываем плагин для подписи
		// 4) получаем подпись и информацию о подписи.
		// 5) добавляем эти данные в архив

		// TODO: проверить, есть ли плагин для подписи

		let plugin = window.g_asc_plugins ? window.g_asc_plugins.getSign() : null;
		if (!plugin)
		{
			this.sendEvent('asc_onError', c_oAscError.ID.Unknown, c_oAscError.Level.NoCritical);
			return;
		}

		let startData = new Asc.CPluginData();
		startData.setAttribute("data", "test_sign_data");

		this.asc_pluginRun(plugin.guid, 0, startData);
	};


	baseEditorsApi.prototype.retrieveFormatPainterData = function() {
		return null;
	};
	baseEditorsApi.prototype.getFormatPainter = function() {
		return this.formatPainter;
	};
	baseEditorsApi.prototype.getFormatPainterState = function() {
		return this.formatPainter.getState();
	};
	baseEditorsApi.prototype.isFormatPainterOn = function() {
		return this.formatPainter.isOn();
	};
	baseEditorsApi.prototype.isFormatPainterMultiple = function() {
		return this.formatPainter.isMultiple();
	};
	baseEditorsApi.prototype.canTurnOffFormatPainter = function() {
		return this.isFormatPainterOn() && !this.isFormatPainterMultiple();
	};
	baseEditorsApi.prototype.checkFormatPainterData = function() {
		return this.formatPainter.checkData();
	};
	baseEditorsApi.prototype.getFormatPainterData = function() {
		return this.formatPainter.data;
	};
	baseEditorsApi.prototype.clearFormatPainterData = function() {
		return this.formatPainter.clearData();
	};
	baseEditorsApi.prototype.sendPaintFormatEvent = function(_value)
	{
		let value = ( true === _value ? AscCommon.c_oAscFormatPainterState.kOn : ( false === _value ? AscCommon.c_oAscFormatPainterState.kOff : _value ) );
		this.formatPainter.putState(value);
		this.sendEvent("asc_onPaintFormatChanged", value);
		if(value === AscCommon.c_oAscFormatPainterState.kOff)
		{
			this.sendEvent('asc_onStopFormatPainter');
		}
	};
	baseEditorsApi.prototype.getEyedropperImgData = function()
	{
		return null;
	};
	baseEditorsApi.prototype.clearEyedropperImgData = function()
	{
		this.eyedropper.clearImageData();
	};
	baseEditorsApi.prototype.asc_startEyedropper = function(fEndCallback)
	{
		this.eyedropper.start(fEndCallback);
		this.stopInkDrawer()
	};
	baseEditorsApi.prototype.finishEyedropper = function()
	{
		this.eyedropper.finish();
	};
	baseEditorsApi.prototype.asc_finishEyedropper = function()
	{
		this.finishEyedropper();
	};
	baseEditorsApi.prototype.cancelEyedropper = function()
	{
		this.eyedropper.cancel();
	};
	baseEditorsApi.prototype.asc_cancelEyedropper = function()
	{
		this.cancelEyedropper();
	};
	baseEditorsApi.prototype.isEyedropperStarted = function()
	{
		return this.eyedropper.isStarted();
	};
	baseEditorsApi.prototype.getEyedropperColor = function()
	{
		return this.eyedropper.getColor();
	};
	baseEditorsApi.prototype.checkEyedropperColor = function(nX, nY)
	{
		this.eyedropper.checkColor(nX, nY);
	};
	baseEditorsApi.prototype.asc_StartDrawInk = function(oAscPen) {
		this.inkDrawer.startDraw(oAscPen);
		if(this.formatPainter.isOn()) {
			this.formatPainter.toggleState();
		}
		if(this.eyedropper.isStarted()) {
			this.eyedropper.end();
		}
	};
	baseEditorsApi.prototype.asc_StartInkEraser = function() {
		this.inkDrawer.startErase();
		if(this.formatPainter.isOn()) {
			this.formatPainter.toggleState();
		}
		if(this.eyedropper.isStarted()) {
			this.eyedropper.end();
		}
	};
	baseEditorsApi.prototype.asc_StopInkDrawer = function() {
		this.stopInkDrawer();
	};
	baseEditorsApi.prototype.stopInkDrawer = function() {
		this.inkDrawer.turnOff();
	};
	baseEditorsApi.prototype.onInkDrawerChangeState = function() {
	};
	baseEditorsApi.prototype.isInkDrawerOn = function() {
		return this.inkDrawer.isOn();
	};
	baseEditorsApi.prototype.isDrawInkMode = function() {
		return this.inkDrawer.isDraw();
	};
	baseEditorsApi.prototype.isEraseInkMode = function() {
		return this.inkDrawer.isErase();
	};
	baseEditorsApi.prototype.getInkPen = function() {
		return this.inkDrawer.getPen();
	};
	baseEditorsApi.prototype.getInkCursorType = function() {
		return this.inkDrawer.getCursorType();
	};
	
	baseEditorsApi.prototype.getSelectionState = function() {
	};
	baseEditorsApi.prototype.getSpeechDescription = function(prevState, action) {
		return null;
	};

	// methods for desktop:
	baseEditorsApi.prototype.isCloudSaveAsLocalToDrawingFormat = function(actionType, fileType) {
		if (!window["AscDesktopEditor"])
			return false;
		if (this.isLocalMode())
			return false;
		if (c_oAscAsyncAction.DownloadAs !== actionType)
			return false;

		switch (fileType)
		{
			case Asc.c_oAscFileType.PDF:
			case Asc.c_oAscFileType.PDFA:
			case Asc.c_oAscFileType.IMG:
			case Asc.c_oAscFileType.JPG:
			case Asc.c_oAscFileType.TIFF:
			case Asc.c_oAscFileType.TGA:
			case Asc.c_oAscFileType.GIF:
			case Asc.c_oAscFileType.PNG:
			case Asc.c_oAscFileType.EMF:
			case Asc.c_oAscFileType.WMF:
			case Asc.c_oAscFileType.BMP:
			case Asc.c_oAscFileType.CR2:
			case Asc.c_oAscFileType.PCX:
			case Asc.c_oAscFileType.RAS:
			case Asc.c_oAscFileType.PSD:
			case Asc.c_oAscFileType.ICO:
				window["AscDesktopEditor"]["emulateCloudPrinting"](true);
				return true;
			default:
				break;
		}
		return false;
	};

	baseEditorsApi.prototype.localSaveToDrawingFormat = function(base64Data, fileType) {
		this.sync_StartAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);

		let documentUrl = this.DocumentUrl;
		if (this.editorId === c_oEditorId.Spreadsheet)
			documentUrl = this.documentId + "/";
		window["AscDesktopEditor"]["localSaveToDrawingFormat"](this.documentTitle, documentUrl || "", this.ThemeLoader ? this.ThemeLoader.ThemesUrl : "", base64Data, fileType);
	};

	baseEditorsApi.prototype["onLocalSaveToDrawingFormat"] = function(error) {
		window["AscDesktopEditor"]["emulateCloudPrinting"](false);
		this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.Save);
		this.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.DownloadAs);
		if (0 !== error)
			this.sendEvent("asc_onError", c_oAscError.ID.ConvertationSaveError, c_oAscError.Level.NoCritical);
	};

	//----------------------------------------------------------export----------------------------------------------------
	window['AscCommon']                = window['AscCommon'] || {};
	window['AscCommon'].baseEditorsApi = baseEditorsApi;

	prot = baseEditorsApi.prototype;
	prot['asc_loadFontsFromServer'] = prot.asc_loadFontsFromServer;
	prot['asc_setRestriction'] = prot.asc_setRestriction;
	prot['asc_addRestriction'] = prot.asc_addRestriction;
	prot['asc_removeRestriction'] = prot.asc_removeRestriction;
	prot['asc_selectSearchingResults'] = prot.asc_selectSearchingResults;
	prot['asc_isSelectSearchingResults'] = prot.asc_isSelectSearchingResults;
	prot['asc_showRevision'] = prot.asc_showRevision;
	prot['asc_getAdvancedOptions'] = prot.asc_getAdvancedOptions;
	prot['asc_Print'] = prot.asc_Print;
	prot['asc_GetCurrentColorSchemeName'] = prot.asc_GetCurrentColorSchemeName;
	prot['asc_GetCurrentColorSchemeIndex'] = prot.asc_GetCurrentColorSchemeIndex;
	prot['asc_runAutostartMacroses'] = prot.asc_runAutostartMacroses;
	prot['asc_runMacros'] = prot.asc_runMacros;
	prot['asc_getAllMacrosNames'] = prot.asc_getAllMacrosNames;
	prot['asc_setVisiblePasteButton'] = prot.asc_setVisiblePasteButton;
	prot['asc_getAutoCorrectMathSymbols'] = prot.asc_getAutoCorrectMathSymbols;
	prot['asc_getAutoCorrectMathFunctions'] = prot.asc_getAutoCorrectMathFunctions;
	prot['asc_resetToDefaultAutoCorrectMathSymbols'] = prot.asc_resetToDefaultAutoCorrectMathSymbols;
	prot['asc_resetToDefaultAutoCorrectMathFunctions'] = prot.asc_resetToDefaultAutoCorrectMathFunctions;
	prot['asc_deleteFromAutoCorrectMathSymbols'] = prot.asc_deleteFromAutoCorrectMathSymbols;
	prot['asc_deleteFromAutoCorrectMathFunctions'] = prot.asc_deleteFromAutoCorrectMathFunctions;
	prot['asc_AddOrEditFromAutoCorrectMathSymbols'] = prot.asc_AddOrEditFromAutoCorrectMathSymbols;
	prot['asc_AddFromAutoCorrectMathFunctions'] = prot.asc_AddFromAutoCorrectMathFunctions;
	prot['asc_refreshOnStartAutoCorrectMathSymbols'] = prot.asc_refreshOnStartAutoCorrectMathSymbols;
	prot['asc_refreshOnStartAutoCorrectMathFunctions'] = prot.asc_refreshOnStartAutoCorrectMathFunctions;
	prot['asc_updateFlagAutoCorrectMathSymbols'] = prot.asc_updateFlagAutoCorrectMathSymbols;
	prot['asc_AddMath'] = prot.asc_AddMath;
	prot['asc_ConvertMathView'] = prot.asc_ConvertMathView;
	prot['asc_GetMathInputType'] = prot.asc_GetMathInputType;
	prot['asc_SetMathInputType'] = prot.asc_SetMathInputType;
	prot['asc_initShortcuts'] = prot.asc_initShortcuts;
	prot['asc_getShortcutAction'] = prot.asc_getShortcutAction;
	prot['asc_removeShortcuts'] = prot.asc_removeShortcuts;
	prot['asc_addCustomShortcutInsertSymbol'] = prot.asc_addCustomShortcutInsertSymbol;
	prot['asc_wopi_renameFile'] = prot.asc_wopi_renameFile;
	prot['asc_getInformationBetweenFrameAndGeneralEditor'] = prot.asc_getInformationBetweenFrameAndGeneralEditor;
	prot['asc_setShapeNames'] = prot.asc_setShapeNames;
	prot['asc_generateChartPreviews'] = prot.asc_generateChartPreviews;
	prot['asc_createSmartArt'] = prot.asc_createSmartArt;
	prot['asc_generateSmartArtPreviews'] = prot.asc_generateSmartArtPreviews;
	prot['asc_addTableOleObject'] = prot.asc_addTableOleObject;
	prot['asc_editTableOleObject'] = prot.asc_editTableOleObject;
	prot['asc_canEditTableOleObject'] = prot.asc_canEditTableOleObject;
	prot['asc_EditSelectAll'] = prot.asc_EditSelectAll;
	prot['setOpenedAt'] = prot.setOpenedAt;
	prot['asc_SaveDrawingAsPicture'] = prot.asc_SaveDrawingAsPicture;
	prot['AddImageUrl'] = prot.AddImageUrl;
	prot['asc_SetDrawImagePreviewBulletForMenu'] = prot['SetDrawImagePreviewBulletForMenu'] = prot.SetDrawImagePreviewBulletForMenu;
	prot['asc_GetPossibleNumberingLanguage'] = prot.asc_GetPossibleNumberingLanguage;
	prot['asc_isCrypto'] = prot.asc_isCrypto;
	prot['asc_startEyedropper'] = prot.asc_startEyedropper;
	prot['asc_finishEyedropper'] = prot.asc_finishEyedropper;
	prot['asc_cancelEyedropper'] = prot.asc_cancelEyedropper;
	prot['asc_StartDrawInk'] = prot.asc_StartDrawInk;
	prot['asc_StartInkEraser'] = prot.asc_StartInkEraser;
	prot['asc_StopInkDrawer'] = prot.asc_StopInkDrawer;
	prot['startGetDocInfo'] = prot.startGetDocInfo;
	prot['stopGetDocInfo'] = prot.stopGetDocInfo;
	prot["can_CopyCut"] = prot.can_CopyCut;
	prot["asc_searchEnabled"] = prot.asc_searchEnabled;
	prot['asc_findText'] = prot.asc_findText;
	prot['asc_endFindText'] = prot.asc_endFindText;

})(window);
