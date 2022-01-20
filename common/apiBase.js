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

(function(window, undefined)
{
	var prot;

	// Import
	var c_oEditorId = AscCommon.c_oEditorId;
	var c_oCloseCode = AscCommon.c_oCloseCode;
	var DownloadType = AscCommon.DownloadType;

	var c_oAscError           = Asc.c_oAscError;
	var c_oAscAsyncAction     = Asc.c_oAscAsyncAction;
	var c_oAscAsyncActionType = Asc.c_oAscAsyncActionType;

	/** @constructor */
	function baseEditorsApi(config, editorId)
	{
		this.editorId      = editorId;
		this.isLoadFullApi = false;
        this.isLoadFonts = false;
		this.openResult    = null;

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

		this.MathMenuLoad          = false;

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

		this.canSave    = true;        // Флаг нужен чтобы не происходило сохранение пока не завершится предыдущее сохранение
		this.IsUserSave = false;    // Флаг, контролирующий сохранение было сделано пользователем или нет (по умолчанию - нет)
		this.isForceSaveOnUserSave = false;
        this.forceSaveButtonTimeout = null;
        this.forceSaveButtonContinue = false;
        this.forceSaveTimeoutTimeout = null;
		this.forceSaveForm = null;
		this.disconnectOnSave = null;
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

		this.isLockTargetUpdate = false;

		this.lastWorkTime = 0;

		this.signatures = [];

		this.currentPassword = "";

		this.disableAutostartMacros = false;
		this.macros = null;
		this.vbaMacros = null;

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

		this.Shortcuts = new AscCommon.CShortcuts();
		this.initDefaultShortcuts();

		this.isUseNativeViewer = true;

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

        AscFonts.load(t, function()
        {
            t.isLoadFonts = true;
            t.onEndLoadFile(null);
        }, function() {
            t.sendEvent("asc_onError", Asc.c_oAscError.ID.LoadingScriptError, c_oAscError.Level.Critical);
        });

		AscCommon.loadChartStyles(function() {}, function(err) {
			t.sendEvent("asc_onError", Asc.c_oAscError.ID.LoadingScriptError, c_oAscError.Level.NoCritical);
		});

		var oldOnError = window.onerror;
		window.onerror = function(errorMsg, url, lineNumber, column, errorObj) {
			//send only first error to reduce number of requests. also following error may be consequences of first
			window.onerror = oldOnError;
			var msg = 'Error: ' + errorMsg + ' Script: ' + url + ' Line: ' + lineNumber + ':' + column +
				' userAgent: ' + (navigator.userAgent || navigator.vendor || window.opera) + ' platform: ' +
				navigator.platform + ' isLoadFullApi: ' + t.isLoadFullApi + ' isDocumentLoadComplete: ' +
				t.isDocumentLoadComplete + ' StackTrace: ' + (errorObj ? errorObj.stack : "");
			t.CoAuthoringApi.sendChangesError(msg);
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

			if (this.documentOpenOptions && this.documentOpenOptions["watermark"])
			{
				this.watermarkDraw = new AscCommon.CWatermarkOnDraw(this.documentOpenOptions["watermark"], this);
				this.watermarkDraw.CheckParams(this);
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
	// Просмотр PDF
	baseEditorsApi.prototype.isPdfViewer                     = function()
	{
		return false;
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
	baseEditorsApi.prototype.asc_setRestriction              = function(val)
	{
		this.restrictions = val;
		this.onUpdateRestrictions();
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
	baseEditorsApi.prototype.isLongAction                    = function()
	{
		return (0 !== this.IsLongActionCurrent);
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

		if (!this.isLongAction())
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
		if (this.isLongAction())
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
				res = !this.isPdfViewer();
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
			rData = {
				"c"             : 'open',
				"id"            : this.documentId,
				"userid"        : this.documentUserId,
				"format"        : this.documentFormat,
				"url"           : this.documentUrl,
				"title"         : this.documentTitle,
				"lcid"          : locale,
				"nobase64"      : true
			};

			if (this.isUseNativeViewer)
			{
				switch (this.documentFormat)
				{
					case "pdf":
					case "xps":
					case "djvu":
						rData["convertToOrigin"] = true;
						break;
					default:
						break;
				}
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
		AscCommon.openFileCommand(this.documentId, data, this.documentUrlChanges, this.documentTokenChanges, AscCommon.c_oSerFormat.Signature, function(error, result)
		{
			var signature = result.data && String.fromCharCode(result.data[0], result.data[1], result.data[2], result.data[3]);
			if (error || (!result.bSerFormat && (t.editorId !== c_oEditorId.Word || 'XLSY' === signature || 'PPTY' === signature)))
			{
				t.sync_EndAction(c_oAscAsyncActionType.BlockInteraction, c_oAscAsyncAction.Open);
				var err = error ? c_oAscError.ID.Unknown : c_oAscError.ID.ConvertationOpenError;
				t.sendEvent("asc_onError",  err, c_oAscError.Level.Critical);
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
	baseEditorsApi.prototype.asyncServerIdEndLoaded              = function()
	{
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

		if (window.g_asc_plugins)
            window.g_asc_plugins.onPluginEvent("onDocumentContentReady");

        if (c_oEditorId.Spreadsheet === this.editorId) {
			this.onUpdateDocumentModified(this.asc_isDocumentModified());
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
		if (this.canSave && !this.isViewMode && (this.canUnlockDocument || 0 !== this.autoSaveGap)) {
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
			this.CoAuthoringApi.sendChangesError(errorMsg);
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
		this._coAuthoringInit();
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

				if (undefined !== this.licenseResult['protectionSupport']) {
					this.isProtectionSupport = this.licenseResult['protectionSupport'];
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
			if (t.isRestrictionView()) {
				t.sync_EndAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.Disconnect);
				t.asc_removeRestriction(Asc.c_oAscRestrictionType.View);
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
		this.CoAuthoringApi.onFirstLoadChangesEnd     = function()
		{
			t.asyncServerIdEndLoaded();
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
			t.licenseResult   = res;
			t.isOnLoadLicense = true;
			var oResult = new AscCommon.asc_CAscEditorPermissions();
			oResult.setLicenseType(res);
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
				t.sendEvent('asc_onError', error, level);
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
				if (t.isRestrictionView()) {
					t.sync_EndAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.Disconnect);
					t.asc_removeRestriction(Asc.c_oAscRestrictionType.View);
				}
				var error = AscCommon.getDisconnectErrorCode(t.isDocumentLoadComplete, opt_closeCode);
				var level = t.isDocumentLoadComplete ? Asc.c_oAscError.Level.NoCritical : Asc.c_oAscError.Level.Critical;
				t.setViewModeDisconnect(AscCommon.getEnableDownloadByCloseCode(opt_closeCode));
				t.sendEvent('asc_onError', error, level);
			} else if (!t.isRestrictionView()) {
				t.sync_StartAction(Asc.c_oAscAsyncActionType.Information, Asc.c_oAscAsyncAction.Disconnect);
				t.asc_addRestriction(Asc.c_oAscRestrictionType.View);
			}
		};
		this.CoAuthoringApi.onDocumentOpen = function (inputWrap) {
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
								var urls = input["data"];
								AscCommon.g_oDocumentUrls.init(urls);
								var documentUrl = urls['Editor.bin'];
								if (t.isUseNativeViewer && !documentUrl)
									documentUrl = urls['origin.pdf'] || urls['origin.xps'] || urls['origin.djvu'];
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
				this.sendEvent('asc_onSpellCheckInit', [
                    "1026",
                    "1027",
                    "1029",
                    "1030",
                    "1031",
                    "1032",
                    "1033",
                    "1036",
                    "1038",
                    "1040",
                    "1042",
                    "1043",
                    "1044",
                    "1045",
                    "1046",
                    "1048",
                    "1049",
                    "1050",
                    "1051",
                    "1053",
                    "1055",
                    "1057",
                    "1058",
                    "1060",
                    "1062",
                    "1063",
                    "1066",
                    "1068",
                    "1069",
                    "1087",
                    "1104",
                    "1110",
                    "1134",
                    "2051",
                    "2055",
                    "2057",
                    "2068",
                    "2070",
                    "3079",
                    "3081",
                    "3082",
                    "4105",
                    "7177",
                    "9242",
                    "10266",
                    "2067"
				]);
			}
		} else {
			if (!this.SpellCheckUrl && !window['NATIVE_EDITOR_ENJINE']) {
				this.SpellCheckApi = {};
				this.SpellCheckApi.log = false;
				this.SpellCheckApi.worker = new CSpellchecker({
					api: this,
					enginePath: "../../../../sdkjs/common/spell/spell",
					dictionariesPath: "./../../../../dictionaries"
				});
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
	baseEditorsApi.prototype.endInsertDocumentUrls              = function ()
	{
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
		if (Asc.c_oAscFileType.HTML === options.fileType && null == options.oDocumentMailMerge && null == options.oMailMergeSendData) {
			options.fileType = Asc.c_oAscFileType.HTML_TODO;
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
		}
		if (options.textParams) {
			oAdditionalData["textParams"] = {"association": options.textParams.asc_getAssociation()};
		}

		if (this._downloadAs(actionType, options, oAdditionalData, dataContainer, downloadType))
		{
			return;
		}

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
	// Images & Charts & TextArts
	baseEditorsApi.prototype.asc_getChartPreviews                = function(chartType)
	{
		return this.chartPreviewManager.getChartPreviews(chartType);
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
			}, false, undefined, token);
		} else {
			callback.call(this, urls);
		}
	};
	baseEditorsApi.prototype._addImageUrl                        = function()
	{
	};
	baseEditorsApi.prototype.asc_addImage                        = function(obj)
	{
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
		}, this.editorId === c_oEditorId.Spreadsheet);
	};

	baseEditorsApi.prototype.asc_addOleObject = function(oPluginData)
	{
		if(this.isViewMode){
			return;
		}
		var oThis      = this;
		var sImgSrc    = oPluginData["imgSrc"];
		var nWidthPix  = oPluginData["widthPix"];
		var nHeightPix = oPluginData["heightPix"];
		var fWidth     = oPluginData["width"];
		var fHeight    = oPluginData["height"];
		var sData      = oPluginData["data"];
		var sGuid      = oPluginData["guid"];
		if (typeof sImgSrc === "string" && sImgSrc.length > 0 && typeof sData === "string"
			&& typeof sGuid === "string" && sGuid.length > 0
			&& AscFormat.isRealNumber(nWidthPix) && AscFormat.isRealNumber(nHeightPix)
			&& AscFormat.isRealNumber(fWidth) && AscFormat.isRealNumber(fHeight)
		)

		this.asc_checkImageUrlAndAction(sImgSrc, function(oImage)
		{
			oThis.asc_addOleObjectAction(AscCommon.g_oDocumentUrls.getImageLocal(oImage.src), sData, sGuid, fWidth, fHeight, nWidthPix, nHeightPix);
		});
	};

	baseEditorsApi.prototype.asc_editOleObject = function(oPluginData)
	{
		if(this.isViewMode){
			return;
		}
		var oThis      = this;
		var bResize    = oPluginData["resize"];
		var sImgSrc    = oPluginData["imgSrc"];
		var oOleObject = AscCommon.g_oTableId.Get_ById(oPluginData["objectId"]);
		var nWidthPix  = oPluginData["widthPix"];
		var nHeightPix = oPluginData["heightPix"];
		var sData      = oPluginData["data"];
		if (typeof sImgSrc === "string" && sImgSrc.length > 0 && typeof sData === "string"
			&& oOleObject && AscFormat.isRealNumber(nWidthPix) && AscFormat.isRealNumber(nHeightPix))
		{
            this.asc_checkImageUrlAndAction(sImgSrc, function(oImage)
			{
				oThis.asc_editOleObjectAction(bResize, oOleObject, AscCommon.g_oDocumentUrls.getImageLocal(oImage.src), sData, nWidthPix, nHeightPix);
			});
		}
	};

	baseEditorsApi.prototype.asc_addOleObjectAction = function(sLocalUrl, sData, sApplicationId, fWidth, fHeight)
	{
	};

	baseEditorsApi.prototype.asc_editOleObjectAction = function(bResize, oOleObject, sImageUrl, sData, nPixWidth, nPixHeight)
	{
	};

	baseEditorsApi.prototype.asc_selectSearchingResults = function(value)
	{
		if (this.selectSearchingResults === value)
		{
			return;
		}
		this.selectSearchingResults = value;
		this._selectSearchingResults(value);
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
		return AscCommon.History.Can_Undo();
	};
	baseEditorsApi.prototype.asc_getCanRedo = function()
	{
		return AscCommon.History.Can_Redo();
	};
	// Offline mode
	baseEditorsApi.prototype.asc_isOffline  = function()
	{
		return (window.location.protocol.indexOf("file") == 0) ? true : false;
	};
	baseEditorsApi.prototype.asc_getUrlType = function(url)
	{
		return AscCommon.getUrlType(url);
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
		if (this.isLoadFullApi && this.DocInfo && this.openResult && this.isLoadFonts)
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

                    if (!transform)
                        window["AscDesktopEditor"]["MediaStart"](sMediaName, pos.X, pos.Y, extX, extY, this.WordControl.m_nZoomValue / 100);
                    else
                        window["AscDesktopEditor"]["MediaStart"](sMediaName, pos.X, pos.Y, extX, extY, this.WordControl.m_nZoomValue / 100, transform.sx, transform.shy, transform.shx, transform.sy, transform.tx, transform.ty);
                }
                else
				{
					var transition = this.WordControl.DemonstrationManager.Transition;
                    if ((manager.SlideNum >= 0 && manager.SlideNum < manager.SlidesCount) && (!transition || !transition.IsPlaying()))
                    {
                        var _w = transition.Rect.w;
                        var _h = transition.Rect.h;
                        var _w_mm = manager.HtmlPage.m_oLogicDocument.GetWidthMM();
                        var _h_mm = manager.HtmlPage.m_oLogicDocument.GetHeightMM();

                        var _x = transition.Rect.x;
                        if (this.isReporterMode)
						{
                            _x += ((this.WordControl.m_oMainParent.AbsolutePosition.L * AscCommon.g_dKoef_mm_to_pix) >> 0);
						}

                        var _zoom = _w / (_w_mm * AscCommon.g_dKoef_mm_to_pix);

                        if (!transform)
                            window["AscDesktopEditor"]["MediaStart"](sMediaName, _x, transition.Rect.y, extX, extY, _zoom);
                        else
                            window["AscDesktopEditor"]["MediaStart"](sMediaName, _x, transition.Rect.y, extX, extY, _zoom, transform.sx, transform.shy, transform.shx, transform.sy, transform.tx, transform.ty);
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
			this.pluginsManager.register(basePath, plugins);
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
	baseEditorsApi.prototype.asc_pluginButtonClick = function(id)
	{
		if (null != this.pluginsManager)
			this.pluginsManager.buttonClick(id);
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
						this.WordControl.m_oLogicDocument.AddTextWithPr(new AscCommon.CUnicodeStringEmulator(arrCharCodes), textPr, true);
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

    baseEditorsApi.prototype["asc_registerPlaceholderCallback"] = function(type, callback)
    {
    	if (this.WordControl && this.WordControl.m_oDrawingDocument && this.WordControl.m_oDrawingDocument.placeholders)
		{
            this.WordControl.m_oDrawingDocument.placeholders.registerCallback(type, callback);
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
		// Disable history for builder
		AscCommon.History.TurnOff();
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

	// drop emulation
    baseEditorsApi.prototype.privateDropEvent = function(obj)
    {
        if (!obj || !obj.type)
            return;

        var e = {
            pageX : obj["x"],
            pageY : obj["y"]
        };

        switch (obj.type)
        {
            case "onbeforedrop":
            {
                this.beginInlineDropTarget(e);
                break;
            }
            case "ondrop":
            {
                this.endInlineDropTarget(e);

                if (obj["html"])
                    this["pluginMethod_PasteHtml"](obj["html"]);
                else if (obj["text"])
                    this["pluginMethod_PasteText"](obj["text"]);

                break;
            }
            default:
                break;
        }
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

		return new Date().getTime() - this.lastWorkTime;
	};

	baseEditorsApi.prototype.checkInterfaceElementBlur = function()
	{
		if (!document.activeElement || !document.createEvent || (document.activeElement.id === "area_id"))
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

    	if (!this.asc_canPaste())
    		return;

    	this._beforeEvalCommand();
		this.macros.runAuto();
		this._afterEvalCommand(undefined);
    };
	baseEditorsApi.prototype.asc_runMacros = function(sGuid)
    {
    	if (!this.macros)
    		return;

    	if (!this.asc_canPaste())
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
	};
	/**
	 * Reset to default version math autocorrect functions list
	 */
	baseEditorsApi.prototype.asc_resetToDefaultAutoCorrectMathFunctions = function()
	{
		window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathFuncs = JSON.parse(JSON.stringify(window['AscCommonWord'].g_AutoCorrectMathsList.DefaultAutoCorrectMathFuncs));
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
	};
	/**
	 * Add item from g_AutoCorrectMathFuncs
	 * @param {string} newEl
	 */
	baseEditorsApi.prototype.asc_AddFromAutoCorrectMathFunctions = function(newEl)
	{
		window['AscCommonWord'].g_AutoCorrectMathsList.AutoCorrectMathFuncs.push(newEl);
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
	baseEditorsApi.prototype.asc_setSkin = function(obj)
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
	//---------------------------------------------------------version----------------------------------------------------
	baseEditorsApi.prototype["GetVersion"] = baseEditorsApi.prototype.GetVersion = function()
	{
		var ver = "@@Version";
		return (ver === "0.0.0" || ver.substr(2) === "Version") ? "develop" : ver;
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
    baseEditorsApi.prototype["attachEvent"] = function(name, callback, listenerId)
    {
    	if (!this.internalEvents.hasOwnProperty(name))
            this.internalEvents[name] = {};
        this.internalEvents[name]["" + ((undefined === listenerId) ? 0 : listenerId)] = callback;
    };
    baseEditorsApi.prototype["detachEvent"] = function(name, listenerId)
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

	//----------------------------------------------------------export----------------------------------------------------
	window['AscCommon']                = window['AscCommon'] || {};
	window['AscCommon'].baseEditorsApi = baseEditorsApi;

	prot = baseEditorsApi.prototype;
	prot['asc_loadFontsFromServer'] = prot.asc_loadFontsFromServer;
	prot['asc_setRestriction'] = prot.asc_setRestriction;
	prot['asc_selectSearchingResults'] = prot.asc_selectSearchingResults;
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
	prot['asc_initShortcuts'] = prot.asc_initShortcuts;
	prot['asc_getShortcutAction'] = prot.asc_getShortcutAction;
	prot['asc_removeShortcuts'] = prot.asc_removeShortcuts;
	prot['asc_addCustomShortcutInsertSymbol'] = prot.asc_addCustomShortcutInsertSymbol;
	prot['asc_wopi_renameFile'] = prot.asc_wopi_renameFile;
	prot['asc_setShapeNames'] = prot.asc_setShapeNames;

	prot['asc_isCrypto'] = prot.asc_isCrypto;

})(window);
