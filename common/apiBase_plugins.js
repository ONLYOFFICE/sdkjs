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
    /**
     * Plugin config.
     * @typedef {Object} Config
	 *
     * @property {string} name
	 * Plugin name which will be visible at the plugin toolbar.
     *
	 * @property {?Object} nameLocale
	 * Translations for the name field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the plugin name translation for each language. Example: { "fr" : "trFr", "de" : "deTr" }.
     *
	 * @property {string} guid
	 * Plugin identifier. It must be of the asc.{UUID} type.
	 *
	 * @property {?string} [baseUrl=""]
	 * Path to the plugin. All the other paths are calculated relative to this path. In case the plugin is installed on the server, an additional parameter (path to the plugins) is added there. If baseUrl == "", the path to all plugins will be used.
     *
	 * @property {Variation[]} variations
	 * Plugin variations, or subplugins, that are created inside the origin plugin.
     */

    /**
     * The editors which the plugin is available for ("word" - text document editor, "cell" - spreadsheet editor, "slide" - presentation editor).
     * @typedef {("word" | "slide" | "cell")} EditorType
     */

    /**
	 * The data type selected in the editor and sent to the plugin: "text" - the text data, "html" - HTML formatted code, "ole" - OLE object data, "desktop" - the desktop editor data, "destop-external" - the main page data of the desktop app (system messages), "none" - no data will be send to the plugin from the editor.
	 * @typedef {("none" | "text" | "ole" | "html" | "desktop")} InitDataType
     */

    /**
     * Plugin event.
     * @typedef {("onDocumentContentReady" | "onTargetPositionChanged" | "onClick" | "onInputHelperClear" | "onInputHelperInput")} EventType
     */

    /**
     * Plugin variations, or subplugins, that are created inside the origin plugin.
     * @typedef {Object} Variation
     * @property {string} description
	 * The description, i.e. what describes your plugin the best way.
	 *
     * @property {?Object} descriptionLocale
	 * Translations for the description field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the plugin description translation for each language.
	 *
     * @property {string} url
	 * Plugin entry point, i.e. an HTML file which connects the plugin.js file (the base file needed for work with plugins) and launches the plugin code.
	 *
     * @property {string[]} icons (with support HiDPI)
	 * Plugin icon image files used in the editors. There can be several scaling types for plugin icons: 100%, 125%, 150%, 175%, 200%, etc.
	 *
     * @property {?boolean} [isViewer=true]
	 * Specifies if the plugin is working when the document is available in the viewer mode only or not.
	 *
	 * @property {boolean} isDisplayedInViewer
	 * Specifies if the plugin will be displayed in the viewer mode as well as in the editor mode (isDisplayedInViewer == true) or in the editor mode only (isDisplayedInViewer == false).
	 *
     * @property {EditorType[]} EditorsSupport
	 * The editors which the plugin is available for ("word" - text document editor, "cell" - spreadsheet editor, "slide" - presentation editor).
	 *
     * @property {boolean} [isVisual=true]
	 * Specifies if the plugin is visual (will open a window for some action, or introduce some additions to the editor panel interface) or non-visual (will provide a button (or buttons) which is going to apply some transformations or manipulations to the document).
     *
	 * @property {boolean} [isModal=true]
	 * Specifies if the opened plugin window is modal (used for visual plugins only, and if isInsideMode is not true).
	 *
	 * @property {boolean} [isInsideMode=false]
	 * Specifies if the plugin must be displayed inside the editor panel instead of its own window.
     *
	 * @property {boolean} isCustomWindow
	 * Specifies if the plugin uses a custom window, without standard borders and buttons (used for modal plugins only).
	 *
	 * @property {boolean} isSystem
	 * Specifies if the plugin is not displayed in the editor interface and is started in the background with the server (or desktop editors start) not interfering with the other plugins, so that they can work simultaneously.
     *
	 * @property {InitDataType} initDataType
	 * The data type selected in the editor and sent to the plugin: "text" - the text data, "html" - HTML formatted code, "ole" - OLE object data, "desktop" - the desktop editor data, "destop-external" - the main page data of the desktop app (system messages), "none" - no data will be send to the plugin from the editor.
     *
	 * @property {string} initData
     * Is usually equal to "" - this is the data which is sent from the editor to the plugin at the plugin start (e.g. if initDataType == "text", the plugin will receive the selected text when run). It may also be equal to encryption in the encryption plugins.
	 *
	 * @property {?boolean} isUpdateOleOnResize
	 * Specifies if an OLE object must be redrawn when resized in the editor using the vector object draw type or not (used for OLE objects only, i.e. initDataType == "ole").
     *
	 * @property {?Button[]} buttons
	 * The list of skinnable plugin buttons used in the plugin interface (used for visual plugins with their own window only, i.e. isVisual == true && isInsideMode == false).
	 *
	 * @property {?boolean} [initOnSelectionChanged=true]
	 * Specifies if the plugin watches the text selection events in the editor window.
	 *
	 * @property {number[]} size
	 * Plugin window size.
	 *
	 * @property {EventType[]} events
     * Plugin events.
     */

    /**
	 * The skinnable plugin button used in the plugin interface (used for visual plugins with their own window only, i.e. isVisual == true && isInsideMode == false).
	 * @typedef {object} Button
	 * @property {string} text - The label which is displayed on the button.
	 * @property {string} textLocale - Translations for the text field. The object keys are the two letter language codes (ru, de, it, etc.) and the values are the button label translation for each language.
	 * @property {boolean} primary - Defines if the button is primary or not. The primary flag affects the button skin only.
	 */

    /**
     * Base class
     * @global
     * @class
     * @name Api
     */

    var Api = window["AscCommon"].baseEditorsApi;

    /**
     * Returns the editor version.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetVersion
     * @returns {string} - Editor version.
     */
    Api.prototype["pluginMethod_GetVersion"] = function() { return this.GetVersion(); };

    /**
     * Adds an OLE object to the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias AddOleObject
	 * @this Api
     * @param {Object} data - The OLE object properties.
     * @param {number} data.width - The object width measured in millimeters.
     * @param {number} data.height - The object height measured in millimeters.
     * @param {string} data.data - OLE object data (internal format).
     * @param {string} data.guid - An OLE object program identifier which must be of the asc.{UUID} type.
     * @param {string} data.imgSrc - A link to the image (its visual representation) stored in the OLE object and used by the plugin.
     * @param {number} data.widthPix - The image width in pixels.
     * @param {number} data.heightPix - The image height in pixels.
    */
    Api.prototype["pluginMethod_AddOleObject"] = function(data) { return this.asc_addOleObject(data); };

    /**
     * Edits an OLE object in the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias EditOleObject
     * @param {Object} data - The OLE object properties.
     * @param {number} data.width - The object width measured in millimeters.
     * @param {number} data.height - The object height measured in millimeters.
     * @param {string} data.data - OLE object data (internal format).
     * @param {string} data.objectId - An OLE object identifier.
     * @param {string} data.imgSrc - A link to the image (its visual representation) stored in the OLE object and used by the plugin.
     * @param {number} data.widthPix - The image width in pixels.
     * @param {number} data.heightPix - The image height in pixels.
     */
    Api.prototype["pluginMethod_EditOleObject"] = function(data) { return this.asc_editOleObject(data); };

    /**
	 * An object containing the font information.
     * @typedef {Object} FontInfo
     * @property {string} m_wsFontName
     * @property {string} m_wsFontPath
	 * @property {number} m_lIndex
	 * @property {boolean} m_bBold
	 * @property {boolean} m_bItalic
	 * @property {boolean} m_bIsFixed
	 * @property {number[]} m_aPanose
	 * @property {number} m_ulUnicodeRange1
	 * @property {number} m_ulUnicodeRange2
	 * @property {number} m_ulUnicodeRange3
	 * @property {number} m_ulUnicodeRange4
	 * @property {number} m_ulCodePageRange1
	 * @property {number} m_ulCodePageRange2
	 * @property {number} m_usWeigth
	 * @property {number} m_usWidth
	 * @property {string} m_sFamilyClass
	 * @property {string} m_eFontFormat
	 * @property {number} m_shAvgCharWidth
	 * @property {number} m_shAscent
	 * @property {number} m_shDescent
	 * @property {number} m_shLineGap
	 * @property {number} m_shXHeight
	 * @property {number} m_shCapHeight
     */

    /**
     * Returns the fonts list.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetFontList
     * @returns {FontInfo[]} - An array of the FontInfo objects.
     */
    Api.prototype["pluginMethod_GetFontList"] = function()
    {
        return AscFonts.g_fontApplication.g_fontSelections.SerializeList();
    };

    /**
     * Inserts text into the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias InputText
	 * @param {string} text - A string value that specifies the text to be inserted into the document.
	 * @param {string} textReplace - A string value that specifies the text to be replaced with a new text.
     */
    Api.prototype["pluginMethod_InputText"] = function(text, textReplace)
    {
        if (this.isViewMode || !AscCommon.g_inputContext)
            return;

        var codes = [];
        for (var i = text.getUnicodeIterator(); i.check(); i.next())
            codes.push(i.value());

        if (textReplace)
        {
            for (var i = 0; i < textReplace.length; i++)
                AscCommon.g_inputContext.emulateKeyDownApi(8);
        }

        AscCommon.g_inputContext.apiInputText(codes);
        AscCommon.g_inputContext.keyPressInput = "";
    };

    /**
     * Pastes text in the HTML format into the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias PasteHtml
     * @param {string} htmlText - A string value that specifies the text in the HTML format to be pasted into the document.
     */
    Api.prototype["pluginMethod_PasteHtml"] = function(htmlText)
    {
        if (!AscCommon.g_clipboardBase)
            return null;

		if (this.isViewMode)
			return null;

        var _elem = document.getElementById("pmpastehtml");
        if (_elem)
            return;

        _elem = document.createElement("div");
        _elem.id = "pmpastehtml";

        if (this.editorId == AscCommon.c_oEditorId.Word || this.editorId == AscCommon.c_oEditorId.Presentation)
        {
            var textPr = this.get_TextProps();
            if (textPr)
            {
                if (undefined !== textPr.TextPr.FontSize)
                    _elem.style.fontSize = textPr.TextPr.FontSize + "pt";

                _elem.style.fontWeight = (true === textPr.TextPr.Bold) ? "bold" : "normal";
                _elem.style.fontStyle = (true === textPr.TextPr.Italic) ? "italic" : "normal";

                var _color = textPr.TextPr.Color;
                if (_color)
                    _elem.style.color = "rgb(" + _color.r + "," + _color.g + "," + _color.b + ")";
                else
                    _elem.style.color = "rgb(0,0,0)";
            }
        }
        else if (this.editorId == AscCommon.c_oEditorId.Spreadsheet)
        {
            var props = this.asc_getCellInfo();

            if (props && props.font)
            {
                if (undefined != props.font.size)
                    _elem.style.fontSize = props.font.size + "pt";

                _elem.style.fontWeight = (true === props.font.bold) ? "bold" : "normal";
                _elem.style.fontStyle = (true === props.font.italic) ? "italic" : "normal";
            }
        }

        _elem.innerHTML = htmlText;
        document.body.appendChild(_elem);
        this.incrementCounterLongAction();
        var b_old_save_format = AscCommon.g_clipboardBase.bSaveFormat;
        AscCommon.g_clipboardBase.bSaveFormat = true;
        this.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.HtmlElement, _elem);
        this.decrementCounterLongAction();

        if (true)
        {
            var fCallback = function ()
            {
                document.body.removeChild(_elem);
                _elem = null;
                AscCommon.g_clipboardBase.bSaveFormat = b_old_save_format;
            };
            if(this.checkLongActionCallback(fCallback, null)){
                fCallback();
            }
        }
        else
        {
            document.body.removeChild(_elem);
            _elem = null;
            AscCommon.g_clipboardBase.bSaveFormat = b_old_save_format;
        }
    };

    /**
     * Pastes text into the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias PasteText
     * @param {string} text - A string value that specifies the text to be pasted into the document.
     */
    Api.prototype["pluginMethod_PasteText"] = function(text)
    {
        if (!AscCommon.g_clipboardBase)
            return null;

        this.asc_PasteData(AscCommon.c_oAscClipboardDataFormat.Text, text);
    };

    /**
     * An object containing the macros data.
     * @typedef {Object} Macros
     * @property {string[]} macrosArray - An array of macros codes ([{"name": "Macros1", "value": "{macrosCode}"}]).
     * @property {number} current - A current macro index.
     */

    /**
     * Returns the document macros.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetMacros
     * @returns {Macros} - Document macros.
     */
    Api.prototype["pluginMethod_GetMacros"] = function()
    {
        return this.asc_getMacros();
    };

    /**
     * Sets macros to the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias SetMacros
     * @param {Macros} data - An object containing the macros data.
     */
    Api.prototype["pluginMethod_SetMacros"] = function(data)
    {
        return this.asc_setMacros(data);
    };

    /**
     * Specifies the start action for long operations.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias StartAction
     * @param {number} type - A value which defines an action type which can take 0 if this is an Information action or 1 if this is a BlockInteraction action.
	 * @param {string} description - A string value that specifies the description text for the start action of the operation.
     */
    Api.prototype["pluginMethod_StartAction"] = function(type, description)
    {
        this.sync_StartAction((type == "Block") ? Asc.c_oAscAsyncActionType.BlockInteraction : Asc.c_oAscAsyncActionType.Information, description);
    };

    /**
     * Specifies the end action for long operations.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias EndAction
     * @param {number} type - A value which defines an action type which can take 0 if this is an Information action or 1 if this is a BlockInteraction action.
     * @param {string} description - A string value that specifies the description text for the operation end action.
     */
    Api.prototype["pluginMethod_EndAction"] = function(type, description, status)
    {
        this.sync_EndAction((type == "Block") ? Asc.c_oAscAsyncActionType.BlockInteraction : Asc.c_oAscAsyncActionType.Information, description);

        if (window["AscDesktopEditor"] && status != null && status != "")
        {
            // error!!!
            if (!window["AscDesktopEditor"]["IsLocalFile"]())
            {
                this.sendEvent("asc_onError", "Encryption error: " + status + ". The file was not compiled.", c_oAscError.Level.Critical);
                window["AscDesktopEditor"]["CryptoMode"] = 0;
            }
            else
            {
                this.sendEvent("asc_onError", "Encryption error: " + status + ". End-to-end encryption mode is disabled.", c_oAscError.Level.NoCritical);
                window["AscDesktopEditor"]["CryptoMode"] = 0;

                if (undefined !== window.LastUserSavedIndex)
                {
                    AscCommon.History.UserSavedIndex = window.LastUserSavedIndex;

                    if (this.editorId == AscCommon.c_oEditorId.Spreadsheet)
                        this.onUpdateDocumentModified(AscCommon.History.Have_Changes());
                    else
                        this.UpdateInterfaceState();
                }
            }

            window.LastUserSavedIndex = undefined;
            setTimeout(function() {

                window["AscDesktopEditor"]["buildCryptedEnd"](false);

            }, 500);

            return;
        }

        window.LastUserSavedIndex = undefined;
        if (this._callbackPluginEndAction)
        {
            this._callbackPluginEndAction.call(this);
        }
    };

    /**
     * Encrypts the document (for crypto plugins).
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias OnEncryption
     * @param {object} obj - The encryption properties. This object can have the following values:
     * * **type** - the type of encrypting operation ("generatePassword" - generates a password for the document, "getPasswordByFile" - sends the password when opening the document, "setPasswordByFile" - sets a password to the document, "encryptData" - encrypts changes when co-editing, "decryptData" - decrypts changes when co-editing),
     * * **password** - a string value specifying the password to access the document,
     * * **data** - encrypted/decrypted changes,
     * * **check** - checks if the encryption/decryption operation is successful or not (used only for "encryptData"/"decryptData" types),
     * * **docinfo** - an unencrypted part of the encrypted file,
     * * **hash** - a string value specifying a file hash (sha256 by default),
     * * **error** - a string value specifying an error that occurs (the "" value means that the operation is successful).
     */
    Api.prototype["pluginMethod_OnEncryption"] = function(obj)
    {
        var _editor = window["Asc"]["editor"] ? window["Asc"]["editor"] : window.editor;
        switch (obj.type)
        {
            case "generatePassword":
            {
                if ("" == obj["password"])
                {
                    _editor.sendEvent("asc_onError", "There is no connection with the blockchain", c_oAscError.Level.Critical);
                    return;
                }
                if ("no_build" === obj["error"])
				{
					// проблемы - но такие, при которых просто не собираем файл...
					window["AscDesktopEditor"]["buildCryptedEnd"](true);
					return;
				}

                var _ret = _editor.getFileAsFromChanges();
                AscCommon.EncryptionWorker.isPasswordCryptoPresent = true;
                _editor.currentDocumentInfoNext = obj["docinfo"];
                window["AscDesktopEditor"]["buildCryptedStart"](_ret.data, _ret.header, obj["password"], obj["docinfo"] ? obj["docinfo"] : "");
                break;
            }
            case "getPasswordByFile":
            {
                if ("" != obj["password"])
                {
                    var _param = ("<m_sPassword>" + AscCommon.CopyPasteCorrectString(obj["password"]) + "</m_sPassword>");
                    _editor.currentPassword = obj["password"];
                    _editor.currentDocumentHash = obj["hash"];
                    _editor.currentDocumentInfo = obj["docinfo"];

                    AscCommon.EncryptionWorker.isPasswordCryptoPresent = true;

                    if (window.isNativeOpenPassword)
                    {
                        window["AscDesktopEditor"]["NativeViewerOpen"](obj["password"]);
                    }
                    else
                    {
                        window["AscDesktopEditor"]["SetAdvancedOptions"](_param);
                    }
                }
                else
                {
                    this._onNeedParams({ message : obj["message"] }, true);
                }
                break;
            }
            case "encryptData":
            case "decryptData":
            {
                AscCommon.EncryptionWorker.receiveChanges(obj);
                break;
            }
        }
    };

    /**
     * Sets the properties to the document.
	 * See {@link https://github.com/ONLYOFFICE/sdkjs-plugins/tree/master/examples/settings examples}
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias SetProperties
     * @param {object} obj - Document properties.
	 * @param {?boolean} obj.copyoutenabled - Disables copying from the editor if it is set to true.
	 * @param {?boolean} obj.hideContentControlTrack - Disables tracking the content control if it is set to true.
	 * @param {?string} obj.watermark_on_draw - A string value for watermark in JSON format.
     */
    Api.prototype["pluginMethod_SetProperties"] = function(obj)
    {
		if (!this.isDocumentLoadComplete && obj)
		{
			if (!this.setPropertiesObj)
			{
				this.setPropertiesObj = obj;
			}
			else
			{
				for (var item in obj)
					this.setPropertiesObj[item] = obj[item];
			}
			return;
		}

		if (this.setPropertiesObj)
		{
			if (!obj)
			{
				obj = this.setPropertiesObj;
			}
			else
			{
				for (var item in this.setPropertiesObj)
					if (!obj[item]) obj[item] = this.setPropertiesObj[item];
			}
			delete this.setPropertiesObj;
		}

        if (!obj)
            return;

        for (var prop in obj)
        {
            switch (prop)
            {
                case "copyoutenabled":
                {
                    this.copyOutEnabled = obj[prop];
                    this.sync_CanCopyCutCallback(this.copyOutEnabled);
                    break;
                }
                case "watermark_on_draw":
                {
                    var sText = "";
                    var tempProp = {};
                    try
                    {
                        tempProp = (typeof obj[prop] === "string") ? JSON.parse(obj[prop]) : obj[prop];
                    }
                    catch (err)
                    {
                        tempProp = {};
                    }
                    if (tempProp["paragraphs"])
                    {
                        tempProp["paragraphs"].forEach(function (el) {
                            if (el["runs"])
                            {
                                sText += el["runs"].reduce(function (accum, curel) {
                                    return accum + (curel["text"] ? curel["text"] : "");
                                }, "");
                            }
                        });
                    }
                    
                    if(!(typeof sText === "string"))
                        sText = "";
                    
                    AscFonts.FontPickerByCharacter.checkText(sText, this, function () {
                        this.watermarkDraw = obj[prop] ? new AscCommon.CWatermarkOnDraw(obj[prop], this) : null;
                        this.watermarkDraw.checkOnReady();
                    });
                    break;
                }
                case "hideContentControlTrack":
                {
                    if (this.editorId === AscCommon.c_oEditorId.Word && this.WordControl && this.WordControl.m_oLogicDocument)
                        this.WordControl.m_oLogicDocument.SetForceHideContentControlTrack(obj[prop]);

                    break;
                }
                case "disableAutostartMacros":
                {
                    this.disableAutostartMacros = true;
					break;
                }
				case "fillForms":
				{
					if (this.editorId !== AscCommon.c_oEditorId.Word
						|| !this.WordControl
						|| !this.WordControl.m_oLogicDocument)
						break;

					let oLogicDocument = this.WordControl.m_oLogicDocument;

					let oMap;
					try
					{
						oMap = (typeof obj[prop] === "string") ? JSON.parse(obj[prop]) : obj[prop];
					}
					catch (err)
					{
						oMap = {};
					}

					if (oMap["tags"])
					{
						let arrControls = oLogicDocument.GetAllContentControls();

						let oTags         = {};
						let arrCheckLocks = [];
						let arrUrls       = [];

						for (let sTag in oMap["tags"])
						{
							oTags[sTag] = [];
							for (let nIndex = 0, nCount = arrControls.length; nIndex < nCount; ++nIndex)
							{
								let oForm = arrControls[nIndex];
								if (oForm
									&& oForm.IsForm()
									&& sTag === oForm.GetTag())
								{
									oTags[sTag].push(oForm);

									let oElement;
									if (oForm.IsInlineLevel())
										oElement = oForm.GetParagraph();
									else
										oElement = oForm;

									if (oElement && -1 === arrCheckLocks.indexOf(oElement))
										arrCheckLocks.push(oElement);

									if (oMap["tags"][sTag]["picture"])
										arrUrls.push(oMap["tags"][sTag]["picture"]);
								}
							}
						}

						function FillForms()
						{
							if (!oLogicDocument.IsSelectionLocked(AscCommon.changestype_None, {
								Type      : AscCommon.changestype_2_ElementsArray_and_Type,
								Elements  : arrCheckLocks,
								CheckType : AscCommon.changestype_Paragraph_Content
							}))
							{
								oLogicDocument.StartAction(AscDFH.historydescription_Document_FillFormsByTags);

								for (let sTag in oTags)
								{
									let oValue = oMap["tags"][sTag];
									for (let nFormIndex = 0, nFormsCount = oTags[sTag].length; nFormIndex < nFormsCount; ++nFormIndex)
									{
										let oForm = oTags[sTag][nFormIndex];
										if (oForm.IsComboBox() || oForm.IsDropDownList())
										{
											if (oValue["comboBox"])
												oForm.SelectListItem(oValue["comboBox"]);
											else if (oForm.IsComboBox() && oValue["text"])
												oForm.SetInnerText(oValue["text"]);
										}
										else if (oForm.IsPictureForm())
										{
											let oPicture = oForm.GetPicture();
											if (oValue["picture"] && oPicture)
												oPicture.setBlipFill(AscFormat.CreateBlipFillRasterImageId(oValue["picture"]));
										}
										else if (oForm.IsCheckBox())
										{
											if (oValue["checkBox"])
												oForm.SetCheckBoxChecked(!!oValue["checkBox"]);
										}
										else
										{
											oForm.SetInnerText(oValue["text"]);
										}
									}
								}

								oLogicDocument.Recalculate();
								oLogicDocument.FinalizeAction();
							}
						}

						if (this.ImageLoader && arrUrls.length)
							this.ImageLoader.LoadImagesWithCallback(arrUrls, FillForms, []);
						else
							FillForms();
					}

					break
				}
                default:
                    break;
            }
        }
    };

    /**
     * Shows the input helper.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias ShowInputHelper
     * @param {string} guid - A string value which specifies a plugin identifier which must be of the asc.{UUID} type.
     * @param {number} w - A number which specifies the window width measured in millimeters.
     * @param {number} h - A number which specifies the window height measured in millimeters.
     * @param {boolean} isKeyboardTake - Defines if the keyboard is caught (true) or not (false).
     */
    Api.prototype["pluginMethod_ShowInputHelper"] = function(guid, w, h, isKeyboardTake)
    {
        var _frame = document.getElementById("iframe_" + guid);
        if (!_frame)
            return;

        var _offset = this.getTargetOnBodyCoords();
        if (w > _offset.W)
            w = _offset.W;
        if (h > _offset.H)
            h = _offset.H;

        var _offsetToFrame = 10;
        var _r = _offset.X + _offsetToFrame + w;
        var _t = _offset.Y - _offsetToFrame - h;
        var _b = _offset.Y + _offset.TargetH + _offsetToFrame + h;

        var _x = _offset.X + _offsetToFrame;
        if (_r > _offset.W)
            _x += (_offset.W - _r);

        var _y = 0;

        if (_b < _offset.H)
        {
            _y = _offset.Y + _offset.TargetH + _offsetToFrame;
        }
        else if (_t > 0)
        {
            _y = _t;
        }
        else
        {
            _y = _offset.Y + _offset.TargetH + _offsetToFrame;
            h += (_offset.H - _b);
        }

        _frame.style.left = _x + "px";
        _frame.style.top = _y + "px";
        _frame.style.width = w + "px";
        _frame.style.height = h + "px";

        if (!this.isMobileVersion)
            _frame.style.zIndex = 1000;
        else
            _frame.style.zIndex = 5001;

        if (!_frame.style.boxShadow)
        {
            _frame.style.boxShadow = "0 6px 12px rgba(0, 0, 0, 0.175)";
            _frame.style.webkitBoxShadow = "0 6px 12px rgba(0, 0, 0, 0.175)";
            //_frame.style.borderRadius = "3px";
        }


        if (isKeyboardTake)
        {
            _frame.setAttribute("oo_editor_input", "true");
            _frame.focus();
        }
        else
        {
            _frame.removeAttribute("oo_editor_input");
            if (AscCommon.g_inputContext)
            {
                AscCommon.g_inputContext.isNoClearOnFocus = true;
                AscCommon.g_inputContext.HtmlArea.focus();
            }
        }

        if (AscCommon.g_inputContext)
        {
            AscCommon.g_inputContext.isInputHelpersPresent = true;
            AscCommon.g_inputContext.isInputHelpers[guid] = true;
        }
    };

    /**
     * Unshows the input helper.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias UnShowInputHelper
     * @param {string} guid - A string value which specifies a plugin identifier which must be of the asc.{UUID} type.
     * @param {string} isclear - Defines if the input context will be cleared (true) or not (false).
     */
    Api.prototype["pluginMethod_UnShowInputHelper"] = function(guid, isclear)
    {
        var _frame = document.getElementById("iframe_" + guid);
        if (!_frame)
            return;

        _frame.style.width = "10px";
        _frame.style.height = "10px";
        _frame.removeAttribute("oo_editor_input");

        _frame.style.zIndex = -1000;

        if (AscCommon.g_inputContext && AscCommon.g_inputContext.HtmlArea)
        {
            AscCommon.g_inputContext.HtmlArea.focus();

            if (AscCommon.g_inputContext.isInputHelpers[guid])
                delete AscCommon.g_inputContext.isInputHelpers[guid];

            var count = 0;
            for (var test in AscCommon.g_inputContext.isInputHelpers)
            {
                if (AscCommon.g_inputContext.isInputHelpers[test])
                    count++;
            }

            AscCommon.g_inputContext.isInputHelpersPresent = (0 != count);
        }

        if (AscCommon.g_inputContext && isclear)
        {
            AscCommon.g_inputContext.keyPressInput = "";
        }
    };
    /**
     * Sends a message to the co-authoring chat.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias CoAuthoringChatSendMessage
     * @param {string} sText - Message text.
     */
    Api.prototype["pluginMethod_CoAuthoringChatSendMessage"] = function(sText)
    {
        return this.CoAuthoringChatSendMessage(sText);
    };

	/**
	 * The current selection type.
	 * @typedef {("none" | "text" | "drawing" | "slide")} SelectionType
	 */

	/**
	 * Returns the type of the current selection.
	 * @memberof Api
	 * @typeofeditors ["CDE", "CSE", "CPE"]
	 * @alias GetSelectionType
	 * @returns {SelectionType} - The selection type.
	 */
	Api.prototype["pluginMethod_GetSelectionType"] = function()
	{
		switch (this.editorId)
		{
			case AscCommon.c_oEditorId.Word:
			{
				if (!this.WordControl || !this.WordControl.m_oLogicDocument)
					return "none";
				var logicDoc = this.WordControl.m_oLogicDocument;

				if (!logicDoc.IsSelectionUse())
					return "none";

				var selectionBounds = logicDoc.GetSelectionBounds();
				var eps = 0.0001;
				if (selectionBounds && selectionBounds.Start && selectionBounds.End &&
					(Math.abs(selectionBounds.Start.W) > eps) &&
					(Math.abs(selectionBounds.End.W) > eps))
				{
					return "text";
				}

				if (logicDoc.DrawingObjects.getSelectedObjectsBounds())
					return "drawing";

				return "none";
			}
			case AscCommon.c_oEditorId.Presentation:
			{
				if (!this.WordControl || !this.WordControl.m_oLogicDocument)
					return "none";
				var logicDoc = this.WordControl.m_oLogicDocument;

				if (-1 === logicDoc.CurPage)
					return "none";

				var _controller = logicDoc.Slides[logicDoc.CurPage].graphicObjects;
				var _elementsCount = _controller.selectedObjects.length;

				var retType = "slide";
				if (!_controller.IsSelectionUse() && _elementsCount > 0)
					retType = "none";

				var selectionBounds = logicDoc.GetSelectionBounds();
				var eps = 0.0001;
				if (selectionBounds && selectionBounds.Start && selectionBounds.End &&
					(Math.abs(selectionBounds.Start.W) > eps) &&
					(Math.abs(selectionBounds.End.W) > eps))
				{
					return "text";
				}
				
				if (retType === "slide" && _controller.getSelectedObjectsBounds())
					retType = "drawing";

				return retType;
			}
			case AscCommon.c_oEditorId.Spreadsheet:
			{
				if (!this.wb || !this.wb.getWorksheet())
					return "none";

				var objectRender = this.wb.getWorksheet().objectRender;
				if (!objectRender)
					return "none";

				var controller = objectRender.controller;
				if (!controller)
					return "none";

				var selection = this.wb.GetSelectionRectsBounds();
				var retType = "none";

				if (!controller.IsSelectionUse() && !selection)
					retType = "none";
				if (controller.GetSelectionBounds() || selection)
					retType = "text";
				if (controller.getSelectedObjectsBounds())
					retType = "drawing";

				return retType;
			}
			default:
				break;
		}
		return "none";
	};
     /**
     * Converts a document to Markdown or HTML text.
     * @memberof Api
     * @typeofeditors ["CDE"]
     * @alias ConvertDocument
     * @param {"markdown" | "html"} [sConvertType="markdown"] - Conversion type ("markdown" or "html").
     * @param {bool} [bHtmlHeadings=false] - Defines if the HTML headings and IDs will be generated when the Markdown renderer of your target platform does not handle Markdown-style IDs.
	 * @param {bool} [bBase64img=false] - Defines if the images will be created in the base64 format.
	 * @param {bool} [bDemoteHeadings=false] - Defines if all heading levels in your document will be demoted to conform with the following standard: single H1 as title, H2 as top-level heading in the text body.
	 * @param {bool} [bRenderHTMLTags=false] - Defines if HTML tags will be preserved in your Markdown. If you just want to use an occasional HTML tag, you can avoid using the opening angle bracket in the following way: \<tag>text\</tag>. By default, the opening angle brackets will be replaced with the special characters.
     */
    Api.prototype["pluginMethod_ConvertDocument"] = function(sConvertType, bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags)
    {
        return this.ConvertDocument(sConvertType, bHtmlHeadings, bBase64img, bDemoteHeadings, bRenderHTMLTags);
    };
    /**
     * Returns the selected text from the document.
     * @memberof Api
     * @typeofeditors ["CDE", "CPE", "CSE"]
     * @alias GetSelectedText
     * @param {prop} numbering - The resulting string display properties:
     * * **NewLine** - defines if the resulting string will include line boundaries or not,
     * * **NewLineParagraph** - defines if the resulting string will include paragraph line boundaries or not,
     * * **Numbering** - defines if the resulting string will include numbering or not.
     * @return {string} - Selected text.
     * @example
     * window.Asc.plugin.executeMethod("GetSelectedText", [{NewLine:true, NewLineParagraph:true, Numbering:true}])
     */
    Api.prototype["pluginMethod_GetSelectedText"] = function(prop)
    {
        var properties;
        if (typeof prop === "object")
        {
            properties =
            {
                NewLine : (prop.hasOwnProperty("NewLine")) ? prop["NewLine"] : true,
                NewLineParagraph : (prop.hasOwnProperty("NewLineParagraph")) ? prop["NewLineParagraph"] : true,
                Numbering : (prop.hasOwnProperty("Numbering")) ? prop["Numbering"] : true,
                Math : (prop.hasOwnProperty("Math")) ? prop["Math"] : true,
                TableCellSeparator: prop["TableCellSeparator"],
                TableRowSeparator: prop["TableRowSeparator"],
                ParaSeparator: prop["ParaSeparator"],
                TabSymbol: prop["TabSymbol"]
            }
        }
        else
        {
            properties =
            {
                NewLine : true,
                NewLineParagraph : true,
                Numbering : true
            }
        }

        return this.asc_GetSelectedText(false, properties);
    };
    /**
     * Replaces each paragraph (or text in cell) in the select with the corresponding text from an array of strings.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias ReplaceTextSmart
     * @param {Array} arrString - An array of replacement strings.
     * @param {string} [sParaTab=" "] - A character which is used to specify the tab in the source text.
     * @param {string} [sParaNewLine=" "] - A character which is used to specify the line break character in the source text.
     * @returns {boolean} - Always returns true.
     */
    Api.prototype["pluginMethod_ReplaceTextSmart"] = function(arrString, sParaTab, sParaNewLine)
    {
		let guid = window.g_asc_plugins ? window.g_asc_plugins.setPluginMethodReturnAsync() : null;
		this.incrementCounterLongAction();

		function ReplaceTextSmart()
		{
			this.asc_canPaste();
			this.ReplaceTextSmart(arrString, sParaTab, sParaNewLine);
			this.asc_Recalculate(true);
			switch (this.editorId)
			{
				case AscCommon.c_oEditorId.Spreadsheet:
					this.asc_endPaste();
					break;
				case AscCommon.c_oEditorId.Word:
				case AscCommon.c_oEditorId.Presentation:
					this.WordControl.m_oLogicDocument.FinalizeAction();
					break;
			}

			this.decrementCounterLongAction();

			if (guid)
				window.g_asc_plugins.onPluginMethodReturn(guid, true);
		}

		let sOverAll = "";
		for (let nIndex = 0, nCount = arrString.length; nIndex < nCount; ++nIndex)
			sOverAll += arrString[nIndex];

		AscFonts.FontPickerByCharacter.checkText(sOverAll, this, ReplaceTextSmart);
        return true;
    };
	/**
     * Returns the current file to download in the specified format.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetFileToDownload
     * @param {string} [format=" "] - The format in which you need to download the file.
     * @returns {string} - URL to download the file in the specified format or error.
     */
	Api.prototype["pluginMethod_GetFileToDownload"] = function(format)
	{
		let guid = window.g_asc_plugins ? window.g_asc_plugins.setPluginMethodReturnAsync() : null;
		let dwnldF = Asc.c_oAscFileType[format] || Asc.c_oAscFileType[this.DocInfo.Format.toUpperCase()];
		let opts = new Asc.asc_CDownloadOptions(dwnldF);
		let _t = this;
		opts.callback = function() {
			_t.sync_EndAction(Asc.c_oAscAsyncActionType.BlockInteraction, Asc.c_oAscAsyncAction.DownloadAs);
			_t.fCurCallback = function(res) {
				let data = (res.status == "ok") ? res.data : "error";
				if (guid)
					window.g_asc_plugins.onPluginMethodReturn(guid, data);
			};
		}
		this.downloadAs(Asc.c_oAscAsyncAction.DownloadAs, opts);
	};


    /**
     * An object containing the font information.
     * @typedef {Object} ImageData
     * @property {string} src
     * @property {number} width
     * @property {number} height
     */

	/**
     * Returns an image data obtained from first of selected drawings.
     * If there are no selected drawings it returns white rect.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetImageDataFromSelection
     * @returns {?ImageData} - ImageData with png image encoded in base64 format or null if there are no selected objects.
     */
	Api.prototype["pluginMethod_GetImageDataFromSelection"] = function()
	{
		return this.getImageDataFromSelection();
	};
	/**
     * Replaces the first selected drawing
     * If there are no selected drawings it inserts the image to the current position.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias PutImageDataToSelection
     * @param {ImageData} oImageData - image encoded in base64 format.
     */
	Api.prototype["pluginMethod_PutImageDataToSelection"] = function(oImageData)
	{
        this._beforeEvalCommand();
		this.putImageToSelection(oImageData["src"], oImageData["width"], oImageData["height"]);
        this._afterEvalCommand();
	};

	function getLocalStorageItem(key)
	{
		try
		{
			return JSON.parse(window.localStorage.getItem(key));
		}
		catch (e)
		{
			return null;
		}
	}
	function setLocalStorageItem(key, value)
	{
		try
		{
			window.localStorage.setItem(key, JSON.stringify(value));
			return true;
		}
		catch (e)
		{
		}
		return false;
	}

	function installPlugin(config, loadFuncName)
	{
		if (!config)
		{
			return {
				"type" : loadFuncName,
				"guid" : ""
			};
		}

		let currentInstalledPlugins = getLocalStorageItem("asc_plugins_installed");
		if (!currentInstalledPlugins)
			currentInstalledPlugins = {};
		currentInstalledPlugins[config["guid"]] = config;
		setLocalStorageItem("asc_plugins_installed", currentInstalledPlugins);

		let currentRemovedPlugins = getLocalStorageItem("asc_plugins_removed");
		if (currentRemovedPlugins && currentRemovedPlugins[config["guid"]])
		{
			delete currentRemovedPlugins[config["guid"]];
			setLocalStorageItem("asc_plugins_removed", currentRemovedPlugins);
		}

		window.g_asc_plugins.api.disableCheckInstalledPlugins = true;
		window.g_asc_plugins.loadExtensionPlugins([config], true);
		delete window.g_asc_plugins.api.disableCheckInstalledPlugins;

		return {
			"type" : loadFuncName,
			"guid" : config["guid"]
		};
	}

	Api.prototype.checkInstalledPlugins = function()
	{
		if (this.disableCheckInstalledPlugins)
			return;

		let arrayPlugins = [];

		let currentInstalledPlugins = getLocalStorageItem("asc_plugins_installed");
		if (currentInstalledPlugins)
		{
			for (let item in currentInstalledPlugins)
			{
				if (currentInstalledPlugins[item]["guid"])
					arrayPlugins.push(currentInstalledPlugins[item]);
			}
		}

		if (window["Asc"]["extensionPlugins"] && window["Asc"]["extensionPlugins"].length)
		{
			let arrayExtensions = window["Asc"]["extensionPlugins"];
			for (let i = 0, len = arrayExtensions.length; i < len; i++)
				arrayPlugins.push(arrayExtensions[i]);
		}

		let isInstalledPresent = window.g_asc_plugins.loadExtensionPlugins(arrayPlugins, undefined, true);

		let isRemovedPresent = false;
		let currentRemovedPlugins = getLocalStorageItem("asc_plugins_removed");

		if (currentRemovedPlugins)
		{
			for (let guid in currentRemovedPlugins)
			{
				if (guid)
				{
					if (window.g_asc_plugins.unregister(guid))
						isRemovedPresent = true;
				}
			}
		}

		// этот метод может быть вызван из интерфейса - нужен таймаут для web-apps
		if (isRemovedPresent || isInstalledPresent) {

			setTimeout(function () {

				// в принципе можно не удалять, так как если ничего не поменялось - то не зайдем второй раз сюда.
				// но зачем еще раз парсить
				window.g_asc_plugins.api.disableCheckInstalledPlugins = true;

				if (isRemovedPresent)
					window.g_asc_plugins.api.sendEvent("asc_onPluginsReset");

				if (isInstalledPresent || isRemovedPresent)
					window.g_asc_plugins.updateInterface();

				delete window.g_asc_plugins.api.disableCheckInstalledPlugins;

			}, 10);

		}
	};

	/**
    * Returns all installed plugins.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @alias GetInstalledPlugins
     * @returns {[]} - Array of all installed plugins.
     */
	Api.prototype["pluginMethod_GetInstalledPlugins"] = function()
	{
		/*
			формат объекта 
			{
				url: url на конфиг (хотя по факту он не нужен, так как конфиг есть в этом объекте и внутри маркетплейса тоже),
				guid: guid плагина,
				canRemoved: флаг, может ли быть удалён плагин или нет (true/false),
				obj: конфиг установленного плагина (от туда берется версия и сравнивается с текущей для проверки обновлений)
			}
		*/

		let baseUrl = window.location.href;
		let posQ = baseUrl.indexOf("?");
		if (-1 !== posQ)
			baseUrl = baseUrl.substr(0, posQ);

		let pluginsArray = window.g_asc_plugins.plugins.concat(window.g_asc_plugins.systemPlugins);
		let returnArray = [];

		for (let i = 0, len = pluginsArray.length; i < len; i++)
		{
			if (pluginsArray[i].isConnector)
				continue;
			returnArray.push({
				"baseUrl" : baseUrl,
				"guid" : pluginsArray[i].guid,
				"canRemoved" : true,
				"obj" : pluginsArray[i].serialize(),
				"removed" : false
			});
		}

		// нужно послать и удаленные. так как удаленный может не быть в сторе. тогда его никак не установить обратно
		let currentRemovedPlugins = getLocalStorageItem("asc_plugins_removed");

		if (currentRemovedPlugins)
		{
			for (let guid in currentRemovedPlugins)
			{
				if (currentRemovedPlugins[guid])
				{
					returnArray.push({
						"baseUrl" : baseUrl,
						"guid": currentRemovedPlugins[guid]["guid"],
						"canRemoved": true,
						"obj": currentRemovedPlugins[guid],
						"removed" : true
					});
				}
			}
		}

		return returnArray;
	};
	/**
    * Remove plugin with such guid.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @param {string} [guid] - The guid of plugins for removing.
     * @alias RemovePlugin
     * @returns {object} - Object with result.
     */
	Api.prototype["pluginMethod_RemovePlugin"] = function(guid)
	{
		let removedPlugin = window.g_asc_plugins.unregister(guid);

		if (removedPlugin)
		{
			let currentRemovedPlugins = getLocalStorageItem("asc_plugins_removed");
			if (!currentRemovedPlugins)
				currentRemovedPlugins = {};
			currentRemovedPlugins[removedPlugin.guid] = removedPlugin.serialize();
			setLocalStorageItem("asc_plugins_removed", currentRemovedPlugins);

			let currentInstalledPlugins = getLocalStorageItem("asc_plugins_installed");
			if (currentInstalledPlugins && currentInstalledPlugins[removedPlugin.guid])
			{
				delete currentInstalledPlugins[removedPlugin.guid];
				setLocalStorageItem("asc_plugins_installed", currentInstalledPlugins);
			}

			this.disableCheckInstalledPlugins = true;

			this.sendEvent("asc_onPluginsReset");
			window.g_asc_plugins.updateInterface();

			delete this.disableCheckInstalledPlugins;
		}

		return {
			type : "Removed",
			guid : removedPlugin ? removedPlugin.guid : ""
		};
	};
	/**
    * Install plugin with by url to config.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @param {string} [url] - The url to plugin config for installing.
     * @alias pluginMethod_InstallPlugin
     * @returns {object} - Object with result.
	 * 
     */
	Api.prototype["pluginMethod_InstallPlugin"] = function(config)
	{
		return installPlugin(config, "Installed");
	};
	/**
    * Update plugin with by url to config.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @param {string} [url] - The url to plugin config for updating.
     * @alias pluginMethod_UpdatePlugin
     * @returns {object} - Object with result.
	 * 
     */
	Api.prototype["pluginMethod_UpdatePlugin"] = function(url, guid)
	{
		return installPlugin(config, "Updated");
	};

	/**
    * Show or hide buttons in header.
     * @memberof Api
     * @typeofeditors ["CDE", "CSE", "CPE"]
     * @param {string} [id] - The id of the button.
     * @param {boolean} [bShow] - The flag show or hide the button.
     * @alias ShowButton 
     */
	Api.prototype["pluginMethod_ShowButton"] = function(id, bShow)
	{
		if (bShow) {
			this.sendEvent("asc_onPluginShowButton", id);
		} else {
			this.sendEvent("asc_onPluginHideButton", id);
		}
	};
})(window);
